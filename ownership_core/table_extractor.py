"""
ownership_core.table_extractor
==============================
Trích xuất bảng biểu có cấu trúc giữ nguyên bố cục không gian (tọa độ):
- Quét toàn bộ các trang (không bỏ vĩnh viễn trang nào).
- Nhận diện bảng qua fitz.find_tables() kết hợp gom hàng/cột từ tọa độ từ (coordinate-based).
- Xử lý tiêu đề nhiều tầng, bảng không đường kẻ, ngày gắn với bảng và đơn vị.
"""

from __future__ import annotations
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Set
import fitz
from ownership_core.config import (
    PRIMARY_KEYWORDS,
    SECONDARY_KEYWORDS,
    COLUMN_SHARE_RATIO_KEYWORDS,
    COLUMN_SHAREHOLDER_COUNT_KEYWORDS,
    normalize_text,
)
from ownership_core.cache_manager import CacheManager
from ownership_core.ocr_engine import OCREngine


@dataclass
class ExtractedCell:
    text: str
    bbox: Optional[Tuple[float, float, float, float]] = None


@dataclass
class StructuredTable:
    page_number: int
    header_rows: List[List[str]]
    data_rows: List[List[str]]
    raw_markdown: str
    table_date: Optional[str] = None
    date_context: str = ""
    share_ratio_col_indices: List[int] = field(default_factory=list)
    shareholder_count_col_indices: List[int] = field(default_factory=list)
    share_count_col_indices: List[int] = field(default_factory=list)
    holder_col_index: int = 0


@dataclass
class ScoredPage:
    page_number: int
    score: int
    matched_keywords: List[str]
    has_text_layer: bool
    text: str


DATE_REGEX = re.compile(r"(?<!\d)(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(20\d{2})(?!\d)")
DATE_LEAD_TERMS = ["tai ngay", "chot ngay", "den ngay", "ngay", "as of", "as at", "thoi diem", "danh sach chot"]


def extract_nearby_date(text: str, report_year: int) -> Tuple[Optional[str], str]:
    """Tìm ngày gắn với bảng từ đoạn văn xung quanh bảng."""
    norm = normalize_text(text)
    for lead in DATE_LEAD_TERMS:
        if lead in norm:
            idx = norm.find(lead)
            sub = text[max(0, idx - 10):min(len(text), idx + 80)]
            matches = list(DATE_REGEX.finditer(sub))
            if matches:
                m = matches[0]
                day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if 1 <= day <= 31 and 1 <= month <= 12 and 2019 <= year <= 2026:
                    iso_date = f"{year:04d}-{month:02d}-{day:02d}"
                    return iso_date, f"Ngày theo ngữ cảnh: '{lead}'"

    # Fallback: tìm bất kỳ ngày nào trong đoạn văn gần năm báo cáo
    for m in DATE_REGEX.finditer(text):
        day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12 and year == report_year:
            iso_date = f"{year:04d}-{month:02d}-{day:02d}"
            return iso_date, "Ngày tìm thấy gần bảng"

    return None, "Không tìm thấy ngày cụ thể"


def reconstruct_table_from_words(words: List[Tuple]) -> List[List[str]]:
    """
    Gom cụm từ tọa độ từ (x0, y0, x1, y1, word, block_no, line_no, word_no)
    thành bảng có hàng và cột thẳng hàng.
    """
    if not words:
        return []

    # 1. Gom các từ theo hàng dựa trên tọa độ y0 (sai số dung sai 3-4pt)
    rows_dict: Dict[int, List[Tuple]] = {}
    sorted_words = sorted(words, key=lambda w: (w[1], w[0]))  # Sắp theo y0, rồi x0
    
    current_y = None
    current_row = []
    
    for w in sorted_words:
        y0 = w[1]
        if current_y is None:
            current_y = y0
            current_row.append(w)
        elif abs(y0 - current_y) <= 4.0:
            current_row.append(w)
        else:
            if current_row:
                current_row.sort(key=lambda x: x[0])
                rows_dict[len(rows_dict)] = current_row
            current_y = y0
            current_row = [w]
            
    if current_row:
        current_row.sort(key=lambda x: x[0])
        rows_dict[len(rows_dict)] = current_row

    # 2. Xác định các cột dựa trên tọa độ x0 của các từ
    raw_lines: List[List[str]] = []
    for r_idx, w_list in sorted(rows_dict.items()):
        # Nối các từ cách nhau hợp lý thành ô
        line_cells: List[str] = []
        cur_cell = []
        last_x1 = None
        for w in w_list:
            word_str = w[4]
            x0 = w[0]
            if last_x1 is None:
                cur_cell.append(word_str)
            elif x0 - last_x1 > 15.0:  # Khoảng cách cột > 15pt
                line_cells.append(" ".join(cur_cell))
                cur_cell = [word_str]
            else:
                cur_cell.append(word_str)
            last_x1 = w[2]
        if cur_cell:
            line_cells.append(" ".join(cur_cell))
        if len(line_cells) >= 2:
            raw_lines.append(line_cells)

    return raw_lines


class TableExtractor:
    """Bộ bóc tách bảng PDF giữ nguyên cấu trúc không gian và ngữ nghĩa cột."""

    def __init__(self, cache_manager: Optional[CacheManager] = None, ocr_engine: Optional[OCREngine] = None):
        self.cache_manager = cache_manager or CacheManager()
        self.ocr_engine = ocr_engine or OCREngine(self.cache_manager)

    def scan_all_pages(self, doc: fitz.Document, file_hash: str) -> List[ScoredPage]:
        """
        Quét lớp chữ TOÀN BỘ các trang của PDF.
        Chấm điểm từ khóa để ưu tiên, KHÔNG loại vĩnh viễn trang nào.
        """
        scored_pages: List[ScoredPage] = []
        total_pages = len(doc)

        for pno in range(total_pages):
            page_no = pno + 1
            # 1. Kiểm tra cache lớp chữ
            cached_text = self.cache_manager.get_page_text(file_hash, page_no)
            if cached_text is None:
                try:
                    p_text = doc[pno].get_text()
                except Exception:
                    p_text = ""
                self.cache_manager.set_page_text(file_hash, page_no, p_text)
            else:
                p_text = cached_text

            char_len = len(p_text.strip())
            has_text = char_len >= 80

            # 2. Chấm điểm từ khóa
            norm = normalize_text(p_text)
            score = 0
            matched = []

            for kw in PRIMARY_KEYWORDS:
                norm_kw = normalize_text(kw)
                cnt = norm.count(norm_kw)
                if cnt > 0:
                    score += cnt * 25
                    matched.append(kw)

            for kw in SECONDARY_KEYWORDS:
                norm_kw = normalize_text(kw)
                cnt = norm.count(norm_kw)
                if cnt > 0:
                    score += cnt * 5
                    matched.append(kw)

            scored_pages.append(ScoredPage(
                page_number=page_no,
                score=score,
                matched_keywords=list(set(matched)),
                has_text_layer=has_text,
                text=p_text
            ))

        return scored_pages

    def extract_structured_tables(
        self,
        doc: fitz.Document,
        scored_pages: List[ScoredPage],
        report_year: int
    ) -> List[StructuredTable]:
        """
        Trích xuất danh sách bảng có cấu trúc từ các trang ứng viên tiềm năng nhất.
        """
        # Ưu tiên các trang có điểm từ khóa > 0
        candidate_pages = [p for p in scored_pages if p.score > 0]
        # Nếu không có trang nào có điểm > 0, chọn các trang đầu và các trang có từ khóa lỏng
        if not candidate_pages:
            candidate_pages = scored_pages[:min(10, len(scored_pages))]

        # Thêm trang liền sau của mỗi trang candidate đề phòng bảng nối trang
        target_page_numbers: Set[int] = set()
        for p in candidate_pages:
            target_page_numbers.add(p.page_number)
            if p.page_number < len(doc):
                target_page_numbers.add(p.page_number + 1)

        extracted_tables: List[StructuredTable] = []

        for p_num in sorted(list(target_page_numbers)):
            pno = p_num - 1
            page = doc[pno]
            page_text = scored_pages[pno].text if pno < len(scored_pages) else page.get_text()

            # Nếu trang thiếu chữ (< 80 ký tự), dùng OCR
            if len(page_text.strip()) < 80 and self.ocr_engine.is_available:
                ocr_txt = self.ocr_engine.ocr_page(page)
                if len(ocr_txt.strip()) > len(page_text.strip()):
                    page_text = ocr_txt

            # 1. Thử dùng fitz.find_tables()
            fitz_tables = []
            try:
                tabs = page.find_tables()
                if tabs and len(tabs.tables) > 0:
                    for tab in tabs:
                        tab_data = tab.extract()
                        if tab_data and len(tab_data) >= 2:
                            fitz_tables.append(tab_data)
            except Exception:
                pass

            # 2. Nếu find_tables() không ra, dùng coordinate words reconstruction
            if not fitz_tables:
                try:
                    words = page.get_text("words")
                    reconstructed = reconstruct_table_from_words(words)
                    if reconstructed and len(reconstructed) >= 3:
                        fitz_tables.append(reconstructed)
                except Exception:
                    pass

            # 3. Phân tích ngữ nghĩa cột và cấu trúc của từng bảng tìm được
            for raw_tab in fitz_tables:
                if not raw_tab or len(raw_tab) < 2:
                    continue

                header_rows = [raw_tab[0]]
                # Kiểm tra nếu dòng 2 là tiêu đề phụ
                data_start_idx = 1
                if len(raw_tab) > 2 and any("%" in str(c) for c in raw_tab[1]):
                    header_rows.append(raw_tab[1])
                    data_start_idx = 2

                data_rows = raw_tab[data_start_idx:]
                header_flat = " ".join([str(c) for row in header_rows for c in row if c])
                
                # Xác định ngày gắn với bảng
                context_window = page_text[:1200]
                t_date, date_ctx = extract_nearby_date(context_window, report_year)

                # Phân tích chỉ số các cột
                share_ratio_cols = []
                count_ratio_cols = []
                share_count_cols = []

                first_row = header_rows[0]
                for c_idx, cell_val in enumerate(first_row):
                    c_norm = normalize_text(str(cell_val or ""))
                    # Nếu có dòng header phụ, ghép thêm
                    if len(header_rows) > 1 and c_idx < len(header_rows[1]):
                        c_norm += " " + normalize_text(str(header_rows[1][c_idx] or ""))

                    # Kiểm tra cột tỷ lệ cổ phần
                    if any(k in c_norm for k in COLUMN_SHARE_RATIO_KEYWORDS) or "%" in c_norm:
                        if not any(bad in c_norm for bad in COLUMN_SHAREHOLDER_COUNT_KEYWORDS):
                            share_ratio_cols.append(c_idx)

                    # Kiểm tra cột tỷ lệ số lượng cổ đông (CẦN LOẠI TRỪ)
                    if any(k in c_norm for k in COLUMN_SHAREHOLDER_COUNT_KEYWORDS):
                        count_ratio_cols.append(c_idx)

                    # Kiểm tra cột số lượng cổ phiếu
                    if any(k in c_norm for k in ["so co phan", "so co phieu", "so luong cp", "number of shares"]):
                        if not any(bad in c_norm for bad in COLUMN_SHAREHOLDER_COUNT_KEYWORDS):
                            share_count_cols.append(c_idx)

                # Tạo định dạng markdown để lưu vết hoặc cấp cho model nếu cần
                md_lines = []
                for row in raw_tab:
                    md_lines.append("| " + " | ".join(str(c or "").strip().replace("\n", " ") for c in row) + " |")
                raw_md = "\n".join(md_lines)

                extracted_tables.append(StructuredTable(
                    page_number=p_num,
                    header_rows=header_rows,
                    data_rows=data_rows,
                    raw_markdown=raw_md,
                    table_date=t_date,
                    date_context=date_ctx,
                    share_ratio_col_indices=share_ratio_cols,
                    shareholder_count_col_indices=count_ratio_cols,
                    share_count_col_indices=share_count_cols,
                    holder_col_index=0
                ))

        return extracted_tables
