"""
ownership_core.inventory
========================
Quản lý kiểm kê dữ liệu 101 mã chứng khoán x 6 năm (2020-2025) = 606 mục.
Đối chiếu danh mục CSV với file thực tế trên đĩa, nhận diện file trùng và chọn file chuẩn.
"""

from __future__ import annotations
import csv
import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
from ownership_core.config import DEFAULT_PDF_DIR, DEFAULT_MANIFEST


@dataclass
class InventoryItem:
    stt: int
    ticker: str
    year: int
    company_name: str
    file_path: Optional[str]
    file_name: Optional[str]
    file_size_bytes: int
    page_count: int
    status_manifest: str
    is_disk_available: bool
    notes: str = ""

    def to_dict(self) -> Dict:
        return asdict(self)


def is_annual_report_filename(name: str) -> bool:
    """Kiểm tra tên file có chứa dấu hiệu Báo cáo thường niên không."""
    low = name.lower()
    return any(k in low for k in ["baocaothuongnien", "bctn", "annual", "báo cáo thường niên", "thuong_nien"])


def score_pdf_candidate(file_path: Path) -> int:
    """Chấm điểm ưu tiên chọn file BCTN chuẩn nếu có nhiều file cùng mã-năm."""
    score = 0
    name = file_path.name.lower()
    if is_annual_report_filename(name):
        score += 50
    if not any(bad in name for bad in ["phu_luc", "soat_xet", "bien_ban", "nghi_quyet", "cv", "tb"]):
        score += 30
    try:
        size = file_path.stat().st_size
        if size > 100 * 1024:  # > 100 KB
            score += 20
        if size > 1024 * 1024:  # > 1 MB
            score += 10
    except Exception:
        pass
    return score


def load_canonical_manifest(
    manifest_csv: Path = DEFAULT_MANIFEST,
    pdf_dir: Path = DEFAULT_PDF_DIR
) -> List[InventoryItem]:
    """
    Tải danh mục 606 mục (101 mã x 6 năm 2020-2025).
    Kiểm tra thực tế từng file trên đĩa.
    """
    if not manifest_csv.exists():
        raise FileNotFoundError(f"Không tìm thấy file danh mục CSV: {manifest_csv}")

    # 1. Đọc dữ liệu từ manifest CSV
    raw_manifest_entries: List[Dict] = []
    with open(manifest_csv, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            ticker = row.get("Mã CK", "").strip().upper()
            try:
                year = int(row.get("Năm", 0))
            except ValueError:
                year = 0
            company = row.get("Tên công ty", "").strip()
            manifest_status = row.get("Trạng thái", "").strip()
            raw_manifest_entries.append({
                "ticker": ticker,
                "year": year,
                "company": company,
                "manifest_status": manifest_status,
                "file_manifest": row.get("File PDF", "").strip()
            })

    # Lấy danh sách 101 mã duy nhất
    unique_tickers = sorted(list({e["ticker"] for e in raw_manifest_entries if e["ticker"]}))
    target_years = [2020, 2021, 2022, 2023, 2024, 2025]

    # Map tên công ty theo ticker
    company_names = {}
    for e in raw_manifest_entries:
        if e["ticker"] and e["company"] and e["ticker"] not in company_names:
            company_names[e["ticker"]] = e["company"]

    # 2. Quét thực tế tất cả các file PDF trên đĩa
    disk_pdfs_by_pair: Dict[Tuple[str, int], List[Path]] = {}
    if pdf_dir.exists():
        for p in pdf_dir.rglob("*.pdf"):
            if not p.is_file():
                continue
            # Xác định ticker từ thư mục cha hoặc tên file
            parent_name = p.parent.name.strip().upper()
            m_ticker = parent_name if len(parent_name) == 3 and parent_name.isalnum() else None
            if not m_ticker:
                match_t = re.search(r"^([A-Z0-9]{3})[_\-\s]", p.name, re.IGNORECASE)
                if match_t:
                    m_ticker = match_t.group(1).upper()
            
            # Xác định năm từ tên file
            match_y = re.search(r"(202[0-5])", p.name)
            m_year = int(match_y.group(1)) if match_y else None

            if m_ticker and m_year and m_year in target_years:
                pair = (m_ticker, m_year)
                disk_pdfs_by_pair.setdefault(pair, []).append(p)

    # 3. Tổng hợp thành danh sách 606 mục chuẩn mực
    inventory: List[InventoryItem] = []
    stt = 1
    for ticker in unique_tickers:
        for year in target_years:
            pair = (ticker, year)
            c_name = company_names.get(ticker, "")
            
            candidates = disk_pdfs_by_pair.get(pair, [])
            selected_pdf: Optional[Path] = None
            if candidates:
                # Nếu có nhiều hơn 1 file, chấm điểm và chọn file tốt nhất
                candidates.sort(key=score_pdf_candidate, reverse=True)
                selected_pdf = candidates[0]

            if selected_pdf and selected_pdf.exists():
                size = selected_pdf.stat().st_size
                # Trang có thể đọc lazy sau
                item = InventoryItem(
                    stt=stt,
                    ticker=ticker,
                    year=year,
                    company_name=c_name,
                    file_path=str(selected_pdf),
                    file_name=selected_pdf.name,
                    file_size_bytes=size,
                    page_count=0,
                    status_manifest="Có file PDF trên đĩa",
                    is_disk_available=True,
                    notes="" if len(candidates) <= 1 else f"Chọn từ {len(candidates)} file ứng viên"
                )
            else:
                item = InventoryItem(
                    stt=stt,
                    ticker=ticker,
                    year=year,
                    company_name=c_name,
                    file_path=None,
                    file_name=f"{ticker}_{year}.pdf",
                    file_size_bytes=0,
                    page_count=0,
                    status_manifest="Thiếu file PDF trên đĩa",
                    is_disk_available=False,
                    notes="Sàn/Nguồn không phát hành hoặc chưa crawl được"
                )
            inventory.append(item)
            stt += 1

    return inventory
