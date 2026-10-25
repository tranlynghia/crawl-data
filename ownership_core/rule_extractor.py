"""
ownership_core.rule_extractor
=============================
Bộ trích xuất thuần code (Code-First Extractor):
- Phân tích trực tiếp các bảng có cấu trúc rõ ràng.
- Bóc tách ứng viên theo cột tỷ lệ và số lượng cổ phiếu đã được định danh.
- Chạy bộ kiểm chứng ngay lập tức. Nếu đạt chuẩn xác minh -> XONG (0 token API, tốc độ tức thì).
"""

from __future__ import annotations
from typing import List, Optional, Tuple
from ownership_core.table_extractor import StructuredTable
from ownership_core.validator import (
    CandidateOwnership,
    classify_holder_string,
    parse_vietnamese_number,
)


def extract_candidates_from_tables(
    ticker: str,
    report_year: int,
    tables: List[StructuredTable]
) -> List[CandidateOwnership]:
    """Trích xuất danh sách ứng viên từ các bảng có cấu trúc bằng quy tắc code thuần túy."""
    candidates: List[CandidateOwnership] = []

    for tab in tables:
        # Nếu bảng không có cột tỷ lệ sở hữu hợp lệ, bỏ qua
        if not tab.share_ratio_col_indices:
            continue

        ratio_col_idx = tab.share_ratio_col_indices[0]
        count_col_idx = tab.share_count_col_indices[0] if tab.share_count_col_indices else None

        header_str = " | ".join(str(c or "").strip() for c in tab.header_rows[0]) if tab.header_rows else ""

        for row in tab.data_rows:
            if not row or len(row) <= ratio_col_idx:
                continue

            # Xác định nhãn cổ đông (thường ở cột 0, hoặc cột 1 nếu cột 0 là STT)
            label = str(row[0] or "").strip()
            if label.isdigit() and len(row) > 1:
                label = str(row[1] or "").strip()

            if not label or len(label) < 2:
                continue

            raw_ratio_str = str(row[ratio_col_idx] or "").strip()
            pct_val, _ = parse_vietnamese_number(raw_ratio_str)

            share_cnt = None
            if count_col_idx is not None and len(row) > count_col_idx:
                raw_count_str = str(row[count_col_idx] or "").strip()
                _, share_cnt = parse_vietnamese_number(raw_count_str)

            classification = classify_holder_string(label)
            row_str = " | ".join(str(c or "").strip() for c in row)
            ev = f"Trang {tab.page_number}:\n{header_str}\n{row_str}"

            cand = CandidateOwnership(
                ticker=ticker,
                report_year=report_year,
                holder_name=label,
                classification=classification,
                raw_percentage=raw_ratio_str,
                percentage=pct_val,
                share_count=share_cnt,
                raw_date=tab.table_date,
                normalized_date=tab.table_date,
                page_number=tab.page_number,
                column_name=str(tab.header_rows[0][ratio_col_idx] or "Tỷ lệ") if tab.header_rows and ratio_col_idx < len(tab.header_rows[0]) else "Tỷ lệ",
                row_label=label,
                evidence=ev,
                source_method="RULE_TABLE",
                grounded=True
            )
            candidates.append(cand)

    return candidates
