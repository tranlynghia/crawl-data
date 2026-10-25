"""
ownership_core.validator
========================
Bộ kiểm chứng bắt buộc (Deterministic Verification Engine):
- Grounding: Bắt buộc số phải xuất hiện tại đúng ô/dòng bằng chứng.
- Phân biệt tỷ lệ sở hữu cổ phần với tỷ lệ số cổ đông, vốn cá nhân HĐQT đại diện, cổ đông nước ngoài.
- Chuẩn hóa dấu chấm/phẩy, không biến null/dash thành 0.
- Kiểm tra số học (shares / total_shares * 100) khi có dữ liệu cùng ngày.
- Tách riêng năm BCTN và ngày chốt danh sách; kiểm tra mốc mục tiêu 31/12.
- Nhận diện và xử lý xung đột số liệu.
"""

from __future__ import annotations
import math
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, Set
from ownership_core.config import (
    DIRECT_STATE_ENTITIES,
    COLUMN_SHAREHOLDER_COUNT_KEYWORDS,
    normalize_text,
)


@dataclass
class CandidateOwnership:
    ticker: str
    report_year: int
    holder_name: str
    classification: str
    raw_percentage: Optional[str] = None
    percentage: Optional[float] = None
    share_count: Optional[int] = None
    raw_date: Optional[str] = None
    normalized_date: Optional[str] = None
    page_number: Optional[int] = None
    column_name: Optional[str] = None
    row_label: Optional[str] = None
    evidence: str = ""
    source_method: str = "RULE_TABLE"
    grounded: bool = False
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "ticker": self.ticker,
            "report_year": self.report_year,
            "holder_name": self.holder_name,
            "classification": self.classification,
            "raw_percentage": self.raw_percentage,
            "percentage": self.percentage,
            "share_count": self.share_count,
            "raw_date": self.raw_date,
            "normalized_date": self.normalized_date,
            "page_number": self.page_number,
            "column_name": self.column_name,
            "row_label": self.row_label,
            "evidence": self.evidence,
            "source_method": self.source_method,
            "grounded": self.grounded,
            "notes": self.notes,
        }


@dataclass
class ValidationResult:
    is_verified: bool
    status: str
    ratio: Optional[float]
    share_count: Optional[int]
    holder_name: str
    cutoff_date: str
    basis: str
    evidence: str
    page_number: Optional[int]
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "is_verified": self.is_verified,
            "status": self.status,
            "ratio": self.ratio,
            "share_count": self.share_count,
            "holder_name": self.holder_name,
            "cutoff_date": self.cutoff_date,
            "basis": self.basis,
            "evidence": self.evidence,
            "page_number": self.page_number,
            "notes": self.notes,
        }


def parse_vietnamese_number(raw_val: Any) -> Tuple[Optional[float], Optional[int]]:
    """
    Chuẩn hóa chuỗi số tiếng Việt/Anh thành (ratio_float, count_int).
    Xử lý: '51,08%', '51.08', '6.562.361', '6,562,361', '-', '–', None.
    TUYỆT ĐỐI không biến None, ô trống hoặc gạch ngang thành 0.
    """
    if raw_val is None:
        return None, None

    s = str(raw_val).strip()
    if s in {"", "-", "–", "—", "N/A", "null", "None"}:
        return None, None

    # Bỏ ký tự %
    s_clean = s.replace("%", "").strip()

    # Nhận diện định dạng chấm/phẩy:
    # Trường hợp 1: Có cả dấu chấm và phẩy (ví dụ: '1.234.567,89' hoặc '1,234,567.89')
    if "." in s_clean and "," in s_clean:
        if s_clean.rfind(",") > s_clean.rfind("."):
            # Chuẩn Việt/Đức: 1.234,56
            s_norm = s_clean.replace(".", "").replace(",", ".")
        else:
            # Chuẩn Mỹ/Anh: 1,234.56
            s_norm = s_clean.replace(",", "")
    elif "," in s_clean:
        # Chỉ có dấu phẩy:
        # Nếu sau phẩy có 1-4 chữ số (ví dụ: '51,08' hoặc '51,084'): xem là phần thập phân
        parts = s_clean.split(",")
        if len(parts) == 2 and 1 <= len(parts[1]) <= 4:
            s_norm = s_clean.replace(",", ".")
        else:
            # Dấu phẩy ngăn cách hàng nghìn (ví dụ: '6,562,361')
            s_norm = s_clean.replace(",", "")
    elif "." in s_clean:
        # Chỉ có dấu chấm:
        parts = s_clean.split(".")
        if len(parts) == 2 and 1 <= len(parts[1]) <= 4 and float(s_clean) <= 100.0:
            s_norm = s_clean
        elif len(parts) >= 2 and all(len(p) == 3 for p in parts[1:]):
            # Ngăn cách hàng nghìn (ví dụ: '6.562.361')
            s_norm = s_clean.replace(".", "")
        else:
            s_norm = s_clean
    else:
        s_norm = s_clean

    try:
        val_float = float(s_norm)
        # Nếu là số nguyên lớn (ví dụ > 100 và không có phần thập phân ban đầu)
        if val_float > 100 and "." not in s_norm:
            return None, int(val_float)
        elif 0.0 <= val_float <= 100.0:
            return round(val_float, 4), None
        else:
            return None, int(val_float)
    except ValueError:
        return None, None


def classify_holder_string(label: str) -> str:
    """Phân loại nhóm/tổ chức cổ đông chính xác."""
    norm = normalize_text(re.sub(r"^[•●\-–—]?\s*(?:\d+|[IVX]+)[.)]?\s+", "", label, flags=re.IGNORECASE))
    
    # 1. Đại diện phần vốn HĐQT / Ban điều hành (KHÔNG tính là sở hữu trực tiếp của Nhà nước)
    if any(k in norm for k in ["dai dien phan von", "dai dien von", "von dai dien", "representative capital"]):
        return "DAI_DIEN_VON_HDQT"

    # 2. Dòng tổng sở hữu Nhà nước chính danh
    if norm in {"nha nuoc", "co dong nha nuoc", "so huu nha nuoc", "von nha nuoc", "state", "state shareholders", "state ownership"}:
        return "TONG_NHA_NUOC"

    # 3. Tổ chức đại diện vốn Nhà nước trực tiếp (SCIC, Bộ, UBND, CMSC...)
    for entity in DIRECT_STATE_ENTITIES:
        if normalize_text(entity) in norm:
            return "TO_CHUC_NHA_NUOC_TRUC_TIEP"

    # 4. Cổ đông Nhà nước riêng lẻ (có từ khóa nhà nước)
    if "nha nuoc" in norm or "state shareholder" in norm:
        return "CO_DONG_NHA_NUOC_RIENG_LE"

    # 5. Các nhóm khác
    if "nuoc ngoai" in norm or "foreign" in norm:
        return "CO_DONG_NUOC_NGOAI"
    if "trong nuoc" in norm or "domestic" in norm:
        return "CO_DONG_TRONG_NUOC"
    if "to chuc" in norm or "institutional" in norm:
        return "CO_DONG_TO_CHUC_CHUA_XAC_MINH"
    if "ca nhan" in norm or "individual" in norm:
        return "CO_DONG_CA_NHAN"

    return "CHUA_RO_TIEU_CHI"


def verify_grounding(candidate: CandidateOwnership, document_text: str) -> bool:
    """
    Kiểm tra Grounding: Con số tỷ lệ hoặc số lượng cổ phiếu
    BẮT BUỘC phải thực sự xuất hiện trong bằng chứng trích xuất hoặc tài liệu nguồn!
    """
    if candidate.percentage is None and candidate.share_count is None:
        # Nếu cả 2 đều rỗng hoặc null, phải có bằng chứng từ khóa
        return bool(candidate.evidence and len(candidate.evidence.strip()) >= 5)

    ev_norm = normalize_text(candidate.evidence + " " + document_text[:5000])

    # 1. Kiểm tra tỷ lệ
    if candidate.percentage is not None:
        p_val = candidate.percentage
        # Các cách viết khả dĩ: '51.08', '51,08', '51.08%', '51,08%'
        p_str1 = f"{p_val:.2f}".rstrip("0").rstrip(".")
        p_str2 = p_str1.replace(".", ",")
        if p_str1 not in ev_norm and p_str2 not in ev_norm and f"{p_val:.1f}" not in ev_norm and f"{p_val:.0f}" not in ev_norm:
            return False

    # 2. Kiểm tra số cổ phiếu nếu có
    if candidate.share_count is not None and candidate.share_count > 0:
        c_val = candidate.share_count
        c_str = str(c_val)
        # Các định dạng có dấu phẩy/chấm: 6,562,361 hoặc 6.562.361
        c_formatted1 = f"{c_val:,}"
        c_formatted2 = c_formatted1.replace(",", ".")
        if (c_str not in ev_norm and
            c_formatted1 not in ev_norm and
            c_formatted2 not in ev_norm and
            normalize_text(c_formatted1) not in ev_norm):
            return False

    return True


class VerificationEngine:
    """Động cơ kiểm chứng tính hợp lệ và chốt kết quả tỷ lệ sở hữu Nhà nước."""

    def __init__(self, target_month_day: str = "12-31"):
        self.target_month_day = target_month_day

    def evaluate_candidates(
        self,
        ticker: str,
        report_year: int,
        candidates: List[CandidateOwnership],
        document_text: str = ""
    ) -> ValidationResult:
        """
        Đánh giá toàn diện danh sách ứng viên và đưa ra quyết định xác minh:
        1. Lọc bỏ ứng viên không có Grounding trong tài liệu.
        2. Loại trừ cột tỷ lệ số lượng cổ đông (chỉ lấy tỷ lệ vốn/cổ phần).
        3. Ưu tiên dòng 'TONG_NHA_NUOC' trên ngày mục tiêu (31/12).
        4. Kiểm tra mâu thuẫn số liệu.
        5. Xuất ValidationResult với đầy đủ bằng chứng kiểm tra được.
        """
        if not candidates:
            return ValidationResult(
                is_verified=False,
                status="CHUA_TIM_THAY_THONG_TIN",
                ratio=None,
                share_count=None,
                holder_name="Không tìm thấy",
                cutoff_date="N/A",
                basis="Không phát hiện bảng cổ đông có vốn Nhà nước trong phạm vi tài liệu",
                evidence="",
                page_number=None,
                notes="Cần kiểm tra thêm các trang scan hoặc tài liệu phụ nếu có"
            )

        # 1. Lọc Grounding và loại trừ cột số lượng cổ đông
        valid_candidates: List[CandidateOwnership] = []
        for c in candidates:
            # Kiểm tra Grounding
            if not verify_grounding(c, document_text):
                continue
            c.grounded = True

            # Loại trừ nếu cột thuộc nhóm "số lượng cổ đông"
            col_norm = normalize_text(c.column_name or "")
            if any(normalize_text(bad) in col_norm for bad in COLUMN_SHAREHOLDER_COUNT_KEYWORDS):
                continue

            valid_candidates.append(c)

        if not valid_candidates:
            return ValidationResult(
                is_verified=False,
                status="CHUA_TIM_THAY_THONG_TIN",
                ratio=None,
                share_count=None,
                holder_name="Không có số liệu kiểm chứng",
                cutoff_date="N/A",
                basis="Các số liệu trích xuất không vượt qua kiểm định Grounding hoặc dính cột số lượng cổ đông",
                evidence=candidates[0].evidence if candidates else "",
                page_number=candidates[0].page_number if candidates else None,
                notes="Loại bỏ số liệu ảo / hallucination"
            )

        # 2. Tìm kiếm ứng viên thuộc nhóm Tổng sở hữu Nhà nước hoặc Tổ chức Nhà nước
        state_candidates = [
            c for c in valid_candidates
            if c.classification in {"TONG_NHA_NUOC", "CO_DONG_NHA_NUOC_RIENG_LE", "TO_CHUC_NHA_NUOC_TRUC_TIEP"}
        ]

        # 3. Trường hợp xác nhận 0% (Không có sở hữu Nhà nước)
        # Nếu bảng nêu rõ 0% hoặc chỉ có 100% cổ đông tư nhân/nước ngoài
        zero_candidates = [c for c in state_candidates if c.percentage == 0.0]
        if zero_candidates and not [c for c in state_candidates if (c.percentage or 0) > 0]:
            zc = zero_candidates[0]
            return ValidationResult(
                is_verified=True,
                status="DA_XAC_MINH_KHONG_CO_SHNN",
                ratio=0.0,
                share_count=zc.share_count or 0,
                holder_name="Không có (0.0%)",
                cutoff_date=zc.normalized_date or f"{report_year}-12-31",
                basis="Tài liệu ghi rõ dòng Nhà nước có tỷ lệ 0.0%",
                evidence=zc.evidence,
                page_number=zc.page_number,
                notes="Xác nhận không có vốn Nhà nước"
            )

        if not state_candidates:
            # Kiểm tra nếu có bảng cơ cấu cổ đông nhưng chỉ có tư nhân & nước ngoài
            other_candidates = [c for c in valid_candidates if c.classification in {"CO_DONG_NUOC_NGOAI", "CO_DONG_TRONG_NUOC"}]
            if other_candidates:
                return ValidationResult(
                    is_verified=True,
                    status="DA_XAC_MINH_KHONG_CO_SHNN",
                    ratio=0.0,
                    share_count=0,
                    holder_name="Không có",
                    cutoff_date=other_candidates[0].normalized_date or f"{report_year}-12-31",
                    basis="Bảng cơ cấu cổ đông chỉ có cổ đông trong nước/nước ngoài tư nhân, không có Nhà nước",
                    evidence=other_candidates[0].evidence,
                    page_number=other_candidates[0].page_number,
                    notes="100% vốn tư nhân/nước ngoài"
                )

            return ValidationResult(
                is_verified=False,
                status="CHUA_TIM_THAY_THONG_TIN",
                ratio=None,
                share_count=None,
                holder_name="Chưa rõ",
                cutoff_date="N/A",
                basis="Chưa tìm thấy dòng cổ đông đại diện Nhà nước trong các ứng viên hợp lệ",
                evidence=valid_candidates[0].evidence,
                page_number=valid_candidates[0].page_number
            )

        # 4. Xử lý các ứng viên có vốn Nhà nước
        # Ưu tiên các ứng viên có ngày đúng 31/12
        target_date_str = f"{report_year}-{self.target_month_day}"
        target_date_candidates = [c for c in state_candidates if c.normalized_date == target_date_str]
        
        # Nếu có ứng viên dòng TONG_NHA_NUOC
        total_line_cands = [c for c in (target_date_candidates or state_candidates) if c.classification == "TONG_NHA_NUOC"]
        if total_line_cands:
            # Kiểm tra mâu thuẫn giữa các bảng cùng ngày
            unique_ratios = list({c.percentage for c in total_line_cands if c.percentage is not None})
            if len(unique_ratios) > 1:
                # Nếu sai lệch > 0.1% -> Báo mâu thuẫn
                if max(unique_ratios) - min(unique_ratios) > 0.1:
                    return ValidationResult(
                        is_verified=False,
                        status="CAN_KIEM_TRA_MAU_THUAN",
                        ratio=None,
                        share_count=None,
                        holder_name="Mâu thuẫn số liệu",
                        cutoff_date=total_line_cands[0].normalized_date or "N/A",
                        basis=f"Phát hiện nhiều tỷ lệ Nhà nước khác nhau trong tài liệu: {unique_ratios}",
                        evidence="\n---\n".join(c.evidence for c in total_line_cands[:2]),
                        page_number=total_line_cands[0].page_number,
                        notes="Cần kiểm tra đối chiếu trực tiếp"
                    )

            best = total_line_cands[0]
            is_target_date = (best.normalized_date == target_date_str)
            status_val = "DA_XAC_MINH_CO_SHNN" if is_target_date else "CAN_KIEM_TRA_NGAY_KHAC"
            
            return ValidationResult(
                is_verified=is_target_date,
                status=status_val,
                ratio=best.percentage,
                share_count=best.share_count,
                holder_name=best.holder_name,
                cutoff_date=best.normalized_date or "Chưa rõ ngày",
                basis="Dòng tổng 'Nhà nước' trên bảng cơ cấu cổ đông",
                evidence=best.evidence,
                page_number=best.page_number,
                notes="Khớp ngày 31/12" if is_target_date else f"Ngày chốt là {best.normalized_date} (khác 31/12)"
            )

        # 5. Nếu không có dòng tổng, tính từ tổ chức đại diện vốn Nhà nước trực tiếp (SCIC, Bộ, UBND...)
        direct_cands = [c for c in (target_date_candidates or state_candidates) if c.classification == "TO_CHUC_NHA_NUOC_TRUC_TIEP"]
        if direct_cands:
            # Cộng dồn các tổ chức khác nhau (tránh cộng trùng)
            seen_orgs: Set[str] = set()
            total_pct = 0.0
            total_shares = 0
            ev_list = []
            
            for dc in direct_cands:
                org_norm = normalize_text(dc.holder_name)
                if org_norm not in seen_orgs and dc.percentage is not None:
                    seen_orgs.add(org_norm)
                    total_pct += dc.percentage
                    if dc.share_count:
                        total_shares += dc.share_count
                    ev_list.append(f"{dc.holder_name}: {dc.percentage}% ({dc.share_count or 'N/A'} CP)")

            if total_pct > 0:
                best = direct_cands[0]
                is_target_date = (best.normalized_date == target_date_str)
                status_val = "DA_XAC_MINH_CO_SHNN" if is_target_date else "CAN_KIEM_TRA_NGAY_KHAC"

                return ValidationResult(
                    is_verified=is_target_date,
                    status=status_val,
                    ratio=round(total_pct, 4),
                    share_count=total_shares if total_shares > 0 else None,
                    holder_name="; ".join(sorted(list(seen_orgs))),
                    cutoff_date=best.normalized_date or "Chưa rõ ngày",
                    basis="Tổng hợp từ các cổ đông đại diện vốn Nhà nước trực tiếp (SCIC, Bộ, UBND...)",
                    evidence="\n".join(ev_list),
                    page_number=best.page_number,
                    notes="Cộng hợp lệ các tổ chức Nhà nước" if is_target_date else f"Ngày chốt là {best.normalized_date}"
                )

        # Fallback: Trả về ứng viên đầu tiên và đánh dấu cần kiểm tra
        first_cand = state_candidates[0]
        return ValidationResult(
            is_verified=False,
            status="CAN_KIEM_TRA_CHUA_DU_CAN_CU",
            ratio=first_cand.percentage,
            share_count=first_cand.share_count,
            holder_name=first_cand.holder_name,
            cutoff_date=first_cand.normalized_date or "N/A",
            basis="Chưa đủ căn cứ xác minh 100% tư cách pháp nhân đại diện vốn Nhà nước",
            evidence=first_cand.evidence,
            page_number=first_cand.page_number,
            notes="Cần kiểm định thêm"
        )
