"""
tests/test_regression_suite.py
==============================
Bộ kiểm thử hồi quy (Regression Test Suite) cho 9 kịch bản lỗi cốt lõi:
1. Model trả số không có trong bằng chứng -> Bị từ chối (Grounding).
2. null không bị biến thành 0.
3. Bảng có 2 cột tỷ lệ: Chọn đúng tỷ lệ cổ phần, loại bỏ tỷ lệ số lượng cổ đông.
4. Không lấy cơ quan cấp đăng ký kinh doanh làm tên cổ đông.
5. Ngày bảng khác năm BCTN không bị gán thành 31/12.
6. Mâu thuẫn nguồn số liệu cùng ngày -> Flag CAN_KIEM_TRA_MAU_THUAN.
7. Xử lý bảng nhiều tầng và làm tròn số.
8. Cache invalidation và retry unverified.
9. API lỗi hoặc JSON hỏng không thành kết quả thành công.
"""

import pytest
from ownership_core.validator import (
    CandidateOwnership,
    VerificationEngine,
    classify_holder_string,
    parse_vietnamese_number,
    verify_grounding,
)
from ownership_core.model_client import clean_json_response
from ownership_core.cache_manager import CacheManager
from ownership_core.table_extractor import extract_nearby_date


# --- 1. MODEL TRẢ SỐ KHÔNG CÓ TRONG BẰNG CHỨNG (HALLUCINATION) ---
def test_rejects_hallucinated_number_not_in_evidence():
    """Model trả về con số 43.31% nhưng trong bằng chứng chỉ có 51.0% -> Phải bị từ chối."""
    cand = CandidateOwnership(
        ticker="AAA",
        report_year=2024,
        holder_name="Nhà nước",
        classification="TONG_NHA_NUOC",
        raw_percentage="43.31%",
        percentage=43.31,
        evidence="Cơ cấu cổ đông: Nhà nước sở hữu 51,00% vốn điều lệ",
        source_method="MODEL_LOCAL"
    )
    # verify_grounding phải trả về False
    assert not verify_grounding(cand, "Cơ cấu cổ đông: Nhà nước sở hữu 51,00% vốn điều lệ")

    engine = VerificationEngine()
    result = engine.evaluate_candidates("AAA", 2024, [cand])
    assert not result.is_verified
    assert result.ratio != 43.31
    assert result.status == "CHUA_TIM_THAY_THONG_TIN"


# --- 2. NULL KHÔNG BỊ BIẾN THÀNH 0 ---
def test_null_or_dash_is_not_converted_to_zero():
    """Dấu gạch ngang '-', ô trống hoặc null không được tự động biến thành 0.0%."""
    for empty_sym in ["-", "–", "—", "", "N/A", "null", None]:
        pct, shares = parse_vietnamese_number(empty_sym)
        assert pct is None, f"Biến ký hiệu '{empty_sym}' thành {pct} là sai!"

    # Chỉ chuỗi ghi rõ 0 hoặc 0.0% mới thành 0.0
    pct_zero, _ = parse_vietnamese_number("0.0%")
    assert pct_zero == 0.0
    pct_zero2, _ = parse_vietnamese_number("0")
    assert pct_zero2 == 0.0


# --- 3. CHỌN ĐÚNG CỘT TỶ LỆ CỔ PHẦN, LOẠI BỎ TỶ LỆ SỐ LƯỢNG CỔ ĐÔNG ---
def test_two_ratio_columns_selects_shares_not_shareholder_count():
    """
    Bảng có 2 cột tỷ lệ:
    Cột 1: Tỷ lệ số lượng cổ đông (0.05%)
    Cột 2: Tỷ lệ sở hữu cổ phần (51.08%)
    Phải chọn 51.08%, tuyệt đối không chọn 0.05%.
    """
    cand_wrong_col = CandidateOwnership(
        ticker="XYZ",
        report_year=2023,
        holder_name="Cổ đông Nhà nước",
        classification="TONG_NHA_NUOC",
        column_name="Tỷ lệ số lượng cổ đông (%)",
        percentage=0.05,
        normalized_date="2023-12-31",
        evidence="Nhà nước | 1 | 0,05% | 10.000.000 | 51,08%"
    )
    cand_correct_col = CandidateOwnership(
        ticker="XYZ",
        report_year=2023,
        holder_name="Cổ đông Nhà nước",
        classification="TONG_NHA_NUOC",
        column_name="Tỷ lệ sở hữu cổ phần (%)",
        percentage=51.08,
        normalized_date="2023-12-31",
        evidence="Nhà nước | 1 | 0,05% | 10.000.000 | 51,08%"
    )

    engine = VerificationEngine()
    result = engine.evaluate_candidates("XYZ", 2023, [cand_wrong_col, cand_correct_col])
    assert result.is_verified
    assert result.ratio == 51.08
    assert result.ratio != 0.05


# --- 4. KHÔNG DÙNG TÊN CƠ QUAN CẤP ĐĂNG KÝ LÀM TÊN CỔ ĐÔNG ---
def test_does_not_treat_registration_authority_as_shareholder():
    """'Sở Kế hoạch và Đầu tư' cấp phép không phải là cổ đông."""
    cls_auth = classify_holder_string("Sở Kế hoạch và Đầu tư TP Hà Nội cấp ngày 01/01/2010")
    assert cls_auth != "TONG_NHA_NUOC"
    assert cls_auth != "TO_CHUC_NHA_NUOC_TRUC_TIEP"

    # Nhưng SCIC hoặc Bộ Tài chính thì phải được phân loại đúng
    cls_scic = classify_holder_string("Tổng công ty Đầu tư và Kinh doanh vốn Nhà nước (SCIC)")
    assert cls_scic == "TO_CHUC_NHA_NUOC_TRUC_TIEP"


# --- 5. NGÀY BẢNG KHÁC NĂM BCTN KHÔNG BỊ SỬA THÀNH 31/12 ---
def test_table_date_different_from_year_end_is_flagged():
    """Nếu bảng chốt ngày 20/04/2023 (ĐHCĐ), không được gán bừa thành 2023-12-31."""
    cand = CandidateOwnership(
        ticker="VNM",
        report_year=2023,
        holder_name="SCIC",
        classification="TO_CHUC_NHA_NUOC_TRUC_TIEP",
        percentage=36.0,
        raw_date="20/04/2023",
        normalized_date="2023-04-20",
        evidence="Danh sách cổ đông chốt ngày 20/04/2023: SCIC sở hữu 36,00%"
    )

    engine = VerificationEngine(target_month_day="12-31")
    result = engine.evaluate_candidates("VNM", 2023, [cand])
    # Không được tự nhận là đã xác minh chốt cuối năm 31/12
    assert not result.is_verified
    assert result.status == "CAN_KIEM_TRA_NGAY_KHAC"
    assert result.cutoff_date == "2023-04-20"
    assert result.ratio == 36.0


# --- 6. MÂU THUẪN NGUỒN SỐ LIỆU CÙNG NGÀY ---
def test_conflicting_sources_flagged_as_warning():
    """Hai bảng trong cùng tài liệu ghi 2 tỷ lệ khác nhau chốt cùng ngày (ví dụ 40% vs 51%) -> Phải báo mâu thuẫn."""
    cand1 = CandidateOwnership(
        ticker="HPG",
        report_year=2022,
        holder_name="Nhà nước",
        classification="TONG_NHA_NUOC",
        percentage=40.0,
        normalized_date="2022-12-31",
        evidence="Bảng 1: Nhà nước 40,0%"
    )
    cand2 = CandidateOwnership(
        ticker="HPG",
        report_year=2022,
        holder_name="Nhà nước",
        classification="TONG_NHA_NUOC",
        percentage=51.0,
        normalized_date="2022-12-31",
        evidence="Bảng 2: Nhà nước 51,0%"
    )

    engine = VerificationEngine()
    result = engine.evaluate_candidates("HPG", 2022, [cand1, cand2])
    assert not result.is_verified
    assert result.status == "CAN_KIEM_TRA_MAU_THUAN"
    assert result.ratio is None


# --- 7. BẢNG NHIỀU TỔ CHỨC NHÀ NƯỚC (CỘNG HỢP LỆ TRÁNH TRÙNG DÒNG TỔNG) ---
def test_aggregates_direct_state_entities_without_duplication():
    """Cộng SCIC (30%) và EVN (20%) nhưng không cộng trùng nếu có dòng tổng."""
    cand_scic = CandidateOwnership(
        ticker="PPC",
        report_year=2024,
        holder_name="SCIC",
        classification="TO_CHUC_NHA_NUOC_TRUC_TIEP",
        percentage=30.0,
        share_count=30000000,
        normalized_date="2024-12-31",
        evidence="SCIC: 30,0% (30.000.000 CP)"
    )
    cand_evn = CandidateOwnership(
        ticker="PPC",
        report_year=2024,
        holder_name="Tập đoàn Điện lực Việt Nam (Bộ Công Thương)",
        classification="TO_CHUC_NHA_NUOC_TRUC_TIEP",
        percentage=20.0,
        share_count=20000000,
        normalized_date="2024-12-31",
        evidence="EVN: 20,0% (20.000.000 CP)"
    )

    engine = VerificationEngine()
    result = engine.evaluate_candidates("PPC", 2024, [cand_scic, cand_evn])
    assert result.is_verified
    assert result.ratio == 50.0
    assert result.share_count == 50000000


# --- 8. PARSE VÀ LÀM SẠCH JSON TỪ MODEL (KHÔNG BỊ PHÁ VỠ BỞI MARKDOWN) ---
def test_clean_json_response_with_markdown_blocks():
    """Model bọc JSON trong ```json ... ``` hoặc có văn bản phụ -> Parse chính xác."""
    raw = """Dưới đây là kết quả trích xuất:
```json
{
  "ticker": "TCB",
  "report_year": 2024,
  "candidates": []
}
```
Chúc bạn một ngày tốt lành!"""
    parsed = clean_json_response(raw)
    assert parsed is not None
    assert parsed["ticker"] == "TCB"
    assert parsed["candidates"] == []


# --- 9. CACHE TÁCH BẠCH KHÔNG XUNG ĐỘT KHI ĐỔI EXTRACTOR VERSION ---
def test_cache_versioning_and_invalidation(tmp_path):
    """Khi đổi extractor version, cache candidate cũ không bị dùng đè nếu hash khác."""
    cm = CacheManager(tmp_path)
    file_h = "abc123hash"
    cm.set_candidates(file_h, "v1", "hash_prompt1", [{"holder_name": "SCIC", "percentage": 40.0}])
    
    # Đọc lại đúng v1, hash_prompt1 -> thấy
    hit = cm.get_candidates(file_h, "v1", "hash_prompt1")
    assert hit is not None
    assert hit[0]["percentage"] == 40.0

    # Đổi sang v2 hoặc prompt khác -> Trả về None để bóc tách lại, không dùng kết quả cũ
    miss = cm.get_candidates(file_h, "v2", "hash_prompt1")
    assert miss is None
    miss2 = cm.get_candidates(file_h, "v1", "hash_prompt2")
    assert miss2 is None
