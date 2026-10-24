from ownership_cli.ownership_101x6 import extract_ocr_candidates, normalize, validated_ratio


def test_normalize_handles_diacritics():
    assert normalize("Cổ đông Nhà nước") == "co dong nha nuoc"


def test_rejects_conflicting_table_date():
    row = {
        "Năm báo cáo": 2020,
        "Tỷ lệ sở hữu Nhà nước (%)": None,
        "Đoạn/bảng bằng chứng": "Tỷ lệ cổ đông Nhà nước 5,3%",
    }
    review = {
        "parsed": {
            "decision": "CORRECT",
            "percentage": 5.3,
            "data_date": "2020-12-31",
            "reason": "Bảng ngày 26/02/2021 ghi tỷ lệ 5,3%",
        }
    }
    assert validated_ratio(review, row)[0] is None


def test_ocr_candidate_requires_state_line_and_late_december_date():
    page = {
        "pdf_file": "ABC_2024.pdf",
        "page_number": 5,
        "text": "Cơ cấu cổ đông tại ngày 20/12/2024\nCổ đông Nhà nước 36,00%",
    }
    candidates = extract_ocr_candidates("ABC", 2024, [page])
    assert len(candidates) == 1
    assert candidates[0]["percentage"] == 36.0
    assert candidates[0]["data_date"] == "2024-12-20"
