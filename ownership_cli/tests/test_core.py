from ownership_cli.ownership_101x6 import (
    choose_state_total,
    extract_ocr_candidates,
    extract_page_candidates,
    normalize,
    ocr_missing_pages,
    validated_ratio,
)
import fitz
import pandas as pd


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


def test_keeps_every_row_and_only_selects_state_total_on_target_date():
    page = {
        "pdf_file": "ABC_2024.pdf",
        "page_number": 5,
        "text": "Cơ cấu cổ đông chốt ngày 31/12/2024\n"
                "Nhóm cổ đông | Số lượng cổ phần | Tỷ lệ (%)\n"
                "Nhà nước | 6.562.361 | 50,08\n"
                "Công ty A | 4.000.000 | 40,00\n"
                "Cổ đông nước ngoài | 1.000.000 | 9,92\n"
                "Cổ đông tổ chức | – | –",
    }
    candidates = extract_page_candidates("ABC", 2024, [page])
    assert len(candidates) == 4
    assert [candidate["percentage"] for candidate in candidates] == [50.08, 40.0, 9.92, None]
    assert candidates[-1]["raw_value"] == "–"
    selected, _ = choose_state_total(candidates, 2024)
    assert selected["holder"] == "Nhà nước"
    assert selected["percentage"] == 50.08


def test_preserves_other_date_and_zero_is_valid():
    page = {"pdf_file": "ABC.pdf", "page_number": 1,
            "text": "Cơ cấu cổ đông tại ngày 20/12/2024\nNhóm | Số cổ phần | Tỷ lệ (%)\nNhà nước | 0 | 0"}
    candidates = extract_page_candidates("ABC", 2024, [page])
    assert candidates[0]["percentage"] == 0
    assert choose_state_total(candidates, 2024)[0] is None
    page["text"] = page["text"].replace("20/12/2024", "31/12/2024")
    assert choose_state_total(extract_page_candidates("ABC", 2024, [page]), 2024)[0]["percentage"] == 0


def test_conflicting_sources_leave_main_ratio_empty():
    pages = [
        {"pdf_file": f"ABC_{index}.pdf", "page_number": 1,
         "text": f"Cơ cấu cổ đông tại ngày 31/12/2024\nNhóm | Số cổ phần | Tỷ lệ (%)\nNhà nước | 100 | {ratio}"}
        for index, ratio in enumerate(("50,08", "51,00"))
    ]
    selected, status = choose_state_total(extract_page_candidates("ABC", 2024, pages), 2024)
    assert selected is None
    assert "Mâu thuẫn" in status


def test_previous_page_date_and_multiple_keywords_are_retained():
    pages = [
        {"pdf_file": "ABC.pdf", "page_number": 1,
         "text": "Cơ cấu cổ đông, vốn cổ phần\nDanh sách chốt ngày 31/12/2024"},
        {"pdf_file": "ABC.pdf", "page_number": 2,
         "text": "Nhóm cổ đông | Số cổ phần | Tỷ lệ (%)\nNhà nước | 100 | 50,08"},
    ]
    candidates = extract_page_candidates("ABC", 2024, pages)
    assert len(candidates) == 2
    assert "cơ cấu cổ đông" in candidates[0]["keywords"]
    assert "vốn cổ phần" in candidates[0]["keywords"]
    assert candidates[1]["date_note"] == "Ngày trên trang trước"
    assert choose_state_total(candidates, 2024)[0]["percentage"] == 50.08


def test_shareholder_count_column_does_not_verify_state_ownership():
    page = {"pdf_file": "ABC.pdf", "page_number": 1,
            "text": "Cơ cấu cổ đông tại ngày 31/12/2024\nNhóm | Số cổ phần | Tỷ lệ số lượng cổ đông (%)\nNhà nước | 100 | 50,08"}
    candidates = extract_page_candidates("ABC", 2024, [page])
    assert candidates[0]["percentage"] == 50.08
    assert choose_state_total(candidates, 2024)[0] is None


def test_multiple_percentages_are_kept_but_not_selected_without_column():
    page = {"pdf_file": "ABC.pdf", "page_number": 1,
            "text": "Cơ cấu cổ đông tại ngày 31/12/2024\n"
                    "Nhóm | Số cổ phần | Tỷ lệ sở hữu | Tỷ lệ biểu quyết\n"
                    "Nhà nước | 100 | 50% | 60%"}
    candidates = extract_page_candidates("ABC", 2024, [page])
    assert [candidate["percentage"] for candidate in candidates] == [50.0, 60.0]
    assert choose_state_total(candidates, 2024)[0] is None


def test_ocr_continues_when_a_fallback_text_page_exists(tmp_path, monkeypatch):
    pdf_root = tmp_path / "pdf"
    ticker_dir = pdf_root / "ABC"
    ticker_dir.mkdir(parents=True)
    with fitz.open() as document:
        page = document.new_page()
        page.insert_text((50, 50), "Shareholder structure " * 10)
        document.new_page()
        document.save(ticker_dir / "ABC_2024.pdf")
    monkeypatch.setattr("pytesseract.image_to_string", lambda *args, **kwargs: "Shareholder structure")
    manifest = pd.DataFrame([{"Mã CK": "ABC", "Năm": 2024}])
    result = ocr_missing_pages(manifest, tmp_path / "text", tmp_path / "ocr", pdf_root,
                               tmp_path / "tessdata", 0, 72)
    assert result["Trang OCR mới"] == 1
    assert ocr_missing_pages(manifest, tmp_path / "text", tmp_path / "ocr", pdf_root,
                             tmp_path / "tessdata", 0, 72)["Trang OCR mới"] == 0
