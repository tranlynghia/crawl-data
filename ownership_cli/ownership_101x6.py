from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
import re
import sys
import time
import unicodedata
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
import fitz
import pytesseract
import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill



VERIFIED = "Đã xác minh có sở hữu nhà nước"
AUDIT_PROMPT = """Bạn kiểm định tỷ lệ sở hữu Nhà nước trong một trang báo cáo thường niên Việt Nam.
Chỉ dùng nội dung trang được đưa vào. Đọc đúng dòng cổ đông Nhà nước và đúng cột tỷ lệ sở hữu cổ phần/vốn, không lấy tỷ lệ số lượng cổ đông, ngưỡng cổ đông lớn (>5%), tỷ lệ tổ chức hay ngày giao dịch.
Không mặc định cổ đông tổ chức/Tổng công ty là Nhà nước. Chỉ tính phần được nêu rõ là Nhà nước; không tự cộng nhiều tổ chức.
Không lấy tỷ lệ phần vốn do một thành viên HĐQT đại diện làm tổng tỷ lệ Nhà nước của doanh nghiệp.
Ngày dữ liệu phải gắn với bảng/dòng sở hữu và đúng 31/12 của năm báo cáo. Không dùng ngày ở footer hoặc năm trên bìa.
Nếu con số đầu vào sai nhưng bảng ghi rõ con số đúng, trả CORRECT. Nếu đầu vào đúng, trả VALID. Nếu không thể xác định chắc chắn, trả UNCLEAR.
Chỉ trả JSON object: decision (VALID|CORRECT|UNCLEAR), percentage (số hoặc null), data_date (YYYY-MM-DD hoặc null), support (trích đoạn rất ngắn từ dòng chứng minh), reason (ngắn)."""
FALLBACK_TERMS = [
    ("cổ đông nhà nước", 140),
    ("sở hữu nhà nước", 140),
    ("vốn nhà nước", 140),
    ("nhà nước", 105),
    ("state shareholder", 140),
    ("state shareholders", 140),
    ("cơ cấu cổ đông", 125),
    ("shareholder structure", 125),
    ("shareholding structure", 125),
    ("cơ cấu sở hữu", 120),
    ("ownership structure", 120),
    ("cơ cấu vốn cổ phần", 115),
    ("vốn cổ phần", 95),
    ("số lượng cổ đông", 110),
    ("number of shareholders", 110),
    ("danh sách cổ đông lớn", 105),
    ("thông tin cổ đông lớn", 105),
    ("major shareholders", 105),
    ("chi tiết vốn góp của chủ sở hữu", 100),
    ("owners capital contribution", 100),
    ("shareholders", 35),
    ("cổ đông", 35),
]


def normalize(value: str) -> str:
    value = value.casefold().replace("đ", "d")
    value = unicodedata.normalize("NFD", value)
    value = "".join(char for char in value if unicodedata.category(char) != "Mn")
    return re.sub(r"[^a-z0-9%]+", " ", value).strip()


NORMALIZED_TERMS = [(term, normalize(term), weight) for term, weight in FALLBACK_TERMS]
DATE_PATTERN = re.compile(r"(?<!\d)(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(20\d{2})(?!\d)")
RATIO_CELL = re.compile(r"^\s*(\d{1,3}(?:[.,]\d{1,4})?)\s*%?\s*$")
EMPTY_CELL = {"", "-", "–", "—"}
RATIO_HEADERS = ("ty le", "percentage", "phan tram")
DATE_CONTEXT = ("tai ngay", "chot ngay", "den ngay", "as of", "as at")


def matched_terms(text: str) -> list[str]:
    normalized = normalize(text)
    return [term for term, normalized_term, _ in NORMALIZED_TERMS if normalized_term in normalized]


def split_cells(line: str) -> list[str]:
    return [cell.strip() for cell in re.split(r"\s*\|\s*|\t+|\s{2,}", line.strip()) if cell.strip()]


def parse_percentage(cell: str) -> float | None:
    match = RATIO_CELL.fullmatch(cell)
    if not match:
        return None
    value = float(match.group(1).replace(",", "."))
    return value if 0 <= value <= 100 else None


def classify_holder(label: str) -> str:
    normalized = normalize(re.sub(r"^[•●\-–—]?\s*(?:\d+|[IVX]+)[.)]?\s+", "", label, flags=re.IGNORECASE))
    if any(term in normalized for term in ("dai dien von", "dai dien phan von", "von dai dien", "representative capital")):
        return "Vốn đại diện HĐQT"
    if normalized in {"nha nuoc", "co dong nha nuoc", "so huu nha nuoc", "von nha nuoc", "state", "state shareholders", "state ownership"}:
        return "Tổng sở hữu Nhà nước"
    if "nha nuoc" in normalized or "state shareholder" in normalized:
        return "Cổ đông Nhà nước riêng lẻ"
    if "nuoc ngoai" in normalized or "foreign" in normalized:
        return "Cổ đông nước ngoài"
    if "trong nuoc" in normalized or "domestic" in normalized:
        return "Cổ đông trong nước"
    if "to chuc" in normalized or "institutional" in normalized:
        return "Cổ đông tổ chức"
    return "Cổ đông/chỉ tiêu chưa xác minh tư cách Nhà nước"


def table_date(lines: list[str], header_index: int, previous_page: str = "") -> tuple[str | None, str]:
    nearby = lines[max(0, header_index - 18):header_index + 1]
    dated = []
    for line in nearby:
        if not any(marker in normalize(line) for marker in DATE_CONTEXT):
            continue
        for match in DATE_PATTERN.finditer(line):
            try:
                dated.append(date(int(match.group(3)), int(match.group(2)), int(match.group(1))).isoformat())
            except ValueError:
                continue
    if not dated and header_index < 12 and previous_page:
        return table_date(previous_page.splitlines()[-12:], 12)
    unique = set(dated)
    if len(unique) == 1:
        return dated[-1], "Ngày gắn với bảng" if not previous_page else "Ngày trên trang trước"
    return None, "Nhiều ngày gần bảng" if unique else "Chưa thấy ngày gắn với bảng"


def extract_page_candidates(ticker: str, year: int, pages: list[dict[str, Any]], source: str | None = None) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    previous_by_file: dict[str, str] = {}
    for page in sorted(pages, key=lambda item: (str(item.get("pdf_file")), int(item.get("page_number") or 0))):
        text = str(page.get("text") or "")
        file_name = str(page.get("pdf_file") or "")
        lines = text.splitlines()
        keywords = matched_terms(text)
        if not keywords:
            previous_by_file[file_name] = text
            continue
        score, _ = page_score(text, year)
        page_data = {
            "ticker": ticker, "year": year, "pdf_file": file_name,
            "page_number": int(page.get("page_number") or 0), "source": source,
            "keywords": "; ".join(keywords), "score": score,
            "method": "OCR" if page.get("ocr") else "Lớp chữ PDF",
        }
        page_candidates = []
        headers = [(index, line, split_cells(line)) for index, line in enumerate(lines)
                   if len(split_cells(line)) >= 2 and any(
                       normalize(cell).startswith(("ty le", "percentage", "phan tram"))
                       for cell in split_cells(line))]
        for header_index, header_line, header_cells in headers:
            column_indices = [index for index, cell in enumerate(header_cells)
                              if normalize(cell).startswith(("ty le", "percentage", "phan tram"))]
            if not column_indices:
                continue
            column_index = column_indices[0]
            data_date, date_note = table_date(lines, header_index)
            if data_date is None and date_note == "Chưa thấy ngày gắn với bảng" and header_index < 12:
                previous = previous_by_file.get(file_name, "")
                if previous:
                    data_date, date_note = table_date(previous.splitlines()[-12:], 12)
                    if data_date:
                        date_note = "Ngày trên trang trước"
            next_header = next((index for index, _, _ in headers if index > header_index), len(lines))
            for row_line in lines[header_index + 1:min(next_header, header_index + 65)]:
                cells = split_cells(row_line)
                if len(cells) < 2 or not re.search(r"[A-Za-zÀ-ỹ]", cells[0]):
                    if len(cells) < 3 or not re.fullmatch(r"(?:\d+|[IVX]+)[.)]?", cells[0], flags=re.IGNORECASE):
                        continue
                if any(marker in normalize(row_line) for marker in RATIO_HEADERS):
                    continue
                label_index = 1 if re.fullmatch(r"(?:\d+|[IVX]+)[.)]?", cells[0], flags=re.IGNORECASE) else 0
                if not re.search(r"[A-Za-zÀ-ỹ]", cells[label_index]):
                    continue
                percent_cells = [cell for cell in cells[1:] if "%" in cell and parse_percentage(cell) is not None]
                if len(percent_cells) == 1:
                    ratio_cell = percent_cells[0]
                elif len(percent_cells) > 1:
                    for raw_value in percent_cells:
                        page_candidates.append({**page_data, "holder": cells[label_index], "column": None,
                                                "raw_value": raw_value, "percentage": parse_percentage(raw_value),
                                                "data_date": data_date, "date_note": "Nhiều cột tỷ lệ; cần xác định đúng cột",
                                                "classification": classify_holder(cells[label_index]),
                                                "ownership_column": False,
                                                "evidence": f"{header_line.strip()}\n{row_line.strip()}"})
                    continue
                elif len(cells) == len(header_cells) and column_index < len(cells):
                    ratio_cell = cells[column_index]
                elif column_index == len(header_cells) - 1 and len(cells) >= 3:
                    ratio_cell = cells[-1]
                else:
                    ratio_cell = ""
                if not ratio_cell:
                    continue
                header_position = header_line.find(header_cells[column_index])
                value_position = row_line.rfind(ratio_cell)
                if header_position >= 0 and value_position >= 0 and abs(header_position - value_position) > 45:
                    continue
                percentage = parse_percentage(ratio_cell)
                if percentage is None and ratio_cell not in EMPTY_CELL:
                    continue
                label = cells[label_index]
                classification = classify_holder(label)
                header_context = normalize(" ".join(lines[max(0, header_index - 2):header_index + 1]))
                ratio_column = normalize(header_cells[column_index])
                ownership_column = (
                    any(marker in ratio_column for marker in ("so huu", "von dieu le", "shareholding", "ownership"))
                    or any(marker in header_context for marker in ("so co phan", "so luong co phan", "so co phieu", "number of shares", "share capital"))
                ) and not any(marker in ratio_column for marker in ("so luong co dong", "number of shareholders"))
                page_candidates.append({**page_data, "holder": label, "column": header_cells[column_index],
                                        "raw_value": ratio_cell, "percentage": percentage,
                                        "data_date": data_date, "date_note": date_note,
                                        "classification": classification, "ownership_column": ownership_column,
                                        "evidence": f"{header_line.strip()}\n{row_line.strip()}"})
        if not page_candidates:
            for line in lines:
                match = re.search(r"(?<!\d)(\d{1,3}(?:[.,]\d{1,4})?)\s*%", line)
                if not match or not any(term in normalize(line) for term in ("co dong", "nha nuoc", "so huu", "shareholder")):
                    continue
                holder = line[:match.start()].strip()
                page_candidates.append({**page_data, "holder": holder, "column": None,
                                        "raw_value": match.group(), "percentage": parse_percentage(match.group()),
                                        "data_date": None, "date_note": "Chưa xác định bảng/ngày/cột",
                                        "classification": classify_holder(holder), "ownership_column": False,
                                        "evidence": line.strip()})
            if not page_candidates:
                page_candidates.append({**page_data, "holder": None, "column": None, "raw_value": None,
                                        "percentage": None, "data_date": None, "date_note": "Chưa xác định bảng/ngày",
                                        "classification": "Trang cổ đông cần kiểm tra", "ownership_column": False,
                                        "evidence": excerpt(text, keywords[0], 1800)})
        candidates.extend(page_candidates)
        previous_by_file[file_name] = text
    return candidates


def validated_ratio(review: dict[str, Any], row: dict[str, Any]) -> tuple[float | None, str]:
    parsed = review.get("parsed") or {}
    decision = parsed.get("decision")
    if decision not in {"VALID", "CORRECT"}:
        return None, str(parsed.get("reason") or "Chưa đủ căn cứ")
    try:
        percentage = float(parsed["percentage"])
        data_date = date.fromisoformat(str(parsed["data_date"]))
    except (KeyError, TypeError, ValueError):
        return None, "API thiếu tỷ lệ hoặc ngày hợp lệ"
    if not (0 <= percentage <= 100 and data_date == date(int(row["Năm báo cáo"]), 12, 31)):
        return None, "Tỷ lệ/ngày ngoài phạm vi"
    reason = str(parsed.get("reason") or "")
    for raw_date in re.findall(r"(?:bảng ngày|danh sách chốt ngày)\s*(\d{1,2}/\d{1,2}/20\d{2})", reason, flags=re.IGNORECASE):
        day, month, year = map(int, raw_date.split("/"))
        if (year, month, day) != (data_date.year, data_date.month, data_date.day):
            return None, "Ngày của bảng/danh sách mâu thuẫn với ngày API chọn"
    evidence = str(row.get("Đoạn/bảng bằng chứng") or "")
    values = [float(raw.replace(",", ".")) for raw in re.findall(r"(?<!\d)(\d{1,3}(?:[.,]\d{1,4})?)\s*%", evidence)]
    has_value = any(abs(value - percentage) < 0.0001 for value in values)
    if not has_value:
        normalized = normalize(evidence)
        if "ty le" in normalized or "percentage" in normalized or "%" in evidence:
            decimals = re.findall(r"(?<!\d)(\d{1,3}[.,]\d{1,4})(?!\d)", evidence)
            has_value = any(abs(float(raw.replace(",", ".")) - percentage) < 0.0001 for raw in decimals)
    if not has_value:
        return None, "Tỷ lệ API không xuất hiện trong đoạn PDF"
    input_percentage = row.get("Tỷ lệ sở hữu Nhà nước (%)")
    if decision == "VALID" and input_percentage is not None and pd.notna(input_percentage):
        if abs(percentage - float(input_percentage)) >= 0.0001:
            return None, "API ghi VALID nhưng đổi tỷ lệ"
    return percentage, reason


def request_review(api_key: str, model: str, row: dict[str, Any], timeout: int) -> dict[str, Any]:
    evidence = str(row.get("Đoạn/bảng bằng chứng") or "")[:6500]
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": AUDIT_PROMPT},
            {"role": "user", "content": json.dumps({
                "ticker": row["Mã CK"],
                "report_year": int(row["Năm báo cáo"]),
                "input_percentage": row.get("Tỷ lệ sở hữu Nhà nước (%)"),
                "input_data_date": row.get("Ngày số liệu"),
                "page_text": evidence,
            }, ensure_ascii=False)},
        ],
        "temperature": 0,
    }
    for attempt in range(3):
        try:
            response = requests.post(
                "https://motapis.com/v1/chat/completions",
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
                json=payload,
                timeout=timeout,
            )
            if response.status_code == 429:
                time.sleep(2 ** (attempt + 1))
                continue
            response.raise_for_status()
            body = response.json()
            content = body["choices"][0]["message"]["content"]
            start, end = content.find("{"), content.rfind("}")
            if start < 0 or end < start:
                raise ValueError("API không trả JSON object")
            return {"parsed": json.loads(content[start : end + 1]), "usage": body.get("usage")}
        except (requests.RequestException, KeyError, ValueError) as exc:
            if attempt == 2:
                raise RuntimeError(f"{type(exc).__name__}: {exc}") from exc
            time.sleep(2 ** attempt)
    raise RuntimeError("API thất bại")


def page_score(text: str, report_year: int) -> tuple[int, str | None]:
    normalized = normalize(text)
    matches = [(term, weight) for term, normalized_term, weight in NORMALIZED_TERMS if normalized_term in normalized]
    if not matches:
        return 0, None
    term, score = max(matches, key=lambda item: item[1])
    if "%" in text:
        score += 15
    if str(report_year) in text:
        score += 8
    if re.search(r"(?:1[5-9]|2\d|3[01])\s*[/.-]\s*12\s*[/.-]\s*20\d{2}", text):
        score += 20
    if any(marker in normalized for marker in ("ty le so huu", "percentage", "so co phieu", "number of shares")):
        score += 12
    return score, term


def excerpt(text: str, term: str | None, limit: int = 900) -> str:
    compact_text = re.sub(r"\s+", " ", text).strip()
    if not compact_text:
        return ""
    if not term:
        return compact_text[:limit]
    normalized_text = normalize(compact_text)
    normalized_term = normalize(term)
    index = normalized_text.find(normalized_term)
    if index < 0:
        return compact_text[:limit]
    ratio = index / max(len(normalized_text), 1)
    center = int(ratio * len(compact_text))
    start = max(0, center - limit // 3)
    return compact_text[start : start + limit]


def load_pages(cache_path: Path, pdf_paths: list[Path] | None = None) -> list[dict[str, Any]]:
    signature = None
    if pdf_paths:
        digest = hashlib.sha256()
        for pdf_path in pdf_paths:
            stat = pdf_path.stat()
            digest.update(f"{pdf_path.name}|{stat.st_size}|{stat.st_mtime_ns}".encode("utf-8"))
        signature = digest.hexdigest()
    if cache_path.exists():
        with gzip.open(cache_path, "rt", encoding="utf-8") as handle:
            payload = json.load(handle)
        if signature is None or payload.get("signature") == signature:
            return payload.get("pages", [])
    if not pdf_paths:
        return []
    pages = []
    for pdf_path in pdf_paths:
        with fitz.open(pdf_path) as document:
            for page_number, page in enumerate(document, start=1):
                page_text = page.get_text("text", sort=True).strip()
                visible_chars = len(re.sub(r"\s", "", page_text))
                pages.append({
                    "pdf_file": pdf_path.name,
                    "page_number": page_number,
                    "text": page_text,
                    "visible_chars": visible_chars,
                    "needs_ocr": visible_chars < 80,
                })
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    with gzip.open(cache_path, "wt", encoding="utf-8", compresslevel=4) as handle:
        json.dump({"signature": signature, "pages": pages}, handle, ensure_ascii=False)
    return pages


def ocr_missing_pages(manifest: pd.DataFrame, cache_root: Path, ocr_cache_root: Path, pdf_root: Path, tessdata: Path, max_pages: int, dpi: int) -> dict[str, int]:
    from io import BytesIO

    from PIL import Image

    if tessdata.is_dir():
        os.environ["TESSDATA_PREFIX"] = str(tessdata.resolve())
    ocr_cache_root.mkdir(parents=True, exist_ok=True)
    reports = 0
    page_count = 0
    for item in manifest.to_dict("records"):
        ticker = str(item["Mã CK"]).strip().upper()
        year = int(item["Năm"])
        pdf_paths = sorted((pdf_root / ticker).glob(f"{ticker}_{year}*.pdf"))
        if not pdf_paths:
            continue
        cache_path = cache_root / f"{ticker}_{year}.json.gz"
        pages = load_pages(cache_path, pdf_paths)
        output_cache = ocr_cache_root / f"{ticker}_{year}.json.gz"
        existing_pages = load_ocr(output_cache).get("pages", [])
        existing_keys = {(page.get("pdf_file"), page.get("page_number")) for page in existing_pages}
        by_file: dict[str, list[dict[str, Any]]] = {}
        for page in pages:
            if page.get("needs_ocr") and (page.get("pdf_file"), page.get("page_number")) not in existing_keys:
                by_file.setdefault(str(page["pdf_file"]), []).append(page)
        selected = []
        for file_pages in by_file.values():
            if max_pages == 0 or len(file_pages) <= max_pages:
                selected.extend(file_pages)
                continue
            indexes = {round(index * (len(file_pages) - 1) / max(max_pages - 1, 1)) for index in range(max_pages)}
            selected.extend(file_pages[index] for index in sorted(indexes))
        if not selected:
            continue
        path_map = {path.name: path for path in pdf_paths}
        ocr_pages = []
        for selected_page in selected:
            with fitz.open(path_map[selected_page["pdf_file"]]) as document:
                page = document[selected_page["page_number"] - 1]
                pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), colorspace=fitz.csRGB, alpha=False)
                image = Image.open(BytesIO(pixmap.tobytes("png")))
                page_text = pytesseract.image_to_string(image, lang="vie+eng", config="--psm 6").strip()
            ocr_pages.append({**selected_page, "text": page_text, "visible_chars": len(re.sub(r"\s", "", page_text)), "needs_ocr": False, "ocr": True})
        all_ocr_pages = existing_pages + ocr_pages
        candidates = extract_ocr_candidates(ticker, year, all_ocr_pages)
        payload = {"ticker": ticker, "year": year, "ocr_pages": len(all_ocr_pages), "pages": all_ocr_pages, "candidates": candidates, "errors": []}
        with gzip.open(output_cache, "wt", encoding="utf-8", compresslevel=4) as handle:
            json.dump(payload, handle, ensure_ascii=False)
        reports += 1
        page_count += len(ocr_pages)
        print(f"OCR_PROGRESS reports={reports} pages={page_count} ticker={ticker} year={year}", flush=True)
    return {"Báo cáo OCR mới": reports, "Trang OCR mới": page_count}


def extract_ocr_candidates(ticker: str, year: int, pages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    candidates = []
    for page in pages:
        text = str(page.get("text") or "")
        dates = re.findall(r"(?<!\d)(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(20\d{2})(?!\d)", text)
        matching_dates = [f"{year}-12-{int(day):02d}" for day, month, found_year in dates if int(found_year) == year and int(month) == 12 and 15 <= int(day) <= 31]
        if not matching_dates:
            continue
        for line in text.splitlines():
            normalized_line = normalize(line)
            if not any(term in normalized_line for term in ("co dong nha nuoc", "so huu nha nuoc", "von nha nuoc", "state shareholder", "state ownership")):
                continue
            values = re.findall(r"(?<!\d)(\d{1,3}(?:[.,]\d{1,4})?)\s*%", line)
            if len(values) != 1:
                continue
            percentage = float(values[0].replace(",", "."))
            if not 0 < percentage <= 100:
                continue
            candidates.append({
                "ticker": ticker,
                "report_year": year,
                "percentage": percentage,
                "data_date": matching_dates[0],
                "pdf_file": page["pdf_file"],
                "page_number": page["page_number"],
                "keyword": "sở hữu Nhà nước",
                "evidence": text[:6500],
                "explicit_state_term": True,
                "score": 1,
            })
    return candidates


def find_fallback_page(pages: list[dict[str, Any]], report_year: int) -> dict[str, Any] | None:
    best = None
    for page in pages:
        text = str(page.get("text") or "")
        if int(page.get("visible_chars") or 0) < 80:
            continue
        score, term = page_score(text, report_year)
        if score <= 0:
            continue
        candidate = {
            "score": score,
            "term": term,
            "pdf_file": page.get("pdf_file"),
            "page_number": int(page.get("page_number") or 0),
            "excerpt": excerpt(text, term),
        }
        if best is None or (candidate["score"], -candidate["page_number"]) > (best["score"], -best["page_number"]):
            best = candidate
    return best


def load_ocr(cache_path: Path) -> dict[str, Any]:
    if not cache_path.exists():
        return {"pages": [], "candidates": [], "ocr_pages": 0, "errors": []}
    with gzip.open(cache_path, "rt", encoding="utf-8") as handle:
        return json.load(handle)


def load_audit(cache_path: Path) -> dict[str, Any]:
    if not cache_path.exists():
        return {"parsed": {"decision": "UNCLEAR", "reason": "Chưa kiểm định lại tỷ lệ cũ"}}
    return json.loads(cache_path.read_text(encoding="utf-8"))


def eligible_ocr_candidate(candidate: dict[str, Any], report_year: int) -> bool:
    try:
        data_date = date.fromisoformat(str(candidate["data_date"]))
        ratio = float(candidate["percentage"])
    except (KeyError, TypeError, ValueError):
        return False
    return (
        bool(candidate.get("explicit_state_term"))
        and 0 <= ratio <= 100
        and data_date == date(report_year, 12, 31)
    )


def audit_candidates(manifest: pd.DataFrame, details: pd.DataFrame, cache_root: Path, ocr_cache_root: Path, audit_cache_root: Path, fallback_audit_root: Path, ocr_audit_root: Path, pdf_root: Path, model: str) -> dict[str, int]:
    api_key = os.environ.get("MOTAPIS_API_KEY")
    if not api_key:
        raise RuntimeError("Thiếu MOTAPIS_API_KEY để dùng --audit")
    for directory in (audit_cache_root, fallback_audit_root, ocr_audit_root):
        directory.mkdir(parents=True, exist_ok=True)
    counts = {"Kiểm định lớp chữ mới": 0, "Kiểm định trang bỏ sót mới": 0, "Kiểm định OCR mới": 0}
    details_by_key = {
        (str(row["Mã CK"]).strip().upper(), int(row["Năm báo cáo"])): row
        for row in details.to_dict("records")
    }
    for item in manifest.to_dict("records"):
        ticker = str(item["Mã CK"]).strip().upper()
        year = int(item["Năm"])
        key = (ticker, year)
        detail = details_by_key.get(key)
        pdf_paths = sorted((pdf_root / ticker).glob(f"{ticker}_{year}*.pdf"))
        if not pdf_paths:
            continue
        if detail and detail.get("Trạng thái") == VERIFIED and pd.notna(detail.get("Tỷ lệ sở hữu Nhà nước (%)")):
            output = audit_cache_root / f"{ticker}_{year}.json"
            if not output.exists():
                review = request_review(api_key, model, detail, 180)
                output.write_text(json.dumps(review, ensure_ascii=False, indent=2), encoding="utf-8")
                counts["Kiểm định lớp chữ mới"] += 1
            continue
        pages = load_pages(cache_root / f"{ticker}_{year}.json.gz", pdf_paths)
        fallback = find_fallback_page(pages, year)
        if fallback:
            page = next((value for value in pages if value["pdf_file"] == fallback["pdf_file"] and value["page_number"] == fallback["page_number"]), None)
            if page:
                page_text = str(page["text"])
                normalized = normalize(page_text)
                explicit = any(term in normalized for term in ("co dong nha nuoc", "so huu nha nuoc", "von nha nuoc", "state shareholder", "state ownership"))
                late_date = bool(re.search(r"(?<!\d)(?:1[5-9]|2\d|3[01])\s*[/.-]\s*12\s*[/.-]\s*" + str(year) + r"(?!\d)", page_text))
                output = fallback_audit_root / f"{ticker}_{year}.json"
                if explicit and late_date and not output.exists():
                    row = {
                        "Mã CK": ticker,
                        "Năm báo cáo": year,
                        "Tỷ lệ sở hữu Nhà nước (%)": None,
                        "Ngày số liệu": None,
                        "Đoạn/bảng bằng chứng": page_text[:6500],
                    }
                    review = request_review(api_key, model, row, 180)
                    review.update({"pdf_file": fallback["pdf_file"], "page_number": fallback["page_number"], "evidence": row["Đoạn/bảng bằng chứng"]})
                    output.write_text(json.dumps(review, ensure_ascii=False, indent=2), encoding="utf-8")
                    counts["Kiểm định trang bỏ sót mới"] += 1
        ocr_payload = load_ocr(ocr_cache_root / f"{ticker}_{year}.json.gz")
        candidates = [candidate for candidate in ocr_payload.get("candidates", []) if eligible_ocr_candidate(candidate, year)]
        if candidates:
            candidate = max(candidates, key=lambda value: int(value.get("score") or 0))
            output = ocr_audit_root / f"{ticker}_{year}.json"
            if not output.exists():
                row = {
                    "Mã CK": ticker,
                    "Năm báo cáo": year,
                    "Tỷ lệ sở hữu Nhà nước (%)": candidate["percentage"],
                    "Ngày số liệu": candidate["data_date"],
                    "Đoạn/bảng bằng chứng": candidate["evidence"],
                }
                review = request_review(api_key, model, row, 180)
                review.update({"pdf_file": candidate["pdf_file"], "page_number": candidate["page_number"], "evidence": candidate["evidence"]})
                output.write_text(json.dumps(review, ensure_ascii=False, indent=2), encoding="utf-8")
                counts["Kiểm định OCR mới"] += 1
    return counts


def build_review(manifest_path: Path, result_path: Path, cache_root: Path, ocr_cache_root: Path, audit_cache_root: Path, fallback_audit_root: Path, ocr_audit_root: Path, pdf_root: Path, output_path: Path) -> dict[str, int]:
    if manifest_path.suffix.lower() == ".csv":
        manifest = pd.read_csv(manifest_path, encoding="utf-8-sig")
    else:
        manifest = pd.read_excel(manifest_path)
    if not result_path.exists():
        details = pd.DataFrame()
    elif result_path.suffix.lower() == ".csv":
        details = pd.read_csv(result_path, encoding="utf-8-sig")
    else:
        details = pd.read_excel(result_path, sheet_name="Chi_tiet")
    detail_by_key = {
        (str(row["Mã CK"]).strip().upper(), int(row["Năm báo cáo"])): row
        for row in details.to_dict("records")
    }

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Kiem_tra"
    sheet.append(["Mã CK", "Năm", "Tỷ lệ sở hữu Nhà nước (%)", "Trang PDF"])
    source_sheet = workbook.create_sheet("Nguon_kiem_tra")
    source_sheet.append(["Mã CK", "Năm", "File PDF", "Link nguồn", "Ngày số liệu", "Phương pháp", "Trạng thái", "Từ khóa", "Đoạn bằng chứng/trang", "Số trang OCR"])
    notes_sheet = workbook.create_sheet("Ghi_chu")
    notes_sheet.append(["Mục", "Giải thích"])
    notes_sheet.append(["Phạm vi", "101 mã CK × 6 năm = 606 dòng; 605 báo cáo đã có PDF, thiếu CHP 2020."])
    notes_sheet.append(["Cuối năm", "Chỉ lấy ngày dữ liệu 15/12–31/12 của năm báo cáo, không mặc định cả năm."])
    notes_sheet.append(["Không có", "Chưa tìm được tỷ lệ đủ điều kiện từ tài liệu đã đọc; KHÔNG có nghĩa tỷ lệ Nhà nước bằng 0%."])
    notes_sheet.append(["OCR", "OCR lấy mẫu trang scan để tìm ứng viên, chưa đọc hết mọi trang scan. Ứng viên OCR phải kiểm tra tay."])
    notes_sheet.append(["Kiểm định", "Tỷ lệ lớp chữ qua kiểm định API vẫn cần người dùng đối chiếu đúng dòng/cột trên trang PDF."])

    statistics = {
        "Tổng dòng mã-năm": 0,
        "Đã tìm kiếm PDF": 0,
        "Thiếu PDF": 0,
        "Tỷ lệ lớp chữ qua kiểm định": 0,
        "Tỷ lệ bổ sung từ trang bỏ sót": 0,
        "Tỷ lệ OCR qua kiểm định": 0,
        "Không có tỷ lệ nhưng có trang cổ đông": 0,
        "Không tìm thấy trang cổ đông bằng lớp chữ": 0,
        "Số trang đã OCR": 0,
    }

    rows = manifest.sort_values(["Mã CK", "Năm"]).to_dict("records")
    for excel_row, item in enumerate(rows, start=2):
        ticker = str(item["Mã CK"]).strip().upper()
        year = int(item["Năm"])
        detail = detail_by_key.get((ticker, year), {})
        cache_path = cache_root / f"{ticker}_{year}.json.gz"
        pdf_paths = sorted((pdf_root / ticker).glob(f"{ticker}_{year}*.pdf"))
        ocr_payload = load_ocr(ocr_cache_root / f"{ticker}_{year}.json.gz")
        statistics["Số trang đã OCR"] += int(ocr_payload.get("ocr_pages") or 0)
        has_pdf = bool(pdf_paths)
        status = detail.get("Trạng thái")
        old_ratio_present = status == VERIFIED and pd.notna(detail.get("Tỷ lệ sở hữu Nhà nước (%)"))
        audit_ratio = None
        audit_reason = ""
        audit_data_date = None
        if old_ratio_present:
            review = load_audit(audit_cache_root / f"{ticker}_{year}.json")
            audit_ratio, audit_reason = validated_ratio(review, detail)
            audit_data_date = (review.get("parsed") or {}).get("data_date")
        verified = old_ratio_present and audit_ratio is not None

        statistics["Tổng dòng mã-năm"] += 1
        statistics["Đã tìm kiếm PDF"] += int(has_pdf)
        statistics["Thiếu PDF"] += int(not has_pdf)

        pdf_file = None
        page_number: int | None = None
        page_comment = ""
        data_date = None
        keyword = None
        method = None
        evidence = None
        result_status = "Không tìm thấy tỷ lệ đủ điều kiện"
        if verified:
            ratio_value: float | str = float(audit_ratio)
            page_number = int(detail["Trang PDF"])
            pdf_file = str(detail["File PDF"])
            page_comment = (
                f"Tỷ lệ lớp chữ đã qua kiểm định API, vẫn cần kiểm tra tay.\nFile: {pdf_file}\nNgày số liệu: {audit_data_date}\n"
                f"Lý do kiểm định: {audit_reason}\n"
                f"Đoạn kiểm tra: {str(detail.get('Đoạn/bảng bằng chứng') or '')[:1200]}"
            )
            data_date = audit_data_date
            keyword = detail.get("Từ khóa tìm thấy")
            method = f"{detail.get('Phương pháp trích')} + kiểm định Motapis"
            evidence = detail.get("Đoạn/bảng bằng chứng")
            result_status = "Có tỷ lệ lớp chữ qua kiểm định; cần kiểm tra tay"
            statistics["Tỷ lệ lớp chữ qua kiểm định"] += 1
        else:
            ratio_value = "Không có"
            fallback_review_path = fallback_audit_root / f"{ticker}_{year}.json"
            fallback_ratio = None
            fallback_review = None
            if fallback_review_path.exists():
                fallback_review = json.loads(fallback_review_path.read_text(encoding="utf-8"))
                fallback_row = {
                    "Năm báo cáo": year,
                    "Tỷ lệ sở hữu Nhà nước (%)": None,
                    "Đoạn/bảng bằng chứng": fallback_review.get("evidence"),
                }
                fallback_ratio, _ = validated_ratio(fallback_review, fallback_row)
            if fallback_ratio is not None and fallback_review:
                ratio_value = fallback_ratio
                pdf_file = str(fallback_review["pdf_file"])
                page_number = int(fallback_review["page_number"])
                data_date = fallback_review["parsed"].get("data_date")
                keyword = "Cổ đông Nhà nước"
                method = "PyMuPDF + kiểm định Motapis"
                evidence = fallback_review.get("evidence")
                result_status = "Tỷ lệ bổ sung qua kiểm định; cần kiểm tra tay"
                page_comment = (
                    "Tỷ lệ bổ sung từ trang lớp chữ; cần kiểm tra trực tiếp dòng/cột/ngày.\n"
                    f"File: {pdf_file}\nNgày số liệu: {data_date}\n"
                    f"Lý do: {fallback_review['parsed'].get('reason')}\n"
                    f"Đoạn kiểm tra: {str(evidence or '')[:1000]}"
                )
                statistics["Tỷ lệ bổ sung từ trang bỏ sót"] += 1
            else:
                candidates = [candidate for candidate in ocr_payload.get("candidates", []) if eligible_ocr_candidate(candidate, year)]
                best_candidate = max(candidates, key=lambda candidate: int(candidate.get("score") or 0), default=None)
                ocr_ratio = None
                ocr_reason = ""
                ocr_review_path = ocr_audit_root / f"{ticker}_{year}.json"
                if best_candidate and ocr_review_path.exists():
                    ocr_review = json.loads(ocr_review_path.read_text(encoding="utf-8"))
                    ocr_ratio, ocr_reason = validated_ratio(ocr_review, {
                        "Năm báo cáo": year,
                        "Tỷ lệ sở hữu Nhà nước (%)": best_candidate["percentage"],
                        "Đoạn/bảng bằng chứng": best_candidate["evidence"],
                    })
                if best_candidate and ocr_ratio is not None:
                    ratio_value = float(ocr_ratio)
                    pdf_file = str(best_candidate["pdf_file"])
                    page_number = int(best_candidate["page_number"])
                    data_date = (ocr_review.get("parsed") or {}).get("data_date")
                    keyword = best_candidate.get("keyword")
                    method = "Tesseract OCR vie+eng"
                    evidence = best_candidate.get("evidence")
                    result_status = "Tỷ lệ OCR qua kiểm định; cần kiểm tra tay"
                    page_comment = (
                        "Tỷ lệ OCR qua kiểm định API, vẫn cần đối chiếu trực tiếp trên PDF.\n"
                        f"File: {pdf_file}\nNgày số liệu: {data_date}\n"
                        f"Lý do: {ocr_reason}\n"
                        f"Đoạn kiểm tra: {str(evidence or '')[:1200]}"
                    )
                    statistics["Tỷ lệ OCR qua kiểm định"] += 1
                else:
                    text_pages = load_pages(cache_path, pdf_paths)
                    fallback = find_fallback_page(text_pages + ocr_payload.get("pages", []), year)
                    if best_candidate and ocr_ratio is None:
                        fallback = {
                            "pdf_file": best_candidate["pdf_file"],
                            "page_number": int(best_candidate["page_number"]),
                            "term": "Ứng viên OCR chưa qua kiểm định",
                            "score": 0,
                            "excerpt": f"{ocr_reason}; {str(best_candidate.get('evidence') or '')[:800]}",
                        }
                    if old_ratio_present and audit_ratio is None and detail.get("File PDF") and pd.notna(detail.get("Trang PDF")):
                        fallback = {
                            "pdf_file": detail["File PDF"],
                            "page_number": int(detail["Trang PDF"]),
                            "term": "Tỷ lệ cũ không qua kiểm định",
                            "score": 0,
                            "excerpt": f"{audit_reason}; {str(detail.get('Đoạn/bảng bằng chứng') or '')[:800]}",
                        }
                    if fallback:
                        pdf_file = str(fallback["pdf_file"])
                        page_number = int(fallback["page_number"])
                        keyword = fallback["term"]
                        method = "PyMuPDF/Tesseract OCR + từ khóa"
                        evidence = fallback["excerpt"]
                        result_status = "Không có tỷ lệ; có trang cổ đông để kiểm tra"
                        page_comment = (
                            "Chưa có tỷ lệ đủ điều kiện; đây là trang cổ đông để kiểm tra tay.\n"
                            f"File: {pdf_file}\nTừ khóa: {fallback['term']}\nĐiểm: {fallback['score']}\n"
                            f"Đoạn kiểm tra: {fallback['excerpt']}"
                        )
                        statistics["Không có tỷ lệ nhưng có trang cổ đông"] += 1
                    elif has_pdf:
                        result_status = "Chưa tìm thấy trang cổ đông; cần OCR thêm/kiểm tra tay"
                        statistics["Không tìm thấy trang cổ đông bằng lớp chữ"] += 1
                    else:
                        result_status = "Thiếu PDF"

        page_value: int | str
        if page_number:
            page_value = page_number
        elif has_pdf:
            page_value = "Không tìm thấy"
        else:
            page_value = "Không có PDF"
        sheet.append([ticker, year, ratio_value, page_value])
        source_sheet.append([ticker, year, pdf_file, item.get("URL"), data_date, method, result_status, keyword, str(evidence or "")[:3000], int(ocr_payload.get("ocr_pages") or 0)])

        page_cell = sheet.cell(excel_row, 4)
        if pdf_file and page_number:
            pdf_path = (pdf_root / ticker / pdf_file).resolve()
            if pdf_path.exists():
                page_cell.hyperlink = f"{pdf_path.as_uri()}#page={page_number}"
                page_cell.style = "Hyperlink"
            page_cell.comment = Comment(page_comment, "Codex")
            if result_status.startswith("Ứng viên OCR"):
                sheet.cell(excel_row, 3).comment = Comment("Tỷ lệ OCR là ứng viên, chưa xác minh độc lập. Kiểm tra trực tiếp trang PDF.", "Codex")

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 12, "B": 10, "C": 30, "D": 18}.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        row[1].number_format = "0"
        if isinstance(row[2].value, (int, float)):
            row[2].number_format = "0.0000"
        if isinstance(row[3].value, (int, float)):
            row[3].number_format = "0"
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for cell in source_sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    source_sheet.freeze_panes = "A2"
    source_sheet.auto_filter.ref = source_sheet.dimensions
    for column, width in {"A": 12, "B": 10, "C": 28, "D": 65, "E": 16, "F": 28, "G": 52, "H": 24, "I": 90, "J": 14}.items():
        source_sheet.column_dimensions[column].width = width
    for cell in notes_sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    notes_sheet.column_dimensions["A"].width = 22
    notes_sheet.column_dimensions["B"].width = 105
    for row in notes_sheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(output_path)
    except PermissionError as exc:
        raise RuntimeError(f"Không thể ghi {output_path}; hãy đóng file Excel đang mở hoặc dùng --output với tên khác") from exc
    return statistics


def choose_state_total(candidates: list[dict[str, Any]], year: int) -> tuple[dict[str, Any] | None, str]:
    target_date = f"{year}-12-31"
    qualified = [candidate for candidate in candidates
                 if candidate.get("classification") == "Tổng sở hữu Nhà nước"
                 and candidate.get("percentage") is not None
                 and candidate.get("data_date") == target_date
                 and candidate.get("column") and candidate.get("ownership_column") and candidate.get("evidence")]
    values = {round(float(candidate["percentage"]), 6) for candidate in qualified}
    if len(values) > 1:
        return None, "Mâu thuẫn tỷ lệ cùng ngày; cần kiểm tra"
    if qualified:
        return max(qualified, key=lambda candidate: (candidate["score"], -candidate["page_number"])), "Đã xác minh từ bảng; cần đối chiếu PDF"
    if any(candidate.get("classification") == "Tổng sở hữu Nhà nước" for candidate in candidates):
        return None, "Có tỷ lệ/bảng Nhà nước nhưng thiếu ngày hoặc cột hợp lệ"
    if any(candidate.get("holder") for candidate in candidates):
        return None, "Có bảng cổ đông; chưa xác minh tỷ lệ Nhà nước"
    return None, "Chưa tìm được bảng cổ đông"


def build_candidate_review(manifest_path: Path, result_path: Path, cache_root: Path, ocr_cache_root: Path,
                           pdf_root: Path, output_path: Path) -> dict[str, int]:
    manifest = pd.read_csv(manifest_path, encoding="utf-8-sig") if manifest_path.suffix.lower() == ".csv" else pd.read_excel(manifest_path)
    if not result_path.exists():
        details = pd.DataFrame()
    elif result_path.suffix.lower() == ".csv":
        details = pd.read_csv(result_path, encoding="utf-8-sig")
    else:
        details = pd.read_excel(result_path, sheet_name="Chi_tiet")
    detail_by_key: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for row in details.to_dict("records"):
        detail_by_key.setdefault((str(row["Mã CK"]).strip().upper(), int(row["Năm báo cáo"])), []).append(row)
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Kiem_tra"
    summary.append(["Mã CK", "Năm mục tiêu", "Tỷ lệ Nhà nước (%)", "Trạng thái", "Tên cổ đông/chỉ tiêu ứng viên",
                    "Tỷ lệ ứng viên (%)", "Ngày số liệu", "Trang PDF", "Từ khóa", "Bằng chứng", "Nguồn"])
    inventory = workbook.create_sheet("Ung_vien")
    inventory.append(["Mã CK", "Năm mục tiêu", "File PDF", "Trang PDF", "Nhóm phân loại", "Tên cổ đông/chỉ tiêu",
                      "Tên cột", "Giá trị gốc", "Tỷ lệ ứng viên (%)", "Ngày số liệu", "Ghi chú ngày",
                      "Từ khóa", "Bằng chứng", "Nguồn", "Phương pháp", "Điểm ưu tiên"])
    notes = workbook.create_sheet("Ghi_chu")
    notes.append(["Mục", "Giải thích"])
    notes.append(["Ngày mục tiêu", "Chỉ chốt tỷ lệ từ bảng ghi rõ ngày 31/12 của năm; ngày khác giữ tại Ung_vien."])
    notes.append(["Ứng viên", "Tỷ lệ cổ đông khác, tổ chức, cá nhân, vốn HĐQT và dữ liệu cũ không tự chuyển thành tỷ lệ Nhà nước."])
    notes.append(["Ô trống", "Không có bằng chứng không tương đương 0%; dấu gạch ngang giữ nguyên tại Giá trị gốc."])
    notes.append(["OCR", "Chỉ các trang đã OCR mới được tìm; dùng --ocr-missing với --max-pages 0 để quét toàn bộ trang scan."])
    statistics = {"Tổng dòng mã-năm": 0, "Có tỷ lệ Nhà nước từ bảng": 0, "Còn cần kiểm tra": 0,
                  "Thiếu PDF": 0, "Số ứng viên/trang": 0}
    for item in manifest.sort_values(["Mã CK", "Năm"]).to_dict("records"):
        ticker, year = str(item["Mã CK"]).strip().upper(), int(item["Năm"])
        pdf_paths = sorted((pdf_root / ticker).glob(f"{ticker}_{year}*.pdf"))
        available_pdf_names = {path.name for path in pdf_paths}
        source = item.get("URL")
        if pd.isna(source):
            source = None
        pages = load_pages(cache_root / f"{ticker}_{year}.json.gz", pdf_paths)
        ocr_pages = [{**page, "ocr": True} for page in load_ocr(ocr_cache_root / f"{ticker}_{year}.json.gz").get("pages", [])]
        ocr_keys = {(page.get("pdf_file"), page.get("page_number")) for page in ocr_pages}
        searchable_pages = [page for page in pages if (page.get("pdf_file"), page.get("page_number")) not in ocr_keys] + ocr_pages
        candidates = extract_page_candidates(ticker, year, searchable_pages, source)
        for detail in detail_by_key.get((ticker, year), []):
            if not pd.notna(detail.get("Tỷ lệ sở hữu Nhà nước (%)")):
                continue
            candidates.append({"ticker": ticker, "year": year, "pdf_file": detail.get("File PDF"),
                               "page_number": int(detail["Trang PDF"]) if pd.notna(detail.get("Trang PDF")) else None,
                               "source": source, "keywords": detail.get("Từ khóa tìm thấy"), "score": -1,
                               "method": "Ứng viên từ kết quả cũ", "holder": "Tỷ lệ cũ chưa đối chiếu lại",
                               "column": None, "raw_value": detail.get("Tỷ lệ sở hữu Nhà nước (%)"),
                               "percentage": float(detail["Tỷ lệ sở hữu Nhà nước (%)"]), "data_date": None,
                               "date_note": "Cần đọc lại ngày trên bảng", "classification": "Dữ liệu cũ cần kiểm tra",
                               "evidence": str(detail.get("Đoạn/bảng bằng chứng") or "")[:1800]})
        candidates.sort(key=lambda candidate: (-int(candidate.get("score") or 0),
                                                str(candidate.get("pdf_file") or ""), int(candidate.get("page_number") or 0)))
        selected, status = choose_state_total(candidates, year)
        if not pdf_paths:
            status = "Thiếu PDF; chưa đủ thông tin"
            statistics["Thiếu PDF"] += 1
        featured = selected or next((candidate for candidate in candidates if candidate.get("holder")), None) or (candidates[0] if candidates else {})
        summary.append([ticker, year, selected["percentage"] if selected else None, status,
                        featured.get("holder"), featured.get("percentage"), featured.get("data_date"),
                        featured.get("page_number"), featured.get("keywords"), featured.get("evidence"), source])
        if selected:
            statistics["Có tỷ lệ Nhà nước từ bảng"] += 1
        else:
            statistics["Còn cần kiểm tra"] += 1
        statistics["Tổng dòng mã-năm"] += 1
        statistics["Số ứng viên/trang"] += len(candidates)
        for candidate in candidates:
            inventory.append([ticker, year, candidate.get("pdf_file"), candidate.get("page_number"),
                              candidate.get("classification"), candidate.get("holder"), candidate.get("column"),
                              candidate.get("raw_value"), candidate.get("percentage"), candidate.get("data_date"),
                              candidate.get("date_note"), candidate.get("keywords"), candidate.get("evidence"),
                              source, candidate.get("method"), candidate.get("score")])
            if candidate.get("pdf_file") and candidate.get("page_number"):
                pdf_path = pdf_root / ticker / str(candidate["pdf_file"])
                if pdf_path.name in available_pdf_names:
                    inventory.cell(inventory.max_row, 4).hyperlink = f"{pdf_path.resolve().as_uri()}#page={candidate['page_number']}"
                    inventory.cell(inventory.max_row, 4).style = "Hyperlink"
        for sheet, row_number, candidate, page_column in ((summary, summary.max_row, featured, 8),):
            file_name, page_number = candidate.get("pdf_file"), candidate.get("page_number")
            if file_name and page_number:
                pdf_path = (pdf_root / ticker / str(file_name)).resolve()
                if pdf_path.exists():
                    sheet.cell(row_number, page_column).hyperlink = f"{pdf_path.as_uri()}#page={page_number}"
                    sheet.cell(row_number, page_column).style = "Hyperlink"
    for sheet in (summary, inventory, notes):
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(wrap_text=True)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = 72 if letter in {"J", "M"} else 25
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(output_path)
    except PermissionError as exc:
        raise RuntimeError(f"Không thể ghi {output_path}; đóng file Excel hoặc chọn --output khác") from exc
    return statistics


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Xử lý PDF sở hữu Nhà nước và tạo bảng 606 mã-năm")
    parser.add_argument("--manifest", type=Path, default=Path("danh_muc_pdf_101x6.csv"))
    parser.add_argument("--results", type=Path, default=Path("ownership_extract_101x6/ratio_candidates.csv"))
    parser.add_argument("--cache-root", type=Path, default=Path("ownership_extract_101x6/page_cache"))
    parser.add_argument("--pdf-root", type=Path, default=Path("pdf_101x6"))
    parser.add_argument("--ocr-cache-root", type=Path, default=Path("ownership_extract_101x6/ocr_cache_v2_160_6"))
    parser.add_argument("--audit-cache-root", type=Path, default=Path("ownership_extract_101x6/audit_text_ratios"))
    parser.add_argument("--fallback-audit-root", type=Path, default=Path("ownership_extract_101x6/audit_fallback_ratios"))
    parser.add_argument("--ocr-audit-root", type=Path, default=Path("ownership_extract_101x6/audit_ocr_ratios"))
    parser.add_argument("--ocr-missing", action="store_true", help="OCR mọi trang scan chưa được OCR, kể cả khi đã thấy trang cổ đông")
    parser.add_argument("--audit", action="store_true", help="Chế độ kiểm định cũ (tạm khóa để tránh chọn một ứng viên rồi bỏ phần còn lại)")
    parser.add_argument("--max-pages", type=int, default=0, help="Số trang scan tối đa mỗi PDF mỗi lượt; 0 = toàn bộ")
    parser.add_argument("--dpi", type=int, default=160)
    parser.add_argument("--tessdata", type=Path, default=Path("ownership_cli/tessdata"))
    parser.add_argument("--audit-model", default="gpt-5.6-sol")
    parser.add_argument("--output", type=Path, default=Path("SO_HUU_NHA_NUOC_101_MA_2020_2025.xlsx"))
    args = parser.parse_args()
    if args.max_pages < 0 or args.dpi <= 0:
        parser.error("--max-pages phải không âm và --dpi phải dương")
    if args.audit:
        parser.error("--audit cũ chỉ kiểm định một ứng viên và chưa tương thích với bảng Ung_vien; chưa gọi API để tránh kết quả thiếu")
    if args.ocr_missing:
        if args.manifest.suffix.lower() == ".csv":
            manifest = pd.read_csv(args.manifest, encoding="utf-8-sig")
        else:
            manifest = pd.read_excel(args.manifest)
        if args.ocr_missing:
            ocr_statistics = ocr_missing_pages(manifest, args.cache_root, args.ocr_cache_root, args.pdf_root, args.tessdata, args.max_pages, args.dpi)
            print(json.dumps(ocr_statistics, ensure_ascii=False), flush=True)
    statistics = build_candidate_review(args.manifest, args.results, args.cache_root, args.ocr_cache_root, args.pdf_root, args.output)
    print(json.dumps({**statistics, "File kết quả": str(args.output)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
