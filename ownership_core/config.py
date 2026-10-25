"""
ownership_core.config
=====================
Cấu hình tập trung, hằng số, từ khóa và bộ quy tắc nhận diện sở hữu Nhà nước.
"""

from __future__ import annotations
import os
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Set, Tuple

# 1. ĐƯỜNG DẪN DỰ ÁN & MẶC ĐỊNH
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PDF_DIR = PROJECT_ROOT / "pdf_101x6"
DEFAULT_MANIFEST = PROJECT_ROOT / "danh_muc_pdf_101x6.csv"
DEFAULT_CHECKPOINT = PROJECT_ROOT / "checkpoint_ownership.jsonl"
DEFAULT_EXCEL_OUT = PROJECT_ROOT / "Ket_Qua_So_Huu_Nha_Nuoc.xlsx"
DEFAULT_CACHE_DIR = PROJECT_ROOT / ".cache_ownership"

# 2. CẤU HÌNH OCR TESSERACT
TESSERACT_EXE_WINDOWS = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
TESSDATA_LOCAL = PROJECT_ROOT / "ownership_cli" / "tessdata"
DEFAULT_OCR_DPI = 160
OCR_LANGUAGES = "vie+eng"

# 3. CẤU HÌNH LOCAL MODEL & GEMINI API
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5-coder:7b")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

# 4. BỘ TỪ KHÓA TÌM BẢNG CƠ CẤU CỔ ĐÔNG (Hỗ trợ có dấu, không dấu, tiếng Anh)
PRIMARY_KEYWORDS = [
    "cổ đông nhà nước",
    "sở hữu nhà nước",
    "vốn nhà nước",
    "state shareholder",
    "state shareholders",
    "state ownership",
    "cơ cấu cổ đông",
    "cơ cấu sở hữu",
    "cơ cấu vốn cổ phần",
    "shareholder structure",
    "ownership structure",
    "danh sách cổ đông lớn",
    "thông tin cổ đông lớn",
    "major shareholders",
    "vốn góp của chủ sở hữu",
    "owners capital contribution",
    "scic",
    "kinh doanh vốn nhà nước",
    "ủy ban quản lý vốn",
]

SECONDARY_KEYWORDS = [
    "cổ đông",
    "cổ phần",
    "vốn điều lệ",
    "shareholders",
    "shares",
    "cổ đông lớn",
    "cổ đông tổ chức",
    "cổ đông sáng lập",
]

# 5. CÁC TỪ KHÓA PHÂN BIỆT CỘT TỶ LỆ CỔ PHẦN VS TỶ LỆ SỐ CỔ ĐÔNG
COLUMN_SHARE_RATIO_KEYWORDS = [
    "tỷ lệ sở hữu",
    "tỷ lệ nắm giữ",
    "tỷ lệ (%)",
    "tỷ lệ vốn",
    "tỷ lệ vđl",
    "tỷ lệ trên vốn điều lệ",
    "tỷ lệ cp",
    "tỷ lệ cổ phần",
    "tỷ lệ cổ phiếu",
    "ownership percentage",
    "shareholding percentage",
    "percentage of shares",
    "percentage of charter capital",
    "% sở hữu",
    "% vốn",
]

COLUMN_SHAREHOLDER_COUNT_KEYWORDS = [
    "số lượng cổ đông",
    "số cổ đông",
    "tỷ lệ số lượng cổ đông",
    "tỷ lệ số cổ đông",
    "number of shareholders",
    "shareholder count",
    "% số cổ đông",
    "% cổ đông",
]

# 6. DANH XÁCH TỔ CHỨC ĐẠI DIỆN VỐN NHÀ NƯỚC CHUẨN XÁC
DIRECT_STATE_ENTITIES = [
    "scic",
    "tổng công ty đầu tư và kinh doanh vốn nhà nước",
    "ủy ban quản lý vốn nhà nước tại doanh nghiệp",
    "cmsc",
    "bộ tài chính",
    "bộ công thương",
    "bộ giao thông vận tải",
    "bộ xây dựng",
    "bộ nông nghiệp và phát triển nông thôn",
    "bộ thông tin và truyền thông",
    "bộ y tế",
    "bộ quốc phòng",
    "ubnd thành phố hà nội",
    "ubnd thành phố hồ chí minh",
    "ubnd tp hà nội",
    "ubnd tp hồ chí minh",
    "ubnd tỉnh",
    "ubnd thành phố",
    "văn phòng tỉnh ủy",
    "văn phòng thành ủy",
    "tổng cục",
    "cục dự trữ nhà nước",
    "ngân hàng nhà nước việt nam",
]

# 7. QUY CHUẨN CHUỖI KHÔNG DẤU (NORMALIZATION)
def normalize_text(text: str) -> str:
    """Chuẩn hóa văn bản tiếng Việt sang chữ thường không dấu, loại bỏ ký tự lạ."""
    if not text:
        return ""
    text = text.casefold().replace("đ", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[^a-z0-9%./\s\-]+", " ", text)
    return " ".join(text.split())

# 8. MẪU PROMPT CHUẨN - TUYỆT ĐỐI KHÔNG CHỨA SỐ MẪU CỐ ĐỊNH NHƯ 43.31 HAY 56626237
STRICT_CANDIDATE_EXTRACTION_PROMPT = """Bạn là chuyên gia kiểm định dữ liệu cổ đông doanh nghiệp niêm yết tại Việt Nam.
Nhiệm vụ: Trích xuất danh sách ứng viên cổ đông đại diện vốn Nhà nước hoặc dòng tổng 'Nhà nước' từ nội dung văn bản báo cáo.

Tên tài liệu: {file_name}
Mã chứng khoán: {ticker}
Năm báo cáo: {year}

VĂN BẢN NGUỒN:
{source_text}

QUY TẮC BẮT BUỘC (TUÂN THỦ 100%):
1. CHỈ trích xuất số liệu thực sự xuất hiện trong văn bản nguồn bên trên. KHÔNG tự bịa, KHÔNG tự suy đoán số liệu.
2. Nếu không tìm thấy thông tin cổ đông Nhà nước, trả về mảng rỗng `candidates: []`.
3. Phân biệt rõ:
   - Cột tỷ lệ sở hữu vốn/cổ phần (Shareholding %) với Cột tỷ lệ số lượng cổ đông (Shareholder count %).
   - Số cổ phiếu (đơn vị: cổ phiếu) với Giá trị vốn góp (đơn vị: VNĐ, đồng).
   - Dòng tổng Nhà nước với phần vốn do từng cá nhân HĐQT đại diện.
4. Trích dẫn chính xác đoạn văn/dòng bảng làm bằng chứng tại `evidence`.
5. ĐẦU RA BẮT BUỘC LÀ JSON THUẦN TÚY (Schema bên dưới, không markdown, không văn bản phụ):

{{
  "ticker": "{ticker}",
  "report_year": {year},
  "has_state_ownership": true_hoac_false_hoac_null,
  "candidates": [
    {{
      "holder_name": "<Tên cổ đông hoặc nhãn dòng chính xác>",
      "row_label": "<Nhãn dòng trong bảng>",
      "column_label": "<Tên cột tỷ lệ sở hữu>",
      "raw_percentage": "<Chuỗi tỷ lệ gốc trong bảng, ví dụ: '51,08%' hoặc '-' hoặc null>",
      "percentage": <Số thực từ 0.0 đến 100.0 hoặc null nếu không có>,
      "share_count": <Số lượng cổ phiếu dạng số nguyên hoặc null nếu không có>,
      "raw_date": "<Ngày chốt danh sách hoặc ngày bảng ghi trong tài liệu>",
      "normalized_date": "<YYYY-MM-DD hoặc null>",
      "page_number": <Số trang nếu biết hoặc null>,
      "evidence": "<Đoạn text chính xác trong nguồn chứng minh con số này>"
    }}
  ],
  "reason": "<Tóm tắt ngắn gọn lý do>"
}}
"""
