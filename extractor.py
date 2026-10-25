#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
========================================================================================
CHƯƠNG TRÌNH TRÍCH XUẤT "TỶ LỆ SỞ HỮU NHÀ NƯỚC" TỪ BÁO CÁO THƯỜNG NIÊN (2020-2025)
========================================================================================
Senior Data Engineering Pipeline:
- Lớp 1 (Local Pre-filtering - 0 token): Quét nhanh text từng trang qua PyMuPDF (fitz),
  tính điểm mật độ từ khóa cơ cấu cổ đông, chọn top trang + trang liền kề.
  Phát hiện PDF Scan thuần ảnh (<100 ký tự) để đánh dấu PDF_SCAN_CAN_OCR, bỏ qua gọi API.
- Lớp 2 (Multi-threaded Gemini API): Chạy song song 8 luồng với 8 API Keys xoay vòng
  (Round-Robin), model gemini-2.5-flash qua thư viện chuẩn google-genai.
  Quản lý Rate Limit (sleep 4.2s), cơ chế bắt lỗi 429 Retry đổi key (chờ 15s).
- Checkpoint & Resume: Lưu tức thời vào JSONL sau mỗi file, tự động tiếp tục khi chạy lại,
  xuất file Excel định dạng chuyên nghiệp mỗi 20 file và khi hoàn tất.
========================================================================================
"""

import os
import sys
import re
import json
import time
import logging
import argparse
import threading
import subprocess
import tempfile
import urllib.request
import urllib.error
from collections import defaultdict, deque
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Any, List, Tuple, Optional, Set

# Hỗ trợ in UTF-8 tiếng Việt trên console Windows
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from tqdm import tqdm
from dotenv import load_dotenv

# Thư viện chuẩn Google GenAI mới nhất
from google import genai
from google.genai import types

# Tắt cảnh báo AFC / httpx không cần thiết để giữ giao diện console gọn gàng
logging.getLogger("google.genai").setLevel(logging.ERROR)
logging.getLogger("httpx").setLevel(logging.WARNING)

# ==============================================================================
# CẤU HÌNH VÀ THAM SỐ TOÀN CỤC
# ==============================================================================
load_dotenv(override=True)

# 1. BỘ TỪ KHÓA LỌC TRANG CƠ CẤU CỔ ĐÔNG (LỚP 1)
KEYWORDS: List[str] = [
    "cơ cấu cổ đông",
    "cổ đông lớn",
    "cổ đông nhà nước",
    "sở hữu nhà nước",
    "danh sách cổ đông",
    "scic",
    "kinh doanh vốn nhà nước",
    "ủy ban quản lý vốn",
    "cổ đông tổ chức",
    "vốn góp của chủ sở hữu"
]

# 2. CẤU HÌNH MÔI TRƯỜNG & ĐƯỜNG DẪN MẶC ĐỊNH
DEFAULT_PDF_FOLDER = os.getenv("PDF_FOLDER_PATH", "pdf_101x6")
DEFAULT_OUTPUT_EXCEL = os.getenv("OUTPUT_EXCEL_PATH", "Ket_Qua_So_Huu_Nha_Nuoc.xlsx")
DEFAULT_CHECKPOINT = os.getenv("CHECKPOINT_PATH", "checkpoint_results.jsonl")
DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
DEFAULT_WORKERS = int(os.getenv("MAX_WORKERS", "8"))

# 3. CẤU HÌNH THUẬT TOÁN SLIDING WINDOW RATE LIMITER & COOLDOWN
MAX_RPM_PER_KEY = int(os.getenv("MAX_RPM_PER_KEY", "15"))  # Giới hạn 15 req/phút/key
MIN_REQUEST_INTERVAL = float(os.getenv("MIN_REQUEST_INTERVAL", "4.1"))  # Giãn cách tối thiểu giữa 2 req cùng key
WINDOW_SECONDS = 60.0  # Độ dài cửa sổ trượt (60 giây)
RATE_LIMIT_SLEEP = float(os.getenv("RATE_LIMIT_SLEEP", "0.0"))
RETRY_DELAY_429 = float(os.getenv("RETRY_DELAY_429", "15.0"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))
EXCEL_SAVE_INTERVAL = int(os.getenv("EXCEL_SAVE_INTERVAL", "20"))
MIN_SCAN_TEXT_LENGTH = 100
MAX_TOP_PAGES = 6

# 4. CẤU HÌNH DỰ PHÒNG HYBRID (OLLAMA LOCAL MODEL)
ENABLE_OLLAMA_FALLBACK = os.getenv("ENABLE_OLLAMA_FALLBACK", "true").lower() == "true"
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5-coder:7b")

# 5. CẤU HÌNH BỎ QUA DOANH NGHIỆP ĐÃ CÓ SỞ HỮU NHÀ NƯỚC
SKIP_IF_STATE_OWNED = os.getenv("SKIP_IF_STATE_OWNED", "true").lower() == "true"
REFERENCE_EXCEL_PATH = os.getenv("REFERENCE_EXCEL_PATH", "SO_HUU_NHA_NUOC_101_MA_UNG_VIEN_2020_2025.xlsx")

# ==============================================================================
# CẤU HÌNH LOGGING
# ==============================================================================
LOG_FILE = "extractor.log"
logger = logging.getLogger("StateOwnershipExtractor")
logger.setLevel(logging.INFO)

# Handler ghi file chi tiết
file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
file_handler.setLevel(logging.INFO)
file_formatter = logging.Formatter("%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s")
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# ==============================================================================
# PROMPT GỬI LÊN GEMINI API
# ==============================================================================
PROMPT_TEMPLATE = """Bạn là chuyên gia phân tích dữ liệu tài chính và quản trị doanh nghiệp tại Việt Nam.
Nhiệm vụ của bạn là trích xuất chính xác "Tỷ lệ sở hữu Nhà nước" (State Ownership Percentage) từ văn bản trích xuất của Báo cáo thường niên hoặc Báo cáo tài chính dưới đây.

Tên file gốc: {file_name}

NỘI DUNG VĂN BẢN TRÍCH XUẤT:
{text_content}

HƯỚNG DẪN PHÂN TÍCH & SUY LUẬN (BẮT BUỘC TUÂN THỦ):
1. Tìm kiếm tại bảng "Cơ cấu cổ đông", "Danh sách cổ đông lớn" hoặc thuyết minh "Vốn chủ sở hữu":
   - Trường hợp 1 (Ghi trực tiếp): Tìm dòng "Cổ đông Nhà nước" / "Sở hữu Nhà nước". Nếu có con số cụ thể, lấy tỷ lệ % đó. Nếu có dòng "Nhà nước" nhưng ô số lượng và tỷ lệ ĐỂ TRỐNG hoặc gạch ngang (-), kết luận tỷ lệ là 0.0%.
   - Trường hợp 2 (Ghi qua tên tổ chức đại diện vốn Nhà nước): Nếu bảng không chia theo dòng "Nhà nước", hãy kiểm tra danh sách "Cổ đông lớn" hoặc "Cổ đông tổ chức". Cộng tổng tỷ lệ sở hữu (%) của các tổ chức đại diện vốn Nhà nước nếu có, bao gồm:
     + Tổng công ty Đầu tư và Kinh doanh vốn Nhà nước (SCIC).
     + Ủy ban Quản lý vốn nhà nước tại doanh nghiệp (CMSC), các Bộ, UBND các tỉnh/thành phố, Văn phòng Thành ủy/Tỉnh ủy.
     + Các Tập đoàn/Tổng công ty Nhà nước nắm vốn chi phối (ví dụ: PVN/Petrovietnam, EVN, TKV/Vinacomin, Petrolimex, Vinachem, Vicem, Vinatex, Vinafood, Becamex IDC, HFIC, ACV, VIMC, Vinataba, Tổng công ty Đầu tư Phát triển Nhà và Đô thị HUD, Tổng công ty Sông Đà, Tổng công ty Lắp máy Lilama, v.v.).
   - Trường hợp 3 (Doanh nghiệp tư nhân / Đã thoái hết vốn): Nếu "Cổ đông lớn (>=5%)" bằng 0% và bảng cơ cấu cổ đông không đề cập đến Nhà nước hay tổ chức Nhà nước nào, kết luận tỷ lệ là 0.0%.

2. CẢNH BÁO TRÁNH NHẦM LẪN:
   - KHÔNG căn cứ vào phần "Lịch sử hình thành" (ví dụ: tiền thân là xí nghiệp quốc doanh hay công ty nhà nước cổ phần hóa năm xưa) để kết luận hiện tại có vốn nhà nước. Chỉ lấy số liệu tại bảng cơ cấu cổ đông của năm báo cáo.
   - KHÔNG nhầm lẫn tỷ lệ sở hữu của công ty mẹ tư nhân hoặc cổ đông nước ngoài thành sở hữu nhà nước.

YÊU CẦU ĐẦU RA (JSON thuần):
{{
    "ma_ck": "Mã chứng khoán 3 chữ cái in hoa",
    "nam_bao_cao": 2024,
    "co_so_huu_nha_nuoc": true,
    "ty_le_so_huu_nha_nuoc_percent": 43.31,
    "so_luong_co_phieu_nha_nuoc": 56626237,
    "ten_co_dong_nha_nuoc": "Tên tổ chức đại diện vốn nhà nước (ví dụ: SCIC). Nếu 0% thì ghi 'Không có'",
    "ngay_chot_danh_sach": "Ngày chốt danh sách cổ đông nếu có (ví dụ: 25/12/2024)",
    "bang_chung_trich_xuat": "Trích dẫn ngắn gọn câu/dòng trong báo cáo làm căn cứ kết luận",
    "trang_thai": "THANH_CONG hoặc KHONG_TIM_THAY_THONG_TIN"
}}"""

# PROMPT DÀNH CHO FILE PDF DẠNG SCAN (GEMINI MULTIMODAL VISION)
SCAN_PROMPT_TEMPLATE = """Bạn là chuyên gia phân tích dữ liệu tài chính và quản trị doanh nghiệp tại Việt Nam.
Dưới đây là tài liệu PDF Báo cáo thường niên / Báo cáo tài chính (dạng scan ảnh) của doanh nghiệp: {file_name}.
Nhiệm vụ của bạn là đọc kỹ toàn bộ tài liệu, tìm phần "Cơ cấu cổ đông", "Danh sách cổ đông lớn", "Vốn góp của chủ sở hữu" để trích xuất chính xác "Tỷ lệ sở hữu Nhà nước" (State Ownership Percentage).

Quy tắc bắt buộc:
1. Tìm tên các cổ đông đại diện vốn Nhà nước: SCIC (Tổng công ty Đầu tư và Kinh doanh vốn Nhà nước), các Tập đoàn/Tổng công ty Nhà nước, Ủy ban Quản lý vốn Nhà nước tại doanh nghiệp, các Bộ, UBND Tỉnh/Thành phố,...
2. Nếu có nhiều cổ đông Nhà nước, hãy cộng tổng tỷ lệ % của tất cả lại.
3. Nếu không có cổ đông Nhà nước nào hoặc 100% tư nhân/nước ngoài, ghi tỷ lệ là 0.0 và co_so_huu_nha_nuoc là false.
4. Trích dẫn chính xác câu văn, tên bảng và số liệu làm bằng chứng.

Chỉ trả về định dạng JSON thuần túy (không markdown, không giải thích):
{{
  "ma_ck": "{ma_ck}",
  "nam_bao_cao": {nam},
  "co_so_huu_nha_nuoc": true,
  "ty_le_so_huu_nha_nuoc_percent": 51.0,
  "so_luong_co_phieu_nha_nuoc": 10000000,
  "ten_co_dong_nha_nuoc": "Tên cơ quan/cổ đông Nhà nước",
  "ngay_chot_danh_sach": "DD/MM/YYYY hoặc N/A",
  "bang_chung_trich_xuat": "Trích dẫn bằng chứng từ tài liệu",
  "trang_thai": "THANH_CONG hoặc KHONG_TIM_THAY_THONG_TIN"
}}"""

# ==============================================================================
# QUẢN LÝ XOAY VÒNG API KEY & RATE LIMIT (ROUND-ROBIN)
# ==============================================================================
class ApiKeyManager:
    """
    Quản lý danh sách API Keys kết hợp thuật toán Sliding Window Rate Limiter:
    - Giới hạn tối đa MAX_RPM_PER_KEY (15 requests/phút/key) trên cửa sổ trượt 60 giây.
    - Làm mịn lưu lượng (Token Smoothing): giãn cách tối thiểu MIN_REQUEST_INTERVAL (4.1s) giữa 2 request cùng key.
    - Khi một key chạm ngưỡng 15 request/phút: Tự động tính toán 'khoảng ngắt'
      (thời gian cần chờ để request cũ nhất rời khỏi khung 60s) và tạm dừng luồng,
      sau đó tiếp tục chu trình lặp lại khi quota phục hồi.
    - Hỗ trợ Cooldown thông minh khi gặp lỗi HTTP 429 và tự động điều phối sang key khả dụng có tải thấp nhất.
    """

    def __init__(self):
        self._keys: List[str] = self._load_keys()
        if not self._keys:
            raise ValueError("KHÔNG tìm thấy API key nào trong .env hoặc file listapi!")
        self._index = 0
        self._global_lock = threading.Lock()
        self._key_locks: Dict[int, threading.Lock] = defaultdict(threading.Lock)
        self._history: Dict[int, deque] = defaultdict(deque)
        self._cooldown: Dict[int, float] = {}
        logger.info(
            f"Đã nạp thành công {len(self._keys)} Gemini API Keys với thuật toán Sliding Window Rate Limiter "
            f"({MAX_RPM_PER_KEY} req/phút/key, giãn cách {MIN_REQUEST_INTERVAL}s)."
        )

    def _load_keys(self) -> List[str]:
        keys = []
        # 1. Ưu tiên đọc trực tiếp từ file listapi cục bộ
        listapi_path = Path("listapi")
        if listapi_path.exists():
            try:
                with open(listapi_path, "r", encoding="utf-8") as f:
                    for line in f:
                        k = line.strip().strip("'\"").rstrip(",").strip()
                        if k and not k.startswith("#") and k not in keys:
                            keys.append(k)
                if keys:
                    logger.info(f"Đã đọc {len(keys)} API keys trực tiếp từ file {listapi_path.resolve()}")
            except Exception as e:
                logger.warning(f"Không thể đọc file listapi: {e}")

        # 2. Đọc bổ sung từ biến môi trường GEMINI_API_KEY_1 -> GEMINI_API_KEY_16
        for i in range(1, 17):
            val = os.getenv(f"GEMINI_API_KEY_{i}")
            if val and val.strip() and not val.startswith("GEMINI_API_KEY"):
                cleaned = val.strip().strip("'\"").rstrip(",").strip()
                if cleaned and cleaned not in keys:
                    keys.append(cleaned)

        # 3. Đọc thêm GEMINI_API_KEY nếu có
        single_key = os.getenv("GEMINI_API_KEY")
        if single_key and single_key.strip():
            for k in single_key.split(","):
                k_clean = k.strip().strip("'\"").rstrip(",").strip()
                if k_clean and k_clean not in keys:
                    keys.append(k_clean)

        return keys

    @property
    def total_keys(self) -> int:
        return len(self._keys)

    def get_key(self, worker_index: Optional[int] = None) -> Tuple[str, int]:
        with self._global_lock:
            if worker_index is not None:
                idx = worker_index % len(self._keys)
            else:
                idx = self._index % len(self._keys)
                self._index += 1
            return self._keys[idx], idx

    def mark_cooldown(self, key_idx: int, cooldown_seconds: float = 30.0):
        """Đánh dấu key bị lỗi 429 tạm thời đưa vào trạng thái nghỉ."""
        with self._global_lock:
            self._cooldown[key_idx] = time.time() + cooldown_seconds
            logger.warning(
                f"⚠️ [RateLimiter] Đánh dấu Cooldown {cooldown_seconds:.1f}s cho Key #{key_idx + 1} do dính giới hạn API (429)."
            )

    def get_best_available_key(
        self,
        preferred_index: Optional[int] = None,
        exclude_indices: Optional[Set[int]] = None
    ) -> Tuple[str, int]:
        """Tìm key tối ưu nhất: không trong cooldown và có số request trong 60s ít nhất."""
        with self._global_lock:
            now = time.time()
            exclude = exclude_indices or set()
            candidates = [i for i in range(len(self._keys)) if i not in exclude]
            if not candidates:
                candidates = list(range(len(self._keys)))

            # Lọc các key không bị cooldown
            available = [i for i in candidates if self._cooldown.get(i, 0) <= now]
            if not available:
                # Nếu tất cả đều bị cooldown, chọn key có thời gian cooldown hết hạn sớm nhất
                available = sorted(candidates, key=lambda i: self._cooldown.get(i, 0))

            # Nếu preferred_index hợp lệ và còn dung lượng (< MAX_RPM_PER_KEY), ưu tiên dùng
            if preferred_index is not None and (preferred_index % len(self._keys)) in available:
                pref_idx = preferred_index % len(self._keys)
                q = self._history[pref_idx]
                while q and q[0] <= now - WINDOW_SECONDS:
                    q.popleft()
                if len(q) < MAX_RPM_PER_KEY:
                    return self._keys[pref_idx], pref_idx

            # Chọn key có số request trong cửa sổ 60s ít nhất (Least Loaded)
            def key_load(idx: int) -> int:
                q = self._history[idx]
                while q and q[0] <= now - WINDOW_SECONDS:
                    q.popleft()
                return len(q)

            best_idx = min(available, key=key_load)
            return self._keys[best_idx], best_idx

    def acquire_rate_limit(self, key_idx: int, worker_name: str = "") -> float:
        """
        Thuật toán Sliding Window Log Rate Limiter:
        - Kiểm tra và chờ cooldown nếu key bị 429 trước đó.
        - Loại bỏ các request đã quá 60s khỏi cửa sổ trượt.
        - Nếu đã đạt 15 request/phút: tính toán khoảng ngắt chính xác và tạm dừng.
        - Đảm bảo giãn cách tối thiểu 4.1s giữa 2 request trên cùng 1 key.
        - Ghi nhận timestamp của request mới vào cửa sổ trượt.
        """
        with self._key_locks[key_idx]:
            while True:
                now = time.time()

                # 1. Kiểm tra Cooldown lỗi 429
                cd_end = self._cooldown.get(key_idx, 0)
                if cd_end > now:
                    sleep_cd = cd_end - now
                    logger.info(
                        f"[{worker_name}] Key #{key_idx + 1} đang trong thời gian Cooldown 429, chờ {sleep_cd:.1f}s..."
                    )
                    time.sleep(sleep_cd)
                    continue

                q = self._history[key_idx]

                # 2. Xóa các request cũ hơn cửa sổ trượt 60 giây
                while q and q[0] <= now - WINDOW_SECONDS:
                    q.popleft()

                # 3. Kiểm tra ngưỡng 15 requests / phút (cho khoảng ngắt khi chạm ngưỡng)
                if len(q) >= MAX_RPM_PER_KEY:
                    oldest_req = q[0]
                    # Khoảng ngắt = thời gian để request cũ nhất trong cửa sổ 60s hết hiệu lực + buffer an toàn
                    pause_time = max(1.0, WINDOW_SECONDS - (now - oldest_req) + 0.5)
                    logger.warning(
                        f"⏸️ [{worker_name}] Key #{key_idx + 1} đã chạm ngưỡng {len(q)}/{MAX_RPM_PER_KEY} requests/phút. "
                        f"Bắt đầu KHOẢNG NGẮT {pause_time:.1f}s để hồi quota (tránh lỗi 429 Too Many Requests)..."
                    )
                    time.sleep(pause_time)
                    continue  # Sau khoảng ngắt, lặp lại kiểm tra để đảm bảo an toàn

                # 4. Làm mịn lưu lượng (Token Smoothing): Giãn cách tối thiểu giữa 2 request trên cùng 1 key
                if q:
                    last_req = q[-1]
                    elapsed = now - last_req
                    if elapsed < MIN_REQUEST_INTERVAL:
                        smooth_sleep = MIN_REQUEST_INTERVAL - elapsed
                        time.sleep(smooth_sleep)

                # 5. Ghi nhận thời điểm request hợp lệ vào cửa sổ trượt
                req_time = time.time()
                q.append(req_time)
                return req_time


# ==============================================================================
# CHECKPOINT & RESUME MANAGER
# ==============================================================================
class CheckpointManager:
    """Quản lý lưu tức thời, phục hồi trạng thái và bỏ qua các file/mã đã có thông tin sở hữu nhà nước."""

    def __init__(
        self,
        checkpoint_path: str,
        reference_excel_path: Optional[str] = "SO_HUU_NHA_NUOC_101_MA_UNG_VIEN_2020_2025.xlsx"
    ):
        self.checkpoint_path = Path(checkpoint_path)
        self.reference_excel_path = Path(reference_excel_path) if reference_excel_path else None
        self._lock = threading.Lock()
        self.completed_files: Set[str] = set()
        self.known_state_pairs: Set[Tuple[str, int]] = set()  # Các cặp (Mã CK, Năm) đã có thông tin sở hữu nhà nước
        self.known_state_tickers: Set[str] = set()  # Các mã CK đã xác nhận có sở hữu Nhà nước
        self.records_by_file: Dict[str, Dict[str, Any]] = {}  # Lưu bản ghi mới nhất theo tên file (chống trùng lặp Excel)
        self._load_existing()

    def _load_existing(self):
        # 1. Đọc từ file checkpoint JSONL
        if self.checkpoint_path.exists():
            try:
                with open(self.checkpoint_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            record = json.loads(line)
                            file_name = record.get("file_goc")
                            status = record.get("trang_thai")
                            if not file_name:
                                continue

                            # Chuẩn hóa Mã CK và Năm Báo Cáo chính quy từ đường dẫn/tên file
                            file_p = Path(record.get("duong_dan_file") or record.get("file_goc") or "")
                            canonical_ma, canonical_nam = extract_metadata_from_filename(file_p)
                            if canonical_ma != "UNKNOWN":
                                record["ma_ck"] = canonical_ma
                            else:
                                record["ma_ck"] = str(record.get("ma_ck", "")).strip().upper()
                            if canonical_nam > 0:
                                record["nam_bao_cao"] = canonical_nam

                            # Giữ bản ghi mới nhất của file (loại bỏ các bản ghi lỗi cũ bị ghi đè)
                            self.records_by_file[file_name] = record

                            # Chỉ coi là hoàn thành nếu trích xuất thành công hoặc không tìm thấy thông tin
                            # Tuyệt đối KHÔNG coi lỗi API hoặc PDF_SCAN là hoàn thành để hệ thống tự động quét lại bằng AI
                            if status in ["THANH_CONG", "THANH_CONG_OLLAMA_LOCAL", "KHONG_TIM_THAY_THONG_TIN"]:
                                self.completed_files.add(file_name)
                            else:
                                self.completed_files.discard(file_name)
                            
                            # Đã có thông tin sở hữu nhà nước hợp lệ
                            if status in ["THANH_CONG", "THANH_CONG_OLLAMA_LOCAL"]:
                                has_state = record.get("co_so_huu_nha_nuoc") is True
                                pct = record.get("ty_le_so_huu_nha_nuoc_percent") or 0.0
                                if has_state or pct > 0:
                                    m = str(record.get("ma_ck", "")).strip().upper()
                                    y = int(record.get("nam_bao_cao", 0) or 0)
                                    if m and y > 0:
                                        self.known_state_pairs.add((m, y))
                                    if m:
                                        self.known_state_tickers.add(m)
                        except json.JSONDecodeError:
                            continue
                logger.info(f"Đã nạp {len(self.completed_files)} file đã xử lý hợp lệ từ checkpoint {self.checkpoint_path.name}.")
            except Exception as e:
                logger.error(f"Lỗi khi đọc file checkpoint {self.checkpoint_path}: {e}")

        # 2. Đọc từ file Excel ứng viên tham chiếu nếu có
        if self.reference_excel_path and self.reference_excel_path.exists():
            try:
                ref_df = pd.read_excel(self.reference_excel_path)
                for _, row in ref_df.iterrows():
                    pct_val = row.get("Tỷ lệ Nhà nước (%)")
                    status_str = str(row.get("Trạng thái", ""))
                    if pd.notna(pct_val) or "Đã xác minh" in status_str:
                        ticker = str(row.get("Mã CK", "")).strip().upper()
                        try:
                            year = int(row.get("Năm mục tiêu", 0))
                            if ticker and year > 0:
                                self.known_state_pairs.add((ticker, year))
                            if ticker:
                                self.known_state_tickers.add(ticker)
                        except (ValueError, TypeError):
                            pass
                logger.info(
                    f"Đã nạp tổng cộng {len(self.known_state_pairs)} cặp (Mã CK, Năm) và "
                    f"{len(self.known_state_tickers)} Mã CK ĐÃ CÓ THÔNG TIN SỞ HỮU NHÀ NƯỚC."
                )
            except Exception as e:
                logger.warning(f"Không thể đọc file tham chiếu {self.reference_excel_path}: {e}")

    def check_file(
        self,
        file_name: str,
        ma_ck: Optional[str] = None,
        nam: Optional[int] = None,
        skip_if_has_state: bool = True,
        skip_ticker_level: bool = False
    ) -> Tuple[bool, str]:
        with self._lock:
            # Kiểm tra 1: File gốc đã được xử lý hoàn tất trong checkpoint
            if file_name in self.completed_files:
                return True, "File đã hoàn thành"

            # Kiểm tra 2: Toàn bộ Mã CK đã được xác định có sở hữu nhà nước (nếu bật chế độ cấp mã CK)
            if skip_ticker_level and ma_ck:
                ticker_upper = ma_ck.strip().upper()
                if ticker_upper in self.known_state_tickers:
                    return True, f"Mã doanh nghiệp {ticker_upper} ĐÃ CÓ thông tin sở hữu Nhà nước"

            # Kiểm tra 3: Cặp (Mã CK, Năm) ĐÃ CÓ thông tin sở hữu nhà nước
            if skip_if_has_state and ma_ck and nam:
                ticker_upper = ma_ck.strip().upper()
                try:
                    year_int = int(nam)
                    if (ticker_upper, year_int) in self.known_state_pairs:
                        return True, f"Mã {ticker_upper} năm {year_int} ĐÃ CÓ thông tin sở hữu Nhà nước"
                except (ValueError, TypeError):
                    pass

            return False, ""

    def is_completed(
        self,
        file_name: str,
        ma_ck: Optional[str] = None,
        nam: Optional[int] = None,
        skip_if_has_state: bool = True,
        skip_ticker_level: bool = False
    ) -> bool:
        done, _ = self.check_file(file_name, ma_ck, nam, skip_if_has_state, skip_ticker_level)
        return done

    def append_record(self, record: Dict[str, Any]):
        with self._lock:
            file_name = record.get("file_goc")
            status = record.get("trang_thai")
            if file_name:
                self.records_by_file[file_name] = record

            if file_name and status in ["THANH_CONG", "THANH_CONG_OLLAMA_LOCAL", "KHONG_TIM_THAY_THONG_TIN"]:
                self.completed_files.add(file_name)
            else:
                if file_name:
                    self.completed_files.discard(file_name)

            if status in ["THANH_CONG", "THANH_CONG_OLLAMA_LOCAL"]:
                has_state = record.get("co_so_huu_nha_nuoc") is True
                pct = record.get("ty_le_so_huu_nha_nuoc_percent") or 0.0
                if has_state or pct > 0:
                    m = str(record.get("ma_ck", "")).strip().upper()
                    y = int(record.get("nam_bao_cao", 0) or 0)
                    if m and y > 0:
                        self.known_state_pairs.add((m, y))
                    if m:
                        self.known_state_tickers.add(m)

            try:
                with open(self.checkpoint_path, "a", encoding="utf-8") as f:
                    f.write(json.dumps(record, ensure_ascii=False) + "\n")
                    f.flush()
            except Exception as e:
                logger.error(f"Không thể ghi checkpoint cho file {file_name}: {e}")

    def get_all_records(self) -> List[Dict[str, Any]]:
        with self._lock:
            return list(self.records_by_file.values())


# ==============================================================================
# HÀM BỔ TRỢ: TÁCH METADATA TỪ TÊN FILE
# ==============================================================================
def extract_metadata_from_filename(file_path: Path) -> Tuple[str, int]:
    """
    Tách mã CK (3 chữ cái/ký tự in hoa) và năm báo cáo (2020-2025) từ đường dẫn file.
    Ưu tiên 1: Tên thư mục cha là mã chứng khoán (chuẩn 100% trong pdf_101x6/<MÃ_CK>/...).
    Ưu tiên 2: Regex trên tên file.
    """
    filename = file_path.name

    # 1. Tìm năm báo cáo (2020-2025)
    match_year = re.search(r"(202[0-5])", filename)
    nam = int(match_year.group(1)) if match_year else 0

    # 2. Tìm mã CK
    # Thư mục cha là mã chứng khoán chuẩn xác nhất (101 thư mục)
    parent_name = file_path.parent.name.strip().upper()
    if len(parent_name) == 3 and parent_name.isalnum():
        return parent_name, nam

    # Fallback regex trên tên file
    pattern_start = re.compile(r"^([A-Za-z0-9]{3})[_\-\s]")
    match_start = pattern_start.search(filename)
    if match_start:
        return match_start.group(1).upper(), nam

    pattern_inner = re.compile(r"([A-Za-z0-9]{3})[_\-\s].*?(202[0-5])")
    match_inner = pattern_inner.search(filename)
    if match_inner:
        return match_inner.group(1).upper(), nam

    match_code = re.search(r"([A-Za-z]{3}|[A-Za-z]{2}[0-9])", filename)
    if match_code:
        return match_code.group(1).upper(), nam

    return parent_name if (len(parent_name) == 3 and parent_name.isalnum()) else "UNKNOWN", nam


# ==============================================================================
# LỚP 1: LOCAL PRE-FILTERING (0 TOKEN)
# ==============================================================================
def extract_relevant_pages_local(
    pdf_path: Path,
    keywords: List[str] = KEYWORDS,
    max_top_pages: int = MAX_TOP_PAGES,
    min_scan_text: int = MIN_SCAN_TEXT_LENGTH
) -> Tuple[bool, str, List[int], int]:
    """
    Quét nhanh text từng trang qua PyMuPDF ở máy cục bộ (0 token API):
    - Tính điểm mật độ từ khóa cơ cấu cổ đông.
    - Chọn tối đa 6 trang có mật độ từ khóa cao nhất.
    - Lấy kèm 1 trang liền sau để chống bảng bị ngắt qua trang kế.
    - Phát hiện PDF Scan (<100 ký tự text cả file).
    Trả về: (is_scan, combined_text, selected_pages_1_indexed, total_pages)
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        logger.error(f"Lỗi không mở được file PDF {pdf_path.name}: {e}")
        return False, "", [], 0

    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        return True, "", [], 0

    total_text_len = 0
    page_texts: Dict[int, str] = {}
    page_scores: List[Tuple[int, int]] = []

    for idx in range(total_pages):
        try:
            p_text = doc[idx].get_text()
        except Exception:
            p_text = ""
        total_text_len += len(p_text.strip())
        page_texts[idx] = p_text

        text_lower = p_text.lower()
        score = sum(text_lower.count(kw) for kw in keywords)
        if score > 0:
            page_scores.append((idx, score))

    doc.close()

    # Phát hiện PDF Scan thuần ảnh
    if total_text_len < min_scan_text:
        return True, "", [], total_pages

    # Chọn tối đa 6 trang có điểm cao nhất
    page_scores.sort(key=lambda x: x[1], reverse=True)
    top_candidates = page_scores[:max_top_pages]

    chosen_indices: Set[int] = set()
    for idx, score in top_candidates:
        chosen_indices.add(idx)
        # Lấy kèm 1 trang liền sau đề phòng bảng ngắt trang
        if idx + 1 < total_pages:
            chosen_indices.add(idx + 1)

    # Nới lỏng từ khóa nếu không có trang nào khớp chính xác
    if not chosen_indices:
        loose_terms = ["cổ đông", "sở hữu", "vốn điều lệ"]
        for idx, text in page_texts.items():
            t_low = text.lower()
            if any(lt in t_low for lt in loose_terms):
                chosen_indices.add(idx)
                if idx + 1 < total_pages:
                    chosen_indices.add(idx + 1)
                if len(chosen_indices) >= max_top_pages * 2:
                    break

    # Nếu vẫn không thấy, lấy 6 trang đầu
    if not chosen_indices:
        for idx in range(min(6, total_pages)):
            chosen_indices.add(idx)

    sorted_indices = sorted(list(chosen_indices))

    text_blocks = []
    selected_pages_1_indexed = []
    for idx in sorted_indices:
        selected_pages_1_indexed.append(idx + 1)
        raw_p_text = page_texts.get(idx, "").strip()
        text_blocks.append(f"--- [TRANG {idx + 1}/{total_pages}] ---\n{raw_p_text}\n")

    combined_text = "\n".join(text_blocks)
    return False, combined_text, selected_pages_1_indexed, total_pages


def extract_scan_pages_with_ocr(pdf_path: Path, max_pages: int = 25) -> Tuple[str, List[int]]:
    """
    Sử dụng Tesseract OCR cục bộ để trích xuất text từ các trang của PDF scan:
    - Quét tối đa 25 trang đầu tiên (nơi chứa cơ cấu cổ đông).
    - Lọc các trang có từ khóa liên quan đến cơ cấu cổ đông / vốn sở hữu.
    """
    extracted_text_blocks = []
    selected_pages = []
    try:
        doc = fitz.open(str(pdf_path))
        num_pages = min(max_pages, len(doc))
        for pno in range(num_pages):
            pix = doc[pno].get_pixmap(dpi=150)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                tmp_name = tmp.name
            pix.save(tmp_name)
            res = subprocess.run(
                ['tesseract', tmp_name, 'stdout', '-l', 'eng'],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            Path(tmp_name).unlink(missing_ok=True)
            txt = (res.stdout or "").strip()
            txt_lower = txt.lower()
            if any(k in txt_lower for k in ["shareholder", "capital", "state", "cổ đông", "sở hữu", "nhà nước", "scic", "vốn", "owner"]):
                extracted_text_blocks.append(f"--- [TRANG {pno + 1}/{len(doc)} (OCR)] ---\n" + txt)
                selected_pages.append(pno + 1)
        doc.close()
    except Exception as e:
        logger.warning(f"Lỗi OCR file {pdf_path.name}: {e}")

    if not extracted_text_blocks:
        try:
            doc = fitz.open(str(pdf_path))
            for pno in range(min(5, len(doc))):
                pix = doc[pno].get_pixmap(dpi=150)
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                    tmp_name = tmp.name
                pix.save(tmp_name)
                res = subprocess.run(
                    ['tesseract', tmp_name, 'stdout', '-l', 'eng'],
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='ignore'
                )
                Path(tmp_name).unlink(missing_ok=True)
                extracted_text_blocks.append(f"--- [TRANG {pno + 1} (OCR)] ---\n" + (res.stdout or "").strip())
                selected_pages.append(pno + 1)
            doc.close()
        except Exception:
            pass

    return "\n\n".join(extracted_text_blocks), selected_pages


# ==============================================================================
# LỚP 2: GỌI GEMINI API VỚI RETRY & XOAY VÒNG KEY
# ==============================================================================
def clean_and_parse_json(raw_text: str) -> Optional[Dict[str, Any]]:
    """Phân tích chuỗi JSON trả về an toàn từ LLM."""
    if not raw_text:
        return None
    cleaned = raw_text.strip()
    if cleaned.startswith("```json"):
        cleaned = cleaned[7:]
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:]
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3]
    cleaned = cleaned.strip()

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"(\{.*\})", cleaned, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                pass
    return None


def call_gemini_with_resilience(
    key_manager: ApiKeyManager,
    worker_idx: int,
    file_name: str,
    text_content: str = "",
    pdf_path: Optional[Path] = None,
    is_scan: bool = False,
    ma_ck: str = "",
    nam: int = 0,
    model_name: str = DEFAULT_MODEL,
    max_retries: int = MAX_RETRIES,
    sleep_after: float = RATE_LIMIT_SLEEP
) -> Tuple[Optional[Dict[str, Any]], str]:
    """
    Gọi Gemini API xử lý cả Text thông thường và PDF Scan Multimodal:
    - Nếu là PDF Scan: Đọc tối đa 100 trang và gửi trực tiếp qua Multimodal Vision AI (không cần OCR ngoài).
    - Áp dụng thuật toán Sliding Window kiểm soát chính xác 15 request/phút/key.
    - Tự động ngắt nghỉ khi chạm ngưỡng 15 request/phút để hồi quota.
    - Nếu dính lỗi 429: đưa key vào Cooldown RETRY_DELAY_429 (15s), chọn key khả dụng khác và retry.
    - Nếu gặp 404: Tự động đổi linh hoạt giữa gemini-3.8-flash và gemini-2.5-flash theo hỗ trợ của key.
    """
    current_key, current_key_idx = key_manager.get_best_available_key(preferred_index=worker_idx)
    active_model = model_name

    # Chuẩn bị payload Multimodal nếu là PDF Scan
    pdf_part = None
    large_pdf_file = None
    if is_scan and pdf_path and pdf_path.exists():
        prompt = SCAN_PROMPT_TEMPLATE.format(file_name=file_name, ma_ck=ma_ck, nam=nam)
        try:
            doc = fitz.open(str(pdf_path))
            total_doc_pages = len(doc)
            # Giới hạn tối đa 100 trang đầu tiên nếu file quá dài
            if total_doc_pages > 100:
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=0, to_page=min(99, total_doc_pages - 1))
                pdf_bytes = new_doc.tobytes()
                new_doc.close()
            else:
                with open(pdf_path, "rb") as f:
                    pdf_bytes = f.read()
            doc.close()

            # Nếu file bytes <= 20MB: Tạo Part inline
            if pdf_bytes and len(pdf_bytes) <= 20 * 1024 * 1024:
                pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
        except Exception as e:
            logger.warning(f"[{file_name}] Lỗi đọc PDF scan để gửi AI: {e}")
            pdf_bytes = None
    else:
        prompt = PROMPT_TEMPLATE.format(file_name=file_name, text_content=text_content)

    for attempt in range(1, max_retries + 1):
        uploaded_file_obj = None
        try:
            # 1. BẮT BUỘC: Đi qua thuật toán Sliding Window Rate Limiter để đảm bảo <= 15 RPM
            key_manager.acquire_rate_limit(current_key_idx, worker_name=f"Worker-{worker_idx + 1}")

            client = genai.Client(api_key=current_key)
            config = types.GenerateContentConfig(
                response_mime_type="application/json",
                temperature=0.0
            )

            # 2. Xây dựng nội dung gửi lên Gemini
            if is_scan:
                if pdf_part:
                    contents = [pdf_part, prompt]
                elif pdf_path and pdf_path.exists():
                    # File lớn hơn 20MB -> Upload qua Google File API
                    uploaded_file_obj = client.files.upload(file=str(pdf_path))
                    contents = [uploaded_file_obj, prompt]
                else:
                    contents = prompt
            else:
                contents = prompt

            response = client.models.generate_content(
                model=active_model,
                contents=contents,
                config=config
            )

            # Nghỉ bổ sung nếu có cấu hình
            if sleep_after > 0:
                time.sleep(sleep_after)

            if not response or not response.text:
                raise ValueError("API trả về phản hồi rỗng.")

            parsed_json = clean_and_parse_json(response.text)
            if not parsed_json:
                raise ValueError(f"Không thể parse JSON từ phản hồi: {response.text[:200]}")

            return parsed_json, "THANH_CONG"

        except Exception as e:
            err_str = str(e)
            is_429 = "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "ResourceExhausted" in err_str
            is_404 = "404" in err_str

            if is_429:
                logger.warning(
                    f"[{file_name}] Gặp lỗi 429 tại Key #{current_key_idx + 1}. "
                    f"Kích hoạt Cooldown {RETRY_DELAY_429}s và chuyển sang key khả dụng khác (Thử lại {attempt}/{max_retries})."
                )
                key_manager.mark_cooldown(current_key_idx, RETRY_DELAY_429)
                current_key, current_key_idx = key_manager.get_best_available_key(exclude_indices={current_key_idx})
                continue

            if is_404:
                alt_model = "gemini-3.8-flash" if active_model == "gemini-2.5-flash" else "gemini-2.5-flash"
                logger.warning(
                    f"[{file_name}] Key #{current_key_idx + 1} gặp lỗi 404 với model {active_model}. "
                    f"Tự động chuyển sang model thay thế nhẹ nhất: {alt_model}..."
                )
                active_model = alt_model
                time.sleep(0.5)
                continue

            logger.warning(
                f"[{file_name}] Lỗi API key #{current_key_idx + 1}: {err_str[:120]}. "
                f"Thử lại {attempt}/{max_retries}..."
            )
            time.sleep(2.0 * attempt)
            current_key, current_key_idx = key_manager.get_best_available_key(exclude_indices={current_key_idx})
        finally:
            if uploaded_file_obj:
                try:
                    client.files.delete(name=uploaded_file_obj.name)
                except Exception:
                    pass

    return None, "LOI_API_VUOT_QUA_SO_LAN_THU"


# ==============================================================================
# LỚP 2 (DỰ PHÒNG): GỌI OLLAMA LOCAL MODEL (QWEN 2.5:7B / LLAMA 3.1)
# ==============================================================================
def check_ollama_available(base_url: str = OLLAMA_BASE_URL) -> bool:
    """Kiểm tra xem Ollama server có đang chạy cục bộ không."""
    try:
        req = urllib.request.Request(f"{base_url.rstrip('/')}/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=1.5) as resp:
            return resp.status == 200
    except Exception:
        return False


def call_ollama_local(
    prompt: str,
    file_name: str,
    model_name: str = OLLAMA_MODEL,
    base_url: str = OLLAMA_BASE_URL,
    timeout: float = 120.0
) -> Tuple[Optional[Dict[str, Any]], str]:
    """
    Gọi model Local qua Ollama API (native /api/chat với format='json'):
    - Tối ưu 100% cho GPU NVIDIA RTX 4050 6GB VRAM, RAM 16GB.
    - Tự động bóc tách và chuẩn hóa JSON.
    """
    url = f"{base_url.rstrip('/')}/api/chat"
    payload = {
        "model": model_name,
        "messages": [{"role": "user", "content": prompt}],
        "format": "json",
        "stream": False,
        "options": {
            "temperature": 0.0,
            "num_predict": 1024
        }
    }
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data_bytes,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    try:
        logger.info(f"[{file_name}] Đang gọi Ollama Local Model ({model_name}) tại {base_url}...")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            result_raw = json.loads(resp.read().decode("utf-8"))
            message_content = result_raw.get("message", {}).get("content", "")
            parsed = clean_and_parse_json(message_content)
            if parsed:
                return parsed, "THANH_CONG_OLLAMA_LOCAL"
            else:
                return None, "LOI_PARSE_JSON_OLLAMA"
    except urllib.error.URLError as e:
        logger.warning(f"[{file_name}] Không thể kết nối tới Ollama tại {base_url} ({e}).")
        return None, "LOI_KET_NOI_OLLAMA"
    except Exception as e:
        logger.warning(f"[{file_name}] Lỗi khi gọi Ollama Local: {e}")
        return None, f"LOI_OLLAMA_{type(e).__name__}"


# ==============================================================================
# QUY TRÌNH XỬ LÝ 1 FILE PDF ĐỘC LẬP
# ==============================================================================
def process_single_pdf(
    file_path: Path,
    key_manager: ApiKeyManager,
    worker_idx: int,
    checkpoint_mgr: CheckpointManager,
    local_mode: bool = False,
    enable_fallback: bool = ENABLE_OLLAMA_FALLBACK,
    ollama_model: str = OLLAMA_MODEL,
    skip_has_state: bool = True,
    skip_ticker_level: bool = False
) -> Dict[str, Any]:
    file_name = file_path.name
    ma_ck, nam = extract_metadata_from_filename(file_path)

    base_record = {
        "ma_ck": ma_ck,
        "nam_bao_cao": nam,
        "file_goc": file_name,
        "duong_dan_file": str(file_path),
        "co_so_huu_nha_nuoc": False,
        "ty_le_so_huu_nha_nuoc_percent": 0.0,
        "so_luong_co_phieu_nha_nuoc": 0,
        "ten_co_dong_nha_nuoc": "Không có",
        "ngay_chot_danh_sach": "N/A",
        "bang_chung_trich_xuat": "",
        "cac_trang_da_doc": "",
        "trang_thai": "CHUA_XU_LY",
        "thoi_gian_xu_ly": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # KIỂM TRA ĐỀ PHÒNG: Nếu luồng khác vừa tìm thấy sở hữu nhà nước cho mã/năm này
    is_done, reason = checkpoint_mgr.check_file(
        file_name, ma_ck, nam, skip_if_has_state=skip_has_state, skip_ticker_level=skip_ticker_level
    )
    if is_done and "sở hữu Nhà nước" in reason:
        logger.info(f"[{file_name}] Bỏ qua vì: {reason}")
        base_record.update({
            "trang_thai": "BO_QUA_DA_CO_SO_HUU_NN",
            "bang_chung_trich_xuat": reason
        })
        return base_record

    # BƯỚC 1: LỚP 1 LOCAL PRE-FILTERING
    is_scan, text_content, selected_pages, total_pages = extract_relevant_pages_local(file_path)

    # NẾU LÀ PDF SCAN: Không bỏ qua làm OCR nữa, mà dùng trực tiếp Gemini Multimodal Vision AI!
    if is_scan:
        logger.info(f"[{file_name}] File scan thuần ảnh ({total_pages} trang). Dùng Gemini Multimodal Vision AI xử lý trực tiếp...")
        base_record["cac_trang_da_doc"] = f"Toàn bộ {min(total_pages, 100)} trang (AI Vision Scan)"
    else:
        base_record["cac_trang_da_doc"] = ", ".join(f"Trang {p}" for p in selected_pages)

    if not is_scan and not text_content.strip():
        base_record.update({
            "trang_thai": "KHONG_TIM_THAY_THONG_TIN",
            "bang_chung_trich_xuat": "Không trích xuất được text liên quan.",
        })
        checkpoint_mgr.append_record(base_record)
        return base_record

    # BƯỚC 2: XỬ LÝ MODEL (GEMINI FLASH HOẶC OLLAMA LOCAL + TESSERACT OCR)
    if is_scan:
        if local_mode:
            logger.info(f"[{file_name}] File scan thuần ảnh. Đang trích xuất qua Tesseract OCR cục bộ...")
            ocr_text, ocr_pages = extract_scan_pages_with_ocr(file_path)
            if ocr_text:
                base_record["cac_trang_da_doc"] = ", ".join(f"Trang {p} (OCR)" for p in ocr_pages)
                prompt = PROMPT_TEMPLATE.format(file_name=file_name, text_content=ocr_text)
                api_result, api_status = call_ollama_local(prompt=prompt, file_name=file_name, model_name=ollama_model)
            else:
                api_result, api_status = None, "KHONG_OCR_DUOC_NOI_DUNG"
        else:
            prompt = SCAN_PROMPT_TEMPLATE.format(file_name=file_name, ma_ck=ma_ck, nam=nam)
            api_result, api_status = call_gemini_with_resilience(
                key_manager=key_manager,
                worker_idx=worker_idx,
                file_name=file_name,
                text_content="",
                pdf_path=file_path,
                is_scan=True,
                ma_ck=ma_ck,
                nam=nam
            )
            # Fallback sang Tesseract OCR + Ollama nếu Gemini không khả dụng
            if not api_result and enable_fallback:
                logger.warning(
                    f"[{file_name}] Gemini Vision không khả dụng ({api_status}). "
                    f"KÍCH HOẠT FALLBACK SANG TESSERACT OCR + OLLAMA LOCAL ({ollama_model})..."
                )
                ocr_text, ocr_pages = extract_scan_pages_with_ocr(file_path)
                if ocr_text:
                    base_record["cac_trang_da_doc"] = ", ".join(f"Trang {p} (OCR)" for p in ocr_pages)
                    prompt_ocr = PROMPT_TEMPLATE.format(file_name=file_name, text_content=ocr_text)
                    api_result, api_status = call_ollama_local(prompt=prompt_ocr, file_name=file_name, model_name=ollama_model)
                else:
                    api_result, api_status = None, "KHONG_OCR_DUOC_NOI_DUNG"
    else:
        prompt = PROMPT_TEMPLATE.format(file_name=file_name, text_content=text_content)
        if local_mode:
            api_result, api_status = call_ollama_local(
                prompt=prompt,
                file_name=file_name,
                model_name=ollama_model
            )
        else:
            api_result, api_status = call_gemini_with_resilience(
                key_manager=key_manager,
                worker_idx=worker_idx,
                file_name=file_name,
                text_content=text_content,
                pdf_path=None,
                is_scan=False,
                ma_ck=ma_ck,
                nam=nam
            )

            # FALLBACK DỰ PHÒNG: Tự động kích hoạt Ollama Local khi Gemini thất bại
            if not api_result and enable_fallback:
                logger.warning(
                    f"[{file_name}] Gemini API không khả dụng ({api_status}). "
                    f"KÍCH HOẠT FALLBACK TỰ ĐỘNG SANG OLLAMA LOCAL ({ollama_model})..."
                )
                api_result, api_status = call_ollama_local(
                    prompt=prompt,
                    file_name=file_name,
                    model_name=ollama_model
                )

    if not api_result:
        base_record.update({
            "trang_thai": api_status,
            "bang_chung_trich_xuat": "Lỗi API sau nhiều lần thử lại (kể cả Fallback Local)."
        })
        checkpoint_mgr.append_record(base_record)
        return base_record

    # BƯỚC 3: CHUẨN HÓA DỮ LIỆU
    ty_le_raw = api_result.get("ty_le_so_huu_nha_nuoc_percent", 0.0)
    try:
        ty_le_val = float(ty_le_raw) if ty_le_raw is not None else 0.0
    except (ValueError, TypeError):
        ty_le_val = 0.0

    so_luong_raw = api_result.get("so_luong_co_phieu_nha_nuoc", 0)
    try:
        so_luong_val = int(float(so_luong_raw)) if so_luong_raw is not None else 0
    except (ValueError, TypeError):
        so_luong_val = 0

    has_state = bool(api_result.get("co_so_huu_nha_nuoc", ty_le_val > 0.0))

    base_record.update({
        "ma_ck": ma_ck,  # LUÔN dùng Mã CK chuẩn từ thư mục gốc, KHÔNG để LLM hallucinate ghi đè
        "nam_bao_cao": nam if nam > 0 else (api_result.get("nam_bao_cao") or nam),
        "co_so_huu_nha_nuoc": has_state,
        "ty_le_so_huu_nha_nuoc_percent": round(ty_le_val, 4),
        "so_luong_co_phieu_nha_nuoc": so_luong_val,
        "ten_co_dong_nha_nuoc": api_result.get("ten_co_dong_nha_nuoc") or ("Không có" if ty_le_val == 0 else "N/A"),
        "ngay_chot_danh_sach": api_result.get("ngay_chot_danh_sach") or "N/A",
        "bang_chung_trich_xuat": api_result.get("bang_chung_trich_xuat") or "",
        "trang_thai": api_result.get("trang_thai") or "THANH_CONG"
    })

    checkpoint_mgr.append_record(base_record)
    return base_record


# ==============================================================================
# XUẤT FILE EXCEL ĐỊNH DẠNG CHUYÊN NGHIỆP
# ==============================================================================
def export_to_excel(records: List[Dict[str, Any]], output_excel_path: str):
    """Xuất danh sách kết quả ra file Excel với định dạng chuyên nghiệp."""
    if not records:
        return

    # Chuẩn hóa mã CK, năm báo cáo và khử trùng lặp file trước khi xuất Excel
    cleaned_records = []
    seen_files = set()
    for r in records:
        rec = dict(r)
        file_name = rec.get("file_goc")
        if file_name and file_name in seen_files:
            continue
        if file_name:
            seen_files.add(file_name)

        file_p = Path(rec.get("duong_dan_file") or rec.get("file_goc") or "")
        canonical_ma, canonical_nam = extract_metadata_from_filename(file_p)
        if canonical_ma != "UNKNOWN":
            rec["ma_ck"] = canonical_ma
        else:
            rec["ma_ck"] = str(rec.get("ma_ck", "")).strip().upper()

        if canonical_nam > 0:
            rec["nam_bao_cao"] = canonical_nam

        cleaned_records.append(rec)

    # Sắp xếp chuẩn xác: 1. Mã CK (A-Z) -> 2. Năm báo cáo (tăng dần) -> 3. Tên file gốc
    sorted_records = sorted(
        cleaned_records,
        key=lambda r: (
            str(r.get("ma_ck", "")).strip().upper(),
            int(r.get("nam_bao_cao", 0) or 0),
            str(r.get("file_goc", ""))
        )
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Sở Hữu Nhà Nước"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
    bold_font = Font(name="Arial", size=10, bold=True)
    state_yes_fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    scan_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )

    headers = [
        "STT",
        "Mã CK",
        "Năm Báo Cáo",
        "Có Sở Hữu Nhà Nước",
        "Tỷ Lệ Sở Hữu (%)",
        "Số Lượng Cổ Phiếu",
        "Tổ Chức Đại Diện Vốn Nhà Nước",
        "Ngày Chốt Danh Sách",
        "Bằng Chứng Trích Xuất",
        "Trạng Thái",
        "Các Trang Đã Đọc",
        "File Gốc",
        "Thời Gian Xử Lý"
    ]

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 28

    for idx, r in enumerate(sorted_records, 1):
        has_state = r.get("co_so_huu_nha_nuoc", False)
        status = r.get("trang_thai", "")
        row_vals = [
            idx,
            r.get("ma_ck", ""),
            r.get("nam_bao_cao", ""),
            "Có" if has_state else "Không",
            r.get("ty_le_so_huu_nha_nuoc_percent", 0.0),
            r.get("so_luong_co_phieu_nha_nuoc", 0),
            r.get("ten_co_dong_nha_nuoc", "Không có"),
            r.get("ngay_chot_danh_sach", "N/A"),
            r.get("bang_chung_trich_xuat", ""),
            status,
            r.get("cac_trang_da_doc", ""),
            r.get("file_goc", ""),
            r.get("thoi_gian_xu_ly", "")
        ]
        ws.append(row_vals)
        current_row = ws.max_row
        ws.row_dimensions[current_row].height = 22

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = data_font
            cell.border = thin_border

            if col_idx in [1, 2, 3, 4, 8, 10, 13]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [5, 6]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if col_idx == 5:
                cell.number_format = "0.00"
            elif col_idx == 6:
                cell.number_format = "#,##0"

            if has_state and col_idx in [4, 5, 7]:
                cell.fill = state_yes_fill
                cell.font = bold_font
            elif status == "THANH_CONG_OLLAMA_LOCAL" and col_idx == 10:
                cell.fill = PatternFill(start_color="D0E1FD", end_color="D0E1FD", fill_type="solid")
            elif status == "PDF_SCAN_CAN_OCR" and col_idx == 10:
                cell.fill = scan_fill

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 50))

    ws.column_dimensions["I"].width = 45  # Bằng chứng trích xuất
    ws.column_dimensions["G"].width = 30  # Tên tổ chức

    out_path = Path(output_excel_path)
    try:
        wb.save(out_path)
        logger.info(f"Đã xuất file Excel: {out_path.resolve()} ({len(sorted_records)} dòng)")
    except PermissionError:
        backup_path = out_path.with_name(f"{out_path.stem}_{int(time.time())}.xlsx")
        wb.save(backup_path)
        logger.warning(f"File {out_path} đang mở, đã lưu ra file backup: {backup_path}")

    try:
        wb.save("Ket_Qua_So_Huu_Nha_Nuoc_606_Chuan.xlsx")
    except Exception:
        pass


# ==============================================================================
# HÀM ĐIỀU PHỐI CHÍNH
# ==============================================================================
def run_pipeline(
    pdf_folder: str = DEFAULT_PDF_FOLDER,
    output_excel: str = DEFAULT_OUTPUT_EXCEL,
    checkpoint_file: str = DEFAULT_CHECKPOINT,
    reference_excel: Optional[str] = REFERENCE_EXCEL_PATH,
    workers: int = DEFAULT_WORKERS,
    limit: Optional[int] = None,
    export_only: bool = False,
    local_mode: bool = False,
    enable_fallback: bool = ENABLE_OLLAMA_FALLBACK,
    ollama_model: str = OLLAMA_MODEL,
    skip_has_state: bool = SKIP_IF_STATE_OWNED,
    skip_ticker_level: bool = False
):
    print("=" * 80)
    print("PIPELINE HYBRID TRÍCH XUẤT TỶ LỆ SỞ HỮU NHÀ NƯỚC (2020-2025)")
    print("=" * 80)

    checkpoint_mgr = CheckpointManager(checkpoint_file, reference_excel_path=reference_excel)

    if export_only:
        print(f"Chế độ xuất Excel từ checkpoint ({len(checkpoint_mgr.get_all_records())} bản ghi)...")
        export_to_excel(checkpoint_mgr.get_all_records(), output_excel)
        print(f"Hoàn thành! Đã lưu tại: {output_excel}")
        return

    key_mgr = ApiKeyManager()
    
    if local_mode:
        num_workers = min(workers, 4)
        print(f"CHẾ ĐỘ: 100% Ollama Local Model ({ollama_model}) | Số luồng song song: {num_workers}")
        if not check_ollama_available(OLLAMA_BASE_URL):
            print(f"⚠️ CẢNH BÁO: Chưa kết nối được Ollama tại {OLLAMA_BASE_URL}!")
            print("Vui lòng mở ứng dụng Ollama hoặc chạy lệnh: ollama serve")
    else:
        num_workers = min(workers, key_mgr.total_keys)
        print(f"CHẾ ĐỘ: Gemini Flash Cloud API ({DEFAULT_MODEL}) ({num_workers} luồng, {key_mgr.total_keys} keys)")
        if enable_fallback:
            ollama_status = "Đang sẵn sàng" if check_ollama_available(OLLAMA_BASE_URL) else "Chưa bật (chạy dự phòng khi cần)"
            print(f"TÍNH NĂNG DỰ PHÒNG: Tự động Fallback sang Ollama Local ({ollama_model}): {ollama_status}")

    pdf_dir = Path(pdf_folder)
    if not pdf_dir.exists():
        print(f"Lỗi: Thư mục PDF không tồn tại: {pdf_dir.resolve()}")
        sys.exit(1)

    all_pdf_files = sorted(list(pdf_dir.rglob("*.pdf")))
    total_files = len(all_pdf_files)
    print(f"Tổng số file PDF tìm thấy: {total_files} file")

    if total_files == 0:
        print("Không có file PDF nào để xử lý.")
        return

    # Lọc file chưa hoàn thành và kiểm tra điều kiện bỏ qua sở hữu nhà nước
    pending_files = []
    skipped_completed = 0
    skipped_has_state = 0

    for f in all_pdf_files:
        ma_ck, nam = extract_metadata_from_filename(f)
        is_done, reason = checkpoint_mgr.check_file(
            file_name=f.name,
            ma_ck=ma_ck,
            nam=nam,
            skip_if_has_state=skip_has_state,
            skip_ticker_level=skip_ticker_level
        )
        if is_done:
            if "sở hữu Nhà nước" in reason:
                skipped_has_state += 1
            else:
                skipped_completed += 1
        else:
            pending_files.append(f)

    print("-" * 80)
    print(f"TIẾN ĐỘ & BỘ LỌC:")
    print(f"  - File đã hoàn thành trong checkpoint: {skipped_completed}/{total_files} file")
    if skip_has_state or skip_ticker_level:
        mode_desc = "Cả Mã CK" if skip_ticker_level else "Cặp (Mã CK, Năm)"
        print(f"  - File BỎ QUA vì ĐÃ CÓ sở hữu Nhà nước ({mode_desc}): {skipped_has_state} file")
    print(f"  - Số file cần xử lý thực tế: {len(pending_files)} file")
    print("-" * 80)

    if limit and limit > 0:
        pending_files = pending_files[:limit]
        print(f"Chế độ giới hạn chạy thử (--limit {limit}): sẽ xử lý {len(pending_files)} file")

    if not pending_files:
        print("Tất cả các file đều đã được xử lý hoàn tất hoặc đã có sở hữu Nhà nước!")
        export_to_excel(checkpoint_mgr.get_all_records(), output_excel)
        return

    processed_in_session = 0

    with ThreadPoolExecutor(max_workers=num_workers) as executor:
        future_to_file = {}
        for idx, pdf_file in enumerate(pending_files):
            worker_id = idx % num_workers
            fut = executor.submit(
                process_single_pdf,
                pdf_file,
                key_mgr,
                worker_id,
                checkpoint_mgr,
                local_mode,
                enable_fallback,
                ollama_model,
                skip_has_state,
                skip_ticker_level
            )
            future_to_file[fut] = pdf_file

        with tqdm(total=len(pending_files), desc="Trích xuất", unit="file", dynamic_ncols=True) as pbar:
            for fut in as_completed(future_to_file):
                pdf_file = future_to_file[fut]
                try:
                    res = fut.result()
                    status = res.get("trang_thai", "UNKNOWN")
                    ticker = res.get("ma_ck", "---")
                    state_pct = res.get("ty_le_so_huu_nha_nuoc_percent", 0.0)

                    pbar.set_postfix({
                        "CK": ticker,
                        "State%": f"{state_pct}%",
                        "TrạngThái": status[:14]
                    })
                except Exception as e:
                    logger.error(f"Lỗi không xác định tại file {pdf_file.name}: {e}")

                pbar.update(1)
                processed_in_session += 1

                if processed_in_session % EXCEL_SAVE_INTERVAL == 0:
                    export_to_excel(checkpoint_mgr.get_all_records(), output_excel)

    print("\nXuất file Excel tổng hợp kết quả...")
    export_to_excel(checkpoint_mgr.get_all_records(), output_excel)
    print("=" * 80)
    print(f"HOÀN THÀNH TOÀN BỘ! Kết quả lưu tại: {Path(output_excel).resolve()}")
    print("=" * 80)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Trích xuất Tỷ lệ sở hữu Nhà nước từ Báo cáo thường niên (2020-2025)")
    parser.add_argument("--folder", type=str, default=DEFAULT_PDF_FOLDER, help="Đường dẫn thư mục chứa PDF")
    parser.add_argument("--excel", type=str, default=DEFAULT_OUTPUT_EXCEL, help="Đường dẫn file Excel đầu ra")
    parser.add_argument("--checkpoint", type=str, default=DEFAULT_CHECKPOINT, help="Đường dẫn file checkpoint JSONL")
    parser.add_argument("--reference-excel", type=str, default=REFERENCE_EXCEL_PATH, help="Đường dẫn file Excel tham chiếu các mã đã có sở hữu Nhà nước")
    parser.add_argument("--workers", type=int, default=DEFAULT_WORKERS, help="Số luồng chạy song song (1-16)")
    parser.add_argument("--limit", type=int, default=None, help="Giới hạn số file cần chạy (dùng khi test thử nghiệm)")
    parser.add_argument("--export-only", action="store_true", help="Chỉ xuất Excel từ file checkpoint hiện có mà không gọi API")
    parser.add_argument("--local", action="store_true", help="Chạy 100%% bằng Ollama Local Model thay vì Gemini Cloud API")
    parser.add_argument("--ollama-model", type=str, default=OLLAMA_MODEL, help="Tên model Ollama local (mặc định: qwen2.5:7b)")
    parser.add_argument("--no-fallback", action="store_true", help="Tắt tính năng tự động fallback sang Ollama khi cạn API")
    parser.add_argument("--skip-has-state", dest="skip_has_state", action="store_true", default=SKIP_IF_STATE_OWNED, help="Bỏ qua các file/cặp (Mã CK, Năm) đã có thông tin sở hữu Nhà nước (mặc định: bật)")
    parser.add_argument("--no-skip-has-state", dest="skip_has_state", action="store_false", help="Không bỏ qua, quét toàn bộ file kể cả mã đã có sở hữu Nhà nước")
    parser.add_argument("--skip-ticker-level", action="store_true", help="Bỏ qua TOÀN BỘ các năm của Mã CK nếu doanh nghiệp đó đã từng có sở hữu Nhà nước")

    args = parser.parse_args()

    run_pipeline(
        pdf_folder=args.folder,
        output_excel=args.excel,
        checkpoint_file=args.checkpoint,
        reference_excel=args.reference_excel,
        workers=args.workers,
        limit=args.limit,
        export_only=args.export_only,
        local_mode=args.local,
        enable_fallback=not args.no_fallback,
        ollama_model=args.ollama_model,
        skip_has_state=args.skip_has_state,
        skip_ticker_level=args.skip_ticker_level
    )
