"""
ownership_core.cache_manager
============================
Hệ thống Cache phân tầng lưu bằng SQLite (WAL mode):
- Tách bạch: Text layer, OCR, Candidate extraction, Verification.
- Cache gắn với SHA-256 của file/ảnh. Sửa prompt/model chỉ vô hiệu hóa phần cần trích lại,
  tuyệt đối không phải đọc lại PDF hay chạy lại OCR 8 GB dữ liệu.
"""

from __future__ import annotations
import hashlib
import json
import sqlite3
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from ownership_core.config import DEFAULT_CACHE_DIR


def get_file_sha256(path: Path) -> str:
    """Tính mã SHA-256 nội dung file với buffer 64KB."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


class CacheManager:
    """Quản lý cache đọc PDF, OCR và Trích xuất phân tầng an toàn."""

    def __init__(self, cache_dir: Path = DEFAULT_CACHE_DIR):
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.cache_dir / "ownership_cache.sqlite3"
        self._local = threading.local()
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        if not hasattr(self._local, "conn") or self._local.conn is None:
            conn = sqlite3.connect(str(self.db_path), timeout=30.0)
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.execute("PRAGMA synchronous = NORMAL;")
            self._local.conn = conn
        return self._local.conn

    def _init_db(self):
        conn = self._get_conn()
        with conn:
            # 1. Bảng lưu lớp chữ PDF từng trang
            conn.execute("""
                CREATE TABLE IF NOT EXISTS page_text_cache (
                    file_hash TEXT,
                    page_no INTEGER,
                    text TEXT,
                    char_count INTEGER,
                    PRIMARY KEY (file_hash, page_no)
                )
            """)

            # 2. Bảng lưu kết quả OCR từng trang ảnh
            conn.execute("""
                CREATE TABLE IF NOT EXISTS page_ocr_cache (
                    image_hash TEXT,
                    lang TEXT,
                    ocr_text TEXT,
                    PRIMARY KEY (image_hash, lang)
                )
            """)

            # 3. Bảng lưu ứng viên đã trích xuất từ văn bản/ảnh
            conn.execute("""
                CREATE TABLE IF NOT EXISTS candidate_cache (
                    file_hash TEXT,
                    extractor_version TEXT,
                    prompt_hash TEXT,
                    candidates_json TEXT,
                    PRIMARY KEY (file_hash, extractor_version, prompt_hash)
                )
            """)

            # 4. Bảng lưu kết quả kiểm chứng (Verification)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS verification_cache (
                    file_hash TEXT,
                    validator_version TEXT,
                    result_json TEXT,
                    PRIMARY KEY (file_hash, validator_version)
                )
            """)

    # --- TEXT CACHE ---
    def get_page_text(self, file_hash: str, page_no: int) -> Optional[str]:
        conn = self._get_conn()
        cur = conn.cursor()
        cur.execute("SELECT text FROM page_text_cache WHERE file_hash = ? AND page_no = ?", (file_hash, page_no))
        row = cur.fetchone()
        return row[0] if row else None

    def set_page_text(self, file_hash: str, page_no: int, text: str):
        conn = self._get_conn()
        with conn:
            conn.execute(
                "INSERT OR REPLACE INTO page_text_cache (file_hash, page_no, text, char_count) VALUES (?, ?, ?, ?)",
                (file_hash, page_no, text, len(text.strip()))
            )

    # --- OCR CACHE ---
    def get_ocr_text(self, image_hash: str, lang: str = "vie+eng") -> Optional[str]:
        conn = self._get_conn()
        cur = conn.cursor()
        cur.execute("SELECT ocr_text FROM page_ocr_cache WHERE image_hash = ? AND lang = ?", (image_hash, lang))
        row = cur.fetchone()
        return row[0] if row else None

    def set_ocr_text(self, image_hash: str, lang: str, ocr_text: str):
        conn = self._get_conn()
        with conn:
            conn.execute(
                "INSERT OR REPLACE INTO page_ocr_cache (image_hash, lang, ocr_text) VALUES (?, ?, ?)",
                (image_hash, lang, ocr_text)
            )

    # --- CANDIDATE CACHE ---
    def get_candidates(self, file_hash: str, extractor_version: str, prompt_hash: str) -> Optional[List[Dict[str, Any]]]:
        conn = self._get_conn()
        cur = conn.cursor()
        cur.execute(
            "SELECT candidates_json FROM candidate_cache WHERE file_hash = ? AND extractor_version = ? AND prompt_hash = ?",
            (file_hash, extractor_version, prompt_hash)
        )
        row = cur.fetchone()
        if row and row[0]:
            try:
                return json.loads(row[0])
            except Exception:
                return None
        return None

    def set_candidates(self, file_hash: str, extractor_version: str, prompt_hash: str, candidates: List[Dict[str, Any]]):
        conn = self._get_conn()
        with conn:
            conn.execute(
                "INSERT OR REPLACE INTO candidate_cache (file_hash, extractor_version, prompt_hash, candidates_json) VALUES (?, ?, ?, ?)",
                (file_hash, extractor_version, prompt_hash, json.dumps(candidates, ensure_ascii=False))
            )
