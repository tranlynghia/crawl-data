"""
ownership_core.ocr_engine
=========================
Bộ máy OCR từng trang/vùng khi lớp chữ thiếu hoặc hỏng.
Hỗ trợ vie+eng (hoặc fallback eng/vie).
Tích hợp Cache SHA-256 hình ảnh tránh quét lại nhiều lần.
"""

from __future__ import annotations
import hashlib
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Tuple, Set
import fitz
from ownership_core.config import (
    TESSERACT_EXE_WINDOWS,
    TESSDATA_LOCAL,
    DEFAULT_OCR_DPI,
    OCR_LANGUAGES,
)
from ownership_core.cache_manager import CacheManager


def find_tesseract_cmd() -> Optional[str]:
    """Tìm đường dẫn tesseract thực thi."""
    if TESSERACT_EXE_WINDOWS.exists():
        return str(TESSERACT_EXE_WINDOWS)
    which = shutil.which("tesseract")
    if which:
        return which
    return None


def get_available_languages(tess_cmd: str, tessdata_dir: Optional[Path] = None) -> Set[str]:
    """Kiểm tra các ngôn ngữ OCR có sẵn."""
    env = os.environ.copy()
    if tessdata_dir and tessdata_dir.exists():
        env["TESSDATA_PREFIX"] = str(tessdata_dir)
    try:
        res = subprocess.run(
            [tess_cmd, "--list-langs"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            env=env,
            timeout=5.0
        )
        if res.returncode == 0:
            langs = set(res.stdout.splitlines()[1:])
            return {l.strip() for l in langs if l.strip()}
    except Exception:
        pass
    return {"eng"}


class OCREngine:
    """Động cơ OCR trang PDF tích hợp bộ đệm (cache)."""

    def __init__(self, cache_manager: Optional[CacheManager] = None):
        self.cache_manager = cache_manager or CacheManager()
        self.tess_cmd = find_tesseract_cmd()
        self.tessdata_dir = TESSDATA_LOCAL if TESSDATA_LOCAL.exists() else None
        
        # Xác định ngôn ngữ OCR tốt nhất có thể dùng
        self.lang = "eng"
        if self.tess_cmd:
            available = get_available_languages(self.tess_cmd, self.tessdata_dir)
            if "vie" in available and "eng" in available:
                self.lang = "vie+eng"
            elif "vie" in available:
                self.lang = "vie"
            elif "eng" in available:
                self.lang = "eng"

    @property
    def is_available(self) -> bool:
        return self.tess_cmd is not None

    def ocr_page(self, page: fitz.Page, dpi: int = DEFAULT_OCR_DPI) -> str:
        """
        OCR một trang PDF đơn lẻ:
        1. Render ra ảnh pixmap.
        2. Tính SHA-256 ảnh. Kiểm tra cache.
        3. Nếu chưa có trong cache: Gọi tesseract và lưu cache.
        """
        if not self.is_available:
            return ""

        # 1. Render ảnh
        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")
        img_hash = hashlib.sha256(img_bytes).hexdigest()

        # 2. Kiểm tra cache
        cached = self.cache_manager.get_ocr_text(img_hash, self.lang)
        if cached is not None:
            return cached

        # 3. Chạy Tesseract OCR qua file tạm thời
        ocr_result = ""
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        
        try:
            tmp_path.write_bytes(img_bytes)
            env = os.environ.copy()
            if self.tessdata_dir:
                env["TESSDATA_PREFIX"] = str(self.tessdata_dir)

            cmd = [self.tess_cmd, str(tmp_path), "stdout", "-l", self.lang]
            res = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                env=env,
                timeout=30.0
            )
            if res.returncode == 0:
                ocr_result = (res.stdout or "").strip()
        except Exception:
            ocr_result = ""
        finally:
            tmp_path.unlink(missing_ok=True)

        # 4. Lưu cache
        self.cache_manager.set_ocr_text(img_hash, self.lang, ocr_result)
        return ocr_result
