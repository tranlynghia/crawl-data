"""
ownership_core.model_client
===========================
Giao tiếp mô hình ngôn ngữ (Local Ollama & Gemini API) có kiểm soát:
- Chỉ gọi khi bóc tách thuần code chưa xác minh được.
- Prompt CHUẨN MỰC: Tuyệt đối không chứa số mẫu (như 43.31 hay 56626237).
- Bắt buộc trả về JSON Schema danh sách ứng viên kèm bằng chứng trích xuất.
- Local model chạy tuần tự (concurrency = 1) để bảo vệ VRAM GPU/RAM.
"""

from __future__ import annotations
import json
import logging
import threading
import time
import urllib.request
import urllib.error
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from ownership_core.config import (
    OLLAMA_BASE_URL,
    OLLAMA_MODEL,
    GEMINI_MODEL,
    STRICT_CANDIDATE_EXTRACTION_PROMPT,
)
from ownership_core.validator import (
    CandidateOwnership,
    classify_holder_string,
    parse_vietnamese_number,
)

logger = logging.getLogger("ModelClient")
_ollama_lock = threading.Lock()


def clean_json_response(raw_text: str) -> Optional[Dict[str, Any]]:
    """Làm sạch và bóc tách an toàn chuỗi JSON trả về từ LLM."""
    if not raw_text:
        return None
    s = raw_text.strip()
    if s.startswith("```json"):
        s = s[7:]
    elif s.startswith("```"):
        s = s[3:]
    if s.endswith("```"):
        s = s[:-3]
    s = s.strip()

    try:
        return json.loads(s)
    except json.JSONDecodeError:
        # Fallback tìm khối JSON đầu tiên
        start = s.find("{")
        end = s.rfind("}")
        if start != -1 and end != -1 and end > start:
            try:
                return json.loads(s[start:end + 1])
            except Exception:
                pass
    return None


class LocalOllamaClient:
    """Gọi mô hình cục bộ qua Ollama API (qwen2.5-coder:7b)."""

    def __init__(self, base_url: str = OLLAMA_BASE_URL, model: str = OLLAMA_MODEL):
        self.base_url = base_url.rstrip("/")
        self.model = model

    def is_available(self) -> bool:
        try:
            req = urllib.request.Request(f"{self.base_url}/api/tags", method="GET")
            with urllib.request.urlopen(req, timeout=1.5) as resp:
                return resp.status == 200
        except Exception:
            return False

    def extract_candidates(
        self,
        ticker: str,
        report_year: int,
        file_name: str,
        source_text: str
    ) -> List[CandidateOwnership]:
        """Gửi văn bản bảng biểu cho Local Ollama và trích xuất danh sách ứng viên."""
        prompt = STRICT_CANDIDATE_EXTRACTION_PROMPT.format(
            file_name=file_name,
            ticker=ticker,
            year=report_year,
            source_text=source_text[:6000]
        )

        payload = {
            "model": self.model,
            "prompt": prompt,
            "stream": False,
            "format": "json",
            "options": {
                "temperature": 0.0,
                "num_ctx": 8192
            }
        }

        # Khóa tuần tự 1 request tại một thời điểm
        with _ollama_lock:
            try:
                req = urllib.request.Request(
                    f"{self.base_url}/api/generate",
                    data=json.dumps(payload).encode("utf-8"),
                    headers={"Content-Type": "application/json"},
                    method="POST"
                )
                with urllib.request.urlopen(req, timeout=60.0) as resp:
                    if resp.status == 200:
                        res_data = json.loads(resp.read().decode("utf-8"))
                        raw_json_str = res_data.get("response", "")
                        parsed = clean_json_response(raw_json_str)
                        if parsed:
                            return self._parse_candidates_from_json(parsed, ticker, report_year)
            except Exception as e:
                logger.warning(f"[{ticker}_{report_year}] Lỗi gọi Ollama: {e}")

        return []

    def _parse_candidates_from_json(
        self,
        parsed_json: Dict[str, Any],
        ticker: str,
        report_year: int
    ) -> List[CandidateOwnership]:
        cands: List[CandidateOwnership] = []
        raw_list = parsed_json.get("candidates") or []
        for item in raw_list:
            h_name = str(item.get("holder_name") or item.get("row_label") or "").strip()
            if not h_name:
                continue

            raw_p = str(item.get("raw_percentage") or item.get("percentage") or "")
            pct_val, _ = parse_vietnamese_number(raw_p)

            share_cnt = item.get("share_count")
            if share_cnt is not None:
                try:
                    share_cnt = int(share_cnt)
                except (ValueError, TypeError):
                    share_cnt = None

            cand = CandidateOwnership(
                ticker=ticker,
                report_year=report_year,
                holder_name=h_name,
                classification=classify_holder_string(h_name),
                raw_percentage=raw_p,
                percentage=pct_val,
                share_count=share_cnt,
                raw_date=item.get("raw_date"),
                normalized_date=item.get("normalized_date"),
                page_number=item.get("page_number"),
                column_name=item.get("column_label"),
                row_label=item.get("row_label"),
                evidence=str(item.get("evidence") or ""),
                source_method="MODEL_LOCAL"
            )
            cands.append(cand)
        return cands
