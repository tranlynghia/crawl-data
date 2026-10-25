"""
ownership_core.pipeline
=======================
Bộ điều phối toàn diện quy trình kiểm kê -> quét PDF -> trích xuất -> kiểm chứng -> xuất Excel:
- Hỗ trợ multiprocessing an toàn trên Windows.
- Tích hợp Cache phân tầng (Text, OCR, Candidate, Validator).
- Mã hóa luồng tuần tự khi gọi Ollama để tránh tràn VRAM.
- Không bao giờ coi kết quả cũ chưa kiểm chứng là đã xác minh.
"""

from __future__ import annotations
import json
import logging
import multiprocessing as mp
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import fitz

from ownership_core.config import (
    DEFAULT_PDF_DIR,
    DEFAULT_MANIFEST,
    DEFAULT_CHECKPOINT,
    DEFAULT_EXCEL_OUT,
    DEFAULT_CACHE_DIR,
    OLLAMA_MODEL,
    OLLAMA_BASE_URL,
)
from ownership_core.inventory import InventoryItem, load_canonical_manifest
from ownership_core.cache_manager import CacheManager, get_file_sha256
from ownership_core.ocr_engine import OCREngine
from ownership_core.table_extractor import TableExtractor, StructuredTable
from ownership_core.rule_extractor import extract_candidates_from_tables
from ownership_core.validator import (
    CandidateOwnership,
    ValidationResult,
    VerificationEngine,
)
from ownership_core.model_client import LocalOllamaClient
from ownership_core.exporter import export_single_sheet_excel

logger = logging.getLogger("Pipeline")


def process_single_pdf_worker(
    item_dict: Dict,
    use_local_model: bool = False,
    cache_dir_str: str = str(DEFAULT_CACHE_DIR)
) -> Tuple[str, int, Dict, List[Dict]]:
    """
    Worker xử lý một file PDF độc lập trong multiprocessing:
    1. Kiểm tra file trên đĩa.
    2. Mở file riêng trong process (tránh xung đột GIL/thread của PyMuPDF).
    3. Quét toàn bộ trang, bóc tách bảng biểu.
    4. Trích xuất ứng viên thuần code.
    5. Chạy bộ kiểm chứng.
    6. Nếu chưa xác minh và có cờ model: gọi Local Ollama.
    7. Chạy bộ kiểm chứng lần cuối và trả về kết quả.
    """
    ticker = item_dict["ticker"]
    year = item_dict["year"]
    file_path_str = item_dict.get("file_path")

    if not file_path_str or not Path(file_path_str).exists():
        val_res = ValidationResult(
            is_verified=False,
            status="THIEU_FILE_TREN_DIA",
            ratio=None,
            share_count=None,
            holder_name="Thiếu file PDF",
            cutoff_date="N/A",
            basis="Không tìm thấy file trên ổ đĩa",
            evidence="",
            page_number=None,
            notes=item_dict.get("notes", "")
        )
        return ticker, year, val_res.to_dict(), []

    pdf_p = Path(file_path_str)
    file_hash = get_file_sha256(pdf_p)
    cache_mgr = CacheManager(Path(cache_dir_str))
    ocr_eng = OCREngine(cache_mgr)
    table_ext = TableExtractor(cache_mgr, ocr_eng)
    validator = VerificationEngine(target_month_day="12-31")

    all_candidates: List[CandidateOwnership] = []
    doc_text_corpus = ""

    try:
        doc = fitz.open(str(pdf_p))
        # 1. Quét tất cả các trang
        scored_pages = table_ext.scan_all_pages(doc, file_hash)
        doc_text_corpus = "\n".join(p.text for p in scored_pages[:15])

        # 2. Bóc tách bảng có cấu trúc
        structured_tables = table_ext.extract_structured_tables(doc, scored_pages, year)

        # 3. Trích xuất ứng viên thuần code từ các bảng biểu rõ ràng
        rule_cands = extract_candidates_from_tables(ticker, year, structured_tables)
        all_candidates.extend(rule_cands)

        # 4. Kiểm chứng các ứng viên thuần code
        val_res = validator.evaluate_candidates(ticker, year, all_candidates, doc_text_corpus)

        # 5. Nếu chưa xác minh được và cho phép gọi model:
        # Gọi Local Ollama trên các trang ứng viên bảng biểu
        if not val_res.is_verified and use_local_model and structured_tables:
            ollama = LocalOllamaClient()
            if ollama.is_available():
                # Chuẩn bị văn bản bảng markdown
                table_md_text = "\n\n".join(t.raw_markdown for t in structured_tables[:3])
                model_cands = ollama.extract_candidates(ticker, year, pdf_p.name, table_md_text)
                if model_cands:
                    all_candidates.extend(model_cands)
                    # Chạy lại bộ kiểm chứng trên toàn bộ ứng viên
                    val_res = validator.evaluate_candidates(ticker, year, all_candidates, doc_text_corpus)

        doc.close()
    except Exception as e:
        logger.error(f"[{ticker}_{year}] Lỗi xử lý file {pdf_p.name}: {e}")
        val_res = ValidationResult(
            is_verified=False,
            status="LOI_XU_LY_FILE",
            ratio=None,
            share_count=None,
            holder_name="Lỗi xử lý",
            cutoff_date="N/A",
            basis=f"Lỗi: {e}",
            evidence="",
            page_number=None,
            notes=str(e)
        )

    return ticker, year, val_res.to_dict(), [c.to_dict() for c in all_candidates]


class OwnershipPipeline:
    """Bộ điều phối chạy toàn bộ hoặc chạy mẫu pipeline kiểm chứng tỷ lệ sở hữu Nhà nước."""

    def __init__(
        self,
        manifest_path: Path = DEFAULT_MANIFEST,
        pdf_dir: Path = DEFAULT_PDF_DIR,
        checkpoint_path: Path = DEFAULT_CHECKPOINT,
        output_excel: Path = DEFAULT_EXCEL_OUT,
        cache_dir: Path = DEFAULT_CACHE_DIR,
        num_workers: int = 2
    ):
        self.manifest_path = manifest_path
        self.pdf_dir = pdf_dir
        self.checkpoint_path = checkpoint_path
        self.output_excel = output_excel
        self.cache_dir = cache_dir
        self.num_workers = max(1, min(num_workers, 4))
        self.results_map: Dict[Tuple[str, int], ValidationResult] = {}
        self._load_checkpoint()

    def _load_checkpoint(self):
        """Đọc kết quả checkpoint đã có."""
        if self.checkpoint_path.exists():
            try:
                with open(self.checkpoint_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        r = json.loads(line)
                        pair = (r.get("ticker"), int(r.get("year", 0)))
                        val_dict = r.get("validation_result") or {}
                        if pair[0] and pair[1] > 0 and val_dict:
                            self.results_map[pair] = ValidationResult(**val_dict)
            except Exception as e:
                logger.warning(f"Lỗi đọc checkpoint: {e}")

    def _save_checkpoint_record(self, ticker: str, year: int, val_res: ValidationResult, candidates: List[Dict]):
        """Ghi kết quả vào checkpoint JSONL ngay lập tức."""
        rec = {
            "ticker": ticker,
            "year": year,
            "validation_result": val_res.to_dict(),
            "candidates": candidates,
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(self.checkpoint_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            f.flush()

    def run(
        self,
        sample_count: Optional[int] = None,
        use_local_model: bool = False,
        retry_unverified: bool = False
    ) -> Dict[str, Any]:
        """Thực thi pipeline."""
        inventory = load_canonical_manifest(self.manifest_path, self.pdf_dir)
        total_items = len(inventory)

        # Lọc danh sách cần xử lý
        tasks: List[InventoryItem] = []
        for item in inventory:
            pair = (item.ticker, item.year)
            if pair in self.results_map:
                existing = self.results_map[pair]
                # Nếu đã xác minh thành công và không yêu cầu retry_unverified -> bỏ qua
                if existing.is_verified and not retry_unverified:
                    continue
                # Nếu chưa xác minh nhưng không bật retry_unverified -> bỏ qua
                if not retry_unverified:
                    continue
            tasks.append(item)

        if sample_count and sample_count > 0:
            tasks = tasks[:sample_count]

        logger.info(f"Bắt đầu xử lý {len(tasks)}/{total_items} mục (Workers: {self.num_workers}, Model Local: {use_local_model})")

        start_time = time.time()
        processed_count = 0
        verified_count = 0

        # Nếu có task cần chạy
        if tasks:
            # Chuẩn bị input cho multiprocessing
            worker_inputs = [
                (task.to_dict(), use_local_model, str(self.cache_dir))
                for task in tasks
            ]

            if self.num_workers > 1:
                with mp.Pool(processes=self.num_workers) as pool:
                    for ticker, year, val_dict, cands in pool.starmap(process_single_pdf_worker, worker_inputs):
                        val_res = ValidationResult(**val_dict)
                        self.results_map[(ticker, year)] = val_res
                        self._save_checkpoint_record(ticker, year, val_res, cands)
                        processed_count += 1
                        if val_res.is_verified:
                            verified_count += 1
            else:
                for inp in worker_inputs:
                    ticker, year, val_dict, cands = process_single_pdf_worker(*inp)
                    val_res = ValidationResult(**val_dict)
                    self.results_map[(ticker, year)] = val_res
                    self._save_checkpoint_record(ticker, year, val_res, cands)
                    processed_count += 1
                    if val_res.is_verified:
                        verified_count += 1

        elapsed = time.time() - start_time

        # Xuất file Excel duy nhất 1 sheet chuẩn 606 dòng
        excel_saved = export_single_sheet_excel(inventory, self.results_map, self.output_excel)

        # Thống kê
        verified_total = sum(1 for v in self.results_map.values() if v.is_verified)
        state_owned_count = sum(1 for v in self.results_map.values() if v.status == "DA_XAC_MINH_CO_SHNN")
        non_state_count = sum(1 for v in self.results_map.values() if v.status == "DA_XAC_MINH_KHONG_CO_SHNN")
        unverified_total = sum(1 for v in self.results_map.values() if not v.is_verified)

        return {
            "total_items": total_items,
            "processed_in_run": processed_count,
            "verified_in_run": verified_count,
            "total_verified": verified_total,
            "total_state_owned": state_owned_count,
            "total_non_state": non_state_count,
            "total_unverified": unverified_total,
            "elapsed_seconds": round(elapsed, 2),
            "excel_path": str(self.output_excel),
            "excel_saved": excel_saved
        }
