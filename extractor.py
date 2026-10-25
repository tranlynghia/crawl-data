#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
========================================================================================
PIPELINE TRÍCH XUẤT & KIỂM CHỨNG TỶ LỆ SỞ HỮU NHÀ NƯỚC (2020-2025)
========================================================================================
Phiên bản 2.0 (High-Precision Verifiable Pipeline):
- Kiểm kê 101 mã x 6 năm = đúng 606 mục; đối chiếu thực tế file trên đĩa.
- Quét toàn bộ trang PDF, trích xuất bảng biểu giữ nguyên cấu trúc không gian (tọa độ).
- Bóc tách thuần Code trước: Phân biệt tỷ lệ cổ phần vs tỷ lệ số lượng cổ đông.
- Bộ kiểm chứng Grounding bắt buộc: Loại bỏ 100% ảo giác (hallucination).
- Chỉ gọi Local Ollama / Gemini khi bố cục bảng chưa rõ ràng; không chứa số mẫu cố định.
- Xuất 1 sheet duy nhất 'Co ty le so huu NN' với đúng 606 dòng dữ liệu chuẩn mực.
========================================================================================
"""

import argparse
import logging
import sys
from pathlib import Path

# Hỗ trợ in UTF-8 tiếng Việt trên console Windows
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from ownership_core.config import (
    DEFAULT_MANIFEST,
    DEFAULT_PDF_DIR,
    DEFAULT_CHECKPOINT,
    DEFAULT_EXCEL_OUT,
    DEFAULT_CACHE_DIR,
)
from ownership_core.pipeline import OwnershipPipeline
from ownership_core.inventory import load_canonical_manifest
from ownership_core.exporter import export_single_sheet_excel

# Cấu hình logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("extractor.log", encoding="utf-8")
    ]
)
logger = logging.getLogger("ExtractorCLI")


def main():
    parser = argparse.ArgumentParser(
        description="Pipeline trích xuất và kiểm chứng tỷ lệ sở hữu Nhà nước từ BCTN (101 mã x 6 năm)"
    )
    parser.add_argument("--sample", type=int, default=None, help="Chạy thử mẫu N file đầu tiên (ví dụ: --sample 25)")
    parser.add_argument("--workers", type=int, default=2, help="Số process chạy song song (mặc định: 2)")
    parser.add_argument("--local", action="store_true", help="Bật Local Ollama model cho bảng chưa rõ bố cục")
    parser.add_argument("--no-model", action="store_true", help="Chỉ dùng bóc tách thuần Code, không gọi LLM")
    parser.add_argument("--retry-unverified", action="store_true", help="Chạy lại các mục chưa xác minh được")
    parser.add_argument("--export-only", action="store_true", help="Chỉ xuất lại file Excel từ checkpoint hiện có")
    parser.add_argument("--manifest", type=str, default=str(DEFAULT_MANIFEST), help="Đường dẫn file manifest CSV")
    parser.add_argument("--pdf-dir", type=str, default=str(DEFAULT_PDF_DIR), help="Thư mục chứa các file PDF BCTN")
    parser.add_argument("--checkpoint", type=str, default=str(DEFAULT_CHECKPOINT), help="File lưu checkpoint kết quả")
    parser.add_argument("--output", type=str, default=str(DEFAULT_EXCEL_OUT), help="File Excel đầu ra")

    args = parser.parse_args()

    pipeline = OwnershipPipeline(
        manifest_path=Path(args.manifest),
        pdf_dir=Path(args.pdf_dir),
        checkpoint_path=Path(args.checkpoint),
        output_excel=Path(args.output),
        num_workers=args.workers
    )

    if args.export_only:
        logger.info(f"Đang xuất lại file Excel chuẩn từ checkpoint: {args.checkpoint}...")
        inventory = load_canonical_manifest(Path(args.manifest), Path(args.pdf_dir))
        ok = export_single_sheet_excel(inventory, pipeline.results_map, Path(args.output))
        if ok:
            logger.info(f"Đã xuất thành công: {args.output} (Sheet: 'Co ty le so huu NN')")
        else:
            logger.warning(f"File {args.output} bị khóa, đã lưu bản dự phòng latest.xlsx")
        return

    # Chạy quy trình
    use_model = args.local and not args.no_model
    stats = pipeline.run(
        sample_count=args.sample,
        use_local_model=use_model,
        retry_unverified=args.retry_unverified
    )

    print("\n" + "=" * 80)
    print("TỔNG KẾT KẾT QUẢ QUY TRÌNH KIỂM CHỨNG SỞ HỮU NHÀ NƯỚC (2020-2025)")
    print("=" * 80)
    print(f"- Tổng số mục danh mục chuẩn:   {stats['total_items']} (101 mã x 6 năm)")
    print(f"- Số file xử lý trong lượt này: {stats['processed_in_run']}")
    print(f"- Tổng số mục đã xác minh:      {stats['total_verified']}/{stats['total_items']}")
    print(f"  + Có sở hữu Nhà nước:         {stats['total_state_owned']}")
    print(f"  + 100% Tư nhân / Nước ngoài:   {stats['total_non_state']}")
    print(f"- Số mục cần kiểm tra / chưa ra:{stats['total_unverified']}")
    print(f"- Thời gian thực thi:           {stats['elapsed_seconds']} giây")
    print(f"- File Excel kết quả:           {stats['excel_path']}")
    print("=" * 80 + "\n")


if __name__ == "__main__":
    main()
