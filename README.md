# Hệ Thống Trích Xuất & Kiểm Chứng Tỷ Lệ Sở Hữu Nhà Nước (2020–2025)

Pipeline chuẩn mực cao (High-Precision Verifiable Pipeline) xử lý **606 mục báo cáo thường niên (101 mã chứng khoán × 6 năm)** với bằng chứng minh bạch, kiểm tra đối chiếu tự động và loại bỏ 100% ảo giác (hallucination).

---

## 1. Yêu Cầu & Cài Đặt

Yêu cầu môi trường Python 3.10+ trên Windows. Cài đặt các thư viện cần thiết:

```powershell
pip install -r requirements.txt
```

*(Tùy chọn OCR)*: Cài đặt [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki) tại `C:\Program Files\Tesseract-OCR\tesseract.exe`. Động cơ OCR sẽ tự động nạp dữ liệu ngôn ngữ tiếng Việt và tiếng Anh từ `ownership_cli/tessdata`.

---

## 2. Kiến Trúc Luồng Xử Lý Thống Nhất

Luồng xử lý chính thống nhất chạy qua script trung tâm: **`extractor.py`** (hoặc `ownership_cli/ownership_101x6.py` - cả hai script đều gọi chung engine `ownership_core`):

1. **Kiểm kê (Inventory)**: Đọc `danh_muc_pdf_101x6.csv`, đối chiếu thực tế file trong thư mục `pdf_101x6/`, tự động chọn bản BCTN chính thức nếu có nhiều file đính kèm.
2. **Quét cấu trúc PDF (Spatial Table Scanner)**: Quét toàn bộ trang, sử dụng tọa độ từ (coordinate grid) và `fitz.find_tables()` để bảo tồn cấu trúc hàng, cột, tiêu đề nhiều tầng và ngày gắn với bảng.
3. **Bóc tách thuần Code (Code-First Extractor)**: Tự động phân tích các bảng rõ ràng, phân biệt chính xác:
   - Cột tỷ lệ sở hữu vốn/cổ phần (Shareholding %) với Cột tỷ lệ số lượng cổ đông (Shareholder count %).
   - Số lượng cổ phiếu với giá trị vốn góp (VNĐ), mã số thuế, địa chỉ.
   - Dòng tổng Nhà nước với phần vốn do cá nhân HĐQT đại diện.
4. **Bộ kiểm chứng bắt buộc (Deterministic Verification Engine)**:
   - **Grounding**: Bắt buộc mọi con số phải xuất hiện thực tế tại đúng ô/dòng nguồn trong tài liệu.
   - **Không biến null/dash thành 0**: Chỉ ghi nhận 0.0% khi bảng ghi rõ 0% hoặc xác nhận 100% tư nhân/nước ngoài.
   - **Mốc mục tiêu 31/12**: Nếu ngày chốt là ngày khác (ví dụ: tháng 4 ĐHCĐ), hệ thống vẫn giữ ứng viên nhưng đánh dấu `CAN_KIEM_TRA_NGAY_KHAC` chứ không tự tiện gán thành 31/12.
5. **Mô hình hỗ trợ (Ollama Local / Gemini API)**: Chỉ gọi khi bảng chưa rõ bố cục. Prompt thuần túy, tuyệt đối không chứa số mẫu cố định (không lặp 43.31 hay 56626237).
6. **Đầu ra duy nhất (Single-Sheet Excel)**: Xuất file `Ket_Qua_So_Huu_Nha_Nuoc.xlsx` với **đúng 1 sheet duy nhất `Co ty le so huu NN`**, gồm đúng 606 dòng dữ liệu cho 101 mã.

---

## 3. Các Lệnh PowerShell Thường Dùng (Copy & Paste)

### A. Kiểm kê dữ liệu thực tế trên đĩa
Kiểm tra đối chiếu 606 mục giữa manifest CSV và file PDF thực tế trong thư mục:
```powershell
python -c "from ownership_core.inventory import load_canonical_manifest; inv = load_canonical_manifest(); print(f'Tong: {len(inv)} | Co file: {sum(1 for i in inv if i.is_disk_available)} | Thieu file: {sum(1 for i in inv if not i.is_disk_available)}')"
```

### B. Chạy thử nghiệm một số file (Sample Benchmark)
Chạy thử mẫu 20 báo cáo đa dạng (chạy thuần code, siêu nhanh, không tốn token/API):
```powershell
python extractor.py --sample 20
```

### C. Chạy toàn bộ 606 file (Thuần Code + Kiểm chứng Grounding)
Chạy kiểm chứng tự động toàn bộ 606 file với 2 process song song (an toàn Windows):
```powershell
python extractor.py --workers 2
```

### D. Bật thêm Model Local (Ollama) cho các bảng phức tạp
Nếu máy có cài Ollama (`qwen2.5-coder:7b`) và muốn hỗ trợ các bảng khó phân tích:
```powershell
python extractor.py --workers 2 --local
```

### E. Thử lại các mục chưa xác minh (Retry Unverified)
Bỏ qua các mục đã xác minh thành công (`DA_XAC_MINH...`), chỉ quét lại các mục cần kiểm tra hoặc chưa tìm thấy:
```powershell
python extractor.py --retry-unverified --local
```

### F. Xuất lại file Excel chuẩn từ Checkpoint hiện có
Tái tạo file `Ket_Qua_So_Huu_Nha_Nuoc.xlsx` chuẩn 606 dòng từ checkpoint bất cứ lúc nào:
```powershell
python extractor.py --export-only
```

### G. Chạy toàn bộ bộ kiểm thử hồi quy (Test Suite)
Chạy 19 bài kiểm thử đơn vị và hồi quy (kiểm tra grounding, chống ảo giác, phân biệt cột cổ đông vs cổ phần):
```powershell
python -m pytest tests/test_regression_suite.py ownership_cli/tests/test_core.py
```

---

## 4. Cấu Trúc Bảng Tính Đầu Ra (`Ket_Qua_So_Huu_Nha_Nuoc.xlsx`)

Sheet duy nhất: **`Co ty le so huu NN`** gồm đúng 606 dòng dữ liệu được sắp xếp theo Mã CK (A-Z) và Năm (2020-2025):
- `STT`: Số thứ tự từ 1 đến 606.
- `Mã CK`: Mã chứng khoán 3 ký tự.
- `Tên công ty`: Tên đầy đủ của doanh nghiệp.
- `Năm Báo Cáo`: 2020 – 2025.
- `Ngày Chốt Danh Sách`: Ngày chốt thực tế ghi trên bảng (YYYY-MM-DD).
- `Tỷ Lệ Sở Hữu Nhà Nước (%)`: Tỷ lệ phần trăm vốn thực tế (định dạng `0.00%`, rỗng nếu chưa xác minh).
- `Số Lượng Cổ Phiếu`: Số cổ phần do Nhà nước nắm giữ (định dạng `#,#00`).
- `Tổ Chức Đại Diện Vốn Nhà Nước`: Tên cơ quan/tổ chức cụ thể (SCIC, Bộ, UBND...).
- `Cơ Sở Xác Định`: Căn cứ dòng tổng hoặc tổng hợp tổ chức.
- `Bằng Chứng Trích Xuất`: Trích đoạn nguyên văn bảng biểu làm bằng chứng.
- `File Nguồn`: Tên file PDF gốc tương ứng.
- `Trang`: Số trang chứa bảng dữ liệu trong PDF.
- `Trạng Thái / Lý Do Chưa Xác Minh`: Trạng thái chi tiết (`DA_XAC_MINH_CO_SHNN`, `DA_XAC_MINH_KHONG_CO_SHNN`, `CAN_KIEM_TRA_NGAY_KHAC`, `CAN_KIEM_TRA_MAU_THUAN`, `CHUA_TIM_THAY_THONG_TIN`).

---

## 5. Quy Định Bảo Mật & Lưu Trữ Git

- Tuyệt đối **không** đưa các file PDF nặng (`pdf_101x6/`), file cache (`.cache_ownership/`), các file log (`*.log`) hoặc API keys (`.env`, `listapi`) lên repository Git.
- Dữ liệu đối chiếu được lưu độc lập tại `checkpoint_ownership.jsonl` và file Excel đầu ra chính thức.
