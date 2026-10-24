# Xử lý sở hữu Nhà nước từ báo cáo thường niên

Chỉ có một script xử lý: `ownership_101x6.py`. Chạy từ `E:\Tools\crawl data` với Python 3.10+:

```powershell
pip install -r ownership_cli/requirements.txt
python ownership_cli/ownership_101x6.py
```

Đầu vào phải tự cung cấp tại máy chạy: `danh_muc_pdf_101x6.csv` và thư mục `pdf_101x6`. Các cache trong `ownership_extract_101x6` là tùy chọn, **không được tải lên GitHub**. Nếu cache lớp chữ thiếu, script đọc lại PDF bằng PyMuPDF. Nếu chưa có `ratio_candidates.csv`, script vẫn xuất đủ dòng theo manifest nhưng tỷ lệ thiếu sẽ để `Không có` đến khi có bằng chứng và kiểm định. Kết quả mặc định là `SO_HUU_NHA_NUOC_101_MA_2020_2025.xlsx`.

Để OCR lấy mẫu các báo cáo chưa tìm thấy trang cổ đông và kiểm định ứng viên mới bằng Motapis:

```powershell
$env:MOTAPIS_API_KEY = '<khóa của bạn>'
python ownership_cli/ownership_101x6.py --ocr-missing --audit
Remove-Item Env:MOTAPIS_API_KEY
```

Có thể chạy riêng `--ocr-missing` hoặc `--audit`. Để OCR, cài Tesseract cùng bộ ngôn ngữ `vie` và `eng`; script dùng `ownership_cli/tessdata` nếu có, nếu không dùng cấu hình Tesseract hệ thống. Mặc định OCR tối đa 6 trang scan mỗi PDF ở 160 DPI; thay bằng `--max-pages` và `--dpi` nếu cần. Cache OCR và kiểm định giúp chạy lại không lặp công việc đã lưu. Khóa API chỉ được đọc từ biến môi trường, không lưu trong mã nguồn.

Quy tắc: chỉ xét tỷ lệ có ngày dữ liệu trong **15/12–31/12** của năm báo cáo; không lấy riêng số liệu năm 2026, không suy đoán tư cách Nhà nước của cổ đông tổ chức, không biến thiếu dữ liệu thành 0%. `Không có` trong Excel nghĩa là chưa tìm được tỷ lệ đủ điều kiện, không chứng minh doanh nghiệp không có vốn Nhà nước. OCR lấy mẫu không tương đương đọc hết mọi trang scan; tỷ lệ do API kiểm định vẫn cần đối chiếu thủ công trên trang PDF được liên kết.
