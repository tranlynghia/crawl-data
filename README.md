# Xử lý sở hữu Nhà nước từ báo cáo thường niên

Chỉ có một script xử lý: `ownership_101x6.py`. Chạy từ `E:\Tools\crawl data` với Python 3.10+:

```powershell
pip install -r ownership_cli/requirements.txt
python ownership_cli/ownership_101x6.py
```

Đầu vào phải tự cung cấp tại máy chạy: `danh_muc_pdf_101x6.csv` và thư mục `pdf_101x6`. Các cache trong `ownership_extract_101x6` là tùy chọn, **không được tải lên GitHub**. Nếu cache lớp chữ thiếu, script đọc lại PDF bằng PyMuPDF. Kết quả mặc định là `SO_HUU_NHA_NUOC_101_MA_2020_2025.xlsx`.

Để quét tất cả trang PDF dạng scan (kể cả khi báo cáo đã có một trang cổ đông dạng văn bản):

```powershell
python ownership_cli/ownership_101x6.py --ocr-missing
```

Để OCR, cài Tesseract cùng bộ ngôn ngữ `vie` và `eng`; script dùng `ownership_cli/tessdata` nếu có, nếu không dùng cấu hình Tesseract hệ thống. Mặc định OCR mọi trang scan ở 160 DPI; dùng `--max-pages N` để giới hạn số trang mới mỗi PDF trong một lượt và `--dpi` để đổi độ phân giải. Cache OCR giúp tiếp tục lần sau mà không bỏ qua các trang chưa đọc. Chạy toàn bộ OCR có thể tốn nhiều thời gian và dung lượng.

`--audit` cũ tạm khóa vì chỉ kiểm định một ứng viên rồi bỏ sót ứng viên khác; lệnh này báo lỗi thay vì tốn phí API và tạo kết luận thiếu.

`Kiem_tra` có một dòng mỗi mã-năm; `Ung_vien` giữ các dòng/trang cổ đông khác cùng cột, giá trị gốc, ngày và nguồn để đối chiếu. Điểm ưu tiên chỉ sắp thứ tự, không loại trang. Cột tỷ lệ Nhà nước chỉ điền khi bảng ghi rõ tổng sở hữu Nhà nước, đúng cột tỷ lệ sở hữu cổ phần/vốn và ngày dữ liệu **31/12** của năm mục tiêu; nếu hai nguồn cùng ngày mâu thuẫn thì để trống và báo cần kiểm tra. Ngày khác vẫn ở `Ung_vien`. Không suy đoán tư cách Nhà nước của cổ đông tổ chức, không cộng phần vốn HĐQT đại diện, không biến thiếu dữ liệu thành 0%. Bảng có thể ghi rõ 0%, nhưng ô trống không chứng minh 0%. Tỷ lệ do máy đọc vẫn nên đối chiếu trực tiếp trên PDF.
