# AI Agent - Phân bổ chi phí điện

## Rule-01: Phân bổ chi phí điện

### Tính năng
- Đọc dữ liệu từ PDF, Excel, Image (OCR), API
- Match theo Mã khách hàng
- Update cột 15 (Tổng tiền của một điểm)
- Shift cột 13, 14 theo tháng (file gốc = tháng 7)
- Bỏ qua dòng "di dời"
- Tự đặt tên file output theo tháng

### Cách dùng
1. Chạy `run_app.bat` để mở GUI app
2. Hoặc chạy script: `python update_excel_auto.py "filename.pdf"`

### Yêu cầu
- Python 3.12+
- Thư viện: pdfplumber, openpyxl, pytesseract, requests, Pillow
