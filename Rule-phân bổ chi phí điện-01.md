# Rule: Phân bổ chi phí điện - 01

## Thông tin chung
- **Tên rule:** Rule-phân bổ chi phí điện-01
- **Loại:** Phân bổ chi phí điện
- **Thư mục dữ liệu:** `D:\File storage\AI\AI Agent\Điện\`

## File template
- **File gốc:** `File gốc phân bổ chi phí điện.xlsx` (luôn dùng làm template)
- **File PDF:** Đặt vào cùng thư mục, tên chứa "tháng X.2026"
- **File output:** `Bảng phân bổ thanh toán chi phí điện tháng X.2026.xlsx`

## Quy tắc xử lý
1. Đọc dữ liệu đầu vào (PDF/Excel/Image/API): trích xuất mã thanh toán + số tiền
2. **Tham chiếu duy nhất: Mã khách hàng** (input ↔ Excel cột 8)
3. Match → điền **cột 15** (Tổng tiền của một điểm) = số tiền từ input
4. **Cột 13, 14 (Kỳ thanh toán):** shift theo tháng
   - File gốc = tháng 7 (mốc tham chiếu)
   - Shift = tháng input - 7
   - Tháng 6 → trừ 1 tháng | Tháng 9 → cộng 2 tháng
   - Giữ nguyên định dạng khoảng ngày (dd/mm/yyyy - dd/mm/yyyy)
5. **Các cột khác:** giữ nguyên công thức Excel
6. **Dòng "di dời":** cột 15 để trống, không shift
7. **Bỏ qua giao dịch MSB** (nội bộ ngân hàng)
8. **File output** đặt tên theo tháng, chủ đề tự đổi theo tháng

## Cấu trúc Excel
- Sheet "Phân bổ" (sheet thứ 2): 133 dòng, 16 cột
- Dòng 4: header, dòng 5+: data
- Cột 8: Mã khách hàng (key match)
- Cột 13: Kỳ thanh toán này (giữ nguyên / Excel formula)
- Cột 14: Kỳ thanh toán trước (giữ nguyên)
- Cột 15: Tổng tiền của một điểm (điền từ PDF)

## Script
- `D:\File storage\AI\AI Agent\Điện\update_excel_auto.py`
- `D:\File storage\AI\AI Agent\Điện\app.py` (GUI app)
- `D:\File storage\AI\AI Agent\Điện\run_app.bat` (launcher)
