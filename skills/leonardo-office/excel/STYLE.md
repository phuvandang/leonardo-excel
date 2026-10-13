# Bộ chuẩn trình bày Excel của Leonardo

Chuẩn này rút từ báo cáo quản trị tháng 09/2026, theo cách trình bày của báo cáo tháng quỹ DCDS và bộ nhận diện Leonardo. Mã nguồn chuẩn là `scripts/leo_style.py`; file này là bản đọc cho người.

## Nguyên lý

Báo cáo đọc tốt khi mắt chỉ phải phân biệt ba thứ: **tiêu đề**, **dòng tổng** và **dòng chi tiết**. Vì vậy chỉ dùng một màu đậm cho tiêu đề cột, một nền rất nhạt cho dòng tổng, còn lại chữ đen trên nền trắng với lưới mảnh. Màu không dùng để trang trí.

## Màu

| Vai trò | Mã màu |
|---|---|
| Tên báo cáo; nền tiêu đề cột (chữ trắng đậm) | `2F5D4A` |
| Nền dòng tổng, dòng nhóm | `EEF3EF` |
| Đường lưới (mảnh, kín bảng) | `4A5A52` |
| Ô cần người nhập | `FBF5E0` |
| Chữ cảnh báo (vượt ngưỡng, lỗ) | `8F2F1D` |
| Nền ngoài vùng báo cáo | `A6A6A6` |
| Chuỗi biểu đồ, theo thứ tự | `2F5D4A`, `6F9A85`, `C5CEC9`, `D8C48C`, `E6D8AC` |

Không tô màu thẻ sheet.

## Phông chữ

Tahoma cỡ 10 cho toàn bộ; tên báo cáo cỡ 13 đậm; dòng kỳ báo cáo và đơn vị in nghiêng.

## Bố cục một sheet

| Dòng | Nội dung |
|---|---|
| 1 | Logo (góc trái) và tên báo cáo in hoa, dạng `TÊN BỘ BÁO CÁO — TÊN PHẦN` |
| 2 | Kỳ: `Tháng 09 năm 2026` hoặc `Tại ngày 30 tháng 09 năm 2026` |
| 3 | `Đơn vị: đồng` (hoặc triệu đồng) |
| 4–9 | Tên công ty, mã số thuế, địa chỉ, kỳ báo cáo, ngày lập, người lập / người duyệt |
| 11 | `A.` + tên phần |
| 12 trở đi | Các khối bảng, đánh số `I`, `II`, `III`…, cách nhau một dòng trống |

Mỗi bảng có ba cột đầu cố định: **STT**, **nội dung**, **mã chỉ tiêu**; cột giá trị bắt đầu từ cột D. Dòng tổng đánh số La Mã, dòng chi tiết đánh số 1, 2, 3 trong từng nhóm.

Xem ở chế độ Normal: trong vùng báo cáo nền trắng, ngoài vùng nền xám. Vùng in: A4 ngang, vừa một trang theo chiều rộng.

Tên sheet: `NN Tên` (ví dụ `01 KQKD`, `05b Cân đối lịch sử`). Sheet tính toán bị ẩn có tiền tố `tính`, sheet dữ liệu thô có tiền tố `DL1`, `DL2`…

## Chữ

| Loại | Cách viết | Ví dụ |
|---|---|---|
| Tên báo cáo, tiêu đề mục | In hoa toàn bộ, kể cả trong ngoặc | `CHI PHÍ TRỰC TIẾP (GÁN THEO MÃ CỬA HÀNG)` |
| Dòng tổng lớn | In hoa | `LỢI NHUẬN GỘP` |
| Tiêu đề cột, dòng chi tiết, dòng "Cộng…" | Chỉ hoa chữ đầu và tên riêng | `Tổng công ty`, `Cộng sàn TMĐT` |
| Danh từ chung giữa câu | Viết thường | `Nợ vay ngân hàng / vốn chủ sở hữu` |

- Dấu thanh kiểu cũ: `hòa`, `hóa`, `lũy`, `xóa`.
- Tháng ghi hai chữ số: `tháng 08`; so sánh viết đủ: `Tháng 09 so với tháng 08`.
- Tên riêng viết đúng: Shopee, TikTok, Lazada, VNPAY, OnePay, MoMo, Facebook, 1Office.
- Viết đầy đủ chữ tắt người ngoài phòng kế toán không hiểu (NCC → nhà cung cấp, MB → mặt bằng, DT → doanh thu, VAT → thuế GTGT). Giữ: GTGT, TMĐT, D2C, KQKD, ROE, HCM, tên ngân hàng, tên đơn vị giao hàng.
- Sau số thứ tự có dấu cách (`3. Phải thu`); dấu ba chấm là `...`; có dấu cách quanh dấu cộng; không có dấu cách trước dấu hai chấm.
- Chữ tắt không rõ nghĩa: để nguyên và hỏi lại, không đoán.

## Số

- Dấu chấm ngăn nghìn, dấu phẩy thập phân (theo máy Việt Nam).
- Số âm trong ngoặc, số 0 hiện dấu gạch `-`.
- Phần trăm một chữ số thập phân.
- Số tổng, chênh lệch, tỷ lệ luôn là công thức; không gõ kết quả đã tính.
