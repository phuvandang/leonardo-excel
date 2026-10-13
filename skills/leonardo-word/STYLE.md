# Bộ chuẩn trình bày văn bản Word của Leonardo

Mã nguồn chuẩn là `scripts/leo_doc.py`; file này là bản đọc cho người. Bản minh hoạ đầy đủ: tạo từ `examples/tieu-chuan.json`.

## Nguyên lý

Văn bản dễ đọc khi mắt chỉ phải phân biệt ba thứ: **tiêu đề**, **nội dung** và **số liệu**. Một màu nhận diện cho tiêu đề, chữ đen cho nội dung, một kiểu bảng duy nhất cho số liệu. Màu không dùng để trang trí.

## Căn cứ tham khảo

| Nguồn | Điều lấy về | Áp dụng |
|---|---|---|
| Nghị định 30/2020/NĐ-CP, Phụ lục I (thể thức văn bản hành chính) | Khổ A4; lề trên, dưới 20–25 mm, trái 30–35 mm, phải 15–20 mm; phông Times New Roman, chữ đen; số trang bằng chữ số Ả Rập | Lề 25 / 20 / 30 / 20 mm; nội dung Times New Roman đen. Nghị định chỉ bắt buộc với cơ quan nhà nước; ở đây dùng làm thông lệ quen mắt người Việt |
| Butterick, *Practical Typography* | Cỡ chữ in 10–12 pt; giãn dòng 120–145% cỡ chữ; 45–90 ký tự mỗi dòng | Cỡ 12, giãn dòng 1,3, vùng chữ 160 mm |
| Hướng dẫn của Microsoft về đánh số tiêu đề | Dùng kiểu tiêu đề có sẵn (Heading 1–3) để có ngăn điều hướng và mục lục | Tiêu đề dùng Heading 1–3 |
| Thông lệ kiểm soát tài liệu của các tổ chức có hệ thống quản lý chất lượng | Bảng nhận diện tài liệu: mã, phiên bản, ngày hiệu lực, người soạn, người duyệt | Bảng kiểm soát ở trang đầu |

Cỡ chữ nội dung, cách đoạn và kiểu tiêu đề là lựa chọn thiết kế của Leonardo trong khung các nguồn trên, không phải quy định pháp lý.

## Trang

A4 dọc. Lề trên 25, dưới 20, trái 30, phải 20 mm; vùng chữ rộng 160 mm.

## Phông chữ

| Vai trò | Phông | Cỡ | Kiểu |
|---|---|---|---|
| Tên văn bản | Tahoma | 16 | Đậm, in hoa, xanh `2F5D4A` |
| Loại văn bản (trên tên) | Tahoma | 11 | Đậm, in hoa, xám |
| Tiêu đề cấp 1 | Tahoma | 13 | Đậm, in hoa, xanh `2F5D4A` |
| Tiêu đề cấp 2 | Tahoma | 12 | Đậm, đen |
| Tiêu đề cấp 3 | Tahoma | 11 | Đậm nghiêng, đen |
| Nội dung | Times New Roman | 12 | Đen, canh đều hai bên, giãn dòng 1,3, cách đoạn 6 pt, không thụt đầu dòng |
| Bảng | Tahoma | 10 | Tiêu đề cột đậm trắng trên nền xanh |
| Đầu trang, chân trang | Tahoma | 8 | Xám `6B6B6B` |

Chọn Times New Roman và Tahoma vì có sẵn trên mọi máy Windows và macOS và hiển thị đúng tiếng Việt. Phông nhận diện của Leonardo (Garamond Pro, SVN-Century Gothic) không dùng ở đây vì máy nào chưa cài sẽ hiển thị sai.

## Màu

Dùng chung với chuẩn Excel: xanh rừng `2F5D4A`, nền nhạt `EEF3EF`, lưới `4A5A52`, đỏ cảnh báo `8F2F1D`, nền ô cần điền `FBF5E0`.

## Cấu trúc văn bản

| Phần | Nội dung |
|---|---|
| Đầu trang, trang đầu | Logo bên trái; tên công ty, mã số thuế, địa chỉ bên phải; đường kẻ xanh bên dưới |
| Khối tên | Loại văn bản; tên văn bản; một dòng mô tả in nghiêng |
| Bảng kiểm soát | Mã tài liệu, phiên bản, ngày hiệu lực, phạm vi, người soạn, người duyệt |
| Thân văn bản | Tiêu đề đánh số 1. / 1.1 / 1.1.1, tối đa ba cấp |
| Khối ký | Nơi và ngày canh phải; chức danh in hoa đậm; "(Ký, ghi rõ họ tên)"; họ tên |
| Đầu trang, các trang sau | Logo nhỏ bên trái; tên ngắn và mã tài liệu bên phải |
| Chân trang | Tên công ty bên trái; "Trang x / y" bên phải |

Tiêu đề luôn đi liền đoạn sau nó. Văn bản dài hơn năm trang thì có mục lục. Khối ký không nằm riêng một trang.

## Bảng

- Tiêu đề cột nền xanh chữ trắng, canh giữa; lưới mảnh; dòng tổng đậm trên nền nhạt.
- Chữ canh trái, số canh phải.
- Tên bảng ở phía trên: "**Bảng n.** Tên bảng", đánh số liên tục; ghi đơn vị ở tên bảng hoặc tiêu đề cột.
- Bảng sang trang thì lặp lại dòng tiêu đề.
- Khung ghi chú: vạch xanh bên trái, nền nhạt, mở đầu bằng nhãn đậm ("Lưu ý:", "Liên quan:").

## Chữ và số

Theo `STYLE.md` của skill leonardo-excel: tiêu đề cấp 1 in hoa, các cấp khác chỉ hoa chữ đầu; dấu thanh kiểu cũ; tháng hai chữ số; tên riêng viết đúng; dấu chấm ngăn nghìn, dấu phẩy thập phân, số âm trong ngoặc. Riêng với văn bản: chữ tắt viết đầy đủ ở lần đầu kèm chữ tắt trong ngoặc; nhấn mạnh bằng chữ đậm, không gạch chân, không in hoa cả câu.

## Tên file

`Mã_Tên-ngắn_vPhiên-bản.docx`, ví dụ `QT-KT-01_Quy-trinh-thanh-toan_v1.0.docx`.
