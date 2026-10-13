# Bộ chuẩn trình bày văn bản Word của Leonardo

Mã nguồn chuẩn là `scripts/leo_doc.py`; file này là bản đọc cho người. Bản minh hoạ đầy đủ: tạo từ `examples/tieu-chuan.json`.

## Nguyên lý

Thứ bậc trong văn bản thể hiện bằng **vị trí, cỡ chữ và độ đậm**, không bằng màu. Cả văn bản dùng một phông. Màu xanh nhận diện chỉ xuất hiện ở tên văn bản, các đường kẻ và bảng số liệu; tiêu đề mục và nội dung đều màu đen.

## Căn cứ tham khảo

| Nguồn | Điều lấy về | Áp dụng |
|---|---|---|
| Nghị định 30/2020/NĐ-CP, Phụ lục I (thể thức văn bản hành chính) | Khổ A4; lề trên, dưới 20–25 mm, trái 30–35 mm, phải 15–20 mm; phông Times New Roman, chữ đen; số trang bằng chữ số Ả Rập | Lề 25 / 20 / 30 / 20 mm; nội dung Times New Roman đen. Nghị định chỉ bắt buộc với cơ quan nhà nước; ở đây dùng làm thông lệ quen mắt người Việt |
| Butterick, *Practical Typography* | Cỡ chữ in 10–12 pt; giãn dòng 120–145% cỡ chữ; 45–90 ký tự mỗi dòng | Cỡ 10,5, giãn dòng 1,3, dòng chữ 145 mm (khoảng 75–80 ký tự) |
| Hai văn bản mẫu của một công ty quản lý quỹ (hợp đồng 16 trang, phiếu lấy ý kiến 6 trang), phân tích ngày 2026-10-13 | Một phông cho cả văn bản; tiêu đề đen, phân cấp bằng cỡ chữ và đường kẻ mảnh; số mục đứng riêng một cột 15 mm; bảng "nhãn – giá trị" không lưới; bìa, mục lục tự động, phần đầu đánh số i, ii | Phiên bản 1.1 của chuẩn này. Chỉ lấy nguyên tắc, không dùng lại bộ khuôn của họ |
| Hướng dẫn của Microsoft về đánh số tiêu đề | Dùng kiểu tiêu đề có sẵn (Heading 1–3) để có ngăn điều hướng và mục lục | Tiêu đề dùng Heading 1–3 |
| Thông lệ kiểm soát tài liệu của các tổ chức có hệ thống quản lý chất lượng | Bảng nhận diện tài liệu: mã, phiên bản, ngày hiệu lực, người soạn, người duyệt | Bảng kiểm soát ở trang đầu |

Cỡ chữ nội dung, cách đoạn và kiểu tiêu đề là lựa chọn thiết kế của Leonardo trong khung các nguồn trên, không phải quy định pháp lý.

## Trang

A4 dọc. Lề trên 25, dưới 20, trái 30, phải 20 mm; vùng trang rộng 160 mm. Số mục nằm trong cột 15 mm đầu tiên; tiêu đề và đoạn văn bắt đầu sau cột đó, nên dòng chữ rộng 145 mm. Bảng số liệu và khung ghi chú dùng trọn 160 mm. Văn bản không có tiêu đề mục (thư, thông báo ngắn) thì đoạn văn không lùi vào.

## Phông chữ

Toàn bộ dùng **Tahoma**, cùng phông với chuẩn Excel.

| Vai trò | Cỡ | Kiểu |
|---|---|---|
| Tên văn bản trên bìa | 24 | Đậm, xanh `2F5D4A`, canh trái |
| Tên văn bản ở trang đầu (văn bản ngắn) | 16 | Đậm, in hoa, xanh `2F5D4A`, canh giữa |
| Loại văn bản (trên tên) | 11 | Đậm, in hoa, xám |
| Tiêu đề cấp 1 | 12 | Đậm, in hoa, đen; đường kẻ xanh mảnh phía trên |
| Tiêu đề cấp 2 | 11 | Đậm, đen |
| Tiêu đề cấp 3 | 10,5 | Đậm nghiêng, đen |
| Nội dung | 10,5 | Đen, canh đều hai bên, giãn dòng 1,3, cách đoạn 6 pt, không thụt đầu dòng |
| Bảng | 9,5 | Tiêu đề cột đậm trắng trên nền xanh |
| Đầu trang, chân trang | 8 | Xám `6B6B6B` |

Tahoma có sẵn trên mọi máy Windows và macOS và hiển thị đúng tiếng Việt. Phông nhận diện của Leonardo (Garamond Pro, SVN-Century Gothic) không dùng ở đây vì máy nào chưa cài sẽ hiển thị sai.

## Màu

Dùng chung với chuẩn Excel: xanh rừng `2F5D4A`, nền nhạt `EEF3EF`, lưới `4A5A52`, đỏ cảnh báo `8F2F1D`, nền ô cần điền `FBF5E0`.

## Cấu trúc văn bản

**Văn bản ngắn (đến năm trang):**

| Phần | Nội dung |
|---|---|
| Đầu trang, trang đầu | Logo bên trái; tên công ty, mã số thuế, địa chỉ bên phải; đường kẻ xanh bên dưới |
| Khối tên | Loại văn bản; tên văn bản; một dòng mô tả in nghiêng |
| Bảng nhận diện tài liệu | Mã, phiên bản, ngày hiệu lực, phạm vi, người soạn, người duyệt; kiểu "nhãn – giá trị", hai cặp cột |
| Thân văn bản | Tiêu đề đánh số 1. / 1.1 / 1.1.1, tối đa ba cấp; ý liệt kê có thứ tự dùng (a) rồi (i) |
| Khối ký | Nơi và ngày canh phải; chức danh in hoa đậm; "(Ký, ghi rõ họ tên)"; họ tên |
| Đầu trang, các trang sau | Logo nhỏ bên trái; tên ngắn và mã tài liệu bên phải |
| Chân trang | Tên công ty bên trái; "Trang x / y" bên phải |

**Văn bản dài (trên năm trang)** có thêm phần đầu:

| Phần | Nội dung |
|---|---|
| Trang bìa | Đầu trang có logo như trên; loại văn bản; tên văn bản cỡ 24 canh trái; dòng mô tả; bảng nhận diện tài liệu. Không hiện số trang |
| Mục lục | Hai cấp, là trường mục lục của Word; số trang do Word điền |
| Số trang | Phần đầu đánh i, ii; thân bài đánh lại từ 1, chân trang ghi "Trang x / y" theo số trang của thân bài |

Số mục do Word tự đánh (không gõ tay), nằm trong cột số 15 mm. Tiêu đề luôn đi liền đoạn sau nó. Khối ký không nằm riêng một trang.

## Bảng

**Bảng số liệu**

- Tiêu đề cột nền xanh chữ trắng, canh giữa; lưới mảnh; dòng tổng đậm trên nền nhạt.
- Chữ canh trái, số canh phải.
- Tên bảng ở phía trên: "**Bảng n.** Tên bảng", đánh số liên tục; ghi đơn vị ở tên bảng hoặc tiêu đề cột.
- Bảng sang trang thì lặp lại dòng tiêu đề.

**Bảng thông tin ("nhãn – giá trị")**

Dùng cho các bên của hợp đồng, thông tin đối tác, nhận diện tài liệu. Không kẻ lưới: nhãn cỡ 8,5 màu xám bên trái, một vạch dọc xanh, giá trị cỡ nội dung bên phải.

**Khung ghi chú:** vạch xanh bên trái, nền nhạt, mở đầu bằng nhãn đậm ("Lưu ý:", "Liên quan:").

## Chữ và số

Theo `STYLE.md` của skill leonardo-excel: tiêu đề cấp 1 in hoa, các cấp khác chỉ hoa chữ đầu; dấu thanh kiểu cũ; tháng hai chữ số; tên riêng viết đúng; dấu chấm ngăn nghìn, dấu phẩy thập phân, số âm trong ngoặc. Riêng với văn bản: chữ tắt viết đầy đủ ở lần đầu kèm chữ tắt trong ngoặc; nhấn mạnh bằng chữ đậm, không gạch chân, không in hoa cả câu.

## Tên file

`Mã_Tên-ngắn_vPhiên-bản.docx`, ví dụ `QT-KT-01_Quy-trinh-thanh-toan_v1.0.docx`.

## Lịch sử phiên bản

| Phiên bản | Ngày | Thay đổi |
|---|---|---|
| 1.0 | 2026-10-13 | Bản đầu: Times New Roman cho nội dung, Tahoma cho tiêu đề; tiêu đề cấp 1 màu xanh; số mục gõ sẵn |
| 1.1 | 2026-10-13 | Một phông Tahoma, nội dung cỡ 10,5; tiêu đề đen với đường kẻ mảnh; số mục do Word tự đánh, cột số 15 mm, thêm cấp (a), (i); bảng "nhãn – giá trị"; trang bìa, mục lục, đánh số trang theo phần |
