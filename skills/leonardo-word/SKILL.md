---
name: leonardo-word
description: Bộ chuẩn trình bày văn bản Word của Leonardo (lề, phông chữ, hệ tiêu đề đánh số, kiểu bảng, đầu trang có logo, chân trang có số trang, bảng kiểm soát tài liệu, khối ký). Dùng skill này mỗi khi soạn một file Word (.docx) cho Leonardo — quy chế, quy trình, báo cáo, tờ trình, biên bản, thông báo, thư gửi đối tác, tài liệu hướng dẫn — và khi người dùng nói "theo chuẩn Leonardo", "theo mẫu công ty", "làm file Word", "soạn văn bản", "xuất ra docx", kể cả khi không nhắc tên skill. Dùng chung bảng màu, logo, thông tin công ty và quy tắc chữ với skill leonardo-excel.
---

# Chuẩn trình bày văn bản Word của Leonardo

Bộ chuẩn bằng chữ nằm ở `STYLE.md` (kèm căn cứ tham khảo). Đọc file đó khi người dùng hỏi về chuẩn hoặc khi cần bố cục mà script chưa hỗ trợ.

Python: `~/.local/share/leonardo-excel-venv/bin/python` (gọi tắt `$PY`, có sẵn python-docx). Luôn chạy với `-I`.

## Tạo văn bản mới

1. Xác định: loại văn bản (quy chế, quy trình, báo cáo, tờ trình, biên bản…), tên, mã tài liệu nếu có, người soạn và người duyệt. Thông tin nào người dùng không cho thì để trống trong bảng kiểm soát, không tự điền.
2. Viết nội dung theo quy tắc chữ trong `STYLE.md`: kết luận trước, câu ngắn, viết đầy đủ chữ tắt ở lần đầu, không bịa số liệu.
3. Viết bản mô tả JSON theo mẫu `examples/tieu-chuan.json` (cấu trúc ở đầu `scripts/new_doc.py`).
   - Tối đa ba cấp tiêu đề; **không gõ số** vào tiêu đề và vào danh sách `numbered`: Word tự đánh 1. / 1.1 / 1.1.1 và (a) / (i).
   - Văn bản trên năm trang: thêm `"cover": true, "toc": true` để có trang bìa, mục lục và số trang i, ii cho phần đầu. Văn bản ngắn thì bỏ `cover`.
   - Thông tin dạng nhãn – giá trị (các bên hợp đồng, thông tin đối tác) dùng khối `info`, không dùng bảng số liệu.
   - Số trong bảng để dạng số (không để chuỗi) để được canh phải, ngăn nghìn, số âm trong ngoặc.
   - Mỗi bảng có `caption`; dòng tổng dùng `"total_row": true`.
   - Nhấn mạnh bằng `**đậm**`; không dùng in hoa cả câu hay gạch chân.
4. Tạo file: `$PY -I scripts/new_doc.py mo-ta.json ket-qua.docx` (không ghi đè file đã có).
5. Kiểm tra cấu trúc (bắt buộc): `$PY -I scripts/check_docx.py ket-qua.docx`. Chỉ giao khi `KẾT LUẬN: ĐẠT`.
6. Xem trước trang đầu: `qlmanage -t -s 1700 -o thu-muc ket-qua.docx` rồi đọc ảnh (chỉ macOS).
7. Nếu có mục lục: báo người dùng rằng khi mở file Word sẽ hỏi có cập nhật các trường không; bấm **Yes** để mục lục hiện số trang.

Bố cục JSON không tả được (hình ảnh, bảng gộp ô): viết mã Python ngắn dùng lớp `Doc` trong `scripts/leo_doc.py`; không gõ lại mã màu, phông, lề ở nơi khác. Phần tử XML tự thêm phải đi qua hàm `_put()` để đúng thứ tự lược đồ.

## Trình bày lại một file Word có sẵn

Nguyên lý: chép nguyên văn chữ và số của bản gốc sang một file mới dựng theo chuẩn, rồi so chữ trước và sau.

1. Chép file gốc vào thư mục tạm, đọc qua để biết đó là văn bản của ai và có gì đặc biệt (chữ ký, bảng gộp ô, hình).
2. Chạy: `$PY -I scripts/restyle_doc.py goc.docx ket-qua.docx [--header-title "tên ngắn ở đầu trang"] [--letterhead]`
   - **Mặc định không đặt logo và thông tin Leonardo.** Chỉ thêm `--letterhead` khi văn bản là của Leonardo. Văn bản của tổ chức khác mà gắn nhận diện Leonardo sẽ gây hiểu nhầm về người phát hành.
   - Script tự so chữ và chỉ giữ file khi mọi đoạn và mọi ô bảng khớp. Đọc dòng "KHÔNG mang theo" và báo lại cho người dùng (chữ ký số, ghi chú, theo dõi thay đổi, chú thích cuối trang, mục lục tự động, hộp chữ).
3. `$PY -I scripts/check_docx.py ket-qua.docx`, rồi xem ảnh trang đầu bằng Quick Look.
4. Đặt file kết quả cạnh file gốc với tên mới. Không ghi đè.

Cách script nhận diện: các dòng in hoa canh giữa ở đầu là tên văn bản; mục đánh số 1., 2.… là tiêu đề cấp 1; bảng đứng trước tên văn bản và bảng cuối cùng không có số liệu là bảng dàn trang (quốc hiệu, khối ký), giữ bố cục và không kẻ viền; các bảng khác nhận kiểu bảng chuẩn. Dòng đầu bảng là tiêu đề cột chỉ khi nó không chứa số. Chữ và cách ghi số trong bảng được giữ nguyên, không đổi sang quy tắc số của Leonardo.

## Ràng buộc cứng

- **Không ghi đè file của người dùng**; luôn ra file mới.
- **Không tự điền** người ký, ngày hiệu lực, số hiệu văn bản khi không có nguồn.
- **Không mở-lưu lại một file Word có sẵn bằng python-docx để "làm đẹp"** nếu chưa so chữ trước và sau bằng `check_docx.py ket-qua.docx goc.docx`: thao tác đó có thể làm mất ghi chú, theo dõi thay đổi, mục lục tự động.
- **Nói thật về việc kiểm tra:** script không mở được Word. Ảnh Quick Look chỉ cho thấy trang đầu, không hiện đúng độ rộng và độ lùi của bảng, các đường kẻ đoạn (kể cả đường kẻ trên tiêu đề cấp 1), ngắt trang, đầu trang các trang sau và số trang. Báo rõ phần nào đã xem, phần nào chưa, và nhắc người dùng mở bằng Word.
- **Văn bản pháp lý gửi cơ quan nhà nước** có thể thức riêng theo Nghị định 30/2020/NĐ-CP; chuẩn này dành cho văn bản nội bộ và văn bản gửi đối tác. Gặp trường hợp đó thì nói rõ với người dùng trước khi làm.

## Chưa hỗ trợ

Soát chữ trong file Word chưa có lệnh riêng. Khi trình bày lại, tiêu đề cấp 2, cấp 3 của bản gốc được giữ làm đoạn thường và số mục giữ dạng chữ gõ sẵn (không chuyển sang số tự động). Chưa có: dẫn chiếu tự cập nhật giữa các điều, dòng chấm để điền, ô đánh dấu, chú thích cuối trang, văn bản song ngữ, chế độ đầu trang gọn.
