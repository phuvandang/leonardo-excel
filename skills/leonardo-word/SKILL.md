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
   - Tối đa ba cấp tiêu đề; **không gõ số** vào tiêu đề, script tự đánh 1. / 1.1 / 1.1.1.
   - Số trong bảng để dạng số (không để chuỗi) để được canh phải, ngăn nghìn, số âm trong ngoặc.
   - Mỗi bảng có `caption`; dòng tổng dùng `"total_row": true`.
   - Nhấn mạnh bằng `**đậm**`; không dùng in hoa cả câu hay gạch chân.
4. Tạo file: `$PY -I scripts/new_doc.py mo-ta.json ket-qua.docx` (không ghi đè file đã có).
5. Kiểm tra cấu trúc (bắt buộc): `$PY -I scripts/check_docx.py ket-qua.docx`. Chỉ giao khi `KẾT LUẬN: ĐẠT`.
6. Xem trước trang đầu: `qlmanage -t -s 1700 -o thu-muc ket-qua.docx` rồi đọc ảnh (chỉ macOS).

Bố cục JSON không tả được (hình ảnh, mục lục, bảng gộp ô): viết mã Python ngắn dùng lớp `Doc` trong `scripts/leo_doc.py`; không gõ lại mã màu, phông, lề ở nơi khác. Phần tử XML tự thêm phải đi qua hàm `_put()` để đúng thứ tự lược đồ.

## Ràng buộc cứng

- **Không ghi đè file của người dùng**; luôn ra file mới.
- **Không tự điền** người ký, ngày hiệu lực, số hiệu văn bản khi không có nguồn.
- **Không mở-lưu lại một file Word có sẵn bằng python-docx để "làm đẹp"** nếu chưa so chữ trước và sau bằng `check_docx.py ket-qua.docx goc.docx`: thao tác đó có thể làm mất ghi chú, theo dõi thay đổi, mục lục tự động.
- **Nói thật về việc kiểm tra:** script không mở được Word. Ảnh Quick Look chỉ cho thấy trang đầu, không hiện đúng độ rộng bảng, đầu trang các trang sau và số trang. Báo rõ phần nào đã xem, phần nào chưa, và nhắc người dùng mở bằng Word.
- **Văn bản pháp lý gửi cơ quan nhà nước** có thể thức riêng theo Nghị định 30/2020/NĐ-CP; chuẩn này dành cho văn bản nội bộ và văn bản gửi đối tác. Gặp trường hợp đó thì nói rõ với người dùng trước khi làm.

## Chưa hỗ trợ

Trình bày lại một file Word có sẵn và soát chữ trong file Word chưa có lệnh riêng. Khi được yêu cầu: đọc nội dung file, dựng lại bằng luồng tạo mới, rồi so chữ với bản gốc bằng `check_docx.py` và báo mọi chỗ khác.
