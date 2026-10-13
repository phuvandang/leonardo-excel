---
name: leonardo-office
description: Bộ chuẩn trình bày file Excel và Word của Leonardo (màu xanh rừng, logo, thông tin công ty, phông Tahoma, quy tắc chữ và số). Dùng skill này mỗi khi tạo một file Excel hoặc Word cho Leonardo — báo cáo, bảng tổng hợp, bảng theo dõi, kế hoạch, quy chế, quy trình, hợp đồng, tờ trình, biên bản, thông báo, thư gửi đối tác — và mỗi khi được yêu cầu "trình bày lại", "làm đẹp", "chuẩn hoá", "soát chính tả", "theo chuẩn Leonardo", "theo mẫu công ty", "theo mẫu tháng 9", "theo mẫu DCDS", "làm file Word", "xuất ra docx" cho một file .xlsx hoặc .docx có sẵn, kể cả khi người dùng không nhắc tên skill. Cũng dùng khi người dùng đưa báo cáo quản trị hằng tháng, bảng lương, bảng thưởng của kế toán và muốn bản trình bày chuyên nghiệp mà số liệu giữ nguyên tuyệt đối.
---

# Chuẩn trình bày Excel và Word của Leonardo

Skill gồm hai phần dùng chung bảng màu, logo, thông tin công ty và quy tắc chữ. Xác định loại file rồi **đọc hướng dẫn của phần đó trước khi làm**:

| Loại file | Đọc | Làm được |
|---|---|---|
| Excel (`.xlsx`) | `excel/GUIDE.md` | Tạo mới; trình bày lại file có sẵn mà số liệu giữ nguyên; soát chữ và trình bày |
| Word (`.docx`) | `word/GUIDE.md` | Tạo văn bản mới (ngắn, hoặc dài có bìa và mục lục); trình bày lại file có sẵn mà chữ giữ nguyên |

Bộ chuẩn bằng chữ: `excel/STYLE.md` và `word/STYLE.md`. Thông tin công ty và logo: `excel/config.json`, `excel/assets/`.

Python của skill: `~/.local/share/leonardo-office-venv/bin/python` (do `install.sh` tạo; có openpyxl, pillow, python-docx). Luôn chạy với `-I`, và chép file của người dùng vào một thư mục tạm riêng trước khi xử lý.

## Ràng buộc chung cho cả hai phần

- **Không ghi đè file của người dùng.** Kết quả luôn là file mới đặt cạnh file gốc.
- **Không đổi số, không đổi chữ của bản gốc** khi trình bày lại; luôn chạy bước tự kiểm tra của từng phần và chỉ giao khi đạt.
- **Không tự điền thông tin không có nguồn**: người lập, người duyệt, người ký, ngày lập, kỳ báo cáo, số hiệu văn bản. File gốc hoặc yêu cầu không ghi rõ thì **hỏi lại người dùng**; không suy ra từ chức danh thông thường, tên file hay người đang trò chuyện. Chưa có câu trả lời thì để trống và nói rõ.
- **Không gắn logo và thông tin Leonardo lên văn bản của tổ chức khác**, trừ khi người dùng yêu cầu rõ; khi đó nhắc rằng bản đó chỉ nên dùng nội bộ.
- **Dữ liệu tài chính, lương, thưởng là mật:** xử lý trên máy, không gửi nội dung file ra dịch vụ bên ngoài.
- **Nói thật về việc kiểm tra:** script không mở được Excel hay Word. Báo rõ phần nào đã xem bằng ảnh, phần nào chưa, và nhắc người dùng mở file để xem lần cuối.

## Báo cáo kết quả

Bằng tiếng Việt, kết luận trước: file nằm ở đâu, kết quả kiểm tra (số ô lệch, số đoạn khớp), những gì bị bỏ qua hoặc để trống và vì sao, những gì chưa được xem bằng mắt.
