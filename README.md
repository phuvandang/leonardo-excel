# leonardo-excel — chuẩn trình bày Excel và Word của Leonardo cho Claude Code

Skill giúp Claude tạo và trình bày lại file Excel theo một chuẩn duy nhất của Leonardo: màu xanh rừng, logo, khối thông tin công ty, bố cục STT / Nội dung / Mã chỉ tiêu, quy tắc chữ và định dạng số.

## Cài đặt

Cần: macOS, Claude Code, python3.

```bash
git clone https://github.com/phuvandang/leonardo-excel.git leonardo-excel-package
cd leonardo-excel-package
bash install.sh
```

Cập nhật về sau: `git pull && bash install.sh`. Bản skill cũ trên máy được đổi tên giữ lại, không bị xoá.

## Cách dùng

Mở Claude Code và nói bằng lời thường, ví dụ:

- "Tạo file Excel theo chuẩn Leonardo: bảng doanh thu theo kênh tháng 10, số liệu như sau…"
- "Trình bày lại file `~/Downloads/bao-cao-thang-10.xlsx` theo chuẩn Leonardo, giữ nguyên số liệu."
- "Soát lỗi chính tả, viết tắt và trình bày của file này theo chuẩn Leonardo."

Claude sẽ tự chọn trong ba luồng:

| Luồng | Khi nào | Kết quả |
|---|---|---|
| Tạo mới | Chưa có file | Một file `.xlsx` mới, số tổng và tỷ lệ là công thức |
| Trình bày lại | Đã có file (thường của kế toán) | File mới cạnh file gốc: sheet trình bày ở đầu, mỗi ô số trỏ về ô gốc; sheet gốc được ẩn, không bị sửa |
| Soát chữ và trình bày | File đã đúng bố cục, cần sạch chữ | File mới đã sửa lỗi chính tả, tên riêng, dấu thanh, viết tắt, cách ghi tháng, chữ bị cắt; kèm danh sách những chỗ cần anh/chị quyết định |

## Văn bản Word

Gói này cài thêm skill `leonardo-word`: soạn file Word theo cùng bộ nhận diện (màu, logo, thông tin công ty, quy tắc chữ). Nói với Claude, ví dụ: "Soạn quy trình thanh toán thành file Word theo chuẩn Leonardo".

Mỗi văn bản có: đầu trang có logo, tên văn bản, bảng kiểm soát tài liệu, tiêu đề đánh số ba cấp, một kiểu bảng duy nhất, khối ký, chân trang có số trang. Bộ chuẩn và căn cứ tham khảo nằm ở `skills/leonardo-word/STYLE.md`; văn bản minh hoạ tạo từ `skills/leonardo-word/examples/tieu-chuan.json`.

Hỗ trợ tạo văn bản mới và trình bày lại file Word có sẵn (chép nguyên văn chữ và số, tự so trước và sau). Soát chữ trong file Word chưa có lệnh riêng.

## Những điều cần biết

- **File gốc không bị ghi đè.** Kết quả luôn là file mới.
- **Số liệu không đổi.** Ở luồng trình bày lại, Claude chạy bước kiểm tra so từng ô với bản gốc và báo số ô lệch (phải bằng 0).
- **Claude không mở được Excel.** Ảnh xem trước dựng bằng Quick Look, không vẽ biểu đồ và không giống Excel tuyệt đối. Luôn mở file bằng Excel để xem lần cuối.
- **Dữ liệu tài chính là mật.** Skill chạy hoàn toàn trên máy; không đưa file báo cáo thật vào kho này.

## Đổi thông tin công ty

Sửa `~/.claude/skills/leonardo-excel/config.json` (tên công ty, mã số thuế, địa chỉ, người ký, đơn vị mặc định, logo). Muốn đổi cho cả team thì sửa `skills/leonardo-excel/config.json` trong kho rồi mọi người cài lại.

## Cấu trúc

```
skills/leonardo-excel/
  SKILL.md        hướng dẫn cho Claude: hai luồng việc, ràng buộc cứng
  STYLE.md        bộ chuẩn bằng chữ: màu, phông, bố cục, quy tắc chữ và số
  config.json     thông tin công ty
  rules.json      bảng quy tắc soát chữ: lỗi gõ, tên riêng, chữ tắt (team bổ sung dần)
  assets/         logo
  scripts/        leo_style.py (thư viện chuẩn), new_report.py, restyle_build.py, restyle_merge.py, verify.py, proofread.py, preview.py
  examples/       vi-du.json: bản mô tả mẫu cho luồng tạo mới (số giả)
```

## Giới hạn hiện tại

- Luồng tạo mới chưa dựng biểu đồ từ bản mô tả; cần biểu đồ thì Claude viết mã riêng dùng `leo_style.py`.
- Luồng trình bày lại không dựng lại biểu đồ của sheet gốc; biểu đồ vẫn nằm ở sheet gốc đang ẩn.
- Bước soát chữ chỉ tự sửa lỗi có trong `rules.json` (lỗi gõ đã biết, tên riêng, dấu thanh, cách ghi tháng, bảng chữ tắt). Câu cú và chữ tắt lạ do Claude đọc và đề xuất; chỗ không chắc sẽ được hỏi lại.
- Bước soát chữ không sửa chữ hiện ra từ công thức và không sửa tiêu đề, chú thích biểu đồ.
- Mã chỉ tiêu được đánh số tự động theo sheet, chưa phải bảng mã cố định.

## Bổ sung quy tắc soát chữ

Gặp lỗi gõ, tên riêng hoặc chữ tắt mới: thêm vào `skills/leonardo-excel/rules.json`, commit, đẩy lên; mọi người `git pull && bash install.sh`. Chữ tắt chỉ đưa vào bảng khi nó có một nghĩa duy nhất trong tài liệu của Leonardo.
