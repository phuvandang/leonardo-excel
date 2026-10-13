# leonardo-excel — chuẩn trình bày Excel của Leonardo cho Claude Code

Skill giúp Claude tạo và trình bày lại file Excel theo một chuẩn duy nhất của Leonardo: màu xanh rừng, logo, khối thông tin công ty, bố cục STT / Nội dung / Mã chỉ tiêu, quy tắc chữ và định dạng số.

## Cài đặt

Cần: macOS, Claude Code, python3.

```bash
git clone <đường dẫn kho> leonardo-excel-package
cd leonardo-excel-package
bash install.sh
```

Cập nhật về sau: `git pull && bash install.sh`. Bản skill cũ trên máy được đổi tên giữ lại, không bị xoá.

## Cách dùng

Mở Claude Code và nói bằng lời thường, ví dụ:

- "Tạo file Excel theo chuẩn Leonardo: bảng doanh thu theo kênh tháng 10, số liệu như sau…"
- "Trình bày lại file `~/Downloads/bao-cao-thang-10.xlsx` theo chuẩn Leonardo, giữ nguyên số liệu."

Claude sẽ tự chọn một trong hai luồng:

| Luồng | Khi nào | Kết quả |
|---|---|---|
| Tạo mới | Chưa có file | Một file `.xlsx` mới, số tổng và tỷ lệ là công thức |
| Trình bày lại | Đã có file (thường của kế toán) | File mới cạnh file gốc: sheet trình bày ở đầu, mỗi ô số trỏ về ô gốc; sheet gốc được ẩn, không bị sửa |

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
  assets/         logo
  scripts/        leo_style.py (thư viện chuẩn), new_report.py, restyle_build.py, restyle_merge.py, verify.py, preview.py
  examples/       vi-du.json: bản mô tả mẫu cho luồng tạo mới (số giả)
```

## Giới hạn hiện tại

- Luồng tạo mới chưa dựng biểu đồ từ bản mô tả; cần biểu đồ thì Claude viết mã riêng dùng `leo_style.py`.
- Luồng trình bày lại không dựng lại biểu đồ của sheet gốc; biểu đồ vẫn nằm ở sheet gốc đang ẩn.
- Chưa có bước soát chính tả tự động; quy tắc chữ nằm trong `STYLE.md` để Claude tuân theo khi viết.
- Mã chỉ tiêu được đánh số tự động theo sheet, chưa phải bảng mã cố định.
