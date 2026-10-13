---
name: leonardo-excel
description: Bộ chuẩn trình bày file Excel của Leonardo (màu xanh rừng, logo, khối thông tin công ty, bố cục STT / Nội dung / Mã chỉ tiêu, quy tắc chữ và định dạng số). Dùng skill này mỗi khi tạo một file Excel mới cho Leonardo (báo cáo, bảng tổng hợp, bảng theo dõi, kế hoạch, bảng tính gửi sếp hoặc gửi đối tác), và mỗi khi được yêu cầu "trình bày lại", "làm đẹp", "áp theme", "đổi màu", "chuẩn hoá", "theo mẫu tháng 9", "theo chuẩn Leonardo", "theo mẫu DCDS" cho một file Excel có sẵn — kể cả khi người dùng không nhắc tên skill. Cũng dùng khi người dùng đưa file báo cáo quản trị hằng tháng của kế toán và muốn bản trình bày chuyên nghiệp mà số liệu giữ nguyên tuyệt đối.
---

# Chuẩn trình bày Excel của Leonardo

Skill có hai luồng việc. Chọn luồng trước khi làm:

| Tình huống | Luồng | Script |
|---|---|---|
| Chưa có file, cần dựng một bảng/báo cáo mới | **Tạo mới** | `new_report.py` |
| Đã có file Excel (thường của kế toán), cần trình bày lại mà không đổi số | **Trình bày lại** | `restyle_build.py` → `restyle_merge.py` → `verify.py` |

Bộ chuẩn bằng chữ (màu, phông, bố cục, quy tắc chữ) nằm ở `STYLE.md`. Đọc file đó khi cần tự viết mã thay vì dùng script, hoặc khi người dùng hỏi về chuẩn.

Python của skill: `~/.local/share/leonardo-excel-venv/bin/python` (do `install.sh` tạo, có sẵn openpyxl và pillow). Gọi tắt là `$PY` bên dưới; thư mục script là `scripts/` cạnh file này. Luôn chạy với `-I`, và chép file của người dùng vào một thư mục tạm riêng trước khi xử lý.

## Luồng 1 — Tạo mới

1. Hỏi hoặc suy ra: tên báo cáo, kỳ, đơn vị tiền, các bảng cần có. Không bịa số; số liệu phải đến từ người dùng hoặc từ nguồn họ chỉ ra.
2. Viết bản mô tả JSON theo mẫu `examples/vi-du.json` (cấu trúc đầy đủ ở phần đầu `scripts/new_report.py`).
   - Mọi số tổng, chênh lệch, tỷ lệ là **công thức** (`"=E{r}-D{r}"`), không gõ kết quả. `{r}` là dòng hiện tại, `{id}` là dòng có `"id"` đó.
   - Tên sheet dạng `NN Tên` (ví dụ `01 Doanh thu`), không quá 31 ký tự.
   - Tiêu đề khối sẽ tự in hoa. Tiêu đề cột và nhãn dòng chỉ viết hoa chữ đầu; dòng tổng lớn (`"kind": "group"`) viết in hoa.
3. Chạy: `$PY -I scripts/new_report.py mo-ta.json ket-qua.xlsx`. Script không ghi đè file đã có.
4. Xem trước: `$PY -I scripts/preview.py ket-qua.xlsx thu-muc-anh` rồi đọc ảnh. Ô công thức hiện 0 trong ảnh là bình thường (Excel sẽ tính khi mở).

File cần bố cục mà JSON không tả được (biểu đồ, bảng lệch chuẩn): viết mã Python ngắn dùng lớp `Sheet` trong `scripts/leo_style.py`, không gõ lại mã màu hay phông ở nơi khác.

## Luồng 2 — Trình bày lại file có sẵn

Nguyên lý: không đụng vào sheet gốc. Dựng sheet trình bày mới, mỗi ô số là công thức trỏ về đúng ô nguồn, rồi ghép các sheet đó vào file gốc ở mức XML. Sheet gốc được ẩn đi và vẫn là nơi tính toán.

1. Chép file gốc vào thư mục tạm. Xem danh sách sheet và chọn các sheet báo cáo cần trình bày (không chọn sheet dữ liệu thô hàng chục nghìn dòng).
2. Dựng sheet trình bày:
   ```
   $PY -I scripts/restyle_build.py goc.xlsx W.xlsx \
     --sheet "Tên sheet nguồn::NN Tên mới::TIÊU ĐỀ BÁO CÁO — TÊN PHẦN::Tháng MM năm YYYY" \
     --sheet "…" --period "Từ 01/MM/YYYY đến DD/MM/YYYY" [--unit "đồng"] [--report-date "…"]
   ```
   Với bảng dạng danh sách (nhân viên, khách hàng, mã hàng): thêm `--label-col C --code-col B --code-title "Mã nhân viên"` để cột tên làm cột nội dung và mã có sẵn trong file gốc thay cho mã tự đánh số.
   Đọc dòng "BỎ QUA" trong kết quả: đó là các dòng nguồn không đưa vào bản trình bày; phải báo lại cho người dùng.
3. Ghép vào file gốc, ra file mới: `$PY -I scripts/restyle_merge.py goc.xlsx W.xlsx ket-qua.xlsx`
4. Kiểm tra (bắt buộc): `$PY -I scripts/verify.py goc.xlsx ket-qua.xlsx`. Chỉ giao file khi dòng cuối là `KẾT LUẬN: ĐẠT`.
5. Xem trước bằng `preview.py`, đọc ảnh các sheet chính.
6. Đặt file kết quả cạnh file gốc với tên mới (ví dụ thêm `_trinh-bay-moi`). Không ghi đè file gốc.

## Ràng buộc cứng

Mỗi điều dưới đây đến từ một sự cố thật; đừng bỏ qua vì file "trông đơn giản".

- **Không mở rồi lưu lại file gốc bằng openpyxl.** Việc đó làm mất giá trị đã tính của công thức, có thể hỏng biểu đồ, bảng tổng hợp động và liên kết ngoài. Chỉ `restyle_merge.py` được chạm vào file gốc, và nó chỉ thêm phần mới.
- **Không ghi đè file gốc**, trừ khi người dùng yêu cầu rõ; khi đó sao lưu trước và nói rõ bản sao lưu nằm ở đâu.
- **Không chạy lại một bước vá lên file đã vá.** Muốn làm lại thì bắt đầu từ file gốc. Vá chồng từng tạo ra khai báo trùng trong biểu đồ khiến Excel không mở được file.
- **Không đổi số, không làm tròn số nguồn, không đổi đơn vị.** Nếu bản gốc tính theo đồng thì bản trình bày ghi "Đơn vị: đồng".
- **Không tự điền thông tin không có nguồn** (ngày lập báo cáo, người ký): để trống và nói rõ.
- **Dữ liệu tài chính và lương là mật:** không gửi nội dung file ra dịch vụ bên ngoài.
- **Nói thật về việc kiểm tra:** script không mở được Excel. Trong báo cáo cuối phải nêu sheet nào đã xem bằng ảnh, sheet nào chưa, và nhắc người dùng mở thử bằng Excel. Nếu file đang mở trong Excel (có file `~$…` bên cạnh) thì nhắc họ đóng không lưu rồi mở lại.

## Báo cáo kết quả cho người dùng

Bằng tiếng Việt, kết luận trước: file ở đâu, số ô liên kết và số ô lệch (luồng 2), những dòng bị bỏ qua và vì sao, những gì chưa được xem bằng mắt, và việc người dùng cần làm tiếp.
