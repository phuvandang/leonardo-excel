"""Dựng ảnh xem trước từng sheet đang hiện bằng Quick Look (chỉ macOS). Cách dùng: python preview.py file.xlsx thu-muc-anh [số sheet tối đa]

Ảnh chỉ cho thấy phần đầu mỗi sheet, không vẽ biểu đồ, không tính công thức chưa có giá trị lưu, và không giống Excel tuyệt đối.
Dùng để bắt lỗi thô (chữ bị cắt, cột lệch, thiếu logo), không thay cho việc mở bằng Excel.
"""
import sys, zipfile, re, copy, os, shutil, subprocess, html
src, outdir = sys.argv[1], sys.argv[2]; limit = int(sys.argv[3]) if len(sys.argv) > 3 else 12
if not shutil.which('qlmanage'): sys.exit('Không có Quick Look (qlmanage) trên máy này: bỏ qua bước xem trước và nói rõ với người dùng là chưa kiểm tra bằng mắt.')
os.makedirs(outdir, exist_ok=True); z = zipfile.ZipFile(src); wb = z.read('xl/workbook.xml').decode()
sheets = re.findall(r'<sheet [^>]*/>', wb); vis = [s for s in sheets if 'idden"' not in s][:limit]
for i, t in enumerate(vis, 1):
    name = html.unescape(re.search(r'name="([^"]*)"', t).group(1)); tmp = os.path.join(outdir, f'{i:02d}.xlsx')
    new = re.sub(r'<sheets>.*</sheets>', '<sheets>' + t + ''.join(s for s in sheets if s != t) + '</sheets>', wb, flags=re.S)
    new = re.sub(r'<definedNames>.*?</definedNames>', '', new, flags=re.S)
    o = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for info in z.infolist(): o.writestr(copy.copy(info), new.encode() if info.filename == 'xl/workbook.xml' else z.read(info.filename))
    o.close()
    subprocess.run(['qlmanage', '-t', '-s', '2400', '-o', outdir, tmp], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(tmp); png = tmp + '.png'
    print(f'{name}: {png}' if os.path.exists(png) else f'{name}: KHÔNG dựng được ảnh')
