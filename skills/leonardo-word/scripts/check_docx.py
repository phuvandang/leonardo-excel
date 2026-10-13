"""Kiểm tra cấu trúc một file .docx trước khi giao. Cách dùng: python check_docx.py file.docx [goc.docx]

1. File nén và mọi phần XML hợp lệ; các phần tử thuộc tính (bảng, ô, đoạn, dòng) đúng thứ tự lược đồ — sai thứ tự là nguyên nhân Word báo file hỏng.
2. Thống kê: số đoạn, bảng, tiêu đề từng cấp, ảnh; phông chữ thực dùng.
3. Nếu có file gốc: so toàn bộ chữ (theo thứ tự đọc) với bản gốc và in những chỗ khác.
"""
import sys, zipfile, re, difflib, os
from xml.dom import minidom
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from leo_doc import ORDER
z = zipfile.ZipFile(sys.argv[1]); fail = []
if z.testzip() is not None: fail.append('file nén hỏng')
parts = [n for n in z.namelist() if n.endswith('.xml') or n.endswith('.rels')]
for n in parts:
    try: minidom.parseString(z.read(n))
    except Exception as e: fail.append(f'XML lỗi: {n}: {e}')
nbad = 0
for n in [p for p in z.namelist() if re.match(r'word/(document|header\d*|footer\d*|styles)\.xml', p)]:
    dom = minidom.parseString(z.read(n))
    for tag, order in ORDER.items():
        for el in dom.getElementsByTagName('w:' + tag):
            kids = [c.tagName[2:] for c in el.childNodes if c.nodeType == 1 and c.tagName.startswith('w:') and c.tagName[2:] in order]
            idx = [order.index(k) for k in kids]
            if idx != sorted(idx) or len(set(kids)) != len(kids):
                nbad += 1
                if nbad <= 5: fail.append(f'{n}: <w:{tag}> sai thứ tự hoặc trùng phần tử con: {kids}')
doc = z.read('word/document.xml').decode()
text_of = lambda xml: [''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)) for p in re.findall(r'<w:p[ >].*?</w:p>', xml, re.S)]
paras = [t for t in text_of(doc) if t.strip()]
heads = {k: len(re.findall(rf'<w:pStyle w:val="Heading{k}"', doc)) for k in (1, 2, 3)}
fonts = sorted(set(re.findall(r'w:ascii="([^"]+)"', doc + z.read('word/styles.xml').decode())))
print(f'1. Phần XML: {len(parts)}, lỗi thứ tự phần tử: {nbad}')
print(f"2. Đoạn có chữ: {len(paras)} | bảng: {doc.count('<w:tbl>')} | tiêu đề cấp 1/2/3: {heads[1]}/{heads[2]}/{heads[3]} | ảnh: {len([n for n in z.namelist() if n.startswith('word/media/')])} | phông: {fonts}")
if len(sys.argv) > 2:
    old = [t for t in text_of(zipfile.ZipFile(sys.argv[2]).read('word/document.xml').decode()) if t.strip()]
    norm = lambda L: [' '.join(t.split()) for t in L]
    d = [l for l in difflib.unified_diff(norm(old), norm(paras), lineterm='', n=0) if l[:1] in '+-' and l[:3] not in ('+++', '---')]
    print(f'3. So chữ với bản gốc: {len(old)} đoạn → {len(paras)} đoạn; dòng khác: {len(d)}')
    for l in d[:20]: print('   ' + l[:160])
    if d: fail.append('chữ khác bản gốc (xem mục 3; chỉ chấp nhận nếu đó là sửa có chủ ý)')
print('KẾT LUẬN:', 'KHÔNG ĐẠT — ' + '; '.join(fail) if fail else 'ĐẠT')
sys.exit(1 if fail else 0)
