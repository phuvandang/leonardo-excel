"""Tạo văn bản Word theo chuẩn Leonardo từ bản mô tả JSON. Cách dùng: python new_doc.py mo-ta.json ket-qua.docx

Bản mô tả (xem examples/tieu-chuan.json):
  { "type": "QUY TRÌNH", "title": "…", "subtitle": "…", "code": "QT-KT-01", "header_title": "tên ngắn hiện ở đầu trang",
    "control": {"Mã tài liệu": "…", "Phiên bản": "1.0", "Ngày hiệu lực": "…", "Người soạn": "…", "Người duyệt": "…", "Phạm vi": "…"},
    "body": [ {"h1": "…"}, {"h2": "…"}, {"h3": "…"}, {"p": "đoạn văn, **chữ đậm**"}, {"bullets": ["…"]}, {"numbered": ["…"]},
              {"table": {"caption": "…", "columns": ["…"], "rows": [["…", 1200]], "widths": [3, 1], "align": ["l", "r"], "total_row": true}},
              {"note": "…", "label": "Lưu ý"}, {"pagebreak": true} ],
    "place_date": "TP. Hồ Chí Minh, ngày … tháng … năm …", "signers": [["Người lập", "Họ tên"], ["Người duyệt", "Họ tên"]] }

Tiêu đề mục được đánh số tự động (1. / 1.1 / 1.1.1); không gõ số vào chữ. Số trong bảng để dạng số (không để chuỗi) để được canh phải và ngăn nghìn.
"""
import json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from leo_doc import Doc

spec = json.load(open(sys.argv[1], encoding='utf-8')); out = sys.argv[2]
assert not os.path.exists(out), f'File đã tồn tại, không ghi đè: {out}'
d = Doc(header_title=spec.get('header_title') or spec.get('title', ''), code=spec.get('code', ''))
d.title_block(spec.get('type'), spec['title'], spec.get('subtitle'), spec.get('control'))
for b in spec.get('body', []):
    if 'h1' in b: d.h1(b['h1'])
    elif 'h2' in b: d.h2(b['h2'])
    elif 'h3' in b: d.h3(b['h3'])
    elif 'p' in b: d.p(b['p'], align=b.get('align', 'justify'))
    elif 'bullets' in b: d.bullets(b['bullets'])
    elif 'numbered' in b: d.bullets(b['numbered'], numbered=True)
    elif 'table' in b: t = b['table']; d.table(t['columns'], t['rows'], t.get('caption'), t.get('widths'), t.get('align'), t.get('total_row', False))
    elif 'note' in b: d.note(b['note'], b.get('label', 'Lưu ý'))
    elif b.get('pagebreak'): d.page_break()
    else: sys.exit(f'Khối không hiểu: {b}')
if spec.get('signers'): d.signers(spec['signers'], spec.get('place_date'))
d.save(out); print('Đã tạo', out)
