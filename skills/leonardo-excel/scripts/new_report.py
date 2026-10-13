"""Luồng "tạo mới": dựng file Excel theo chuẩn Leonardo từ một bản mô tả JSON. Cách dùng: python new_report.py mo-ta.json ket-qua.xlsx

Bản mô tả (xem examples/vi-du.json):
  { "period": "...", "report_date": "...", "unit": "đồng", "info": true,
    "sheets": [ { "name": "01 Doanh thu", "title": "BÁO CÁO … — …", "subtitle": "Tháng 10 năm 2026",
                  "label_width": 40, "value_width": 16,
                  "blocks": [ { "title": "Doanh thu theo kênh", "columns": ["Kênh", "Tháng 09", "Tháng 10", "Chênh lệch"],
                                "formats": ["number", "number", "number"],
                                "rows": [ {"id": "shopee", "label": "Shopee", "values": [100, 120, "=E{r}-D{r}"]},
                                          {"label": "Cộng", "kind": "group", "values": ["=SUM(D{shopee}:D{tiktok})", …]} ],
                                "status_column": 4, "note": "Ghi chú: …" } ] } ] }

Quy tắc: số tổng, chênh lệch, tỷ lệ phải là công thức ("=..."), không gõ kết quả. Trong công thức, {r} là dòng hiện tại,
{id} là số dòng của dòng có "id" đó (trong cùng sheet). Cột giá trị đầu tiên là cột D.
kind: item (mặc định) | group (dòng tổng, đậm) | sub (dòng tên nhóm, không có số). Định dạng: number | decimal | percent | date | general | text | status | input.
"""
import json, sys, os, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from leo_style import Sheet, FORMATS, rom

spec = json.load(open(sys.argv[1], encoding='utf-8')); out = sys.argv[2]
assert not os.path.exists(out), f'File đã tồn tại, không ghi đè: {out}'


def build(ids_in):
    wb = Workbook(); wb.remove(wb.active); ids_out = {}
    for k, s in enumerate(spec['sheets'], 1):
        assert len(s['name']) <= 31 and not re.search(r'[\[\]:*?/\\]', s['name']), f"Tên sheet không hợp lệ: {s['name']}"
        ncol = max(len(b['columns']) - 1 for b in s['blocks'])
        sh = Sheet(wb, s['name'], s['title'], [7, s.get('label_width', 44), 11] + [s.get('value_width', 16)] * max(ncol, 3), f'{k:02d}', s.get('subtitle'),
                   unit=s.get('unit', spec.get('unit')), period=spec.get('period'), report_date=spec.get('report_date'), info=spec.get('info', True))
        ids = ids_in.get(s['name'], {}); mine = ids_out.setdefault(s['name'], {})
        sh.sec('A.', s['title'].split('— ')[-1].upper())
        for b in s['blocks']:
            cols = b['columns']; fmts = b.get('formats', [])
            sh.block(b['title'], ncols=len(cols) - 1); sh.header(['STT', cols[0], 'Mã chỉ tiêu'] + cols[1:]); first = sh.r; gi = ii = 0
            for row in b['rows']:
                kind = row.get('kind', 'item'); r = sh.r
                if 'id' in row: mine[row['id']] = r
                vals = []
                for j, v in enumerate(row.get('values', [])):
                    f = (row.get('formats') or fmts + [None] * 99)[j] or 'number'
                    if isinstance(v, str) and v.startswith('='):
                        v = re.sub(r'\{(\w+)\}', lambda m: str(r if m.group(1) == 'r' else ids.get(m.group(1), 0)), v)
                    vals.append((v, 'input' if f == 'input' else ('status' if f == 'status' else (None if f == 'text' else FORMATS[f]))))
                if kind == 'sub': sh.row('', row['label'], [], 'sub'); ii = 0; continue
                if kind == 'group': stt = rom(gi); gi += 1; ii = 0
                else: ii += 1; stt = ii
                sh.row(stt, row['label'], vals, kind)
            if b.get('status_column'): c = CL(3 + b['status_column']); sh.status_cf(f'{c}{first}:{c}{sh.r - 1}')
            if b.get('note'): sh.note(b['note'])
        sh.finish()
    return wb, ids_out


_, ids = build({})          # lượt 1: lấy số dòng của các "id"
wb, _ = build(ids)          # lượt 2: điền công thức
wb.save(out); print('Đã tạo', out, '|', [(ws.title, ws.dimensions) for ws in wb])
