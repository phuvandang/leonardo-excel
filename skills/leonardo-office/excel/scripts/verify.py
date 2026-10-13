"""Kiểm tra file sau khi trình bày lại. Cách dùng: python verify.py goc.xlsx ket-qua.xlsx

Ba phép kiểm:
  1. Các phần của file gốc còn nguyên (chỉ được khác: workbook, bảng kiểu, danh mục phần, cờ "đang chọn").
  2. Mỗi ô liên kết ở sheet trình bày bằng đúng giá trị đã lưu của ô nguồn.
  3. Độ phủ: ô số nào của sheet nguồn chưa xuất hiện trong bản trình bày.
Kết thúc bằng dòng KẾT LUẬN: ĐẠT hoặc KHÔNG ĐẠT (mã thoát 1).
"""
import sys, zipfile, re, html, collections
from xml.dom import minidom
a, b = zipfile.ZipFile(sys.argv[1]), zipfile.ZipFile(sys.argv[2]); fail = []
if b.testzip() is not None: fail.append('file nén hỏng')
an, bn = set(a.namelist()), set(b.namelist())
ALLOWED = {'[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml'}
print('1. Phần bị bỏ:', sorted(an - bn), '| phần thêm:', len(bn - an))
if (an - bn) - {'xl/calcChain.xml'}: fail.append('mất phần của file gốc')
for n in sorted(an & bn):
    if a.getinfo(n).CRC == b.getinfo(n).CRC or n in ALLOWED: continue
    if n.startswith('xl/worksheets/') and a.read(n).replace(b' tabSelected="1"', b'', 1) == b.read(n): print('   ', n, ': chỉ khác cờ "đang chọn"'); continue
    fail.append(f'phần gốc bị đổi: {n}')
xf = lambda z: re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', re.search(r'<cellXfs.*?</cellXfs>', z.read('xl/styles.xml').decode(), re.S).group(0), re.S)
xa, xb = xf(a), xf(b)
if xb[:len(xa)] != xa: fail.append('kiểu ô cũ bị đổi chỉ số')
for n in sorted(bn - an) + sorted(ALLOWED):
    if n.endswith(('.xml', '.rels')):
        try: minidom.parseString(b.read(n))
        except Exception as e: fail.append(f'XML lỗi: {n}: {e}')
rels = {}
for m in re.finditer(r'<Relationship ([^>]*)/>', b.read('xl/_rels/workbook.xml.rels').decode()):
    d = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); rels[d['Id']] = 'xl/' + d['Target'].lstrip('/').replace('xl/', '')
sheets = []
for m in re.finditer(r'<sheet ([^>]*)/>', b.read('xl/workbook.xml').decode()):
    d = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1))); sheets.append((html.unescape(d['name']), rels[d['r:id']], d.get('state', '')))
print('   Sheet:', [n + (' (ẩn)' if s else '') for n, p, s in sheets])
part = {n: p for n, p, s in sheets}
ss = re.findall(r'<si>(.*?)</si>', a.read('xl/sharedStrings.xml').decode(), re.S) if 'xl/sharedStrings.xml' in an else []
sst = lambda i: html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', ss[i])))
CELL = re.compile(r'<c ([^>]*?)(?:/>|>(.*?)</c>)', re.S); cache = {}
def src_cells(name):
    if name not in cache:
        d = {}
        for m in CELL.finditer(a.read(part[name]).decode()):
            at = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); v = re.search(r'<v>([^<]*)</v>', m.group(2) or '')
            it = re.findall(r'<t[^>]*>([^<]*)</t>', m.group(2) or '') if at.get('t') == 'inlineStr' else None
            d[at['r']] = (at.get('t', 'n'), v.group(1) if v else None, it)
        cache[name] = d
    return cache[name]
tot = bad = left = 0; linked = collections.defaultdict(set)
for n, p, s in sheets:
    if p in an: continue
    x = b.read(p).decode(); left += x.count('t="s"')
    for m in CELL.finditer(x):
        body = m.group(2) or ''; f = re.search(r'<f>([^<]*)</f>', body)
        if not f: continue
        fm = re.fullmatch(r"'((?:[^']|'')+)'!([A-Z]+\d+)", html.unescape(f.group(1)))
        if not fm: continue
        sn = fm.group(1).replace("''", "'"); t, v, it = src_cells(sn).get(fm.group(2), ('n', None, None)); linked[sn].add(fm.group(2))
        at = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); nv = re.search(r'<v>([^<]*)</v>', body); tot += 1
        if nv is None or v is None and t != 'inlineStr': ok = False
        elif t == 's': ok = at.get('t') == 'str' and html.unescape(nv.group(1)) == sst(int(v))
        elif t in ('str', 'e', 'b'): ok = html.unescape(nv.group(1)) == html.unescape(v)
        elif t == 'inlineStr': ok = html.unescape(nv.group(1)) == html.unescape(''.join(it))
        else: ok = at.get('t') is None and float(nv.group(1)) == float(v)
        if not ok:
            bad += 1
            if bad <= 8: print('   LỆCH', n, at['r'], html.unescape(f.group(1)))
print(f'2. Ô liên kết: {tot}, lệch: {bad}, ô chữ trỏ nhầm bảng chuỗi chung: {left}')
if bad or left or not tot: fail.append('ô liên kết lệch hoặc không có ô liên kết nào')
print('3. Độ phủ ô số của từng sheet nguồn:')
for name in linked:
    nums = [dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))) for m in CELL.finditer(a.read(part[name]).decode()) if '<v>' in (m.group(2) or '')]
    nums = [d['r'] for d in nums if d.get('t', 'n') == 'n']; miss = [r for r in nums if r not in linked[name]]
    by_row = collections.Counter(re.sub(r'[A-Z]+', '', r) for r in miss).most_common(4); by_col = collections.Counter(re.sub(r'\d+', '', r) for r in miss).most_common(3)
    print(f'   {name}: {len(nums)} ô số, đã có {len(nums) - len(miss)}, chưa có {len(miss)}' + (f' (nhiều nhất ở dòng {by_row}, cột {by_col})' if miss else ''))
print('KẾT LUẬN:', 'KHÔNG ĐẠT — ' + '; '.join(fail) if fail else 'ĐẠT (ô chưa có ở mục 3, nếu có, phải được giải thích cho người dùng)')
sys.exit(1 if fail else 0)
