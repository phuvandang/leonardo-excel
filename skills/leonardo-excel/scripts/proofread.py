"""Luồng "soát chữ và trình bày": tìm và sửa lỗi chữ, lỗi hiển thị trong file Excel mà không đổi số, công thức, kiểu ô hay bố cục.

  python proofread.py file.xlsx --dump  [--sheets "A,B"]                 in toàn bộ chữ để đọc
  python proofread.py file.xlsx         [--sheets "A,B"] [--abbrev]      chỉ kiểm tra, in đề xuất và cảnh báo
  python proofread.py file.xlsx --apply ket-qua.xlsx [--abbrev] [--edits sua.json] [--skip-rows 10]

Máy tự sửa (theo rules.json): lỗi gõ đã biết, tên riêng, dấu thanh kiểu cũ (hoà → hòa), khoảng trắng và dấu câu, cách ghi tháng (tháng 8 → tháng 08).
--abbrev: viết đầy đủ chữ tắt trong rules.json (mặc định chỉ liệt kê làm gợi ý, vì một chữ tắt có thể mang nghĩa khác).
--edits: các sửa do người đọc quyết định, dạng [{"sheet": "01 KQKD", "cell": "B14", "text": "Chữ mới"}].
--skip-rows N: không đụng N dòng đầu mỗi sheet (khối đầu trang).

Chỉ sửa ô chữ gõ trực tiếp. Chữ hiện ra từ công thức được liệt kê kèm công thức để sửa tại nguồn.
Ô có nội dung đang được công thức dùng làm chuỗi so khớp thì bị bỏ qua và cảnh báo. Không chạy trên sheet dữ liệu thô.
Khi --apply: ghi ra file mới, nới chiều cao dòng nếu chữ bị cắt, rồi tự so trước/sau; chỉ báo ĐẠT khi công thức và giá trị số giống hệt.
"""
import argparse, copy, html, json, math, os, re, sys, zipfile
from xml.dom import minidom
from xml.sax.saxutils import escape

ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument('src'); ap.add_argument('--sheets'); ap.add_argument('--dump', action='store_true'); ap.add_argument('--apply')
ap.add_argument('--abbrev', action='store_true'); ap.add_argument('--edits'); ap.add_argument('--skip-rows', type=int, default=0)
ap.add_argument('--max-rows', type=int, default=3000, help='bỏ qua sheet dài hơn số dòng này (dữ liệu thô)')
A = ap.parse_args()
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = json.load(open(os.path.join(ROOT, 'rules.json'), encoding='utf-8'))
VN = 'A-Za-zÀ-ỹ'
# ---------- quy tắc
TONE = {}
for a, b in (('oàoáoảoãoạ', 'òaóaỏaõaọa'), ('oèoéoẻoẽoẹ', 'òeóeỏeõeọe'), ('uỳuýuỷuỹuỵ', 'ùyúyủyũyụy')):
    for i in range(0, 10, 2):
        TONE[a[i:i + 2]] = b[i:i + 2]; TONE[a[i:i + 2].upper()] = b[i:i + 2].upper(); TONE[a[i].upper() + a[i + 1]] = b[i].upper() + b[i + 1]
tone_re = re.compile(r'(?<![qQ])(' + '|'.join(map(re.escape, TONE)) + rf')(?![{VN}\d])')
BRANDS = [(re.compile(rf'(?<![{VN}]){p}(?![{VN}])'), r) for p, r in R['brands']]
ABBR = {k: v for k, v in R['abbreviations'].items()}
abbr_re = re.compile(rf'(?<![{VN}\d&])(' + '|'.join(sorted(map(re.escape, ABBR), key=len, reverse=True)) + rf')(?![{VN}\d&])') if ABBR else None


def mechanical(t):
    o = t
    for a, b in R['typos']: t = t.replace(a, b)
    for rx, rep in BRANDS: t = rx.sub(rep, t)
    t = tone_re.sub(lambda m: TONE[m.group(1)], t)
    t = t.replace('…', '...'); t = re.sub(r'(?<!\.)\.\.(?!\.)', '...', t)
    t = re.sub(r'^(\d{1,2}|[A-D])\.(?=[^\s\d.])', r'\1. ', t)                       # "3.Phải thu" → "3. Phải thu"
    t = re.sub(r'\(\s+', '(', t); t = re.sub(r'\s+\)', ')', t)
    t = re.sub(rf'(?<=[{VN}])\((?=[{VN}])', ' (', t) if 'Ctrl' not in t and '=' not in t else t
    t = re.sub(r' +([:;,])(?= |$)', r'\1', t)                                         # bỏ dấu cách trước dấu câu
    t = re.sub(r'(?i)(?<=tháng )(\d)(?!\d)(?!,\d)', r'0\1', t)                       # tháng 8 → tháng 08
    t = re.sub(r'[ \t]{2,}', ' ', t) if not o.startswith(' ') else t                 # giữ thụt đầu dòng có chủ ý
    return t.rstrip() if t.strip() else t


def expand(t):
    def rep(m):
        full = ABBR[m.group(1)]
        if t.isupper(): return full.upper()
        return full[0].upper() + full[1:] if m.start() == 0 or t[m.start() - 1] == '_' else full
    return abbr_re.sub(rep, t) if abbr_re else t


# ---------- đọc file
z = zipfile.ZipFile(A.src); names = z.namelist()
rels = {}
for m in re.finditer(r'<Relationship ([^>]*)/>', z.read('xl/_rels/workbook.xml.rels').decode()):
    d = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); rels[d['Id']] = 'xl/' + d['Target'].lstrip('/').replace('xl/', '')
SHEETS = []
for m in re.finditer(r'<sheet ([^>]*)/>', z.read('xl/workbook.xml').decode()):
    d = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1))); SHEETS.append((html.unescape(d['name']), rels[d['r:id']], 'idden' in d.get('state', '')))
want = [s.strip() for s in A.sheets.split(',')] if A.sheets else None
if want:
    miss = [w for w in want if w not in [n for n, _, _ in SHEETS]]
    if miss: sys.exit(f'Không có sheet: {miss}. Các sheet: {[n for n, _, _ in SHEETS]}')
SS = re.findall(r'<si>(.*?)</si>', z.read('xl/sharedStrings.xml').decode(), re.S) if 'xl/sharedStrings.xml' in names else []
sst = lambda i: html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', SS[i])))
CELL = re.compile(r'<c ([^>]*?)(?:/>|>(.*?)</c>)', re.S); ROW = re.compile(r'<row ([^>]*?)(?<!/)>(.*?)</row>', re.S)
raw = {}; targets = []
for n, p, hidden in SHEETS:
    data = z.read(p).decode(); raw[p] = data; nrows = data.count('<row ')
    if want: ok = n in want
    else: ok = not hidden and nrows <= A.max_rows
    if ok and nrows > A.max_rows and not want: ok = False
    if ok: targets.append((n, p))
    elif not want and not hidden: print(f'(bỏ qua sheet "{n}": {nrows} dòng, coi là dữ liệu thô; dùng --sheets nếu vẫn muốn soát)')
ALLF = [html.unescape(f) for d in raw.values() for f in re.findall(r'<f[^>]*>([^<]*)</f>', d)]
_LK = re.compile(r'\b(MATCH|VLOOKUP|HLOOKUP|XLOOKUP|SUMIFS?|COUNTIFS?|AVERAGEIFS?|INDEX|SEARCH|FIND)\(', re.I)
is_looked = lambda n: any(_LK.search(f) and (f"'{n}'!" in f or f'{n}!' in f) for f in ALLF)
if not want:                                 # sheet đang bị công thức dò chữ: mặc định không soát
    for n, p in list(targets):
        if is_looked(n): targets.remove((n, p)); print(f'(bỏ qua sheet "{n}": đang được công thức dò/so khớp tham chiếu tới; dùng --sheets nếu vẫn muốn soát)')
LIT = set(re.findall(r'"([^"]+)"', ' '.join(ALLF)))                                  # chuỗi đang được công thức dùng để so khớp
LOOKUP = re.compile(r'\b(MATCH|VLOOKUP|HLOOKUP|XLOOKUP|SUMIFS?|COUNTIFS?|AVERAGEIFS?|INDEX|SEARCH|FIND)\(', re.I)


def cell_text(attrs, inner):
    at = dict(re.findall(r'(\w+)="([^"]*)"', attrs)); t = at.get('t'); inner = inner or ''
    if t == 'inlineStr': return at, 'inline', html.unescape(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', inner)))
    if t == 's':
        v = re.search(r'<v>(\d+)</v>', inner); return at, 'shared', sst(int(v.group(1))) if v else ''
    if t == 'str': return at, 'formula', html.unescape((re.search(r'<v>([^<]*)</v>', inner) or [0, ''])[1])
    return at, None, None


rownum = lambda ref: int(re.sub(r'\D', '', ref))
if A.dump:
    for n, p in targets:
        print(f'===== {n}')
        for rm in ROW.finditer(raw[p]):
            parts = []
            for cm in CELL.finditer(rm.group(2)):
                at, kind, txt = cell_text(cm.group(1), cm.group(2))
                if kind and txt.strip():
                    f = re.search(r'<f[^>]*>([^<]*)</f>', cm.group(2) or '')
                    parts.append(f"{at['r']}{'=' if kind == 'formula' else ''}:{txt}" + (f"  ⟵ {html.unescape(f.group(1))[:60]}" if kind == 'formula' and f else ''))
            if parts: print(' | '.join(parts))
    sys.exit(0)

EDITS = {}
if A.edits:
    for e in json.load(open(A.edits, encoding='utf-8')): EDITS[(e['sheet'], e['cell'].upper())] = e['text']
# ---------- soát
out = {}; log = {}; suggest = {}; guarded = []; fromf = []; used_edits = set()
for n, p in targets:
    data = raw[p]
    looked = any(LOOKUP.search(f) and (f"'{n}'!" in f or f'{n}!' in f) for f in ALLF)
    if looked: print(f'⚠ Sheet "{n}" đang được công thức dò/so khớp tham chiếu tới: ô nào thuộc vùng dò thì không được sửa chữ. Xem kỹ danh sách thay đổi.')

    def fix(cm):
        attrs, inner = cm.group(1), cm.group(2)
        at, kind, old = cell_text(attrs, inner)
        if not kind or not old.strip(): return cm.group(0)
        ref = at['r']
        if rownum(ref) <= A.skip_rows: return cm.group(0)
        if kind == 'formula':
            new = mechanical(old)
            if new != old or (n, ref) in EDITS: fromf.append((n, ref, old, html.unescape((re.search(r'<f[^>]*>([^<]*)</f>', inner or '') or [0, '?'])[1])[:70]))
            return cm.group(0)
        new = EDITS.get((n, ref))
        if new is not None: used_edits.add((n, ref)); cat = 'người sửa'
        else:
            new = mechanical(old); cat = 'máy sửa'
            ex = expand(new)
            if ex != new:
                if A.abbrev: new = ex; cat = 'máy sửa + viết tắt'
                else: suggest.setdefault((new, ex), []).append(f'{n}!{ref}')
        if new == old: return cm.group(0)
        if old in LIT: guarded.append((n, ref, old)); return cm.group(0)
        log.setdefault((old, new, cat), []).append(f'{n}!{ref}')
        a2 = re.sub(r' t="[a-zA-Z]+"', '', attrs).rstrip() + ' t="inlineStr"'
        sp = ' xml:space="preserve"' if new != new.strip() or '  ' in new else ''
        return f'<c {a2}><is><t{sp}>{escape(new)}</t></is></c>'
    out[p] = CELL.sub(fix, data)
for key in EDITS:
    if key not in used_edits: print(f'⚠ Mục trong --edits không khớp ô chữ gõ trực tiếp nào: {key}')
# ---------- hiển thị: chiều cao dòng và cảnh báo
styles = z.read('xl/styles.xml').decode()
xfl = re.findall(r'<xf [^>]*?(?:/>|>.*?</xf>)', re.search(r'<cellXfs[^>]*>(.*?)</cellXfs>', styles, re.S).group(1), re.S)
wrap = ['wrapText="1"' in x or 'wrapText="true"' in x for x in xfl]
fonts = re.findall(r'<font>.*?</font>|<font/>', re.search(r'<fonts[^>]*>(.*?)</fonts>', styles, re.S).group(1), re.S)
size = [float((re.search(r'<sz val="([\d.]+)"', fonts[min(int((re.search(r'fontId="(\d+)"', x) or [0, '0'])[1]), len(fonts) - 1)]) or [0, '10'])[1]) for x in xfl]


def col2n(c):
    k = 0
    for ch in c: k = k * 26 + ord(ch) - 64
    return k


def n2col(k):
    r = ''
    while k: k, m = divmod(k - 1, 26); r = chr(65 + m) + r
    return r


def lines(text, px, sz):
    cap = max(1, int((px - 6) / (6.5 * sz / 10))); k = 0
    for para in text.split('\n'):
        cur = 0; k += 1
        for w in para.split(' '):
            L = len(w)
            if cur and cur + 1 + L > cap: k += 1; cur = 0
            while L > cap: k += 1; L -= cap
            cur = cur + (1 if cur else 0) + L
    return k


hlog = []; warn = []
for n, p in targets:
    s = out[p]; dflt = float((re.search(r'defaultColWidth="([\d.]+)"', s) or [0, '8.43'])[1]); widths = {}
    for c in re.finditer(r'<col ([^>]*)/>', s):
        a = dict(re.findall(r'(\w+)="([^"]*)"', c.group(1)))
        for i in range(int(a['min']), min(int(a['max']), 300) + 1): widths[i] = float(a.get('width', dflt))
    merges = {(r[0], int(r[1])): (col2n(r[2]), int(r[3]) - int(r[1]) + 1) for r in re.findall(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', s)}

    def fit(rm):
        attrs, body = rm.group(1), rm.group(2); rn = int(re.search(r'\br="(\d+)"', attrs).group(1))
        if rn <= A.skip_rows: return rm.group(0)
        hm = re.search(r'\bht="([\d.]+)"', attrs); cur = float(hm.group(1)) if hm else 15.0; need = 0; filled = set(re.findall(r'<c r="([A-Z]+)\d+"[^>]*>(?=.*?<(?:v|t)[ >])', body))
        for cm in CELL.finditer(body):
            at, kind, txt = cell_text(cm.group(1), cm.group(2)); si = min(int(at.get('s', 0)), len(xfl) - 1)
            col = re.match(r'[A-Z]+', at['r']).group(0); c0 = col2n(col); c1, rows = merges.get((col, rn), (c0, 1))
            px = sum(7 * widths.get(i, dflt) + 5 for i in range(c0, c1 + 1))
            if kind is None:
                v = re.search(r'<v>(-?[\d.]+(?:E-?\d+)?)</v>', cm.group(2) or '')
                if v and len(f'{abs(float(v.group(1))):,.0f}') * 7.2 + 14 > px: warn.append(f'{n}!{at["r"]}: số có thể hiện #### (cột hẹp)')
                continue
            if not txt: continue
            if not wrap[si]:
                if len(txt) * 5.6 * size[si] / 10 > px and n2col(c1 + 1) in filled: warn.append(f'{n}!{at["r"]}: chữ bị ô bên phải che ("{txt[:30]}…")')
                continue
            k = lines(txt, px, size[si])
            if k >= 2: need = max(need, (k * size[si] * 1.32 + 5) / rows)
        if need > cur + 0.5:
            new = math.ceil(need); hlog.append((n, rn, cur, new))
            attrs = attrs.replace(hm.group(0), f'ht="{new}"') if hm else attrs + f' ht="{new}"'
            if 'customHeight' not in attrs: attrs += ' customHeight="1"'
            return f'<row {attrs}>{body}</row>'
        return rm.group(0)
    out[p] = ROW.sub(fit, s)
# ---------- báo cáo
short = lambda refs: ', '.join(refs[:3]) + (f' … ({len(refs)} ô)' if len(refs) > 3 else '')
print(f'\nSHEET ĐÃ SOÁT: {[n for n, _ in targets]}')
print(f'\nTHAY ĐỔI ({sum(len(v) for v in log.values())} ô, {len(log)} kiểu):')
for (o, nw, cat), refs in log.items(): print(f'  [{cat}] {short(refs)}\n      - {o}\n      + {nw}')
if suggest:
    print(f'\nGỢI Ý VIẾT TẮT (chưa áp dụng; thêm --abbrev để áp dụng, hoặc đưa mục chọn lọc vào --edits):')
    for (o, nw), refs in suggest.items(): print(f'  {short(refs)}\n      - {o}\n      + {nw}')
if fromf:
    print('\nCHỮ LẤY TỪ CÔNG THỨC CẦN SỬA TẠI NGUỒN (không tự sửa):')
    for n, ref, old, f in fromf[:40]: print(f'  {n}!{ref}: "{old[:60]}"  ⟵ ={f}')
if guarded:
    print('\nBỎ QUA VÌ CHUỖI ĐANG ĐƯỢC CÔNG THỨC DÙNG ĐỂ SO KHỚP:')
    for n, ref, old in guarded[:40]: print(f'  {n}!{ref}: "{old[:60]}"')
print(f'\nCHIỀU CAO DÒNG CẦN NỚI: {len(hlog)}', [(n, f'dòng {r}', f'{a:g}→{b}') for n, r, a, b in hlog[:30]])
print(f'CẢNH BÁO HIỂN THỊ: {len(warn)}' + (' (in 25 mục đầu)' if len(warn) > 25 else '')); [print('  ' + w) for w in warn[:25]]
if not A.apply:
    print('\n(Chưa ghi gì. Dùng --apply ket-qua.xlsx để ghi ra file mới.)'); sys.exit(0)
# ---------- ghi và tự kiểm tra
assert os.path.abspath(A.apply) != os.path.abspath(A.src), 'Không ghi đè file nguồn: chọn tên file kết quả khác.'
assert not os.path.exists(A.apply), f'File đã tồn tại, không ghi đè: {A.apply}'
changed = {p: d for p, d in out.items() if d != raw[p]}
with zipfile.ZipFile(A.apply, 'w', zipfile.ZIP_DEFLATED) as zo:
    for info in zipfile.ZipFile(A.src).infolist():
        zo.writestr(copy.copy(info), changed[info.filename].encode() if info.filename in changed else z.read(info.filename))
b = zipfile.ZipFile(A.apply); fail = []
if b.testzip() is not None or b.namelist() != names: fail.append('cấu trúc file nén khác')
tot = dict(o=0, f=0, n=0, t=0)
for p in names:
    if p not in changed:
        if z.getinfo(p).CRC != b.getinfo(p).CRC: fail.append(f'phần không liên quan bị đổi: {p}')
        continue
    x, y = raw[p], b.read(p).decode()
    try: minidom.parseString(y.encode())
    except Exception as e: fail.append(f'XML lỗi {p}: {e}'); continue
    sx, sy = x.split('<sheetData>'), y.split('<sheetData>')
    if sx[0] != sy[0] or sx[1].split('</sheetData>')[1] != sy[1].split('</sheetData>')[1]: fail.append(f'{p}: phần ngoài dữ liệu ô (cột, ô gộp, định dạng) bị đổi')
    strip = lambda s: re.sub(r' (ht|customHeight)="[^"]*"', '', s)
    if [strip(r) for r in re.findall(r'<row ([^>]*?)(?<!/)>', x)] != [strip(r) for r in re.findall(r'<row ([^>]*?)(?<!/)>', y)]: fail.append(f'{p}: thuộc tính dòng bị đổi')
    cx, cy = CELL.findall(x), CELL.findall(y)
    if len(cx) != len(cy): fail.append(f'{p}: số ô khác'); continue
    for (ax, ix), (ay, iy) in zip(cx, cy):
        tot['o'] += 1; dx, dy = dict(re.findall(r'(\w+)="([^"]*)"', ax)), dict(re.findall(r'(\w+)="([^"]*)"', ay))
        if dx.get('r') != dy.get('r') or dx.get('s') != dy.get('s'): fail.append(f'{p}: ô {dx.get("r")} đổi vị trí hoặc kiểu'); break
        fx, fy = re.findall(r'<f[^>]*>[^<]*</f>|<f[^>]*/>', ix or ''), re.findall(r'<f[^>]*>[^<]*</f>|<f[^>]*/>', iy or '')
        if fx != fy: fail.append(f'{p}: công thức ô {dx["r"]} bị đổi'); break
        tot['f'] += len(fx)
        if dx.get('t') in ('s', 'inlineStr'):
            if (ax, ix) != (ay, iy): tot['t'] += 1
        elif (ax, ix) != (ay, iy): fail.append(f'{p}: ô số/công thức {dx["r"]} bị đổi'); break
        else: tot['n'] += 1
print(f"\nTỰ KIỂM TRA: {tot['o']} ô ở các sheet đã sửa; {tot['f']} công thức và {tot['n']} ô số/ô công thức giống hệt; {tot['t']} ô chữ đổi; {len(hlog)} dòng nới chiều cao.")
if fail:
    os.remove(A.apply); print('KẾT LUẬN: KHÔNG ĐẠT — đã xoá file kết quả:', '; '.join(fail[:6])); sys.exit(1)
print(f'KẾT LUẬN: ĐẠT — đã ghi {A.apply}')
