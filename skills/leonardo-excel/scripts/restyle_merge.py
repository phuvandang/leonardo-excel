"""Luồng "trình bày lại", bước 2: ghép các sheet trình bày (file W của restyle_build.py) vào file gốc ở mức XML.

Không mở-lưu file gốc bằng openpyxl, không đổi tên, không sửa ô nào của sheet gốc: chỉ thêm sheet mới lên đầu,
thêm kiểu vào cuối bảng kiểu, bỏ calcChain và (mặc định) ẩn các sheet nguồn. Luôn ghi ra file mới.

Cách dùng: python restyle_merge.py goc.xlsx W.xlsx ket-qua.xlsx [--keep-visible]
"""
import sys, zipfile, re, copy, json, html
from xml.dom import minidom
args = [a for a in sys.argv[1:] if not a.startswith('--')]; KEEP = '--keep-visible' in sys.argv
big, wpath, dst = args[:3]; vjson = wpath + '.values.json'
import os
assert os.path.abspath(dst) != os.path.abspath(big), 'Không ghi đè file gốc: hãy chọn tên file kết quả khác.'
zb = zipfile.ZipFile(big); zw = zipfile.ZipFile(wpath); VAL = json.load(open(vjson))
rb = lambda n: zb.read(n).decode('utf-8'); rw = lambda n: zw.read(n).decode('utf-8')
repl, new, drop = {}, {}, {'xl/calcChain.xml'}
SRC = [] if KEEP else json.load(open(wpath + '.meta.json', encoding='utf-8'))['sources']
wwb = rw('xl/workbook.xml'); NAMES = [html.unescape(x) for x in re.findall(r'<sheet [^>]*?name="([^"]*)"', wwb)]; NN = len(NAMES)
ERR = ('#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!')
esc = lambda s: s.replace('&', '&amp;')
wbx = rb('xl/workbook.xml')
for n in SRC:
    m = re.search(r'<sheet [^>]*?name="%s"[^>]*/>' % re.escape(esc(n)), wbx); assert m, n
    if 'state=' not in m.group(0): wbx = wbx.replace(m.group(0), m.group(0).replace('<sheet ', '<sheet state="hidden" ', 1))
for n in NAMES: assert f'name="{esc(n)}"' not in wbx, n
# bỏ đánh dấu "đang chọn" ở sheet gốc để không bị chọn nhóm cùng sheet mới
for part in [x for x in zb.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', x)]:
    with zb.open(part) as fh: headb = fh.read(20000)
    if b' tabSelected="1"' in headb:
        data = zb.read(part); repl[part] = data.replace(b' tabSelected="1"', b'', 1); print('bỏ tabSelected:', part)
# ---------- gộp styles (chỉ thêm vào cuối, không đổi chỉ số cũ)
st = rb('xl/styles.xml'); ws_ = rw('xl/styles.xml')
def children(x, tag, child):
    m = re.search(rf'<{tag}[ >].*?</{tag}>|<{tag}[^>]*/>', x, re.S); return re.findall(rf'<{child}(?: [^>]*)?/>|<{child}(?: [^>]*)?>.*?</{child}>', m.group(0), re.S) if m else []
count = lambda tag: len(children(st, tag, {'fonts': 'font', 'fills': 'fill', 'borders': 'border', 'cellXfs': 'xf'}[tag]))
nf, nfi, nb, nx = count('fonts'), count('fills'), count('borders'), count('cellXfs')
wf, wfi, wb_, wxf = children(ws_, 'fonts', 'font'), children(ws_, 'fills', 'fill'), children(ws_, 'borders', 'border'), children(ws_, 'cellXfs', 'xf')
wnum = dict(re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', ws_)); maxid = max([163] + [int(i) for i in re.findall(r'<numFmt numFmtId="(\d+)"', st)])
nmap = {}; addnum = ''
for k, (i, code) in enumerate(wnum.items(), 1): nmap[i] = str(maxid + k); addnum += f'<numFmt numFmtId="{maxid + k}" formatCode="{code}"/>'
def fix_xf(x):
    x = re.sub(r'fontId="(\d+)"', lambda m: f'fontId="{int(m.group(1)) + nf}"', x); x = re.sub(r'fillId="(\d+)"', lambda m: f'fillId="{int(m.group(1)) + nfi}"', x)
    x = re.sub(r'borderId="(\d+)"', lambda m: f'borderId="{int(m.group(1)) + nb}"', x); x = re.sub(r'numFmtId="(\d+)"', lambda m: f'numFmtId="{nmap.get(m.group(1), m.group(1))}"', x)
    return re.sub(r'xfId="\d+"', 'xfId="0"', x)
def append(tag, items):
    global st
    m = re.search(rf'<{tag}( count="(\d+)")?', st)
    if m.group(1): st = st.replace(m.group(0), f'<{tag} count="{int(m.group(2)) + len(items)}"', 1)
    k = st.index(f'</{tag}>'); st = st[:k] + ''.join(items) + st[k:]
append('fonts', wf); append('fills', wfi); append('borders', wb_); append('cellXfs', [fix_xf(x) for x in wxf])
m = re.search(r'<numFmts count="(\d+)">', st)
if m: st = st.replace(m.group(0), f'<numFmts count="{int(m.group(1)) + len(nmap)}">').replace('</numFmts>', addnum + '</numFmts>')
else: st = re.sub(r'(<styleSheet[^>]*>)', lambda mm: mm.group(1) + f'<numFmts count="{len(nmap)}">{addnum}</numFmts>', st, count=1)
minidom.parseString(st.encode()); repl['xl/styles.xml'] = st.encode()
print('styles: +fonts', len(wf), '+fills', len(wfi), '+borders', len(wb_), '+xfs', len(wxf), 'từ chỉ số', nx)
# ---------- chuyển sheet, drawing, ảnh
ss = re.findall(r'<si>(.*?)</si>', rw('xl/sharedStrings.xml'), re.S) if 'xl/sharedStrings.xml' in zw.namelist() else []
ct_add = wrels_add = sheets_add = ''
maxsid = max(int(i) for i in re.findall(r'sheetId="(\d+)"', wbx))
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
stats = {'formula': 0, 'cached': 0, 'str': 0}
exist = set(zb.namelist()); wrels0 = rb('xl/_rels/workbook.xml.rels')
for k, name in enumerate(NAMES, 1):
    s = rw(f'xl/worksheets/sheet{k}.xml')
    s = re.sub(r'(<c [^>]*?\bs=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2)) + nx}{m.group(3)}', s)
    s = re.sub(r'(<row [^>]*?\bs=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2)) + nx}{m.group(3)}', s)
    s = re.sub(r'(<col [^>]*?\bstyle=")(\d+)(")', lambda m: f'{m.group(1)}{int(m.group(2)) + nx}{m.group(3)}', s)
    def cell(m):
        attrs, body = m.group(1), m.group(2); ref = re.search(r'r="([A-Z]+\d+)"', attrs).group(1)
        if 't="s"' in attrs:
            idx = int(re.search(r'<v>(\d+)</v>', body).group(1)); stats['str'] += 1
            return f'<c {attrs.replace(chr(116) + "=" + chr(34) + "s" + chr(34), "t=" + chr(34) + "inlineStr" + chr(34))}><is>{ss[idx]}</is></c>'
        if '<f>' in body:
            stats['formula'] += 1; f = re.search(r'<f>.*?</f>', body, re.S).group(0); v = VAL.get(f'{name}|{ref}')
            assert v is not None, (name, ref)
            stats['cached'] += 1; a2 = re.sub(r' t="[a-zA-Z]+"', '', attrs)
            if isinstance(v, bool): return f'<c {a2} t="b">{f}<v>{int(v)}</v></c>'
            if isinstance(v, str): return f'<c {a2} t="{"e" if v in ERR else "str"}">{f}<v>{html.escape(v, quote=False)}</v></c>'
            return f'<c {a2}>{f}<v>{v!r}</v></c>'
        return m.group(0)
    s = re.sub(r'<c ([^>]*?[^/])>(.*?)</c>', cell, s, flags=re.S)
    if k > 1: s = s.replace(' tabSelected="1"', '')
    minidom.parseString(s.encode())
    pn = f'xl/worksheets/sheet{100 + k}.xml'; assert pn not in exist and f'rId9{k:02d}"' not in wrels0
    new[pn] = s.encode()
    ct_add += f'<Override PartName="/{pn}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    wrels_add += f'<Relationship Id="rId9{k:02d}" Type="{REL}/worksheet" Target="worksheets/sheet{100 + k}.xml"/>'
    sheets_add += f'<sheet name="{esc(name)}" sheetId="{maxsid + k}" r:id="rId9{k:02d}"/>'
    if f'xl/worksheets/_rels/sheet{k}.xml.rels' not in zw.namelist(): continue
    sr = rw(f'xl/worksheets/_rels/sheet{k}.xml.rels'); dn = re.search(r'Target="/xl/drawings/(drawing\d+)\.xml"', sr).group(1); dnum = int(dn[7:])
    assert f'xl/drawings/drawing{300 + dnum}.xml' not in exist
    new[f'xl/worksheets/_rels/sheet{100 + k}.xml.rels'] = sr.replace(f'/xl/drawings/{dn}.xml', f'../drawings/drawing{300 + dnum}.xml').encode()
    new[f'xl/drawings/drawing{300 + dnum}.xml'] = zw.read(f'xl/drawings/{dn}.xml')
    ct_add += f'<Override PartName="/xl/drawings/drawing{300 + dnum}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    dr = rw(f'xl/drawings/_rels/{dn}.xml.rels')
    for im in re.findall(r'Target="/xl/media/(image\d+\.png)"', dr):
        new[f'xl/media/dcds_{im}'] = zw.read(f'xl/media/{im}'); dr = dr.replace(f'/xl/media/{im}', f'../media/dcds_{im}')
    new[f'xl/drawings/_rels/drawing{300 + dnum}.xml.rels'] = dr.encode()
print('ô:', stats)
# ---------- workbook, rels, content types
wbx = wbx.replace('<sheets>', '<sheets>' + sheets_add, 1)
wbx = re.sub(r'localSheetId="(\d+)"', lambda m: f'localSheetId="{int(m.group(1)) + NN}"', wbx)
pa = ''.join(re.findall(r'<definedName [^>]*>.*?</definedName>', re.search(r'<definedNames>.*?</definedNames>', wwb, re.S).group(0), re.S))
if '<definedNames>' in wbx: wbx = wbx.replace('<definedNames>', '<definedNames>' + pa, 1)
else: wbx = wbx.replace('</sheets>', '</sheets><definedNames>' + pa + '</definedNames>', 1)
wbx = re.sub(r' activeTab="\d+"', '', wbx); wbx = re.sub(r' firstSheet="\d+"', '', wbx)
minidom.parseString(wbx.encode()); repl['xl/workbook.xml'] = wbx.encode()
wr = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', wrels0).replace('</Relationships>', wrels_add + '</Relationships>'); repl['xl/_rels/workbook.xml.rels'] = wr.encode()
ct = rb('[Content_Types].xml'); ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
if 'Extension="png"' not in ct: ct = ct.replace('<Default ', '<Default Extension="png" ContentType="image/png"/><Default ', 1)
ct = ct.replace('</Types>', ct_add + '</Types>'); minidom.parseString(ct.encode()); repl['[Content_Types].xml'] = ct.encode()
with zipfile.ZipFile(dst, 'w') as zout:
    for info in zb.infolist():
        if info.filename in drop: continue
        zout.writestr(copy.copy(info), repl[info.filename] if info.filename in repl else zb.read(info.filename), compress_type=zipfile.ZIP_DEFLATED)
    for k, v in new.items(): zout.writestr(k, v, compress_type=zipfile.ZIP_DEFLATED)
print('ok, thêm', len(new), 'phần; sửa', sorted(repl))
