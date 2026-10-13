"""Trình bày lại một file Word có sẵn theo chuẩn Leonardo, giữ nguyên toàn bộ chữ và số.

  python restyle_doc.py goc.docx ket-qua.docx [--letterhead] [--header-title "tên ngắn"]

Đọc văn bản gốc theo thứ tự (đoạn, bảng, hình), dựng lại bằng lớp Doc: lề, phông, tiêu đề, kiểu bảng, đầu trang và chân trang theo chuẩn.
Chữ trong đoạn và trong ô bảng được chép nguyên văn (kể cả cách ghi số); số thứ tự tự động của Word được chuyển thành chữ.
Mặc định KHÔNG đặt logo và thông tin Leonardo, vì văn bản có thể của tổ chức khác; thêm --letterhead khi đó là văn bản của Leonardo.
Không mang theo: chữ ký số, ghi chú, theo dõi thay đổi, chú thích cuối trang, mục lục tự động, hộp chữ. Script báo nếu phát hiện các thành phần này.
Cuối cùng tự so chữ trước và sau; chỉ giữ file khi mọi đoạn và mọi ô bảng khớp.
"""
import io, os, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from leo_doc import Doc, no_number, _font, _borders, _widths, _no_borders, _shade, P, BODY_FONT, HEAD_FONT, BODY_PT, TABLE_PT, TEXT_W

args = [a for a in sys.argv[1:] if not a.startswith('--')]; src, out = args[0], args[1]
LETTER = '--letterhead' in sys.argv
HT = sys.argv[sys.argv.index('--header-title') + 1] if '--header-title' in sys.argv else None
if HT in args: args.remove(HT)
assert os.path.abspath(src) != os.path.abspath(out) and not os.path.exists(out), 'Không ghi đè: chọn tên file kết quả khác.'
z = zipfile.ZipFile(src); S = docx.Document(src); body = S.element.body
# ---------- cảnh báo thành phần không mang theo
docxml = z.read('word/document.xml').decode(); lost = []
if any(n.startswith('_xmlsignatures/') for n in z.namelist()): lost.append('chữ ký số (file mới không còn giá trị ký)')
for tag, name in (('w:ins ', 'theo dõi thay đổi'), ('w:del ', 'theo dõi thay đổi (phần xoá)'), ('w:commentReference', 'ghi chú'), ('w:footnoteReference', 'chú thích cuối trang'),
                  ('w:endnoteReference', 'chú thích cuối văn bản'), ('w:txbxContent', 'hộp chữ'), ('TOC \\', 'mục lục tự động')):
    if '<' + tag in docxml or tag in docxml and tag.startswith('TOC'): lost.append(name)
# ---------- số thứ tự tự động
NUM = z.read('word/numbering.xml').decode() if 'word/numbering.xml' in z.namelist() else ''
ABS = dict(re.findall(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', NUM)); counters = {}


def level_def(numid, il):
    m = re.search(r'<w:abstractNum [^>]*w:abstractNumId="%s".*?</w:abstractNum>' % ABS.get(numid, '-'), NUM, re.S)
    l = re.search(r'<w:lvl w:ilvl="%s".*?</w:lvl>' % il, m.group(0), re.S) if m else None
    if not l: return None
    return re.search(r'<w:numFmt w:val="([^"]+)"', l.group(0)).group(1), re.search(r'<w:lvlText w:val="([^"]*)"', l.group(0)).group(1)


def roman(n):
    r = ''
    for v, s in ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'), (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'), (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
        while n >= v: r += s; n -= v
    return r


def fmt_num(n, f):
    return {'decimal': str(n), 'lowerLetter': chr(96 + n), 'upperLetter': chr(64 + n), 'lowerRoman': roman(n), 'upperRoman': roman(n).upper()}.get(f, str(n))


def auto_label(p_el):
    """Trả về (nhãn, kiểu) của số thứ tự tự động, ví dụ ('1.', 'decimal'); ('–', 'bullet'); (None, None) nếu đoạn không đánh số."""
    np_ = p_el.find('.//' + qn('w:numPr'))
    if np_ is None or np_.find(qn('w:numId')) is None: return None, None
    nid = np_.find(qn('w:numId')).get(qn('w:val')); il = int(np_.find(qn('w:ilvl')).get(qn('w:val'))) if np_.find(qn('w:ilvl')) is not None else 0
    if nid == '0': return None, None
    d = level_def(nid, il)
    if not d: return None, None
    f, text = d
    if f == 'bullet': return '–', 'bullet'
    c = counters.setdefault(nid, {}); c[il] = c.get(il, 0) + 1
    for k in [k for k in c if k > il]: del c[k]
    lab = text
    for k in range(il + 1):
        dk = level_def(nid, k); lab = lab.replace(f'%{k + 1}', fmt_num(c.get(k, 1), dk[0] if dk else 'decimal'))
    return lab, f


def run_bold(r, par):
    if r.bold is not None: return r.bold
    st = par.style
    while st is not None:
        if st.font.bold is not None: return st.font.bold
        st = st.base_style
    return False


def copy_runs(src_par, dst_par, font, size, force_bold=False, upper=False, prefix=None, color=None):
    if prefix: _font(dst_par.add_run(prefix + ' '), font, size, bold=force_bold, color=color)
    runs = []
    for item in src_par.iter_inner_content():          # gồm cả chữ nằm trong liên kết
        runs += item.runs if hasattr(item, 'runs') else [item]
    for r in runs:
        t = r.text
        if t: _font(dst_par.add_run(t.upper() if upper else t), font, size, bold=force_bold or bool(run_bold(r, src_par)), italic=bool(r.italic), color=color)
    for d_ in src_par._p.findall('.//' + qn('w:drawing')):
        blip = d_.find('.//' + qn('a:blip')); ext = d_.find('.//' + qn('wp:extent'))
        if blip is None: continue
        part = S.part.related_parts[blip.get(qn('r:embed'))]; w = min(int(ext.get('cx')) / 36000 if ext is not None else TEXT_W, TEXT_W)
        dst_par.add_run().add_picture(io.BytesIO(part.blob), width=Mm(w)); stats['hình'] += 1


isnumtxt = lambda t: bool(re.fullmatch(r'[-–(]?[\d.,]+%?\)?|-', t.strip())) if t.strip() else False
elements = [(c.tag.split('}')[1], c) for c in body.iterchildren() if c.tag.split('}')[1] in ('p', 'tbl')]
first_text = next((i for i, (t, e) in enumerate(elements) if t == 'p' and Paragraph(e, S).text.strip()), 0)
# ---------- khối tên văn bản: các đoạn canh giữa liền nhau ở đầu
title_idx = []
for i in range(first_text, len(elements)):
    t, e = elements[i]
    if t != 'p': break
    p = Paragraph(e, S)
    if not p.text.strip(): continue
    if p.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(p.text) < 120 and p.text.strip().isupper(): title_idx.append(i)
    else: break
titles = [Paragraph(elements[i][1], S).text.strip() for i in title_idx]
D = Doc(header_title=HT or (' '.join(titles)[:70] if titles else ''), letterhead=LETTER); d = D.d
stats = dict(đoạn=0, bảng=0, hình=0, tiêu_đề=0); last_tbl = max((i for i, (t, e) in enumerate(elements) if t == 'tbl'), default=-1)


def layout_table(el):
    """Bảng dùng để dàn trang (quốc hiệu, khối ký): giữ bố cục, không kẻ viền."""
    t = Table(el, S); nc = len(t.columns); nt = d.add_table(rows=len(t.rows), cols=nc); _no_borders(nt)
    gw = [int(g.get(qn('w:w')) or 0) for g in el.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]; tot = sum(gw) or 1
    _widths(nt, [TEXT_W * w / tot for w in gw] if len(gw) == nc and tot > 1 else [TEXT_W / nc] * nc)
    for i, row in enumerate(t.rows):
        for j, cell in enumerate(row.cells):
            dst = nt.cell(i, j); first = True
            for sp in cell.paragraphs:
                dp = dst.paragraphs[0] if first else dst.add_paragraph(); first = False
                dp.alignment = sp.alignment if sp.alignment is not None else WD_ALIGN_PARAGRAPH.CENTER
                dp.paragraph_format.space_after = Pt(2); dp.paragraph_format.line_spacing = 1.15
                copy_runs(sp, dp, BODY_FONT, BODY_PT)
    d.add_paragraph().paragraph_format.space_after = Pt(2)


def data_table(el):
    t = Table(el, S); trs = el.findall(qn('w:tr')); gw = [int(g.get(qn('w:w')) or 0) for g in el.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]
    nc = len(gw) or len(t.columns); nt = d.add_table(rows=len(trs), cols=nc); nt.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(nt._tbl.tblPr, 'w:tblBorders', {s: (4, P['grid']) for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')})
    # lưới ô: grid[i] = [(cột bắt đầu, số cột gộp, phần tử ô)]
    grid = []
    for tr in trs:
        col = 0; row = []
        for tc in tr.findall(qn('w:tc')):
            gs = tc.find('.//' + qn('w:gridSpan')); span = int(gs.get(qn('w:val'))) if gs is not None else 1
            row.append((col, span, tc)); col += span
        grid.append(row)
    ctext = lambda tc: ' '.join(''.join(x.text or '' for x in pe.iter(qn('w:t'))) for pe in tc.findall(qn('w:p'))).strip()
    row_has_num = lambda row: any(isnumtxt(ctext(tc)) and ctext(tc) != '-' for _, _, tc in row)
    # dòng đầu là tiêu đề cột chỉ khi nó không chứa số; dòng thứ hai là tiêu đề phụ khi dòng đầu có ô gộp và dòng hai cũng không chứa số
    nhead = 0 if (not grid or row_has_num(grid[0])) else 1
    if nhead and len(grid) > 2 and any(sp > 1 for _, sp, _ in grid[0]) and not row_has_num(grid[1]) and not any(isnumtxt(ctext(tc)) for _, _, tc in grid[1][:1]): nhead = 2
    # độ rộng cột theo nội dung dài nhất (ô không gộp), cỡ chữ nhỏ dần khi bảng nhiều cột
    size = TABLE_PT if nc <= 4 else (9 if nc <= 6 else 8)
    wt = [4.0] * nc
    for row in grid:
        for c0, sp, tc in row:
            if sp == 1 and c0 < nc: wt[c0] = max(wt[c0], min(max((len(w) for w in ctext(tc).split()), default=0) + 2, 40) if not isnumtxt(ctext(tc)) else len(ctext(tc)) + 2, min(len(ctext(tc)), 34) * 0.55)
    _widths(nt, [TEXT_W * w / sum(wt) for w in wt])
    for i, tr in enumerate(trs):
        col = 0; cells = tr.findall(qn('w:tc')); row_text = ' '.join(''.join(x.text or '' for x in tc.iter(qn('w:t'))) for tc in cells)
        total = i >= nhead and i == len(trs) - 1 and re.search(r'(?i)\btổng\b', row_text) is not None
        for tc in cells:
            gs = tc.find('.//' + qn('w:gridSpan')); span = int(gs.get(qn('w:val'))) if gs is not None else 1
            if col >= nc: break
            dst = nt.cell(i, col)
            if span > 1: dst = dst.merge(nt.cell(i, min(col + span - 1, nc - 1)))
            first = True
            for pe in tc.findall(qn('w:p')):
                sp = Paragraph(pe, S); txt = sp.text
                dp = dst.paragraphs[0] if first else dst.add_paragraph(); first = False
                dp.paragraph_format.space_after = Pt(2); dp.paragraph_format.space_before = Pt(2); dp.paragraph_format.line_spacing = 1.1
                if i < nhead: dp.alignment = WD_ALIGN_PARAGRAPH.CENTER; copy_runs(sp, dp, HEAD_FONT, size, force_bold=True, color='FFFFFF')
                else:
                    dp.alignment = WD_ALIGN_PARAGRAPH.RIGHT if isnumtxt(txt) and col > 0 else (WD_ALIGN_PARAGRAPH.CENTER if isnumtxt(txt) else WD_ALIGN_PARAGRAPH.LEFT)
                    copy_runs(sp, dp, HEAD_FONT, size, force_bold=total)
            if i < nhead: _shade(dst, P['green'])
            elif total: _shade(dst, P['soft'])
            col += span
    for k in range(nhead):
        trp = nt.rows[k]._tr.get_or_add_trPr(); th = docx.oxml.OxmlElement('w:tblHeader'); th.set(qn('w:val'), 'true'); trp.append(th)
    d.add_paragraph().paragraph_format.space_after = Pt(2)


for i, (tag, el) in enumerate(elements):
    if tag == 'tbl':
        stats['bảng'] += 1
        x = Table(el, S); txt = ' '.join(c.text for r in x.rows for c in r.cells)
        is_layout = i < first_text or (i == last_tbl and len(x.rows) <= 3 and not re.search(r'\d{3}', txt))
        layout_table(el) if is_layout else data_table(el)
        continue
    sp = Paragraph(el, S); text = sp.text.strip(); has_img = el.find('.//' + qn('w:drawing')) is not None
    label, kind = auto_label(el)
    if not text and not has_img: continue
    stats['đoạn'] += 1
    if i in title_idx:
        k = title_idx.index(i); dp = d.add_paragraph(); dp.alignment = WD_ALIGN_PARAGRAPH.CENTER; dp.paragraph_format.line_spacing = 1.15
        dp.paragraph_format.space_after = Pt(2 if k < len(title_idx) - 1 else 10); dp.paragraph_format.space_before = Pt(8 if k == 0 else 0)
        big = len(title_idx) == 1 or k == 1
        copy_runs(sp, dp, HEAD_FONT, 16 if big else 11, force_bold=True, upper=True, color=P['green'] if big else P['grey']); continue
    typed = re.match(r'^(\d{1,2})\.\s+\S', text)
    if (kind == 'decimal' and re.fullmatch(r'\d+\.', label or '')) or (typed and len(text) < 140 and not text.rstrip().endswith((';', ','))):
        h = d.add_heading('', level=1); no_number(h); copy_runs(sp, h, HEAD_FONT, 12, force_bold=True, upper=True, prefix=label); stats['tiêu_đề'] += 1; continue
    dp = d.add_paragraph()
    dp.alignment = WD_ALIGN_PARAGRAPH.CENTER if (sp.alignment == WD_ALIGN_PARAGRAPH.CENTER or (has_img and not text)) else (WD_ALIGN_PARAGRAPH.RIGHT if sp.alignment == WD_ALIGN_PARAGRAPH.RIGHT else WD_ALIGN_PARAGRAPH.JUSTIFY)
    if kind == 'bullet' or re.match(r'^[-–•]\s', text):
        dp.paragraph_format.left_indent = Mm(10); dp.paragraph_format.first_line_indent = Mm(-4.5); dp.paragraph_format.space_after = Pt(3)
    elif label or re.match(r'^[a-z]\)\s', text):
        dp.paragraph_format.left_indent = Mm(6); dp.paragraph_format.first_line_indent = Mm(-6)
    copy_runs(sp, dp, BODY_FONT, BODY_PT, prefix=label)
D.save(out)
# ---------- tự so chữ trước và sau
strip = lambda t: re.sub(r'^(\d+\.|[a-zA-Z]\)|[ivxIVX]+[.)]|–)\s+', '', ' '.join(t.split())).casefold()
def texts(path):
    dd = docx.Document(path); ps = []; cs = []
    for c in dd.element.body.iterchildren():
        tg = c.tag.split('}')[1]
        if tg == 'p':
            t = Paragraph(c, dd).text
            if t.strip(): ps.append(strip(t))
        elif tg == 'tbl':
            for tc in c.iter(qn('w:tc')):
                t = ' '.join(''.join(x.text or '' for x in pe.iter(qn('w:t'))) for pe in tc.findall(qn('w:p')))
                if t.strip(): cs.append(' '.join(t.split()))
    return ps, cs
(pa, ca), (pb, cb) = texts(src), texts(out)
print(f"Đã dựng: {stats['đoạn']} đoạn ({stats['tiêu_đề']} tiêu đề cấp 1), {stats['bảng']} bảng, {stats['hình']} hình | nhận diện Leonardo: {'có' if LETTER else 'không'}")
if lost: print('⚠ KHÔNG mang theo:', '; '.join(sorted(set(lost))))
ok = pa == pb and ca == cb
print(f'So chữ: đoạn {len(pa)} → {len(pb)} ({"khớp" if pa == pb else "KHÁC"}); ô bảng {len(ca)} → {len(cb)} ({"khớp" if ca == cb else "KHÁC"})')
if not ok:
    import difflib
    for l in [l for l in difflib.unified_diff(pa + ca, pb + cb, lineterm='', n=0) if l[:1] in '+-' and l[:3] not in ('+++', '---')][:12]: print('   ' + l[:150])
    os.remove(out); print('KẾT LUẬN: KHÔNG ĐẠT — đã xoá file kết quả'); sys.exit(1)
print('KẾT LUẬN: ĐẠT —', out)
