# -*- coding: utf-8 -*-
from pathlib import Path
from copy import deepcopy as dc
import zipfile, xml.etree.ElementTree as E, io, re, json, shutil
from PIL import Image

QA = Path(__file__).resolve().parent
TASK = QA.parent
OUT = TASK / 'Ho_so_nop'
MAIN = OUT / 'De_cuong_Tin12_Bai9_Lop12-4.docx'
HAND = OUT / '02_Phieu_hoc_tap_va_dap_an_Tin12_Bai9.docx'
SOURCE = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_template_5512.docx')
HSOURCE = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_worksheet_5512.docx')
CROPPED_DIR = QA / 'cropped_figures'

def package(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}

PKG = package(SOURCE)
HPKG = package(HSOURCE)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
w = lambda t: '{' + W + '}' + t

for data in [HPKG['word/document.xml'], PKG['word/document.xml']]:
    for _, (prefix, uri) in E.iterparse(io.BytesIO(data), events=['start-ns']):
        if prefix and not re.fullmatch(r'ns\d+', prefix):
            E.register_namespace(prefix, uri)
E.register_namespace('w', W)
E.register_namespace('wp', WP)
E.register_namespace('a', A)
E.register_namespace('pic', PIC)
E.register_namespace('r', R)

root = E.fromstring(PKG['word/document.xml'])
oldbody = root.find(w('body'))

def txt(e):
    return ''.join(t.text or '' for t in e.iter(w('t')))

def node(t, attrs=None):
    return E.Element(w(t), {w(k): str(v) for k, v in (attrs or {}).items()})

def font(rp, size_val=26, font_name='Times New Roman', color_val='000000'):
    rp.append(node('rFonts', {'ascii': font_name, 'hAnsi': font_name, 'cs': font_name, 'eastAsia': font_name}))
    rp.append(node('color', {'val': color_val}))
    rp.append(node('sz', {'val': size_val}))
    rp.append(node('szCs', {'val': size_val}))

bullet_model = next(p for p in oldbody.iter(w('p')) if p.find(w('pPr') + '/' + w('numPr')) is not None)

def p(text='', role='body', align=None, keep=False, first=None, italic=False, sz=26):
    out = node('p')
    pp = node('pPr')
    out.append(pp)
    if role == 'title':
        pp.append(node('pStyle', {'val': 'Title'}))
    if role == 'bullet':
        pp.append(dc(bullet_model.find(w('pPr') + '/' + w('numPr'))))
        pp.append(node('ind', {'left': 567, 'hanging': 283, 'right': 0}))
    elif role == 'code':
        pp.append(node('ind', {'left': 283, 'right': 0, 'firstLine': 0}))
    else:
        pp.append(node('ind', {'left': 0, 'right': 0, 'firstLine': 567 if role == 'body' and first is None else (first or 0)}))
    if role in ['title', 'heading', 'sub', 'label'] or keep:
        pp.append(node('keepNext'))
    pp.append(node('widowControl', {'val': 1}))
    pp.append(node('snapToGrid', {'val': 0}))
    pp.append(node('spacing', {'before': 60 if role in ['heading', 'sub'] else 0, 'after': 40 if role in ['heading', 'sub'] else 0, 'line': 276 if role in ['cell', 'code'] else 288, 'lineRule': 'auto'}))
    pp.append(node('jc', {'val': align or ('center' if role in ['title', 'caption'] else 'left' if role in ['heading', 'sub', 'label', 'code', 'cell'] else 'both')}))
    rp = node('rPr')
    font(rp, size_val=22 if role == 'code' else sz, font_name='Courier New' if role == 'code' else 'Times New Roman')
    pp.append(rp)

    # Normalize nested markdown markers and parse **bold**, *italic*, `code`
    text = re.sub(r'\*\*`([^`]+)`\*\*', r'**\1**', text)
    text = re.sub(r'`(\[[A-Za-z0-9_. –-]+\])`', r'**\1**', text)
    tokens = re.split(r'(\*\*.*?\*\*|\*[^*\n]+?\*|`[^`\n]+?`)', text)
    for token in tokens:
        if not token:
            continue
        is_bold = token.startswith('**') and token.endswith('**') and len(token) > 4
        is_italic = token.startswith('*') and token.endswith('*') and not is_bold and len(token) > 2
        is_code = token.startswith('`') and token.endswith('`') and len(token) > 2
        value = token[2:-2] if is_bold else token[1:-1] if (is_italic or is_code) else token
        run = node('r')
        rp = node('rPr')
        use_code_font = is_code or role == 'code'
        font(rp, size_val=22 if use_code_font else sz, font_name='Courier New' if use_code_font else 'Times New Roman')
        if is_bold or role in ['title', 'heading', 'sub', 'label']:
            rp.append(node('b'))
        if italic or is_italic or role == 'caption':
            rp.append(node('i'))
        run.append(rp)
        for i, line in enumerate(value.split('\n')):
            if i:
                run.append(node('br'))
            t = node('t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = line
            run.append(t)
        out.append(run)
    return out

pending_breaks = set()
def add(body, text='', role='body', **kwargs):
    out = p(text, role, **kwargs)
    if id(body) in pending_breaks:
        out.find(w('pPr')).append(node('pageBreakBefore'))
        pending_breaks.remove(id(body))
    body.append(out)
    return out

def br(body):
    pending_breaks.add(id(body))

def tabline(body, left, right):
    out = p('', first=0, align='left')
    tabs = node('tabs')
    tabs.append(node('tab', {'val': 'left', 'pos': 3800}))
    out.find(w('pPr')).append(tabs)
    for i, value in enumerate([left, right]):
        run = node('r')
        rp = node('rPr')
        font(rp)
        rp.append(node('b'))
        run.append(rp)
        if i:
            run.append(node('tab'))
        t = node('t')
        t.text = value
        run.append(t)
        out.append(run)
    body.append(out)

# Register 8 figures in PKG
FIG_MAP = {
    'hinh_9_1.png': ('rId21', 4500000, 'Hình 9.1. Danh sách có thứ tự trong HTML và kết quả hiển thị (SGK trang 52)'),
    'hinh_9_2.png': ('rId22', 4500000, 'Hình 9.2. Danh sách không có thứ tự dùng kí tự đầu dòng hình vuông (SGK trang 52)'),
    'hinh_9_3.png': ('rId23', 4680000, 'Hình 9.3. Ví dụ tạo danh sách mô tả bằng cặp thẻ <dl>, <dt>, <dd> (SGK trang 53)'),
    'hinh_9_4.png': ('rId24', 4680000, 'Hình 9.4. Cấu trúc HTML cơ bản tạo bảng và kết quả hiển thị (SGK trang 54)'),
    'hinh_9_5.png': ('rId25', 4680000, 'Hình 9.5. Ví dụ sử dụng thuộc tính style điều chỉnh kích thước bảng, hàng, ô (SGK trang 54)'),
    'hinh_9_6.png': ('rId26', 4680000, 'Hình 9.6. Ví dụ gộp hàng (rowspan) và gộp cột (colspan) của bảng (SGK trang 55)'),
    'hinh_9_7.png': ('rId27', 2050000, 'Hình 9.7. Mẫu danh sách các Câu lạc bộ lồng nhau cần tạo ở Nhiệm vụ 1 (SGK trang 55)'),
    'hinh_9_8.png': ('rId28', 3960000, 'Hình 9.8. Mẫu bảng Lịch hoạt động CLB Thể thao cần tạo ở Nhiệm vụ 2 (SGK trang 56)'),
}

doc_pr_counter = [100]
def add_figure(body, fig_name):
    rid, cx, caption_text = FIG_MAP[fig_name]
    img_path = CROPPED_DIR / fig_name
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    cy = int(cx * h_px / w_px)
    doc_pr_counter[0] += 1
    did = doc_pr_counter[0]

    para = node('p')
    pp = node('pPr')
    pp.append(node('keepNext'))
    pp.append(node('widowControl', {'val': 1}))
    pp.append(node('spacing', {'before': 60, 'after': 20, 'line': 240, 'lineRule': 'auto'}))
    pp.append(node('jc', {'val': 'center'}))
    para.append(pp)

    drawing_xml = f'''<w:r xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}">
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="{cx}" cy="{cy}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="{did}" name="{fig_name}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="{did}" name="{fig_name}"/>
                  <pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>
                  <pic:nvPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="{rid}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>'''
    para.append(E.fromstring(drawing_xml))
    body.append(para)
    add(body, caption_text, role='caption', first=0, italic=True, sz=24)

def table(rows, widths, header=True, cell_sz=24):
    t = node('tbl')
    pr = node('tblPr')
    t.append(pr)
    pr.append(node('tblW', {'w': sum(widths), 'type': 'dxa'}))
    pr.append(node('jc', {'val': 'center'}))
    pr.append(node('tblInd', {'w': 0, 'type': 'dxa'}))
    borders = node('tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        borders.append(node(side, {'val': 'single', 'sz': 4, 'color': '404040'}))
    pr.append(borders)
    pr.append(node('tblLayout', {'type': 'fixed'}))
    margins = node('tblCellMar')
    for side, value in [('top', 60), ('bottom', 60), ('left', 90), ('right', 90)]:
        margins.append(node(side, {'w': value, 'type': 'dxa'}))
    pr.append(margins)
    grid = node('tblGrid')
    t.append(grid)
    for width in widths:
        grid.append(node('gridCol', {'w': width}))
    for i, values in enumerate(rows):
        row = node('tr')
        trp = node('trPr')
        row.append(trp)
        trp.append(node('cantSplit'))
        if header and i == 0:
            trp.append(node('tblHeader'))
        for j, value in enumerate(values):
            cell = node('tc')
            cp = node('tcPr')
            cp.append(node('tcW', {'w': widths[j], 'type': 'dxa'}))
            cp.append(node('vAlign', {'val': 'top'}))
            if header and i == 0:
                cp.append(node('shd', {'val': 'clear', 'color': 'auto', 'fill': 'E8EEF5'}))
            cell.append(cp)
            lines = str(value).split('\n') if not isinstance(value, list) else value
            for line in lines:
                is_code_line = line.strip().startswith('`') and line.strip().endswith('`') and len(line.strip()) > 2 and '**' not in line
                if header and i == 0:
                    cell.append(p('**' + line + '**', 'cell', align='center', sz=cell_sz))
                elif is_code_line:
                    cell.append(p(line.strip()[1:-1], 'code', align='left', sz=21))
                else:
                    cell.append(p(line, 'cell', align='left', sz=cell_sz))
            row.append(cell)
        t.append(row)
    return t

def signature_table():
    widths = [4500, 4500]
    t = node('tbl')
    pr = node('tblPr')
    t.append(pr)
    pr.append(node('tblW', {'w': sum(widths), 'type': 'dxa'}))
    pr.append(node('jc', {'val': 'center'}))
    borders = node('tblBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        borders.append(node(side, {'val': 'none', 'sz': 0, 'color': 'auto'}))
    pr.append(borders)
    grid = node('tblGrid')
    t.append(grid)
    for width in widths:
        grid.append(node('gridCol', {'w': width}))

    row = node('tr')
    trp = node('trPr')
    trp.append(node('cantSplit'))
    row.append(trp)

    # Left cell: GVHD
    c1 = node('tc')
    cp1 = node('tcPr')
    cp1.append(node('tcW', {'w': widths[0], 'type': 'dxa'}))
    c1.append(cp1)
    c1.append(p('', 'cell', align='center', keep=True))
    c1.append(p('**GIÁO VIÊN HƯỚNG DẪN**', 'cell', align='center', keep=True))
    c1.append(p('(Ký và ghi rõ họ tên)', 'cell', align='center', italic=True, keep=True, sz=24))
    for _ in range(3):
        c1.append(p('', 'cell', align='center', keep=True))
    c1.append(p('**Thầy Mai Xuân Văn**', 'cell', align='center'))

    # Right cell: GV giảng dạy / SVTT
    c2 = node('tc')
    cp2 = node('tcPr')
    cp2.append(node('tcW', {'w': widths[1], 'type': 'dxa'}))
    c2.append(cp2)
    c2.append(p('Huế, ngày 04 tháng 11 năm 2026', 'cell', align='center', italic=True, keep=True))
    c2.append(p('**GIÁO VIÊN GIẢNG DẠY**', 'cell', align='center', keep=True))
    c2.append(p('(Sinh viên thực hiện ký và ghi rõ họ tên)', 'cell', align='center', italic=True, keep=True, sz=24))
    for _ in range(3):
        c2.append(p('', 'cell', align='center', keep=True))
    c2.append(p('**Hồ Quang Huy**', 'cell', align='center'))

    row.append(c1)
    row.append(c2)
    t.append(row)
    return t

def ending(pkg, start=1):
    sect = dc(E.fromstring(pkg['word/document.xml']).find(w('body'))[-1])
    for n in list(sect.findall(w('pgNumType'))):
        sect.remove(n)
    sect.append(node('pgNumType', {'start': start}))
    return sect

def write_doc(pkg, body, path, include_figures=False):
    data = dict(pkg)
    if include_figures:
        rels_xml = data['word/_rels/document.xml.rels'].decode('utf-8')
        for fig_name, (rid, _, _) in FIG_MAP.items():
            data['word/media/' + fig_name] = (CROPPED_DIR / fig_name).read_bytes()
            rel_tag = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fig_name}"/>'
            if f'Id="{rid}"' not in rels_xml:
                rels_xml = rels_xml.replace('</Relationships>', rel_tag + '</Relationships>')
        data['word/_rels/document.xml.rels'] = rels_xml.encode('utf-8')

    src = data['word/document.xml'].decode('utf-8')
    prefix = src[:src.index('<w:body')]
    data['word/document.xml'] = (prefix + E.tostring(body, encoding='unicode') + '</w:document>').encode('utf-8')
    if 'word/settings.xml' in data:
        s = data['word/settings.xml'].decode('utf-8')
        if '<w:updateFields' not in s:
            s = s.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>')
        data['word/settings.xml'] = s.encode('utf-8')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, value in data.items():
            z.writestr(n, value)

# ==============================================================================
# BUILD MAIN LESSON PLAN DOCUMENT (De_cuong_Tin12_Bai9_Lop12-4.docx)
# ==============================================================================
b = node('body')
for idx, el in enumerate(list(oldbody)[:25]):
    if idx in (11, 12):
        continue
    b.append(dc(el))

cover = list(b)
for el in cover:
    for rp in el.iter(w('rPr')):
        for tag in ['sz', 'szCs']:
            size = rp.find(w(tag))
            if size is not None:
                size.set(w('val'), '26')

# Fix Comment #0 on Cover Title: ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY & BỔ SUNG TIẾT PPCT
cover_h1 = cover[6]
for child in list(cover_h1):
    if child.tag != w('pPr'):
        cover_h1.remove(child)
for run in p('ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY', 'title').findall(w('r')):
    cover_h1.append(run)

cover_h2 = cover[7]
for child in list(cover_h2):
    if child.tag != w('pPr'):
        cover_h2.remove(child)
for run in p('BÀI 9. TẠO DANH SÁCH, BẢNG (02 TIẾT – TIẾT PPCT: 20, 21)', 'title').findall(w('r')):
    cover_h2.append(run)

# Update cover table row 6 & row 7
cover_table = cover[10]
c_rows = cover_table.findall(w('tr'))
if len(c_rows) >= 8:
    # Row 6: Lịch dự giờ
    tc_sched_lbl = c_rows[6].findall(w('tc'))[0]
    for ch in list(tc_sched_lbl):
        if ch.tag != w('tcPr'):
            tc_sched_lbl.remove(ch)
    tc_sched_lbl.append(p('**Lịch dự giờ**', 'cell', sz=26))
    # Row 7: Tên bài dạy + Tiết PPCT
    tc_lesson_val = c_rows[7].findall(w('tc'))[2]
    for ch in list(tc_lesson_val):
        if ch.tag != w('tcPr'):
            tc_lesson_val.remove(ch)
    tc_lesson_val.append(p('Bài 9. Tạo danh sách, bảng (02 tiết – Tiết PPCT: 20, 21)', 'cell', sz=26))

# Page 1 Header (Fixes Comment #0 & Comment #1)
tabline(b, 'Trường: THPT Thuận Hoá', 'Giáo viên giảng dạy (SVTT): Hồ Quang Huy')
tabline(b, 'Tổ: Tin học', 'Giáo viên hướng dẫn: Thầy Mai Xuân Văn')
add(b, 'ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY', 'title')
add(b, 'TÊN BÀI DẠY: BÀI 9. TẠO DANH SÁCH, BẢNG', 'title')
add(b, '(Thời lượng: 02 tiết – Tiết thứ 1, 2 của bài; Tiết PPCT: 20, 21)', align='center', first=0, keep=True, italic=True)
add(b, 'Môn học: Tin học; Lớp: 12/4', align='center', first=0, keep=True)
add(b, 'Thời gian thực hiện: 02 tiết (90 phút – Tiết PPCT: 20, 21; mỗi tiết 45 phút)', align='center', first=0, keep=True)
add(b, 'Lịch dự giờ: Thứ Năm, ngày 05/11/2026 – Tiết 3 (09h00 – 09h45, dự giờ Tiết 1 / Tiết PPCT: 20) – Phòng H.II.4', align='center', first=0, keep=True)
add(b, 'Hình thức tổ chức: Dạy học kết hợp thực hành tại phòng máy', align='center', first=0)

# ==============================================================================
# I. MỤC TIÊU (Fixes Comments #8, #14, #17, #21, #25, #31)
# ==============================================================================
add(b, 'I. MỤC TIÊU', 'heading')
add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT), lớp 12, Chủ đề F (Giải quyết vấn đề với sự trợ giúp của máy tính – Tạo trang web): Sử dụng được các thẻ HTML để trình bày trang web, trong đó có tạo danh sách và tạo bảng.')

add(b, '1. Về kiến thức', 'sub')
add(b, 'Sau bài học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu `[KT1]` – `[KT6]` để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
KT_LIST = [
    '**`[KT1]`**: Nhận biết được ý nghĩa, ưu điểm của việc sử dụng danh sách và bảng trong trình bày văn bản trên trang web; lựa chọn được dạng trình bày (danh sách có thứ tự, danh sách không có thứ tự, danh sách mô tả hoặc bảng) phù hợp với đặc điểm dữ liệu.',
    '**`[KT2]`**: Trình bày được cú pháp và sử dụng đúng các cặp thẻ `<ol>...</ol>`, `<ul>...</ul>`, `<li>...</li>` cùng thuộc tính `type`, `start` và đặc tính `list-style-type` trong thuộc tính `style` để tạo danh sách có thứ tự và danh sách không có thứ tự.',
    '**`[KT3]`**: Trình bày được cú pháp và sử dụng đúng bộ thẻ `<dl>`, `<dt>`, `<dd>` để tạo danh sách mô tả; thực hiện được cách đặt danh sách con bên trong phần tử `<li>...</li>` của danh sách cha để tạo danh sách lồng nhau nhiều cấp.',
    '**`[KT4]`**: Trình bày được cú pháp và sử dụng đúng các cặp thẻ `<table>`, `<tr>`, `<th>`, `<td>`, `<caption>` để khởi tạo cấu trúc bảng HTML cơ bản có tiêu đề bảng, hàng tiêu đề và các hàng dữ liệu.',
    '**`[KT5]`**: Sử dụng được thuộc tính `style` (`border`, `width`, `height`) để tạo khung viền cho bảng/ô, điều chỉnh kích thước và sử dụng đúng thuộc tính `rowspan`, `colspan` để gộp hàng, gộp cột trong bảng.',
    '**`[KT6]`**: Vận dụng tổng hợp các thẻ tạo danh sách và bảng để hoàn thiện trang web giới thiệu và lịch hoạt động Câu lạc bộ theo mẫu SGK (Hình 9.7, Hình 9.8); phát hiện, sửa được lỗi cấu trúc thẻ bảng HTML và biết kiểm chứng gợi ý của AI trước khi áp dụng.'
]
for item in KT_LIST:
    add(b, item, 'bullet')

add(b, '2. Về năng lực', 'sub')
add(b, 'a) Năng lực chung', 'sub')
NLC_LIST = [
    '**`[NLC1 – Tự chủ và tự học]`**: Chủ động đọc hiểu Sách giáo khoa, quan sát hình minh họa (Hình 9.1 – 9.8), tự thực hành gõ mã HTML, lưu tệp và đối chiếu kết quả hiển thị trên trình duyệt với yêu cầu đề ra.',
    '**`[NLC2 – Giao tiếp và hợp tác]`**: Tích cực trao đổi, thảo luận theo cặp đôi tại phòng máy, phân công rõ vai người thao tác và người kiểm tra; trình bày, giải thích rõ ràng lựa chọn thẻ, cấu trúc danh sách lồng hoặc cách tính số ô sau khi gộp.',
    '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Phân tích yêu cầu trình bày dữ liệu thực tế để lựa chọn cấu trúc danh sách/bảng phù hợp; phát hiện nguyên nhân lỗi cú pháp HTML và đề xuất cách khắc phục chính xác.'
]
for item in NLC_LIST:
    add(b, item, 'bullet')

add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
NLDT_LIST = [
    '**`[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]`**: Sử dụng thành thạo phần mềm soạn thảo mã nguồn/văn bản thuần (Notepad / VS Code / Sublime Text) và trình duyệt web để tạo, lưu tệp `.html` (mã hóa UTF-8), mở, tải lại trang và quản lí tệp thực hành xuyên suốt hai tiết học.',
    '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông]`**: Viết đúng mã HTML tạo các kiểu danh sách, danh sách lồng và thiết lập bảng có định dạng viền, kích thước, gộp ô (`rowspan`, `colspan`); kiểm tra số hàng/cột quy đổi và sửa lỗi quan hệ phân cấp `<table>` – `<tr>` – `<th>`/`<td>`.',
    '**`[NLd – Ứng dụng CNTT và truyền thông trong học và tự học]`**: Khai thác hiệu quả Sách giáo khoa, phiếu học tập, tệp học liệu HTML mẫu và công cụ trợ lý AI để tìm hiểu cú pháp, kiểm chứng lỗi cấu trúc mã HTML dựa trên bằng chứng thực nghiệm.'
]
for item in NLDT_LIST:
    add(b, item, 'bullet')

add(b, 'c) Năng lực số và năng lực AI', 'sub')
NLS_LIST = [
    '**`[NLS 3.1.NC1.a]`**: Tạo và phát triển nội dung số có cấu trúc trên trang web thông qua hệ thống thẻ danh sách và bảng HTML.',
    '**`[NLS 3.2.NC1.a]`**: Tích hợp, tinh chỉnh định dạng danh sách và bảng dữ liệu trong cùng một trang web hoàn chỉnh.',
    '**`[NLAI 12.C2.2]`**: Vận dụng trợ lý AI để hỗ trợ phát hiện lỗi cú pháp mã HTML; biết đối chiếu với chuẩn kiến thức SGK và chạy thử trên trình duyệt để kiểm chứng gợi ý của AI trước khi sử dụng.'
]
for item in NLS_LIST:
    add(b, item, 'bullet')

add(b, '3. Về phẩm chất', 'sub')
PC_LIST = [
    '**`[PC1 – Chăm chỉ]`**: Tích cực nghiên cứu bài học, kiên trì thực hành gõ mã lệnh, rà soát và chỉnh sửa từng cặp thẻ HTML cho đến khi trang web hiển thị đúng mẫu.',
    '**`[PC2 – Trách nhiệm]`**: Hoàn thành đầy đủ nhiệm vụ cá nhân và cặp đôi đúng thời gian quy định; có ý thức lưu trữ, bảo quản tệp thực hành và tuân thủ nội quy phòng máy.',
    '**`[PC3 – Trung thực]`**: Tự giác thực hiện nhiệm vụ, báo cáo đúng kết quả chạy thử, giải thích trung thực phần mã do bản thân viết và phần tham khảo/kiểm chứng từ gợi ý hỗ trợ.'
]
for item in PC_LIST:
    add(b, item, 'bullet')

# ==============================================================================
# II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU (Fixes Comments #2, #3)
# ==============================================================================
add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
for item in [
    'Máy tính của giáo viên, máy chiếu (Projector), bảng viết; phòng máy tính thực hành có cài đặt trình soạn thảo văn bản thuần / mã nguồn (Notepad, Notepad++ hoặc Visual Studio Code) và trình duyệt web (Google Chrome, Microsoft Edge hoặc CocCoc).',
    'Bố trí học sinh ngồi theo cặp (01 – 02 học sinh/máy), luân phiên đảm nhận vai người thao tác viết mã và người đối chiếu, kiểm tra kết quả.',
    'Hệ thống tệp học liệu HTML hoạt động hoàn toàn ngoại tuyến (Offline). Trong hoạt động kiểm chứng AI ở cuối bài, giáo viên sử dụng tình huống phản hồi giả định đã chuẩn bị sẵn trong Phiếu học tập số 5 (nếu phòng máy có kết nối Internet, giáo viên có thể minh họa trực tiếp trên máy chiếu).'
]:
    add(b, item, 'bullet')

add(b, '2. Học liệu (Kèm tham chiếu Phụ lục và đường dẫn lưu trữ)', 'sub')
add(b, 'a) Giáo viên chuẩn bị:', 'label')
for item in [
    '**Sách giáo khoa:** *Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng* (và *Sách giáo khoa Tin học 12 – Định hướng Khoa học máy tính*), **Bộ sách Kết nối tri thức với cuộc sống**, Nhà xuất bản Giáo dục Việt Nam (Phần 1 – Nội dung chung cả hai định hướng, Chủ đề 4: Giải quyết vấn đề với sự trợ giúp của máy tính, **Bài 9. Tạo danh sách, bảng**, trang 52 – 56; các Hình 9.1 – 9.8).',
    '**Sách giáo viên & Kế hoạch bài dạy tham khảo:** *Sách giáo viên Tin học 12* và *Kế hoạch bài dạy môn Tin học 12 – Bộ sách Kết nối tri thức với cuộc sống*, Nhà xuất bản Giáo dục Việt Nam.',
    '**Hệ thống Phiếu học tập và Công cụ kiểm tra, đánh giá (Đính kèm đầy đủ tại Mục IV. PHỤ LỤC cuối văn bản này):** Gồm Phiếu học tập số 1, 2, 3, 4, 5 kèm Đáp án và 07 Công cụ kiểm tra, đánh giá (`CCĐG 1` → `CCĐG 7`: Bảng kiểm đánh giá từng nhiệm vụ và Phiếu chấm Rubric 10 điểm tại **Phụ lục 1, Phụ lục 2, Phụ lục 3**).',
    '**Bộ tệp học liệu HTML thực hành (Đính kèm toàn văn mã nguồn tại Phụ lục 4 và đóng gói trong thư mục `Hoc_lieu_HTML/` của hồ sơ nộp):** Gồm `00_Trinh_chieu.html` (bài trình chiếu trực quan 2 tiết), `01_Khung_CLB.html` (tệp khung khởi tạo cho học sinh), `02_CLB_Mau_CoBan.html` (mẫu bảng cơ bản), `03_Ma_Can_Sua.html` & `04_Ma_Da_Sua.html` (học liệu kiểm chứng lỗi cấu trúc bảng), `05_CLB_HoanChinh_SGK.html` (sản phẩm mẫu hoàn chỉnh theo Hình 9.7 và Hình 9.8 SGK). Đường dẫn lưu trữ học liệu trực tuyến dự phòng trên Google Drive: `https://drive.google.com/drive/folders/1rwm5SGx-sOww3XrSkvKNfglCtejI0lA0`.'
]:
    add(b, item, 'bullet')

add(b, 'b) Học sinh chuẩn bị:', 'label')
for item in [
    '*Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (hoặc Định hướng Khoa học máy tính), Bộ sách Kết nối tri thức với cuộc sống*, vở ghi bài, bút viết.',
    'Ôn lại kiến thức Bài 7 (*HTML và cấu trúc trang web*) và Bài 8 (*Định dạng văn bản*, đặc biệt là cú pháp thuộc tính `style="thuộc_tính: giá_trị;"` và Hình 8.3 trang 50).',
    'Tạo thư mục lưu trữ cá nhân trên máy tính phòng thực hành để lưu tệp sản phẩm `HoTen_12-4_Bai9.html` xuyên suốt hai tiết học (Tiết PPCT 20 và Tiết PPCT 21).'
]:
    add(b, item, 'bullet')

# ==============================================================================
# III. TIẾN TRÌNH DẠY HỌC (Fixes Comments #4–#32)
# ==============================================================================
add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
add(b, '**Bảng tổng hợp phân phối thời gian và cấu trúc tiến trình bài dạy (02 tiết = 90 phút; Tiết PPCT: 20, 21):**', keep=True)

TIME_OVERVIEW = [
    ['Tiết học\n(Tiết PPCT)', 'Hoạt động dạy học (Theo Công văn 5512)', 'Tiểu mục / Nhiệm vụ học tập', 'Thời lượng', 'Công cụ ĐG\n(Phụ lục)'],
    ['Tiết 1 của bài\n(Tiết PPCT: 20)\n[45 phút]',
     '1. Hoạt động 1: Mở đầu',
     'Xác định vấn đề: Lựa chọn cách trình bày dữ liệu bằng danh sách và bảng trên trang web',
     '05 phút',
     'CCĐG 1\n(Phụ lục 1.1)'],
    ['Tiết 1 của bài\n(Tiết PPCT: 20)\n[45 phút]',
     '2. Hoạt động 2: Hình thành kiến thức mới\n(Tổng: 62 phút)',
     'Nhiệm vụ 1: Tìm hiểu danh sách có thứ tự (<ol>, <li>) và danh sách không có thứ tự (<ul>, <li>)',
     '20 phút',
     'CCĐG 2\n(Phụ lục 1.2)'],
    ['Tiết 1 của bài\n(Tiết PPCT: 20)\n[45 phút]',
     '2. Hoạt động 2: Hình thành kiến thức mới\n(Tổng: 62 phút)',
     'Nhiệm vụ 2: Tìm hiểu danh sách mô tả (<dl>, <dt>, <dd>) và danh sách lồng nhau',
     '20 phút',
     'CCĐG 3\n(Phụ lục 1.3)'],
    ['Tiết 2 của bài\n(Tiết PPCT: 21)\n[45 phút]',
     '2. Hoạt động 2: Hình thành kiến thức mới\n(Tiếp theo)',
     'Nhiệm vụ 3: Tìm hiểu cấu trúc cơ bản của bảng trong HTML (<table>, <tr>, <th>, <td>, <caption>)',
     '10 phút',
     'CCĐG 4\n(Phụ lục 1.4)'],
    ['Tiết 2 của bài\n(Tiết PPCT: 21)\n[45 phút]',
     '2. Hoạt động 2: Hình thành kiến thức mới\n(Tiếp theo)',
     'Nhiệm vụ 4: Tìm hiểu cách định dạng bảng (style: border, width, height) và gộp ô (rowspan, colspan)',
     '12 phút',
     'CCĐG 5\n(Phụ lục 1.5)'],
    ['Tiết 2 của bài\n(Tiết PPCT: 21)\n[45 phút]',
     '3. Hoạt động 3: Luyện tập',
     'Thực hành tổng hợp: Tạo danh sách các CLB (Hình 9.7), bảng lịch hoạt động CLB (Hình 9.8) và bài tập Luyện tập SGK trang 55–56',
     '18 phút',
     'CCĐG 6 – Rubric\n(Phụ lục 2)'],
    ['Tiết 2 của bài\n(Tiết PPCT: 21)\n[45 phút]',
     '4. Hoạt động 4: Vận dụng',
     'Kiểm chứng gợi ý sửa lỗi cấu trúc bảng HTML của AI và giao nhiệm vụ Vận dụng thực tế (SGK trang 56)',
     '05 phút',
     'CCĐG 7\n(Phụ lục 3)']
]
b.append(table(TIME_OVERVIEW, [1450, 2100, 3250, 950, 1250], header=True, cell_sz=23))

# ------------------------------------------------------------------------------
# 1. HOẠT ĐỘNG 1: MỞ ĐẦU (05 phút – Tiết 1 của bài / Tiết PPCT: 20)
# ------------------------------------------------------------------------------
add(b, '1. Hoạt động 1: Mở đầu – Xác định vấn đề cần giải quyết (Thời lượng: 05 phút – Tiết 1 của bài / Tiết PPCT: 20)', 'heading')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT1]`):** Nhận biết được ý nghĩa, ưu điểm của việc sử dụng danh sách và bảng trong trình bày văn bản trên trang web; lựa chọn được dạng trình bày (danh sách hoặc bảng) phù hợp với đặc điểm dữ liệu.',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLd]`):** Chủ động quan sát tình huống so sánh, trao đổi nhanh với bạn cùng bàn để rút ra sự cần thiết của phần tử danh sách và bảng trong HTML.',
    '**Về phẩm chất (`[PC1]`):** Tích cực suy nghĩ, hào hứng tham gia phát biểu xây dựng bài ngay từ đầu tiết học.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Giáo viên trình chiếu hai tình huống trực quan từ Sách giáo khoa Tin học 12 (trang 52 – 53): (1) Tình huống Hoạt động 1 SGK trang 52: So sánh một trang web viết các mục nối tiếp dàn trải trong đoạn văn `<p>` (hoặc đánh số thủ công như Hình 8.3 trang 50) với trang web trình bày tách dòng có đánh số/kí hiệu đầu dòng tự động; (2) Tình huống Hoạt động 2 SGK trang 53 (Hội chợ ẩm thực của lớp 12E): Cần đăng lên trang web thông tin các món ăn gồm 4 thuộc tính (*Tên món ăn, Đơn giá, Số lượng, Tổng số tiền*). Học sinh thảo luận cặp đôi trả lời 02 câu hỏi định hướng:')
add(b, '- **Câu hỏi 1:** Khi liệt kê danh sách các câu lạc bộ của trường hoặc các bước giải một bài toán, cách trình bày theo danh sách có ưu điểm gì so với viết liền trong một đoạn văn bản?', first=283)
add(b, '- **Câu hỏi 2:** Trong tình huống Hội chợ ẩm thực (Hoạt động 2 SGK trang 53), lớp 12E nên dùng danh sách hay dùng bảng để biểu diễn các thông tin (*món ăn, đơn giá, số lượng, tổng số tiền*)? Vì sao?', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Câu trả lời miệng của học sinh chỉ ra được: (1) Danh sách giúp tách bạch từng mục tương tự nhau, tự động đánh số hoặc tạo kí hiệu đầu dòng rõ ràng, dễ theo dõi; (2) Tình huống Hội chợ ẩm thực cần dùng **bảng** vì mỗi đối tượng (món ăn) có tới 4 thuộc tính cần căn chỉnh thẳng hàng theo cột và hàng để dễ so sánh, tính toán.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 1 – Bảng kiểm đánh giá câu trả lời tình huống mở đầu** *(Xem chi tiết tại Phụ lục 1.1)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
HD1_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (01 phút)**\n'
        '- Giáo viên chiếu lên màn hình hai phương án trình bày trang web giới thiệu Câu lạc bộ và tình huống *Hội chợ ẩm thực lớp 12E* (SGK trang 52–53).\n'
        '- Giáo viên nêu 02 câu hỏi ở mục Nội dung và yêu cầu học sinh suy nghĩ cá nhân 30 giây, sau đó trao đổi nhanh theo cặp đôi trong 30 giây để lựa chọn phương án trình bày phù hợp.\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (02 phút)**\n'
        '- Học sinh quan sát màn hình chiếu và SGK trang 52–53, đối chiếu Hình 8.3 (trang 50) để nhận diện hạn chế của việc gõ dấu cộng/số thứ tự thủ công.\n'
        '- Học sinh trao đổi với bạn cùng máy để thống nhất câu trả lời cho Câu hỏi 1 và Câu hỏi 2.\n'
        '- Giáo viên quan sát và gợi mở cho các cặp còn lúng túng: *"Nếu mỗi món ăn có tới 4 thông tin đi kèm thì viết dạng danh sách một chiều có dễ cộng tổng tiền không?"*\n\n'
        '**Bước 3. Báo cáo sản phẩm (01 phút)**\n'
        '- Giáo viên mời đại diện 02 cặp đôi trình bày ngắn gọn câu trả lời trước lớp.\n'
        '- Các cặp đôi khác lắng nghe, nhận xét và bổ sung ý kiến.\n\n'
        '**Bước 4. Đánh giá, nhận xét (01 phút)**\n'
        '- Giáo viên đối chiếu với **Bảng kiểm CCĐG 1 (Phụ lục 1.1)** để nhận xét, biểu dương câu trả lời của học sinh.\n'
        '- Giáo viên chuẩn hóa ý nghĩa của danh sách và bảng trong HTML, từ đó dẫn dắt vào **Bài 9. Tạo danh sách, bảng** (thực hiện trong 02 tiết: Tiết PPCT 20 và 21) với nhiệm vụ xuyên suốt là xây dựng *Trang thông tin và Lịch hoạt động Câu lạc bộ trường THPT Thuận Hoá*.',

        '**BÀI 9. TẠO DANH SÁCH, BẢNG**\n'
        '*(Thời lượng: 02 tiết – Tiết PPCT: 20, 21)*\n\n'
        '**I. Ý nghĩa của danh sách và bảng trên trang web:**\n'
        '- **Danh sách (List):** Dùng để nhóm và liệt kê tập hợp các mục tương tự nhau theo một trình tự nhất định hoặc theo từng nhóm nội dung, giúp văn bản trên trang web mạch lạc, dễ quan sát.\n'
        '- **Bảng (Table):** Dùng để sắp xếp và biểu diễn dữ liệu có nhiều thuộc tính theo cấu trúc hàng và cột, giúp người đọc dễ dàng đối chiếu, so sánh thông tin.\n\n'
        '**Đáp án tình huống mở đầu:**\n'
        '- *Câu 1:* Dùng danh sách giúp tách rõ từng mục, tự động đánh số thứ tự hoặc tạo kí hiệu đầu dòng đồng nhất mà không cần gõ thủ công.\n'
        '- *Câu 2:* Tình huống Hội chợ ẩm thực lớp 12E nên dùng **bảng** vì mỗi món ăn gồm 4 thuộc tính (*món ăn, đơn giá, số lượng, tổng số tiền*) cần đặt trên cùng một hàng và gióng thẳng theo từng cột.'
    ]
]
b.append(table(HD1_ROWS, [4700, 4300], header=True, cell_sz=24))

# ------------------------------------------------------------------------------
# 2. HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (62 phút – Gồm 4 Nhiệm vụ: NV1, NV2, NV3, NV4)
# (Fixes Comments #4, #5, #6, #16, #20, #24, #27, #29)
# ------------------------------------------------------------------------------
add(b, '2. Hoạt động 2: Hình thành kiến thức mới – Tìm hiểu cách tạo danh sách và thiết lập bảng trong HTML (Tổng thời lượng: 62 phút; Thực hiện ở Tiết 1 [40 phút – Tiết PPCT: 20] và Tiết 2 [22 phút – Tiết PPCT: 21])', 'heading')
add(b, 'Hoạt động Hình thành kiến thức mới được tổ chức thống nhất xuyên suốt bài học và chia thành **04 tiểu mục (Nhiệm vụ 1, Nhiệm vụ 2, Nhiệm vụ 3, Nhiệm vụ 4)** theo đúng mạch kiến thức SGK trang 52 – 55, cụ thể như sau:')

# --- NHIỆM VỤ 1 (20 phút - Tiết 1 / Tiết PPCT: 20) ---
add(b, '2.1. Nhiệm vụ 1: Tìm hiểu danh sách có thứ tự và danh sách không có thứ tự (Thời lượng: 20 phút – Tiết 1 của bài / Tiết PPCT: 20)', 'sub')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT2]`):** Trình bày được cú pháp và sử dụng đúng các cặp thẻ `<ol>...</ol>`, `<ul>...</ul>`, `<li>...</li>` cùng thuộc tính `type`, `start` và đặc tính `list-style-type` trong thuộc tính `style` để tạo danh sách có thứ tự và danh sách không có thứ tự.',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLa]`, `[NLc]`, `[NLS 3.1.NC1.a]`):** Tự đọc hiểu mục 1a SGK trang 52, hoàn thành Phiếu học tập số 1 và thực hành gõ, lưu, chạy thử đoạn mã tạo danh sách có thứ tự / không có thứ tự trên máy tính.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`):** Chăm chỉ, cẩn thận khi đóng – mở đúng các cặp thẻ HTML và khai báo đúng tên thuộc tính.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Học sinh đọc mục **1a) Danh sách có hoặc không có thứ tự** (SGK trang 52), quan sát trực tiếp **Hình 9.1** và **Hình 9.2** dưới đây, kết hợp thảo luận cặp đôi để hoàn thành **Phiếu học tập số 1 (Phụ lục 1.2)** và gõ chạy thử trên máy tính:')
add_figure(b, 'hinh_9_1.png')
add_figure(b, 'hinh_9_2.png')
add(b, '- **Yêu cầu trong Phiếu học tập số 1:** (1) Xác định cặp thẻ bao ngoài của danh sách có thứ tự, danh sách không có thứ tự và cặp thẻ xác định từng mục bên trong danh sách; (2) Nêu các giá trị của thuộc tính `type` (`"1"`, `"A"`, `"a"`, `"I"`, `"i"`), ý nghĩa của thuộc tính `start` trong thẻ `<ol>` và dự đoán kí hiệu hiển thị của 2 mục đầu tiên khi viết `<ol type="a" start="3">`; (3) Nêu 4 giá trị của đặc tính `list-style-type` (`disc`, `circle`, `square`, `none`) trong thuộc tính `style` của thẻ `<ul>` và gõ chạy thử trên trình duyệt.', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Phiếu học tập số 1 đã điền đầy đủ đáp án đúng và đoạn mã HTML thử nghiệm `<ol type="a" start="3">`, `<ul style="list-style-type:square">` hiển thị đúng trên trình duyệt của học sinh.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 2 – Phiếu học tập số 1 và Bảng kiểm đánh giá sản phẩm Nhiệm vụ 1** *(Xem chi tiết tại Phụ lục 1.2)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
NV1_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (03 phút)**\n'
        '- Giáo viên liên hệ tính năng *Numbering* (danh sách đánh số) và *Bullets* (danh sách kí hiệu đầu dòng) trong MS Word, sau đó chiếu Hình 9.1 và Hình 9.2 (SGK trang 52).\n'
        '- Giáo viên phát **Phiếu học tập số 1 (Phụ lục 1.2)** và giao nhiệm vụ cho các cặp đôi: nghiên cứu mục 1a SGK trang 52, hoàn thành 3 câu hỏi trong Phiếu số 1 (05 phút), sau đó mở tệp `01_Khung_CLB.html` gõ thử nghiệm hai danh sách `<ol>` và `<ul>` với các thuộc tính `type`, `start`, `style="list-style-type:square"` (07 phút).\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (10 phút)**\n'
        '- Học sinh đọc SGK trang 52, trao đổi cặp đôi để điền câu trả lời vào Phiếu học tập số 1.\n'
        '- Học sinh mở trình soạn thảo mã nguồn, gõ đoạn mã tạo danh sách có thứ tự và không có thứ tự, lưu tệp (Ctrl + S) và mở trên trình duyệt (F5) để kiểm chứng kết quả của `type="a" start="3"` và `list-style-type:square`.\n'
        '- Giáo viên đi quanh phòng máy quan sát, nhắc học sinh viết đúng dấu ngoặc kép bao quanh giá trị thuộc tính và đóng đủ thẻ `</li>`, `</ol>`, `</ul>`.\n\n'
        '**Bước 3. Báo cáo sản phẩm (04 phút)**\n'
        '- Giáo viên mời đại diện 02 cặp đôi trình bày kết quả Phiếu học tập số 1 và chiếu màn hình kết quả chạy thử của máy mình.\n'
        '- Các cặp khác đối chiếu kết quả, nhận xét và giải thích vì sao `<ol type="a" start="3">` lại bắt đầu bằng chữ `c.` và `d.`.\n\n'
        '**Bước 4. Đánh giá, nhận xét (03 phút)**\n'
        '- Giáo viên nhận xét, đánh giá quá trình làm việc của học sinh theo **CCĐG 2 (Phụ lục 1.2)**.\n'
        '- Giáo viên chốt cú pháp chuẩn vào cột ghi bài (bên phải) và lưu ý học sinh ý nghĩa viết tắt tiếng Anh (`ol` = *Ordered List*, `ul` = *Unordered List*, `li` = *List Item*) để ghi nhớ chính xác.',

        '**1. TẠO DANH SÁCH TRONG HTML**\n'
        '**a) Danh sách có thứ tự và danh sách không có thứ tự**\n'
        '- Mỗi mục trong danh sách đều được đặt trong cặp thẻ `<li>...</li>` (*List Item*).\n\n'
        '***(1) Danh sách có thứ tự (`<ol>` – Ordered List):***\n'
        '- **Cú pháp:**\n'
        '`<ol type="kiểu_đánh_số" start="số_bắt_đầu">`\n'
        '`  <li>Mục thứ 1</li>`\n'
        '`  <li>Mục thứ 2</li>`\n'
        '`</ol>`\n'
        '- **Thuộc tính của `<ol>`:**\n'
        '  + `type`: Chọn kiểu đánh thứ tự, nhận 1 trong 5 giá trị: `"1"` (mặc định: 1, 2, 3...), `"A"` (chữ hoa: A, B, C...), `"a"` (chữ thường: a, b, c...), `"I"` (số La Mã hoa: I, II, III...), `"i"` (số La Mã thường: i, ii, iii...).\n'
        '  + `start`: Số nguyên xác định vị trí bắt đầu đếm (Ví dụ: `<ol type="a" start="3">` sẽ bắt đầu từ chữ **`c`**, mục tiếp theo là **`d`**).\n\n'
        '***(2) Danh sách không có thứ tự (`<ul>` – Unordered List):***\n'
        '- **Cú pháp:**\n'
        '`<ul style="list-style-type: kiểu_kí_tự;">`\n'
        '`  <li>Mục thứ 1</li>`\n'
        '`  <li>Mục thứ 2</li>`\n'
        '`</ul>`\n'
        '- **Đặc tính `list-style-type` trong thuộc tính `style`:**\n'
        '  Chọn kí hiệu đầu dòng gồm 4 giá trị: `disc` (chấm tròn đặc – mặc định), `circle` (vòng tròn rỗng), `square` (hình vuông đặc) và `none` (không hiển thị kí tự đầu dòng).'
    ]
]
b.append(table(NV1_ROWS, [4500, 4500], header=True, cell_sz=24))

# --- NHIỆM VỤ 2 (20 phút - Tiết 1 / Tiết PPCT: 20) ---
add(b, '2.2. Nhiệm vụ 2: Tìm hiểu danh sách mô tả và danh sách lồng nhau (Thời lượng: 20 phút – Tiết 1 của bài / Tiết PPCT: 20)', 'sub')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT3]`):** Trình bày được cú pháp và sử dụng đúng bộ thẻ `<dl>`, `<dt>`, `<dd>` để tạo danh sách mô tả; thực hiện được cách đặt danh sách con bên trong phần tử `<li>...</li>` của danh sách cha để tạo danh sách lồng nhau nhiều cấp.',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLC3]`, `[NLa]`, `[NLc]`, `[NLS 3.1.NC1.a]`):** Phân tích được Hình 9.3, trả lời được câu hỏi củng cố SGK trang 53 (tạo danh sách lồng 2 cấp) và lưu được tệp thực hành danh sách cuối Tiết 1.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`):** Kiên trì, tỉ mỉ khi căn lề thụt dòng mã nguồn để kiểm soát đúng vị trí đóng – mở thẻ `<li>` cha.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Học sinh đọc mục **1b) Danh sách mô tả**, phần *Lưu ý* về danh sách lồng nhau và câu hỏi củng cố cuối mục 1 (SGK trang 53), quan sát trực tiếp **Hình 9.3** dưới đây để hoàn thành **Phiếu học tập số 2 (Phụ lục 1.3)** và thực hành viết mã danh sách mô tả kết hợp danh sách lồng:')
add_figure(b, 'hinh_9_3.png')
add(b, '- **Yêu cầu trong Phiếu học tập số 2:** (1) Phân biệt vai trò của 3 thẻ `<dl>`, `<dt>`, `<dd>` trong Hình 9.3 và cho biết danh sách mô tả có sử dụng thẻ `<li>` hay không; (2) Giải quyết câu hỏi SGK trang 53: *"Làm thế nào để tạo một danh sách lồng nhau, danh sách mức 1 đánh số dạng 1, 2, 3,... và danh sách mức 2 đánh số dạng a, b, c?"* – Chỉ rõ danh sách mức 2 (`<ol type="a">`) phải đặt ở vị trí nào so với thẻ đóng `</li>` của mục cha; (3) Gõ thử nghiệm 01 danh sách mô tả giới thiệu Câu lạc bộ và 01 danh sách lồng 2 cấp cho nhóm *Thể thao* (mục *Võ* gồm *a. Karatedo, b. Taekwondo, c. Vovinam*), lưu tệp `HoTen_12-4_Bai9.html` trước khi kết thúc Tiết 1.', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Phiếu học tập số 2 đã hoàn thành chính xác; tệp `HoTen_12-4_Bai9.html` lưu trên máy có đoạn mã danh sách mô tả `<dl>/<dt>/<dd>` và danh sách lồng 2 cấp hiển thị đúng cấu trúc phân cấp.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 3 – Phiếu học tập số 2 và Bảng kiểm đánh giá sản phẩm Nhiệm vụ 2** *(Xem chi tiết tại Phụ lục 1.3)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
NV2_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (03 phút)**\n'
        '- Giáo viên chiếu Hình 9.3 (SGK trang 53) và nêu vấn đề: *"Khi cần liệt kê tên một món ăn hoặc tên một Câu lạc bộ kèm theo lời giải thích chi tiết bên dưới, hoặc khi trong mục Võ lại có các môn con a, b, c thì ta dùng cấu trúc thẻ nào?"*\n'
        '- Giáo viên phát **Phiếu học tập số 2 (Phụ lục 1.3)**, yêu cầu học sinh đọc mục 1b SGK trang 53, trả lời câu hỏi trong phiếu (05 phút) và thực hành gõ danh sách mô tả + danh sách lồng 2 cấp vào tệp `HoTen_12-4_Bai9.html` (07 phút).\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (10 phút)**\n'
        '- Học sinh nghiên cứu Hình 9.3 và phần *Lưu ý* trang 53, thảo luận cặp đôi hoàn thành Phiếu học tập số 2.\n'
        '- Học sinh thực hành gõ cặp thẻ `<dl>`, `<dt>`, `<dd>` và thử nghiệm lồng danh sách `<ol type="a">` vào bên trong phần tử `<li>Võ ... </li>`, lưu tệp và tải lại trình duyệt.\n'
        '- Giáo viên quan sát, phát hiện lỗi phổ biến (đóng `</li>` của mục cha trước rồi mới mở `<ol>` con) để hướng dẫn học sinh tự đối chiếu và sửa lỗi.\n\n'
        '**Bước 3. Báo cáo sản phẩm (04 phút)**\n'
        '- Giáo viên mời 01 cặp đôi trình bày đáp án câu hỏi củng cố SGK trang 53 và chỉ trực tiếp trên đoạn mã vị trí đặt danh sách con.\n'
        '- Một cặp đôi khác nhận xét và phân biệt sự khác nhau giữa `<dl>` (dùng `<dt>`, `<dd>`) với `<ol>`/`<ul>` (dùng `<li>`).\n\n'
        '**Bước 4. Đánh giá, nhận xét (03 phút)**\n'
        '- Giáo viên đánh giá kết quả theo **CCĐG 3 (Phụ lục 1.3)**, chuẩn hóa kiến thức ghi bài (cột phải) và nhắc học sinh kiểm tra việc lưu tệp `HoTen_12-4_Bai9.html` để đầu Tiết 2 (Tiết PPCT 21) tiếp tục tìm hiểu phần Thiết lập bảng.',

        '**b) Danh sách mô tả (`<dl>` – Description List)**\n'
        '- Dùng để liệt kê các mục (thuật ngữ, tên gọi) kèm theo lời mô tả, giải thích cho từng mục (**không** dùng thẻ `<li>`).\n'
        '- **Cú pháp:**\n'
        '`<dl>`\n'
        '`  <dt>Tên mục 1</dt>`\n'
        '`  <dd>Mô tả chi tiết của mục 1</dd>`\n'
        '`  <dt>Tên mục 2</dt>`\n'
        '`  <dd>Mô tả chi tiết của mục 2</dd>`\n'
        '`</dl>`\n'
        '*(Trong đó: `<dl>` khai báo danh sách mô tả; `<dt>` xác định tên mục/thuật ngữ; `<dd>` xác định nội dung mô tả).* \n\n'
        '**c) Danh sách lồng nhau (Nested Lists) & Đáp án câu hỏi SGK trang 53:**\n'
        '- **Nguyên tắc:** Để tạo danh sách lồng nhau, ta đặt toàn bộ danh sách con (`<ol>` hoặc `<ul>`) **nằm bên trong cặp thẻ `<li>...</li>` của mục cha tương ứng** (trước thẻ đóng `</li>` của mục cha).\n'
        '- **Mẫu mã trả lời câu hỏi SGK trang 53 (Mức 1 đánh số 1, 2, 3; Mức 2 đánh chữ a, b, c):**\n'
        '`<ol type="1">`\n'
        '`  <li>Mục lớn 1`\n'
        '`    <ol type="a">`\n'
        '`      <li>Mục con a</li>`\n'
        '`      <li>Mục con b</li>`\n'
        '`    </ol>`\n'
        '`  </li>`\n'
        '`  <li>Mục lớn 2</li>`\n'
        '`</ol>`'
    ]
]
b.append(table(NV2_ROWS, [4500, 4500], header=True, cell_sz=24))

# --- NHIỆM VỤ 3 (10 phút - Tiết 2 / Tiết PPCT: 21) ---
add(b, '2.3. Nhiệm vụ 3: Tìm hiểu cấu trúc cơ bản của bảng trong HTML (Thời lượng: 10 phút – Tiết 2 của bài / Tiết PPCT: 21)', 'sub')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT4]`):** Trình bày được cú pháp và sử dụng đúng các cặp thẻ `<table>`, `<tr>`, `<th>`, `<td>`, `<caption>` để khởi tạo cấu trúc bảng HTML cơ bản có tiêu đề bảng, hàng tiêu đề và các hàng dữ liệu.',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLa]`, `[NLc]`, `[NLS 3.1.NC1.a]`):** Đọc hiểu sơ đồ cấu trúc bảng ở Hình 9.4 (SGK trang 54), hoàn thành Phiếu học tập số 3 và tạo được bảng HTML cơ bản trên máy tính.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`):** Cẩn thận, tuân thủ đúng thứ tự phân cấp `<table>` $\\rightarrow$ `<tr>` $\\rightarrow$ `<th>`/`<td>`.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Mở đầu Tiết 2 (Tiết PPCT 21), học sinh nghiên cứu phần đầu mục **2. Thiết lập bảng** (SGK trang 53 – 54), quan sát trực tiếp **Hình 9.4** dưới đây và hoàn thành **Phiếu học tập số 3 (Phụ lục 1.4)**:')
add_figure(b, 'hinh_9_4.png')
add(b, '- **Yêu cầu trong Phiếu học tập số 3:** (1) Ghép đúng chức năng của 5 cặp thẻ tạo bảng: `<table>...</table>`, `<caption>...</caption>`, `<tr>...</tr>`, `<th>...</th>`, `<td>...</td>`; (2) Xác định vị trí bắt buộc của cặp thẻ `<caption>` trong bảng; (3) Mở tệp `02_CLB_Mau_CoBan.html`, bổ sung thêm 01 hàng dữ liệu mới bằng cặp thẻ `<tr>` và `<td>`, lưu và quan sát sự khác biệt về định dạng hiển thị mặc định giữa ô tiêu đề `<th>` (chữ in đậm, căn giữa) với ô dữ liệu `<td>` (chữ thường, căn trái).', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Phiếu học tập số 3 hoàn chỉnh và đoạn mã bảng cơ bản có `<caption>`, hàng tiêu đề `<th>` và các hàng dữ liệu `<td>` chạy đúng trên trình duyệt.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 4 – Phiếu học tập số 3 và Bảng kiểm đánh giá sản phẩm Nhiệm vụ 3** *(Xem chi tiết tại Phụ lục 1.4)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
NV3_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (02 phút)**\n'
        '- Giáo viên chiếu Hình 9.4 (SGK trang 54), phát **Phiếu học tập số 3 (Phụ lục 1.4)** và yêu cầu học sinh quan sát cách 1 bảng dữ liệu được xây dựng lần lượt theo từng hàng (`<tr>`) từ trên xuống dưới, trong mỗi hàng gồm các ô (`<th>` hoặc `<td>`) từ trái sang phải.\n'
        '- Giáo viên giao nhiệm vụ hoàn thành Phiếu số 3 và thử nghiệm thêm hàng dữ liệu vào tệp `02_CLB_Mau_CoBan.html`.\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (04 phút)**\n'
        '- Học sinh đọc SGK trang 53–54, trao đổi cặp đôi hoàn thành Phiếu học tập số 3.\n'
        '- Học sinh mở tệp `02_CLB_Mau_CoBan.html`, thêm thẻ `<caption>` và 01 hàng `<tr><td>...</td></tr>`, lưu tệp và mở trên trình duyệt.\n'
        '- Giáo viên quan sát, nhắc học sinh tuyệt đối không đặt thẻ `<td>` hoặc `<th>` nằm ngoài cặp thẻ hàng `<tr>...</tr>`.\n\n'
        '**Bước 3. Báo cáo sản phẩm (02 phút)**\n'
        '- Một cặp đôi trình bày kết quả ghép thẻ trong Phiếu số 3, chỉ rõ vị trí đặt thẻ `<caption>` và sự khác nhau giữa `<th>` với `<td>`.\n'
        '- Các cặp khác nhận xét, bổ sung.\n\n'
        '**Bước 4. Đánh giá, nhận xét (02 phút)**\n'
        '- Giáo viên đánh giá theo **CCĐG 4 (Phụ lục 1.4)** và chốt cú pháp khởi tạo bảng cơ bản vào cột ghi bài.\n'
        '- Giáo viên đặt vấn đề chuyển tiếp: *"Bảng ở Hình 9.4 mặc định chưa có đường viền ngăn cách các ô và chưa gộp được các ô tiêu đề chung. Làm thế nào để kẻ khung viền và gộp ô? Chúng ta sang Nhiệm vụ 4."*',

        '**2. THIẾT LẬP BẢNG TRONG HTML**\n'
        '**a) Cấu trúc cơ bản của bảng HTML**\n'
        '- Trong HTML, bảng được tạo bởi cặp thẻ `<table>...</table>`, xây dựng theo từng **hàng** từ trên xuống dưới; mỗi hàng gồm các **ô** từ trái sang phải:\n'
        '  + `<table>...</table>`: Khai báo phần tử bảng.\n'
        '  + `<caption>...</caption>`: Tạo tiêu đề (chú thích) cho bảng; đặt **ngay sau thẻ mở `<table>`** và trước thẻ `<tr>` đầu tiên.\n'
        '  + `<tr>...</tr>` (*Table Row*): Tạo một hàng trong bảng.\n'
        '  + `<th>...</th>` (*Table Header*): Tạo ô tiêu đề (mặc định chữ **in đậm**, căn giữa ô).\n'
        '  + `<td>...</td>` (*Table Data*): Tạo ô dữ liệu thông thường (mặc định chữ thường, căn trái).\n\n'
        '- **Cú pháp tổng quát:**\n'
        '`<table>`\n'
        '`  <caption>Tiêu đề bảng</caption>`\n'
        '`  <tr><th>Cột 1</th><th>Cột 2</th></tr>`\n'
        '`  <tr><td>Ô 1</td><td>Ô 2</td></tr>`\n'
        '`</table>`'
    ]
]
b.append(table(NV3_ROWS, [4500, 4500], header=True, cell_sz=24))

# --- NHIỆM VỤ 4 (12 phút - Tiết 2 / Tiết PPCT: 21) ---
add(b, '2.4. Nhiệm vụ 4: Tìm hiểu cách định dạng bảng và gộp ô trong bảng HTML (Thời lượng: 12 phút – Tiết 2 của bài / Tiết PPCT: 21)', 'sub')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT5]`):** Sử dụng được thuộc tính `style` (`border`, `width`, `height`) để tạo khung viền cho bảng/ô, điều chỉnh kích thước và sử dụng đúng thuộc tính `rowspan`, `colspan` để gộp hàng, gộp cột trong bảng.',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLC3]`, `[NLa]`, `[NLc]`, `[NLS 3.2.NC1.a]`):** Phân tích được Hình 9.5, Hình 9.6, trả lời câu hỏi SGK trang 55 và thiết lập đúng 2 hàng tiêu đề có gộp ô của Hình 9.8.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`):** Cẩn thận khi tính toán số ô còn lại ở các hàng tiếp theo sau khi đã sử dụng `rowspan` và `colspan`.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Học sinh đọc phần định dạng bảng và gộp ô (SGK trang 54 – 55), quan sát trực tiếp **Hình 9.5** và **Hình 9.6** dưới đây, trả lời câu hỏi củng cố trang 55 và hoàn thành **Phiếu học tập số 4 (Phụ lục 1.5)**:')
add_figure(b, 'hinh_9_5.png')
add_figure(b, 'hinh_9_6.png')
add(b, '- **Yêu cầu trong Phiếu học tập số 4:** (1) Nêu cú pháp tạo đường viền bằng đặc tính `border` trong thuộc tính `style` và trả lời câu hỏi SGK trang 55: *"Bảng trong ví dụ trên Hình 9.6 có nhược điểm gì khi dùng `border=\"1\"`? Cần làm thế nào để giải quyết nhược điểm đó?"*; (2) Phân biệt `rowspan="k"` (gộp `k` hàng theo chiều dọc) với `colspan="m"` (gộp `m` cột theo chiều ngang) và nguyên tắc lược bỏ thẻ `<th>`/`<td>` ở vị trí đã bị ô gộp chiếm chỗ; (3) Áp dụng ngay để viết đoạn mã tạo 02 hàng tiêu đề của bảng *Lịch hoạt động CLB Thể thao* (Hình 9.8 trang 56): hàng 1 có `Bộ môn` (`rowspan="2"`), `GV phụ trách` (`rowspan="2"`), `Ngày` (`colspan="5"`); giải thích vì sao hàng 2 chỉ viết 5 cặp thẻ `<th>` (*Thứ 2* đến *Thứ 6*).', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Phiếu học tập số 4 hoàn chỉnh; đoạn mã HTML tạo 2 hàng tiêu đề bảng Lịch hoạt động CLB có viền `style="border: 1px solid ..."` và gộp ô `rowspan="2"`, `colspan="5"` hiển thị chính xác 7 cột quy đổi.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 5 – Phiếu học tập số 4 và Bảng kiểm đánh giá sản phẩm Nhiệm vụ 4** *(Xem chi tiết tại Phụ lục 1.5)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
NV4_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (02 phút)**\n'
        '- Giáo viên chiếu Hình 9.5 và Hình 9.6 (SGK trang 54–55), phát **Phiếu học tập số 4 (Phụ lục 1.5)**.\n'
        '- Giáo viên yêu cầu các cặp đôi tìm hiểu cách dùng thuộc tính `style` (`border`, `width`, `height`), quy tắc gộp ô `rowspan`, `colspan` và gõ thử nghiệm 2 hàng tiêu đề của bảng Lịch CLB.\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (06 phút)**\n'
        '- Học sinh đọc SGK trang 54–55, thảo luận trả lời các câu hỏi trong Phiếu học tập số 4.\n'
        '- Học sinh gõ thử nghiệm đoạn mã tạo khung viền bằng `style="border: 1px solid black;"` và 2 hàng tiêu đề có `rowspan="2"`, `colspan="5"` trên máy tính.\n'
        '- Giáo viên hướng dẫn học sinh quy tắc kiểm tra nhanh: *"Tổng số cột quy đổi của mọi hàng trong bảng phải luôn bằng nhau (bằng 7 cột). Ở hàng 2, vì 2 cột đầu đã bị Bộ môn và GV phụ trách ở hàng 1 gộp xuống nên ta chỉ được viết đúng 5 thẻ `<th>` từ Thứ 2 đến Thứ 6."*\n\n'
        '**Bước 3. Báo cáo sản phẩm (02 phút)**\n'
        '- Hai cặp đôi trình bày đáp án Phiếu số 4, giải thích câu hỏi SGK trang 55 và chiếu đoạn mã 2 hàng tiêu đề bảng.\n'
        '- Giáo viên chiếu 01 trường hợp lỗi thừa ô ở hàng thứ hai (khi học sinh quên không bỏ qua vị trí đã gộp `rowspan="2"`) để cả lớp nhận diện.\n\n'
        '**Bước 4. Đánh giá, nhận xét (02 phút)**\n'
        '- Giáo viên đánh giá theo **CCĐG 5 (Phụ lục 1.5)** và chuẩn hóa kiến thức mục 2b vào cột ghi bài.\n'
        '- Như vậy học sinh đã hoàn thành toàn bộ phần **Hoạt động 2: Hình thành kiến thức mới** của cả bài học, sẵn sàng bước sang **Hoạt động 3: Luyện tập**.',

        '**b) Định dạng bảng và gộp ô trong bảng HTML**\n'
        '***(1) Tạo khung viền và điều chỉnh kích thước bảng (Thuộc tính `style`):***\n'
        '- **Tạo viền (`border`):** Đặt trong thuộc tính `style` của thẻ `<table>`, `<th>`, `<td>` theo cú pháp:\n'
        '`style="border: độ_dày_px kiểu_viền [màu_viền];"`\n'
        '  *(Kiểu viền gồm: `solid` – nét liền, `dotted` – chấm rời, `double` – nét đôi, `none` – không viền).* \n'
        '  *Lưu ý (Đáp án câu hỏi SGK trang 55):* Trong HTML5, thuộc tính `border="1"` không còn được khuyến khích và tạo viền kép thô; cần dùng thuộc tính `style="border: 1px solid ...;"` cho cả `<table>` và từng ô `<th>`, `<td>` (kết hợp `border-collapse: collapse;` nếu muốn gộp viền đơn).\n'
        '- **Kích thước (`width`, `height`):** Đặt chiều rộng, chiều cao theo tỉ lệ `%` (khuyến nghị) hoặc điểm ảnh `px`:\n'
        '`<table style="width: 80%; height: 400px;">`\n\n'
        '***(2) Gộp hàng và gộp cột (`rowspan` và `colspan`):***\n'
        '- Đặt trong thẻ mở `<th>` hoặc `<td>` tại vị trí ô đầu tiên cần gộp:\n'
        '  + `rowspan="số_hàng"`: Gộp các ô liên tiếp theo chiều dọc (hàng). Ở các hàng kế tiếp bên dưới, **phải bỏ qua** không viết lại cặp thẻ `<th>`/`<td>` tại vị trí đã bị chiếm.\n'
        '  + `colspan="số_cột"`: Gộp các ô liên tiếp theo chiều ngang (cột). Trong cùng hàng đó, giảm số thẻ `<th>`/`<td>` tương ứng.\n'
        '- **Ví dụ (Hình 9.6 SGK trang 55):**\n'
        '`<tr>`\n'
        '`  <th rowspan="2">Họ và tên</th>`\n'
        '`  <th colspan="3">Điểm thi</th>`\n'
        '`</tr>`\n'
        '`<tr>`\n'
        '`  <td>Toán</td><td>Vật lí</td><td>Hoá học</td>`\n'
        '`</tr>`'
    ]
]
b.append(table(NV4_ROWS, [4500, 4500], header=True, cell_sz=24))

# ------------------------------------------------------------------------------
# 3. HOẠT ĐỘNG 3: LUYỆN TẬP (18 phút – Tiết 2 của bài / Tiết PPCT: 21)
# (Fixes Comments #5, #20, #22, #29, #30, #33, #34)
# ------------------------------------------------------------------------------
add(b, '3. Hoạt động 3: Luyện tập – Thực hành tạo danh sách và bảng cho trang web Câu lạc bộ (Thời lượng: 18 phút – Tiết 2 của bài / Tiết PPCT: 21)', 'heading')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT2]`, `[KT3]`, `[KT4]`, `[KT5]`, `[KT6]`):** Vận dụng tổng hợp toàn bộ kiến thức về danh sách (`<ul>`, `<ol>`, `<li>`, `<dl>`, `<dt>`, `<dd>`, danh sách lồng) và bảng (`<table>`, `<caption>`, `<tr>`, `<th>`, `<td>`, `style`, `rowspan`, `colspan`) đã học ở Hoạt động 2 để hoàn thiện trang web Câu lạc bộ theo đúng Nhiệm vụ 1 (Hình 9.7), Nhiệm vụ 2 (Hình 9.8) và bài tập Luyện tập (SGK trang 55 – 56).',
    '**Về năng lực (`[NLC1]`, `[NLC2]`, `[NLC3]`, `[NLa]`, `[NLc]`, `[NLS 3.1.NC1.a]`, `[NLS 3.2.NC1.a]`):** Tạo, chỉnh sửa, tích hợp danh sách và bảng trong cùng một tệp HTML (`HoTen_12-4_Bai9.html`), lưu mã hóa UTF-8 và kiểm tra chéo sản phẩm theo bảng tiêu chí Rubric.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`, `[PC3]`):** Chăm chỉ, trách nhiệm hoàn thành sản phẩm thực hành và trung thực khi tự đánh giá, đánh giá đồng đẳng.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Học sinh thực hiện trọn vẹn phần **3. Thực hành tạo danh sách và bảng** (Nhiệm vụ 1, Nhiệm vụ 2 SGK trang 55 – 56) và bài tập **Luyện tập** (SGK trang 56) trực tiếp trên máy tính theo **Hình 9.7** và **Hình 9.8** dưới đây:')
add_figure(b, 'hinh_9_7.png')
add_figure(b, 'hinh_9_8.png')
add(b, '- **Yêu cầu 1 (Hoàn thiện Nhiệm vụ 1 SGK trang 55 – Hình 9.7):** Rà soát và hoàn thiện phần *Danh sách Câu lạc bộ* trong tệp `HoTen_12-4_Bai9.html` gồm 2 mục lớn không có thứ tự (`<ul>`: *THỂ THAO*, *NGHỆ THUẬT*); bên trong mỗi mục lớn là danh sách có thứ tự kiểu số (`<ol type="1">`); riêng mục *4. Võ* (*a. Karatedo, b. Taekwondo, c. Vovinam*) và mục *3. Âm nhạc* (*a. Thanh nhạc, b. Piano, c. Violin*) có danh sách con lồng bên trong đánh chữ cái thường (`<ol type="a">`), kèm ít nhất 01 cặp mô tả `<dl>/<dt>/<dd>`.', first=283)
add(b, '- **Yêu cầu 2 (Thực hiện Nhiệm vụ 2 SGK trang 56 – Hình 9.8 & Bài tập Luyện tập trang 56):** Bổ sung vào cùng tệp `HoTen_12-4_Bai9.html` bảng *Lịch hoạt động CLB Thể thao* gồm 7 hàng, 7 cột quy đổi (02 hàng tiêu đề có `rowspan="2"` cho cột *Bộ môn*, *GV phụ trách* và `colspan="5"` cho cột *Ngày*; 05 hàng dữ liệu cho *Bóng đá, Bóng chuyền, Bóng rổ, Võ thuật, Bơi*). Sử dụng thuộc tính `style` thay cho `border="1"` để kẻ viền bảng/ô, trong đó **tô viền màu xanh (`border: 1px solid green;`) cho các ô thuộc hai dòng tiêu đề đầu** và đổi màu chữ (đỏ, vàng/cam đậm, xanh) theo yêu cầu mục **Luyện tập SGK trang 56**.', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Tệp sản phẩm hoàn chỉnh `HoTen_12-4_Bai9.html` của mỗi học sinh/cặp đôi (đối chiếu với tệp mẫu chuẩn `05_CLB_HoanChinh_SGK.html` tại **Phụ lục 4**), mở trên trình duyệt hiển thị đầy đủ cả Danh sách Câu lạc bộ (Hình 9.7) và Bảng Lịch hoạt động CLB Thể thao (Hình 9.8) có định dạng màu viền/màu chữ bằng `style`.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 6 – Phiếu đánh giá sản phẩm thực hành (Bảng tiêu chí Rubric 10 điểm)** *(Xem chi tiết tại Phụ lục 2)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
HD3_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (02 phút)**\n'
        '- Giáo viên chiếu đồng thời Hình 9.7, Hình 9.8 (SGK trang 55–56) và công bố **Phiếu đánh giá sản phẩm thực hành – Rubric 10 điểm (CCĐG 6 tại Phụ lục 2)**.\n'
        '- Giáo viên giao nhiệm vụ cho học sinh mở lại tệp `HoTen_12-4_Bai9.html`, tích hợp phần danh sách câu lạc bộ và phần tiêu đề bảng đã khởi tạo ở Hoạt động 2, tiếp tục gõ hoàn thiện 5 hàng dữ liệu lịch hoạt động và áp dụng định dạng `style` của bài tập Luyện tập trang 56.\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (12 phút)**\n'
        '- Học sinh thực hành trên máy tính: hoàn thiện danh sách 3 cấp (Hình 9.7) và bảng 7 hàng × 7 cột quy đổi (Hình 9.8), thêm thuộc tính `style="border: 1px solid green;"` cho 2 hàng tiêu đề đầu và định dạng màu chữ.\n'
        '- Giáo viên thực hiện dạy học phân hóa: cấp khung dữ liệu văn bản thô (`01_Khung_CLB.html`) cho các học sinh gõ phím chậm để các em tập trung viết đúng cấu trúc thẻ HTML; khuyến khích học sinh hoàn thành sớm thực hiện đầy đủ bài tập đổi màu viền/màu chữ ở mục Luyện tập trang 56.\n'
        '- Trong 02 phút cuối của bước thực hiện, hai cặp đôi ngồi cạnh nhau đổi máy kiểm tra chéo sản phẩm dựa theo 7 tiêu chí trong **Rubric (Phụ lục 2)**.\n\n'
        '**Bước 3. Báo cáo sản phẩm (02 phút)**\n'
        '- Giáo viên mời 02 cặp đôi chiếu trang web `HoTen_12-4_Bai9.html` và mã nguồn HTML lên màn hình lớn, giải thích cách lồng danh sách ở mục *Võ / Âm nhạc* và cách gộp ô ở 2 hàng tiêu đề bảng.\n'
        '- Cặp đôi chấm chéo báo cáo điểm đánh giá theo Rubric.\n\n'
        '**Bước 4. Đánh giá, nhận xét (02 phút)**\n'
        '- Giáo viên đối chiếu sản phẩm của học sinh với tệp mẫu `05_CLB_HoanChinh_SGK.html` và **Rubric 10 điểm (CCĐG 6 – Phụ lục 2)** để nhận xét, chấm điểm thường xuyên cho các cặp làm tốt.\n'
        '- Giáo viên chọn 01 tình huống lỗi cấu trúc bảng điển hình để dẫn sang **Hoạt động 4: Vận dụng**.',

        '**3. THỰC HÀNH TẠO DANH SÁCH VÀ BẢNG & LUYỆN TẬP (SGK trang 55 – 56)**\n\n'
        '**a) Cấu trúc mã HTML tạo Danh sách Câu lạc bộ (Nhiệm vụ 1 – Hình 9.7):**\n'
        '`<h2>Danh sách Câu lạc bộ</h2>`\n'
        '`<ul>`\n'
        '`  <li>THỂ THAO`\n'
        '`    <ol type="1">`\n'
        '`      <li>Bóng đá</li>`\n'
        '`      <li>Bóng chuyền</li>`\n'
        '`      <li>Bóng rổ</li>`\n'
        '`      <li>Võ`\n'
        '`        <ol type="a">`\n'
        '`          <li>Karatedo</li>`\n'
        '`          <li>Taekwondo</li>`\n'
        '`          <li>Vovinam</li>`\n'
        '`        </ol>`\n'
        '`      </li>`\n'
        '`      <li>Bơi</li>`\n'
        '`    </ol>`\n'
        '`  </li>`\n'
        '`  <li>NGHỆ THUẬT ... </li>`\n'
        '`</ul>`\n\n'
        '**b) Cấu trúc mã HTML tạo Bảng Lịch hoạt động CLB & Luyện tập (Hình 9.8 & Luyện tập trang 56):**\n'
        '`<table style="border: 1px solid black; width: 90%;">`\n'
        '`  <caption>Lịch hoạt động CLB Thể thao</caption>`\n'
        '`  <tr>`\n'
        '`    <th rowspan="2" style="border: 1px solid green;">Bộ môn</th>`\n'
        '`    <th rowspan="2" style="border: 1px solid green;">GV phụ trách</th>`\n'
        '`    <th colspan="5" style="border: 1px solid green;">Ngày</th>`\n'
        '`  </tr>`\n'
        '`  <tr>`\n'
        '`    <th style="border: 1px solid green;">Thứ 2</th>`\n'
        '`    <th style="border: 1px solid green;">Thứ 3</th>`\n'
        '`    <th style="border: 1px solid green;">Thứ 4</th>`\n'
        '`    <th style="border: 1px solid green;">Thứ 5</th>`\n'
        '`    <th style="border: 1px solid green;">Thứ 6</th>`\n'
        '`  </tr>`\n'
        '`  <tr>`\n'
        '`    <td style="border: 1px solid black; color: red;">Bóng đá</td>`\n'
        '`    <td style="border: 1px solid black;">Thầy Khải</td>`\n'
        '`    <td style="border: 1px solid black;"></td>`\n'
        '`    <td style="border: 1px solid black;">X</td>`\n'
        '`    <td style="border: 1px solid black;"></td>`\n'
        '`    <td style="border: 1px solid black;"></td>`\n'
        '`    <td style="border: 1px solid black;">X</td>`\n'
        '`  </tr>`\n'
        '`  <!-- Các hàng Bóng chuyền, Bóng rổ, Võ thuật, Bơi tương tự -->`\n'
        '`</table>`'
    ]
]
b.append(table(HD3_ROWS, [4300, 4700], header=True, cell_sz=23))

# ------------------------------------------------------------------------------
# 4. HOẠT ĐỘNG 4: VẬN DỤNG (05 phút – Tiết 2 của bài / Tiết PPCT: 21)
# (Fixes Comments #31, #32, #33)
# ------------------------------------------------------------------------------
add(b, '4. Hoạt động 4: Vận dụng – Kiểm chứng gợi ý sửa lỗi cấu trúc HTML của AI và mở rộng bài toán thực tiễn (Thời lượng: 05 phút – Tiết 2 của bài / Tiết PPCT: 21)', 'heading')
add(b, 'a) Mục tiêu', 'label')
for item in [
    '**Về kiến thức (`[KT6]`):** Phát hiện và sửa được lỗi cấu trúc thẻ bảng HTML (`<td>` đặt ngoài `<tr>`); biết đánh giá, kiểm chứng gợi ý của công cụ AI dựa trên chuẩn kiến thức SGK và định hướng được cách giải bài toán Vận dụng SGK trang 56.',
    '**Về năng lực (`[NLC1]`, `[NLC3]`, `[NLc]`, `[NLd]`, `[NLAI 12.C2.2]`):** Viết được câu lệnh hỏi (Prompt) rõ ràng yêu cầu AI kiểm tra cấu trúc mã HTML; biết bác bỏ gợi ý sai/chưa đủ và tự kiểm chứng bằng thực nghiệm chạy thử.',
    '**Về phẩm chất (`[PC1]`, `[PC2]`, `[PC3]`):** Trung thực, có tư duy phản biện và trách nhiệm khi sử dụng công cụ AI trong học tập.'
]:
    add(b, item, 'bullet')

add(b, 'b) Nội dung', 'label')
add(b, 'Học sinh thực hiện **Phiếu học tập số 5 (Phụ lục 3)** với 02 nhiệm vụ gắn kết giữa kiểm chứng mã HTML và mở rộng thực tiễn:')
add(b, '- **Nhiệm vụ 4.1 (Kiểm chứng gợi ý AI trên lớp – Tệp `03_Ma_Can_Sua.html`):** Cho đoạn mã bảng bị lỗi cấu trúc (`<td>Bóng đá</td><td>Thứ Ba</td>` đặt trực tiếp trong `<table>` mà thiếu cặp thẻ bao `<tr>...</tr>`). Một công cụ AI đưa ra gợi ý giả định: *"Chỉ cần thêm thuộc tính `style=\"border: 1px solid black;\"` vào thẻ `<table>` là bảng sẽ đúng."* Học sinh phát hiện lỗi cấu trúc thật sự, giải thích vì sao gợi ý trên chưa giải quyết đúng bản chất lỗi, viết lại Prompt chuẩn và sửa mã thành `04_Ma_Da_Sua.html`.', first=283)
add(b, '- **Nhiệm vụ 4.2 (Nhiệm vụ Vận dụng về nhà – Mục Vận dụng SGK trang 56):** Cho trước bảng dữ liệu điểm thi cỡ $n \\times 4$ (*Họ và tên, Điểm Toán, Điểm Vật lí, Điểm Hoá học*). Em hãy tạo trang web Thời khóa biểu của lớp 12/4 hoặc viết chương trình Python đọc dữ liệu và tự động sinh ra tệp `.html` chứa mã bảng có cấu trúc gộp ô như Hình 9.6b.', first=283)

add(b, 'c) Sản phẩm', 'label')
for item in [
    'Phiếu học tập số 5 hoàn chỉnh (chỉ rõ lỗi thiếu thẻ `<tr>...</tr>`, bác bỏ gợi ý chỉ thêm viền, viết được Prompt chuẩn) và định hướng giải bài toán Vận dụng SGK trang 56.',
    '**Công cụ kiểm tra, đánh giá sản phẩm của học sinh:** **CCĐG 7 – Phiếu học tập số 5 và Bảng kiểm đánh giá năng lực vận dụng & kiểm chứng AI** *(Xem chi tiết tại Phụ lục 3)*.'
]:
    add(b, item, 'bullet')

add(b, 'd) Tổ chức thực hiện', 'label')
HD4_ROWS = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập (01 phút)**\n'
        '- Giáo viên chiếu đoạn mã lỗi trong tệp `03_Ma_Can_Sua.html` cùng phản hồi giả định của AI trong **Phiếu học tập số 5 (Phụ lục 3)**.\n'
        '- Giáo viên yêu cầu học sinh quan sát mã nguồn, chỉ ra lỗi cấu trúc bảng và nhận xét gợi ý giả định của AI; đồng thời giao nhiệm vụ Vận dụng SGK trang 56 về nhà.\n\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ (02 phút)**\n'
        '- Học sinh đối chiếu đoạn mã lỗi với cấu trúc chuẩn `<table>` $\\rightarrow$ `<tr>` $\\rightarrow$ `<td>` đã học ở Nhiệm vụ 3, phát hiện hai thẻ `<td>` bị đặt ngoài `<tr>`.\n'
        '- Học sinh sửa lại mã trong tệp `03_Ma_Can_Sua.html`, lưu thành `04_Ma_Da_Sua.html` và ghi câu trả lời ngắn vào Phiếu số 5.\n\n'
        '**Bước 3. Báo cáo sản phẩm (01 phút)**\n'
        '- Một học sinh trình bày trước lớp vì sao gợi ý *"chỉ thêm border"* là chưa đủ và đọc câu lệnh Prompt chuẩn để yêu cầu AI kiểm tra đúng quan hệ phân cấp thẻ HTML.\n\n'
        '**Bước 4. Đánh giá, nhận xét (01 phút)**\n'
        '- Giáo viên đánh giá theo **CCĐG 7 (Phụ lục 3)**, chốt nguyên tắc kiểm chứng AI vào cột ghi bài.\n'
        '- Giáo viên thu tệp sản phẩm `HoTen_12-4_Bai9.html`, hướng dẫn học sinh thực hiện bài tập **Vận dụng SGK trang 56** ở nhà và đọc trước **Bài 10. Tạo liên kết** (Tiết PPCT: 22, 23).',

        '**4. VẬN DỤNG & NGUYÊN TẮC KIỂM CHỨNG MÃ HTML VỚI AI**\n\n'
        '**a) Sửa lỗi cấu trúc bảng và kiểm chứng gợi ý AI:**\n'
        '- **Lỗi trong đoạn mã:** Hai phần tử `<td>Bóng đá</td>` và `<td>Thứ Ba</td>` đặt trực tiếp trong `<table>` mà thiếu cặp thẻ hàng `<tr>...</tr>` bao ngoài.\n'
        '- **Đánh giá gợi ý giả định:** Chưa đúng bản chất vì thêm `border` chỉ kẻ khung viền ngoài, không sửa được lỗi sai phân cấp `<table>` $\\rightarrow$ `<tr>` $\\rightarrow$ `<td>`.\n'
        '- **Mã sửa đúng:**\n'
        '`<tr><td>Bóng đá</td><td>Thứ Ba</td></tr>`\n'
        '- **Mẫu Prompt chuẩn khi hỏi AI:**\n'
        '*"Hãy kiểm tra cấu trúc đoạn mã bảng HTML sau, chỉ rõ lỗi phân cấp giữa các thẻ table, tr, th, td; gợi ý cách sửa tối thiểu giữ nguyên dữ liệu và giải thích cách kiểm tra lại trên trình duyệt."*\n\n'
        '**b) Hướng dẫn nhiệm vụ Vận dụng về nhà (SGK trang 56):**\n'
        '- Tự thiết kế bảng *Thời khóa biểu lớp 12/4* bằng HTML.\n'
        '- *(Khuyến khích)* Dùng vòng lặp `for` trong Python để ghép chuỗi sinh tự động các dòng `<tr><td>...</td></tr>` từ danh sách n × 4 và ghi ra tệp `bang_diem.html`.'
    ]
]
b.append(table(HD4_ROWS, [4400, 4600], header=True, cell_sz=24))

# ==============================================================================
# IV. PHỤ LỤC (HỌC LIỆU VÀ CÔNG CỤ KIỂM TRA, ĐÁNH GIÁ THEO TỪNG HOẠT ĐỘNG)
# (Fixes Comments #3, #33, #34, #35 - Replaces old IV & removes old V)
# ==============================================================================
br(b)
add(b, 'IV. PHỤ LỤC (HỆ THỐNG PHIẾU HỌC TẬP, CÔNG CỤ KIỂM TRA ĐÁNH GIÁ VÀ MÃ NGUỒN HỌC LIỆU)', 'heading')
add(b, 'Hệ thống Phụ lục dưới đây đính kèm đầy đủ các **Phiếu học tập**, **Đáp án**, **Công cụ kiểm tra, đánh giá (`CCĐG 1` – `CCĐG 7`)** cho từng Hoạt động / Nhiệm vụ trong Tiến trình dạy học và **Mã nguồn các tệp HTML** sử dụng trong bài dạy:')

# --- PHỤ LỤC 1 ---
add(b, 'PHỤ LỤC 1. PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ QUÁ TRÌNH (HOẠT ĐỘNG 1 & HOẠT ĐỘNG 2)', 'sub')

add(b, '1.1. Phụ lục 1.1 – Công cụ đánh giá Hoạt động 1 (CCĐG 1: Bảng kiểm đánh giá câu trả lời tình huống mở đầu)', 'label')
CCDG1_TABLE = [
    ['STT', 'Tiêu chí quan sát / Yêu cầu cần đạt ở Hoạt động 1 (Mở đầu)', 'Đạt', 'Chưa đạt', 'Ghi chú'],
    ['1', 'Nêu được ưu điểm của danh sách: tách rõ các mục tương tự nhau, tự động đánh số hoặc tạo kí hiệu đầu dòng.', '[  ]', '[  ]', 'Đánh giá `[KT1]`'],
    ['2', 'Chọn đúng phần tử Bảng (Table) cho tình huống Hội chợ ẩm thực lớp 12E (Hoạt động 2 SGK trang 53).', '[  ]', '[  ]', 'Đánh giá `[KT1]`'],
    ['3', 'Giải thích được lí do chọn bảng: mỗi món ăn có 4 thuộc tính cần căn thẳng hàng theo cột và dòng để dễ đối chiếu.', '[  ]', '[  ]', 'Đánh giá `[NLC2]`'],
    ['4', 'Thảo luận cặp đôi tích cực, trình bày câu trả lời rõ ràng, đúng thời gian quy định (05 phút).', '[  ]', '[  ]', 'Đánh giá `[PC1]`']
]
b.append(table(CCDG1_TABLE, [700, 5100, 900, 1000, 1300], header=True, cell_sz=23))

add(b, '1.2. Phụ lục 1.2 – Phiếu học tập số 1 & Công cụ đánh giá Nhiệm vụ 1 (CCĐG 2)', 'label')
add(b, '**[PHIẾU HỌC TẬP SỐ 1 – TÌM HIỂU DANH SÁCH CÓ THỨ TỰ VÀ KHÔNG CÓ THỨ TỰ (Tiết 1 / Tiết PPCT: 20)]**', keep=True)
add(b, 'Trường: THPT Thuận Hoá | Lớp: 12/4 | Nhóm/Máy: ......... | Học sinh: .........................................................', first=0)
add(b, '- **Câu 1:** Điền tên cặp thẻ HTML tương ứng: Tạo danh sách có thứ tự: `................`; Tạo danh sách không có thứ tự: `................`; Tạo từng mục trong danh sách: `................`.', first=283)
add(b, '- **Câu 2:** Thuộc tính `type` của thẻ `<ol>` nhận những giá trị nào? Nếu viết `<ol type="a" start="3">` thì hai mục `<li>` đầu tiên sẽ hiển thị kí hiệu thứ tự là gì?', first=283)
add(b, '- **Câu 3:** Đặc tính `list-style-type` trong thuộc tính `style` của thẻ `<ul>` nhận những giá trị nào? Hãy gõ và chạy thử `<ul style="list-style-type:square">` trên máy tính.', first=283)
add(b, '**[ĐÁP ÁN PHIẾU HỌC TẬP SỐ 1]:** Câu 1: `<ol>...</ol>`, `<ul>...</ul>`, `<li>...</li>`. Câu 2: 5 giá trị `"1"`, `"A"`, `"a"`, `"I"`, `"i"`; hai mục đầu hiển thị là **`c.`** và **`d.`**. Câu 3: 4 giá trị `disc`, `circle`, `square`, `none`; kết quả hiển thị dấu đầu dòng hình vuông đặc.', first=0)

CCDG2_TABLE = [
    ['STT', 'Tiêu chí đánh giá sản phẩm Nhiệm vụ 1 (CCĐG 2 – Thang điểm 10)', 'Điểm tối đa', 'Điểm đạt', 'Minh chứng mục tiêu'],
    ['1', 'Điền đúng 3 cặp thẻ `<ol>...</ol>`, `<ul>...</ul>` và `<li>...</li>` (Câu 1 Phiếu số 1).', '3,0 điểm', '........', '`[KT2]`, `[NLC1]`'],
    ['2', 'Nêu đủ 5 giá trị của `type` và xác định đúng 2 mục đầu của `<ol type="a" start="3">` là `c.` và `d.` (Câu 2).', '3,0 điểm', '........', '`[KT2]`, `[NLc]`'],
    ['3', 'Nêu đủ 4 giá trị của `list-style-type` và chạy thử thành công danh sách trên trình duyệt (Câu 3).', '4,0 điểm', '........', '`[NLa]`, `[NLS 3.1.NC1.a]`'],
    ['', '**TỔNG ĐIỂM ĐÁNH GIÁ NHIỆM VỤ 1**', '**10,0 điểm**', '........', '']
]
b.append(table(CCDG2_TABLE, [700, 4900, 1100, 1000, 1300], header=True, cell_sz=23))

add(b, '1.3. Phụ lục 1.3 – Phiếu học tập số 2 & Công cụ đánh giá Nhiệm vụ 2 (CCĐG 3)', 'label')
add(b, '**[PHIẾU HỌC TẬP SỐ 2 – TÌM HIỂU DANH SÁCH MÔ TẢ VÀ DANH SÁCH LỒNG NHAU (Tiết 1 / Tiết PPCT: 20)]**', keep=True)
add(b, 'Trường: THPT Thuận Hoá | Lớp: 12/4 | Nhóm/Máy: ......... | Học sinh: .........................................................', first=0)
add(b, '- **Câu 1:** Trong danh sách mô tả (Hình 9.3 SGK trang 53), thẻ nào dùng để khai báo toàn bộ danh sách mô tả, thẻ nào xác định tên mục và thẻ nào xác định lời mô tả? Có dùng thẻ `<li>` trong `<dl>` không?', first=283)
add(b, '- **Câu 2 (Câu hỏi SGK trang 53):** Viết đoạn mã tạo danh sách lồng nhau 2 cấp: mức 1 đánh số `1, 2` (`THỂ THAO`, `NGHỆ THUẬT`), bên trong mục `1. THỂ THAO` có danh sách mức 2 đánh chữ `a, b, c` (`Karatedo`, `Taekwondo`, `Vovinam`).', first=283)
add(b, '**[ĐÁP ÁN PHIẾU HỌC TẬP SỐ 2]:** Câu 1: `<dl>` khai báo danh sách mô tả, `<dt>` xác định tên mục, `<dd>` xác định lời mô tả; **không** dùng thẻ `<li>`. Câu 2: Đặt cặp thẻ `<ol type="a">...</ol>` nằm hoàn toàn bên trong cặp thẻ `<li>...</li>` của mục cha `THỂ THAO` (trước thẻ đóng `</li>`).', first=0)

CCDG3_TABLE = [
    ['STT', 'Tiêu chí đánh giá sản phẩm Nhiệm vụ 2 (CCĐG 3 – Bảng kiểm)', 'Đạt', 'Chưa đạt', 'Minh chứng mục tiêu'],
    ['1', 'Phân biệt đúng 3 thẻ `<dl>`, `<dt>`, `<dd>` và khẳng định không dùng `<li>` trong danh sách mô tả.', '[  ]', '[  ]', '`[KT3]`, `[NLC1]`'],
    ['2', 'Đặt đúng danh sách con `<ol type="a">` nằm bên trong cặp thẻ `<li>...</li>` của danh sách cha.', '[  ]', '[  ]', '`[KT3]`, `[NLc]`'],
    ['3', 'Viết được ít nhất 01 cặp `<dt>/<dd>` và 01 danh sách lồng 2 cấp lưu vào tệp `HoTen_12-4_Bai9.html`.', '[  ]', '[  ]', '`[NLa]`, `[PC1]`, `[PC2]`']
]
b.append(table(CCDG3_TABLE, [700, 5100, 900, 1000, 1300], header=True, cell_sz=23))

add(b, '1.4. Phụ lục 1.4 – Phiếu học tập số 3 & Công cụ đánh giá Nhiệm vụ 3 (CCĐG 4)', 'label')
add(b, '**[PHIẾU HỌC TẬP SỐ 3 – TÌM HIỂU CẤU TRÚC CƠ BẢN CỦA BẢNG HTML (Tiết 2 / Tiết PPCT: 21)]**', keep=True)
add(b, 'Trường: THPT Thuận Hoá | Lớp: 12/4 | Nhóm/Máy: ......... | Học sinh: .........................................................', first=0)
add(b, '- **Câu 1:** Nêu chức năng của các cặp thẻ: `<table>...</table>`, `<caption>...</caption>`, `<tr>...</tr>`, `<th>...</th>`, `<td>...</td>`.', first=283)
add(b, '- **Câu 2:** Cặp thẻ `<caption>` phải đặt ở vị trí nào trong bảng? Điểm khác biệt hiển thị mặc định giữa `<th>` và `<td>` là gì?', first=283)
add(b, '**[ĐÁP ÁN PHIẾU HỌC TẬP SỐ 3]:** Câu 1: `<table>` tạo bảng, `<caption>` tạo tiêu đề bảng, `<tr>` tạo hàng, `<th>` tạo ô tiêu đề, `<td>` tạo ô dữ liệu. Câu 2: `<caption>` đặt ngay sau thẻ mở `<table>` và trước hàng `<tr>` đầu tiên; `<th>` hiển thị chữ in đậm căn giữa, `<td>` hiển thị chữ thường căn trái.', first=0)

CCDG4_TABLE = [
    ['STT', 'Tiêu chí đánh giá sản phẩm Nhiệm vụ 3 (CCĐG 4 – Bảng kiểm)', 'Đạt', 'Chưa đạt', 'Minh chứng mục tiêu'],
    ['1', 'Nêu đúng chức năng của cả 5 cặp thẻ `<table>`, `<caption>`, `<tr>`, `<th>`, `<td>`.', '[  ]', '[  ]', '`[KT4]`, `[NLC1]`'],
    ['2', 'Đặt đúng thẻ `<caption>` ngay dưới thẻ mở `<table>` và đặt mọi thẻ `<th>`, `<td>` bên trong `<tr>`.', '[  ]', '[  ]', '`[KT4]`, `[NLc]`'],
    ['3', 'Thêm được hàng dữ liệu mới vào tệp `02_CLB_Mau_CoBan.html` và mở hiển thị đúng trên trình duyệt.', '[  ]', '[  ]', '`[NLa]`, `[NLS 3.1.NC1.a]`']
]
b.append(table(CCDG4_TABLE, [700, 5100, 900, 1000, 1300], header=True, cell_sz=23))

add(b, '1.5. Phụ lục 1.5 – Phiếu học tập số 4 & Công cụ đánh giá Nhiệm vụ 4 (CCĐG 5)', 'label')
add(b, '**[PHIẾU HỌC TẬP SỐ 4 – ĐỊNH DẠNG BẢNG VÀ GỘP Ô (Tiết 2 / Tiết PPCT: 21)]**', keep=True)
add(b, 'Trường: THPT Thuận Hoá | Lớp: 12/4 | Nhóm/Máy: ......... | Học sinh: .........................................................', first=0)
add(b, '- **Câu 1 (Câu hỏi SGK trang 55):** Bảng ở Hình 9.6 dùng `border="1"` có nhược điểm gì? Viết cú pháp thuộc tính `style` để tạo viền nét liền dày `1px` màu xanh (`green`) cho bảng và từng ô.', first=283)
add(b, '- **Câu 2:** Điền giá trị thuộc tính gộp ô cho hàng tiêu đề 1 của Hình 9.8 (SGK trang 56): `<th rowspan="...">Bộ môn</th> <th rowspan="...">GV phụ trách</th> <th colspan="...">Ngày</th>`. Ở hàng tiêu đề thứ 2 ngay bên dưới, ta chỉ viết mấy cặp thẻ `<th>`? Vì sao?', first=283)
add(b, '**[ĐÁP ÁN PHIẾU HỌC TẬP SỐ 4]:** Câu 1: `border="1"` chỉ tạo viền mặc định nét đôi thô và không linh hoạt; cần dùng `style="border: 1px solid green;"` cho cả `<table>`, `<th>`, `<td>`. Câu 2: Điền lần lượt `2`, `2`, `5`; hàng tiêu đề thứ 2 chỉ viết **5 cặp thẻ `<th>`** (*Thứ 2* đến *Thứ 6*) vì 2 cột đầu đã bị hai ô `rowspan="2"` ở hàng 1 chiếm chỗ.', first=0)

CCDG5_TABLE = [
    ['STT', 'Tiêu chí đánh giá sản phẩm Nhiệm vụ 4 (CCĐG 5 – Bảng kiểm)', 'Đạt', 'Chưa đạt', 'Minh chứng mục tiêu'],
    ['1', 'Viết đúng cú pháp `style="border: 1px solid ...; width: ...;"` cho cả `<table>` và `<th>`, `<td>`.', '[  ]', '[  ]', '`[KT5]`, `[NLc]`'],
    ['2', 'Khai báo đúng `rowspan="2"` cho 2 cột đầu và `colspan="5"` cho cột Ngày ở hàng tiêu đề 1.', '[  ]', '[  ]', '`[KT5]`, `[NLC3]`'],
    ['3', 'Viết đúng 5 thẻ `<th>` ở hàng tiêu đề 2 (không viết thừa ô tại vị trí đã gộp `rowspan="2"`).', '[  ]', '[  ]', '`[NLc]`, `[NLS 3.2.NC1.a]`']
]
b.append(table(CCDG5_TABLE, [700, 5100, 900, 1000, 1300], header=True, cell_sz=23))

# --- PHỤ LỤC 2: RUBRIC 10 ĐIỂM (Fixes Comment #34 completely!) ---
br(b)
add(b, 'PHỤ LỤC 2. CÔNG CỤ ĐÁNH GIÁ SẢN PHẨM THỰC HÀNH TẠI HOẠT ĐỘNG 3 – LUYỆN TẬP (CCĐG 6: RUBRIC 10 ĐIỂM)', 'sub')
add(b, '**PHIẾU ĐÁNH GIÁ SẢN PHẨM THỰC HÀNH (RUBRIC 10 ĐIỂM)**', 'title')

# Administrative metadata table required by Comment #34
RUBRIC_INFO = [
    ['Thông tin bài dạy & Lớp học', 'Thông tin Giáo viên & Học sinh / Nhóm thực hiện'],
    [
        '**Tên bài dạy:** Bài 9. Tạo danh sách, bảng\n'
        '**Môn học:** Tin học 12 – **Lớp:** 12/4\n'
        '**Tiết dạy:** Tiết 1, 2 của bài (**Tiết PPCT: 20, 21**)\n'
        '**Tên tệp sản phẩm:** `HoTen_12-4_Bai9.html`',
        '**Giáo viên giảng dạy (SVTT):** Hồ Quang Huy\n'
        '**Giáo viên hướng dẫn:** Thầy Mai Xuân Văn\n'
        '**Họ và tên học sinh:** ........................................................\n'
        '**Nhóm / Máy số:** ..................... **Điểm tổng:** ......... / 10 điểm'
    ]
]
b.append(table(RUBRIC_INFO, [4400, 4600], header=True, cell_sz=23))
add(b, '', first=0)

RUBRIC_FULL = [
    ['STT', 'Tiêu chí\nđánh giá', 'Câu lệnh / Cú pháp HTML\nbắt buộc kiểm tra', 'Mô tả chi tiết các mức độ đạt của sản phẩm\n(Mức Đạt tối đa / Mức Đạt một phần / Mức Chưa đạt)', 'Điểm\ntối đa', 'Điểm\nđạt'],
    ['1',
     'Cấu trúc trang HTML và lưu tệp\n(`[NLa]`)',
     '`<!DOCTYPE html>`\n`<html lang="vi">`\n`<head><meta charset="utf-8">`\n`<title>...</title></head>`\n`<body>...</body></html>`',
     '- **Mức Đạt tối đa (1,0 điểm):** Có đầy đủ khung cấu trúc HTML5 và `<meta charset="utf-8">`; lưu đúng tên tệp `HoTen_12-4_Bai9.html`, mở trên trình duyệt không lỗi phông tiếng Việt.\n'
     '- **Mức Đạt một phần (0,5 điểm):** Có khung HTML và mở được tệp nhưng đặt sai tên tệp hoặc thiếu khai báo `charset="utf-8"`.\n'
     '- **Mức Chưa đạt (0 điểm):** Không lưu được tệp `.html` hoặc tệp không mở được trên trình duyệt.',
     '1,0',
     '......'],
    ['2',
     'Danh sách không thứ tự và có thứ tự\n(`[KT2]`)',
     '`<h2>Danh sách Câu lạc bộ</h2>`\n`<ul><li>THỂ THAO`\n`<ol type="1"><li>...</li></ol>`\n`</li><li>NGHỆ THUẬT...</li></ul>`',
     '- **Mức Đạt tối đa (1,5 điểm):** Sử dụng đúng cặp thẻ `<ul>...</ul>` cho 2 nhóm lớn (*THỂ THAO*, *NGHỆ THUẬT*) và `<ol type="1">...<li>...</li></ol>` cho danh sách các bộ môn theo đúng Hình 9.7 SGK trang 55.\n'
     '- **Mức Đạt một phần (0,75 điểm):** Tạo được danh sách `<ul>` và `<ol>` nhưng thiếu mục hoặc nhầm thuộc tính `type`.\n'
     '- **Mức Chưa đạt (0 điểm):** Không sử dụng thẻ `<ul>`, `<ol>`, `<li>` (gõ số thứ tự thủ công bằng `<p>`).',
     '1,5',
     '......'],
    ['3',
     'Danh sách lồng nhau và Danh sách mô tả\n(`[KT3]`)',
     '`<li>Võ <ol type="a">`\n`  <li>Karatedo</li>...`\n`</ol></li>`\n`<dl><dt>...</dt><dd>...</dd></dl>`',
     '- **Mức Đạt tối đa (1,5 điểm):** Đặt đúng danh sách con `<ol type="a">` (gồm *a. Karatedo, b. Taekwondo, c. Vovinam* và *a. Thanh nhạc, b. Piano, c. Violin*) nằm **bên trong** thẻ `<li>...</li>` của mục cha (*1,0 điểm*); có ít nhất 01 cặp `<dl><dt>...</dt><dd>...</dd></dl>` hợp lệ (*0,5 điểm*).\n'
     '- **Mức Đạt một phần (0,75 điểm):** Có danh sách con hoặc `<dl>/<dt>/<dd>` nhưng đặt danh sách con ngoài thẻ đóng `</li>` của mục cha.\n'
     '- **Mức Chưa đạt (0 điểm):** Không có danh sách lồng và không có danh sách mô tả.',
     '1,5',
     '......'],
    ['4',
     'Khởi tạo bảng và Tiêu đề bảng\n(`[KT4]`)',
     '`<table>`\n`<caption>Lịch CLB</caption>`\n`<tr><th>...</th></tr>`\n`<tr><td>...</td></tr>`\n`</table>`',
     '- **Mức Đạt tối đa (2,0 điểm):** Có thẻ `<caption>` đặt ngay sau `<table>` (*0,5 điểm*); bảng có đủ 7 hàng (`<tr>`: 2 hàng tiêu đề `<th>` + 5 hàng dữ liệu `<td>` cho *Bóng đá, Bóng chuyền, Bóng rổ, Võ thuật, Bơi*) và đủ 7 cột quy đổi theo Hình 9.8 SGK trang 56 (*1,5 điểm*).\n'
     '- **Mức Đạt một phần (1,0 điểm):** Tạo được bảng có `<tr>`, `<th>`, `<td>` nhưng thiếu `<caption>` hoặc mới nhập được 2–3 hàng dữ liệu.\n'
     '- **Mức Chưa đạt (0 điểm):** Chưa tạo được bảng hoặc đặt `<td>` nằm ngoài `<tr>`.',
     '2,0',
     '......'],
    ['5',
     'Gộp hàng và Gộp cột tiêu đề bảng\n(`[KT5]`)',
     '`<th rowspan="2">Bộ môn</th>`\n`<th rowspan="2">GV...</th>`\n`<th colspan="5">Ngày</th>`\n`<tr><th>Thứ 2</th>...</tr>`',
     '- **Mức Đạt tối đa (2,0 điểm):** Hàng tiêu đề 1 dùng đúng `rowspan="2"` cho 2 ô đầu (*Bộ môn*, *GV phụ trách*) và `colspan="5"` cho ô *Ngày* (*1,0 điểm*); hàng tiêu đề 2 viết đúng 5 cặp thẻ `<th>` (*Thứ 2* → *Thứ 6*), không bị thừa ô lệch cột (*1,0 điểm*).\n'
     '- **Mức Đạt một phần (1,0 điểm):** Có sử dụng `rowspan` và `colspan` nhưng viết thừa 2 ô trống ở hàng tiêu đề 2 làm bảng bị lệch thành 9 cột.\n'
     '- **Mức Chưa đạt (0 điểm):** Không sử dụng `rowspan` và `colspan`.',
     '2,0',
     '......'],
    ['6',
     'Định dạng viền và màu sắc bằng thuộc tính style\n(`[KT5]`, Luyện tập tr.56)',
     '`style="border: 1px solid green;"`\n`style="border: 1px solid black;"`\n`style="color: red;"`',
     '- **Mức Đạt tối đa (1,0 điểm):** Sử dụng thuộc tính `style` để tạo khung viền cho cả `<table>` và từng ô `<th>`, `<td>` (*0,5 điểm*); áp dụng đúng viền màu xanh (`green`) cho 2 hàng tiêu đề đầu và đổi màu chữ theo yêu cầu Luyện tập SGK trang 56 (*0,5 điểm*).\n'
     '- **Mức Đạt một phần (0,5 điểm):** Chỉ đặt `border` ở thẻ `<table>` mà chưa đặt `border` cho các ô `<th>`, `<td>`.\n'
     '- **Mức Chưa đạt (0 điểm):** Bảng không có đường viền.',
     '1,0',
     '......'],
    ['7',
     'Kiểm chứng lỗi cấu trúc HTML với AI\n(`[KT6]`, `[NLAI 12.C2.2]`)',
     'Sửa lỗi `<td>` ngoài `<tr>`:\n`<tr><td>Bóng đá</td>`\n`<td>Thứ Ba</td></tr>`',
     '- **Mức Đạt tối đa (1,0 điểm):** Chỉ ra đúng lỗi hai thẻ `<td>` nằm ngoài `<tr>` trong tệp `03_Ma_Can_Sua.html`, giải thích được vì sao gợi ý chỉ thêm `border` là chưa đủ và sửa mã chạy đúng.\n'
     '- **Mức Đạt một phần (0,5 điểm):** Sửa được mã thêm `<tr>...</tr>` nhưng chưa giải thích được lí do bác bỏ gợi ý giả định.\n'
     '- **Mức Chưa đạt (0 điểm):** Không phát hiện và không sửa được lỗi cấu trúc.',
     '1,0',
     '......'],
    ['',
     '**TỔNG CỘNG**',
     '',
     '**Tổng điểm đánh giá sản phẩm thực hành (Thang điểm 10)**',
     '**10,0**',
     '......']
]
b.append(table(RUBRIC_FULL, [650, 1500, 2150, 3300, 700, 700], header=True, cell_sz=22))

# --- PHỤ LỤC 3 ---
add(b, 'PHỤ LỤC 3. PHIẾU HỌC TẬP SỐ 5 VÀ CÔNG CỤ ĐÁNH GIÁ HOẠT ĐỘNG 4 – VẬN DỤNG (CCĐG 7)', 'sub')
add(b, '**[PHIẾU HỌC TẬP SỐ 5 – KIỂM CHỨNG GỢI Ý SỬA LỖI CẤU TRÚC BẢNG CỦA AI (Tiết 2 / Tiết PPCT: 21)]**', keep=True)
add(b, 'Trường: THPT Thuận Hoá | Lớp: 12/4 | Nhóm/Máy: ......... | Học sinh: .........................................................', first=0)
add(b, 'Cho đoạn mã HTML bị lỗi trong tệp `03_Ma_Can_Sua.html`:')
for ln in ['<table>', '  <caption>Lịch CLB</caption>', '  <tr><th>CLB</th><th>Lịch</th></tr>', '  <td>Bóng đá</td>', '  <td>Thứ Ba</td>', '</table>']:
    add(b, '`' + ln + '`', 'code')
add(b, '- **Gợi ý giả định của AI:** *"Bạn chỉ cần thêm thuộc tính `style=\"border: 1px solid black;\"` vào thẻ `<table>`."*')
add(b, '- **Câu hỏi 1:** Gợi ý trên đã sửa được lỗi cấu trúc của bảng chưa? Hãy chỉ rõ lỗi sai trong đoạn mã trên và viết lại dòng mã sửa tối thiểu.', first=283)
add(b, '- **Câu hỏi 2:** Hãy viết 01 câu lệnh Prompt chuẩn để yêu cầu công cụ AI kiểm tra đúng lỗi quan hệ phân cấp giữa các thẻ bảng HTML.', first=283)
add(b, '**[ĐÁP ÁN PHIẾU HỌC TẬP SỐ 5]:** Câu 1: Chưa đủ; lỗi cấu trúc là hai thẻ `<td>Bóng đá</td>` và `<td>Thứ Ba</td>` đang nằm ngoài cặp thẻ hàng `<tr>...</tr>`; cách sửa tối thiểu là bọc hai ô đó vào trong một hàng: `<tr><td>Bóng đá</td><td>Thứ Ba</td></tr>`. Câu 2: Prompt chuẩn: *"Hãy kiểm tra cấu trúc đoạn mã HTML sau, chỉ ra lỗi phân cấp giữa các thẻ table, tr, th, td; gợi ý cách sửa tối thiểu giữ nguyên dữ liệu và nêu cách kiểm tra lại."*', first=0)

CCDG7_TABLE = [
    ['STT', 'Tiêu chí đánh giá Hoạt động 4 – Vận dụng (CCĐG 7 – Bảng kiểm)', 'Đạt', 'Chưa đạt', 'Minh chứng mục tiêu'],
    ['1', 'Phát hiện chính xác lỗi hai phần tử `<td>` đặt ngoài cặp thẻ hàng `<tr>...</tr>`.', '[  ]', '[  ]', '`[KT6]`, `[NLc]`'],
    ['2', 'Bác bỏ được gợi ý giả định chỉ thêm `border` và sửa đúng mã trong tệp `04_Ma_Da_Sua.html`.', '[  ]', '[  ]', '`[NLd]`, `[PC3]`'],
    ['3', 'Viết được câu lệnh Prompt rõ yêu cầu kiểm tra quan hệ phân cấp thẻ và nắm được yêu cầu bài Vận dụng tr.56.', '[  ]', '[  ]', '`[NLAI 12.C2.2]`']
]
b.append(table(CCDG7_TABLE, [700, 5100, 900, 1000, 1300], header=True, cell_sz=23))

# --- PHỤ LỤC 4: MÃ NGUỒN CÁC TỆP HỌC LIỆU HTML (Fixes Comment #3) ---
add(b, 'PHỤ LỤC 4. MÃ NGUỒN CÁC TỆP HỌC LIỆU HTML ĐÍNH KÈM TRONG HỒ SƠ BÀI DẠY', 'sub')
add(b, 'Toàn bộ các tệp học liệu HTML dưới đây được đóng gói sẵn trong thư mục `Hoc_lieu_HTML/` đi kèm hồ sơ bài dạy (và đồng bộ trên Google Drive học liệu):')

add(b, '4.1. Tệp `01_Khung_CLB.html` (Tệp khung khởi tạo cấp cho học sinh thực hành xuyên suốt 2 tiết):', 'label')
for ln in [
    '<!DOCTYPE html>',
    '<html lang="vi">',
    '<head>',
    '  <meta charset="utf-8">',
    '  <title>Câu lạc bộ trường THPT Thuận Hoá - Lớp 12/4</title>',
    '</head>',
    '<body>',
    '  <h1>TRANG THÔNG TIN CÂU LẠC BỘ HỌC SINH</h1>',
    '  <!-- PHẦN 1 (TIẾT 1 - TIẾT PPCT 20): TẠO DANH SÁCH CÂU LẠC BỘ THEO HÌNH 9.7 -->',
    '  <h2>Danh sách Câu lạc bộ</h2>',
    '  <!-- Học sinh gõ mã <ul>, <ol>, <li> và <dl>, <dt>, <dd> tại đây -->',
    '',
    '  <!-- PHẦN 2 (TIẾT 2 - TIẾT PPCT 21): THIẾT LẬP BẢNG LỊCH HOẠT ĐỘNG THEO HÌNH 9.8 -->',
    '  <!-- Học sinh gõ mã <table>, <caption>, <tr>, <th>, <td> tại đây -->',
    '</body>',
    '</html>'
]:
    add(b, '`' + ln + '`', 'code')

add(b, '4.2. Tệp `05_CLB_HoanChinh_SGK.html` (Sản phẩm mẫu hoàn chỉnh tích hợp Hình 9.7, Hình 9.8 và Luyện tập SGK trang 55–56):', 'label')
for ln in [
    '<!DOCTYPE html>',
    '<html lang="vi">',
    '<head>',
    '  <meta charset="utf-8">',
    '  <title>Danh sách và Lịch hoạt động Câu lạc bộ - Bài 9 Tin học 12</title>',
    '</head>',
    '<body>',
    '  <h2>Danh sách Câu lạc bộ</h2>',
    '  <dl>',
    '    <dt><b>Câu lạc bộ năng khiếu THPT Thuận Hoá</b></dt>',
    '    <dd>Nơi rèn luyện thể chất, phát triển năng khiếu nghệ thuật và kỹ năng hợp tác cho học sinh.</dd>',
    '  </dl>',
    '  <ul>',
    '    <li>THỂ THAO',
    '      <ol type="1">',
    '        <li>Bóng đá</li>',
    '        <li>Bóng chuyền</li>',
    '        <li>Bóng rổ</li>',
    '        <li>Võ',
    '          <ol type="a">',
    '            <li>Karatedo</li>',
    '            <li>Taekwondo</li>',
    '            <li>Vovinam</li>',
    '          </ol>',
    '        </li>',
    '        <li>Bơi</li>',
    '      </ol>',
    '    </li>',
    '    <li>NGHỆ THUẬT',
    '      <ol type="1">',
    '        <li>Mĩ thuật</li>',
    '        <li>Nhiếp ảnh</li>',
    '        <li>Âm nhạc',
    '          <ol type="a">',
    '            <li>Thanh nhạc</li>',
    '            <li>Piano</li>',
    '            <li>Violin</li>',
    '          </ol>',
    '        </li>',
    '        <li>Khiêu vũ</li>',
    '      </ol>',
    '    </li>',
    '  </ul>',
    '  <table style="border: 1px solid black; border-collapse: collapse; width: 90%;">',
    '    <caption><b>Lịch hoạt động CLB Thể thao</b></caption>',
    '    <tr>',
    '      <th rowspan="2" style="border: 1px solid green; padding: 4px;">Bộ môn</th>',
    '      <th rowspan="2" style="border: 1px solid green; padding: 4px;">GV phụ trách</th>',
    '      <th colspan="5" style="border: 1px solid green; padding: 4px;">Ngày</th>',
    '    </tr>',
    '    <tr>',
    '      <th style="border: 1px solid green; padding: 4px;">Thứ 2</th>',
    '      <th style="border: 1px solid green; padding: 4px;">Thứ 3</th>',
    '      <th style="border: 1px solid green; padding: 4px;">Thứ 4</th>',
    '      <th style="border: 1px solid green; padding: 4px;">Thứ 5</th>',
    '      <th style="border: 1px solid green; padding: 4px;">Thứ 6</th>',
    '    </tr>',
    '    <tr>',
    '      <td style="border: 1px solid black; color: red; padding: 4px;">Bóng đá</td>',
    '      <td style="border: 1px solid black; padding: 4px;">Thầy Khải</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '    </tr>',
    '    <tr>',
    '      <td style="border: 1px solid black; color: darkorange; padding: 4px;">Bóng chuyền</td>',
    '      <td style="border: 1px solid black; padding: 4px;">Cô Lan</td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '    </tr>',
    '    <tr>',
    '      <td style="border: 1px solid black; color: blue; padding: 4px;">Bóng rổ</td>',
    '      <td style="border: 1px solid black; padding: 4px;">Thầy Vũ</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '    </tr>',
    '    <tr>',
    '      <td style="border: 1px solid black; padding: 4px;">Võ thuật</td>',
    '      <td style="border: 1px solid black; padding: 4px;">Thầy Hoàng - Karatedo<br>Thầy Đức - Taekwondo<br>Thầy Thắng - Vovinam</td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '    </tr>',
    '    <tr>',
    '      <td style="border: 1px solid black; padding: 4px;">Bơi</td>',
    '      <td style="border: 1px solid black; padding: 4px;">Cô Bích</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '      <td style="border: 1px solid black; text-align: center;"></td>',
    '      <td style="border: 1px solid black; text-align: center;">X</td>',
    '    </tr>',
    '  </table>',
    '</body>',
    '</html>'
]:
    add(b, '`' + ln + '`', 'code')

# ==============================================================================
# V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT (Fixes Comments #36, #37, #38)
# ==============================================================================
add(b, 'V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT', 'heading', keep=True)
add(b, '**Nhận xét của Giáo viên hướng dẫn:**', first=0, keep=True)
dots = '.' * 115
for _ in range(4):
    add(b, dots, first=0, align='left', keep=True)
add(b, dots, first=0, align='left', keep=True)
b.append(signature_table())

b.append(ending(PKG))
write_doc(PKG, b, MAIN, include_figures=True)

# Also copy to 06_thanh_pham_docx for easy access
THANH_PHAM = TASK.parent / '06_thanh_pham_docx' / 'Đề cương dự giờ Tin12_Bai9_Lop12-4_HoQuangHuy.docx'
shutil.copy2(MAIN, THANH_PHAM)

# ==============================================================================
# BUILD COMPANION HANDOUT & RUBRIC DOCUMENT (02_Phieu_hoc_tap_va_dap_an_Tin12_Bai9.docx)
# ==============================================================================
h = node('body')
add(h, 'PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ BÀI 9: TẠO DANH SÁCH, BẢNG', 'title')
add(h, '(Thời lượng: 02 tiết – Tiết thứ 1, 2 của bài; Tiết PPCT: 20, 21 – Lớp 12/4)', align='center', first=0, italic=True, keep=True)
add(h, 'Trường: THPT Thuận Hoá | Giáo viên giảng dạy (SVTT): Hồ Quang Huy | GVHD: Thầy Mai Xuân Văn', align='center', first=0, keep=True)
add(h, 'Họ và tên học sinh: ....................................................... Lớp: 12/4 | Nhóm / Máy số: ...............', first=0)

add(h, 'PHẦN I. PHIẾU HỌC TẬP TIẾT 1 (TIẾT PPCT: 20 – TẠO DANH SÁCH TRONG HTML)', 'heading')
add(h, '1. Phiếu học tập số 1 – Hoạt động 2, Nhiệm vụ 1: Danh sách có thứ tự và không có thứ tự (20 phút)', 'sub')
add(h, '- **Câu 1:** Điền tên cặp thẻ HTML tương ứng: Danh sách có thứ tự: `................`; Danh sách không có thứ tự: `................`; Từng mục trong danh sách: `................`.', first=283)
add(h, '- **Câu 2:** Thuộc tính `type` của thẻ `<ol>` nhận 5 giá trị nào? Nếu viết `<ol type="a" start="3">` thì hai mục `<li>` đầu tiên hiển thị kí hiệu gì? ....................................................................................', first=283)
add(h, '- **Câu 3:** Đặc tính `list-style-type` trong thuộc tính `style` của thẻ `<ul>` nhận 4 giá trị nào? Kết quả khi chạy `<ul style="list-style-type:square">` là gì? .......................................................', first=283)

add(h, '2. Phiếu học tập số 2 – Hoạt động 2, Nhiệm vụ 2: Danh sách mô tả và danh sách lồng nhau (20 phút)', 'sub')
add(h, '- **Câu 1:** Điền thẻ tương ứng trong danh sách mô tả (Hình 9.3 SGK trang 53): Khai báo danh sách mô tả: `........`; Tên mục: `........`; Lời mô tả: `........`. Trong `<dl>` có dùng thẻ `<li>` không? `........`', first=283)
add(h, '- **Câu 2 (Câu hỏi SGK trang 53):** Khi tạo danh sách lồng nhau (Hình 9.7 SGK trang 55), cặp thẻ `<ol type="a">...</ol>` của danh sách con (*Karatedo, Taekwondo, Vovinam*) phải đặt ở vị trí nào so với cặp thẻ `<li>...</li>` của mục cha *4. Võ*?', first=283)
add(h, '......................................................................................................................................................................', first=0)
add(h, '- **Thực hành trên máy cuối Tiết 1:** Mở tệp khung `01_Khung_CLB.html`, lưu thành `HoTen_12-4_Bai9.html` (UTF-8), gõ 01 cặp `<dl>/<dt>/<dd>` và nhánh danh sách lồng nhau *THỂ THAO* theo Hình 9.7 SGK trang 55.', first=283)

br(h)
add(h, 'PHẦN II. PHIẾU HỌC TẬP TIẾT 2 (TIẾT PPCT: 21 – THIẾT LẬP BẢNG VÀ LUYỆN TẬP, VẬN DỤNG)', 'heading')
add(h, 'Họ và tên học sinh: ....................................................... Lớp: 12/4 | Nhóm / Máy số: ...............', first=0)
add(h, '3. Phiếu học tập số 3 – Hoạt động 2, Nhiệm vụ 3: Cấu trúc cơ bản của bảng trong HTML (10 phút)', 'sub')
add(h, '- **Câu 1:** Nêu chức năng của các cặp thẻ: `<table>`: .......................; `<caption>`: .......................; `<tr>`: .......................; `<th>`: .......................; `<td>`: .......................', first=283)
add(h, '- **Câu 2:** Thẻ `<caption>` đặt ở đâu trong `<table>`? Khác biệt hiển thị mặc định giữa `<th>` và `<td>` là gì?', first=283)
add(h, '......................................................................................................................................................................', first=0)

add(h, '4. Phiếu học tập số 4 – Hoạt động 2, Nhiệm vụ 4: Định dạng bảng và gộp ô (12 phút)', 'sub')
add(h, '- **Câu 1 (Câu hỏi SGK trang 55):** Chỉ đặt `border` ở thẻ `<table>` đã tạo đủ viền cho từng ô chưa? Viết cú pháp `style` tạo viền nét liền dày `1px` màu xanh (`green`): ...............................................................', first=283)
add(h, '- **Câu 2:** Điền giá trị gộp ô cho hàng tiêu đề 1 của Hình 9.8 (SGK trang 56): `<th rowspan="...">Bộ môn</th> <th rowspan="...">GV phụ trách</th> <th colspan="...">Ngày</th>`. Hàng tiêu đề 2 ngay bên dưới chỉ viết mấy cặp thẻ `<th>`? Vì sao?', first=283)
add(h, '......................................................................................................................................................................', first=0)

add(h, '5. Yêu cầu Thực hành tại Hoạt động 3 – Luyện tập (18 phút, SGK trang 55–56)', 'sub')
add(h, 'Tiếp tục hoàn thiện tệp `HoTen_12-4_Bai9.html` gồm: (1) Danh sách lồng nhau đầy đủ 2 nhóm *THỂ THAO*, *NGHỆ THUẬT* theo Hình 9.7 (trang 55); (2) Bảng *Lịch hoạt động CLB Thể thao* 7 hàng × 7 cột quy đổi theo Hình 9.8 (trang 56); (3) Áp dụng bài Luyện tập trang 56: dùng `style` tô viền màu xanh (`green`) cho 2 hàng tiêu đề và đổi màu chữ các bộ môn.', first=283)

add(h, '6. Phiếu học tập số 5 – Hoạt động 4: Vận dụng – Kiểm chứng gợi ý AI (05 phút)', 'sub')
add(h, 'Cho đoạn mã lỗi trong tệp `03_Ma_Can_Sua.html`: `<table> <caption>Lịch CLB</caption> <tr><th>CLB</th><th>Lịch</th></tr> <td>Bóng đá</td> <td>Thứ Ba</td> </table>`. Gợi ý giả định của AI: *"Chỉ cần thêm `border` vào `table`."*', first=283)
add(h, '- **Câu hỏi:** Gợi ý trên đã sửa được lỗi cấu trúc chưa? Chỉ rõ lỗi và viết lại dòng mã sửa tối thiểu:', first=283)
add(h, '......................................................................................................................................................................', first=0)

br(h)
add(h, 'PHẦN III. PHIẾU ĐÁNH GIÁ SẢN PHẨM THỰC HÀNH (CCĐG 6 – RUBRIC 10 ĐIỂM)', 'heading')
h.append(table(RUBRIC_INFO, [4400, 4600], header=True, cell_sz=23))
add(h, '', first=0)
h.append(table(RUBRIC_FULL, [650, 1500, 2150, 3300, 700, 700], header=True, cell_sz=22))

br(h)
add(h, 'PHẦN IV. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM DÀNH CHO GIÁO VIÊN', 'heading')
add(h, '- **Đáp án Phiếu số 1 (CCĐG 2):** Câu 1: `<ol>...</ol>`, `<ul>...</ul>`, `<li>...</li>`. Câu 2: 5 giá trị `"1"`, `"A"`, `"a"`, `"I"`, `"i"`; hai mục đầu của `<ol type="a" start="3">` là **`c.`** và **`d.`**. Câu 3: 4 giá trị `disc`, `circle`, `square`, `none`; tạo kí tự đầu dòng hình vuông đặc.', first=283)
add(h, '- **Đáp án Phiếu số 2 (CCĐG 3):** Câu 1: `<dl>`, `<dt>`, `<dd>`; không dùng `<li>`. Câu 2: Danh sách con `<ol type="a">...</ol>` phải đặt nằm hoàn toàn **bên trong** cặp thẻ `<li>...</li>` của mục cha *4. Võ* (trước thẻ đóng `</li>`).', first=283)
add(h, '- **Đáp án Phiếu số 3 (CCĐG 4):** Câu 1: `<table>` tạo bảng, `<caption>` tiêu đề bảng, `<tr>` hàng, `<th>` ô tiêu đề, `<td>` ô dữ liệu. Câu 2: `<caption>` đặt ngay sau thẻ mở `<table>`; `<th>` in đậm căn giữa, `<td>` chữ thường căn trái.', first=283)
add(h, '- **Đáp án Phiếu số 4 (CCĐG 5):** Câu 1: Chưa đủ, cần đặt `style="border: 1px solid green;"` cho cả `<table>`, `<th>`, `<td>`. Câu 2: Điền `2`, `2`, `5`; hàng tiêu đề 2 chỉ viết **5 cặp thẻ `<th>`** (*Thứ 2* → *Thứ 6*) vì 2 cột đầu đã bị hai ô `rowspan="2"` chiếm chỗ.', first=283)
add(h, '- **Đáp án Phiếu số 5 (CCĐG 7):** Gợi ý giả định chưa đủ vì hai thẻ `<td>` đang đặt ngoài `<tr>`; cần bọc trong hàng `<tr><td>Bóng đá</td><td>Thứ Ba</td></tr>`. Prompt chuẩn yêu cầu AI kiểm tra quan hệ phân cấp `table - tr - th/td` và giữ nguyên dữ liệu.', first=283)

h.append(ending(HPKG))
write_doc(HPKG, h, HAND, include_figures=False)

# ==============================================================================
# UPDATE SLIDE DECK (00_Trinh_chieu.html) & TEXT GUIDES + ZIP ARCHIVE
# ==============================================================================
A_DIR = OUT / 'Hoc_lieu_HTML'
deckpath = A_DIR / '00_Trinh_chieu.html'
olddeck = deckpath.read_text(encoding='utf-8')
style_m = re.search(r'<style>(.*?)</style>', olddeck, re.S).group(1)
script_m = re.search(r'<script>(.*?)</script>', olddeck, re.S).group(1)
slides = [
    ('Bài 9 · 02 tiết (Tiết PPCT: 20, 21) · 90 phút', 'Tạo danh sách, bảng', '<p class="lead">Tiết 1 (Tiết PPCT: 20): Hoạt động 1 (Mở đầu) & Hoạt động 2 (Nhiệm vụ 1, Nhiệm vụ 2)<br>Tiết 2 (Tiết PPCT: 21): Hoạt động 2 (Nhiệm vụ 3, Nhiệm vụ 4), Hoạt động 3 (Luyện tập) & Hoạt động 4 (Vận dụng)</p><p class="note">Tin học 12/4 · SGK Kết nối tri thức với cuộc sống trang 52–56</p>'),
    ('Tiết 1 (PPCT: 20) · Hoạt động 1: Mở đầu · 05 phút', 'Lựa chọn cách trình bày dữ liệu trên web', '<div class="cols"><div><h3>Tình huống 1: Danh sách CLB</h3><p>Liệt kê các mục cùng loại (Thể thao, Nghệ thuật).</p><details><summary>Phương án chọn</summary><p>Dùng Danh sách (List).</p></details></div><div><h3>Tình huống 2: Hội chợ ẩm thực 12E</h3><p>Đối chiếu 4 thuộc tính: Món ăn · Đơn giá · Số lượng · Tổng tiền.</p><details><summary>Phương án chọn</summary><p>Dùng Bảng (Table) theo hàng và cột.</p></details></div></div><p class="note">SGK trang 52–53 · Công cụ đánh giá: CCĐG 1.</p>'),
    ('Tiết 1 (PPCT: 20) · HĐ 2 – Nhiệm vụ 1 · 20 phút', 'Danh sách có thứ tự và không có thứ tự', '<div class="cols"><pre>&lt;ol type="a" start="3"&gt;\n  &lt;li&gt;Bóng đá&lt;/li&gt;\n  &lt;li&gt;Bóng chuyền&lt;/li&gt;\n&lt;/ol&gt;</pre><div><p><code>&lt;ol&gt;</code>: Danh sách có thứ tự (<code>type="1|A|a|I|i"</code>, <code>start</code>)<br><code>&lt;ul&gt;</code>: Danh sách không có thứ tự (<code>list-style-type: disc|circle|square|none</code>)<br><code>&lt;li&gt;</code>: Từng mục trong danh sách</p></div></div><p class="note">Hình 9.1, Hình 9.2 SGK trang 52 · Phiếu học tập số 1 (CCĐG 2).</p>'),
    ('Tiết 1 (PPCT: 20) · HĐ 2 – Nhiệm vụ 2 · 20 phút', 'Danh sách mô tả và danh sách lồng nhau', '<div class="cols"><pre>&lt;dl&gt;\n &lt;dt&gt;Thể thao&lt;/dt&gt;\n &lt;dd&gt;Rèn luyện sức khoẻ.&lt;/dd&gt;\n&lt;/dl&gt;</pre><pre>&lt;ul&gt;\n &lt;li&gt;THỂ THAO\n  &lt;ol type="1"&gt;\n   &lt;li&gt;Võ\n    &lt;ol type="a"&gt;&lt;li&gt;Karatedo&lt;/li&gt;&lt;/ol&gt;\n   &lt;/li&gt;\n  &lt;/ol&gt;\n &lt;/li&gt;\n&lt;/ul&gt;</pre></div><p class="note">Hình 9.3 SGK trang 53 · Danh sách con đặt bên trong thẻ &lt;li&gt;...&lt;/li&gt; của mục cha · Phiếu số 2 (CCĐG 3).</p>'),
    ('Tiết 1 (PPCT: 20) · HĐ 2 – Nhiệm vụ 2', 'Khởi tạo khung danh sách CLB (Hình 9.7)', '<div class="cols"><div><ul><li>Mở <code>01_Khung_CLB.html</code>, lưu <code>HoTen_12-4_Bai9.html</code>.</li><li>Tạo 01 cặp mô tả <code>&lt;dl&gt;/&lt;dt&gt;/&lt;dd&gt;</code>.</li><li>Tạo nhánh <strong>THỂ THAO</strong> có danh sách con <strong>Võ (a, b, c)</strong>.</li></ul></div><div><p>Hình 9.7 · SGK trang 55</p><p>Kiểm tra đóng thẻ <code>&lt;/ol&gt;</code> trước <code>&lt;/li&gt;</code> của mục cha.</p></div></div>'),
    ('Tiết 1 (PPCT: 20) · Chốt Tiết 1', 'Tổng kết kiến thức tạo danh sách HTML', '<ol><li>Phân biệt khi nào dùng <code>&lt;ol&gt;</code>, <code>&lt;ul&gt;</code>, <code>&lt;dl&gt;</code>?</li><li>Danh sách con phải đặt ở vị trí nào trong danh sách cha?</li></ol><p class="callout">Giữ tệp <code>HoTen_12-4_Bai9.html</code> để tiếp tục học phần Bảng và Luyện tập ở Tiết 2 (Tiết PPCT: 21).</p>'),
    ('Tiết 2 (PPCT: 21) · HĐ 2 – Nhiệm vụ 3 · 10 phút', 'Cấu trúc cơ bản của bảng trong HTML', '<div class="cols"><pre>&lt;table&gt;\n &lt;caption&gt;Lịch CLB&lt;/caption&gt;\n &lt;tr&gt;&lt;th&gt;CLB&lt;/th&gt;&lt;th&gt;Lịch&lt;/th&gt;&lt;/tr&gt;\n &lt;tr&gt;&lt;td&gt;Bóng đá&lt;/td&gt;&lt;td&gt;Thứ Ba&lt;/td&gt;&lt;/tr&gt;\n&lt;/table&gt;</pre><div><p><code>&lt;table&gt;</code> → Bảng<br><code>&lt;caption&gt;</code> → Tiêu đề bảng<br><code>&lt;tr&gt;</code> → Hàng<br><code>&lt;th&gt;</code> → Ô tiêu đề · <code>&lt;td&gt;</code> → Ô dữ liệu</p><a href="02_CLB_Mau_CoBan.html" target="_blank">Mở bảng cơ bản</a></div></div><p class="note">Hình 9.4 SGK trang 54 · Phiếu học tập số 3 (CCĐG 4).</p>'),
    ('Tiết 2 (PPCT: 21) · HĐ 2 – Nhiệm vụ 4 · 12 phút', 'Định dạng bảng và gộp ô (rowspan, colspan)', '<pre>&lt;table style="border: 1px solid black; border-collapse: collapse; width: 90%;"&gt;\n&lt;th style="border: 1px solid green;" rowspan="2"&gt;Bộ môn&lt;/th&gt;</pre><div class="cols"><div><h3>rowspan="k"</h3><p>Gộp k ô theo chiều dọc (hàng).</p></div><div><h3>colspan="m"</h3><p>Gộp m ô theo chiều ngang (cột).</p></div></div><p class="note">Hình 9.5, Hình 9.6 SGK trang 54–55 · Phiếu học tập số 4 (CCĐG 5).</p>'),
    ('Tiết 2 (PPCT: 21) · HĐ 2 – Nhiệm vụ 4', 'Hai hàng tiêu đề bảng Lịch CLB (Hình 9.8)', '<pre>&lt;tr&gt;\n &lt;th rowspan="2"&gt;Bộ môn&lt;/th&gt;\n &lt;th rowspan="2"&gt;GV phụ trách&lt;/th&gt;\n &lt;th colspan="5"&gt;Ngày&lt;/th&gt;\n&lt;/tr&gt;\n&lt;tr&gt;&lt;th&gt;Thứ 2&lt;/th&gt;&lt;th&gt;Thứ 3&lt;/th&gt;&lt;th&gt;Thứ 4&lt;/th&gt;&lt;th&gt;Thứ 5&lt;/th&gt;&lt;th&gt;Thứ 6&lt;/th&gt;&lt;/tr&gt;</pre><p class="note">Hàng tiêu đề 2 chỉ viết 5 thẻ &lt;th&gt; vì 2 cột đầu đã được 2 ô rowspan="2" chiếm chỗ.</p>'),
    ('Tiết 2 (PPCT: 21) · Hoạt động 3: Luyện tập · 18 phút', 'Thực hành hoàn thiện trang web Câu lạc bộ', '<div class="cols"><div><ul><li>Hoàn thiện danh sách lồng nhau Hình 9.7 (trang 55).</li><li>Hoàn thiện bảng 7 hàng × 7 cột Hình 9.8 (trang 56).</li><li>Luyện tập tr.56: viền xanh (<code>green</code>) cho 2 hàng tiêu đề, đổi màu chữ bộ môn.</li></ul></div><div><p>Đánh giá theo Rubric 10 điểm (CCĐG 6 – Phụ lục 2)</p><a href="05_CLB_HoanChinh_SGK.html" target="_blank">Mở sản phẩm mẫu hoàn chỉnh</a></div></div>'),
    ('Tiết 2 (PPCT: 21) · Hoạt động 4: Vận dụng · 05 phút', 'Kiểm chứng gợi ý AI và mở rộng thực tiễn', '<div class="cols"><pre>&lt;table&gt;\n &lt;tr&gt;&lt;th&gt;CLB&lt;/th&gt;&lt;th&gt;Lịch&lt;/th&gt;&lt;/tr&gt;\n &lt;td&gt;Bóng đá&lt;/td&gt;\n &lt;td&gt;Thứ Ba&lt;/td&gt;\n&lt;/table&gt;</pre><div><p class="callout">Gợi ý giả định của AI:<br>“Chỉ cần thêm border vào table.”</p><p>Phát hiện lỗi <code>&lt;td&gt;</code> ngoài <code>&lt;tr&gt;</code> · Phiếu số 5 (CCĐG 7).</p></div></div>'),
    ('Tổng kết toàn bài · Tiết PPCT: 20, 21', 'Hoàn thành Bài 9: Tạo danh sách, bảng', '<p class="callout">Đọc mã cấu trúc → Đối chiếu SGK → Lưu UTF-8 → Chạy thử trình duyệt.</p><p>Nhiệm vụ về nhà (Vận dụng SGK trang 56): Tạo bảng Thời khóa biểu lớp 12/4 hoặc dùng Python sinh bảng HTML từ dữ liệu n × 4.</p>')
]
html_deck = '<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Bài 9 Tạo danh sách, bảng (Tiết PPCT 20, 21)</title><style>' + style_m + '</style></head><body><main>'
for i, (k, title, content) in enumerate(slides):
    html_deck += '<section class="slide' + (' active' if i == 0 else '') + '"><div class="kicker">' + k + '</div><h1>' + title + '</h1>' + content + '</section>'
html_deck += '</main><nav class="controls" aria-label="Chuyển trang"><button id="prev" aria-label="Trang trước">←</button><span id="count"></span><button id="next" aria-label="Trang sau">→</button><button id="full">Toàn màn hình</button></nav><script>' + script_m + '</script></body></html>'
deckpath.write_text(html_deck, encoding='utf-8')

(OUT / '03_Doi_chieu_gop_y_cua_thay.txt').write_text('''BẢNG ĐỐI CHIẾU CHỈNH SỬA CHI TIẾT THEO 39 NHẬN XÉT CỦA CÔ TUYẾT NGÂN TÔN VÀ THẦY THỐNG NHẤT
Nguồn nhận xét: De_cuong_Tin12_Bai9_Lop12-4 (2).docx (39 comments từ #0 đến #38)
File sản phẩm đã hoàn thiện: De_cuong_Tin12_Bai9_Lop12-4.docx (và bản đồng bộ tại 06_thanh_pham_docx)

1. NHẬN XÉT #0 & #1 (Cô Tuyết Ngân Tôn – Tiêu đề & Phân phối chương trình):
- Đã sửa tên tiêu đề trên trang bìa và đầu trang 1 thành: "ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY".
- Đã tra cứu Phân phối chương trình Tin học 12 KNTT và ghi rõ: "Thời lượng: 02 tiết – Tiết thứ 1, 2 của bài; Tiết PPCT: 20, 21" trên cả trang bìa, đầu trang 1, bảng tổng quan tiến trình và từng hoạt động/nhiệm vụ.
- Đã lưu trữ bảng Phân phối chương trình Tin học 10, 11, 12 KNTT và bộ KHBD tham khảo vào đúng ngăn 01_chuong_trinh_gdpt/ và 04_form_mau_khbd/.

2. NHẬN XÉT #2 (Cô Tuyết Ngân Tôn – Tên Sách giáo khoa):
- Đã sửa chính xác tên SGK tại mục II.2. Học liệu: "Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (và Sách giáo khoa Tin học 12 – Định hướng Khoa học máy tính), Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam, Bài 9: Tạo danh sách, bảng (trang 52 – 56)".

3. NHẬN XÉT #3 (Thầy Thống Nhất – Đính kèm tệp và đường dẫn học liệu):
- Đã bổ sung đầy đủ đường dẫn lưu trữ học liệu trực tuyến và đính kèm toàn bộ nội dung Phiếu học tập, Bảng kiểm, Rubric cùng mã nguồn các tệp HTML (01_Khung_CLB.html, 05_CLB_HoanChinh_SGK.html...) trực tiếp vào Mục IV. PHỤ LỤC của KHBD.

4. NHẬN XÉT #4, #5, #6, #7, #16, #20, #24, #27, #29 (Cô Tuyết Ngân Tôn – Cấu trúc 4 hoạt động theo Công văn 5512):
- Đã bỏ dòng tiêu đề thừa "4. Các hoạt động dạy học" ở đầu mục III (Nhận xét #7).
- Đã cấu trúc lại toàn bộ Tiến trình dạy học theo đúng 4 hoạt động của Công văn 5512:
  + 1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 / Tiết PPCT: 20).
  + 2. Hoạt động 2: Hình thành kiến thức mới (62 phút), gộp chung và chia thành 4 tiểu mục (Nhiệm vụ 1, 2, 3, 4) ghi rõ thời lượng và tiết thực hiện:
    * 2.1. Nhiệm vụ 1: Tìm hiểu danh sách có thứ tự và không có thứ tự (20 phút – Tiết 1 / Tiết PPCT: 20).
    * 2.2. Nhiệm vụ 2: Tìm hiểu danh sách mô tả và danh sách lồng nhau (20 phút – Tiết 1 / Tiết PPCT: 20).
    * 2.3. Nhiệm vụ 3: Tìm hiểu cấu trúc cơ bản của bảng trong HTML (10 phút – Tiết 2 / Tiết PPCT: 21).
    * 2.4. Nhiệm vụ 4: Tìm hiểu cách định dạng bảng và gộp ô trong bảng HTML (12 phút – Tiết 2 / Tiết PPCT: 21).
  + 3. Hoạt động 3: Luyện tập (18 phút – Tiết 2 / Tiết PPCT: 21) – đặt sau khi hoàn thành toàn bộ Hoạt động 2.
  + 4. Hoạt động 4: Vận dụng (05 phút – Tiết 2 / Tiết PPCT: 21).

5. NHẬN XÉT #8, #14, #17, #21, #25, #31 (Cô Tuyết Ngân Tôn – Đối chiếu Mục tiêu 1-1):
- Đã mã hóa từng mục tiêu ở Mục I ([KT1]–[KT6], [NLC1]–[NLC3], [NLa/c/d], [NLS/NLAI], [PC1]–[PC3]) và trích dẫn khớp 1-1 tại mục "a) Mục tiêu" của tất cả các Hoạt động / Nhiệm vụ, đảm bảo phủ kín 100% mục tiêu toàn bài.

6. NHẬN XÉT #9, #10, #11, #12, #13, #15, #19, #23, #26, #28, #30, #32 (Cô Tuyết Ngân Tôn – Bảng 2 cột và tên 4 bước ở mục d) Tổ chức thực hiện):
- Đã chuyển toàn bộ các bảng "d) Tổ chức thực hiện" sang bảng 2 cột chuẩn CV 5512:
  + Cột 1: HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH (trình bày đúng tên 4 bước: Bước 1. Giáo viên chuyển giao nhiệm vụ học tập; Bước 2. Học sinh thực hiện nhiệm vụ; Bước 3. Báo cáo sản phẩm; Bước 4. Đánh giá, nhận xét).
  + Cột 2: SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI) (trình bày đầy đủ nội dung kiến thức chốt, cú pháp HTML, ví dụ và đáp án).

7. NHẬN XÉT #18 & #22 (Thầy Thống Nhất – Chèn trực tiếp các hình SGK vào mục b) Nội dung):
- Đã trích xuất và chèn trực tiếp cả 8 hình minh họa từ SGK Tin học 12 KNTT (Hình 9.1, 9.2, 9.3, 9.4, 9.5, 9.6, 9.7, 9.8) kèm chú thích rõ nét ngay tại mục "b) Nội dung" của từng nhiệm vụ tương ứng.

8. NHẬN XÉT #33, #34, #35 (Cô Tuyết Ngân Tôn – Đổi mục IV thành PHỤ LỤC, bỏ mục V thừa và chi tiết hóa Rubric):
- Đã ghi rõ công cụ đánh giá (CCĐG 1 – CCĐG 7) ngay sau mục "c) Sản phẩm" của từng Hoạt động / Nhiệm vụ.
- Đã đổi tên Mục IV thành "IV. PHỤ LỤC" chứa đầy đủ Phiếu học tập số 1–5, Bảng kiểm/Thang điểm (CCĐG 1–7) và mã nguồn HTML.
- Đã mở rộng bảng Rubric (Phụ lục 2 – CCĐG 6) có đầy đủ bảng thông tin hành chính (Tên bài dạy, Lớp, Tiết PPCT, Giáo viên, Học sinh, Nhóm), cột STT, Tiêu chí, Câu lệnh/Cú pháp HTML bắt buộc kiểm tra, Mô tả chi tiết 3 mức độ đạt và Điểm số.
- Đã xóa bỏ hoàn toàn mục "V. KẾT QUẢ DỰ KIẾN" bị trùng lặp với Mục tiêu (Nhận xét #35).

9. NHẬN XÉT #36, #37, #38 (Cô Tuyết Ngân Tôn – Phần Nhận xét của GVHD và Ký duyệt):
- Đã đổi tiêu đề thành "V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT" và dòng "Nhận xét của Giáo viên hướng dẫn:".
- Đã in nghiêng dòng ngày tháng "Huế, ngày 04 tháng 11 năm 2026".
- Đã trình bày cân đối 2 cột chữ ký: "GIÁO VIÊN HƯỚNG DẪN (Thầy Mai Xuân Văn)" và "GIÁO VIÊN GIẢNG DẠY (Sinh viên thực hiện: Hồ Quang Huy)".
''', encoding='utf-8')

(OUT / 'DOC_TRUOC_KHI_NOP.txt').write_text('''HỒ SƠ DỰ GIỜ TIN HỌC 12 – BÀI 9: TẠO DANH SÁCH, BẢNG
(ĐÃ CHỈNH SỬA HOÀN CHỈNH THEO 39 NHẬN XÉT CỦA CÔ TUYẾT NGÂN TÔN & THẦY THỐNG NHẤT)

Sinh viên thực hiện: Hồ Quang Huy – MSV: 24S1020030
Giáo viên hướng dẫn: Thầy Mai Xuân Văn – Trường THPT Thuận Hoá
Lớp giảng dạy: 12/4 – Thời lượng bài học: 02 tiết (Tiết thứ 1, 2 của bài; Tiết PPCT: 20, 21)
Lịch dự giờ: Thứ Năm, ngày 05/11/2026 – Tiết 3 (09h00 – 09h45, dự giờ Tiết 1 / Tiết PPCT: 20) – Phòng H.II.4

1. BẢN CHÍNH KẾ HOẠCH BÀI DẠY ĐỂ IN VÀ NỘP
- Tệp: De_cuong_Tin12_Bai9_Lop12-4.docx (đồng bộ tại ngăn 06_thanh_pham_docx/Đề cương dự giờ Tin12_Bai9_Lop12-4_HoQuangHuy.docx)
- Đã tích hợp đầy đủ:
  + Trang bìa và trang đầu ghi rõ: ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY và Tiết PPCT: 20, 21.
  + Mục tiêu mã hóa [KT1]–[KT6], [NLC1]–[NLC3], [NLa/c/d], [NLS/NLAI], [PC1]–[PC3] khớp 1-1 với từng hoạt động.
  + Tiến trình dạy học chuẩn Công văn 5512 gồm 4 hoạt động:
    * Hoạt động 1: Mở đầu (05 phút – Tiết 1 / Tiết PPCT: 20)
    * Hoạt động 2: Hình thành kiến thức mới (62 phút: Nhiệm vụ 1 [20p, Tiết 1], Nhiệm vụ 2 [20p, Tiết 1], Nhiệm vụ 3 [10p, Tiết 2], Nhiệm vụ 4 [12p, Tiết 2])
    * Hoạt động 3: Luyện tập (18 phút – Tiết 2 / Tiết PPCT: 21)
    * Hoạt động 4: Vận dụng (05 phút – Tiết 2 / Tiết PPCT: 21)
  + Đã chèn trực tiếp 8 hình SGK (Hình 9.1 đến Hình 9.8) vào mục b) Nội dung.
  + Bảng d) Tổ chức thực hiện 2 cột chuẩn: HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH (4 bước chuẩn) | SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI).
  + Mục IV. PHỤ LỤC đính kèm trọn bộ Phiếu học tập số 1–5, Công cụ đánh giá CCĐG 1–7, Rubric chi tiết và mã nguồn HTML.
  + Mục V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT đúng quy chuẩn.

2. TỆP PHIẾU HỌC TẬP & RUBRIC RỜI ĐỂ IN PHÁT HỌC SINH
- Tệp: 02_Phieu_hoc_tap_va_dap_an_Tin12_Bai9.docx
- Trang 1: Phiếu học tập Tiết 1 (Phiếu số 1 & số 2 – Tiết PPCT: 20).
- Trang 2: Phiếu học tập Tiết 2 (Phiếu số 3, số 4, Luyện tập & Phiếu số 5 – Tiết PPCT: 21).
- Trang 3: Phiếu đánh giá Rubric 10 điểm (CCĐG 6).
- Trang 4: Đáp án và hướng dẫn chấm dành cho giáo viên.

3. THƯ MỤC HỌC LIỆU THỰC HÀNH (Hoc_lieu_HTML/)
- 00_Trinh_chieu.html: Bài trình chiếu 12 slide ngoại tuyến đồng bộ theo cấu trúc mới.
- 01_Khung_CLB.html: Tệp khung HTML cấp cho học sinh thực hành xuyên suốt 2 tiết.
- 02_CLB_Mau_CoBan.html: Tệp bảng cơ bản minh họa cho Nhiệm vụ 3.
- 03_Ma_Can_Sua.html & 04_Ma_Da_Sua.html: Học liệu kiểm chứng lỗi cấu trúc bảng cho Hoạt động 4 (Vận dụng).
- 05_CLB_HoanChinh_SGK.html: Sản phẩm mẫu hoàn chỉnh tích hợp Hình 9.7, Hình 9.8 và bài Luyện tập trang 56.
''', encoding='utf-8')

# Rebuild ZIP package
archive = TASK / 'HoQuangHuy_24S1020030_DuGio_Tin12_Bai9_05-11-2026.zip'
files_to_zip = sorted(p for p in OUT.rglob('*') if p.is_file())
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
    for fpath in files_to_zip:
        z.write(fpath, fpath.relative_to(OUT).as_posix())

print('Successfully built revised', MAIN)
print('Successfully built revised', HAND)
print('Copied to', THANH_PHAM)
print('Updated archive', archive)
