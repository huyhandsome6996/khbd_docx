# -*- coding: utf-8 -*-
"""
Core Word (.docx) XML Builder Engine for Công văn 5512 Lesson Plans (Khối 12 - THPT Thuận Hoá).
Applies 100% of the 42 reviewer comments from Cô Tuyết Ngân Tôn, Thầy Thống Nhất, and Thầy Nguyễn Nhất Tân,
plus all formatting and pagination rules (100% Times New Roman, no mid-sentence bold/caps/italics, no blank page gaps).
"""
from pathlib import Path
from copy import deepcopy as dc
import zipfile, xml.etree.ElementTree as E, io, re, shutil
from PIL import Image

TEMPLATE_DOCX = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_template_5512.docx')
HANDOUT_TEMPLATE_DOCX = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_worksheet_5512.docx')
CROPPED_DIR = Path(r'c:\Users\PC\khbd_docx\khoi12\_scripts_build\cropped_figures')

def load_package(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}

PKG_TEMPLATE = load_package(TEMPLATE_DOCX)
HPKG_TEMPLATE = load_package(HANDOUT_TEMPLATE_DOCX)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
w = lambda t: '{' + W + '}' + t

for data in [HPKG_TEMPLATE['word/document.xml'], PKG_TEMPLATE['word/document.xml']]:
    for _, (prefix, uri) in E.iterparse(io.BytesIO(data), events=['start-ns']):
        if prefix and not re.fullmatch(r'ns\d+', prefix):
            E.register_namespace(prefix, uri)
E.register_namespace('w', W)
E.register_namespace('wp', WP)
E.register_namespace('a', A)
E.register_namespace('pic', PIC)
E.register_namespace('r', R)

_root_model = E.fromstring(PKG_TEMPLATE['word/document.xml'])
_oldbody_model = _root_model.find(w('body'))
_bullet_model = next(p for p in _oldbody_model.iter(w('p')) if p.find(w('pPr') + '/' + w('numPr')) is not None)

def node(t, attrs=None):
    return E.Element(w(t), {w(k): str(v) for k, v in (attrs or {}).items()})

def font(rp, size_val=26, font_name='Times New Roman', color_val='000000'):
    rp.append(node('rFonts', {'ascii': font_name, 'hAnsi': font_name, 'cs': font_name, 'eastAsia': font_name}))
    rp.append(node('color', {'val': color_val}))
    rp.append(node('sz', {'val': size_val}))
    rp.append(node('szCs', {'val': size_val}))

def p(text='', role='body', align=None, keep=False, first=None, italic=False, sz=26):
    out = node('p')
    pp = node('pPr')
    out.append(pp)
    if role == 'title':
        pp.append(node('pStyle', {'val': 'Title'}))
    if role == 'bullet':
        pp.append(dc(_bullet_model.find(w('pPr') + '/' + w('numPr'))))
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
    if align is None and role in ['body', 'bullet']:
        _plain = re.sub(r'\*+', '', str(text))
        if re.search(r'<[A-Za-z/!]|=\"|https?://|\.(?:html|svg|docx|css|js|png)\b|\w_\w|#\w|/\w+/|\S{20,}', _plain):
            align = 'left'
    pp.append(node('jc', {'val': align or ('center' if role in ['title', 'caption'] else 'left' if role in ['heading', 'sub', 'label', 'code', 'cell'] else 'both')}))
    rp = node('rPr')
    font(rp, size_val=sz, font_name='Times New Roman')
    pp.append(rp)

    text = re.sub(r'\*\*`([^`]+)`\*\*', r'**\1**', str(text))
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
        font(rp, size_val=sz, font_name='Times New Roman')
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

_pending_breaks = set()
def add(body, text='', role='body', **kwargs):
    if '\n' in str(text) and role in ['body', 'bullet']:
        last_out = None
        for line in str(text).split('\n'):
            line = line.strip()
            if line:
                last_out = add(body, line, role=role, **kwargs)
        return last_out
    out = p(text, role, **kwargs)
    if id(body) in _pending_breaks:
        out.find(w('pPr')).append(node('pageBreakBefore'))
        _pending_breaks.remove(id(body))
    body.append(out)
    return out

def br(body):
    _pending_breaks.add(id(body))

def tabline(body, left, right):
    out = p('', first=0, align='left')
    tabs = node('tabs')
    tabs.append(node('tab', {'val': 'left', 'pos': 4200}))
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

# All cropped SGK Grade 12 figures for Bài 10
FIG_MAP_KHOI12 = {
    'hinh_10_1.png': ('rId71', 4500000, 'Hình 10.1. Cấu trúc một website minh hoạ các trường hợp liên kết tương đối (SGK Tin học 12 KNTT, Bài 10, trang 59)'),
    'hinh_10_2.png': ('rId72', 4500000, 'Hình 10.2. Nội dung giới thiệu hoạt động ngoại khoá bổ sung vào đầu trang web CLB.html (SGK Tin học 12 KNTT, Bài 10, trang 61)'),
}

_doc_pr_counter = [300]
def add_figure(body, fig_name):
    rid, cx, caption_text = FIG_MAP_KHOI12[fig_name]
    img_path = CROPPED_DIR / fig_name
    with Image.open(img_path) as im:
        w_px, h_px = im.size
    cy = int(cx * h_px / w_px)
    _doc_pr_counter[0] += 1
    did = _doc_pr_counter[0]

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

def table(rows, widths, header=True, cell_sz=24, cant_split=False):
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
        if cant_split or (header and i == 0):
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
                    cell.append(p(line.strip()[1:-1], 'code', align='left', sz=22))
                else:
                    cell.append(p(line, 'cell', align='left', sz=cell_sz))
            row.append(cell)
        t.append(row)
    return t

def signature_table(
    sign_date_str='Huế, ngày 12 tháng 11 năm 2026',
    gvhd_name='Thầy Mai Xuân Văn',
    gvgd_name='Thầy Nguyễn Nhất Tân',
    sv_title='SINH VIÊN THỰC HIỆN',
    sv_name='Hồ Quang Huy'
):
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

    # Row 1: Giáo viên hướng dẫn (trái) và Giáo viên giảng dạy (phải)
    row = node('tr')
    trp = node('trPr')
    trp.append(node('cantSplit'))
    row.append(trp)

    c1 = node('tc')
    cp1 = node('tcPr')
    cp1.append(node('tcW', {'w': widths[0], 'type': 'dxa'}))
    c1.append(cp1)
    c1.append(p('', 'cell', align='center', keep=True))
    c1.append(p('**GIÁO VIÊN HƯỚNG DẪN**', 'cell', align='center', keep=True))
    c1.append(p('(Ký và ghi rõ họ tên)', 'cell', align='center', italic=True, keep=True, sz=24))
    for _ in range(2):
        c1.append(p('', 'cell', align='center', keep=True))
    c1.append(p(f'**{gvhd_name}**', 'cell', align='center'))

    c2 = node('tc')
    cp2 = node('tcPr')
    cp2.append(node('tcW', {'w': widths[1], 'type': 'dxa'}))
    c2.append(cp2)
    c2.append(p(sign_date_str, 'cell', align='center', italic=True, keep=True))
    c2.append(p('**GIÁO VIÊN GIẢNG DẠY**', 'cell', align='center', keep=True))
    c2.append(p('(Ký và ghi rõ họ tên)', 'cell', align='center', italic=True, keep=True, sz=24))
    for _ in range(2):
        c2.append(p('', 'cell', align='center', keep=True))
    c2.append(p(f'**{gvgd_name}**', 'cell', align='center'))

    row.append(c1)
    row.append(c2)
    t.append(row)

    # Row 2: Sinh viên thực hiện ký phía dưới
    row2 = node('tr')
    trp2 = node('trPr')
    trp2.append(node('cantSplit'))
    row2.append(trp2)

    c2_1 = node('tc')
    cp2_1 = node('tcPr')
    cp2_1.append(node('tcW', {'w': widths[0], 'type': 'dxa'}))
    c2_1.append(cp2_1)
    c2_1.append(p('', 'cell'))

    c2_2 = node('tc')
    cp2_2 = node('tcPr')
    cp2_2.append(node('tcW', {'w': widths[1], 'type': 'dxa'}))
    c2_2.append(cp2_2)
    c2_2.append(p('', 'cell', align='center', keep=True))
    c2_2.append(p(f'**{sv_title}**', 'cell', align='center', keep=True))
    c2_2.append(p('(Ký và ghi rõ họ tên)', 'cell', align='center', italic=True, keep=True, sz=24))
    for _ in range(2):
        c2_2.append(p('', 'cell', align='center', keep=True))
    c2_2.append(p(f'**{sv_name}**', 'cell', align='center'))

    row2.append(c2_1)
    row2.append(c2_2)
    t.append(row2)
    return t

def init_body_with_cover(
    cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
    cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT – TIẾT PPCT: 22)',
    class_str='Tin học 12 – Lớp 12/2 – Trường THPT Thuận Hoá',
    schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.II.2',
    lesson_row_str='Bài 10: Tạo liên kết (Tiết 1 - Lí thuyết – Tiết PPCT: 22)',
    page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT)',
    page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 22, 23] – Kế hoạch này chi tiết hoá Tiết thứ 1 của bài / Tiết PPCT: 22)',
    page1_class_line='Môn học: Tin học; Lớp: 12/2',
    page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 1 - LT / Tiết PPCT: 22)',
    page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 12/2 – Phòng H.II.2',
    page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp quan sát trực quan và thảo luận nhóm',
    gvgd_name='Thầy Nguyễn Nhất Tân',
    gvhd_name='Thầy Mai Xuân Văn',
    sv_name='Hồ Quang Huy'
):
    b = node('body')
    for idx, el in enumerate(list(_oldbody_model)[:25]):
        if idx in (11, 12, 13, 14, 15):
            continue
        b.append(dc(el))

    cover = list(b)
    for el in cover:
        for rp in el.iter(w('rPr')):
            for tag in ['sz', 'szCs']:
                size = rp.find(w(tag))
                if size is not None:
                    size.set(w('val'), '26')

    cover_h1 = cover[6]
    for child in list(cover_h1):
        if child.tag != w('pPr'):
            cover_h1.remove(child)
    for run in p(cover_title, 'title').findall(w('r')):
        cover_h1.append(run)

    cover_h2 = cover[7]
    for child in list(cover_h2):
        if child.tag != w('pPr'):
            cover_h2.remove(child)
    for run in p(cover_subtitle, 'title').findall(w('r')):
        cover_h2.append(run)

    cover_table = cover[10]
    c_rows = cover_table.findall(w('tr'))
    if len(c_rows) >= 8:
        def set_cell_text(row_idx, cell_idx, text_val, bold=False):
            tc = c_rows[row_idx].findall(w('tc'))[cell_idx]
            for ch in list(tc):
                if ch.tag != w('tcPr'):
                    tc.remove(ch)
            tc.append(p(f'**{text_val}**' if bold else text_val, 'cell', sz=26))

        set_cell_text(0, 0, 'GV giảng dạy / GV hướng dẫn', bold=True)
        set_cell_text(0, 2, f'GVGD: {gvgd_name} – GVHD: {gvhd_name}', bold=True)
        set_cell_text(1, 0, 'Sinh viên thực hiện', bold=True)
        set_cell_text(1, 2, sv_name, bold=True)
        set_cell_text(5, 2, class_str)
        set_cell_text(6, 0, 'Lịch dự giờ', bold=True)
        set_cell_text(6, 2, schedule_str)
        set_cell_text(7, 2, lesson_row_str)

    # Page 1 Header
    tabline(b, 'Trường: THPT Thuận Hoá', f'Giáo viên giảng dạy: {gvgd_name}')
    tabline(b, 'Tổ: Tin học', f'Giáo viên hướng dẫn: {gvhd_name}')
    tabline(b, f'Sinh viên thực hiện: {sv_name}', 'Mã sinh viên: 24S1020030')
    add(b, cover_title, 'title')
    add(b, page1_title, 'title')
    add(b, page1_subtitle, align='center', first=0, keep=True, italic=True)
    add(b, page1_class_line, align='center', first=0, keep=True)
    add(b, page1_time_line, align='center', first=0, keep=True)
    add(b, page1_sched_line, align='center', first=0, keep=True)
    add(b, page1_mode_line, align='center', first=0)
    return b

def add_approval_section(
    b,
    sign_date_str='Huế, ngày 12 tháng 11 năm 2026',
    gvhd_name='Thầy Mai Xuân Văn',
    gvgd_name='Thầy Nguyễn Nhất Tân',
    sv_title='SINH VIÊN THỰC HIỆN',
    sv_name='Hồ Quang Huy',
    section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
):
    add(b, section_title, 'heading')
    add(b, '**Nhận xét của Giáo viên hướng dẫn và Giáo viên giảng dạy:**', first=0, keep=True)
    for _ in range(3):
        add(b, '.' * 115, first=0, keep=True)
    b.append(signature_table(sign_date_str, gvhd_name, gvgd_name, sv_title, sv_name))

def ending(pkg, start=1):
    sect = dc(E.fromstring(pkg['word/document.xml']).find(w('body'))[-1])
    for n in list(sect.findall(w('pgNumType'))):
        sect.remove(n)
    sect.append(node('pgNumType', {'start': start}))
    return sect

def save_docx(body, out_path, used_figures=None, is_handout=False):
    pkg = HPKG_TEMPLATE if is_handout else PKG_TEMPLATE
    body.append(ending(pkg, start=1))
    data = dict(pkg)
    if used_figures:
        rels_xml = data['word/_rels/document.xml.rels'].decode('utf-8')
        for fig_name in used_figures:
            rid, _, _ = FIG_MAP_KHOI12[fig_name]
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
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, value in data.items():
            z.writestr(n, value)
    return out_path

