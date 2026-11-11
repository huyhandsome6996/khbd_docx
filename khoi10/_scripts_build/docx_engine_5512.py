# -*- coding: utf-8 -*-
"""
Core Word (.docx) XML Builder Engine for Công văn 5512 Lesson Plans (Khối 10 - THPT Thuận Hoá).
Applies 100% of the 42 reviewer comments from Cô Tuyết Ngân Tôn, Thầy Thống Nhất, and Thầy Nguyễn Nhất Tân.
"""
from pathlib import Path
from copy import deepcopy as dc
import zipfile, xml.etree.ElementTree as E, io, re, shutil
from PIL import Image

TEMPLATE_DOCX = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_template_5512.docx')
HANDOUT_TEMPLATE_DOCX = Path(r'c:\Users\PC\khbd_docx\04_form_mau_khbd\_engine_5512\base_worksheet_5512.docx')
CROPPED_DIR = Path(r'c:\Users\PC\khbd_docx\khoi10\_scripts_build\cropped_figures')

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

# All cropped SGK Grade 10 figures
FIG_MAP_KHOI10 = {
    # On tap GHKI
    'hinh_1_1_xu_li_tt.png': ('rId31', 4200000, 'Hình 1.1. Quá trình xử lí thông tin/dữ liệu bằng máy tính (SGK Tin học 10 KNTT, Bài 1, trang 6)'),
    'bang_1_1_don_vi_luu_tru.png': ('rId32', 4680000, 'Bảng 1.1. Các đơn vị lưu trữ dữ liệu hơn kém nhau 2^10 = 1024 lần (SGK Tin học 10 KNTT, Bài 1, trang 8)'),
    'hinh_8_2_bang_so_sanh_mang.png': ('rId33', 4500000, 'Bảng đối chiếu phạm vi, cách kết nối và sở hữu giữa mạng LAN và Internet (SGK Tin học 10 KNTT, Bài 8, trang 39)'),
    'hop_9_1_bao_ve_tt_ca_nhan.png': ('rId34', 3600000, 'Hộp kiến thức: Các biện pháp bảo vệ thông tin cá nhân trên không gian mạng (SGK Tin học 10 KNTT, Bài 9, trang 44)'),
    'hop_11_1_vi_pham_tren_mang.png': ('rId35', 4400000, 'Hộp kiến thức: Những hành vi vi phạm về đạo đức, pháp luật và văn hoá khi sử dụng mạng (SGK Tin học 10 KNTT, Bài 11, trang 56)'),
    # Bai 12
    'hinh_12_1_bitmap_vector.png': ('rId41', 4300000, 'Hình 12.1a (Ảnh chụp điểm ảnh bitmap) và Hình 12.1b (Hình vẽ đồ hoạ vectơ) (SGK Tin học 10 KNTT, Bài 12, trang 63)'),
    'hinh_12_2_va_bang_so_sanh.png': ('rId42', 4680000, 'Hình 12.2. Phóng to một phần ảnh gốc và Bảng so sánh đồ hoạ điểm ảnh với đồ hoạ vectơ (SGK Tin học 10 KNTT, Bài 12, trang 64)'),
    'hinh_12_3_man_hinh_inkscape.png': ('rId43', 4680000, 'Hình 12.3. Màn hình làm việc của phần mềm thiết kế đồ hoạ Inkscape (SGK Tin học 10 KNTT, Bài 12, trang 65)'),
    'hinh_12_4_cong_cu_inkscape.png': ('rId44', 4500000, 'Hình 12.4. Một số công cụ vẽ hình cơ bản có sẵn trong Hộp công cụ của Inkscape (SGK Tin học 10 KNTT, Bài 12, trang 66)'),
    'hinh_12_5_ve_bong_hoa.png': ('rId45', 4680000, 'Hình 12.5. Các bước thực hành vẽ bông hoa từ hình tròn và những hình elip (SGK Tin học 10 KNTT, Bài 12, trang 67)'),
    'hinh_12_6_ve_quoc_ki.png': ('rId46', 4400000, 'Hình 12.6. Các bước thực hành vẽ Quốc kì Việt Nam trong Inkscape (SGK Tin học 10 KNTT, Bài 12, trang 68)'),
    'hinh_12_van_dung_quoc_ki.png': ('rId47', 4500000, 'Hình minh hoạ bài tập Vận dụng 1 & 2: Vẽ Quốc kì các nước và ghép hình cơ bản (SGK Tin học 10 KNTT, Bài 12, trang 68)'),
    # Bai 13
    'hinh_13_1_2_dua_hau_thuoc_tinh.png': ('rId51', 4500000, 'Hình 13.1 (Miếng dưa hấu) và Hình 13.2 (Các đối tượng hình khối biến đổi từ công cụ hình sao/đa giác) (SGK Tin học 10 KNTT, Bài 13, trang 69)'),
    'bang_13_1_thuoc_tinh_hinh_khoi.png': ('rId52', 4680000, 'Bảng các thuộc tính cơ bản của một số đối tượng hình khối có sẵn trong Inkscape (SGK Tin học 10 KNTT, Bài 13, trang 69–70)'),
    'hinh_13_3_4_fill_stroke.png': ('rId53', 4680000, 'Hình 13.3 (Các kiểu tô màu) và Hình 13.4 (Hộp thoại Fill and Stroke trong Inkscape) (SGK Tin học 10 KNTT, Bài 13, trang 70)'),
    'hinh_13_5_ket_hop_doi_tuong.png': ('rId54', 4500000, 'Hình 13.5. Ghép hai hình tròn và một hình vuông xoay góc để tạo thành hình trái tim (SGK Tin học 10 KNTT, Bài 13, trang 71)'),
    'hinh_13_6_cac_phep_ghep_hinh.png': ('rId55', 4680000, 'Hình 13.6. Kết quả thực hiện 6 phép ghép đối tượng đồ hoạ trong bảng chọn Path (SGK Tin học 10 KNTT, Bài 13, trang 72)'),
    'hinh_13_7_8_dam_may_logo.png': ('rId56', 3600000, 'Hình 13.7. Ghép các hình tròn bằng phép hợp (Path/Union) để tạo hình đám mây (SGK Tin học 10 KNTT, Bài 13, trang 72)'),
    'hinh_13_9_ve_logo.png': ('rId57', 4400000, 'Hình 13.8 & Hình 13.9. Các bước thực hành vẽ biểu tượng (logo) bằng cung tròn và độ dày nét vẽ (SGK Tin học 10 KNTT, Bài 13, trang 72–73)'),
    'hinh_13_10_11_ve_cay.png': ('rId58', 4680000, 'Hình 13.10 & Hình 13.11. Các bước thực hành vẽ hình cây bằng phép hiệu (Difference) và phép hợp (Union) (SGK Tin học 10 KNTT, Bài 13, trang 73)'),
    'hinh_13_12_13_hinh_cau.png': ('rId59', 4400000, 'Hình 13.12 & Hình 13.13. Các bước thực hành vẽ hình cầu 3D bằng tô màu chuyển (Gradient) (SGK Tin học 10 KNTT, Bài 13, trang 73–74)'),
    'hinh_13_14_15_16_luyen_tap.png': ('rId60', 4500000, 'Hình 13.14, 13.15, 13.16. Các sản phẩm thực hành phần Luyện tập Bài 13 (SGK Tin học 10 KNTT, Bài 13, trang 74)'),
}

_doc_pr_counter = [200]
def add_figure(body, fig_name):
    rid, cx, caption_text = FIG_MAP_KHOI10[fig_name]
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
                    cell.append(p(line.strip()[1:-1], 'code', align='left', sz=21))
                else:
                    cell.append(p(line, 'cell', align='left', sz=cell_sz))
            row.append(cell)
        t.append(row)
    return t

def signature_table(
    sign_date_str='Huế, ngày 01 tháng 11 năm 2026',
    gvhd_name='Thầy Mai Xuân Văn',
    gvgd_name='Thầy Mai Xuân Văn',
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
    cover_subtitle='ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I (01 TIẾT – TIẾT PPCT: 15)',
    class_str='Lớp 10/2 – Trường THPT Thuận Hoá',
    schedule_str='Thứ Hai, ngày 02/11/2026 – Tiết 2 (08h00 – 08h45) – Phòng H.I.4',
    lesson_row_str='Ôn tập kiểm tra Giữa học kì I (01 tiết – Tiết PPCT: 15)',
    page1_title='TÊN BÀI DẠY: ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I',
    page1_subtitle='(Thời lượng: 01 tiết – Tiết PPCT: 15 theo Kế hoạch dạy học môn Tin học Khối 10 Trường THPT Thuận Hoá)',
    page1_class_line='Môn học: Tin học; Lớp: 10/2',
    page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết PPCT: 15)',
    page1_sched_line='Lịch dự giờ: Thứ Hai, ngày 02/11/2026 – Tiết 2 (08h00 – 08h45) – Lớp 10/2 – Phòng H.I.4',
    page1_mode_line='Hình thức tổ chức: Dạy học trên lớp kết hợp thảo luận nhóm và sơ đồ tư duy'
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

        set_cell_text(0, 0, 'Giáo viên giảng dạy / GVHD', bold=True)
        set_cell_text(0, 2, 'Thầy Mai Xuân Văn', bold=True)
        set_cell_text(1, 0, 'Sinh viên thực hiện', bold=True)
        set_cell_text(1, 2, 'Hồ Quang Huy', bold=True)
        set_cell_text(5, 2, class_str)
        set_cell_text(6, 0, 'Lịch dự giờ', bold=True)
        set_cell_text(6, 2, schedule_str)
        set_cell_text(7, 2, lesson_row_str)

    # Page 1 Header
    tabline(b, 'Trường: THPT Thuận Hoá', 'Giáo viên giảng dạy: Thầy Mai Xuân Văn')
    tabline(b, 'Tổ: Tin học', 'Sinh viên thực hiện: Hồ Quang Huy')
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
    sign_date_str='Huế, ngày 01 tháng 11 năm 2026',
    gvhd_name='Thầy Mai Xuân Văn',
    gvgd_name='Thầy Mai Xuân Văn',
    sv_title='SINH VIÊN THỰC HIỆN',
    sv_name='Hồ Quang Huy',
    section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT'
):
    add(b, section_title, 'heading')
    add(b, '**Nhận xét của Giáo viên hướng dẫn:**', first=0, keep=True)
    for _ in range(3):
        add(b, '.' * 115, first=0, keep=True)
    b.append(signature_table(sign_date_str, gvhd_name, gvgd_name, sv_title, sv_name))

def ending(pkg, start=1):
    sect = dc(E.fromstring(pkg['word/document.xml']).find(w('body'))[-1])
    for n in list(sect.findall(w('pgNumType'))):
        sect.remove(n)
    sect.append(node('pgNumType', {'start': start}))
    return sect


_OOXML_ORDER = {
    'pPr': 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(),
    'rPr': 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath'.split(),
    'tcPr': 'cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark'.split(),
    'tblPr': 'tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook'.split(),
    'sectPr': 'headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings'.split(),
}

def _local(tag):
    return str(tag).split('}')[-1].split(':')[-1]

def sort_ooxml_children(root):
    """Sắp xếp lại thứ tự phần tử con theo đúng lược đồ OOXML để Microsoft Word mở file không báo lỗi."""
    for el in root.iter():
        order = _OOXML_ORDER.get(_local(el.tag))
        if not order or len(el) < 2:
            continue
        kids = list(el)
        keyed = [(order.index(_local(c.tag)) if _local(c.tag) in order else len(order), i, c) for i, c in enumerate(kids)]
        if [k[0] for k in keyed] == sorted(k[0] for k in keyed):
            continue
        for c in kids:
            el.remove(c)
        for _, _, c in sorted(keyed, key=lambda k: (k[0], k[1])):
            el.append(c)

def save_docx(body, out_path, used_figures=None, is_handout=False):
    pkg = HPKG_TEMPLATE if is_handout else PKG_TEMPLATE
    body.append(ending(pkg, start=1))
    data = dict(pkg)
    if used_figures:
        rels_xml = data['word/_rels/document.xml.rels'].decode('utf-8')
        for fig_name in used_figures:
            rid, _, _ = FIG_MAP_KHOI10[fig_name]
            data['word/media/' + fig_name] = (CROPPED_DIR / fig_name).read_bytes()
            rel_tag = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fig_name}"/>'
            if f'Id="{rid}"' not in rels_xml:
                rels_xml = rels_xml.replace('</Relationships>', rel_tag + '</Relationships>')
        data['word/_rels/document.xml.rels'] = rels_xml.encode('utf-8')

    sort_ooxml_children(body)
    src = data['word/document.xml'].decode('utf-8')
    prefix = src[:src.index('<w:body')]
    data['word/document.xml'] = (prefix + E.tostring(body, encoding='unicode') + '</w:document>').encode('utf-8')
    if 'word/settings.xml' in data:
        s = data['word/settings.xml'].decode('utf-8')
        s = re.sub(r'<w:updateFields[^>]*/>', '', s)  # tránh hộp thoại cập nhật trường khi mở bằng Word
        data['word/settings.xml'] = s.encode('utf-8')
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, value in data.items():
            z.writestr(n, value)
    return out_path
