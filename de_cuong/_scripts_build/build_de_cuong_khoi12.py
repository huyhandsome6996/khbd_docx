# -*- coding: utf-8 -*-
"""
Tao bo De cuong du gio (Kien tap su pham - co cac buoc ngan gon theo 4 hoat dong)
theo y kien gop y cua Thay Nguyen Nhat Tan va bam sat:
- Phan phoi chuong trinh Khoi 12 (khoi_12_tin_hoc_07112026.xls)
- Sach giao khoa Tin hoc 12 - Dinh huong Tin hoc ung dung (KNTT)
- Cong van 5512/BGDDT-GDTrH & Thong tu 02/2025/TT-BGDDT

Xuat ra thu muc: c:/Users/PC/khbd_docx/de_cuong/
  1. de_cuong_du_gio_12-4.docx (Bai 9: Tao danh sach, bang - Tiet PPCT 19, 20 - Lop 12/4 - Phong H.II.4)
  2. de_cuong_du_gio_12-2.docx (Bai 10: Tao lien ket - Tiet 1 LT - Tiet PPCT 21 - Lop 12/2 - Phong H.II.2)
  3. de_cuong_du_gio_12-1.docx (Bai 10: Tao lien ket - Tiet 2 TH - Tiet PPCT 22 - Lop 12/1 - Phong H.II.1)
"""

import sys
import shutil
import re
import zipfile
from pathlib import Path
from lxml import etree as LX

ROOT_DIR = Path(r'c:\Users\PC\khbd_docx')
SCRIPTS_12_DIR = ROOT_DIR / 'khoi12' / '_scripts_build'
sys.path.insert(0, str(SCRIPTS_12_DIR))

import docx_engine_5512 as eng

DE_CUONG_DIR = ROOT_DIR / 'de_cuong'
KHOI12_DIR = ROOT_DIR / 'khoi12'
BAI9_FIG_DIR = ROOT_DIR / '09_nop_du_gio_Tin12_Bai9_05-11-2026' / '_scripts_build' / 'cropped_figures'

BAI9_FIGS = {
    'hinh_9_1.png': ('rId21', 4500000, 'Hình 9.1. Danh sách có thứ tự trong HTML và kết quả hiển thị (SGK trang 52)'),
    'hinh_9_2.png': ('rId22', 4500000, 'Hình 9.2. Danh sách không có thứ tự dùng kí tự đầu dòng hình vuông (SGK trang 52)'),
    'hinh_9_3.png': ('rId23', 4680000, 'Hình 9.3. Ví dụ tạo danh sách mô tả bằng cặp thẻ <dl>, <dt>, <dd> (SGK trang 53)'),
    'hinh_9_4.png': ('rId24', 4680000, 'Hình 9.4. Cấu trúc HTML cơ bản tạo bảng và kết quả hiển thị (SGK trang 54)'),
    'hinh_9_5.png': ('rId25', 4680000, 'Hình 9.5. Ví dụ sử dụng thuộc tính style điều chỉnh kích thước bảng, hàng, ô (SGK trang 54)'),
    'hinh_9_6.png': ('rId26', 4680000, 'Hình 9.6. Ví dụ gộp hàng (rowspan) và gộp cột (colspan) của bảng (SGK trang 55)'),
    'hinh_9_7.png': ('rId27', 2050000, 'Hình 9.7. Mẫu danh sách các Câu lạc bộ lồng nhau cần tạo ở Nhiệm vụ 1 (SGK trang 55)'),
    'hinh_9_8.png': ('rId28', 3960000, 'Hình 9.8. Mẫu bảng Lịch hoạt động CLB Thể thao cần tạo ở Nhiệm vụ 2 (SGK trang 56)'),
}

for fname, meta in BAI9_FIGS.items():
    src_fig = BAI9_FIG_DIR / fname
    dst_fig = eng.CROPPED_DIR / fname
    if src_fig.exists() and not dst_fig.exists():
        shutil.copy2(src_fig, dst_fig)
    eng.FIG_MAP_KHOI12[fname] = meta

GVGD_NAME = 'Thầy Nguyễn Nhất Tân'
GVHD_NAME = 'Thầy Mai Xuân Văn'
SV_NAME = 'Hồ Quang Huy'


def _set_run_sz(para, sz):
    """Đặt cỡ chữ cho toàn bộ run trong một đoạn."""
    for run in para.iter(eng.w('r')):
        rp = run.find(eng.w('rPr'))
        if rp is None:
            continue
        for tag in ('sz', 'szCs'):
            node = rp.find(eng.w(tag))
            if node is None:
                node = eng.node(tag)
                rp.append(node)
            node.set(eng.w('val'), str(sz))


def add_figure_sgk(b, fig_name):
    """Chèn ảnh cắt từ SGK; ảnh đã có chú thích in sẵn nên bỏ caption trùng."""
    eng.add_figure(b, fig_name)
    b.remove(b[-1])
    pp = b[-1].find(eng.w('pPr'))
    for node in list(pp.findall(eng.w('keepNext'))):
        pp.remove(node)


def add_steps(b, b1, b2, b3, b4):
    """Mô tả cách thức tiến hành hoạt động bằng một đoạn văn liền mạch, không chia bước."""
    para = eng.add(b, f'**d) Cách thức tiến hành:** {b1} {b2} {b3} {b4}', first=0)
    pp = para.find(eng.w('pPr'))
    for node in list(pp.findall(eng.w('keepLines'))):
        pp.remove(node)


def save_de_cuong(body, out_path, used_figures):
    """Áp dụng quy tắc riêng của đề cương, không thay đổi engine giáo án."""
    w = eng.w
    for para in body.findall(w('p')):
        text = ''.join(t.text or '' for t in para.iter(w('t')))
        tabs = para.find(w('pPr') + '/' + w('tabs'))
        if tabs is not None and any(tab.get(w('pos')) in ('3800', '4200') for tab in tabs.findall(w('tab'))):
            for tab in tabs.findall(w('tab')):
                if tab.get(w('pos')) in ('3800', '4200'):
                    tab.set(w('val'), 'right')
                    tab.set(w('pos'), '9071')
            _set_run_sz(para, 24)
        if text.startswith(('(Đề cương', 'Môn học:', 'Thời lượng', 'Lịch dự giờ:', 'Hình thức:')):
            _set_run_sz(para, 24)

    cover = None
    for tbl in body.findall(w('tbl')):
        grid = tbl.find(w('tblGrid'))
        if grid is not None and sum(int(c.get(w('w'), 0)) for c in grid) == 7370:
            cover = tbl
            break
    if cover is not None:
        for t in cover.iter(w('t')):
            if (t.text or '').strip() == 'GV giảng dạy / GV hướng dẫn':
                t.text = 'Giáo viên'
        body.insert(list(body).index(cover), eng.p('', first=0, sz=18))

    for para in body.iter(w('p')):
        pp = para.find(w('pPr'))
        if pp is None:
            pp = eng.node('pPr')
            para.insert(0, pp)
        jc = pp.find(w('jc'))
        if jc is not None and jc.get(w('val')) in ('both', 'distribute'):
            jc.set(w('val'), 'left')
        para_text = ''.join(t.text or '' for t in para.iter(w('t')))
        # Đoạn mô tả cách thức tiến hành dài nên cho phép ngắt trang tự nhiên.
        if not para_text.startswith('d) Cách thức tiến hành') and pp.find(w('keepLines')) is None:
            pp.append(eng.node('keepLines'))
        for br in para.iter(w('br')):
            if br.get(w('type')) is None:
                if jc is not None and jc.get(w('val')) != 'center':
                    jc.set(w('val'), 'left')

    for table in body.iter(w('tbl')):
        for row in table.findall(w('tr')):
            rp = row.find(w('trPr'))
            if rp is None:
                rp = eng.node('trPr')
                row.insert(0, rp)
            if rp.find(w('cantSplit')) is None:
                rp.append(eng.node('cantSplit'))
        grid = table.find(w('tblGrid'))
        cols = list(grid) if grid is not None else []
        total = sum(int(c.get(w('w'), 0)) for c in cols)
        if total > 9071:
            ratio = 9071 / total
            widths = [round(int(c.get(w('w'))) * ratio) for c in cols]
            widths[-1] += 9071 - sum(widths)
            for col, width in zip(cols, widths):
                col.set(w('w'), str(width))
            tw = table.find(w('tblPr') + '/' + w('tblW'))
            if tw is not None:
                tw.set(w('w'), '9071')
            for cw in table.iter(w('tcW')):
                if cw.get(w('type')) == 'dxa':
                    cw.set(w('w'), str(round(int(cw.get(w('w'))) * ratio)))

    approval = False
    signature_paras = []
    for element in body:
        text = ''.join(t.text or '' for t in element.iter(w('t')))
        if text.startswith('V. NHẬN XÉT'):
            approval = True
        if approval:
            signature_paras.extend(element.iter(w('p')))
    for para in signature_paras[:-1]:
        pp = para.find(w('pPr'))
        if pp.find(w('keepNext')) is None:
            pp.append(eng.node('keepNext'))
    if signature_paras:
        pp = signature_paras[-1].find(w('pPr'))
        for n in list(pp.findall(w('keepNext'))):
            pp.remove(n)

    eng.save_docx(body, out_path, used_figures=used_figures)
    with zipfile.ZipFile(out_path) as archive:
        package = {name: archive.read(name) for name in archive.namelist()}
    ns = {'w': eng.W, 'a': eng.A}
    for name, data in list(package.items()):
        if not (name.startswith('word/') and name.endswith('.xml')):
            continue
        root = LX.fromstring(data)
        for rp in root.xpath('//w:rPr', namespaces=ns):
            fonts = rp.find(w('rFonts'))
            if fonts is None:
                fonts = LX.Element(w('rFonts'))
                rp.insert(0, fonts)
            fonts.attrib.clear()
            for script in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                fonts.set(w(script), 'Times New Roman')
            for tag in ('spacing', 'w', 'fitText', 'position'):
                for node in list(rp.findall(w(tag))):
                    rp.remove(node)
        for fonts in root.xpath('//w:rFonts', namespaces=ns):
            fonts.attrib.clear()
            for script in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                fonts.set(w(script), 'Times New Roman')
        for node in root.xpath('//a:latin | //a:ea | //a:cs | //a:font', namespaces=ns):
            node.set('typeface', 'Times New Roman')
        if name == 'word/fontTable.xml':
            for node in list(root):
                if node.get(w('name')) != 'Times New Roman':
                    root.remove(node)
        if name == 'word/settings.xml':
            compat = root.find(w('compat'))
            if compat is None:
                compat = LX.SubElement(root, w('compat'))
            if compat.find(w('doNotExpandShiftReturn')) is None:
                LX.SubElement(compat, w('doNotExpandShiftReturn'))
        package[name] = LX.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in package.items():
            archive.writestr(name, data)
    validate_de_cuong(out_path, used_figures)


def validate_de_cuong(out_path, used_figures):
    """Kiểm tra XML: đủ 4 hoạt động, không chia bước, không bảng kịch bản và lỗi font."""
    ns = {'w': eng.W, 'a': eng.A, 'r': eng.R}
    with zipfile.ZipFile(out_path) as archive:
        root = LX.fromstring(archive.read('word/document.xml'))
        text = ' '.join(root.xpath('//w:t/text()', namespaces=ns))
        forbidden = re.search(r'HOẠT ĐỘNG / THỜI LƯỢNG|HOẠT ĐỘNG CỦA GV VÀ HS|Tổ chức thực hiện|Bước [1-4]|[`*\u200b\ufffd]', text, re.I)
        if forbidden:
            raise ValueError(f'{out_path.name}: nội dung không phù hợp: {forbidden.group()}')
        for req in ('Hoạt động 1', 'Hoạt động 2', 'Hoạt động 3', 'Hoạt động 4',
                    'd) Cách thức tiến hành',
                    GVGD_NAME, GVHD_NAME, SV_NAME, '24S1020030'):
            if req not in text:
                raise ValueError(f'{out_path.name}: thiếu {req}')
        for name in archive.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                part = LX.fromstring(archive.read(name))
                for fonts in part.xpath('//w:rFonts', namespaces=ns):
                    if set(fonts.attrib.values()) != {'Times New Roman'}:
                        raise ValueError(f'{out_path.name}: font chưa chuẩn trong {name}')
        for figure in used_figures:
            if 'word/media/' + figure not in archive.namelist():
                raise ValueError(f'{out_path.name}: thiếu ảnh {figure}')
        print(f'XML OK: {out_path.name}; {len(used_figures)} hình SGK')


# ==============================================================================
# 1. DE CUONG LOP 12/4 - BAI 9: TAO DANH SACH, BANG (TIET PPCT: 19, 20)
# ==============================================================================
def build_12_4_bai9():
    b = eng.init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ',
        cover_subtitle='BÀI 9: TẠO DANH SÁCH, BẢNG (TIẾT PPCT: 19, 20)',
        class_str='Tin học 12 – Lớp 12/4 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Năm, ngày 05/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.II.4',
        lesson_row_str='Bài 9: Tạo danh sách, bảng (SGK trang 52–56 – Tiết PPCT: 19, 20)',
        page1_title='TÊN BÀI DẠY: BÀI 9. TẠO DANH SÁCH, BẢNG',
        page1_subtitle='(Đề cương dự giờ kiến tập – Thời lượng bài học theo PPCT: 02 tiết [Tiết PPCT: 19, 20] – Dự giờ tại Lớp 12/4)',
        page1_class_line='Môn học: Tin học; Lớp: 12/4',
        page1_time_line='Thời lượng bài học theo PPCT: 02 tiết (Tiết 19: Lí thuyết – dự giờ; Tiết 20: Thực hành)',
        page1_sched_line='Lịch dự giờ: Thứ Năm, 05/11/2026 – Tiết 3 (09h00–09h45) – Phòng H.II.4',
        page1_mode_line='Hình thức: Tóm tắt mục tiêu, nội dung trọng tâm, phương pháp tổ chức dạy học và đánh giá',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- **[KT1]:** Trình bày được cú pháp và chức năng các thẻ HTML tạo danh sách có thứ tự (<ol>, <li>), danh sách không có thứ tự (<ul>, <li>), danh sách mô tả (<dl>, <dt>, <dd>) và danh sách lồng nhau; giải thích được ý nghĩa các thuộc tính type, start.', first=0)
    eng.add(b, '- **[KT2]:** Trình bày được cấu trúc cơ bản của bảng trong HTML gồm các thẻ <table>, <tr>, <th>, <td>, <caption>; vận dụng thuộc tính style để định dạng kích thước, đường viền và các thuộc tính colspan, rowspan để gộp ô.', first=0)
    eng.add(b, '- **[KT3]:** Vận dụng phối hợp thẻ tạo danh sách và thẻ tạo bảng để xây dựng trang web hoàn chỉnh giới thiệu Câu lạc bộ và Lịch hoạt động câu lạc bộ đúng yêu cầu Sách giáo khoa Tin học 12 trang 55–56.', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'label', italic=True)
    eng.add(b, '- **[NLC1 – Tự chủ và tự học]:** Chủ động đọc Sách giáo khoa, phân tích ví dụ mã nguồn HTML và tự kiểm chứng kết quả hiển thị trên trình duyệt web.', first=0)
    eng.add(b, '- **[NLC2 – Giao tiếp và hợp tác]:** Thảo luận cặp đôi hoặc nhóm nhỏ để nhận diện cấu trúc thẻ danh sách, bảng biểu và hỗ trợ nhau sửa lỗi cú pháp.', first=0)
    eng.add(b, '- **[NLC3 – Giải quyết vấn đề và sáng tạo]:** Lựa chọn cấu trúc danh sách hoặc bảng hợp lí, thiết kế bảng có gộp ô (rowspan, colspan) rõ ràng, thẩm mĩ cho trang web thực tế.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'label', italic=True)
    eng.add(b, '- **[NLa]:** Sử dụng và quản lí hiệu quả trình soạn thảo mã nguồn HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web.', first=0)
    eng.add(b, '- **[NLc]:** Phân tích cấu trúc dữ liệu dạng danh sách, bảng biểu và viết đúng mã HTML chuẩn cấu trúc lồng thẻ.', first=0)
    eng.add(b, '- **[NLd]:** Khai thác ví dụ trong học liệu số để tự học, kiểm tra và hoàn thiện sản phẩm HTML.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'label', italic=True)
    eng.add(b, '- **[NLS 3.1.NC1.a]:** Tạo lập, chỉnh sửa và định dạng nội dung số dạng danh sách và bảng trên trang web bằng HTML.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- **[PC1 – Chăm chỉ]:** Kiên trì thực hành viết mã HTML, cẩn trọng kiểm tra cặp thẻ đóng – mở và các thuộc tính bảng.', first=0)
    eng.add(b, '- **[PC2 – Trách nhiệm]:** Chấp hành nghiêm túc nội quy lớp học, bảo quản thiết bị và hoàn thành sản phẩm thực hành đúng hạn.', first=0)
    eng.add(b, '- **[PC3 – Trung thực]:** Tự tay biên tập mã nguồn HTML và tự đánh giá trung thực mức độ hoàn thành sản phẩm.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Hệ thống máy tính giáo viên, máy chiếu/màn hình lớn tại phòng học H.II.4 (Lớp 12/4) cài đặt trình soạn thảo HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web (Chrome / Edge).', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), NXB Giáo dục Việt Nam, Bài 9: Tạo danh sách, bảng (trang 52–56).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 19: Bài 4.9 – Tạo danh sách, bảng, tiết 1 – LT; Tiết PPCT 20: tiết 2 – TH).', first=0)
    eng.add(b, '- Hình ảnh minh họa SGK (Hình 9.1 đến Hình 9.8), Phiếu học tập và tệp HTML mẫu (CLB.html).', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM VÀ PHƯƠNG PHÁP TỔ CHỨC DẠY HỌC', 'heading')

    eng.add(b, '1. Hoạt động 1: Mở đầu (Khởi động – Khoảng 5 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Tạo hứng thú, giúp học sinh nhận thấy nhu cầu tổ chức thông tin trên trang web dưới dạng danh sách và bảng biểu thay cho đoạn văn xuôi.', first=0)
    eng.add(b, '**b) Nội dung:** Quan sát hai cách trình bày thông tin Câu lạc bộ và Lịch hoạt động CLB (một đoạn văn bản thuần tuý và một trang web có danh sách, bảng biểu rõ ràng) và trả lời câu hỏi gợi mở.', first=0)
    eng.add(b, '**c) Sản phẩm:** Nhận định của học sinh về tính trực quan, dễ tra cứu của danh sách và bảng biểu; xác định được mục tiêu Bài 9.', first=0)
    add_steps(
        b,
        'Giáo viên chiếu tình huống so sánh hai cách trình bày lịch CLB và nêu câu hỏi gợi mở.',
        'Học sinh quan sát, so sánh ưu nhược điểm của từng cách trình bày.',
        'Đại diện 1–2 học sinh nêu nhận xét về sự cần thiết của danh sách và bảng trên trang web.',
        'Giáo viên nhận xét, chốt vấn đề và dẫn dắt vào Bài 9: Tạo danh sách, bảng.'
    )

    eng.add(b, '2. Hoạt động 2: Hình thành kiến thức mới (Bám sát SGK trang 52–56)', 'sub')
    eng.add(b, '2.1. Nội dung 1: Tạo danh sách trong HTML (Mục 1 – SGK trang 52–53)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Học sinh viết và giải thích được 3 loại danh sách trong HTML (<ol>, <ul>, <dl>) và danh sách lồng nhau.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm kiến thức trọng tâm:**', first=0)
    eng.add(b, '- **Danh sách có thứ tự:** Dùng cặp thẻ <ol>...</ol> bao ngoài và <li>...</li> cho từng mục; thuộc tính type ("1", "A", "a", "I", "i") chọn kiểu đánh số, thuộc tính start chọn giá trị bắt đầu.', first=0)
    eng.add(b, '- **Danh sách không có thứ tự:** Dùng <ul>...</ul> và <li>...</li>; định dạng kí hiệu đầu dòng bằng style với list-style-type nhận giá trị disc, circle, square hoặc none (ví dụ: <ul style="list-style-type:square">).', first=0)
    eng.add(b, '- **Danh sách mô tả:** Dùng <dl>...</dl>, trong đó <dt>...</dt> chứa thuật ngữ và <dd>...</dd> chứa mô tả.', first=0)
    eng.add(b, '- **Danh sách lồng nhau & Củng cố trang 53:** Đặt thẻ <ol> hoặc <ul> bên trong thẻ <li> của danh sách cha (đáp án câu hỏi tr. 53: đặt <ol type="a"> bên trong thẻ <li> của <ol type="1">).', first=0)
    add_figure_sgk(b, 'hinh_9_1.png')
    add_figure_sgk(b, 'hinh_9_2.png')
    add_figure_sgk(b, 'hinh_9_3.png')
    add_steps(
        b,
        'Giáo viên giao nhiệm vụ nghiên cứu Mục 1 SGK và các ví dụ Hình 9.1–9.3.',
        'Học sinh đọc SGK, thảo luận cặp đôi phân tích cú pháp <ol>, <ul>, <dl> và câu hỏi củng cố tr. 53.',
        'Đại diện cặp đôi trình bày cấu trúc từng loại danh sách và đáp án danh sách lồng nhau.',
        'Giáo viên chuẩn hoá cú pháp, lưu ý thuộc tính style cho <ul> và quy tắc lồng thẻ vào bên trong <li>.'
    )

    eng.add(b, '2.2. Nội dung 2: Thiết lập bảng trong HTML (Mục 2 – SGK trang 53–55)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Trình bày được cấu trúc bảng cơ bản, cách điều chỉnh kích thước, đường viền và gộp ô bằng colspan, rowspan.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm kiến thức trọng tâm:**', first=0)
    eng.add(b, '- **Cấu trúc bảng cơ bản:** Dùng <table>...</table>, hàng <tr>...</tr>, ô tiêu đề <th>...</th>, ô dữ liệu <td>...</td>, tiêu đề bảng <caption>...</caption>.', first=0)
    eng.add(b, '- **Định dạng kích thước và viền bảng:** Dùng thuộc tính style với width, height và border (ví dụ style="border:1px solid") cho <table> và các ô.', first=0)
    eng.add(b, '- **Gộp ô trong bảng & Củng cố trang 55:** Dùng colspan="n" để gộp ngang n ô, rowspan="m" để gộp dọc m ô và bỏ bớt thẻ <td>/<th> bị gộp (câu hỏi tr. 55: bảng Hình 9.6 chưa phân biệt rõ viền các ô và độ rộng cột; khắc phục bằng thuộc tính style với border và width).', first=0)
    add_figure_sgk(b, 'hinh_9_4.png')
    add_figure_sgk(b, 'hinh_9_5.png')
    add_figure_sgk(b, 'hinh_9_6.png')
    add_steps(
        b,
        'Giáo viên chiếu Hình 9.4–9.6, yêu cầu học sinh phân tích cấu trúc hàng, ô và cách gộp ô.',
        'Học sinh nghiên cứu Mục 2 SGK, đếm số thẻ <td>/<th> trên từng hàng khi sử dụng rowspan, colspan.',
        'Học sinh trả lời câu hỏi về vai trò <th>, colspan, rowspan và nhược điểm Hình 9.6.',
        'Giáo viên chốt kiến thức cấu trúc bảng và nhấn mạnh quy tắc lược bỏ ô bị gộp để không lệch cột.'
    )

    eng.add(b, '2.3. Nội dung 3: Thực hành tạo danh sách và bảng (Mục 3 – SGK trang 55–56)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Xây dựng được trang web CLB.html có danh sách câu lạc bộ lồng nhau và bảng lịch hoạt động có gộp ô.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm:**', first=0)
    eng.add(b, '- **Nhiệm vụ 1 (Hình 9.7):** Tạo danh sách lồng nhau gồm Thể thao (Bóng đá, Bóng chuyền, Bóng rổ, Võ [Karatedo, Taekwondo, Vovinam], Bơi) và Nghệ thuật (Mĩ thuật, Nhiếp ảnh, Âm nhạc [Thanh nhạc, Piano, Violin], Khiêu vũ).', first=0)
    eng.add(b, '- **Nhiệm vụ 2 (Hình 9.8):** Tạo bảng Lịch hoạt động CLB Thể thao lưu vào tệp CLB.html có tiêu đề, hàng tiêu đề gộp ô cho các buổi và các môn thể thao từ Thứ 2 đến Thứ 6.', first=0)
    add_figure_sgk(b, 'hinh_9_7.png')
    add_figure_sgk(b, 'hinh_9_8.png')
    add_steps(
        b,
        'Giáo viên hướng dẫn mẫu bố cục Hình 9.7–9.8 và giao hai nhiệm vụ thực hành trong SGK.',
        'Học sinh thực hành viết mã HTML trên máy (hoặc phác thảo khung thẻ trên phiếu nếu dạy lí thuyết tại lớp).',
        'Một số học sinh trình chiếu mã nguồn và kết quả hiển thị trang CLB.html.',
        'Giáo viên kiểm tra cấu trúc lồng thẻ, cách gộp ô colspan/rowspan và tổng kết sản phẩm.'
    )

    eng.add(b, '3. Hoạt động 3: Luyện tập (SGK trang 56)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Củng cố kĩ năng dùng thuộc tính style để tạo viền bảng và định dạng màu chữ trong HTML.', first=0)
    eng.add(b, '**b) Nội dung:** Thực hiện bài Luyện tập SGK trang 56: sửa mã nguồn Hình 9.6a dùng style thay border="1", đặt viền màu xanh cho hai hàng đầu và tô màu đỏ, vàng, xanh cho chữ Toán, Vật lí, Hoá học.', first=0)
    eng.add(b, '**c) Sản phẩm:** Đoạn mã HTML sử dụng style="border:1px solid blue" cho hai hàng đầu và style="color:red/gold/green" cho từng ô tiêu đề môn học.', first=0)
    add_steps(
        b,
        'Giáo viên nêu yêu cầu bài Luyện tập trang 56 và gợi ý tên màu cơ bản trong CSS inline.',
        'Học sinh viết lại đoạn mã định dạng bằng style theo cặp đôi.',
        'Đại diện học sinh đọc hoặc trình bày đoạn mã đã chỉnh sửa; lớp nhận xét.',
        'Giáo viên nhận xét, chốt đáp án chuẩn và nhắc ưu tiên dùng style thay thuộc tính border cũ.'
    )

    eng.add(b, '4. Hoạt động 4: Vận dụng (SGK trang 56)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Vận dụng kiến thức tạo bảng HTML kết hợp lập trình Python sinh tệp bảng điểm tự động.', first=0)
    eng.add(b, '**b) Nội dung:** Thực hiện bài Vận dụng SGK trang 56: với bảng dữ liệu n × 4 (Họ tên và điểm Toán, Vật lí, Hoá học), viết chương trình Python tạo tệp HTML có bảng tương tự Hình 9.6b.', first=0)
    eng.add(b, '**c) Sản phẩm:** Tệp kịch bản Python sinh mã HTML tạo bảng điểm hoàn chỉnh (học sinh hoàn thiện và nộp ở giờ học sau).', first=0)
    add_steps(
        b,
        'Giáo viên giao nhiệm vụ Vận dụng trang 56 và gợi ý cấu trúc vòng lặp ghép chuỗi thẻ <tr><td>...</td></tr>.',
        'Học sinh tiếp nhận yêu cầu, phác thảo ý tưởng thuật toán sinh mã HTML từ danh sách dữ liệu.',
        '1–2 học sinh nêu hướng xử lí phần tiêu đề bảng gộp ô và các hàng dữ liệu động.',
        'Giáo viên dặn dò học sinh hoàn thiện bài Vận dụng tại nhà và chuẩn bị Bài 10.'
    )

    eng.add(b, 'IV. HỆ THỐNG CÂU HỎI CỦNG CỐ VÀ TIÊU CHÍ ĐÁNH GIÁ', 'heading')
    eng.add(b, '- **Câu hỏi củng cố 1 (SGK tr. 53):** Tạo danh sách lồng nhau có cấp 1 đánh số 1, 2, 3 và cấp 2 đánh số a, b, c như thế nào?', first=0)
    eng.add(b, '- **Câu hỏi củng cố 2 (SGK tr. 55):** Bảng trong Hình 9.6 có nhược điểm gì và cách khắc phục?', first=0)
    eng.add(b, '- **Tiêu chí đánh giá sản phẩm thực hành (Thang điểm 10):** Đúng cấu trúc HTML và lồng thẻ (3,0 điểm); Tạo đúng danh sách lồng nhau theo Hình 9.7 (2,5 điểm); Tạo đúng bảng lịch hoạt động có gộp ô chuẩn theo Hình 9.8 (3,0 điểm); Trình bày mã nguồn sạch đẹp, lưu tệp đúng quy định (1,5 điểm).', first=0)

    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 05 tháng 11 năm 2026',
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
    )

    out_path = DE_CUONG_DIR / 'de_cuong_du_gio_12-4.docx'
    save_de_cuong(b, out_path, used_figures=list(BAI9_FIGS.keys()))


# ==============================================================================
# 2. DE CUONG LOP 12/2 - BAI 10 (TIET 1 - LI THUYET - TIET PPCT: 21)
# ==============================================================================
def build_12_2_bai10_tiet1():
    b = eng.init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ',
        cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT – TIẾT PPCT: 21)',
        class_str='Tin học 12 – Lớp 12/2 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.II.2',
        lesson_row_str='Bài 10: Tạo liên kết (Tiết 1 – Lí thuyết – Tiết PPCT: 21)',
        page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT)',
        page1_subtitle='(Đề cương dự giờ kiến tập – Thời lượng toàn bài theo PPCT: 02 tiết [Tiết PPCT: 21, 22] – Trọng tâm Tiết 1: Lí thuyết)',
        page1_class_line='Môn học: Tin học; Lớp: 12/2',
        page1_time_line='Thời lượng: 01 tiết (45 phút – Tiết 1 - Lí thuyết / Tiết PPCT: 21)',
        page1_sched_line='Lịch dự giờ: Thứ Sáu, 13/11/2026 – Tiết 3 (09h00–09h45) – Phòng H.II.2',
        page1_mode_line='Hình thức: Tóm tắt mục tiêu, nội dung trọng tâm, phương pháp tổ chức dạy học và đánh giá',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- **[KT1]:** Trình bày được khái niệm siêu văn bản (hypertext), siêu liên kết (hyperlink), cú pháp thẻ <a> với thuộc tính href và phân biệt rõ đường dẫn tuyệt đối với đường dẫn tương đối.', first=0)
    eng.add(b, '- **[KT2]:** Phân biệt được liên kết ngoài (tới trang web khác trên Internet) và liên kết trong (tới trang web khác trong cùng website); viết đúng đường dẫn tương đối cho 4 trường hợp cấu trúc thư mục theo Hình 10.1 (cùng thư mục, dưới 1 cấp, dưới nhiều cấp và ngược lên thư mục mức trên bằng ../).', first=0)
    eng.add(b, '- **[KT3]:** Trình bày được cách tạo liên kết tới một vị trí khác trong cùng trang web bằng thuộc tính id và cú pháp #mã_định_danh_đoạn, cũng như cách tạo liên kết tới tệp hình ảnh.', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'label', italic=True)
    eng.add(b, '- **[NLC1 – Tự chủ và tự học]:** Học sinh tự đọc Sách giáo khoa Tin học 12 (trang 57–60), phân tích sơ đồ cây thư mục Hình 10.1 để tự rút ra quy tắc xác định đường dẫn tương đối.', first=0)
    eng.add(b, '- **[NLC2 – Giao tiếp và hợp tác]:** Thảo luận cặp đôi, trao đổi kết quả phân tích đường dẫn trên Phiếu học tập và phản biện lẫn nhau.', first=0)
    eng.add(b, '- **[NLC3 – Giải quyết vấn đề và sáng tạo]:** Giải quyết chính xác bài toán định vị đường dẫn khi chuyển đổi cấu trúc thư mục website và thiết kế trước các mã định danh id phục vụ Tiết 2 Thực hành.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'label', italic=True)
    eng.add(b, '- **[NLb]:** Giải thích được bản chất địa chỉ URL, đường dẫn tuyệt đối và đường dẫn tương đối trong môi trường web.', first=0)
    eng.add(b, '- **[NLc]:** Viết đúng cú pháp thẻ <a> liên kết ngoài, liên kết trong cùng website, liên kết nội bộ trong trang (#id) và liên kết hình ảnh.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'label', italic=True)
    eng.add(b, '- **[NLS 3.1.NC1.a]:** Nắm vững kĩ thuật thiết kế và liên kết tài nguyên siêu văn bản trên môi trường số chuẩn xác.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- **[PC1 – Chăm chỉ]:** Tích cực nghiên cứu Sách giáo khoa và hoàn thành đầy đủ các câu hỏi củng cố lí thuyết.', first=0)
    eng.add(b, '- **[PC2 – Trách nhiệm]:** Chuẩn bị kĩ phương án thiết kế liên kết cho tệp CLB.html để phục vụ giờ thực hành Tiết 2.', first=0)
    eng.add(b, '- **[PC3 – Trung thực]:** Tự kiểm tra, đối chiếu kết quả bài làm với đáp án chuẩn.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Máy tính giáo viên kết nối máy chiếu/màn hình tại phòng học H.II.2 (Lớp 12/2), bảng viết, phấn.', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), NXB Giáo dục Việt Nam, Bài 10: Tạo liên kết (Mục 1 và Mục 2, trang 57–60).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 21: Bài 4.10: Tạo liên kết – Tiết 1 - LT).', first=0)
    eng.add(b, '- Hình ảnh Hình 10.1 cắt từ SGK trang 59, Phiếu học tập và tệp minh họa CLB.html.', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM VÀ PHƯƠNG PHÁP TỔ CHỨC DẠY HỌC', 'heading')

    eng.add(b, '1. Hoạt động 1: Mở đầu (Tình huống mở đầu & Hoạt động 1 – SGK trang 57, 5 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Giúp học sinh nhận biết sự khác biệt giữa trang web (siêu văn bản) với tệp văn bản thuần tuý (.txt) và hình thành nhu cầu tạo liên kết.', first=0)
    eng.add(b, '**b) Nội dung:** So sánh một trang web giới thiệu hoạt động CLB với một tệp văn bản thuần tuý (.txt) và trả lời câu hỏi Hoạt động 1 (SGK tr. 57).', first=0)
    eng.add(b, '**c) Sản phẩm:** Trang web là siêu văn bản vì ngoài văn bản còn chứa hình ảnh, âm thanh, video và các siêu liên kết cho phép di chuyển phi tuyến tính giữa các nội dung.', first=0)
    add_steps(
        b,
        'Giáo viên chiếu trang web mẫu và đặt câu hỏi so sánh trang web với tệp văn bản thuần tuý (.txt).',
        'Học sinh quan sát và trao đổi nhanh theo cặp để tìm điểm khác biệt.',
        '1–2 học sinh nêu câu trả lời về tính đa phương tiện và khả năng chuyển trang qua liên kết.',
        'Giáo viên nhận xét, chốt khái niệm mở đầu và dẫn vào Bài 10: Tạo liên kết.'
    )

    eng.add(b, '2. Hoạt động 2: Hình thành kiến thức mới (Bám sát SGK trang 57–60)', 'sub')
    eng.add(b, '2.1. Nội dung 1: Siêu văn bản và đường dẫn (Mục 1 – SGK trang 57–58)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Nắm vững khái niệm siêu liên kết, cú pháp thẻ <a> và phân biệt đường dẫn tuyệt đối với đường dẫn tương đối.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm kiến thức trọng tâm:**', first=0)
    eng.add(b, '- **Siêu liên kết và cú pháp thẻ <a>:** Siêu liên kết (hyperlink) là tham chiếu đến tài liệu ở vị trí khác; cú pháp: <a href="URL">Nội dung hiển thị</a> (mặc định chữ xanh, gạch chân).', first=0)
    eng.add(b, '- **Đường dẫn tuyệt đối:** Ghi đầy đủ giao thức (http:// hoặc https://), tên miền và đường dẫn chi tiết (ví dụ: https://google.com).', first=0)
    eng.add(b, '- **Đường dẫn tương đối & Củng cố tr. 58:** Không có giao thức, tên miền; xác định tính từ vị trí tệp hiện tại (đáp án tr. 58: a) tương đối; b) tuyệt đối; c) tương đối theo cú pháp HTML do thiếu giao thức).', first=0)
    add_steps(
        b,
        'Giáo viên giao nhiệm vụ đọc Mục 1 SGK và phân loại 3 đường dẫn ở câu hỏi củng cố trang 58.',
        'Học sinh đọc SGK, thảo luận cặp đôi phân tích cú pháp thẻ <a> và hai loại đường dẫn.',
        'Đại diện nhóm báo cáo cú pháp thẻ <a> và giải thích kết quả phân loại 3 ví dụ a, b, c.',
        'Giáo viên chuẩn hoá kiến thức và nhấn mạnh dấu hiệu nhận biết đường dẫn tuyệt đối (có giao thức http/https).'
    )

    eng.add(b, '2.2. Nội dung 2: Các cách tạo liên kết tới một trang web (Mục 2 – SGK trang 58–60)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Viết đúng câu lệnh liên kết ngoài, liên kết trong website (4 trường hợp Hình 10.1), liên kết nội bộ (#id) và liên kết hình ảnh.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm kiến thức trọng tâm:**', first=0)
    eng.add(b, '- **Liên kết ngoài:** Dùng đường dẫn tuyệt đối, ví dụ <a href="https://hanhtrangso.nxbgd.vn/sach-dien-tu?book_active=0">Sách điện tử KNTT</a>.', first=0)
    eng.add(b, '- **Liên kết trong cùng website (4 trường hợp theo Hình 10.1, tr. 59–60):** (1) Cùng thư mục: href="thong_tin.html"; (2) Dưới 1 cấp: href="bai_tap/bai_tap_1.html"; (3) Dưới 2 cấp: href="bai_tap/on_tap/bai_tap_on_tap.html"; (4) Ngược lên mức trên bằng "../": href="../index.html".', first=0)
    add_figure_sgk(b, 'hinh_10_1.png')
    eng.add(b, '- **Liên kết tới vị trí trong cùng trang (#id):** Gắn id="tên_định_danh" cho phần tử đích (ví dụ <h2 id="CLBTT">Lịch hoạt động</h2>) và đặt liên kết <a href="#CLBTT">Xem lịch hoạt động</a>.', first=0)
    eng.add(b, '- **Liên kết tới hình ảnh & Củng cố tr. 60:** Đặt thẻ <img> trong cặp thẻ <a>...</a> (ví dụ <a href="thong_tin.html"><img src="images/sun.png"></a>); câu hỏi tr. 60: từ bai_tap_on_tap.html về index.html dùng <a href="../../index.html">Về trang chủ</a>.', first=0)
    add_steps(
        b,
        'Giáo viên chiếu sơ đồ cây thư mục Hình 10.1, giao nhiệm vụ tìm đường dẫn cho 4 trường hợp và cách tạo liên kết #id.',
        'Học sinh phân tích Hình 10.1 theo nhóm đôi, xác định số lần lùi thư mục "../" và cú pháp gắn id.',
        'Học sinh lên bảng viết câu lệnh <a> cho các trường hợp và câu hỏi củng cố trang 60.',
        'Giáo viên nhận xét, sửa lỗi cú pháp thường gặp (thiếu "../", trùng id, quên dấu "#") và chốt kiến thức.'
    )

    eng.add(b, '3. Hoạt động 3: Luyện tập (Củng cố lí thuyết – 10 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Vận dụng viết thành thạo câu lệnh tạo liên kết trong website và liên kết nội bộ trên giấy/phiếu học tập.', first=0)
    eng.add(b, '**b) Nội dung:** Học sinh hoàn thành nhanh bài tập viết thẻ <a> liên kết từ bai_tap_1.html tới sun.png trong thư mục images, liên kết quay về index.html và liên kết nội bộ tới đoạn có id="gioi_thieu".', first=0)
    eng.add(b, '**c) Sản phẩm:** Các câu lệnh HTML chính xác: <a href="../images/sun.png">...</a>, <a href="../index.html">...</a>, <a href="#gioi_thieu">...</a>.', first=0)
    add_steps(
        b,
        'Giáo viên giao đề bài luyện tập trên màn chiếu (3 câu hỏi viết nhanh câu lệnh thẻ <a>).',
        'Học sinh tự thực hiện độc lập trong 4 phút rồi đối chiếu chéo với bạn cùng bàn.',
        '2–3 học sinh trình bày kết quả trên bảng; cả lớp nhận xét, đối chiếu đáp án.',
        'Giáo viên chốt đáp án chuẩn và đánh giá mức độ nắm vững cú pháp của cả lớp.'
    )

    eng.add(b, '4. Hoạt động 4: Vận dụng (Chuẩn bị cho Tiết 2 Thực hành – 5 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Định hướng cấu trúc liên kết và danh sách mã định danh id cho bài thực hành Tiết 2.', first=0)
    eng.add(b, '**b) Nội dung:** Phác thảo sơ đồ liên kết hai chiều giữa tệp CLB.html và thong_tin.html (cùng thư mục) và dự kiến các giá trị id (#CLBTT, #bongda, #bongchuyen, #bongro, #vo, #boi) theo yêu cầu SGK trang 61.', first=0)
    eng.add(b, '**c) Sản phẩm:** Bản phác thảo cấu trúc thư mục và các cặp id – href chuẩn bị sẵn sàng cho tiết thực hành trên máy.', first=0)
    add_steps(
        b,
        'Giáo viên giới thiệu mục tiêu Tiết 2 Thực hành và giao nhiệm vụ phác thảo sơ đồ liên kết cho bộ tệp CLB.',
        'Học sinh làm việc cá nhân, ghi chú các đường dẫn và mã id dự kiến vào vở.',
        '1–2 học sinh nêu phương án tổ chức tệp và liên kết hai chiều giữa CLB.html và thong_tin.html.',
        'Giáo viên nhận xét, dặn dò học sinh chuẩn bị tệp CLB.html từ Bài 9 để thực hành ở Tiết 2.'
    )

    eng.add(b, 'IV. HỆ THỐNG CÂU HỎI CỦNG CỐ VÀ TIÊU CHÍ ĐÁNH GIÁ', 'heading')
    eng.add(b, '- **Câu hỏi củng cố 1 (SGK tr. 58):** Phân biệt đường dẫn tuyệt đối và tương đối qua 3 đường dẫn mẫu.', first=0)
    eng.add(b, '- **Câu hỏi củng cố 2 (SGK tr. 60):** Từ bai_tap_on_tap.html (trong bai_tap/on_tap/) về index.html dùng href="../../index.html".', first=0)
    eng.add(b, '- **Định hướng đánh giá Tiết 1:** Đánh giá thường xuyên qua câu trả lời, phiếu học tập (chính xác cú pháp thẻ <a>, 4 trường hợp đường dẫn Hình 10.1 và liên kết nội bộ #id).', first=0)

    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 13 tháng 11 năm 2026',
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
    )

    out_path = DE_CUONG_DIR / 'de_cuong_du_gio_12-2.docx'
    save_de_cuong(b, out_path, used_figures=['hinh_10_1.png'])


# ==============================================================================
# 3. DE CUONG LOP 12/1 - BAI 10 (TIET 2 - THUC HANH - TIET PPCT: 22)
# ==============================================================================
def build_12_1_bai10_tiet2():
    b = eng.init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ',
        cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TIẾT 2 – THỰC HÀNH – TIẾT PPCT: 22)',
        class_str='Tin học 12 – Lớp 12/1 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 2 (08h00 – 08h45) – Phòng H.II.1',
        lesson_row_str='Bài 10: Tạo liên kết (Tiết 2 – Thực hành – Tiết PPCT: 22)',
        page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT (TIẾT 2 – THỰC HÀNH)',
        page1_subtitle='(Đề cương dự giờ kiến tập – Thời lượng toàn bài theo PPCT: 02 tiết [Tiết PPCT: 21, 22] – Trọng tâm Tiết 2: Thực hành)',
        page1_class_line='Môn học: Tin học; Lớp: 12/1',
        page1_time_line='Thời lượng: 01 tiết (45 phút – Tiết 2 - Thực hành / Tiết PPCT: 22)',
        page1_sched_line='Lịch dự giờ: Thứ Sáu, 13/11/2026 – Tiết 2 (08h00–08h45) – Phòng H.II.1',
        page1_mode_line='Hình thức: Tóm tắt mục tiêu, nội dung thực hành, phương pháp tổ chức và Rubric đánh giá',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- **[KT1]:** Vận dụng thành thạo kĩ thuật gắn thuộc tính id và đặt liên kết nội bộ trong cùng trang web (#CLBTT) trên tệp mã nguồn CLB.html theo đúng yêu cầu Nhiệm vụ 1 (SGK trang 61).', first=0)
    eng.add(b, '- **[KT2]:** Tạo lập được tệp trang web mới thong_tin.html nằm cùng thư mục với CLB.html và thiết lập liên kết hai chiều giữa hai trang web bằng đường dẫn tương đối theo Nhiệm vụ 2 và Bài Luyện tập 1 (SGK trang 61).', first=0)
    eng.add(b, '- **[KT3]:** Hoàn thiện trọn vẹn trang web đa liên kết cho Bài Luyện tập 2 (gioi_thieu_ban_than.html) và phần Vận dụng (gắn id riêng cho 5 đoạn giới thiệu câu lạc bộ dưới bảng trong tệp CLB.html và tạo liên kết từ tên từng CLB trong bảng cuộn xuống đoạn giới thiệu tương ứng).', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'label', italic=True)
    eng.add(b, '- **[NLC1 – Tự chủ và tự học]:** Học sinh tự đọc hướng dẫn thực hành SGK trang 61, viết mã và tự kiểm thử hoạt động của liên kết trên trình duyệt web.', first=0)
    eng.add(b, '- **[NLC2 – Giao tiếp và hợp tác]:** Trao đổi, hỗ trợ bạn cùng máy phát hiện và khắc phục lỗi sai đường dẫn hoặc đặt tên id không khớp.', first=0)
    eng.add(b, '- **[NLC3 – Giải quyết vấn đề và sáng tạo]:** Tổ chức bộ tệp trang web khoa học, thiết kế nội dung giới thiệu sinh động và điều hướng liên kết mượt mà.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'label', italic=True)
    eng.add(b, '- **[NLa]:** Sử dụng thành thạo trình soạn thảo mã nguồn HTML (Sublime Text / VS Code / Notepad++) và công cụ kiểm tra liên kết trên trình duyệt web.', first=0)
    eng.add(b, '- **[NLc]:** Viết mã nguồn HTML chuẩn xác cho liên kết nội bộ (#id) và liên kết tương đối giữa các tệp cùng thư mục.', first=0)
    eng.add(b, '- **[NLd]:** Lưu trữ, quản lí cấu trúc thư mục thực hành Bai10_TaoLienKet sạch sẽ, đúng quy định.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'label', italic=True)
    eng.add(b, '- **[NLS 3.1.NC1.a]:** Tạo lập, biên tập và liên kết thành công bộ trang web đa trang (.html) theo đúng yêu cầu kĩ thuật.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- **[PC1 – Chăm chỉ]:** Kiên trì biên tập đầy đủ nội dung giới thiệu, danh sách thành tích và thông tin chi tiết của từng câu lạc bộ.', first=0)
    eng.add(b, '- **[PC2 – Trách nhiệm]:** Chấp hành nghiêm túc nội quy phòng máy, lưu trữ thư mục bài làm đúng quy định.', first=0)
    eng.add(b, '- **[PC3 – Trung thực]:** Tự tay viết mã nguồn HTML và tự đánh giá trung thực sản phẩm theo Rubric.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Máy tính tại phòng học H.II.1 (Lớp 12/1) có cài đặt trình soạn thảo văn bản HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web (Chrome / Edge), máy chiếu.', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), NXB Giáo dục Việt Nam, Bài 10: Tạo liên kết (Mục 3. Thực hành tạo liên kết, Luyện tập và Vận dụng – trang 61).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 22: Bài 4.10: Tạo liên kết – Tiết 2 - TH).', first=0)
    eng.add(b, '- Thư mục học liệu chứa tệp gốc CLB.html kế thừa từ Bài 9, Hình 10.2 (SGK trang 61) và Bảng tiêu chí đánh giá thực hành (Rubric).', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM VÀ PHƯƠNG PHÁP TỔ CHỨC DẠY HỌC', 'heading')

    eng.add(b, '1. Hoạt động 1: Mở đầu (Kiểm tra nhanh cú pháp & khởi động thực hành – 5 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Hệ thống nhanh kiến thức Tiết 1 (cú pháp thẻ <a>, thuộc tính id, đường dẫn tương đối) và kiểm tra tệp CLB.html sẵn sàng trên máy.', first=0)
    eng.add(b, '**b) Nội dung:** Trả lời nhanh 2 câu hỏi ôn tập cú pháp liên kết nội bộ (#CLBTT) và mở thư mục Bai10_TaoLienKet chứa tệp CLB.html.', first=0)
    eng.add(b, '**c) Sản phẩm:** Học sinh nhắc đúng cú pháp gắn id và href="#id"; mở được tệp CLB.html trong trình soạn thảo.', first=0)
    add_steps(
        b,
        'Giáo viên kiểm tra việc chuẩn bị tệp CLB.html trên máy và nêu câu hỏi nhắc lại cú pháp thẻ <a>, thuộc tính id.',
        'Học sinh khởi động phần mềm soạn thảo, mở tệp CLB.html và nhớ lại kiến thức Tiết 1.',
        '1–2 học sinh trả lời nhanh vai trò của id="CLBTT" và href="#CLBTT".',
        'Giáo viên nhận xét, chốt kiến thức nền và chuyển sang nội dung thực hành chính.'
    )

    eng.add(b, '2. Hoạt động 2: Thực hành tạo liên kết (Mục 3 – SGK trang 61)', 'sub')
    eng.add(b, '2.1. Nhiệm vụ 1: Tạo liên kết trong cùng trang (Hình 10.2 – SGK trang 61)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Gắn được id="CLBTT" cho bảng Lịch hoạt động CLB Thể thao và đặt liên kết từ đầu trang CLB.html cuộn xuống bảng.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm:** Bổ sung phần giới thiệu hoạt động ngoại khoá vào đầu tệp CLB.html theo Hình 10.2; gắn <table id="CLBTT" ...> và đặt câu lệnh <a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a>.', first=0)
    add_figure_sgk(b, 'hinh_10_2.png')
    add_steps(
        b,
        'Giáo viên chiếu Hình 10.2, thị phạm nhanh cách gắn id="CLBTT" và đặt thẻ <a> ở đầu tệp CLB.html.',
        'Học sinh bổ sung phần giới thiệu và tạo liên kết nội bộ trên tệp CLB.html, chạy thử trên trình duyệt.',
        'Một vài học sinh báo cáo kết quả, kiểm chứng thao tác nháy liên kết cuộn xuống đúng bảng lịch.',
        'Giáo viên kiểm tra trên một số máy, lưu ý lỗi chính tả giữa id và href rồi chốt Nhiệm vụ 1.'
    )

    eng.add(b, '2.2. Nhiệm vụ 2: Tạo liên kết sang trang khác (SGK trang 61)', 'label', italic=True)
    eng.add(b, '**a) Mục tiêu:** Tạo mới tệp thong_tin.html cùng thư mục và tạo liên kết mở sang trang này từ CLB.html.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm:** Tạo tệp thong_tin.html giới thiệu chi tiết các CLB (Bóng đá, Bóng chuyền, Bóng rổ, Âm nhạc, Mĩ thuật...), bổ sung <p><a href="thong_tin.html">Xem thông tin chi tiết các câu lạc bộ</a></p> trong CLB.html.', first=0)
    add_steps(
        b,
        'Giáo viên giao Nhiệm vụ 2 và nhấn mạnh yêu cầu lưu tệp thong_tin.html trong cùng thư mục với CLB.html.',
        'Học sinh soạn thảo tệp thong_tin.html, thêm thẻ liên kết trong CLB.html và kiểm thử chuyển trang.',
        'Đại diện học sinh trình bày trang web sau khi liên kết hai tệp HTML thành công.',
        'Giáo viên nhận xét, giải đáp nhanh lỗi không tìm thấy tệp do lưu nhầm thư mục hoặc sai tên tệp.'
    )

    eng.add(b, '3. Hoạt động 3: Luyện tập thực hành (Luyện tập 1 & 2 – SGK trang 61)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Thiết lập liên kết hai chiều giữa CLB.html và thong_tin.html; thực hành đặt đường dẫn tuyệt đối.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm:**', first=0)
    eng.add(b, '- **Luyện tập 1:** Thêm <a href="CLB.html">Quay lại trang Câu lạc bộ</a> ở cuối tệp thong_tin.html để điều hướng hai chiều.', first=0)
    eng.add(b, '- **Luyện tập 2:** Tạo tệp gioi_thieu_ban_than.html (thông tin cá nhân lớp 12/1) có ít nhất 03 liên kết ngoài bằng đường dẫn tuyệt đối.', first=0)
    add_steps(
        b,
        'Giáo viên giao hai bài Luyện tập SGK trang 61.',
        'Học sinh hoàn thiện liên kết quay lại trên thong_tin.html và xây dựng tệp gioi_thieu_ban_than.html.',
        'Học sinh kiểm tra chéo bài với bạn cùng máy và trình bày sản phẩm tiêu biểu trước lớp.',
        'Giáo viên chấm nhanh một số bài đạt yêu cầu và lưu ý định dạng URL đầy đủ giao thức cho liên kết ngoài.'
    )

    eng.add(b, '4. Hoạt động 4: Vận dụng thực hành (Vận dụng – SGK trang 61)', 'sub')
    eng.add(b, '**a) Mục tiêu:** Vận dụng thành thạo liên kết nội bộ từ các ô trong bảng dữ liệu đến nhiều vị trí khác nhau trên trang.', first=0)
    eng.add(b, '**b) Nội dung & c) Sản phẩm:** Ở cuối CLB.html (dưới bảng), bổ sung 05 đoạn giới thiệu CLB có gắn mã định danh (id="bongda", id="bongchuyen", id="bongro", id="vo", id="boi"); trong bảng lịch hoạt động, biến tên từng CLB thành thẻ <a> dẫn tới id tương ứng.', first=0)
    add_steps(
        b,
        'Giáo viên giao bài Vận dụng trang 61 và gợi ý quy tắc đặt tên id không dấu cho 5 đoạn văn bản.',
        'Học sinh thực hành thêm 5 đoạn giới thiệu cuối trang CLB.html và gắn thẻ <a> cho tên từng CLB trong bảng.',
        'Học sinh chạy thử trên trình duyệt, kiểm tra từng liên kết trong bảng cuộn đúng tới đoạn giới thiệu tương ứng.',
        'Giáo viên tổng kết tiết thực hành, nhận xét tinh thần làm việc và công bố tiêu chí đánh giá.'
    )

    eng.add(b, 'IV. CÂU HỎI KIỂM TRA VÀ TIÊU CHÍ ĐÁNH GIÁ THỰC HÀNH', 'heading')
    eng.add(b, '- **Câu hỏi kiểm tra nhanh đầu giờ:** Phân biệt vai trò thuộc tính id="CLBTT" và giá trị href="#CLBTT"; giải thích vì sao thong_tin.html phải lưu cùng thư mục với CLB.html khi viết href="thong_tin.html".', first=0)
    rubric_intro = eng.add(b, '- **Tiêu chí đánh giá sản phẩm thực hành trên máy tính (Rubric 10 điểm):**', first=0)
    rubric_intro.find(eng.w('pPr')).append(eng.node('keepNext'))
    rubric_table = eng.table([
        ['TIÊU CHÍ ĐÁNH GIÁ', 'MÔ TẢ YÊU CẦU ĐẠT', 'ĐIỂM TỐI ĐA'],
        ['1. Nhiệm vụ 1 (Hình 10.2)', 'Hiển thị đủ phần giới thiệu đầu trang CLB.html; nháy vào liên kết cuộn đúng tới bảng có id="CLBTT".', '3,0 điểm'],
        ['2. Nhiệm vụ 2 (Trang mới)', 'Tạo tệp thong_tin.html cùng thư mục, đủ nội dung; nháy liên kết từ CLB.html mở đúng trang mới.', '2,5 điểm'],
        ['3. Luyện tập 1 & 2', 'Có liên kết quay về từ thong_tin.html; tệp gioi_thieu_ban_than.html chạy đúng liên kết ngoài.', '2,0 điểm'],
        ['4. Vận dụng & Tác phong', 'Gắn đủ 5 id và 5 thẻ <a> trong bảng CLB.html hoạt động chính xác; lưu trữ thư mục khoa học.', '2,5 điểm']
    ], [2500, 5200, 1650], header=True, cell_sz=22, cant_split=True)
    b.append(rubric_table)

    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 13 tháng 11 năm 2026',
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
    )

    out_path = DE_CUONG_DIR / 'de_cuong_du_gio_12-1.docx'
    save_de_cuong(b, out_path, used_figures=['hinh_10_2.png'])


def sync_to_khoi12():
    """Đồng bộ 2 file lớp 12/1 và 12/2 sang khoi12/ để tránh lệch phiên bản."""
    for fname in ('de_cuong_du_gio_12-1.docx', 'de_cuong_du_gio_12-2.docx'):
        src = DE_CUONG_DIR / fname
        dst = KHOI12_DIR / fname
        if src.exists() and dst.parent.exists():
            shutil.copy2(src, dst)


if __name__ == '__main__':
    DE_CUONG_DIR.mkdir(parents=True, exist_ok=True)
    build_12_4_bai9()
    build_12_2_bai10_tiet1()
    build_12_1_bai10_tiet2()
    sync_to_khoi12()
    print('Successfully built and synced all 3 De cuong files in:', DE_CUONG_DIR)
