# -*- coding: utf-8 -*-
"""
Bo tao ho so DAY THAT cho tiet giang day thuc te Lop 10/3:
- Thoi gian: Thu Hai, ngay 16/11/2026 - Tiet 3 (09h00 - 09h45)
- Dia diem: Phong H.I.5 (1 may tinh GV + may chieu + bang phan)
- Bai day: Tin hoc 10 (KNTT) - Bai 13: Bo sung cac doi tuong do hoa (Tiet 1 - Li thuyet / Tiet PPCT: 19, SGK tr. 69-72, lam tien de cho Tiet 2 - TH tr. 72-74)
- GVHD / GVGD: Thay Mai Xuan Van | SVTT: Ho Quang Huy (24S1020030)

Xuat ra 03 tep trong thu muc c:/Users/PC/khbd_docx/Day_that/:
  1. 01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx
     (Ke hoach bai day chuan CV 5512, van phong su pham chuan muc, dung pham vi Tiet 1 Li thuyet tr. 69-72 lam tien de cho Tiet 2)
  2. 02_Kich_ban_len_lop_chi_tiet_vui_nhon_Tin10_Bai13_Tiet1.docx
     (Cam nang cam tay chi viec: noi gi, lam gi, bam nut nao tren mau_bai13.svg, xu ly tinh huong & lien he bai cu/thuc te)
  3. 03_Bai_thuyet_trinh_giang_day_lien_mach_45_phut_Tin10_Bai13_Tiet1.docx
     (Kich ban thuyet trinh giang day lien mach tu phut 00 den phut 45 khop 100% KHBD so 01)
"""

import sys
import re
import zipfile
from pathlib import Path
from lxml import etree as LX

ROOT_DIR = Path(r'c:\Users\PC\khbd_docx')
SCRIPTS_10_DIR = ROOT_DIR / 'khoi10' / '_scripts_build'
sys.path.insert(0, str(SCRIPTS_10_DIR))

import docx_engine_5512 as eng

DAY_THAT_DIR = ROOT_DIR / 'Day_that'

GV_NAME = 'Thầy Mai Xuân Văn'
SV_NAME = 'Hồ Quang Huy'
MSV = '24S1020030'


def _set_run_sz(para, sz):
    """Đặt cỡ chữ đồng bộ cho toàn bộ run trong một đoạn."""
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


def clean_and_validate_docx(body, out_path, used_figures=None, is_khbd=True):
    """Chuẩn hoá trình bày XML (100% Times New Roman, không rách bảng/chữ ký, không giãn chữ Shift+Enter)."""
    w = eng.w

    # 1. Căn phải dòng hành chính ở đầu trang 1 và cố định cỡ chữ 12pt (24 half-pt)
    for para in body.findall(w('p')):
        text = ''.join(t.text or '' for t in para.iter(w('t')))
        tabs = para.find(w('pPr') + '/' + w('tabs'))
        if tabs is not None and any(tab.get(w('pos')) in ('3800', '4200') for tab in tabs.findall(w('tab'))):
            for tab in tabs.findall(w('tab')):
                if tab.get(w('pos')) in ('3800', '4200'):
                    tab.set(w('val'), 'right')
                    tab.set(w('pos'), '9071')
            _set_run_sz(para, 24)
        if text.startswith(('(Thời lượng', '(Tài liệu', 'Môn học:', 'Thời gian thực hiện:', 'Lịch dự giờ:', 'Hình thức tổ chức:', 'Mục tiêu sử dụng:', 'Lịch dạy:', 'Khớp 100%:', 'Bối cảnh phòng', 'Chiến thuật dạy:')):
            _set_run_sz(para, 24)

    # 2. Trang bìa: rút gọn nhãn hàng đầu không wrap để toàn bộ trang bìa gọn trong 1 trang
    cover = None
    for tbl in body.findall(w('tbl')):
        grid = tbl.find(w('tblGrid'))
        if grid is not None and sum(int(c.get(w('w'), 0)) for c in grid) == 7370:
            cover = tbl
            break
    if cover is not None:
        for t in cover.iter(w('t')):
            if (t.text or '').strip() in ('GV giảng dạy / GV hướng dẫn', 'Giáo viên giảng dạy / GVHD'):
                t.text = 'Giáo viên'

    # 3. Tránh giãn chữ khó đọc khi đoạn chứa <w:br/> hoặc ký tự mã lệnh; giữ dòng trong đoạn
    for para in body.iter(w('p')):
        pp = para.find(w('pPr'))
        if pp is None:
            pp = eng.node('pPr')
            para.insert(0, pp)
        jc = pp.find(w('jc'))
        text = ''.join(t.text or '' for t in para.iter(w('t')))
        has_br = any(br.get(w('type')) is None for br in para.iter(w('br')))
        if jc is not None and jc.get(w('val')) in ('both', 'distribute'):
            if has_br or re.search(r'[A-Za-z_]{3,}|[+/\\=]|Ctrl|Shift|Path|Fill|Stroke', text):
                jc.set(w('val'), 'left')
        if pp.find(w('keepLines')) is None:
            pp.append(eng.node('keepLines'))

    # 4. Ngăn hàng tiêu đề và các hàng ngắn bị xé lẻ; không đặt cantSplit cho hàng dài để bảng tự chảy mượt qua trang; không để hàng tiêu đề đứng mồ côi ở cuối trang, và bề rộng <= 9071 twip
    for table in body.iter(w('tbl')):
        rows = table.findall(w('tr'))
        for idx_r, row in enumerate(rows):
            rp = row.find(w('trPr'))
            if rp is None:
                rp = eng.node('trPr')
                row.insert(0, rp)
            row_text = ''.join(t.text or '' for t in row.iter(w('t')))
            if idx_r == 0 or len(row_text) < 550:
                if rp.find(w('cantSplit')) is None:
                    rp.append(eng.node('cantSplit'))
            else:
                for cs in list(rp.findall(w('cantSplit'))):
                    rp.remove(cs)
            if idx_r == 0 and len(rows) > 1 and table is not cover:
                for para in row.iter(w('p')):
                    pp = para.find(w('pPr'))
                    if pp.find(w('keepNext')) is None:
                        pp.append(eng.node('keepNext'))
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

    # 5. Giữ khối chữ ký V. NHẬN XÉT không tách trang
    if is_khbd:
        approval = False
        sig_paras = []
        for el in body:
            text = ''.join(t.text or '' for t in el.iter(w('t')))
            if text.startswith('V. NHẬN XÉT'):
                approval = True
            if approval:
                sig_paras.extend(el.iter(w('p')))
        for para in sig_paras[:-1]:
            pp = para.find(w('pPr'))
            if pp.find(w('keepNext')) is None:
                pp.append(eng.node('keepNext'))
        if sig_paras:
            pp = sig_paras[-1].find(w('pPr'))
            for n in list(pp.findall(w('keepNext'))):
                pp.remove(n)

    eng.save_docx(body, out_path, used_figures=used_figures or [])

    # 6. Chuẩn hoá 100% font Times New Roman toàn bộ gói XML và thêm doNotExpandShiftReturn
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

    # 7. Kiểm tra XML
    with zipfile.ZipFile(out_path) as archive:
        root = LX.fromstring(archive.read('word/document.xml'))
        text = ' '.join(root.xpath('//w:t/text()', namespaces=ns))
        bad = re.search(r'[`$\u200b\ufffd]|\*\*|\\rightarrow|\§|->|→', text)
        if bad:
            raise ValueError(f'{out_path.name}: phát hiện ký tự thô {bad.group()}')
        for req in ('10/3', 'H.I.5', '16/11/2026', SV_NAME, GV_NAME):
            if req not in text:
                raise ValueError(f'{out_path.name}: thiếu thông tin {req}')
        if is_khbd:
            for forbidden in ('không có máy tính', 'mách có chứng', 'không phát giấy', 'Khả thi 100%', 'trung bình/yếu', 'vừa sức', 'Hình 13.14', 'Hình 13.15', 'Hình 13.16', 'Hình 13.11'):
                if forbidden.lower() in text.lower():
                    raise ValueError(f'{out_path.name}: còn chứa cụm từ không phù hợp KHBD "{forbidden}"')
        for name in archive.namelist():
            if name.startswith('word/') and name.endswith('.xml'):
                part = LX.fromstring(archive.read(name))
                for fonts in part.xpath('//w:rFonts', namespaces=ns):
                    if set(fonts.attrib.values()) != {'Times New Roman'}:
                        raise ValueError(f'{out_path.name}: font chưa chuẩn ở {name}')
        for fig in (used_figures or []):
            if 'word/media/' + fig not in archive.namelist():
                raise ValueError(f'{out_path.name}: thiếu hình {fig}')
    print(f'XML OK: {out_path.name} ({len(used_figures or [])} hình SGK)')


# ==============================================================================
# FILE 1: KE HOACH BAI DAY CHUAN CV 5512 - TIET 1 LI THUYET (10/3)
# ==============================================================================
def build_khbd_day_that():
    b = eng.init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 - LÍ THUYẾT – TIẾT PPCT: 19)',
        class_str='Tin học 10 – Lớp 10/3 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.I.5',
        lesson_row_str='Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT) – Tiết PPCT: 19 (Tổng bài: 02 tiết – PPCT: 19, 20)',
        page1_title='TÊN BÀI DẠY: BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 – LÍ THUYẾT)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 19, 20] – Kế hoạch này chi tiết hoá Tiết thứ 1 của bài / Tiết PPCT: 19)',
        page1_class_line='Môn học: Tin học; Lớp: 10/3',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 1 - LT / Tiết PPCT: 19)',
        page1_sched_line='Lịch dự giờ: Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 10/3 – Phòng H.I.5',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết trên lớp kết hợp trình chiếu, minh hoạ trực quan trên Inkscape và vấn đáp SGK'
    )

    # I. MỤC TIÊU
    eng.add(b, 'I. MỤC TIÊU', 'heading')
    eng.add(b, 'Căn cứ Yêu cầu cần đạt Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E, khung Năng lực số (Thông tư số 02/2025/TT-BGDĐT) và Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá (Tiết PPCT 19 – Mục 4.13. Bổ sung các đối tượng đồ họa, tiết 1 - LT; chi tiết trích lục đính kèm tại Phụ lục 1, Mục IV. PHỤ LỤC):')

    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu **`[KT1]`** – **`[KT3]`**):', keep=True)
    for item in [
        '**`[KT1]`**: Nhận biết và nêu được ý nghĩa các thuộc tính cơ bản của ba nhóm đối tượng hình khối trên Thanh điều khiển thuộc tính theo Bảng trang 69–70 SGK: hình vuông, hình chữ nhật (`W, H`, `Rx, Ry`); hình tròn, hình elip (`Rx, Ry`, `Start, End` và ba chế độ vẽ hình quạt, cung tròn, hình tròn kín); hình đa giác, hình sao (`Corners`, `Rounded`, `Spoke ratio`, `Randomized`).',
        '**`[KT2]`**: Trình bày được cách mở và chức năng ba trang của hộp thoại **Fill and Stroke** (`Object / Fill and Stroke`, nháy nút phải chuột chọn `Fill and Stroke` hoặc nhấn `Shift + Ctrl + F`) theo Hình 13.4 SGK trang 70–71: `Fill` (màu tô nền), `Stroke paint` (màu vẽ đường viền), `Stroke style` (độ dày và kiểu nét vẽ); phân biệt được các kiểu tô (không màu `X`, màu đồng nhất, `Linear gradient`, `Radial gradient`, hoa văn `Pattern`), nút `?` (huỷ đặt), công cụ chỉnh sửa màu chuyển (biểu tượng cây bút), công cụ ống hút sao chép màu và độ trong suốt (`Opacity`, giá trị `A` alpha).',
        '**`[KT3]`**: Phân biệt được 06 phép ghép đối tượng đồ hoạ trong bảng chọn **Path** (Hình 13.6 SGK trang 71–72): Phép hợp (`Union`, `Ctrl + +`), Phép hiệu (`Difference`, `Ctrl + -`), Phép giao (`Intersection`, `Ctrl + *`), Phép hiệu đối xứng (`Exclusion`, `Ctrl + ^`), Phép chia (`Division`, `Ctrl + /`) và Phép cắt (`Cut Path`, `Ctrl + Alt + /`); nắm vững quy tắc đối tượng lớp dưới bị cắt bởi đối tượng lớp trên để làm cơ sở thực hành ở Tiết 2.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**`[NLC1 – Tự chủ và tự học]`**: Chủ động quan sát hình ảnh trình chiếu, đọc hiểu và tra cứu bảng biểu, hình minh hoạ trong Sách giáo khoa Tin học 10 trang 69–72 để hoàn thành nhiệm vụ học tập.',
        '**`[NLC2 – Giao tiếp và hợp tác]`**: Tích cực trao đổi theo cặp đôi, lắng nghe và thống nhất ý kiến trình bày khi phân tích thuộc tính hình khối, các kiểu tô màu và các phép ghép đối tượng.',
        '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Vận dụng được kiến thức lí thuyết về thuộc tính hình khối, hộp thoại `Fill and Stroke` và bảng chọn `Path` để giải quyết các câu hỏi củng cố (trang 70–72), đề xuất quy trình vẽ miếng dưa hấu (Hình 13.1) và định hướng thực hiện các nhiệm vụ thực hành ở Tiết 2.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**`[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]`**: Nhận biết chính xác vị trí, chức năng của Thanh điều khiển thuộc tính, hộp thoại `Fill and Stroke` và bảng chọn `Path` trên giao diện phần mềm `Inkscape`.',
        '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông]`**: Phân tích được đối tượng đồ hoạ đơn giản (trái tim, đám mây, miếng dưa hấu) thành các hình khối cơ bản, thứ tự lớp và phép ghép đối tượng tương ứng.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'sub')
    eng.add(b, '**`[NLS 3.2.NC1.a]`**: Hiểu rõ nguyên lí điều chỉnh thuộc tính, màu tô, độ trong suốt và kết hợp các đối tượng đồ hoạ vectơ làm nền tảng cho hoạt động thực hành tạo lập sản phẩm số ở Tiết 2 (Tiết PPCT: 20).', 'bullet')

    eng.add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**`[PC1 – Chăm chỉ]`**: Chú ý theo dõi phần trình chiếu và thao tác minh hoạ của giáo viên, chủ động đọc Sách giáo khoa và ghi chép bài đầy đủ.',
        '**`[PC2 – Trách nhiệm]`**: Tự giác thực hiện nhiệm vụ cá nhân, hợp tác nghiêm túc khi thảo luận cặp đôi và tích cực tham gia phát biểu xây dựng bài.',
        '**`[PC3 – Trung thực]`**: Trung thực trong quá trình tham gia trả lời câu hỏi trắc nghiệm, vấn đáp và thẳng thắn trao đổi những nội dung chưa rõ.'
    ]:
        eng.add(b, item, 'bullet')

    # II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính giáo viên cài đặt phần mềm **Inkscape**, máy chiếu (hoặc màn hình trình chiếu) tại lớp học, bảng viết, phấn.',
        '**Học liệu chuẩn (SGK):** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống (Nhà xuất bản Giáo dục Việt Nam)*, **Bài 13: Bổ sung các đối tượng đồ hoạ** (trọng tâm Tiết 1 – Lí thuyết: trang 69–72; định hướng chuẩn bị Tiết 2 – Thực hành: trang 72–74).',
        '**Học liệu số và công cụ đánh giá của giáo viên:** Bộ hình ảnh trực quan trích từ SGK, hệ thống câu hỏi vấn đáp – trắc nghiệm củng cố (`CCĐG 1` đến `CCĐG 4`), bảng tiêu chí đánh giá Rubric (`CCĐG 5`) và bản trích lục căn cứ đính kèm tại Mục `IV. PHỤ LỤC`; tệp đồ hoạ minh hoạ chuẩn bị sẵn để giáo viên thao tác thị phạm trực quan trên phần mềm `Inkscape`.'
    ]:
        eng.add(b, item, 'bullet')

    # III. TIẾN TRÌNH DẠY HỌC
    eng.add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    eng.add(b, '**Bảng tổng quan tiến trình dạy học Tiết 1 – Lí thuyết (45 phút – Tiết 1 của bài / Tiết PPCT: 19)**', first=0, keep=True)
    ov_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng', 'Mã mục tiêu', 'Sản phẩm cần đạt & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Quan sát Hình 13.1 – Miếng dưa hấu)', '05 phút\n(PPCT 19)', '[KT1], [NLC1],\n[PC1]', 'Câu trả lời của học sinh nêu đúng tên hình khối, màu sắc và thứ tự 4 lớp từ dưới lên trên.\nCông cụ đánh giá: CCĐG 1 (Bảng kiểm vấn đáp ngắn).'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Đối tượng hình khối & thuộc tính (09p)\n- 2.2. Nhiệm vụ 2: Thiết lập màu tô, màu vẽ – Fill and Stroke (09p)\n- 2.3. Nhiệm vụ 3: Các phép ghép đối tượng Path (10p)', '28 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2], [NLC3],\n[NLa], [NLc],\n[NLS 3.2.NC1.a],\n[PC1], [PC2], [PC3]', 'Học sinh nêu được ý nghĩa các thuộc tính hình khối (tr. 69–70), phân biệt chức năng hộp thoại Fill and Stroke (tr. 70–71), phân biệt 6 phép ghép trong bảng chọn Path (tr. 71–72) và ghi vở ý chính.\nCông cụ đánh giá: CCĐG 2, CCĐG 3, CCĐG 4.'],
        ['3. Hoạt động 3: Luyện tập\n(Giải quyết 04 câu hỏi củng cố lí thuyết SGK tr. 70, 71, 72 & Hình 13.7)', '07 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC2], [NLC3], [NLc],\n[PC2], [PC3]', 'Câu trả lời chính xác cho câu hỏi tham số Rounded (tr. 70), 02 câu trắc nghiệm Fill and Stroke (tr. 71) và phép ghép cùng các bước vẽ đám mây Hình 13.7 (tr. 72).\nCông cụ đánh giá: CCĐG 5 (Rubric quan sát vấn đáp).'],
        ['4. Hoạt động 4: Vận dụng\n(Đề xuất cách vẽ Miếng dưa hấu H.13.1 & Định hướng Tiết 2 Thực hành)', '05 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC3], [NLc],\n[NLS 3.2.NC1.a],\n[PC1], [PC2]', 'Học sinh đề xuất được 3 bước vẽ Miếng dưa hấu Hình 13.1 và xác định được công cụ/lệnh tương ứng cho 03 nhiệm vụ thực hành (tr. 72–74) ở Tiết 2.\nCông cụ đánh giá: CCĐG 5 (Rubric quan sát vấn đáp).']
    ]
    b.append(eng.table(ov_rows, [2700, 1150, 1950, 3271], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 1 ---
    eng.add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[NLC1]`, `[PC1]` – Học sinh ôn lại quy tắc thứ tự lớp đối tượng đã học ở Bài 12 thông qua quan sát Hình 13.1, từ đó nhận thấy nhu cầu tìm hiểu cách điều chỉnh thuộc tính hình khối, thiết lập màu tô nâng cao và ghép các đối tượng để tạo hình vẽ hoàn chỉnh.', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên trình chiếu **Hình 13.1 (Miếng dưa hấu – SGK trang 69)** và nêu nhiệm vụ quan sát, phân tích cho cả lớp:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời của học sinh nêu đúng tên các đối tượng hình học, màu sắc cấu thành và thứ tự 4 lớp từ dưới lên trên của miếng dưa hấu (Hình 13.1).', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 1` *(Bảng kiểm vấn đáp ngắn – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu Hình 13.1 (SGK trang 69) lên màn hình, yêu cầu học sinh mở SGK trang 69 và quan sát, suy nghĩ cá nhân trong 01 phút để trả lời 02 câu hỏi:\n'
            '  + Câu hỏi 1: Quan sát miếng dưa hấu ở Hình 13.1, em hãy kể tên các đối tượng hình học và màu sắc cấu tạo nên hình vẽ đó.\n'
            '  + Câu hỏi 2: Vận dụng quy tắc thứ tự lớp đã học ở Bài 12 (đối tượng vẽ trước nằm ở lớp dưới, vẽ sau nằm ở lớp trên), em hãy xác định thứ tự 4 lớp (vỏ xanh, cùi trắng, ruột đỏ, hạt đen) từ lớp dưới cùng đến lớp trên cùng.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 69, quan sát Hình 13.1 trên màn chiếu kết hợp với hình minh hoạ trong sách, đối chiếu kiến thức Bài 12 để xác định các hình khối và thứ tự từng lớp đối tượng.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời lần lượt 02 học sinh đứng tại chỗ trả lời Câu hỏi 1 và Câu hỏi 2; các học sinh khác lắng nghe, nhận xét và bổ sung.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá câu trả lời của học sinh theo `CCĐG 1`, thao tác kéo tách nhẹ các lớp của miếng dưa hấu mẫu trên phần mềm `Inkscape` qua màn chiếu để học sinh kiểm chứng trực quan thứ tự 4 lớp, rồi đặt vấn đề vào bài mới: *"Ở Bài 12, các em mới chỉ vẽ được hình tròn nguyên vẹn với màu tô đồng nhất. Để cắt được nửa hình tròn như miếng dưa hấu, tô màu chuyển tạo chiều sâu hoặc ghép nhiều hình thành một đối tượng phức tạp, bài học hôm nay sẽ giúp các em nắm vững các lệnh và thuộc tính bổ sung trong Inkscape."*',
            '**Khởi động – Phân tích Hình 13.1 (SGK trang 69):**\n'
            '- **Các đối tượng cấu thành:**\n'
            '  + 03 nửa hình tròn có kích thước giảm dần (vỏ xanh lá đậm, cùi xanh nhạt/trắng, ruột đỏ).\n'
            '  + Các hình giọt nước/elip nhỏ màu đen (hạt dưa).\n'
            '- **Thứ tự 4 lớp (từ dưới lên trên):**\n'
            '  1. Vỏ xanh (lớp dưới cùng), tiếp đến 2. Cùi trắng, tiếp theo 3. Ruột đỏ và trên cùng là 4. Các hạt đen.'
        ]
    ]
    b.append(eng.table(hd1_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 2 ---
    eng.add(b, '2. Hoạt động 2: Hình thành kiến thức mới (28 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')

    # Nhiệm vụ 2.1
    eng.add(b, '2.1. Nhiệm vụ 1: Tìm hiểu các đối tượng hình khối và thuộc tính cơ bản (09 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[PC1]` – Học sinh nhận biết được công cụ tạo hình ở Hoạt động 1 SGK (Hình 13.2) và trình bày được ý nghĩa các thuộc tính cơ bản của ba nhóm hình khối trên Thanh điều khiển thuộc tính theo Bảng trang 69–70 SGK.', first=0)
    eng.add(b, '**b) Nội dung:** Học sinh quan sát **Hình 13.2 (SGK trang 69)** và nghiên cứu **Bảng các thuộc tính cơ bản của một số hình có sẵn (SGK trang 69–70)** để thực hiện nhiệm vụ thảo luận cặp đôi:', first=0, keep=True)
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời của học sinh nêu đúng công cụ vẽ 4 hình ở Hình 13.2, giải thích được ý nghĩa các thuộc tính `W, H`, `Rx, Ry`, `Start, End`, `Corners`, `Rounded`, `Spoke ratio`, `Randomized` và nội dung ghi chép vào vở.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 2` *(Bảng câu hỏi tra cứu Bảng thuộc tính SGK trang 69–70 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu Hình 13.2 (SGK trang 69) và nêu câu hỏi Hoạt động 1 SGK: *"Bốn hình màu đỏ ở Hình 13.2 được vẽ từ cùng một công cụ có sẵn nào của Inkscape? Theo em, để tạo ra 4 hình dạng khác nhau như vậy từ công cụ đó, ta cần thay đổi những thông số nào?"*\n'
            '- Tiếp theo, giáo viên giao nhiệm vụ trao đổi cặp đôi trong 02 phút: Học sinh nghiên cứu Bảng các thuộc tính cơ bản của một số hình có sẵn (SGK trang 69–70) để hoàn thành 03 yêu cầu:\n'
            '  1. Nêu ý nghĩa các thuộc tính `W, H` và `Rx, Ry` của hình vuông, hình chữ nhật.\n'
            '  2. Nêu ý nghĩa các thuộc tính `Rx, Ry`, `Start, End` cùng 3 chế độ vẽ của hình tròn, hình elip; xác định góc `Start` và `End` để tạo nửa hình tròn (như miếng dưa hấu).\n'
            '  3. Nêu ý nghĩa 4 thuộc tính `Corners`, `Rounded`, `Spoke ratio`, `Randomized` của hình đa giác, hình sao.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 69–70, quan sát Hình 13.2, đọc kĩ nội dung Bảng thuộc tính và trao đổi với bạn cùng bàn để xác định tên công cụ cùng ý nghĩa từng thuộc tính.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời lần lượt đại diện 02–03 cặp học sinh đứng tại chỗ trả lời câu hỏi Hoạt động 1 (Hình 13.2) và ý nghĩa các thuộc tính của 3 nhóm hình khối; các học sinh khác lắng nghe, nhận xét và bổ sung.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, đánh giá theo `CCĐG 2`, sau đó **thao tác minh hoạ trực quan trên phần mềm Inkscape qua màn chiếu (02 phút)**: chọn công cụ Hình đa giác, hình sao biến đổi lần lượt 4 kiểu như Hình 13.2; điều chỉnh `Rx = 25` để bo tròn 4 góc hình chữ nhật; điều chỉnh `Start = 0, End = 180` và giới thiệu 3 nút chế độ (hình quạt, cung tròn hở, hình tròn kín) trên hình tròn; chuẩn hoá nội dung ghi bài lên bảng.',
            '**1. Các đối tượng hình khối và thuộc tính cơ bản (SGK trang 69–70)**\n'
            '- **Hoạt động 1 (Hình 13.2):** Cả 4 hình đều vẽ từ công cụ **Hình đa giác, hình sao** trên hộp công cụ, bằng cách chuyển đổi chế độ đa giác/sao (`Corners = 5`) và điều chỉnh `Rounded`, `Randomized`.\n'
            '- **Bảng các thuộc tính cơ bản trên Thanh điều khiển thuộc tính:**\n'
            '  1. **Hình vuông, hình chữ nhật:** `W, H` (chiều rộng, chiều dài); `Rx, Ry` (bán kính của góc bo, bằng 0 nếu góc của hình là góc vuông).\n'
            '  2. **Hình tròn, hình elip:** `Rx, Ry` (bán kính theo phương ngang và thẳng đứng); `Start, End` (góc điểm đầu và điểm cuối tính theo độ khi sử dụng các chế độ vẽ hình quạt, cung tròn hở hoặc hình tròn kín; đặt `Start = 0, End = 180` tạo nửa hình tròn).\n'
            '  3. **Hình đa giác, hình sao:** `Corners` (số đỉnh của đa giác hoặc số cánh của hình sao); `Rounded` (độ cong tại các đỉnh của hình); `Spoke ratio` (tỉ lệ bán kính từ tâm đến góc trong và từ tâm tới đỉnh đầu hình sao); `Randomized` (tham số làm méo hình ngẫu nhiên).'
        ]
    ]
    b.append(eng.table(nv1_rows, [4650, 4421], header=True, cell_sz=22))

    # Nhiệm vụ 2.2
    eng.add(b, '2.2. Nhiệm vụ 2: Tìm hiểu cách thiết lập màu tô, màu vẽ và hộp thoại Fill and Stroke (09 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT2]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh nhận biết sự khác nhau giữa các kiểu tô màu ở Hình 13.3, nêu được cách mở hộp thoại `Fill and Stroke`, chức năng ba trang (`Fill`, `Stroke paint`, `Stroke style`), các kiểu tô, công cụ chỉnh sửa màu chuyển, sao chép màu và độ trong suốt (`Opacity`, kênh `A`).', first=0)
    eng.add(b, '**b) Nội dung:** Học sinh quan sát **Hình 13.3**, **Hình 13.4** và đọc mục 2 *(SGK trang 70–71)* để thảo luận và trả lời các câu hỏi tìm hiểu kiến thức mới:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời của học sinh phân biệt được 5 kiểu tô trong Hình 13.3, nêu đúng cách mở, chức năng 3 trang của hộp thoại `Fill and Stroke`, ý nghĩa `Opacity` / kênh `A` và nội dung ghi chép vào vở.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 3` *(Bảng câu hỏi tìm hiểu hộp thoại Fill and Stroke SGK trang 70–71 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu Hình 13.3 (SGK trang 70) và nêu câu hỏi Hoạt động 2 SGK: *"Quan sát 5 hình vuông trong Hình 13.3 và nhận xét phần màu tô bên trong các hình có điểm gì khác nhau?"*\n'
            '- Tiếp theo, giáo viên trình chiếu Hình 13.4 (Hộp thoại `Fill and Stroke` – SGK trang 70), yêu cầu học sinh đọc mục 2 (trang 70–71) theo cặp đôi trong 02 phút để trả lời 03 câu hỏi:\n'
            '  1. Nêu các cách mở hộp thoại `Fill and Stroke` và chức năng của 3 trang: `Fill`, `Stroke paint`, `Stroke style`.\n'
            '  2. Phân biệt các biểu tượng kiểu tô màu trong Hình 13.4 (dấu `X`, màu đồng nhất, màu chuyển giữa hai hay nhiều màu `Linear gradient`, màu chuyển từ tâm `Radial gradient`, hoa văn `Pattern`) và tác dụng nút `?`.\n'
            '  3. Nêu vai trò của biểu tượng cây bút, công cụ ống hút màu phía dưới hộp thoại và ý nghĩa của giá trị `Opacity` (kênh `A` alpha).\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 70–71, quan sát Hình 13.3, Hình 13.4, đối chiếu 5 biểu tượng kiểu tô màu và 3 bước thực hiện chỉnh sửa nền, đường nét trong SGK để thảo luận.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên gọi 01 học sinh nhận xét sự khác nhau của 5 hình ở Hình 13.3; mời 02 học sinh trình bày cách mở, chức năng 3 trang, các kiểu tô màu và độ trong suốt; học sinh khác nhận xét, bổ sung.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo `CCĐG 3`, **thao tác minh hoạ trực quan trên Inkscape qua màn chiếu (02 phút)**: mở hộp thoại `Fill and Stroke` (`Shift + Ctrl + F`), tô màu chuyển `Linear gradient` và `Radial gradient` (chỉ biểu tượng cây bút điều chỉnh điểm dừng màu), chuyển trang `Stroke style` thay đổi độ dày `Width` và kiểu nét đứt, rồi kéo thanh `Opacity` để học sinh quan sát hiệu ứng trong suốt giữa hai hình đè lớp; chuẩn hoá nội dung ghi bảng.',
            '**2. Thiết lập màu tô, màu vẽ và tô màu cho đối tượng (SGK trang 70–71)**\n'
            '- **Nhận xét Hình 13.3:** Cả 5 hình đều có màu viền cam giống nhau nhưng khác nhau về phần màu tô bên trong: (1) không màu; (2) màu đồng nhất; (3) chuyển màu theo chiều ngang (`Linear gradient`); (4) chuyển màu từ tâm ra ngoài (`Radial gradient`); (5) chuyển màu có độ trong suốt.\n'
            '- **Cách mở hộp thoại Fill and Stroke:** Chọn đối tượng, chọn lệnh `Object / Fill and Stroke`, hoặc nháy nút phải chuột vào đối tượng chọn `Fill and Stroke`, hoặc nhấn tổ hợp phím `Shift + Ctrl + F`.\n'
            '- **Ba trang chính của hộp thoại:**\n'
            '  + **Fill (Màu tô):** Chọn kiểu tô cho nền bên trong đối tượng gồm: `X` (không màu/trong suốt), màu đồng nhất, màu chuyển giữa hai hoặc nhiều màu (`Linear gradient`), màu chuyển từ tâm đối tượng (`Radial gradient`), hoa văn (`Pattern`) và nút `?` (huỷ đặt đối tượng trở về trạng thái ban đầu).\n'
            '  + **Stroke paint (Màu vẽ):** Thiết lập màu cho đường viền (có công cụ ống hút sao chép màu từ vị trí bất kì trong vùng làm việc).\n'
            '  + **Stroke style (Kiểu nét vẽ):** Thay đổi độ dày mỏng (`Width`) và kiểu đường viền (nét liền, nét đứt...).\n'
            '- **Opacity và A (alpha):** Xác định độ trong suốt của màu khi tô lên đối tượng.'
        ]
    ]
    b.append(eng.table(nv2_rows, [4650, 4421], header=True, cell_sz=22))

    # Nhiệm vụ 2.3
    eng.add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các phép ghép đối tượng đồ hoạ trong bảng chọn Path (10 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT3]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[NLc]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh đề xuất được cách ghép 2 hình tròn và 1 hình vuông thành trái tim (Hình 13.5), phân biệt đúng 06 phép ghép đối tượng đồ hoạ trong bảng chọn `Path` (Hình 13.6) và nắm vững quy tắc đối tượng lớp dưới bị cắt bởi đối tượng lớp trên.', first=0)
    eng.add(b, '**b) Nội dung:** Học sinh quan sát **Hình 13.5** và **Hình 13.6** *(SGK trang 71–72)*, đọc nội dung mục 3 để giải quyết nhiệm vụ cặp đôi:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời nêu được cách xếp ghép hình trái tim (Hình 13.5), phân biệt đúng tên tiếng Anh, phím tắt và kết quả của 06 phép ghép trong Hình 13.6 cùng quy tắc thứ tự lớp dưới/lớp trên.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 4` *(Bảng hỏi – đáp phân tích Hình 13.5 và 06 phép ghép Hình 13.6 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu Hình 13.5 (SGK trang 71) nêu câu hỏi Hoạt động 3 SGK: *"Từ ba mẫu giấy gồm 2 hình tròn và 1 hình vuông như Hình 13.5, em hãy nêu cách xoay hình và xếp chồng các hình lên nhau để tạo thành hình trái tim?"*\n'
            '- Tiếp theo, giáo viên trình chiếu Hình 13.6 (SGK trang 72) với quy ước hình vuông ở lớp dưới, hình tròn ở lớp trên, yêu cầu học sinh đọc mục 3 (trang 71–72) theo cặp đôi trong 03 phút để thực hiện 2 yêu cầu:\n'
            '  1. Đối chiếu mô tả ở trang 71–72 với hình kết quả trong Hình 13.6, gọi đúng tên tiếng Anh và tổ hợp phím tắt của 06 phép ghép trong bảng chọn `Path`: Phép hợp, Phép hiệu, Phép giao, Phép hiệu đối xứng, Phép chia và Phép cắt.\n'
            '  2. Chỉ ra quy tắc chung về thứ tự lớp trong các phép `Difference`, `Division` và `Cut Path`: đối tượng ở lớp nào là đối tượng bị cắt?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 71–72, thảo luận cặp đôi để phân tích cách xếp Hình 13.5 và đối chiếu từng hình kết quả ở Hình 13.6 với các lệnh trong bảng chọn `Path`.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02–03 học sinh lần lượt trình bày cách ghép hình trái tim (Hình 13.5), tên và kết quả của 06 phép ghép ở Hình 13.6 cùng quy tắc thứ tự lớp; lớp nhận xét, bổ sung.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo `CCĐG 4`, **thao tác minh hoạ trực quan trên Inkscape qua màn chiếu (02 phút)**: thực hiện các lệnh `Path / Union`, `Path / Difference` và `Path / Intersection` trên cặp hình vuông (lớp dưới) và hình tròn (lớp trên), đồng thời nhấn mạnh quy tắc trọng tâm: với Phép hiệu (`Difference`), Phép chia (`Division`) và Phép cắt (`Cut Path`), **hình ở lớp dưới luôn bị cắt bởi hình ở lớp trên**; chốt kiến thức lên bảng để học sinh ghi bài.',
            '**3. Các phép ghép đối tượng đồ hoạ trong bảng chọn Path (SGK trang 71–72)**\n'
            '- **Hoạt động 3 (Hình 13.5):** Xoay hình vuông nghiêng 45 độ, đặt 2 hình tròn chồng lên hai cạnh phía trên của hình vuông rồi thực hiện phép hợp (`Path / Union`) để thu được hình trái tim.\n'
            '- **Sáu phép ghép trong bảng chọn Path (Hình 13.6):**\n'
            '  1. **Phép hợp (Union, Ctrl + +):** Gồm tất cả các phần thuộc một trong các hình đơn.\n'
            '  2. **Phép hiệu (Difference, Ctrl + -):** Phần thuộc hình lớp dưới nhưng không thuộc hình lớp trên.\n'
            '  3. **Phép giao (Intersection, Ctrl + *):** Phần thuộc cả hai hình được chọn.\n'
            '  4. **Phép hiệu đối xứng (Exclusion, Ctrl + ^):** Phần hình thuộc các hình trừ phần giao nhau.\n'
            '  5. **Phép chia (Division, Ctrl + /):** Hình lớp dưới được chia thành các phần bởi đường nét của hình lớp trên.\n'
            '  6. **Phép cắt (Cut Path, Ctrl + Alt + /):** Cắt hình lớp dưới thành các phần bởi hai điểm giao ở viền với hình lớp trên (kết quả là các hình mới không có màu).\n'
            '- **Quy tắc cần nhớ:** Trong `Difference`, `Division` và `Cut Path`, đối tượng lớp dưới luôn bị cắt bởi đối tượng lớp trên.'
        ]
    ]
    b.append(eng.table(nv3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 3 ---
    eng.add(b, '3. Hoạt động 3: Luyện tập (07 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[KT2]`, `[KT3]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[PC2]`, `[PC3]` – Học sinh vận dụng trực tiếp các kiến thức lí thuyết vừa học ở Mục 1, 2, 3 để hoàn thành 04 câu hỏi củng cố chính thức trong SGK trang 70, 71, 72.', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên tổ chức cho học sinh trả lời hệ thống **04 câu hỏi củng cố lí thuyết trong SGK Tin học 10 (trang 70, 71, 72)** kết hợp **Hình 13.7 (Vẽ đám mây – SGK trang 72)**:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời chính xác của học sinh cho 04 câu hỏi củng cố SGK trang 70–72 (tham số `Rounded`, phương án `C. Stroke style`, phương án `B` cây bút chỉnh `gradient`, và phép hợp `Path / Union` cùng 3 bước vẽ đám mây Hình 13.7).', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 5` *(Rubric quan sát vấn đáp Hoạt động Luyện tập & Vận dụng – Phụ lục 3)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu lần lượt 04 câu hỏi củng cố lí thuyết trong SGK trang 70, 71, 72 lên màn hình:\n'
            '  1. *Câu hỏi củng cố Mục 1 (SGK trang 70):* Nếu trên hình vẽ có sẵn một hình sao 5 cánh nhọn, em cần thay đổi tham số nào để các đỉnh ngôi sao trở nên cong?\n'
            '  2. *Câu hỏi trắc nghiệm 1 Mục 2 (SGK trang 71):* Để xác định đường viền của đối tượng dạng nét đứt, em cần chọn trang nào trong hộp thoại `Fill and Stroke`? (`A. Fill`, `B. Stroke paint`, `C. Stroke style`, `D. Cả A và B`).\n'
            '  3. *Câu hỏi trắc nghiệm 2 Mục 2 (SGK trang 71):* Để chỉnh thông số của `gradient`, em cần chọn biểu tượng nào trong 4 phương án `A, B, C, D` (trang 71 SGK)?\n'
            '  4. *Câu hỏi củng cố Mục 3 (SGK trang 72 – Hình 13.7):* Em hãy nêu phép ghép hình và các bước thực hiện để vẽ đám mây như Hình 13.7.\n'
            '- Với Câu 1, 2, 3, giáo viên yêu cầu cả lớp suy nghĩ cá nhân (với Câu 2, 3 cả lớp đồng loạt biểu quyết chọn đáp án bằng cách giơ ngón tay: 1 ngón chọn A, 2 ngón chọn B, 3 ngón chọn C, 4 ngón chọn D). Với Câu 4, học sinh trao đổi nhanh theo cặp đôi trong 02 phút.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh suy nghĩ cá nhân, đối chiếu lại kiến thức Mục 1, 2, 3 trong SGK trang 70–72 để chọn đáp án cho Câu 1, 2, 3 và thống nhất quy trình 3 bước vẽ đám mây Hình 13.7 ở Câu 4.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 01 học sinh trả lời Câu 1; cho cả lớp đồng loạt giơ ngón tay chọn đáp án Câu 2, Câu 3 rồi mời 01 học sinh giải thích ngắn gọn; mời 01–02 học sinh trình bày phép ghép và các bước vẽ đám mây Hình 13.7 ở Câu 4.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo Tiêu chí 1 và Tiêu chí 2 của Rubric `CCĐG 5` (Phụ lục 3); thao tác kiểm chứng nhanh trên phần mềm Inkscape qua màn chiếu (tăng `Rounded` trên ngôi sao, chọn nét đứt trong `Stroke style` và dùng `Path / Union` hợp nhất các hình tròn thành đám mây Hình 13.7) rồi chốt đáp án chuẩn.',
            '**Giải đáp 04 câu hỏi củng cố lí thuyết (SGK trang 70–72):**\n'
            '- **Câu 1 (tr. 70):** Thay đổi tham số **`Rounded`** (độ cong tại các đỉnh của hình).\n'
            '- **Câu 2 (tr. 71):** Chọn **`C. Stroke style`**.\n'
            '- **Câu 3 (tr. 71):** Chọn **`B`** (biểu tượng cây bút chỉnh sửa gradient).\n'
            '- **Câu 4 (tr. 72 – Vẽ đám mây Hình 13.7):** Sử dụng **Phép hợp** (`Path / Union`, `Ctrl + +`) gồm 3 bước:\n'
            '  1. Dùng công cụ Hình tròn/elip vẽ các hình tròn xếp chồng lấp lên nhau tạo dáng đám mây.\n'
            '  2. Chọn tất cả các hình tròn vừa vẽ.\n'
            '  3. Chọn lệnh **`Path / Union`** (hoặc nhấn `Ctrl + +`) để hợp nhất thành hình đám mây hoàn chỉnh.'
        ]
    ]
    b.append(eng.table(hd3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 4 ---
    eng.add(b, '4. Hoạt động 4: Vận dụng (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[KT2]`, `[KT3]`, `[NLC3]`, `[NLc]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh vận dụng tổng hợp kiến thức Tiết 1 để giải quyết tình huống mở đầu (đề xuất cách vẽ Miếng dưa hấu Hình 13.1) và liên hệ chuẩn bị kiến thức nền tảng cho 03 nhiệm vụ thực hành ở Tiết 2 (Tiết PPCT: 20).', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên yêu cầu học sinh quay lại giải quyết vấn đề đặt ra ở **Hình 13.1 (Miếng dưa hấu – SGK trang 69)** *(kết nối Vận dụng 1 trang 74)* và định hướng nhận diện công cụ, lệnh dùng cho **03 nhiệm vụ Thực hành (SGK trang 72–74)** ở Tiết 2.', first=0)
    eng.add(b, '**c) Sản phẩm:** Đề xuất bằng lời của học sinh nêu đủ 3 bước vẽ miếng dưa hấu Hình 13.1 (dùng `Start = 0, End = 180`, thứ tự lớp) và bảng đối chiếu công cụ, lệnh cho 03 bài thực hành Tiết 2.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 5` *(Rubric quan sát vấn đáp Hoạt động Luyện tập & Vận dụng – Phụ lục 3)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd4_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên trình chiếu lại Hình 13.1 (Miếng dưa hấu) và đặt câu hỏi tổng kết mạch bài học:\n'
            '  1. Dựa vào thuộc tính hình tròn và quy tắc thứ tự lớp đã học hôm nay, em hãy đề xuất các bước để vẽ được miếng dưa hấu Hình 13.1?\n'
            '  2. Quan sát lướt tên 03 nhiệm vụ Thực hành trong SGK trang 72–74 (Vẽ logo Hình 13.8, Vẽ hình cây Hình 13.10, Vẽ hình cầu Hình 13.12), hãy cho biết mỗi nhiệm vụ sẽ vận dụng những kiến thức nào của Tiết 1?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh suy nghĩ cá nhân trong 01 phút, liên kết thuộc tính `Start, End`, hộp thoại `Fill and Stroke` và phép ghép `Path` để phác thảo câu trả lời.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 01–02 học sinh phát biểu cách vẽ miếng dưa hấu Hình 13.1 và chỉ ra công cụ/lệnh tương ứng cho 3 nhiệm vụ Thực hành của Tiết 2.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo Tiêu chí 3 của Rubric `CCĐG 5`, thao tác minh hoạ nhanh Miếng dưa hấu trên màn chiếu (30 giây: đặt `Start = 0, End = 180` ở chế độ hình quạt, nhân bản `Ctrl + D` thu nhỏ đổi màu cùi trắng và ruột đỏ, thêm hạt elip đen) để kiểm chứng câu trả lời của học sinh; đồng thời dặn dò học sinh ôn lại lí thuyết Mục 1, 2, 3 và đọc trước 03 nhiệm vụ Thực hành trang 72–74 chuẩn bị cho **Tiết 2 – Thực hành (Tiết PPCT: 20)**.',
            '**1. Đề xuất các bước vẽ Miếng dưa hấu Hình 13.1 (SGK tr. 69, liên hệ Vận dụng 1 tr. 74):**\n'
            '1. Chọn công cụ Hình tròn (`E`), đặt `Start = 0`, `End = 180` (chế độ hình quạt) tạo nửa hình tròn màu xanh lá đậm làm vỏ dưa (lớp dưới cùng).\n'
            '2. Nhân bản đối tượng (`Ctrl + D`), thu nhỏ và tô màu trắng làm cùi; tiếp tục nhân bản, thu nhỏ và tô màu đỏ làm ruột dưa.\n'
            '3. Vẽ các hình elip nhỏ màu đen đặt ở lớp trên cùng làm hạt dưa.\n'
            '**2. Tiền đề chuẩn bị cho Tiết 2 – Thực hành (Tiết PPCT: 20 – SGK tr. 72–74):**\n'
            '- *Nhiệm vụ 1 (Vẽ logo – Hình 13.8):* Áp dụng chế độ cung tròn hở và thuộc tính góc mở của Hình tròn + trang `Stroke style` (tăng `Width`), `Stroke paint`.\n'
            '- *Nhiệm vụ 2 (Vẽ hình cây – Hình 13.10):* Áp dụng phép hiệu (`Path / Difference`) tạo thân cây + phép hợp (`Path / Union`) gộp tán lá và thân cây.\n'
            '- *Nhiệm vụ 3 (Vẽ hình cầu – Hình 13.12):* Áp dụng bỏ viền (`X`) + kiểu tô màu chuyển `Linear gradient` và `Radial gradient` trong trang `Fill`.'
        ]
    ]
    b.append(eng.table(hd4_rows, [4650, 4421], header=True, cell_sz=22))

    # IV. PHỤ LỤC
    eng.add(b, 'IV. PHỤ LỤC (TRÍCH LỤC CĂN CỨ CHƯƠNG TRÌNH, PHÂN PHỐI CHƯƠNG TRÌNH VÀ CÔNG CỤ ĐÁNH GIÁ)', 'heading')

    eng.add(b, 'Phụ lục 1. Trích lục Yêu cầu cần đạt Chương trình GDPT 2018, Thông tư 02/2025/TT-BGDĐT và Phân phối chương trình Khối 10', 'sub')
    can_cu_rows = [
        ['Văn bản căn cứ được viện dẫn ở Mục I', 'Nội dung trích lục nguyên văn áp dụng cho Bài 13 (Tiết 1 – Lí thuyết / Tiết PPCT: 19)'],
        ['**1. Chương trình GDPT 2018 môn Tin học**\n*(Thông tư số 32/2018/TT-BGDĐT, Lớp 10, Chủ đề E. Ứng dụng tin học – Phần mềm thiết kế đồ hoạ, định hướng ICT)*',
         '- **Yêu cầu cần đạt 1:** *"Sử dụng được một số chức năng cơ bản của phần mềm thiết kế đồ hoạ."*\n'
         '- **Yêu cầu cần đạt 2:** *"Tạo được sản phẩm số đơn giản, hữu ích và thực tế như thiết kế logo, tạo banner, topic quảng cáo, băng-rôn, áp phích, poster và thiệp chúc mừng,…"*'],
        ['**2. Khung năng lực số cho người học**\n*(Thông tư số 02/2025/TT-BGDĐT)*',
         '- **Mã năng lực số `[NLS 3.2.NC1.a]` (Tích hợp và tạo lập lại nội dung số):** Hiểu rõ nguyên lí tinh chỉnh thuộc tính, màu sắc, độ trong suốt và kết hợp các đối tượng đồ hoạ vectơ làm cơ sở thực hành tạo lập sản phẩm mới.'],
        ['**3. Kế hoạch dạy học môn Tin học Khối 10**\n*(Phân phối chương trình năm học 2026–2027 của Tổ Tin học – Trường THPT Thuận Hoá)*',
         '- **Tiết PPCT 19 (Học kì I):** `Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT)` – Thời lượng: 01 tiết (45 phút). Nội dung lí thuyết trọng tâm theo SGK Tin học 10 (KNTT, trang 69–72): (1) Các đối tượng hình khối và thuộc tính cơ bản; (2) Thiết lập màu tô, màu vẽ và hộp thoại `Fill and Stroke`; (3) Các phép ghép đối tượng trong bảng chọn `Path`; giải quyết các câu hỏi củng cố (trang 70–72) và định hướng chuẩn bị cho Tiết 2 Thực hành.\n'
         '- **Tiết PPCT 20 (Học kì I):** `Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 2 - TH)` – Thời lượng: 01 tiết (45 phút). Học sinh thực hành các nhiệm vụ trong SGK trang 72–74 (Logo 3 cánh, Hình cây xanh, Hình cầu 3D), phần Luyện tập và Vận dụng để hoàn thiện sản phẩm `.svg`.']
    ]
    b.append(eng.table(can_cu_rows, [3150, 5921], header=True, cell_sz=21))
    eng.add(b, '', first=0)

    eng.add(b, 'Phụ lục 2. Hệ thống câu hỏi vấn đáp và tìm hiểu kiến thức mới (CCĐG 1 – CCĐG 4)', 'sub')
    pl1_rows = [
        ['Công cụ đánh giá', 'Câu hỏi / Yêu cầu thực hiện trên lớp', 'Đáp án chuẩn & Tiêu chí đánh giá'],
        ['**CCĐG 1**\n(Hoạt động 1:\nMở đầu)',
         '- Quan sát Hình 13.1 (SGK tr. 69): Kể tên các đối tượng hình học, màu sắc cấu tạo nên miếng dưa hấu và xác định thứ tự 4 lớp từ dưới lên trên.',
         '- Nêu đúng 3 nửa hình tròn giảm dần kích thước (vỏ xanh dưới cùng, đến cùi trắng, tiếp đến ruột đỏ) và các hạt elip đen ở lớp trên cùng *(Đạt / Chưa đạt)*.'],
        ['**CCĐG 2**\n(Nhiệm vụ 2.1:\nThuộc tính)',
         '- Quan sát Hình 13.2 (tr. 69) và đọc Bảng thuộc tính SGK tr. 69–70: (1) Bốn hình màu đỏ ở Hình 13.2 vẽ từ công cụ nào? (2) Nêu ý nghĩa `W, H`, `Rx, Ry` của hình vuông/chữ nhật; (3) Nêu ý nghĩa `Rx, Ry`, `Start, End` và 3 chế độ vẽ hình tròn/elip; (4) Nêu ý nghĩa `Corners`, `Rounded`, `Spoke ratio`, `Randomized`.',
         '- Nêu đúng công cụ *Hình đa giác, hình sao* và giải thích chính xác ý nghĩa các thuộc tính của 3 nhóm hình khối, trong đó xác định được `Start = 0, End = 180` tạo nửa hình tròn *(Đạt từ 3/4 ý trở lên)*.'],
        ['**CCĐG 3**\n(Nhiệm vụ 2.2:\nFill & Stroke)',
         '- Quan sát Hình 13.3, Hình 13.4 và đọc Mục 2 (SGK tr. 70–71): (1) Nhận xét sự khác nhau của 5 hình trong Hình 13.3; (2) Nêu cách mở và chức năng 3 trang của hộp thoại `Fill and Stroke`; (3) Phân biệt 5 kiểu tô màu, nút `?`, ống hút màu và ý nghĩa `Opacity` / kênh `A`.',
         '- Phân biệt đúng 5 hình trong Hình 13.3, nêu đủ cách mở và 3 trang (`Fill`, `Stroke paint`, `Stroke style`), hiểu đúng `Linear/Radial gradient`, nút `?` và độ trong suốt `Opacity` *(Đạt từ 2/3 ý trở lên)*.'],
        ['**CCĐG 4**\n(Nhiệm vụ 2.3:\nPhép ghép Path)',
         '- Quan sát Hình 13.5, Hình 13.6 (SGK tr. 71–72): (1) Nêu cách xoay, xếp ghép thành trái tim ở Hình 13.5; (2) Gọi tên tiếng Anh, tổ hợp phím tắt và kết quả của 06 phép ghép `Path` trong Hình 13.6; (3) Cho biết đối tượng lớp nào bị cắt trong `Difference`, `Division`, `Cut Path`.',
         '- Nêu đúng cách xoay hình vuông 45 độ và ghép với 2 hình tròn bằng `Path / Union` (H.13.5); phân biệt đúng 06 phép ghép H.13.6 và quy tắc đối tượng lớp dưới bị cắt bởi đối tượng lớp trên *(Đạt)*.']
    ]
    b.append(eng.table(pl1_rows, [1600, 4200, 3271], header=True, cell_sz=21))

    eng.add(b, 'Phụ lục 3. Công cụ đánh giá CCĐG 5 – Rubric quan sát vấn đáp Hoạt động Luyện tập & Vận dụng', 'sub')
    rubric_info = [
        ['THÔNG TIN HÀNH CHÍNH CÔNG CỤ ĐÁNH GIÁ (CCĐG 5)', 'NỘI DUNG CHI TIẾT'],
        ['Tên bài dạy & Tiết PPCT', 'Bài 13. Bổ sung các đối tượng đồ hoạ (Tiết 1 - Lí thuyết / Tiết PPCT: 19)'],
        ['Lớp giảng dạy & Thời gian', 'Lớp 10/3 – Trường THPT Thuận Hoá (Thứ Hai, ngày 16/11/2026)'],
        ['Giáo viên hướng dẫn & SVTT', 'GVHD: Thầy Mai Xuân Văn – Sinh viên thực hiện: Hồ Quang Huy (24S1020030)'],
        ['Hình thức đánh giá', 'Quan sát câu trả lời vấn đáp và biểu quyết trắc nghiệm trực tiếp của học sinh tại lớp']
    ]
    b.append(eng.table(rubric_info, [3700, 5371], header=True, cell_sz=21))
    eng.add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Nội dung yêu cầu kiểm tra (SGK tr. 69–72)', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Tỉ trọng', 'Mức đạt'],
        ['1', 'Củng cố thuộc tính hình khối & hộp thoại Fill and Stroke', 'Trả lời câu hỏi SGK tr. 70 (tham số làm cong đỉnh ngôi sao) và 02 câu trắc nghiệm SGK tr. 71.', '- **Tốt:** Nêu chính xác tham số `Rounded` và chọn đúng cả 2 câu trắc nghiệm (`C. Stroke style`; `B` cây bút chỉnh gradient) kèm giải thích rõ ràng.\n- **Đạt:** Trả lời đúng ít nhất 2/3 câu hỏi (có thể cần gợi ý đối chiếu SGK).\n- **Chưa đạt:** Nhầm lẫn giữa trang `Fill` và `Stroke style` hoặc không xác định được tham số `Rounded`.', '30%', '......'],
        ['2', 'Củng cố phép ghép đối tượng trong bảng chọn Path', 'Trả lời câu hỏi SGK tr. 72 (Hình 13.7): Nêu phép ghép hình và các bước thực hiện để vẽ đám mây.', '- **Tốt:** Xác định đúng phép hợp (`Path / Union`, `Ctrl + +`) và trình bày đủ 3 bước thực hiện (vẽ các hình tròn chồng lên nhau, chọn tất cả, chọn `Path / Union`).\n- **Đạt:** Gọi đúng tên phép hợp `Union` nhưng trình bày các bước thực hiện chưa thật đầy đủ.\n- **Chưa đạt:** Nhầm sang phép hiệu/giao hoặc chưa nêu được các bước.', '35%', '......'],
        ['3', 'Vận dụng vẽ Miếng dưa hấu (H.13.1) & Tiền đề cho Tiết 2 Thực hành', 'Đề xuất 3 bước vẽ Miếng dưa hấu Hình 13.1 và chỉ ra công cụ/lệnh tương ứng cho 3 nhiệm vụ Thực hành Tiết 2 (tr. 72–74).', '- **Tốt:** Nêu mạch lạc 3 bước vẽ miếng dưa hấu (`Start = 0, End = 180` chế độ quạt, nhân bản `Ctrl + D` thu nhỏ đổi màu theo đúng thứ tự lớp, thêm hạt elip đen) và xác định đúng kiến thức nền áp dụng cho Logo (H.13.8), Cây (H.13.10), Hình cầu (H.13.12).\n- **Đạt:** Nêu được ý chính cách cắt nửa hình tròn và thứ tự lớp, nhận biết được phần lớn công cụ cho Tiết 2.\n- **Chưa đạt:** Chưa kết nối được kiến thức Tiết 1 vào bài toán mở đầu và nhiệm vụ Tiết 2.', '35%', '......']
    ]
    b.append(eng.table(rubric_rows, [650, 1550, 1800, 3571, 750, 750], header=True, cell_sz=21))

    # V. NHẬN XÉT VÀ KÝ DUYỆT
    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 15 tháng 11 năm 2026',
        gvhd_name=GV_NAME,
        gvgd_name=GV_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT'
    )

    out_path = DAY_THAT_DIR / '01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx'
    used_figs = [
        'hinh_13_1_2_dua_hau_thuoc_tinh.png',
        'bang_13_1_thuoc_tinh_hinh_khoi.png',
        'hinh_13_3_4_fill_stroke.png',
        'hinh_13_5_ket_hop_doi_tuong.png',
        'hinh_13_6_cac_phep_ghep_hinh.png',
        'hinh_13_7_8_dam_may_logo.png',
    ]
    clean_and_validate_docx(b, out_path, used_figures=used_figs, is_khbd=True)


# ==============================================================================
# FILE 2: KICH BAN LEN LOP CHI TIET - VUI NHON - DE DOC, DE HIEU, DE DAY
# ==============================================================================
def build_kich_ban_vui_nhon():
    b = eng.init_body_with_cover(
        cover_title='CẨM NANG LÊN LỚP CHI TIẾT',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 – LỚP 10/3)',
        class_str='Lớp 10/3 – Phòng H.I.5 – Thứ Hai 16/11/2026',
        schedule_str='Khớp 100% KHBD số 01 (Tiết 3: 09h00 – 09h45)',
        lesson_row_str='Tin học 10 (KNTT) – Bài 13 (Tiết 1 LT – PPCT 19)',
        page1_title='CẨM NANG GIẢNG DẠY & KỊCH BẢN LÊN LỚP CHI TIẾT (45 PHÚT)',
        page1_subtitle='(Tài liệu tự học cấp tốc & cầm tay lên lớp cho Sinh viên Hồ Quang Huy [24S1020030] – Dạy thật Lớp 10/3)',
        page1_class_line='Lịch dạy: Thứ Hai, 16/11/2026 – Tiết 3 (09h00–09h45) – Phòng H.I.5 | GVHD: Thầy Mai Xuân Văn',
        page1_time_line='Khớp 100%: 01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx & bảng demo mau_bai13.svg',
        page1_sched_line='Bối cảnh phòng H.I.5: Có 01 máy tính giáo viên, máy chiếu và bảng phấn',
        page1_mode_line='Chiến thuật dạy: Giải thích bình dân – Bám sát SGK tr. 69–72 – Bấm sẵn mau_bai13.svg – Tạo tiền đề vững cho Tiết 2'
    )

    # MỤC I: CHECKLIST 5 PHÚT TRƯỚC GIỜ VÀ SƠ ĐỒ BẢNG PHẤN
    eng.add(b, 'I. CHECKLIST 5 PHÚT TRƯỚC GIỜ & SƠ ĐỒ BẢNG PHẤN (LÀM NGAY KHI CẮM MÁY CHIẾU)', 'heading')
    eng.add(b, '1. Bốn việc bắt buộc làm trước khi vào tiết (Để lên bục bấm phát ăn ngay, không bị quê)', 'sub')
    for item in [
        '**Bước 1 – Mở sẵn tệp mau_bai13.svg bằng Inkscape:** Mở ngay trước giờ chuông reo để Inkscape tải xong trước (tránh mất 10–15 giây chờ mở phần mềm trước mặt Thầy Văn và cả lớp). Tệp này đã chia sẵn các khu vực thị phạm khớp từng phút giảng trên lớp (lớp khung nền đã được khoá an toàn để bạn quét chuột không bị xê dịch khung).',
        '**Bước 2 – Vẽ nháp thử để Reset công cụ Inkscape:** Inkscape hay nhớ lại chế độ của lần dùng trước. Hãy thử bấm vẽ 1 hình tròn nháp ra ngoài lề trắng xem có bị tàng hình (`Opacity = 0%`) hoặc khuyết hình (`Start/End`) hay không. Nếu bị khuyết, bấm nút **Hình tròn nguyên vẹn (nút ngoài cùng bên phải thanh thuộc tính)**; nếu trong suốt, kéo ô `O:` ở góc dưới trái lên `100`, rồi bấm `Delete` xoá hình nháp.',
        '**Bước 3 – Bật sẵn hộp thoại Fill and Stroke:** Nhấn sẵn `Shift + Ctrl + F` để bảng điều khiển màu hiện cố định bên phải màn hình từ đầu đến cuối tiết.',
        '**Bước 4 – Xếp sẵn cửa sổ:** Mở sẵn cuốn Kịch bản hoặc Slide ảnh SGK cạnh cửa sổ Inkscape đang mở `mau_bai13.svg`, kiểm tra `Alt + Tab` chuyển qua lại mượt trong 1 giây.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Sơ đồ trình bày Bảng phấn (Viết gọn 1 lần, dùng suốt cả tiết)', 'sub')
    board_rows = [
        ['GÓC TRÁI BẢNG (Thứ Hai, 16/11/2026)', 'GIỮA BẢNG (Nội dung HS ghi vào vở)', 'GÓC PHẢI BẢNG (Nháp & Mẹo nhớ)'],
        [
            'Tiết 19: BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ\n'
            '-----------------\n'
            'Mục tiêu hôm nay:\n'
            '1. Thuộc tính hình khối\n'
            '2. Tô màu & nét vẽ (Fill & Stroke)\n'
            '3. Phép ghép/cắt hình (Path)',
            '**1. Đối tượng hình khối & thuộc tính (tr. 69–70)**\n'
            '- Chữ nhật: `W, H`, `Rx, Ry` (bo góc)\n'
            '- Tròn: `Start, End` (`0 và 180`: nửa tròn), 3 chế độ vẽ\n'
            '- Sao/Đa giác: `Corners` (cánh), `Rounded` (cong), `Spoke ratio`, `Randomized`\n\n'
            '**2. Màu tô, màu vẽ – Fill and Stroke (Shift+Ctrl+F)**\n'
            '- `Fill`: Màu nền (`X`, đồng nhất, `Linear/Radial gradient`, `Pattern`, `?`)\n'
            '- `Stroke paint` / `Stroke style`: Màu viền (ống hút) / Kiểu nét\n'
            '- `Opacity` / kênh `A`: Độ trong suốt\n\n'
            '**3. Các phép ghép hình trong bảng chọn Path (tr. 71–72)**\n'
            '- `Union` (`Ctrl + +`): Phép hợp\n'
            '- `Difference` (`Ctrl + -`): Phép hiệu (Dưới bị Trên cắt)\n'
            '- `Intersection` (`Ctrl + *`): Phép giao\n'
            '- `Exclusion`, `Division`, `Cut Path` (SGK tr. 72)',
            '**Mẹo nhớ Bài 12:**\n'
            '- Vẽ trước nằm DƯỚI\n'
            '- Vẽ sau nằm TRÊN\n\n'
            '**Đáp án Luyện tập (tr. 70–72):**\n'
            '- Câu 1 (tr. 70): `Rounded`\n'
            '- Câu 2 (tr. 71): `C. Stroke style`\n'
            '- Câu 3 (tr. 71): `B` (cây bút)\n'
            '- Câu 4 (tr. 72): `Path / Union`\n\n'
            '**Quy tắc cắt Path:**\n'
            'HÌNH DƯỚI bị HÌNH TRÊN cắt!'
        ]
    ]
    b.append(eng.table(board_rows, [2150, 4671, 2250], header=True, cell_sz=21))

    # MỤC II: HỆ THỐNG KIẾN THỨC CỐT LÕI TỪ CON SỐ 0 CHO NGƯỜI CHƯA NẮM VỮNG
    eng.add(b, 'II. HỆ THỐNG KIẾN THỨC CỐT LÕI BÀI 13 DỄ HIỂU (ĐỌC 5 PHÚT LÀ NẮM CHẮC)', 'heading')
    eng.add(b, 'Phần này diễn giải toàn bộ thuật ngữ tiếng Anh trong SGK Tin học 10 (tr. 69–74) sang tiếng Việt dễ hiểu, giúp bạn tự tin giải thích trôi chảy dù học sinh hay Thầy Mai Xuân Văn có hỏi thêm chi tiết nào:', first=0)

    eng.add(b, '1. Thuộc tính hình khối (Mục 1 SGK tr. 69–70): Khác nhau giữa Hình vuông, Hình tròn, Ngôi sao', 'sub')
    for item in [
        '**Mẹo nhớ vị trí:** Nháy đúp chuột vào bất kỳ hình nào (hoặc chọn công cụ tương ứng ở thanh bên trái) thì thanh ngang trên cùng (**Thanh điều khiển thuộc tính** của Bài 12) sẽ lập tức hiện ra các ô thông số của hình đó.',
        '**Hoạt động 1 (Hình 13.2 tr. 69):** Cả 4 hình màu đỏ đều vẽ từ công cụ **Hình đa giác, hình sao**: (1) ngôi sao 5 cánh nhọn (`Corners = 5`); (2) chuyển chế độ đa giác đều 5 cạnh (`Corners = 5`); (3) tăng `Rounded` để làm cong đỉnh; (4) tăng `Randomized` để làm méo hình ngẫu nhiên.',
        '**Hình vuông / Hình chữ nhật (phím R):** `W, H` là chiều rộng, chiều dài; **`Rx, Ry`** là **bán kính của góc bo** (`Rx = Ry = 0` thì góc hình là góc vuông; tăng `Rx, Ry` lên thì góc bo tròn mềm mại).',
        '**Hình tròn / Hình elip (phím E):** `Rx, Ry` là bán kính theo phương ngang và thẳng đứng. Có **Start và End** (góc điểm đầu và điểm cuối tính theo độ, ví dụ `0` và `180` là cắt đúng một nửa hình tròn) cùng **3 nút biểu tượng hình tròn** ở cuối thanh thuộc tính: **Hình quạt (Slice)** nối 2 đầu vào tâm như miếng dưa hấu; **Cung tròn hở (Arc)** chỉ giữ nét cong; **Hình tròn khép kín (Whole)** đóng tròn nguyên vẹn.',
        '**Hình đa giác / Hình sao (phím *):** Có 2 nút chuyển chế độ đa giác đều hoặc ngôi sao nhọn; **`Corners`** là số đỉnh đa giác hoặc số cánh hình sao; **`Spoke ratio`** là tỉ lệ bán kính từ tâm đến góc trong và từ tâm tới đỉnh đầu hình sao; **`Rounded`** là độ cong tại các đỉnh (đây là đáp án câu hỏi tr. 70); **`Randomized`** là tham số làm méo hình ngẫu nhiên.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Hộp thoại Fill and Stroke (Mục 2 SGK tr. 70–71): Phân biệt 3 trang & 5 kiểu tô màu', 'sub')
    for item in [
        '**Cách gọi hộp thoại (theo SGK tr. 70):** Chọn đối tượng, vào **`Object / Fill and Stroke`**, hoặc **nháy nút phải chuột vào hình chọn** `Fill and Stroke`, hoặc nhấn **`Shift + Ctrl + F`**. Nhớ phân biệt: **`Fill`** là màu tô **bên trong nền**, còn **`Stroke`** là màu vẽ **đường viền bên ngoài**.',
        '**Ba trang (Tab) trong hộp thoại:** Trang 1 **`Fill`** (màu tô nền); Trang 2 **`Stroke paint`** (chọn màu cho đường viền — có biểu tượng **ống hút màu** phía dưới để sao chép màu ở vị trí bất kỳ); Trang 3 **`Stroke style`** (chỉnh độ dày mỏng `Width` và kiểu nét vẽ đứt/liền — đây là đáp án Câu 1 tr. 71).',
        '**Các biểu tượng kiểu tô màu trong Hình 13.4:** Dấu **`X`** (không màu/trong suốt); **Ô vuông đặc** (màu đồng nhất); **Hai ô vuông chuyển** (`Linear gradient` – màu chuyển giữa hai hoặc nhiều màu theo chiều thẳng); **Ô vuông sáng giữa tâm** (`Radial gradient` – màu chuyển từ tâm của đối tượng ra ngoài giúp tạo cảm giác khối cầu 3D); **Ô bàn cờ** (`Pattern` – hoa văn); Dấu **`?`** (huỷ đặt đối tượng trở về trạng thái ban đầu, rất tiện khi nhân bản đối tượng).',
        '**Độ trong suốt (Opacity và kênh A – alpha):** Kéo giảm từ `100%` xuống `50% – 80%` thì đối tượng trở nên **trong suốt một phần** (giống tấm kính mờ), nhìn xuyên qua được các hình nằm lớp bên dưới.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '3. Trọn bộ 06 phép ghép đối tượng trong bảng chọn Path (Mục 3 SGK tr. 71–72)', 'sub')
    eng.add(b, 'Trong Hình 13.6 SGK trang 72 luôn giả định có **Hình vuông nằm lớp DƯỚI** và **Hình tròn nằm lớp TRÊN**. Khi chọn từ 2 hình trở lên rồi mở menu **`Path`**, hãy nhớ ý nghĩa trực quan của cả 6 lệnh sau (trong đó 3 lệnh đầu là trọng tâm nhất):', first=0, keep=True)
    path_rows = [
        ['STT & Tên lệnh trong Path', 'Phím tắt', 'Giải thích nhanh & Kết quả (Hình vuông DƯỚI, Hình tròn TRÊN)', 'Ứng dụng & Câu hỏi SGK'],
        ['1. **Union** *(Phép hợp)*', '`Ctrl + +`', '**Gộp làm một**: Gộp chung cả hình vuông và hình tròn thành một khối duy nhất (giống tập hợp A ∪ B).', 'Ghép Trái tim (H.13.5), Đám mây (H.13.7), Tán cây (H.13.11).'],
        ['2. **Difference** *(Phép hiệu)*', '`Ctrl + -`', '**Khoét khuyết (cắt bỏ)**: Hình **DƯỚI** (vuông) bị hình **TRÊN** (tròn) khoét mất phần đè nhau, hình tròn biến mất (giống A \\ B).', 'Khoét Logo Apple "cắn dở", gọt thân cây thắt eo (H.13.11), chân Quân Bích (H.13.15).'],
        ['3. **Intersection** *(Phép giao)*', '`Ctrl + *`', '**Lấy phần giao nhau**: Chỉ giữ lại đúng phần chồng lên nhau ở giữa hai hình (giống tập hợp A ∩ B).', 'Vẽ chiếc lá trên đầu táo Apple từ 2 hình tròn giao nhau.'],
        ['4. **Exclusion** *(Hiệu đối xứng)*', '`Ctrl + ^`', '**Bỏ phần giao nhau ở giữa**: Giữ hai phần rìa ngoài nhưng khoét thủng phần giao nhau ở giữa.', 'Tạo chữ lồng, khung viền rỗng giữa khi hai hình đè nhau.'],
        ['5. **Division** *(Phép chia)*', '`Ctrl + /`', '**Chia nhỏ thành nhiều mảnh rời**: Lấy mép hình **TRÊN** cắt hình **DƯỚI** thành các mảnh rời nhau nhưng **vẫn giữ màu tô**, kéo tách từng mảnh ra được!', 'Tách 1 hình tròn thành nhiều miếng màu ghép khác nhau.'],
        ['6. **Cut Path** *(Phép cắt đường)*', '`Ctrl + Alt + /`', '**Cắt đứt đường viền**: Chỉ cắt đứt đường viền hình **DƯỚI** tại điểm giao với hình **TRÊN**, **mất hết màu tô nền bên trong**.', 'Cắt lấy một đoạn cung viền hở để chỉnh đường nét.']
    ]
    b.append(eng.table(path_rows, [2100, 1250, 3421, 2300], header=True, cell_sz=21))

    eng.add(b, '4. Nắm nhanh 03 bài Thực hành SGK trang 72–74 (Để tự tin dặn dò chuẩn bị Tiết 2)', 'sub')
    for item in [
        '**Nhiệm vụ 1 – Logo 3 cánh (H.13.8–13.9, tr. 72–73):** Vẽ hình tròn, kéo nút tạo cung hở và bấm nút **Arc**, sau đó mở thẻ `Stroke style` tăng độ dày `Width`, nhấn `Ctrl + D` nhân 3 bản xoay ghép khít rồi tô `Stroke paint` màu đỏ đậm.',
        '**Nhiệm vụ 2 – Hình cây xanh (H.13.10–13.11, tr. 73):** Xếp 3 hình tròn làm tán, đặt hình chữ nhật (dưới) và 2 hình tròn (trên), vào `Path` chọn `Difference` **từng cặp 2 hình một** để khoét eo thân cây, rồi quét cả tán và thân vào `Path` chọn `Union` và tô màu xanh lá.',
        '**Nhiệm vụ 3 – Hình cầu 3D (H.13.12–13.13, tr. 73–74):** Tắt viền cả 3 hình (`Stroke = X`); vẽ chữ nhật nền vàng, hình tròn to đen tô **`Linear gradient`** chéo sáng tối và hình tròn nhỏ trắng lệch góc trên phải tô **`Radial gradient`** tạo quầng sáng phản chiếu.'
    ]:
        eng.add(b, item, 'bullet')

    # MỤC III: LIÊN HỆ BÀI CŨ & MỞ RỘNG THỰC TẾ + AI
    eng.add(b, 'III. LIÊN HỆ BÀI CŨ SGK & ỨNG DỤNG THỰC TẾ (CHÈN KHI GIẢNG)', 'heading')
    rel_rows = [
        ['Chặng giảng', 'Liên hệ Bài cũ trong SGK Tin 10 (Nói 20 giây cho Thầy Văn chấm điểm)', 'Ví dụ gần gũi & Liên hệ AI / Thực tế bên ngoài'],
        ['**Chặng 1:**\nMở đầu',
         '- **Bài 11 (Bản quyền phần mềm):** Giải thích ngắn gọn vì sao dùng **Inkscape**: vì Inkscape là phần mềm nguồn mở miễn phí, tôn trọng bản quyền.\n- **Bài 12 (Thứ tự lớp):** Hình vẽ trước nằm dưới, vẽ sau nằm trên.',
         '- Liên hệ thứ tự lớp bằng **xếp chồng các tờ giấy màu thủ công**: Tờ đặt trước nằm dưới (lớp 1), các tờ đặt sau đè lên trên! Muốn thấy hạt dưa trên miếng dưa hấu thì hạt phải nằm ở lớp trên cùng.'],
        ['**Chặng 2A:**\nThuộc tính',
         '- **Bài 12 (Màn hình Inkscape H.12.3):** Nhắc lại tên thanh ngang ngay dưới thanh bảng chọn chính là **Thanh điều khiển thuộc tính** — chọn công cụ nào là thanh đó hiển thị thông số tương ứng.',
         '- So sánh `Rx, Ry` giữa **điện thoại bo góc mềm** và **khối gạch vuông góc 90 độ**.\n- Giải thích `Rounded`: điều chỉnh độ cong tại các đỉnh nhọn để biến ngôi sao thành bông hoa cánh tròn.'],
        ['**Chặng 2B:**\nFill & Stroke',
         '- **Bài 12 (Dải màu dưới đáy):** Nháy chuột trái = đổi màu tô (`Fill`); giữ `Shift` + nháy chuột = đổi màu viền (`Stroke`). Nay mở `Shift + Ctrl + F` để chỉnh màu chuyển 3D và nét đứt.',
         '- `Fill` là **màu nền bên trong**, `Stroke` là **đường viền bao quanh**.\n- `Linear gradient`: chuyển màu dần theo phương thẳng.\n- `Radial gradient`: chuyển màu toả tròn từ tâm tạo khối cầu 3D.\n- `Opacity`: chỉnh độ mờ trong suốt giống tấm kính mờ (tương tự trên **Canva, Figma**).'],
        ['**Chặng 2C:**\nPhép ghép Path & AI',
         '- **Toán lớp 10 (Tập hợp)** & **Bài 12 (lớp trên/dưới)**: `Union` = `A ∪ B`, `Intersection` = `A ∩ B`, `Difference` = `A \\ B`.\n- Mở đường sang **Bài 14**: Bấm lệnh `Path` biến hình khối thành đối tượng đường.',
         '- **Cách tạo Logo Apple:** Quả táo nguyên + hình tròn lệch phải + `Path / Difference` để khoét phần khuyết; chiếc lá trên đỉnh = 2 hình tròn + `Intersection`.\n- **Cắt Laser / Mica:** Máy CNC đọc trực tiếp đường `Path` trong file `.svg`.\n- **AI ảnh vs AI Vector:** Ảnh `.png` từ AI phóng to dễ vỡ nét, còn mã `.svg` đưa vào Inkscape vẫn mở từng lớp chỉnh sửa dễ dàng.']
    ]
    b.append(eng.table(rel_rows, [1450, 3200, 4421], header=True, cell_sz=21))

    # MỤC IV: KỊCH BẢN TỪNG CHẶNG (2 CỘT RỘNG RÃI + HÌNH SGK NGAY TRÊN MỖI CHẶNG)
    eng.add(b, 'IV. KỊCH BẢN LÊN LỚP TỪNG CHẶNG: NÓI GÌ – BẤM GÌ TRÊN `mau_bai13.svg` – XỬ LÝ TÌNH HUỐNG (45 PHÚT)', 'heading')
    eng.add(b, 'Nguyên tắc lên lớp: Giữ nhịp giảng tự nhiên, rõ ràng, bám đúng quy trình "đặt câu hỏi để học sinh tra trực tiếp trong SGK – thao tác trực tiếp trên tệp mau_bai13.svg cho cả lớp quan sát", giúp lớp theo kịp bài và giáo viên dự giờ đánh giá cao tính sư phạm.', first=0)

    # --- Chặng 1 ---
    eng.add(b, 'Chặng 1 (Phút 00–05): Mở đầu – Phân tích Miếng dưa hấu (Hình 13.1 SGK tr. 69)', 'sub')
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    c1_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Chào cả lớp! Hôm nay chúng ta học Bài 13: Bổ sung các đối tượng đồ hoạ. Trước hết, các em mở SGK Tin học 10 trang 69 và quan sát Hình 13.1."*\n'
            '- *"Nhìn vào hình miếng dưa hấu ở Hình 13.1, bạn nào nêu giúp thầy nó được ghép từ những lớp màu và hình cơ bản nào?"*\n'
            '- *(HS trả lời xong)*: *"Chính xác! Bây giờ nhắc lại quy tắc ở Bài 12: hình vẽ trước nằm dưới, vẽ sau nằm trên, giống như xếp chồng các tờ giấy màu. Vậy giữa vỏ xanh, cùi trắng, ruột đỏ và hạt đen, phần nào nằm dưới cùng và phần nào nằm trên cùng?"*\n'
            '- *(Chốt vào bài)*: *"Ở Bài 12 mình mới vẽ được hình tròn nguyên vẹn, chưa biết cắt thành nửa hình tròn, chưa biết tô màu chuyển khối hay cắt ghép các hình lại với nhau. Bài 13 hôm nay sẽ giải quyết cả 3 nội dung đó."*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 1 bên trái):** Kéo tách nhẹ từng lớp của miếng dưa hấu mẫu (vỏ xanh, cùi trắng, ruột đỏ, hạt đen) để lớp thấy rõ thứ tự lớp dưới/trên, sau đó nhấn `Ctrl + Z` trả về vị trí cũ. Ghi tựa Bài 13 lên bảng.\n'
            '- **Xử lý nếu lớp im lặng:** Hỏi gợi mở trực tiếp: *"Nếu đặt lớp hạt đen nằm dưới lớp ruột đỏ thì khi nhìn từ trên xuống mình có thấy được hạt dưa nữa không?"* Khi đó cả lớp sẽ dễ dàng trả lời được là hạt phải nằm trên cùng.'
        ]
    ]
    b.append(eng.table(c1_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2A ---
    eng.add(b, 'Chặng 2A (Phút 05–14): Nhiệm vụ 1 – Các đối tượng hình khối & thuộc tính (SGK tr. 69–70)', 'sub')
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    c2a_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Mời cả lớp nhìn Hình 13.2 trang 69: gồm ngôi sao nhọn, ngũ giác đều, ngôi sao cánh tròn và hình méo ngẫu nhiên. Cả 4 hình đều được tạo từ cùng một công cụ — nhìn vào biểu tượng ngôi sao ở ô cuối bảng đầu trang 70, một bạn đọc giúp thầy tên công cụ đó là gì?"* (Công cụ **Hình đa giác, hình sao** — từ trái qua: sao 5 cánh, chuyển chế độ đa giác, tăng `Rounded`, tăng `Randomized`).\n'
            '- *"Bây giờ hai bạn cùng bàn đọc Bảng thuộc tính trang 69–70 trong 2 phút và gạch chân các từ khoá trả lời 3 ý sau:*\n'
            '  *1. Hình vuông, hình chữ nhật có những thuộc tính nào để chỉnh kích thước và bo tròn 4 góc?* (`W, H` và `Rx, Ry`)\n'
            '  *2. Hình tròn, hình elip có thuộc tính Start, End dùng để làm gì và chọn giá trị nào để cắt thành nửa hình tròn?* (`0` và `180`, cùng 3 chế độ quạt/cung tròn/hình tròn kín)\n'
            '  *3. Hình đa giác, hình sao có 4 thuộc tính nào?* (`Corners`, `Rounded`, `Spoke ratio`, `Randomized`)\n'
            '- *"Bây giờ các em nhìn lên máy chiếu để xem kết quả khi thay đổi các thông số này."*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 2):**\n'
            '  1. Nháy đúp vào **Hình chữ nhật xanh**, gõ `Rx = 25` rồi nhấn `Enter`: quan sát 4 góc bo tròn.\n'
            '  2. Nháy đúp vào **Hình tròn cam**, gõ `Start = 0`, `End = 180` rồi nhấn `Enter`: thu được nửa hình tròn (chỉ 3 nút hình quạt, cung tròn hở, hình tròn kín).\n'
            '  3. Nháy đúp vào **Ngôi sao đỏ**, sửa `Rounded = 0.25` rồi nhấn `Enter`: các đỉnh nhọn uốn thành cánh hoa tròn.\n'
            '- **Xử lý nếu HS im lặng hoặc ngại đọc tiếng Anh:** Chỉ rõ *"Các em nhìn đúng cột Thuộc tính bên trái dòng mô tả ở trang 69–70 và đọc từng chữ là được"*, không bắt chấm phát âm tiếng Anh.'
        ]
    ]
    b.append(eng.table(c2a_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2B ---
    eng.add(b, 'Chặng 2B (Phút 14–23): Nhiệm vụ 2 – Màu tô, màu vẽ & Fill and Stroke (SGK tr. 70–71)', 'sub')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    c2b_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Ở Bài 12 muốn đổi màu nhanh thì nháy vào dải màu dưới đáy. Muốn tô màu chuyển hoặc đổi kiểu nét viền, ta chọn Object / Fill and Stroke, hoặc nháy nút phải chuột chọn Fill and Stroke, hay nhấn Shift + Ctrl + F."*\n'
            '- *"Nhìn 5 ô vuông ở Hình 13.3 trang 70 và đối chiếu 5 dòng giải thích bên dưới hình: một bạn mô tả giúp thầy phần màu tô bên trong của 5 hình này khác nhau thế nào?"* (Ô 1: không màu; Ô 2: đồng nhất; Ô 3: chuyển màu theo chiều ngang; Ô 4: chuyển màu từ tâm; Ô 5: chuyển màu kết hợp độ trong suốt).\n'
            '- *"Nhìn Hình 13.4, các em phân biệt 3 trang: Fill là màu tô nền (có nút ? huỷ đặt về trạng thái ban đầu), Stroke paint là màu đường viền (có công cụ ống hút sao chép màu), Stroke style là độ dày và kiểu nét vẽ. Ngoài ra có biểu tượng cây bút chỉnh thông số gradient và giá trị Opacity / kênh A chỉnh độ trong suốt."*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 3):**\n'
            '  + Hộp thoại `Fill and Stroke` đã mở sẵn bên phải: bấm vào **Hình tròn xanh**, chọn tab `Fill`, bấm nút **`Radial gradient`** (nút thứ 4) cho lớp thấy hiệu ứng chuyển màu toả tròn 3D (chỉ thêm công cụ cây bút chỉnh điểm dừng màu).\n'
            '  + Chuyển tab `Stroke style`, bấm `Dashes` đổi viền sang nét đứt.\n'
            '  + Bấm **Hình tròn xanh dương** đè trên chữ nhật vàng, kéo thanh `Opacity` xuống khoảng `50%` cho HS thấy hiệu ứng nhìn xuyên qua.\n'
            '- **Xử lý nếu HS im lặng:** Hỏi từng câu ngắn theo 3 ý: *Trang nào chỉnh nền? Trang nào chỉnh màu viền? Trang nào chỉnh độ dày và kiểu nét?*'
        ]
    ]
    b.append(eng.table(c2b_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2C ---
    eng.add(b, 'Chặng 2C (Phút 23–33): Nhiệm vụ 3 – Các phép ghép hình trong Path (SGK tr. 71–72)', 'sub')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    c2c_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Nhìn Hình 13.5 trang 71: để xếp 3 mẫu giấy thành trái tim, ta xoay hình vuông nghiêng 45 độ, đặt 2 hình tròn chồng lên hai cạnh trên rồi dùng phép hợp trong menu Path."*\n'
            '- Chiếu Hình 13.6 (tr. 72): *"Quy ước hình vuông nằm lớp dưới, hình tròn nằm lớp trên. Các em nhìn nhãn tiếng Việt in sẵn dưới từng hình ở Hình 13.6 và đọc chữ in đậm ở trang 71–72 để phân biệt 06 phép ghép trong menu Path:"*\n'
            '  1. **Phép hợp** (`Union`, `Ctrl + +`): Gộp chung cả hai hình thành một.\n'
            '  2. **Phép hiệu** (`Difference`, `Ctrl + -`): Lấy phần hình dưới không thuộc hình trên.\n'
            '  3. **Phép giao** (`Intersection`, `Ctrl + *`): Chỉ giữ phần chung của hai hình.\n'
            '  4. **Phép hiệu đối xứng** (`Exclusion`, `Ctrl + ^`): Bỏ phần giao ở giữa.\n'
            '  5. **Phép chia** (`Division`, `Ctrl + /`): Chia hình dưới thành nhiều mảnh bởi đường nét hình trên.\n'
            '  6. **Phép cắt** (`Cut Path`, `Ctrl + Alt + /`): Cắt đường viền hình dưới tại giao điểm, kết quả không có màu tô.\n'
            '- Hỏi chốt quy tắc: *"Trong phép hiệu, phép chia và phép cắt, đối tượng lớp nào bị cắt?"* (**Lớp dưới bị lớp trên cắt**).',
            '- **Thao tác AN TOÀN TUYỆT ĐỐI (mau_bai13.svg – Khu vực 4):**\n'
            '  + **B & C (Cắt Vuông-Tròn & Logo Apple):** Quét chọn hình vuông đỏ với hình tròn cam, thử lần lượt chọn menu **Path** rồi chọn **Intersection** (hoặc **Difference** trên quả táo đen + hình tròn xám ở ô C) để cả lớp quan sát trực tiếp kết quả không cần dùng phím tắt.\n'
            '- **Xử lý tình huống:** Nếu bấm `Difference` mà thấy mất hình gốc, bình tĩnh bấm `Ctrl + Z` và nhắc lại quy tắc vàng: *"Hình dưới luôn bị hình trên cắt; chọn hình muốn giữ nhấn phím End đưa xuống dưới trước khi cắt."*'
        ]
    ]
    b.append(eng.table(c2c_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 3 ---
    eng.add(b, 'Chặng 3 (Phút 33–40): Luyện tập – Giải 04 câu hỏi củng cố SGK (tr. 70, 71, 72 & Hình 13.7)', 'sub')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    c3_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Bây giờ chúng ta luyện tập củng cố lí thuyết qua đúng 04 câu hỏi trong SGK từ trang 70 đến trang 72:"*\n'
            '  + *"Câu 1 (SGK tr. 70): Nếu trên hình vẽ có sẵn một hình sao 5 cánh nhọn, em cần thay đổi tham số nào để các đỉnh ngôi sao trở nên cong?"* (`Rounded`).\n'
            '  + *"Câu 2 (SGK tr. 71): Để xác định đường viền của đối tượng dạng nét đứt, chọn trang nào trong Fill and Stroke?"* (Cả lớp giơ 3 ngón: **`C. Stroke style`**).\n'
            '  + *"Câu 3 (SGK tr. 71): Để chỉnh thông số của gradient, chọn biểu tượng nào trong 4 hình A, B, C, D?"* (Cả lớp giơ 2 ngón: **`B`** – biểu tượng cây bút).\n'
            '  + *"Câu 4 (SGK tr. 72 – Hình 13.7): Thảo luận cặp đôi 2 phút, nêu tên phép ghép hình và các bước thực hiện để vẽ đám mây như Hình 13.7?"* (**Phép hợp** `Path / Union`, gồm 3 bước: vẽ các hình tròn chồng lấp lên nhau, chọn tất cả các hình, chọn lệnh `Path / Union`).',
            '- **Thao tác kiểm chứng trên mau_bai13.svg:**\n'
            '  + Ở Câu 1–3: Sau khi lớp trả lời, nháy chuột vào ô ngôi sao (Khu vực 2) hoặc đường viền nét đứt (Khu vực 3) để kiểm chứng.\n'
            '  + Ở Câu 4 (**Khu vực 4A**): Quét chọn 4 hình xanh ở ô A, mở menu **Path rồi bấm Union** (KHÔNG bấm phím tắt `Ctrl + +` để tránh kẹt Shift) cho cả lớp thấy đám mây hoàn chỉnh hiện ra ngay.\n'
            '- **Xử lý nếu HS im lặng:** Hướng dẫn học sinh mở lại đúng trang 70, 71, 72 đối chiếu bảng biểu và hình vừa học để trả lời chính xác.'
        ]
    ]
    b.append(eng.table(c3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 4 ---
    eng.add(b, 'Chặng 4 (Phút 40–45): Vận dụng – Miếng dưa hấu Hình 13.1 & Tiền đề chuẩn bị cho Tiết 2 Thực hành', 'sub')
    c4_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Quay lại bài toán Mở đầu đầu giờ — Miếng dưa hấu Hình 13.1 (trang 69): Sau khi học xong Tiết 1, em hãy nêu 3 bước ngắn gọn để vẽ được miếng dưa hấu đó?"*\n'
            '  1. *"Tạo nửa hình tròn vỏ xanh bằng thông số gì?"* (`Start = 0, End = 180`, chế độ hình quạt `Slice`).\n'
            '  2. *"Nhân bản thu nhỏ thành cùi trắng và ruột đỏ thế nào?"* (`Ctrl + D` nhân bản, thu nhỏ và tô trắng, rồi tiếp tục `Ctrl + D` thu nhỏ tô đỏ).\n'
            '  3. *"Hạt dưa đặt ở đâu?"* (Vẽ elip nhỏ màu đen ở lớp trên cùng).\n'
            '- *(Liên hệ làm tiền đề cho Tiết 2 – Thực hành trang 72–74)*: *"Để Tiết 2 thực hành nhanh, các em nhìn lướt trang 72–74 xem 3 bài Thực hành dùng lại kiến thức nào hôm nay:"*\n'
            '  + *Logo 3 cánh (Hình 13.8):* dùng **cung tròn hở** + trang **`Stroke style`** (tăng `Width`).\n'
            '  + *Hình cây (Hình 13.10):* dùng **`Path / Difference`** (tạo thân) + **`Path / Union`** (ghép tán và thân).\n'
            '  + *Hình cầu (Hình 13.12):* bỏ viền (`X`) + tô **`Linear gradient`** và **`Radial gradient`**.\n'
            '- *(Dặn dò)*: *"Về nhà các em ôn kĩ 3 nội dung lí thuyết Tiết 1 và đọc trước 3 nhiệm vụ Thực hành trang 72–74 để Tiết 2 thực hành thuận lợi. Em cảm ơn thầy và chào cả lớp!"*',
            '- **Thao tác minh hoạ (mau_bai13.svg – Khu vực 1 bên phải):** Chọn hình tròn xanh chờ sẵn bên phải Khu vực 1: nhấn `E`, chỉnh `Start = 0, End = 180`, bấm `Ctrl + D` thu nhỏ tô trắng, tiếp tục `Ctrl + D` thu nhỏ tô đỏ, thêm 2 hạt elip đen trong khoảng 25 giây để minh hoạ kết quả trọn vẹn.\n'
            '- *(Nếu muốn cho HS xem trước mục tiêu Tiết 2)*: Chỉ nhẹ qua hình cầu 3D minh hoạ ở góc dưới phải màn hình để tạo hứng thú chuẩn bị cho Tiết 2.\n'
            '- **Lưu ý giờ:** Nếu còn dưới 2 phút, chỉ minh hoạ cắt nửa hình tròn và `Ctrl + D` 1 lần tô đỏ rồi dặn dò kết thúc đúng giờ.'
        ]
    ]
    b.append(eng.table(c4_rows, [4650, 4421], header=True, cell_sz=22))

    # MỤC V: BẢNG CỨU HỘ KỸ THUẬT 3 GIÂY
    eng.add(b, 'V. BẢNG CỨU HỘ KỸ THUẬT 3 GIÂY KHI ĐỨNG MÁY TRÊN BỤC GIẢNG', 'heading')
    rescue_rows = [
        ['Sự cố bất ngờ trên máy chiếu', 'Nguyên nhân', 'Cách xử lý trong 3 giây (Không hoảng!)'],
        ['Vẽ hình ra bị **tàng hình** (chỉ thấy khung chấm nét đứt, không thấy màu)', 'Inkscape nhớ lại lần trước có ai đặt `Opacity = 0%` hoặc bỏ màu nền `Fill = X` & viền `Stroke = X`.', 'Nhìn góc dưới cùng bên trái màn hình chỉnh `O: 100` (hoặc kéo `Opacity` lên 100%) + nháy chuột trái vào một ô màu xanh/đỏ trên Bảng màu dưới đáy.'],
        ['Vẽ hình tròn ra chỉ thấy **nửa hình / miếng bánh** chứ không tròn kín', 'Inkscape đang nhớ góc `Start/End` của lần vẽ trước.', 'Nhìn lên cuối Thanh điều khiển thuộc tính phía trên, bấm vào **biểu tượng hình tròn khép kín (nút ngoài cùng bên phải trong 3 nút tròn)** là tròn lại ngay!'],
        ['Bấm `Corners / Start / End` không thấy ô nhập đâu cả', 'Đang chọn công cụ `Select` (mũi tên đen) nên không hiện thuộc tính riêng của hình.', 'Nháy **đúp chuột** thẳng vào hình đó (hoặc bấm lại biểu tượng hình tròn `E` / hình sao `*` ở thanh công cụ bên trái).'],
        ['Bấm `Path / Difference` mà mất tiêu cả hình gốc hoặc không có tác dụng', 'Chọn nhầm **3 hình cùng lúc** (chỉ được chọn đúng 2 hình) hoặc **hình cần giữ đang nằm ở lớp trên**.', 'Bấm `Ctrl + Z` (hoàn tác). Nháy chọn hình cần giữ bấm phím `End` (đưa xuống dưới cùng). Rồi giữ `Shift` chọn **từng cặp 2 hình** một và mở menu `Path`, chọn `Difference`.'],
        ['Bấm phím tắt `Ctrl + +` trên bàn phím laptop không ăn', 'Do dấu `+` trên laptop phải bấm chung phím `Shift` nên hay bị kẹt.', 'Cứ mở chuột lên menu **Path rồi bấm chữ Union** (vừa chắc ăn 100%, vừa giúp học sinh dưới lớp nhìn thấy tên lệnh trên máy chiếu!).'],
        ['Lỡ bấm nhầm kéo lệch khung bảng của `mau_bai13.svg`', 'Hiếm gặp vì Layer 1 đã khoá; có thể do mở khoá nhầm.', 'Bấm `Ctrl + Z` ngay lập tức (hoặc đóng file không lưu rồi mở lại `mau_bai13.svg`).']
    ]
    b.append(eng.table(rescue_rows, [2350, 2700, 4021], header=True, cell_sz=21))

    out_path = DAY_THAT_DIR / '02_Kich_ban_len_lop_chi_tiet_vui_nhon_Tin10_Bai13_Tiet1.docx'
    used_figs = [
        'hinh_13_1_2_dua_hau_thuoc_tinh.png',
        'bang_13_1_thuoc_tinh_hinh_khoi.png',
        'hinh_13_3_4_fill_stroke.png',
        'hinh_13_5_ket_hop_doi_tuong.png',
        'hinh_13_6_cac_phep_ghep_hinh.png',
        'hinh_13_7_8_dam_may_logo.png',
    ]
    clean_and_validate_docx(b, out_path, used_figures=used_figs, is_khbd=False)


# ==============================================================================
# FILE 3: BAI THUYET TRINH GIANG DAY LIEN MACH 45 PHUT (PHIEN BAN DAY HOC)
# ==============================================================================
def build_bai_thuyet_trinh_day_hoc():
    b = eng.init_body_with_cover(
        cover_title='KỊCH BẢN THUYẾT TRÌNH GIẢNG DẠY (45 PHÚT)',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 – LỚP 10/3)',
        class_str='Lớp 10/3 – Phòng H.I.5 – Thứ Hai 16/11/2026',
        schedule_str='Tiết 3 (09h00 – 09h45) – Phiên bản nói & tương tác liền mạch',
        lesson_row_str='Tin học 10 (KNTT) – Bài 13 (Tiết 1 LT – PPCT 19)',
        page1_title='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 – LỚP 10/3 – PHIÊN BẢN DẠY HỌC)',
        page1_subtitle='(Tài liệu cầm tay đọc trực tiếp theo nhịp thuyết trình cho Sinh viên Hồ Quang Huy [24S1020030] – Lớp 10/3)',
        page1_class_line='Lịch dạy: Thứ Hai, 16/11/2026 – Tiết 3 (09h00–09h45) – Phòng H.I.5 | GVHD: Thầy Mai Xuân Văn',
        page1_time_line='Khớp 100%: 01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx & bảng demo mau_bai13.svg',
        page1_sched_line='Bối cảnh phòng học: 01 máy tính giáo viên, máy chiếu, bảng phấn (học sinh sử dụng trực tiếp SGK Tin học 10 tr. 69–72)',
        page1_mode_line='Chiến thuật dạy: Dẫn dắt như một bài thuyết trình sinh động – Bám sát SGK tr. 69–72 – Bấm sẵn mau_bai13.svg – Tạo tiền đề cho Tiết 2'
    )

    # HƯỚNG DẪN ĐỌC KỊCH BẢN
    eng.add(b, 'QUY ƯỚC ĐỌC NHANH TRÊN BỤC GIẢNG', 'heading')
    for item in [
        '**[Hành động & Thao tác]:** Thời điểm di chuyển, chỉ màn chiếu hoặc thao tác vào các khu vực đã dựng sẵn trong tệp mau_bai13.svg trên Inkscape.',
        '**[Lời nói trực tiếp]:** Phần lời trình bày tự nhiên theo nhịp một bài thuyết trình trên lớp, rõ ràng, dễ nói.',
        '**[Tương tác với học sinh]:** Câu hỏi ngắn bám sát SGK để học sinh đọc hoặc trả lời được ngay, tránh không khí im lặng kéo dài.',
        '**[Xử lý tình huống]:** Cách gợi ý nhanh nếu lớp ngập ngừng hoặc thao tác gặp sự cố.',
        '**[Viết bảng & Chốt ghi vở]:** Tóm tắt ngắn gọn ghi lên bảng giữa và đọc chậm cho học sinh chép bài.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 0
    eng.add(b, 'PHẦN 0. CHUẨN BỊ TRƯỚC GIỜ LÊN LỚP (PHÒNG H.I.5)', 'heading')
    for item in [
        '**[Hành động & Thao tác]:** Trước khi vào tiết, kết nối máy chiếu và mở sẵn tệp mau_bai13.svg bằng Inkscape để phần mềm nạp sẵn giao diện, tránh chờ đợi trong giờ học.',
        '**[Hành động & Thao tác]:** Nhấn Shift + Ctrl + F để mở cố định hộp thoại Fill and Stroke ở cạnh phải màn hình. Vẽ thử một hình tròn nháp ở vùng trống bên ngoài để kiểm tra trạng thái mặc định (Opacity đạt 100%, chế độ hình tròn khép kín), sau đó nhấn Delete xoá hình nháp.',
        '**[Viết bảng sẵn]:** Ở góc trái bảng ghi thứ ngày và tiêu đề "Tiết 19: Bài 13. Bổ sung các đối tượng đồ hoạ"; để trống phần bảng giữa ghi 3 ý trọng tâm.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 1
    eng.add(b, 'PHẦN 1. ỔN ĐỊNH LỚP & MỞ ĐẦU BÀI HỌC (PHÚT 00 – 05)', 'heading')
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    for item in [
        '**[Hành động & Thao tác]:** Đứng ở vị trí giữa bục giảng, quan sát cả lớp để ổn định trật tự, chào lớp và mời học sinh mở SGK Tin học 10 trang 69.',
        '**[Lời nói trực tiếp]:** *"Chào cả lớp 10/3. Hôm nay chúng ta học Bài 13: Bổ sung các đối tượng đồ hoạ. Trước khi đi vào nội dung mới, các em mở SGK trang 69 và quan sát Hình 13.1 trên màn chiếu."*',
        '**[Hành động & Thao tác]:** Phóng to vừa tầm nhìn vào Khu vực 1 bên trái của tệp mau_bai13.svg nơi có sẵn hình miếng dưa hấu 4 lớp.',
        '**[Tương tác với học sinh]:** *"Nhìn vào miếng dưa hấu ở Hình 13.1, một bạn nêu giúp thầy hình này được tạo thành từ những lớp màu và hình cơ bản nào?"* — Mời 01 học sinh trả lời nhanh (vỏ xanh đậm, cùi trắng, ruột đỏ là các nửa hình tròn; các hạt đen là hình elip nhỏ).',
        '**[Lời nói trực tiếp]:** *"Đúng rồi. Bây giờ chúng ta nhắc lại quy tắc thứ tự lớp ở Bài 12: trong Inkscape, hình nào vẽ trước thì nằm ở lớp dưới, hình nào vẽ sau thì đè lên lớp trên, giống như khi xếp chồng các lớp giấy màu lên nhau. Vậy giữa vỏ xanh, cùi trắng, ruột đỏ và các hạt đen, phần nào nằm dưới cùng và phần nào phải nằm trên cùng?"*',
        '**[Hành động & Thao tác]:** Dùng chuột kéo tách nhẹ từng lớp của miếng dưa hấu tại Khu vực 1 trên màn chiếu để học sinh quan sát trực quan từ dưới lên trên (vỏ xanh, cùi trắng, ruột đỏ, hạt đen), sau đó nhấn Ctrl + Z liên tiếp đưa các lớp trở về vị trí ban đầu.',
        '**[Xử lý tình huống]:** Nếu lớp ngập ngừng chưa trả lời, hỏi gợi mở trực tiếp: *"Nếu ta đặt các hạt đen nằm dưới lớp ruột đỏ thì nhìn từ trên xuống có thấy hạt dưa nữa không?"* — Học sinh sẽ đáp ngay hạt phải đặt ở lớp trên cùng.',
        '**[Lời nói trực tiếp – Dẫn vào bài]:** *"Ở Bài 12, các em đã vẽ được hình vuông, hình tròn nguyên vẹn và tô một màu đơn giản. Để xẻ đôi hình tròn thành nửa hình như miếng dưa hấu, tô màu có độ nổi khối hoặc ghép cắt nhiều hình với nhau, chúng ta sẽ tìm hiểu 3 nội dung chính trong Bài 13 hôm nay: chỉnh thuộc tính hình khối, sử dụng hộp thoại Fill and Stroke và các phép ghép đối tượng trong menu Path."*',
        '**[Viết bảng & Chốt ghi vở]:** Viết tên bài lên phần bảng giữa và hướng dẫn học sinh ghi ngắn gọn tựa Bài 13 cùng nhận xét thứ tự 4 lớp của Hình 13.1.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 2
    eng.add(b, 'PHẦN 2. MỤC 1 – THUỘC TÍNH CÁC ĐỐI TƯỢNG HÌNH KHỐI (PHÚT 05 – 14)', 'heading')
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Chúng ta vào phần 1: Các đối tượng hình khối và thuộc tính. Trước hết, các em quan sát Hình 13.2 trong SGK trang 69."*',
        '**[Tương tác với học sinh – Đọc thẳng trong SGK]:** *"Trên Hình 13.2 có 4 hình: ngôi sao nhọn, ngũ giác đều, ngôi sao cánh tròn và hình méo ngẫu nhiên. Cả 4 hình đều được tạo từ cùng một công cụ ở Bài 12. Các em nhìn vào biểu tượng ngôi sao ở ô cuối bảng đầu trang 70 sẽ đọc được ngay tên công cụ đó là gì?"* — Gọi 01 học sinh đọc (*"Công cụ Hình đa giác, hình sao"*), rồi chốt nhanh cách biến đổi tạo 4 hình từ trái sang phải: hình 1 là sao 5 cánh nhọn (Corners = 5); hình 2 bấm chuyển chế độ đa giác 5 cạnh; hình 3 tăng độ cong đỉnh (Rounded); hình 4 tăng độ méo ngẫu nhiên (Randomized).',
        '**[Lời nói trực tiếp – Giao nhiệm vụ cặp đôi đọc Bảng thuộc tính SGK tr. 69–70]:** *"Bây giờ các em dành 2 phút cùng bạn bên cạnh đọc Bảng các thuộc tính cơ bản ở cuối trang 69 và đầu trang 70 SGK, gạch chân từ khoá để nêu ý nghĩa thuộc tính của 3 nhóm hình khối:"*',
        '**[Tương tác với học sinh – 3 ý đọc Bảng SGK tr. 69–70]:**\n'
        '  Ý 1 (Hình vuông, hình chữ nhật): *"W, H là chiều rộng, chiều dài; còn Rx, Ry dùng để làm gì?"* (**Bán kính góc bo**, bằng 0 nếu góc vuông).\n'
        '  Ý 2 (Hình tròn, hình elip): *"Start, End dùng để làm gì và nhập bao nhiêu độ để cắt thành nửa hình tròn như miếng dưa hấu?"* (**0** và **180**, cùng 3 chế độ vẽ hình quạt, cung tròn hở, hình tròn kín).\n'
        '  Ý 3 (Hình đa giác, hình sao): *"Gọi tên 4 thuộc tính ở bảng đầu trang 70?"* (**Corners** số đỉnh/cánh, **Rounded** độ cong tại các đỉnh, **Spoke ratio** tỉ lệ bán kính trong/ngoài, **Randomized** làm méo hình ngẫu nhiên).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 2]:** Sau khi mời lần lượt 03 học sinh trả lời từng nhóm hình, chuyển màn chiếu sang Khu vực 2 của mau_bai13.svg: *"Bây giờ các em quan sát trực tiếp kết quả trên màn chiếu khi thầy thay đổi đúng các thông số này."*\n'
        '  Thao tác 1: Nháy đúp chuột vào **Hình chữ nhật màu xanh**, chỉ lên thanh thuộc tính nhập Rx = 25 rồi nhấn Enter để bốn góc bo tròn lại.\n'
        '  Thao tác 2: Nháy đúp vào **Hình tròn màu cam**, nhập Start = 0, End = 180 rồi nhấn Enter để tạo nửa hình tròn; di chuột chỉ 3 nút tròn cuối thanh thuộc tính (hình quạt nối tâm, cung tròn hở, và nút đóng kín thành hình tròn nguyên vẹn).\n'
        '  Thao tác 3: Nháy đúp vào **Ngôi sao màu đỏ**, sửa ô Rounded = 0.25 rồi nhấn Enter để các đỉnh nhọn uốn cong thành cánh hoa.',
        '**[Xử lý tình huống]:** Nếu học sinh ngại đọc chữ tiếng Anh, nhắc ngay: *"Không sao, các em chỉ cần nhìn đúng cột Thuộc tính bên trái dòng mô tả ở trang 69–70 và đọc từng chữ là được."*',
        '**[Viết bảng & Chốt ghi vở]:** Ghi tóm tắt Mục 1 lên bảng giữa và đọc cho lớp chép đúng theo bảng SGK: *"1. Thuộc tính hình khối (SGK tr. 69–70): Hình vuông, chữ nhật có W, H (chiều rộng, chiều dài) và Rx, Ry (bán kính góc bo; bằng 0 nếu góc vuông). Hình tròn, elip có Rx, Ry (bán kính ngang và thẳng đứng) và Start, End (góc điểm đầu và cuối; 0 và 180 tạo nửa hình tròn). Đa giác, hình sao có Corners (số đỉnh hoặc cánh), Rounded (độ cong tại các đỉnh), Spoke ratio (tỉ lệ bán kính từ tâm đến góc trong và tới đỉnh đầu) và Randomized (làm méo hình ngẫu nhiên)."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 3
    eng.add(b, 'PHẦN 3. MỤC 2 – TÔ MÀU, MÀU VẼ & ĐỘ TRONG SUỐT (PHÚT 14 – 23)', 'heading')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    for item in [
        '**[Lời nói trực tiếp – Liên hệ Bài 12]:** *"Chúng ta chuyển sang phần 2: Thiết lập màu tô, màu vẽ và độ trong suốt. Trước hết, một bạn nhắc lại giúp thầy ở Bài 12, muốn đổi màu nền và màu viền nhanh dưới dải màu đáy màn hình mình thao tác thế nào?"* — Gọi nhanh 1 học sinh nhắc lại (*"Nháy chuột đổi màu nền, giữ Shift và nháy chuột đổi màu viền"*).',
        '**[Lời nói trực tiếp]:** *"Đúng rồi. Tuy nhiên bảng màu dưới đáy chỉ tô được màu đồng nhất. Khi muốn tô màu chuyển tạo cảm giác nổi khối 3D hoặc đổi kiểu nét đứt cho đường viền, ta dùng hộp thoại Fill and Stroke bằng cách chọn lệnh Object / Fill and Stroke, hoặc nháy nút phải chuột vào đối tượng chọn Fill and Stroke, hay nhấn tổ hợp phím Shift + Ctrl + F."*',
        '**[Tương tác với học sinh – Quan sát Hình 13.3 tr. 70]:** *"Các em quan sát 5 ô vuông ở Hình 13.3 trang 70 và đọc 5 dòng giải thích bên dưới hình: một bạn mô tả giúp thầy phần màu tô bên trong của 5 hình này khác nhau thế nào?"* — Mời 01 học sinh đọc và nhận xét (*"Ô 1 không màu dấu X; Ô 2 màu đồng nhất; Ô 3 chuyển giữa hai hoặc nhiều màu theo chiều ngang; Ô 4 chuyển màu từ tâm ra ngoài; Ô 5 chuyển màu kết hợp giảm độ trong suốt"*).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 3]:** Chuyển sang Khu vực 3 và chỉ vào hộp thoại Fill and Stroke đang mở sẵn bên phải màn hình, đối chiếu với Hình 13.4 trong sách:\n'
        '  Giải thích cấu trúc 3 trang và biểu tượng phụ: *"Hộp thoại gồm 3 trang rõ ràng: Fill chỉnh màu tô nền bên trong (nút dấu hỏi chấm dùng huỷ đặt đối tượng về trạng thái ban đầu khi sao chép đối tượng); Stroke paint chỉnh màu đường viền (có công cụ hình ống hút để sao chép màu bất kỳ trong vùng làm việc); Stroke style chỉnh độ dày Width và kiểu nét liền hoặc nét đứt; biểu tượng cây bút phía dưới dùng chỉnh thông số của gradient."*\n'
        '  Thao tác tô chuyển sắc: Bấm vào **Hình tròn màu xanh** ở Khu vực 3, chọn trang Fill, nháy biểu tượng thứ tư **Radial gradient** để thấy màu chuyển từ tâm ra ngoài; chỉ thêm nút **Linear gradient** bên cạnh để thấy kiểu chuyển màu theo phương thẳng.\n'
        '  Thao tác kiểu đường viền: Chuyển sang trang Stroke style, chỉ ô Width (độ dày) và mục Dashes chọn kiểu nét đứt để lớp thấy đường viền thay đổi ngay.\n'
        '  Thao tác độ trong suốt: Bấm vào **Hình tròn xanh dương** đang nằm đè trên hình chữ nhật vàng, kéo thanh trượt **Opacity** phía dưới từ 100% xuống khoảng 50%: *"Khi giảm giá trị Opacity hoặc kênh A (alpha), màu trở nên trong suốt một phần giúp nhìn xuyên qua lớp bên dưới."*',
        '**[Viết bảng & Chốt ghi vở]:** Ghi gọn Mục 2 lên bảng giữa: *"2. Hộp thoại Fill and Stroke (Object / Fill and Stroke, nháy phải chuột chọn Fill and Stroke, hoặc Shift + Ctrl + F): Gồm 3 trang là Fill (không màu X, đồng nhất, chuyển thẳng Linear, chuyển từ tâm Radial, hoa văn Pattern; dấu ? huỷ đặt), Stroke paint (màu viền, ống hút sao chép màu), Stroke style (độ dày Width và kiểu nét vẽ). Biểu tượng cây bút chỉnh thông số gradient; Opacity và A (alpha) chỉnh độ trong suốt."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 4
    eng.add(b, 'PHẦN 4. MỤC 3 – CÁC PHÉP GHÉP ĐỐI TƯỢNG TRONG MENU Path (PHÚT 23 – 33)', 'heading')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Chúng ta sang phần 3: Các phép ghép đối tượng đồ hoạ. Các em nhìn Hình 13.5 trang 71: muốn ghép 3 mẫu giấy rời thành trái tim cân đối, ta xoay hình vuông nghiêng một góc 45 độ, đặt hai hình tròn chồng lên hai cạnh phía trên của hình vuông rồi dùng lệnh ghép hình trong bảng chọn Path."*',
        '**[Liên hệ Toán lớp 10 & Hướng dẫn đọc Hình 13.6 tr. 72]:** *"Các lệnh trong Path tương ứng trực tiếp với phép Hợp, Hiệu và Giao trong môn Toán lớp 10. Ở Hình 13.6 trang 72 quy ước hình vuông vẽ trước (lớp dưới) và hình tròn vẽ sau (lớp trên). Các em đối chiếu nhãn tiếng Việt dưới từng hình ở Hình 13.6 với chữ in đậm ở trang 71–72 để phân biệt 06 phép ghép:"*\n'
        '  1. *Phép hợp (Union, Ctrl + +):* Lấy tất cả các phần thuộc một trong các hình đơn.\n'
        '  2. *Phép hiệu (Difference, Ctrl + -):* Lấy phần thuộc hình lớp dưới nhưng không thuộc hình lớp trên.\n'
        '  3. *Phép giao (Intersection, Ctrl + *):* Chỉ giữ phần thuộc cả hai hình được chọn.\n'
        '  4. *Phép hiệu đối xứng (Exclusion, Ctrl + ^):* Phần hình thuộc các hình trừ phần giao nhau.\n'
        '  5. *Phép chia (Division, Ctrl + /):* Hình lớp dưới được chia thành các phần bởi đường nét của hình lớp trên.\n'
        '  6. *Phép cắt (Cut Path, Ctrl + Alt + /):* Cắt hình lớp dưới thành các phần bởi hai điểm giao ở viền với hình lớp trên (kết quả là các hình mới không có màu).',
        '**[Tương tác chốt quy tắc lớp]:** *"Trong các phép hiệu (Difference), phép chia (Division) và phép cắt (Cut Path), đối tượng ở lớp nào luôn là đối tượng bị cắt?"* — Mời 01 học sinh trả lời: **Đối tượng lớp dưới luôn bị cắt bởi đối tượng lớp trên**.',
        '**[Hành động & Thao tác trên mau_bai13.svg – Khu vực 4]:** Chuyển sang Khu vực 4 để thao tác trực tiếp bằng chuột qua menu Path (tránh kẹt phím Shift với dấu + trên laptop):\n'
        '  Thao tác 1 (Thử phép giao Intersection / phép hiệu Difference): Quét chọn hình vuông đỏ và hình tròn cam ở ô B, vào menu **Path chọn Intersection** (xem phần giao), rồi bấm Ctrl + Z và thử chọn **Difference** để cả lớp thấy hình vuông lớp dưới bị khoét khuyết đúng phần đè.\n'
        '  Thao tác 2 (Cắt khuyết logo Apple): Quét chọn hình quả táo đen ở lớp dưới và hình tròn xám đè ở mép phải tại ô C, mở menu **Path** rồi chọn **Difference** để tạo vết cắt khuyết tròn đều trên thân quả táo; giải thích thêm phần chiếc lá phía trên là kết quả lấy phần giao của 2 hình tròn bằng **Intersection**.',
        '**[Mở rộng kiến thức thực tế & AI (30 giây)]:** *"Trong thực tế, tệp đồ hoạ vectơ (.svg) dùng trực tiếp cho in ấn khổ lớn hay máy cắt CNC, cắt Laser vì đường Path không vỡ nét khi phóng to. Tương tự khi dùng công cụ AI tạo ảnh, ảnh .png phóng to dễ mờ, còn nếu yêu cầu AI xuất mã .svg thì ta đưa vào Inkscape mở từng đường Path ra sửa rất thuận tiện."*',
        '**[Viết bảng & Chốt ghi vở]:** Đọc cho học sinh ghi tóm tắt Mục 3: *"3. Các phép ghép đối tượng trong bảng chọn Path (SGK tr. 71–72): Trọng tâm gồm Union (Phép hợp, Ctrl + +), Difference (Phép hiệu, Ctrl + -), Intersection (Phép giao, Ctrl + *); cùng Exclusion, Division, Cut Path. Quy tắc cần nhớ: Với Difference, Division và Cut Path, ĐỐI TƯỢNG LỚP DƯỚI luôn bị ĐỐI TƯỢNG LỚP TRÊN cắt. Nếu chọn nhầm thứ tự lớp, nhấn Ctrl + Z, chọn hình cần giữ rồi nhấn phím End đưa xuống dưới cùng trước khi cắt lại."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 5
    eng.add(b, 'PHẦN 5. LUYỆN TẬP – GIẢI 04 CÂU HỎI CỦNG CỐ SGK TRANG 70, 71, 72 (PHÚT 33 – 40)', 'heading')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Để củng cố trọn vẹn 3 nội dung lí thuyết vừa học, chúng ta cùng giải quyết 04 câu hỏi củng cố chính thức trong SGK từ trang 70 đến trang 72."*',
        '**[Tương tác với học sinh – Luyện tập 4 câu hỏi SGK tr. 70–72]:**\n'
        '  Câu 1 (SGK tr. 70): *"Nếu trên hình vẽ có sẵn một hình sao 5 cánh nhọn, em cần thay đổi tham số nào để các đỉnh ngôi sao trở nên cong?"* — Mời 01 học sinh trả lời nhanh: Tham số **Rounded**.\n'
        '  Câu 2 (Trắc nghiệm 1 – SGK tr. 71): *"Để xác định đường viền của đối tượng dạng nét đứt, em cần chọn trang nào trong hộp thoại Fill and Stroke: A. Fill, B. Stroke paint, C. Stroke style, hay D. Cả A và B?"* — Đếm 1, 2, 3 cho cả lớp đồng loạt giơ ngón tay chọn đáp án: **3 ngón – C. Stroke style**.\n'
        '  Câu 3 (Trắc nghiệm 2 – SGK tr. 71): *"Để chỉnh thông số của gradient, em cần chọn biểu tượng nào trong 4 phương án A, B, C, D?"* — Đếm 1, 2, 3 cả lớp đồng loạt giơ ngón tay: **2 ngón – B** (biểu tượng cây bút chỉnh gradient).\n'
        '  Câu 4 (SGK tr. 72 – Hình 13.7): Chiếu Hình 13.7 lên màn hình: *"Hai bạn cùng bàn trao đổi trong 2 phút, nêu giúp thầy tên phép ghép hình VÀ các bước thực hiện để vẽ đám mây hoàn chỉnh như Hình 13.7?"* — Mời 01–02 học sinh trình bày đủ 3 bước dùng **Phép hợp (Path / Union)**: Bước 1 vẽ các hình tròn/elip chồng lấp lên nhau tạo khung đám mây; Bước 2 chọn tất cả các hình vừa vẽ; Bước 3 chọn **Path / Union** (hoặc nhấn Ctrl + +).',
        '**[Hành động & Thao tác kiểm chứng trên mau_bai13.svg – Khu vực 4A]:** Quét chọn 4 hình xanh ở ô A của Khu vực 4, mở menu **Path rồi bấm Union** để ghép thành hình đám mây hoàn chỉnh ngay trước mắt cả lớp, sau đó ghi ngắn gọn đáp án 4 câu củng cố lên góc phải bảng.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 6
    eng.add(b, 'PHẦN 6. VẬN DỤNG & TIỀN ĐỀ CHUẨN BỊ CHO TIẾT 2 THỰC HÀNH (PHÚT 40 – 45)', 'heading')
    for item in [
        '**[Lời nói trực tiếp]:** *"Còn 5 phút cuối, chúng ta quay lại giải quyết câu hỏi Vẽ miếng dưa hấu Hình 13.1 đã nêu ở đầu giờ, đồng thời chuẩn bị kiến thức cho giờ Thực hành ở Tiết 2."*',
        '**[Tương tác với học sinh – Giải quyết Hình 13.1 (liên hệ Vận dụng 1 tr. 74)]:** Mời 01 hoặc 02 học sinh trình bày đề xuất 3 bước vẽ miếng dưa hấu:\n'
        '  Bước 1: Chọn công cụ Hình tròn (E), nhập **Start = 0, End = 180** ở chế độ hình quạt (Slice) tô xanh lá đậm làm vỏ dưa (lớp dưới cùng).\n'
        '  Bước 2: Nhấn **Ctrl + D** nhân bản, thu nhỏ và tô trắng làm cùi; tiếp tục nhấn **Ctrl + D** nhân bản, thu nhỏ và tô đỏ làm ruột dưa.\n'
        '  Bước 3: Vẽ các hình elip nhỏ màu đen đặt lên lớp trên cùng làm hạt dưa.',
        '**[Hành động & Thao tác kiểm chứng trên mau_bai13.svg – Khu vực 1 bên phải]:** *"Thầy thao tác nhanh quy trình các em vừa nêu ngay trên hình tròn chờ sẵn ở Khu vực 1."* — Chọn hình tròn xanh bên phải Khu vực 1, nhấn E, gõ Start = 0, End = 180, nhấn Ctrl + D thu nhỏ tô trắng, Ctrl + D thu nhỏ tô đỏ và thêm 2 hạt elip đen trong khoảng 25 giây (nếu sắp hết giờ thì minh hoạ bước cắt nửa hình tròn và Ctrl + D tô đỏ cho kịp giờ chuông).',
        '**[Tương tác – Xây dựng tiền đề trực tiếp cho Tiết 2 Thực hành (SGK tr. 72–74)]:** *"Để tiết Thực hành sau (Tiết 2 – PPCT 20) làm nhanh và chính xác, các em xem lướt tên 3 bài Thực hành trang 72 đến 74 và nêu giúp thầy mỗi bài sẽ dùng lại kiến thức nào của hôm nay?"*\n'
        '  1. *Nhiệm vụ 1 – Vẽ Logo (Hình 13.8):* Áp dụng chế độ **cung tròn hở** của Hình tròn và trang **Stroke style** (tăng độ dày Width) + **Stroke paint**.\n'
        '  2. *Nhiệm vụ 2 – Vẽ Hình cây (Hình 13.10):* Áp dụng phép hiệu **Path / Difference** (cắt thân cây từng cặp 2 hình) và phép hợp **Path / Union** (gộp tán lá với thân).\n'
        '  3. *Nhiệm vụ 3 – Vẽ Hình cầu (Hình 13.12):* Áp dụng tắt đường viền (X) và kiểu tô chuyển **Linear gradient**, **Radial gradient** để tạo độ nổi khối sáng tối 3D.',
        '**[Lời nói trực tiếp – Dặn dò & Chào kết thúc]:** *"Về nhà, các em ôn lại 3 nội dung lí thuyết hôm nay và đọc trước hướng dẫn 3 bài thực hành trang 72 đến 74 để Tiết 2 thực hành thao tác cho nhanh và chuẩn xác. Cảm ơn các em đã chú ý lắng nghe và hợp tác trong tiết học hôm nay. Chào cả lớp."*',
        '**[Hành động kết thúc]:** Tắt màn chiếu gọn gàng, xoá bảng sạch sẽ và kết thúc tiết học đúng giờ.'
    ]:
        eng.add(b, item, 'bullet')

    out_path = DAY_THAT_DIR / '03_Bai_thuyet_trinh_giang_day_lien_mach_45_phut_Tin10_Bai13_Tiet1.docx'
    used_figs = [
        'hinh_13_1_2_dua_hau_thuoc_tinh.png',
        'bang_13_1_thuoc_tinh_hinh_khoi.png',
        'hinh_13_3_4_fill_stroke.png',
        'hinh_13_5_ket_hop_doi_tuong.png',
        'hinh_13_6_cac_phep_ghep_hinh.png',
        'hinh_13_7_8_dam_may_logo.png',
    ]
    clean_and_validate_docx(b, out_path, used_figures=used_figs, is_khbd=False)


if __name__ == '__main__':
    DAY_THAT_DIR.mkdir(parents=True, exist_ok=True)
    build_khbd_day_that()
    build_kich_ban_vui_nhon()
    build_bai_thuyet_trinh_day_hoc()
    print('Done building all 3 Day_that docx files in:', DAY_THAT_DIR)
