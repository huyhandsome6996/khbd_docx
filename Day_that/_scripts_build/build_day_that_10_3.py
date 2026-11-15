# -*- coding: utf-8 -*-
"""
Bo tao ho so DAY THAT cho tiet giang day thuc te Lop 10/3:
- Thoi gian: Thu Hai, ngay 16/11/2026 - Tiet 3 (09h00 - 09h45)
- Dia diem: Phong H.I.5 (1 may tinh GV + may chieu + bang phan)
- Bai day: Tin hoc 10 (KNTT) - Bai 13: Bo sung cac doi tuong do hoa (Tiet 1 - Li thuyet / Tiet PPCT: 19, SGK tr. 69-72, lam tien de cho Tiet 2 - TH tr. 72-74)
- GVHD / GVGD: Thay Mai Xuan Van | SVTT: Ho Quang Huy (24S1020030)

Xuat ra 02 tep trong thu muc c:/Users/PC/khbd_docx/Day_that/:
  1. khbd10-3.docx
     (Ke hoach bai day chuan CV 5512 dung de nop va kiem tra, dung pham vi Tiet 1 Li thuyet tr. 69-72)
  2. Kich_ban_thuyet_trinh_va_tap_duyet_Bai13_Tiet1.docx
     (Tai lieu tap duyet o nha va cam tay len lop: giai thich ban chat kien thuc, huong dan bam Inkscape tu con so 0 tren giao dien may, loi giang chi tiet 45 phut & bang phan)
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
        cover_title='KẾ HOẠCH BÀI DẠY',
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

    out_path = DAY_THAT_DIR / 'khbd10-3.docx'
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
# FILE 2: KICH BAN THUYET TRINH & TAP DUYET INKSCAPE TOAN DIEN (KHONG TRANG BIA)
# ==============================================================================
def build_kich_ban_tap_duyet_toan_dien():
    # Khởi tạo trực tiếp từ body, KHÔNG tạo trang bìa hành chính (vì tài liệu dùng đọc và tập duyệt ở nhà)
    b = eng.node('body')

    # Tiêu đề mở đầu gọn gàng ngay đầu Trang 1
    eng.add(b, 'KỊCH BẢN THUYẾT TRÌNH & HƯỚNG DẪN TẬP DẠY BÀI 13 (TIẾT 1 – LÍ THUYẾT)', 'title', keep=True)
    eng.add(b, 'Môn Tin học 10 (Kết nối tri thức) – Trang 69–72 SGK – Tự học Inkscape & Giảng sâu bản chất', align='center', first=0, italic=True, sz=24, keep=True)
    eng.add(b, f'Lớp dạy: 10/3 (Phòng H.I.5 – Thứ Hai, 16/11/2026, Tiết 3: 09h00–09h45) | GVHD: {GV_NAME} | SVTH: {SV_NAME}', align='center', first=0, sz=22)

    # ==========================================================================
    # PHẦN I. BỨC TRANH TỔNG THỂ
    # ==========================================================================
    eng.add(b, 'PHẦN I. BỨC TRANH TỔNG THỂ: BÀI NÀY DẠY CÁI GÌ, MỤC ĐÍCH LÀ GÌ?', 'heading')
    eng.add(b, 'Nếu chỉ đọc lướt tên mục và thuật ngữ tiếng Anh thì sẽ thấy bài rất rời rạc và khô khan. Trước khi mở sách hay mở phần mềm, bạn hãy nắm chắc câu chuyện xuyên suốt của cả Tiết 1 trong 3 dòng sau:', first=0)
    for item in [
        '**Điểm xuất phát (Bài 12 đã học):** Ở Bài 12, học sinh mới chỉ biết vẽ ra hình vuông, hình tròn **nguyên vẹn**, tô **một màu bệt** dưới đáy màn hình và hiểu quy tắc **vẽ trước nằm lớp dưới, vẽ sau nằm lớp trên**.',
        '**Vấn đề đặt ra ở Bài 13 (Tại sao phải học bài này?):** Khi muốn vẽ một vật thật đẹp như **miếng dưa hấu (Hình 13.1)**, **trái tim**, **đám mây** hay **quả cầu nổi khối 3D**, bộ kỹ năng ở Bài 12 bị bế tắc hoàn toàn: không cắt được nửa hình tròn, không tô được màu sáng tối 3D, không nhìn xuyên thấu được qua lớp trên, và không đục khoét hay hàn dính nhiều hình thành một hình mới được.',
        '**Đích đến của Tiết 1 (45 phút hôm nay):** Học sinh nắm được **bản chất và công dụng của 03 nhóm công cụ** để giải quyết trọn vẹn 3 bế tắc trên, chứng minh ngay bằng 04 câu hỏi củng cố SGK (tr. 70–72) và ghép xong quy trình vẽ miếng dưa hấu; đồng thời **chuẩn bị sẵn hành trang** để tiết sau xuống phòng máy thực hành vẽ Logo, Hình cây và Quả cầu 3D (tr. 72–74) mà không bị bỡ ngỡ.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, 'Bảng tóm tắt mạch kiến thức Bài 13 Tiết 1 (Bế tắc từ Bài 12 đến Giải pháp ở Bài 13)', 'sub')
    overview_script_rows = [
        ['Điều Bài 12 CHƯA làm được', 'Công cụ Bài 13 giải quyết (Nội dung truyền đạt)', 'Sản phẩm cụ thể HS hiểu và làm được'],
        ['**Bế tắc 1:** Hình vẽ ra luôn thô cứng (góc vuông sắc, hình tròn kín bưng, sao 5 cánh nhọn y như nhau).',
         '**Mục 1 – Thanh điều khiển thuộc tính** (tr. 69–70): Chỉnh trực tiếp các con số **W, H, Rx, Ry** (bo góc), **Start, End + 3 nút tròn** (cắt miếng bánh/cung hở), **Corners, Rounded, Spoke ratio, Randomized** (biến hoá đa giác/sao).',
         'Hiểu cách cắt đúng **nửa hình tròn (0 đến 180 độ)** làm vỏ dưa hấu; làm cong đỉnh ngôi sao thành bông hoa (**Rounded**); cung tròn hở làm Logo (H.13.8 ở Tiết 2).'],
        ['**Bế tắc 2:** Chỉ tô được màu đồng nhất bệt, viền chỉ là nét liền trơn, màu lớp trên che kín mít lớp dưới.',
         '**Mục 2 – Hộp thoại Fill and Stroke** (Shift + Ctrl + F, tr. 70–71): Mở kho đồ nghề gồm **Fill** (5 kiểu tô, màu chuyển Linear/Radial, cây bút chỉnh hướng, dấu ?), **Stroke paint** (màu viền, ống hút), **Stroke style** (độ dày, nét đứt) và **Opacity / kênh A** (độ trong suốt).',
         'Phân biệt 5 ô màu Hình 13.3; chọn được **Stroke style** làm viền nét đứt; hiểu cách tô chuyển sắc tạo khối cầu 3D (H.13.12 ở Tiết 2) và làm lớp kính mờ nhìn xuyên thấu.'],
        ['**Bế tắc 3:** Muốn ghép trái tim hay đám mây mà xếp đè lên nhau vẫn bị lộ đường viền cắt ngang bên trong; không biết khoét khuyết một hình.',
         '**Mục 3 – 06 phép ghép trong bảng chọn Path** (tr. 71–72): Hàn dính (**Union**), khuôn dập cắt bỏ (**Difference**), lấy phần chung (**Intersection**), đục thủng giữa (**Exclusion**), cắt nhiều mảnh (**Division**), cắt đứt viền (**Cut Path**).',
         'Thuộc lòng quy tắc **hình lớp dưới luôn bị hình lớp trên cắt**; ghép được **trái tim (H.13.5)**, **đám mây (H.13.7)** và chuẩn bị cho cắt thân cây xanh (H.13.10 ở Tiết 2).']
    ]
    b.append(eng.table(overview_script_rows, [2650, 3621, 2800], header=True, cell_sz=21, cant_split=True))

    # ==========================================================================
    # PHẦN II. HƯỚNG DẪN TẬP DÙNG INKSCAPE TỪ CON SỐ 0 TẠI NHÀ
    # ==========================================================================
    eng.add(b, 'PHẦN II. HƯỚNG DẪN TẬP DÙNG INKSCAPE TỪ CON SỐ 0 TẠI NHÀ (10 PHÚT LÀM CHỦ)', 'heading')
    eng.add(b, 'Trước khi đứng lớp, bạn hãy mở Inkscape trên laptop lên giống hệt ảnh màn hình bạn vừa chụp và tập bấm tuần tự theo 3 bước dưới đây. Tập một lần ở nhà sẽ thấy phần mềm cực kỳ đơn giản.', first=0)

    eng.add(b, '1. Cách mở tệp mẫu dựng sẵn mau_bai13.svg và 4 phím điều hướng sống còn', 'sub')
    for item in [
        '**Mở tệp mẫu:** Trên cửa sổ Inkscape, bấm menu **File** (góc trên bên trái) chọn **Open...** (hoặc nhấn phím **Ctrl + O**), tìm tới tệp `c:/Users/PC/khbd_docx/Day_that/mau_bai13.svg` rồi bấm **Open**. (Trong tệp này tôi đã dựng sẵn 4 khu vực khớp từng phần giảng ở lớp, khung bảng được khoá chặt nên bạn kéo chuột thoả thích không sợ hỏng khung).',
        '**Phím số 5:** Nhấn phím **5** một cái là toàn bộ trang bảng mẫu nhảy ra giữa, vừa khít màn hình.',
        '**Giữ Ctrl + Lăn con lăn chuột:** Phóng to vào đúng Khu vực đang giảng (hoặc thu nhỏ lại để nhìn tổng thể).',
        '**Giữ phím Space (dấu cách dài) + Di chuột (hoặc kéo thanh cuộn):** Trượt bảng sang khu vực tiếp theo rất mượt.',
        '**Phím Ctrl + Z:** Lỡ tay kéo lệch, cắt nhầm hay xoá nhầm thì bấm **Ctrl + Z** — quay ngược lại như chưa từng có gì xảy ra. Bấm bao nhiêu lần cũng được!'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Bản đồ 5 vị trí nút bấm trên đúng giao diện Inkscape của bạn', 'sub')
    inkscape_map_rows = [
        ['Vị trí trên màn hình Inkscape của bạn', 'Nút cần biết & Phím tắt', 'Bấm sao cho đúng? (Quy trình thao tác tay)'],
        ['**1. Hộp công cụ dọc bên trái**\n*(Cột biểu tượng góc trên trái)*',
         '- **Hàng 1, Cột 1 (Mũi tên trắng – phím S):** Chọn & Di chuyển hình.\n'
         '- **Hàng 2, Cột 1 (Hình vuông hồng – phím R):** Vẽ vuông / chữ nhật.\n'
         '- **Hàng 2, Cột 2 (Hình tròn hồng – phím E):** Vẽ tròn / elip / nửa tròn.\n'
         '- **Hàng 2, Cột 3 (Ngôi sao hồng – phím dấu sao):** Vẽ sao / đa giác.\n'
         '- **Hàng 4, Cột 1 (Ô chuyển màu xanh – phím G):** Chỉnh màu chuyển Gradient.\n'
         '- **Hàng 4, Cột 3 (Ống hút – phím D):** Hút sao chép màu từ hình khác.',
         '- Bấm vào **Mũi tên (S)** rồi bấm vào 1 hình bất kỳ trên giấy để chọn.\n'
         '- *Bấm 1 lần* vào hình: hiện 8 mũi tên kéo to/nhỏ. *Bấm thêm 1 lần nữa* vào hình đó: 4 góc đổi thành mũi tên cong để **xoay nghiêng** (ví dụ xoay hình vuông 45 độ ghép trái tim).\n'
         '- *Nhân bản cực nhanh:* Chọn hình bằng mũi tên S rồi bấm **Ctrl + D**, hình bản sao đè khít lên hình cũ, chỉ việc kéo góc thu nhỏ lại.\n'
         '- *Đổi vị trí lớp trên/dưới:* Chọn hình muốn đẩy xuống dưới cùng nhấn phím **End**; muốn đưa lên trên cùng nhấn phím **Home**.'],
        ['**2. Thanh điều khiển thuộc tính**\n*(Thanh ngang nằm ngay dưới chữ File, Edit, View...)*',
         '- Khi chọn Mũi tên (S): chỉ hiện **X, Y, W, H**.\n'
         '- Khi đang vẽ/chọn Hình vuông (R): hiện **W, H, Rx, Ry**.\n'
         '- Khi đang vẽ/chọn Hình tròn (E): hiện **Rx, Ry, Start, End** + **3 nút tròn nhỏ**.\n'
         '- Khi đang vẽ/chọn Ngôi sao (công cụ sao): hiện **2 nút chuyển Đa giác/Sao, Corners, Spoke ratio, Rounded, Randomized**.',
         '- **Quy tắc vàng chống bối rối:** Muốn sửa góc bo `Rx`, góc cắt `Start/End` hay độ cong `Rounded` của hình nào có sẵn trên giấy, chỉ cần **nháy đúp chuột (2 cái liên tiếp)** thẳng vào hình đó — lập tức thanh ngang phía trên hiện đủ các ô nhập số tương ứng! Gõ số xong bắt buộc **nhấn Enter** thì hình mới đổi.'],
        ['**3. Dải ô màu nhanh ở đáy màn hình**',
         '- Các ô màu trải dài sát đáy.\n'
         '- Ô **chữ X đỏ** ở góc trái cùng để xoá màu.\n'
         '- Ô **O: 100** bên dưới chữ X là độ trong suốt nhanh.',
         '- **Nháy chuột trái** vào 1 ô màu = Tô **màu nền bên trong (Fill)**.\n'
         '- **Giữ Shift + Nháy chuột trái** vào 1 ô màu = Tô **màu đường viền (Stroke)**.\n'
         '- Chuột trái vào **X** = mất màu nền; **Shift + chuột trái vào X** = mất đường viền.'],
        ['**4. Bảng Fill and Stroke ở khung bên phải**\n*(Nếu tắt bấm Shift + Ctrl + F hoặc Object / Fill and Stroke)*',
         '- **Thẻ Fill:** 7 nút vuông nhỏ (`X` không màu, ô đặc đồng nhất, **Linear gradient**, **Radial gradient**, Mesh, **Pattern** hoa văn, **?** huỷ đặt).\n'
         '- **Thẻ Stroke paint:** Chọn màu đường viền.\n'
         '- **Thẻ Stroke style:** Ô **Width** (độ dày viền) và mục **Dashes** (chọn kiểu nét đứt).\n'
         '- Thanh **Opacity (%)** ở góc dưới cùng bên phải.',
         '- **Lưu ý lệch phiên bản với SGK:** Trong SGK (Hình 13.4 tr. 71) có minh hoạ **biểu tượng cây bút** để sửa dải màu chuyển Gradient. Trên bản Inkscape mới của bạn, sau khi chọn tô màu chuyển, bạn bấm phím **G** (hoặc nút Gradient ở Hàng 4 Cột 1) là thấy ngay thanh 2 đầu chấm tròn nối trên hình để kéo đổi màu và chỉnh hướng sáng tối trực tiếp!'],
        ['**5. Bảng chọn Path trên thanh menu trên cùng**',
         '- **Union** (Phép hợp)\n'
         '- **Difference** (Phép hiệu)\n'
         '- **Intersection** (Phép giao)\n'
         '- **Exclusion** (Hiệu đối xứng)\n'
         '- **Division** (Phép chia)\n'
         '- **Cut Path** (Phép cắt)',
         '- Bấm **Mũi tên (S)**, nhấp chuột ra khoảng trắng ngoài 2 hình, **kéo quét thành khung nét đứt bao trọn 2 hình** (hoặc bấm hình 1, giữ phím **Shift**, bấm hình 2).\n'
         '- Sau đó bấm menu **Path** trên cùng rồi nhấp chuột thẳng vào chữ **Union / Difference / Intersection** — **KHÔNG bấm phím tắt Ctrl + + trên laptop** vì dễ kẹt phím Shift trước mặt học sinh!']
    ]
    b.append(eng.table(inkscape_map_rows, [2150, 2950, 3971], header=True, cell_sz=21, cant_split=True))

    eng.add(b, '3. Bài tập làm quen 3 phút trên trang trắng & Bảng cứu hộ 6 lỗi hay gặp nhất', 'sub')
    for item in [
        '**Tập tay 3 phút ở nhà (trên New document trống trước khi mở mau_bai13.svg):** (1) Nhấn `E` kéo vẽ 1 hình tròn, bấm ô màu đỏ dưới đáy cho ra ruột dưa; (2) Gõ `Start = 0`, `End = 180`, nhấn `Enter` xem nửa hình tròn hiện ra, rồi bấm thử cả **3 nút tròn nhỏ** bên cạnh ô `End` để phân biệt *hình quạt kín* với *cung tròn hở* và *tròn nguyên*; (3) Nhấn `R` vẽ hình chữ nhật đè lệch lên, nhấn `S` kéo quét cả 2 hình, mở menu **Path** bấm thử lần lượt **Union** (rồi `Ctrl + Z`), **Difference** (rồi `Ctrl + Z`), **Intersection** (rồi `Ctrl + Z`) — bấm 3 lần như vậy là hiểu thông suốt 100% quy tắc hình nào bị cắt!',
        '**Lưu ý khu vực 5 trong mau_bai13.svg:** Ở góc dưới phải của `mau_bai13.svg` có Khu vực 5 chứa hình mẫu **Quả cầu 3D** dùng để chiếu gợi mở 15 giây cuối tiết chuẩn bị cho Tiết 2 (các hình hoa/bóng bay bên cạnh là đồ dùng thực hành tr. 74 để dành sang Tiết 2, trong Tiết 1 bạn bỏ qua không giảng sâu vào đó).'
    ]:
        eng.add(b, item, 'bullet')

    rescue_rows = [
        ['Sự cố trên màn chiếu khi đang dạy', 'Tại sao lại bị như vậy?', 'Cách xử lý gọn trong 3 giây'],
        ['Vẽ hình ra bị **tàng hình** (chỉ thấy viền chấm chọn)', 'Inkscape nhớ lại lần trước đặt `Opacity = 0%` hoặc bỏ màu nền `X` lẫn viền `X`.', 'Nhìn góc dưới cùng bên trái đặt lại `O: 100`, rồi nháy chuột trái vào một ô màu xanh/đỏ ở dải màu đáy.'],
        ['Vẽ hình tròn mới ra chỉ thấy **nửa miếng bánh**', 'Inkscape đang nhớ lại góc `Start / End` của miếng dưa hấu vừa cắt trước đó.', 'Nhìn lên cuối thanh ngang trên cùng, bấm **nút hình tròn kín (nút thứ 3 ngoài cùng bên phải)** là tròn nguyên lại ngay.'],
        ['Bấm mãi không thấy ô `Start / End / Rounded` đâu', 'Đang để công cụ Mũi tên chọn (`S`) nên thanh trên chỉ hiện toạ độ `X, Y, W, H`.', '**Nháy đúp chuột** thẳng vào hình trên màn hình là thanh thuộc tính bật ra ngay.'],
        ['Bấm `Path / Difference` mà mất hình cần giữ hoặc không thấy cắt', 'Chọn nhầm **3 hình cùng lúc** (chỉ được chọn đúng 2 hình cho `Difference`), hoặc hình cần giữ lại nằm **lớp trên**.', 'Nhấn `Ctrl + Z`. Chọn riêng hình muốn giữ rồi bấm phím **End** đưa nó xuống lớp dưới cùng. Sau đó chọn lại đúng **2 hình** rồi vào menu **Path / Difference**.'],
        ['Bấm `Ctrl + +` ghép hình trên laptop không ăn', 'Dấu `+` dùng chung phím `=` nên phải giữ thêm `Shift`, rất dễ nhấn hụt.', 'Dùng chuột bấm thẳng lên menu **Path rồi bấm Union** — vừa chắc ăn 100%, vừa giúp cả lớp nhìn thấy tên lệnh trên màn chiếu.'],
        ['Nhỡ tay làm rối tung hình mẫu trong `mau_bai13.svg`', 'Thao tác thử nghiệm bị lệch khỏi vị trí ban đầu.', 'Nhấn `Ctrl + Z` vài lần; nếu rối quá chỉ việc đóng file chọn **Close without saving** rồi mở lại `mau_bai13.svg` là mới tinh.']
    ]
    b.append(eng.table(rescue_rows, [2400, 2750, 3921], header=True, cell_sz=21, cant_split=True))

    # ==========================================================================
    # PHẦN III. KỊCH BẢN GIẢNG BÀI CHI TIẾT TỪNG PHÚT (00 – 45)
    # ==========================================================================
    eng.add(b, 'PHẦN III. KỊCH BẢN GIẢNG BÀI CHI TIẾT TỪNG PHÚT (GIẢNG SÂU BẢN CHẤT)', 'heading')
    eng.add(b, 'Cấu trúc mỗi chặng gồm đủ 4 phần rõ ràng: (1) Mục đích sư phạm để bạn biết mình đang dẫn học sinh đi đâu; (2) Bản chất kiến thức giải thích bằng ví dụ đời thường dễ hiểu; (3) Lời giảng liền mạch, tự nhiên kết hợp thao tác bấm trên mau_bai13.svg; (4) Nội dung chép bảng chuẩn xác theo KHBD khbd10-3.docx.', first=0)

    # --- CHẶNG 1 ---
    eng.add(b, 'Chặng 1 (Phút 00 – 05): Mở đầu – Vì sao biết Bài 12 rồi vẫn chưa vẽ được Miếng dưa hấu (Hình 13.1)?', 'sub')
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    eng.add(b, '**1. Mục đích của chặng này:** Tạo lý do thuyết phục để học sinh thấy **bắt buộc phải học Bài 13**. Thay vì vào lớp đọc chép *"Hôm nay học thuộc tính..."*, ta đưa ngay một hình quen thuộc (miếng dưa hấu) để học sinh vừa ôn quy tắc lớp trên/dưới của Bài 12, vừa tự nhận ra Bài 12 chưa đủ công cụ để vẽ được hình này.', first=0)
    eng.add(b, '**2. Bản chất cần truyền đạt:** Một hình phức tạp trên máy tính thực chất chỉ là **nhiều hình hình học đơn giản xếp chồng lên nhau theo thứ tự lớp** (như xếp chồng các tờ giấy màu cắt thủ công: tờ đặt trước nằm dưới cùng, tờ đặt sau đè lên trên). Muốn hạt đen hiện rõ trên miếng dưa đỏ thì hạt đen buộc phải nằm ở lớp trên cùng.', first=0)
    eng.add(b, '**3. Lời giảng đầy đủ & Thao tác Inkscape (Khu vực 1 bên trái của mau_bai13.svg):**', first=0)
    for item in [
        '*(Ổn định lớp, mời học sinh mở SGK Tin học 10 trang 69, phóng to màn hình vào Khu vực 1 bên trái nơi có sẵn miếng dưa hấu 4 lớp).*',
        '**Nói:** *"Chào cả lớp 10/3. Trước khi vào bài mới, các em nhìn miếng dưa hấu ở Hình 13.1 trang 69 và trả lời nhanh giúp thầy hai câu nhỏ: Thứ nhất, miếng dưa hấu này thực chất được ghép từ những hình học đơn giản nào và màu sắc ra sao? Thứ hai, nhớ lại quy tắc ở Bài 12 — hình nào vẽ trước thì nằm lớp dưới, hình vẽ sau nằm lớp trên — vậy trong 4 phần vỏ xanh, cùi trắng, ruột đỏ và hạt đen, phần nào nằm dưới cùng và phần nào nằm trên cùng?"*',
        '*(Gọi 1–2 học sinh trả lời. Nếu lớp im lặng, hỏi gợi mở: "Nếu ta vẽ hạt đen xong rồi mới phủ miếng ruột đỏ lên trên thì nhìn từ trên xuống còn thấy hạt dưa nữa không?" — lớp sẽ trả lời ngay là hạt đen phải nằm trên cùng).* ',
        '*(Thao tác Inkscape kiểm chứng):* Dùng mũi tên chọn (**S**) **kéo tách nhẹ từng lớp** của miếng dưa hấu ở Khu vực 1 ra lệch nhau một chút cho cả lớp nhìn thấy rõ 4 miếng rời xếp chồng từ dưới lên (vỏ xanh dưới cùng, cùi trắng, ruột đỏ, hạt đen trên cùng), rồi nhấn **Ctrl + Z** trả lại nguyên vẹn.',
        '**Dẫn vào bài (điểm chốt tạo động lực học):** *"Rất chính xác! Nhưng có một vấn đề: nếu bây giờ thầy yêu cầu các em dùng đúng kiến thức Bài 12 để vẽ lại miếng dưa hấu này, các em sẽ kẹt ngay ở hai chỗ. Một là Bài 12 chỉ vẽ được hình tròn kín bưng chứ chưa biết cách xẻ đôi thành nửa hình tròn như miếng dưa. Hai là sau này gặp hình trái tim hay đám mây, chỉ xếp đè giấy lên thôi thì không hàn dính lại được. Vì vậy hôm nay chúng ta học Bài 13: Bổ sung các đối tượng đồ hoạ để nắm 3 công cụ giải quyết triệt để chuyện đó: tinh chỉnh thông số hình khối, phối màu với Fill and Stroke, và cắt ghép hình với menu Path."*'
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**4. Viết bảng & Cho học sinh ghi vở:** Ghi tên bài ở giữa bảng: `Bài 13. Bổ sung các đối tượng đồ hoạ`. Cho HS ghi ngắn mục Mở đầu: Phân tích Hình 13.1 gồm 03 nửa hình tròn giảm dần kích thước và các hình elip nhỏ màu đen; thứ tự 4 lớp từ dưới lên trên là Vỏ xanh, Cùi trắng, Ruột đỏ, Hạt đen.', first=0)

    # --- CHẶNG 2A ---
    eng.add(b, 'Chặng 2A (Phút 05 – 14): Mục 1 – Giải mã Thanh điều khiển thuộc tính (SGK tr. 69–70)', 'sub')
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    eng.add(b, '**1. Mục đích của chặng này:** Không bắt học sinh học vẹt tên tiếng Anh, mà giúp các em hiểu: mỗi hình vẽ ra trong Inkscape đều giống một món đồ có gắn **bảng điều khiển thông số** ngay phía trên màn hình. Muốn bẻ góc nhọn thành góc bo, xẻ nửa vòng tròn hay biến ngôi sao thành bông hoa, ta không cần tự tay vẽ nét mới mà chỉ cần **nhập đúng con số điều khiển**.', first=0)
    eng.add(b, '**2. Bản chất từng thông số để giảng cho học sinh dễ hình dung nhất:**', first=0)
    for item in [
        '**Hình vuông, Hình chữ nhật** (`W, H` và `Rx, Ry`): `W` (Width) là chiều rộng, `H` (Height) là chiều cao. Còn `Rx, Ry` là **độ tròn ở 4 góc**: tưởng tượng cái biển số xe hay khung ứng dụng trên màn hình điện thoại — góc nó không nhọn hoắt như viên gạch mà lượn cong mềm mại, phần lượn cong đó là một góc cung tròn có bán kính ngang `Rx`, bán kính dọc `Ry`. Khi để `Rx = Ry = 0` thì không lượn chút nào, góc vuông lại 90 độ.',
        '**Hình tròn, Hình elip** (`Start, End` và 3 nút chế độ): Hãy tưởng tượng vòng tròn là mặt đồng hồ chia 360 độ, trong Inkscape **vạch 0 độ nằm ở hướng 3 giờ bên phải**, rồi quay xuôi xuống đáy (90 độ là hướng 6 giờ) sang trái (180 độ là hướng 9 giờ). Vì thế khi nhập `Start = 0`, `End = 180`, Inkscape cắt đúng **nửa vòng dưới** tạo luôn dáng miếng dưa hấu đang mỉm cười mà không cần xoay! Còn **3 nút tròn nhỏ** bên cạnh ô `End` giải quyết hình dáng đường đóng: nút 1 (*Hình quạt*) đóng 2 nét thẳng vào tâm giống miếng bánh pizza hay miếng dưa hấu; nút 2 (*Cung tròn hở*) bỏ hết màu nền và nét nối tâm, chỉ còn đúng nét cong như cầu vồng; nút 3 (*Hình tròn kín*) đóng kín toàn bộ lại thành hình tròn nguyên.',
        '**Hình đa giác, Hình sao** (`Corners`, `Spoke ratio`, `Rounded`, `Randomized`): Nhìn 4 hình đỏ ở **Hình 13.2 (tr. 69)**: tất cả không phải 4 công cụ khác nhau đâu, mà sinh ra từ **cùng 1 công cụ Ngôi sao/Đa giác** nhờ vặn 4 nút chỉnh: `Corners` chọn số đỉnh hoặc cánh (ví dụ 5 cánh); bật nút Đa giác đều là cạnh thẳng nối 5 đỉnh thành hình thứ 2 (ngũ giác); ở ngôi sao nhọn, tăng `Rounded` lên thì các đầu nhọn phồng cong thành **cánh hoa** ở hình thứ 3; còn tăng `Randomized` giống như làm rung tay khiến các cánh to nhỏ méo mó ngẫu nhiên thành hình thứ 4. Riêng `Spoke ratio` quyết định ngôi sao mập hay gầy (tỉ lệ khoảng cách từ tâm đến khe lõm trong so với từ tâm ra mũi nhọn ngoài).'
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**3. Lời giảng + Thao tác bấm trên mau_bai13.svg (Khu vực 2) + Tương tác học sinh:**', first=0)
    for item in [
        '**Bước giảng dẫn & khai thác Hình 13.2:** *(Mời cả lớp nhìn Hình 13.2 trang 69 trên màn chiếu)*: *"Nhìn 4 hình màu đỏ ở Hình 13.2: ngôi sao 5 cánh nhọn, hình ngũ giác đều, bông hoa cánh tròn và hình sao méo ngẫu nhiên. Mới nhìn tưởng phải dùng 4 công cụ vẽ khác nhau, nhưng thực chất cả 4 đều do một công cụ duy nhất tạo ra. Các em nhìn biểu tượng ngôi sao cuối bảng đầu trang 70 xem công cụ đó tên gì?"* — Gọi 1 HS đọc tên (**Hình đa giác, hình sao**), rồi giảng ngắn gọn nguyên lý đổi từ hình 1 sang hình 2, 3, 4 như phần Bản chất ở trên.',
        '**Giao việc cặp đôi (02 phút):** *"Bây giờ hai bạn cùng bàn đọc Bảng các thuộc tính cơ bản ở cuối trang 69 và đầu trang 70 SGK. Không cần thuộc vẹt chữ tiếng Anh, các em chỉ cần gạch chân 3 câu trả lời: (1) Hình chữ nhật chỉnh gì để đổi kích thước và bo tròn góc? (2) Hình tròn chỉnh gì và nhập góc bao nhiêu để cắt thành nửa hình tròn như miếng dưa? (3) Hình đa giác, ngôi sao có 4 thông số gì và thông số nào làm cong đỉnh nhọn?"* — Mời lần lượt 3 cặp đứng tại chỗ trả lời (nếu các em ngại đọc tiếng Anh, nhắc cứ nhìn đúng cột bên trái trang 69–70 đọc từng chữ).',
        '**Thao tác 1 trên Khu vực 2 (Chữ nhật xanh):** Nháy **đúp chuột** vào Hình chữ nhật xanh, chỉ chuột lên thanh ngang trên cùng cho lớp thấy các ô `W, H, Rx, Ry` vừa hiện ra. Gõ `25` vào ô `Rx` rồi nhấn **Enter** — cả lớp thấy bốn góc vuông lập tức bo tròn mềm mại.',
        '**Thao tác 2 trên Khu vực 2 (Hình tròn cam):** Nháy **đúp chuột** vào Hình tròn cam, gõ `Start = 0`, `End = 180`, nhấn **Enter** — nửa hình tròn hiện ra ngay! Chỉ chuột sang **3 nút tròn nhỏ** bên cạnh ô `End`, bấm thử sang nút 2 (*Cung tròn hở*) rồi bấm nút 3 (*Hình tròn kín*) rồi hoàn tác về hình quạt.',
        '**Thao tác 3 trên Khu vực 2 (Ngôi sao đỏ):** Nháy **đúp chuột** vào Ngôi sao đỏ, gõ `0.25` vào ô `Rounded`, nhấn **Enter** — 5 cánh nhọn uốn cong thành cánh hoa y hệt Hình 13.2.'
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**4. Nội dung ghi bảng & Đọc cho học sinh chép vở (Ghi ở cột Giữa bảng):**', first=0)
    eng.add(b, '**1. Các đối tượng hình khối và thuộc tính cơ bản (SGK tr. 69–70):**', first=0)
    for item in [
        '`Hình vuông, chữ nhật: W, H (rộng, dài); Rx, Ry (bán kính góc bo; bằng 0 là góc vuông).`',
        '`Hình tròn, elip: Rx, Ry (bán kính ngang, đứng); Start, End (góc điểm đầu, cuối; đặt 0 và 180 tạo nửa hình tròn) và 3 chế độ vẽ (hình quạt, cung tròn hở, hình tròn kín).`',
        '`Hình đa giác, hình sao: Corners (số đỉnh/cánh); Rounded (độ cong đỉnh); Spoke ratio (tỉ lệ bán kính trong/ngoài); Randomized (làm méo ngẫu nhiên).`'
    ]:
        eng.add(b, item, 'bullet')

    # --- CHẶNG 2B ---
    eng.add(b, 'Chặng 2B (Phút 14 – 23): Mục 2 – Tô màu, đường viền & độ trong suốt với Fill and Stroke (SGK tr. 70–71)', 'sub')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    eng.add(b, '**1. Mục đích của chặng này:** Giúp học sinh nâng cấp từ kiểu tô màu bệt một sắc (ở Bài 12) lên hiểu cơ chế phối màu chuyên nghiệp của dân thiết kế: biết tách riêng **màu ruột** (`Fill`), **màu viền** (`Stroke`), **hiệu ứng sáng tối 3D** (`Gradient`), **kiểu nét viền** (`Style`), và **độ xuyên thấu** (`Opacity`).', first=0)
    eng.add(b, '**2. Bản chất cần làm rõ để không bị lẫn lộn:**', first=0)
    for item in [
        '*Tại sao phải mở hộp thoại Fill and Stroke khi dưới đáy đã có dải màu?* Dải màu dưới đáy chỉ giống hộp bút sáp: chọn 1 màu thì tô đặc một màu đều tăm tắp. Muốn làm màu chuyển nhạt dần để tạo chiều sâu khối cầu, làm viền nét đứt, hoặc làm hình mờ xuyên thấu như tấm kính thì buộc phải mở bảng **Fill and Stroke** (vào `Object / Fill and Stroke`, chuột phải chọn `Fill and Stroke`, hoặc nhanh nhất nhấn **Shift + Ctrl + F**).',
        '*Hiểu đúng nghĩa từng chữ tiếng Anh trong 3 thẻ (Tab):* **Fill** nghĩa là *lấp đầy* (màu sơn phủ kín ruột bên trong); **Stroke paint** là *màu nét vẽ* (màu của đường viền bao quanh); **Stroke style** là *kiểu dáng nét vẽ* (đường viền dày hay mảnh, nét liền hay nét đứt khúc).',
        '*Phân biệt Linear gradient và Radial gradient:* `Linear` là đường thẳng — màu chuyển từ trái sang phải hoặc trên xuống dưới giống ánh nắng chiếu ngang; `Radial` là toả tròn từ tâm — giữa sáng chói rồi đậm dần ra rìa giúp một hình tròn phẳng biến thành **quả cầu 3D có độ nổi khối**.',
        '*Hiểu đúng nút dấu ? và độ trong suốt Opacity / A:* Dấu `?` (*Unset*) dùng để **gỡ bỏ thiết lập màu riêng của đối tượng** về trạng thái chưa đặt (hay dùng khi copy hoặc tái sử dụng nhiều đối tượng giống nhau). Còn `Opacity` (và kênh `A` – Alpha) giống độ đục của **tấm kính**: 100% là ván gỗ đặc không nhìn xuyên được, giảm xuống 50% thành kính mờ thấy hình lớp dưới, về 0% là tàng hình hoàn toàn. *(Điểm khác nhỏ trên máy: SGK cũ vẽ biểu tượng cây bút chỉnh gradient trong hộp thoại, còn bản Inkscape mới trên máy tính thầy dùng phím tắt chữ G hoặc nút Gradient Hàng 4 Cột 1 để kéo trực tiếp thanh chỉnh màu trên hình).* '
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**3. Lời giảng + Thao tác bấm trên mau_bai13.svg (Khu vực 3) + Tương tác học sinh:**', first=0)
    for item in [
        '**Kết nối Bài 12:** *"Ở Bài 12 muốn đổi màu nhanh, các em chỉ cần nháy chuột vào dải màu đáy màn hình để đổi màu nền, giữ Shift nháy chuột để đổi màu viền. Nhưng nhìn lên 5 hình vuông ở Hình 13.3 trang 70 này: tại sao 5 hình đều có viền cam mà phần ruột bên trong lại trông hoàn toàn khác nhau — có hình rỗng ruột, có hình chuyển màu đậm sang nhạt, có hình toả sáng từ giữa tâm?"*',
        '**Tương tác khai thác Hình 13.3 (tr. 70):** Mời 01 học sinh nhìn 5 hình vuông và đọc 5 dòng chú thích ngay dưới Hình 13.3 để phân biệt: (1) Không màu dấu X; (2) Màu đồng nhất; (3) Chuyển màu theo chiều ngang (**Linear gradient**); (4) Chuyển màu từ tâm ra ngoài (**Radial gradient**); (5) Chuyển màu có độ trong suốt (nhờ kéo `Opacity` hoặc kênh `A`).',
        '**Giải thích Hình 13.4 trên bảng Fill and Stroke bên phải màn hình:** *"Để chỉnh sâu như vậy, ta nhấn Shift + Ctrl + F mở hộp thoại Fill and Stroke gồm 3 ngăn rõ ràng: ngăn 1 là Fill chỉnh màu nền bên trong (có dấu X là không màu, đồng nhất, chuyển thẳng Linear, chuyển từ tâm Radial, hoa văn Pattern và dấu hỏi chấm để huỷ đặt); ngăn 2 là Stroke paint chỉnh màu đường viền (có thêm ống hút sao chép màu từ bất kỳ điểm nào); ngăn 3 là Stroke style chỉnh độ dày Width và kiểu nét đứt."*',
        '**Thao tác 1 trên Khu vực 3 (Tô chuyển màu 3D):** Bấm chọn **Hình tròn xanh** ở Khu vực 3, trong thẻ `Fill` ở bảng bên phải bấm nút thứ tư **Radial gradient** — cả lớp thấy ngay tâm hình tròn rực sáng nổi khối 3D! Bấm phím **G** trên bàn phím cho cả lớp thấy 2 chấm tròn và đường điều hướng màu chuyển hiện ra trên hình, kéo nhẹ đầu chấm tròn để minh hoạ việc đổi vị trí tâm sáng (rồi bấm `S`).',
        '**Thao tác 2 trên Khu vực 3 (Đổi viền nét đứt):** Chuyển sang thẻ **Stroke style** ở khung bên phải, tăng nhẹ `Width` rồi bấm mục **Dashes** chọn một kiểu nét gạch đứt — đường viền xung quanh biến thành nét đứt ngay lập tức.',
        '**Thao tác 3 trên Khu vực 3 (Tấm kính trong suốt):** Chọn **Hình tròn xanh dương** đang nằm đè lên hình chữ nhật vàng. Kéo thanh trượt **Opacity (%)** ở góc dưới cùng bên phải từ 100% xuống khoảng 50% — hình chữ nhật vàng ở lớp dưới lập tức hiện mờ xuyên qua hình tròn xanh!'
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**4. Nội dung ghi bảng & Đọc cho học sinh chép vở (Ghi ở cột Giữa bảng):**', first=0)
    eng.add(b, '**2. Thiết lập màu tô, màu vẽ – Hộp thoại Fill and Stroke (SGK tr. 70–71):**', first=0)
    for item in [
        '`Cách mở: Object / Fill and Stroke, chuột phải chọn Fill and Stroke, hoặc Shift + Ctrl + F.`',
        '`3 trang chính: Fill (X không màu, đồng nhất, Linear gradient, Radial gradient, hoa văn Pattern, ? huỷ đặt); Stroke paint (màu viền, ống hút sao chép màu); Stroke style (độ dày Width, kiểu nét liền/đứt).`',
        '`Biểu tượng cây bút (công cụ Gradient) chỉnh thông số chuyển màu; Opacity và kênh A (alpha) chỉnh độ trong suốt.`'
    ]:
        eng.add(b, item, 'bullet')

    # --- CHẶNG 2C ---
    eng.add(b, 'Chặng 2C (Phút 23 – 33): Mục 3 – 06 phép ghép cắt đối tượng trong menu Path (SGK tr. 71–72)', 'sub')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add(b, '**1. Mục đích của chặng này:** Đây là vũ khí mạnh nhất của Bài 13. Học sinh hiểu được thay vì phải vẽ tự do những đường cong khó nhằn (trái tim, đám mây, quả táo bị cắn khuyết, thân cây uốn eo), dân thiết kế chỉ vẽ **những hình tròn, hình vuông đơn giản xếp đè lên nhau rồi dùng lệnh trong menu** `Path` để hàn lại hoặc đục cắt.', first=0)
    eng.add(b, '**2. Bản chất cần khắc sâu (Ví dụ khuôn làm bánh + Tập hợp Toán 10):**', first=0)
    for item in [
        '*Tại sao xếp đè 3 hình vào nhau chưa phải ghép Path?* Nếu chỉ xếp 2 hình tròn + 1 hình vuông nghiêng 45 độ như **Hình 13.5** mà không bấm `Union`, khi di chuyển từng hình vẫn bị rơi rời ra hoặc đường viền của hình bên trong cắt ngang giữa trái tim rất xấu. Bấm `Path / Union` sẽ nấu chảy và hàn kín cả 3 hình thành **một miếng trái tim liền khối duy nhất**.',
        '*Quy tắc Khuôn dập và Tấm bột (CỰC QUAN TRỌNG để giải thích Difference / Division / Cut Path):* Hãy dặn học sinh tưởng tượng **Hình nằm ở LỚP TRÊN là chiếc khuôn dập bằng sắt (hoặc con dao)**, còn **Hình nằm ở LỚP DƯỚI là tấm bột bánh (hoặc tờ giấy)**. Khi ép khuôn xuống (bấm `Difference`, `Division`, `Cut Path`), **tấm bột ở LỚP DƯỚI luôn là vật bị cắt**, còn khuôn sắt ở lớp trên sau khi cắt xong sẽ biến mất!',
        '*Liên hệ Tập hợp Toán lớp 10 để nhớ Hình 13.6 trong 10 giây:* Quy ước Hình vuông nằm lớp **DƯỚI**, Hình tròn nằm lớp **TRÊN** (như ô đầu Hình 13.6): `Union` giống phép **Hợp (A ∪ B)** lấy hết cả hai; `Difference` giống phép **Hiệu (A \\ B)** lấy phần vuông trừ bỏ phần tròn đè lên; `Intersection` giống phép **Giao (A ∩ B)** chỉ giữ phần bụng chung ở giữa.'
    ]:
        eng.add(b, item, 'bullet')

    path_deep_rows = [
        ['Phép ghép (Menu Path)', 'Phím tắt', 'Ví von trực quan (Vuông DƯỚI – Tròn TRÊN)', 'Kết quả nhìn thấy & Ứng dụng'],
        ['**1. Phép hợp**\n(Union)', 'Ctrl + +', 'Hàn hai miếng kim loại nóng chảy thành một tấm liền.', 'Giữ toàn bộ phần của cả hai hình, xoá ranh giới bên trong. Dùng ghép **Trái tim H.13.5**, **Đám mây H.13.7**.'],
        ['**2. Phép hiệu**\n(Difference)', 'Ctrl + -', 'Dùng hình tròn TRÊN làm **khuôn đục** khoét một miếng trên hình vuông DƯỚI.', 'Hình vuông bị khoét khuyết hình tròn; hình tròn biến mất. Dùng khoét **Logo Apple cắn dở**, gọt thân cây.'],
        ['**3. Phép giao**\n(Intersection)', 'Ctrl + *', 'Chỉ giữ lại đúng **phần đè lên nhau** của hai hình.', 'Cả hai phần dư bên ngoài đều mất sạch. Hai hình tròn chồng nhau giao ra hình chiếc lá/con mắt.'],
        ['**4. Phép hiệu đối xứng**\n(Exclusion)', 'Ctrl + ^', 'Giữ hai cánh bên ngoài, **đục thủng phần giao nhau** ở chính giữa.', 'Ngược lại hoàn toàn với phép Giao: ở giữa bị thủng lỗ trống.'],
        ['**5. Phép chia**\n(Division)', 'Ctrl + /', 'Lấy viền hình tròn TRÊN làm **dao cắt pizza** xẻ hình vuông DƯỚI.', 'Hình vuông dưới bị cắt rời thành các mảnh độc lập **vẫn giữ nguyên màu tô**, kéo tách từng mảnh ra được.'],
        ['**6. Phép cắt**\n(Cut Path)', 'Ctrl + Alt + /', 'Chỉ cắt đứt **sợi dây viền** của hình vuông DƯỚI tại 2 giao điểm.', 'Chỉ còn đường viền bị chặt đứt thành đoạn rời, **mất sạch màu tô nền** bên trong.']
    ]
    b.append(eng.table(path_deep_rows, [1950, 1250, 2900, 2971], header=True, cell_sz=21, cant_split=True))

    eng.add(b, '**3. Lời giảng + Thao tác bấm trên mau_bai13.svg (Khu vực 4B & 4C) + Tương tác học sinh:**', first=0)
    for item in [
        '**Dẫn từ Hình 13.5 (tr. 71):** *"Các em nhìn Hình 13.5: có hai hình tròn và một hình vuông. Làm sao xếp 3 hình giấy này thành hình trái tim? Đầu tiên ta bấm hai lần vào hình vuông để xoay nghiêng 45 độ thành hình thoi, đặt 2 hình tròn lên hai cạnh trên của hình thoi rồi quét chọn cả 3 hình, vào Path chọn Union để hàn dính thành trái tim liền nét."*',
        '**Khai thác Hình 13.6 (tr. 72):** Yêu cầu cặp đôi đối chiếu Hình 13.6 với chữ in đậm trang 71–72 (có nhãn tiếng Việt sẵn dưới mỗi hình) để gọi đúng tên tiếng Anh + tổ hợp phím tắt + mô tả kết quả của cả 6 phép ghép. Sau đó **đặt câu hỏi chốt sống còn:** *"Quan sát Phép hiệu, Phép chia và Phép cắt trong Hình 13.6: trong hình gốc có hình vuông dưới và hình tròn trên, sau khi thực hiện lệnh thì hình nào bị cắt xẻ, hình nào là khuôn cắt?"* — Mời 1 HS trả lời: **Đối tượng lớp dưới bị cắt bởi đối tượng lớp trên**.',
        '**Thao tác trực tiếp trên Khu vực 4B (KHÔNG bấm Ctrl + + mà bấm thẳng menu Path cho cả lớp nhìn rõ):** Chọn Mũi tên (**S**), kéo chuột quét khung nét đứt bao trọn hình vuông đỏ (lớp dưới) và hình tròn cam (lớp trên). Mở menu **Path** trên cùng, bấm **Intersection** — cả lớp thấy chỉ còn đúng một mẩu phần giao ở giữa! Bấm **Ctrl + Z** trả lại ban đầu, rồi mở menu **Path** bấm **Difference** — hình vuông bị khoét khuyết tròn ngay! *(Nhắc thêm cho lớp nhớ: Nếu lỡ thấy mất nhầm hình muốn giữ, cứ nhấn Ctrl + Z, chọn hình cần giữ nhấn phím End để đẩy nó xuống lớp dưới cùng rồi cắt lại).* ',
        '**Thao tác trực tiếp trên Khu vực 4C (Logo Apple thực tế — cực hút học sinh):** Chỉ vào hình quả táo đen (lớp dưới) có hình tròn xám đè lệch bên mép phải (lớp trên). Quét chọn cả 2 hình, vào menu **Path** bấm **Difference** — vết khuyết cắn dở của quả táo Apple hiện ra sắc nét ngay tức khắc! Giải thích thêm chiếc lá trên đầu là kết quả phép Giao (**Intersection**) của 2 hình tròn.',
        '**Mở rộng thực tế ngắn (30 giây cho bài học có chiều sâu):** *"Vì sao dân quảng cáo làm biển hiệu, máy cắt Laser/CNC hay cần file vẽ từ Inkscape (đuôi .svg) thay vì ảnh chụp bình thường? Vì máy cắt chạy theo đúng các đường Path vectơ toán học vừa học nên phóng to cỡ nào cạnh cắt cũng sắc lẹm không vỡ hạt, khác hoàn toàn ảnh điểm ảnh khi nhờ AI vẽ ra nếu chỉ tải file .png."* *(Cũng chính các đường Path này sẽ là cầu nối sang Bài 14 học chỉnh từng điểm nốt).* '
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**4. Nội dung ghi bảng & Đọc cho học sinh chép vở (Ghi ở cột Giữa bảng):**', first=0)
    eng.add(b, '**3. Các phép ghép đối tượng đồ hoạ trong bảng chọn Path (SGK tr. 71–72):**', first=0)
    for item in [
        '`Phép hợp (Union, Ctrl + +): Gồm tất cả phần thuộc các hình đơn.`',
        '`Phép hiệu (Difference, Ctrl + -): Phần thuộc lớp dưới nhưng không thuộc lớp trên.`',
        '`Phép giao (Intersection, Ctrl + *): Chỉ giữ phần thuộc cả hai hình.`',
        '`Ngoài ra còn: Phép hiệu đối xứng (Exclusion, Ctrl + ^); Phép chia (Division, Ctrl + /); Phép cắt (Cut Path, Ctrl + Alt + /).`',
        '`Quy tắc cần nhớ: Trong Difference, Division và Cut Path, đối tượng lớp DƯỚI luôn bị cắt bởi đối tượng lớp TRÊN.`'
    ]:
        eng.add(b, item, 'bullet')

    # --- CHẶNG 3 ---
    eng.add(b, 'Chặng 3 (Phút 33 – 40): Luyện tập – Kiểm chứng kiến thức qua 04 câu hỏi SGK (tr. 70, 71, 72 & Hình 13.7)', 'sub')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    eng.add(b, '**1. Mục đích của chặng này:** Kiểm tra xem học sinh có thực sự biết vận dụng 3 phần lí thuyết vừa học vào đúng **04 câu hỏi củng cố chính thức của SGK** hay không (khớp hoàn toàn tiêu chí 1 & 2 của Rubric trong `khbd10-3.docx`). Không kiểm tra đánh đố ngoài bài, dùng ngay máy chiếu kiểm chứng đáp án trước mắt lớp.', first=0)
    eng.add(b, '**2. Cách dẫn dắt từng câu hỏi & Kiểm chứng trên Inkscape (Khu vực 4A):**', first=0)
    for item in [
        '**Câu 1 (SGK tr. 70):** *"Nếu trên hình vẽ đang có sẵn một ngôi sao 5 cánh nhọn, muốn các đỉnh nhọn cong tròn mềm mại như bông hoa, ta sửa thông số nào?"* — Mời 01 HS đứng tại chỗ trả lời: Tham số **Rounded**. *(Bấm nhanh vào Ngôi sao ở Khu vực 2 kiểm chứng lại nếu cần).* ',
        '**Câu 2 (Trắc nghiệm 1 – SGK tr. 71):** *"Muốn đổi đường viền thành nét đứt, chọn ngăn nào trong hộp thoại Fill and Stroke? A. Fill, B. Stroke paint, C. Stroke style, D. Cả A và B. Thầy đếm 1, 2, 3 cả lớp đồng loạt giơ số ngón tay nhé!"* — Cả lớp giơ **3 ngón: C. Stroke style**, gọi 1 HS giải thích ngắn (vì `Stroke style` chỉnh độ dày và kiểu đường viền nét liền hay nét đứt).',
        '**Câu 3 (Trắc nghiệm 2 – SGK tr. 71):** *"Để chỉnh thông số của màu chuyển gradient, chọn biểu tượng nào trong 4 hình A, B, C, D ở trang 71?"* — Cả lớp giơ **2 ngón: B** (biểu tượng cây bút chỉnh gradient).',
        '**Câu 4 (SGK tr. 72 – Hình 13.7):** Chiếu **Hình 13.7 (Vẽ đám mây)** lên màn chiếu: *"Hai bạn cùng bàn trao đổi 02 phút: để vẽ được đám mây hoàn chỉnh như Hình 13.7, ta dùng phép ghép nào trong menu Path và nêu đủ 3 bước thực hiện?"* — Mời 1–2 HS trình bày đủ: Dùng **Phép hợp** (`Path / Union`), gồm 3 bước: (1) Vẽ các hình tròn xếp chồng lấp lên nhau tạo dáng đám mây; (2) Quét chọn tất cả các hình tròn vừa vẽ; (3) Chọn **Path / Union** để hợp nhất thành một khối đám mây.',
        '**Bấm kiểm chứng ngay trên Khu vực 4A của mau_bai13.svg:** Chọn Mũi tên (**S**), kéo khung nét đứt quét chọn cả 4 hình tròn màu xanh ở **ô A**, bấm menu **Path rồi bấm Union** — toàn bộ đường viền rối bên trong biến mất, hoá thành một đám mây xanh liền nét chỉ trong 1 cái nhấp chuột!'
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**3. Xử lý nếu lớp im lặng:** Mở lại đúng số trang trong SGK (`tr. 70`, `tr. 71`, `tr. 72`), chỉ rõ dòng bảng hoặc hình minh hoạ để học sinh đối chiếu đọc đáp án.', first=0)
    eng.add(b, '**4. Ghi góc phải bảng:** Ghi nhanh 4 đáp án chốt: `Câu 1: Rounded | Câu 2: C (Stroke style) | Câu 3: B (cây bút) | Câu 4: Path / Union (3 bước vẽ đám mây).`', first=0)

    # --- CHẶNG 4 ---
    eng.add(b, 'Chặng 4 (Phút 40 – 45): Vận dụng vẽ Miếng dưa hấu (Hình 13.1) & Mở đường sang Tiết 2 Thực hành', 'sub')
    eng.add(b, '**1. Mục đích của chặng này:** Khép lại trọn vẹn vòng tròn sư phạm đã mở ra từ phút 00 (lúc đầu chưa vẽ được miếng dưa hấu, đến cuối giờ giải quyết gọn trong 3 bước), đồng thời **giao cầu nối rõ ràng sang Tiết 2 (Thực hành tr. 72–74)** để học sinh biết tiết sau mình sẽ ngồi máy thực hành những gì.', first=0)
    eng.add(b, '**2. Lời giảng, Tương tác chốt Miếng dưa hấu & Minh hoạ Inkscape (Khu vực 1 bên phải):**', first=0)
    for item in [
        '**Giải quyết trọn vẹn bài toán Mở đầu (Hình 13.1 – liên hệ Vận dụng 1 tr. 74):** *"Còn 5 phút cuối, ta quay lại miếng dưa hấu đầu giờ. Với 3 mảng kiến thức vừa học, một bạn đứng lên đề xuất 3 bước rõ ràng để vẽ được miếng dưa hấu này?"* — Chốt 3 bước: (1) Chọn công cụ Hình tròn (**E**), đặt `Start = 0`, `End = 180` ở chế độ hình quạt (Slice) tô xanh lá đậm làm **vỏ dưa** (lớp dưới cùng); (2) Nhấn `Ctrl + D` nhân bản, kéo thu nhỏ và tô màu trắng làm **cùi dưa**, rồi tiếp tục nhấn `Ctrl + D`, kéo thu nhỏ và tô màu đỏ làm **ruột dưa**; (3) Vẽ thêm các hình elip nhỏ màu đen đặt lên **lớp trên cùng** làm **hạt dưa**.',
        '**Thao tác dựng thật trước mặt lớp (Khu vực 1 bên phải của mau_bai13.svg – khoảng 25–30 giây):** Chọn hình tròn xanh chờ sẵn bên phải Khu vực 1, nhấn **E**, gõ `Start = 0, End = 180` nhấn **Enter**; bấm **S**; bấm **Ctrl + D** kéo thu nhỏ rồi nhấp ô màu trắng dưới đáy; bấm **Ctrl + D** kéo thu nhỏ rồi nhấp ô màu đỏ; vẽ hoặc kéo 2 hạt elip đen đặt lên trên! *(Nếu chuông sắp reo còn dưới 2 phút, chỉ cần cắt nửa hình tròn và Ctrl + D tô ruột đỏ 1 lần rồi chốt bài để không lố giờ).* '
    ]:
        eng.add(b, item, 'bullet')
    eng.add(b, '**3. Định hướng công cụ cho 03 bài Thực hành Tiết 2 (SGK trang 72–74) & Kết thúc:**', first=0)
    for item in [
        '**Nói (cho HS lật lướt trang 72–74 xem tên 3 nhiệm vụ):** *"Để Tiết 2 xuống phòng thực hành các em làm nhanh mà không bỡ ngỡ, hãy nhìn lướt tên 3 bài thực hành từ trang 72 đến 74 và chỉ ra xem mỗi bài dùng lại công cụ nào ta vừa học hôm nay?"*',
        '*Nhiệm vụ 1 – Vẽ Logo 3 cánh (Hình 13.8, tr. 72–73):* Dùng chế độ **cung tròn hở** (cắt bởi `Start, End`) cùng với tăng độ dày `Width` trong `Stroke style` và màu viền `Stroke paint`.',
        '*Nhiệm vụ 2 – Vẽ Hình cây xanh (Hình 13.10, tr. 73):* Dùng `Path / Difference` khoét thắt eo thân cây và `Path / Union` hàn tán lá với thân thành 1 khối.',
        '*Nhiệm vụ 3 – Vẽ Hình cầu 3D (Hình 13.12, tr. 73–74):* Bỏ đường viền (`X`) và phối màu chuyển `Linear gradient`, `Radial gradient` trong thẻ `Fill` (chỉ lướt qua hình Quả cầu 3D mẫu ở góc dưới phải Khu vực 5 để HS thấy hiệu ứng ánh sáng 3D cực đẹp!).',
        '**Dặn dò kết thúc đúng giờ:** *"Về nhà các em xem lại 3 mục lí thuyết trang 69–72 trong vở ghi và đọc trước hướng dẫn 3 nhiệm vụ thực hành trang 72–74 để tiết sau thao tác ngay trên máy. Cảm ơn các em, mời cả lớp nghỉ!"* — Tắt màn chiếu gọn gàng, xoá bảng sạch sẽ.'
    ]:
        eng.add(b, item, 'bullet')

    # ==========================================================================
    # PHẦN IV. BỐ CỤC TRÌNH BÀY BẢNG PHẤN & CHECKLIST 5 PHÚT TRƯỚC GIỜ
    # ==========================================================================
    eng.br(b)
    eng.add(b, 'PHẦN IV. BỐ CỤC BẢNG PHẤN (3 CỘT) & CHECKLIST 5 PHÚT TRƯỚC KHI VÀO TIẾT', 'heading')
    eng.add(b, '1. Sơ đồ trình bày Bảng phấn trên lớp (Viết tuần tự từ đầu đến cuối tiết, gọn sạch không phải xoá bảng)', 'sub')
    board_rows = [
        ['GÓC TRÁI BẢNG (Thứ Hai, 16/11/2026)', 'GIỮA BẢNG (Nội dung HS ghi vào vở)', 'GÓC PHẢI BẢNG (Nháp & Đáp án)'],
        [
            'Tiết PPCT 19:\n'
            '**BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ**\n'
            '*(Tiết 1 – Lí thuyết)*\n'
            '-----------------\n'
            '3 trọng tâm hôm nay:\n'
            '1. Thuộc tính hình khối\n'
            '2. Tô màu & nét vẽ (Fill and Stroke)\n'
            '3. Phép ghép cắt hình (Path)',
            '**1. Thuộc tính hình khối (SGK tr. 69–70):**\n'
            '- Vuông, chữ nhật: W, H (rộng, dài); Rx, Ry (bán kính góc bo; bằng 0 nếu góc vuông).\n'
            '- Tròn, elip: Rx, Ry (bán kính ngang, đứng); Start, End (0 và 180 tạo nửa tròn) + 3 chế độ quạt, cung hở, tròn kín.\n'
            '- Đa giác, sao: Corners (số đỉnh/cánh), Rounded (cong đỉnh), Spoke ratio (tỉ lệ trong/ngoài), Randomized (méo ngẫu nhiên).\n\n'
            '**2. Hộp thoại Fill and Stroke (Shift + Ctrl + F):**\n'
            '- Fill: Màu tô nền (X, đồng nhất, Linear/Radial gradient, Pattern, dấu ?).\n'
            '- Stroke paint: Màu viền (có ống hút sao chép màu).\n'
            '- Stroke style: Độ dày (Width) và kiểu nét đứt/liền.\n'
            '- Biểu tượng cây bút (Gradient) chỉnh dải màu; Opacity / A chỉnh độ trong suốt.\n\n'
            '**3. Các phép ghép trong bảng chọn Path (tr. 71–72):**\n'
            '- Trọng tâm: Union (hợp, Ctrl + +), Difference (hiệu, Ctrl + -), Intersection (giao, Ctrl + *).\n'
            '- Khác: Exclusion, Division, Cut Path.\n'
            '- Quy tắc: Trong Difference, Division, Cut Path thì **ĐỐI TƯỢNG LỚP DƯỚI luôn bị cắt bởi ĐỐI TƯỢNG LỚP TRÊN**.',
            '**Nhắc lại Bài 12:**\n'
            '- Vẽ trước nằm DƯỚI\n'
            '- Vẽ sau nằm TRÊN\n\n'
            '**Đáp án Luyện tập (tr. 70–72):**\n'
            '- Câu 1 (tr. 70): Rounded\n'
            '- Câu 2 (tr. 71): C. Stroke style\n'
            '- Câu 3 (tr. 71): B (cây bút)\n'
            '- Câu 4 (tr. 72 – H.13.7): Path / Union (3 bước vẽ đám mây)\n\n'
            '**Chuẩn bị Tiết 2 (tr. 72–74):**\n'
            'Logo (H.13.8) | Cây (H.13.10) | Cầu 3D (H.13.12)'
        ]
    ]
    b.append(eng.table(board_rows, [2150, 4671, 2250], header=True, cell_sz=21, cant_split=True))

    eng.add(b, '2. Checklist 5 phút chuẩn bị trước khi vào tiết (Để lên bục giảng tự tin 100%)', 'sub')
    for item in [
        '**Bước 1 – Mở sẵn mau_bai13.svg:** Mở Inkscape trước giờ vào tiết để nạp sẵn phần mềm, mở tệp `mau_bai13.svg` và nhấn phím **5** cho vừa màn hình.',
        '**Bước 2 – Kiểm tra bảng Fill and Stroke:** Nhìn khung bên phải xem `Fill and Stroke` đã mở sẵn chưa (nếu chưa thì nhấn **Shift + Ctrl + F**).',
        '**Bước 3 – Vẽ thử 1 hình tròn nháp:** Kiểm tra nhanh góc dưới trái xem `O:` đang là **100** và trên thanh ngang đang để nút **hình tròn khép kín**, sau đó bấm **Delete** xoá hình nháp.',
        '**Bước 4 – Xếp sẵn tài liệu:** Mở tệp kịch bản này (hoặc bản in cầm tay) gọn gàng trên bàn giáo viên để liếc nhanh theo nhịp 6 chặng.'
    ]:
        eng.add(b, item, 'bullet')

    out_path = DAY_THAT_DIR / 'Kich_ban_thuyet_trinh_va_tap_duyet_Bai13_Tiet1.docx'
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
    build_kich_ban_tap_duyet_toan_dien()
    print('Done building all 2 Day_that docx files in:', DAY_THAT_DIR)

