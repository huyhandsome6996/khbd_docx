# -*- coding: utf-8 -*-
"""
Bo tao ho so DAY THAT cho tiet giang day thuc te Lop 10/3:
- Thoi gian: Thu Hai, ngay 16/11/2026 - Tiet 3 (09h00 - 09h45)
- Dia diem: Phong H.I.5 (Lop hoc thuong: 1 may tinh GV + may chieu + bang phan; HS khong co may tinh ca nhan)
- Bai day: Tin hoc 10 (KNTT) - Bai 13: Bo sung cac doi tuong do hoa (Tiet 1 - Li thuyet / Tiet PPCT: 19)
- GVHD / GVGD: Thay Mai Xuan Van | SVTT: Ho Quang Huy (24S1020030)

Xuat ra 02 tep trong thu muc c:/Users/PC/khbd_docx/Day_that/:
  1. 01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx
     (Ke hoach bai day chuan CV 5512, thuc te 100%, de thuc hien voi moi trinh do HS, loai bo hinh thuc bat kha thi)
  2. 02_Kich_ban_len_lop_chi_tiet_vui_nhon_Tin10_Bai13_Tiet1.docx
     (Cam nang cam tay chi viec: noi gi, lam gi, bam nut nao, pha tro gi, chua chay khi HS im re & kien thuc AI/Apple/CNC)
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
# FILE 1: KE HOACH BAI DAY CHUAN CV 5512 - THUC TE 100% - DE THUC HIEN (10/3)
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
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết tại lớp kết hợp thao tác trực quan trên máy chiếu và vấn đáp SGK'
    )

    # I. MỤC TIÊU
    eng.add(b, 'I. MỤC TIÊU', 'heading')
    eng.add(b, 'Căn cứ Yêu cầu cần đạt Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E, khung Năng lực số (Thông tư số 02/2025/TT-BGDĐT) và Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá (Tiết PPCT 19 – Mục 4.13. Bổ sung các đối tượng đồ họa, tiết 1 - LT; chi tiết trích lục đính kèm tại Phụ lục 1, Mục IV. PHỤ LỤC), phù hợp điều kiện phòng học lí thuyết (không có máy tính học sinh):')

    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu `[KT1]` – `[KT3]`):', keep=True)
    for item in [
        '**`[KT1]`**: Nhận biết và nêu được ý nghĩa các thuộc tính cơ bản của ba nhóm hình khối trên Thanh điều khiển thuộc tính: `W, H, Rx, Ry` (hình vuông/chữ nhật); `Rx, Ry, Start, End` (hình tròn/elip); `Corners, Rounded, Spoke ratio, Randomized` (hình đa giác/ngôi sao) theo bảng trang 69–70 SGK.',
        '**`[KT2]`**: Nêu được cách mở và chức năng ba trang của hộp thoại **`Fill and Stroke`** (`Object / Fill and Stroke` hoặc `Shift + Ctrl + F`): trang `Fill` (màu tô nền), `Stroke paint` (màu viền), `Stroke style` (độ dày, kiểu nét vẽ); phân biệt được màu đồng nhất (`Flat color`), màu chuyển (`Linear gradient`, `Radial gradient`) và độ trong suốt (`Opacity`).',
        '**`[KT3]`**: Nhận diện được kết quả của 06 phép ghép đối tượng trong bảng chọn **`Path`** (Hình 13.6 SGK trang 72), trọng tâm là ba phép phổ biến: Phép hợp (`Union`, `Ctrl + +`), Phép hiệu (`Difference`, `Ctrl + -`) và Phép giao (`Intersection`, `Ctrl + *`), nắm được quy tắc hình lớp dưới bị cắt bởi hình lớp trên.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**`[NLC1 – Tự chủ và tự học]`**: Chủ động quan sát hình ảnh trên màn chiếu và tra cứu trực tiếp bảng, hình minh hoạ trong SGK trang 69–72 để trả lời câu hỏi.',
        '**`[NLC2 – Giao tiếp và hợp tác]`**: Trao đổi nhanh với bạn cùng bàn trong 2 phút để thống nhất câu trả lời trắc nghiệm và nhận diện phép ghép hình.',
        '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Đề xuất được công cụ, tham số hoặc phép ghép phù hợp để tạo hình đơn giản từ các bài tập trong SGK (củng cố trang 70, 71, 72, Luyện tập trang 74).'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**`[NLa – Sử dụng và quản lí các phương tiện CNTT]`**: Nhận biết đúng vị trí Thanh điều khiển thuộc tính, hộp thoại `Fill and Stroke` và bảng chọn `Path` trên màn hình `Inkscape` khi giáo viên thị phạm.',
        '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT]`**: Phân tích được một hình vẽ đơn giản (miếng dưa hấu, đám mây, bông hoa) thành các hình khối cơ bản, lớp vẽ và phép ghép tương ứng.'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'sub')
    eng.add(b, '**`[NLS 3.2.NC1.a]`**: Hiểu cách điều chỉnh thuộc tính đối tượng vectơ (kích thước, độ trong suốt, màu chuyển, ghép khối) làm cơ sở cho giờ thực hành tạo sản phẩm số ở Tiết 2.', 'bullet')

    eng.add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**`[PC1 – Chăm chỉ]`**: Chú ý theo dõi phần thao tác mẫu trên màn chiếu, mở đúng trang SGK và ghi chép đầy đủ ý chính.',
        '**`[PC2 – Trách nhiệm]`**: Tích cực tham gia giơ tay chọn đáp án, giữ trật tự và hợp tác trong các hoạt động ngắn tại lớp.',
        '**`[PC3 – Trung thực]`**: Tự suy nghĩ chọn đáp án trắc nghiệm và thẳng thắn nêu chỗ chưa hiểu để giáo viên giải đáp.'
    ]:
        eng.add(b, item, 'bullet')

    # II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính giáo viên cài sẵn phần mềm **`Inkscape`** (đã đặt lại chế độ tô màu mặc định `Opacity = 100%`, mở sẵn hộp thoại `Fill and Stroke`), máy chiếu tại lớp học, bảng viết, phấn.',
        '**Học liệu chuẩn (SGK):** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống (NXB Giáo dục Việt Nam)*, **Bài 13: Bổ sung các đối tượng đồ hoạ** (trang 69–74).',
        '**Tệp trình chiếu & minh hoạ của giáo viên:** Bộ câu hỏi hỏi-đáp tra cứu SGK và câu hỏi trắc nghiệm hiển thị trực tiếp trên màn chiếu (sử dụng ngay SGK học sinh mang theo, không phát phiếu giấy rời gây mất thời gian; hệ thống câu hỏi và trích lục căn cứ đính kèm tại Phụ lục 1, Phụ lục 2, Phụ lục 3 ở Mục IV. PHỤ LỤC); tệp `mau_bai13.svg` chuẩn bị sẵn để giáo viên thao tác thị phạm trực tiếp trên `Inkscape`.'
    ]:
        eng.add(b, item, 'bullet')

    # III. TIẾN TRÌNH DẠY HỌC
    eng.add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    eng.add(b, '**Bảng tổng quan tiến trình dạy học Tiết 1 – Lí thuyết (45 phút – Tiết 1 của bài / Tiết PPCT: 19)**', first=0, keep=True)
    ov_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng', 'Mã mục tiêu', 'Sản phẩm & Công cụ đánh giá (Khả thi 100%)'],
        ['1. Hoạt động 1: Mở đầu\n(Quan sát Hình 13.1 – Miếng dưa hấu)', '05 phút\n(PPCT 19)', '[KT1], [NLC1],\n[PC1]', 'Câu trả lời miệng nêu tên hình và thứ tự 4 lớp từ dưới lên.\nCông cụ đánh giá: CCĐG 1 (Quan sát trả lời miệng).'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Thuộc tính hình khối (09p)\n- 2.2. Nhiệm vụ 2: Fill and Stroke (09p)\n- 2.3. Nhiệm vụ 3: 6 phép ghép Path (10p)', '28 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2], [NLC3],\n[NLa], [NLc],\n[NLS 3.2.NC1.a],\n[PC1], [PC2], [PC3]', 'Học sinh tra bảng SGK, giơ tay chọn trắc nghiệm, ghép cặp hình và ghi vở ý chính.\nCông cụ đánh giá: CCĐG 2, CCĐG 3, CCĐG 4.'],
        ['3. Hoạt động 3: Luyện tập\n(Nhìn Hình 13.14, 13.15, 13.16 đoán lệnh)', '07 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC3], [NLc], [PC2]', 'Học sinh nêu đúng từ khoá công cụ/thuộc tính/phép ghép cho 3 hình SGK tr. 74.\nCông cụ đánh giá: CCĐG 5 (Rubric quan sát vấn đáp).'],
        ['4. Hoạt động 4: Vận dụng\n(Chốt cách vẽ Miếng dưa hấu Hình 13.1)', '05 phút\n(PPCT 19)', '[KT1], [KT2], [KT3],\n[NLC3], [NLc], [PC1]', 'Học sinh nêu được 3 ý vẽ Miếng dưa hấu; xem GV demo xong tại lớp; chuẩn bị Tiết 2 TH.']
    ]
    eng.add(b, '', first=0)
    b.remove(b[-1])
    b.append(eng.table(ov_rows, [2700, 1150, 1950, 3271], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 1 ---
    eng.add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[NLC1]`, `[PC1]` – Ôn lại quy tắc thứ tự lớp của Bài 12 qua hình ảnh quen thuộc và tạo nhu cầu tìm hiểu cách cắt nửa hình tròn, tô màu và ghép hình.', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên chiếu **Hình 13.1 (Miếng dưa hấu – SGK trang 69)** và đặt hai câu hỏi trực quan dễ trả lời cho cả lớp quan sát:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    eng.add(b, '**c) Sản phẩm:** Câu trả lời miệng của học sinh nêu đúng màu sắc/hình khối và thứ tự 4 lớp từ dưới lên của miếng dưa hấu.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 1` *(Bảng kiểm vấn đáp ngắn – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 13.1 (SGK trang 69) lên màn hình và đặt câu hỏi ngắn cho cả lớp quan sát trong 01 phút:\n'
            '  + Câu hỏi 1: Quan sát miếng dưa hấu Hình 13.1, em thấy nó được ghép từ những hình khối và màu sắc nào?\n'
            '  + Câu hỏi 2: Nhắc lại quy tắc thứ tự lớp ở Bài 12 (vẽ trước ở dưới, vẽ sau ở trên): phần vỏ xanh, cùi trắng, ruột đỏ và hạt đen lớp nào nằm dưới cùng, lớp nào nằm trên cùng?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 69 hoặc nhìn màn chiếu, suy nghĩ nhanh tại chỗ (không cần thảo luận phức tạp).\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên gọi 01 học sinh trả lời Câu hỏi 1 và 01 học sinh trả lời Câu hỏi 2.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo `CCĐG 1`, khen ngợi và đặt vấn đề vào bài: *"Ở Bài 12 các em chỉ mới vẽ được hình tròn nguyên vẹn. Muốn cắt đôi ra nửa hình tròn làm miếng dưa hấu, tô màu bóng 3D hay ghép hình phức tạp, hôm nay chúng ta học Bài 13."*',
            '**Khởi động – Phân tích Hình 13.1 (SGK trang 69):**\n'
            '- **Các hình cấu thành:**\n'
            '  + 03 nửa hình tròn kích thước giảm dần (vỏ xanh đậm, cùi trắng/xanh nhạt, ruột đỏ).\n'
            '  + Các hình giọt nước/elip nhỏ màu đen (hạt dưa).\n'
            '- **Thứ tự 4 lớp (từ dưới lên trên):**\n'
            '  1. Vỏ xanh (lớp dưới cùng), tiếp đến 2. Cùi trắng, tiếp theo 3. Ruột đỏ và trên cùng là 4. Hạt đen (lớp trên cùng).'
        ]
    ]
    b.append(eng.table(hd1_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 2 ---
    eng.add(b, '2. Hoạt động 2: Hình thành kiến thức mới (28 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')

    # Nhiệm vụ 2.1
    eng.add(b, '2.1. Nhiệm vụ 1: Tìm hiểu các đối tượng hình khối và thuộc tính cơ bản (09 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[NLC1]`, `[NLa]`, `[PC1]` – Học sinh tra cứu được bảng trang 69–70 SGK để nêu đúng 3 thuộc tính trọng tâm (`Rx, Ry`; `Start, End`; `Rounded`) và hiểu qua thao tác minh hoạ của giáo viên.', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên chiếu **Hình 13.2** và **Bảng thuộc tính trang 69–70 SGK**, yêu cầu học sinh đọc bảng trong sách giáo khoa để trả lời nhanh 03 câu hỏi trọng tâm:', first=0, keep=True)
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    eng.add(b, '**c) Sản phẩm:** Học sinh tìm đúng từ khoá thuộc tính trong bảng SGK và ghi tóm tắt vào vở.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 2` *(Câu hỏi tra bảng SGK tr. 69–70 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 13.2 (tr. 69) hỏi cả lớp: *"Bốn hình màu đỏ này đều vẽ từ cùng một công cụ ở Bài 12, nhìn vào đỉnh nhọn của chúng em đoán xem là công cụ nào?"*\n'
            '- Tiếp theo, giáo viên yêu cầu học sinh đọc Bảng thuộc tính trang 69–70 SGK trong 02 phút theo cặp đôi để tìm đúng 3 thông số:\n'
            '  1. Để bo tròn 4 góc hình chữ nhật, chỉnh cặp thông số nào?\n'
            '  2. Để vẽ nửa hình tròn (như miếng dưa hấu), chỉnh góc `Start` và `End` bằng bao nhiêu độ?\n'
            '  3. (Câu hỏi SGK trang 70): Để các đỉnh ngôi sao 5 cánh nhọn trở nên cong tròn như bông hoa, thay đổi tham số nào?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 69–70, dùng bút gạch chân thẳng vào 3 từ khoá trong bảng.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời lần lượt 03 học sinh đọc to tên thuộc tính vừa tìm được trong SGK (kể cả học sinh trung bình/yếu chỉ cần nhìn sách cũng đọc đúng).\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét theo `CCĐG 2`, sau đó **mở phần mềm Inkscape thao tác trực tiếp trên máy chiếu (01 phút)**: gõ `Start = 0, End = 180` biến hình tròn thành nửa miếng dưa hấu và tăng `Rounded = 0.2` biến ngôi sao thành bông hoa để học sinh tận mắt thấy sự thay đổi.',
            '**1. Đối tượng hình khối và thuộc tính cơ bản (SGK tr. 69–70)**\n'
            '- **Hoạt động 1 (Hình 13.2):** Cả 4 hình đều vẽ bằng công cụ **Hình đa giác / Hình sao (phím *)** và chỉnh thuộc tính trên Thanh điều khiển.\n'
            '- **Thuộc tính trọng tâm cần nhớ:**\n'
            '  1. **Hình vuông / chữ nhật (R):** `W, H` (rộng, cao); `Rx, Ry` (độ bo tròn góc; bằng 0 nếu góc vuông).\n'
            '  2. **Hình tròn / elip (E):** `Rx, Ry` (hai bán kính); `Start, End` (góc đầu, góc cuối khi vẽ hình quạt/cung tròn; ví dụ `Start = 0, End = 180` tạo nửa hình tròn).\n'
            '  3. **Hình đa giác / sao (phím *):** `Corners` (số đỉnh/cánh); `Spoke ratio` (tỉ lệ bán kính trong/ngoài); `Rounded` (độ cong tại các đỉnh); `Randomized` (độ méo ngẫu nhiên).\n'
            '- **Đáp án câu hỏi tr. 70:** Thay đổi tham số **`Rounded`** để đỉnh ngôi sao cong tròn.'
        ]
    ]
    b.append(eng.table(nv1_rows, [4650, 4421], header=True, cell_sz=22))

    # Nhiệm vụ 2.2
    eng.add(b, '2.2. Nhiệm vụ 2: Thiết lập màu tô, màu vẽ và hộp thoại Fill and Stroke (09 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT2]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[NLS 3.2.NC1.a]`, `[PC2]`, `[PC3]` – Học sinh nhận biết sự khác nhau giữa các kiểu tô màu (đồng nhất, chuyển màu, trong suốt), chức năng 3 trang của `Fill and Stroke` và chọn đúng 02 câu trắc nghiệm SGK tr. 71.', first=0)
    eng.add(b, '**b) Nội dung:** Học sinh quan sát **Hình 13.3, Hình 13.4 (SGK trang 70)** và tham gia chọn đáp án nhanh cho 02 câu trắc nghiệm cuối mục ở trang 71:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    eng.add(b, '**c) Sản phẩm:** Học sinh phân biệt được `Fill` – `Stroke`, ý nghĩa `Opacity` và giơ tay chọn đúng đáp án `C` và `B` ở trang 71.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 3` *(02 câu trắc nghiệm nhanh trang 71 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 13.3 (5 hình vuông viền cam – tr. 70) và đặt câu hỏi: *"Bằng mắt thường, em hãy mô tả bằng lời xem phần màu bên trong của 5 ô vuông này khác nhau điểm nào?"*\n'
            '- Tiếp theo, giáo viên chiếu Hình 13.4 (Hộp thoại `Fill and Stroke`), giới thiệu ngắn gọn 3 thẻ `Fill`, `Stroke paint`, `Stroke style` và thanh trượt `Opacity`, rồi **bấm Shift + Ctrl + F trên Inkscape** thị phạm tô màu chuyển toả tròn (`Radial gradient`) và giảm `Opacity` cho cả lớp xem.\n'
            '- Cuối cùng, giáo viên chiếu **02 câu hỏi trắc nghiệm trang 71 SGK** và yêu cầu cả lớp giơ ngón tay (1 ngón = A, 2 = B, 3 = C, 4 = D) chọn đáp án.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh quan sát hình, đối chiếu SGK trang 70–71 và chọn phương án A/B/C/D.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- 01 học sinh mô tả sự khác nhau của 5 hình trong Hình 13.3; cả lớp đồng loạt giơ tay chọn đáp án 2 câu trắc nghiệm tr. 71; 01 học sinh giải thích ngắn gọn.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét theo `CCĐG 3`, chốt kiến thức chuẩn lên bảng.',
            '**2. Màu tô, màu vẽ và hộp thoại Fill and Stroke (SGK tr. 70–71)**\n'
            '- **Cách mở:** Menu **`Object / Fill and Stroke`** (phím tắt `Shift + Ctrl + F`).\n'
            '- **Ba trang (Tab) chính:**\n'
            '  + **Fill (Màu tô):** Tô nền bên trong đối tượng (không màu `X`, đồng nhất `Flat color`, chuyển tuyến tính `Linear gradient`, chuyển từ tâm `Radial gradient`, hoa văn `Pattern`).\n'
            '  + **Stroke paint (Màu vẽ):** Đặt màu cho đường viền.\n'
            '  + **Stroke style (Kiểu nét vẽ):** Chỉnh độ dày (`Width`) và kiểu đường nét (nét liền, nét đứt...).\n'
            '- **Opacity (%) và Alpha (A):** Điều chỉnh độ trong suốt (nhìn xuyên qua các lớp).\n'
            '- **Đáp án trắc nghiệm trang 71:**\n'
            '  + Câu 1 (đường viền nét đứt): **`C. Stroke style`**.\n'
            '  + Câu 2 (chỉnh điểm dừng màu gradient): **`B`** (cây bút chỉnh gradient).'
        ]
    ]
    b.append(eng.table(nv2_rows, [4650, 4421], header=True, cell_sz=22))

    # Nhiệm vụ 2.3
    eng.add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các phép ghép đối tượng đồ hoạ trong bảng chọn Path (10 phút)', 'sub')
    eng.add(b, '**a) Mục tiêu:** `[KT3]`, `[NLC1]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh nhận biết được quy luật của 6 phép ghép trong bảng chọn `Path` (đặc biệt là `Union`, `Difference`, `Intersection`), trả lời được cách ghép hình trái tim (tr. 71) và đám mây (tr. 72).', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên chiếu **Hình 13.5**, **Hình 13.6** và **Hình 13.7 (SGK trang 71–72)** để học sinh quan sát hình kết quả và ghép đôi tên gọi tương ứng:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    eng.add(b, '**c) Sản phẩm:** Học sinh phân biệt đúng 6 phép ghép ở Hình 13.6, nêu được cách ghép đám mây Hình 13.7 và quy tắc "hình dưới bị hình trên cắt".', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 4` *(Bài tập nhận diện hình kết quả Hình 13.6 & Câu hỏi đám mây tr. 72 – Phụ lục 2)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 13.5 (2 hình tròn và 1 hình vuông ghép thành trái tim) và Hình 13.6 (SGK tr. 72) với 1 hình vuông (lớp dưới) + 1 hình tròn (lớp trên), giao bài tập cực nhanh theo cặp bàn (02 phút):\n'
            '  + Câu hỏi 1: Quan sát Hình 13.6 trong SGK trang 72, hãy chỉ ra phép nào gộp hết hai hình (`Union`), phép nào lấy hình vuông bị hình tròn khoét mất một miếng (`Difference`), và phép nào chỉ lấy phần chung ở giữa (`Intersection`)?\n'
            '  + Câu hỏi 2 (SGK tr. 72): Để vẽ đám mây Hình 13.7 từ nhiều hình tròn chồng lên nhau, ta chọn phép ghép nào trong 6 phép trên?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở SGK trang 72, nhìn trực tiếp vào Hình 13.6 và đối chiếu hình vẽ với tên lệnh bên dưới hình.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh đứng tại chỗ đọc đáp án; sau đó giáo viên giải thích nốt 3 phép còn lại (`Exclusion`, `Division`, `Cut Path`) bằng trực quan.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo `CCĐG 4`, **demo trực tiếp trên Inkscape (01 phút)**: gộp 4 hình tròn bằng `Path / Union` thành đám mây Hình 13.7 và dùng `Path / Difference` cắt hình vuông bằng hình tròn (liên hệ ngay Logo quả táo cắn dở Apple và máy cắt Laser/AI Vector `.svg` để mở rộng hứng thú). Nhấn mạnh quy tắc: Trong `Difference`, `Division`, `Cut Path`, **hình lớp dưới luôn bị hình lớp trên cắt**.',
            '**3. Các phép ghép đối tượng trong bảng chọn Path (SGK tr. 71–72)**\n'
            '- **Cách thực hiện:** Chọn từ 2 đối tượng trở lên, mở bảng chọn **`Path`** rồi chọn phép ghép:\n'
            '  1. **Phép hợp (Union, Ctrl + +):** Gộp tất cả thành một hình duy nhất (dùng ghép hình trái tim H.13.5 và **đám mây H.13.7**).\n'
            '  2. **Phép hiệu (Difference, Ctrl + -):** Lấy **hình lớp dưới** trừ phần giao với **hình lớp trên**.\n'
            '  3. **Phép giao (Intersection, Ctrl + *):** Chỉ lấy phần giao nhau của các hình.\n'
            '  4. **Phép hiệu đối xứng (Exclusion, Ctrl + ^):** Lấy tất cả trừ phần giao nhau.\n'
            '  5. **Phép chia (Division, Ctrl + /):** Cắt hình lớp dưới thành nhiều mảnh rời theo nét hình lớp trên.\n'
            '  6. **Phép cắt đường (Cut Path, Ctrl + Alt + /):** Cắt đường viền hình lớp dưới tại giao điểm (mất màu tô nền).\n'
            '- **Quy tắc cần nhớ:** Với `Difference`, `Division`, `Cut Path`, **hình lớp dưới làm gốc bị cắt bởi hình lớp trên**.'
        ]
    ]
    b.append(eng.table(nv3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 3 ---
    eng.add(b, '3. Hoạt động 3: Luyện tập (07 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[KT2]`, `[KT3]`, `[NLC3]`, `[NLc]`, `[PC2]` – Học sinh vận dụng được kiến thức vừa học ở mức độ nhận diện công cụ, thuộc tính và phép ghép cho 03 hình Luyện tập SGK trang 74 (vừa sức, không yêu cầu viết văn bản dài).', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên tổ chức hoạt động nhanh **"Nhìn hình đoán lệnh"** qua **Hình 13.14, Hình 13.15, Hình 13.16 (SGK trang 74)** trên màn chiếu:', first=0, keep=True)
    eng.add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    eng.add(b, '**c) Sản phẩm:** Học sinh nêu đúng 1–2 từ khoá công cụ/thuộc tính/phép ghép cho từng hình (không yêu cầu hoàn chỉnh quy trình phức tạp).', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 5` *(Rubric 6 cột quan sát vấn đáp Luyện tập & Vận dụng – Phụ lục 3)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu 3 hình Luyện tập ở SGK trang 74 (Hình 13.14 Bông hoa, Hình 13.15 Quân Bích đỏ, Hình 13.16 Chùm bóng bay) và đặt 03 câu hỏi ngắn, mỗi câu chỉ cần trả lời 1 cụm từ khoá:\n'
            '  1. *Hình 13.14 (Bông hoa):* Bắt đầu từ công cụ Hình sao (`*`), ta tăng thuộc tính nào để đầu cánh sao uốn cong tròn? Vẽ thêm hình gì làm nhụy ở giữa?\n'
            '  2. *Hình 13.15 (Quân Bích đỏ):* Phần tán trên ghép giống hình gì (trái tim lật ngược bằng `Path / Union`), còn phần đế dưới bị lõm hai bên ta khoét bằng phép gì (`Path / Difference`)?\n'
            '  3. *Hình 13.16 (Chùm bóng bay nhìn xuyên qua nhau):* Chỉnh thông số nào trong hộp thoại `Fill and Stroke` để quả bóng trở nên trong suốt một phần?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh trao đổi với bạn cùng bàn 02 phút, nhìn vào 3 hình trong SGK tr. 74 để chọn từ khoá.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên gọi 03 học sinh xung phong (hoặc gọi ngẫu nhiên trả lời từng hình một); lớp nhận xét.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên đánh giá theo Rubric `CCĐG 5` (Phụ lục 3) và biểu diễn nhanh thao tác tạo bông hoa Hình 13.14 trên Inkscape (30 giây) để chốt đáp án.',
            '**Giải nhanh 3 hình Luyện tập (SGK tr. 74):**\n'
            '- **Hình 13.14 (Bông hoa):** Công cụ Hình sao (`Corners = 10`), tăng **`Rounded`** để bo tròn cánh + thêm **1 hình tròn vàng** ở lớp trên làm nhụy.\n'
            '- **Hình 13.15:** Tán trên dùng **`Path / Union`** (ghép 1 hình vuông xoay 45 độ + 2 hình tròn); chân đế dùng **`Path / Difference`** khoét 2 bên hình chữ nhật; cuối cùng dùng **`Path / Union`** gộp lại.\n'
            '- **Hình 13.16 (Chùm bóng):** Vẽ 3 hình elip, giảm độ trong suốt **`Opacity`** khoảng `80%` để nhìn xuyên nhau + vẽ dây bằng Bút vẽ.'
        ]
    ]
    b.append(eng.table(hd3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- HOẠT ĐỘNG 4 ---
    eng.add(b, '4. Hoạt động 4: Vận dụng (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    eng.add(b, '**a) Mục tiêu:** `[KT1]`, `[KT2]`, `[KT3]`, `[NLC3]`, `[NLc]`, `[PC1]` – Học sinh tự ghép nối lại toàn bộ kiến thức bài học để trả lời câu hỏi Vận dụng 1 (vẽ Miếng dưa hấu Hình 13.1 đầu giờ) và nắm nhiệm vụ cho Tiết 2 Thực hành.', first=0)
    eng.add(b, '**b) Nội dung:** Giáo viên quay lại **Hình 13.1 (Miếng dưa hấu)** ở Hoạt động Mở đầu và đặt câu hỏi tổng kết bài học (Vận dụng 1 SGK trang 74).', first=0)
    eng.add(b, '**c) Sản phẩm:** Học sinh nêu miệng được 3 ý chính để vẽ miếng dưa hấu bằng `Start = 0, End = 180` và xem giáo viên thực hiện trực tiếp trên `Inkscape`.', first=0)
    eng.add(b, '**Công cụ đánh giá:** `CCĐG 5` *(Rubric quan sát vấn đáp – Phụ lục 3)*.', first=0, keep=True)
    eng.add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd4_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu lại Hình 13.1 (Miếng dưa hấu) và hỏi: *"Bây giờ đã biết thuộc tính hình tròn và thứ tự lớp rồi, đố cả lớp nêu nhanh trong 3 câu ngắn làm sao vẽ được miếng dưa hấu này?"*\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh suy nghĩ 01 phút, liên hệ lại mục `Start, End` vừa học.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 01–02 học sinh phát biểu ngắn gọn.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, thao tác mẫu nhanh Miếng dưa hấu ngay trên Inkscape (30 giây) để khép lại mạch bài học từ Khởi động đến Vận dụng, rồi dặn dò học sinh đọc trước 3 bài thực hành trang 72–74 (Logo, Hình cây, Hình cầu) chuẩn bị cho **Tiết 2 – Thực hành (Tiết PPCT: 20)**.',
            '**Quy trình 3 ý vẽ Miếng dưa hấu Hình 13.1 (Vận dụng 1 tr. 74):**\n'
            '1. Dùng Hình tròn (`E`), đặt **`Start = 0, End = 180`** (chế độ hình quạt `Slice`) tạo nửa hình tròn màu xanh đậm làm vỏ dưa.\n'
            '2. Nhấn `Ctrl + D` nhân bản, thu nhỏ tô trắng làm cùi; tiếp tục `Ctrl + D`, thu nhỏ tô đỏ làm ruột.\n'
            '3. Vẽ elip nhỏ màu đen, `Ctrl + D` xếp rải lên ruột đỏ làm hạt dưa.'
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
         '- **Mã năng lực số `[NLS 3.2.NC1.a]` (Tích hợp và tạo lập lại nội dung số):** Nhận biết và hiểu được cách tinh chỉnh thuộc tính, màu sắc, độ trong suốt và kết hợp các đối tượng đồ hoạ số (vectơ) để hình thành sản phẩm mới.'],
        ['**3. Kế hoạch dạy học môn Tin học Khối 10**\n*(Phân phối chương trình năm học 2026–2027 của Tổ Tin học – Trường THPT Thuận Hoá)*',
         '- **Tiết PPCT 19 (Học kì I):** `Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT)` – Thời lượng: 01 tiết (45 phút). Nội dung trọng tâm theo SGK Tin học 10 (KNTT, trang 69–74): (1) Các đối tượng hình khối và thuộc tính cơ bản; (2) Thiết lập màu tô, màu vẽ và hộp thoại `Fill and Stroke`; (3) Các phép ghép đối tượng trong bảng chọn `Path`; Luyện tập & Vận dụng lí thuyết.\n'
         '- **Tiết PPCT 20 (Học kì I):** `Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 2 - TH)` – Thời lượng: 01 tiết (45 phút). Học sinh thực hành tại phòng máy 03 bài trong SGK trang 72–74 (Logo 3 cánh, Hình cây xanh, Hình cầu 3D) và hoàn thiện sản phẩm `.svg`.']
    ]
    b.append(eng.table(can_cu_rows, [3150, 5921], header=True, cell_sz=21))
    eng.add(b, '', first=0)

    eng.add(b, 'Phụ lục 2. Hệ thống câu hỏi vấn đáp tra bảng SGK và Trắc nghiệm chiếu màn hình (CCĐG 1 – CCĐG 4)', 'sub')
    pl1_rows = [
        ['Công cụ đánh giá', 'Câu hỏi / Yêu cầu thực hiện trực tiếp tại lớp (Không phát giấy)', 'Đáp án chuẩn & Tiêu chí đạt'],
        ['**CCĐG 1**\n(Hoạt động 1:\nMở đầu)',
         '- Quan sát Hình 13.1 (tr. 69): Kể tên các hình/màu và lớp nằm dưới cùng, trên cùng.',
         '- Nêu đúng 3 nửa hình tròn (vỏ xanh dưới cùng, đến cùi trắng, tiếp đến ruột đỏ) và hạt đen ở lớp trên cùng *(Đạt / Chưa đạt)*.'],
        ['**CCĐG 2**\n(Nhiệm vụ 2.1:\nThuộc tính)',
         '- Tra bảng SGK tr. 69–70: (1) Bo góc hình chữ nhật dùng thông số nào? (2) Vẽ nửa hình tròn chỉnh góc nào? (3) Câu hỏi tr. 70: Làm cong đỉnh ngôi sao chỉnh tham số nào?',
         '- Trả lời đúng 3 từ khoá: (1) `Rx, Ry`; (2) `Start = 0, End = 180`; (3) `Rounded` *(Đạt từ 2/3 ý trở lên)*.'],
        ['**CCĐG 3**\n(Nhiệm vụ 2.2:\nFill & Stroke)',
         '- Chiếu 02 câu trắc nghiệm SGK tr. 71 cho cả lớp giơ tay chọn A/B/C/D:\n  + Câu 1: Để chọn nét đứt cho đường viền chọn trang nào?\n  + Câu 2: Chỉnh điểm dừng màu gradient chọn biểu tượng nào?',
         '- Câu 1: **`C. Stroke style`**.\n- Câu 2: **`B`** (cây bút chỉnh gradient).\n*(Đánh giá tỉ lệ cả lớp giơ tay đúng)*.'],
        ['**CCĐG 4**\n(Nhiệm vụ 2.3:\nPhép ghép Path)',
         '- Quan sát Hình 13.6 & Hình 13.7 (tr. 72): Chỉ ra lệnh `Union`, `Difference`, `Intersection` và nêu phép ghép dùng vẽ đám mây Hình 13.7.',
         '- Nêu đúng ý nghĩa gộp (`Union`), cắt lớp dưới (`Difference`), lấy phần chung (`Intersection`) và đáp án đám mây là **`Path / Union`** *(Đạt)*.']
    ]
    b.append(eng.table(pl1_rows, [1600, 4200, 3271], header=True, cell_sz=21))

    eng.add(b, 'Phụ lục 3. Công cụ đánh giá CCĐG 5 – Rubric quan sát trả lời nhanh Hoạt động Luyện tập & Vận dụng', 'sub')
    rubric_info = [
        ['THÔNG TIN HÀNH CHÍNH CÔNG CỤ ĐÁNH GIÁ (CCĐG 5)', 'NỘI DUNG CHI TIẾT'],
        ['Tên bài dạy & Tiết PPCT', 'Bài 13. Bổ sung các đối tượng đồ hoạ (Tiết 1 - Lí thuyết / Tiết PPCT: 19)'],
        ['Lớp giảng dạy & Thời gian', 'Lớp 10/3 – Trường THPT Thuận Hoá (Thứ Hai, ngày 16/11/2026)'],
        ['Giáo viên hướng dẫn & SVTT', 'GVHD: Thầy Mai Xuân Văn – Sinh viên thực hiện: Hồ Quang Huy (24S1020030)'],
        ['Hình thức đánh giá thực tế', 'Giáo viên gọi tên, quan sát câu trả lời miệng trực tiếp tại lớp qua màn chiếu & SGK']
    ]
    b.append(eng.table(rubric_info, [3700, 5371], header=True, cell_sz=21))
    eng.add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Nội dung câu hỏi SGK tr. 74', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Tỉ trọng', 'Mức đạt'],
        ['1', 'Nhận diện thuộc tính hình khối (Hình 13.14)', 'Chọn công cụ vẽ cánh hoa và tham số làm cong đỉnh ngôi sao.', '- **Tốt:** Nêu đúng Hình sao (`*`) và tham số `Rounded` + Hình tròn làm nhụy.\n- **Đạt:** Nêu được Hình sao nhưng cần gợi ý mới nhớ tên `Rounded`.\n- **Chưa đạt:** Nhầm sang công cụ khác.', '30%', '......'],
        ['2', 'Nhận diện phép ghép Path (Hình 13.15)', 'Chọn phép ghép phần trên và phần đế của quân Bích đỏ.', '- **Tốt:** Nêu đúng dùng `Union` cho tán và `Difference` cho đế.\n- **Đạt:** Chỉ ra được có gộp hình (`Union`) nhưng chưa phân biệt rõ `Difference`.\n- **Chưa đạt:** Không xác định được phép ghép.', '30%', '......'],
        ['3', 'Độ trong suốt & Vẽ dưa hấu (H.13.16 & H.13.1)', 'Chọn `Opacity` cho bóng bay và góc `Start=0, End=180` vẽ dưa hấu.', '- **Tốt:** Nêu đúng từ khoá `Opacity` và góc `Start=0, End=180` + `Ctrl + D`.\n- **Đạt:** Mô tả được bằng lời ("làm mờ đi", "cắt nửa hình tròn").\n- **Chưa đạt:** Chưa liên hệ được cách làm.', '40%', '......']
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
        'hinh_13_14_15_16_luyen_tap.png',
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
        page1_sched_line='Bối cảnh phòng H.I.5: Có 01 máy tính GV + máy chiếu + bảng phấn (HS không có máy tính, dễ ngại nói)',
        page1_mode_line='Chiến thuật dạy: Giải thích bình dân – Hỏi câu nhìn SGK đọc được – Bấm sẵn mau_bai13.svg – Không đơ máy!'
    )

    # MỤC I: CHECKLIST 5 PHÚT TRƯỚC GIỜ VÀ SƠ ĐỒ BẢNG PHẤN
    eng.add(b, 'I. CHECKLIST 5 PHÚT TRƯỚC GIỜ & SƠ ĐỒ BẢNG PHẤN (LÀM NGAY KHI CẮM MÁY CHIẾU)', 'heading')
    eng.add(b, '1. Bốn việc bắt buộc làm trước khi vào tiết (Để lên bục bấm phát ăn ngay, không bị quê)', 'sub')
    for item in [
        '**Bước 1 – Mở sẵn tệp mau_bai13.svg bằng Inkscape:** Mở ngay trước giờ chuông reo để Inkscape tải xong trước (tránh mất 10–15 giây chờ mở phần mềm trước mặt Thầy Văn và cả lớp). Tệp này đã chia sẵn **5 khung thị phạm** khớp đúng từng phút giảng trên lớp (lớp khung nền đã được khoá an toàn để bạn quét chuột không bị xê dịch khung).',
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
            '1. Chỉnh góc/cánh (Thuộc tính)\n'
            '2. Tô màu 3D, mờ (Fill & Stroke)\n'
            '3. Ghép/Cắt hình (Path)',
            '**1. Đối tượng hình khối & thuộc tính (tr. 69-70)**\n'
            '- Chữ nhật: `W, H`, `Rx, Ry` (bo góc)\n'
            '- Tròn: `Start, End` (từ `0 đến 180`: nửa tròn)\n'
            '- Sao/Đa giác: `Corners` (cánh), `Rounded` (cong)\n\n'
            '**2. Màu tô, màu vẽ – Fill and Stroke (Shift+Ctrl+F)**\n'
            '- `Fill`: Màu nền (`Flat`, `Linear/Radial gradient`)\n'
            '- `Stroke paint` / `Stroke style`: Màu viền / Kiểu nét\n'
            '- `Opacity`: Độ trong suốt (%)\n\n'
            '**3. Các phép ghép hình trong bảng chọn Path**\n'
            '- `Union` (`Ctrl + +`): Hợp (gộp lại)\n'
            '- `Difference` (`Ctrl + -`): Hiệu (Dưới bị Trên cắt)\n'
            '- `Intersection` (`Ctrl + *`): Giao (phần chung)\n'
            '- `Exclusion`, `Division`, `Cut Path` (SGK tr. 72)',
            '**Mẹo nhớ Bài 12:**\n'
            '- Vẽ trước thì nằm DƯỚI\n'
            '- Vẽ sau thì nằm TRÊN\n\n'
            '**Nháp đáp án tr. 71:**\n'
            'Câu 1: C (`Stroke style`)\n'
            'Câu 2: B (cây bút)\n\n'
            '**Quy tắc cắt Path:**\n'
            'HÌNH DƯỚI bị HÌNH TRÊN cắt!'
        ]
    ]
    b.append(eng.table(board_rows, [2150, 4671, 2250], header=True, cell_sz=21))

    # MỤC II: GIẢI NGỐ KIẾN THỨC CỐT LÕI TỪ CON SỐ 0 CHO NGƯỜI TRÁI NGÀNH
    eng.add(b, 'II. TỰ HỌC & GIẢI NGỐ KIẾN THỨC CỐT LÕI BÀI 13 TỪ CON SỐ 0 (ĐỌC 5 PHÚT LÀ NẮM CHẮC)', 'heading')
    eng.add(b, 'Phần này dịch toàn bộ thuật ngữ tiếng Anh khó nhớ trong SGK Tin học 10 (tr. 69–74) sang tiếng Việt bình dân, giúp bạn tự tin giải thích trôi chảy dù học sinh hay Thầy Mai Xuân Văn có hỏi vặn bất kỳ chi tiết nào:', first=0)

    eng.add(b, '1. Thuộc tính hình khối (Mục 1 SGK tr. 69–70): Khác nhau giữa Hình vuông, Hình tròn, Ngôi sao', 'sub')
    for item in [
        '**Mẹo nhớ vị trí:** Nháy đúp chuột vào bất kỳ hình nào (hoặc chọn công cụ tương ứng ở thanh bên trái) thì thanh ngang trên cùng (**Thanh điều khiển thuộc tính** của Bài 12) sẽ lập tức hiện ra các ô thông số của hình đó.',
        '**Hình vuông / Hình chữ nhật (phím R):** `W, H` là chiều rộng, chiều cao; **`Rx, Ry`** là **độ bo tròn 4 góc** (`Rx = Ry = 0` thì góc nhọn vuông vức như viên gạch thẻ; tăng `Rx, Ry` lên thì góc bo tròn mềm mại như viền điện thoại iPhone).',
        '**Hình tròn / Hình elip (phím E):** `Rx, Ry` là bán kính ngang/dọc. Có **Start và End** (góc điểm đầu và điểm cuối tính theo độ, ví dụ `0` và `180` là cắt đúng một nửa vòng tròn) cùng **3 nút biểu tượng hình tròn** ở cuối thanh thuộc tính: **Hình quạt (Slice)** nối 2 đầu vào tâm như miếng bánh/dưa hấu; **Cung tròn hở (Arc)** chỉ giữ nét viền cong không nối tâm như chữ C; **Hình tròn khép kín (Whole)** đóng tròn nguyên vẹn.',
        '**Hình đa giác / Hình sao (phím *):** Có 2 nút chuyển chế độ hình đa giác (cạnh đều) hoặc ngôi sao nhọn; **`Corners`** là số đỉnh/số cánh; **`Spoke ratio`** là tỉ lệ bán kính trong/ngoài quyết định cánh sao mập hay gầy nhọn; **`Rounded`** là độ uốn cong tại các đỉnh ("cho ngôi sao đi spa mài góc thành bông hoa"); **`Randomized`** là độ méo ngẫu nhiên (bằng `0` thì đối xứng đều đặn, tăng lên thì hình méo mó tự do).'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Hộp thoại Fill and Stroke (Mục 2 SGK tr. 70–71): Phân biệt 3 trang & 5 kiểu tô màu', 'sub')
    for item in [
        '**Cách gọi hộp thoại:** Menu **`Object / Fill and Stroke`** hoặc bấm **`Shift + Ctrl + F`**. Nhớ phân biệt: **`Fill`** là màu sơn **bên trong nền**, còn **`Stroke`** là màu chỉ may **đường viền bên ngoài**.',
        '**Ba trang (Tab) trong hộp thoại:** Trang 1 **`Fill`** (tô nền); Trang 2 **`Stroke paint`** (chọn màu cho đường viền); Trang 3 **`Stroke style`** (chỉnh độ dày mỏng `Width` và kiểu nét đứt/nét liền ở mục `Dashes` — đây là đáp án Câu 1 tr. 71).',
        '**Các biểu tượng kiểu tô màu trong Hình 13.4:** Dấu **`X`** (không tô/trong suốt); **Ô vuông đặc** (`Flat color` – đồng nhất 1 màu); **Hai ô vuông chuyển** (`Linear gradient` – chuyển màu thẳng, ví dụ từ sáng sang tối); **Ô vuông sáng giữa tâm** (`Radial gradient` – chuyển màu toả tròn giúp hình tròn phẳng biến thành quả cầu 3D); **Ô bàn cờ** (`Pattern` – tô hoa văn); Dấu **`?`** (huỷ đặt).',
        '**Độ trong suốt (Opacity và kênh A – Alpha):** Kéo giảm từ `100%` xuống `50% – 80%` thì đối tượng giống như **tấm kính râm**, nhìn xuyên thấu được các hình nằm lớp bên dưới (ứng dụng vẽ chùm bóng bay Hình 13.16).'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '3. Trọn bộ 06 phép ghép đối tượng trong bảng chọn Path (Mục 3 SGK tr. 71–72)', 'sub')
    eng.add(b, 'Trong Hình 13.6 SGK trang 72 luôn giả định có **Hình vuông nằm lớp DƯỚI** và **Hình tròn nằm lớp TRÊN**. Khi chọn từ 2 hình trở lên rồi mở menu **`Path`**, hãy nhớ ý nghĩa bình dân của cả 6 lệnh sau (trong đó 3 lệnh đầu là trọng tâm nhất):', first=0, keep=True)
    path_rows = [
        ['STT & Tên lệnh trong Path', 'Phím tắt', 'Ví von bình dân & Kết quả (Hình vuông DƯỚI, Hình tròn TRÊN)', 'Ứng dụng & Câu hỏi SGK'],
        ['1. **Union** *(Phép hợp)*', '`Ctrl + +`', '**"Hàn xì hai đứa thành 1 cục"**: Gộp chung cả hình vuông và hình tròn thành một khối duy nhất (giống tập hợp A ∪ B).', 'Ghép Trái tim (H.13.5), Đám mây (H.13.7), Tán cây (H.13.11).'],
        ['2. **Difference** *(Phép hiệu)*', '`Ctrl + -`', '**"Khuôn bánh cắn đứt 1 miếng"**: Hình **DƯỚI** (vuông) bị hình **TRÊN** (tròn) khoét mất phần đè nhau, hình tròn biến mất (giống A \\ B).', 'Khoét Logo Apple "cắn dở", gọt thân cây thắt eo (H.13.11), chân Quân Bích (H.13.15).'],
        ['3. **Intersection** *(Phép giao)*', '`Ctrl + *`', '**"Chỉ chừa khúc đè lên nhau"**: Chỉ giữ lại phần giao nhau ở giữa hai hình (giống tập hợp A ∩ B).', 'Vẽ chiếc lá trên đầu táo Apple từ 2 hình tròn giao nhau.'],
        ['4. **Exclusion** *(Hiệu đối xứng)*', '`Ctrl + ^`', '**"Đục lỗ xuyên thấu ở giữa"**: Giữ hai phần rìa ngoài nhưng khoét thủng phần giao nhau ở giữa.', 'Tạo chữ lồng, khung viền rỗng giữa khi hai hình đè nhau.'],
        ['5. **Division** *(Phép chia)*', '`Ctrl + /`', '**"Dao thái bánh"**: Lấy mép hình **TRÊN** cắt hình **DƯỚI** thành các mảnh rời nhau nhưng **vẫn giữ màu tô**, kéo tách từng mảnh ra được!', 'Tách 1 hình tròn thành nhiều miếng màu ghép khác nhau.'],
        ['6. **Cut Path** *(Phép cắt đường)*', '`Ctrl + Alt + /`', '**"Kéo cắt sợi dây"**: Chỉ cắt đứt đường viền hình **DƯỚI** tại điểm giao với hình **TRÊN**, **mất hết màu tô nền bên trong**.', 'Cắt lấy một đoạn cung viền hở để chỉnh đường nét.']
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
    eng.add(b, 'III. KHO LIÊN HỆ BÀI CŨ SGK & CHUYỆN NGOÀI ĐỜI CỰC CUỐN (CHÈN KHI GIẢNG)', 'heading')
    rel_rows = [
        ['Chặng giảng', 'Liên hệ Bài cũ trong SGK Tin 10 (Nói 20 giây cho Thầy Văn chấm điểm)', 'Ví von vui nhộn & Liên hệ AI / Thực tế bên ngoài'],
        ['**Chặng 1:**\nMở đầu',
         '- **Bài 11 (Bản quyền phần mềm):** Hỏi vui sao trường không dạy Photoshop/Illustrator đắt tiền mà dạy **Inkscape**? Vì Inkscape miễn phí nguồn mở, tôn trọng bản quyền!\n- **Bài 12 (Thứ tự lớp):** Hình vẽ trước nằm dưới, vẽ sau nằm trên.',
         '- Ví von thứ tự lớp bằng **ly trà sữa trân châu**: Trân châu bỏ trước nằm đáy (lớp 1), đá và kem cheese bỏ sau đè lên trên! Muốn thấy hạt dưa trên miếng dưa hấu thì hạt phải rắc lên lớp trên cùng!'],
        ['**Chặng 2A:**\nThuộc tính',
         '- **Bài 12 (Màn hình Inkscape H.12.3):** Hỏi lại tên cái thanh ngang ngay dưới thanh bảng chọn: chính là **Thanh điều khiển thuộc tính** — chọn công cụ nào là thanh đó biến hình theo thông số của công cụ đó.',
         '- Ví von `Rx, Ry` như **điện thoại iPhone (bo góc mềm)** so với **viên gạch thẻ (góc vuông 90 độ)**.\n- Ví von `Rounded` như **cho ngôi sao 5 cánh nhọn đi spa mài góc** để biến thành bông hoa 5 cánh tròn!'],
        ['**Chặng 2B:**\nFill & Stroke',
         '- **Bài 12 (Dải màu dưới đáy):** Nháy chuột trái = đổi màu tô (`Fill`); giữ `Shift` + nháy chuột = đổi màu viền (`Stroke`). Nay mở `Shift + Ctrl + F` để chỉnh màu chuyển 3D và nét đứt.',
         '- `Fill` là **cái áo**, `Stroke` là **đường chỉ viền cổ áo**.\n- `Linear gradient` như **ly trà đào cam sả** chuyển từ vàng xuống cam.\n- `Radial gradient` như **đèn rọi vào quả bi-a 3D**.\n- `Opacity` giống như **đeo kính râm** (cũng dùng chữ này trong **Canva, Figma**).'],
        ['**Chặng 2C:**\nPhép ghép Path & AI',
         '- **Toán lớp 10 (Tập hợp)** & **Bài 12 (lớp trên/dưới)**: `Union` = `A ∪ B`, `Intersection` = `A ∩ B`, `Difference` = `A \\ B`.\n- Mở đường sang **Bài 14**: Bấm lệnh `Path` biến hình khối thành đối tượng đường.',
         '- **Bí mật Logo Apple:** Quả táo nguyên + hình tròn bên phải + `Path / Difference` = "cắn 1 miếng" chuẩn xác! Chiếc lá trên đỉnh = 2 hình tròn + `Intersection`!\n- **Cắt Laser / Mica:** Máy CNC đọc trực tiếp đường `Path` trong file `.svg`.\n- **AI ảnh vs AI Vector:** Midjourney vẽ `.png` phóng to sẽ vỡ hạt, còn hỏi **ChatGPT/Gemini viết code SVG** mở vào Inkscape vẫn sửa từng lớp vô tư!']
    ]
    b.append(eng.table(rel_rows, [1450, 3200, 4421], header=True, cell_sz=21))

    # MỤC IV: KỊCH BẢN TỪNG CHẶNG (2 CỘT RỘNG RÃI + HÌNH SGK NGAY TRÊN MỖI CHẶNG)
    eng.add(b, 'IV. KỊCH BẢN LÊN LỚP TỪNG CHẶNG: NÓI GÌ – BẤM GÌ TRÊN `mau_bai13.svg` – CHỮA CHÁY SAO (45 PHÚT)', 'heading')
    eng.add(b, 'Mẹo tâm lý: Không cần gồng mình tỏ ra hiểu hết mọi ngóc ngách đồ hoạ! Bạn cứ nhập vai một người thầy gần gũi, hài hước, bám đúng quy luật "hỏi câu mở SGK ra là đọc được đáp án – bấm chuột trực tiếp trên tệp mau_bai13.svg cho cả lớp xem tận mắt", bảo đảm cả lớp lẫn Thầy Mai Xuân Văn đều hài lòng.', first=0)

    # --- Chặng 1 ---
    eng.add(b, 'Chặng 1 (Phút 00–05): Mở đầu – Phân tích Miếng dưa hấu (Hình 13.1 SGK tr. 69)', 'sub')
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    c1_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Chào cả lớp! Hôm nay trời mát/nóng quá, thầy mời cả lớp ăn... dưa hấu bằng mắt trước nha!"*\n'
            '- *"Thử thách đầu giờ siêu dễ: Nhìn miếng dưa hấu Hình 13.1 trang 69, kể giúp thầy từ ngoài vào trong có mấy lớp màu và gồm những hình gì?"*\n'
            '- *(HS trả lời xong)*: *"Chuẩn rồi! Giờ nhớ lại Bài 12: Hình vẽ trước nằm dưới, vẽ sau nằm trên, giống như ly trà sữa trân châu bỏ vào trước nằm đáy. Vậy vỏ xanh nằm đáy (lớp 1) hay hạt đen nằm đáy?"*\n'
            '- *(Chốt vào bài)*: *"Bài 12 mình mới vẽ được cả cái bánh pizza hình tròn chứ chưa biết xẻ nửa ra làm miếng dưa hấu. Hôm nay Bài 13 sẽ dạy cách chỉnh góc, tô màu 3D và cắt ghép hình!"*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 1 bên trái):** Kéo thử từng lớp của miếng dưa hấu mẫu (vỏ xanh, cùi trắng, ruột đỏ, hạt đen) ra cho cả lớp cười ồ thấy rõ lớp dưới/trên, xong nhấn `Ctrl + Z` trả về chỗ cũ. Ghi tựa Bài 13 lên bảng.\n'
            '- **Chữa cháy nếu lớp im re:** Cười tươi hỏi trắc nghiệm bằng miệng luôn: *"Đố nha: lớp hạt đen nằm dưới lớp ruột đỏ hay trên lớp ruột đỏ? Nếu hạt nằm dưới thì mình có thấy cái hạt nào không?"* Khi đó cả lớp tự khắc đồng thanh đáp "Trên!".'
        ]
    ]
    b.append(eng.table(c1_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2A ---
    eng.add(b, 'Chặng 2A (Phút 05–14): Nhiệm vụ 1 – Các đối tượng hình khối & thuộc tính (SGK tr. 69–70)', 'sub')
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    c2a_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Nhìn 4 hình màu đỏ ở Hình 13.2 trang 69: Nhìn khác hẳn nhau vậy nhưng sinh ra từ CÙNG MỘT MẸ là một công cụ ở Bài 12! Đố các bạn là công cụ nào trong 3 hình: Chữ nhật, Hình tròn hay Hình sao?"*, sau đó gọi 1 HS trả lời (Hình sao!).\n'
            '- *"Chính xác! Giờ cho cả lớp 2 phút rà mắt vào Bảng thuộc tính trang 69–70 SGK, không cần học thuộc lòng, chỉ cần nhìn sách đọc to cho thầy đúng 3 mật mã sau:*\n'
            '  *1. Điện thoại iPhone bo tròn góc, còn gạch thẻ vuông góc: muốn bo 4 góc hình chữ nhật chỉnh chữ cái nào?* (`Rx, Ry`)\n'
            '  *2. Muốn xẻ nửa hình tròn thành miếng dưa hấu chỉnh Start và End bao nhiêu độ?* (`0` và `180`)\n'
            '  *3. Câu hỏi trang 70 SGK: Cho ngôi sao 5 cánh nhọn đi spa mài cong các đỉnh thành cánh hoa thì chỉnh tham số nào?* (`Rounded`)\n'
            '- *"Nói suông sợ chưa tin, nhìn máy chiếu thầy bấm phép biến hình trong 30 giây nè!"*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 2):**\n'
            '  1. Nháy đúp vào **Hình chữ nhật xanh**, gõ `Rx = 25` rồi nhấn `Enter`: bo tròn 4 góc ngay!\n'
            '  2. Nháy đúp vào **Hình tròn cam**, gõ `Start = 0`, `End = 180` rồi nhấn `Enter`: xẻ ngay nửa hình tròn!\n'
            '  3. Nháy đúp vào **Ngôi sao đỏ**, sửa `Rounded = 0.25` rồi nhấn `Enter`: biến ngay thành bông hoa 5 cánh bầu bĩnh!\n'
            '- **Chữa cháy:** Nếu HS đọc sai tên tiếng Anh `Rx, Ry` hay `Rounded`, cười bảo *"Đọc theo kiểu phiên âm Việt hoá cũng tính điểm, miễn chỉ đúng dòng trong SGK trang 70!"* rồi ghi chữ chuẩn lên bảng.'
        ]
    ]
    b.append(eng.table(c2a_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2B ---
    eng.add(b, 'Chặng 2B (Phút 14–23): Nhiệm vụ 2 – Màu tô, màu vẽ & Fill and Stroke (SGK tr. 70–71)', 'sub')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    c2b_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Ở Bài 12 muốn tô màu thì nháy ở dải màu dưới đáy. Nhưng tô vậy hơi phẳng, chưa ra chất 3D! Nay bấm Shift + Ctrl + F gọi hộp thoại chuyên nghiệp Fill and Stroke."*\n'
            '- *"Nhìn 5 ô vuông Hình 13.3 trang 70 bằng mắt thường, bạn nào tả hộ thầy phần ruột bên trong khác gì nhau? Nói tiếng Việt bình dân thôi!"* (Ô 1 rỗng, Ô 2 đặc 1 màu, Ô 3 nhạt dần ngang, Ô 4 sáng giữa tâm, Ô 5 mờ nhìn xuyên).\n'
            '- *"Hãy nhớ: Fill là màu áo, Stroke paint là màu chỉ may viền, Stroke style là độ dày chỉ may hay nét đứt. Còn Opacity giống như đeo kính râm, kéo xuống là nhìn xuyên thấu qua lớp dưới!"*\n'
            '- *"Giờ cả lớp bình chọn bằng ngón tay nha (1 ngón = A, 2 = B, 3 = C, 4 = D) cho 2 câu trắc nghiệm SGK trang 71!"*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 3):**\n'
            '  + Hộp thoại `Fill and Stroke` đã mở sẵn bên phải: bấm vào **Hình tròn xanh**, chọn tab `Fill`, bấm nút **`Radial gradient`** (nút thứ 4) sẽ biến ngay thành quả bóng 3D!\n'
            '  + Chuyển tab `Stroke style`, bấm `Dashes` đổi viền sang nét đứt.\n'
            '  + Bấm **Hình tròn xanh dương** đè trên chữ nhật vàng, kéo thanh `Opacity` xuống `40% – 80%` cho HS thấy nhìn xuyên qua.\n'
            '- **Đáp án tr. 71:** Câu 1 giơ 3 ngón (**C. Stroke style**); Câu 2 giơ 2 ngón (**B** – cây bút).\n'
            '- **Chữa cháy:** Nếu HS ngại giơ tay, hô *"1, 2, 3 cả lớp cùng giơ một lượt nhé!"* để không ai bị ngại.'
        ]
    ]
    b.append(eng.table(c2b_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2C ---
    eng.add(b, 'Chặng 2C (Phút 23–33): Nhiệm vụ 3 – Các phép ghép hình trong Path (SGK tr. 71–72)', 'sub')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    c2c_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Các bạn nhìn Hình 13.5 trang 71: Từ 1 hình vuông với 2 hình tròn mà ra hình trái tim — không lẽ ngồi vẽ tay cho méo? Dân thiết kế và AI người ta dùng phép toán trong menu Path!"*\n'
            '- Chiếu Hình 13.6 (tr. 72): *"Trong Toán lớp 10 mới học Tập hợp A ∪ B, A ∩ B, A \\ B đúng không? Ở Hình 13.6 có hình vuông vẽ trước (lớp dưới) và hình tròn vẽ sau (lớp trên). Nhìn hình đọc tên giúp thầy:"*\n'
            '  1. *"Phép nào hàn xì hai đứa thành một cục (Union)?"*\n'
            '  2. *"Phép nào lấy hình tròn cắn đứt 1 miếng của hình vuông (Difference)? — Bí mật: Logo Quả táo cắn dở Apple làm đúng 1 nhát bằng lệnh này!"*\n'
            '  3. *"Phép nào chỉ lấy đúng phần đè lên nhau (Intersection)? — Đó là cách vẽ chiếc lá trên đầu táo Apple!"*\n'
            '- *(Giới thiệu nhanh 3 phép phụ còn lại trong H.13.6: Exclusion đục lỗ giữa, Division cắt rời giữ màu, Cut Path cắt đứt viền mất màu)*.\n'
            '- Hỏi câu trang 72: *"Vậy muốn ghép nhiều hình tròn thành Đám mây Hình 13.7, chọn phép nào?"* (`Path / Union`!).',
            '- **Thao tác AN TOÀN TUYỆT ĐỐI (mau_bai13.svg – Khu vực 4):**\n'
            '  + **A (Đám mây):** Quét chọn 3 hình tròn và 1 chữ nhật xanh nhạt, **KHÔNG bấm phím tắt Ctrl + +** (hay kẹt Shift trên laptop), mà mở thẳng menu **Path chọn Union** cho cả lớp đọc rõ chữ, lập tức biến thành Đám mây H.13.7!\n'
            '  + **B & C (Cắt Vuông-Tròn & Logo Apple):** Quét chọn hình vuông đỏ với hình tròn cam (hoặc quả táo đen với hình tròn trắng xám bên cạnh), mở menu **Path chọn Difference**, táo cắn dở hiện ra ngay lập tức (cả lớp vỗ tay liền)!\n'
            '- **Chữa cháy:** Nếu bấm `Difference` mà thấy mất hình gốc? Bình tĩnh bấm `Ctrl + Z`, cười giải thích: *"Hình nào vẽ sau nằm trên sẽ đóng vai dao cắt; muốn hình gốc không mất thì bấm phím End cho nó xuống đáy rồi cắt lại!"*.'
        ]
    ]
    b.append(eng.table(c2c_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 3 ---
    eng.add(b, 'Chặng 3 (Phút 33–40): Luyện tập – Trò chơi "Nhìn hình đoán 1 từ khoá" (SGK tr. 74)', 'sub')
    eng.add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    c3_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Không kiểm tra giấy! Thầy mời cả lớp chơi trò Nhìn hình đoán mật mã ở SGK trang 74! Cực dễ, mỗi hình chỉ cần nói đúng 1 từ khoá:"*\n'
            '  + *"Hình 13.14 (Bông hoa 10 cánh): Cho ngôi sao đi spa uốn tròn cánh dùng chữ tiếng Anh nào? Nhụy ở giữa vẽ bằng hình gì?"* (`Rounded` + Hình tròn vàng).\n'
            '  + *"Hình 13.15 (Quân Bích đỏ): Tán trên giống hình gì lộn ngược? Chân đế bị khoét lõm hai bên dùng phép gì trong Path?"* (`Union` ghép trái tim + `Difference` khoét đế + `Union`).\n'
            '  + *"Hình 13.16 (Chùm bóng bay nhìn xuyên nhau): Đeo kính râm cho quả bóng bằng thông số gì?"* (`Opacity` khoảng 75%–80%!).',
            '- **Thao tác (mau_bai13.svg – Khu vực 5):** Chỉ vào hình mẫu Bông hoa H.13.14 và Chùm bóng bay H.13.16 dựng sẵn ở Khu vực 5 để đối chiếu ngay cho lớp xem (không mất công vẽ lại từ đầu).\n'
            '- Gọi luân phiên 3 bạn ở các dãy bàn khác nhau, khen ngay *"Chuẩn không cần chỉnh!"* để giữ sinh khí vui vẻ cuối giờ.\n'
            '- **Chữa cháy:** Nếu HS quên từ tiếng Anh của quân Bích (Hình 13.15), chỉ cần gợi *"Trái tim lật ngược hồi nãy mình ghép bằng lệnh gì?"*, khi đó HS nhìn bảng giữa sẽ nhớ ngay ra `Union` và `Difference`.'
        ]
    ]
    b.append(eng.table(c3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 4 ---
    eng.add(b, 'Chặng 4 (Phút 40–45): Vận dụng – Chốt hạ Miếng dưa hấu & Dặn dò Tiết 2 (SGK tr. 74)', 'sub')
    c4_rows = [
        ['MÌNH NÓI GÌ? (Lời thoại tự nhiên, vui nhộn)', 'TAY BẤM GÌ (mau_bai13.svg) & CHỮA CHÁY KHI LỚP IM RE'],
        [
            '- *"Quay lại miếng dưa hấu đầu giờ! Bây giờ đã học xong thuộc tính và phím tắt rồi, ai tóm tắt nhanh trong 3 ý ngắn làm sao vẽ được miếng dưa hấu này trên máy?"*\n'
            '  1. *"Xẻ nửa hình tròn xanh bằng thông số gì?"* (`Start = 0, End = 180`, chế độ quạt `Slice`).\n'
            '  2. *"Nhân đôi thu nhỏ thành cùi trắng và ruột đỏ bằng phím tắt gì ở Bài 12?"* (`Ctrl + D` rồi kéo thu nhỏ, đổi màu).\n'
            '  3. *"Hạt dưa màu đen vẽ bằng hình gì và nằm lớp nào?"* (Elip nhỏ màu đen ở lớp trên cùng).\n'
            '- *(Tổng kết & Dặn dò)*: *"Cả lớp hôm nay nắm bài cực chắc! Về nhà xem trước 3 bài thực hành trang 72–74 (Logo 3 cánh, Hình cây xanh, Hình cầu 3D) để Tiết 2 Thực hành làm cho mượt nha. Thầy cảm ơn Thầy Văn và cả lớp!"*',
            '- **Thao tác chốt hạ (mau_bai13.svg – Khu vực 1 bên phải):** Tận dụng ngay hình tròn xanh đang chờ sẵn ở Khu vực 1: nhấn `E`, chỉnh `Start = 0, End = 180`, rồi bấm `Ctrl + D` thu nhỏ tô trắng, tiếp tục `Ctrl + D` thu nhỏ tô đỏ, sau đó rải 2 chấm elip đen: chưa tới 25 giây hiện ngay miếng dưa hấu mới tinh khép lại tiết học cực ấn tượng!\n'
            '- **Lưu ý giờ:** Nếu còn dưới 2 phút mà chuông sắp reo thì không cần vẽ hạt, chỉ bấm cắt nửa hình tròn và `Ctrl + D` 1 lần tô đỏ minh hoạ rồi chốt dặn dò ngay cho kịp giờ.'
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
        'hinh_13_14_15_16_luyen_tap.png',
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
        page1_sched_line='Bối cảnh phòng học: 01 máy tính GV + máy chiếu + bảng phấn (HS sử dụng trực tiếp SGK Tin học 10 tr. 69–74)',
        page1_mode_line='Chiến thuật dạy: Dẫn dắt như một bài thuyết trình sinh động – Hỏi câu nhìn SGK đọc được – Bấm sẵn mau_bai13.svg'
    )

    # HƯỚNG DẪN ĐỌC KỊCH BẢN
    eng.add(b, 'QUY ƯỚC ĐỌC NHANH TRÊN BỤC GIẢNG (LIẾC MẮT LÀ BIẾT LÀM GÌ)', 'heading')
    for item in [
        '**[Hành động & Thao tác]:** Cho bạn biết lúc nào bước ra giữa bục, lúc nào chỉ tay lên màn chiếu, lúc nào bấm vào khu vực nào của mau_bai13.svg trên Inkscape.',
        '**[Lời nói trực tiếp]:** Câu thoại bạn nói nguyên văn hoặc biến tấu tự nhiên với lớp, giọng hào hứng, thân thiện như đang thuyết trình kể chuyện.',
        '**[Tương tác với học sinh]:** Cách đặt câu hỏi ngắn, tổ chức giơ ngón tay, trò chơi đoán từ khoá để lớp không bị chìm vào im lặng.',
        '**[Chốt bảng & Ghi vở]:** Câu ngắn gọn đọc chậm cho học sinh chép vào vở và tóm tắt nhanh lên phần bảng giữa.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 0
    eng.add(b, 'PHẦN 0. CHUẨN BỊ 3 PHÚT TRƯỚC KHI CHUÔNG REO (KHÔNG ĐỂ LỖI KỸ THUẬT)', 'heading')
    for item in [
        '**[Hành động & Thao tác]:** Vừa bước vào phòng H.I.5 giờ ra chơi, bạn cắm dây máy chiếu, mở ngay tệp mau_bai13.svg bằng phần mềm Inkscape (để máy tải trước, tránh lúc vào tiết phải đứng đợi 15 giây).',
        '**[Hành động & Thao tác]:** Nhấn sẵn tổ hợp phím Shift + Ctrl + F để hộp thoại Fill and Stroke mở cố định ở góc phải màn hình. Vẽ thử 1 hình tròn nháp ra vùng trắng bên ngoài để chắc chắn không bị tàng hình (Opacity phải là 100%) và là hình tròn khép kín, rồi nhấn Delete xoá hình nháp.',
        '**[Viết bảng sẵn]:** Ở góc trái bảng viết sẵn ngày tháng năm và cụm *"Tiết 19: Bài 13. Bổ sung các đối tượng đồ hoạ"*; chừa khoảng bảng giữa rộng để ghi 3 ý chính trong giờ.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 1
    eng.add(b, 'PHẦN 1. CHÀO HỎI & MỞ ĐẦU "MIẾNG DƯA HẤU" (PHÚT 00 – 05)', 'heading')
    eng.add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    for item in [
        '**[Hành động & Thao tác]:** Đứng thẳng giữa bục giảng, mỉm cười nhìn bao quát cả lớp và hướng về phía Thầy Mai Xuân Văn ngồi cuối lớp (nếu thầy dự giờ), ra hiệu cho lớp ổn định chỗ ngồi trong 10 giây.',
        '**[Lời nói trực tiếp]:** *"Thầy chào cả lớp 10/3! Hôm nay trời có chút oi bức, nên mở đầu tiết Tin học thầy chưa vội kiểm tra bài cũ khô khan đâu, mà thầy xin mời cả lớp ăn... một miếng dưa hấu bằng mắt trước nha! Các em mở giúp thầy Sách giáo khoa Tin học 10 trang 69 hoặc nhìn thẳng lên màn chiếu Hình 13.1."*',
        '**[Hành động & Thao tác]:** Chuyển màn hình sang Khu vực 1 bên trái của mau_bai13.svg, nơi bạn đã dựng sẵn Miếng dưa hấu 4 lớp.',
        '**[Tương tác với học sinh]:** *"Thử thách 30 giây siêu dễ dành cho cả lớp: Nhìn vào miếng dưa hấu Hình 13.1 này, bạn nào xung phong kể giúp thầy nó được ghép từ mấy lớp màu và gồm những hình cơ bản nào?"* — Mời 01 học sinh đứng dậy trả lời (vỏ xanh đậm, cùi trắng, ruột đỏ, hạt đen; các nửa hình tròn và hạt nhỏ).',
        '**[Lời nói trực tiếp]:** *"Quá chuẩn, cảm ơn em! Giờ cả lớp nhớ lại quy tắc vàng ở Bài 12 tuần trước nè: trong Inkscape, hình nào vẽ trước thì nằm ở lớp dưới, hình nào vẽ sau thì đè lên lớp trên, giống y như ly trà sữa trân châu: trân châu bỏ trước nằm đáy, đá và kem cheese bỏ sau đè lên trên. Vậy bạn nào trả lời nhanh giúp thầy: giữa lớp vỏ xanh, cùi trắng, ruột đỏ và mấy hạt dưa màu đen, lớp nào nằm dưới đáy (lớp 1) và lớp nào phải nằm trên cùng (lớp 4)?"*',
        '**[Hành động & Thao tác]:** Vừa nói vừa dùng chuột kéo tách rời từng lớp của miếng dưa hấu ở Khu vực 1 trên máy chiếu ra cho cả lớp nhìn thấy tận mắt: vỏ xanh ở dưới cùng, tới cùi trắng, tới ruột đỏ, còn hạt đen nằm trên cùng. Sau đó nhấn Ctrl + Z liên tiếp cho các lớp nhảy về khớp lại như cũ.',
        '**[Cứu cánh nếu lớp im lặng]:** Cười tươi hỏi mẹo bằng giọng vui: *"Ủa, đố cả lớp nè: nếu thầy lỡ xếp hạt đen nằm dưới lớp ruột đỏ thì mắt mình còn thấy cái hạt nào nữa không?"*, khi đó lớp sẽ tự khắc cười và đồng thanh đáp *"Không, phải nằm trên cùng!"*.',
        '**[Lời nói trực tiếp – Dẫn vào bài]:** *"Tuyệt vời! Ở Bài 12 các em mới vẽ được hình tròn nguyên vẹn như nguyên cái bánh pizza, tô được một màu phẳng lì chứ chưa biết xẻ đôi thành nửa miếng dưa hấu, chưa biết tô màu bóng bẩy 3D hay cắt ghép các hình lại với nhau. Hôm nay, Bài 13: Bổ sung các đối tượng đồ hoạ sẽ dạy chúng ta đúng 3 bí kíp đó!"*',
        '**[Viết bảng & Chốt ghi vở]:** Viết to tên bài ở giữa bảng và dặn nhanh: *"Các em mở vở ghi tựa Bài 13 và 4 dòng ngắn gọn thứ tự lớp của Hình 13.1 nhé: Vỏ xanh ở lớp dưới cùng, tiếp đến cùi trắng, tiếp theo ruột đỏ và trên cùng là hạt đen."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 2
    eng.add(b, 'PHẦN 2. THUYẾT TRÌNH MỤC 1 – THUỘC TÍNH BIẾN HÌNH CỦA HÌNH KHỐI (PHÚT 05 – 14)', 'heading')
    eng.add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Chúng ta bước vào bí kíp thứ nhất: Chỉnh thuộc tính để biến hình các khối cơ bản. Trước khi đọc chữ, cả lớp nhìn lên Hình 13.2 trên màn chiếu hoặc trong SGK trang 69 giúp thầy."*',
        '**[Tương tác với học sinh]:** *"Nhìn 4 hình màu đỏ ở Hình 13.2: một ngũ giác đều, một ngôi sao nhọn, một ngôi sao bầu bĩnh như bông hoa và một hình méo ngẫu nhiên. Nhìn mặt mũi khác hẳn nhau như vậy nhưng thực ra cả 4 đứa đều sinh ra từ CÙNG MỘT MẸ là một công cụ đã học ở Bài 12. Đố cả lớp là công cụ nào trong 3 hình: Hình chữ nhật, Hình tròn hay Hình đa giác và Hình sao?"* — Gọi 01 học sinh đọc to đáp án (*"Hình đa giác / Hình sao, phím sao!"*).',
        '**[Lời nói trực tiếp – Giao việc tra SGK cực dễ]:** *"Chính xác 100%! Vậy làm sao từ 1 công cụ mà nặn ra được 4 kiểu khác nhau như thế? Bí mật nằm ở Thanh điều khiển thuộc tính nằm ngang phía trên cùng màn hình. Thầy không bắt các em học vẹt lý thuyết đâu, bây giờ cho hai bạn cùng bàn chụm đầu rà mắt thẳng vào cái Bảng ở trang 69 và đầu trang 70 SGK trong 2 phút, lấy bút chì gạch chân giúp thầy đúng 3 từ khoá tiếng Anh trả lời 3 câu hỏi sau:"*',
        '**[Tương tác với học sinh – 3 câu hỏi tra bảng SGK]:** Đọc từng câu rõ ràng để học sinh gạch thẳng vào sách:\n'
        '  Câu 1: *"Điện thoại iPhone thì bo tròn 4 góc mềm mại, còn viên gạch thẻ thì góc vuông 90 độ. Muốn bo tròn 4 góc của hình chữ nhật, các em sửa hai thông số nào?"* (**Rx, Ry**).\n'
        '  Câu 2: *"Muốn cắt đôi hình tròn thành nửa hình tròn như miếng dưa hấu lúc nãy, các em nhập hai thông số Start và End bằng bao nhiêu độ?"* (**0** và **180**).\n'
        '  Câu 3 (Câu hỏi SGK trang 70): *"Muốn cho ngôi sao 5 cánh nhọn đi spa mài cong các đỉnh thành cánh bông hoa mềm mại như hình thứ ba, các em thay đổi tham số nào?"* (**Rounded**).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 2]:** Sau khi mời 03 học sinh đọc to 3 từ khoá vừa gạch chân trong sách (kể cả bạn học yếu nhìn sách đọc vẫn trúng), bạn chuyển ngay sang Khu vực 2 của mau_bai13.svg: *"Nói suông sợ các bạn chưa tin, nhìn lên máy chiếu thầy bấm phép biến hình trong đúng 30 giây nè!"*\n'
        '  Thao tác 1: Nháy đúp chuột vào **Hình chữ nhật màu xanh**, chỉ chuột vào thanh ngang trên cùng gõ Rx = 25 rồi nhấn Enter: *"Đó, từ viên gạch thẻ thành màn hình iPhone bo góc liền!"*\n'
        '  Thao tác 2: Nháy đúp vào **Hình tròn màu cam**, gõ Start = 0, End = 180 rồi nhấn Enter: *"Xong luôn nửa cái bánh, chỉ chuột lên 3 nút hình tròn nhỏ cuối thanh nè: nút Hình quạt có nối tâm làm miếng dưa hấu, nút Cung tròn hở mất viền đáy dùng vẽ logo, còn nút bên phải đóng tròn nguyên vẹn lại."* (nhấn nhẹ giới thiệu rồi hoàn tác lại).\n'
        '  Thao tác 3: Nháy đúp vào **Ngôi sao màu đỏ**, sửa ô Rounded = 0.25 rồi nhấn Enter: *"Tèn ten! Ngôi sao 5 cánh vừa biến thành bông hoa 5 cánh tròn trịa rồi đó!"*',
        '**[Cứu cánh nếu học sinh đọc sai phát âm tiếng Anh]:** Vui vẻ khích lệ: *"Không sao, môn Tin không chấm điểm phát âm chuẩn như môn Tiếng Anh, miễn các em chỉ đúng chữ Rx, Ry, Start, End và Rounded trong sách trang 70 là thầy cộng điểm cộng ngay!"*',
        '**[Viết bảng & Chốt ghi vở]:** Viết gọn Mục 1 vào phần bảng giữa và đọc chậm: *"Các em ghi ngắn gọn vào vở Mục 1. Thuộc tính hình khối (SGK tr. 69–70): Chữ nhật có W, H (rộng, cao) và Rx, Ry (độ bo góc). Hình tròn có Start, End (độ mở góc, từ 0 đến 180 tạo nửa hình tròn). Đa giác và sao có Corners (số đỉnh hoặc cánh), Spoke ratio (tỉ lệ nhánh), Rounded (độ cong đỉnh, dùng biến ngôi sao thành cánh hoa) và Randomized (độ méo ngẫu nhiên)."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 3
    eng.add(b, 'PHẦN 3. THUYẾT TRÌNH MỤC 2 – TÔ MÀU 3D, NÉT VẼ & ĐỘ TRONG SUỐT (PHÚT 14 – 23)', 'heading')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    for item in [
        '**[Lời nói trực tiếp – Liên hệ Bài 12]:** *"Vừa xong bí kíp nặn hình. Giờ qua bí kíp thứ hai: Làm cho hình đẹp lên bằng màu sắc và đường nét! Bạn nào còn nhớ ở Bài 12, muốn tô màu nền và màu viền nhanh dưới dải màu đáy màn hình mình bấm chuột sao không?"* — Gọi nhanh 1 bạn nhắc lại (*"Nháy thường đổi màu nền, giữ Shift và nháy chuột đổi màu viền!"*).',
        '**[Lời nói trực tiếp]:** *"Chuẩn luôn! Nhưng tô ở dải màu dưới đáy chỉ ra 1 màu phẳng lì thôi, nhìn chưa xịn, chưa có hiệu ứng 3D lung linh. Bây giờ mời cả lớp nhìn 5 ô vuông viền cam ở Hình 13.3 trang 70 SGK."*',
        '**[Tương tác với học sinh]:** *"Bằng mắt thường rất bình dân thôi, không cần thuật ngữ cao siêu gì hết, bạn nào mô tả giúp thầy phần ruột màu xanh bên trong 5 ô vuông Hình 13.3 khác nhau điểm nào?"* — Mời 01 học sinh tả theo mắt nhìn (*"Ô 1 trắng trơn không có màu; Ô 2 xanh đặc một màu; Ô 3 chuyển màu từ đậm sang nhạt nằm ngang; Ô 4 sáng ở giữa rồi đậm ra ngoài; Ô 5 bị mờ đi..."*).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 3]:** Chuyển sang Khu vực 3 và chỉ vào bảng Fill and Stroke đã mở sẵn ở mép phải màn hình (chỉ luôn phím tắt Shift + Ctrl + F và Hình 13.4):\n'
        '  Giải thích nhanh bằng ví von: *"Để làm được 5 kiểu màu vi diệu đó, chúng ta dùng hộp thoại Fill and Stroke. Hãy nhớ câu thần chú bình dân này cho không bao giờ nhầm: Fill là cái áo mặc bên trong, Stroke paint là màu sợi chỉ may đường viền bên ngoài, còn Stroke style là độ dày mỏng của đường chỉ và chọn đường nét liền hay nét đứt!"*\n'
        '  Bấm vào **Hình tròn màu xanh** ở Khu vực 3, chọn thẻ Fill, bấm vào biểu tượng thứ tư **Radial gradient** (hình ô vuông sáng ở giữa tâm): *"Các bạn thấy từ hình tròn phẳng lì nó biến thành quả bóng 3D nổi khối chưa? Còn nút kế bên trái là Linear gradient – chuyển màu thẳng như ly trà đào cam sả từ vàng xuống cam đó!"*\n'
        '  Bấm qua thẻ Stroke style, chỉ vào ô Width và mục Dashes chọn nét đứt cho lớp xem đường viền biến thành nét đứt.\n'
        '  Bấm vào **Hình tròn xanh dương** đang đè lên hình chữ nhật vàng, kéo thanh trượt **Opacity** phía dưới từ 100% xuống khoảng 50%: *"Còn chữ Opacity và kênh Alpha này y hệt như đeo kính râm vậy, giảm phần trăm xuống thì hình trở nên trong suốt, nhìn xuyên thấu thấy lớp nằm bên dưới! Bạn nào hay thiết kế Canva hay Figma trên điện thoại cũng sẽ thấy đúng chữ Opacity này luôn!"*',
        '**[Tương tác với học sinh – Bình chọn ngón tay cho 2 câu trắc nghiệm SGK tr. 71]:** *"Bây giờ kiểm tra độ tỉnh táo nè! Cả lớp nhìn 02 câu trắc nghiệm ở giữa trang 71 SGK. Thay vì chỉ một bạn giơ tay cho ngại, thầy đếm 1, 2, 3 cả lớp đồng loạt giơ ngón tay chọn đáp án nhé: 1 ngón là A, 2 ngón là B, 3 ngón là C, 4 ngón là D!"*\n'
        '  Câu 1 (chọn đường viền nét đứt ở trang nào): Đếm *"1, 2, 3 giơ tay!"* — Chốt đáp án **C. Stroke style** (3 ngón).\n'
        '  Câu 2 (chọn biểu tượng chỉnh điểm dừng màu gradient): Đếm *"1, 2, 3 giơ tay!"* — Chốt đáp án **B** (cây bút chỉnh sửa gradient, 2 ngón).',
        '**[Viết bảng & Chốt ghi vở]:** *"Tuyệt vời, lớp mình bắt nhịp rất nhanh! Các em ghi nhanh vào vở Mục 2. Hộp thoại Fill and Stroke (mở bằng phím Shift + Ctrl + F): gồm 3 trang là Fill (màu tô nền: không màu, đồng nhất Flat, chuyển thẳng Linear, chuyển toả tròn Radial, hoa văn Pattern), Stroke paint (màu viền), Stroke style (độ dày và nét đứt). Opacity và Alpha (A) chỉnh độ trong suốt. Ghi thêm đáp án tr. 71: Câu 1 C, Câu 2 B để cuối kì ôn thi khỏi tìm lại nhé!"*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 4
    eng.add(b, 'PHẦN 4. THUYẾT TRÌNH MỤC 3 – CÁC PHÉP GHÉP ĐỐI TƯỢNG TRONG MENU Path (PHÚT 23 – 33)', 'heading')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    for item in [
        '**[Lời nói trực tiếp – Dẫn nhập vui]:** *"Mình đã biết nặn hình, biết tô màu 3D rồi. Bây giờ qua bí kíp đỉnh cao nhất của Bài 13: Ghép hình và cắt gọt hình! Các em nhìn Hình 13.5 trang 71 trên màn chiếu: Từ hai hình tròn đỏ và một hình vuông xoay nghiêng mà ghép thành hình trái tim cân đối tăm tắp — không lẽ ngồi cầm bút vẽ tay cho méo xẹo đúng không? Dân thiết kế đồ hoạ người ta dùng các phép toán ghép hình trong menu Path!"*',
        '**[Liên hệ Toán lớp 10 & Khám phá Hình 13.6 tr. 72]:** *"Nói phép toán đừng sợ, nó chính là 3 phép Tập hợp các em mới học trong môn Toán lớp 10 đó: Hợp, Hiệu và Giao! Ở Hình 13.6 trang 72 luôn có quy ước sẵn: Hình vuông vẽ trước nằm ở lớp DƯỚI, còn Hình tròn vẽ sau nằm đè ở lớp TRÊN. Hai bạn cùng bàn liếc nhìn 6 hình kết quả ở Hình 13.6 trong 2 phút và trả lời giúp thầy 2 câu:"*\n'
        '  Câu 1: *"Trong 6 hình ở Hình 13.6, phép nào hàn xì hai hình lại thành một cục duy nhất (Union)? Phép nào lấy hình tròn làm khuôn bánh cắn khuyết mất một góc của hình vuông (Difference)? Và phép nào chỉ giữ lại đúng khúc giao nhau ở giữa (Intersection)?"*\n'
        '  Câu 2 (Câu hỏi SGK trang 72): *"Vậy nhìn qua Hình 13.7, để ghép mấy hình tròn chồng lên nhau thành Đám mây hoàn chỉnh không còn nét cắt bên trong, ta chọn phép nào?"*',
        '**[Tương tác & Chuẩn hoá]:** Mời 02 học sinh phát biểu nhanh (Đám mây dùng Path / Union), rồi bạn giải thích nốt 3 phép còn lại ở Hình 13.6 bằng ví von dễ nhớ: *"Ngoài 3 phép chính vừa nêu, các em nhớ thêm 3 phép phụ trong bảng trang 72: Exclusion (hiệu đối xứng) là giữ hai rìa ngoài nhưng đục thủng lỗ ở giữa; Division (phép chia) như con dao thái bánh cắt hình dưới ra nhiều mảnh rời mà vẫn còn màu tô; còn Cut Path là cái kéo chỉ cắt đứt đường viền nên mất sạch màu tô bên trong!"*',
        '**[Hành động & Thao tác "Ảo thuật" trên mau_bai13.svg – Khu vực 4]:** Chuyển sang Khu vực 4 để biểu diễn cho cả lớp trầm trồ (bấm menu bằng chuột cho chắc ăn, không bấm phím tắt Ctrl + + kẻo kẹt Shift):\n'
        '  **Màn 1 (Đám mây H.13.7):** Quét chuột chọn 4 hình xanh ở ô A, mở menu **Path** rồi bấm **Union**: *"Một, hai, ba... Biến thành đám mây mịn màng ngay lập tức!"*\n'
        '  **Màn 2 (Logo quả táo cắn dở Apple – Liên hệ thực tế):** *"Các em dùng điện thoại có biết Logo quả táo cắn dở của hãng Apple vẽ sao cho cái vết cắn bên phải tròn đều tăm tắp không? Người ta xếp 1 quả táo nguyên ở lớp dưới, đặt đè 1 hình tròn nhỏ lên mép phải ở lớp trên, rồi làm đúng 1 nhát!"* — Quét chọn quả táo đen và hình tròn trắng xám ở ô C, mở menu **Path** rồi bấm **Difference**: *"Rắc! Cắn một miếng chuẩn như hãng Apple luôn! Còn cái cuống lá trên đỉnh quả táo chính là 2 hình tròn lấy phần giao nhau ở giữa bằng lệnh Intersection đó!"*',
        '**[Mở rộng kiến thức AI & Công nghệ thực tế (30 giây)]:** *"Chưa hết đâu nha: mấy máy cắt chữ mica hay cắt Laser biển quảng cáo ngoài tiệm là đọc trực tiếp các đường Path trong tệp vectơ này đó. Sau này các em nhờ AI vẽ ảnh như ChatGPT hay Gemini: nếu AI xuất ảnh PNG thông thường phóng to hay bị mờ, còn nhờ AI xuất ra mã vectơ SVG thì bỏ vào Inkscape vẫn mở từng lớp Path ra sửa y chang như mình đang làm hôm nay!"*',
        '**[Viết bảng & Chốt ghi vở]:** Nhấn mạnh thật rõ rồi đọc cho lớp ghi Mục 3: *"Các em đóng khung đỏ hoặc gạch chân thật đậm quy tắc sống còn này vào vở: Với lệnh cắt Difference, Division và Cut Path, HÌNH LỚP DƯỚI luôn bị HÌNH LỚP TRÊN cắt! Còn nhớ mẹo lỡ bấm nhầm mà thấy hình gốc biến mất thì cứ nhấn Ctrl + Z, chọn hình cần giữ rồi nhấn phím End đưa nó xuống lớp dưới cùng xong cắt lại nhé!"*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 5
    eng.add(b, 'PHẦN 5. MINI GAME LUYỆN TẬP "NHÌN HÌNH ĐOÁN MẬT MÃ" (PHÚT 33 – 40)', 'heading')
    eng.add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Học lý thuyết đủ 3 bí kíp rồi, không kiểm tra giấy, thầy mời cả lớp thư giãn 7 phút cuối bằng trò chơi Nhìn hình đoán mật mã ở trang 74 SGK! Luật chơi cực nhẹ nhàng: thầy chiếu 3 bức hình Hình 13.14, Hình 13.15, Hình 13.16; các em chỉ cần nói đúng 1 hoặc 2 từ khoá tiếng Anh vừa học để vẽ ra hình đó!"*',
        '**[Tương tác với học sinh – Gọi luân phiên 3 dãy bàn]:**\n'
        '  **Lượt 1 – Hình 13.14 (Bông hoa 10 cánh):** *"Hình này vẽ từ Ngôi sao 10 cánh. Đố bạn dãy trái: dùng từ khoá gì để uốn các cánh nhọn thành cánh tròn mượt, rồi vẽ thêm hình gì ở giữa làm nhụy?"* — Chốt: **Rounded** và **Hình tròn màu vàng**.\n'
        '  **Lượt 2 – Hình 13.15 (Quân Bích đỏ):** *"Tán trên giống trái tim lộn ngược mình vừa ghép ở Hình 13.5 bằng lệnh gì? Còn phần chân đế bị khoét lõm hai bên ta gọt bằng lệnh nào?"* — Nếu lớp ngập ngừng, gợi ngay *"Trái tim lúc nãy ghép bằng phép hợp gì?"*, học sinh nhìn bảng giữa sẽ đọc ngay: **Union** và **Difference**!\n'
        '  **Lượt 3 – Hình 13.16 (Chùm bóng bay nhìn xuyên nhau):** *"Quả bóng vàng với bóng xanh nhìn xuyên thấu qua nhau như đeo kính râm là mình kéo thông số gì ở góc hộp thoại màu?"* — Chốt: Kéo giảm **Opacity** xuống khoảng 80%!',
        '**[Hành động & Thao tác trên mau_bai13.svg – Khu vực 5]:** Chỉ chuột sang Khu vực 5 nơi đã có sẵn mẫu Bông hoa và Chùm bóng bay cho lớp kiểm chứng trực quan, khen ngợi cả 3 bạn vừa trả lời bằng một tràng pháo tay khích lệ.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 6
    eng.add(b, 'PHẦN 6. VẬN DỤNG CHỐT HẠ "MIẾNG DƯA HẤU" & CHÀO KẾT THÚC (PHÚT 40 – 45)', 'heading')
    for item in [
        '**[Lời nói trực tiếp – Khép vòng tròn bài học]:** *"Còn 5 phút cuối cùng, chúng ta quay lại đúng câu đố đầu giờ: Miếng dưa hấu Hình 13.1 ở trang 69! Bây giờ nắm đủ 3 bí kíp trong tay rồi, bạn nào tóm tắt nhanh giúp thầy trong 3 ý ngắn làm sao vẽ được miếng dưa hấu này trên Inkscape?"*',
        '**[Tương tác với học sinh]:** Mời 01 hoặc 02 học sinh phát biểu miệng 3 bước:\n'
        '  **Bước 1:** Vẽ hình tròn bằng công cụ E, nhập **Start = 0, End = 180** (chế độ hình quạt Slice) tô xanh đậm làm vỏ dưa.\n'
        '  **Bước 2:** Bấm **Ctrl + D** nhân đôi rồi thu nhỏ lại tô trắng làm cùi; tiếp tục **Ctrl + D** thu nhỏ lần nữa tô đỏ làm ruột.\n'
        '  **Bước 3:** Vẽ các hình elip nhỏ màu đen rải lên lớp trên cùng làm hạt dưa.',
        '**[Hành động & Thao tác chốt màn cuối trên mau_bai13.svg – Khu vực 1 bên phải]:** *"Để kiểm chứng lời bạn nói chuẩn 100%, thầy làm trực tiếp trên máy chiếu trong đúng 25 giây nha!"* — Chọn hình tròn xanh đang chờ sẵn bên phải Khu vực 1, nhấn phím E, gõ Start = 0, End = 180, nhấn Ctrl + D thu nhỏ tô trắng, Ctrl + D thu nhỏ tô đỏ, thêm 2 hạt elip đen (nếu gần sát chuông hết giờ thì cắt nửa hình tròn và Ctrl + D 1 lần tô đỏ minh hoạ nhanh cho kịp). *"Xong! Vậy là mình giải quyết trọn vẹn thử thách từ đầu giờ rồi nha!"*',
        '**[Lời nói trực tiếp – Dặn dò & Chào kết thúc]:** *"Về nhà, các em đọc trước giúp thầy 3 bài thực hành ở trang 72 đến 74: Logo 3 cánh (Hình 13.8), Hình cây xanh (Hình 13.10) và Quả cầu 3D (Hình 13.12) để sang Tiết 2 xuống phòng máy thực hành làm thật nhanh và đẹp nhé. Tiết học hôm nay lớp 10/3 học rất sôi nổi và hợp tác! Em xin trân trọng cảm ơn Thầy Mai Xuân Văn và chúc các em những tiết học tiếp theo thật vui!"*',
        '**[Hành động kết thúc]:** Mỉm cười gật đầu chào lớp, tắt máy chiếu (hoặc lưu giữ nếu thầy dặn), dọn bảng phấn gọn gàng.'
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
        'hinh_13_14_15_16_luyen_tap.png',
    ]
    clean_and_validate_docx(b, out_path, used_figures=used_figs, is_khbd=False)


if __name__ == '__main__':
    DAY_THAT_DIR.mkdir(parents=True, exist_ok=True)
    build_khbd_day_that()
    build_kich_ban_vui_nhon()
    build_bai_thuyet_trinh_day_hoc()
    print('Done building all 3 Day_that docx files in:', DAY_THAT_DIR)
