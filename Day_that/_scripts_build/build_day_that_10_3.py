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

    # MỤC II: HỆ THỐNG KIẾN THỨC CỐT LÕI TỪ CON SỐ 0 CHO NGƯỜI CHƯA NẮM VỮNG
    eng.add(b, 'II. HỆ THỐNG KIẾN THỨC CỐT LÕI BÀI 13 DỄ HIỂU (ĐỌC 5 PHÚT LÀ NẮM CHẮC)', 'heading')
    eng.add(b, 'Phần này diễn giải toàn bộ thuật ngữ tiếng Anh trong SGK Tin học 10 (tr. 69–74) sang tiếng Việt dễ hiểu, giúp bạn tự tin giải thích trôi chảy dù học sinh hay Thầy Mai Xuân Văn có hỏi thêm chi tiết nào:', first=0)

    eng.add(b, '1. Thuộc tính hình khối (Mục 1 SGK tr. 69–70): Khác nhau giữa Hình vuông, Hình tròn, Ngôi sao', 'sub')
    for item in [
        '**Mẹo nhớ vị trí:** Nháy đúp chuột vào bất kỳ hình nào (hoặc chọn công cụ tương ứng ở thanh bên trái) thì thanh ngang trên cùng (**Thanh điều khiển thuộc tính** của Bài 12) sẽ lập tức hiện ra các ô thông số của hình đó.',
        '**Hình vuông / Hình chữ nhật (phím R):** `W, H` là chiều rộng, chiều cao; **`Rx, Ry`** là **độ bo tròn 4 góc** (`Rx = Ry = 0` thì góc nhọn vuông vức; tăng `Rx, Ry` lên thì góc bo tròn mềm mại như viền điện thoại).',
        '**Hình tròn / Hình elip (phím E):** `Rx, Ry` là bán kính ngang/dọc. Có **Start và End** (góc điểm đầu và điểm cuối tính theo độ, ví dụ `0` và `180` là cắt đúng một nửa vòng tròn) cùng **3 nút biểu tượng hình tròn** ở cuối thanh thuộc tính: **Hình quạt (Slice)** nối 2 đầu vào tâm như miếng bánh/dưa hấu; **Cung tròn hở (Arc)** chỉ giữ nét viền cong không nối tâm như chữ C; **Hình tròn khép kín (Whole)** đóng tròn nguyên vẹn.',
        '**Hình đa giác / Hình sao (phím *):** Có 2 nút chuyển chế độ hình đa giác (cạnh đều) hoặc ngôi sao nhọn; **`Corners`** là số đỉnh/số cánh; **`Spoke ratio`** là tỉ lệ bán kính trong/ngoài quyết định cánh sao mập hay gầy nhọn; **`Rounded`** là độ uốn cong tại các đỉnh (làm cong đỉnh nhọn của ngôi sao thành cánh hoa); **`Randomized`** là độ méo ngẫu nhiên (bằng `0` thì đối xứng đều đặn, tăng lên thì hình méo mó tự do).'
    ]:
        eng.add(b, item, 'bullet')

    eng.add(b, '2. Hộp thoại Fill and Stroke (Mục 2 SGK tr. 70–71): Phân biệt 3 trang & 5 kiểu tô màu', 'sub')
    for item in [
        '**Cách gọi hộp thoại:** Menu **`Object / Fill and Stroke`** hoặc bấm **`Shift + Ctrl + F`**. Nhớ phân biệt: **`Fill`** là màu sơn **bên trong nền**, còn **`Stroke`** là màu **đường viền bên ngoài**.',
        '**Ba trang (Tab) trong hộp thoại:** Trang 1 **`Fill`** (tô nền); Trang 2 **`Stroke paint`** (chọn màu cho đường viền); Trang 3 **`Stroke style`** (chỉnh độ dày mỏng `Width` và kiểu nét đứt/nét liền ở mục `Dashes` — đây là đáp án Câu 1 tr. 71).',
        '**Các biểu tượng kiểu tô màu trong Hình 13.4:** Dấu **`X`** (không tô/trong suốt); **Ô vuông đặc** (`Flat color` – đồng nhất 1 màu); **Hai ô vuông chuyển** (`Linear gradient` – chuyển màu thẳng, ví dụ từ sáng sang tối); **Ô vuông sáng giữa tâm** (`Radial gradient` – chuyển màu toả tròn giúp hình tròn phẳng có chiều sâu khối cầu 3D); **Ô bàn cờ** (`Pattern` – tô hoa văn); Dấu **`?`** (huỷ đặt).',
        '**Độ trong suốt (Opacity và kênh A – Alpha):** Kéo giảm từ `100%` xuống `50% – 80%` thì đối tượng trở nên **trong suốt một phần** (giống tấm kính mờ), nhìn xuyên qua được các hình nằm lớp bên dưới (ứng dụng vẽ chùm bóng bay Hình 13.16).'
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
            '- *"Mời cả lớp nhìn Hình 13.2 trang 69: gồm một ngũ giác, một ngôi sao nhọn, một ngôi sao cánh tròn và một hình đa giác ngẫu nhiên. Tuy hình dáng khác nhau nhưng cả 4 hình đều được tạo ra từ CÙNG MỘT CÔNG CỤ đã học ở Bài 12. Đố cả lớp đó là công cụ nào?"*, sau đó gọi 1 HS trả lời (Công cụ vẽ đa giác / hình sao).\n'
            '- *"Đúng rồi! Bây giờ hai bạn cùng bàn đọc Bảng thuộc tính ở trang 69–70 SGK trong 2 phút và gạch chân 3 từ khoá trả lời các câu sau:*\n'
            '  *1. Muốn bo tròn 4 góc hình chữ nhật thay vì để góc vuông 90 độ, ta chỉnh hai thông số nào?* (`Rx, Ry`)\n'
            '  *2. Muốn cắt hình tròn thành nửa hình tròn như miếng dưa hấu, ta nhập Start và End bằng bao nhiêu độ?* (`0` và `180`)\n'
            '  *3. Câu hỏi trang 70 SGK: Để uốn cong các đỉnh nhọn của ngôi sao thành cánh hoa tròn, ta thay đổi tham số nào?* (`Rounded`)\n'
            '- *"Bây giờ các em nhìn lên máy chiếu để xem kết quả khi thay đổi các thông số này."*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 2):**\n'
            '  1. Nháy đúp vào **Hình chữ nhật xanh**, gõ `Rx = 25` rồi nhấn `Enter`: quan sát 4 góc bo tròn.\n'
            '  2. Nháy đúp vào **Hình tròn cam**, gõ `Start = 0`, `End = 180` rồi nhấn `Enter`: thu được nửa hình tròn.\n'
            '  3. Nháy đúp vào **Ngôi sao đỏ**, sửa `Rounded = 0.25` rồi nhấn `Enter`: các đỉnh nhọn uốn thành cánh hoa tròn.\n'
            '- **Xử lý nếu HS ngại đọc tiếng Anh:** Nhắc nhẹ *"Các em không cần lo phát âm tiếng Anh, chỉ cần đọc đúng hoặc chỉ đúng vị trí các chữ Rx, Ry, Start, End, Rounded trong SGK trang 70 là được."* rồi viết rõ lên bảng.'
        ]
    ]
    b.append(eng.table(c2a_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2B ---
    eng.add(b, 'Chặng 2B (Phút 14–23): Nhiệm vụ 2 – Màu tô, màu vẽ & Fill and Stroke (SGK tr. 70–71)', 'sub')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    c2b_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Ở Bài 12 muốn đổi màu nhanh thì nháy vào dải màu dưới đáy màn hình. Nhưng cách đó chỉ tô được một màu đơn sắc phẳng. Để tạo màu chuyển khối 3D và chỉnh nét viền, ta nhấn Shift + Ctrl + F mở hộp thoại Fill and Stroke."*\n'
            '- *"Nhìn 5 ô vuông ở Hình 13.3 trang 70, bạn nào nhận xét giúp thầy phần tô màu bên trong của 5 hình này khác nhau ở điểm nào?"* (Ô 1 rỗng, Ô 2 đặc 1 màu, Ô 3 chuyển màu thẳng ngang, Ô 4 sáng giữa tâm, Ô 5 mờ nhìn xuyên).\n'
            '- *"Các em nhớ 3 trang trong hộp thoại: Fill là màu nền bên trong, Stroke paint là màu đường viền bên ngoài, Stroke style là độ dày và kiểu nét liền hoặc nét đứt. Còn Opacity chỉnh độ trong suốt để nhìn xuyên qua lớp bên dưới."*\n'
            '- *"Để củng cố, cả lớp giơ ngón tay chọn đáp án (1 ngón = A, 2 = B, 3 = C, 4 = D) cho 2 câu trắc nghiệm SGK trang 71 nhé."*',
            '- **Thao tác Inkscape (mau_bai13.svg – Khu vực 3):**\n'
            '  + Hộp thoại `Fill and Stroke` đã mở sẵn bên phải: bấm vào **Hình tròn xanh**, chọn tab `Fill`, bấm nút **`Radial gradient`** (nút thứ 4) cho lớp thấy hiệu ứng chuyển màu toả tròn 3D.\n'
            '  + Chuyển tab `Stroke style`, bấm `Dashes` đổi viền sang nét đứt.\n'
            '  + Bấm **Hình tròn xanh dương** đè trên chữ nhật vàng, kéo thanh `Opacity` xuống khoảng `50%` cho HS thấy nhìn xuyên qua.\n'
            '- **Đáp án tr. 71:** Câu 1 giơ 3 ngón (**C. Stroke style**); Câu 2 giơ 2 ngón (**B** – biểu tượng chỉnh sửa gradient).\n'
            '- **Xử lý nếu HS ngại giơ tay:** Đếm *"1, 2, 3 cả lớp cùng giơ một lượt"* để mọi học sinh đều tham gia đồng loạt.'
        ]
    ]
    b.append(eng.table(c2b_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 2C ---
    eng.add(b, 'Chặng 2C (Phút 23–33): Nhiệm vụ 3 – Các phép ghép hình trong Path (SGK tr. 71–72)', 'sub')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    c2c_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Các em nhìn Hình 13.5 trang 71: Từ 1 hình vuông xoay nghiêng và 2 hình tròn mà ghép được thành hình trái tim cân đối. Trong Inkscape, ta làm việc này bằng các phép ghép hình trong menu Path."*\n'
            '- Chiếu Hình 13.6 (tr. 72): *"Trong môn Toán lớp 10 các em đã học các phép Tập hợp A ∪ B, A ∩ B, A \\ B. Ở Hình 13.6 quy ước hình vuông vẽ trước (lớp dưới) và hình tròn vẽ sau (lớp trên). Nhìn hình và xác định giúp thầy:"*\n'
            '  1. *"Phép nào gộp hai hình thành một khối thống nhất (Union)?"*\n'
            '  2. *"Phép nào lấy hình tròn cắt khuyết một phần của hình vuông (Difference)? Đây cũng là cách người ta khoét phần khuyết trên logo quả táo của hãng Apple."*\n'
            '  3. *"Phép nào chỉ giữ lại đúng phần giao nhau giữa hai hình (Intersection)?"*\n'
            '- *(Giới thiệu nhanh 3 phép còn lại trong H.13.6: Exclusion giữ rìa bỏ phần giao; Division chia thành các mảnh rời còn màu; Cut Path cắt đứt đường viền và bỏ màu tô)*.\n'
            '- Hỏi câu trang 72: *"Vậy muốn ghép nhiều hình tròn thành Đám mây Hình 13.7, ta chọn phép nào?"* (`Path / Union`!).',
            '- **Thao tác AN TOÀN TUYỆT ĐỐI (mau_bai13.svg – Khu vực 4):**\n'
            '  + **A (Đám mây):** Quét chọn 3 hình tròn và 1 chữ nhật xanh nhạt, **KHÔNG bấm phím tắt Ctrl + +** (dễ kẹt phím Shift trên laptop), mà mở thẳng menu **Path chọn Union** cho cả lớp đọc rõ chữ, các hình gộp lại thành Đám mây H.13.7.\n'
            '  + **B & C (Cắt Vuông-Tròn & Logo Apple):** Quét chọn hình vuông đỏ với hình tròn cam (hoặc quả táo đen với hình tròn trắng xám bên cạnh), mở menu **Path chọn Difference** để tạo vết cắt khuyết ngay trên màn chiếu.\n'
            '- **Xử lý tình huống:** Nếu bấm `Difference` mà thấy mất hình gốc, bình tĩnh bấm `Ctrl + Z` và giải thích: *"Hình nào nằm trên sẽ cắt hình nằm dưới; muốn giữ lại hình nào thì chọn hình đó rồi nhấn phím End đưa xuống lớp dưới cùng trước khi cắt."*'
        ]
    ]
    b.append(eng.table(c2c_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 3 ---
    eng.add(b, 'Chặng 3 (Phút 33–40): Luyện tập – Trả lời nhanh 3 hình SGK trang 74', 'sub')
    eng.add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    c3_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Bây giờ chúng ta vận dụng ngay để giải nhanh 3 bài ở phần Luyện tập SGK trang 74. Với mỗi hình, các em chỉ cần nêu đúng 1 hoặc 2 từ khoá vừa học:"*\n'
            '  + *"Hình 13.14 (Bông hoa 10 cánh): Dùng thông số nào để uốn tròn cánh ngôi sao 10 cánh, và vẽ thêm hình gì ở giữa làm nhụy?"* (`Rounded` + Hình tròn vàng).\n'
            '  + *"Hình 13.15 (Quân Bích đỏ): Phần trên giống hình trái tim lật ngược ta ghép bằng lệnh gì, phần chân đế khoét lõm hai bên ta cắt bằng lệnh nào?"* (`Union` ghép trên + `Difference` khoét đế + `Union`).\n'
            '  + *"Hình 13.16 (Chùm bóng bay nhìn xuyên nhau): Chỉnh thông số nào để nhìn xuyên qua các quả bóng?"* (`Opacity` khoảng 75%–80%).',
            '- **Thao tác (mau_bai13.svg – Khu vực 5):** Chỉ vào mẫu Bông hoa H.13.14 và Chùm bóng bay H.13.16 dựng sẵn ở Khu vực 5 để học sinh đối chiếu trực quan ngay.\n'
            '- Mời luân phiên 3 học sinh ở các dãy bàn khác nhau trả lời và nhận xét, chốt đáp án rõ ràng.\n'
            '- **Xử lý nếu lớp chậm:** Với Hình 13.15, gợi lại *"Lúc nãy hình trái tim Hình 13.5 ghép bằng lệnh gì trong Path?"* để học sinh nhìn bảng giữa đọc ra `Union` và `Difference`.'
        ]
    ]
    b.append(eng.table(c3_rows, [4650, 4421], header=True, cell_sz=22))

    # --- Chặng 4 ---
    eng.add(b, 'Chặng 4 (Phút 40–45): Vận dụng – Vẽ lại Miếng dưa hấu & Dặn dò Tiết 2 (SGK tr. 74)', 'sub')
    c4_rows = [
        ['LỜI DẪN & CÂU HỎI TƯƠNG TÁC', 'THAO TÁC (mau_bai13.svg) & XỬ LÝ KHI LỚP IM LẶNG'],
        [
            '- *"Quay lại phần Vận dụng ở cuối trang 74: Tóm tắt lại quy trình vẽ miếng dưa hấu Hình 13.1 trong 3 bước ngắn gọn:"*\n'
            '  1. *"Tạo nửa hình tròn vỏ xanh bằng thông số gì?"* (`Start = 0, End = 180`, chế độ quạt `Slice`).\n'
            '  2. *"Nhân đôi thu nhỏ thành cùi trắng và ruột đỏ bằng phím tắt nào?"* (`Ctrl + D` rồi kéo thu nhỏ, đổi màu).\n'
            '  3. *"Hạt dưa màu đen vẽ bằng hình gì và đặt ở lớp nào?"* (Elip nhỏ màu đen ở lớp trên cùng).\n'
            '- *(Tổng kết & Dặn dò)*: *"Cả lớp hôm nay nắm bài rất tốt. Về nhà các em đọc trước 3 nhiệm vụ thực hành trang 72–74 (Logo 3 cánh, Hình cây xanh, Hình cầu 3D) để sang Tiết 2 xuống phòng máy thực hành thuận lợi hơn. Em cảm ơn thầy và chào cả lớp!"*',
            '- **Thao tác minh hoạ (mau_bai13.svg – Khu vực 1 bên phải):** Chọn hình tròn xanh chờ sẵn bên phải Khu vực 1: nhấn `E`, chỉnh `Start = 0, End = 180`, bấm `Ctrl + D` thu nhỏ tô trắng, tiếp tục `Ctrl + D` thu nhỏ tô đỏ, thêm 2 hạt elip đen trong khoảng 25 giây để minh hoạ kết quả trọn vẹn.\n'
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
        '**[Tương tác với học sinh]:** *"Trên Hình 13.2 có 4 hình khác nhau: một ngũ giác đều, một ngôi sao nhọn, một ngôi sao cánh tròn và một hình đa giác ngẫu nhiên. Tuy hình dáng khác nhau, nhưng cả 4 hình này đều được vẽ từ cùng một công cụ đã học ở Bài 12. Bạn nào nhớ đó là công cụ nào trong ba công cụ: hình chữ nhật, hình tròn hay đa giác và hình sao?"* — Gọi 01 học sinh trả lời (*"Công cụ vẽ hình đa giác và hình sao"*).',
        '**[Lời nói trực tiếp – Giao nhiệm vụ đọc SGK]:** *"Chính xác. Để biến đổi một hình cơ bản ra nhiều dạng khác nhau như vậy, ta thay đổi các thông số trên Thanh điều khiển thuộc tính phía trên màn hình. Các em dành 2 phút cùng bạn bên cạnh xem Bảng tóm tắt thuộc tính ở cuối trang 69 và đầu trang 70 SGK, dùng bút chì gạch chân các từ khoá trả lời 3 câu hỏi sau:"*',
        '**[Tương tác với học sinh – 3 câu hỏi tra bảng SGK]:** Đọc chậm từng câu để lớp tìm trực tiếp trong bảng:\n'
        '  Câu 1: *"Muốn bo tròn 4 góc của hình chữ nhật thay vì để góc vuông 90 độ, ta thay đổi hai thông số nào?"* (**Rx, Ry**).\n'
        '  Câu 2: *"Muốn cắt hình tròn thành nửa hình tròn như miếng dưa hấu, ta nhập thông số Start và End bằng bao nhiêu độ?"* (**0** và **180**).\n'
        '  Câu 3 (Câu hỏi SGK trang 70): *"Để làm cong các đỉnh nhọn của ngôi sao thành hình cánh hoa tròn, ta điều chỉnh tham số nào?"* (**Rounded**).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 2]:** Sau khi mời lần lượt 03 học sinh đọc 3 từ khoá vừa tìm được trong sách, chuyển màn chiếu sang Khu vực 2 của mau_bai13.svg: *"Bây giờ các em quan sát trực tiếp kết quả trên màn chiếu khi thầy thay đổi đúng các thông số này."*\n'
        '  Thao tác 1: Nháy đúp chuột vào **Hình chữ nhật màu xanh**, chỉ lên thanh thuộc tính nhập Rx = 25 rồi nhấn Enter để bốn góc bo tròn lại.\n'
        '  Thao tác 2: Nháy đúp vào **Hình tròn màu cam**, nhập Start = 0, End = 180 rồi nhấn Enter để tạo nửa hình tròn; di chuột chỉ 3 nút tròn cuối thanh thuộc tính (hình quạt nối tâm, cung tròn hở nét đáy, và nút đóng kín thành hình tròn nguyên vẹn).\n'
        '  Thao tác 3: Nháy đúp vào **Ngôi sao màu đỏ**, sửa ô Rounded = 0.25 rồi nhấn Enter để các đỉnh nhọn uốn cong thành cánh hoa.',
        '**[Xử lý tình huống]:** Nếu học sinh lúng túng khi đọc tên tiếng Anh, động viên ngắn gọn: *"Các em không cần lo phát âm, chỉ cần đọc đúng mặt chữ hoặc chỉ đúng vị trí Rx, Ry, Start, End, Rounded trong sách trang 70 là đạt yêu cầu."*',
        '**[Viết bảng & Chốt ghi vở]:** Ghi tóm tắt Mục 1 lên bảng giữa và đọc cho lớp chép: *"1. Thuộc tính hình khối (SGK tr. 69–70): Hình chữ nhật có W, H (rộng, cao) và Rx, Ry (độ bo góc). Hình tròn có Start, End (góc đầu và cuối; 0 và 180 tạo nửa hình tròn). Đa giác và sao có Corners (số đỉnh hoặc cánh), Spoke ratio (tỉ lệ nhánh), Rounded (độ cong đỉnh) và Randomized (độ méo ngẫu nhiên)."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 3
    eng.add(b, 'PHẦN 3. MỤC 2 – TÔ MÀU, MÀU VẼ & ĐỘ TRONG SUỐT (PHÚT 14 – 23)', 'heading')
    eng.add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    for item in [
        '**[Lời nói trực tiếp – Liên hệ Bài 12]:** *"Chúng ta chuyển sang phần 2: Thiết lập màu tô, màu vẽ và độ trong suốt. Trước hết, một bạn nhắc lại giúp thầy ở Bài 12, khi muốn đổi màu nền và màu viền nhanh dưới dải màu đáy màn hình mình thao tác thế nào?"* — Gọi nhanh 1 học sinh nhắc lại (*"Nháy chuột đổi màu nền, giữ Shift và nháy chuột đổi màu viền"*).',
        '**[Lời nói trực tiếp]:** *"Đúng rồi. Tuy nhiên bảng màu dưới đáy chỉ chọn được một màu đơn sắc phẳng. Khi muốn tô màu chuyển sắc tạo cảm giác nổi khối 3D hoặc đổi kiểu nét đứt cho viền, ta dùng hộp thoại Fill and Stroke. Các em quan sát 5 ô vuông ở Hình 13.3 trang 70 SGK."*',
        '**[Tương tác với học sinh]:** *"Nhìn bằng mắt thường vào 5 ô vuông ở Hình 13.3, bạn nào mô tả nhanh giúp thầy phần màu bên trong của 5 ô khác nhau như thế nào?"* — Mời 01 học sinh nhận xét (*"Ô 1 không có màu; Ô 2 tô đặc một màu; Ô 3 chuyển màu từ đậm sang nhạt theo chiều ngang; Ô 4 sáng ở giữa tâm và đậm dần ra ngoài; Ô 5 mờ đi nhìn xuyên xuống dưới"*).',
        '**[Lời nói trực tiếp & Thao tác trên mau_bai13.svg – Khu vực 3]:** Chuyển sang Khu vực 3 và chỉ vào hộp thoại Fill and Stroke đang mở sẵn bên phải màn hình (nhắc lại phím tắt Shift + Ctrl + F và đối chiếu Hình 13.4):\n'
        '  Giải thích cấu trúc 3 trang: *"Hộp thoại Fill and Stroke có 3 trang rất rõ ràng: trang Fill chỉnh màu tô nền bên trong; trang Stroke paint chỉnh màu đường viền bao quanh; trang Stroke style chỉnh độ dày của nét vẽ và kiểu nét liền hay nét đứt."*\n'
        '  Thao tác tô chuyển sắc: Bấm vào **Hình tròn màu xanh** ở Khu vực 3, chọn trang Fill, nháy biểu tượng thứ tư **Radial gradient** để biến hình tròn phẳng thành khối cầu 3D có điểm sáng ở tâm; chỉ thêm nút **Linear gradient** bên cạnh để thấy kiểu chuyển màu theo phương thẳng.\n'
        '  Thao tác kiểu đường viền: Chuyển sang trang Stroke style, chỉ ô Width (độ dày) và mục Dashes chọn kiểu nét đứt để lớp thấy đường viền thay đổi ngay.\n'
        '  Thao tác độ trong suốt: Bấm vào **Hình tròn xanh dương** đang nằm đè trên hình chữ nhật vàng, kéo thanh trượt **Opacity** phía dưới từ 100% xuống khoảng 50%: *"Khi giảm Opacity hoặc kênh Alpha, hình trở nên trong suốt một phần nên ta nhìn xuyên qua được đối tượng nằm ở lớp dưới. Bạn nào từng dùng Canva hay Figma sẽ thấy tên gọi Opacity này hoàn toàn giống nhau."*',
        '**[Tương tác với học sinh – Trả lời nhanh 2 câu trắc nghiệm SGK tr. 71]:** *"Các em nhìn 2 câu hỏi củng cố ở giữa trang 71 SGK. Thầy sẽ đếm từ 1 đến 3, cả lớp cùng giơ ngón tay chọn đáp án để khỏi mất thời gian: 1 ngón là A, 2 ngón là B, 3 ngón là C, 4 ngón là D."*\n'
        '  Câu 1 (chọn kiểu nét đứt của đường viền ở trang nào): Đếm 1, 2, 3 cho lớp giơ tay và chốt đáp án **C. Stroke style** (3 ngón).\n'
        '  Câu 2 (biểu tượng nào dùng để điều chỉnh điểm dừng chuyển màu): Đếm 1, 2, 3 cho lớp giơ tay và chốt đáp án **B** (biểu tượng cây bút chỉnh sửa gradient, 2 ngón).',
        '**[Viết bảng & Chốt ghi vở]:** Ghi gọn Mục 2 lên bảng giữa: *"2. Hộp thoại Fill and Stroke (Shift + Ctrl + F): Gồm 3 trang là Fill (không tô, đồng nhất Flat color, chuyển thẳng Linear, chuyển toả tròn Radial, hoa văn Pattern), Stroke paint (màu viền), Stroke style (độ dày Width và nét đứt Dashes). Opacity và kênh Alpha (A) chỉnh độ trong suốt. Đáp án tr. 71: Câu 1 C, Câu 2 B."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 4
    eng.add(b, 'PHẦN 4. MỤC 3 – CÁC PHÉP GHÉP ĐỐI TƯỢNG TRONG MENU Path (PHÚT 23 – 33)', 'heading')
    eng.add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    eng.add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    eng.add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Chúng ta sang phần 3, cũng là phần rất quan trọng của Bài 13: Các phép ghép đối tượng đồ hoạ. Các em nhìn Hình 13.5 trang 71 trên màn chiếu: từ 2 hình tròn và 1 hình vuông xoay nghiêng mà ta ghép được thành hình trái tim cân đối chính xác thay vì tự vẽ tay. Trong Inkscape, ta thực hiện việc này bằng các lệnh trong bảng chọn Path."*',
        '**[Liên hệ Toán lớp 10 & Phân tích Hình 13.6 tr. 72]:** *"Các lệnh này tương ứng trực tiếp với các phép toán Tập hợp các em đã học trong môn Toán lớp 10: phép Hợp, phép Hiệu và phép Giao. Ở Hình 13.6 trang 72 luôn có quy ước sẵn: hình vuông vẽ trước nằm ở lớp dưới, hình tròn vẽ sau nằm đè ở lớp trên. Các em quan sát Hình 13.6 trong 2 phút và trả lời giúp thầy 2 câu sau:"*\n'
        '  Câu 1: *"Trong Hình 13.6, phép nào gộp cả hai hình thành một hình thống nhất (Union)? Phép nào lấy hình tròn cắt khuyết một phần của hình vuông (Difference)? Và phép nào chỉ giữ lại phần giao nhau ở giữa hai hình (Intersection)?"*\n'
        '  Câu 2 (Câu hỏi SGK trang 72): *"Vậy nhìn sang Hình 13.7, để ghép nhiều hình tròn chồng lên nhau thành hình Đám mây không còn đường cắt bên trong, ta chọn phép nào?"*',
        '**[Tương tác & Chuẩn hoá]:** Mời 02 học sinh phát biểu nhanh (Đám mây dùng Path / Union), sau đó giải thích ngắn gọn 3 phép còn lại ở Hình 13.6: *"Ngoài 3 phép chính vừa nêu, bảng trang 72 còn có 3 phép: Exclusion (hiệu đối xứng) giữ phần rìa ngoài nhưng bỏ phần giao ở giữa; Division (phép chia) lấy hình trên cắt hình dưới thành các mảnh rời vẫn giữ nguyên màu tô; còn Cut Path chỉ cắt đứt đường viền tại điểm giao và làm mất màu tô nền."*',
        '**[Hành động & Thao tác trên mau_bai13.svg – Khu vực 4]:** Chuyển sang Khu vực 4 để thao tác trực tiếp bằng chuột qua menu Path (không dùng phím tắt Ctrl + + để tránh lỗi kẹt Shift trên bàn phím laptop):\n'
        '  Thao tác 1 (Đám mây H.13.7): Quét chọn 4 hình màu xanh ở ô A, mở menu **Path** rồi chọn **Union** để các hình hợp nhất ngay thành hình đám mây.\n'
        '  Thao tác 2 (Cắt khuyết logo Apple): Quét chọn hình quả táo đen ở lớp dưới và hình tròn xám đè ở mép phải tại ô C, mở menu **Path** rồi chọn **Difference** để tạo vết cắt khuyết tròn đều trên thân quả táo; giải thích thêm phần chiếc lá phía trên chính là kết quả lấy phần giao nhau của 2 hình tròn bằng **Intersection**.',
        '**[Mở rộng kiến thức thực tế & AI (30 giây)]:** *"Trong thực tế, tệp đồ hoạ vectơ (.svg) rất hữu ích cho in ấn khổ lớn hay máy cắt CNC, cắt Laser bảng hiệu vì các đường Path không bị vỡ nét khi phóng to. Tương tự, khi sử dụng các công cụ AI tạo hình ảnh hiện nay, ảnh thông thường (.png) phóng to dễ bị mờ, còn khi yêu cầu AI xuất ra mã vectơ .svg thì ta có thể mở ngay vào Inkscape để tách và sửa từng đường Path như vừa thao tác."*',
        '**[Viết bảng & Chốt ghi vở]:** Nhấn mạnh quy tắc rồi đọc cho học sinh ghi Mục 3: *"Các em gạch chân thật rõ quy tắc này vào vở: Với các phép cắt như Difference, Division và Cut Path, đối tượng nằm ở LỚP DƯỚI luôn bị đối tượng ở LỚP TRÊN cắt. Nếu lỡ thao tác nhầm làm mất hình cần giữ, ta nhấn Ctrl + Z, chọn hình cần giữ rồi nhấn phím End đưa nó xuống lớp dưới cùng trước khi thực hiện lại lệnh cắt."*'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 5
    eng.add(b, 'PHẦN 5. LUYỆN TẬP NHANH 3 BÀI TẬP SGK TRANG 74 (PHÚT 33 – 40)', 'heading')
    eng.add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    for item in [
        '**[Lời nói trực tiếp]:** *"Chúng ta chuyển sang phần Luyện tập ở trang 74 SGK. Quan sát 3 hình là Hình 13.14, Hình 13.15 và Hình 13.16, các em không cần chép đề mà chỉ cần nêu đúng các từ khoá hoặc lệnh Inkscape dùng để tạo ra từng hình."*',
        '**[Tương tác với học sinh – Mời đại diện 3 dãy bàn]:**\n'
        '  Lượt 1 – Hình 13.14 (Bông hoa 10 cánh): *"Với bông hoa vẽ từ ngôi sao 10 cánh này, dãy bên trái cho thầy biết ta điều chỉnh tham số nào để uốn cong các cánh nhọn thành cánh tròn, và thêm hình gì ở giữa làm nhụy?"* — Chốt: Chỉnh **Rounded** và vẽ **Hình tròn màu vàng** đặt lên trên.\n'
        '  Lượt 2 – Hình 13.15 (Hình quân Bích màu đỏ): *"Phần tán phía trên có hình dạng giống trái tim lật ngược ta ghép bằng lệnh gì, còn phần chân đế bị khoét lõm hai bên ta cắt bằng lệnh nào?"* — Nếu lớp ngập ngừng, gợi ý lại hình trái tim Hình 13.5 ở phần trước, học sinh sẽ đọc ngay: **Union** để ghép phần trên và **Difference** (thực hiện từng cặp 2 hình) để khoét hai bên chân đế, rồi ghép chung bằng **Union**.\n'
        '  Lượt 3 – Hình 13.16 (Chùm bóng bay): *"Để các quả bóng vàng, cam, xanh nhìn xuyên thấu qua nhau ở vùng chồng lấn, ta kéo giảm thông số nào?"* — Chốt: Giảm **Opacity** xuống khoảng 75% đến 80%.',
        '**[Hành động & Thao tác trên mau_bai13.svg – Khu vực 5]:** Chuyển sang Khu vực 5 nơi đã chuẩn bị sẵn mẫu Bông hoa và Chùm bóng bay để cả lớp kiểm chứng kết quả trực quan.'
    ]:
        eng.add(b, item, 'bullet')

    # PHẦN 6
    eng.add(b, 'PHẦN 6. VẬN DỤNG & DẶN DÒ KẾT THÚC TIẾT HỌC (PHÚT 40 – 45)', 'heading')
    for item in [
        '**[Lời nói trực tiếp]:** *"Còn 5 phút cuối, chúng ta giải quyết bài Vận dụng ở cuối trang 74: Tóm tắt cách vẽ hoàn chỉnh miếng dưa hấu Hình 13.1 bằng đúng các kiến thức vừa học trong tiết này."*',
        '**[Tương tác với học sinh]:** Mời 01 hoặc 02 học sinh trình bày miệng quy trình 3 bước:\n'
        '  Bước 1: Chọn công cụ vẽ hình tròn (E), nhập **Start = 0, End = 180** ở chế độ hình quạt (Slice) và tô xanh đậm làm vỏ dưa.\n'
        '  Bước 2: Nhấn **Ctrl + D** nhân đôi, thu nhỏ kích thước và tô trắng làm cùi; tiếp tục nhấn **Ctrl + D**, thu nhỏ lần nữa và tô đỏ làm ruột.\n'
        '  Bước 3: Vẽ các hình elip nhỏ màu đen đặt lên lớp trên cùng làm hạt dưa.',
        '**[Hành động & Thao tác chốt giờ trên mau_bai13.svg – Khu vực 1 bên phải]:** *"Thầy sẽ thao tác nhanh quy trình các em vừa nêu ngay trên hình tròn chờ sẵn ở Khu vực 1."* — Chọn hình tròn xanh bên phải Khu vực 1, nhấn E, gõ Start = 0, End = 180, nhấn Ctrl + D thu nhỏ tô trắng, Ctrl + D thu nhỏ tô đỏ và thêm 2 hạt elip đen trong khoảng 25 giây (nếu sắp hết giờ thì minh hoạ bước cắt nửa hình tròn và Ctrl + D tô đỏ cho kịp giờ chuông).',
        '**[Lời nói trực tiếp – Dặn dò & Chào kết thúc]:** *"Về nhà, các em ôn lại 3 nội dung chính hôm nay và đọc trước 3 bài thực hành ở trang 72 đến 74 gồm Logo 3 cánh (Hình 13.8), Hình cây xanh (Hình 13.10) và Quả cầu 3D (Hình 13.12) để Tiết 2 xuống phòng máy thực hành thao tác cho nhanh. Cảm ơn các em đã chú ý lắng nghe và hợp tác trong tiết học hôm nay. Chào cả lớp."*',
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
        'hinh_13_14_15_16_luyen_tap.png',
    ]
    clean_and_validate_docx(b, out_path, used_figures=used_figs, is_khbd=False)


if __name__ == '__main__':
    DAY_THAT_DIR.mkdir(parents=True, exist_ok=True)
    build_khbd_day_that()
    build_kich_ban_vui_nhon()
    build_bai_thuyet_trinh_day_hoc()
    print('Done building all 3 Day_that docx files in:', DAY_THAT_DIR)
