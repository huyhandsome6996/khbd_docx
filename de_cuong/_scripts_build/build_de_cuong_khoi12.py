# -*- coding: utf-8 -*-
"""
Tao bo De cuong du gio (Kien tap su pham - KHONG co tien trinh len lop gom cac buoc cu the)
theo dung y kien gop y cua Thay Nguyen Nhat Tan va bam sat:
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
from pathlib import Path

ROOT_DIR = Path(r'c:\Users\PC\khbd_docx')
SCRIPTS_12_DIR = ROOT_DIR / 'khoi12' / '_scripts_build'
sys.path.insert(0, str(SCRIPTS_12_DIR))

import docx_engine_5512 as eng

DE_CUONG_DIR = ROOT_DIR / 'de_cuong'
BAI9_FIG_DIR = ROOT_DIR / '09_nop_du_gio_Tin12_Bai9_05-11-2026' / '_scripts_build' / 'cropped_figures'

# Copy hinh_9_*.png into eng.CROPPED_DIR and register into FIG_MAP_KHOI12 if not already present
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
        page1_time_line='Thời lượng bài học theo PPCT: 02 tiết (Tiết 19: Lí thuyết; Tiết 20: Thực hành)',
        page1_sched_line='Lịch dự giờ: Thứ Năm, ngày 05/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 12/4 – Phòng H.II.4',
        page1_mode_line='Hình thức: Đề cương trọng tâm bài học (không trình bày tiến trình lên lớp gồm các bước cụ thể)',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- [KT1]: Trình bày được cú pháp và chức năng các thẻ HTML tạo danh sách có thứ tự (<ol>, <li>), danh sách không có thứ tự (<ul>, <li>), danh sách mô tả (<dl>, <dt>, <dd>) và danh sách lồng nhau; giải thích được ý nghĩa các thuộc tính type, start.', first=0)
    eng.add(b, '- [KT2]: Trình bày được cấu trúc cơ bản của bảng trong HTML gồm các thẻ <table>, <tr>, <th>, <td>, <caption>; hiểu và vận dụng được các thuộc tính định dạng kích thước (style="width: ...; height: ..."), viền bảng (border) và gộp ô (colspan, rowspan).', first=0)
    eng.add(b, '- [KT3]: Vận dụng phối hợp thẻ tạo danh sách và thẻ tạo bảng để xây dựng trang web hoàn chỉnh giới thiệu Câu lạc bộ và Lịch hoạt động câu lạc bộ đúng yêu cầu Sách giáo khoa Tin học 12 trang 55–56.', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'italic_bold')
    eng.add(b, '- [NLC1 – Tự chủ và tự học]: Chủ động đọc Sách giáo khoa, phân tích ví dụ mã nguồn HTML và tự kiểm chứng kết quả hiển thị trên trình duyệt web.', first=0)
    eng.add(b, '- [NLC2 – Giao tiếp và hợp tác]: Thảo luận cặp đôi hoặc nhóm nhỏ để nhận diện cấu trúc thẻ danh sách, bảng biểu và hỗ trợ nhau sửa lỗi cú pháp.', first=0)
    eng.add(b, '- [NLC3 – Giải quyết vấn đề và sáng tạo]: Lựa chọn cấu trúc danh sách hoặc bảng hợp lí, thiết kế bảng có gộp ô (rowspan, colspan) rõ ràng, thẩm mĩ cho trang web thực tế.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'italic_bold')
    eng.add(b, '- [NLa]: Sử dụng và quản lí hiệu quả trình soạn thảo mã nguồn HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web.', first=0)
    eng.add(b, '- [NLc]: Phân tích cấu trúc dữ liệu dạng danh sách, bảng biểu và viết đúng mã HTML chuẩn cấu trúc lồng thẻ.', first=0)
    eng.add(b, '- [NLd]: Lưu trữ, kiểm tra sản phẩm số và chia sẻ tệp mã nguồn HTML đúng quy chuẩn thư mục học tập.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'italic_bold')
    eng.add(b, '- [NLS 3.1.NC1.a]: Tạo lập, chỉnh sửa và định dạng nội dung số dạng danh sách và bảng biểu trên trang web bằng ngôn ngữ HTML.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- [PC1 – Chăm chỉ]: Kiên trì thực hành viết mã HTML, cẩn trọng kiểm tra cặp thẻ đóng – mở và các thuộc tính bảng.', first=0)
    eng.add(b, '- [PC2 – Trách nhiệm]: Chấp hành nghiêm túc nội quy phòng học/phòng máy H.II.4, bảo quản thiết bị và hoàn thành sản phẩm thực hành đúng hạn.', first=0)
    eng.add(b, '- [PC3 – Trung thực]: Tự tay biên tập mã nguồn HTML và tự đánh giá trung thực mức độ hoàn thành sản phẩm.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Hệ thống máy tính giáo viên, máy chiếu/màn hình lớn tại phòng học/phòng máy H.II.4 (Lớp 12/4) cài đặt trình soạn thảo HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web (Chrome / Edge).', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam, Bài 9: Tạo danh sách, bảng (trang 52–56).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 19: Lí thuyết; Tiết PPCT 20: Thực hành).', first=0)
    eng.add(b, '- Hệ thống hình ảnh minh họa SGK (Hình 9.1 đến Hình 9.8), Phiếu học tập và tệp mã nguồn HTML mẫu (CLB.html).', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM BÀI HỌC VÀ ĐỊNH HƯỚNG PHƯƠNG PHÁP DẠY HỌC', 'heading')
    eng.add(b, '1. Định hướng phương pháp và hình thức tổ chức dạy học (theo Công văn 5512)', 'sub')
    eng.add(b, 'Khác với giáo án chi tiết từng bước lên lớp, đề cương hệ thống hóa 04 hoạt động dạy học chủ đạo theo đúng định hướng phát triển năng lực học sinh:', first=0)
    b.append(eng.table([
        ['HOẠT ĐỘNG / THỜI LƯỢNG', 'MỤC TIÊU ĐÁP ỨNG', 'PHƯƠNG PHÁP & KĨ THUẬT DẠY HỌC', 'HÌNH THỨC TỔ CHỨC & ĐÁNH GIÁ'],
        ['1. Mở đầu (05 phút)', '[KT1], [NLC1], [NLa]', 'Nêu vấn đề, trực quan qua trang web thực tế (so sánh danh sách và bảng biểu).', 'Cá nhân / Cặp đôi; đánh giá qua câu trả lời miệng (Bảng kiểm).'],
        ['2. Hình thành kiến thức mới (52 phút – Gồm danh sách & bảng HTML)', '[KT1], [KT2], [NLC1], [NLC2], [NLc], [NLS 3.1.NC1.a]', 'Dạy học hợp tác theo nhóm đôi, trực quan phân tích mã nguồn SGK (Hình 9.1 – 9.6), vấn đáp gợi mở.', 'Thảo luận cặp đôi trên Phiếu học tập số 1–4 kết hợp đối chiếu mẫu trình chiếu.'],
        ['3. Luyện tập – Thực hành (25 phút – Nhiệm vụ 1 & 2 trang 55–56)', '[KT1], [KT2], [KT3], [NLa], [NLc], [PC1], [PC2]', 'Thực hành theo mẫu hướng dẫn (Hình 9.7, 9.8), học qua làm (Learning by doing).', 'Thực hành cá nhân/cặp đôi trên máy tính tạo tệp CLB.html; chấm điểm theo Rubric.'],
        ['4. Vận dụng (08 phút)', '[KT3], [NLC3], [NLc], [NLd], [PC3]', 'Giao nhiệm vụ gắn với thực tiễn (bổ sung thông tin hoặc thiết kế thời khóa biểu/bảng điểm).', 'Hoàn thiện mã nguồn và trình bày kết quả trên trình duyệt web.']
    ], [2100, 1700, 2900, 2650], header=True, cell_sz=22))
    eng.br(b)

    eng.add(b, '2. Nội dung kiến thức trọng tâm bài học (Bám sát SGK Tin học 12 trang 52–56)', 'sub')
    eng.add(b, 'a) Tạo danh sách trong HTML (Mục 1 – SGK trang 52–53)', 'italic_bold')
    eng.add(b, '- Danh sách có thứ tự (Ordered List): Sử dụng cặp thẻ <ol>...</ol> bao ngoài và <li>...</li> cho từng mục. Thuộc tính type quy định kiểu đánh số ("1", "A", "a", "I", "i") và thuộc tính start quy định giá trị bắt đầu.', first=0)
    eng.add(b, '- Danh sách không có thứ tự (Unordered List): Sử dụng cặp thẻ <ul>...</ul> và <li>...</li>. Thuộc tính type quy định kí hiệu đầu dòng ("disc", "circle", "square") hoặc dùng CSS list-style-type.', first=0)
    eng.add(b, '- Danh sách mô tả (Description List): Sử dụng <dl>...</dl>, trong đó <dt>...</dt> chứa thuật ngữ cần mô tả và <dd>...</dd> chứa nội dung mô tả tương ứng.', first=0)
    eng.add(b, '- Danh sách lồng nhau: Cho phép đặt một phần tử <ol> hoặc <ul> bên trong một phần tử <li> của danh sách cha để biểu diễn dữ liệu phân cấp nhiều mức.', first=0)
    eng.add_figure(b, 'hinh_9_1.png')
    eng.add_figure(b, 'hinh_9_2.png')
    eng.add_figure(b, 'hinh_9_3.png')

    eng.add(b, 'b) Thiết lập bảng trong HTML (Mục 2 – SGK trang 54–55)', 'italic_bold')
    eng.add(b, '- Cấu trúc bảng cơ bản: Sử dụng cặp thẻ <table>...</table>; mỗi hàng khai báo bằng <tr>...</tr>; ô tiêu đề cột/hàng dùng <th>...</th> (chữ đậm, căn giữa); ô dữ liệu dùng <td>...</td>; tiêu đề bảng dùng <caption>...</caption>.', first=0)
    eng.add(b, '- Định dạng kích thước và viền bảng: Thuộc tính border="1" (hoặc CSS border) để hiển thị khung viền; thuộc tính style="width: ...; height: ..." để chỉnh độ rộng, chiều cao của bảng, hàng hoặc từng ô.', first=0)
    eng.add(b, '- Gộp ô trong bảng: Thuộc tính colspan="n" dùng để gộp n ô theo chiều ngang (trên cùng một hàng); thuộc tính rowspan="m" dùng để gộp m ô theo chiều dọc (trên m hàng liên tiếp). Cần xóa bớt thẻ <td>/<th> bị gộp để bảng không bị lệch cột.', first=0)
    eng.add_figure(b, 'hinh_9_4.png')
    eng.add_figure(b, 'hinh_9_5.png')
    eng.add_figure(b, 'hinh_9_6.png')

    eng.add(b, 'c) Nội dung thực hành trọng tâm (Mục 3 – SGK trang 55–56)', 'italic_bold')
    eng.add(b, '- Nhiệm vụ 1: Tạo danh sách lồng nhau liệt kê các Câu lạc bộ của trường (CLB Thể thao gồm Bóng đá, Bóng chuyền, Bóng rổ, Võ; CLB Nghệ thuật gồm Mĩ thuật, Nhiếp ảnh, Âm nhạc, Khiêu vũ) theo mẫu Hình 9.7.', first=0)
    eng.add(b, '- Nhiệm vụ 2: Tạo bảng Lịch hoạt động CLB Thể thao (lưu vào tệp CLB.html) có tiêu đề bảng, hàng tiêu đề gộp cột/hàng phù hợp cho các môn thể thao và thời gian sinh hoạt từ Thứ 2 đến Thứ 6 theo mẫu Hình 9.8.', first=0)
    eng.add_figure(b, 'hinh_9_7.png')
    eng.add_figure(b, 'hinh_9_8.png')

    eng.add(b, 'IV. HỆ THỐNG CÂU HỎI CỦNG CỐ VÀ TIÊU CHÍ ĐÁNH GIÁ', 'heading')
    eng.add(b, '- Câu hỏi củng cố 1 (SGK trang 53): Phân biệt cặp thẻ tạo danh sách có thứ tự (<ol>) và danh sách không có thứ tự (<ul>); viết đoạn mã HTML tạo danh sách lồng nhau gồm 2 cấp.', first=0)
    eng.add(b, '- Câu hỏi củng cố 2 (SGK trang 55): Trình bày sự khác nhau giữa thẻ <th> và <td>; giải thích cách tính số thẻ <td> cần giữ lại khi sử dụng thuộc tính colspan="3" hoặc rowspan="2".', first=0)
    eng.add(b, '- Bài tập Luyện tập & Vận dụng (SGK trang 56): Chuyển đổi danh sách bảng từ dạng chuẩn sang bảng có gộp ô theo hàng/cột và bổ sung nội dung giới thiệu CLB vào tệp CLB.html chuẩn bị cho Bài 10 (Tạo liên kết).', first=0)
    eng.add(b, '- Tiêu chí đánh giá sản phẩm thực hành (Thang điểm 10): Đúng cấu trúc HTML và lồng thẻ (3,0 điểm); Tạo đúng danh sách lồng nhau theo Hình 9.7 (2,5 điểm); Tạo đúng bảng lịch hoạt động có gộp ô chuẩn theo Hình 9.8 (3,0 điểm); Trình bày mã nguồn sạch đẹp, lưu tệp đúng quy định (1,5 điểm).', first=0)

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
    eng.save_docx(b, out_path, used_figures=list(BAI9_FIGS.keys()))


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
        page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 12/2 – Phòng H.II.2',
        page1_mode_line='Hình thức: Đề cương trọng tâm bài học (không trình bày tiến trình lên lớp gồm các bước cụ thể)',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- [KT1]: Trình bày được khái niệm siêu văn bản (hypertext), siêu liên kết (hyperlink), cú pháp thẻ <a> với thuộc tính href và phân biệt rõ đường dẫn tuyệt đối với đường dẫn tương đối.', first=0)
    eng.add(b, '- [KT2]: Phân biệt được liên kết ngoài (tới trang web khác trên Internet) và liên kết trong (tới trang web khác trong cùng website); viết đúng đường dẫn tương đối cho 4 trường hợp cấu trúc thư mục theo Hình 10.1 (cùng thư mục, dưới 1 cấp, dưới nhiều cấp và ngược lên thư mục mức trên bằng ../).', first=0)
    eng.add(b, '- [KT3]: Trình bày được quy trình 2 bước tạo liên kết tới một vị trí khác trong cùng trang web bằng thuộc tính id và cú pháp #mã_định_danh_đoạn, cũng như cách tạo liên kết tới tệp hình ảnh.', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'italic_bold')
    eng.add(b, '- [NLC1 – Tự chủ và tự học]: Học sinh tự đọc Sách giáo khoa Tin học 12 (trang 57–60), phân tích sơ đồ cây thư mục Hình 10.1 để tự rút ra quy tắc xác định đường dẫn tương đối.', first=0)
    eng.add(b, '- [NLC2 – Giao tiếp và hợp tác]: Thảo luận cặp đôi, trao đổi kết quả phân tích đường dẫn trên Phiếu học tập và phản biện lẫn nhau.', first=0)
    eng.add(b, '- [NLC3 – Giải quyết vấn đề và sáng tạo]: Giải quyết chính xác bài toán định vị đường dẫn khi chuyển đổi cấu trúc thư mục website và thiết kế trước các mã định danh id phục vụ Tiết 2 Thực hành.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'italic_bold')
    eng.add(b, '- [NLb]: Giải thích được bản chất địa chỉ URL, đường dẫn tuyệt đối và đường dẫn tương đối trong môi trường web.', first=0)
    eng.add(b, '- [NLc]: Viết đúng cú pháp thẻ <a> liên kết ngoài, liên kết trong cùng website, liên kết nội bộ trong trang (#id) và liên kết hình ảnh.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'italic_bold')
    eng.add(b, '- [NLS 3.1.NC1.a]: Nắm vững kĩ thuật thiết kế và liên kết tài nguyên siêu văn bản trên môi trường số chuẩn xác.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- [PC1 – Chăm chỉ]: Tích cực nghiên cứu Sách giáo khoa và hoàn thành đầy đủ các câu hỏi củng cố lí thuyết.', first=0)
    eng.add(b, '- [PC2 – Trách nhiệm]: Chuẩn bị kĩ phương án thiết kế liên kết cho tệp CLB.html để phục vụ giờ thực hành Tiết 2.', first=0)
    eng.add(b, '- [PC3 – Trung thực]: Tự kiểm tra, đối chiếu kết quả bài làm với đáp án chuẩn.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Máy tính giáo viên kết nối máy chiếu/màn hình tại phòng học H.II.2 (Lớp 12/2), bảng viết, phấn.', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam, Bài 10: Tạo liên kết (Mục 1 và Mục 2, trang 57–60).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 21: Bài 4.10: Tạo liên kết – Tiết 1 - LT).', first=0)
    eng.add(b, '- Hình ảnh Hình 10.1 cắt từ SGK trang 59, Phiếu học tập số 1, 2, 3, 4 và tệp minh họa CLB.html.', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM BÀI HỌC VÀ ĐỊNH HƯỚNG PHƯƠNG PHÁP DẠY HỌC', 'heading')
    eng.add(b, '1. Định hướng phương pháp và hình thức tổ chức dạy học (theo Công văn 5512)', 'sub')
    b.append(eng.table([
        ['HOẠT ĐỘNG / THỜI LƯỢNG', 'MỤC TIÊU ĐÁP ỨNG', 'PHƯƠNG PHÁP & KĨ THUẬT DẠY HỌC', 'HÌNH THỨC TỔ CHỨC & ĐÁNH GIÁ'],
        ['1. Mở đầu (05 phút)', '[KT1], [NLC1], [PC1]', 'Trực quan so sánh báo điện tử và sách báo giấy; nêu vấn đề (Khởi động trang 57).', 'Cặp đôi trả lời nhanh; đánh giá bằng Bảng kiểm.'],
        ['2. Hình thành kiến thức mới (28 phút – Mục 1 & Mục 2 SGK)', '[KT1], [KT2], [KT3], [NLC1], [NLC2], [NLb], [NLc], [NLS 3.1.NC1.a]', 'Dạy học giải quyết vấn đề, thảo luận cặp đôi trên Phiếu học tập số 1–3, phân tích sơ đồ cây thư mục Hình 10.1.', 'Nhóm đôi nghiên cứu SGK trang 57–60; chữa bài tập mẫu trên bảng.'],
        ['3. Luyện tập (07 phút)', '[KT1], [KT2], [KT3], [NLC3], [NLc], [PC3]', 'Luyện tập viết câu lệnh HTML ngắn trên giấy/phiếu cho các tình huống đường dẫn.', 'Cá nhân hoàn thành Phiếu học tập số 4; giáo viên chấm chữa nhanh.'],
        ['4. Vận dụng – Chuẩn bị Tiết 2 (05 phút)', '[KT3], [NLC3], [NLc], [PC2]', 'Hướng dẫn học sinh lập bảng thiết kế id và thẻ <a> cho tệp CLB.html (Hình 10.2).', 'Cá nhân ghi thiết kế vào vở để thực hành ngay ở Tiết 2.']
    ], [2100, 1700, 2900, 2650], header=True, cell_sz=22))
    eng.br(b)

    eng.add(b, '2. Nội dung kiến thức trọng tâm bài học (Bám sát SGK Tin học 12 trang 57–60)', 'sub')
    eng.add(b, 'a) Siêu văn bản và đường dẫn (Mục 1 – SGK trang 57–58)', 'italic_bold')
    eng.add(b, '- Tình huống mở đầu & Hoạt động 1 (SGK trang 57): Trang web là siêu văn bản vì ngoài văn bản thuần túy còn tích hợp hình ảnh, âm thanh, video và đặc biệt là các siêu liên kết giúp truy cập thông tin phi tuyến tính (khác với tệp .txt chỉ chứa văn bản thuần túy).', first=0)
    eng.add(b, '- Khái niệm siêu liên kết và cú pháp thẻ <a>:', first=0)
    eng.add(b, '  + Siêu liên kết (hyperlink) là một tham chiếu đến một tài liệu nằm tại một vị trí khác (trong cùng trang web, cùng website hoặc thuộc website khác).', first=0)
    eng.add(b, '  + Cú pháp: <a href="URL">Nội dung hiển thị tại vị trí đặt liên kết</a> (mặc định có màu xanh và gạch chân).', first=0)
    eng.add(b, '- Phân biệt hai loại đường dẫn trong thuộc tính href:', first=0)
    eng.add(b, '  + Đường dẫn tuyệt đối: Ghi đầy đủ giao thức (http:// hoặc https://), tên miền và đường dẫn chi tiết. Ví dụ: https://google.com.', first=0)
    eng.add(b, '  + Đường dẫn tương đối: Không có giao thức hay tên miền, chỉ ghi đường dẫn tính từ vị trí tệp hiện tại. Ví dụ: html/cach_tao_lien_ket.html.', first=0)

    eng.add(b, 'b) Các cách tạo liên kết tới một trang web (Mục 2 – SGK trang 58–60)', 'italic_bold')
    eng.add(b, '- Liên kết tới một trang web khác (Liên kết ngoài): Sử dụng đường dẫn tuyệt đối. Ví dụ: <a href="https://hanhtrangso.nxbgd.vn/sach-dien-tu?book_active=0">Sách điện tử KNTT</a>.', first=0)
    eng.add(b, '- Liên kết đến một vị trí khác trong cùng website (Liên kết trong – Hình 10.1, trang 59–60):', first=0)
    eng.add(b, '  + Trường hợp 1 (Cùng thư mục): <a href="thong_tin.html">Giới thiệu về trang web</a>', first=0)
    eng.add(b, '  + Trường hợp 2 (Thư mục khác, dưới 1 cấp): <a href="bai_tap/bai_tap_1.html">Bài tập 1</a>', first=0)
    eng.add(b, '  + Trường hợp 3 (Thư mục khác, dưới 2 cấp): <a href="bai_tap/on_tap/bai_tap_on_tap.html">Bài tập ôn tập</a>', first=0)
    eng.add(b, '  + Trường hợp 4 (Thư mục mức trên): <a href="../index.html">Quay lại trang chủ</a> (mỗi "../" lùi lên 1 cấp thư mục mẹ).', first=0)
    eng.add_figure(b, 'hinh_10_1.png')
    eng.add(b, '- Liên kết tới vị trí khác trong cùng trang web (Liên kết nội bộ):', first=0)
    eng.add(b, '  + Bước 1: Gắn thuộc tính id="tên_định_danh" cho phần tử đích (tên id viết liền không dấu, không trùng lặp). Ví dụ: <h2 id="CLBTT">Lịch hoạt động</h2>.', first=0)
    eng.add(b, '  + Bước 2: Đặt thẻ liên kết với href="#tên_định_danh". Ví dụ: <a href="#CLBTT">Xem lịch hoạt động</a>.', first=0)
    eng.add(b, '- Liên kết tới hình ảnh: Đặt đường dẫn tệp ảnh trong href (ví dụ <a href="images/sun.png">Mặt trời</a>) hoặc lồng thẻ <img> vào giữa cặp thẻ <a>...</a>.', first=0)

    eng.add(b, 'IV. HỆ THỐNG CÂU HỎI CỦNG CỐ VÀ TIÊU CHÍ ĐÁNH GIÁ', 'heading')
    eng.add(b, '- Câu hỏi củng cố 1 (SGK trang 58): Phân loại đường dẫn tuyệt đối và tương đối cho 3 trường hợp: a) html/cach_tao_lien_ket.html (tương đối); b) http://google.com (tuyệt đối); c) mail.google.com/mail/u/0/#inbox/FMfcgzGrcjHzMWKllZqRrT (tương đối theo cú pháp HTML do thiếu giao thức https://).', first=0)
    eng.add(b, '- Câu hỏi củng cố 2 (SGK trang 60): Dựa theo Hình 10.1, viết câu lệnh liên kết từ bai_tap_on_tap.html (ở mức sâu 2 cấp trong bai_tap/on_tap/) về trang chủ index.html: <a href="../../index.html">Về trang chủ</a>.', first=0)
    eng.add(b, '- Bài tập Luyện tập trên lớp: Viết mã liên kết tới hình ảnh sun.png và liên kết nội bộ tới đoạn có id="gioi_thieu".', first=0)
    eng.add(b, '- Định hướng đánh giá Tiết 1: Đánh giá qua Phiếu học tập số 1, 2, 3, 4 (chính xác cú pháp thẻ <a>, phân biệt đúng 4 trường hợp đường dẫn tương đối ở Hình 10.1 và quy trình 2 bước gắn id).', first=0)

    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 12 tháng 11 năm 2026',
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
    )

    out_path = DE_CUONG_DIR / 'de_cuong_du_gio_12-2.docx'
    eng.save_docx(b, out_path, used_figures=['hinh_10_1.png'])


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
        page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Tiết 2 (08h00 – 08h45) – Lớp 12/1 – Phòng H.II.1',
        page1_mode_line='Hình thức: Đề cương trọng tâm bài học (không trình bày tiến trình lên lớp gồm các bước cụ thể)',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    eng.add(b, 'I. MỤC TIÊU BÀI HỌC (THEO CÔNG VĂN 5512 & THÔNG TƯ 02/2025)', 'heading')
    eng.add(b, '1. Về kiến thức', 'sub')
    eng.add(b, '- [KT1]: Vận dụng thành thạo kĩ thuật gắn thuộc tính id và đặt liên kết nội bộ trong cùng trang web (#CLBTT) trên tệp mã nguồn CLB.html theo đúng yêu cầu Nhiệm vụ 1 (SGK trang 61).', first=0)
    eng.add(b, '- [KT2]: Tạo lập được tệp trang web mới thong_tin.html nằm cùng thư mục với CLB.html và thiết lập liên kết hai chiều giữa hai trang web bằng đường dẫn tương đối theo Nhiệm vụ 2 và Bài Luyện tập 1 (SGK trang 61).', first=0)
    eng.add(b, '- [KT3]: Hoàn thiện trọn vẹn trang web đa liên kết cho Bài Luyện tập 2 (gioi_thieu_ban_than.html) và phần Vận dụng (gắn id riêng cho 5 đoạn giới thiệu câu lạc bộ dưới bảng trong tệp CLB.html và tạo liên kết từ tên từng CLB trong bảng cuộn xuống đoạn giới thiệu tương ứng).', first=0)

    eng.add(b, '2. Về năng lực', 'sub')
    eng.add(b, 'a) Năng lực chung', 'italic_bold')
    eng.add(b, '- [NLC1 – Tự chủ và tự học]: Học sinh tự đọc hướng dẫn thực hành SGK trang 61, viết mã và tự kiểm thử hoạt động của liên kết trên trình duyệt web.', first=0)
    eng.add(b, '- [NLC2 – Giao tiếp và hợp tác]: Trao đổi, hỗ trợ bạn cùng máy phát hiện và khắc phục lỗi sai đường dẫn hoặc đặt tên id không khớp.', first=0)
    eng.add(b, '- [NLC3 – Giải quyết vấn đề và sáng tạo]: Tổ chức bộ tệp trang web khoa học, thiết kế nội dung giới thiệu sinh động và điều hướng liên kết mượt mà.', first=0)

    eng.add(b, 'b) Năng lực đặc thù môn Tin học', 'italic_bold')
    eng.add(b, '- [NLa]: Sử dụng thành thạo trình soạn thảo mã nguồn HTML (Sublime Text / VS Code / Notepad++) và công cụ kiểm tra liên kết trên trình duyệt web.', first=0)
    eng.add(b, '- [NLc]: Viết mã nguồn HTML chuẩn xác cho liên kết nội bộ (#id) và liên kết tương đối giữa các tệp cùng thư mục.', first=0)
    eng.add(b, '- [NLd]: Lưu trữ, quản lí cấu trúc thư mục thực hành Bai10_TaoLienKet sạch sẽ, đúng quy định.', first=0)

    eng.add(b, 'c) Năng lực số (Thông tư 02/2025/TT-BGDĐT)', 'italic_bold')
    eng.add(b, '- [NLS 3.1.NC1.a]: Tạo lập, biên tập và liên kết thành công bộ trang web đa trang (.html) theo đúng yêu cầu kĩ thuật.', first=0)

    eng.add(b, '3. Về phẩm chất', 'sub')
    eng.add(b, '- [PC1 – Chăm chỉ]: Kiên trì biên tập đầy đủ nội dung giới thiệu, danh sách thành tích và thông tin chi tiết của từng câu lạc bộ.', first=0)
    eng.add(b, '- [PC2 – Trách nhiệm]: Chấp hành nghiêm túc nội quy phòng học/phòng máy H.II.1, lưu trữ thư mục bài làm đúng quy định.', first=0)
    eng.add(b, '- [PC3 – Trung thực]: Tự tay viết mã nguồn HTML và tự đánh giá trung thực sản phẩm theo Rubric.', first=0)

    eng.add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    eng.add(b, '1. Thiết bị dạy học và điều kiện tổ chức', 'sub')
    eng.add(b, '- Máy tính tại phòng học/phòng máy H.II.1 (Lớp 12/1) có cài đặt trình soạn thảo văn bản HTML (Sublime Text / VS Code / Notepad++) và trình duyệt web (Chrome / Edge), máy chiếu.', first=0)
    eng.add(b, '2. Học liệu', 'sub')
    eng.add(b, '- Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam, Bài 10: Tạo liên kết (Mục 3. Thực hành tạo liên kết, Luyện tập và Vận dụng – trang 61).', first=0)
    eng.add(b, '- Phân phối chương trình Khối 12 Trường THPT Thuận Hoá (Tiết PPCT 22: Bài 4.10: Tạo liên kết – Tiết 2 - TH).', first=0)
    eng.add(b, '- Thư mục học liệu chứa tệp gốc CLB.html kế thừa từ Bài 9, Hình 10.2 (SGK trang 61) và Bảng tiêu chí đánh giá thực hành (Rubric).', first=0)

    eng.add(b, 'III. NỘI DUNG TRỌNG TÂM BÀI HỌC VÀ ĐỊNH HƯỚNG PHƯƠNG PHÁP DẠY HỌC', 'heading')
    eng.add(b, '1. Định hướng phương pháp và hình thức tổ chức dạy học (theo Công văn 5512)', 'sub')
    b.append(eng.table([
        ['HOẠT ĐỘNG / THỜI LƯỢNG', 'MỤC TIÊU ĐÁP ỨNG', 'PHƯƠNG PHÁP & KĨ THUẬT DẠY HỌC', 'HÌNH THỨC TỔ CHỨC & ĐÁNH GIÁ'],
        ['1. Mở đầu – Ôn tập cú pháp (05 phút)', '[KT1], [KT2], [NLC1], [NLa]', 'Vấn đáp nhanh tái hiện cú pháp #id và liên kết tương đối; khởi động thư mục thực hành Bai10_TaoLienKet.', 'Cá nhân mở tệp CLB.html và trả lời 03 câu hỏi nhanh.'],
        ['2. Thực hành Nhiệm vụ 1 & Nhiệm vụ 2 (24 phút – SGK trang 61)', '[KT1], [KT2], [NLC2], [NLc], [NLS 3.1.NC1.a], [PC1]', 'Thực hành theo mẫu trực quan (Hình 10.2), giáo viên thị phạm lỗi thường gặp và hỗ trợ trực tiếp tại máy.', 'Cá nhân/cặp đôi thực hành trên máy tính tạo liên kết nội bộ và liên kết tới thong_tin.html.'],
        ['3. Luyện tập (08 phút – Bài 1 & 2 trang 61)', '[KT2], [KT3], [NLC3], [NLc], [PC3]', 'Luyện tập mở rộng trang web đa trang, tạo liên kết hai chiều và trang cá nhân.', 'Thực hành tạo liên kết quay về từ thong_tin.html và tạo gioi_thieu_ban_than.html.'],
        ['4. Vận dụng (08 phút – SGK trang 61)', '[KT3], [NLC3], [NLc], [NLd], [PC2]', 'Giải quyết bài toán thực tế hoàn chỉnh trang CLB.html với 5 liên kết nội bộ từ bảng.', 'Hoàn thiện sản phẩm và đánh giá theo Rubric 10 điểm.']
    ], [2100, 1700, 2900, 2650], header=True, cell_sz=22))
    eng.br(b)

    eng.add(b, '2. Nội dung thực hành trọng tâm (Bám sát SGK Tin học 12 trang 61)', 'sub')
    eng.add(b, 'a) Nhiệm vụ 1: Tạo liên kết trong cùng trang (Hình 10.2 – SGK trang 61)', 'italic_bold')
    eng.add(b, '- Bổ sung vào đầu tệp CLB.html phần giới thiệu các hoạt động ngoại khoá của trường gồm 2 đoạn văn và hình ảnh minh họa theo đúng bố cục Hình 10.2.', first=0)
    eng.add(b, '- Gắn thuộc tính id="CLBTT" cho phần tiêu đề hoặc thẻ <table> của bảng Lịch hoạt động CLB Thể thao.', first=0)
    eng.add(b, '- Đặt liên kết tại cụm từ "Theo dõi lịch hoạt động các CLB Thể thao" bằng câu lệnh: <a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a> để khi nháy chuột trình duyệt tự động cuộn xuống bảng lịch.', first=0)
    eng.add_figure(b, 'hinh_10_2.png')

    eng.add(b, 'b) Nhiệm vụ 2: Tạo liên kết sang trang khác (SGK trang 61)', 'italic_bold')
    eng.add(b, '- Tạo mới tệp thong_tin.html trong cùng thư mục chứa CLB.html, biên tập nội dung giới thiệu chi tiết các câu lạc bộ (Bóng đá, Bóng chuyền, Bóng rổ, Âm nhạc, Mĩ thuật...) kèm địa chỉ liên hệ và danh sách thành tích nổi bật.', first=0)
    eng.add(b, '- Tại vị trí thích hợp trong tệp CLB.html, thêm dòng liên kết bằng đường dẫn tương đối: <p><a href="thong_tin.html">Xem thông tin chi tiết các câu lạc bộ</a></p>.', first=0)

    eng.add(b, 'c) Nội dung Luyện tập và Vận dụng (SGK trang 61)', 'italic_bold')
    eng.add(b, '- Luyện tập 1: Bổ sung liên kết quay lại trang chính CLB.html ở cuối tệp thong_tin.html (<a href="CLB.html">Quay lại trang Câu lạc bộ</a>) để tạo điều hướng hai chiều hoàn chỉnh.', first=0)
    eng.add(b, '- Luyện tập 2: Tạo mới tệp gioi_thieu_ban_than.html (giới thiệu họ tên, lớp 12/1, sở thích, định hướng nghề nghiệp) và đặt liên kết từ trang này tới website yêu thích bằng đường dẫn tuyệt đối.', first=0)
    eng.add(b, '- Vận dụng: Ở cuối tệp CLB.html (phía dưới bảng), bổ sung 05 đoạn văn bản giới thiệu chi tiết cho 05 câu lạc bộ và gắn mã định danh riêng (id="bongda", id="bongchuyen", id="bongro", id="vo", id="boiloi"); trong bảng lịch hoạt động, biến tên mỗi câu lạc bộ thành thẻ <a> dẫn tới id tương ứng.', first=0)

    eng.add(b, 'IV. HỆ THỐNG CÂU HỎI KIỂM TRA VÀ TIÊU CHÍ ĐÁNH GIÁ THỰC HÀNH', 'heading')
    eng.add(b, '- Câu hỏi kiểm tra nhanh đầu giờ: Phân biệt chức năng của id="CLBTT" và href="#CLBTT"; giải thích vì sao tệp thong_tin.html phải lưu cùng thư mục với CLB.html khi viết href="thong_tin.html".', first=0)
    eng.add(b, '- Tiêu chí đánh giá sản phẩm thực hành trên máy tính (Rubric 10 điểm):', first=0)
    b.append(eng.table([
        ['TIÊU CHÍ ĐÁNH GIÁ', 'MÔ TẢ YÊU CẦU ĐẠT', 'ĐIỂM TỐI ĐA'],
        ['1. Nhiệm vụ 1 (Hình 10.2)', 'Hiển thị đủ phần giới thiệu đầu trang CLB.html; nháy vào liên kết cuộn đúng tới bảng có id="CLBTT".', '3,0 điểm'],
        ['2. Nhiệm vụ 2 (Trang mới)', 'Tạo tệp thong_tin.html cùng thư mục, đủ nội dung; nháy liên kết từ CLB.html mở đúng trang mới.', '2,5 điểm'],
        ['3. Luyện tập 1 & 2', 'Có liên kết quay về từ thong_tin.html; tệp gioi_thieu_ban_than.html chạy đúng liên kết ngoài.', '2,0 điểm'],
        ['4. Vận dụng & Tác phong', 'Gắn đủ 5 id và 5 thẻ <a> trong bảng CLB.html hoạt động chính xác; lưu trữ thư mục khoa học.', '2,5 điểm']
    ], [2500, 5200, 1650], header=True, cell_sz=22))
    eng.br(b)

    eng.add_approval_section(
        b,
        sign_date_str='Huế, ngày 12 tháng 11 năm 2026',
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_title='SINH VIÊN THỰC HIỆN',
        sv_name=SV_NAME,
        section_title='V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ KÝ DUYỆT'
    )

    out_path = DE_CUONG_DIR / 'de_cuong_du_gio_12-1.docx'
    eng.save_docx(b, out_path, used_figures=['hinh_10_2.png'])


if __name__ == '__main__':
    DE_CUONG_DIR.mkdir(parents=True, exist_ok=True)
    build_12_4_bai9()
    build_12_2_bai10_tiet1()
    build_12_1_bai10_tiet2()
    print('Successfully built all 3 De cuong files in:', DE_CUONG_DIR)
