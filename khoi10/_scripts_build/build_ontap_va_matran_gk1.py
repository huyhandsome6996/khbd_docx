# -*- coding: utf-8 -*-
"""
Builds:
1. 01_De_cuong_du_gio_Tin10_OnTap_GHKI_TietPPCT15_Lop10-2.docx
2. 02_Ma_tran_Ban_dac_ta_De_kiem_tra_GHKI_Tin10_TietPPCT16.docx
Strictly aligned with khoi_10_tin_hoc_07092026.xls of Trường THPT Thuận Hoá (7 lessons: Mục 1.1, Mục 1.2, Mục 1.7, Mục 2.8, Mục 2.9, Mục 2.10, Mục 3.11; NO Bài 3, 4, 5, 6) and all 42 reviewer comments.
"""
from pathlib import Path
import shutil
from docx_engine_5512 import (
    init_body_with_cover, add, br, table, add_figure, add_approval_section, save_docx
)

KHOI10_DIR = Path(r'c:\Users\PC\khbd_docx\khoi10')
THANH_PHAM_DIR = Path(r'c:\Users\PC\khbd_docx\06_thanh_pham_docx')
GOI_NOP_DIR = Path(r'c:\Users\PC\khbd_docx\08_goi_nop_du_gio\HoQuangHuy_24S1020030_HoSo_DuGio_Tin10')

def build_ontap_ghki():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I (01 TIẾT – TIẾT PPCT: 15)',
        class_str='Lớp 10/2 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Hai, ngày 02/11/2026 – Tiết 2 (08h00 – 08h45) – Phòng H.I.4',
        lesson_row_str='Ôn tập kiểm tra Giữa học kì I (01 tiết – Tiết PPCT: 15)',
        page1_title='TÊN BÀI DẠY: ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I (CHỦ ĐỀ 1, CHỦ ĐỀ 2, CHỦ ĐỀ 3)',
        page1_subtitle='(Thời lượng: 01 tiết – Tiết PPCT: 15 theo Kế hoạch dạy học môn Tin học Khối 10 Trường THPT Thuận Hoá)',
        page1_class_line='Môn học: Tin học; Lớp: 10/2',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết PPCT: 15)',
        page1_sched_line='Lịch dự giờ: Thứ Hai, ngày 02/11/2026 – Tiết 2 (08h00 – 08h45) – Lớp 10/2 – Phòng H.I.4',
        page1_mode_line='Hình thức tổ chức: Dạy học trên lớp kết hợp thảo luận nhóm theo 3 trạm kiến thức'
    )

    # I. MỤC TIÊU
    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT) và Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá, tiết Ôn tập kiểm tra Giữa học kì I (Tiết PPCT: 15) hệ thống hoá toàn bộ 07 bài học đã thực hiện từ Tiết PPCT 01 đến Tiết PPCT 14 gồm: Mục 1.1 (Bài 1, trang 6–10), Mục 1.2 (Bài 2, trang 11–15), Mục 1.7 (Bài 7, trang 33–37), Mục 2.8 (Bài 8, trang 38–43), Mục 2.9 (Bài 9, trang 44–49), Mục 2.10 (Bài 10, trang 50–54) và Mục 3.11 (Bài 11, trang 55–60).')

    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau bài học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu [KT1] – [KT4] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]**: Hệ thống hoá và phân biệt rõ mối quan hệ giữa thông tin, dữ liệu và vật mang tin; trình bày được 3 bước của quá trình xử lí thông tin trong máy tính (Hình 1.1 SGK trang 6); thực hiện đúng phép quy đổi giữa các đơn vị lưu trữ dữ liệu (bit, Byte, KB, MB, GB, TB theo Bảng 1.1 SGK trang 8); nêu được vai trò của thiết bị thông minh (Mục 1.2, Bài 2 trang 11–15) và cách sử dụng thiết bị số thông dụng, quản lí tệp/thư mục (Mục 1.7, Bài 7 trang 33–37).',
        '**[KT2]**: Trình bày và phân biệt được mạng cục bộ (LAN) với mạng Internet về phạm vi, cách kết nối và quyền sở hữu (Bảng so sánh SGK Bài 8 trang 39); nêu được vai trò của dịch vụ đám mây (Cloud), Internet vạn vật (IoT), các nguy cơ mất an toàn thông tin, phần mềm độc hại (Virus, Trojan, Worm, Ransomware, Spyware) và quy trình phòng vệ, khai thác tài nguyên số an toàn (Mục 2.8, Mục 2.9, Mục 2.10 – Bài 8, 9, 10 trang 38–54).',
        '**[KT3]**: Nhận diện được 6 nhóm hành vi vi phạm đạo đức, pháp luật và văn hoá khi giao tiếp trên mạng (SGK Bài 11 trang 56); phân biệt được quyền nhân thân và quyền tài sản trong quyền tác giả theo Luật Sở hữu trí tuệ (SGK Bài 11 trang 59), dẫn chiếu đúng quy định pháp luật (Luật An ninh mạng năm 2018, Nghị định 15/2020/NĐ-CP) và nghĩa vụ tôn trọng bản quyền sản phẩm số (Mục 3.11).',
        '**[KT4]**: Vận dụng tổng hợp kiến thức của cả 3 chủ đề để giải quyết các câu hỏi trắc nghiệm nhiều lựa chọn, câu hỏi đúng/sai và bài tập xử lí tình huống thực tiễn theo đúng cấu trúc Ma trận đề kiểm tra Giữa học kì I (Tiết PPCT: 16).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**[NLC1 – Tự chủ và tự học]**: Học sinh chủ động rà soát hệ thống khái niệm cốt lõi trong Sách giáo khoa Tin học 10 (trang 6–15 và trang 33–60), tự đánh giá mức độ nắm vững kiến thức của bản thân trước kì kiểm tra định kì.',
        '**[NLC2 – Giao tiếp và hợp tác]**: Học sinh tích cực thảo luận nhóm theo 3 trạm chuyên đề, phân công rõ người điều phối, thư ký và người báo cáo để hoàn thành Phiếu học tập và bảo vệ lập luận trước lớp.',
        '**[NLC3 – Giải quyết vấn đề và sáng tạo]**: Học sinh phân tích được các tình huống thực tiễn về dung lượng lưu trữ, sự cố an ninh mạng và vi phạm bản quyền để đề xuất phương án xử lí có căn cứ khoa học, pháp lí.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]**: Giải thích được chức năng của thiết bị số thông dụng, thông số dung lượng bộ nhớ và các biện pháp cấu hình bảo mật tài khoản, thiết bị mạng cơ bản.',
        '**[NLb – Ứng xử phù hợp trong môi trường số]**: Thực hiện và tuyên truyền được quy tắc ứng xử văn minh, phòng chống bắt nạt qua mạng, kiểm chứng thông tin chống tin giả và tôn trọng bản quyền phần mềm, tài liệu số.',
        '**[NLd – Ứng dụng CNTT và truyền thông trong học và tự học]**: Khai thác hiệu quả sơ đồ hệ thống hoá kiến thức, bảng tổng hợp và công cụ số để ôn tập, chuẩn bị tốt cho bài kiểm tra Giữa học kì I.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    for item in [
        '**[NLS 1.1.NC1.a]**: Xác định và khai thác đúng nguồn tài nguyên học tập số hợp pháp phục vụ ôn tập.',
        '**[NLS 4.2.NC1.a]**: Bảo vệ thông tin cá nhân, mật khẩu và quyền riêng tư khi tham gia môi trường mạng.'
    ]:
        add(b, item, 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**[PC1 – Chăm chỉ]**: Học sinh nghiêm túc tổng hợp kiến thức, chủ động hoàn thành đầy đủ các câu hỏi trong Phiếu học tập và đề cương ôn tập.',
        '**[PC2 – Trách nhiệm]**: Học sinh có ý thức bảo vệ dữ liệu cá nhân, tuân thủ pháp luật và chịu trách nhiệm về mỗi phát ngôn, chia sẻ của bản thân trên không gian mạng.',
        '**[PC3 – Trung thực]**: Học sinh tự giác làm bài luyện tập, trung thực khi tự đánh giá và đánh giá đồng đẳng kết quả ôn tập của bạn.'
    ]:
        add(b, item, 'bullet')

    # II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    add(b, '1. Thiết bị dạy học', 'sub')
    for item in [
        '**Đối với giáo viên:** Máy tính xách tay, máy chiếu (hoặc màn hình tương tác tại phòng học H.I.4), bút chỉ laser, bảng phụ/giấy A3 cho 3 trạm thảo luận nhóm và đồng hồ bấm giờ hiển thị trên màn chiếu.',
        '**Đối với học sinh:** Sách giáo khoa, vở ghi môn Tin học 10, bút viết, máy tính bỏ túi (hỗ trợ tính toán luỹ thừa cơ số 2 khi đổi đơn vị dữ liệu).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Học liệu (Nói có sách, mách có chứng)', 'sub')
    for item in [
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, cụ thể gồm 07 bài học thuộc phạm vi kiểm tra Giữa học kì I:\n  + **Chủ đề 1 (Máy tính và xã hội tri thức):** Bài 1 (*Thông tin và xử lí thông tin*, trang 6–10), Bài 2 (*Vai trò của thiết bị thông minh và tin học đối với xã hội*, trang 11–15), Bài 7 (*Thực hành sử dụng thiết bị số thông dụng*, trang 33–37).\n  + **Chủ đề 2 (Mạng máy tính và Internet):** Bài 8 (*Mạng máy tính trong cuộc sống hiện đại*, trang 38–43), Bài 9 (*An toàn trên không gian mạng*, trang 44–49), Bài 10 (*Thực hành khai thác tài nguyên trên không gian mạng*, trang 50–54).\n  + **Chủ đề 3 (Đạo đức, pháp luật và văn hoá trong môi trường số):** Bài 11 (*Ứng xử trên môi trường số. Nghĩa vụ tôn trọng bản quyền*, trang 55–60).',
        '**Văn bản Phân phối chương trình của nhà trường:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (từ Tiết PPCT 01 đến Tiết PPCT 15).',
        '**Hệ thống Phiếu học tập và Công cụ đánh giá (đính kèm toàn văn tại Mục IV. Phụ lục):**\n  + **Phiếu học tập số 1 (Mảnh ghép Chuyên đề 1, 2, 3):** Hệ thống hoá kiến thức cốt lõi của 07 bài học (CCĐG 2).\n  + **Phiếu học tập số 2:** Bộ 10 câu hỏi trắc nghiệm khách quan rà soát nhanh chuẩn Ma trận đề kiểm tra Giữa học kì I (CCĐG 3).\n  + **Phiếu học tập số 3:** Bộ 04 bài tập xử lí tình huống thực tiễn và tự luận vận dụng (CCĐG 4).'
    ]:
        add(b, item, 'bullet')

    # III. TIẾN TRÌNH DẠY HỌC
    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học (Thời lượng: 01 tiết = 45 phút – Tiết PPCT: 15)**', first=0, keep=True)
    overview_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng và Tiết PPCT', 'Mục tiêu cốt lõi', 'Sản phẩm và Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Kích hoạt trí nhớ: Ai nhanh hơn?)', '06 phút\n(Tiết 1 của bài /\nTiết PPCT: 15)', '[KT1], [KT2], [KT3],\n[NLC1], [PC1]', 'Câu trả lời nhanh 05 từ khoá thuộc Chủ đề 1, 2, 3.\nCông cụ đánh giá: CCĐG 1 (Bảng kiểm quan sát).'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n(Hệ thống hoá kiến thức cốt lõi Chủ đề 1, 2, 3)\n- 2.1. Nhiệm vụ 1: Hệ thống hoá Chủ đề 1 (Mục 1.1, Mục 1.2, Mục 1.7)\n- 2.2. Nhiệm vụ 2: Hệ thống hoá Chủ đề 2 (Mục 2.8, Mục 2.9, Mục 2.10)\n- 2.3. Nhiệm vụ 3: Hệ thống hoá Chủ đề 3 (Mục 3.11)', '18 phút\n(Tiết 1 của bài /\nTiết PPCT: 15)\n- NV1: 06 phút\n- NV2: 06 phút\n- NV3: 06 phút', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2],\n[NLa], [NLb], [NLd],\n[NLS 1.1.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1 hoàn chỉnh của 3 trạm chuyên đề và bảng kiến thức chuẩn hoá trên bảng.\nCông cụ đánh giá: CCĐG 2 (Rubric chấm Phiếu số 1).'],
        ['3. Hoạt động 3: Luyện tập\n(Chinh phục 10 câu hỏi trắc nghiệm chuẩn cấu trúc đề GHKI)', '12 phút\n(Tiết 1 của bài /\nTiết PPCT: 15)', '[KT1], [KT2], [KT3], [KT4],\n[NLC1], [NLC3],\n[NLa], [NLb], [PC3]', 'Phiếu trả lời 10 câu hỏi trắc nghiệm (Phiếu học tập số 2) và giải thích từ khoá loại trừ nhiễu.\nCông cụ đánh giá: CCĐG 3 (Thang điểm Phiếu số 2).'],
        ['4. Hoạt động 4: Vận dụng\n(Xử lí tình huống thực tiễn và Dặn dò kiểm tra GHKI)', '09 phút\n(Tiết 1 của bài /\nTiết PPCT: 15)', '[KT1], [KT2], [KT3], [KT4],\n[NLC2], [NLC3], [NLb],\n[NLS 4.2.NC1.a],\n[PC2], [PC3]', 'Lời giải 04 bài tập tình huống thực tiễn (Phiếu học tập số 3) và kế hoạch tự ôn tập ở nhà.\nCông cụ đánh giá: CCĐG 4 (Rubric 10 điểm chấm sản phẩm vận dụng).']
    ]
    b.append(table(overview_rows, [2600, 1400, 2100, 2900], header=True, cell_sz=22))

    # HOẠT ĐỘNG 1
    add(b, '1. Hoạt động 1: Mở đầu (06 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [NLC1], [PC1] – Học sinh huy động nhanh các khái niệm cốt lõi đã học trong 14 tiết đầu năm học (Chủ đề 1, Chủ đề 2, Chủ đề 3) và nắm rõ cấu trúc bài kiểm tra Giữa học kì I (Tiết PPCT: 16).')
    add(b, '**b) Nội dung:** Giáo viên tổ chức trò chơi khởi động *"Giải mã từ khoá Tin học 10"* gồm 05 câu hỏi phản xạ nhanh tương ứng với các bài học từ Mục 1.1 đến Mục 3.11 trong Sách giáo khoa Tin học 10.')
    add(b, '**c) Sản phẩm:** Câu trả lời miệng chính xác của học sinh cho 05 từ khoá khởi động và tâm thế sẵn sàng bước vào hoạt động hệ thống hoá kiến thức.')
    add(b, '**Công cụ đánh giá:** CCĐG 1 – Bảng kiểm quan sát mức độ tham gia và phản xạ kiến thức ở Hoạt động Mở đầu *(Đính kèm chi tiết tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    hd1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu lên màn hình 05 câu hỏi nhận diện nhanh từ khoá thuộc 07 bài học đã học (Mục 1.1, Mục 1.2, Mục 1.7, Mục 2.8, Mục 2.9, Mục 2.10, Mục 3.11) và nêu rõ thể lệ: Học sinh suy nghĩ cá nhân trong 10 giây cho mỗi câu và giơ tay giành quyền trả lời:\n'
            '  + Câu 1 (Bài 1, trang 8): Đơn vị dữ liệu nhỏ nhất mà máy tính truy cập được, gồm 8 bit, có tên gọi là gì?\n'
            '  + Câu 2 (Bài 2, trang 13): Hệ thống kết nối các thiết bị, đồ vật thông minh qua mạng Internet để tự động thu thập và trao đổi dữ liệu được viết tắt là gì?\n'
            '  + Câu 3 (Bài 8, trang 39): Thiết bị mạng nào dùng để kết nối các mạng LAN với nhau và định tuyến gói tin ra mạng Internet?\n'
            '  + Câu 4 (Bài 9, trang 46): Loại phần mềm độc hại nào chuyên mã hoá dữ liệu của người dùng để đòi tiền chuộc?\n'
            '  + Câu 5 (Bài 11, trang 59): Quyền của tổ chức, cá nhân đối với tác phẩm do mình sáng tạo ra hoặc sở hữu được gọi là gì?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh lắng nghe từng câu hỏi, đối chiếu nhanh với kiến thức đã học trong Sách giáo khoa và giơ tay phát biểu.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời lần lượt 05 học sinh trả lời nhanh tại chỗ và giải thích ngắn gọn từ khoá vừa nêu.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét câu trả lời theo bảng kiểm CCĐG 1, tuyên dương các học sinh có phản xạ tốt và dẫn dắt vào bài: *"Trong 14 tiết học vừa qua, lớp chúng ta đã hoàn thành 07 bài học thuộc Chủ đề 1, Chủ đề 2 và Chủ đề 3. Tiết học hôm nay (Tiết PPCT: 15) sẽ giúp các em hệ thống hoá toàn bộ kiến thức trọng tâm để chuẩn bị tốt nhất cho bài Kiểm tra Giữa học kì I ở Tiết PPCT 16."*',
            '**Đáp án 05 câu hỏi khởi động:**\n'
            '- **Câu 1:** Byte (ký hiệu là B, 1 Byte = 8 bit – SGK Bài 1, trang 8).\n'
            '- **Câu 2:** IoT (*Internet of Things* – Internet vạn vật – SGK Bài 2, trang 13).\n'
            '- **Câu 3:** Bộ định tuyến (Router – SGK Bài 8, trang 39).\n'
            '- **Câu 4:** Phần mềm tống tiền (Ransomware – SGK Bài 9, trang 46).\n'
            '- **Câu 5:** Quyền tác giả (SGK Bài 11, trang 59).\n\n'
            '**Giới thiệu phạm vi Ôn tập Kiểm tra Giữa học kì I (Tiết PPCT: 15):**\n'
            '- **Chủ đề 1:** Bài 1 (Mục 1.1), Bài 2 (Mục 1.2), Bài 7 (Mục 1.7).\n'
            '- **Chủ đề 2:** Bài 8 (Mục 2.8), Bài 9 (Mục 2.9), Bài 10 (Mục 2.10).\n'
            '- **Chủ đề 3:** Bài 11 (Mục 3.11).'
        ]
    ]
    b.append(table(hd1_rows, [4800, 4200], header=True, cell_sz=23))

    # HOẠT ĐỘNG 2
    add(b, '2. Hoạt động 2: Hình thành kiến thức mới – Hệ thống hoá kiến thức cốt lõi Chủ đề 1, 2, 3 (18 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'heading')
    add(b, '*(Lưu ý sư phạm: Trong tiết Ôn tập, Hoạt động 2 đóng vai trò hệ thống hoá, cấu trúc lại các mạch kiến thức cốt lõi của 03 chủ đề đã học thành sơ đồ và bảng đối chiếu chuẩn xác).*', first=0, italic=True, keep=True)

    # 2.1 Nhiệm vụ 1
    add(b, '2.1. Nhiệm vụ 1: Hệ thống hoá Chủ đề 1 – Máy tính và xã hội tri thức: Bài 1, Bài 2, Bài 7 (06 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [NLa], [NLd], [PC1] – Học sinh tóm tắt chính xác mối quan hệ giữa thông tin và dữ liệu, quy trình 3 bước xử lí thông tin, bảng quy đổi đơn vị lưu trữ dữ liệu (Bài 1), vai trò của thiết bị thông minh trong cuộc cách mạng công nghiệp 4.0 (Bài 2) và kỹ năng sử dụng thiết bị số thông dụng, quản lí tệp/thư mục (Bài 7).')
    add(b, '**b) Nội dung:** Giáo viên yêu cầu các nhóm Trạm 1 hoàn thành Phần I của **Phiếu học tập số 1** dựa trên **Hình 1.1 (SGK trang 6)**, **Bảng 1.1 (SGK trang 8)**, nội dung Bài 2 (trang 11–15) và Bài 7 (trang 33–37):')
    add_figure(b, 'hinh_1_1_xu_li_tt.png')
    add_figure(b, 'bang_1_1_don_vi_luu_tru.png')
    add(b, '**c) Sản phẩm:** Nội dung Phần I trên Phiếu học tập số 1 đã điền đầy đủ từ khoá, công thức quy đổi đơn vị lưu trữ dữ liệu và phân biệt rõ thiết bị thông minh.')
    add(b, '**Công cụ đánh giá:** CCĐG 2 – Rubric đánh giá Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    nv1_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chia lớp 10/2 thành 3 cụm trạm (mỗi cụm 2 nhóm), phát **Phiếu học tập số 1** và giao nhiệm vụ cho Cụm Trạm 1 (đồng thời cả lớp cùng rà soát Phần I trong 03 phút):\n'
            '  + Phân biệt 3 khái niệm: Thông tin, Dữ liệu, Vật mang tin và nêu 3 bước xử lí thông tin theo **Hình 1.1 (SGK trang 6)**.\n'
            '  + Nêu quy tắc quy đổi các đơn vị lưu trữ dữ liệu theo **Bảng 1.1 (SGK trang 8)**.\n'
            '  + Chỉ ra 2 đặc điểm cốt lõi giúp phân biệt thiết bị thông minh với thiết bị điện tử thông thường (**Bài 2, trang 11–12**) và các thao tác quản lí tệp, xem thông số hệ thống (**Bài 7, trang 33–37**).\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Các nhóm học sinh thảo luận, đối chiếu Sách giáo khoa trang 6–15 và trang 33–37 để điền câu trả lời vào Phần I của Phiếu học tập số 1. Giáo viên quan sát, hỗ trợ các nhóm kiểm tra luỹ thừa 2¹⁰ = 1024.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Đại diện Nhóm 1 (Cụm Trạm 1) trình bày nhanh kết quả trong 02 phút; đại diện Nhóm 2 nhận xét, bổ sung ví dụ thực tế về phần mở rộng tệp (.docx, .xlsx, .jpg, .mp4, .exe).\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, chuẩn hoá kiến thức Chủ đề 1 lên màn chiếu và lưu ý học sinh lỗi thường gặp khi nhầm lẫn giữa bit (nhỏ nhất) và Byte (1 B = 8 bit), trong khi từ B -> KB -> MB -> GB -> TB đều nhân/chia cho 1024 (2¹⁰).',
            '**1. Tổng hợp kiến thức Chủ đề 1 (Bài 1, Bài 2, Bài 7):**\n'
            '**1. Thông tin và xử lí thông tin (Mục 1.1 – Bài 1, SGK tr. 6–10):**\n'
            '- **Thông tin:** Là tất cả những gì mang lại hiểu biết cho con người.\n'
            '- **Dữ liệu:** Là thông tin đã được đưa vào máy tính (dưới dạng văn bản, số, hình ảnh, âm thanh) để máy tính nhận biết và xử lí được.\n'
            '- **Vật mang tin:** Là phương tiện dùng để lưu trữ và truyền tải thông tin (giấy, đĩa CD, thẻ nhớ, ổ cứng).\n'
            '- **Quá trình xử lí thông tin của máy tính (Hình 1.1):** Gồm 3 bước: Tiếp nhận dữ liệu (Input) -> Xử lí dữ liệu (Processing) -> Đưa ra kết quả (Output), gắn liền với Lưu trữ (Storage).\n'
            '- **Đơn vị lưu trữ dữ liệu (Bảng 1.1, tr. 8):**\n'
            '  + 1 Byte (B) = 8 bit (Byte là đơn vị dữ liệu nhỏ nhất mà máy tính truy cập được).\n'
            '  + Các đơn vị kế tiếp hơn kém nhau 2¹⁰ = 1024 lần: 1 KB = 1024 B; 1 MB = 1024 KB; 1 GB = 1024 MB; 1 TB = 1024 GB.\n'
            '**2. Thiết bị thông minh và thành tựu Tin học (Mục 1.2 – Bài 2, SGK tr. 11–15):**\n'
            '- **Thiết bị thông minh:** Là thiết bị điện tử có thể hoạt động một cách tự chủ trong một mức độ nhất định nhờ phần mềm điều khiển và có khả năng kết nối với các thiết bị khác để trao đổi dữ liệu (Ví dụ: Điện thoại thông minh, đồng hồ thông minh, robot hút bụi thông minh; Máy tính bỏ túi Casio thông thường không phải là thiết bị thông minh).\n'
            '- **Vai trò trong Cách mạng công nghiệp 4.0:** Trụ cột gồm IoT (Internet vạn vật), Big Data (Dữ liệu lớn), AI (Trí tuệ nhân tạo) và Cloud Computing (Điện toán đám mây).\n'
            '**3. Thực hành sử dụng thiết bị số thông dụng (Mục 1.7 – Bài 7, SGK tr. 33–37):**\n'
            '- Đọc được các thông số kỹ thuật cơ bản của máy tính/điện thoại: Tốc độ CPU (GHz), dung lượng bộ nhớ trong RAM (GB), dung lượng ổ cứng (GB/TB), độ phân giải màn hình.\n'
            '- Quản lí tệp và thư mục khoa học; nhận biết loại tệp qua **phần mở rộng** (.docx, .pdf, .png, .mp3, .py).'
        ]
    ]
    b.append(table(nv1_rows, [4600, 4400], header=True, cell_sz=23))

    # 2.2 Nhiệm vụ 2
    add(b, '2.2. Nhiệm vụ 2: Hệ thống hoá Chủ đề 2 – Mạng máy tính và Internet: Bài 8, Bài 9, Bài 10 (06 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLC2], [NLa], [NLb], [NLS 1.1.NC1.a], [PC1], [PC2] – Học sinh phân biệt chính xác mạng LAN và Internet, dịch vụ đám mây và IoT (Bài 8), các loại phần mềm độc hại cùng biện pháp bảo vệ thông tin cá nhân (Bài 9), và cách khai thác học liệu, dịch thuật, lưu trữ trực tuyến trên không gian mạng (Bài 10).')
    add(b, '**b) Nội dung:** Giáo viên yêu cầu Cụm Trạm 2 và cả lớp đối chiếu **Bảng so sánh mạng LAN và Internet (SGK Bài 8 trang 39)** và **Hộp kiến thức bảo vệ thông tin cá nhân (SGK Bài 9 trang 44)** để hoàn thành Phần II của **Phiếu học tập số 1**:')
    add_figure(b, 'hinh_8_2_bang_so_sanh_mang.png')
    add_figure(b, 'hop_9_1_bao_ve_tt_ca_nhan.png')
    add(b, '**c) Sản phẩm:** Nội dung Phần II trên Phiếu học tập số 1 hoàn chỉnh bảng so sánh mạng LAN – Internet và quy trình phòng chống phần mềm độc hại.')
    add(b, '**Công cụ đánh giá:** CCĐG 2 – Rubric đánh giá Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    nv2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên giao nhiệm vụ cho Cụm Trạm 2 báo cáo Phần II của **Phiếu học tập số 1** dựa trên **Bảng so sánh mạng cục bộ và Internet (SGK trang 39)** và **Hộp kiến thức Bài 9 (SGK trang 44)**:\n'
            '  + So sánh mạng cục bộ (LAN) và Internet theo 3 tiêu chí: Phạm vi, quy mô; Cách kết nối; Quyền sở hữu.\n'
            '  + Phân biệt cơ chế lây lan của Virus, Worm (sâu máy tính), Trojan (ngựa thành Troia) và nêu 3 nguyên tắc bảo vệ thông tin cá nhân trên mạng.\n'
            '  + Nêu các bước khai thác tài nguyên học tập và chia sẻ tệp qua ổ đĩa đám mây (Google Drive, OneDrive) ở **Bài 10 (SGK trang 50–54)**.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh trong Cụm Trạm 2 thống nhất đáp án; các nhóm khác đối chiếu với bài làm của nhóm mình để chuẩn bị nhận xét.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Đại diện Nhóm 3 (Cụm Trạm 2) trình bày bảng đối chiếu LAN – Internet và cơ chế phòng vệ trên không gian mạng. Học sinh Nhóm 4 bổ sung ví dụ về xác thực hai yếu tố (2FA).\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, chốt kiến thức trọng tâm Chủ đề 2 và nhấn mạnh điểm phân biệt cốt lõi: LAN kết nối trực tiếp qua Hub/Switch/Wi-Fi và thuộc sở hữu của một cá nhân/tổ chức xác định; còn Internet kết nối các mạng qua Router thông qua nhà cung cấp dịch vụ (ISP) và không thuộc sở hữu riêng của ai.',
            '**2. Tổng hợp kiến thức Chủ đề 2 (Bài 8, Bài 9, Bài 10):**\n'
            '**1. Mạng máy tính trong cuộc sống hiện đại (Mục 2.8 – Bài 8, SGK tr. 38–43):**\n'
            '- **So sánh mạng LAN và Internet (SGK tr. 39):**\n'
            '  + Phạm vi, quy mô: LAN ở quy mô nhỏ (cơ quan, trường học, gia đình); Internet có quy mô toàn cầu.\n'
            '  + Cách kết nối: LAN kết nối trực tiếp qua Hub, Switch, Wi-Fi; Internet kết nối các mạng qua Router thông qua nhà cung cấp dịch vụ Internet (ISP như Viettel, VNPT, FPT).\n'
            '  + Sở hữu: LAN có chủ sở hữu xác định; Internet không thuộc quyền sở hữu của riêng cá nhân hay tổ chức nào.\n'
            '- **Điện toán đám mây (Cloud):** Cung cấp dịch vụ lưu trữ (Google Drive, Dropbox) và phần mềm trực tuyến (Google Docs) mọi lúc, mọi nơi.\n'
            '**2. An toàn trên không gian mạng (Mục 2.9 – Bài 9, SGK tr. 44–49):**\n'
            '- **Các nguy cơ trên mạng:** Tin giả/tin phản văn hoá, lừa đảo qua mạng (Phishing), lộ thông tin cá nhân, bắt nạt qua mạng (Cyberbullying), nghiện mạng.\n'
            '- **Phần mềm độc hại (Malware):** Virus (ẩn trong tệp, cần người dùng kích hoạt), Worm (sâu máy tính tự nhân bản qua mạng), Trojan (ngụy trang phần mềm hữu ích để đánh cắp dữ liệu/chiếm quyền điều khiển).\n'
            '- **Biện pháp bảo vệ:** Đặt mật khẩu mạnh + bật xác thực 2 lớp (2FA); không mở liên kết/tệp đính kèm lạ; cẩn trọng khi dùng Wi-Fi công cộng; cài phần mềm diệt virus có bản quyền.\n'
            '**3. Thực hành khai thác tài nguyên trên mạng (Mục 2.10 – Bài 10, SGK tr. 50–54):**\n'
            '- Sử dụng máy tìm kiếm, công cụ dịch tự động (Google Translate), kho học liệu mở và chia sẻ thư mục trên Google Drive có phân quyền (Chỉ xem, Nhận xét, Chỉnh sửa).'
        ]
    ]
    b.append(table(nv2_rows, [4600, 4400], header=True, cell_sz=23))

    # 2.3 Nhiệm vụ 3
    add(b, '2.3. Nhiệm vụ 3: Hệ thống hoá Chủ đề 3 – Đạo đức, pháp luật và văn hoá trong môi trường số: Bài 11 (06 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'sub')
    add(b, '**a) Mục tiêu:** [KT3], [NLC1], [NLC2], [NLb], [PC2], [PC3] – Học sinh liệt kê được các hành vi vi phạm đạo đức, pháp luật trên mạng, phân biệt rõ quyền nhân thân và quyền tài sản trong quyền tác giả, và nắm vững nghĩa vụ tôn trọng bản quyền sản phẩm thông tin, phần mềm.')
    add(b, '**b) Nội dung:** Giáo viên yêu cầu Cụm Trạm 3 đối chiếu **Hộp kiến thức Bài 11 (SGK trang 56)** và mục **3. Quyền tác giả và bản quyền (SGK trang 59–60)** để hoàn thành Phần III của **Phiếu học tập số 1**:')
    add_figure(b, 'hop_11_1_vi_pham_tren_mang.png')
    add(b, '**c) Sản phẩm:** Nội dung Phần III trên Phiếu học tập số 1 phân biệt chính xác vi phạm đạo đức với vi phạm pháp luật, quyền nhân thân với quyền tài sản.')
    add(b, '**Công cụ đánh giá:** CCĐG 2 – Rubric đánh giá Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    nv3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên yêu cầu Cụm Trạm 3 trình bày kết quả thảo luận Phần III của **Phiếu học tập số 1**:\n'
            '  + Nêu 06 hành vi vi phạm đạo đức, pháp luật và văn hoá khi sử dụng mạng theo **Hộp kiến thức trang 56**.\n'
            '  + Phân biệt Quyền nhân thân và Quyền tài sản của tác giả theo Luật Sở hữu trí tuệ (**SGK trang 59**).\n'
            '  + Nêu các hành vi vi phạm bản quyền phần mềm cần tránh (bẻ khoá crack, dùng khoá lậu, sao chép trái phép).\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh Cụm Trạm 3 rà soát nội dung Phần III; học sinh cả lớp đối chiếu với SGK trang 55–60.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Đại diện Nhóm 5 (Cụm Trạm 3) trình bày; học sinh Nhóm 6 lấy ví dụ thực tế về việc trích dẫn nguồn khi sử dụng hình ảnh, bài viết trên Internet.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét по CCĐG 2, khen ngợi sự chuẩn bị của cả 3 cụm trạm và chốt toàn bộ hệ thống lý thuyết của 07 bài học trước khi chuyển sang phần Luyện tập.',
            '**3. Tổng hợp kiến thức Chủ đề 3 (Bài 11, SGK tr. 55–60):**\n'
            '**1. Những hành vi vi phạm trên mạng (SGK tr. 56):**\n'
            '- Đưa tin không phù hợp lên mạng (tin giả, tin sai sự thật, bạo lực);\n'
            '- Công bố thông tin cá nhân/tổ chức mà không được phép;\n'
            '- Phát tán thư điện tử, tin nhắn rác (Spam);\n'
            '- Vi phạm bản quyền khi sử dụng dữ liệu và phần mềm;\n'
            '- Bắt nạt qua mạng (Cyberbullying) và lừa đảo qua mạng;\n'
            '- Ứng xử thiếu văn hoá (chửi tục, công kích, sỉ nhục người khác).\n'
            '**2. Văn bản pháp luật điều chỉnh (SGK tr. 57–58):**\n'
            '- **Luật An ninh mạng năm 2018** (Điều 8, Điều 16, Điều 18) và **Nghị định 15/2020/NĐ-CP** (Điều 101 quy định xử phạt hành vi đăng tải, chia sẻ thông tin giả mạo, sai sự thật, xúc phạm danh dự).\n'
            '**3. Quyền tác giả và bản quyền (Luật Sở hữu trí tuệ – SGK tr. 59–60):**\n'
            '- **Quyền nhân thân:** Đặt tên cho tác phẩm; đứng tên thật/bút danh; công bố tác phẩm; bảo vệ sự toàn vẹn của tác phẩm.\n'
            '- **Quyền tài sản:** Làm tác phẩm phái sinh; sao chép tác phẩm; truyền đạt tác phẩm đến công chúng; cho thuê bản gốc/bản sao chương trình máy tính.\n'
            '- **Tôn trọng bản quyền:** Chỉ sử dụng phần mềm có bản quyền hoặc phần mềm nguồn mở miễn phí (Inkscape, LibreOffice, Ubuntu); ghi rõ nguồn và tên tác giả khi trích dẫn.'
        ]
    ]
    b.append(table(nv3_rows, [4600, 4400], header=True, cell_sz=23))

    # HOẠT ĐỘNG 3: LUYỆN TẬP
    add(b, '3. Hoạt động 3: Luyện tập (12 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [KT4], [NLC1], [NLC3], [NLa], [NLb], [PC3] – Học sinh giải quyết chính xác bộ 10 câu hỏi trắc nghiệm khách quan bao phủ đủ 07 bài học (Mục 1.1, Mục 1.2, Mục 1.7, Mục 2.8, Mục 2.9, Mục 2.10, Mục 3.11) và phân tích được cách loại trừ các phương án nhiễu.')
    add(b, '**b) Nội dung:** Học sinh làm việc cá nhân trong 06 phút để hoàn thành **Phiếu học tập số 2** gồm 10 câu hỏi trắc nghiệm khách quan (8 câu nhiều lựa chọn và 2 câu trắc nghiệm Đúng/Sai chuẩn định dạng đề kiểm tra Giữa học kì I), sau đó đổi chéo phiếu kiểm tra và giải thích đáp án.')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 2 đã khoanh đáp án, chấm điểm chéo và lời giải thích căn cứ Sách giáo khoa cho từng câu.')
    add(b, '**Công cụ đánh giá:** CCĐG 3 – Bảng đáp án và thang điểm chấm Phiếu học tập số 2 *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    hd3_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên phát **Phiếu học tập số 2** (đồng thời chiếu 10 câu hỏi lên màn hình) và yêu cầu học sinh làm bài độc lập trong 06 phút:\n'
            '  + Câu 1 (Mục 1.1): Phân biệt thông tin và dữ liệu.\n'
            '  + Câu 2 (Mục 1.1): Quy đổi dung lượng từ GB sang MB và tính số lượng tệp lưu trữ.\n'
            '  + Câu 3 (Mục 1.2): Nhận diện thiết bị thông minh.\n'
            '  + Câu 4 (Mục 1.7): Nhận diện thông số kỹ thuật và phần mở rộng tệp trên thiết bị số.\n'
            '  + Câu 5 (Mục 2.8): Phân biệt chức năng của Switch và Router, mạng LAN và Internet.\n'
            '  + Câu 6 (Mục 2.9): Phân biệt cơ chế hoạt động của Trojan, Worm, Virus.\n'
            '  + Câu 7 (Mục 2.10): Phân quyền chia sẻ thư mục học tập trên Google Drive.\n'
            '  + Câu 8 (Mục 3.11): Phân biệt quyền nhân thân và quyền tài sản theo Luật Sở hữu trí tuệ.\n'
            '  + Câu 9 và Câu 10 (Đúng/Sai tổng hợp Chủ đề 1, 2, 3): Đánh giá 4 nhận định trong mỗi tình huống thực tế.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh suy nghĩ độc lập, tính toán và khoanh tròn đáp án vào Phiếu học tập số 2. Hết 06 phút, hai học sinh cùng bàn đổi phiếu cho nhau để chuẩn bị chấm chéo.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời lần lượt học sinh giải thích đáp án và chỉ rõ trang SGK tham chiếu cho các câu hỏi dễ nhầm (Câu 2, Câu 5, Câu 6, Câu 8, Câu 9, Câu 10).\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Học sinh chấm điểm chéo theo thang điểm CCĐG 3 (1,0 điểm/câu). Giáo viên thống kê nhanh số học sinh đạt từ 8/10 điểm trở lên và giải đáp triệt để các thắc mắc.',
            '**Đáp án tóm tắt Phiếu học tập số 2 (10 câu trắc nghiệm):**\n'
            '- **Câu 1 (Bài 1, tr. 7):** **B** *(Dữ liệu là thông tin đã được đưa vào máy tính; thông tin là ý nghĩa rút ra từ dữ liệu)*.\n'
            '- **Câu 2 (Bài 1, tr. 8):** **C** *(2 GB = 2 × 1024 = 2048 MB; mỗi tệp 4 MB nên chứa tối đa 2048 : 4 = 512 tệp)*.\n'
            '- **Câu 3 (Bài 2, tr. 11):** **D** *(Máy tính bỏ túi thông thường không có hệ điều hành tự chủ và không kết nối mạng nên không phải thiết bị thông minh)*.\n'
            '- **Câu 4 (Bài 7, tr. 34):** **A** *(Thông số 8 GB đi kèm RAM chỉ dung lượng bộ nhớ truy cập ngẫu nhiên)*.\n'
            '- **Câu 5 (Bài 8, tr. 39):** **C** *(Switch chỉ chuyển tiếp dữ liệu trong nội bộ mạng LAN, còn Router dùng để kết nối các mạng LAN với nhau và ra Internet)*.\n'
            '- **Câu 6 (Bài 9, tr. 46):** **B** *(Trojan là phần mềm nội gián ngụy trang dưới dạng phần mềm hợp pháp để đánh cắp thông tin hoặc chiếm quyền điều khiển)*.\n'
            '- **Câu 7 (Bài 10, tr. 52):** **B** *(Chế độ **Người xem / Viewer** cho phép các bạn đọc và tải tài liệu nhưng không thể xoá hay sửa tệp gốc)*.\n'
            '- **Câu 8 (Bài 11, tr. 59):** **C** *(Sao chép tác phẩm và làm tác phẩm phái sinh thuộc Quyền tài sản, không phải quyền nhân thân)*.\n'
            '- **Câu 9 (Đúng/Sai – Chủ đề 1 và 2):** **a) Đúng; b) Sai; c) Đúng; d) Sai.**\n'
            '- **Câu 10 (Đúng/Sai – Chủ đề 2 và 3):** **a) Đúng; b) Sai; c) Đúng; d) Đúng.**'
        ]
    ]
    b.append(table(hd3_rows, [4600, 4400], header=True, cell_sz=23))

    # HOẠT ĐỘNG 4: VẬN DỤNG
    add(b, '4. Hoạt động 4: Vận dụng (09 phút – Tiết 1 của bài / Tiết PPCT: 15)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [KT4], [NLC2], [NLC3], [NLb], [NLS 4.2.NC1.a], [PC2], [PC3] – Học sinh vận dụng kiến thức tổng hợp của 07 bài học để giải quyết 04 bài tập tình huống thực tiễn (tính toán dung lượng lưu trữ, thiết kế kết nối mạng, xử lí sự cố lừa đảo/bắt nạt qua mạng và bảo vệ bản quyền số).')
    add(b, '**b) Nội dung:** Giáo viên giao **Phiếu học tập số 3** gồm 04 tình huống tự luận thực tiễn. Các nhóm thảo luận nhanh và trình bày quy trình xử lí tình huống trọng tâm ngay tại lớp; các bài còn lại hoàn thiện vào vở để chuẩn bị cho tiết Kiểm tra Giữa học kì I (Tiết PPCT: 16).')
    add(b, '**c) Sản phẩm:** Lời giải chi tiết 04 bài tập tình huống trong Phiếu học tập số 3 và kế hoạch tự ôn tập cá nhân của học sinh.')
    add(b, '**Công cụ đánh giá:** CCĐG 4 – Rubric 10 điểm đánh giá bài tập xử lí tình huống Vận dụng *(Đính kèm chi tiết tại Mục IV. Phụ lục 3)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', 'heading')
    hd4_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu nội dung **Phiếu học tập số 3** gồm 04 bài tập tự luận chuẩn cấu trúc phần Tự luận (3,0 điểm) của Đề kiểm tra Giữa học kì I:\n'
            '  + Bài toán 1 (Mục 1.1 và Mục 1.7): Thẻ nhớ máy ảnh còn trống 4 GB. Học sinh cần lưu một video kỉ niệm của lớp nặng 1,5 GB và các bức ảnh chụp định dạng .jpg, trung bình mỗi bức ảnh nặng 4 MB. Hỏi sau khi lưu video, thẻ nhớ còn lưu thêm được tối đa bao nhiêu bức ảnh?\n'
            '  + Tình huống 2 (Mục 2.8 và Mục 2.10): Nhóm học sinh cần làm báo cáo chung nhưng mỗi bạn ở một nhà khác nhau. Em hãy tư vấn giải pháp công cụ đám mây và cách phân quyền để cả nhóm cùng soạn thảo mà không sợ mất dữ liệu.\n'
            '  + Tình huống 3 (Mục 2.9 và Mục 3.11): Bạn cùng lớp nhận được tin nhắn từ tài khoản lạ đe doạ tung ảnh chế ghép xúc phạm lên mạng xã hội và gửi kèm tệp Danh_sach_anh.exe yêu cầu mở ra xem. Em hãy hướng dẫn bạn 4 bước xử lí an toàn và đúng pháp luật.\n'
            '  + Tình huống 4 (Mục 3.11): Một bạn tải phần mềm thiết kế đồ hoạ thương mại có gắn mã bẻ khoá (crack) trên mạng về máy để làm bài tập. Hãy phân tích 2 rủi ro của hành vi này và đề xuất phần mềm thay thế hợp pháp.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh thảo luận theo cặp đôi trong 04 phút để giải Bài toán 1 và Tình huống 3 tại lớp, đồng thời ghi dàn ý cho Tình huống 2 và Tình huống 4.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh lên bảng trình bày lời giải Bài toán 1 và Tình huống 3; các học sinh khác nhận xét, đối chiếu theo tiêu chí của Rubric CCĐG 4.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, chấm điểm mẫu theo Rubric CCĐG 4 và dặn dò lịch Kiểm tra Giữa học kì I (Tiết PPCT: 16, thời gian làm bài 45 phút, gồm 70% trắc nghiệm và 30% tự luận phủ đúng 07 bài Mục 1.1, Mục 1.2, Mục 1.7, Mục 2.8, Mục 2.9, Mục 2.10, Mục 3.11).',
            '**Hướng dẫn giải bài tập vận dụng (Phiếu học tập số 3):**\n'
            '**1. Lời giải Bài toán 1 (Mục 1.1 và Mục 1.7):**\n'
            '- Dung lượng thẻ nhớ còn lại sau khi lưu video 1,5 GB là:\n'
            '  4 GB - 1,5 GB = 2,5 GB.\n'
            '- Đổi sang đơn vị MB (theo Bảng 1.1 SGK trang 8, 1 GB = 1024 MB):\n'
            '  2,5 × 1024 = 2560 MB.\n'
            '- Số bức ảnh tối đa lưu thêm được là:\n'
            '  2560 : 4 = 640 (bức ảnh).\n'
            '**2. Giải pháp Tình huống 2 (Mục 2.8 và Mục 2.10):**\n'
            '- Sử dụng dịch vụ đám mây Google Drive / Google Docs; tạo thư mục chung, phân quyền **Chỉnh sửa (Editor)** cho các thành viên trong nhóm và quyền **Chỉ xem (Viewer)** cho các bạn nhóm khác.\n'
            '**3. Quy trình 4 bước xử lí Tình huống 3 (Mục 2.9 và Mục 3.11):**\n'
            '- **Bước 1 (An toàn kỹ thuật):** Tuyệt đối không mở tệp Danh_sach_anh.exe vì phần mở rộng .exe là tệp thực thi có nguy cơ cao chứa Trojan/Ransomware.\n'
            '- **Bước 2 (Không phản ứng tiêu cực):** Không chửi bới, công kích lại kẻ đe doạ trên mạng và không làm theo yêu cầu tống tiền/ép buộc.\n'
            '- **Bước 3 (Lưu bằng chứng):** Chụp ảnh màn hình tin nhắn đe doạ, lưu đường dẫn tài khoản của đối tượng và chặn/báo cáo (Report) tài khoản vi phạm.\n'
            '- **Bước 4 (Tìm kiếm sự hỗ trợ hợp pháp):** Báo ngay cho cha mẹ, giáo viên chủ nhiệm hoặc cơ quan chức năng; hành vi vu khống, làm nhục người khác trên mạng vi phạm Luật An ninh mạng năm 2018 và Điều 101 Nghị định 15/2020/NĐ-CP.\n'
            '**4. Giải pháp Tình huống 4 (Mục 3.11):**\n'
            '- **Hai rủi ro:** (1) Vi phạm quyền tác giả/bản quyền phần mềm theo Luật Sở hữu trí tuệ; (2) Tệp crack thường bị kẻ xấu cài sẵn mã độc (Trojan, Spyware).\n'
            '- **Giải pháp hợp pháp:** Gỡ bỏ bản crack và cài đặt phần mềm nguồn mở miễn phí Inkscape (sẽ học ngay ở Chủ đề 4 – Bài 12 và Bài 13 sau kì kiểm tra Giữa kì I).'
        ]
    ]
    b.append(table(hd4_rows, [4600, 4400], header=True, cell_sz=23))

    # IV. PHỤ LỤC
    add(b, 'IV. PHỤ LỤC (HỒ SƠ HỌC LIỆU, PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ)', 'heading')
    add(b, 'Phụ lục 1. Công cụ đánh giá CCĐG 1, CCĐG 2 và Phiếu học tập số 1 (Hệ thống hoá 3 Chủ đề)', 'sub')
    ccdg1_rows = [
        ['STT', 'Tiêu chí quan sát (CCĐG 1 – Hoạt động 1: Mở đầu)', 'Đạt', 'Chưa đạt', 'Ghi chú'],
        ['1', 'Học sinh chú ý lắng nghe và chủ động giơ tay tham gia giải mã 05 từ khoá khởi động.', '[  ]', '[  ]', ''],
        ['2', 'Học sinh trả lời đúng từ khoá và nêu được bài học tương ứng trong SGK Tin học 10.', '[  ]', '[  ]', ''],
        ['3', 'Học sinh xác định đúng phạm vi 07 bài học (Mục 1.1 đến Mục 3.11) của bài kiểm tra Giữa kì I.', '[  ]', '[  ]', '']
    ]
    b.append(table(ccdg1_rows, [700, 5100, 900, 1100, 1200], header=True, cell_sz=22))
    add(b, '', first=0)

    add(b, '**Phiếu học tập số 1 – Hệ thống hoá kiến thức Chủ đề 1, Chủ đề 2, Chủ đề 3 (Hoạt động 2)**', first=0, keep=True)
    pht1_rows = [
        ['Trạm chuyên đề', 'Câu hỏi và yêu cầu hệ thống hoá (Dựa theo SGK Tin học 10 KNTT)', 'Phần trả lời của nhóm học sinh'],
        ['PHẦN I.\nCHỦ ĐỀ 1\n(Bài 1, Bài 2, Bài 7\nTiết PPCT: 01–06)',
         '1. Nêu 3 bước của quá trình xử lí thông tin trong máy tính (Hình 1.1 trang 6).\n'
         '2. Điền số thích hợp (Bảng 1.1 trang 8): 1 Byte = ...... bit; 1 KB = ...... B; 1 MB = ...... KB; 1 GB = ...... MB; 1 TB = ...... GB.\n'
         '3. Nêu 2 đặc điểm cốt lõi của thiết bị thông minh (Bài 2 trang 11) và ý nghĩa của phần mở rộng tệp (Bài 7 trang 35).',
         '1. Ba bước: ........................................................................\n'
         '2. Quy đổi: ........................................................................\n'
         '3. Thiết bị thông minh và quản lí tệp: ............................................'],
        ['PHẦN II.\nCHỦ ĐỀ 2\n(Bài 8, Bài 9, Bài 10\nTiết PPCT: 07–12)',
         '1. So sánh mạng LAN và mạng Internet về: Phạm vi, Cách kết nối, Quyền sở hữu (Bảng trang 39).\n'
         '2. Phân biệt Virus, Worm, Trojan (Bài 9 trang 46) và nêu 3 biện pháp bảo vệ thông tin cá nhân (Hộp kiến thức trang 44).\n'
         '3. Nêu 3 mức phân quyền khi chia sẻ tệp trên Google Drive (Bài 10 trang 52).',
         '1. So sánh LAN và Internet: .....................................................\n'
         '2. Malware và Bảo vệ thông tin cá nhân: .....................................\n'
         '3. Ba mức phân quyền đám mây: .................................................'],
        ['PHẦN III.\nCHỦ ĐỀ 3\n(Bài 11\nTiết PPCT: 13–14)',
         '1. Kể tên 6 hành vi vi phạm đạo đức, pháp luật và văn hoá trên mạng (Hộp kiến thức trang 56).\n'
         '2. Phân biệt Quyền nhân thân và Quyền tài sản trong quyền tác giả (Bài 11 trang 59).\n'
         '3. Nêu 2 việc học sinh cần làm để tôn trọng bản quyền phần mềm và sản phẩm số.',
         '1. Sáu hành vi vi phạm: .........................................................\n'
         '2. Quyền nhân thân vs Quyền tài sản: .......................................\n'
         '3. Nghĩa vụ tôn trọng bản quyền: ..............................................']
    ]
    b.append(table(pht1_rows, [1700, 3900, 3400], header=True, cell_sz=22))

    add(b, 'Phụ lục 2. Phiếu học tập số 2 và Công cụ đánh giá CCĐG 3 (Bộ 10 câu hỏi trắc nghiệm Hoạt động 3)', 'sub')
    for q_line in [
        '- **Câu 1 (Mục 1.1 – Nhận biết):** Phát biểu nào sau đây phản ánh đúng mối quan hệ giữa thông tin và dữ liệu trong Tin học? **A.** Dữ liệu là hiểu biết của con người, thông tin là các dãy bit trong máy tính. **B.** Thông tin là ý nghĩa đem lại hiểu biết cho con người, còn dữ liệu là thông tin đã được đưa vào máy tính. **C.** Thông tin và dữ liệu hoàn toàn giống nhau trong mọi trường hợp. **D.** Máy tính chỉ lưu trữ thông tin chứ không lưu trữ dữ liệu.',
        '- **Câu 2 (Mục 1.1 – Vận dụng):** Một thẻ nhớ có dung lượng trống là 2 GB. Hỏi thẻ nhớ đó có thể lưu trữ tối đa bao nhiêu tệp ảnh kỉ yếu, biết dung lượng trung bình của mỗi tệp ảnh là 4 MB? **A.** 500 tệp. **B.** 256 tệp. **C.** 512 tệp. **D.** 1024 tệp.',
        '- **Câu 3 (Mục 1.2 – Nhận biết):** Thiết bị nào sau đây **không phải** là thiết bị thông minh? **A.** Đồng hồ thông minh (Smartwatch). **B.** Điện thoại thông minh (Smartphone). **C.** Robot hút bụi tự động lập bản đồ phòng. **D.** Máy tính bỏ túi thông thường chỉ thực hiện các phép tính số học cơ bản.',
        '- **Câu 4 (Mục 1.7 – Thông hiểu):** Khi xem thông tin cấu hình của một máy tính có ghi RAM: 8.00 GB, thông số này cho biết điều gì? **A.** Dung lượng bộ nhớ truy cập ngẫu nhiên hỗ trợ CPU xử lí tạm thời dữ liệu đang chạy. **B.** Tổng dung lượng lưu trữ vĩnh viễn của ổ đĩa cứng. **C.** Tốc độ xử lí của vi xử lí trung tâm. **D.** Độ phân giải của màn hình máy tính.',
        '- **Câu 5 (Mục 2.8 – Thông hiểu):** Điểm khác biệt cơ bản giữa bộ chuyển mạch (Switch) và bộ định tuyến (Router) là gì? **A.** Switch dùng để kết nối Internet, Router chỉ dùng trong phòng máy. **B.** Switch phát sóng Wi-Fi, còn Router chỉ truyền tín hiệu qua cáp đồng. **C.** Switch kết nối các thiết bị trong cùng một mạng cục bộ (LAN), còn Router dùng để kết nối các mạng LAN với nhau và định tuyến ra Internet. **D.** Switch và Router có chức năng hoàn toàn giống nhau.',
        '- **Câu 6 (Mục 2.9 – Thông hiểu):** Loại phần mềm độc hại nào dưới đây không tự nhân bản nhưng ngụy trang dưới dạng một phần mềm tiện ích hợp pháp để lừa người dùng cài đặt, từ đó mở cửa hậu đánh cắp thông tin cá nhân? **A.** Virus máy tính. **B.** Ngựa thành Troia (Trojan). **C.** Sâu máy tính (Worm). **D.** Tường lửa (Firewall).',
        '- **Câu 7 (Mục 2.10 – Vận dụng):** Để chia sẻ một thư mục tài liệu ôn tập trên Google Drive cho cả lớp cùng đọc và tải về máy nhưng không ai được tự ý xoá hoặc chỉnh sửa tệp gốc của mình, em cần thiết lập quyền truy cập nào? **A.** Người chỉnh sửa (Editor). **B.** Người xem (Viewer). **C.** Chuyển quyền sở hữu (Transfer ownership). **D.** Công khai mật khẩu tài khoản Google cho cả lớp.',
        '- **Câu 8 (Mục 3.11 – Thông hiểu):** Theo Luật Sở hữu trí tuệ (SGK Bài 11 trang 59), quyền nào sau đây thuộc về Quyền tài sản của quyền tác giả? **A.** Đặt tên cho tác phẩm. **B.** Đứng tên thật hoặc bút danh trên tác phẩm. **C.** Sao chép tác phẩm, làm tác phẩm phái sinh và cho thuê bản sao chương trình máy tính. **D.** Bảo vệ sự toàn vẹn của tác phẩm.',
        '- **Câu 9 (Mục 1.1, Mục 1.7, Mục 2.8 – Trắc nghiệm Đúng/Sai):** Bạn An sử dụng máy tính cá nhân có ổ cứng trống 50 GB kết nối vào mạng Wi-Fi của gia đình để tải các video bài giảng Tin học 10. Hãy xác định tính **Đúng (Đ)** hoặc **Sai (S)** của các nhận định sau:\n  a) **[Đ]** Mạng Wi-Fi trong gia đình bạn An thuộc phạm vi mạng cục bộ (LAN) và có chủ sở hữu xác định.\n  b) **[S]** Để kết nối mạng gia đình ra Internet, bạn An chỉ cần một thiết bị Switch mà không cần đăng kí với nhà cung cấp dịch vụ ISP.\n  c) **[Đ]** Các tệp video bài giảng có phần mở rộng là .mp4 và thuộc loại dữ liệu đa phương tiện.\n  d) **[S]** Với 50 GB dung lượng trống, ổ cứng của An chỉ chứa được tối đa 40 tệp video mà mỗi tệp nặng 1024 MB *(Sai vì 1024 MB = 1 GB, nên chứa được tới 50 tệp)*.',
        '- **Câu 10 (Mục 2.9, Mục 3.11 – Trắc nghiệm Đúng/Sai):** Trong quá trình tham gia mạng xã hội và học tập trực tuyến, học sinh cần tuân thủ quy tắc an toàn và pháp luật. Hãy xác định tính **Đúng (Đ)** hoặc **Sai (S)** của các nhận định sau:\n  a) **[Đ]** Không nên tự ý đăng tải số căn cước công dân, số điện thoại và địa chỉ nhà riêng của người khác lên mạng xã hội khi chưa được phép.\n  b) **[S]** Khi thấy một thông tin giật gân chưa rõ nguồn gốc, học sinh có thể chia sẻ ngay lên trang cá nhân kèm dòng chữ *"Chưa kiểm chứng"* thì sẽ không vi phạm pháp luật.\n  c) **[Đ]** Sử dụng phần mềm nguồn mở miễn phí Inkscape để vẽ đồ hoạ là hành vi tôn trọng bản quyền phần mềm.\n  d) **[Đ]** Khi trích dẫn một đoạn văn bản hoặc bức ảnh trên mạng vào bài thuyết trình của nhóm, học sinh cần ghi rõ tên tác giả và nguồn trích dẫn.'
    ]:
        add(b, q_line, first=0)

    add(b, 'Phụ lục 3. Công cụ đánh giá CCĐG 4 – Rubric 10 điểm chấm sản phẩm Vận dụng (Phiếu học tập số 3)', 'sub')
    rubric_info = [
        ['Thông tin hành chính phiếu đánh giá (CCĐG 4)', 'Nội dung chi tiết'],
        ['Tên bài dạy và Tiết PPCT', 'Ôn tập kiểm tra Giữa học kì I (01 tiết – Tiết PPCT: 15)'],
        ['Lớp giảng dạy và Lịch dự giờ', 'Lớp 10/2 – Trường THPT Thuận Hoá (Thứ Hai, 02/11/2026 – Tiết 2, Phòng H.I.4)'],
        ['Giáo viên giảng dạy / GVHD', 'Giáo viên giảng dạy & GVHD: Thầy Mai Xuân Văn – Sinh viên thực hiện: Hồ Quang Huy'],
        ['Họ và tên học sinh / Nhóm thực hiện', '........................................................................ Nhóm / Bàn số: ...................']
    ]
    b.append(table(rubric_info, [3600, 5400], header=True, cell_sz=22))
    add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Căn cứ SGK và Pháp lí bắt buộc trích dẫn', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Điểm tối đa', 'Điểm đạt'],
        ['1', 'Bài toán 1:\nQuy đổi đơn vị lưu trữ dữ liệu (Mục 1.1 và Mục 1.7)', 'Bảng 1.1 SGK trang 8:\n1 GB = 1024 MB;\nPhép trừ dung lượng và chia nguyên.', '- **Mức 3 (2,5 đ):** Tính đúng dung lượng còn lại 2,5 GB = 2560 MB và số ảnh tối đa 640 ảnh, trình bày rõ đơn vị.\n- **Mức 2 (1,5 đ):** Nêu đúng công thức 1 GB = 1024 MB nhưng nhầm một bước tính số học.\n- **Mức 1 (0,5 đ):** Nhầm hệ số quy đổi là 1000 và không ra kết quả đúng.', '2,5 điểm', '......'],
        ['2', 'Tình huống 2:\nKhai thác đám mây và phân quyền (Mục 2.8 và Mục 2.10)', 'SGK Bài 8 (tr. 41) và Bài 10 (tr. 52):\nGoogle Drive / Docs; quyền Editor và Viewer.', '- **Mức 3 (2,5 đ):** Chọn đúng dịch vụ đám mây và phân biệt rõ quyền Chỉnh sửa (Editor) cho thành viên nhóm với Chỉ xem (Viewer).\n- **Mức 2 (1,5 đ):** Nêu được tên dịch vụ đám mây nhưng chưa phân biệt rõ các mức phân quyền.\n- **Mức 1 (0,5 đ):** Chỉ nêu gửi tệp rời rạc qua tin nhắn, chưa áp dụng lưu trữ đám mây.', '2,5 điểm', '......'],
        ['3', 'Tình huống 3:\nPhòng vệ trước mã độc và bắt nạt qua mạng (Mục 2.9 và Mục 3.11)', 'SGK Bài 9 (tr. 44–48) và Bài 11 (tr. 56–58):\nTệp .exe, Luật An ninh mạng 2018, NĐ 15/2020/NĐ-CP.', '- **Mức 3 (2,5 đ):** Trình bày đủ 4 bước: Không mở tệp .exe; không đôi co; chụp màn hình lưu bằng chứng; báo người lớn/cơ quan chức năng kèm căn cứ pháp lí.\n- **Mức 2 (1,5 đ):** Nêu được 2–3 bước xử lí an toàn nhưng thiếu trích dẫn quy định pháp luật.\n- **Mức 1 (0,5 đ):** Chỉ nêu cách xử lí cảm tính, thiếu biện pháp kỹ thuật phòng chống mã độc.', '2,5 điểm', '......'],
        ['4', 'Tình huống 4:\nTôn trọng bản quyền phần mềm (Mục 3.11)', 'SGK Bài 11 (tr. 59–60):\nLuật Sở hữu trí tuệ;\nPhần mềm nguồn mở Inkscape.', '- **Mức 3 (2,5 đ):** Phân tích đủ 2 rủi ro (vi phạm quyền tài sản của tác giả và nguy cơ nhiễm Trojan từ tệp crack); đề xuất đúng phần mềm miễn phí Inkscape.\n- **Mức 2 (1,5 đ):** Nêu được hành vi vi phạm bản quyền nhưng chưa đề xuất được phần mềm nguồn mở thay thế.\n- **Mức 1 (0,5 đ):** Chưa nhận diện được rủi ro pháp lí và an toàn thông tin của phần mềm bẻ khoá.', '2,5 điểm', '......'],
        ['', '**Tổng điểm**', '', '**Xếp loại:** Đạt từ 8,0 – 10,0: Tốt | 5,0 – 7,9: Đạt | < 5,0: Cần cố gắng', '**10,0 điểm**', '......']
    ]
    b.append(table(rubric_rows, [650, 1500, 1950, 3500, 700, 700], header=True, cell_sz=21))

    # V. NHẬN XÉT CỦA GVHD VÀ KÝ DUYỆT
    add_approval_section(b, sign_date_str='Huế, ngày 01 tháng 11 năm 2026')

    used_figs = [
        'hinh_1_1_xu_li_tt.png',
        'bang_1_1_don_vi_luu_tru.png',
        'hinh_8_2_bang_so_sanh_mang.png',
        'hop_9_1_bao_ve_tt_ca_nhan.png',
        'hop_11_1_vi_pham_tren_mang.png'
    ]
    out_main = KHOI10_DIR / '01_De_cuong_du_gio_Tin10_OnTap_GHKI_TietPPCT15_Lop10-2.docx'
    save_docx(b, out_main, used_figures=used_figs)
    shutil.copy2(out_main, THANH_PHAM_DIR / 'Đề cương bài dạy Tin10_OntapGK1.docx')
    shutil.copy2(out_main, GOI_NOP_DIR / '01_KHBD_Tin10_On_tap_GK1.docx')
    print('Built:', out_main)

def build_matran_ghki():
    b = init_body_with_cover(
        cover_title='HỒ SƠ KIỂM TRA ĐÁNH GIÁ ĐỊNH KÌ – KHỐI 10',
        cover_subtitle='MA TRẬN, BẢN ĐẶC TẢ, ĐỀ KIỂM TRA VÀ ĐÁP ÁN GIỮA HỌC KÌ I (TIẾT PPCT: 16)',
        class_str='Khối 10 (Lớp 10/1, 10/2, 10/3) – Trường THPT Thuận Hoá',
        schedule_str='Tuần 8 – Tiết PPCT: 16 (Sau tiết Ôn tập Tiết PPCT: 15)',
        lesson_row_str='Kiểm tra giữa kỳ (01 tiết = 45 phút – Tiết PPCT: 16)',
        page1_title='MA TRẬN, BẢN ĐẶC TẢ, ĐỀ KIỂM TRA VÀ HƯỚNG DẪN CHẤM KIỂM TRA GIỮA HỌC KÌ I',
        page1_subtitle='(Thời lượng: 01 tiết = 45 phút – Tiết PPCT: 16 theo Kế hoạch dạy học môn Tin học Khối 10 Trường THPT Thuận Hoá)',
        page1_class_line='Môn học: Tin học; Khối: 10 (Bộ sách Kết nối tri thức với cuộc sống)',
        page1_time_line='Thời gian làm bài: 45 phút (Không kể thời gian phát đề)',
        page1_sched_line='Phạm vi kiểm tra: 07 bài học từ Tiết PPCT 01 đến Tiết PPCT 14 (Mục 1.1, Mục 1.2, Mục 1.7, Mục 2.8, Mục 2.9, Mục 2.10, Mục 3.11)',
        page1_mode_line='Hình thức kiểm tra: Kết hợp Trắc nghiệm khách quan (7,0 điểm) và Tự luận (3,0 điểm)'
    )

    add(b, 'I. CĂN CỨ XÂY DỰNG MA TRẬN VÀ BẢN ĐẶC TẢ ĐỀ KIỂM TRA GIỮA HỌC KÌ I', 'heading')
    for item in [
        '**Căn cứ Chương trình Giáo dục phổ thông 2018 môn Tin học** (Ban hành kèm theo Thông tư số 32/2018/TT-BGDĐT ngày 26/12/2018 của Bộ trưởng Bộ Giáo dục và Đào tạo).',
        '**Căn cứ Công văn số 7991/BGDĐT-GDTrH** về việc hướng dẫn thực hiện kiểm tra, đánh giá định kì cấp Trung học theo định hướng phát triển phẩm chất và năng lực học sinh (cấu trúc đề gồm: Trắc nghiệm nhiều lựa chọn 3,0 điểm; Trắc nghiệm Đúng/Sai 4,0 điểm; Tự luận 3,0 điểm; tỉ lệ mức độ nhận thức: **Nhận biết 40% – Thông hiểu 30% – Vận dụng 30%**).',
        '**Căn cứ Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá:** Tính đến hết Tiết PPCT 14 (trước tiết Ôn tập – Tiết PPCT 15 và tiết Kiểm tra giữa kỳ – Tiết PPCT 16), học sinh Khối 10 học đúng 07 bài học thuộc 03 chủ đề:\n  + **Chủ đề 1 (06 tiết – Tiết PPCT 01–06):** Mục 1.1 Thông tin và xử lí thông tin (02 tiết – SGK Bài 1, tr. 6–10); Mục 1.2 Vai trò của thiết bị thông minh và tin học đối với xã hội (02 tiết – SGK Bài 2, tr. 11–15); Mục 1.7 Thực hành sử dụng thiết bị số thông dụng (02 tiết – SGK Bài 7, tr. 33–37). *(Nhà trường không phân phối Bài 3, Bài 4, Bài 5, Bài 6 nên đề kiểm tra tuyệt đối không đưa các nội dung đó vào đề)*.\n  + **Chủ đề 2 (06 tiết – Tiết PPCT 07–12):** Mục 2.8 Mạng máy tính trong cuộc sống hiện đại (02 tiết – SGK Bài 8, tr. 38–43); Mục 2.9 An toàn trên không gian mạng (02 tiết – SGK Bài 9, tr. 44–49); Mục 2.10 Thực hành khai thác tài nguyên trên không gian mạng (02 tiết – SGK Bài 10, tr. 50–54).\n  + **Chủ đề 3 (02 tiết – Tiết PPCT 13–14):** Mục 3.11 Ứng xử trên môi trường số, nghĩa vụ tôn trọng bản quyền (02 tiết – SGK Bài 11, tr. 55–60).'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. KHUNG MA TRẬN ĐỀ KIỂM TRA GIỮA HỌC KÌ I MÔN TIN HỌC 10 (TIẾT PPCT: 16)', 'heading')
    add(b, '*(Quy ước: Phần I gồm 12 câu trắc nghiệm nhiều lựa chọn × 0,25 đ = 3,0 điểm; Phần II gồm 04 câu trắc nghiệm Đúng/Sai = 16 ý × 0,25 đ = 4,0 điểm; Phần III gồm 03 câu Tự luận = 3,0 điểm. Tổng cộng: 10,0 điểm).*', first=0, italic=True, keep=True)
    matran_rows = [
        ['TT', 'Chủ đề / Mạch nội dung\n(Số tiết PPCT)', 'Đơn vị kiến thức (Bài học trong SGK Tin học 10 KNTT)', 'Phần I: TN nhiều lựa chọn (12 câu = 3,0 đ)\n[NB / TH / VD]', 'Phần II: TN Đúng/Sai (4 câu = 16 ý = 4,0 đ)\n[NB / TH / VD]', 'Phần III: Tự luận (3 câu = 3,0 đ)\n[NB / TH / VD]', 'Tổng điểm theo bài\n(Tỉ lệ %)'],
        ['1', 'Chủ đề 1.\nMáy tính và xã hội tri thức\n(06 tiết:\nPPCT 01–06)', 'Mục 1.1. Bài 1: Thông tin và xử lí thông tin (02 tiết – tr. 6–10)', '2 NB (C1, C2)\n1 TH (C3)', 'Câu 13 (4 ý):\n2 ý NB (a, b)\n1 ý TH (c)\n1 ý VD (d)', 'Câu 17a (0,5 đ – TH)\nCâu 17b (0,5 đ – VD)', '2,75 điểm\n(27,5%)'],
        ['2', 'Chủ đề 1.\n(tiếp theo)', 'Mục 1.2. Bài 2: Vai trò của thiết bị thông minh và tin học đối với xã hội (02 tiết – tr. 11–15)', '1 NB (C4)\n1 TH (C5)', '—', '—', '0,50 điểm\n(5,0%)'],
        ['3', 'Chủ đề 1.\n(tiếp theo)', 'Mục 1.7. Bài 7: TH: Sử dụng thiết bị số thông dụng (02 tiết – tr. 33–37)', '1 NB (C6)', 'Câu 14 (2 ý đầu):\n1 ý NB (a)\n1 ý TH (b)', '—', '0,75 điểm\n(7,5%)'],
        ['4', 'Chủ đề 2.\nMạng máy tính và Internet\n(06 tiết:\nPPCT 07–12)', 'Mục 2.8. Bài 8: Mạng máy tính trong cuộc sống hiện đại (02 tiết – tr. 38–43)', '1 NB (C7)\n1 TH (C8)', 'Câu 14 (2 ý sau):\n1 ý TH (c)\n1 ý VD (d)', '—', '1,00 điểm\n(10,0%)'],
        ['5', 'Chủ đề 2.\n(tiếp theo)', 'Mục 2.9. Bài 9: An toàn trên không gian mạng (02 tiết – tr. 44–49)', '1 NB (C9)\n1 TH (C10)', 'Câu 15 (4 ý):\n2 ý NB (a, b)\n1 ý TH (c)\n1 ý VD (d)', 'Câu 18 (1,0 đ – VD)', '2,50 điểm\n(25,0%)'],
        ['6', 'Chủ đề 2.\n(tiếp theo)', 'Mục 2.10. Bài 10: TH: Khai thác tài nguyên trên không gian mạng (02 tiết – tr. 50–54)', '1 NB (C11)', '—', '—', '0,25 điểm\n(2,5%)'],
        ['7', 'Chủ đề 3.\nĐạo đức, pháp luật và văn hoá trong môi trường số\n(02 tiết: PPCT 13–14)', 'Mục 3.11. Bài 11: Ứng xử trên môi trường số, nghĩa vụ tôn trọng bản quyền (02 tiết – tr. 55–60)', '1 NB (C12)', 'Câu 16 (4 ý):\n2 ý NB (a, b)\n1 ý TH (c)\n1 ý VD (d)', 'Câu 19 (1,0 đ – VD)', '2,25 điểm\n(22,5%)'],
        ['', '**TỔNG CỘNG**', '**07 bài học (14 tiết thực học: Tiết PPCT 01–14)**', '**8 NB (2,0 đ)**\n**4 TH (1,0 đ)**\n**0 VD (0,0 đ)**', '**7 ý NB (1,75 đ)**\n**5 ý TH (1,25 đ)**\n**4 ý VD (1,00 đ)**', '**0,25 đ NB**\n**0,75 đ TH**\n**2,00 đ VD**', '**10,0 điểm**\n**(40% NB – 30% TH – 30% VD)**']
    ]
    b.append(table(matran_rows, [450, 1450, 2100, 1300, 1400, 1300, 1000], header=True, cell_sz=20))

    add(b, 'III. BẢN ĐẶC TẢ ĐỀ KIỂM TRA GIỮA HỌC KÌ I MÔN TIN HỌC 10', 'heading')
    dacta_rows = [
        ['TT', 'Bài học theo PPCT Trường THPT Thuận Hoá', 'Mức độ đánh giá và Yêu cầu cần đạt (Theo CT GDPT 2018 và SGK Tin học 10 KNTT)', 'Câu hỏi tương ứng trong đề'],
        ['1', 'Mục 1.1. Thông tin và xử lí thông tin\n(Tiết PPCT: 01, 02\nSGK Bài 1, tr. 6–10)',
         '- **Nhận biết:** Nêu được khái niệm thông tin, dữ liệu, vật mang tin; nhận biết được đơn vị dữ liệu nhỏ nhất mà máy tính truy cập được (Byte = 8 bit) và 3 bước xử lí thông tin (Hình 1.1).\n'
         '- **Thông hiểu:** Phân biệt được thông tin và dữ liệu qua ví dụ thực tế; giải thích được sự ưu việt của việc lưu trữ, xử lí và truyền thông tin bằng thiết bị số.\n'
         '- **Vận dụng:** Thực hiện được phép quy đổi giữa các đơn vị lưu trữ dữ liệu (B, KB, MB, GB, TB theo Bảng 1.1 trang 8) để tính toán dung lượng bộ nhớ và số lượng tệp lưu trữ.',
         '- Phần I: Câu 1 (NB), Câu 2 (NB), Câu 3 (TH)\n- Phần II: Câu 13 (a-NB, b-NB, c-TH, d-VD)\n- Phần III: Câu 17a (TH), Câu 17b (VD)'],
        ['2', 'Mục 1.2. Vai trò của thiết bị thông minh và tin học đối với xã hội\n(Tiết PPCT: 03, 04\nSGK Bài 2, tr. 11–15)',
         '- **Nhận biết:** Nhận diện được thiết bị thông minh và các thành tựu nổi bật của ngành Tin học (Hệ điều hành, Mạng Internet, Ngôn ngữ lập trình bậc cao, AI).\n'
         '- **Thông hiểu:** Giải thích được hai đặc trưng cơ bản của thiết bị thông minh (khả năng xử lí tự chủ và khả năng kết nối trao đổi dữ liệu) và đóng góp của Tin học đối với kinh tế, giáo dục, y tế.',
         '- Phần I: Câu 4 (NB), Câu 5 (TH)'],
        ['3', 'Mục 1.7. TH: Sử dụng thiết bị số thông dụng\n(Tiết PPCT: 05, 06\nSGK Bài 7, tr. 33–37)',
         '- **Nhận biết:** Nhận biết được các thông số kỹ thuật cơ bản của thiết bị số (CPU, RAM, ROM/Ổ cứng) và phần mở rộng của các loại tệp thông dụng.\n'
         '- **Thông hiểu:** Giải thích được ý nghĩa của từng thông số hệ thống và cách tổ chức cây thư mục hợp lí trên máy tính/điện thoại.',
         '- Phần I: Câu 6 (NB)\n- Phần II: Câu 14 (a-NB, b-TH)'],
        ['4', 'Mục 2.8. Mạng máy tính trong cuộc sống hiện đại\n(Tiết PPCT: 07, 08\nSGK Bài 8, tr. 38–43)',
         '- **Nhận biết:** Nêu được khái niệm mạng LAN, Internet, dịch vụ điện toán đám mây (Cloud) và Internet vạn vật (IoT).\n'
         '- **Thông hiểu:** Phân biệt được mạng LAN và Internet theo 3 tiêu chí tại Bảng trang 39; phân biệt được chức năng của Hub/Switch và Router.\n'
         '- **Vận dụng:** Lựa chọn được mô hình kết nối mạng và dịch vụ đám mây phù hợp với nhu cầu học tập, làm việc nhóm.',
         '- Phần I: Câu 7 (NB), Câu 8 (TH)\n- Phần II: Câu 14 (c-TH, d-VD)'],
        ['5', 'Mục 2.9. An toàn trên không gian mạng\n(Tiết PPCT: 09, 10\nSGK Bài 9, tr. 44–49)',
         '- **Nhận biết:** Kể tên được các nguy cơ trên không gian mạng (tin giả, lừa đảo, lộ thông tin cá nhân, bắt nạt qua mạng) và các loại phần mềm độc hại (Virus, Worm, Trojan).\n'
         '- **Thông hiểu:** Phân biệt được cơ chế lây nhiễm của Virus, Worm và Trojan; giải thích được tác dụng của tường lửa và phần mềm diệt virus.\n'
         '- **Vận dụng:** Đề xuất được quy trình xử lí cụ thể, an toàn khi gặp tình huống lừa đảo trực tuyến, tệp đính kèm nghi ngờ chứa mã độc hoặc bị bắt nạt trên không gian mạng.',
         '- Phần I: Câu 9 (NB), Câu 10 (TH)\n- Phần II: Câu 15 (a-NB, b-NB, c-TH, d-VD)\n- Phần III: Câu 18 (VD)'],
        ['6', 'Mục 2.10. TH: Khai thác tài nguyên trên không gian mạng\n(Tiết PPCT: 11, 12\nSGK Bài 10, tr. 50–54)',
         '- **Nhận biết:** Nhận biết được các công cụ tìm kiếm, dịch thuật trực tuyến, kho học liệu mở và các chế độ phân quyền chia sẻ tệp trên ổ đĩa đám mây (Google Drive).',
         '- Phần I: Câu 11 (NB)'],
        ['7', 'Mục 3.11. Ứng xử trên môi trường số, nghĩa vụ tôn trọng bản quyền\n(Tiết PPCT: 13, 14\nSGK Bài 11, tr. 55–60)',
         '- **Nhận biết:** Nhận diện được 6 hành vi vi phạm đạo đức, pháp luật và văn hoá khi sử dụng mạng (SGK trang 56); nêu được khái niệm quyền tác giả.\n'
         '- **Thông hiểu:** Phân biệt được Quyền nhân thân và Quyền tài sản trong quyền tác giả theo Luật Sở hữu trí tuệ (SGK trang 59); dẫn chiếu được Luật An ninh mạng 2018 và Nghị định 15/2020/NĐ-CP.\n'
         '- **Vận dụng:** Phân tích được tính hợp pháp của các hành vi chia sẻ thông tin, sử dụng phần mềm và đề xuất cách ứng xử văn minh, tôn trọng bản quyền sản phẩm số.',
         '- Phần I: Câu 12 (NB)\n- Phần II: Câu 16 (a-NB, b-NB, c-TH, d-VD)\n- Phần III: Câu 19 (VD)']
    ]
    b.append(table(dacta_rows, [450, 2150, 4400, 2000], header=True, cell_sz=21))

    br(b)
    add(b, 'IV. ĐỀ KIỂM TRA CHÍNH THỨC GIỮA HỌC KÌ I MÔN TIN HỌC 10 (TIẾT PPCT: 16)', 'heading')
    add(b, 'TRƯỜNG THPT THUẬN HOÁ – TỔ TIN HỌC', align='center', first=0, keep=True)
    add(b, '**ĐỀ KIỂM TRA ĐỊNH KÌ GIỮA HỌC KÌ I – NĂM HỌC 2026–2027**', align='center', first=0, keep=True)
    add(b, '**Môn: TIN HỌC – Khối 10 – Thời gian làm bài: 45 phút (Tiết PPCT: 16)**', align='center', first=0, keep=True)
    add(b, 'Họ và tên học sinh: ................................................................... Lớp: 10/...... Số báo danh: ...............', first=0, keep=True)

    add(b, '**PHẦN I. TRẮC NGHIỆM NHIỀU LỰA CHỌN** (3,0 điểm – Gồm 12 câu, từ Câu 1 đến Câu 12; mỗi câu đúng được 0,25 điểm)', 'heading')
    p1_questions = [
        '**Câu 1 (Mục 1.1 – NB):** Trong quá trình xử lí thông tin của máy tính (Hình 1.1 SGK trang 6), thứ tự đúng của 3 bước cơ bản là gì?\nA. Xử lí dữ liệu -> Tiếp nhận dữ liệu -> Đưa ra kết quả.\nB. Tiếp nhận dữ liệu -> Xử lí dữ liệu -> Đưa ra kết quả.\nC. Đưa ra kết quả -> Tiếp nhận dữ liệu -> Xử lí dữ liệu.\nD. Tiếp nhận dữ liệu -> Đưa ra kết quả -> Xử lí dữ liệu.',
        '**Câu 2 (Mục 1.1 – NB):** Theo Bảng 1.1 (SGK trang 8), đơn vị dữ liệu nhỏ nhất mà máy tính truy cập được là Byte (B). Một Byte gồm bao nhiêu bit?\nA. 2 bit.\nB. 4 bit.\nC. 8 bit.\nD. 1024 bit.',
        '**Câu 3 (Mục 1.1 – TH):** Bạn Minh nhìn tấm biển báo giao thông có hình chữ "P" kèm dấu gạch chéo màu đỏ và hiểu rằng đoạn đường này cấm đỗ xe. Trong tình huống trên, điều hiểu biết *"đoạn đường này cấm đỗ xe"* là gì?\nA. Là vật mang tin.\nB. Là thông tin đối với bạn Minh.\nC. Là dữ liệu dạng tệp văn bản trong máy tính.\nD. Là thiết bị tiếp nhận dữ liệu.',
        '**Câu 4 (Mục 1.2 – NB):** Thiết bị nào sau đây là thiết bị thông minh?\nA. Bóng đèn sợi đốt bật/tắt bằng công tắc cơ học.\nB. Máy tính bỏ túi thông thường không có kết nối mạng.\nC. Điện thoại thông minh (Smartphone) có hệ điều hành và kết nối Internet.\nD. Quạt bàn cơ học chỉnh 3 mức gió bằng nút bấm.',
        '**Câu 5 (Mục 1.2 – TH):** Phát biểu nào sau đây giải thích đúng nhất lí do thiết bị thông minh đóng vai trò quan trọng trong cuộc Cách mạng công nghiệp lần thứ tư (4.0)?\nA. Vì thiết bị thông minh luôn có kích thước rất lớn và tiêu thụ nhiều điện năng.\nB. Vì thiết bị thông minh có khả năng hoạt động tự chủ theo phần mềm và kết nối mạng (IoT) để thu thập, trao đổi, xử lí dữ liệu thông minh.\nC. Vì thiết bị thông minh hoàn toàn không cần sử dụng phần mềm máy tính.\nD. Vì thiết bị thông minh chỉ dùng để chơi trò chơi giải trí.',
        '**Câu 6 (Mục 1.7 – NB):** Khi quản lí tệp trên máy tính (Bài 7 SGK trang 35), tệp có tên Bao_cao_nhom1.docx thuộc loại tệp nào và phần mở rộng của tệp là gì?\nA. Tệp văn bản Microsoft Word, phần mở rộng là docx.\nB. Tệp hình ảnh đồ hoạ, phần mở rộng là Bao_cao.\nC. Tệp chương trình thực thi, phần mở rộng là nhom1.\nD. Tệp âm thanh số, phần mở rộng là doc.',
        '**Câu 7 (Mục 2.8 – NB):** Theo Bảng so sánh mạng cục bộ và Internet (SGK Bài 8 trang 39), thiết bị nào dưới đây được dùng để kết nối các mạng cục bộ (LAN) ở xa nhau lại với nhau tạo thành mạng Internet?\nA. Chuột quang máy tính.\nB. Bộ định tuyến (Router) thông qua các nhà cung cấp dịch vụ kết nối (ISP).\nC. Máy in văn phòng.\nD. Màn hình hiển thị.',
        '**Câu 8 (Mục 2.8 – TH):** Điểm khác biệt đúng về quyền sở hữu giữa mạng cục bộ (LAN) và mạng Internet là gì?\nA. Mạng LAN không có chủ sở hữu, còn Internet thuộc sở hữu của một công ty duy nhất.\nB. Cả mạng LAN và mạng Internet đều thuộc sở hữu riêng của một cá nhân.\nC. Mạng LAN thuộc quyền sở hữu của một cá nhân, gia đình hoặc cơ quan, tổ chức xác định; còn Internet không thuộc quyền sở hữu của riêng ai.\nD. Mạng LAN có quy mô toàn cầu, còn Internet chỉ dùng trong một phòng học.',
        '**Câu 9 (Mục 2.9 – NB):** Trong các phần mềm độc hại (Bài 9 SGK trang 46), loại mã độc nào lợi dụng các lỗ hổng bảo mật mạng để tự động nhân bản và lây lan từ máy tính này sang máy tính khác mà không cần người dùng kích hoạt tệp?\nA. Sâu máy tính (Worm).\nB. Phần mềm soạn thảo văn bản.\nC. Trình duyệt web.\nD. Hệ điều hành Windows.',
        '**Câu 10 (Mục 2.9 – TH):** Biện pháp nào sau đây giúp bảo vệ thông tin cá nhân **an toàn nhất** khi sử dụng mạng máy tính (theo Hộp kiến thức SGK Bài 9 trang 44)?\nA. Ghi chép mật khẩu tài khoản ngân hàng và email lên mặt bàn học ở lớp để khỏi quên.\nB. Sử dụng mạng Wi-Fi công cộng không mật khẩu để đăng nhập tài khoản ngân hàng trực tuyến.\nC. Không ghi chép thông tin cá nhân ở nơi người khác dễ đọc, giữ máy tính không nhiễm phần mềm gián điệp và cẩn trọng khi dùng Wi-Fi công cộng.\nD. Chia sẻ mật khẩu tài khoản mạng xã hội cho tất cả bạn bè trên mạng.',
        '**Câu 11 (Mục 2.10 – NB):** Khi sử dụng ổ đĩa đám mây Google Drive để nộp bài tập thực hành cho giáo viên (Bài 10 SGK trang 52), thao tác nào giúp em gửi liên kết tệp cho thầy/cô xem được bài làm?\nA. Nháy nút phải chuột vào tệp -> chọn Chia sẻ (Share) -> thiết lập quyền truy cập và sao chép đường liên kết gửi cho giáo viên.\nB. Đổi phần mở rộng của tệp thành .exe.\nC. Tắt kết nối mạng Internet của máy tính trước khi tải tệp lên.\nD. Xoá tệp khỏi thùng rác của máy tính.',
        '**Câu 12 (Mục 3.11 – NB):** Theo Hộp kiến thức Bài 11 (SGK trang 56), hành vi nào sau đây là hành vi vi phạm đạo đức, pháp luật và văn hoá khi sử dụng mạng?\nA. Tự giác ghi rõ tên tác giả và nguồn trích dẫn khi sử dụng tư liệu trên mạng.\nB. Công bố thông tin cá nhân của người khác mà không được phép và phát tán tin nhắn rác (Spam).\nC. Sử dụng phần mềm nguồn mở miễn phí phục vụ học tập.\nD. Tuyên truyền kiến thức phòng chống lừa đảo trên không gian mạng.'
    ]
    for q in p1_questions:
        add(b, q, first=0)

    add(b, '**PHẦN II. TRẮC NGHIỆM ĐÚNG / SAI** (4,0 điểm – Gồm 04 câu, từ Câu 13 đến Câu 16; mỗi câu có 4 ý a, b, c, d, học sinh chọn Đúng [Đ] hoặc Sai [S] cho từng ý)', 'heading')
    p2_questions = [
        '**Câu 13 (Mục 1.1 – Bài 1: Thông tin và xử lí thông tin):** Nhóm học sinh lớp 10 thực hiện dự án làm kỷ yếu số của lớp. Các bạn dùng máy ảnh số chụp ảnh, chép vào máy tính để chỉnh sửa rồi lưu vào một ổ cứng di động còn trống 4 GB. Xét tính Đúng (Đ) hoặc Sai (S) của các phát biểu sau:\n'
        'a) Các tệp ảnh kỷ yếu lưu trong ổ cứng di động là dữ liệu, còn ổ cứng di động là vật mang tin.\n'
        'b) Theo Bảng 1.1 (SGK trang 8), hai đơn vị lưu trữ dữ liệu liên tiếp nhau từ Byte trở lên (B, KB, MB, GB, TB) hơn kém nhau đúng 1000 lần.\n'
        'c) Việc lưu trữ và gửi bộ ảnh kỷ yếu số qua mạng Internet giúp chia sẻ nhanh chóng cho cả lớp mà không làm giảm chất lượng bản gốc khi sao chép.\n'
        'd) Nếu nhóm lưu 01 video tổng kết lớp nặng 2 GB và các bức ảnh kỷ yếu mà mỗi bức ảnh nặng 4 MB, thì phần dung lượng còn lại của ổ cứng 4 GB lưu được tối đa 512 bức ảnh.',
        '**Câu 14 (Mục 1.7 và Mục 2.8 – Bài 7 và Bài 8):** Phòng Tin học của Trường THPT Thuận Hoá có 40 máy tính để bàn được kết nối với nhau qua bộ chuyển mạch (Switch) và nối với bộ định tuyến (Router) của nhà trường để truy cập Internet. Cấu hình mỗi máy tính ghi: CPU Core i5 3.0 GHz, RAM 8 GB, SSD 256 GB. Xét tính Đúng (Đ) hoặc Sai (S) của các phát biểu sau:\n'
        'a) Thông số SSD 256 GB cho biết dung lượng ổ cứng thể rắn dùng để lưu trữ lâu dài hệ điều hành, phần mềm và tệp dữ liệu của người dùng.\n'
        'b) Khi học sinh tắt nguồn máy tính, toàn bộ dữ liệu đang lưu trong ổ cứng SSD 256 GB sẽ bị mất trắng giống như bộ nhớ RAM.\n'
        'c) Hệ thống 40 máy tính kết nối qua Switch trong phòng máy tạo thành một mạng cục bộ (LAN) thuộc quyền sở hữu của nhà trường.\n'
        'd) Nếu cáp quang từ Router ra nhà cung cấp dịch vụ ISP bị đứt, các máy tính trong phòng học vẫn có thể chia sẻ tệp cho nhau trong nội bộ mạng LAN qua Switch.',
        '**Câu 15 (Mục 2.9 – Bài 9: An toàn trên không gian mạng):** Bạn Lan nhận được một thư điện tử lạ thông báo *"Chúc mừng bạn trúng học bổng 50 triệu đồng"*, bên trong yêu cầu Lan nhập mật khẩu email, mã OTP ngân hàng của bố mẹ và tải tệp đính kèm Nhan_thuong_ngay.exe về máy tính để xác nhận. Xét tính Đúng (Đ) hoặc Sai (S) của các phát biểu sau:\n'
        'a) Đây là dấu hiệu điển hình của hình thức lừa đảo trực tuyến (Phishing) kết hợp phát tán phần mềm độc hại trên không gian mạng.\n'
        'b) Tệp Nhan_thuong_ngay.exe là tệp văn bản thuần tuý nên hoàn toàn vô hại khi mở trên máy tính.\n'
        'c) Nếu Lan tải và chạy tệp .exe đó, máy tính có nguy cơ bị nhiễm phần mềm gián điệp (Spyware/Trojan) hoặc mã độc tống tiền (Ransomware).\n'
        'd) Cách xử lí đúng là Lan tuyệt đối không mở tệp đính kèm, không cung cấp mật khẩu/mã OTP, đánh dấu thư rác/lừa đảo và báo cho bố mẹ hoặc thầy cô biết.',
        '**Câu 16 (Mục 3.11 – Bài 11: Ứng xử trên môi trường số, nghĩa vụ tôn trọng bản quyền):** Hoà là người viết ra một chương trình máy tính hỗ trợ học từ vựng tiếng Anh. Sau đó, Hoà chuyển giao quyền tài sản của chương trình này cho Công ty Phần mềm X theo hợp đồng hợp pháp. Xét tính Đúng (Đ) hoặc Sai (S) của các phát biểu sau theo Luật Sở hữu trí tuệ (SGK trang 59–60):\n'
        'a) Quyền được đứng tên thật hoặc bút danh tác giả trên chương trình máy tính thuộc về Quyền nhân thân của Hoà.\n'
        'b) Quyền sao chép, phân phối và cho thuê bản sao chương trình máy tính thuộc về Quyền tài sản.\n'
        'c) Sau khi đã mua quyền tài sản, Công ty X có toàn quyền xoá tên tác giả Hoà khỏi chương trình và tự nhận mình là người sáng tạo ra tác phẩm gốc.\n'
        'd) Nếu một học sinh tự ý dùng công cụ bẻ khoá (crack) phần mềm thương mại của Công ty X rồi đăng lên mạng cho nhiều người tải miễn phí thì học sinh đó đã vi phạm bản quyền.'
    ]
    for q in p2_questions:
        add(b, q, first=0)

    add(b, '**PHẦN III. TỰ LUẬN** (3,0 điểm – Gồm 03 câu, từ Câu 17 đến Câu 19)', 'heading')
    p3_questions = [
        '**Câu 17 (1,0 điểm – Mục 1.1: Thông tin và xử lí thông tin, SGK Bài 1 trang 6–10):**\n'
        'a) *(0,5 điểm – Thông hiểu)* Em hãy phân biệt khái niệm **Thông tin** và **Dữ liệu**, đồng thời cho 01 ví dụ minh hoạ trong học tập có chỉ rõ đâu là dữ liệu, đâu là thông tin.\n'
        'b) *(0,5 điểm – Vận dụng)* Một thiết bị lưu trữ USB có dung lượng trống thực tế là 8 GB. Em cần sao chép một thư mục tài liệu học tập gồm 01 tệp video bài giảng nặng 3 GB và 1200 tệp ảnh tư liệu (mỗi tệp ảnh có dung lượng 4 MB). Hãy tính toán cụ thể (theo Bảng 1.1 SGK trang 8) để cho biết chiếc USB đó có đủ dung lượng để chứa toàn bộ thư mục trên hay không? Vì sao?',
        '**Câu 18 (1,0 điểm – Mục 2.9: An toàn trên không gian mạng, SGK Bài 9 trang 44–49):**\n'
        'Khi sử dụng mạng xã hội, một người bạn cùng lớp của em bị kẻ xấu lập tài khoản giả mạo, cắt ghép hình ảnh sai sự thật để bắt nạt, xúc phạm và nhắn tin đe doạ. Dựa vào kiến thức Bài 9 (SGK trang 45–48), em hãy nêu rõ **04 việc cần làm (và không nên làm)** để tư vấn giúp bạn mình phòng vệ và giải quyết sự việc một cách an toàn, hiệu quả.',
        '**Câu 19 (1,0 điểm – Mục 3.11: Ứng xử trên môi trường số, nghĩa vụ tôn trọng bản quyền, SGK Bài 11 trang 55–60):**\n'
        'Để chuẩn bị cho Chủ đề 4 (Thiết kế đồ hoạ trên máy tính), bạn Nam định lên mạng tìm tải bản bẻ khoá (crack) của một phần mềm đồ hoạ thương mại đắt tiền về cài vào máy tính cá nhân. Dựa vào kiến thức Bài 11 (SGK trang 59–60), em hãy:\n'
        'a) *(0,5 điểm)* Giải thích vì sao hành vi tải và sử dụng phần mềm bẻ khoá (crack) là không đúng pháp luật và tiềm ẩn mất an toàn thông tin.\n'
        'b) *(0,5 điểm)* Đề xuất giải pháp lựa chọn phần mềm hợp pháp, miễn phí giúp bạn Nam hoàn thành tốt nhiệm vụ học tập môn Tin học 10.'
    ]
    for q in p3_questions:
        add(b, q, first=0)

    add(b, 'V. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM CHI TIẾT ĐỀ KIỂM TRA GIỮA HỌC KÌ I (TIẾT PPCT: 16)', 'heading')
    add(b, '**1. Đáp án Phần I – Trắc nghiệm nhiều lựa chọn** (3,0 điểm: Mỗi câu đúng được 0,25 điểm)', 'heading')
    ans_p1 = [
        ['Câu 1', 'Câu 2', 'Câu 3', 'Câu 4', 'Câu 5', 'Câu 6', 'Câu 7', 'Câu 8', 'Câu 9', 'Câu 10', 'Câu 11', 'Câu 12'],
        ['**B**', '**C**', '**B**', '**C**', '**B**', '**A**', '**B**', '**C**', '**A**', '**C**', '**A**', '**B**'],
        ['tr. 6', 'tr. 8', 'tr. 7', 'tr. 11', 'tr. 12', 'tr. 35', 'tr. 39', 'tr. 39', 'tr. 46', 'tr. 44', 'tr. 52', 'tr. 56']
    ]
    b.append(table(ans_p1, [750]*12, header=True, cell_sz=21))

    add(b, '**2. Đáp án Phần II – Trắc nghiệm Đúng / Sai** (4,0 điểm: Mỗi ý đúng được 0,25 điểm)', 'heading')
    ans_p2 = [
        ['Câu hỏi và Căn cứ SGK', 'Ý a)', 'Ý b)', 'Ý c)', 'Ý d)', 'Giải thích căn cứ Sách giáo khoa Tin học 10 KNTT'],
        ['**Câu 13**\n(Mục 1.1 – Bài 1, tr. 6–10)', '**Đúng**', '**Sai**', '**Đúng**', '**Đúng**', '- Ý b **Sai** vì theo Bảng 1.1 (SGK trang 8), các đơn vị hơn kém nhau 2¹⁰ = 1024 lần.\n- Ý d **Đúng** vì 4 GB - 2 GB = 2 GB = 2048 MB; số ảnh lưu được là 2048 : 4 = 512 bức ảnh.'],
        ['**Câu 14**\n(Mục 1.7 và Mục 2.8 – Bài 7, 8)', '**Đúng**', '**Sai**', '**Đúng**', '**Đúng**', '- Ý b **Sai** vì ổ cứng SSD là bộ nhớ ngoài lưu trữ lâu dài, không bị mất dữ liệu khi tắt nguồn.\n- Ý d **Đúng** vì Switch vẫn duy trì kết nối nội bộ mạng LAN (SGK trang 39).'],
        ['**Câu 15**\n(Mục 2.9 – Bài 9, tr. 44–49)', '**Đúng**', '**Sai**', '**Đúng**', '**Đúng**', '- Ý b **Sai** vì phần mở rộng .exe là tệp chương trình thực thi, thường bị kẻ xấu lợi dụng để phát tán Virus, Trojan, Ransomware.'],
        ['**Câu 16**\n(Mục 3.11 – Bài 11, tr. 55–60)', '**Đúng**', '**Đúng**', '**Sai**', '**Đúng**', '- Ý c **Sai** vì đứng tên tác giả và bảo vệ sự toàn vẹn của tác phẩm là **Quyền nhân thân** gắn liền với tác giả Hoà (SGK trang 59).']
    ]
    b.append(table(ans_p2, [1500, 750, 750, 750, 750, 4500], header=True, cell_sz=21))

    add(b, '**3. Hướng dẫn chấm Phần III – Tự luận** (3,0 điểm)', 'heading')
    ans_p3 = [
        ['Câu', 'Nội dung đáp án và căn cứ chấm chi tiết (Nói có sách, mách có chứng)', 'Điểm'],
        ['**Câu 17**\n(1,0 điểm)\nMục 1.1 – Bài 1\n(SGK tr. 6–8)',
         '- **Ý a (0,5 điểm):**\n'
         '  + Nêu đúng: **Thông tin** là tất cả những gì mang lại hiểu biết cho con người; **Dữ liệu** là thông tin đã được đưa vào máy tính dưới dạng văn bản, số, hình ảnh, âm thanh *(0,25 điểm)*.\n'
         '  + Lấy được 01 ví dụ đúng và chỉ rõ đâu là dữ liệu (ví dụ: bảng điểm số 8, 9, 10 nhập vào máy tính) và đâu là thông tin (kết luận học sinh đạt danh hiệu Học sinh Giỏi) *(0,25 điểm)*.\n'
         '- **Ý b (0,5 điểm):**\n'
         '  + Tính đúng dung lượng trống còn lại của USB sau khi chép tệp video 3 GB: 8 GB - 3 GB = 5 GB; đổi sang MB theo Bảng 1.1 (SGK trang 8): 5 × 1024 = 5120 MB *(0,25 điểm)*.\n'
         '  + Tính tổng dung lượng của 1200 tệp ảnh: 1200 × 4 MB = 4800 MB. Vì 4800 MB < 5120 MB (còn dư 320 MB), nên kết luận chiếc USB 8 GB **đủ dung lượng** để chứa toàn bộ thư mục *(0,25 điểm)*.',
         '**1,0 đ**'],
        ['**Câu 18**\n(1,0 điểm)\nMục 2.9 – Bài 9\n(SGK tr. 45)',
         '- Nêu đủ 04 biện pháp xử lí đúng theo mục b) Bắt nạt qua mạng (SGK Bài 9 trang 45) và Bài 11 (trang 56–58) *(mỗi ý đúng được 0,25 điểm)*:\n'
         '  1. **Giữ bình tĩnh và không đáp trả:** Tuyệt đối không chửi bới, đôi co trên mạng và không làm theo các yêu cầu đe doạ của kẻ xấu *(0,25 điểm)*.\n'
         '  2. **Lưu giữ bằng chứng:** Chụp ảnh màn hình các tin nhắn đe doạ, bài đăng vu khống và lưu lại đường dẫn tài khoản giả mạo *(0,25 điểm)*.\n'
         '  3. **Sử dụng công cụ bảo vệ trên mạng:** Chặn (Block) tài khoản quấy rối và gửi báo cáo vi phạm (Report) tới ban quản trị mạng xã hội để gỡ bỏ nội dung giả mạo *(0,25 điểm)*.\n'
         '  4. **Chia sẻ và nhờ can thiệp kịp thời:** Báo ngay cho cha mẹ, thầy cô giáo chủ nhiệm hoặc cơ quan công an khi sự việc nghiêm trọng (vì hành vi làm nhục, vu khống người khác vi phạm Luật An ninh mạng năm 2018 và Nghị định 15/2020/NĐ-CP) *(0,25 điểm)*.',
         '**1,0 đ**'],
        ['**Câu 19**\n(1,0 điểm)\nMục 3.11 – Bài 11\n(SGK tr. 59–60)',
         '- **Ý a (0,5 điểm):**\n'
         '  + Về pháp lí: Sử dụng phần mềm bẻ khoá (crack) là hành vi xâm phạm **Quyền tài sản** của chủ sở hữu quyền tác giả theo Luật Sở hữu trí tuệ (SGK trang 59–60) *(0,25 điểm)*.\n'
         '  + Về an toàn thông tin: Các tệp bẻ khoá (crack/patch) không rõ nguồn gốc thường bị tin tặc cài cắm sẵn phần mềm độc hại (Trojan, Spyware, Ransomware) gây mất dữ liệu và mất quyền kiểm soát máy tính *(0,25 điểm)*.\n'
         '- **Ý b (0,5 điểm):**\n'
         '  + Tư vấn bạn Nam sử dụng phần mềm nguồn mở / phần mềm miễn phí hợp pháp tương đương, cụ thể là phần mềm thiết kế đồ hoạ vectơ **Inkscape** (hoặc GIMP cho ảnh bitmap) được cung cấp miễn phí tại trang chủ chính thức https://inkscape.org *(0,50 điểm)*.',
         '**1,0 đ**']
    ]
    b.append(table(ans_p3, [1200, 6800, 1000], header=True, cell_sz=21))

    add_approval_section(b, sign_date_str='Huế, ngày 02 tháng 11 năm 2026', section_title='VI. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT')
    out_matran = KHOI10_DIR / '02_Ma_tran_Ban_dac_ta_De_kiem_tra_GHKI_Tin10_TietPPCT16.docx'
    save_docx(b, out_matran, used_figures=None)
    shutil.copy2(out_matran, THANH_PHAM_DIR / 'Ma-tran-Ban-dac-ta-De-KT-GK1_Tin10.docx')
    shutil.copy2(out_matran, GOI_NOP_DIR / '02_Ma_tran_Ban_dac_ta_De_KT_GK1_Tin10.docx')
    print('Built:', out_matran)

if __name__ == '__main__':
    build_ontap_ghki()
    build_matran_ghki()