# -*- coding: utf-8 -*-
"""
Builds all deliverables for Grade 12 Bài 10: Tạo liên kết (SGK Tin học 12 KNTT, trang 57–61; Tiết PPCT: 21, 22):
1. 01_De_cuong_du_gio_Tin12_Bai10_Tiet1_LT_TietPPCT21_Lop12-2.docx
   (Dự giờ Thứ Sáu, ngày 13/11/2026 – Tiết 3, 09h00 – 09h45, Lớp 12/2, Phòng H.II.2 – Tiết 1: Lí thuyết)
2. 02_De_cuong_du_gio_Tin12_Bai10_Tiet2_TH_TietPPCT22_Lop12-1.docx
   (Dự giờ Thứ Sáu, ngày 13/11/2026 – Tiết 2, 08h00 – 08h45, Lớp 12/1, Phòng H.II.1 – Tiết 2: Thực hành)
3. 03_De_cuong_du_gio_Tin12_Bai10_CaBai_2Tiet_TietPPCT21-22.docx
   (Bản gộp trọn vẹn cả Bài 10 – 02 tiết = 90 phút, Tiết PPCT: 21, 22)
4. 04_Phieu_hoc_tap_va_Rubric_Tin12_Bai10.docx
   (Bộ Phiếu học tập số 1, 2, 3, 4 và Rubric CCĐG 6 rời để in phát tại lớp)
"""
from pathlib import Path
import shutil, zipfile
from docx_engine_5512 import (
    init_body_with_cover, add, br, table, add_figure, add_approval_section, save_docx, node, tabline
)

KHOI12_DIR = Path(r'c:\Users\PC\khbd_docx\khoi12')
THANH_PHAM_DIR = Path(r'c:\Users\PC\khbd_docx\06_thanh_pham_docx')
GOI_NOP_DIR = Path(r'c:\Users\PC\khbd_docx\08_goi_nop_du_gio\HoQuangHuy_24S1020030_HoSo_DuGio_Tin12_Bai10')

GVGD_NAME = 'Thầy Nguyễn Nhất Tân'
GVHD_NAME = 'Thầy Mai Xuân Văn'
SV_NAME = 'Hồ Quang Huy'
SIGN_DATE = 'Huế, ngày 12 tháng 11 năm 2026'

# ==============================================================================
# SHARED STEP TABLES FOR BÀI 10: TẠO LIÊN KẾT (SGK TIN HỌC 12 KNTT, TRANG 57–61)
# ==============================================================================

HD1_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu lên màn hình một trang báo điện tử (hoặc trang web giới thiệu Câu lạc bộ đã học ở Bài 9) đặt cạnh hình ảnh một trang sách, báo giấy truyền thống và nêu câu hỏi Mở đầu (SGK Tin học 12 trang 57):\n'
        '  + Câu hỏi 1: Em hãy nêu những điểm khác biệt khi em đọc thông tin trên các trang web với việc em đọc sách, báo giấy.\n'
        '  + Câu hỏi 2: Theo em, điểm khác biệt nào là quan trọng nhất giúp người đọc di chuyển tức thời giữa các trang thông tin?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh suy nghĩ cá nhân 01 phút, sau đó trao đổi nhanh với bạn cùng bàn trong 02 phút để liệt kê các điểm khác biệt giữa trang web và sách, báo giấy.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh đại diện trả lời tại chỗ; mời 01 học sinh khác nhận xét, bổ sung.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá theo bảng kiểm CCĐG 1, biểu dương câu trả lời đúng và dẫn dắt vào bài mới: *"Khi đọc sách báo giấy, ta phải lật tuần tự từng trang. Ngược lại, trên trang web có các vùng văn bản hoặc hình ảnh gắn liên kết giúp người dùng nháy chuột chuyển ngay tới nội dung mong muốn ở vị trí khác hoặc trang web khác. Bài 10 sẽ giúp các em nắm vững bản chất của siêu văn bản, đường dẫn và sử dụng thẻ HTML để tạo các loại liên kết."*',
        '**Nhận xét mở đầu (SGK Bài 10, trang 57):**\n'
        '- **Điểm khác biệt khi đọc trang web so với sách, báo giấy:**\n'
        '  + Trang web không chỉ có văn bản và hình ảnh tĩnh mà còn tích hợp âm thanh, video, hoạt ảnh và khả năng tương tác.\n'
        '  + Cách tiếp cận thông tin trên sách, báo giấy chủ yếu là tuần tự (lật từng trang), còn trên trang web là phi tuyến tính thông qua các liên kết.\n'
        '- **Điểm khác biệt quan trọng nhất:** Trang web có các siêu liên kết (hyperlink) giúp người đọc chỉ cần nháy chuột là di chuyển nhanh chóng từ một nội dung thông tin này đến một nội dung thông tin khác trong cùng trang, cùng website hoặc sang website khác trên Internet.'
    ]
]

NV1_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên phát Phiếu học tập số 1 và yêu cầu học sinh làm việc theo cặp đôi (06 phút), nghiên cứu mục 1. Siêu văn bản và đường dẫn (SGK trang 57–58) để thực hiện 03 yêu cầu:\n'
        '  + Yêu cầu 1 (Hoạt động 1 SGK trang 57): Các tệp có phần mở rộng .docx được tạo bởi Microsoft Word có thể là một siêu văn bản còn tệp có phần mở rộng .txt tạo bởi Notepad thì không. Theo em, siêu văn bản có những đặc điểm gì?\n'
        '  + Yêu cầu 2: Trình bày khái niệm siêu liên kết, cú pháp thẻ <a> tạo liên kết trong HTML và phân biệt Đường dẫn tuyệt đối với Đường dẫn tương đối.\n'
        '  + Yêu cầu 3 (Câu hỏi củng cố SGK trang 58): Trong ba đường dẫn sau, đường dẫn nào là đường dẫn tương đối, đường dẫn nào là đường dẫn tuyệt đối: a) html/cach_tao_lien_ket.html; b) http://google.com; c) mail.google.com/mail/u/0/#inbox/FMfcgzGrcjHzMWKllZqRrT.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đọc kĩ SGK trang 57–58, thảo luận nhóm đôi và hoàn thành Phiếu học tập số 1. Giáo viên quan sát, gợi ý học sinh chú ý dấu hiệu nhận biết giao thức (http:// hoặc https://) ở đầu đường dẫn.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời đại diện 02 nhóm trình bày kết quả Phiếu học tập số 1; các nhóm khác nhận xét, bổ sung.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo CCĐG 2, chuẩn hoá kiến thức mục 1 và nhấn mạnh hộp ghi nhớ SGK trang 58.',
        '**1. Siêu văn bản và đường dẫn (SGK trang 57–58)**\n'
        '- **Giải quyết Hoạt động 1 (SGK trang 57):** Tệp .txt chỉ chứa văn bản thuần tuý (plain text), còn tệp .docx hoặc trang web (.html) là siêu văn bản vì ngoài văn bản còn chứa được hình ảnh, âm thanh, bảng biểu và đặc biệt là các siêu liên kết.\n'
        '- **Khái niệm siêu văn bản:** Siêu văn bản (hypertext) là loại văn bản mà nội dung của nó không chỉ chứa văn bản mà còn có thể chứa nhiều dạng dữ liệu khác như âm thanh, hình ảnh,... và đặc biệt là chứa các siêu liên kết tới siêu văn bản khác.\n'
        '- **Khái niệm siêu liên kết và thẻ <a>:**\n'
        '  + Siêu liên kết (hyperlink, gọi tắt là liên kết) là một tham chiếu đến một tài liệu nằm tại một vị trí khác (trong cùng trang web, cùng website hoặc thuộc website khác).\n'
        '  + Cú pháp tạo liên kết trong HTML:\n'
        '    <a href="URL">Nội dung hiển thị tại vị trí đặt liên kết</a>\n'
        '  + Trong đó: thuộc tính href cung cấp đường dẫn (URL) tới tài liệu/vị trí đích; nội dung phần tử hiển thị trên màn hình (mặc định có màu xanh và gạch chân).\n'
        '- **Hai loại đường dẫn trong thuộc tính href:**\n'
        '  + **Đường dẫn tuyệt đối:** Bao gồm đầy đủ giao thức (thường là http:// hoặc https://), tên miền (tên website) và đường dẫn chi tiết đến trang web. Ví dụ: https://google.com.\n'
        '  + **Đường dẫn tương đối:** Dùng khi liên kết tới tài liệu khác trên cùng trang web hoặc cùng một website (cùng máy chủ); không yêu cầu giao thức hay tên miền mà chỉ cần tên đường dẫn tương đối từ tệp hiện tại. Ví dụ: bai_tap_1.html, tailieu/kehoach.pdf.\n'
        '- **Trả lời Câu hỏi củng cố SGK trang 58:**\n'
        '  + a) html/cach_tao_lien_ket.html: Đường dẫn tương đối (chỉ gồm thư mục con và tên tệp, không có giao thức và tên miền).\n'
        '  + b) http://google.com: Đường dẫn tuyệt đối (có đầy đủ giao thức http:// và tên miền google.com).\n'
        '  + c) mail.google.com/mail/u/0/#inbox/FMfcgzGrcjHzMWKllZqRrT: Đường dẫn tương đối theo cú pháp HTML (do khuyết giao thức https:// ở đầu; nếu gõ trực tiếp lên thanh địa chỉ trình duyệt thì trình duyệt tự thêm https:// để hiểu là địa chỉ web).'
    ]
]

NV2_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên nêu câu hỏi Hoạt động 2 (SGK trang 58): *"Hãy kể tên các kiểu liên kết mà em đã gặp khi duyệt web hoặc khi đọc văn bản"*, sau đó giao Phiếu học tập số 2 yêu cầu học sinh nghiên cứu mục 2a, 2b và Hình 10.1 (SGK trang 58–60) trong 06 phút:\n'
        '  + Yêu cầu 1: Phân biệt Liên kết ngoài (liên kết tới một trang web khác trên Internet) và Liên kết trong (liên kết tới vị trí khác trong cùng website); viết câu lệnh HTML tạo liên kết tới trang Hành trang số của NXB Giáo dục Việt Nam.\n'
        '  + Yêu cầu 2: Quan sát Hình 10.1 – Cấu trúc một website (SGK trang 59), viết đoạn mã HTML cho 4 trường hợp liên kết tương đối trong cùng website: (1) Từ index.html tới thong_tin.html (cùng thư mục); (2) Từ index.html tới bai_tap_1.html (thư mục khác dưới 1 cấp); (3) Từ index.html tới bai_tap_on_tap.html (thư mục khác dưới 2 cấp); (4) Từ bai_tap_1.html quay lại index.html ở thư mục mẹ web/ (thư mục mức trên).\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh quan sát kĩ sơ đồ cây thư mục ở Hình 10.1 (SGK trang 59), thảo luận cặp đôi để xác định đường đi từ tệp nguồn tới tệp đích và viết đúng chuỗi href vào Phiếu học tập số 2.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 nhóm lên bảng viết 04 câu lệnh thẻ <a> tương ứng với 4 trường hợp của Hình 10.1; các nhóm khác đối chiếu, nhận xét.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo CCĐG 3 và chốt quy tắc: Xuống thư mục con dùng dấu gạch chéo "/" (tên_thư_mục/tên_tệp); ngược lên thư mục mẹ một mức dùng kí tự "../" (số cụm "../" bằng đúng số mức thư mục cần quay lên).',
        '**2. Các cách tạo liên kết tới một trang web (SGK trang 58–60)**\n'
        '**a) Liên kết tới một trang web khác (Liên kết ngoài):**\n'
        '- Sử dụng đường dẫn tuyệt đối trong thuộc tính href để dẫn tới một trang web khác trên Internet.\n'
        '- Ví dụ (SGK trang 58):\n'
        '  <p><a href="https://hanhtrangso.nxbgd.vn/sach-dien-tu?book_active=0">Sách điện tử Kết nối tri thức với cuộc sống</a></p>\n'
        '**b) Liên kết đến một vị trí khác trong cùng website (Liên kết trong – Hình 10.1, trang 59–60):**\n'
        '- Sử dụng đường dẫn tương đối dựa theo cấu trúc thư mục của website (Hình 10.1: thư mục gốc web/ chứa thư mục images/, tệp index.html, tệp thong_tin.html và thư mục bai_tap/; trong bai_tap/ có bai_tap_1.html, bai_tap_2.html và thư mục con on_tap/ chứa bai_tap_on_tap.html):\n'
        '  + **Trường hợp 1 – Liên kết tới trang web cùng thư mục:** Chỉ cần cung cấp tên tệp đích. Ví dụ từ index.html tới thong_tin.html:\n'
        '    <a href="thong_tin.html">Giới thiệu về trang web</a>\n'
        '  + **Trường hợp 2 – Liên kết tới trang web thuộc thư mục khác, dưới một cấp:** Ghép tên thư mục con và tên tệp phân cách bởi dấu "/". Ví dụ từ index.html tới bai_tap_1.html:\n'
        '    <a href="bai_tap/bai_tap_1.html">Bài tập 1</a>\n'
        '  + **Trường hợp 3 – Liên kết tới trang web thuộc thư mục khác, dưới hai (hay nhiều) cấp:** Liệt kê lần lượt các cấp thư mục từ trên xuống và tên tệp phân cách bởi dấu "/". Ví dụ từ index.html tới bai_tap_on_tap.html:\n'
        '    <a href="bai_tap/on_tap/bai_tap_on_tap.html">Bài tập ôn tập</a>\n'
        '  + **Trường hợp 4 – Liên kết tới trang web nằm ở thư mục mức trên:** Sử dụng cụm kí tự "../" để trở lại thư mục mức trên một cấp. Ví dụ từ bai_tap_1.html (trong thư mục bai_tap) quay lại index.html (trong thư mục mẹ web):\n'
        '    <a href="../index.html">Quay lại trang chủ</a>'
    ]
]

NV3_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên giao Phiếu học tập số 3 yêu cầu học sinh đọc phần tiếp theo của mục 2b (SGK trang 60) trong 06 phút để trả lời 03 câu hỏi:\n'
        '  + Câu 1: Nêu quy trình 2 bước để tạo liên kết tới một vị trí khác trong cùng trang web (liên kết nội bộ trang). Thuộc tính id có vai trò gì và cú pháp href khi liên kết tới id đó viết như thế nào?\n'
        '  + Câu 2: Làm thế nào để sử dụng một hình ảnh (ví dụ ảnh Mặt Trời images/sun.png trong Hình 10.1) làm vị trí đặt liên kết dẫn tới trang thong_tin.html?\n'
        '  + Câu 3 (Câu hỏi củng cố SGK trang 60): Viết một đoạn mã HTML để hiển thị một danh sách không có thứ tự trong tệp index.html gồm ba mục: bài tập 1, bài tập 2 và ôn tập; ba mục này liên kết tới ba trang tương ứng với ba tệp bai_tap_1.html, bai_tap_2.html và bai_tap_on_tap.html.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đọc SGK trang 60, kết hợp kiến thức tạo danh sách <ul><li> đã học ở Bài 9 với thẻ <a> của Bài 10 để hoàn thành Phiếu học tập số 3.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 01 học sinh trình bày quy trình 2 bước tạo mã định danh đoạn (id) và 01 học sinh lên bảng viết đoạn mã danh sách liên kết của Câu hỏi trang 60.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá theo CCĐG 4, chốt hộp kiến thức ghi nhớ trang 60 và lưu ý: Tên id (mã định danh đoạn) phải là duy nhất trong toàn bộ trang web và khi gọi trong href bắt buộc phải có dấu thăng "#" đứng trước (#mã_định_danh_đoạn).',
        '**c) Liên kết tới vị trí khác trong cùng trang web và liên kết hình ảnh (SGK trang 60)**\n'
        '- **Liên kết tới vị trí khác trong cùng trang web (2 bước):**\n'
        '  + **Bước 1:** Tạo phần tử HTML có thuộc tính id nằm tại vị trí đích cần liên kết đến (thuộc tính id được gọi là mã định danh đoạn và chỉ xuất hiện duy nhất một lần trong toàn bộ trang web).\n'
        '  + **Bước 2:** Tạo liên kết tới phần tử trên bằng cách thiết lập thuộc tính href="#mã_định_danh_đoạn".\n'
        '  + Ví dụ (SGK trang 60 – tạo bảng có id là Thong_tin và đặt liên kết tới bảng đó):\n'
        '    1. <table id="Thong_tin"> ...</table>\n'
        '    2. <a href="#Thong_tin"> Thông tin chi tiết </a>\n'
        '- **Tạo liên kết tới hình ảnh / dùng hình ảnh làm liên kết (SGK trang 60):**\n'
        '  + Sử dụng thẻ <img src="đường dẫn tới tệp ảnh"> đặt bên trong thẻ <a>.\n'
        '  + Ví dụ (Hình 10.1): Trong trang index.html, hiển thị ảnh Mặt Trời (images/sun.png) và đặt liên kết từ ảnh tới trang thong_tin.html:\n'
        '    <a href="thong_tin.html"><img src="images/sun.png"></a>\n'
        '- **Trả lời Câu hỏi củng cố SGK trang 60:**\n'
        '  Đoạn mã HTML trong tệp index.html tạo danh sách không có thứ tự liên kết tới 3 tệp bài tập:\n'
        '  <ul>\n'
        '    <li><a href="bai_tap/bai_tap_1.html">bài tập 1</a></li>\n'
        '    <li><a href="bai_tap/bai_tap_2.html">bài tập 2</a></li>\n'
        '    <li><a href="bai_tap/on_tap/bai_tap_on_tap.html">ôn tập</a></li>\n'
        '  </ul>'
    ]
]

# Bước thực hành trọn vẹn Nhiệm vụ 1 và Nhiệm vụ 2 (dùng cho Tiết 2 và Bản gộp Cả bài)
NV4_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh mở thư mục thực hành chứa tệp CLB.html (sản phẩm đã tạo ở Bài 9 gồm danh sách các CLB và bảng Lịch hoạt động các CLB Thể thao), quan sát mục 3. Thực hành tạo liên kết và Hình 10.2 (SGK trang 61) để thực hiện tuần tự 02 nhiệm vụ trên máy tính:\n'
        '  + Nhiệm vụ 1 – Tạo liên kết trong cùng trang (Hình 10.2, trang 61): Thêm phần "Giới thiệu các hoạt động ngoại khoá" lên đầu tệp CLB.html; gắn mã định danh id="CLBTT" cho bảng Lịch hoạt động các CLB Thể thao và đặt liên kết từ dòng "Theo dõi lịch hoạt động các CLB Thể thao" trượt thẳng xuống bảng đó.\n'
        '  + Nhiệm vụ 2 – Tạo liên kết sang trang khác (trang 61): Tạo thêm tệp thong_tin.html trong cùng thư mục với CLB.html để giới thiệu chi tiết về các câu lạc bộ, sau đó đặt đường liên kết tương đối từ trang chủ CLB.html sang trang thong_tin.html.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thực hành viết mã HTML trên máy tính, lưu tệp (Ctrl + S) và mở trên trình duyệt web để kiểm thử cả hai liên kết. Giáo viên quan sát các nhóm học sinh, hỗ trợ kiểm tra lỗi quên dấu "#" trước mã định danh CLBTT hoặc lưu tệp thong_tin.html khác thư mục.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên chiếu 02 sản phẩm thực hành hoàn chỉnh của học sinh lên màn chiếu; học sinh nháy chuột trực tiếp vào liên kết nội bộ (#CLBTT) và liên kết chuyển trang (thong_tin.html) để cả lớp kiểm chứng.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá sản phẩm thực hành theo Tiêu chí 1 và Tiêu chí 2 của Rubric CCĐG 6 (Mục IV. Phụ lục 2).',
        '**3. Thực hành tạo liên kết (SGK trang 61)**\n'
        '**a) Nhiệm vụ 1: Tạo liên kết trong cùng trang (Hình 10.2, trang 61):**\n'
        '- **Bước 1:** Mở tệp CLB.html và thêm nội dung giới thiệu hoạt động ngoại khoá của trường lên đầu phần thân trang (<body>) như Hình 10.2:\n'
        '  <h2>Giới thiệu các hoạt động ngoại khoá</h2>\n'
        '  <p>Trường THPT Nguyễn Bỉnh Khiêm có bề dày thành tích về học tập và các hoạt động thể thao, văn nghệ của thành phố. Các câu lạc bộ ngoại khoá hoạt động sôi nổi và luôn được nhà trường tạo điều kiện để sinh hoạt. Những năm qua, trường đạt nhiều thành tích trong các cuộc thi cấp thành phố, quận:</p>\n'
        '  <ul>\n'
        '    <li>03 HCV cấp thành phố môn Bơi lội</li>\n'
        '    <li>02 HCV, 06 HCB và 05 HCĐ cấp thành phố các môn Võ thuật</li>\n'
        '    <li>01 giải ba môn Bóng đá nam cấp thành phố</li>\n'
        '    <li>01 giải nhì, 02 giải ba Hội thi biểu diễn âm nhạc quận</li>\n'
        '  </ul>\n'
        '- **Bước 2:** Thêm mã định danh id="CLBTT" cho thẻ mở <table> của bảng Lịch hoạt động các CLB Thể thao:\n'
        '  <table id="CLBTT" border="1">\n'
        '    <caption>Lịch hoạt động các CLB Thể thao</caption>\n'
        '- **Bước 3:** Thêm liên kết nội bộ cho dòng "Theo dõi lịch hoạt động các CLB Thể thao":\n'
        '  <p><a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a></p>\n'
        '**b) Nhiệm vụ 2: Tạo liên kết sang trang khác (SGK trang 61):**\n'
        '- **Bước 1:** Trong cùng thư mục với tệp CLB.html, tạo tệp mới thong_tin.html và thêm nội dung giới thiệu thông tin về các câu lạc bộ.\n'
        '- **Bước 2:** Trong tệp CLB.html, thêm đoạn mã tạo liên kết bằng đường dẫn tương đối đến tệp thong_tin.html:\n'
        '  <p><a href="thong_tin.html">Thông tin của các CLB</a></p>'
    ]
]

# Hoạt động 3 (Luyện tập) và Hoạt động 4 (Vận dụng) cho Tiết 2 Thực hành và Bản gộp Cả bài
HD3_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh hoàn thành Phiếu học tập số 4 kết hợp viết mã kiểm chứng trực tiếp trên máy tính cho 02 bài tập phần Luyện tập (SGK trang 61):\n'
        '  + Bài Luyện tập 1 (SGK trang 61): Viết đoạn mã HTML để tạo liên kết quay về trang chủ CLB.html được đặt bên trong trang thong_tin.html.\n'
        '  + Bài Luyện tập 2 (SGK trang 61): Viết một trang web giới thiệu bản thân (lưu tên gioi_thieu_ban_than.html), trong đó đặt ít nhất ba liên kết tới trang web (có thể là trang Facebook hoặc trang cá nhân) của ba bạn cùng lớp.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh bổ sung thẻ <a> quay về CLB.html ở cuối tệp thong_tin.html, sau đó tạo nhanh tệp gioi_thieu_ban_than.html có danh sách 3 liên kết tuyệt đối và ghi câu lệnh vào Phiếu học tập số 4.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình bày mã nguồn và chạy thử thao tác chuyển trang khép kín từ CLB.html sang thong_tin.html rồi quay lại CLB.html.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét, chấm điểm theo CCĐG 5 và nhấn mạnh sự khác biệt: Bài Luyện tập 1 dùng đường dẫn tương đối (vì hai tệp nằm cùng thư mục), còn Bài Luyện tập 2 dùng đường dẫn tuyệt đối (vì liên kết tới trang mạng xã hội trên Internet).',
        '**Đáp án phần Luyện tập (SGK Bài 10, trang 61):**\n'
        '- **Giải Bài Luyện tập 1 (SGK trang 61):**\n'
        '  + Do tệp thong_tin.html và tệp CLB.html nằm trong cùng một thư mục nên ta dùng đường dẫn tương đối. Đoạn mã đặt trong tệp thong_tin.html là:\n'
        '    <p><a href="CLB.html">Quay về trang chủ CLB</a></p>\n'
        '- **Giải Bài Luyện tập 2 (SGK trang 61):**\n'
        '  + Tạo tệp gioi_thieu_ban_than.html sử dụng thẻ <a> với đường dẫn tuyệt đối (bắt đầu bằng https://) dẫn tới trang Facebook của 3 bạn cùng lớp:\n'
        '    <h2>Giới thiệu bản thân</h2>\n'
        '    <p>Họ và tên: Nguyễn Văn A – Lớp 12 – Trường THPT Thuận Hoá.</p>\n'
        '    <p>Danh sách trang cá nhân của ba bạn cùng lớp:</p>\n'
        '    <ul>\n'
        '      <li><a href="https://www.facebook.com/ban.thu.nhat">Trang Facebook của bạn Trần Minh B</a></li>\n'
        '      <li><a href="https://www.facebook.com/ban.thu.hai">Trang Facebook của bạn Lê Thị C</a></li>\n'
        '      <li><a href="https://www.facebook.com/ban.thu.ba">Trang Facebook của bạn Phạm Quốc D</a></li>\n'
        '    </ul>'
    ]
]

HD4_ROWS_BAI10 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên giao nhiệm vụ phần Vận dụng (SGK trang 61): Viết thêm các đoạn giới thiệu (địa điểm, số thành viên, thành tích, phí hội viên, điện thoại liên hệ) cho từng câu lạc bộ ở dưới bảng Lịch hoạt động các CLB Thể thao trong tệp CLB.html, rồi đặt liên kết từ tên câu lạc bộ trong bảng đến đoạn giới thiệu tương ứng.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thực hành ngay trên tệp CLB.html: tạo các khối hoặc đoạn văn có gắn thuộc tính id riêng cho từng bộ môn (ví dụ: id="clb-bongda", id="clb-bongchuyen", id="clb-bongro", id="clb-vothuat", id="clb-boi") ở bên dưới bảng, rồi bọc tên bộ môn trong các ô <td> của bảng bằng thẻ <a href="#..."> tương ứng.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình chiếu trang CLB.html hoàn chỉnh, nháy chuột vào tên từng CLB trong bảng để kiểm tra trang tự động di chuyển đến đúng đoạn giới thiệu chi tiết phía dưới.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên chấm điểm tổng kết theo Rubric CCĐG 6, khen ngợi các học sinh hoàn thành xuất sắc và dặn dò chuẩn bị Bài 11: Chèn tệp tin đa phương tiện và khung nội tuyến vào trang web (Bài 4.11 – Tiết PPCT: 23, 24).',
        '**Hướng dẫn giải phần Vận dụng (SGK Bài 10, trang 61):**\n'
        '- **Bước 1 – Đặt liên kết tại các ô tên bộ môn trong bảng Lịch hoạt động các CLB Thể thao (tệp CLB.html):**\n'
        '  + Ô Bóng đá: <td><a href="#clb-bongda">Bóng đá</a></td>\n'
        '  + Ô Bóng chuyền: <td><a href="#clb-bongchuyen">Bóng chuyền</a></td>\n'
        '  + Ô Bóng rổ: <td><a href="#clb-bongro">Bóng rổ</a></td>\n'
        '  + Ô Võ thuật: <td><a href="#clb-vothuat">Võ thuật</a></td>\n'
        '  + Ô Bơi: <td><a href="#clb-boi">Bơi</a></td>\n'
        '- **Bước 2 – Viết các đoạn giới thiệu có gắn thuộc tính id tương ứng ở phía dưới bảng (</tbody></table>):**\n'
        '  <h3 id="clb-bongda">1. Câu lạc bộ Bóng đá</h3>\n'
        '  <p>Địa điểm: Sân bóng cỏ nhân tạo trường. Số thành viên: 45 học sinh. Thành tích: 01 giải Ba cấp thành phố. Phí hội viên: 50.000đ/tháng. Điện thoại liên hệ: Thầy Khải (0905.111.222).</p>\n'
        '  <h3 id="clb-bongchuyen">2. Câu lạc bộ Bóng chuyền</h3>\n'
        '  <p>Địa điểm: Sân thể chất số 1. Số thành viên: 32 học sinh. Thành tích: Giải Khuyến khích cấp cụm. Phí hội viên: 40.000đ/tháng. Điện thoại liên hệ: Cô Lan (0905.222.333).</p>\n'
        '  <h3 id="clb-bongro">3. Câu lạc bộ Bóng rổ</h3>\n'
        '  <p>Địa điểm: Nhà thi đấu đa năng. Số thành viên: 38 học sinh. Thành tích: Top 4 giải học sinh thành phố. Phí hội viên: 50.000đ/tháng. Điện thoại liên hệ: Thầy Vũ (0905.333.444).</p>\n'
        '  <h3 id="clb-vothuat">4. Câu lạc bộ Võ thuật</h3>\n'
        '  <p>Địa điểm: Phòng tập Võ đường trường. Số thành viên: 60 học sinh. Thành tích: 02 HCV, 06 HCB, 05 HCĐ cấp thành phố. Phí hội viên: 60.000đ/tháng. Điện thoại liên hệ: Thầy Hoàng (0905.444.555).</p>\n'
        '  <h3 id="clb-boi">5. Câu lạc bộ Bơi lội</h3>\n'
        '  <p>Địa điểm: Bể bơi tiêu chuẩn trường. Số thành viên: 40 học sinh. Thành tích: 03 HCV cấp thành phố. Phí hội viên: 70.000đ/tháng. Điện thoại liên hệ: Cô Bích (0905.555.666).</p>'
    ]
]

# Tiết 1 (Lí thuyết – Lớp 12/2, Phòng H.II.2): học sinh làm việc trên Phiếu học tập và vở ghi, giáo viên minh hoạ trên máy chiếu
HD3_ROWS_T1 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên phát Phiếu học tập số 4 và yêu cầu mỗi học sinh tự suy luận trên giấy (không dùng máy tính) để giải quyết 02 bài tập phần Luyện tập (SGK trang 61) cùng 02 câu hỏi trắc nghiệm củng cố sơ đồ Hình 10.1:\n'
        '  + Câu 1 (Bài Luyện tập 1 SGK trang 61): Viết đoạn mã HTML đặt trong tệp thong_tin.html để tạo liên kết quay về trang chủ CLB.html (biết hai tệp nằm cùng thư mục).\n'
        '  + Câu 2 (Bài Luyện tập 2 SGK trang 61): Viết cấu trúc đoạn mã HTML tạo danh sách ít nhất 03 liên kết tới trang Facebook của ba bạn cùng lớp.\n'
        '  + Câu 3 (Củng cố Hình 10.1): Từ tệp bai_tap_2.html (trong thư mục bai_tap/), muốn liên kết đến tệp bai_tap_on_tap.html (trong thư mục con bai_tap/on_tap/), cần viết thuộc tính href như thế nào?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh vận dụng kiến thức về đường dẫn tương đối và đường dẫn tuyệt đối vừa học ở mục 1, mục 2 để viết câu lệnh HTML vào Phiếu học tập số 4.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh lên bảng viết câu lệnh của Bài Luyện tập 1, Bài Luyện tập 2 và giải thích Câu 3; sau đó giáo viên mở sẵn tệp mẫu trên máy chiếu để cả lớp quan sát kết quả chạy thử.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét, chấm điểm theo CCĐG 5 và lưu ý học sinh tránh lỗi quên dấu ngoặc kép bao quanh giá trị của thuộc tính href.',
        '**Đáp án Phiếu học tập số 4 – Phần Luyện tập (SGK Bài 10, trang 59–61):**\n'
        '- **Câu 1 – Bài Luyện tập 1 (SGK trang 61):**\n'
        '  + Vì thong_tin.html và CLB.html đặt trong cùng thư mục nên dùng đường dẫn tương đối:\n'
        '    <p><a href="CLB.html">Quay về trang chủ CLB</a></p>\n'
        '- **Câu 2 – Bài Luyện tập 2 (SGK trang 61):**\n'
        '  + Vì liên kết tới trang Facebook trên mạng Internet nên dùng đường dẫn tuyệt đối (có giao thức https://):\n'
        '    <ul>\n'
        '      <li><a href="https://www.facebook.com/ban.mot">Facebook bạn thứ nhất</a></li>\n'
        '      <li><a href="https://www.facebook.com/ban.hai">Facebook bạn thứ hai</a></li>\n'
        '      <li><a href="https://www.facebook.com/ban.ba">Facebook bạn thứ ba</a></li>\n'
        '    </ul>\n'
        '- **Câu 3 – Liên kết từ bai_tap_2.html đến bai_tap_on_tap.html (Hình 10.1, trang 59):**\n'
        '  + Tệp bai_tap_2.html đang nằm sẵn trong thư mục bai_tap/, còn tệp đích nằm trong thư mục con on_tap/ dưới một cấp so với bai_tap/, do đó câu lệnh đúng là:\n'
        '    <a href="on_tap/bai_tap_on_tap.html">Bài tập ôn tập</a>'
    ]
]

HD4_ROWS_T1 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu lên màn hình Hình 10.2 cùng nội dung mục 3. Thực hành tạo liên kết và mục Vận dụng (SGK trang 61), yêu cầu học sinh thiết kế trước sơ đồ mã định danh (id) và câu lệnh liên kết vào vở ghi để chuẩn bị cho Tiết 2 Thực hành (Tiết PPCT: 22):\n'
        '  + Xác định vị trí gắn id="CLBTT" và vị trí đặt thẻ <a href="#CLBTT"> trong tệp CLB.html (Nhiệm vụ 1).\n'
        '  + Xác định câu lệnh liên kết từ CLB.html sang thong_tin.html (Nhiệm vụ 2).\n'
        '  + Thiết kế bảng mã định danh đoạn (id) cho 5 câu lạc bộ thể thao ở phần Vận dụng (trang 61).\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh quan sát Hình 10.2 trên màn chiếu, ghi tóm tắt cấu trúc thẻ và các mã định danh id vào vở. Ở tiết lí thuyết này học sinh chưa thực hành trên máy tính.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình bày cách đặt tên id và cú pháp liên kết nội bộ cho 5 câu lạc bộ trong bảng Lịch hoạt động CLB Thể thao.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét phần chuẩn bị của học sinh, tổng kết Tiết 1 (Tiết PPCT: 21) và dặn học sinh mang theo vở ghi để thực hành ở Tiết 2 (Tiết PPCT: 22).',
        '**Bản thiết kế chuẩn bị cho Tiết 2 Thực hành (SGK Bài 10, trang 61):**\n'
        '1. **Nhiệm vụ 1 (Liên kết trong cùng trang CLB.html – Hình 10.2):**\n'
        '   - Tại thẻ mở của bảng: <table id="CLBTT" border="1">\n'
        '   - Tại dòng cuối phần giới thiệu đầu trang: <a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a>\n'
        '2. **Nhiệm vụ 2 (Liên kết sang trang khác cùng thư mục):**\n'
        '   - Tạo tệp thong_tin.html cùng thư mục với CLB.html.\n'
        '   - Đặt liên kết trong CLB.html: <a href="thong_tin.html">Thông tin của các CLB</a>\n'
        '3. **Vận dụng (Liên kết từ tên CLB trong bảng xuống đoạn giới thiệu cuối trang CLB.html):**\n'
        '   - Gắn thuộc tính id cho 5 tiêu đề đoạn giới thiệu dưới bảng: id="clb-bongda", id="clb-bongchuyen", id="clb-bongro", id="clb-vothuat", id="clb-boi".\n'
        '   - Đặt thẻ <a href="#clb-..."> tương ứng vào 5 ô tên bộ môn trong bảng.'
    ]
]


def add_bai10_appendix(b, class_label='12/2 và 12/1', ppct_label='21, 22', mode='cabai'):
    add(b, 'IV. PHỤ LỤC (HỒ SƠ HỌC LIỆU, PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ)', 'heading')
    all_rows = [
        ['Tên Phiếu / Công cụ', 'Nội dung câu hỏi và yêu cầu (Dựa sát SGK Tin học 12 KNTT, Bài 10 trang 57–61)', 'Đáp án chuẩn và Thang điểm chấm'],
        ['**CCĐG 1**\n(Hoạt động 1:\nMở đầu)',
         '- Câu hỏi Khởi động SGK trang 57: Nêu những điểm khác biệt khi đọc thông tin trên trang web so với đọc sách, báo giấy và chỉ ra điểm khác biệt quan trọng nhất.',
         '- Nêu được trang web có đa phương tiện và siêu liên kết (hyperlink) giúp di chuyển phi tuyến tính tức thời giữa các nội dung (Đạt / Chưa đạt).'],
        ['**Phiếu học tập số 1**\n**và CCĐG 2**\n(Hoạt động 2 –\nNhiệm vụ 1)',
         '1. Giải Hoạt động 1 (SGK tr. 57): Vì sao tệp .docx có thể là siêu văn bản còn tệp .txt thì không? Nêu đặc điểm của siêu văn bản.\n'
         '2. Viết cú pháp thẻ <a> và phân biệt Đường dẫn tuyệt đối với Đường dẫn tương đối.\n'
         '3. Giải Câu hỏi tr. 58: Phân loại 3 đường dẫn a) html/cach_tao_lien_ket.html; b) http://google.com; c) mail.google.com/mail/u/0/#inbox/...',
         '- Nêu đúng đặc điểm siêu văn bản và cú pháp <a href="URL">Nội dung</a> *(4,0 điểm)*.\n'
         '- Phân biệt đúng đường dẫn tuyệt đối (có giao thức + tên miền) và tương đối *(3,0 điểm)*.\n'
         '- Câu hỏi tr. 58: b là đường dẫn tuyệt đối; a và c là đường dẫn tương đối *(3,0 điểm)*.'],
        ['**Phiếu học tập số 2**\n**và CCĐG 3**\n(Hoạt động 2 –\nNhiệm vụ 2)',
         '1. Viết câu lệnh HTML tạo liên kết ngoài tới trang Hành trang số (SGK tr. 58).\n'
         '2. Dựa vào Hình 10.1 (SGK tr. 59), viết 4 câu lệnh thẻ <a> cho 4 trường hợp: (a) Từ index.html tới thong_tin.html; (b) Từ index.html tới bai_tap_1.html; (c) Từ index.html tới bai_tap_on_tap.html; (d) Từ bai_tap_1.html quay lại index.html.',
         '- Viết đúng liên kết ngoài với đường dẫn tuyệt đối https://... *(2,0 điểm)*.\n'
         '- Viết đúng 4 câu lệnh theo Hình 10.1: href="thong_tin.html", href="bai_tap/bai_tap_1.html", href="bai_tap/on_tap/bai_tap_on_tap.html", href="../index.html" *(8,0 điểm – mỗi ý 2,0 điểm)*.'],
        ['**Phiếu học tập số 3**\n**và CCĐG 4**\n(Hoạt động 2 –\nNhiệm vụ 3)',
         '1. Nêu 2 bước tạo liên kết tới vị trí khác trong cùng trang web bằng thuộc tính id và kí tự "#".\n'
         '2. Viết câu lệnh đặt ảnh Mặt Trời (images/sun.png) làm liên kết tới trang thong_tin.html (SGK tr. 60).\n'
         '3. Giải Câu hỏi củng cố SGK tr. 60: Viết danh sách <ul> gồm 3 mục liên kết tới bai_tap_1.html, bai_tap_2.html, bai_tap_on_tap.html.',
         '- Trình bày đúng 2 bước: gắn id="Tên_id" và gọi href="#Tên_id" *(3,0 điểm)*.\n'
         '- Viết đúng <a href="thong_tin.html"><img src="images/sun.png"></a> *(3,0 điểm)*.\n'
         '- Viết đúng danh sách <ul> chứa 3 thẻ <li><a href="...">...</a></li> *(4,0 điểm)*.'],
        ['**Phiếu học tập số 4**\n**và CCĐG 5**\n(Hoạt động 3:\nLuyện tập)',
         '1. Giải Bài Luyện tập 1 (SGK tr. 61): Viết đoạn mã tạo liên kết quay về trang chủ CLB.html trong trang thong_tin.html.\n'
         '2. Giải Bài Luyện tập 2 (SGK tr. 61): Viết đoạn mã trang web giới thiệu bản thân có ít nhất 3 liên kết tới trang web (Facebook) của 3 bạn cùng lớp.',
         '- Bài 1: Viết đúng <a href="CLB.html">Quay về trang chủ CLB</a> *(5,0 điểm)*.\n'
         '- Bài 2: Viết đúng cấu trúc HTML có ít nhất 3 thẻ <a href="https://..."> hợp lệ *(5,0 điểm)*.']
    ]
    ccdg1_t2 = [
        '**CCĐG 1**\n(Hoạt động 1:\nMở đầu Tiết 2)',
        '- Mở tệp CLB.html (Bài 9) trên trình soạn thảo mã và trả lời nhanh 3 câu hỏi thao tác cốt lõi: (1) Cú pháp gắn mã định danh đoạn và liên kết tới vị trí trong cùng trang? (2) Cú pháp liên kết tới tệp khác cùng thư mục? (3) Cách lưu và kiểm thử liên kết trên trình duyệt?',
        '- (1) Gắn id="CLBTT" tại thẻ đích và gọi href="#CLBTT" tại thẻ <a>; (2) Dùng href="thong_tin.html"; (3) Nhấn Ctrl + S và tải lại trang (F5) trên trình duyệt (Đạt / Chưa đạt).'
    ]
    if mode == 'tiet1':
        pht_rows = all_rows
        add(b, 'Phụ lục 1. Hệ thống Phiếu học tập số 1, 2, 3, 4 và Công cụ đánh giá CCĐG 1 – CCĐG 5', 'sub')
    elif mode == 'tiet2':
        pht_rows = [all_rows[0], ccdg1_t2, all_rows[5]]
        add(b, 'Phụ lục 1. Phiếu học tập số 4 và Công cụ đánh giá CCĐG 1, CCĐG 5', 'sub')
    else:
        pht_rows = all_rows
        add(b, 'Phụ lục 1. Hệ thống Phiếu học tập số 1, 2, 3, 4 và Công cụ đánh giá CCĐG 1 – CCĐG 5', 'sub')
    b.append(table(pht_rows, [1600, 4200, 3200], header=True, cell_sz=21))

    add(b, 'Phụ lục 2. Công cụ đánh giá CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành HTML tạo liên kết (Nhiệm vụ 1, Nhiệm vụ 2, Luyện tập và Vận dụng SGK trang 61)', 'sub')
    rubric_info = [
        ['Thông tin hành chính phiếu đánh giá sản phẩm thực hành (CCĐG 6)', 'Nội dung chi tiết'],
        ['Tên bài dạy và Tiết PPCT', f'Bài 10: Tạo liên kết (Tiết PPCT: {ppct_label})'],
        ['Lớp giảng dạy và Trường thực tập', f'Lớp {class_label} – Trường THPT Thuận Hoá'],
        ['Giáo viên giảng dạy', GVGD_NAME],
        ['Giáo viên hướng dẫn', GVHD_NAME],
        ['Sinh viên thực hiện', f'{SV_NAME} (Mã sinh viên: 24S1020030)'],
        ['Họ và tên học sinh / Nhóm', '........................................................................ Lớp: ......................']
    ]
    b.append(table(rubric_info, [3800, 5200], header=True, cell_sz=22))
    add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Phần tử và Thuộc tính HTML bắt buộc kiểm tra', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Điểm tối đa', 'Điểm đạt'],
        ['1', 'Nhiệm vụ 1:\nLiên kết trong cùng trang\n(Hình 10.2 SGK tr. 61)', 'Nội dung giới thiệu đầu trang CLB.html; thuộc tính id="CLBTT" tại thẻ <table>; thẻ <a href="#CLBTT">.', '- **Mức 3 (3,5 đ):** Thêm đầy đủ phần giới thiệu và danh sách thành tích như Hình 10.2; gắn đúng id="CLBTT" cho bảng và nháy vào liên kết cuộn mượt xuống bảng.\n- **Mức 2 (2,5 đ):** Tạo được liên kết #CLBTT hoạt động đúng nhưng phần văn bản giới thiệu còn thiếu ý.\n- **Mức 1 (1,0 đ):** Quên dấu "#" trong href="#CLBTT" hoặc chưa gắn id cho thẻ <table>.', '3,5 điểm', '......'],
        ['2', 'Nhiệm vụ 2 và Luyện tập 1:\nLiên kết hai chiều giữa CLB.html và thong_tin.html', 'Tệp thong_tin.html cùng thư mục; <a href="thong_tin.html"> trong CLB.html và <a href="CLB.html"> trong thong_tin.html.', '- **Mức 3 (3,5 đ):** Tạo đúng tệp thong_tin.html cùng thư mục; chuyển trang hai chiều thông suốt giữa CLB.html và thong_tin.html.\n- **Mức 2 (2,5 đ):** Liên kết từ CLB.html sang thong_tin.html hoạt động đúng nhưng chưa có liên kết quay lại.\n- **Mức 1 (1,0 đ):** Lưu sai thư mục hoặc viết sai tên tệp khiến trình duyệt báo lỗi không tìm thấy tệp.', '3,5 điểm', '......'],
        ['3', 'Luyện tập 2 và Vận dụng:\nLiên kết ngoài và Liên kết từng CLB trong bảng (tr. 61)', 'Đường dẫn tuyệt đối https://...; các thuộc tính id cho từng CLB dưới bảng và thẻ <a href="#..."> trong các ô <td>.', '- **Mức 3 (2,0 đ):** Hoàn thành liên kết từ tên các CLB trong bảng xuống đoạn giới thiệu chi tiết dưới bảng và tạo đúng liên kết tuyệt đối.\n- **Mức 2 (1,25 đ):** Tạo được đoạn giới thiệu và liên kết cho ít nhất 02 câu lạc bộ trong bảng.\n- **Mức 1 (0,5 đ):** Mới viết được đoạn văn giới thiệu dưới bảng nhưng chưa gắn được liên kết từ bảng xuống.', '2,0 điểm', '......'],
        ['4', 'Cú pháp mã nguồn HTML và Ý thức thực hành', 'Đóng/mở thẻ chuẩn, khai báo <meta charset="UTF-8">, lưu tệp đúng quy định.', '- **Mức 3 (1,0 đ):** Mã HTML thụt lề rõ ràng, hiển thị tiếng Việt có dấu chuẩn, hoàn thành đúng thời gian.\n- **Mức 2 (0,5 đ):** Trang web chạy được nhưng trình bày mã nguồn chưa gọn.\n- **Mức 1 (0,25 đ):** Quên lưu tệp hoặc đặt tên tệp sai phần mở rộng .html.', '1,0 điểm', '......'],
        ['', '**Tổng điểm**', '', '**Xếp loại:** Đạt từ 8,0 – 10,0: Tốt | 5,0 – 7,9: Đạt | < 5,0: Cần cố gắng', '**10,0 điểm**', '......']
    ]
    b.append(table(rubric_rows, [650, 1500, 2000, 3450, 700, 700], header=True, cell_sz=21))


# ==============================================================================
# 1. FILE 01: TIẾT 1 – LÍ THUYẾT (TIẾT PPCT: 21 – LỚP 12/2 – PHÒNG H.II.2)
# ==============================================================================
def build_bai10_tiet1():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT – TIẾT PPCT: 21)',
        class_str='Tin học 12 – Lớp 12/2 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.II.2',
        lesson_row_str='Bài 4.10: Tạo liên kết (tiết 1 - LT) – Tiết PPCT: 21 (Tổng bài: 02 tiết – PPCT: 21, 22)',
        page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT (TIẾT 1 – LÍ THUYẾT)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 21, 22] – Kế hoạch này chi tiết hoá Tiết thứ 1 của bài / Tiết PPCT: 21)',
        page1_class_line='Môn học: Tin học; Lớp: 12/2',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 1 – LT / Tiết PPCT: 21)',
        page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 12/2 – Phòng H.II.2',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp quan sát trực quan và thảo luận nhóm',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 12, Chủ đề F (Giải quyết vấn đề với sự trợ giúp của máy tính – Tạo trang web) và Kế hoạch dạy học môn Tin học Khối 12 của Trường THPT Thuận Hoá (Tiết PPCT 21 – Bài 4.10: Tạo liên kết (tiết 1 - LT)):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu [KT1] – [KT3] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]:** Nêu được khái niệm siêu văn bản (hypertext), siêu liên kết (hyperlink) và cú pháp thẻ <a> tạo liên kết trong HTML (<a href="URL">Nội dung hiển thị</a>); phân biệt chính xác đường dẫn tuyệt đối với đường dẫn tương đối (SGK trang 57–58).',
        '**[KT2]:** Trình bày được cách tạo liên kết tới một trang web khác trên Internet (liên kết ngoài dùng đường dẫn tuyệt đối) và cách tạo liên kết tới các trang web trong cùng một website theo 4 trường hợp cấu trúc thư mục ở Hình 10.1 (cùng thư mục, dưới một cấp, dưới hai hay nhiều cấp, và lên thư mục mức trên bằng kí tự "../") (SGK trang 58–60).',
        '**[KT3]:** Giải thích được quy trình 2 bước tạo liên kết tới một vị trí khác trong cùng trang web bằng thuộc tính định danh id và tham chiếu href="#mã_định_danh_đoạn", cách đặt hình ảnh làm liên kết bằng thẻ <img> bên trong thẻ <a>, và kết hợp được thẻ <a> với danh sách <ul><li> (SGK trang 60).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Học sinh chủ động đọc hiểu Sách giáo khoa Tin học 12 trang 57–60, phân tích sơ đồ cây thư mục Hình 10.1 để hoàn thành các Phiếu học tập.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Học sinh thảo luận cặp đôi sôi nổi để phân loại các dạng đường dẫn và kiểm tra chéo cú pháp câu lệnh HTML.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Học sinh xác định đúng đường dẫn tương đối giữa hai tệp bất kì trong cấu trúc cây thư mục nhiều cấp.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực ứng xử phù hợp trong môi trường số [NLb]:** Có ý thức kiểm chứng tính chính xác và an toàn của địa chỉ URL trước khi chia sẻ hoặc nhúng liên kết vào trang web.',
        '**Năng lực giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông [NLc]:** Viết đúng các đoạn mã HTML sử dụng thẻ <a>, thuộc tính href, thuộc tính id và thẻ <img> để kết nối các tài nguyên web.',
        '**Năng lực hợp tác trong môi trường số [NLe]:** Phối hợp hiệu quả với bạn cùng nhóm trong hoạt động phân tích cấu trúc website.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Nắm vững quy tắc tổ chức siêu văn bản và tạo lập liên kết điều hướng giữa các trang nội dung số.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Học sinh tập trung theo dõi bài giảng, ghi chép đầy đủ cú pháp thẻ <a> và các trường hợp đường dẫn.',
        '**Phẩm chất trách nhiệm [PC2]:** Học sinh hoàn thành nghiêm túc các Phiếu học tập cá nhân và nhóm đôi đúng thời gian quy định.',
        '**Phẩm chất trung thực [PC3]:** Học sinh tự giác làm bài tập củng cố và nhận xét khách quan kết quả của bạn.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính của giáo viên có trình duyệt web và trình soạn thảo mã HTML, máy chiếu, bảng viết và bộ Phiếu học tập số 1, 2, 3, 4.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam*, Bài 10: Tạo liên kết (mục 1, mục 2 và phần Luyện tập, Vận dụng – trang 57 đến trang 61).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 12 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 21 – Bài 4.10: Tạo liên kết (tiết 1 - LT)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 1 – Lí thuyết (45 phút – Tiết thứ 1 của bài / Tiết PPCT: 21)**', first=0, keep=True)
    ov_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng và Tiết PPCT', 'Mục tiêu', 'Sản phẩm và Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(So sánh đọc trang web với đọc sách, báo giấy – SGK tr. 57)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 21)', '[KT1], [NLC1],\n[PC1]', 'Câu trả lời miệng nêu bật vai trò của siêu liên kết.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Siêu văn bản, thẻ <a> và Đường dẫn tuyệt đối/tương đối (10p)\n- 2.2. Nhiệm vụ 2: Liên kết ngoài và Liên kết trong cùng website theo Hình 10.1 (10p)\n- 2.3. Nhiệm vụ 3: Liên kết vị trí trong cùng trang (id) và Liên kết hình ảnh (08p)', '28 phút\n(Tiết 1 của bài /\nTiết PPCT: 21)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2],\n[NLb], [NLc], [NLe],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, số 2, số 3 đã hoàn thành.\nCông cụ đánh giá: CCĐG 2, CCĐG 3, CCĐG 4.'],
        ['3. Hoạt động 3: Luyện tập\n(Viết mã HTML giải Bài Luyện tập 1, 2 SGK tr. 61 và Củng cố Hình 10.1)', '07 phút\n(Tiết 1 của bài /\nTiết PPCT: 21)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC3],\n[NLc], [PC3]', 'Lời giải trên Phiếu học tập số 4.\nCông cụ đánh giá: CCĐG 5.'],
        ['4. Hoạt động 4: Vận dụng\n(Thiết kế sơ đồ mã định danh id cho dự án CLB.html chuẩn bị Tiết 2)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 21)', '[KT3], [NLC3],\n[NLc], [PC2]', 'Bản thiết kế mã định danh id và liên kết trong vở ghi chuẩn bị cho Tiết 2 Thực hành (Tiết PPCT: 22).']
    ]
    b.append(table(ov_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [PC1] – Học sinh nhận biết được sự khác biệt cốt lõi giữa việc đọc thông tin trên trang web với đọc sách, báo giấy, từ đó hiểu được ý nghĩa của siêu liên kết.', first=0)
    add(b, '**b) Nội dung:** Học sinh thảo luận cặp đôi trả lời câu hỏi Khởi động trong Sách giáo khoa Tin học 12 trang 57.', first=0)
    add(b, '**c) Sản phẩm:** Câu trả lời miệng của học sinh chỉ ra điểm khác biệt quan trọng nhất trên trang web là có các siêu liên kết (hyperlink) giúp di chuyển nhanh chóng giữa các trang thông tin.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 1 – Bảng kiểm quan sát Hoạt động Mở đầu *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD1_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (28 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu siêu văn bản, siêu liên kết, thẻ <a> và hai loại đường dẫn (10 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [NLb], [PC1] – Học sinh trình bày được đặc điểm của siêu văn bản, cú pháp thẻ <a> và phân biệt được đường dẫn tuyệt đối với đường dẫn tương đối.', first=0)
    add(b, '**b) Nội dung:** Học sinh nghiên cứu mục 1. Siêu văn bản và đường dẫn (SGK trang 57–58), trả lời Hoạt động 1 (trang 57) và Câu hỏi củng cố (trang 58) trên Phiếu học tập số 1.', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 1 ghi đầy đủ khái niệm siêu văn bản, cú pháp thẻ <a> và phân loại đúng 3 đường dẫn a, b, c ở trang 58.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 2 – Thang điểm chấm Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV1_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu liên kết tới trang web khác trên Internet và liên kết tới trang web trong cùng website theo Hình 10.1 (10 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLC2], [NLC3], [NLc], [NLe], [PC1], [PC2] – Học sinh viết đúng câu lệnh HTML cho liên kết ngoài (đường dẫn tuyệt đối) và 4 trường hợp liên kết tương đối trong cùng website theo Hình 10.1.', first=0)
    add(b, '**b) Nội dung:** Học sinh đọc mục 2a, 2b (SGK trang 58–60), quan sát Hình 10.1 – Cấu trúc một website (SGK trang 59) và hoàn thành Phiếu học tập số 2:', first=0)
    add_figure(b, 'hinh_10_1.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 2 hoàn chỉnh các đoạn mã HTML tạo liên kết ngoài và 4 trường hợp liên kết trong cùng website (cùng thư mục, dưới 1 cấp, dưới 2 cấp, lên thư mục mức trên bằng "../").', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 3 – Thang điểm chấm Phiếu học tập số 2 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV2_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu liên kết tới vị trí khác trong cùng trang web (id) và tạo liên kết tới hình ảnh (08 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT3], [NLC1], [NLc], [NLS 3.1.NC1.a], [PC2] – Học sinh nắm vững quy trình 2 bước tạo liên kết nội bộ trong cùng trang bằng thuộc tính id và kí tự "#", cách đặt ảnh làm liên kết và giải đúng Câu hỏi củng cố trang 60.', first=0)
    add(b, '**b) Nội dung:** Học sinh đọc phần cuối mục 2b (SGK trang 60) và hoàn thành Phiếu học tập số 3 (bao gồm Câu hỏi củng cố cuối trang 60).', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 3 trình bày đúng 2 bước tạo liên kết tới mã định danh đoạn (#id), câu lệnh liên kết hình ảnh Mặt Trời (images/sun.png) và đoạn mã danh sách <ul> liên kết tới 3 tệp bài tập.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 4 – Thang điểm chấm Phiếu học tập số 3 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV3_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (07 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [NLC1], [NLC3], [NLc], [PC3] – Học sinh vận dụng lí thuyết vừa học để viết chính xác đoạn mã HTML cho Bài Luyện tập 1, Bài Luyện tập 2 (SGK trang 61) và bài toán đường dẫn tương đối trên Hình 10.1.', first=0)
    add(b, '**b) Nội dung:** Học sinh làm việc cá nhân hoàn thành Phiếu học tập số 4.', first=0)
    add(b, '**c) Sản phẩm:** Lời giải chính xác trên Phiếu học tập số 4.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 5 – Đáp án và thang điểm chấm Phiếu học tập số 4 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_T1, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (05 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [NLC3], [NLc], [PC2] – Học sinh phân tích trước yêu cầu của Nhiệm vụ 1, Nhiệm vụ 2 và phần Vận dụng (SGK trang 61), thiết kế sơ đồ mã định danh đoạn (id) để sẵn sàng thực hành trên máy tính ở Tiết 2 (Tiết PPCT: 22).', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 10.2 (SGK trang 61) và lập bảng thiết kế các thuộc tính id cùng thẻ <a> cho dự án trang web CLB.html:', first=0)
    add_figure(b, 'hinh_10_2.png')
    add(b, '**c) Sản phẩm:** Bản ghi thiết kế các mã định danh id và câu lệnh liên kết trong vở để sử dụng trực tiếp ở Tiết 2 Thực hành (Tiết PPCT: 22).', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Rubric đánh giá sản phẩm thực hành và vận dụng *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_T1, [4600, 4400], header=True, cell_sz=23))

    add_bai10_appendix(b, class_label='12/2', ppct_label='21', mode='tiet1')
    add_approval_section(b, sign_date_str=SIGN_DATE, gvhd_name=GVHD_NAME, gvgd_name=GVGD_NAME, sv_name=SV_NAME)

    used_figs = ['hinh_10_1.png', 'hinh_10_2.png']
    out_p = KHOI12_DIR / 'de_cuong_du_gio_12-2.docx'
    save_docx(b, out_p, used_figures=used_figs)
    THANH_PHAM_DIR.mkdir(parents=True, exist_ok=True)
    GOI_NOP_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin12_Bai10_Tiet1_Lop12-2.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / out_p.name)
    shutil.copy2(out_p, GOI_NOP_DIR / '01_De_cuong_du_gio_Tin12_Bai10_Tiet1_LT_TietPPCT21_Lop12-2.docx')
    print('Built:', out_p)


# ==============================================================================
# 2. FILE 02: TIẾT 2 – THỰC HÀNH (TIẾT PPCT: 22 – LỚP 12/1 – PHÒNG H.II.1)
# ==============================================================================
def build_bai10_tiet2():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TIẾT 2 – THỰC HÀNH – TIẾT PPCT: 22)',
        class_str='Tin học 12 – Lớp 12/1 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 2 (08h00 – 08h45) – Phòng H.II.1',
        lesson_row_str='Bài 4.10: Tạo liên kết (tiết 2 - TH) – Tiết PPCT: 22 (Tổng bài: 02 tiết – PPCT: 21, 22)',
        page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT (TIẾT 2 – THỰC HÀNH)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 21, 22] – Kế hoạch này chi tiết hoá Tiết thứ 2 của bài / Tiết PPCT: 22)',
        page1_class_line='Môn học: Tin học; Lớp: 12/1',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 2 – TH / Tiết PPCT: 22)',
        page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Tiết 2 (08h00 – 08h45) – Lớp 12/1 – Phòng H.II.1',
        page1_mode_line='Hình thức tổ chức: Dạy học thực hành tạo liên kết HTML và kiểm thử trên trình duyệt web',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 12, Chủ đề F (Giải quyết vấn đề với sự trợ giúp của máy tính – Tạo trang web) và Kế hoạch dạy học môn Tin học Khối 12 của Trường THPT Thuận Hoá (Tiết PPCT 22 – Bài 4.10: Tạo liên kết (tiết 2 - TH)):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức và kĩ năng thực hành (ký hiệu [KT1] – [KT3] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]:** Vận dụng chuẩn xác quy trình 2 bước tạo liên kết nội bộ trong cùng trang web (gắn thuộc tính định danh id="CLBTT" vào thẻ <table> và tạo thẻ <a href="#CLBTT">) kết hợp bổ sung khối văn bản giới thiệu theo đúng Hình 10.2 (Nhiệm vụ 1 – SGK trang 61).',
        '**[KT2]:** Tạo được tệp trang web mới thong_tin.html trong cùng thư mục với tệp CLB.html và thiết lập thành công liên kết hai chiều bằng đường dẫn tương đối giữa trang chủ CLB.html với trang thong_tin.html (Nhiệm vụ 2 và Bài Luyện tập 1 – SGK trang 61).',
        '**[KT3]:** Hoàn thiện được trang web giới thiệu bản thân có ít nhất 03 liên kết ngoài (Bài Luyện tập 2 – SGK trang 61) và mở rộng tệp CLB.html với các đoạn giới thiệu chi tiết từng câu lạc bộ ở cuối trang được liên kết trực tiếp từ tên từng câu lạc bộ trong bảng Lịch hoạt động các CLB Thể thao (Vận dụng – SGK trang 61).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Học sinh chủ động đọc hướng dẫn mục 3. Thực hành tạo liên kết (SGK trang 61), tự viết mã HTML và kiểm thử trên trình duyệt web.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Học sinh trao đổi theo nhóm đôi, hỗ trợ nhau phát hiện và sửa lỗi cú pháp đường dẫn hoặc mã định danh.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Học sinh biết cách tổ chức hệ thống liên kết điều hướng khoa học giữa các phần trong cùng trang và giữa nhiều trang web trong một website.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực sử dụng và quản lí các phương tiện CNTT và truyền thông [NLa]:** Quản lí đúng cấu trúc thư mục chứa các tệp .html trên máy tính và sử dụng thành thạo trình soạn thảo mã kết hợp trình duyệt web.',
        '**Năng lực giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông [NLc]:** Viết và gỡ lỗi thành công các thẻ <a href="...">, thuộc tính id="..." để hoàn thiện sản phẩm website Câu lạc bộ trường em.',
        '**Năng lực ứng dụng CNTT và truyền thông trong học và tự học [NLd]:** Chủ động tra cứu và kiểm thử nhanh sự hoạt động của các siêu liên kết.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Tạo lập, biên tập và liên kết thành công bộ trang web đa trang (.html) theo đúng yêu cầu kĩ thuật.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Học sinh kiên trì biên tập đầy đủ nội dung giới thiệu, danh sách thành tích và thông tin chi tiết của từng câu lạc bộ.',
        '**Phẩm chất trách nhiệm [PC2]:** Học sinh chấp hành nghiêm túc nội quy lớp học, lưu trữ và bảo quản thư mục bài làm đúng quy định.',
        '**Phẩm chất trung thực [PC3]:** Học sinh tự tay viết mã nguồn HTML và tự đánh giá trung thực sản phẩm theo Rubric CCĐG 6.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính có cài đặt trình soạn thảo văn bản HTML (Sublime Text, VS Code hoặc Notepad++) và trình duyệt web (Chrome hoặc Edge), máy chiếu, bảng viết, thư mục học liệu chứa tệp gốc CLB.html từ Bài 9.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam*, Bài 10: Tạo liên kết (mục 3. Thực hành tạo liên kết, Luyện tập và Vận dụng – trang 61).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 12 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 22 – Bài 4.10: Tạo liên kết (tiết 2 - TH)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 2 – Thực hành (45 phút – Tiết thứ 2 của bài / Tiết PPCT: 22)**', first=0, keep=True)
    ov2_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng và Tiết PPCT', 'Mục tiêu', 'Sản phẩm và Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Kích hoạt kĩ năng tạo liên kết và Mở tệp dự án CLB.html)', '05 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT1], [KT2],\n[NLC1], [NLa]', 'Mở sẵn tệp CLB.html trên trình soạn thảo và trả lời đúng 3 câu hỏi thao tác cốt lõi.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới (Hướng dẫn và Thực hành cốt lõi SGK tr. 61)\n- 2.1. Nhiệm vụ 1: Tạo liên kết trong cùng trang CLB.html (Hình 10.2 – 12 phút)\n- 2.2. Nhiệm vụ 2: Tạo liên kết sang trang khác thong_tin.html (12 phút)', '24 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT1], [KT2],\n[NLC1], [NLC2],\n[NLa], [NLc], [NLd],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Tệp CLB.html (có liên kết #CLBTT và liên kết sang thong_tin.html) và tệp thong_tin.html.\nCông cụ đánh giá: CCĐG 6 (Rubric thực hành).'],
        ['3. Hoạt động 3: Luyện tập\n(Thực hành Bài Luyện tập 1 và 2 SGK tr. 61)', '08 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT2], [KT3],\n[NLC1], [NLC3],\n[NLc], [PC3]', 'Liên kết quay về CLB.html trong thong_tin.html và tệp gioi_thieu_ban_than.html.\nCông cụ đánh giá: CCĐG 5 và CCĐG 6.'],
        ['4. Hoạt động 4: Vận dụng\n(Tạo liên kết từ tên từng CLB trong bảng đến đoạn giới thiệu cuối trang – SGK tr. 61)', '08 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT3], [NLC2],\n[NLC3], [NLc],\n[PC1], [PC2]', 'Tệp CLB.html hoàn chỉnh có 5 đoạn giới thiệu CLB gắn id và liên kết từ bảng xuống.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov2_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [NLC1], [NLa] – Học sinh tái hiện nhanh cú pháp tạo liên kết nội bộ và liên kết tương đối đã học ở Tiết 1 (Tiết PPCT: 21), đồng thời mở sẵn tệp CLB.html để bắt đầu thực hành.', first=0)
    add(b, '**b) Nội dung:** Giáo viên yêu cầu học sinh mở thư mục thực hành Bai10_TaoLienKet (chứa tệp CLB.html từ Bài 9) và trả lời nhanh 03 câu hỏi kiểm tra kĩ năng nền tảng.', first=0)
    add(b, '**c) Sản phẩm:** Cửa sổ trình soạn thảo mã HTML mở sẵn tệp CLB.html và câu trả lời đúng 03 câu hỏi thao tác.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 1 – Bảng kiểm quan sát chuẩn bị thực hành *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd1_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên yêu cầu học sinh lớp 12/1 mở thư mục Bai10_TaoLienKet trên máy tính, mở tệp CLB.html bằng trình soạn thảo mã và trình duyệt web, sau đó trả lời nhanh 03 câu hỏi:\n'
            '  + Câu 1: Muốn tạo liên kết nhảy đến một bảng ở cuối trang web hiện tại, ta cần thêm thuộc tính gì vào thẻ <table> và viết giá trị href của thẻ <a> như thế nào?\n'
            '  + Câu 2: Khi tệp thong_tin.html nằm cùng thư mục với tệp CLB.html, ta viết thuộc tính href như thế nào để liên kết từ CLB.html sang thong_tin.html?\n'
            '  + Câu 3: Sau khi sửa mã HTML, ta cần nhấn tổ hợp phím nào để lưu và làm mới trang trên trình duyệt?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở tệp CLB.html trên máy và giơ tay trả lời nhanh 03 câu hỏi.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh trả lời tại chỗ và chiếu cấu trúc tệp CLB.html gốc lên màn hình.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét và giao nhiệm vụ thực hành trọng tâm của Tiết 2 (Tiết PPCT: 22): Hoàn thành Nhiệm vụ 1, Nhiệm vụ 2, Luyện tập và Vận dụng trong SGK trang 61.',
            '**Kiến thức thao tác cốt lõi cần nhớ khi thực hành:**\n'
            '1. **Liên kết trong cùng trang (2 bước):**\n'
            '   - Bước 1 (Gắn cờ đích): Thêm thuộc tính id="Tên_định_danh" vào thẻ mở của phần tử đích (ví dụ: <table id="CLBTT" border="1">).\n'
            '   - Bước 2 (Tạo liên kết): Đặt dấu "#" trước tên định danh trong thuộc tính href (ví dụ: <a href="#CLBTT">...</a>).\n'
            '2. **Liên kết tới tệp cùng thư mục:** Chỉ cần viết tên tệp đích trong href (ví dụ: <a href="thong_tin.html">...</a>).\n'
            '3. **Lưu và kiểm thử:** Nhấn Ctrl + S trong trình soạn thảo, sau đó chuyển sang trình duyệt nhấn F5 (hoặc Ctrl + R) và nháy chuột vào liên kết để kiểm tra.'
        ]
    ]
    b.append(table(hd1_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới – Hướng dẫn và Thực hành Nhiệm vụ 1, Nhiệm vụ 2 (24 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tạo liên kết trong cùng trang – Hình 10.2 SGK trang 61 (12 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [NLa], [NLc], [NLS 3.1.NC1.a], [PC1] – Học sinh bổ sung thành công phần giới thiệu hoạt động ngoại khoá vào đầu tệp CLB.html theo Hình 10.2, gắn mã định danh id="CLBTT" cho bảng và tạo liên kết nội bộ nhảy tới bảng.', first=0)
    add(b, '**b) Nội dung:** Học sinh đọc hướng dẫn Nhiệm vụ 1 (SGK trang 61), quan sát Hình 10.2 và thực hành trực tiếp trên tệp CLB.html:', first=0)
    add_figure(b, 'hinh_10_2.png')
    add(b, '**c) Sản phẩm:** Tệp CLB.html hiển thị đầy đủ nội dung Hình 10.2 ở đầu trang; khi nháy chuột vào dòng "Theo dõi lịch hoạt động các CLB Thể thao", trình duyệt cuộn ngay tới bảng Lịch hoạt động các CLB Thể thao.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Tiêu chí 1 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv1_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 10.2 (SGK trang 61), phân tích nhanh 3 bước thực hiện trong 02 phút và yêu cầu học sinh thực hành trên máy tính trong 08 phút:\n'
            '  + Bước 1: Mở tệp CLB.html và thêm nội dung giới thiệu hoạt động ngoại khoá (gồm tiêu đề, đoạn văn và danh sách 4 thành tích) như Hình 10.2.\n'
            '  + Bước 2: Thêm thuộc tính id="CLBTT" vào thẻ <table> của bảng Lịch hoạt động các CLB Thể thao.\n'
            '  + Bước 3: Thêm thẻ <a href="#CLBTT"> cho dòng chữ "Theo dõi lịch hoạt động các CLB Thể thao".\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh gõ mã HTML vào tệp CLB.html, nhấn Ctrl + S lưu tệp và mở trên trình duyệt để kiểm thử. Giáo viên quan sát và hướng dẫn học sinh kiểm tra sự trùng khớp chữ hoa/chữ thường giữa id="CLBTT" và href="#CLBTT".\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên chọn 02 bài thực hành tốt chiếu lên màn hình chung, mời học sinh thao tác nháy liên kết và giải thích tác dụng của thuộc tính id.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét, chấm điểm theo Tiêu chí 1 của Rubric CCĐG 6 (3,5 điểm).',
            '**1. Thực hành Nhiệm vụ 1: Tạo liên kết trong cùng trang (SGK trang 61):**\n'
            '- **Bước 1:** Mở tệp CLB.html và chèn thêm khối nội dung giới thiệu vào đầu phần <body> theo đúng Hình 10.2:\n'
            '  <h2>Giới thiệu các hoạt động ngoại khoá</h2>\n'
            '  <p>Trường THPT Nguyễn Bỉnh Khiêm có bề dày thành tích về học tập và các hoạt động thể thao, văn nghệ của thành phố. Các câu lạc bộ ngoại khoá hoạt động sôi nổi và luôn được nhà trường tạo điều kiện để sinh hoạt. Những năm qua, trường đạt nhiều thành tích trong các cuộc thi cấp thành phố, quận:</p>\n'
            '  <ul>\n'
            '    <li>03 HCV cấp thành phố môn Bơi lội</li>\n'
            '    <li>02 HCV, 06 HCB và 05 HCĐ cấp thành phố các môn Võ thuật</li>\n'
            '    <li>01 giải ba môn Bóng đá nam cấp thành phố</li>\n'
            '    <li>01 giải nhì, 02 giải ba Hội thi biểu diễn âm nhạc quận</li>\n'
            '  </ul>\n'
            '- **Bước 2:** Thêm mã định danh id="CLBTT" cho bảng Lịch hoạt động các CLB Thể thao đã tạo ở bài trước:\n'
            '  <table id="CLBTT" border="1">\n'
            '    <caption>Lịch hoạt động các CLB Thể thao</caption>\n'
            '- **Bước 3:** Thêm liên kết cho dòng "Theo dõi lịch hoạt động các CLB Thể thao" ngay dưới danh sách thành tích:\n'
            '  <p><a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a></p>'
        ]
    ]
    b.append(table(nv1_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tạo liên kết sang trang khác – SGK trang 61 (12 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLC2], [NLa], [NLc], [NLd], [NLS 3.1.NC1.a], [PC1], [PC2] – Học sinh tạo được tệp thong_tin.html cùng thư mục với CLB.html và thiết lập thành công đường liên kết tương đối từ trang chủ CLB.html sang trang thong_tin.html.', first=0)
    add(b, '**b) Nội dung:** Học sinh thực hiện đúng 2 bước của Nhiệm vụ 2 (SGK trang 61) trên máy tính.', first=0)
    add(b, '**c) Sản phẩm:** Tệp thong_tin.html chứa thông tin giới thiệu các câu lạc bộ và liên kết <a href="thong_tin.html">Thông tin của các CLB</a> hoạt động chính xác trên trang CLB.html.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Tiêu chí 2 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv2_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên nêu yêu cầu Nhiệm vụ 2 (SGK trang 61) và hướng dẫn học sinh thực hiện 2 bước trên máy tính:\n'
            '  + Bước 1: Trong cùng thư mục với tệp CLB.html, tạo tệp mới tên là thong_tin.html và biên soạn nội dung giới thiệu thông tin về các câu lạc bộ.\n'
            '  + Bước 2: Trong tệp CLB.html, viết đoạn mã tạo liên kết bằng đường dẫn tương đối đến tệp thong_tin.html.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh tạo tệp thong_tin.html, lưu cùng thư mục với CLB.html, sau đó thêm thẻ <a href="thong_tin.html">Thông tin của các CLB</a> vào CLB.html và kiểm thử trên trình duyệt.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh trình chiếu kết quả nháy chuột chuyển từ trang CLB.html sang trang thong_tin.html và giải thích vì sao chỉ cần viết href="thong_tin.html".\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét theo Tiêu chí 2 của Rubric CCĐG 6 và chuyển sang hoạt động Luyện tập.',
            '**2. Thực hành Nhiệm vụ 2: Tạo liên kết sang trang khác (SGK trang 61):**\n'
            '- **Bước 1:** Trong cùng thư mục chứa tệp CLB.html, tạo tệp mới thong_tin.html có cấu trúc chuẩn HTML5 giới thiệu quy chế, thời gian sinh hoạt và quyền lợi khi tham gia các Câu lạc bộ Thể thao và Nghệ thuật của trường.\n'
            '- **Bước 2:** Mở tệp CLB.html và thêm đoạn mã liên kết bằng đường dẫn tương đối:\n'
            '  <p><a href="thong_tin.html">Thông tin của các CLB</a></p>\n'
            '- **Kiểm thử:** Mở CLB.html trên trình duyệt web, nháy chuột vào dòng "Thông tin của các CLB", trình duyệt chuyển hướng sang hiển thị nội dung tệp thong_tin.html.'
        ]
    ]
    b.append(table(nv2_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (08 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'heading')
    add(b, '**a) Mục tiêu:** [KT2], [KT3], [NLC1], [NLC3], [NLc], [PC3] – Học sinh hoàn thành Bài Luyện tập 1 (tạo liên kết từ thong_tin.html quay về CLB.html) và Bài Luyện tập 2 (tạo trang web giới thiệu bản thân có ít nhất 03 liên kết ngoài) trong SGK trang 61.', first=0)
    add(b, '**b) Nội dung:** Học sinh thực hành viết mã trên máy tính và ghi câu lệnh kết quả vào Phiếu học tập số 4.', first=0)
    add(b, '**c) Sản phẩm:** Liên kết hai chiều hoàn chỉnh giữa CLB.html và thong_tin.html, tệp gioi_thieu_ban_than.html và Phiếu học tập số 4 đã điền đáp án.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 5 và CCĐG 6 *(Đính kèm tại Mục IV. Phụ lục 1 và Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (08 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [NLC2], [NLC3], [NLc], [PC1], [PC2] – Học sinh giải quyết trọn vẹn bài tập Vận dụng (SGK trang 61): viết thêm các đoạn giới thiệu cho từng câu lạc bộ ở dưới bảng Lịch hoạt động các CLB Thể thao và đặt liên kết từ tên câu lạc bộ trong bảng đến đoạn giới thiệu tương ứng.', first=0)
    add(b, '**b) Nội dung:** Học sinh bổ sung 5 đoạn giới thiệu có gắn thuộc tính id riêng cho từng CLB ở cuối tệp CLB.html và gắn thẻ <a href="#..."> cho tên các CLB trong bảng.', first=0)
    add(b, '**c) Sản phẩm:** Tệp CLB.html hoàn chỉnh toàn bộ yêu cầu của Bài 10, nháy vào tên bất kì CLB nào trong bảng đều cuộn chính xác tới đoạn giới thiệu chi tiết của CLB đó.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add_bai10_appendix(b, class_label='12/1', ppct_label='22', mode='tiet2')
    add_approval_section(b, sign_date_str=SIGN_DATE, gvhd_name=GVHD_NAME, gvgd_name=GVGD_NAME, sv_name=SV_NAME)

    used_figs = ['hinh_10_2.png']
    out_p = KHOI12_DIR / 'de_cuong_du_gio_12-1.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin12_Bai10_Tiet2_Lop12-1.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / out_p.name)
    shutil.copy2(out_p, GOI_NOP_DIR / '02_De_cuong_du_gio_Tin12_Bai10_Tiet2_TH_TietPPCT22_Lop12-1.docx')
    print('Built:', out_p)


# ==============================================================================
# 3. FILE 03: BẢN GỘP TRỌN BỘ CẢ BÀI 10 (02 TIẾT = 90 PHÚT – TIẾT PPCT: 21, 22)
# ==============================================================================
def build_bai10_cabai():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 10: TẠO LIÊN KẾT (TRỌN BỘ 02 TIẾT – TIẾT PPCT: 21, 22)',
        class_str='Tin học 12 – Lớp 12/2 và Lớp 12/1 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Sáu, ngày 13/11/2026 – Tiết 3 (12/2, Phòng H.II.2) và Tiết 2 (12/1, Phòng H.II.1)',
        lesson_row_str='Bài 4.10: Tạo liên kết (02 tiết: Tiết 1 Lí thuyết [PPCT 21] và Tiết 2 Thực hành [PPCT 22])',
        page1_title='TÊN BÀI DẠY: BÀI 10. TẠO LIÊN KẾT',
        page1_subtitle='(Thời lượng thực hiện toàn bài: 02 tiết = 90 phút – Tiết PPCT: 21, 22 theo Kế hoạch dạy học môn Tin học Khối 12)',
        page1_class_line='Môn học: Tin học; Lớp: 12/2 (Tiết 1 – LT) và 12/1 (Tiết 2 – TH)',
        page1_time_line='Thời gian thực hiện: 02 tiết (90 phút – Tiết 1: Lí thuyết [PPCT 21]; Tiết 2: Thực hành [PPCT 22])',
        page1_sched_line='Lịch dự giờ: Thứ Sáu, ngày 13/11/2026 – Lớp 12/2 (Tiết 3: 09h00–09h45, P. H.II.2) và Lớp 12/1 (Tiết 2: 08h00–08h45, P. H.II.1)',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết (Tiết 1) kết hợp thực hành (Tiết 2)',
        gvgd_name=GVGD_NAME,
        gvhd_name=GVHD_NAME,
        sv_name=SV_NAME
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 12, Chủ đề F (Giải quyết vấn đề với sự trợ giúp của máy tính – Tạo trang web) và Kế hoạch dạy học môn Tin học Khối 12 của Trường THPT Thuận Hoá (Bài 4.10: Tạo liên kết – 02 tiết, Tiết PPCT: 21, 22):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau bài học này (02 tiết = 90 phút), học sinh đạt được các yêu cầu về kiến thức và kĩ năng (ký hiệu [KT1] – [KT4] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]:** Phát biểu được khái niệm siêu văn bản (hypertext), siêu liên kết (hyperlink), cú pháp thẻ <a> và phân biệt rõ đường dẫn tuyệt đối với đường dẫn tương đối (SGK trang 57–58).',
        '**[KT2]:** Viết đúng câu lệnh HTML tạo liên kết ngoài (đường dẫn tuyệt đối) và liên kết tới các trang web trong cùng website theo 4 trường hợp cấu trúc thư mục ở Hình 10.1 (cùng thư mục, dưới 1 cấp, dưới nhiều cấp, và lên thư mục mức trên bằng "../") (SGK trang 58–60).',
        '**[KT3]:** Thiết lập được liên kết tới vị trí khác trong cùng trang web bằng thuộc tính id và tham chiếu href="#mã_định_danh_đoạn", tạo liên kết từ hình ảnh bằng thẻ <img> đặt trong thẻ <a>, và kết hợp thẻ <a> với danh sách <ul><li> (SGK trang 60).',
        '**[KT4]:** Thực hành hoàn thiện bộ trang web Câu lạc bộ (CLB.html, thong_tin.html) và trang giới thiệu bản thân (gioi_thieu_ban_than.html) đáp ứng trọn vẹn Nhiệm vụ 1 (Hình 10.2), Nhiệm vụ 2, phần Luyện tập và phần Vận dụng (SGK trang 61).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Học sinh chủ động đọc hiểu Sách giáo khoa Tin học 12 trang 57–61, tự hoàn thành Phiếu học tập và các nhiệm vụ thực hành trên máy tính.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Học sinh thảo luận nhóm đôi hiệu quả để phân tích cấu trúc cây thư mục Hình 10.1 và kiểm thử chéo sản phẩm trang web.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Học sinh tự phát hiện, gỡ lỗi sai đường dẫn tương đối hoặc sai mã định danh id khi liên kết không hoạt động.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực sử dụng và quản lí các phương tiện CNTT và truyền thông [NLa]:** Tổ chức và quản lí khoa học thư mục dự án web gồm nhiều tệp .html và thư mục ảnh trên máy tính.',
        '**Năng lực ứng xử phù hợp trong môi trường số [NLb]:** Có trách nhiệm khi chia sẻ và gắn liên kết tới các nguồn tài nguyên hợp pháp, an toàn trên không gian mạng.',
        '**Năng lực giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông [NLc]:** Vận dụng phối hợp các thẻ HTML (<a>, <img>, <ul>, <li>, <table>) để xây dựng website đa trang có điều hướng mạch lạc.',
        '**Năng lực ứng dụng CNTT và truyền thông trong học và tự học [NLd]:** Sử dụng thành thạo trình soạn thảo mã và trình duyệt để xây dựng trang web giới thiệu bản thân và câu lạc bộ trường học.',
        '**Năng lực hợp tác trong môi trường số [NLe]:** Phối hợp nhóm đôi hoàn thành tốt nhiệm vụ học tập.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Thiết kế, biên tập mã nguồn HTML và liên kết hoàn chỉnh các trang nội dung số.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Học sinh tích cực ghi chép ở tiết lí thuyết và kiên trì hoàn thiện mã nguồn ở tiết thực hành.',
        '**Phẩm chất trách nhiệm [PC2]:** Học sinh chấp hành nghiêm túc nội quy lớp học, lưu trữ tệp đúng quy định và hoàn thành sản phẩm đúng thời hạn.',
        '**Phẩm chất trung thực [PC3]:** Học sinh tự giác thực hiện bài làm và đánh giá công tâm theo Rubric CCĐG 6.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính của giáo viên và học sinh có cài đặt trình soạn thảo mã HTML và trình duyệt web, máy chiếu, bảng viết, bộ Phiếu học tập và Rubric đánh giá.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 12 – Định hướng Tin học ứng dụng (Bộ sách Kết nối tri thức với cuộc sống), Nhà xuất bản Giáo dục Việt Nam*, Bài 10: Tạo liên kết (trang 57 đến trang 61).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 12 năm học 2026–2027 của Trường THPT Thuận Hoá* (Bài 4.10: Tạo liên kết – 02 tiết: Tiết PPCT 21 [tiết 1 - LT] và Tiết PPCT 22 [tiết 2 - TH]).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học toàn Bài 10 (02 tiết = 90 phút – Tiết PPCT: 21, 22)**', first=0, keep=True)
    ov_all = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng và Phân bổ Tiết PPCT', 'Mục tiêu', 'Sản phẩm và Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(So sánh đọc trang web với đọc sách, báo giấy – SGK tr. 57)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 21)', '[KT1], [NLC1],\n[PC1]', 'Câu trả lời miệng nêu bật vai trò của siêu liên kết.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Siêu văn bản, thẻ <a> và Đường dẫn (12p – Tiết 1)\n- 2.2. Nhiệm vụ 2: Liên kết ngoài và Liên kết trong theo Hình 10.1 (15p – Tiết 1)\n- 2.3. Nhiệm vụ 3: Liên kết vị trí cùng trang (id) và Liên kết ảnh (13p – Tiết 1)\n- 2.4. Nhiệm vụ 4: Thực hành Nhiệm vụ 1 (Hình 10.2) và Nhiệm vụ 2 (25p – Tiết 2)', '65 phút\n(40p ở Tiết 1 [PPCT 21]\n+ 25p đầu Tiết 2 [PPCT 22])', '[KT1], [KT2],\n[KT3], [KT4],\n[NLC1], [NLC2],\n[NLa], [NLb], [NLc],\n[NLd], [NLe],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, 2, 3 (Tiết 1) và hai tệp thực hành CLB.html, thong_tin.html (Tiết 2).\nCông cụ đánh giá: CCĐG 2, 3, 4 và CCĐG 6.'],
        ['3. Hoạt động 3: Luyện tập\n(Giải và thực hành Bài Luyện tập 1 và 2 SGK trang 61)', '10 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT2], [KT3], [KT4],\n[NLC1], [NLC3],\n[NLc], [PC3]', 'Phiếu học tập số 4, liên kết quay về CLB.html và tệp gioi_thieu_ban_than.html.\nCông cụ đánh giá: CCĐG 5.'],
        ['4. Hoạt động 4: Vận dụng\n(Thực hành liên kết từ tên từng CLB trong bảng đến đoạn giới thiệu cuối trang – SGK tr. 61)', '10 phút\n(Tiết 2 của bài /\nTiết PPCT: 22)', '[KT3], [KT4],\n[NLC2], [NLC3],\n[NLc], [PC1], [PC2]', 'Tệp CLB.html hoàn chỉnh có 5 đoạn giới thiệu CLB gắn id và liên kết từ bảng xuống.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov_all, [2700, 1500, 1900, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [PC1] – Học sinh nhận biết sự khác biệt giữa đọc thông tin trên trang web với đọc sách, báo giấy và nêu được vai trò của siêu liên kết.', first=0)
    add(b, '**b) Nội dung:** Học sinh thảo luận cặp đôi trả lời câu hỏi Khởi động trong Sách giáo khoa Tin học 12 trang 57.', first=0)
    add(b, '**c) Sản phẩm:** Câu trả lời miệng của học sinh chỉ ra vai trò của siêu liên kết (hyperlink).', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 1 – Bảng kiểm quan sát Hoạt động Mở đầu *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD1_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (65 phút: 40 phút thuộc Tiết 1 [Tiết PPCT: 21] + 25 phút đầu thuộc Tiết 2 [Tiết PPCT: 22])', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu siêu văn bản, siêu liên kết, thẻ <a> và hai loại đường dẫn (12 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [NLb], [PC1] – Học sinh phát biểu được khái niệm siêu văn bản, cú pháp thẻ <a> và phân biệt được đường dẫn tuyệt đối với đường dẫn tương đối.', first=0)
    add(b, '**b) Nội dung:** Học sinh nghiên cứu mục 1. Siêu văn bản và đường dẫn (SGK trang 57–58) và hoàn thành Phiếu học tập số 1.', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 1 hoàn chỉnh khái niệm và trả lời đúng Câu hỏi củng cố trang 58.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 2 – Thang điểm chấm Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV1_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu liên kết ngoài và liên kết tới trang web trong cùng website theo Hình 10.1 (15 phút – Tiết 1 của bài / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLC2], [NLC3], [NLc], [NLe], [PC1], [PC2] – Học sinh viết đúng câu lệnh HTML cho liên kết ngoài và 4 trường hợp liên kết tương đối trong cùng website theo Hình 10.1.', first=0)
    add(b, '**b) Nội dung:** Học sinh đọc mục 2a, 2b (SGK trang 58–60), quan sát Hình 10.1 (SGK trang 59) và hoàn thành Phiếu học tập số 2:', first=0)
    add_figure(b, 'hinh_10_1.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 2 hoàn chỉnh các câu lệnh thẻ <a> theo cấu trúc Hình 10.1.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 3 – Thang điểm chấm Phiếu học tập số 2 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV2_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu liên kết tới vị trí khác trong cùng trang web (id) và tạo liên kết tới hình ảnh (13 phút cuối Tiết 1 / Tiết PPCT: 21)', 'sub')
    add(b, '**a) Mục tiêu:** [KT3], [NLC1], [NLc], [NLS 3.1.NC1.a], [PC2] – Học sinh nắm vững quy trình 2 bước tạo liên kết nội bộ bằng thuộc tính id và tham chiếu "#id", cách đặt hình ảnh làm liên kết và giải đúng Câu hỏi củng cố trang 60.', first=0)
    add(b, '**b) Nội dung:** Học sinh nghiên cứu phần cuối mục 2b (SGK trang 60) và hoàn thành Phiếu học tập số 3.', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 3 hoàn chỉnh quy trình 2 bước tạo mã định danh đoạn và đoạn mã danh sách liên kết.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 4 – Thang điểm chấm Phiếu học tập số 3 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV3_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.4. Nhiệm vụ 4: Thực hành tạo liên kết trong cùng trang (Hình 10.2) và liên kết sang trang khác (25 phút đầu Tiết 2 / Tiết PPCT: 22)', 'sub')
    add(b, '**a) Mục tiêu:** [KT4], [NLC1], [NLC2], [NLa], [NLc], [NLd], [NLS 3.1.NC1.a], [PC1], [PC2] – Học sinh thực hành hoàn thiện Nhiệm vụ 1 (Hình 10.2) và Nhiệm vụ 2 (SGK trang 61) trên bộ tệp CLB.html và thong_tin.html.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 10.2 (SGK trang 61) và thực hành trực tiếp trên máy tính:', first=0)
    add_figure(b, 'hinh_10_2.png')
    add(b, '**c) Sản phẩm:** Hai tệp CLB.html và thong_tin.html hoạt động trơn tru cả liên kết nội bộ (#CLBTT) và liên kết chuyển trang.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Rubric đánh giá sản phẩm thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV4_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (10 phút – Tiết 2 của bài / Tiết PPCT: 22)', 'heading')
    add(b, '**a) Mục tiêu:** [KT2], [KT3], [KT4], [NLC1], [NLC3], [NLc], [PC3] – Học sinh giải quyết trọn vẹn Bài Luyện tập 1 và Bài Luyện tập 2 (SGK trang 61).', first=0)
    add(b, '**b) Nội dung:** Học sinh hoàn thành Phiếu học tập số 4 và thực hành bổ sung liên kết quay về trang chủ CLB.html cùng tệp gioi_thieu_ban_than.html.', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 4 và các tệp HTML đã kiểm thử thành công trên trình duyệt.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 5 – Đáp án và thang điểm chấm Phiếu học tập số 4 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (10 phút cuối Tiết 2 / Tiết PPCT: 22)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [KT4], [NLC2], [NLC3], [NLc], [PC1], [PC2] – Học sinh vận dụng kĩ thuật tạo mã định danh đoạn (id) để liên kết từ tên từng câu lạc bộ trong bảng Lịch hoạt động các CLB Thể thao xuống đoạn giới thiệu tương ứng ở cuối trang CLB.html (SGK trang 61).', first=0)
    add(b, '**b) Nội dung:** Học sinh thực hành phần Vận dụng (SGK trang 61) ngay trên tệp CLB.html.', first=0)
    add(b, '**c) Sản phẩm:** Tệp CLB.html hoàn chỉnh có 5 đoạn giới thiệu CLB gắn thuộc tính id và các liên kết điều hướng chính xác từ bảng xuống.', first=0)
    add(b, '**Công cụ kiểm tra, đánh giá sản phẩm học tập:** CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_BAI10, [4600, 4400], header=True, cell_sz=23))

    add_bai10_appendix(b, class_label='12/2 (Tiết 1 – LT) và 12/1 (Tiết 2 – TH)', ppct_label='21, 22', mode='cabai')
    add_approval_section(b, sign_date_str=SIGN_DATE, gvhd_name=GVHD_NAME, gvgd_name=GVGD_NAME, sv_name=SV_NAME)

    used_figs = ['hinh_10_1.png', 'hinh_10_2.png']
    out_p = KHOI12_DIR / '03_De_cuong_du_gio_Tin12_Bai10_CaBai_2Tiet_TietPPCT21-22.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin12_Bai10_CaBai_2Tiet.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / out_p.name)
    print('Built:', out_p)


# ==============================================================================
# 4. FILE 04: BỘ PHIẾU HỌC TẬP VÀ RUBRIC RỜI ĐỂ IN PHÁT CHO HỌC SINH
# ==============================================================================
def build_bai10_worksheets():
    b = node('body')
    tabline(b, 'Trường: THPT Thuận Hoá – Tổ: Tin học', f'Giáo viên giảng dạy: {GVGD_NAME}')
    tabline(b, f'Giáo viên hướng dẫn: {GVHD_NAME}', f'Sinh viên thực hiện: {SV_NAME} (24S1020030)')
    add(b, 'BỘ PHIẾU HỌC TẬP, ĐÁP ÁN VÀ RUBRIC ĐÁNH GIÁ', 'title')
    add(b, 'BÀI 10: TẠO LIÊN KẾT (SGK TIN HỌC 12 KNTT, TRANG 57–61 – TIẾT PPCT: 21, 22)', 'title')
    add(b, '(Dùng cho Tiết 1 Lí thuyết tại Lớp 12/2 và Tiết 2 Thực hành tại Lớp 12/1 – Ngày 13/11/2026)', align='center', first=0, italic=True)
    add_bai10_appendix(b, class_label='12/2 và 12/1', ppct_label='21, 22', mode='cabai')
    add_approval_section(
        b,
        sign_date_str=SIGN_DATE,
        gvhd_name=GVHD_NAME,
        gvgd_name=GVGD_NAME,
        sv_name=SV_NAME,
        section_title='XÁC NHẬN CỦA GIÁO VIÊN HƯỚNG DẪN, GIÁO VIÊN GIẢNG DẠY VÀ SINH VIÊN THỰC HIỆN'
    )
    out_p = KHOI12_DIR / '04_Phieu_hoc_tap_va_Rubric_Tin12_Bai10.docx'
    save_docx(b, out_p, is_handout=True)
    shutil.copy2(out_p, GOI_NOP_DIR / out_p.name)
    print('Built:', out_p)


# ==============================================================================
# 5. BỘ HỌC LIỆU HTML THỰC HÀNH BÀI 10 (KẾ THỪA BÀI 9 và CHUẨN SGK TRANG 59–61)
# ==============================================================================
def build_html_learning_materials():
    html_dir = KHOI12_DIR / 'Hoc_lieu_HTML'
    html_dir.mkdir(parents=True, exist_ok=True)

    # 1. Tệp gốc CLB_Bai9_Goc.html (để học sinh mở đầu giờ Tiết 2)
    clb_goc = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Câu lạc bộ ngoại khoá - Bản gốc từ Bài 9</title>
</head>
<body>
  <h1>Câu lạc bộ trường em</h1>
  <h2>Danh sách các Câu lạc bộ</h2>
  <ul>
    <li>THỂ THAO
      <ol>
        <li>Bóng đá</li>
        <li>Bóng chuyền</li>
        <li>Bóng rổ</li>
        <li>Võ thuật
          <ol type="a">
            <li>Karatedo</li>
            <li>Taekwondo</li>
            <li>Vovinam</li>
          </ol>
        </li>
        <li>Bơi</li>
      </ol>
    </li>
    <li>NGHỆ THUẬT
      <ol>
        <li>Mĩ thuật</li>
        <li>Nhiếp ảnh</li>
        <li>Âm nhạc
          <ol type="a">
            <li>Thanh nhạc</li>
            <li>Piano</li>
            <li>Violin</li>
          </ol>
        </li>
        <li>Khiêu vũ</li>
      </ol>
    </li>
  </ul>

  <table border="1" style="border-collapse: collapse; width: 90%;">
    <caption>Lịch hoạt động các CLB Thể thao</caption>
    <tr>
      <th rowspan="2">Bộ môn</th>
      <th rowspan="2">GV phụ trách</th>
      <th colspan="5">Ngày</th>
    </tr>
    <tr>
      <th>Thứ 2</th><th>Thứ 3</th><th>Thứ 4</th><th>Thứ 5</th><th>Thứ 6</th>
    </tr>
    <tr><td>Bóng đá</td><td>Thầy Khải</td><td></td><td>X</td><td></td><td></td><td>X</td></tr>
    <tr><td>Bóng chuyền</td><td>Cô Lan</td><td>X</td><td></td><td></td><td></td><td></td></tr>
    <tr><td>Bóng rổ</td><td>Thầy Vũ</td><td></td><td></td><td>X</td><td></td><td></td></tr>
    <tr><td>Võ thuật</td><td>Thầy Hoàng - Karatedo<br>Thầy Đức - Taekwondo<br>Thầy Thắng - Vovinam</td><td>X</td><td>X</td><td></td><td>X</td><td></td></tr>
    <tr><td>Bơi</td><td>Cô Bích</td><td></td><td></td><td>X</td><td></td><td>X</td></tr>
  </table>
</body>
</html>
"""
    (html_dir / '01_CLB_Bai9_Goc.html').write_text(clb_goc, encoding='utf-8')

    # 2. Tệp CLB.html hoàn chỉnh sau Bài 10 (Nhiệm vụ 1 + Nhiệm vụ 2 + Vận dụng SGK trang 61)
    clb_hoan_chinh = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Giới thiệu các hoạt động ngoại khoá - Bài 10 Hoàn chỉnh</title>
</head>
<body>
  <!-- Nhiệm vụ 1 (Hình 10.2 SGK trang 61): Phần giới thiệu bổ sung vào đầu trang -->
  <h2>Giới thiệu các hoạt động ngoại khoá</h2>
  <p>Trường THPT Nguyễn Bỉnh Khiêm có bề dày thành tích về học tập và các hoạt động thể thao, văn nghệ của thành phố.
  Các câu lạc bộ ngoại khoá hoạt động sôi nổi và luôn được nhà trường tạo điều kiện để sinh hoạt.
  Những năm qua, trường đạt nhiều thành tích trong các cuộc thi cấp thành phố, quận:</p>
  <ul>
    <li>03 HCV cấp thành phố môn Bơi lội</li>
    <li>02 HCV, 06 HCB và 05 HCĐ cấp thành phố các môn Võ thuật</li>
    <li>01 giải ba môn Bóng đá nam cấp thành phố</li>
    <li>01 giải nhì, 02 giải ba Hội thi biểu diễn âm nhạc quận</li>
  </ul>
  <p><a href="#CLBTT">Theo dõi lịch hoạt động các CLB Thể thao</a></p>
  <!-- Nhiệm vụ 2 (SGK trang 61): Liên kết tương đối sang trang thong_tin.html -->
  <p><a href="thong_tin.html">Thông tin của các CLB</a></p>

  <hr>
  <h2>Danh sách các Câu lạc bộ</h2>
  <ul>
    <li>THỂ THAO
      <ol>
        <li>Bóng đá</li>
        <li>Bóng chuyền</li>
        <li>Bóng rổ</li>
        <li>Võ thuật</li>
        <li>Bơi</li>
      </ol>
    </li>
    <li>NGHỆ THUẬT
      <ol>
        <li>Mĩ thuật</li>
        <li>Nhiếp ảnh</li>
        <li>Âm nhạc</li>
        <li>Khiêu vũ</li>
      </ol>
    </li>
  </ul>

  <!-- Nhiệm vụ 1 (Bước 2) & Vận dụng (SGK trang 61): Bảng có id="CLBTT" và liên kết từ tên từng CLB xuống đoạn giới thiệu -->
  <table id="CLBTT" border="1" style="border-collapse: collapse; width: 90%;">
    <caption>Lịch hoạt động các CLB Thể thao</caption>
    <tr>
      <th rowspan="2">Bộ môn</th>
      <th rowspan="2">GV phụ trách</th>
      <th colspan="5">Ngày</th>
    </tr>
    <tr>
      <th>Thứ 2</th><th>Thứ 3</th><th>Thứ 4</th><th>Thứ 5</th><th>Thứ 6</th>
    </tr>
    <tr><td><a href="#clb-bongda">Bóng đá</a></td><td>Thầy Khải</td><td></td><td>X</td><td></td><td></td><td>X</td></tr>
    <tr><td><a href="#clb-bongchuyen">Bóng chuyền</a></td><td>Cô Lan</td><td>X</td><td></td><td></td><td></td><td></td></tr>
    <tr><td><a href="#clb-bongro">Bóng rổ</a></td><td>Thầy Vũ</td><td></td><td></td><td>X</td><td></td><td></td></tr>
    <tr><td><a href="#clb-vothuat">Võ thuật</a></td><td>Thầy Hoàng - Karatedo<br>Thầy Đức - Taekwondo<br>Thầy Thắng - Vovinam</td><td>X</td><td>X</td><td></td><td>X</td><td></td></tr>
    <tr><td><a href="#clb-boi">Bơi</a></td><td>Cô Bích</td><td></td><td></td><td>X</td><td></td><td>X</td></tr>
  </table>

  <hr>
  <!-- Phần Vận dụng (SGK trang 61): Các đoạn giới thiệu chi tiết từng CLB ở dưới bảng -->
  <h2>Chi tiết các Câu lạc bộ Thể thao</h2>

  <h3 id="clb-bongda">1. Câu lạc bộ Bóng đá</h3>
  <p>Địa điểm: Sân bóng cỏ nhân tạo trường. Số thành viên: 45 học sinh. Thành tích: 01 giải Ba môn Bóng đá nam cấp thành phố. Phí hội viên: 50.000đ/tháng. Điện thoại liên hệ: Thầy Khải (0905.111.222).</p>

  <h3 id="clb-bongchuyen">2. Câu lạc bộ Bóng chuyền</h3>
  <p>Địa điểm: Sân thể chất ngoài trời số 1. Số thành viên: 32 học sinh. Thành tích: Giải Khuyến khích cấp cụm trường. Phí hội viên: 40.000đ/tháng. Điện thoại liên hệ: Cô Lan (0905.222.333).</p>

  <h3 id="clb-bongro">3. Câu lạc bộ Bóng rổ</h3>
  <p>Địa điểm: Nhà thi đấu đa năng. Số thành viên: 38 học sinh. Thành tích: Top 4 giải Bóng rổ học sinh thành phố. Phí hội viên: 50.000đ/tháng. Điện thoại liên hệ: Thầy Vũ (0905.333.444).</p>

  <h3 id="clb-vothuat">4. Câu lạc bộ Võ thuật (Karatedo – Taekwondo – Vovinam)</h3>
  <p>Địa điểm: Phòng tập Võ đường trường. Số thành viên: 60 học sinh. Thành tích: 02 HCV, 06 HCB và 05 HCĐ cấp thành phố các môn Võ thuật. Phí hội viên: 60.000đ/tháng. Điện thoại liên hệ: Thầy Hoàng (0905.444.555).</p>

  <h3 id="clb-boi">5. Câu lạc bộ Bơi lội</h3>
  <p>Địa điểm: Bể bơi tiêu chuẩn trường. Số thành viên: 40 học sinh. Thành tích: 03 HCV cấp thành phố môn Bơi lội. Phí hội viên: 70.000đ/tháng. Điện thoại liên hệ: Cô Bích (0905.555.666).</p>
</body>
</html>
"""
    (html_dir / 'CLB.html').write_text(clb_hoan_chinh, encoding='utf-8')

    # 3. Tệp thong_tin.html (Nhiệm vụ 2 + Bài Luyện tập 1 SGK trang 61)
    thong_tin_html = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Thông tin của các Câu lạc bộ</title>
</head>
<body>
  <h2>Thông tin chung về các Câu lạc bộ ngoại khoá</h2>
  <p>Các câu lạc bộ Thể thao và Nghệ thuật sinh hoạt định kỳ từ 16h30 đến 18h00 các ngày trong tuần tại khuôn viên nhà trường.</p>
  <ul>
    <li>Đối tượng tham gia: Toàn bộ học sinh Khối 10, Khối 11 và Khối 12.</li>
    <li>Quyền lợi: Được huấn luyện bài bản và đại diện nhà trường tham gia các giải đấu cấp quận, thành phố.</li>
    <li>Đăng ký trực tiếp tại Văn phòng Đoàn trường hoặc liên hệ giáo viên phụ trách từng bộ môn.</li>
  </ul>
  <!-- Bài Luyện tập 1 (SGK trang 61): Liên kết quay về trang chủ CLB.html -->
  <p><a href="CLB.html">Quay về trang chủ CLB</a></p>
</body>
</html>
"""
    (html_dir / 'thong_tin.html').write_text(thong_tin_html, encoding='utf-8')

    # 4. Tệp gioi_thieu_ban_than.html (Bài Luyện tập 2 SGK trang 61)
    gt_html = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Giới thiệu bản thân - Bài Luyện tập 2 trang 61</title>
</head>
<body>
  <h2>Giới thiệu bản thân</h2>
  <p>Họ và tên: Học sinh Lớp 12 – Trường THPT Thuận Hoá, thành phố Huế.</p>
  <p>Sở thích: Lập trình web HTML, thiết kế đồ hoạ và tham gia CLB Thể thao.</p>
  <h3>Liên kết tới trang mạng xã hội của ba bạn cùng lớp:</h3>
  <ul>
    <li><a href="https://www.facebook.com/ban.cung.lop.1">Trang Facebook của bạn thứ nhất</a></li>
    <li><a href="https://www.facebook.com/ban.cung.lop.2">Trang Facebook của bạn thứ hai</a></li>
    <li><a href="https://www.facebook.com/ban.cung.lop.3">Trang Facebook của bạn thứ ba</a></li>
  </ul>
</body>
</html>
"""
    (html_dir / 'gioi_thieu_ban_than.html').write_text(gt_html, encoding='utf-8')

    if (GOI_NOP_DIR / 'Hoc_lieu_HTML').exists():
        shutil.rmtree(GOI_NOP_DIR / 'Hoc_lieu_HTML')
    shutil.copytree(html_dir, GOI_NOP_DIR / 'Hoc_lieu_HTML')
    print('Built HTML learning materials in:', html_dir)


def make_zips():
    for old_name in [
        '01_De_cuong_du_gio_Tin12_Bai10_Tiet1_LT_TietPPCT22_Lop12-2.docx',
        '02_De_cuong_du_gio_Tin12_Bai10_Tiet2_TH_TietPPCT23_Lop12-1.docx',
        '03_De_cuong_du_gio_Tin12_Bai10_CaBai_2Tiet_TietPPCT22-23.docx'
    ]:
        for d in (KHOI12_DIR, GOI_NOP_DIR):
            p = d / old_name
            if p.exists():
                p.unlink()
    for dup_name in [
        '01_De_cuong_du_gio_Tin12_Bai10_Tiet1_LT_TietPPCT21_Lop12-2.docx',
        '02_De_cuong_du_gio_Tin12_Bai10_Tiet2_TH_TietPPCT22_Lop12-1.docx'
    ]:
        p = KHOI12_DIR / dup_name
        if p.exists():
            p.unlink()

    readme_p = KHOI12_DIR / 'README_NGUON_VA_GIAI_TRINH_KHOI12.md'
    if readme_p.exists():
        shutil.copy2(readme_p, GOI_NOP_DIR / readme_p.name)

    xls_p = KHOI12_DIR / 'khoi_12_tin_hoc_07112026.xls'
    if xls_p.exists():
        shutil.copy2(xls_p, GOI_NOP_DIR / xls_p.name)

    zip1_path = KHOI12_DIR / 'HoQuangHuy_24S1020030_TronBo_HoSo_Khoi12.zip'
    with zipfile.ZipFile(zip1_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(KHOI12_DIR.glob('*.docx')):
            if not f.name.startswith('~$') and not f.name.startswith('.~'):
                z.write(f, f.name)
        if xls_p.exists():
            z.write(xls_p, xls_p.name)
        if readme_p.exists():
            z.write(readme_p, readme_p.name)
        for hf in sorted((KHOI12_DIR / 'Hoc_lieu_HTML').glob('*.html')):
            z.write(hf, f'Hoc_lieu_HTML/{hf.name}')

    zip2_path = Path(r'c:\Users\PC\khbd_docx\08_goi_nop_du_gio\HoQuangHuy_24S1020030_HoSo_DuGio_Tin12_Bai10.zip')
    with zipfile.ZipFile(zip2_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in sorted(GOI_NOP_DIR.rglob('*')):
            if f.is_file() and not f.name.startswith('~$') and not f.name.startswith('.~'):
                z.write(f, f.relative_to(GOI_NOP_DIR).as_posix())
    print('Packaged ZIPs:', zip1_path, 'and', zip2_path)


if __name__ == '__main__':
    build_bai10_tiet1()
    build_bai10_tiet2()
    build_bai10_cabai()
    build_bai10_worksheets()
    build_html_learning_materials()
    make_zips()

