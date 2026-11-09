# -*- coding: utf-8 -*-
"""
Builds all 3 deliverables for Grade 10 Bài 12: Phần mềm thiết kế đồ hoạ (Mục 4.12 - Tiết PPCT: 17, 18):
1. 03_De_cuong_du_gio_Tin10_Bai12_Tiet1_LT_TietPPCT17_Lop10-1.docx (Dự giờ Thứ Ba 10/11/2026 - Tiết 1, 7h10-7h55, Lớp 10/1, Phòng H.I.3)
2. 04_De_cuong_du_gio_Tin10_Bai12_Tiet2_TH_TietPPCT18_Lop10-2.docx (Dự giờ Thứ Ba 10/11/2026 - Tiết 2, 8h00-8h45, Lớp 10/2, Phòng H.I.4)
3. 05_De_cuong_du_gio_Tin10_Bai12_CaBai_2Tiet_TietPPCT17-18.docx (Bản gộp trọn vẹn cả Bài 12 - 02 tiết = 90 phút, Tiết PPCT: 17, 18)
"""
from pathlib import Path
import shutil
from docx_engine_5512 import (
    init_body_with_cover, add, br, table, add_figure, add_approval_section, save_docx
)

KHOI10_DIR = Path(r'c:\Users\PC\khbd_docx\khoi10')
THANH_PHAM_DIR = Path(r'c:\Users\PC\khbd_docx\06_thanh_pham_docx')
GOI_NOP_DIR = Path(r'c:\Users\PC\khbd_docx\08_goi_nop_du_gio\HoQuangHuy_24S1020030_HoSo_DuGio_Tin10')

# Shared step tables for Bài 12 so all 3 files have 100% rigorous, polite, evidence-backed content
HD1_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu lên màn hình hai hình bông hoa cúc ở phần Khởi động Hình 12.1a và Hình 12.1b (SGK Tin học 10 trang 63) cùng trạng thái phóng to một phần của hai hình đó ở Hình 12.2 (SGK trang 64).\n'
        '- Giáo viên nêu câu hỏi yêu cầu học sinh quan sát và trao đổi theo cặp trong 02 phút:\n'
        '  + Câu hỏi 1: Em hãy nhận xét về sự khác nhau giữa màu sắc, độ chân thực và độ sắc nét của Hình 12.1a (ảnh chụp) và Hình 12.1b (hình vẽ bằng phần mềm).\n'
        '  + Câu hỏi 2: Khi phóng to nhiều lần (Hình 12.2), hình nào bị nhoè và lộ các ô vuông răng cưa, hình nào vẫn giữ nguyên đường viền sắc nét?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh quan sát kỹ Hình 12.1a, 12.1b và Hình 12.2 trong Sách giáo khoa trang 63–64, thảo luận với bạn cùng bàn để tìm điểm khác biệt.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh đại diện trả lời tại chỗ; mời 01 học sinh khác nhận xét, bổ sung.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá theo bảng kiểm CCĐG 1, biểu dương câu trả lời đúng và dẫn dắt vào bài mới: *"Trong thực tế, để tạo ra các ấn phẩm truyền thông như biểu trưng (logo), áp phích, bảng hiệu có thể in ấn ở mọi kích thước mà không bị vỡ nét như Hình 12.1b, người ta sử dụng phần mềm thiết kế đồ hoạ vectơ. Bài 12 sẽ giúp các em phân biệt hai loại đồ hoạ và làm chủ phần mềm Inkscape."*',
        '**Nhận xét mở đầu (SGK Bài 12, trang 63–64):**\n'
        '- **Hình 12.1a (Ảnh chụp):** Có màu sắc chuyển tiếp tự nhiên, độ chi tiết rất cao và chân thực; tuy nhiên khi phóng to (Hình 12.2 trái) thì hình bị nhoè, vỡ nét thành các ô vuông nhỏ (điểm ảnh – pixel).\n'
        '- **Hình 12.1b (Hình vẽ bằng phần mềm):** Được tạo từ các hình khối, đường nét và mảng màu ít hơn nhưng khi phóng to hay thu nhỏ ở mức bất kì (Hình 12.2 phải) thì đường nét vẫn sắc nét, không bị vỡ hình.'
    ]
]

NV1_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên phát Phiếu học tập số 1 và yêu cầu học sinh làm việc theo nhóm đôi (05 phút), nghiên cứu mục 1. Thiết kế đồ hoạ và Bảng so sánh giữa đồ hoạ điểm ảnh và đồ hoạ vectơ (SGK trang 63–64) để thực hiện hai yêu cầu:\n'
        '  + Yêu cầu 1: Nêu khái niệm thiết kế đồ hoạ và hoàn thành bảng phân biệt Đồ hoạ điểm ảnh (bitmap) với Đồ hoạ vectơ (vector) theo 5 tiêu chí trong SGK trang 64.\n'
        '  + Yêu cầu 2 (Câu hỏi củng cố SGK trang 64): Trả lời 2 câu hỏi: (1) Ảnh chụp là loại đồ hoạ nào? (2) Tại sao dùng đồ hoạ vectơ phù hợp hơn dùng đồ hoạ điểm ảnh khi thiết kế logo?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đọc hiểu Sách giáo khoa trang 63–64, thảo luận nhóm đôi và điền đầy đủ từ khoá vào Phiếu học tập số 1. Giáo viên quan sát, gợi ý học sinh liên hệ với kiến thức dung lượng tệp đã học ở Bài 1 (Mục 1.1).\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời đại diện 02 nhóm trình bày bảng so sánh và giải thích 02 câu hỏi củng cố trang 64; các nhóm khác nhận xét, bổ sung.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo CCĐG 2, chuẩn hoá bảng kiến thức và nhấn mạnh hộp ghi nhớ trang 64.',
        '**1. Thiết kế đồ hoạ (SGK trang 63–64)**\n'
        '- **Khái niệm:** Thiết kế đồ hoạ (graphic design) là quá trình thiết kế các thông điệp truyền thông bằng hình ảnh; giải quyết vấn đề thông qua sự kết hợp giữa hình ảnh, kiểu chữ với ý tưởng để truyền tải thông tin đến người xem.\n'
        '- **Phân biệt hai loại đồ hoạ cơ bản (Bảng trang 64):**\n'
        '  + **Đồ hoạ điểm ảnh (Bitmap):** (1) Định nghĩa bằng tập hợp các điểm ảnh (pixel), mỗi điểm có màu riêng; (2) Phù hợp chỉnh sửa ảnh chụp, nhiều chi tiết, màu sắc liền mạch, chân thực; (3) Phóng to làm ảnh hưởng chất lượng (nhoè, vỡ hình); (4) Ảnh lớn, độ chi tiết cao thì tệp có kích thước lớn (.jpg, .png, .bmp); (5) Không thể chuyển thẳng sang đồ hoạ vectơ mà giữ nguyên chất lượng.\n'
        '  + **Đồ hoạ vectơ (Vector):** (1) Định nghĩa bằng các phương trình toán học (đường nét có điểm đầu, điểm cuối); (2) Phù hợp tạo logo, minh hoạ, bản vẽ kĩ thuật; (3) Có thể co dãn tuỳ ý mà không bị vỡ hình; (4) Tạo bản in kích thước tuỳ ý nhưng độ lớn của tệp không thay đổi (.svg); (5) Dễ dàng xuất/chuyển sang đồ hoạ điểm ảnh.\n'
        '- **Trả lời Câu hỏi SGK trang 64:**\n'
        '  + Câu 1: Ảnh chụp bằng máy ảnh, điện thoại là đồ hoạ điểm ảnh (bitmap).\n'
        '  + Câu 2: Thiết kế logo nên dùng đồ hoạ vectơ vì logo thường gồm các hình khối/đường nét cô đọng và cần phóng to/thu nhỏ trên nhiều ấn phẩm (từ danh thiếp nhỏ đến biển quảng cáo lớn) mà vẫn giữ nguyên độ sắc nét, tệp nhẹ và dễ chỉnh sửa từng thành phần.'
    ]
]

NV2_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu tình huống Hoạt động 2 (SGK trang 64): *"Minh muốn thiết kế logo cho câu lạc bộ bóng đá của trường nhưng bạn ấy không biết bắt đầu từ đâu"* và giao Phiếu học tập số 2 (06 phút) yêu cầu học sinh đọc mục 2. Phần mềm đồ hoạ (SGK trang 64–66):\n'
        '  + Yêu cầu 1: Phân loại các phần mềm đồ hoạ thông dụng (Adobe Illustrator, CorelDRAW, Inkscape, Adobe Photoshop, GIMP) thành 2 nhóm (vectơ và bitmap); giải thích vì sao chọn Inkscape trong nhà trường phổ thông (liên hệ Bài 11 – bản quyền phần mềm).\n'
        '  + Yêu cầu 2: Quan sát Hình 12.3 (SGK trang 65), xác định vị trí và chức năng của các thành phần chính trên màn hình làm việc của Inkscape.\n'
        '  + Yêu cầu 3: Trả lời 02 câu hỏi trắc nghiệm củng cố ở đầu trang 66 SGK.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đối chiếu Hình 12.3 (SGK trang 65) với giao diện phần mềm Inkscape đang mở trên màn chiếu/máy tính, thảo luận và điền vào Phiếu học tập số 2.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 01 học sinh lên chỉ trực tiếp trên màn hình 8 vùng giao diện của Hình 12.3 và mời 01 học sinh trả lời 2 câu hỏi củng cố trang 66.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo CCĐG 3, chốt kiến thức mục 2 và nhấn mạnh sự khác biệt giữa Vùng làm việc (Canvas) và Trang in (Page).',
        '**2. Phần mềm đồ hoạ và giao diện Inkscape (SGK trang 64–66)**\n'
        '- **Hai nhóm phần mềm đồ hoạ:**\n'
        '  + Phần mềm tạo, chỉnh sửa hình vectơ: Adobe Illustrator, CorelDRAW, Inkscape.\n'
        '  + Phần mềm xử lí ảnh bitmap: Adobe Photoshop, GIMP.\n'
        '  + **Inkscape** là phần mềm nguồn mở miễn phí (tải tại https://inkscape.org), giúp học sinh thiết kế đồ hoạ vectơ chuyên nghiệp và tuân thủ nghĩa vụ tôn trọng bản quyền (đã học ở Bài 11).\n'
        '- **Các thành phần trên màn hình làm việc của Inkscape (Hình 12.3, trang 65):**\n'
        '  1. **Thanh bảng chọn (Menu bar):** Chứa các lệnh liên quan đến tệp tin (File), chỉnh sửa (Edit), đối tượng (Object), đường (Path)...\n'
        '  2. **Hộp công cụ (Tool box – bên trái):** Chứa các công cụ chính để chọn, khởi tạo, vẽ và điều chỉnh đối tượng đồ hoạ.\n'
        '  3. **Thanh điều khiển thuộc tính (Tool Controls Bar):** Chứa các thuộc tính của công cụ/đối tượng đang được chọn (thay đổi linh hoạt theo từng công cụ).\n'
        '  4. **Vùng làm việc (Canvas) & Trang in (Page):** Vùng làm việc là toàn bộ nền trắng; khung chữ nhật ở giữa là Trang in ứng với kích thước khi xuất/in sản phẩm.\n'
        '  5. **Bảng màu (Color Palette – phía dưới):** Chứa các ô màu để thiết lập nhanh màu tô (Fill) và màu vẽ đường viền (Stroke).\n'
        '  6. **Thanh trạng thái (Status bar):** Hiển thị thông tin lớp vẽ, màu đối tượng, toạ độ con trỏ, tỉ lệ thu phóng.\n'
        '- **Trả lời Câu hỏi SGK trang 66:**\n'
        '  + Câu 1: Thiết kế bộ sản phẩm bút, sổ, danh thiếp, bì thư, túi giấy nên dùng **B. Inkscape**.\n'
        '  + Câu 2: Có thể vẽ hình vào **A. Toàn bộ vùng làm việc** (nhưng khi in chỉ phần nằm trong Trang in được in ra).'
    ]
]

NV3_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên nêu bài toán Hoạt động 3: Sắp xếp các đối tượng (SGK trang 66) về quả dưa hấu và quả cam đặt trên bàn, đồng thời yêu cầu học sinh đọc mục 3. Các đối tượng đồ hoạ của hình vẽ (SGK trang 66–67) để trả lời các câu hỏi trong Phiếu học tập số 3:\n'
        '  + Câu 1: Xác định cách xếp thứ tự trước/sau của quả dưa hấu và quả cam cho 4 trường hợp a, b, c, d ở Hoạt động 3 (trang 66); từ đó rút ra quy tắc thứ tự lớp (Layer order) của các đối tượng trong Inkscape.\n'
        '  + Câu 2: Nêu quy trình 3 bước vẽ một đối tượng cơ bản từ Hộp công cụ (Hình 12.4 trang 66), cách dùng công cụ Select (phím F1 hoặc S) để phóng to/thu nhỏ (nháy lần 1) và xoay đối tượng (nháy lần 2).\n'
        '  + Câu 3: Phân biệt thao tác chọn Màu tô (Fill Color) và Màu vẽ đường viền (Stroke Color) trên Bảng màu và cách lưu tệp chuẩn .svg (trả lời kèm câu hỏi trắc nghiệm đầu trang 67).\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thảo luận nhóm đôi, đọc kỹ SGK trang 66–67 và hoàn thành Phiếu học tập số 3.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Đại diện học sinh trả lời bài toán quả dưa hấu – quả cam và thao tác với công cụ Select, bảng màu; giáo viên thao tác thị phạm trực tiếp trên Inkscape.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo CCĐG 4 và chốt quy tắc vàng: Nháy chuột vào ô màu = Đổi màu tô (Fill); Giữ phím Shift + Nháy chuột vào ô màu = Đổi màu viền (Stroke).',
        '**3. Các đối tượng đồ hoạ của hình vẽ (SGK trang 66–67)**\n'
        '- **Giải quyết Hoạt động 3 (SGK trang 66 – Quả dưa hấu và quả cam):**\n'
        '  + a) Nhìn thấy đủ cả hai quả: Đặt hai quả tách rời nhau, không chồng lấp.\n'
        '  + b) Chỉ nhìn thấy quả dưa hấu: Đặt quả dưa hấu (lớn hơn) ở lớp trên che khuất hoàn toàn quả cam ở lớp dưới.\n'
        '  + c) Nhìn thấy đủ quả cam, dưa hấu bị che một phần: Đặt quả cam ở lớp trên, chồng một phần lên quả dưa hấu ở lớp dưới.\n'
        '  + d) Nhìn thấy đủ quả dưa hấu, cam bị che một phần: Đặt quả dưa hấu ở lớp trên, chồng một phần lên quả cam ở lớp dưới.\n'
        '- **Quy tắc thứ tự lớp trong Inkscape:** Đối tượng vẽ trước nằm ở lớp dưới, đối tượng vẽ sau nằm ở lớp trên (có thể che khuất một phần hoặc toàn bộ đối tượng lớp dưới). Có thể thay đổi thứ tự lớp bằng các nút lệnh trên Thanh điều khiển thuộc tính (Lower selection to bottom, Raise selection to top).\n'
        '- **Ba bước thêm đối tượng có sẵn (Hình 12.4, trang 66):**\n'
        '  + Bước 1: Chọn công cụ tương ứng trên Hộp công cụ (Hình vuông/chữ nhật F4/R, Hình tròn/elip F5/E, Hình đa giác/ngôi sao phím dấu hoa thị, Bút vẽ, Văn bản F8/T).\n'
        '  + Bước 2: Chỉnh tuỳ chọn trên Thanh điều khiển thuộc tính (ví dụ: Corners, Spoke ratio).\n'
        '  + Bước 3: Kéo thả chuột trên Vùng làm việc để vẽ hình.\n'
        '- **Thao tác chọn, biến đổi, tô màu và lưu tệp (trang 66–67):**\n'
        '  + Dùng công cụ Select (mũi tên chọn): Nháy chuột lần 1 vào đối tượng để di chuyển hoặc kéo mũi tên 2 chiều ở góc để phóng to/thu nhỏ; Nháy chuột lần 2 vào đối tượng để xuất hiện mũi tên cong ở góc phục vụ xoay hình.\n'
        '  + **Màu tô (Fill Color):** Chọn đối tượng -> Nháy chuột trái vào ô màu trên Bảng màu.\n'
        '  + **Màu vẽ viền (Stroke Color):** Chọn đối tượng -> Nhấn giữ phím Shift + Nháy chuột trái vào ô màu trên Bảng màu.\n'
        '  + **Lưu tệp:** Chọn File / Save (hoặc nhấn Ctrl + S), phần mở rộng mặc định của tệp đồ hoạ vectơ Inkscape là .svg.\n'
        '- **Trả lời câu hỏi SGK trang 67:** Thanh công cụ được sử dụng nhiều nhất trong Inkscape để khởi tạo và chỉnh sửa đối tượng là **D. Hộp công cụ**.'
    ]
]

NV4_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh quan sát mục Thực hành (SGK trang 67–68) gồm Hình 12.5 (Vẽ bông hoa) và Hình 12.6 (Vẽ Quốc kì Việt Nam), sau đó thực hành tại máy tính cá nhân/cặp đôi để hoàn thành 02 sản phẩm trong tệp HoTen_Lop_Bai12.svg:\n'
        '  + Sản phẩm 1 – Nhiệm vụ 1 (SGK trang 67 – Hình 12.5): Vẽ một bông hoa 10–12 cánh từ 01 hình tròn (nhụy hoa) và các hình elip (cánh hoa) bằng lệnh nhân bản Duplicate (Ctrl + D), thao tác nháy đúp công cụ Select để xoay cánh hoa và ghép quanh nhụy.\n'
        '  + Sản phẩm 2 – Nhiệm vụ 2 (SGK trang 67–68 – Hình 12.6): Vẽ lá cờ Tổ quốc Việt Nam bằng công cụ Hình sao (Corners = 5, Spoke ratio = 0.4, màu vàng), công cụ Hình chữ nhật (tỉ lệ 2 : 3, nền màu đỏ) và dùng nút lệnh Đưa xuống lớp dưới (Lower selection to bottom) để ngôi sao vàng nổi lên giữa nền cờ đỏ.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thực hành trực tiếp trên phần mềm Inkscape theo đúng quy trình 3 bước của Hình 12.5 và Hình 12.6. Giáo viên đi từng dãy máy quan sát, hỗ trợ học sinh xử lí tình huống ngôi sao bị che khuất bởi hình chữ nhật vẽ sau.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên chiếu 02 bài thực hành hoàn chỉnh của học sinh lên màn hình chung; học sinh giải thích cách dùng Duplicate (Ctrl + D) và nút hạ lớp (Lower selection to bottom).\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá sản phẩm thực hành theo Rubric CCĐG 6 (Mục IV. Phụ lục 2), khen ngợi các nhóm căn chỉnh cánh hoa đều và đúng tỉ lệ Quốc kì.',
        '**4. Thực hành vẽ hình đơn giản trong Inkscape (SGK trang 67–68)**\n'
        '**a) Nhiệm vụ 1: Vẽ bông hoa từ hình tròn và các hình elip (Hình 12.5, trang 67):**\n'
        '- **Bước 1:** Chọn công cụ Hình tròn/elip trên Hộp công cụ, vẽ 01 hình tròn làm đài/nhụy hoa và 01 hình elip đứng làm cánh hoa; chọn màu vàng trên Bảng màu (Hình 12.5a).\n'
        '- **Bước 2:** Chọn hình elip -> nháy nút phải chuột chọn Duplicate (hoặc nhấn Ctrl + D) nhiều lần để tạo các bản sao cánh hoa đồng nhất về kích thước (Hình 12.5b).\n'
        '- **Bước 3:** Chọn công cụ Select: nháy chuột lần 1 lên một cánh hoa (chế độ phóng to/thu nhỏ – Hình 12.5c), nháy chuột lần 2 để hiển thị mũi tên xoay ở 4 góc (Hình 12.5d); kéo mũi tên góc để xoay cánh hoa theo hướng phù hợp (Hình 12.5e) rồi di chuyển xếp đều quanh nhụy hoa tạo thành bông hoa hoàn chỉnh (Hình 12.5g).\n'
        '**b) Nhiệm vụ 2: Vẽ Quốc kì Việt Nam (Hình 12.6, trang 67–68):**\n'
        '- **Bước 1 (Vẽ ngôi sao 5 cánh):** Chọn công cụ Hình đa giác/ngôi sao -> trên Thanh điều khiển thuộc tính chọn biểu tượng Hình sao, đặt Corners = 5 và Spoke ratio = 0.4 -> kéo thả vẽ ngôi sao và tô màu vàng (Hình 12.6a).\n'
        '- **Bước 2 (Vẽ nền cờ đỏ):** Chọn công cụ Hình vuông/chữ nhật, vẽ hình chữ nhật có tỉ lệ chiều rộng và chiều dài là 2 : 3, tô nền màu đỏ (Hình 12.6b). (Lưu ý: Do hình chữ nhật vẽ sau nên nằm ở lớp trên, khi đưa ngôi sao vào giữa sẽ bị hình chữ nhật che khuất).\n'
        '- **Bước 3 (Đổi thứ tự lớp):** Chọn công cụ Select -> nháy chọn hình chữ nhật đỏ -> nháy nút Lower selection to bottom (hoặc nhấn phím End) trên Thanh điều khiển thuộc tính để đưa hình chữ nhật xuống lớp dưới, giúp ngôi sao vàng hiện rõ chính giữa nền đỏ (Hình 12.6c).\n'
        '- **Bước 4:** Lưu tệp với tên HoTen_Lop_Bai12.svg (Ctrl + S).'
    ]
]

HD3_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh hoàn thành Phiếu học tập số 4 (giải quyết trọn vẹn 02 bài tập phần Luyện tập (SGK trang 68) kết hợp kiểm chứng trực tiếp trên phần mềm Inkscape):\n'
        '  + Bài Luyện tập 1 (SGK trang 68): Sau khi vẽ một hình tròn trong Inkscape, ta thực hiện liên tục các thao tác: Nháy công cụ Select -> nháy chọn hình tròn -> nháy chuột chọn lần lượt các ô màu Đỏ, Nâu, Cam trên bảng màu -> giữ phím Shift và nháy chuột vào ô màu Đỏ đậm trên bảng màu. Em hãy xác định kết quả cuối cùng hình tròn sẽ có màu tô nền và màu viền như thế nào?\n'
        '  + Bài Luyện tập 2 (SGK trang 68): Để thay một ngôi sao thành một khối lập phương (hoặc vẽ hình mới), em sẽ tìm công cụ ở thanh công cụ nào? (A. Bảng màu / B. Thanh điều khiển thuộc tính / C. Hộp công cụ / D. Hộp thoại lệnh).\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh suy nghĩ, ghi câu trả lời vào Phiếu học tập số 4 và thực hiện chuỗi thao tác của Bài Luyện tập 1 ngay trên phần mềm Inkscape để kiểm chứng kết quả.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình bày kết quả và giải thích cơ chế ghi đè màu tô (Fill) khi nháy chuột thường so với thao tác giữ phím Shift để đổi màu vẽ (Stroke).\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét, chấm điểm theo CCĐG 5 và chốt lại quy tắc thao tác chuẩn trong Inkscape.',
        '**Đáp án phần Luyện tập (SGK Bài 12, trang 68):**\n'
        '- **Giải Bài Luyện tập 1 (SGK trang 68):**\n'
        '  + Khi nháy chuột thường lần lượt vào các ô màu Đỏ -> Nâu -> Cam trên Bảng màu, màu tô bên trong (Fill Color) của hình tròn sẽ thay đổi liên tiếp và dừng lại ở màu cuối cùng được nháy là màu Cam.\n'
        '  + Khi nhấn giữ phím Shift và nháy chuột vào ô màu Đỏ đậm, thao tác này không làm thay đổi màu tô nền mà thiết lập màu vẽ đường viền (Stroke Color) của hình tròn thành màu Đỏ đậm.\n'
        '  + **Kết luận:** Hình tròn thu được có phần nền bên trong (Fill) màu Cam và đường viền bao quanh (Stroke) màu Đỏ đậm.\n'
        '- **Giải Bài Luyện tập 2 (SGK trang 68):**\n'
        '  + Chọn đáp án **C. Hộp công cụ** (Vì công cụ vẽ hình hộp 3D / khối lập phương nằm trong Hộp công cụ ở cạnh trái màn hình làm việc của Inkscape).'
    ]
]

HD4_ROWS_BAI12 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu hình minh hoạ phần Vận dụng (SGK trang 68) và yêu cầu học sinh thực hiện:\n'
        '  + Bài Vận dụng 1 (SGK trang 68): Phân tích cấu tạo các lớp hình học và chọn vẽ ít nhất 01 Quốc kì trong số 4 nước: Thái Lan, Lào, Myanmar (hoặc Anh) ngay trên tệp thực hành.\n'
        '  + Bài Vận dụng 2 (SGK trang 68): Phân tích các hình cơ bản (hình chữ nhật, hình tam giác/đa giác 3 đỉnh, hình tròn) tạo nên hình chú chó và chú tuần lộc để hoàn thiện ở nhà.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh phân tích thứ tự xếp lớp và thực hành vẽ nhanh Quốc kì Lào hoặc Myanmar/Thái Lan trên máy tính.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình bày quy trình xếp lớp khi vẽ Quốc kì Lào (3 dải chữ nhật Đỏ – Xanh – Đỏ và 1 hình tròn Trắng ở lớp trên cùng) và Quốc kì Thái Lan (3 hình chữ nhật Đỏ – Trắng – Xanh lam xếp chồng đồng tâm từ lớn đến nhỏ).\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo Rubric CCĐG 6, tổng kết toàn bài và hướng dẫn học sinh chuẩn bị Bài 13: Bổ sung các đối tượng đồ hoạ (Tiết PPCT: 19, 20).',
        '**Hướng dẫn giải phần Vận dụng (SGK Bài 12, trang 68):**\n'
        '**1. Bài Vận dụng 1 – Vẽ Quốc kì các nước (SGK trang 68):**\n'
        '- **Quốc kì Thái Lan:** Vẽ 01 hình chữ nhật lớn màu Đỏ ở lớp dưới cùng -> vẽ 01 hình chữ nhật màu Trắng thấp hơn đặt chính giữa ở lớp trung gian -> vẽ 01 hình chữ nhật màu Xanh đậm đặt chính giữa ở lớp trên cùng.\n'
        '- **Quốc kì Lào:** Vẽ hình chữ nhật lớn màu Đỏ (lớp dưới) -> vẽ hình chữ nhật màu Xanh lam cao bằng 1/2 đặt chính giữa (lớp giữa) -> dùng công cụ Hình tròn vẽ 01 hình tròn màu Trắng đặt ở tâm (lớp trên cùng).\n'
        '- **Quốc kì Myanmar:** Vẽ 3 dải hình chữ nhật ngang màu Vàng (trên), Xanh lá (giữa), Đỏ (dưới) -> dùng công cụ Hình sao (Corners = 5, Spoke ratio = 0.4) vẽ ngôi sao 5 cánh màu Trắng đặt ở lớp trên cùng chính giữa lá cờ.\n'
        '**2. Bài Vận dụng 2 – Ghép các hình cơ bản thành con vật (SGK trang 68):**\n'
        '- **Hình chú chó & chú tuần lộc:** Kết hợp công cụ Hình chữ nhật (thân, chân, cổ, sừng), Hình đa giác Corners = 3 (tai hình tam giác) và Hình tròn (mắt, chóp mũi), sử dụng lệnh Duplicate (Ctrl + D) để nhân bản các chi tiết đối xứng.'
    ]
]

def add_bai12_appendix(b, class_label='10/1 & 10/2', ppct_label='17, 18'):
    add(b, 'IV. PHỤ LỤC (HỒ SƠ HỌC LIỆU, PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ)', 'heading')
    add(b, 'Phụ lục 1. Hệ thống Phiếu học tập số 1, 2, 3, 4 và Công cụ đánh giá CCĐG 1 – CCĐG 5', 'sub')
    pht_rows = [
        ['Tên Phiếu / Công cụ', 'Nội dung câu hỏi và yêu cầu (Dựa sát SGK Tin học 10 KNTT, Bài 12 trang 63–68)', 'Đáp án chuẩn & Thang điểm chấm'],
        ['**CCĐG 1**\n(Hoạt động 1:\nMở đầu)',
         '- Quan sát Hình 12.1a, 12.1b (tr. 63) và Hình 12.2 (tr. 64): So sánh độ chi tiết màu sắc và hiện tượng xảy ra khi phóng to hai bức hình.',
         '- Chỉ rõ Hình 12.1a là ảnh chụp bị vỡ hạt (pixel) khi phóng to; Hình 12.1b là hình vẽ giữ nguyên độ sắc nét (Đạt / Chưa đạt).'],
        ['**Phiếu học tập số 1**\n**và CCĐG 2**\n(Hoạt động 2 –\nNhiệm vụ 1)',
         '1. Điền bảng so sánh Đồ hoạ điểm ảnh (Bitmap) và Đồ hoạ vectơ (Vector) theo 5 tiêu chí (SGK tr. 64): Định nghĩa; Ứng dụng phù hợp; Khi phóng to; Kích thước tệp; Khả năng chuyển đổi.\n'
         '2. Trả lời 2 câu hỏi SGK tr. 64: (a) Ảnh chụp là loại đồ hoạ nào? (b) Vì sao thiết kế logo dùng đồ hoạ vectơ?',
         '- Điền đúng 5 tiêu chí so sánh theo Bảng trang 64 *(6,0 điểm)*.\n'
         '- Trả lời đúng: (a) Đồ hoạ điểm ảnh; (b) Co dãn không vỡ hình, tệp nhẹ, dễ in ấn mọi kích cỡ *(4,0 điểm)*.'],
        ['**Phiếu học tập số 2**\n**và CCĐG 3**\n(Hoạt động 2 –\nNhiệm vụ 2)',
         '1. Phân loại phần mềm: Illustrator, CorelDRAW, Inkscape, Photoshop, GIMP vào 2 cột (Vectơ / Bitmap).\n'
         '2. Ghi tên 6 thành phần giao diện chính trên Hình 12.3 (SGK tr. 65).\n'
         '3. Khoanh đáp án 2 câu hỏi trắc nghiệm đầu trang 66 SGK.',
         '- Phân loại đúng 5 phần mềm *(3,0 điểm)*.\n'
         '- Nêu đúng chức năng 6 vùng giao diện Inkscape *(4,0 điểm)*.\n'
         '- Câu 1 tr. 66: **B. Inkscape**; Câu 2 tr. 66: **A. Toàn bộ vùng làm việc** *(3,0 điểm)*.'],
        ['**Phiếu học tập số 3**\n**và CCĐG 4**\n(Hoạt động 2 –\nNhiệm vụ 3)',
         '1. Giải bài toán sắp xếp quả dưa hấu và quả cam cho 4 ý a, b, c, d (Hoạt động 3 tr. 66).\n'
         '2. Nêu sự khác nhau khi nháy chuột lần 1 và nháy chuột lần 2 bằng công cụ Select lên đối tượng.\n'
         '3. Nêu thao tác chọn màu tô (Fill), màu vẽ (Stroke) và phần mở rộng tệp Inkscape.',
         '- Giải đúng 4 trường hợp xếp lớp a, b, c, d *(4,0 điểm)*.\n'
         '- Nháy lần 1: Di chuyển/co dãn; Nháy lần 2: Xoay đối tượng *(3,0 điểm)*.\n'
         '- Nháy thường = Fill; Shift + Nháy = Stroke; tệp .svg; Câu hỏi tr. 67 chọn **D. Hộp công cụ** *(3,0 điểm)*.'],
        ['**Phiếu học tập số 4**\n**và CCĐG 5**\n(Hoạt động 3:\nLuyện tập)',
         '1. Giải Bài Luyện tập 1 (SGK tr. 68): Xác định màu tô và màu viền của hình tròn sau chuỗi thao tác.\n'
         '2. Giải Bài Luyện tập 2 (SGK tr. 68): Chọn thanh công cụ chứa khối lập phương.',
         '- Bài 1: Màu tô (Fill) là màu Cam; màu viền (Stroke) là màu Đỏ đậm *(5,0 điểm)*.\n'
         '- Bài 2: Chọn **C. Hộp công cụ** *(5,0 điểm)*.']
    ]
    b.append(table(pht_rows, [1600, 4200, 3200], header=True, cell_sz=21))

    add(b, 'Phụ lục 2. Công cụ đánh giá CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành Inkscape (Hình 12.5, Hình 12.6 & Vận dụng)', 'sub')
    rubric_info = [
        ['Thông tin hành chính phiếu đánh giá sản phẩm thực hành (CCĐG 6)', 'Nội dung chi tiết'],
        ['Tên bài dạy & Tiết PPCT', f'Bài 12. Phần mềm thiết kế đồ hoạ (Tiết PPCT: {ppct_label})'],
        ['Lớp giảng dạy & Địa điểm', f'Lớp {class_label} – Trường THPT Thuận Hoá (Phòng H.I.3 / H.I.4)'],
        ['Giáo viên giảng dạy và Giáo viên hướng dẫn', 'Thầy Mai Xuân Văn'],
        ['Sinh viên thực hiện', 'Hồ Quang Huy'],
        ['Họ và tên học sinh / Nhóm máy', '........................................................................ Máy số: ...................']
    ]
    b.append(table(rubric_info, [3800, 5200], header=True, cell_sz=22))
    add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Công cụ & Thao tác Inkscape bắt buộc kiểm tra', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Điểm tối đa', 'Điểm đạt'],
        ['1', 'Nhiệm vụ 1:\nVẽ bông hoa\n(Hình 12.5 SGK tr. 67)', 'Công cụ Hình tròn/elip (E), lệnh Duplicate (Ctrl + D), nháy lần 2 công cụ Select để xoay cánh hoa.', '- **Mức 3 (3,5 đ):** Vẽ đúng nhụy tròn, dùng Duplicate tạo các cánh hoa elip bằng nhau, xoay và xếp đều quanh nhụy như Hình 12.5g.\n- **Mức 2 (2,5 đ):** Vẽ được nhụy và các cánh hoa nhưng góc xoay chưa đều.\n- **Mức 1 (1,0 đ):** Các cánh hoa vẽ rời rạc bằng tay, không dùng Duplicate và chưa biết xoay.', '3,5 điểm', '......'],
        ['2', 'Nhiệm vụ 2:\nVẽ Quốc kì Việt Nam\n(Hình 12.6 SGK tr. 68)', 'Công cụ Hình sao (Corners = 5, Spoke ratio = 0.4), Hình chữ nhật tỉ lệ 2 : 3, nút Lower selection to bottom.', '- **Mức 3 (3,5 đ):** Ngôi sao 5 cánh chuẩn (Corners = 5, Spoke ratio = 0.4) màu vàng nổi ở lớp trên chính giữa nền cờ chữ nhật đỏ tỉ lệ 2 : 3.\n- **Mức 2 (2,5 đ):** Vẽ được nền đỏ và sao vàng đúng lớp nhưng thông số Spoke ratio hoặc vị trí chưa cân.\n- **Mức 1 (1,0 đ):** Ngôi sao bị hình chữ nhật đỏ che khuất do chưa biết đổi thứ tự lớp.', '3,5 điểm', '......'],
        ['3', 'Luyện tập & Vận dụng:\nPhối màu Fill/Stroke & Quốc kì các nước (tr. 68)', 'Nháy chuột chọn Fill, giữ Shift + Nháy chuột chọn Stroke; xếp lớp vẽ Quốc kì Lào/Thái Lan/Myanmar.', '- **Mức 3 (2,0 đ):** Thực hiện đúng thao tác đổi màu nền/màu viền Bài Luyện tập 1 và vẽ thêm được ít nhất 01 Quốc kì ở phần Vận dụng 1.\n- **Mức 2 (1,25 đ):** Làm đúng Bài Luyện tập 1 nhưng chưa hoàn thiện hình Vận dụng.\n- **Mức 1 (0,5 đ):** Còn nhầm lẫn giữa thao tác chọn màu nền (Fill) và màu viền (Stroke).', '2,0 điểm', '......'],
        ['4', 'Lưu tệp & Ý thức thực hành phòng máy', 'Lệnh File / Save (Ctrl + S) đúng định dạng .svg.', '- **Mức 3 (1,0 đ):** Lưu tệp đúng tên HoTen_Lop_Bai12.svg, hoàn thành đúng giờ, giữ gìn thiết bị phòng máy.\n- **Mức 2 (0,5 đ):** Lưu tệp đúng định dạng .svg nhưng đặt tên chưa chuẩn.\n- **Mức 1 (0,25 đ):** Quên lưu tệp thực hành.', '1,0 điểm', '......'],
        ['', '**Tổng điểm**', '', '**Xếp loại:** Đạt từ 8,0 – 10,0: Tốt | 5,0 – 7,9: Đạt | < 5,0: Cần cố gắng', '**10,0 điểm**', '......']
    ]
    b.append(table(rubric_rows, [650, 1500, 2000, 3450, 700, 700], header=True, cell_sz=21))

def build_bai12_tiet1():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (TIẾT 1 - LÍ THUYẾT – TIẾT PPCT: 17)',
        class_str='Lớp 10/1 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Ba, ngày 10/11/2026 – Tiết 1 (07h10 – 07h55) – Phòng H.I.3',
        lesson_row_str='Mục 4.12. Phần mềm thiết kế đồ họa (tiết 1 - LT) – Tiết PPCT: 17 (Tổng bài: 02 tiết – PPCT: 17, 18)',
        page1_title='TÊN BÀI DẠY: BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (TIẾT 1 – LÍ THUYẾT)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 17, 18] – Kế hoạch này chi tiết hoá Tiết thứ 1 của bài / Tiết PPCT: 17)',
        page1_class_line='Môn học: Tin học; Lớp: 10/1',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 1 - LT / Tiết PPCT: 17)',
        page1_sched_line='Lịch dự giờ: Thứ Ba, ngày 10/11/2026 – Tiết 1 (07h10 – 07h55) – Lớp 10/1 – Phòng H.I.3',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp quan sát trực quan và thảo luận nhóm'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E (Ứng dụng tin học – Phần mềm thiết kế đồ hoạ) và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 17 – Mục 4.12. Phần mềm thiết kế đồ họa (tiết 1 - LT)):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu [KT1] – [KT3] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]:** Nêu được khái niệm thiết kế đồ hoạ; phân biệt được hai loại đồ hoạ cơ bản là đồ hoạ điểm ảnh (bitmap) và đồ hoạ vectơ (vector) theo 5 tiêu chí trong Bảng so sánh SGK trang 64; giải thích được lí do chọn đồ hoạ vectơ khi thiết kế logo, bảng hiệu.',
        '**[KT2]:** Phân loại được các phần mềm đồ hoạ thông dụng (Inkscape, Illustrator, CorelDRAW, Photoshop, GIMP); nhận diện đúng các thành phần trên màn hình làm việc của Inkscape (Hình 12.3 SGK trang 65).',
        '**[KT3]:** Trình bày được quy tắc thứ tự lớp (Layer order – đối tượng vẽ trước ở lớp dưới, vẽ sau ở lớp trên), quy trình 3 bước vẽ đối tượng hình học cơ bản (Hình 12.4 SGK trang 66), thao tác chọn/co dãn/xoay bằng công cụ Select, cách chọn màu tô (Fill), màu vẽ (Stroke với phím Shift) và lưu tệp .svg.'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Học sinh chủ động đọc hiểu Sách giáo khoa Tin học 10 trang 63–67, quan sát Hình 12.1 đến Hình 12.4 để hoàn thành các Phiếu học tập.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Học sinh tích cực thảo luận cặp đôi trong các hoạt động so sánh đồ hoạ bitmap – vectơ và bài toán sắp xếp thứ tự lớp.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Học sinh giải thích được cách xử lí tình huống các đối tượng đồ hoạ che khuất nhau và cách kết hợp các hình cơ bản.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực sử dụng và quản lí các phương tiện CNTT và truyền thông [NLa]:** Nhận biết và gọi đúng tên các vùng chức năng, công cụ vẽ cơ bản của phần mềm thiết kế đồ hoạ Inkscape.',
        '**Năng lực ứng xử phù hợp trong môi trường số [NLb]:** Có ý thức sử dụng phần mềm nguồn mở miễn phí Inkscape thay cho phần mềm bẻ khoá nhằm tôn trọng bản quyền phần mềm (kế thừa Bài 11 – Mục 3.11).',
        '**Năng lực hợp tác trong môi trường số [NLe]:** Phối hợp hiệu quả với bạn cùng nhóm để hoàn thành nhiệm vụ học tập đúng tiến độ.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Bước đầu nắm vững nguyên lí tạo và chỉnh sửa nội dung đồ hoạ số định dạng vectơ (.svg).', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Học sinh tập trung quan sát, ghi chép đầy đủ hệ thống khái niệm và quy tắc thao tác trong Sách giáo khoa.',
        '**Phẩm chất trách nhiệm [PC2]:** Học sinh hoàn thành nghiêm túc các Phiếu học tập cá nhân và cặp đôi đúng thời gian quy định.',
        '**Phẩm chất trung thực [PC3]:** Học sinh tự giác làm bài tập củng cố và đánh giá khách quan kết quả của bạn.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính của giáo viên đã cài đặt phần mềm Inkscape (phiên bản 1.0 trở lên), máy chiếu tại phòng học H.I.3, bảng viết và bộ Phiếu học tập.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, Bài 12: Phần mềm thiết kế đồ hoạ (mục 1, mục 2, mục 3 và phần Luyện tập, Vận dụng – trang 63 đến trang 68).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 17 – Mục 4.12. Phần mềm thiết kế đồ họa (tiết 1 - LT)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 1 - Lí thuyết (45 phút – Tiết thứ 1 của bài / Tiết PPCT: 17)**', first=0, keep=True)
    ov_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng & Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Quan sát Hình 12.1a, 12.1b & Hình 12.2)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 17)', '[KT1], [NLC1],\n[PC1]', 'Nhận xét sự khác biệt giữa ảnh chụp và hình vẽ khi phóng to.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Thiết kế đồ hoạ & Phân biệt Bitmap – Vector (09p)\n- 2.2. Nhiệm vụ 2: Phần mềm đồ hoạ & Giao diện Inkscape (09p)\n- 2.3. Nhiệm vụ 3: Các đối tượng đồ hoạ, thứ tự lớp & tô màu (10p)', '28 phút\n(Tiết 1 của bài /\nTiết PPCT: 17)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2],\n[NLa], [NLb], [NLe],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, số 2, số 3 đã hoàn thành.\nCông cụ đánh giá: CCĐG 2, CCĐG 3, CCĐG 4.'],
        ['3. Hoạt động 3: Luyện tập\n(Giải Bài Luyện tập 1 & 2 SGK trang 68)', '07 phút\n(Tiết 1 của bài /\nTiết PPCT: 17)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC3],\n[NLa], [PC3]', 'Lời giải Phiếu học tập số 4 (Bài Luyện tập 1 & 2 trang 68).\nCông cụ đánh giá: CCĐG 5.'],
        ['4. Hoạt động 4: Vận dụng\n(Phân tích thuật toán xếp lớp vẽ Quốc kì Việt Nam & các nước)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 17)', '[KT3], [NLC3],\n[NLa], [PC2]', 'Sơ đồ phân tích các bước vẽ Hình 12.5, 12.6 chuẩn bị cho Tiết 2 Thực hành (Tiết PPCT: 18).']
    ]
    b.append(table(ov_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [PC1] – Học sinh nhận biết trực quan sự khác biệt về chất lượng hiển thị khi phóng to giữa ảnh chụp (bitmap) và hình vẽ bằng phần mềm (vector).', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.1a, Hình 12.1b (SGK trang 63) và Hình 12.2 (SGK trang 64) để nhận xét về màu sắc, độ nét khi phóng to:', first=0)
    add_figure(b, 'hinh_12_1_bitmap_vector.png')
    add(b, '**c) Sản phẩm:** Câu trả lời miệng của học sinh chỉ ra Hình 12.1a bị vỡ nét khi phóng to còn Hình 12.1b giữ nguyên độ sắc nét.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 1 – Bảng kiểm quan sát Hoạt động Mở đầu *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD1_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (28 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu khái niệm thiết kế đồ hoạ và phân biệt đồ hoạ điểm ảnh với đồ hoạ vectơ (09 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [PC1] – Học sinh phát biểu được khái niệm thiết kế đồ hoạ và lập được bảng so sánh 5 tiêu chí giữa đồ hoạ điểm ảnh (bitmap) và đồ hoạ vectơ (vector).', first=0)
    add(b, '**b) Nội dung:** Học sinh nghiên cứu mục 1. Thiết kế đồ hoạ (SGK trang 63–64), quan sát Hình 12.2 và Bảng so sánh (SGK trang 64) để hoàn thành Phiếu học tập số 1:', first=0)
    add_figure(b, 'hinh_12_2_va_bang_so_sanh.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 1 hoàn chỉnh bảng so sánh và trả lời đúng 2 câu hỏi củng cố trang 64.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 2 – Thang điểm chấm Phiếu học tập số 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV1_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu phần mềm đồ hoạ và màn hình làm việc của Inkscape (09 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLa], [NLb], [PC1], [PC2] – Học sinh phân loại đúng hai nhóm phần mềm đồ hoạ và xác định chính xác các thành phần trên màn hình làm việc của phần mềm Inkscape.', first=0)
    add(b, '**b) Nội dung:** Học sinh đọc mục 2. Phần mềm đồ hoạ (SGK trang 64–66), quan sát Hình 12.3 (SGK trang 65) và hoàn thành Phiếu học tập số 2:', first=0)
    add_figure(b, 'hinh_12_3_man_hinh_inkscape.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 2 xác định đúng chức năng các thanh công cụ trên màn hình Inkscape và đáp án 2 câu hỏi đầu trang 66.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 3 – Thang điểm chấm Phiếu học tập số 2 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV2_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các đối tượng đồ hoạ của hình vẽ, thứ tự lớp và cách tô màu (10 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT3], [NLC2], [NLC3], [NLa], [NLe], [NLS 3.1.NC1.a], [PC2] – Học sinh giải thích được quy tắc thứ tự lớp (Layer order), quy trình 3 bước vẽ đối tượng có sẵn (Hình 12.4) và cách chọn màu tô (Fill), màu vẽ (Stroke).', first=0)
    add(b, '**b) Nội dung:** Học sinh nghiên cứu mục 3. Các đối tượng đồ hoạ của hình vẽ (SGK trang 66–67), quan sát Hình 12.4 (SGK trang 66) và hoàn thành Phiếu học tập số 3:', first=0)
    add_figure(b, 'hinh_12_4_cong_cu_inkscape.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 3 giải quyết đúng 4 trường hợp xếp lớp quả dưa hấu – quả cam và các quy tắc thao tác đối tượng.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 4 – Thang điểm chấm Phiếu học tập số 3 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV3_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (07 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [NLC1], [NLC3], [NLa], [PC3] – Học sinh vận dụng lý thuyết vừa học để giải chính xác Bài Luyện tập 1 và Bài Luyện tập 2 trong Sách giáo khoa trang 68.', first=0)
    add(b, '**b) Nội dung:** Học sinh làm việc cá nhân hoàn thành Phiếu học tập số 4 (gồm Bài Luyện tập 1 và Bài Luyện tập 2 SGK trang 68).', first=0)
    add(b, '**c) Sản phẩm:** Lời giải chính xác trên Phiếu học tập số 4 (hình tròn có nền màu Cam, viền màu Đỏ đậm; chọn đáp án C. Hộp công cụ).', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 5 – Đáp án và thang điểm chấm Phiếu học tập số 4 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (05 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [NLC3], [NLa], [PC2] – Học sinh vận dụng quy tắc thứ tự lớp và thuộc tính đối tượng để lập sơ đồ các bước vẽ bông hoa (Hình 12.5), Quốc kì Việt Nam (Hình 12.6) và Quốc kì các nước (Vận dụng trang 68), sẵn sàng cho Tiết 2 Thực hành (Tiết PPCT: 18).', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát trước Hình 12.5, Hình 12.6 và hình Vận dụng (SGK trang 67–68) để phân tích quy trình xếp lớp:', first=0)
    add_figure(b, 'hinh_12_van_dung_quoc_ki.png')
    add(b, '**c) Sản phẩm:** Bản ghi tóm tắt quy trình thực hành trong vở để thực hiện trên phòng máy ở Tiết 2 (Tiết PPCT: 18).', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Rubric đánh giá sản phẩm thực hành và vận dụng *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add_bai12_appendix(b, class_label='10/1 (Dự giờ Tiết 1 – Phòng H.I.3)', ppct_label='17')
    add_approval_section(b, sign_date_str='Huế, ngày 09 tháng 11 năm 2026')

    used_figs = [
        'hinh_12_1_bitmap_vector.png',
        'hinh_12_2_va_bang_so_sanh.png',
        'hinh_12_3_man_hinh_inkscape.png',
        'hinh_12_4_cong_cu_inkscape.png',
        'hinh_12_van_dung_quoc_ki.png'
    ]
    out_p = KHOI10_DIR / '03_De_cuong_du_gio_Tin10_Bai12_Tiet1_LT_TietPPCT17_Lop10-1.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin10_Bai12_Tiet1.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / '03_KHBD_Tin10_Bai12_Tiet1.docx')
    print('Built:', out_p)

def build_bai12_tiet2():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (TIẾT 2 - THỰC HÀNH – TIẾT PPCT: 18)',
        class_str='Lớp 10/2 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Ba, ngày 10/11/2026 – Tiết 2 (08h00 – 08h45) – Phòng H.I.4',
        lesson_row_str='Mục 4.12. Phần mềm thiết kế đồ họa (tiết 2 - TH) – Tiết PPCT: 18 (Tổng bài: 02 tiết – PPCT: 17, 18)',
        page1_title='TÊN BÀI DẠY: BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (TIẾT 2 – THỰC HÀNH)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 17, 18] – Kế hoạch này chi tiết hoá Tiết thứ 2 của bài / Tiết PPCT: 18)',
        page1_class_line='Môn học: Tin học; Lớp: 10/2',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 2 - TH / Tiết PPCT: 18)',
        page1_sched_line='Lịch dự giờ: Thứ Ba, ngày 10/11/2026 – Tiết 2 (08h00 – 08h45) – Lớp 10/2 – Phòng H.I.4',
        page1_mode_line='Hình thức tổ chức: Dạy học thực hành trên phần mềm Inkscape'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E (Ứng dụng tin học – Phần mềm thiết kế đồ hoạ) và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 18 – Mục 4.12. Phần mềm thiết kế đồ họa (tiết 2 - TH)):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức và kỹ năng thực hành (ký hiệu [KT1] – [KT3] để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**[KT1]:** Vận dụng chuẩn xác quy trình tạo đối tượng hình tròn, hình elip, nhân bản đối tượng (Duplicate / Ctrl + D), chọn và xoay đối tượng bằng công cụ Select để vẽ hoàn chỉnh bông hoa theo Hình 12.5 (SGK trang 67).',
        '**[KT2]:** Thiết lập đúng tham số của công cụ Hình đa giác/ngôi sao (Corners = 5, Spoke ratio = 0.4), kết hợp công cụ Hình chữ nhật (tỉ lệ 2 : 3), tô màu nền/viền và điều chỉnh thứ tự lớp (Lower selection to bottom) để vẽ hoàn chỉnh Quốc kì Việt Nam theo Hình 12.6 (SGK trang 67–68).',
        '**[KT3]:** Giải quyết được các bài tập thực hành nâng cao thuộc phần Luyện tập và Vận dụng (SGK trang 68): phối màu Fill/Stroke, vẽ Quốc kì các nước (Thái Lan, Lào, Myanmar, Anh) và ghép các hình cơ bản thành hình con vật.'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Học sinh tự đọc hướng dẫn thực hành trong Sách giáo khoa trang 67–68 và chủ động thao tác trên phần mềm Inkscape để tạo sản phẩm.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Học sinh trao đổi theo cặp máy, hỗ trợ nhau kiểm tra thông số Corners, Spoke ratio và thứ tự lớp.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Học sinh tự phát hiện và xử lí lỗi khi đối tượng vẽ sau che khuất đối tượng vẽ trước hoặc khi tâm xoay cánh hoa bị lệch.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực sử dụng và quản lí các phương tiện CNTT và truyền thông [NLa]:** Sử dụng thành thạo phần mềm Inkscape để khởi tạo tệp mới (Ctrl + N), vẽ hình, biến đổi đối tượng và lưu tệp định dạng chuẩn .svg (Ctrl + S).',
        '**Năng lực giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông [NLc]:** Phân tích một hình vẽ phức tạp (bông hoa, lá cờ, con vật) thành các đối tượng hình học đơn giản và sắp xếp đúng thứ tự các lớp.',
        '**Năng lực ứng dụng CNTT và truyền thông trong học và tự học [NLd]:** Khai thác hiệu quả các phím tắt và thanh điều khiển thuộc tính của Inkscape để tạo sản phẩm thẩm mỹ.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Tạo và hoàn thiện sản phẩm đồ hoạ vectơ số hoàn chỉnh (.svg) theo yêu cầu kỹ thuật.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Học sinh kiên trì căn chỉnh từng cánh hoa, góc xoay và tỉ lệ lá cờ cho đến khi sản phẩm cân đối, đẹp mắt.',
        '**Phẩm chất trách nhiệm [PC2]:** Học sinh tuân thủ nội quy phòng học/phòng máy H.I.4, lưu trữ tệp đúng quy định và hoàn thành sản phẩm đúng thời gian.',
        '**Phẩm chất trung thực [PC3]:** Học sinh tự tay thực hiện sản phẩm thực hành và đánh giá công tâm theo bảng tiêu chí Rubric CCĐG 6.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính phòng học/phòng máy H.I.4 đã cài đặt sẵn phần mềm Inkscape, máy chiếu, tệp mẫu .svg và phiếu đánh giá Rubric.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, Bài 12: Phần mềm thiết kế đồ hoạ (mục Thực hành, Luyện tập và Vận dụng – trang 67–68).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 18 – Mục 4.12. Phần mềm thiết kế đồ họa (tiết 2 - TH)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 2 - Thực hành (45 phút – Tiết thứ 2 của bài / Tiết PPCT: 18)**', first=0, keep=True)
    ov2_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng & Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Kích hoạt kỹ năng thao tác Inkscape)', '05 phút\n(Tiết 2 của bài /\nTiết PPCT: 18)', '[KT1], [KT2],\n[NLC1], [NLa]', 'Khởi động Inkscape, trả lời nhanh 3 thao tác cốt lõi (Select, Fill/Stroke, Thứ tự lớp).\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới (Hướng dẫn & Thực hành cốt lõi SGK tr. 67–68)\n- 2.1. Nhiệm vụ 1: Thực hành vẽ bông hoa (Hình 12.5 – 12 phút)\n- 2.2. Nhiệm vụ 2: Thực hành vẽ Quốc kì Việt Nam (Hình 12.6 – 12 phút)', '24 phút\n(Tiết 2 của bài /\nTiết PPCT: 18)', '[KT1], [KT2],\n[NLC1], [NLC2],\n[NLa], [NLc], [NLd],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Tệp HoTen_10-2_Bai12.svg chứa Bông hoa (Hình 12.5) và Quốc kì Việt Nam (Hình 12.6).\nCông cụ đánh giá: CCĐG 6 (Rubric thực hành).'],
        ['3. Hoạt động 3: Luyện tập\n(Thực hành kiểm chứng Bài Luyện tập 1 & 2 SGK tr. 68)', '08 phút\n(Tiết 2 của bài /\nTiết PPCT: 18)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC3],\n[NLa], [PC3]', 'Hình tròn nền Cam viền Đỏ đậm trên tệp .svg và phiếu trả lời Luyện tập.\nCông cụ đánh giá: CCĐG 5 & CCĐG 6.'],
        ['4. Hoạt động 4: Vận dụng\n(Thực hành vẽ Quốc kì các nước & Ghép hình cơ bản SGK tr. 68)', '08 phút\n(Tiết 2 của bài /\nTiết PPCT: 18)', '[KT3], [NLC2],\n[NLC3], [NLc],\n[PC1], [PC2]', 'Hình vẽ Quốc kì Lào / Thái Lan / Myanmar bổ sung vào tệp HoTen_10-2_Bai12.svg.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov2_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [NLC1], [NLa] – Học sinh tái hiện nhanh các quy tắc thao tác đối tượng đã học ở Tiết 1 (Tiết PPCT: 17), khởi động phần mềm Inkscape và tạo tệp làm việc mới.', first=0)
    add(b, '**b) Nội dung:** Giáo viên yêu cầu học sinh mở phần mềm Inkscape, lưu tệp mới với tên HoTen_10-2_Bai12.svg và trả lời nhanh 03 câu hỏi kiểm tra kỹ năng nền tảng trước khi thực hành.', first=0)
    add(b, '**c) Sản phẩm:** Cửa sổ phần mềm Inkscape sẵn sàng trên máy và câu trả lời đúng 03 thao tác cốt lõi.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 1 – Bảng kiểm quan sát thao tác chuẩn bị thực hành *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    hd1_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên yêu cầu học sinh lớp 10/2 khởi động phần mềm Inkscape, nhấn Ctrl + S lưu tệp vào thư mục thực hành với tên HoTen_10-2_Bai12.svg và trả lời nhanh 03 câu hỏi:\n'
            '  + Câu 1: Muốn xoay một đối tượng đã vẽ, ta dùng công cụ Select nháy vào đối tượng mấy lần?\n'
            '  + Câu 2: Làm thế nào để đổi màu đường viền (Stroke) của đối tượng từ Bảng màu?\n'
            '  + Câu 3: Khi đối tượng vẽ sau che khuất đối tượng vẽ trước, ta xử lí như thế nào?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh khởi động Inkscape, tạo tệp .svg và giơ tay trả lời nhanh 3 câu hỏi.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh trả lời và thao tác thử nghiệm 10 giây trên máy chiếu.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét và giao nhiệm vụ thực hành trọng tâm của Tiết 2 (Tiết PPCT: 18): Hoàn thành Hình 12.5, Hình 12.6 và các bài Luyện tập, Vận dụng (SGK trang 67–68).',
            '**Kiến thức thao tác cốt lõi cần nhớ khi thực hành:**\n'
            '1. **Xoay đối tượng:** Chọn công cụ Select -> nháy chuột lần 2 vào đối tượng -> kéo mũi tên cong ở góc.\n'
            '2. **Đổi màu:** Nháy chuột trái = Màu tô nền (Fill); Giữ Shift + Nháy chuột trái = Màu vẽ viền (Stroke).\n'
            '3. **Thứ tự lớp:** Chọn đối tượng -> nháy nút Lower selection to bottom (hạ xuống lớp dưới cùng) hoặc Raise selection to top (đưa lên lớp trên cùng).'
        ]
    ]
    b.append(table(hd1_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới – Hướng dẫn và Thực hành Nhiệm vụ 1, Nhiệm vụ 2 (24 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Thực hành vẽ bông hoa từ hình tròn và những hình elip – Hình 12.5 SGK trang 67 (12 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [NLa], [NLc], [NLS 3.1.NC1.a], [PC1] – Học sinh sử dụng thành thạo công cụ Hình tròn/elip, lệnh nhân bản Duplicate (Ctrl + D) và chế độ xoay của công cụ Select để vẽ hoàn chỉnh bông hoa theo Hình 12.5.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát hướng dẫn Nhiệm vụ 1 – Hình 12.5 (SGK trang 67) và thực hành trực tiếp trên phần mềm Inkscape:', first=0)
    add_figure(b, 'hinh_12_5_ve_bong_hoa.png')
    add(b, '**c) Sản phẩm:** Hình vẽ bông hoa hoàn chỉnh (nhụy tròn + 10 đến 12 cánh hoa elip xoay đều) được lưu trong tệp HoTen_10-2_Bai12.svg.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Tiêu chí 1 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv1_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 12.5 (SGK trang 67), làm mẫu nhanh thao tác nhân bản cánh hoa bằng Duplicate (Ctrl + D) và xoay cánh hoa trong 02 phút, sau đó yêu cầu học sinh thực hành trong 08 phút theo đúng 3 bước trong SGK trang 67.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh thực hành vẽ nhụy hoa hình tròn, vẽ 01 cánh hoa hình elip, nhấn Ctrl + D tạo các bản sao cánh hoa, nháy lần 2 công cụ Select để xoay từng cánh hoa và di chuyển ghép quanh nhụy hoa.\n'
            '- Giáo viên quan sát từng dãy máy, hướng dẫn học sinh giữ phím Ctrl khi vẽ hình tròn để nhụy hoa tròn đều.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên chọn 02 bài làm tiêu biểu chiếu lên màn hình, mời học sinh nêu lại tác dụng của lệnh Duplicate so với việc vẽ từng cánh hoa bằng tay.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên nhận xét theo Tiêu chí 1 của Rubric CCĐG 6 (3,5 điểm) và nhắc học sinh nhấn Ctrl + S cập nhật tệp.',
            '**1. Thực hành Nhiệm vụ 1: Vẽ bông hoa (Hình 12.5, SGK trang 67):**\n'
            '- **Bước 1:** Chọn công cụ Hình tròn/elip (E), vẽ 01 hình tròn (nhụy hoa) và 01 hình elip đứng (cánh hoa), chọn màu vàng trên Bảng màu (Hình 12.5a).\n'
            '- **Bước 2:** Chọn hình elip -> nháy nút phải chuột chọn Duplicate (hoặc nhấn Ctrl + D) để tạo các bản sao cánh hoa có kích thước bằng hệt nhau (Hình 12.5b).\n'
            '- **Bước 3:** Chọn công cụ Select (S): nháy chuột lần 1 lên cánh hoa (chế độ co dãn – Hình 12.5c), nháy chuột lần 2 để mở chế độ xoay (Hình 12.5d); kéo mũi tên góc xoay cánh hoa theo hướng phù hợp (Hình 12.5e) và xếp đều quanh nhụy hoa (Hình 12.5g).'
        ]
    ]
    b.append(table(nv1_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Thực hành vẽ Quốc kì Việt Nam – Hình 12.6 SGK trang 67–68 (12 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLC3], [NLa], [NLc], [NLd], [NLS 3.1.NC1.a], [PC1], [PC2] – Học sinh thiết lập chính xác thuộc tính ngôi sao 5 cánh (Corners = 5, Spoke ratio = 0.4), vẽ nền cờ chữ nhật tỉ lệ 2 : 3 và sử dụng thành thạo nút hạ lớp Lower selection to bottom để hoàn thiện Quốc kì Việt Nam.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.6 (SGK trang 68) và thực hành vẽ Quốc kì Việt Nam ngay bên cạnh hình bông hoa trên tệp HoTen_10-2_Bai12.svg:', first=0)
    add_figure(b, 'hinh_12_6_ve_quoc_ki.png')
    add(b, '**c) Sản phẩm:** Hình vẽ Quốc kì Việt Nam đúng tỉ lệ, ngôi sao 5 cánh vàng nằm ở lớp trên chính giữa nền cờ đỏ.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Tiêu chí 2 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    nv2_t2_rows = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên chiếu Hình 12.6 (SGK trang 68) và yêu cầu học sinh thực hiện đúng 3 bước trong SGK trang 67–68 để vẽ Quốc kì Việt Nam:\n'
            '  + Thiết lập công cụ Hình sao với Corners = 5, Spoke ratio = 0.4, tô màu vàng.\n'
            '  + Vẽ hình chữ nhật tỉ lệ 2 : 3 (ví dụ: rộng 120 mm, dài 180 mm trên thanh thuộc tính W, H), tô nền màu đỏ.\n'
            '  + Xử lí hiện tượng hình chữ nhật vẽ sau che mất ngôi sao bằng nút Lower selection to bottom.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh thực hành trên máy tính. Giáo viên hỗ trợ các học sinh đang để chế độ Đa giác (Polygon) chuyển sang chế độ Hình sao (Star) trên Thanh điều khiển thuộc tính.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh trình bày sản phẩm và giải thích vì sao ở Bước 2 ngôi sao lại bị che khuất và nút lệnh nào ở Bước 3 đã giải quyết vấn đề đó.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên chấm điểm theo Tiêu chí 2 của Rubric CCĐG 6 (3,5 điểm).',
            '**2. Thực hành Nhiệm vụ 2: Vẽ Quốc kì Việt Nam (Hình 12.6, SGK trang 67–68):**\n'
            '- **Bước 1 (Vẽ ngôi sao 5 cánh):** Chọn công cụ Hình đa giác/ngôi sao (phím dấu hoa thị) -> trên Thanh điều khiển thuộc tính chọn nút Hình sao, đặt Corners = 5 và Spoke ratio = 0.4 -> vẽ ngôi sao và tô màu vàng (Hình 12.6a).\n'
            '- **Bước 2 (Vẽ nền cờ đỏ):** Chọn công cụ Hình vuông/chữ nhật (R), vẽ hình chữ nhật tỉ lệ chiều rộng : chiều dài là 2 : 3, tô nền màu đỏ (Hình 12.6b). Vì hình chữ nhật vẽ sau nên nằm ở lớp trên và che khuất ngôi sao khi xếp chồng.\n'
            '- **Bước 3 (Đưa hình chữ nhật xuống lớp dưới):** Chọn công cụ Select -> chọn hình chữ nhật đỏ -> nháy nút Lower selection to bottom trên Thanh điều khiển thuộc tính để đưa hình chữ nhật xuống lớp dưới, giúp ngôi sao vàng nổi bật chính giữa nền đỏ (Hình 12.6c).'
        ]
    ]
    b.append(table(nv2_t2_rows, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (08 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [KT2], [KT3], [NLC1], [NLC3], [NLa], [PC3] – Học sinh thực hành kiểm chứng trực tiếp Bài Luyện tập 1 và trả lời Bài Luyện tập 2 (SGK trang 68).', first=0)
    add(b, '**b) Nội dung:** Học sinh vẽ 01 hình tròn trên tệp thực hành, thực hiện đúng chuỗi thao tác của Bài Luyện tập 1 (SGK trang 68) và trả lời Bài Luyện tập 2 (SGK trang 68) vào Phiếu học tập số 4.', first=0)
    add(b, '**c) Sản phẩm:** Hình tròn có nền màu Cam, viền dày màu Đỏ đậm trên tệp .svg và đáp án Phiếu học tập số 4.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 5 & CCĐG 6 *(Đính kèm tại Mục IV. Phụ lục 1 & 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (08 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [NLC2], [NLC3], [NLc], [PC1], [PC2] – Học sinh vận dụng kỹ năng vẽ hình cơ bản, nhân bản và sắp xếp thứ tự lớp để vẽ Quốc kì các nước (Thái Lan, Lào, Myanmar) và phân tích cách ghép hình con vật ở phần Vận dụng SGK trang 68.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát hình mẫu phần Vận dụng (SGK trang 68) và thực hành vẽ ít nhất 01 Quốc kì (Thái Lan, Lào hoặc Myanmar) ngay tại lớp:', first=0)
    add_figure(b, 'hinh_12_van_dung_quoc_ki.png')
    add(b, '**c) Sản phẩm:** Tệp HoTen_10-2_Bai12.svg hoàn chỉnh gồm Bông hoa, Quốc kì Việt Nam, Hình tròn Luyện tập và Quốc kì tự chọn ở phần Vận dụng.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành Inkscape *(Đính kèm chi tiết tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add_bai12_appendix(b, class_label='10/2 (Dự giờ Tiết 2 – Phòng H.I.4)', ppct_label='18')
    add_approval_section(b, sign_date_str='Huế, ngày 09 tháng 11 năm 2026')

    used_figs = [
        'hinh_12_5_ve_bong_hoa.png',
        'hinh_12_6_ve_quoc_ki.png',
        'hinh_12_van_dung_quoc_ki.png'
    ]
    out_p = KHOI10_DIR / '04_De_cuong_du_gio_Tin10_Bai12_Tiet2_TH_TietPPCT18_Lop10-2.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin10_Bai12_Tiet2.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / '04_KHBD_Tin10_Bai12_Tiet2.docx')
    print('Built:', out_p)

def build_bai12_cabai():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (02 TIẾT – TIẾT PPCT: 17, 18)',
        class_str='Lớp 10/1 (Tiết 1) & Lớp 10/2 (Tiết 2) – Trường THPT Thuận Hoá',
        schedule_str='Thứ Ba, ngày 10/11/2026 – Tiết 1 (7h10–7h55, 10/1, H.I.3) & Tiết 2 (8h00–8h45, 10/2, H.I.4)',
        lesson_row_str='Bài 12. Phần mềm thiết kế đồ hoạ (02 tiết – Tiết PPCT: 17, 18)',
        page1_title='TÊN BÀI DẠY: BÀI 12. PHẦN MỀM THIẾT KẾ ĐỒ HOẠ',
        page1_subtitle='(Thời lượng: 02 tiết = 90 phút – Tiết thứ 1, 2 của bài; Tiết PPCT: 17, 18)',
        page1_class_line='Môn học: Tin học; Lớp: 10/1 & 10/2',
        page1_time_line='Thời gian thực hiện: 02 tiết (90 phút – Tiết PPCT: 17 [Lí thuyết], Tiết PPCT: 18 [Thực hành])',
        page1_sched_line='Lịch dự giờ: Thứ Ba, 10/11/2026 – Tiết 1 (7h10–7h55, Lớp 10/1, Phòng H.I.3) & Tiết 2 (8h00–8h45, Lớp 10/2, Phòng H.I.4)',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp thực hành tại phòng máy'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E (Ứng dụng tin học – Phần mềm thiết kế đồ hoạ) và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 17, 18 – Mục 4.12. Phần mềm thiết kế đồ họa):')
    add(b, '1. Về kiến thức', 'sub')
    for item in [
        '**[KT1]:** Nêu được khái niệm thiết kế đồ hoạ; phân biệt được đồ hoạ điểm ảnh (bitmap) và đồ hoạ vectơ (vector) theo 5 tiêu chí tại Bảng trang 64 SGK.',
        '**[KT2]:** Phân loại được các phần mềm đồ hoạ thông dụng và nhận diện đúng các thành phần trên màn hình làm việc của phần mềm Inkscape (Hình 12.3 trang 65).',
        '**[KT3]:** Trình bày và thực hiện được quy tắc thứ tự lớp (Layer order), 3 bước vẽ đối tượng có sẵn (Hình 12.4 trang 66), thao tác chọn/co dãn/xoay bằng công cụ Select, chọn màu tô (Fill), màu vẽ (Stroke) và lưu tệp .svg.',
        '**[KT4]:** Thực hành vẽ hoàn chỉnh bông hoa (Hình 12.5 trang 67), Quốc kì Việt Nam (Hình 12.6 trang 68) và giải quyết các bài tập Luyện tập, Vận dụng (trang 68).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**Năng lực tự chủ và tự học [NLC1]:** Chủ động nghiên cứu Sách giáo khoa trang 63–68 và tự thực hành trên phần mềm Inkscape.',
        '**Năng lực giao tiếp và hợp tác [NLC2]:** Thảo luận cặp đôi hiệu quả khi hoàn thành Phiếu học tập và hỗ trợ nhau trong giờ thực hành.',
        '**Năng lực giải quyết vấn đề và sáng tạo [NLC3]:** Phát hiện và xử lí đúng các tình huống đối tượng bị che khuất hoặc sai tỉ lệ hình học.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**Năng lực sử dụng và quản lí các phương tiện CNTT và truyền thông [NLa]:** Làm chủ giao diện, hộp công cụ, thanh điều khiển thuộc tính, bảng màu và quản lí tệp .svg trong Inkscape.',
        '**Năng lực ứng xử phù hợp trong môi trường số [NLb]:** Tôn trọng bản quyền phần mềm thông qua việc lựa chọn phần mềm nguồn mở miễn phí Inkscape.',
        '**Năng lực giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông [NLc]:** Phân rã hình mẫu phức tạp thành các hình cơ bản và sắp xếp đúng thứ tự lớp để tạo sản phẩm đồ hoạ.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**[NLS 3.1.NC1.a]:** Tạo và phát triển sản phẩm đồ hoạ vectơ số hoàn chỉnh bằng phần mềm chuyên dụng.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**Phẩm chất chăm chỉ [PC1]:** Tích cực học tập lý thuyết và kiên trì rèn luyện kỹ năng thực hành vẽ hình.',
        '**Phẩm chất trách nhiệm [PC2]:** Hoàn thành đầy đủ nhiệm vụ học tập ở cả 2 tiết và tuân thủ nội quy phòng học/phòng máy.',
        '**Phẩm chất trung thực [PC3]:** Tự giác thực hiện sản phẩm và đánh giá trung thực theo Rubric.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính giáo viên và học sinh cài đặt phần mềm Inkscape, máy chiếu tại phòng H.I.3 và H.I.4.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, Bài 12: Phần mềm thiết kế đồ hoạ (trang 63–68).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 17 [tiết 1 - LT] và Tiết PPCT 18 [tiết 2 - TH]).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học toàn bài (02 tiết = 90 phút – Tiết PPCT: 17, 18)**', first=0, keep=True)
    ov_all = [
        ['Hoạt động dạy học (Chuẩn Công văn 5512)', 'Thời lượng & Phân bổ Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(So sánh ảnh chụp và hình vẽ – Hình 12.1a, 12.1b)', '05 phút\n(Thuộc Tiết 1 của bài /\nTiết PPCT: 17)', '[KT1], [NLC1], [PC1]', 'Câu trả lời nhận xét Hình 12.1a & 12.1b.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới (65 phút)\n- 2.1. Nhiệm vụ 1: Thiết kế đồ hoạ & Bitmap – Vector (12p – Tiết 1 / PPCT: 17)\n- 2.2. Nhiệm vụ 2: Phần mềm đồ hoạ & Giao diện Inkscape (13p – Tiết 1 / PPCT: 17)\n- 2.3. Nhiệm vụ 3: Đối tượng đồ hoạ, thứ tự lớp & tô màu (15p – Tiết 1 / PPCT: 17)\n- 2.4. Nhiệm vụ 4: Thực hành vẽ bông hoa & Quốc kì VN (25p – Tiết 2 / PPCT: 18)', '65 phút\n(40p ở Tiết 1 / PPCT: 17\n+ 25p ở Tiết 2 / PPCT: 18)', '[KT1], [KT2], [KT3], [KT4],\n[NLC1], [NLC2], [NLC3],\n[NLa], [NLb], [NLc],\n[NLS 3.1.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, 2, 3 (Tiết 1) và tệp HoTen_Lop_Bai12.svg chứa Hình 12.5, 12.6 (Tiết 2).\nCông cụ đánh giá: CCĐG 2, 3, 4, 6.'],
        ['3. Hoạt động 3: Luyện tập\n(Giải và kiểm chứng Bài Luyện tập 1 & 2 SGK tr. 68)', '10 phút\n(Thuộc Tiết 2 của bài /\nTiết PPCT: 18)', '[KT3], [KT4],\n[NLC1], [NLC3],\n[NLa], [PC3]', 'Phiếu học tập số 4 và hình tròn nền Cam viền Đỏ đậm.\nCông cụ đánh giá: CCĐG 5.'],
        ['4. Hoạt động 4: Vận dụng\n(Thực hành vẽ Quốc kì các nước & Ghép hình SGK tr. 68)', '10 phút\n(Thuộc Tiết 2 của bài /\nTiết PPCT: 18)', '[KT4], [NLC2], [NLC3],\n[NLc], [PC1], [PC2]', 'Sản phẩm Quốc kì các nước trong tệp .svg.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov_all, [2800, 1500, 1900, 2800], header=True, cell_sz=21))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'heading')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [PC1] – Học sinh nhận biết sự khác nhau giữa ảnh chụp và hình vẽ bằng phần mềm khi phóng to.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.1a, Hình 12.1b (SGK trang 63) và trao đổi cặp đôi:', first=0)
    add_figure(b, 'hinh_12_1_bitmap_vector.png')
    add(b, '**c) Sản phẩm:** Câu trả lời miệng của học sinh chỉ rõ ảnh chụp bị vỡ hạt khi phóng to còn hình vẽ giữ nguyên độ nét.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 1 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD1_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (65 phút: Nhiệm vụ 1, 2, 3 thuộc Tiết 1 / Tiết PPCT: 17; Nhiệm vụ 4 thuộc Tiết 2 / Tiết PPCT: 18)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu khái niệm thiết kế đồ hoạ và phân biệt đồ hoạ điểm ảnh với đồ hoạ vectơ (12 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT1], [NLC1], [NLC2], [PC1] – Học sinh nêu được khái niệm thiết kế đồ hoạ và phân biệt đồ hoạ điểm ảnh với đồ hoạ vectơ theo 5 tiêu chí.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.2 và Bảng so sánh (SGK trang 64) để hoàn thành Phiếu học tập số 1:', first=0)
    add_figure(b, 'hinh_12_2_va_bang_so_sanh.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 1 đã điền đủ 5 tiêu chí so sánh và đáp án 2 câu hỏi củng cố trang 64.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 2 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV1_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu phần mềm đồ hoạ và giao diện làm việc của Inkscape (13 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT2], [NLC1], [NLa], [NLb], [PC1], [PC2] – Học sinh phân loại được phần mềm đồ hoạ và nhận diện các thành phần trên màn hình Inkscape.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.3 (SGK trang 65) và hoàn thành Phiếu học tập số 2:', first=0)
    add_figure(b, 'hinh_12_3_man_hinh_inkscape.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 2 hoàn chỉnh.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 3 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV2_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các đối tượng đồ hoạ của hình vẽ, thứ tự lớp và cách tô màu (15 phút – Tiết 1 của bài / Tiết PPCT: 17)', 'sub')
    add(b, '**a) Mục tiêu:** [KT3], [NLC2], [NLC3], [NLa], [NLS 3.1.NC1.a], [PC2] – Học sinh nắm vững quy tắc thứ tự lớp, 3 bước vẽ đối tượng cơ bản và cách chọn màu tô/màu vẽ.', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát Hình 12.4 (SGK trang 66) và hoàn thành Phiếu học tập số 3:', first=0)
    add_figure(b, 'hinh_12_4_cong_cu_inkscape.png')
    add(b, '**c) Sản phẩm:** Phiếu học tập số 3 hoàn chỉnh.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 4 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV3_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.4. Nhiệm vụ 4: Thực hành vẽ bông hoa (Hình 12.5) và vẽ Quốc kì Việt Nam (Hình 12.6) (25 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'sub')
    add(b, '**a) Mục tiêu:** [KT4], [NLC1], [NLC2], [NLa], [NLc], [NLS 3.1.NC1.a], [PC1], [PC2] – Học sinh thực hành vẽ hoàn chỉnh bông hoa và Quốc kì Việt Nam trên phần mềm Inkscape.', first=0)
    add(b, '**b) Nội dung:** Học sinh thực hành theo hướng dẫn tại Hình 12.5 (SGK trang 67) và Hình 12.6 (SGK trang 68):', first=0)
    add_figure(b, 'hinh_12_5_ve_bong_hoa.png')
    add_figure(b, 'hinh_12_6_ve_quoc_ki.png')
    add(b, '**c) Sản phẩm:** Tệp HoTen_Lop_Bai12.svg chứa hình bông hoa và Quốc kì Việt Nam đúng chuẩn.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Rubric đánh giá sản phẩm thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(NV4_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (10 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '**a) Mục tiêu:** [KT3], [KT4], [NLC1], [NLC3], [NLa], [PC3] – Học sinh giải và thực hành kiểm chứng Bài Luyện tập 1 & 2 (SGK trang 68).', first=0)
    add(b, '**b) Nội dung:** Học sinh hoàn thành Phiếu học tập số 4 và thực hành kiểm chứng màu Fill/Stroke trên máy tính.', first=0)
    add(b, '**c) Sản phẩm:** Phiếu học tập số 4 và hình tròn nền Cam, viền Đỏ đậm.', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 5 *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD3_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (10 phút – Tiết 2 của bài / Tiết PPCT: 18)', 'heading')
    add(b, '**a) Mục tiêu:** [KT4], [NLC2], [NLC3], [NLc], [PC1], [PC2] – Học sinh vận dụng tổng hợp kiến thức để vẽ Quốc kì các nước và ghép hình cơ bản (SGK trang 68).', first=0)
    add(b, '**b) Nội dung:** Học sinh quan sát hình mẫu phần Vận dụng (SGK trang 68) và thực hành trên phần mềm Inkscape:', first=0)
    add_figure(b, 'hinh_12_van_dung_quoc_ki.png')
    add(b, '**c) Sản phẩm:** Tệp HoTen_Lop_Bai12.svg bổ sung Quốc kì các nước (Lào, Thái Lan, Myanmar).', first=0)
    add(b, '**Công cụ đánh giá:** CCĐG 6 – Rubric 10 điểm *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, '**d) Tổ chức thực hiện:**', first=0, keep=True)
    b.append(table(HD4_ROWS_BAI12, [4600, 4400], header=True, cell_sz=23))

    add_bai12_appendix(b, class_label='10/1 & 10/2 (Trọn bộ 02 tiết)', ppct_label='17, 18')
    add_approval_section(b, sign_date_str='Huế, ngày 09 tháng 11 năm 2026')

    used_figs = [
        'hinh_12_1_bitmap_vector.png',
        'hinh_12_2_va_bang_so_sanh.png',
        'hinh_12_3_man_hinh_inkscape.png',
        'hinh_12_4_cong_cu_inkscape.png',
        'hinh_12_5_ve_bong_hoa.png',
        'hinh_12_6_ve_quoc_ki.png',
        'hinh_12_van_dung_quoc_ki.png'
    ]
    out_p = KHOI10_DIR / '05_De_cuong_du_gio_Tin10_Bai12_CaBai_2Tiet_TietPPCT17-18.docx'
    save_docx(b, out_p, used_figures=used_figs)
    print('Built:', out_p)

if __name__ == '__main__':
    build_bai12_tiet1()
    build_bai12_tiet2()
    build_bai12_cabai()
