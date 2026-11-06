# -*- coding: utf-8 -*-
"""
Builds all deliverables for Grade 10 Bài 13: Bổ sung các đối tượng đồ hoạ (Mục 4.13 - Tiết PPCT: 19, 20)
and the standalone Student Handout & Rubric compilation:
1. 06_De_cuong_du_gio_Tin10_Bai13_Tiet1_LT_TietPPCT19_Lop10-3.docx (Dự giờ Thứ Hai 16/11/2026 - Tiết 3, 9h00-9h45, Lớp 10/3, Phòng H.I.5)
2. 07_De_cuong_du_gio_Tin10_Bai13_Tiet2_TH_TietPPCT20_Lop10-3.docx (Tiết 2 Thực hành - Tiết PPCT: 20, Lớp 10/3)
3. 08_De_cuong_du_gio_Tin10_Bai13_CaBai_2Tiet_TietPPCT19-20.docx (Bản gộp trọn vẹn cả Bài 13 - 02 tiết = 90 phút, Tiết PPCT: 19, 20)
4. 09_Phieu_hoc_tap_va_Rubric_Tin10_OnTap_Bai12_Bai13.docx (Bộ Phiếu học tập rời, Rubric và Đáp án in phát học sinh)
"""
from pathlib import Path
import shutil
from docx_engine_5512 import (
    init_body_with_cover, node, add, br, table, add_figure, add_approval_section, save_docx
)

KHOI10_DIR = Path(r'c:\Users\PC\khbd_docx\khoi10')
THANH_PHAM_DIR = Path(r'c:\Users\PC\khbd_docx\06_thanh_pham_docx')
GOI_NOP_DIR = Path(r'c:\Users\PC\khbd_docx\08_goi_nop_du_gio\HoQuangHuy_24S1020030_HoSo_DuGio_Tin10')

HD1_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu lên màn hình hình vẽ miếng dưa hấu ở phần Khởi động **Hình 13.1 (SGK Tin học 10 trang 69)** và nêu 02 câu hỏi yêu cầu học sinh trao đổi theo cặp trong 02 phút:\n'
        '  + *Câu hỏi 1:* Em hãy quan sát Hình 13.1 và kể tên các đối tượng hình học tạo nên miếng dưa hấu (vỏ xanh, cùi trắng, ruột đỏ và các hạt dưa).\n'
        '  + *Câu hỏi 2:* Xác định thứ tự các lớp từ dưới lên trên của các đối tượng trong Hình 13.1. Làm thế nào từ công cụ vẽ Hình tròn của Bài 12 lại vẽ được hình nửa vòng tròn (hình quạt/cung tròn) như miếng dưa hấu?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh quan sát kỹ Hình 13.1 trong Sách giáo khoa trang 69, thảo luận cặp đôi để phân tích các lớp đối tượng.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh phát biểu tại chỗ; học sinh khác nhận xét, bổ sung.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá theo bảng kiểm `CCĐG 1` và dẫn dắt vào bài mới: *"Để vẽ được miếng dưa hấu Hình 13.1 hay các biểu trưng, hình cây, hình cầu 3D phức tạp, chúng ta cần biết cách tuỳ chỉnh thuộc tính đặc trưng của từng hình khối, sử dụng hộp thoại tô màu nâng cao `Fill and Stroke` và thực hiện các phép ghép hình trong bảng chọn `Path`. Đó chính là nội dung của Bài 13."*',
        '**Phân tích Hình 13.1 – Miếng dưa hấu (SGK trang 69):**\n'
        '- **Các đối tượng cấu thành:**\n'
        '  1. Một nửa hình tròn lớn màu xanh lá đậm làm **vỏ dưa**;\n'
        '  2. Một nửa hình tròn nhỡ màu trắng/xanh nhạt làm **cùi dưa**;\n'
        '  3. Một nửa hình tròn nhỏ hơn màu đỏ làm **ruột dưa**;\n'
        '  4. Các hình giọt nước/elip nhỏ màu đen làm **hạt dưa**.\n'
        '- **Thứ tự các lớp (từ lớp dưới lên lớp trên):**\n'
        '  *Lớp 1 (dưới cùng):* Vỏ dưa màu xanh -> *Lớp 2:* Cùi dưa màu trắng -> *Lớp 3:* Ruột dưa màu đỏ -> *Lớp 4 (trên cùng):* Các hạt dưa màu đen.'
    ]
]

NV1_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên phát **Phiếu học tập số 1**, yêu cầu học sinh quan sát **Hình 13.2 (SGK trang 69)** và nghiên cứu **Bảng các thuộc tính cơ bản của một số hình có sẵn (SGK trang 69–70)** trong 05 phút để trả lời:\n'
        '  + *Câu 1 (Hoạt động 1 SGK trang 69):* Bốn hình màu đỏ trong Hình 13.2 được vẽ từ cùng một công cụ nào trong Hộp công cụ của Inkscape? Nhờ thay đổi những thông số nào mà tạo ra được 4 hình dạng khác nhau đó?\n'
        '  + *Câu 2:* Hoàn thành bảng giải thích ý nghĩa các thuộc tính: `W, H`, `Rx, Ry` (của hình vuông/chữ nhật); `Rx, Ry`, `Start, End` (của hình tròn/elip); `Corners`, `Rounded`, `Spoke ratio`, `Randomized` (của hình đa giác/hình sao).\n'
        '  + *Câu 3 (Câu hỏi củng cố SGK trang 70):* Nếu trên hình vẽ có sẵn một hình sao 5 cánh nhọn, em cần thay đổi tham số nào để các đỉnh ngôi sao trở nên cong tròn?\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đọc SGK trang 69–70, thảo luận cặp đôi và hoàn thành Phiếu học tập số 1.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Đại diện 02 nhóm trình bày kết quả; giáo viên minh hoạ trực tiếp việc thay đổi `Start = 0, End = 180` (tạo nửa hình tròn làm miếng dưa hấu) và tăng `Rounded` trên màn chiếu.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo `CCĐG 2` và chốt bảng thuộc tính chuẩn.',
        '**1. CÁC ĐỐI TƯỢNG HÌNH KHỐI VÀ THUỘC TÍNH CƠ BẢN (SGK trang 69–70)**\n'
        '- **Giải Hoạt động 1 – Hình 13.2 (trang 69):** Cả 4 hình đều được tạo từ công cụ **Hình đa giác / Hình sao (`*`)** bằng cách chuyển đổi chế độ sao/đa giác và điều chỉnh các thuộc tính `Corners`, `Spoke ratio`, `Rounded` trên Thanh điều khiển thuộc tính.\n'
        '- **Bảng các thuộc tính cơ bản của hình khối (SGK trang 69–70):**\n'
        '  1. **Hình vuông, hình chữ nhật (`R`):**\n'
        '     + `W, H`: Chiều rộng (`Width`) và chiều cao/dài (`Height`) của hình.\n'
        '     + `Rx, Ry`: Bán kính bo góc theo phương ngang và phương thẳng đứng (bằng `0` nếu góc vuông).\n'
        '  2. **Hình tròn, hình elip (`E`):**\n'
        '     + `Rx, Ry`: Bán kính theo phương ngang và phương thẳng đứng.\n'
        '     + `Start, End`: Góc của điểm đầu và điểm cuối (đơn vị: độ) khi vẽ hình quạt (`Slice`), cung tròn (`Arc`) hoặc hình tròn khép kín.\n'
        '  3. **Hình đa giác, hình sao (`*`):**\n'
        '     + `Corners`: Số đỉnh của đa giác hoặc số cánh của hình sao.\n'
        '     + `Rounded`: Độ cong (làm tròn) tại các đỉnh của hình.\n'
        '     + `Spoke ratio`: Tỉ lệ bán kính từ tâm đến góc trong và từ tâm tới đỉnh đầu hình sao.\n'
        '     + `Randomized`: Tham số làm biến dạng/méo hình ngẫu nhiên.\n'
        '- **Trả lời câu hỏi SGK trang 70:** Để các đỉnh của ngôi sao 5 cánh nhọn trở nên cong, ta thay đổi tham số **`Rounded`** (độ cong tại các đỉnh).'
    ]
]

NV2_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh quan sát **Hình 13.3**, **Hình 13.4 (SGK trang 70)** và đọc mục **2. Thiết lập màu tô, màu vẽ và tô màu cho đối tượng (SGK trang 70–71)** để hoàn thành **Phiếu học tập số 2** (06 phút):\n'
        '  + *Câu 1 (Hoạt động 2 SGK trang 70):* Quan sát 5 hình vuông trong Hình 13.3 và chỉ ra sự khác nhau về màu tô (`Fill`) và màu vẽ đường viền (`Stroke`) của từng hình.\n'
        '  + *Câu 2:* Nêu cách mở hộp thoại **`Fill and Stroke`** (`Object / Fill and Stroke` hoặc `Shift + Ctrl + F`), chức năng của 3 trang **`Fill`**, **`Stroke paint`**, **`Stroke style`** và ý nghĩa của các kiểu tô màu, thông số **`Opacity`** / **`Alpha (A)`**.\n'
        '  + *Câu 3:* Trả lời 02 câu hỏi trắc nghiệm củng cố ở giữa trang 71 SGK.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh đối chiếu Hình 13.3, Hình 13.4 và trang 70–71 SGK để điền vào Phiếu học tập số 2.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Đại diện học sinh trình bày sự khác nhau của 5 hình trong Hình 13.3, 3 bước chỉnh sửa nền/đường nét và đáp án 2 câu hỏi củng cố trang 71.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo `CCĐG 3` và minh hoạ trực tiếp thao tác tô màu chuyển (`Linear/Radial gradient`) cùng độ trong suốt (`Opacity`) trên hộp thoại `Fill and Stroke`.',
        '**2. THIẾT LẬP MÀU TÔ, MÀU VẼ VÀ HỘP THOẠI FILL AND STROKE (SGK trang 70–71)**\n'
        '- **Nhận xét Hình 13.3 (SGK trang 70):** Cả 5 hình vuông đều có viền màu đỏ cam nhưng khác nhau về kiểu màu tô bên trong: (1) Không màu (trong suốt); (2) Màu xanh lá đồng nhất (`Flat color`); (3) Màu chuyển tuyến tính từ xanh sang trắng (`Linear gradient`); (4) Màu chuyển toả tròn từ tâm (`Radial gradient`); (5) Màu chuyển có độ trong suốt (`Opacity`).\n'
        '- **Hộp thoại `Fill and Stroke` (Hình 13.4 – Mở bằng lệnh `Object / Fill and Stroke` hoặc `Shift + Ctrl + F`):**\n'
        '  + **Trang `Fill` (Màu tô):** Thiết lập kiểu tô cho phần nền bên trong đối tượng.\n'
        '  + **Trang `Stroke paint` (Màu vẽ):** Thiết lập màu cho đường viền của đối tượng.\n'
        '  + **Trang `Stroke style` (Kiểu nét vẽ):** Thay đổi độ dày mỏng (`Width`) và kiểu đường viền (nét liền, nét đứt, chấm gạch...).\n'
        '- **Các tuỳ chọn kiểu tô màu:** `X` (Không màu/trong suốt); Màu đồng nhất (`Flat color`); Màu chuyển tuyến tính (`Linear gradient`); Màu chuyển từ tâm (`Radial gradient`); Hoa văn (`Pattern`); `?` (Huỷ đặt).\n'
        '- **Độ trong suốt:** Giá trị **`Opacity`** (tính theo `%`) và kênh **`A (Alpha)`** xác định độ trong suốt của đối tượng/màu tô.\n'
        '- **Trả lời Câu hỏi củng cố SGK trang 71:**\n'
        '  + *Câu 1:* Để xác định đường viền của đối tượng dạng nét đứt, chọn trang **`C. Stroke style`**.\n'
        '  + *Câu 2:* Để chỉnh thông số điểm dừng màu của `gradient`, chọn biểu tượng cây bút chỉnh sửa gradient **`B`**.'
    ]
]

NV3_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên chiếu **Hình 13.5 (SGK trang 71)**, **Hình 13.6** và **Hình 13.7 (SGK trang 72)**, yêu cầu học sinh nghiên cứu mục **3. Các phép ghép đối tượng đồ hoạ (SGK trang 71–72)** để hoàn thành **Phiếu học tập số 3** (07 phút):\n'
        '  + *Câu 1 (Hoạt động 3 SGK trang 71):* Làm thế nào để ghép 2 hình tròn đỏ và 1 hình vuông đỏ trong Hình 13.5 thành một hình trái tim hoàn chỉnh?\n'
        '  + *Câu 2:* Lập bảng tổng hợp tên tiếng Anh, tổ hợp phím tắt và kết quả thu được của **06 phép ghép hình** trong bảng chọn **`Path`** (theo Hình 13.6 trang 72).\n'
        '  + *Câu 3 (Câu hỏi SGK trang 72):* Em hãy nêu phép ghép hình và các bước thực hiện để vẽ đám mây như **Hình 13.7 (SGK trang 72)**.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh quan sát kỹ Hình 13.6 (đối chiếu hình vuông ở lớp dưới và hình tròn ở lớp trên) để phân biệt 6 phép ghép trong bảng chọn `Path`.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Đại diện học sinh trình bày bảng 6 phép ghép hình và quy trình vẽ đám mây Hình 13.7; các nhóm khác nhận xét.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét theo `CCĐG 4` và lưu ý học sinh quy tắc cực kì quan trọng: Trong phép hiệu (`Difference`), phép chia (`Division`) và phép cắt (`Cut Path`), kết quả luôn lấy **hình ở lớp dưới** làm gốc bị cắt bởi **hình ở lớp trên**.',
        '**3. CÁC PHÉP GHÉP ĐỐI TƯỢNG ĐỒ HOẠ TRONG BẢNG CHỌN PATH (SGK trang 71–72)**\n'
        '- **Giải Hoạt động 3 – Hình 13.5 (trang 71):** Xoay hình vuông một góc `45°` thành hình thoi, đặt 2 hình tròn chồng lên hai cạnh trên của hình vuông rồi thực hiện phép hợp (`Path / Union`) để thu được hình trái tim.\n'
        '- **Sáu phép ghép đối tượng trong bảng chọn `Path` (Hình 13.6, trang 71–72):**\n'
        '  1. **Phép hợp (`Path / Union`, phím tắt `Ctrl + +`):** Lấy tất cả các phần thuộc một trong các hình đơn (gộp thành một đối tượng duy nhất).\n'
        '  2. **Phép hiệu (`Path / Difference`, phím tắt `Ctrl + -`):** Lấy phần thuộc **hình lớp dưới** nhưng không thuộc hình lớp trên.\n'
        '  3. **Phép giao (`Path / Intersection`, phím tắt `Ctrl + *`):** Chỉ lấy phần giao nhau thuộc cả hai hình được chọn.\n'
        '  4. **Phép hiệu đối xứng (`Path / Exclusion`, phím tắt `Ctrl + ^`):** Lấy phần hình thuộc các hình trừ đi phần giao nhau ở giữa.\n'
        '  5. **Phép chia (`Path / Division`, phím tắt `Ctrl + /`):** Hình lớp dưới được chia cắt thành các mảnh rời nhau bởi đường nét của hình lớp trên.\n'
        '  6. **Phép cắt (`Path / Cut Path`, phím tắt `Ctrl + Alt + /`):** Cắt đường viền của hình lớp dưới thành các đoạn rời nhau tại giao điểm với hình lớp trên (kết quả không có màu tô nền).\n'
        '- **Trả lời câu hỏi SGK trang 72 (Vẽ đám mây Hình 13.7):**\n'
        '  + *Phép ghép sử dụng:* **Phép hợp (`Path / Union`, tổ hợp phím `Ctrl + +`)**.\n'
        '  + *Các bước thực hiện:* (1) Dùng công cụ Hình tròn/elip vẽ 5–6 hình tròn có kích thước khác nhau xếp chồng lấp lên nhau tạo dáng đám mây; (2) Dùng công cụ `Select` quét chọn tất cả các hình tròn đó; (3) Chọn lệnh **`Path / Union`** (hoặc nhấn `Ctrl + +`) để hợp nhất thành hình đám mây chỉ có một đường viền bao ngoài như Hình 13.7.'
    ]
]

NV4_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh quan sát mục **THỰC HÀNH (SGK trang 72–74)** gồm 3 nhiệm vụ chuẩn mực (**Hình 13.8–13.9: Vẽ logo**, **Hình 13.10–13.11: Vẽ hình cây**, **Hình 13.12–13.13: Vẽ hình cầu**) và thực hành trên máy tính lưu vào tệp `HoTen_10-3_Bai13.svg`:\n'
        '  + *Nhiệm vụ thực hành 1 (SGK trang 72–73 – Hình 13.8 & 13.9):* Vẽ biểu tượng (logo) 3 cánh cung tròn bằng cách tạo cung tròn hở, chỉnh độ dày nét vẽ (`Width`) trong `Stroke style`, nhân bản (`Ctrl + D`), xoay và ghép 3 cung tròn.\n'
        '  + *Nhiệm vụ thực hành 2 (SGK trang 73 – Hình 13.10 & 13.11):* Vẽ hình cây xanh bằng cách ghép 3 hình tròn làm tán lá, dùng phép hiệu (`Path / Difference`) giữa hình chữ nhật và 2 hình tròn để tạo gốc cây uốn cong, rồi dùng phép hợp (`Path / Union`) ghép tán lá với gốc cây và tô màu xanh lá.\n'
        '  + *Nhiệm vụ thực hành 3 (SGK trang 73–74 – Hình 13.12 & 13.13):* Vẽ hình cầu 3D có bóng sáng bằng kết hợp hình chữ nhật nền vàng, hình tròn lớn tô màu chuyển tuyến tính (`Linear gradient`) và hình tròn nhỏ màu trắng tô màu chuyển toả tròn (`Radial gradient`).\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thực hành tuần tự cả 3 nhiệm vụ trên phần mềm `Inkscape`. Giáo viên đi từng dãy máy hướng dẫn kỹ thuật dùng `Path / Difference` từng cặp 2 đối tượng khi tạo thân cây (Hình 13.11c) và cách kéo cần điều khiển `Gradient` (Hình 13.13b).\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên chiếu 03 bài thực hành hoàn chỉnh của học sinh lên màn chiếu; học sinh giải thích tại sao ở Nhiệm vụ 2 phải thực hiện `Path / Difference` hai lần trước khi `Path / Union`.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên chấm điểm sản phẩm theo Rubric `CCĐG 6` (Mục IV. Phụ lục 2).',
        '**4. HƯỚNG DẪN CHI TIẾT 03 NHIỆM VỤ THỰC HÀNH (SGK trang 72–74)**\n'
        '**a) Nhiệm vụ 1: Vẽ logo (Hình 13.8 & Hình 13.9, trang 72–73):**\n'
        '- *Bước 1–3:* Dùng công cụ Hình tròn vẽ hình tròn (H.13.9a) -> kéo nút tròn điều khiển trên hình để mở góc cung tròn (H.13.9b) -> trên Thanh điều khiển thuộc tính chọn biểu tượng **Cung tròn hở (`Arc`)** để chỉ giữ lại đường viền không khép kín (H.13.9c).\n'
        '- *Bước 4:* Mở hộp thoại **`Fill and Stroke` (`Shift + Ctrl + F`)** -> chọn trang **`Stroke style`** -> tăng giá trị ô **`Width`** để nét cung tròn dày lên (H.13.9d).\n'
        '- *Bước 5–6:* Nhấn **`Ctrl + D`** tạo 2 bản sao, dùng công cụ `Select` xoay và di chuyển 3 cung tròn ghép khít tạo thành logo 3 cánh rồi đổi màu vẽ (`Stroke paint`) màu đỏ đậm như Hình 13.8 (H.13.9e).\n'
        '**b) Nhiệm vụ 2: Vẽ hình cây (Hình 13.10 & Hình 13.11, trang 73):**\n'
        '- *Bước 1 (Tán lá):* Vẽ 01 vòng tròn, sao chép (`Ctrl + D`) thành 3 hình tròn xếp chụm vào nhau như Hình 13.11a.\n'
        '- *Bước 2–3 (Tạo thân cây bằng phép hiệu):* Vẽ 01 hình chữ nhật ở lớp dưới và 02 hình tròn ép hai bên ở lớp trên (H.13.11b). Chọn hình chữ nhật và hình tròn bên trái -> chọn **`Path / Difference` (`Ctrl + -`)**; tiếp tục chọn kết quả vừa tạo với hình tròn bên phải -> chọn **`Path / Difference`** lần 2 để thu được thân cây thắt eo (H.13.11c).\n'
        '- *Bước 4–5 (Hợp nhất thành cây):* Di chuyển thân cây vào dưới 3 hình tròn tán lá (H.13.11d) -> quét chọn cả 4 phần và chọn lệnh **`Path / Union` (`Ctrl + +`)** (H.13.11e) -> tô màu xanh lá hoàn chỉnh (H.13.11f).\n'
        '**c) Nhiệm vụ 3: Vẽ hình cầu 3D (Hình 13.12 & Hình 13.13, trang 73–74):**\n'
        '- *Bước 1:* Vẽ các hình có màu vẽ viền trong suốt (`Stroke = X`): hình chữ nhật nền vàng, hình tròn to màu đen và hình tròn nhỏ màu trắng đặt lệch góc trên phải (H.13.13a).\n'
        '- *Bước 2:* Chọn hình tròn đen -> trong hộp thoại `Fill and Stroke` chọn **`Linear gradient`** và công cụ chỉnh sửa gradient, kéo trục chuyển màu từ trái trên (sáng) xuống phải dưới (tối) (H.13.13b).\n'
        '- *Bước 3–4:* Chọn hình tròn trắng -> chọn **`Radial gradient`** (màu chuyển từ tâm – H.13.13c), tăng kích thước và căn chỉnh để tạo vầng sáng phản chiếu 3D tự nhiên cho hình cầu (H.13.13d).'
    ]
]

HD3_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên yêu cầu học sinh hoàn thành **Phiếu học tập số 4** và thực hành các bài tập phần **LUYỆN TẬP (SGK trang 74 – Hình 13.14, 13.15, 13.16)**:\n'
        '  + *Bài Luyện tập 1 (Hình 13.14 – Vẽ bông hoa cách điệu):* Vẽ ngôi sao nhiều cánh, tuỳ chỉnh thuộc tính **`Corners`**, **`Rounded`**, **`Spoke ratio`** để các cánh hoa uốn tròn mềm mại rồi đặt hình tròn màu vàng ở lớp trên làm nhụy hoa.\n'
        '  + *Bài Luyện tập 2 (Hình 13.15 – Biểu tượng quân Bích đỏ):* Kết hợp phép ghép hình trái tim ngược (từ 2 hình tròn + 1 hình thoi như Hình 13.5) với phần chân đế (tạo bằng `Path / Difference` như thân cây Hình 13.11c) rồi hợp nhất bằng `Path / Union`.\n'
        '  + *Bài Luyện tập 3 (Hình 13.16 – Chùm bóng bay):* Dùng hình elip có độ trong suốt (`Opacity` khoảng `75%–85%`), cắt tạo vệt sáng và dùng công cụ Bút vẽ tự do (`Freehand lines`) vẽ dây bóng.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh điền phương án thiết lập thuộc tính/phép ghép vào Phiếu học tập số 4 và thực hành vẽ Hình 13.14 (và Hình 13.15 hoặc 13.16) vào tệp `HoTen_10-3_Bai13.svg`.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên mời 02 học sinh trình bày thông số `Corners = 10`, `Rounded ≈ 0.25`, `Spoke ratio ≈ 0.45` khi vẽ bông hoa Hình 13.14 và cách ghép Hình 13.15.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên nhận xét và chấm điểm theo `CCĐG 5` và `CCĐG 6`.',
        '**HƯỚNG DẪN GIẢI PHẦN LUYỆN TẬP (SGK BÀI 13, TRANG 74):**\n'
        '- **Bài 1 – Vẽ bông hoa Hình 13.14:**\n'
        '  + Dùng công cụ **Hình sao (`*`)**, đặt **`Corners = 10`** (10 cánh), tăng **`Rounded ≈ 0.2 – 0.3`** để bo tròn đầu cánh hoa, đặt **`Spoke ratio ≈ 0.45`**; tô nền màu đỏ và viền đen đậm.\n'
        '  + Dùng công cụ **Hình tròn (`E`)** vẽ nhụy hoa màu vàng, viền đỏ đặt ở lớp trên chính giữa ngôi sao.\n'
        '- **Bài 2 – Vẽ Hình 13.15:**\n'
        '  + *Phần tán trên:* Ghép 1 hình vuông xoay `45°` và 2 hình tròn ở hai cạnh dưới bằng **`Path / Union`** (tạo hình trái tim lật ngược).\n'
        '  + *Phần chân đế:* Vẽ hình chữ nhật bị khoét hai bên bởi 2 hình tròn bằng **`Path / Difference`** (tương tự bước vẽ thân cây Hình 13.11c), sau đó gộp với phần trên bằng **`Path / Union`** và tô màu đỏ.\n'
        '- **Bài 3 – Vẽ chùm bóng Hình 13.16:**\n'
        '  + Vẽ 3 quả bóng bằng hình elip tô màu Xanh, Cam, Vàng với độ trong suốt **`Opacity ≈ 80%`** để thấy phần giao nhau; thêm vệt sáng trắng cong ở góc bóng và dùng công cụ **Bút vẽ (`F6`/`P`)** vẽ 3 sợi dây.'
    ]
]

HD4_ROWS_BAI13 = [
    ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
    [
        '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
        '- Giáo viên giao nhiệm vụ phần **VẬN DỤNG (SGK trang 74)**:\n'
        '  + *Bài Vận dụng 1 (SGK trang 74):* Vận dụng thuộc tính góc `Start = 0, End = 180` của công cụ Hình tròn (chế độ hình quạt `Slice`) kết hợp nhân bản `Duplicate` để hoàn thiện hình vẽ **Miếng dưa hấu (Hình 13.1 trang 69)**.\n'
        '  + *Bài Vận dụng 2 (SGK trang 74):* Thiết kế một hình minh hoạ hệ Mặt Trời / các hành tinh trong không gian bằng hiệu ứng hình cầu `Gradient` đã học ở Nhiệm vụ 3.\n'
        '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
        '- Học sinh thực hành vẽ nhanh Miếng dưa hấu Hình 13.1 trên máy tính và phác thảo ý tưởng Bài Vận dụng 2.\n'
        '**Bước 3. Báo cáo sản phẩm**\n'
        '- Giáo viên chiếu sản phẩm Miếng dưa hấu của học sinh lên màn chiếu, tổng kết toàn bộ chuỗi kiến thức từ Khởi động (Hình 13.1) đến Vận dụng.\n'
        '**Bước 4. Đánh giá, nhận xét**\n'
        '- Giáo viên đánh giá theo Rubric `CCĐG 6` và dặn dò học sinh chuẩn bị **Bài 14: Làm việc với các đối tượng đường và văn bản (`Tiết PPCT: 21, 22`)**.',
        '**HƯỚNG DẪN GIẢI PHẦN VẬN DỤNG (SGK BÀI 13, TRANG 74):**\n'
        '- **Bài Vận dụng 1 – Quy trình vẽ Miếng dưa hấu Hình 13.1:**\n'
        '  + *Bước 1:* Chọn công cụ **Hình tròn (`E`)**, trên Thanh điều khiển thuộc tính đặt **`Start = 0.0`**, **`End = 180.0`** và chọn chế độ khép kín hình quạt (`Slice`) để tạo nửa hình tròn.\n'
        '  + *Bước 2:* Tô màu xanh lá đậm làm **vỏ dưa** (lớp 1) -> nhấn `Ctrl + D` nhân bản, thu nhỏ nhẹ và tô màu xanh nhạt/trắng làm **cùi dưa** (lớp 2) -> nhấn `Ctrl + D` nhân bản lần nữa, thu nhỏ và tô màu đỏ tươi làm **ruột dưa** (lớp 3).\n'
        '  + *Bước 3:* Vẽ 01 hình elip nhỏ màu đen làm hạt dưa, dùng `Ctrl + D` nhân bản và xếp rải đều trên nền ruột đỏ (lớp 4 trên cùng).\n'
        '- **Bài Vận dụng 2 – Vẽ hình trong không gian:**\n'
        '  + Vẽ nền chữ nhật màu xanh đen vũ trụ, thêm các ngôi sao nhỏ (`Corners = 5`) và các hành tinh hình cầu 3D tô màu chuyển `Linear / Radial gradient`.'
    ]
]

def add_bai13_appendix(b, class_label='10/3', ppct_label='19, 20'):
    br(b)
    add(b, 'IV. PHỤ LỤC (HỒ SƠ HỌC LIỆU, PHIẾU HỌC TẬP VÀ CÔNG CỤ ĐÁNH GIÁ)', 'heading')
    add(b, 'Phụ lục 1. Hệ thống Phiếu học tập số 1, 2, 3, 4 và Công cụ đánh giá CCĐG 1 – CCĐG 5', 'sub')
    pht_rows = [
        ['Tên Phiếu / Công cụ', 'Nội dung câu hỏi và yêu cầu (Dựa sát SGK Tin học 10 KNTT, Bài 13 trang 69–74)', 'Đáp án chuẩn & Thang điểm chấm'],
        ['**CCĐG 1**\n(Hoạt động 1:\nMở đầu)',
         '- Quan sát Hình 13.1 (Miếng dưa hấu – tr. 69): Kể tên các đối tượng hình học và xác định thứ tự 4 lớp từ dưới lên trên.',
         '- Nêu đúng 3 nửa hình tròn đồng tâm (vỏ xanh, cùi trắng, ruột đỏ) và các hạt đen ở lớp trên cùng *(Đạt / Chưa đạt)*.'],
        ['**Phiếu học tập số 1\n& CCĐG 2**\n(Hoạt động 2 –\nNhiệm vụ 1)',
         '1. Điền ý nghĩa các thuộc tính vào bảng (SGK tr. 69–70): `W, H`; `Rx, Ry`; `Start, End`; `Corners`; `Rounded`; `Spoke ratio`; `Randomized`.\n'
         '2. Trả lời câu hỏi SGK tr. 70: Thay đổi tham số nào để các đỉnh ngôi sao 5 cánh nhọn trở nên cong?',
         '- Giải thích đúng ý nghĩa 7 nhóm thuộc tính của 3 loại hình khối *(7,0 điểm)*.\n'
         '- Trả lời đúng câu hỏi tr. 70: Tham số **`Rounded`** *(3,0 điểm)*.'],
        ['**Phiếu học tập số 2\n& CCĐG 3**\n(Hoạt động 2 –\nNhiệm vụ 2)',
         '1. Nêu chức năng 3 trang `Fill`, `Stroke paint`, `Stroke style` trong hộp thoại `Fill and Stroke` (Hình 13.4 tr. 70) và ý nghĩa của `Opacity`.\n'
         '2. Khoanh đáp án 2 câu hỏi trắc nghiệm củng cố trang 71 SGK.',
         '- Nêu đúng chức năng 3 trang và độ trong suốt `Opacity` *(6,0 điểm)*.\n'
         '- Câu 1 tr. 71: **`C. Stroke style`**; Câu 2 tr. 71: **`B`** (cây bút chỉnh gradient) *(4,0 điểm)*.'],
        ['**Phiếu học tập số 3\n& CCĐG 4**\n(Hoạt động 2 –\nNhiệm vụ 3)',
         '1. Hoàn thành bảng 6 phép ghép đối tượng trong bảng chọn `Path` (Hình 13.6 tr. 72): `Union`, `Difference`, `Intersection`, `Exclusion`, `Division`, `Cut Path` (Phím tắt & Kết quả).\n'
         '2. Trả lời câu hỏi SGK tr. 72: Nêu phép ghép và các bước vẽ đám mây Hình 13.7.',
         '- Điền đúng phím tắt và quy tắc của 6 phép ghép hình *(7,0 điểm)*.\n'
         '- Vẽ đám mây Hình 13.7: Vẽ các hình tròn chồng lấp rồi dùng **`Path / Union (Ctrl + +)`** *(3,0 điểm)*.'],
        ['**Phiếu học tập số 4\n& CCĐG 5**\n(Hoạt động 3:\nLuyện tập)',
         '- Nêu thông số thuộc tính và phép ghép cần dùng để vẽ: (1) Bông hoa Hình 13.14; (2) Hình 13.15; (3) Chùm bóng Hình 13.16 (SGK tr. 74).',
         '- Hình 13.14: Hình sao (`Corners = 10, Rounded ≈ 0.25`) + Hình tròn; Hình 13.15: `Union` + `Difference`; Hình 13.16: `Opacity` + Bút vẽ *(10,0 điểm)*.']
    ]
    b.append(table(pht_rows, [1600, 4200, 3200], header=True, cell_sz=21))

    add(b, 'Phụ lục 2. Công cụ đánh giá CCĐG 6 – Rubric 10 điểm chấm sản phẩm thực hành Bài 13 (Hình 13.8–13.16 & Vận dụng)', 'sub')
    rubric_info = [
        ['THÔNG TIN HÀNH CHÍNH PHIẾU ĐÁNH GIÁ SẢN PHẨM THỰC HÀNH (CCĐG 6)', 'NỘI DUNG CHI TIẾT'],
        ['Tên bài dạy & Tiết PPCT', f'Bài 13. Bổ sung các đối tượng đồ hoạ (Tiết PPCT: {ppct_label})'],
        ['Lớp giảng dạy & Địa điểm', f'Lớp {class_label} – Trường THPT Thuận Hoá (Phòng H.I.5)'],
        ['Giáo viên giảng dạy / GVHD', 'SVTT: Hồ Quang Huy – Giáo viên hướng dẫn: Thầy Mai Xuân Văn'],
        ['Họ và tên học sinh / Nhóm máy', '........................................................................ Máy số: ...................']
    ]
    b.append(table(rubric_info, [3800, 5200], header=True, cell_sz=22))
    add(b, '', first=0)

    rubric_rows = [
        ['STT', 'Tiêu chí đánh giá', 'Thuộc tính, Hộp thoại & Phép ghép bắt buộc kiểm tra', 'Mô tả 3 mức độ đạt (Mức 3: Tốt / Mức 2: Đạt / Mức 1: Chưa đạt)', 'Điểm tối đa', 'Điểm đạt'],
        ['1', 'Nhiệm vụ 1:\nVẽ biểu tượng Logo\n(Hình 13.8 & 13.9 tr. 72–73)', 'Công cụ Hình tròn chế độ cung tròn hở (`Arc`), `Fill and Stroke -> Stroke style -> Width`, `Ctrl + D`, xoay ghép 3 cánh.', '- **Mức 3 (2,5 đ):** Tạo đúng cung tròn hở, tăng độ dày nét vẽ (`Width`), nhân bản và xoay ghép khít 3 cánh logo màu đỏ đậm như Hình 13.8.\n- **Mức 2 (1,75 đ):** Vẽ được 3 cung tròn dày nhưng chỗ tiếp giáp giữa 3 cánh còn lệch nhẹ.\n- **Mức 1 (0,75 đ):** Chưa chuyển sang chế độ cung tròn hở (`Arc`) hoặc chưa tăng được độ dày nét.', '2,5 điểm', '......'],
        ['2', 'Nhiệm vụ 2:\nVẽ hình cây\n(Hình 13.10 & 13.11 tr. 73)', 'Phép hiệu `Path / Difference (Ctrl + -)` tạo thân cây và phép hợp `Path / Union (Ctrl + +)` gộp tán lá.', '- **Mức 3 (3,0 đ):** Thực hiện đúng 2 lần `Path / Difference` tạo thân cây cong đều, dùng `Path / Union` hợp nhất với 3 hình tròn tán lá và tô màu xanh lá chuẩn Hình 13.10.\n- **Mức 2 (2,0 đ):** Tạo được tán lá và thân cây nhưng thao tác `Difference` còn lệch tỉ lệ.\n- **Mức 1 (1,0 đ):** Chỉ xếp chồng các hình mà chưa thực hiện phép ghép trong bảng chọn `Path`.', '3,0 điểm', '......'],
        ['3', 'Nhiệm vụ 3:\nVẽ hình cầu 3D\n(Hình 13.12 & 13.13 tr. 73–74)', 'Bỏ viền (`Stroke = X`), tô màu chuyển `Linear gradient` và `Radial gradient` trong `Fill and Stroke`.', '- **Mức 3 (2,5 đ):** Phối hợp mượt mà `Linear gradient` trên hình tròn lớn và `Radial gradient` trên hình tròn trắng tạo hiệu ứng khối cầu 3D chân thực như Hình 13.12.\n- **Mức 2 (1,75 đ):** Áp dụng được `Gradient` nhưng hướng chuyển màu chưa tạo hiệu ứng bóng sáng rõ.\n- **Mức 1 (0,75 đ):** Chỉ tô màu đồng nhất (`Flat color`), chưa dùng `Gradient`.', '2,5 điểm', '......'],
        ['4', 'Luyện tập & Vận dụng:\nHình 13.14–13.16 & Miếng dưa hấu Hình 13.1 (tr. 74)', 'Thuộc tính `Rounded`, `Start = 0, End = 180`, `Opacity`, lưu tệp `HoTen_10-3_Bai13.svg`.', '- **Mức 3 (2,0 đ):** Hoàn thành thêm ít nhất 01 hình Luyện tập (Hình 13.14/13.15/13.16) hoặc Miếng dưa hấu Hình 13.1 và lưu tệp `.svg` đúng chuẩn.\n- **Mức 2 (1,25 đ):** Vẽ được phần khung hình Luyện tập/Vận dụng nhưng chưa tinh chỉnh xong màu sắc.\n- **Mức 1 (0,5 đ):** Chưa thực hiện bài tập Luyện tập/Vận dụng.', '2,0 điểm', '......'],
        ['', '**TỔNG ĐIỂM**', '', '**Xếp loại:** Đạt từ `8,0 – 10,0`: Tốt | `5,0 – 7,9`: Đạt | `< 5,0`: Cần cố gắng', '**10,0 điểm**', '......']
    ]
    b.append(table(rubric_rows, [650, 1500, 2000, 3450, 700, 700], header=True, cell_sz=21))

def build_bai13_tiet1():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 - LÍ THUYẾT – TIẾT PPCT: 19)',
        class_str='Lớp 10/3 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45) – Phòng H.I.5',
        lesson_row_str='Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT) – Tiết PPCT: 19 (Tổng bài: 02 tiết – PPCT: 19, 20)',
        page1_title='TÊN BÀI DẠY: BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 1 – LÍ THUYẾT)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 19, 20] – Kế hoạch này chi tiết hoá Tiết thứ 1 của bài / Tiết PPCT: 19)',
        page1_class_line='Môn học: Tin học; Lớp: 10/3',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 1 - LT / Tiết PPCT: 19)',
        page1_sched_line='Lịch dự giờ: Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45) – Lớp 10/3 – Phòng H.I.5',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp minh hoạ trực quan và thảo luận nhóm'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E (Ứng dụng tin học – Phần mềm thiết kế đồ hoạ) và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 19 – Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT)):')
    add(b, '1. Về kiến thức', 'sub')
    add(b, 'Sau tiết học này, học sinh đạt được các yêu cầu về kiến thức (ký hiệu `[KT1]` – `[KT3]` để đối chiếu xuyên suốt các hoạt động dạy học):', keep=True)
    for item in [
        '**`[KT1]`**: Nêu được ý nghĩa và cách điều chỉnh các thuộc tính cơ bản của từng đối tượng hình khối trên Thanh điều khiển thuộc tính: `W, H, Rx, Ry` (hình vuông/chữ nhật); `Rx, Ry, Start, End` (hình tròn/elip); `Corners, Rounded, Spoke ratio, Randomized` (hình đa giác/hình sao) theo Bảng trang 69–70 SGK.',
        '**`[KT2]`**: Trình bày được chức năng của hộp thoại **`Fill and Stroke`** (Hình 13.4 SGK trang 70) gồm 3 trang `Fill` (màu tô), `Stroke paint` (màu vẽ), `Stroke style` (kiểu và độ dày nét vẽ), các kiểu tô màu chuyển (`Linear gradient`, `Radial gradient`) và độ trong suốt (`Opacity`).',
        '**`[KT3]`**: Phân biệt được 06 phép ghép đối tượng đồ hoạ trong bảng chọn **`Path`** (Hình 13.6 SGK trang 72): Phép hợp (`Union`), phép hiệu (`Difference`), phép giao (`Intersection`), phép hiệu đối xứng (`Exclusion`), phép chia (`Division`) và phép cắt (`Cut Path`).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**`[NLC1 – Tự chủ và tự học]`**: Học sinh chủ động đọc hiểu Sách giáo khoa trang 69–72, quan sát Hình 13.1 đến Hình 13.7 để hoàn thành các Phiếu học tập.',
        '**`[NLC2 – Giao tiếp và hợp tác]`**: Học sinh thảo luận nhóm đôi sôi nổi để phân tích thứ tự lớp và kết quả của 6 phép ghép hình.',
        '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Học sinh đề xuất được quy trình kết hợp các phép ghép (`Union`, `Difference`) và thuộc tính hình khối để tạo ra các hình phức tạp (trái tim, đám mây, hình cây, miếng dưa hấu).'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**`[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]`**: Nhận biết chính xác vị trí các ô thông số thuộc tính, hộp thoại `Fill and Stroke` và các lệnh trong bảng chọn `Path` của `Inkscape`.',
        '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông]`**: Phân tích được một đối tượng đồ hoạ phức tạp thành các hình khối cơ bản và lựa chọn đúng phép ghép tương ứng.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**`[NLS 3.2.NC1.a]`**: Biết cách tinh chỉnh, biến đổi và kết hợp các đối tượng đồ hoạ số để tạo ra hình vẽ mới.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**`[PC1 – Chăm chỉ]`**: Học sinh chú ý theo dõi bài giảng, ghi chép đầy đủ bảng thuộc tính và 6 phép ghép hình.',
        '**`[PC2 – Trách nhiệm]`**: Học sinh tích cực hoàn thành nhiệm vụ thảo luận cặp đôi trong giờ dự giờ.',
        '**`[PC3 – Trung thực]`**: Học sinh tự giác trả lời câu hỏi củng cố và đánh giá khách quan bài làm của bạn.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính giáo viên cài đặt `Inkscape`, máy chiếu tại phòng học `H.I.5`, bộ Phiếu học tập số 1, 2, 3, 4.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, **Bài 13: Bổ sung các đối tượng đồ hoạ** (trang 69–74).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 19 – Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 1 - LT)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 1 - Lí thuyết (45 phút – Tiết thứ 1 của bài / Tiết PPCT: 19)**', first=0, keep=True)
    ov_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng & Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Phân tích Hình 13.1 – Miếng dưa hấu)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 19)', '[KT1], [NLC1],\n[PC1]', 'Phân tích các đối tượng và thứ tự 4 lớp của miếng dưa hấu.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới\n- 2.1. Nhiệm vụ 1: Các đối tượng hình khối & thuộc tính (09p)\n- 2.2. Nhiệm vụ 2: Thiết lập màu tô, màu vẽ & Fill and Stroke (09p)\n- 2.3. Nhiệm vụ 3: Các phép ghép đối tượng đồ hoạ Path (10p)', '28 phút\n(Tiết 1 của bài /\nTiết PPCT: 19)', '[KT1], [KT2], [KT3],\n[NLC1], [NLC2], [NLC3],\n[NLa], [NLc],\n[NLS 3.2.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, số 2, số 3 đã hoàn thành.\nCông cụ đánh giá: CCĐG 2, CCĐG 3, CCĐG 4.'],
        ['3. Hoạt động 3: Luyện tập\n(Phân tích quy trình vẽ Hình 13.14, 13.15, 13.16 SGK tr. 74)', '07 phút\n(Tiết 1 của bài /\nTiết PPCT: 19)', '[KT1], [KT2], [KT3],\n[NLC3], [NLc], [PC3]', 'Phiếu học tập số 4 xác định thuộc tính và phép ghép cho 3 bài Luyện tập.\nCông cụ đánh giá: CCĐG 5.'],
        ['4. Hoạt động 4: Vận dụng\n(Thiết kế quy trình vẽ Miếng dưa hấu Hình 13.1 & chuẩn bị Tiết 2 TH)', '05 phút\n(Tiết 1 của bài /\nTiết PPCT: 19)', '[KT1], [KT2], [KT3],\n[NLC3], [NLc], [PC2]', 'Quy trình 3 bước vẽ Miếng dưa hấu bằng góc `Start = 0, End = 180` chuẩn bị cho Tiết 2 (Tiết PPCT: 20).']
    ]
    b.append(table(ov_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[NLC1]`, `[PC1]` – Học sinh nhận diện được các đối tượng hình học và thứ tự các lớp trong hình vẽ miếng dưa hấu (Hình 13.1), từ đó nảy sinh nhu cầu tìm hiểu thuộc tính hình khối và phép ghép hình.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.1 và Hình 13.2 (SGK trang 69)** để thảo luận câu hỏi khởi động:', 'sub')
    add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    add(b, 'c) Sản phẩm: Câu trả lời miệng của học sinh nêu đúng 4 lớp của miếng dưa hấu Hình 13.1.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 1` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD1_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (28 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu các đối tượng hình khối và thuộc tính cơ bản (09 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT1]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[PC1]` – Học sinh trình bày được ý nghĩa các thuộc tính đặc trưng của hình vuông/chữ nhật (`W, H, Rx, Ry`), hình tròn/elip (`Rx, Ry, Start, End`) và hình đa giác/ngôi sao (`Corners, Rounded, Spoke ratio, Randomized`).', 'sub')
    add(b, 'b) Nội dung: Học sinh nghiên cứu **Bảng các thuộc tính cơ bản của một số hình có sẵn (SGK trang 69–70)** và hoàn thành **Phiếu học tập số 1**:', 'sub')
    add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 1 hoàn chỉnh.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 2` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV1_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu cách thiết lập màu tô, màu vẽ và hộp thoại Fill and Stroke (09 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT2]`, `[NLC1]`, `[NLa]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh phân biệt được các kiểu tô màu (đồng nhất, chuyển màu tuyến tính/toả tròn, hoa văn), độ trong suốt `Opacity` và chức năng 3 trang của hộp thoại `Fill and Stroke`.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.3 và Hình 13.4 (SGK trang 70)**, đọc mục 2 (trang 70–71) và hoàn thành **Phiếu học tập số 2**:', 'sub')
    add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 2 hoàn chỉnh và trả lời đúng 2 câu hỏi củng cố trang 71.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 3` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV2_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các phép ghép đối tượng đồ hoạ trong bảng chọn Path (10 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT3]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[NLS 3.2.NC1.a]`, `[PC2]` – Học sinh phân biệt chính xác 6 phép ghép đối tượng đồ hoạ (`Union`, `Difference`, `Intersection`, `Exclusion`, `Division`, `Cut Path`) và nêu được các bước vẽ đám mây Hình 13.7.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.5 (SGK trang 71)**, **Hình 13.6** và **Hình 13.7 (SGK trang 72)** để hoàn thành **Phiếu học tập số 3**:', 'sub')
    add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 3 hoàn chỉnh bảng 6 phép ghép và quy trình vẽ đám mây Hình 13.7.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 4` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV3_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (07 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[KT2]`, `[KT3]`, `[NLC3]`, `[NLc]`, `[PC3]` – Học sinh vận dụng lý thuyết vừa học để chỉ ra thuộc tính và phép ghép cần dùng cho 03 bài tập Luyện tập (Hình 13.14, 13.15, 13.16 SGK trang 74).', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.14, Hình 13.15, Hình 13.16 (SGK trang 74)** và hoàn thành **Phiếu học tập số 4**:', 'sub')
    add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 4 phân tích đúng cách vẽ Hình 13.14, 13.15, 13.16.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 5` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD3_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[KT2]`, `[KT3]`, `[NLC3]`, `[NLc]`, `[PC2]` – Học sinh thiết kế được quy trình chi tiết để vẽ Miếng dưa hấu Hình 13.1 và chuẩn bị cho Tiết 2 Thực hành (`Tiết PPCT: 20`).', 'sub')
    add(b, 'b) Nội dung: Học sinh liên hệ lại câu hỏi mở đầu để viết quy trình 3 bước vẽ Miếng dưa hấu (Bài Vận dụng 1 SGK trang 74).', 'sub')
    add(b, 'c) Sản phẩm: Quy trình vẽ Miếng dưa hấu trong vở ghi và sự chuẩn bị cho giờ thực hành ở Tiết 2.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD4_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add_bai13_appendix(b, class_label='10/3 (Dự giờ Tiết 1 – Phòng H.I.5)', ppct_label='19')
    add_approval_section(b, sign_date_str='Huế, ngày 15 tháng 11 năm 2026')

    used_figs = [
        'hinh_13_1_2_dua_hau_thuoc_tinh.png',
        'bang_13_1_thuoc_tinh_hinh_khoi.png',
        'hinh_13_3_4_fill_stroke.png',
        'hinh_13_5_ket_hop_doi_tuong.png',
        'hinh_13_6_cac_phep_ghep_hinh.png',
        'hinh_13_7_8_dam_may_logo.png',
        'hinh_13_14_15_16_luyen_tap.png'
    ]
    out_p = KHOI10_DIR / '06_De_cuong_du_gio_Tin10_Bai13_Tiet1_LT_TietPPCT19_Lop10-3.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin10_Bai13_Tiet1.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / '05_KHBD_Tin10_Bai13_Tiet1.docx')
    print('Built:', out_p)

def build_bai13_tiet2():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 2 - THỰC HÀNH – TIẾT PPCT: 20)',
        class_str='Lớp 10/3 – Trường THPT Thuận Hoá',
        schedule_str='Tuần 10 – Tiết thứ 2 của bài / Tiết PPCT: 20 – Phòng máy Tin học',
        lesson_row_str='Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 2 - TH) – Tiết PPCT: 20 (Tổng bài: 02 tiết – PPCT: 19, 20)',
        page1_title='TÊN BÀI DẠY: BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT 2 – THỰC HÀNH)',
        page1_subtitle='(Thời lượng toàn bài: 02 tiết [Tiết PPCT: 19, 20] – Kế hoạch này chi tiết hoá Tiết thứ 2 của bài / Tiết PPCT: 20)',
        page1_class_line='Môn học: Tin học; Lớp: 10/3',
        page1_time_line='Thời gian thực hiện: 01 tiết (45 phút – Tiết 2 - TH / Tiết PPCT: 20)',
        page1_sched_line='Phân phối chương trình: Tiết PPCT 20 (Thực hành nối tiếp Tiết PPCT 19)',
        page1_mode_line='Hình thức tổ chức: Dạy học thực hành tại phòng máy trên phần mềm Inkscape'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 20 – Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 2 - TH)):')
    add(b, '1. Về kiến thức', 'sub')
    for item in [
        '**`[KT1]`**: Thực hành tạo được biểu tượng (`Logo` – Hình 13.8 & 13.9 SGK trang 72–73) bằng cách điều chỉnh thuộc tính cung tròn hở (`Arc`), độ dày nét vẽ (`Width` trong `Stroke style`), nhân bản và xoay đối tượng.',
        '**`[KT2]`**: Thực hành vẽ được **Hình cây** (Hình 13.10 & 13.11 SGK trang 73) bằng sự kết hợp giữa phép hiệu (`Path / Difference`) và phép hợp (`Path / Union`); vẽ được **Hình cầu 3D** (Hình 13.12 & 13.13 SGK trang 73–74) bằng hiệu ứng tô màu chuyển `Linear gradient` và `Radial gradient`.',
        '**`[KT3]`**: Vận dụng tổng hợp thuộc tính hình khối, hộp thoại `Fill and Stroke` và các phép ghép hình để hoàn thành các sản phẩm phần **Luyện tập** (Hình 13.14, 13.15, 13.16) và **Vận dụng** (Miếng dưa hấu Hình 13.1 – SGK trang 74).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**`[NLC1 – Tự chủ và tự học]`**: Học sinh tự đọc hướng dẫn thực hành SGK trang 72–74 và chủ động hoàn thiện 3 nhiệm vụ thực hành trên máy tính.',
        '**`[NLC2 – Giao tiếp và hợp tác]`**: Học sinh trao đổi cặp đôi để kiểm tra thứ tự lớp trước khi thực hiện phép hiệu `Path / Difference`.',
        '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Học sinh khắc phục được lỗi mất hình khi chọn sai lớp trong phép hiệu và sáng tạo màu sắc cho sản phẩm.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**`[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]`**: Sử dụng thành thạo hộp thoại `Fill and Stroke` (`Shift + Ctrl + F`), công cụ `Gradient` và bảng chọn `Path` trong `Inkscape`.',
        '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông]`**: Kết hợp nhuần nhuyễn các phép ghép hình học để tạo ra các biểu tượng đồ hoạ phức tạp.',
        '**`[NLd – Ứng dụng CNTT và truyền thông trong học và tự học]`**: Lưu trữ và quản lí tệp đồ hoạ vectơ `HoTen_10-3_Bai13.svg` khoa học.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**`[NLS 3.2.NC1.a]`**: Tích hợp và tinh chỉnh các đối tượng đồ hoạ vectơ nâng cao để tạo sản phẩm số hoàn chỉnh.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**`[PC1 – Chăm chỉ]`**: Kiên trì rèn luyện từng bước thực hành từ Nhiệm vụ 1 đến Nhiệm vụ 3 và bài tập Luyện tập.',
        '**`[PC2 – Trách nhiệm]`**: Tuân thủ quy định an toàn phòng máy và hoàn thành sản phẩm đúng thời gian.',
        '**`[PC3 – Trung thực]`**: Tự thực hiện bài thực hành và đánh giá trung thực theo Rubric `CCĐG 6`.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Phòng máy tính đã cài đặt phần mềm `Inkscape`, máy chiếu, phiếu đánh giá Rubric `CCĐG 6`.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, **Bài 13: Bổ sung các đối tượng đồ hoạ** (mục Thực hành, Luyện tập, Vận dụng – trang 72–74).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 20 – Mục 4.13. Bổ sung các đối tượng đồ họa (tiết 2 - TH)).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học Tiết 2 - Thực hành (45 phút – Tiết thứ 2 của bài / Tiết PPCT: 20)**', first=0, keep=True)
    ov2_rows = [
        ['Hoạt động dạy học (Công văn 5512)', 'Thời lượng & Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Ôn nhanh phím tắt Path & mở hộp thoại Fill and Stroke)', '04 phút\n(Tiết 2 của bài /\nTiết PPCT: 20)', '[KT1], [KT2],\n[NLC1], [NLa]', 'Khởi động Inkscape, mở hộp thoại `Fill and Stroke` và tạo tệp `HoTen_10-3_Bai13.svg`.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới – Thực hành 3 nhiệm vụ SGK tr. 72–74\n- 2.1. Nhiệm vụ 1: Vẽ Logo 3 cánh (Hình 13.8 & 13.9 – 08 phút)\n- 2.2. Nhiệm vụ 2: Vẽ hình cây (Hình 13.10 & 13.11 – 09 phút)\n- 2.3. Nhiệm vụ 3: Vẽ hình cầu 3D (Hình 13.12 & 13.13 – 08 phút)', '25 phút\n(Tiết 2 của bài /\nTiết PPCT: 20)', '[KT1], [KT2],\n[NLC1], [NLC2],\n[NLa], [NLc], [NLd],\n[NLS 3.2.NC1.a],\n[PC1], [PC2]', 'Tệp `HoTen_10-3_Bai13.svg` chứa đủ 3 sản phẩm: Logo (Hình 13.8), Hình cây (Hình 13.10) và Hình cầu 3D (Hình 13.12).\nCông cụ đánh giá: CCĐG 6 (Rubric thực hành).'],
        ['3. Hoạt động 3: Luyện tập\n(Thực hành vẽ Hình 13.14, 13.15, 13.16 SGK tr. 74)', '09 phút\n(Tiết 2 của bài /\nTiết PPCT: 20)', '[KT3], [NLC1],\n[NLC3], [NLc], [PC3]', 'Sản phẩm Bông hoa Hình 13.14, Hình 13.15 hoặc Chùm bóng Hình 13.16.\nCông cụ đánh giá: CCĐG 5 & CCĐG 6.'],
        ['4. Hoạt động 4: Vận dụng\n(Thực hành vẽ Miếng dưa hấu Hình 13.1 SGK tr. 74)', '07 phút\n(Tiết 2 của bài /\nTiết PPCT: 20)', '[KT3], [NLC2],\n[NLC3], [NLc],\n[PC1], [PC2]', 'Hình vẽ Miếng dưa hấu hoàn chỉnh trong tệp `HoTen_10-3_Bai13.svg`.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov2_rows, [2700, 1400, 2000, 2900], header=True, cell_sz=22))

    add(b, '1. Hoạt động 1: Mở đầu (04 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[KT2]`, `[NLC1]`, `[NLa]` – Học sinh khởi động phần mềm `Inkscape`, tạo tệp thực hành mới và nhắc lại chính xác tổ hợp phím tắt của phép hợp (`Ctrl + +`), phép hiệu (`Ctrl + -`) và hộp thoại `Fill and Stroke` (`Shift + Ctrl + F`).', 'sub')
    add(b, 'b) Nội dung: Giáo viên yêu cầu học sinh khởi động `Inkscape`, lưu tệp `HoTen_10-3_Bai13.svg` và trả lời nhanh 3 câu hỏi kiểm tra bài cũ.', 'sub')
    add(b, 'c) Sản phẩm: Tệp `HoTen_10-3_Bai13.svg` đã mở sẵn kèm hộp thoại `Fill and Stroke` ở cạnh phải màn hình.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 1` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    hd1_t2 = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên yêu cầu học sinh mở phần mềm `Inkscape`, nhấn `Ctrl + S` lưu tệp `HoTen_10-3_Bai13.svg`, mở hộp thoại `Fill and Stroke` và trả lời nhanh:\n'
            '  + Tổ hợp phím tắt của phép hợp (`Union`) và phép hiệu (`Difference`) trong bảng chọn `Path` là gì?\n'
            '  + Trong phép hiệu (`Difference`), hình nào bị cắt và hình nào dùng làm khuôn cắt?\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh mở phần mềm, tạo tệp và giơ tay trả lời.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên mời 02 học sinh trả lời nhanh.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên chốt lại quy tắc và giao 3 nhiệm vụ thực hành trọng tâm trang 72–74 SGK.',
            '**Tổ hợp phím tắt và quy tắc cần nhớ khi thực hành:**\n'
            '- Mở hộp thoại **`Fill and Stroke`:** `Object / Fill and Stroke` (hoặc `Shift + Ctrl + F`).\n'
            '- **Phép hợp (`Path / Union`):** `Ctrl + +` (gộp các hình thành một hình chung).\n'
            '- **Phép hiệu (`Path / Difference`):** `Ctrl + -` (lấy **hình lớp dưới** trừ đi phần giao với **hình lớp trên**).'
        ]
    ]
    b.append(table(hd1_t2, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới – Thực hành 03 nhiệm vụ SGK trang 72–74 (25 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Thực hành vẽ biểu tượng Logo – Hình 13.8 & Hình 13.9 SGK trang 72–73 (08 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'sub')
    add(b, 'a) Mục tiêu: `[KT1]`, `[NLC1]`, `[NLa]`, `[NLS 3.2.NC1.a]`, `[PC1]` – Học sinh tạo được cung tròn hở (`Arc`), chỉnh độ dày nét vẽ (`Width`) trong `Stroke style`, nhân bản và xoay ghép thành logo 3 cánh.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.9 (SGK trang 73)** và thực hành theo 6 bước hướng dẫn:', 'sub')
    add_figure(b, 'hinh_13_9_ve_logo.png')
    add(b, 'c) Sản phẩm: Hình vẽ Logo 3 cánh màu đỏ đậm hoàn chỉnh.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` – Tiêu chí 1 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV4_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2 & Nhiệm vụ 3: Thực hành vẽ hình cây (Hình 13.10–13.11) và vẽ hình cầu 3D (Hình 13.12–13.13) (17 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'sub')
    add(b, 'a) Mục tiêu: `[KT2]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[NLc]`, `[NLd]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh thực hành thành thạo phép hiệu (`Difference`), phép hợp (`Union`) để vẽ hình cây và dùng màu chuyển (`Linear/Radial gradient`) để vẽ hình cầu 3D.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.10, 13.11 (SGK trang 73)** và **Hình 13.12, 13.13 (SGK trang 73–74)** để thực hành trên máy tính:', 'sub')
    add_figure(b, 'hinh_13_10_11_ve_cay.png')
    add_figure(b, 'hinh_13_12_13_hinh_cau.png')
    add(b, 'c) Sản phẩm: Hình cây xanh và Hình cầu 3D hoàn chỉnh trên tệp `HoTen_10-3_Bai13.svg`.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` – Tiêu chí 2 & Tiêu chí 3 của Rubric thực hành *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    nv23_t2 = [
        ['HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH', 'SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)'],
        [
            '**Bước 1. Giáo viên chuyển giao nhiệm vụ học tập**\n'
            '- Giáo viên hướng dẫn nhanh điểm mấu chốt của **Nhiệm vụ 2 (Vẽ hình cây – Hình 13.11)** và **Nhiệm vụ 3 (Vẽ hình cầu – Hình 13.13)**, sau đó yêu cầu học sinh thực hành trong 14 phút:\n'
            '  + Ở Nhiệm vụ 2: Hình chữ nhật làm thân cây phải đặt ở **lớp dưới** hai hình tròn trước khi nhấn `Path / Difference` từng bên.\n'
            '  + Ở Nhiệm vụ 3: Cả 3 hình đều phải tắt màu viền (`Stroke paint = X`), hình tròn đen dùng `Linear gradient`, hình tròn trắng dùng `Radial gradient`.\n'
            '**Bước 2. Học sinh thực hiện nhiệm vụ**\n'
            '- Học sinh thực hành vẽ Hình cây và Hình cầu 3D trên máy tính. Giáo viên hỗ trợ trực tiếp từng bàn.\n'
            '**Bước 3. Báo cáo sản phẩm**\n'
            '- Giáo viên chiếu 02 bài làm tốt nhất lên màn chiếu để cả lớp quan sát và học tập.\n'
            '**Bước 4. Đánh giá, nhận xét**\n'
            '- Giáo viên chấm điểm theo Tiêu chí 2 (3,0 điểm) và Tiêu chí 3 (2,5 điểm) của Rubric `CCĐG 6`.',
            '**Quy trình chuẩn thực hành Nhiệm vụ 2 & Nhiệm vụ 3 (SGK trang 73–74):**\n'
            '- **Nhiệm vụ 2 (Vẽ hình cây – Hình 13.11):**\n'
            '  + *Tán lá:* Vẽ 3 hình tròn xếp chụm vào nhau (H.13.11a).\n'
            '  + *Thân cây:* Đặt hình chữ nhật ở lớp dưới 2 hình tròn hai bên (H.13.11b) -> thực hiện `Path / Difference` (`Ctrl + -`) lần lượt với từng hình tròn để tạo thân cây uốn cong (H.13.11c).\n'
            '  + *Ghép hoàn thiện:* Đặt thân cây vào tán lá -> chọn cả 4 hình -> nhấn `Path / Union` (`Ctrl + +`) và tô màu xanh lá (H.13.11d–f).\n'
            '- **Nhiệm vụ 3 (Vẽ hình cầu 3D – Hình 13.13):**\n'
            '  + Vẽ hình chữ nhật vàng, hình tròn lớn màu đen và hình tròn nhỏ màu trắng không viền (H.13.13a).\n'
            '  + Tô `Linear gradient` cho hình tròn đen (H.13.13b) và `Radial gradient` cho hình tròn trắng (H.13.13c–d).'
        ]
    ]
    b.append(table(nv23_t2, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (09 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, 'a) Mục tiêu: `[KT3]`, `[NLC1]`, `[NLC3]`, `[NLc]`, `[PC3]` – Học sinh thực hành vẽ được Bông hoa cách điệu (Hình 13.14), biểu tượng (Hình 13.15) và chùm bóng bay (Hình 13.16) ở trang 74 SGK.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.14, 13.15, 13.16 (SGK trang 74)** và thực hành trên máy tính:', 'sub')
    add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    add(b, 'c) Sản phẩm: Hình 13.14 (và Hình 13.15 hoặc 13.16) được vẽ bổ sung vào tệp `HoTen_10-3_Bai13.svg`.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 5` & `CCĐG 6` *(Đính kèm tại Mục IV. Phụ lục 1 & 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD3_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (07 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, 'a) Mục tiêu: `[KT3]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[PC1]`, `[PC2]` – Học sinh hoàn thành sản phẩm vẽ Miếng dưa hấu (Hình 13.1) bằng thuộc tính góc `Start = 0, End = 180` ở phần Vận dụng SGK trang 74.', 'sub')
    add(b, 'b) Nội dung: Học sinh thực hành Bài Vận dụng 1 & 2 (SGK trang 74) và lưu tệp hoàn chỉnh.', 'sub')
    add(b, 'c) Sản phẩm: Tệp `HoTen_10-3_Bai13.svg` hoàn chỉnh toàn bài.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` – Rubric 10 điểm *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD4_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add_bai13_appendix(b, class_label='10/3 (Tiết 2 Thực hành)', ppct_label='20')
    add_approval_section(b, sign_date_str='Huế, ngày 15 tháng 11 năm 2026')

    used_figs = [
        'hinh_13_9_ve_logo.png',
        'hinh_13_10_11_ve_cay.png',
        'hinh_13_12_13_hinh_cau.png',
        'hinh_13_14_15_16_luyen_tap.png'
    ]
    out_p = KHOI10_DIR / '07_De_cuong_du_gio_Tin10_Bai13_Tiet2_TH_TietPPCT20_Lop10-3.docx'
    save_docx(b, out_p, used_figures=used_figs)
    shutil.copy2(out_p, THANH_PHAM_DIR / 'Đề cương bài dạy Tin10_Bai13_Tiet2.docx')
    shutil.copy2(out_p, GOI_NOP_DIR / '06_KHBD_Tin10_Bai13_Tiet2.docx')
    print('Built:', out_p)

def build_bai13_cabai():
    b = init_body_with_cover(
        cover_title='ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY',
        cover_subtitle='BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (02 TIẾT – TIẾT PPCT: 19, 20)',
        class_str='Lớp 10/3 – Trường THPT Thuận Hoá',
        schedule_str='Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45, dự giờ Tiết 1 / Tiết PPCT: 19) – Phòng H.I.5',
        lesson_row_str='Bài 13. Bổ sung các đối tượng đồ hoạ (02 tiết – Tiết PPCT: 19, 20)',
        page1_title='TÊN BÀI DẠY: BÀI 13. BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ',
        page1_subtitle='(Thời lượng: 02 tiết = 90 phút – Tiết thứ 1, 2 của bài; Tiết PPCT: 19, 20)',
        page1_class_line='Môn học: Tin học; Lớp: 10/3',
        page1_time_line='Thời gian thực hiện: 02 tiết (90 phút – Tiết PPCT: 19 [Lí thuyết], Tiết PPCT: 20 [Thực hành])',
        page1_sched_line='Lịch dự giờ: Thứ Hai, ngày 16/11/2026 – Tiết 3 (09h00 – 09h45, dự giờ Tiết 1 / Tiết PPCT: 19) – Lớp 10/3 – Phòng H.I.5',
        page1_mode_line='Hình thức tổ chức: Dạy học lí thuyết kết hợp thực hành tại phòng máy'
    )

    add(b, 'I. MỤC TIÊU', 'heading')
    add(b, 'Căn cứ Yêu cầu cần đạt của Chương trình Giáo dục phổ thông 2018 môn Tin học (Thông tư số 32/2018/TT-BGDĐT), lớp 10, Chủ đề E và Kế hoạch dạy học môn Tin học Khối 10 của Trường THPT Thuận Hoá (Tiết PPCT 19, 20 – Mục 4.13. Bổ sung các đối tượng đồ họa):')
    add(b, '1. Về kiến thức', 'sub')
    for item in [
        '**`[KT1]`**: Nêu và sử dụng được các thuộc tính cơ bản của từng đối tượng hình khối trên Thanh điều khiển thuộc tính (`W, H, Rx, Ry`, `Start, End`, `Corners, Rounded, Spoke ratio, Randomized` – SGK trang 69–70).',
        '**`[KT2]`**: Thiết lập được màu tô (`Fill`), màu vẽ (`Stroke paint`), kiểu và độ dày nét vẽ (`Stroke style`), màu chuyển (`Linear/Radial gradient`) và độ trong suốt (`Opacity`) bằng hộp thoại `Fill and Stroke` (SGK trang 70–71).',
        '**`[KT3]`**: Trình bày và thực hiện đúng 06 phép ghép đối tượng đồ hoạ trong bảng chọn `Path` (`Union, Difference, Intersection, Exclusion, Division, Cut Path` – SGK trang 71–72).',
        '**`[KT4]`**: Thực hành vẽ hoàn chỉnh Logo (Hình 13.8–13.9), Hình cây (Hình 13.10–13.11), Hình cầu 3D (Hình 13.12–13.13) và các bài tập Luyện tập, Vận dụng (SGK trang 72–74).'
    ]:
        add(b, item, 'bullet')

    add(b, '2. Về năng lực', 'sub')
    add(b, 'a) Năng lực chung', 'sub')
    for item in [
        '**`[NLC1 – Tự chủ và tự học]`**: Chủ động đọc hiểu Sách giáo khoa trang 69–74 và tự thực hành trên phần mềm `Inkscape`.',
        '**`[NLC2 – Giao tiếp và hợp tác]`**: Thảo luận cặp đôi hiệu quả khi phân tích các phép ghép hình và hỗ trợ nhau trong giờ thực hành.',
        '**`[NLC3 – Giải quyết vấn đề và sáng tạo]`**: Vận dụng linh hoạt các phép ghép và hiệu ứng chuyển màu để tạo các sản phẩm đồ hoạ phức tạp.'
    ]:
        add(b, item, 'bullet')

    add(b, 'b) Năng lực đặc thù môn Tin học', 'sub')
    for item in [
        '**`[NLa – Sử dụng và quản lí các phương tiện CNTT và truyền thông]`**: Sử dụng thành thạo Thanh điều khiển thuộc tính, hộp thoại `Fill and Stroke` và bảng chọn `Path` trong `Inkscape`.',
        '**`[NLc – Giải quyết vấn đề với sự hỗ trợ của CNTT và truyền thông]`**: Phân tích cấu trúc hình học của sản phẩm mẫu để chọn đúng tổ hợp thuộc tính và phép ghép.',
        '**`[NLd – Ứng dụng CNTT và truyền thông trong học và tự học]`**: Lưu trữ và quản lí tệp thực hành `.svg` đúng quy chuẩn.'
    ]:
        add(b, item, 'bullet')

    add(b, 'c) Năng lực số', 'sub')
    add(b, '**`[NLS 3.2.NC1.a]`**: Tích hợp, biến đổi và tinh chỉnh các đối tượng đồ hoạ số nâng cao.', 'bullet')

    add(b, '3. Về phẩm chất', 'sub')
    for item in [
        '**`[PC1 – Chăm chỉ]`**: Kiên trì nghiên cứu lý thuyết và rèn luyện kỹ năng thực hành.',
        '**`[PC2 – Trách nhiệm]`**: Hoàn thành đầy đủ các nhiệm vụ học tập ở cả 2 tiết học.',
        '**`[PC3 – Trung thực]`**: Tự giác làm bài và đánh giá trung thực theo Rubric.'
    ]:
        add(b, item, 'bullet')

    add(b, 'II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU', 'heading')
    for item in [
        '**Thiết bị dạy học:** Máy tính giáo viên và học sinh cài đặt `Inkscape`, máy chiếu tại phòng `H.I.5`.',
        '**Sách giáo khoa chính thức:** *Sách giáo khoa Tin học 10 – Bộ sách Kết nối tri thức với cuộc sống, Nhà xuất bản Giáo dục Việt Nam*, **Bài 13: Bổ sung các đối tượng đồ hoạ** (trang 69–74).',
        '**Phân phối chương trình:** *Kế hoạch dạy học môn Tin học Khối 10 năm học 2026–2027 của Trường THPT Thuận Hoá* (Tiết PPCT 19 [tiết 1 - LT] và Tiết PPCT 20 [tiết 2 - TH]).'
    ]:
        add(b, item, 'bullet')

    add(b, 'III. TIẾN TRÌNH DẠY HỌC', 'heading')
    add(b, '**Bảng tổng quan tiến trình dạy học toàn bài (02 tiết = 90 phút – Tiết PPCT: 19, 20)**', first=0, keep=True)
    ov_all = [
        ['Hoạt động dạy học (Chuẩn Công văn 5512)', 'Thời lượng & Phân bổ Tiết PPCT', 'Mục tiêu', 'Sản phẩm & Công cụ đánh giá'],
        ['1. Hoạt động 1: Mở đầu\n(Phân tích Hình 13.1 – Miếng dưa hấu)', '05 phút\n(Thuộc Tiết 1 của bài /\nTiết PPCT: 19)', '[KT1], [NLC1], [PC1]', 'Phân tích 4 lớp đối tượng của miếng dưa hấu Hình 13.1.\nCông cụ đánh giá: CCĐG 1.'],
        ['2. Hoạt động 2: Hình thành kiến thức mới (65 phút)\n- 2.1. Nhiệm vụ 1: Các đối tượng hình khối & thuộc tính (12p – Tiết 1 / PPCT: 19)\n- 2.2. Nhiệm vụ 2: Thiết lập màu tô, màu vẽ & Fill and Stroke (13p – Tiết 1 / PPCT: 19)\n- 2.3. Nhiệm vụ 3: Các phép ghép đối tượng đồ hoạ Path (15p – Tiết 1 / PPCT: 19)\n- 2.4. Nhiệm vụ 4: Thực hành vẽ Logo, Hình cây & Hình cầu 3D (25p – Tiết 2 / PPCT: 20)', '65 phút\n(40p ở Tiết 1 / PPCT: 19\n+ 25p ở Tiết 2 / PPCT: 20)', '[KT1], [KT2], [KT3], [KT4],\n[NLC1], [NLC2], [NLC3],\n[NLa], [NLc], [NLd],\n[NLS 3.2.NC1.a],\n[PC1], [PC2]', 'Phiếu học tập số 1, 2, 3 (Tiết 1) và tệp `HoTen_10-3_Bai13.svg` chứa Hình 13.8–13.13 (Tiết 2).\nCông cụ đánh giá: CCĐG 2, 3, 4, 6.'],
        ['3. Hoạt động 3: Luyện tập\n(Giải và thực hành Hình 13.14, 13.15, 13.16 SGK tr. 74)', '12 phút\n(Thuộc Tiết 2 của bài /\nTiết PPCT: 20)', '[KT1], [KT2], [KT3], [KT4],\n[NLC1], [NLC3], [NLc], [PC3]', 'Phiếu học tập số 4 và sản phẩm Hình 13.14–13.16.\nCông cụ đánh giá: CCĐG 5 & CCĐG 6.'],
        ['4. Hoạt động 4: Vận dụng\n(Thực hành vẽ Miếng dưa hấu Hình 13.1 SGK tr. 74)', '08 phút\n(Thuộc Tiết 2 của bài /\nTiết PPCT: 20)', '[KT4], [NLC2], [NLC3],\n[NLc], [PC1], [PC2]', 'Sản phẩm Miếng dưa hấu Hình 13.1 trong tệp `.svg`.\nCông cụ đánh giá: CCĐG 6 (Rubric 10 điểm).']
    ]
    b.append(table(ov_all, [2800, 1500, 1900, 2800], header=True, cell_sz=21))

    add(b, '1. Hoạt động 1: Mở đầu (05 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[NLC1]`, `[PC1]` – Học sinh phân tích được các lớp đối tượng của miếng dưa hấu Hình 13.1.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.1 và Hình 13.2 (SGK trang 69)**:', 'sub')
    add_figure(b, 'hinh_13_1_2_dua_hau_thuoc_tinh.png')
    add(b, 'c) Sản phẩm: Câu trả lời nhận diện 4 lớp của miếng dưa hấu.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 1` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD1_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2. Hoạt động 2: Hình thành kiến thức mới (65 phút: Nhiệm vụ 1, 2, 3 thuộc Tiết 1 / Tiết PPCT: 19; Nhiệm vụ 4 thuộc Tiết 2 / Tiết PPCT: 20)', 'heading')
    add(b, '2.1. Nhiệm vụ 1: Tìm hiểu các đối tượng hình khối và thuộc tính cơ bản (12 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT1]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[PC1]` – Học sinh nắm vững bảng thuộc tính cơ bản của hình vuông/chữ nhật, hình tròn/elip và hình đa giác/ngôi sao.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Bảng các thuộc tính cơ bản (SGK trang 69–70)** và hoàn thành **Phiếu học tập số 1**:', 'sub')
    add_figure(b, 'bang_13_1_thuoc_tinh_hinh_khoi.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 1 hoàn chỉnh.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 2` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV1_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.2. Nhiệm vụ 2: Tìm hiểu thiết lập màu tô, màu vẽ và hộp thoại Fill and Stroke (13 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT2]`, `[NLC1]`, `[NLa]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh nắm vững chức năng 3 trang của hộp thoại `Fill and Stroke`, các kiểu tô màu chuyển và `Opacity`.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.3 và Hình 13.4 (SGK trang 70)** và hoàn thành **Phiếu học tập số 2**:', 'sub')
    add_figure(b, 'hinh_13_3_4_fill_stroke.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 2 hoàn chỉnh.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 3` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV2_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.3. Nhiệm vụ 3: Tìm hiểu các phép ghép đối tượng đồ hoạ trong bảng chọn Path (15 phút – Tiết 1 của bài / Tiết PPCT: 19)', 'sub')
    add(b, 'a) Mục tiêu: `[KT3]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[NLS 3.2.NC1.a]`, `[PC2]` – Học sinh phân biệt chính xác 6 phép ghép đối tượng đồ hoạ trong bảng chọn `Path`.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.5, Hình 13.6, Hình 13.7 (SGK trang 71–72)** và hoàn thành **Phiếu học tập số 3**:', 'sub')
    add_figure(b, 'hinh_13_5_ket_hop_doi_tuong.png')
    add_figure(b, 'hinh_13_6_cac_phep_ghep_hinh.png')
    add_figure(b, 'hinh_13_7_8_dam_may_logo.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 3 hoàn chỉnh.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 4` *(Đính kèm tại Mục IV. Phụ lục 1)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV3_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '2.4. Nhiệm vụ 4: Thực hành vẽ Logo (Hình 13.8–13.9), Hình cây (Hình 13.10–13.11) và Hình cầu 3D (Hình 13.12–13.13) (25 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'sub')
    add(b, 'a) Mục tiêu: `[KT4]`, `[NLC1]`, `[NLC2]`, `[NLa]`, `[NLc]`, `[NLd]`, `[NLS 3.2.NC1.a]`, `[PC1]`, `[PC2]` – Học sinh thực hành hoàn chỉnh cả 3 nhiệm vụ trong SGK trang 72–74.', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.9, 13.10, 13.11, 13.12, 13.13 (SGK trang 73–74)** và thực hành trên máy tính:', 'sub')
    add_figure(b, 'hinh_13_9_ve_logo.png')
    add_figure(b, 'hinh_13_10_11_ve_cay.png')
    add_figure(b, 'hinh_13_12_13_hinh_cau.png')
    add(b, 'c) Sản phẩm: Tệp `HoTen_10-3_Bai13.svg` chứa đủ Logo, Hình cây và Hình cầu 3D.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(NV4_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '3. Hoạt động 3: Luyện tập (12 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, 'a) Mục tiêu: `[KT1]`, `[KT2]`, `[KT3]`, `[KT4]`, `[NLC1]`, `[NLC3]`, `[NLc]`, `[PC3]` – Học sinh giải và thực hành vẽ Hình 13.14, 13.15, 13.16 (SGK trang 74).', 'sub')
    add(b, 'b) Nội dung: Học sinh quan sát **Hình 13.14, 13.15, 13.16 (SGK trang 74)** và thực hành:', 'sub')
    add_figure(b, 'hinh_13_14_15_16_luyen_tap.png')
    add(b, 'c) Sản phẩm: Phiếu học tập số 4 và hình vẽ Luyện tập trong tệp `.svg`.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 5` & `CCĐG 6` *(Đính kèm tại Mục IV. Phụ lục 1 & 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD3_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add(b, '4. Hoạt động 4: Vận dụng (08 phút – Tiết 2 của bài / Tiết PPCT: 20)', 'heading')
    add(b, 'a) Mục tiêu: `[KT4]`, `[NLC2]`, `[NLC3]`, `[NLc]`, `[PC1]`, `[PC2]` – Học sinh thực hành vẽ hoàn chỉnh Miếng dưa hấu Hình 13.1.', 'sub')
    add(b, 'b) Nội dung: Học sinh thực hành Bài Vận dụng 1 & 2 (SGK trang 74).', 'sub')
    add(b, 'c) Sản phẩm: Sản phẩm Miếng dưa hấu Hình 13.1 trong tệp `HoTen_10-3_Bai13.svg`.', 'sub')
    add(b, '**Công cụ đánh giá:** `CCĐG 6` – Rubric 10 điểm *(Đính kèm tại Mục IV. Phụ lục 2)*.', first=0, keep=True)
    add(b, 'd) Tổ chức thực hiện:', 'sub')
    b.append(table(HD4_ROWS_BAI13, [4600, 4400], header=True, cell_sz=23))

    add_bai13_appendix(b, class_label='10/3 (Trọn bộ 02 tiết)', ppct_label='19, 20')
    add_approval_section(b, sign_date_str='Huế, ngày 15 tháng 11 năm 2026')

    used_figs = [
        'hinh_13_1_2_dua_hau_thuoc_tinh.png',
        'bang_13_1_thuoc_tinh_hinh_khoi.png',
        'hinh_13_3_4_fill_stroke.png',
        'hinh_13_5_ket_hop_doi_tuong.png',
        'hinh_13_6_cac_phep_ghep_hinh.png',
        'hinh_13_7_8_dam_may_logo.png',
        'hinh_13_9_ve_logo.png',
        'hinh_13_10_11_ve_cay.png',
        'hinh_13_12_13_hinh_cau.png',
        'hinh_13_14_15_16_luyen_tap.png'
    ]
    out_p = KHOI10_DIR / '08_De_cuong_du_gio_Tin10_Bai13_CaBai_2Tiet_TietPPCT19-20.docx'
    save_docx(b, out_p, used_figures=used_figs)
    print('Built:', out_p)

def build_handout_khoi10():
    h = node('body')
    add(h, 'TRƯỜNG THPT THUẬN HOÁ – TỔ TIN HỌC', 'title')
    add(h, 'BỘ PHIẾU HỌC TẬP, BẢNG KIỂM, RUBRIC VÀ ĐÁP ÁN MÔN TIN HỌC 10', 'title')
    add(h, '(Dùng cho các tiết Dự giờ Khối 10: Ôn tập GHKI [Tiết PPCT: 15], Bài 12 [Tiết PPCT: 17, 18], Bài 13 [Tiết PPCT: 19, 20])', align='center', first=0, italic=True, keep=True)
    add(h, 'Giáo viên giảng dạy (SVTT): Hồ Quang Huy – Giáo viên hướng dẫn: Thầy Mai Xuân Văn', align='center', first=0, keep=True)

    add(h, 'PHẦN I. BỘ PHIẾU HỌC TẬP TIẾT ÔN TẬP KIỂM TRA GIỮA HỌC KÌ I (TIẾT PPCT: 15 – LỚP 10/2)', 'heading')
    add(h, 'Họ và tên học sinh: ................................................................... Lớp: 10/2 | Nhóm / Bàn: ...............', first=0)
    pht_ontap = [
        ['Trạm chuyên đề', 'Câu hỏi rà soát kiến thức cốt lõi (SGK Tin học 10 KNTT)', 'Phần ghi bài làm của học sinh'],
        ['Trạm 1:\nChủ đề 1\n(Mục 1.1, Mục 1.2, Mục 1.7\nPPCT: 01–06)',
         '1. Ba bước xử lí thông tin của máy tính (Hình 1.1 tr. 6): ......................\n'
         '2. Quy đổi (Bảng 1.1 tr. 8): 1 B = .... bit; 1 KB = ...... B; 1 GB = ...... MB.\n'
         '3. Hai đặc điểm của thiết bị thông minh (Bài 2 tr. 11): ...........................',
         '1. ..................................................................\n'
         '2. ..................................................................\n'
         '3. ..................................................................'],
        ['Trạm 2:\nChủ đề 2\n(Mục 2.8, Mục 2.9, Mục 2.10\nPPCT: 07–12)',
         '1. Khác biệt giữa LAN và Internet về Phạm vi, Kết nối, Sở hữu (tr. 39).\n'
         '2. Phân biệt Virus, Worm, Trojan (tr. 46) & 3 cách bảo vệ thông tin (tr. 44).\n'
         '3. Ba mức phân quyền chia sẻ tệp trên Google Drive (tr. 52).',
         '1. ..................................................................\n'
         '2. ..................................................................\n'
         '3. ..................................................................'],
        ['Trạm 3:\nChủ đề 3\n(Mục 3.11 – Bài 11\nPPCT: 13–14)',
         '1. Sáu hành vi vi phạm đạo đức, pháp luật trên mạng (Hộp tr. 56).\n'
         '2. Phân biệt Quyền nhân thân và Quyền tài sản (Luật SHTT – tr. 59).\n'
         '3. Nghĩa vụ tôn trọng bản quyền phần mềm và thông tin số (tr. 60).',
         '1. ..................................................................\n'
         '2. ..................................................................\n'
         '3. ..................................................................']
    ]
    h.append(table(pht_ontap, [1600, 4100, 3300], header=True, cell_sz=21))

    br(h)
    add(h, 'PHẦN II. BỘ PHIẾU HỌC TẬP & RUBRIC BÀI 12: PHẦN MỀM THIẾT KẾ ĐỒ HOẠ (TIẾT PPCT: 17, 18 – LỚP 10/1 & 10/2)', 'heading')
    add(h, 'Họ và tên học sinh: ................................................................... Lớp: 10/.... | Máy số: ...............', first=0)
    pht_b12 = [
        ['Phiếu học tập', 'Câu hỏi và nhiệm vụ (SGK Tin học 10 KNTT, Bài 12 trang 63–68)', 'Phần trả lời của học sinh / Đáp án đối chiếu'],
        ['Phiếu số 1\n(Mục 1 – tr. 63–64)',
         '1. So sánh Đồ hoạ điểm ảnh (`Bitmap`) và Đồ hoạ vectơ (`Vector`) theo 5 tiêu chí (Bảng tr. 64).\n'
         '2. Ảnh chụp là loại đồ hoạ nào? Tại sao thiết kế logo dùng đồ hoạ vectơ?',
         '- Bitmap: tập điểm ảnh (`pixel`), phóng to bị vỡ hình, tệp lớn.\n'
         '- Vector: phương trình toán học, co dãn không vỡ hình, tệp nhẹ (`.svg`), rất hợp tạo logo.'],
        ['Phiếu số 2\n(Mục 2 – tr. 64–66)',
         '1. Phân loại phần mềm: `Illustrator, CorelDRAW, Inkscape, Photoshop, GIMP`.\n'
         '2. Kể tên 6 vùng giao diện chính của Inkscape (Hình 12.3 tr. 65).\n'
         '3. Câu 1 & Câu 2 trang 66 SGK.',
         '- Vector: `Illustrator, CorelDRAW, Inkscape`; Bitmap: `Photoshop, GIMP`.\n'
         '- Câu 1 tr. 66: **B. Inkscape**; Câu 2 tr. 66: **A. Toàn bộ vùng làm việc**.'],
        ['Phiếu số 3\n(Mục 3 – tr. 66–67)',
         '1. Giải bài toán xếp lớp quả dưa hấu và quả cam (Hoạt động 3 tr. 66).\n'
         '2. Phân biệt nháy lần 1 và nháy lần 2 công cụ `Select`.\n'
         '3. Cách chọn màu `Fill`, màu `Stroke` và lưu tệp `.svg`.',
         '- Đối tượng vẽ trước ở lớp dưới, vẽ sau ở lớp trên.\n'
         '- Nháy lần 1: Di chuyển/co dãn; Nháy lần 2: Xoay.\n'
         '- Nháy chuột = `Fill`; `Shift + Nháy chuột` = `Stroke`.'],
        ['Phiếu số 4\n(Thực hành & LT\ntr. 67–68)',
         '1. Thực hành vẽ Bông hoa (Hình 12.5) và Quốc kì Việt Nam (Hình 12.6).\n'
         '2. Giải Bài Luyện tập 1 & 2 (SGK tr. 68).',
         '- Bài LT 1 tr. 68: Hình tròn có nền **màu Cam**, viền **màu Đỏ đậm**.\n'
         '- Bài LT 2 tr. 68: Chọn **C. Hộp công cụ**.']
    ]
    h.append(table(pht_b12, [1600, 4100, 3300], header=True, cell_sz=21))

    br(h)
    add(h, 'PHẦN III. BỘ PHIẾU HỌC TẬP & RUBRIC BÀI 13: BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ (TIẾT PPCT: 19, 20 – LỚP 10/3)', 'heading')
    add(h, 'Họ và tên học sinh: ................................................................... Lớp: 10/3 | Máy số: ...............', first=0)
    pht_b13 = [
        ['Phiếu học tập', 'Câu hỏi và nhiệm vụ (SGK Tin học 10 KNTT, Bài 13 trang 69–74)', 'Phần trả lời của học sinh / Đáp án đối chiếu'],
        ['Phiếu số 1\n(Mục 1 – tr. 69–70)',
         '1. Giải thích các thuộc tính: `W, H, Rx, Ry` (chữ nhật); `Start, End` (tròn/elip); `Corners, Rounded, Spoke ratio, Randomized` (đa giác/sao).\n'
         '2. Câu hỏi SGK tr. 70: Làm cong đỉnh ngôi sao 5 cánh bằng tham số nào?',
         '- `W, H`: rộng, cao; `Rx, Ry`: bán kính bo góc / bán kính elip; `Start, End`: góc đầu, cuối; `Corners`: số đỉnh/cánh; `Rounded`: độ cong đỉnh; `Spoke ratio`: tỉ lệ cánh sao.\n'
         '- Câu hỏi tr. 70: Tham số **`Rounded`**.'],
        ['Phiếu số 2\n(Mục 2 – tr. 70–71)',
         '1. Nêu chức năng 3 trang `Fill`, `Stroke paint`, `Stroke style` trong hộp thoại `Fill and Stroke` (Hình 13.4 tr. 70) và thông số `Opacity`.\n'
         '2. Trả lời Câu 1 & Câu 2 trang 71 SGK.',
         '- `Fill`: màu tô; `Stroke paint`: màu vẽ viền; `Stroke style`: kiểu nét và độ dày `Width`; `Opacity`: độ trong suốt.\n'
         '- Câu 1 tr. 71: **`C. Stroke style`**; Câu 2 tr. 71: **`B`** (bút chỉnh gradient).'],
        ['Phiếu số 3\n(Mục 3 – tr. 71–72)',
         '1. Nêu phím tắt và kết quả của 6 phép ghép trong `Path` (Hình 13.6 tr. 72).\n'
         '2. Câu hỏi SGK tr. 72: Nêu phép ghép và các bước vẽ đám mây Hình 13.7.',
         '- `Union (Ctrl++)`, `Difference (Ctrl+-)`, `Intersection (Ctrl+*)`, `Exclusion (Ctrl+^)`, `Division (Ctrl+/)`, `Cut Path (Ctrl+Alt+/)`.\n'
         '- Đám mây Hình 13.7: Vẽ các hình tròn chồng lấp -> `Path / Union`.'],
        ['Phiếu số 4\n(Thực hành, LT & VD\ntr. 72–74)',
         '1. Thực hành vẽ Logo (Hình 13.8), Hình cây (Hình 13.10), Hình cầu (Hình 13.12).\n'
         '2. Luyện tập Hình 13.14–13.16 & Vận dụng vẽ Miếng dưa hấu Hình 13.1.',
         '- Logo: cung tròn hở + tăng `Width` + `Ctrl+D`.\n'
         '- Hình cây: `Path/Difference` (thân) + `Path/Union` (tán lá).\n'
         '- Hình cầu: `Linear gradient` + `Radial gradient`.']
    ]
    h.append(table(pht_b13, [1600, 4100, 3300], header=True, cell_sz=21))

    out_h = KHOI10_DIR / '09_Phieu_hoc_tap_va_Rubric_Tin10_OnTap_Bai12_Bai13.docx'
    save_docx(h, out_h, used_figures=None, is_handout=True)
    print('Built:', out_h)

if __name__ == '__main__':
    build_bai13_tiet1()
    build_bai13_tiet2()
    build_bai13_cabai()
    build_handout_khoi10()
