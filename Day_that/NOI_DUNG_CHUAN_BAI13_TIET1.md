# NỘI DUNG CHUẨN BÀI 13 (TIẾT 1 – LÍ THUYẾT): BỔ SUNG CÁC ĐỐI TƯỢNG ĐỒ HOẠ

- **Môn / Lớp:** Tin học 10 (SGK Kết nối tri thức với cuộc sống, trang 69–72) – Lớp **10/3** – Phòng **H.I.5**
- **Thời gian:** Thứ Hai, ngày **16/11/2026** – Tiết 3 (09h00–09h45) – Thời lượng: **45 phút**
- **Phân phối chương trình:** Tiết PPCT **19** (`Mục 4.13. Bổ sung các đối tượng đồ họa, tiết 1 - LT`; tổng bài 2 tiết: PPCT 19, 20)
- **Đơn vị & Người thực hiện:** Trường THPT Thuận Hoá | GVHD: Thầy Mai Xuân Văn | SVTT: Hồ Quang Huy
- **Hình thức:** Dạy học lí thuyết tại lớp học bình thường (có máy chiếu, bảng phấn, GV thao tác minh hoạ trên Inkscape). *Không dùng các bài tập tr. 74 ở phần Luyện tập của Tiết 1 vì phần đó thuộc Tiết 2 Thực hành.*

---

## I. MỤC TIÊU KIẾN THỨC TRỌNG TÂM (`[KT1]` – `[KT3]`)
- **`[KT1]`**: Nêu được ý nghĩa các thuộc tính cơ bản của 3 nhóm hình khối (Bảng tr. 69–70): Vuông/chữ nhật (`W, H, Rx, Ry`); Tròn/elip (`Rx, Ry, Start, End` + 3 chế độ); Đa giác/sao (`Corners, Rounded, Spoke ratio, Randomized`).
- **`[KT2]`**: Trình bày được cách mở, chức năng 3 trang của hộp thoại `Fill and Stroke` (`Fill`, `Stroke paint`, `Stroke style`), các kiểu tô (`X`, đồng nhất, `Linear gradient`, `Radial gradient`, `Pattern`, dấu `?`), cây bút chỉnh gradient, ống hút sao chép màu và độ trong suốt (`Opacity`, `A`).
- **`[KT3]`**: Phân biệt được 06 phép ghép trong bảng chọn `Path` (`Union`, `Difference`, `Intersection`, `Exclusion`, `Division`, `Cut Path`) và quy tắc **đối tượng lớp dưới luôn bị cắt bởi đối tượng lớp trên**.

---

## II. TIẾN TRÌNH NỘI DUNG 45 PHÚT (CHUẨN `khbd10-3.docx`)

### 1. Hoạt động 1: Mở đầu (05 phút) – Phân tích Miếng dưa hấu (Hình 13.1, SGK tr. 69)
- **Câu hỏi cho HS:**
  1. Hình 13.1 được ghép từ những hình học cơ bản và màu sắc nào?
  2. Theo quy tắc Bài 12 (*vẽ trước nằm dưới, vẽ sau nằm trên*), sắp xếp thứ tự 4 lớp từ dưới lên trên?
- **Đáp án chốt:**
  - Cấu tạo: **03 nửa hình tròn** nhỏ dần (vỏ xanh đậm, cùi trắng/xanh nhạt, ruột đỏ) + **các hình elip/giọt nước nhỏ màu đen** (hạt dưa).
  - Thứ tự **4 lớp từ dưới lên trên:** **1. Vỏ xanh (dưới cùng)** → **2. Cùi trắng** → **3. Ruột đỏ** → **4. Hạt đen (trên cùng)**.
- **Vấn đề dẫn vào bài:** Bài 12 chỉ vẽ được hình tròn kín nguyên vẹn, tô màu bệt và xếp đè lên nhau → Cần học Bài 13 để chỉnh thuộc tính (cắt nửa hình tròn), tô màu nâng cao (`Fill and Stroke`) và ghép cắt hình (`Path`).

### 2. Hoạt động 2: Hình thành kiến thức mới (28 phút)

#### Mục 1. Các đối tượng hình khối và thuộc tính cơ bản (09 phút – `[KT1]`, SGK tr. 69–70)
- **Dẫn nhập trực quan (Ví dụ siêu xe Rolls-Royce):** Tốc độ, dung tích bình xăng, màu sơn, độ bo mui xe, số nan bánh xe... là **thuộc tính** của chiếc xe. Trong Inkscape, mỗi hình cũng có bảng thông số riêng nằm trên **Thanh điều khiển thuộc tính** (thanh ngang phía trên).
- **Bảng thuộc tính 3 nhóm hình cơ bản (Bảng tr. 69–70 & Hình 13.2 tr. 69):**
  1. **Hình vuông, hình chữ nhật (`R`):**
     - `W` (Width): Chiều rộng | `H` (Height): Chiều dài.
     - `Rx, Ry`: Bán kính góc bo tròn *(đặt `Rx = Ry = 0` là góc vuông)*.
  2. **Hình tròn, hình elip (`E`):**
     - `Rx, Ry`: Bán kính ngang, bán kính đứng.
     - `Start, End`: Góc điểm đầu, điểm cuối theo độ *(đặt `Start = 0, End = 180` tạo nửa hình tròn)*.
     - **3 nút chế độ vẽ:** (1) Hình quạt kín (*Slice*), (2) Cung tròn hở (*Arc*), (3) Hình tròn kín nguyên vẹn (*Whole*).
  3. **Hình đa giác, hình sao (`*` – giải thích 4 hình ở Hình 13.2):**
     - 2 nút chuyển chế độ: Đa giác đều / Hình sao.
     - `Corners`: Số đỉnh / số cánh | `Spoke ratio`: Tỉ lệ bán kính trong/ngoài.
     - `Rounded`: Độ cong đỉnh *(tăng lên uốn cong đỉnh nhọn thành cánh hoa)*.
     - `Randomized`: Làm méo ngẫu nhiên các đỉnh.

#### Mục 2. Thiết lập màu tô, màu vẽ – Hộp thoại `Fill and Stroke` (09 phút – `[KT2]`, SGK tr. 70–71)
- **Hoạt động 2 (Hình 13.3 tr. 70 – So sánh 5 hình vuông):** Cả 5 hình đều có viền cam giống nhau nhưng khác màu tô bên trong: (1) Không màu `X`; (2) Màu đồng nhất; (3) Chuyển ngang (`Linear gradient`); (4) Chuyển từ tâm (`Radial gradient`); (5) Chuyển màu có độ trong suốt.
- **3 cách mở hộp thoại `Fill and Stroke`:** `Object / Fill and Stroke` | Nháy chuột phải chọn `Fill and Stroke` | Phím tắt **`Shift + Ctrl + F`**.
- **3 trang chính (Hình 13.4 tr. 70–71):**
  1. **`Fill` (Màu tô nền):** `X` (không màu), Đồng nhất, `Linear gradient` (chuyển thẳng/ngang), `Radial gradient` (chuyển từ tâm), `Pattern` (hoa văn), nút **`?`** (huỷ đặt về trạng thái ban đầu).
  2. **`Stroke paint` (Màu vẽ đường viền):** Chọn màu viền + công cụ **ống hút** sao chép màu.
  3. **`Stroke style` (Kiểu nét vẽ):** Độ dày mỏng (`Width`) + kiểu đường viền (nét liền, **nét đứt**...).
- **Chỉnh nâng cao:** **Biểu tượng cây bút** chỉnh thông số màu chuyển `gradient`; thanh **`Opacity`** và giá trị **`A` (alpha)** chỉnh độ trong suốt.

#### Mục 3. Các phép ghép đối tượng đồ hoạ trong bảng chọn `Path` (10 phút – `[KT3]`, SGK tr. 71–72)
- **Hoạt động 3 (Hình 13.5 tr. 71 – Ghép trái tim):** Xoay hình vuông nghiêng 45°, xếp 2 hình tròn lên 2 cạnh trên rồi chọn **`Path / Union`**.
- **06 phép ghép trong bảng chọn `Path` (Hình 13.6 tr. 72 – Quy ước: Hình vuông lớp DƯỚI, Hình tròn lớp TRÊN):**
  1. **Phép hợp (`Union`, `Ctrl + +`):** Gồm tất cả phần thuộc các hình đơn (hàn thành 1 khối).
  2. **Phép hiệu (`Difference`, `Ctrl + -`):** Phần thuộc hình lớp dưới nhưng không thuộc hình lớp trên.
  3. **Phép giao (`Intersection`, `Ctrl + *`):** Chỉ giữ phần thuộc cả hai hình.
  4. **Phép hiệu đối xứng (`Exclusion`, `Ctrl + ^`):** Phần thuộc các hình trừ phần giao nhau.
  5. **Phép chia (`Division`, `Ctrl + /`):** Hình lớp dưới được chia thành các mảnh bởi đường nét hình lớp trên (còn màu tô).
  6. **Phép cắt (`Cut Path`, `Ctrl + Alt + /`):** Cắt đường viền hình lớp dưới tại giao điểm với hình lớp trên (mất màu tô nền).
- **Quy tắc cần nhớ:** Trong `Difference`, `Division` và `Cut Path`, **ĐỐI TƯỢNG LỚP DƯỚI luôn bị cắt bởi ĐỐI TƯỢNG LỚP TRÊN**.

### 3. Hoạt động 3: Luyện tập (07 phút) – 04 câu hỏi củng cố SGK (tr. 70, 71, 72)
1. **Câu 1 (tr. 70):** Muốn các đỉnh ngôi sao 5 cánh nhọn trở nên cong tròn, sửa tham số nào? → **Đáp án:** `Rounded`.
2. **Câu 2 (Trắc nghiệm 1 tr. 71):** Muốn đổi đường viền thành nét đứt, chọn trang nào trong `Fill and Stroke`? (`A. Fill`, `B. Stroke paint`, `C. Stroke style`, `D. Cả A và B`) → **Đáp án:** `C. Stroke style`.
3. **Câu 3 (Trắc nghiệm 2 tr. 71):** Chỉnh thông số `gradient` chọn biểu tượng nào trong `A, B, C, D`? → **Đáp án:** `B` *(biểu tượng cây bút)*.
4. **Câu 4 (tr. 72 – Hình 13.7 Vẽ đám mây):** Nêu phép ghép và các bước vẽ đám mây Hình 13.7 → **Đáp án:** Dùng **Phép hợp (`Path / Union`, `Ctrl + +`)**, gồm 3 bước: (1) Vẽ các hình tròn xếp chồng lên nhau tạo dáng đám mây; (2) Chọn tất cả các hình tròn; (3) Chọn `Path / Union`.

### 4. Hoạt động 4: Vận dụng & Định hướng Tiết 2 Thực hành (05 phút)
- **Chốt 3 bước vẽ Miếng dưa hấu (Hình 13.1):**
  - Bước 1: Vẽ hình tròn, đặt `Start = 0, End = 180` (chế độ hình quạt) tô xanh đậm làm vỏ dưa (lớp dưới cùng).
  - Bước 2: Nhân bản (`Ctrl + D`) thu nhỏ tô trắng làm cùi; nhân bản tiếp thu nhỏ tô đỏ làm ruột.
  - Bước 3: Vẽ thêm các hình elip nhỏ màu đen đặt ở **lớp trên cùng** làm hạt dưa.
- **Định hướng công cụ cho 03 bài Thực hành Tiết 2 (Tiết PPCT 20, SGK tr. 72–74):**
  - *Nhiệm vụ 1 – Vẽ Logo 3 cánh (H.13.8):* Cung tròn hở (`Start, End`) + tăng `Width` trong `Stroke style` + `Stroke paint`.
  - *Nhiệm vụ 2 – Vẽ Hình cây xanh (H.13.10):* `Path / Difference` (cắt eo thân cây) + `Path / Union` (hàn tán lá).
  - *Nhiệm vụ 3 – Vẽ Hình cầu 3D (H.13.12):* Bỏ viền (`X`) + phối màu chuyển `Linear / Radial gradient` trong `Fill`.

---

## III. DANH SÁCH ẢNH MINH HOẠ SGK (ĐƯỜNG DẪN GITHUB RAW & LOCAL)
Base GitHub Raw: `https://raw.githubusercontent.com/huyhandsome6996/khbd_docx/main/khoi10/_scripts_build/cropped_figures/`
Base Local (từ thư mục `Day_that/`): `../khoi10/_scripts_build/cropped_figures/`

1. `hinh_13_1_2_dua_hau_thuoc_tinh.png` — Hình 13.1 (Miếng dưa hấu) & Hình 13.2 (4 biến thể ngôi sao/đa giác)
2. `bang_13_1_thuoc_tinh_hinh_khoi.png` — Bảng thuộc tính 3 nhóm hình khối (tr. 69–70)
3. `hinh_13_3_4_fill_stroke.png` — Hình 13.3 (5 kiểu tô) & Hình 13.4 (Hộp thoại `Fill and Stroke`)
4. `hinh_13_5_ket_hop_doi_tuong.png` — Hình 13.5 (Ghép trái tim)
5. `hinh_13_6_cac_phep_ghep_hinh.png` — Hình 13.6 (06 phép ghép `Path`)
6. `hinh_13_7_8_dam_may_logo.png` — Hình 13.7 (Vẽ đám mây bằng `Path / Union`)
7. `hinh_13_9_ve_logo.png` — Hình 13.8 & 13.9 (Logo Tiết 2)
8. `hinh_13_10_11_ve_cay.png` — Hình 13.10 & 13.11 (Cây xanh Tiết 2)
9. `hinh_13_12_13_hinh_cau.png` — Hình 13.12 & 13.13 (Quả cầu 3D Tiết 2)
