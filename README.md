# HỆ THỐNG HỌC LIỆU & QUY TRÌNH SOẠN KẾ HOẠCH BÀI DẠY MÔN TIN HỌC THPT (DÀNH CHO AI AGENT)

Kho lưu trữ (`khbd_docx`) này được cấu trúc hóa thành **5 ngăn (compartments)** chuyên biệt dựa trên toàn bộ tài liệu pháp quy, giáo trình lý luận & phương pháp dạy học, sách giáo khoa và biểu mẫu chuẩn do người dùng cung cấp.

Mục đích của kiến trúc này là giúp **AI Agent** định vị tức thì tài liệu nguồn, hiểu rõ chuẩn mực sư phạm và tự động biên soạn **Kế hoạch bài dạy (Đề cương bài dạy)** môn Tin học cấp THPT (Lớp 10, 11, 12) đúng chuẩn 100%.

---

## 1. Sơ đồ Phân hoạch 5 Ngăn của Kho lưu trữ (Repository Architecture)

```text
khbd_docx/
├── README.md                                   <-- [TÀI LIỆU NÀY] Bản đồ tổng thể & Quy trình 5 bước cho Agent
├── .gitattributes                              <-- Cấu hình Git LFS cho các tệp PDF dung lượng lớn (>100 MB)
│
├── 01_chuong_trinh_gdpt/                       <-- [NGĂN 01] Chương trình GDPT 2018 (Căn cứ pháp lý & YCCĐ)
│   ├── README.md                                   Hướng dẫn tra cứu YCCĐ, mã Năng lực (NLa–NLe) & Phẩm chất
│   └── CHƯƠNG TRÌNH GIÁO DỤC PHỔ THÔNG MÔN TIN HỌC.docx
│
├── 02_ly_luan_va_ppdh/                         <-- [NGĂN 02] Lý luận & Phương pháp dạy học Tin học
│   ├── README.md                                   Hướng dẫn áp dụng công thức mục tiêu, chuỗi hoạt động, thực hành & AI
│   ├── 17_8_Giao trinh LL-PPDH Tin học - NTD.pdf   Giáo trình Lý luận và PPDH Tin học (TS. Nguyễn Thế Dũng)
│   ├── Chương 1 - Lý luận và Phương pháp Dạy học Tin học.pptx
│   ├── Chuong2_DinhHuongDayHocMonTinHoc.pptx
│   ├── Chuong3_PhuongPhapDayHocMonTinHoc.pptx
│   ├── Chuong4_HinhThucToChucVaKyThuatDayHocMonTinHoc.pptx
│   ├── Chuong5_ThietKeKeHoachDayHocMonTinHoc.pptx
│   └── Chuong6_ToChucDayHocTinHoc.pptx
│
├── 03_sach_giao_khoa/                          <-- [NGĂN 03] Sách giáo khoa Tin học 10, 11, 12 (Bộ KNTT)
│   ├── README.md                                   Mục lục chi tiết từng bài/trang của 5 cuốn SGK & cách trích xuất
│   ├── trich_xuat_trang_sgk.ps1                    Script xuất trang PDF (>100MB) ra ảnh .png để Agent đọc trực tiếp
│   ├── Tin-học-10-KNTT.pdf
│   ├── SGK Tin học 11 - Định hướng khoa học máy tính-mau.pdf
│   ├── SGK Tin học 11 - Định hướng tin học ứng dụng-mau.pdf
│   ├── SGK Tin học 12 - Định hướng Khoa học máy tính-mau.pdf
│   └── SGK Tin học 12 - Định hướng Tin học ứng dụng-mau.pdf
│
├── 04_form_mau_khbd/                           <-- [NGĂN 04] Form mẫu Kế hoạch bài dạy (Chuẩn cấu trúc & định dạng)
│   ├── README.md                                   Đặc tả cấu trúc 5 phần, 4 bảng biểu và thông số định dạng .docx
│   └── Đề cương bài dạy Tin12_Bai8_Tiet2 (sv Phan Trần Hồng Anh).docx
│
├── 05_ke_hoach_bai_day_output/                 <-- [NGĂN 05] Nơi lưu trữ bản thảo (.md) Kế hoạch bài dạy do Agent biên soạn
│   └── README.md                                   Quy ước đặt tên file, cách xuất .docx và Checklist 10 tiêu chí QA
│
└── 06_thanh_pham_docx/                         <-- [NGĂN 06] MỤC RIÊNG CHỨA THÀNH PHẨM FILE WORD (.DOCX) ĐỂ TẢI VỀ DÙNG NGAY
    ├── README.md                                   Hướng dẫn & phân luồng tài liệu cho AI Agent
    ├── Đề cương bài dạy Tin10_OntapGK1.docx        [Tin 10] KHBD Ôn tập GK1
    ├── Ma-tran-Ban-dac-ta-De-KT-GK1_Tin10.docx     [Tin 10] Ma trận & Đặc tả Đề kiểm tra GK1
    ├── Đề cương bài dạy Tin10_Bai12_Tiet1.docx     [Tin 10] KHBD Bài 12 - Tiết 1
    ├── Đề cương bài dạy Tin10_Bai12_Tiet2.docx     [Tin 10] KHBD Bài 12 - Tiết 2
    ├── Đề cương bài dạy Tin10_Bai13_Tiet1.docx     [Tin 10] KHBD Bài 13 - Tiết 1
    ├── Đề cương bài dạy Tin10_Bai13_Tiet2.docx     [Tin 10] KHBD Bài 13 - Tiết 2
    ├── Đề cương dự giờ SHCN_Tuan8_12-2_HoQuangHuy.docx  [Hồ sơ thực tập] Đề cương dự giờ Sinh hoạt chủ nhiệm lớp 12/2
    └── De_cuong_du_gio_SHCN_Tuan8_12-2_HoQuangHuy.md    [Bản Markdown] Kèm chỉ dẫn loại trừ dành cho AI Agent
```

> ⚠️ **LƯU Ý DÀNH CHO CODING AGENT:** File `Đề cương dự giờ SHCN_Tuan8_12-2_HoQuangHuy.docx` và `De_cuong_du_gio_SHCN_Tuan8_12-2_HoQuangHuy.md` thuộc hồ sơ nghiệp vụ chủ nhiệm lớp (Thực tập sư phạm I), **KHÔNG THUỘC KẾ HOẠCH BÀI DẠY MÔN TIN HỌC**. Agent khi biên soạn giáo án Tin học tuyệt đối không dùng tài liệu này làm dữ liệu huấn luyện hoặc form mẫu.

---

## 2. Lịch sử Nguồn gốc Học liệu (Commit Lineage)

Toàn bộ tài liệu trong 5 ngăn được phân hoạch từ các mốc cung cấp học liệu của người dùng trên nhánh `main`:
1. **Học liệu Sách giáo khoa (`b9ab453`):** 5 cuốn SGK Tin học Lớp 10, Lớp 11 (KHMT & THƯD), Lớp 12 (KHMT & THƯD) – bộ *Kết nối tri thức với cuộc sống* $\rightarrow$ Phân vào **Ngăn 03 (`03_sach_giao_khoa/`)**.
2. **Biểu mẫu Kế hoạch bài dạy chuẩn (`db9b450`):** `Đề cương bài dạy Tin12_Bai8_Tiet2 (sv Phan Trần Hồng Anh).docx` $\rightarrow$ Phân vào **Ngăn 04 (`04_form_mau_khbd/`)**.
3. **Hệ thống Slide & Giáo trình Lý luận và PPDH Tin học (`32ec8b7`, `9335040`):** 6 chương Slide (`Chương 1` $\rightarrow$ `Chương 6`) và Giáo trình PDF `17_8_Giao trinh LL-PPDH Tin học - NTD.pdf` $\rightarrow$ Phân vào **Ngăn 02 (`02_ly_luan_va_ppdh/`)**.
4. **Chương trình Giáo dục phổ thông môn Tin học (`5dd11af`):** Văn bản Chương trình GDPT 2018 môn Tin học (`CHƯƠNG TRÌNH GIÁO DỤC PHỔ THÔNG MÔN TIN HỌC.docx`) $\rightarrow$ Phân vào **Ngăn 01 (`01_chuong_trinh_gdpt/`)**.

---

## 3. Quy trình 5 Bước Bắt Buộc Dành Cho AI Agent Khi Soạn Kế Hoạch Bài Dạy

Khi người dùng yêu cầu: *"Hãy soạn cho tôi Kế hoạch bài dạy / Đề cương bài dạy môn Tin học lớp [X], Bài [Y], Tiết [Z]..."*, AI Agent **BẮT BUỘC** thực hiện tuần tự 5 bước sau:

### Bước 1: Đọc Yêu cầu cần đạt & Mã năng lực tại Ngăn 01 (`01_chuong_trinh_gdpt/`)
* Mở [`01_chuong_trinh_gdpt/README.md`](01_chuong_trinh_gdpt/README.md) để xác định mạch kiến thức (`DL`, `ICT`, `CS`), Chủ đề (`A` $\rightarrow$ `G`), các **Yêu cầu cần đạt (YCCĐ)** bắt buộc của bài học và chọn đúng mã **Năng lực đặc thù (`NLa` – `NLe`)**, **Năng lực chung**, **Phẩm chất**.

### Bước 2: Tra cứu Mục lục và Trích xuất trang SGK tại Ngăn 03 (`03_sach_giao_khoa/`)
* Mở [`03_sach_giao_khoa/README.md`](03_sach_giao_khoa/README.md) để tra cứu số trang của Bài `[Y]` trong cuốn SGK tương ứng.
* Chạy script [`03_sach_giao_khoa/trich_xuat_trang_sgk.ps1`](03_sach_giao_khoa/trich_xuat_trang_sgk.ps1) để trích xuất đúng các trang PDF của bài học đó ra ảnh `.png` (giúp vượt qua giới hạn 100 MB của công cụ đọc file trực tiếp), sau đó dùng `read_file` đọc kỹ từng trang sách để lấy chính xác tên các hoạt động, bảng biểu, hình ảnh, hộp kiến thức và bài tập trong SGK.

### Bước 3: Áp dụng Nguyên tắc Sư phạm & Thiết kế Hoạt động tại Ngăn 02 (`02_ly_luan_va_ppdh/`)
* Mở [`02_ly_luan_va_ppdh/README.md`](02_ly_luan_va_ppdh/README.md) để áp dụng:
  1. **Công thức viết mục tiêu 4 thành phần** (Động từ hành động quan sát/đo lường được + Nội dung + Tiêu chí).
  2. Xây dựng một **Tình huống học tập thực tiễn** xuyên suốt tiết dạy.
  3. Thiết kế chuỗi 4–5 hoạt động học tập liên kết, tăng tiến (*Scaffolding*), đảm bảo **tối giản lý thuyết – tối đa thực hành** và tích hợp hoạt động **sử dụng AI kiểm tra mã / phát hiện lỗi có kiểm chứng**.

### Bước 4: Định dạng theo đúng Form mẫu chuẩn tại Ngăn 04 (`04_form_mau_khbd/`)
* Mở [`04_form_mau_khbd/README.md`](04_form_mau_khbd/README.md) để biên soạn đầy đủ 5 phần (`I. MỤC TIÊU`, `II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU`, `III. TIẾN TRÌNH DẠY HỌC`, `IV. KIỂM TRA, ĐÁNH GIÁ`, `V. KẾT QUẢ DỰ KIẾN`) và 4 bảng biểu chuẩn (Bảng kiến thức tiết trước, Bảng phân bố thời gian 45 phút, Bảng chính 4 cột các hoạt động dạy học, Bảng Rubric tiêu chí đánh giá 10 điểm).

### Bước 5: Xuất file kết quả và Kiểm duyệt tại Ngăn 05 (`05_ke_hoach_bai_day_output/`)
* Lưu file Kế hoạch bài dạy hoàn chỉnh (`.md` và/hoặc `.docx` theo yêu cầu người dùng) vào [`05_ke_hoach_bai_day_output/`](05_ke_hoach_bai_day_output/) và đối chiếu **Checklist 10 tiêu chí QA** trong [`05_ke_hoach_bai_day_output/README.md`](05_ke_hoach_bai_day_output/README.md) trước khi bàn giao.
