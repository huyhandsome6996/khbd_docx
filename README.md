# HỆ THỐNG HỌC LIỆU, CẨM NANG QUY CHUẨN & KHO HỒ SƠ THỰC TẬP SƯ PHẠM (DÀNH CHO AI AGENT)

Kho lưu trữ [`c:\Users\PC\khbd_docx`](file:///c:/Users/PC/khbd_docx) được tổ chức thành hệ thống ngăn chuyên biệt (từ `01` đến `09` cùng các ngăn thực hành theo đợt dự giờ: [`khoi10`](file:///c:/Users/PC/khbd_docx/khoi10), [`khoi12`](file:///c:/Users/PC/khbd_docx/khoi12), [`de_cuong`](file:///c:/Users/PC/khbd_docx/de_cuong) và [`Day_that`](file:///c:/Users/PC/khbd_docx/Day_that)), tích hợp toàn bộ văn bản pháp quy, phân phối chương trình chính thức của **Trường THPT Thuận Hoá**, giáo trình lý luận & phương pháp dạy học, 5 cuốn sách giáo khoa, bộ quy chuẩn **42 nhận xét của 03 Giảng viên**, bộ quy tắc **Sinh hoạt chủ nhiệm lớp 12/2**, cẩm nang **Dạy thật tại lớp** và bộ máy Python tạo file Word (`.docx`) tự động.

> [!IMPORTANT]
> **CHỈ DẪN BẮT BUỘC DÀNH CHO MỌI AI AGENT KHI BẮT ĐẦU LÀM VIỆC TRONG KHO LƯU TRỮ NÀY:**
> 1. **Tuyệt đối tuân thủ nguyên tắc "Nói có sách, mách có chứng":** Mọi số tiết, `Tiết PPCT`, tên bài học, yêu cầu cần đạt, số trang sách giáo khoa, hình ảnh minh họa đều phải trích dẫn chính xác từ tài liệu gốc trong Ngăn 01 và Ngăn 03 (PPCT Khối 10: `khoi_10_tin_hoc_07092026.xls`; PPCT Khối 12: `khoi12/khoi_12_tin_hoc_07112026.xls`).
> 2. **Tuyệt đối tuân thủ nguyên tắc "Văn phong con người lịch sự, có đầu có cuối, đầy đủ chủ ngữ – vị ngữ":** Giáo viên hướng dẫn và Giảng viên chấm bài rất nghiêm khắc, cực kỳ ghét lối viết cộc lốc, trống không, thiếu chủ ngữ. Mọi câu văn, kịch bản giao nhiệm vụ của Giáo viên, lời báo cáo của Học sinh hay bài học kinh nghiệm của Sinh viên thực tập đều phải mở đầu bằng chủ ngữ trang trọng, lễ phép (*"Giáo viên..."*, *"Học sinh..."*, *"Em thưa thầy/cô..."*, *"Cô giáo chủ nhiệm Nguyễn Thị Thu Phương..."*).
> 3. **Phân biệt rõ 04 luồng hồ sơ nghiệp vụ theo từng Giáo viên:**
>    - **Luồng 1 – Hồ sơ KHBD đầy đủ chuẩn CV 5512 (Khối 10 & Khối 12):** Tuân thủ 100% [`04_form_mau_khbd/QUY_CHUAN_SOAN_KHBD_5512_THEO_GOP_Y_GV.md`](file:///c:/Users/PC/khbd_docx/04_form_mau_khbd/QUY_CHUAN_SOAN_KHBD_5512_THEO_GOP_Y_GV.md) (đúc kết từ **42 nhận xét của Cô Tuyết Ngân Tôn, Thầy Thống Nhất, Thầy Nguyễn Nhất Tân**).
>    - **Luồng 2 – Đề cương dự giờ kiến tập Khối 12 (nộp Thầy Nguyễn Nhất Tân):** Lưu tại [`de_cuong/`](file:///c:/Users/PC/khbd_docx/de_cuong) — không kẻ bảng 2 cột kịch bản chi tiết lên lớp, mà trình bày gọn mục `d) Cách thức tiến hành:` bằng 01 đoạn văn 4 câu cô đọng.
>    - **Luồng 3 – Hồ sơ DẠY THẬT trực tiếp trên lớp (nộp Thầy Mai Xuân Văn – Khối 10):** Lưu tại [`Day_that/`](file:///c:/Users/PC/khbd_docx/Day_that) — viết KHBD đầy đủ 5512 nhưng **bắt buộc 100% khả thi tại phòng học thường không có máy tính HS** (loại bỏ hoạt động giấy rời rườm rà không thực hiện được), đi kèm file thứ 2 là **Kịch bản lên lớp vui nhộn cầm tay chỉ việc**. Đọc kỹ [`Day_that/README_KINH_NGHIEM_VA_LUU_Y_CHO_AGENT.md`](file:///c:/Users/PC/khbd_docx/Day_that/README_KINH_NGHIEM_VA_LUU_Y_CHO_AGENT.md) trước khi soạn bất kỳ tiết dạy thật nào!
>    - **Luồng 4 – Hồ sơ Dự giờ Sinh hoạt chủ nhiệm (Lớp 12/2):** Tuân thủ 100% [`04_form_mau_khbd/QUY_TAC_SOAN_DE_CUONG_SHCN_12-2.md`](file:///c:/Users/PC/khbd_docx/04_form_mau_khbd/QUY_TAC_SOAN_DE_CUONG_SHCN_12-2.md) (nộp **Cô Nguyễn Thị Thu Phương**).

---

## 1. Sơ đồ Kiến trúc Kho lưu trữ (`khbd_docx`)

```text
c:\Users\PC\khbd_docx\
├── README.md                                       <-- [TÀI LIỆU NÀY] Cẩm nang tổng thể & Bản đồ 09 ngăn cho AI Agent
│
├── 01_chuong_trinh_gdpt/                           <-- [NGĂN 01] Chương trình GDPT 2018 & Phân phối chương trình (PPCT)
│   ├── README.md                                       Quy tắc tra cứu YCCĐ, thứ tự Năng lực (NL chung -> NL đặc thù -> NLS/AI)
│   ├── khoi_10_tin_hoc_07092026.xls                    [GỐC] Kế hoạch dạy học (PPCT) Khối 10 Trường THPT Thuận Hoá (70 tiết)
│   ├── PHAN_PHOI_CHUONG_TRINH_TIN_10_11_12_KNTT.md     Bảng tra cứu Tiết PPCT Lớp 10, 11, 12 & Link Google Drive giải trình
│   └── CHƯƠNG TRÌNH GIÁO DỤC PHỔ THÔNG MÔN TIN HỌC.docx Toàn văn Thông tư 32/2018/TT-BGDĐT môn Tin học
│
├── 02_ly_luan_va_ppdh/                             <-- [NGĂN 02] Lý luận & Phương pháp dạy học Tin học (TS. Nguyễn Thế Dũng)
│   ├── README.md                                       Công thức mục tiêu 4 thành phần, chuỗi 4 hoạt động CV 5512 & Rubric
│   ├── 17_8_Giao trinh LL-PPDH Tin học - NTD.pdf       Giáo trình Lý luận và PPDH Tin học (ĐHSP Huế)
│   └── Chuong1..Chuong6 (*.pptx)                       6 bộ Slide bài giảng Lý luận & PPDH Tin học
│
├── 03_sach_giao_khoa/                              <-- [NGĂN 03] Sách giáo khoa Tin học 10, 11, 12 (Bộ Kết nối tri thức)
│   ├── README.md                                       Mục lục chi tiết từng bài/trang của 5 cuốn SGK & quy tắc cắt hình SGK
│   ├── trich_xuat_trang_sgk.ps1                        Script xuất trang PDF SGK ra ảnh .png để đọc và cắt hình minh họa
│   ├── Tin-học-10-KNTT.pdf                             SGK Tin học 10 (Trang PDF = Trang sách + 1)
│   ├── SGK Tin học 11 - Định hướng KHMT / THƯD.pdf     2 cuốn SGK Tin học 11 (Trang PDF = Trang sách + 2)
│   └── SGK Tin học 12 - Định hướng KHMT / THƯD.pdf     2 cuốn SGK Tin học 12 (Trang PDF = Trang sách + 2)
│
├── 04_form_mau_khbd/                               <-- [NGĂN 04] Trung tâm Quy chuẩn, Biểu mẫu Vàng (.docx) & Engine 5512
│   ├── README.md                                       Hướng dẫn cấu trúc chuẩn Công văn 5512 & cách dùng _engine_5512
│   ├── QUY_CHUAN_SOAN_KHBD_5512_THEO_GOP_Y_GV.md       [BẮT BUỘC ĐỌC] Quy chuẩn 42 nhận xét của 3 Giảng viên cho môn Tin học
│   ├── QUY_TAC_SOAN_DE_CUONG_SHCN_12-2.md              [BẮT BUỘC ĐỌC] Quy chuẩn soạn Đề cương SHCN lớp 12/2 nộp Cô Thu Phương
│   ├── MAU_CHUAN_KHBD_5512_2Tiet_Tin12_Bai9.docx       Biểu mẫu Vàng #1: KHBD bài 02 tiết gộp chuẩn CV 5512
│   ├── MAU_CHUAN_KHBD_5512_1Tiet_Tin10_Bai12_Tiet1.docx Biểu mẫu Vàng #2: KHBD bài 01 tiết rời chuẩn CV 5512
│   ├── MAU_CHUAN_MA_TRAN_DAC_TA_DE_KIEM_TRA_7991_Tin10_GHKI.docx Biểu mẫu Vàng #3: Ma trận, Đặc tả & Đề kiểm tra CV 7991
│   ├── MAU_CHUAN_DE_CUONG_SHCN_Lop12-2.docx            Biểu mẫu Vàng #4: Đề cương dự giờ Sinh hoạt chủ nhiệm lớp 12/2
│   ├── _engine_5512/                                   Bộ máy Python (docx_engine_5512.py + base templates) sinh file .docx
│   └── khbd_kntt_tham_khao/                            Bộ KHBD tham khảo cả năm Lớp 10, 11, 12 của NXB Giáo dục Việt Nam
│
├── 05_ke_hoach_bai_day_output/                     <-- [NGĂN 05] Kho lưu trữ bản văn bản Markdown (.md) & Checklist 15 tiêu chí QA
│   ├── README.md                                       Danh mục 11 tệp .md (Tin 10, Tin 12, SHCN) & Master QA Checklist
│   └── *.md                                            Bản văn bản .md đồng bộ 1-1 với các tệp .docx thành phẩm
│
├── 06_thanh_pham_docx/                             <-- [NGĂN 06] Kho chứa toàn bộ tệp Word (.docx) thành phẩm sẵn sàng in ấn
│   ├── README.md                                       Bảng phân luồng 11 tệp .docx Tin học (10 & 12) + 02 tệp .docx SHCN
│   └── *.docx                                          Các tệp Word thành phẩm (Trang bìa 1 trang, A4 dọc, đánh số trang)
│
├── 08_goi_nop_du_gio/                              <-- [NGĂN 08] Gói đóng gói (.zip) Hồ sơ dự giờ Khối 10
│   ├── README.md
│   ├── HoQuangHuy_24S1020030_HoSo_DuGio_Tin10/
│   └── HoQuangHuy_24S1020030_HoSo_DuGio_Tin10.zip
│
├── 09_nop_du_gio_Tin12_Bai9_05-11-2026/            <-- [NGĂN 09] Gói Hồ sơ dự giờ Tin học 12 – Bài 9 (Lớp 12/4 – 05/11/2026)
│   ├── README.md
│   ├── Ho_so_nop/                                      Chứa De_cuong_Tin12_Bai9_Lop12-4.docx, Phiếu HT, Học liệu HTML & Đối chiếu
│   ├── _scripts_build/                                 Script build_revised_5512.py & 8 hình cắt SGK (Hình 9.1 – 9.8)
│   └── HoQuangHuy_24S1020030_DuGio_Tin12_Bai9_05-11-2026.zip
│
├── khoi10/                                         <-- [NGĂN KHOI10] Ngăn riêng biệt chứa trọn bộ Hồ sơ Khối 10 & Cẩm nang giải trình
│   ├── README_NGUON_VA_GIAI_TRINH_KHOI10.md            Cẩm nang giải trình "Nói có sách, mách có chứng" cho Khối 10
│   ├── PHAN_PHOI_CHUONG_TRINH_TIN_10_11_12_KNTT.md     Bản sao tra cứu nhanh PPCT tại ngăn khoi10
│   ├── khoi_10_tin_hoc_07092026.xls                    Bản gốc Kế hoạch dạy học Khối 10 Trường THPT Thuận Hoá
│   ├── 01_..09_*.docx                                  09 tệp .docx thành phẩm Khối 10 (Tiết PPCT 15 -> 20, cả bản 1 tiết & 2 tiết)
│   ├── _scripts_build/                                 Bộ script Python & 19 hình cắt SGK Tin 10 (docx_engine_5512.py, build_*.py)
│   └── HoQuangHuy_24S1020030_TronBo_HoSo_Khoi10.zip    Tệp nén trọn bộ hồ sơ Khối 10
│
├── khoi12/                                         <-- [NGĂN KHOI12] Ngăn riêng chứa trọn bộ Hồ sơ Khối 12 (Bài 10 & SHCN)
│   ├── README_NGUON_VA_GIAI_TRINH_KHOI12.md            Cẩm nang nguồn PPCT Khối 12 (Bài 10 – Tiết PPCT: 21, 22)
│   ├── khoi_12_tin_hoc_07112026.xls                    Bản gốc Kế hoạch dạy học Khối 12 Trường THPT Thuận Hoá
│   ├── 03_..04_*.docx & de_cuong_du_gio_12-*.docx      Bản gộp Bài 10, Phiếu HT & bản đồng bộ đề cương 12/1, 12/2
│   ├── Hoc_lieu_HTML/                                  Bộ tệp mã nguồn HTML chuẩn SGK Bài 9–10
│   └── _scripts_build/                                 Bộ script Python & hình cắt SGK Tin 12
│
├── de_cuong/                                       <-- [NGĂN ĐỀ CƯƠNG] Đề cương dự giờ kiến tập Khối 12 (theo góp ý Thầy Nguyễn Nhất Tân)
│   ├── README.md                                       Mô tả cấu trúc đề cương dự giờ kiến tập (12/4, 12/1, 12/2)
│   ├── de_cuong_du_gio_12-4.docx                       Lớp 12/4 (H.II.4) – Bài 9: Tạo danh sách, bảng (Tiết PPCT: 19, 20)
│   ├── de_cuong_du_gio_12-1.docx                       Lớp 12/1 (H.II.1) – Bài 10: Tạo liên kết [Tiết 2 – TH] (Tiết PPCT: 22)
│   ├── de_cuong_du_gio_12-2.docx                       Lớp 12/2 (H.II.2) – Bài 10: Tạo liên kết [Tiết 1 – LT] (Tiết PPCT: 21)
│   └── _scripts_build/build_de_cuong_khoi12.py         Script build đồng thời 3 đề cương & đồng bộ sang khoi12/
│
└── Day_that/                                       <-- [NGĂN DẠY THẬT] Hồ sơ đứng lớp giảng dạy trực tiếp (nộp Thầy Mai Xuân Văn)
    ├── README_KINH_NGHIEM_VA_LUU_Y_CHO_AGENT.md        [BẮT BUỘC ĐỌC] Bài học phân biệt Thầy Tân vs Thầy Văn & thiết kế 100% khả thi
    ├── 01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx  KHBD 5512 đầy đủ khả thi 100% (Bài 13 Tiết 1 – Lớp 10/3 – H.I.5 – 16/11/2026)
    ├── 02_Kich_ban_len_lop_chi_tiet_vui_nhon_Tin10_Bai13_Tiet1.docx  Cẩm nang cầm tay nói gì – bấm gì – chữa cháy gì từng phút
    └── _scripts_build/build_day_that_10_3.py           Script build bộ đôi hồ sơ Dạy thật & kiểm tra XML
```

---

## 2. Thông tin Hành chính Cố định & Lịch Dự giờ Thực tập của Sinh viên Hồ Quang Huy

Mọi tài liệu được tạo ra trong kho lưu trữ này đều thuộc hồ sơ Thực tập sư phạm của:
- **Sinh viên thực hiện (Giáo viên giảng dạy):** **Hồ Quang Huy** – Mã sinh viên: `24S1020030` (Ngành Sư phạm Tin học, Khoa Tin học, Trường Đại học Sư phạm – Đại học Huế).
- **Trường thực tập:** **Trường THPT Thuận Hoá** (TP. Huế).
- **Giáo viên hướng dẫn giảng dạy môn Tin học (Tổ Tin học):** **Thầy Mai Xuân Văn**.
- **Giáo viên hướng dẫn công tác chủ nhiệm (Lớp 12/2):** **Cô Nguyễn Thị Thu Phương** (Sinh hoạt chủ nhiệm cố định vào **Tiết 5 sáng Thứ Bảy hằng tuần**; danh sách Ban cán sự và 4 tổ trưởng lớp 12/2 xem tại [`04_form_mau_khbd/QUY_TAC_SOAN_DE_CUONG_SHCN_12-2.md`](file:///c:/Users/PC/khbd_docx/04_form_mau_khbd/QUY_TAC_SOAN_DE_CUONG_SHCN_12-2.md)).

### Bảng Lịch Dự giờ Giảng dạy Môn Tin học (Khối 10 & Khối 12)

| Ngày dự giờ | Thứ | Tiết TKB | Khung giờ | Lớp | Phòng | Bài dạy theo lịch đăng ký | Tiết PPCT chuẩn | Hồ sơ thành phẩm tương ứng |
| :---: | :---: | :---: | :---: | :---: | :---: | :--- | :---: | :--- |
| `02/11/2026` | Thứ Hai | Tiết 2 | `08h00 – 08h45` | `10/2` | `H.I.4` | **Ôn tập kiểm tra Giữa học kì I** | **Tiết 15** | [`khoi10/01_De_cuong_du_gio_Tin10_OnTap_GHKI_TietPPCT15_Lop10-2.docx`](file:///c:/Users/PC/khbd_docx/khoi10/01_De_cuong_du_gio_Tin10_OnTap_GHKI_TietPPCT15_Lop10-2.docx) |
| *(Định kì)* | — | — | *(45 phút)* | Khối 10 | — | **Kiểm tra Giữa học kì I** *(Ma trận, Đặc tả & Đề KT)* | **Tiết 16** | [`khoi10/02_Ma_tran_Ban_dac_ta_De_kiem_tra_GHKI_Tin10_TietPPCT16.docx`](file:///c:/Users/PC/khbd_docx/khoi10/02_Ma_tran_Ban_dac_ta_De_kiem_tra_GHKI_Tin10_TietPPCT16.docx) |
| `05/11/2026` | Thứ Năm | Tiết 3 | `09h00 – 09h45` | `12/4` | `H.II.4` | **Bài 9. Tạo danh sách, bảng** *(Dự giờ Tiết 1)* | **Tiết 19, 20** | [`de_cuong/de_cuong_du_gio_12-4.docx`](file:///c:/Users/PC/khbd_docx/de_cuong/de_cuong_du_gio_12-4.docx) *(kèm [`09_nop_du_gio_Tin12_Bai9_05-11-2026`](file:///c:/Users/PC/khbd_docx/09_nop_du_gio_Tin12_Bai9_05-11-2026/Ho_so_nop/De_cuong_Tin12_Bai9_Lop12-4.docx))* |
| `10/11/2026` | Thứ Ba | Tiết 1 | `07h10 – 07h55` | `10/1` | `H.I.3` | **Bài 12. Phần mềm thiết kế đồ hoạ (Tiết 1 – LT)** | **Tiết 17** | [`khoi10/03_De_cuong_du_gio_Tin10_Bai12_Tiet1_LT_TietPPCT17_Lop10-1.docx`](file:///c:/Users/PC/khbd_docx/khoi10/03_De_cuong_du_gio_Tin10_Bai12_Tiet1_LT_TietPPCT17_Lop10-1.docx) *(kèm bản gộp 2 tiết số `05`)* |
| `10/11/2026` | Thứ Ba | Tiết 2 | `08h00 – 08h45` | `10/2` | `H.I.4` | **Bài 12. Phần mềm thiết kế đồ hoạ (Tiết 2 – TH)** | **Tiết 18** | [`khoi10/04_De_cuong_du_gio_Tin10_Bai12_Tiet2_TH_TietPPCT18_Lop10-2.docx`](file:///c:/Users/PC/khbd_docx/khoi10/04_De_cuong_du_gio_Tin10_Bai12_Tiet2_TH_TietPPCT18_Lop10-2.docx) *(kèm bản gộp 2 tiết số `05`)* |
| `13/11/2026` | Thứ Sáu | Tiết 2 | `08h00 – 08h45` | `12/1` | `H.II.1` | **Bài 10. Tạo liên kết (Tiết 2 – Thực hành)** | **Tiết 22** | [`de_cuong/de_cuong_du_gio_12-1.docx`](file:///c:/Users/PC/khbd_docx/de_cuong/de_cuong_du_gio_12-1.docx) *(đồng bộ tại [`khoi12`](file:///c:/Users/PC/khbd_docx/khoi12/de_cuong_du_gio_12-1.docx))* |
| `13/11/2026` | Thứ Sáu | Tiết 3 | `09h00 – 09h45` | `12/2` | `H.II.2` | **Bài 10. Tạo liên kết (Tiết 1 – Lí thuyết)** | **Tiết 21** | [`de_cuong/de_cuong_du_gio_12-2.docx`](file:///c:/Users/PC/khbd_docx/de_cuong/de_cuong_du_gio_12-2.docx) *(đồng bộ tại [`khoi12`](file:///c:/Users/PC/khbd_docx/khoi12/de_cuong_du_gio_12-2.docx))* |
| `16/11/2026` | Thứ Hai | Tiết 3 | `09h00 – 09h45` | `10/3` | `H.I.5` | **Bài 13. Bổ sung các đối tượng đồ hoạ (Tiết 1 – LT)** *(DẠY THẬT)* | **Tiết 19** *(Tiết 2 TH là 20)* | [`Day_that/01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx`](file:///c:/Users/PC/khbd_docx/Day_that/01_KHBD_Tin10_Bai13_Tiet1_Lop10-3_DayThat.docx) & [`02_Kich_ban...docx`](file:///c:/Users/PC/khbd_docx/Day_that/02_Kich_ban_len_lop_chi_tiet_vui_nhon_Tin10_Bai13_Tiet1.docx) *(bộ mẫu chung tại [`khoi10`](file:///c:/Users/PC/khbd_docx/khoi10/06_De_cuong_du_gio_Tin10_Bai13_Tiet1_LT_TietPPCT19_Lop10-3.docx))* |

---

## 3. Tóm tắt 10 Kinh nghiệm "Xương máu" Đúc kết từ 42 Nhận xét của Giảng viên (Dành cho AI Agent)

Khi soạn bất kỳ bài dạy hay đề cương nào tiếp theo, Agent tuyệt đối không được lặp lại các lỗi cũ mà phải áp dụng ngay 10 quy tắc vàng sau (chi tiết tại [`04_form_mau_khbd/QUY_CHUAN_SOAN_KHBD_5512_THEO_GOP_Y_GV.md`](file:///c:/Users/PC/khbd_docx/04_form_mau_khbd/QUY_CHUAN_SOAN_KHBD_5512_THEO_GOP_Y_GV.md)):

1. **Tên tài liệu chuẩn:** Luôn ghi đầy đủ **`ĐỀ CƯƠNG DỰ GIỜ – KẾ HOẠCH BÀI DẠY`** kèm rõ **`Thời lượng` + `Tiết thứ mấy của bài` + `Tiết PPCT`** ngay dưới tên bài trên cả Trang bìa và Đầu trang 1.
2. **Đặc thù PPCT Khối 10 Trường THPT Thuận Hoá (`khoi_10_tin_hoc_07092026.xls`):** Trường dạy theo định hướng ICT nên trong Chủ đề 1 **chỉ dạy Bài 1 (`§1.1`), Bài 2 (`§1.2`), Bài 7 (`§1.7`) và BỎ QUA Bài 3, 4, 5, 6**. Tuyệt đối không đưa kiến thức hệ nhị phân, phép toán logic `AND/OR/NOT`, bảng mã ASCII/Unicode vào bài Ôn tập hay Đề kiểm tra Giữa kì I của Khối 10!
3. **Thứ tự Năng lực tại Mục `I.2`:** Bắt buộc đặt **`a) Năng lực chung` lên trước**, rồi mới đến **`b) Năng lực đặc thù môn Tin học` (`NLa` – `NLe`)**, và cuối cùng là **`c) Năng lực số và Năng lực AI`**.
4. **Mã hóa và khớp nối 1–1 Mục tiêu:** Mọi mã `[KT1..n]`, `[NLC1..3]`, `[NLa..e]`, `[NLS/NLAI]`, `[PC1..3]` ở Mục `I` phải được trích dẫn và giải quyết hết 100% tại mục `a) Mục tiêu` của các Hoạt động / Nhiệm vụ.
5. **Đúng 04 Hoạt động chuẩn Công văn 5512 tại Mục `III`:**
   - Không có tiêu đề thừa `Các hoạt động dạy học`.
   - Chỉ có đúng 4 hoạt động: `1. Hoạt động 1: Mở đầu`, `2. Hoạt động 2: Hình thành kiến thức mới`, `3. Hoạt động 3: Luyện tập`, `4. Hoạt động 4: Vận dụng`.
   - Toàn bộ các đơn vị kiến thức mới phải **gộp chung vào duy nhất `Hoạt động 2`**, chia nhỏ bên trong thành `2.1. Nhiệm vụ 1`, `2.2. Nhiệm vụ 2`, `2.3. Nhiệm vụ 3`...
   - Với bài dạy 2 tiết gộp, không đưa `Luyện tập` vào Tiết 1 khi chưa dạy xong toàn bộ `Hình thành kiến thức mới`.
6. **Cấu trúc ngoài bảng & trong bảng của mỗi Hoạt động / Nhiệm vụ:**
   - Đặt `a) Mục tiêu`, `b) Nội dung`, `c) Sản phẩm` **ở NGOÀI bảng**.
   - Chỉ đặt **Bảng 2 cột** ở mục `d) Tổ chức thực hiện`:
     - Cột 1: **`HOẠT ĐỘNG CỦA GIÁO VIÊN VÀ HỌC SINH`** (viết đúng tên 4 bước: `Bước 1. Giáo viên chuyển giao nhiệm vụ học tập`, `Bước 2. Học sinh thực hiện nhiệm vụ`, `Bước 3. Báo cáo sản phẩm`, `Bước 4. Đánh giá, nhận xét`).
     - Cột 2: **`SẢN PHẨM DỰ KIẾN (NỘI DUNG GHI BÀI)`**.
7. **Nhúng trực tiếp Hình ảnh / Bảng SGK vào `b) Nội dung`:** Cắt (crop) sắc nét mọi hình ảnh/bảng biểu SGK được nhắc đến trong nhiệm vụ và chèn trực tiếp vào `.docx` (kèm chú thích hình).
8. **Ghi chú Công cụ đánh giá (`CCĐG`) & Phụ lục Rubric 6 cột:**
   - Cuối mục `c) Sản phẩm` của mọi hoạt động/nhiệm vụ phải ghi rõ tên Công cụ kiểm tra đánh giá (`CCĐG 1..n`) đính kèm tại `IV. PHỤ LỤC`.
   - Xóa bỏ mục `KẾT QUẢ DỰ KIẾN` ở cuối bài; đặt `IV. PHỤ LỤC` và `V. NHẬN XÉT CỦA GIÁO VIÊN HƯỚNG DẪN VÀ KÝ DUYỆT`.
   - Bảng Rubric tại Phụ lục phải có khung thông tin hành chính ở đầu và đủ 6 cột (cột `STT` rộng tối thiểu `650` dxa để không bị rớt chữ).
9. **Trang bìa đúng 01 trang & Đánh số trang chân trang:**
   - Trang bìa phải gói gọn trong đúng Trang 1 (bỏ qua các đoạn trống thừa `idx in (11,12,13,14,15)` trong template gốc).
   - Phần nội dung chính (Section 2) bắt buộc đánh số trang căn giữa ở Footer bắt đầu từ `1`.
10. **Kiểm tra trực quan (`x2t.exe` + `pdftoppm`) trước khi bàn giao:** Luôn xuất `.docx` ra `.pdf` và `.png` trong thư mục `scratch` để kiểm tra bằng mắt thường trang bìa, bảng biểu, hình ảnh và chữ ký cuối bài trước khi gửi người dùng.


---

## 4. Kinh nghiệm Thực chiến Phân luồng theo Giảng viên & Kỹ thuật Dạy thật (`Day_that/`)

Đọc chi tiết tại [`Day_that/README_KINH_NGHIEM_VA_LUU_Y_CHO_AGENT.md`](file:///c:/Users/PC/khbd_docx/Day_that/README_KINH_NGHIEM_VA_LUU_Y_CHO_AGENT.md):

1. **Không dùng nhầm form giữa hai thầy:**
   - **Thầy Nguyễn Nhất Tân (`de_cuong/` – Dự giờ kiến tập Khối 12):** Dùng Đề cương cô đọng, không kẻ bảng 2 cột `Bước 1..4` chi tiết; mục `d) Cách thức tiến hành:` viết thành **01 đoạn văn 4 câu ngắn gọn**.
   - **Thầy Mai Xuân Văn (`Day_that/` – Đứng lớp DẠY THẬT Khối 10):** Dùng **KHBD đầy đủ 5512** (bảng 2 cột `Bước 1..4`), nhưng mọi hoạt động **phải khả thi 100%** tại phòng học thường (`H.I.5`, không có máy tính HS) với học sinh trung bình/yếu (tra bảng SGK, giơ ngón tay chọn trắc nghiệm, nhìn hình đoán 1 từ khoá, GV demo 30s–1 phút trên máy chiếu — không bịa phát 4 phiếu giấy rời rườm rà).
2. **Xuất Bộ đôi Hồ sơ khi Dạy thật:** Luôn đi kèm `01_KHBD_...docx` (nộp GVHD) và `02_Kich_ban_len_lop_chi_tiet_vui_nhon_...docx` (cẩm nang nói gì, bấm gì, reset Inkscape tránh tàng hình `Opacity=0%`/nửa vòng `Start/End`, liên hệ Bài 11–12 & Apple/CNC/AI `.svg`, chữa cháy khi lớp im re).
3. **Bẫy kỹ thuật `docx_engine_5512.py` mới ghi nhận:** Không lồng `` `...` `` trong `**...**`/`*...*`; không trải `**...**` qua `\n` trong ô bảng; đặt `<w:keepNext/>` ở hàng tiêu đề bảng (`idx_r == 0`) và chỉ đặt `<w:cantSplit/>` cho hàng ngắn (`< 550` ký tự) để không tạo nửa trang trắng.
