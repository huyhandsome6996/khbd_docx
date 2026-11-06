# NGĂN 03: SÁCH GIÁO KHOA TIN HỌC 10, 11, 12 (BỘ KẾT NỐI TRI THỨC VỚI CUỘC SỐNG)

> **Vai trò của ngăn này đối với AI Agent:** Cung cấp toàn bộ **ngữ liệu gốc, bảng biểu, hình ảnh minh họa, hoạt động khám phá, hộp kiến thức, bài tập thực hành, luyện tập và vận dụng** của 5 cuốn Sách giáo khoa Tin học cấp THPT (Lớp 10, Lớp 11 KHMT/THƯD, Lớp 12 KHMT/THƯD) thuộc bộ sách *Kết nối tri thức với cuộc sống* (NXB Giáo dục Việt Nam).

---

## 1. Danh mục 5 cuốn Sách giáo khoa & Công cụ trong Ngăn 03

| STT | Tên file | Khối lớp & Định hướng | Dung lượng | Tổng số trang PDF | Công thức quy đổi số trang |
| :-: | :--- | :--- | :---: | :---: | :--- |
| 1 | [`Tin-học-10-KNTT.pdf`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/Tin-h%E1%BB%8Dc-10-KNTT.pdf) | **Lớp 10** (Chung cả 2 định hướng ICT & CS) | `27.37 MB` | 170 trang | **`Trang PDF = Trang sách + 1`** |
| 2 | [`SGK Tin học 11 - Định hướng khoa học máy tính-mau.pdf`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/SGK%20Tin%20h%E1%BB%8Dc%2011%20-%20%C4%90%E1%BB%8Bnh%20h%C6%B0%E1%BB%9Bng%20khoa%20h%E1%BB%8Dc%20m%C3%A1y%20t%C3%ADnh-mau.pdf) | **Lớp 11 – KHMT (`CS`)** | `105.50 MB` | 151 trang | **`Trang PDF = Trang sách + 2`** |
| 3 | [`SGK Tin học 11 - Định hướng tin học ứng dụng-mau.pdf`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/SGK%20Tin%20h%E1%BB%8Dc%2011%20-%20%C4%90%E1%BB%8Bnh%20h%C6%B0%E1%BB%9Bng%20tin%20h%E1%BB%8Dc%20%E1%BB%A9ng%20d%E1%BB%A5ng-mau.pdf) | **Lớp 11 – THƯD (`ICT`)** | `112.20 MB` | 155 trang | **`Trang PDF = Trang sách + 2`** |
| 4 | [`SGK Tin học 12 - Định hướng Khoa học máy tính-mau.pdf`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/SGK%20Tin%20h%E1%BB%8Dc%2012%20-%20%C4%90%E1%BB%8Bnh%20h%C6%B0%E1%BB%9Bng%20Khoa%20h%E1%BB%8Dc%20m%C3%A1y%20t%C3%ADnh-mau.pdf) | **Lớp 12 – KHMT (`CS`)** | `127.75 MB` | 168 trang | **`Trang PDF = Trang sách + 2`** |
| 5 | [`SGK Tin học 12 - Định hướng Tin học ứng dụng-mau.pdf`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/SGK%20Tin%20h%E1%BB%8Dc%2012%20-%20%C4%90%E1%BB%8Bnh%20h%C6%B0%E1%BB%9Bng%20Tin%20h%E1%BB%8Dc%20%E1%BB%A9ng%20d%E1%BB%A5ng-mau.pdf) | **Lớp 12 – THƯD (`ICT`)** | `115.85 MB` | 160 trang | **`Trang PDF = Trang sách + 2`** |
| 6 | [`trich_xuat_trang_sgk.ps1`](file:///c:/Users/PC/khbd_docx/03_sach_giao_khoa/trich_xuat_trang_sgk.ps1) | Script PowerShell trích xuất trang PDF thành ảnh `.png` để Agent đọc trực tiếp và cắt (crop) hình minh họa SGK. | `2.3 KB` | — | — |

---

## 2. Hướng dẫn BẮT BUỘC cho Agent khi đọc trang sách và Cắt (Crop) Hình ảnh SGK nhúng vào KHBD

> [!IMPORTANT]
> **Quy tắc từ Nhận xét `#18 & #22` của Thầy Thống Nhất:**
> Khi trong mục `b) Nội dung` của bất kỳ Hoạt động / Nhiệm vụ nào có nhắc tới *Hình X.Y* hoặc *Bảng X.Y* trong SGK, Agent **BẮT BUỘC phải trích xuất trang SGK đó ra ảnh `.png`, cắt (crop) đúng khung Hình/Bảng bằng Python (`PIL.Image`) và chèn (embed) trực tiếp vào file `.docx`**. Tuyệt đối không chỉ viết chữ *"Quan sát Hình 12.3 trong SGK"* mà không có hình!

### Bước 1: Tra cứu số trang sách của bài học trong Mục lục tại Mục 3 bên dưới
* Với **SGK Tin học 10**: `Trang PDF = Trang sách + 1` (Ví dụ: Bài 12 ở trang sách `63–68` $\rightarrow$ `Trang PDF = 64–69`).
* Với **SGK Tin học 11 và 12**: `Trang PDF = Trang sách + 2` (Ví dụ: Bài 9 Tin học 12 ở trang sách `52–56` $\rightarrow$ `Trang PDF = 54–58`).
* **Lưu ý khi trích dẫn trong văn bản Word:** Luôn ghi theo **`Trang sách` in trên góc trang SGK** (ví dụ: *SGK Tin học 10, trang 63*), không ghi số trang PDF.

### Bước 2: Chạy script `trich_xuat_trang_sgk.ps1` (hoặc `pdftoppm`) xuất các trang của bài học ra ảnh `.png` tại thư mục `scratch`
```powershell
powershell -ExecutionPolicy Bypass -File "c:\Users\PC\khbd_docx\03_sach_giao_khoa\trich_xuat_trang_sgk.ps1" `
    -PdfPath "c:\Users\PC\khbd_docx\03_sach_giao_khoa\Tin-học-10-KNTT.pdf" `
    -StartPage 64 -EndPage 69 `
    -OutputDir "C:\Users\PC\.gemini\antigravity\brain\<conversation-id>\scratch\sgk_pages"
```

### Bước 3: Dùng `read_file` đọc ảnh trang sách và dùng `PIL.Image` cắt các Hình/Bảng cần nhúng vào `.docx`
* Lưu các hình đã cắt vào thư mục `_scripts_build/cropped_figures/` tương ứng của gói bài dạy để nhúng vào mục `b) Nội dung` thông qua hàm `doc.fig(filename, caption)` của `docx_engine_5512.py`.

---

## 3. Bảng Tra Cứu Mục Lục Chi Tiết 5 Cuốn Sách Giáo Khoa

### 3.1. Mục lục SGK Tin học 10 (`Tin-học-10-KNTT.pdf`)
*(Quy đổi: `Trang PDF = Trang sách + 1`)*

* **CHỦ ĐỀ 1. MÁY TÍNH VÀ XÃ HỘI TRI THỨC**
  * **Bài 1 (`§1.1`).** Thông tin và xử lí thông tin *(Trang sách: 6–10 | Trang PDF: 7–11)* — **Có dạy tại THPT Thuận Hoá (Tiết PPCT 01, 02)**
  * **Bài 2 (`§1.2`).** Vai trò của thiết bị thông minh và tin học đối với xã hội *(Trang sách: 11–15 | Trang PDF: 12–16)* — **Có dạy tại THPT Thuận Hoá (Tiết PPCT 03, 04)**
  * **Bài 3.** Một số kiểu dữ liệu và dữ liệu văn bản *(Trang sách: 16 | Trang PDF: 17)* — *(THPT Thuận Hoá KHÔNG phân phối)*
  * **Bài 4.** Hệ nhị phân và dữ liệu số nguyên *(Trang sách: 20 | Trang PDF: 21)* — *(THPT Thuận Hoá KHÔNG phân phối)*
  * **Bài 5.** Dữ liệu lôgic *(Trang sách: 24 | Trang PDF: 25)* — *(THPT Thuận Hoá KHÔNG phân phối)*
  * **Bài 6.** Dữ liệu âm thanh và hình ảnh *(Trang sách: 28 | Trang PDF: 29)* — *(THPT Thuận Hoá KHÔNG phân phối)*
  * **Bài 7 (`§1.7`).** Thực hành sử dụng thiết bị số thông dụng *(Trang sách: 33–37 | Trang PDF: 34–38)* — **Có dạy tại THPT Thuận Hoá (Tiết PPCT 05, 06)**
* **CHỦ ĐỀ 2. MẠNG MÁY TÍNH VÀ INTERNET**
  * **Bài 8 (`§2.8`).** Mạng máy tính trong cuộc sống hiện đại *(Trang sách: 38–43 | Trang PDF: 39–44)* — **Tiết PPCT 07, 08**
  * **Bài 9 (`§2.9`).** An toàn trên không gian mạng *(Trang sách: 44–49 | Trang PDF: 45–50)* — **Tiết PPCT 09, 10**
  * **Bài 10 (`§2.10`).** Thực hành khai thác tài nguyên trên Internet *(Trang sách: 50–54 | Trang PDF: 51–55)* — **Tiết PPCT 11, 12**
* **CHỦ ĐỀ 3. ĐẠO ĐỨC, PHÁP LUẬT VÀ VĂN HOÁ TRONG MÔI TRƯỜNG SỐ**
  * **Bài 11 (`§3.11`).** Ứng xử trên môi trường số. Nghĩa vụ tôn trọng bản quyền *(Trang sách: 55–60 | Trang PDF: 56–61)* — **Tiết PPCT 13, 14**
  * *(Ôn tập GHKI: Tiết PPCT 15 | Kiểm tra GHKI: Tiết PPCT 16)*
* **CHỦ ĐỀ 4. ỨNG DỤNG TIN HỌC** *(Định hướng ICT – Thiết kế đồ họa Inkscape)*
  * **Bài 12 (`§4.12`).** Phần mềm thiết kế đồ hoạ *(Trang sách: 63–68 | Trang PDF: 64–69)* — **Tiết PPCT 17 (LT), 18 (TH)**
  * **Bài 13 (`§4.13`).** Bổ sung các đối tượng đồ hoạ *(Trang sách: 69–74 | Trang PDF: 70–75)* — **Tiết PPCT 19 (LT), 20 (TH)**
  * **Bài 14 (`§4.14`).** Làm việc với đối tượng đường và văn bản *(Trang sách: 75–81 | Trang PDF: 76–82)* — **Tiết PPCT 21 (LT), 22 (TH)**
  * **Bài 15 (`§4.15`).** Hoàn thiện hình ảnh đồ hoạ *(Trang sách: 82–85 | Trang PDF: 83–86)* — **Tiết PPCT 23 (TH), 24 (TH)**
* **CHỦ ĐỀ 5. GIẢI QUYẾT VẤN ĐỀ VỚI SỰ TRỢ GIÚP CỦA MÁY TÍNH** *(Lập trình Python)*
  * **Bài 16 (`§5.16`).** Ngôn ngữ lập trình bậc cao và Python *(Trang sách: 86 | Trang PDF: 87)* — **Tiết PPCT 25, 26**
  * **Bài 17 (`§5.17`).** Biến và lệnh gán *(Trang sách: 91 | Trang PDF: 92)* — **Tiết PPCT 27, 28**
  * **Bài 18 (`§5.18`).** Các lệnh vào ra đơn giản *(Trang sách: 97 | Trang PDF: 98)* — **Tiết PPCT 29, 30**
  * **Bài 19 (`§5.19`).** Câu lệnh rẽ nhánh `if` *(Trang sách: 101 | Trang PDF: 102)* — **Tiết PPCT 31, 32**
  * *(Ôn tập CHK1: Tiết PPCT 33 | Kiểm tra CHK1: Tiết PPCT 34)*
  * **Bài 20 (`§5.20`).** Câu lệnh lặp `for` *(Trang sách: 105 | Trang PDF: 106)* — **Tiết PPCT 35, 36**
  * **Bài 21 (`§5.21`).** Câu lệnh lặp `while` *(Trang sách: 108 | Trang PDF: 109)* — **Tiết PPCT 37, 38**
  * **Bài 22 (`§5.22`).** Kiểu dữ liệu danh sách *(Trang sách: 111 | Trang PDF: 112)* — **Tiết PPCT 39, 40**
  * **Bài 23 (`§5.23`).** Một số lệnh làm việc với dữ liệu danh sách *(Trang sách: 115 | Trang PDF: 116)* — **Tiết PPCT 41, 42, 43**
  * **Bài 24 (`§5.24`).** Xâu kí tự *(Trang sách: 119 | Trang PDF: 120)* — **Tiết PPCT 44, 45**
  * **Bài 25 (`§5.25`).** Một số lệnh làm việc với xâu kí tự *(Trang sách: 123 | Trang PDF: 124)* — **Tiết PPCT 46, 47, 48**
  * **Bài 26 (`§5.26`).** Hàm trong Python *(Trang sách: 127 | Trang PDF: 128)* — **Tiết PPCT 49, 50**
  * *(Ôn tập GHK2: Tiết PPCT 51 | Kiểm tra GHK2: Tiết PPCT 52)*
  * **Bài 27 (`§5.27`).** Tham số của hàm *(Trang sách: 131 | Trang PDF: 132)* — **Tiết PPCT 53, 54, 55**
  * **Bài 28 (`§5.28`).** Phạm vi của biến *(Trang sách: 136 | Trang PDF: 137)* — **Tiết PPCT 56, 57, 58**
  * **Bài 29 (`§5.29`).** Nhận biết lỗi chương trình *(Trang sách: 141 | Trang PDF: 142)* — **Tiết PPCT 59, 60**
  * **Bài 30 (`§5.30`).** Kiểm thử và gỡ lỗi chương trình *(Trang sách: 145 | Trang PDF: 146)* — **Tiết PPCT 61, 62**
  * **Bài 31 (`§5.31`).** Thực hành viết chương trình đơn giản *(Trang sách: 149 | Trang PDF: 150)* — **Tiết PPCT 63, 64**
  * **Bài 32 (`§5.32`).** Ôn tập lập trình Python *(Trang sách: 153 | Trang PDF: 154)* — **Tiết PPCT 65, 66**
  * *(Ôn tập CHK2: Tiết PPCT 67 | Kiểm tra CHK2: Tiết PPCT 68)*
* **CHỦ ĐỀ 6. HƯỚNG NGHIỆP VỚI TIN HỌC**
  * **Bài 33 (`§6.33`).** Nghề thiết kế đồ hoạ máy tính *(Trang sách: 156 | Trang PDF: 157)* — **Tiết PPCT 69**
  * **Bài 34 (`§6.34`).** Nghề phát triển phần mềm *(Trang sách: 160 | Trang PDF: 161)* — **Tiết PPCT 70**

---

### 3.2. Mục lục SGK Tin học 11 – Định hướng Khoa học máy tính & Tin học ứng dụng
*(Quy đổi cho cả 2 cuốn Lớp 11: `Trang PDF = Trang sách + 2`)*

* **Phần chung cho cả 2 cuốn Lớp 11 (Từ Bài 1 đến Bài 16):**
  * **Chủ đề 1:** Bài 1 (tr. 5 / PDF 7), Bài 2 (tr. 10 / PDF 12), Bài 3 (tr. 15 / PDF 17), Bài 4 (tr. 21 / PDF 23), Bài 5 (tr. 27 / PDF 29).
  * **Chủ đề 2:** Bài 6 (tr. 32 / PDF 34), Bài 7 (tr. 36 / PDF 38), Bài 8 (tr. 39 / PDF 41).
  * **Chủ đề 3:** Bài 9 (tr. 43 / PDF 45).
  * **Chủ đề 4:** Bài 10 (tr. 49 / PDF 51), Bài 11 (tr. 53 / PDF 55), Bài 12 (tr. 58 / PDF 60), Bài 13 (tr. 64 / PDF 66), Bài 14 (tr. 69 / PDF 71), Bài 15 (tr. 73 / PDF 75).
  * **Chủ đề 5:** Bài 16 (tr. 77 / PDF 79).
* **Phần riêng Lớp 11 – KHMT (`CS`):** Bài 17 → Bài 31 (Trang sách 81 → 145 | Trang PDF 83 → 147).
* **Phần riêng Lớp 11 – THƯD (`ICT`):** Bài 17 → Bài 31 (Trang sách 81 → 150 | Trang PDF 83 → 152).

---

### 3.3. Mục lục SGK Tin học 12 – Định hướng Khoa học máy tính & Tin học ứng dụng
*(Quy đổi cho cả 2 cuốn Lớp 12: `Trang PDF = Trang sách + 2`)*

* **Phần chung cho cả 2 cuốn Lớp 12 (Từ Bài 1 đến Bài 21):**
  * **Chủ đề 1:** Bài 1. Làm quen với Trí tuệ nhân tạo *(tr. 5 | PDF 7 – Tiết PPCT 01, 02)*; Bài 2. Trí tuệ nhân tạo trong khoa học và đời sống *(tr. 9 | PDF 11 – Tiết PPCT 03, 04)*.
  * **Chủ đề 2:** Bài 3. Một số thiết bị mạng thông dụng *(tr. 14 | PDF 16 – Tiết PPCT 05, 06)*; Bài 4. Giao thức mạng *(tr. 21 | PDF 23 – Tiết PPCT 07, 08)*; Bài 5. Thực hành chia sẻ tài nguyên trên mạng *(tr. 26 | PDF 28 – Tiết PPCT 09–12)*.
  * **Chủ đề 3:** Bài 6. Giao tiếp và ứng xử trong không gian mạng *(tr. 34 | PDF 36 – Tiết PPCT 13, 14)*.
  * *(Kiểm tra Giữa học kì I: Tiết PPCT 15)*.
  * **Chủ đề 4 (Thiết kế trang web HTML & CSS):**
    * **Bài 7.** HTML và cấu trúc trang web *(Trang sách: 39–45 | Trang PDF: 41–47 | Tiết PPCT: 16, 17)*
    * **Bài 8.** Định dạng văn bản *(Trang sách: 46–51 | Trang PDF: 48–53 | Tiết PPCT: 18, 19)*
    * **Bài 9.** Tạo danh sách, bảng *(Trang sách: 52–56 | Trang PDF: 54–58 | **Tiết PPCT: 20, 21**)*
    * **Bài 10.** Tạo liên kết *(Trang sách: 57–61 | Trang PDF: 59–63 | **Tiết PPCT: 22, 23**)*
    * **Bài 11.** Chèn tệp tin đa phương tiện và khung nội tuyến vào trang web *(Trang sách: 62–66 | Trang PDF: 64–68 | Tiết PPCT: 24, 25)*
    * **Bài 12.** Tạo biểu mẫu *(Trang sách: 67–70 | Trang PDF: 69–72 | Tiết PPCT: 26, 27)*
    * **Bài 13.** Khái niệm, vai trò của CSS *(Trang sách: 71–75 | Trang PDF: 73–77 | Tiết PPCT: 28, 29)*
    * **Bài 14.** Định dạng văn bản bằng CSS *(Trang sách: 76–82 | Trang PDF: 78–84 | Tiết PPCT: 30, 31)*
    * **Bài 15.** Tạo màu cho chữ và nền *(Trang sách: 83–88 | Trang PDF: 85–90 | Tiết PPCT: 32, 33)*
    * **Bài 16.** Định dạng khung *(Trang sách: 89–95 | Trang PDF: 91–97 | Tiết PPCT: 37, 38)*
    * **Bài 17.** Các mức ưu tiên của bộ chọn *(Trang sách: 96–101 | Trang PDF: 98–103 | Tiết PPCT: 39, 40)*
    * **Bài 18.** Thực hành tổng hợp thiết kế trang web *(Trang sách: 102–105 | Trang PDF: 104–107 | Tiết PPCT: 41, 42, 43)*
  * **Chủ đề 5:** Bài 19 (tr. 106 / PDF 108), Bài 20 (tr. 110 / PDF 112), Bài 21 (tr. 113 / PDF 115).
* **Phần riêng Lớp 12 – KHMT (`CS`):** Bài 22 → Bài 30 (Trang sách 118 → 161 | Trang PDF 120 → 163).
* **Phần riêng Lớp 12 – THƯD (`ICT`):** Bài 22 → Bài 28 (Trang sách 118 → 153 | Trang PDF 120 → 155).

---

## 4. Quy tắc Ánh xạ Cấu trúc Bài học SGK sang 04 Hoạt động Chuẩn Công văn 5512

| Thành phần trong SGK KNTT | Cách chuyển hóa vào KHBD chuẩn Công văn 5512 (Theo 42 nhận xét của Giảng viên) |
| :--- | :--- |
| **Hộp "SAU BÀI NÀY EM SẼ:"** | Kết hợp với YCCĐ trong Ngăn 01 để viết mục `I. MỤC TIÊU -> 1. Về kiến thức` (gắn mã `[KT1]`, `[KT2]`...). **Không tạo mục `V. KẾT QUẢ DỰ KIẾN` ở cuối bài.** |
| **Mục Khởi động (Biểu tượng màn hình xanh)** | Chuyển thành **`1. Hoạt động 1: Mở đầu (Xác định vấn đề / nhiệm vụ học tập)`**. |
| **Các mục nội dung 1, 2... + Hộp "Hoạt động 1, 2..." + Hộp Kiến thức + Câu hỏi củng cố + Hình/Bảng minh họa** | Chuyển toàn bộ vào duy nhất mục **`2. Hoạt động 2: Hình thành kiến thức mới`**, chia thành các tiểu mục **`2.1. Nhiệm vụ 1`**, **`2.2. Nhiệm vụ 2`**, **`2.3. Nhiệm vụ 3`**... (bắt buộc cắt và nhúng trực tiếp các Hình/Bảng SGK vào mục `b) Nội dung`). |
| **Mục Thực hành + Luyện tập** | Chuyển thành **`3. Hoạt động 3: Luyện tập`** (chỉ thực hiện sau khi đã hoàn thành xong toàn bộ các Nhiệm vụ của Hoạt động 2; đi kèm Phiếu học tập hoặc Rubric 6 cột tại `IV. PHỤ LỤC`). |
| **Mục Vận dụng** | Chuyển thành **`4. Hoạt động 4: Vận dụng`** (giao nhiệm vụ vận dụng thực tiễn hoặc kiểm chứng/sửa lỗi mã nguồn kèm hướng dẫn nộp bài). |
