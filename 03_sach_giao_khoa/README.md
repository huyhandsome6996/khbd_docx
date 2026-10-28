# NGĂN 03: SÁCH GIÁO KHOA TIN HỌC 10, 11, 12 (BỘ KẾT NỐI TRI THỨC VỚI CUỘC SỐNG)

> **Vai trò của ngăn này đối với AI Agent:** Cung cấp toàn bộ **ngữ liệu, bảng biểu, hình ảnh minh họa, hoạt động khám phá, hộp kiến thức, bài tập thực hành, luyện tập và vận dụng** của 5 cuốn Sách giáo khoa Tin học cấp THPT (Lớp 10, Lớp 11 KHMT/THƯD, Lớp 12 KHMT/THƯD) thuộc bộ sách *Kết nối tri thức với cuộc sống* (NXB Giáo dục Việt Nam).

---

## 1. Danh mục 5 cuốn Sách giáo khoa & Công cụ trong ngăn

| STT | Tên file | Khối lớp & Định hướng | Dung lượng | Tổng số trang PDF | Công thức quy đổi số trang |
| :-: | :--- | :--- | :---: | :---: | :--- |
| 1 | `Tin-học-10-KNTT.pdf` | **Lớp 10** (Chung cả 2 định hướng ICT & CS) | `27.37 MB` | 170 trang | `Trang PDF = Trang sách + 1` |
| 2 | `SGK Tin học 11 - Định hướng khoa học máy tính-mau.pdf` | **Lớp 11 – KHMT (`CS`)** | `105.50 MB` *(Git LFS)* | 151 trang | `Trang PDF = Trang sách + 2` |
| 3 | `SGK Tin học 11 - Định hướng tin học ứng dụng-mau.pdf` | **Lớp 11 – THƯD (`ICT`)** | `112.20 MB` *(Git LFS)* | 155 trang | `Trang PDF = Trang sách + 2` |
| 4 | `SGK Tin học 12 - Định hướng Khoa học máy tính-mau.pdf` | **Lớp 12 – KHMT (`CS`)** | `127.75 MB` *(Git LFS)* | 168 trang | `Trang PDF = Trang sách + 2` |
| 5 | `SGK Tin học 12 - Định hướng Tin học ứng dụng-mau.pdf` | **Lớp 12 – THƯD (`ICT`)** | `115.85 MB` *(Git LFS)* | 160 trang | `Trang PDF = Trang sách + 2` |
| 6 | `trich_xuat_trang_sgk.ps1` | Script PowerShell trích xuất trang PDF thành ảnh `.png` để Agent đọc trực tiếp các file > 100 MB. | `2.3 KB` | — | — |

---

## 2. Hướng dẫn BẮT BUỘC cho Agent khi đọc trang sách trong ngăn này

> [!IMPORTANT]
> **Lưu ý kỹ thuật cực kỳ quan trọng cho AI Agent:**
> Công cụ `read_file` mặc định có giới hạn kích thước tệp tối đa là **100 MB**, và khi gọi `read_file` trực tiếp lên file PDF có thể tải toàn bộ hàng trăm trang gây tràn bộ nhớ ngữ cảnh (Context Window).
> Vì vậy, để đọc chính xác nội dung của bất kỳ bài học nào trong 5 cuốn SGK (đặc biệt là 4 cuốn SGK lớp 11 và 12 nặng trên 100 MB), Agent **BẮT BUỘC** làm theo 3 bước sau:

### Bước 1: Tra cứu số trang sách của bài học trong Mục lục tại المục 3 bên dưới
* Với **SGK Tin học 10**: `Trang PDF = Trang sách + 1` (Ví dụ: Bài 20 ở trang sách `105` $\rightarrow$ `Trang PDF = 106`).
* Với **SGK Tin học 11 và 12**: `Trang PDF = Trang sách + 2` (Ví dụ: Bài 8 Tin học 12 ở trang sách `46` đến `51` $\rightarrow$ `Trang PDF = 48` đến `53`).

### Bước 2: Chạy script `trich_xuat_trang_sgk.ps1` để xuất đúng các trang của bài học ra ảnh `.png`
Sử dụng công cụ `run_command` chạy lệnh:
```powershell
powershell -ExecutionPolicy Bypass -File "c:\Users\PC\khbd_docx\03_sach_giao_khoa\trich_xuat_trang_sgk.ps1" `
    -PdfPath "c:\Users\PC\khbd_docx\03_sach_giao_khoa\SGK Tin học 12 - Định hướng Khoa học máy tính-mau.pdf" `
    -StartPage 48 -EndPage 53 `
    -OutputDir "C:\Users\PC\.gemini\antigravity\brain\<conversation-id>\scratch\sgk_pages"
```

### Bước 3: Dùng công cụ `read_file` đọc các file ảnh `page_048.png` ... `page_053.png`
* Sau khi đọc xong nội dung chi tiết từng trang (Bảng biểu, Hình ảnh, Hoạt động 1, Hoạt động 2, Hộp kiến thức, Câu hỏi, Thực hành, Luyện tập, Vận dụng), Agent sử dụng chính xác ngữ liệu đó (ví dụ: *"nghiên cứu Bảng 8.1 SGK"*, *"quan sát Hình 19.2 SGK"*) để đưa vào Kế hoạch bài dạy.
* Tuyệt đối không lưu thư mục ảnh tạm `extracted_pages` vào Git commit.

---

## 3. Bảng Tra Cứu Mục Lục Chi Tiết 5 Cuốn Sách Giáo Khoa

### 3.1. Mục lục SGK Tin học 10 (`Tin-học-10-KNTT.pdf`)
*(Quy đổi: `Trang PDF = Trang sách + 1`)*

* **CHỦ ĐỀ 1. MÁY TÍNH VÀ XÃ HỘI TRI THỨC**
  * **Bài 1.** Thông tin và xử lí thông tin *(Trang sách: 6 | Trang PDF: 7)*
  * **Bài 2.** Vai trò của thiết bị thông minh và tin học đối với xã hội *(Trang sách: 11 | Trang PDF: 12)*
  * **Bài 3.** Một số kiểu dữ liệu và dữ liệu văn bản *(Trang sách: 16 | Trang PDF: 17)*
  * **Bài 4.** Hệ nhị phân và dữ liệu số nguyên *(Trang sách: 20 | Trang PDF: 21)*
  * **Bài 5.** Dữ liệu lôgic *(Trang sách: 24 | Trang PDF: 25)*
  * **Bài 6.** Dữ liệu âm thanh và hình ảnh *(Trang sách: 28 | Trang PDF: 29)*
  * **Bài 7.** Thực hành sử dụng thiết bị số thông dụng *(Trang sách: 33 | Trang PDF: 34)*
* **CHỦ ĐỀ 2. MẠNG MÁY TÍNH VÀ INTERNET**
  * **Bài 8.** Mạng máy tính trong cuộc sống hiện đại *(Trang sách: 38 | Trang PDF: 39)*
  * **Bài 9.** An toàn trên không gian mạng *(Trang sách: 44 | Trang PDF: 45)*
  * **Bài 10.** Thực hành khai thác tài nguyên trên Internet *(Trang sách: 50 | Trang PDF: 51)*
* **CHỦ ĐỀ 3. ĐẠO ĐỨC, PHÁP LUẬT VÀ VĂN HOÁ TRONG MÔI TRƯỜNG SỐ**
  * **Bài 11.** Ứng xử trên môi trường số. Nghĩa vụ tôn trọng bản quyền *(Trang sách: 55 | Trang PDF: 56)*
* **CHỦ ĐỀ 4. ỨNG DỤNG TIN HỌC** *(Định hướng ICT – Thiết kế đồ họa Inkscape)*
  * **Bài 12.** Phần mềm thiết kế đồ hoạ *(Trang sách: 63 | Trang PDF: 64)*
  * **Bài 13.** Bổ sung các đối tượng đồ hoạ *(Trang sách: 69 | Trang PDF: 70)*
  * **Bài 14.** Làm việc với đối tượng đường và văn bản *(Trang sách: 75 | Trang PDF: 76)*
  * **Bài 15.** Hoàn thiện hình ảnh đồ hoạ *(Trang sách: 82 | Trang PDF: 83)*
* **CHỦ ĐỀ 5. GIẢI QUYẾT VẤN ĐỀ VỚI SỰ TRỢ GIÚP CỦA MÁY TÍNH** *(Lập trình Python)*
  * **Bài 16.** Ngôn ngữ lập trình bậc cao và Python *(Trang sách: 86 | Trang PDF: 87)*
  * **Bài 17.** Biến và lệnh gán *(Trang sách: 91 | Trang PDF: 92)*
  * **Bài 18.** Các lệnh vào ra đơn giản *(Trang sách: 97 | Trang PDF: 98)*
  * **Bài 19.** Câu lệnh rẽ nhánh `if` *(Trang sách: 101 | Trang PDF: 102)*
  * **Bài 20.** Câu lệnh lặp `for` *(Trang sách: 105 | Trang PDF: 106)*
  * **Bài 21.** Câu lệnh lặp `while` *(Trang sách: 108 | Trang PDF: 109)*
  * **Bài 22.** Kiểu dữ liệu danh sách *(Trang sách: 111 | Trang PDF: 112)*
  * **Bài 23.** Một số lệnh làm việc với dữ liệu danh sách *(Trang sách: 115 | Trang PDF: 116)*
  * **Bài 24.** Xâu kí tự *(Trang sách: 119 | Trang PDF: 120)*
  * **Bài 25.** Một số lệnh làm việc với xâu kí tự *(Trang sách: 123 | Trang PDF: 124)*
  * **Bài 26.** Hàm trong Python *(Trang sách: 127 | Trang PDF: 128)*
  * **Bài 27.** Tham số của hàm *(Trang sách: 131 | Trang PDF: 132)*
  * **Bài 28.** Phạm vi của biến *(Trang sách: 136 | Trang PDF: 137)*
  * **Bài 29.** Nhận biết lỗi chương trình *(Trang sách: 141 | Trang PDF: 142)*
  * **Bài 30.** Kiểm thử và gỡ lỗi chương trình *(Trang sách: 145 | Trang PDF: 146)*
  * **Bài 31.** Thực hành viết chương trình đơn giản *(Trang sách: 149 | Trang PDF: 150)*
  * **Bài 32.** Ôn tập lập trình Python *(Trang sách: 153 | Trang PDF: 154)*
* **CHỦ ĐỀ 6. HƯỚNG NGHIỆP VỚI TIN HỌC**
  * **Bài 33.** Nghề thiết kế đồ hoạ máy tính *(Trang sách: 156 | Trang PDF: 157)*
  * **Bài 34.** Nghề phát triển phần mềm *(Trang sách: 160 | Trang PDF: 161)*

---

### 3.2. Mục lục SGK Tin học 11 – Định hướng Khoa học máy tính & Tin học ứng dụng
*(Quy đổi cho cả 2 cuốn Lớp 11: `Trang PDF = Trang sách + 2`)*

#### A. Phần chung cho cả 2 cuốn Lớp 11 (Từ Bài 1 đến Bài 16)
* **CHỦ ĐỀ 1. MÁY TÍNH VÀ XÃ HỘI TRI THỨC**
  * **Bài 1.** Hệ điều hành *(Trang sách: 5 | Trang PDF: 7)*
  * **Bài 2.** Thực hành sử dụng hệ điều hành *(Trang sách: 10 | Trang PDF: 12)*
  * **Bài 3.** Phần mềm nguồn mở và phần mềm chạy trên Internet *(Trang sách: 15 | Trang PDF: 17)*
  * **Bài 4.** Bên trong máy tính *(Trang sách: 21 | Trang PDF: 23)*
  * **Bài 5.** Kết nối máy tính với các thiết bị số *(Trang sách: 27 | Trang PDF: 29)*
* **CHỦ ĐỀ 2. TỔ CHỨC LƯU TRỮ, TÌM KIẾM VÀ TRAO ĐỔI THÔNG TIN**
  * **Bài 6.** Lưu trữ và chia sẻ tệp tin trên Internet *(Trang sách: 32 | Trang PDF: 34)*
  * **Bài 7.** Thực hành tìm kiếm thông tin trên Internet *(Trang sách: 36 | Trang PDF: 38)*
  * **Bài 8.** Thực hành nâng cao sử dụng thư điện tử và mạng xã hội *(Trang sách: 39 | Trang PDF: 41)*
* **CHỦ ĐỀ 3. ĐẠO ĐỨC, PHÁP LUẬT VÀ VĂN HOÁ TRONG MÔI TRƯỜNG SỐ**
  * **Bài 9.** Giao tiếp an toàn trên Internet *(Trang sách: 43 | Trang PDF: 45)*
* **CHỦ ĐỀ 4. GIỚI THIỆU CÁC HỆ CƠ SỞ DỮ LIỆU**
  * **Bài 10.** Lưu trữ dữ liệu và khai thác thông tin phục vụ quản lí *(Trang sách: 49 | Trang PDF: 51)*
  * **Bài 11.** Cơ sở dữ liệu *(Trang sách: 53 | Trang PDF: 55)*
  * **Bài 12.** Hệ quản trị cơ sở dữ liệu và hệ cơ sở dữ liệu *(Trang sách: 58 | Trang PDF: 60)*
  * **Bài 13.** Cơ sở dữ liệu quan hệ *(Trang sách: 64 | Trang PDF: 66)*
  * **Bài 14.** SQL – Ngôn ngữ truy vấn có cấu trúc *(Trang sách: 69 | Trang PDF: 71)*
  * **Bài 15.** Bảo mật và an toàn hệ cơ sở dữ liệu *(Trang sách: 73 | Trang PDF: 75)*
* **CHỦ ĐỀ 5. HƯỚNG NGHIỆP VỚI TIN HỌC**
  * **Bài 16.** Công việc quản trị cơ sở dữ liệu *(Trang sách: 77 | Trang PDF: 79)*

#### B. Phần riêng Lớp 11 – Định hướng Khoa học máy tính (`SGK Tin học 11 - Định hướng khoa học máy tính-mau.pdf`)
* **CHỦ ĐỀ 6. KĨ THUẬT LẬP TRÌNH**
  * **Bài 17.** Dữ liệu mảng một chiều và hai chiều *(Trang sách: 81 | Trang PDF: 83)*
  * **Bài 18.** Thực hành dữ liệu mảng một chiều và hai chiều *(Trang sách: 86 | Trang PDF: 88)*
  * **Bài 19.** Bài toán tìm kiếm *(Trang sách: 89 | Trang PDF: 91)*
  * **Bài 20.** Thực hành bài toán tìm kiếm *(Trang sách: 94 | Trang PDF: 96)*
  * **Bài 21.** Các thuật toán sắp xếp đơn giản *(Trang sách: 99 | Trang PDF: 101)*
  * **Bài 22.** Thực hành bài toán sắp xếp *(Trang sách: 104 | Trang PDF: 106)*
  * **Bài 23.** Kiểm thử và đánh giá chương trình *(Trang sách: 106 | Trang PDF: 108)*
  * **Bài 24.** Đánh giá độ phức tạp thời gian thuật toán *(Trang sách: 111 | Trang PDF: 113)*
  * **Bài 25.** Thực hành xác định độ phức tạp thời gian thuật toán *(Trang sách: 115 | Trang PDF: 117)*
  * **Bài 26.** Phương pháp làm mịn dần trong thiết kế chương trình *(Trang sách: 118 | Trang PDF: 120)*
  * **Bài 27.** Thực hành thiết kế chương trình theo phương pháp làm mịn dần *(Trang sách: 123 | Trang PDF: 125)*
  * **Bài 28.** Thiết kế chương trình theo mô đun *(Trang sách: 127 | Trang PDF: 129)*
  * **Bài 29.** Thực hành thiết kế chương trình theo mô đun *(Trang sách: 132 | Trang PDF: 134)*
  * **Bài 30.** Thiết lập thư viện cho chương trình *(Trang sách: 137 | Trang PDF: 139)*
  * **Bài 31.** Thực hành thiết lập thư viện chương trình *(Trang sách: 143 | Trang PDF: 145)*

#### C. Phần riêng Lớp 11 – Định hướng Tin học ứng dụng (`SGK Tin học 11 - Định hướng tin học ứng dụng-mau.pdf`)
* **CHỦ ĐỀ 6. THỰC HÀNH TẠO VÀ KHAI THÁC CƠ SỞ DỮ LIỆU**
  * **Bài 17.** Quản trị cơ sở dữ liệu trên máy tính *(Trang sách: 81 | Trang PDF: 83)*
  * **Bài 18.** Thực hành xác định cấu trúc bảng và các trường khoá *(Trang sách: 86 | Trang PDF: 88)*
  * **Bài 19.** Thực hành tạo lập cơ sở dữ liệu và các bảng *(Trang sách: 91 | Trang PDF: 93)*
  * **Bài 20.** Thực hành tạo lập các bảng có khoá ngoài *(Trang sách: 95 | Trang PDF: 97)*
  * **Bài 21.** Thực hành cập nhật và truy xuất dữ liệu các bảng *(Trang sách: 100 | Trang PDF: 102)*
  * **Bài 22.** Thực hành cập nhật bảng dữ liệu có tham chiếu *(Trang sách: 105 | Trang PDF: 107)*
  * **Bài 23.** Thực hành truy xuất dữ liệu qua liên kết các bảng *(Trang sách: 109 | Trang PDF: 111)*
  * **Bài 24.** Thực hành sao lưu dữ liệu *(Trang sách: 113 | Trang PDF: 115)*
* **CHỦ ĐỀ 7. PHẦN MỀM CHỈNH SỬA ẢNH VÀ LÀM VIDEO**
  * **Bài 25.** Phần mềm chỉnh sửa ảnh *(Trang sách: 116 | Trang PDF: 118)*
  * **Bài 26.** Công cụ tinh chỉnh màu sắc và công cụ chọn *(Trang sách: 122 | Trang PDF: 124)*
  * **Bài 27.** Công cụ vẽ và một số ứng dụng *(Trang sách: 128 | Trang PDF: 130)*
  * **Bài 28.** Tạo ảnh động *(Trang sách: 133 | Trang PDF: 135)*
  * **Bài 29.** Khám phá phần mềm làm phim *(Trang sách: 137 | Trang PDF: 139)*
  * **Bài 30.** Biên tập phim *(Trang sách: 143 | Trang PDF: 145)*
  * **Bài 31.** Thực hành tạo phim hoạt hình *(Trang sách: 148 | Trang PDF: 150)*

---

### 3.3. Mục lục SGK Tin học 12 – Định hướng Khoa học máy tính & Tin học ứng dụng
*(Quy đổi cho cả 2 cuốn Lớp 12: `Trang PDF = Trang sách + 2`)*

#### A. Phần chung cho cả 2 cuốn Lớp 12 (Từ Bài 1 đến Bài 21)
* **CHỦ ĐỀ 1. MÁY TÍNH VÀ XÃ HỘI TRI THỨC**
  * **Bài 1.** Làm quen với Trí tuệ nhân tạo *(Trang sách: 5 | Trang PDF: 7)*
  * **Bài 2.** Trí tuệ nhân tạo trong khoa học và đời sống *(Trang sách: 9 | Trang PDF: 11)*
* **CHỦ ĐỀ 2. MẠNG MÁY TÍNH VÀ INTERNET**
  * **Bài 3.** Một số thiết bị mạng thông dụng *(Trang sách: 14 | Trang PDF: 16)*
  * **Bài 4.** Giao thức mạng *(Trang sách: 21 | Trang PDF: 23)*
  * **Bài 5.** Thực hành chia sẻ tài nguyên trên mạng *(Trang sách: 26 | Trang PDF: 28)*
* **CHỦ ĐỀ 3. ĐẠO ĐỨC, PHÁP LUẬT VÀ VĂN HOÁ TRONG MÔI TRƯỜNG SỐ**
  * **Bài 6.** Giao tiếp và ứng xử trong không gian mạng *(Trang sách: 34 | Trang PDF: 36)*
* **CHỦ ĐỀ 4. GIẢI QUYẾT VẤN ĐỀ VỚI SỰ TRỢ GIÚP CỦA MÁY TÍNH** *(Thiết kế trang web HTML & CSS)*
  * **Bài 7.** HTML và cấu trúc trang web *(Trang sách: 39 | Trang PDF: 41)*
  * **Bài 8.** Định dạng văn bản *(Trang sách: 46 | Trang PDF: 48)*
  * **Bài 9.** Tạo danh sách, bảng *(Trang sách: 52 | Trang PDF: 54)*
  * **Bài 10.** Tạo liên kết *(Trang sách: 57 | Trang PDF: 59)*
  * **Bài 11.** Chèn tệp tin đa phương tiện và khung nội tuyến vào trang web *(Trang sách: 62 | Trang PDF: 64)*
  * **Bài 12.** Tạo biểu mẫu *(Trang sách: 67 | Trang PDF: 69)*
  * **Bài 13.** Khái niệm, vai trò của CSS *(Trang sách: 71 | Trang PDF: 73)*
  * **Bài 14.** Định dạng văn bản bằng CSS *(Trang sách: 76 | Trang PDF: 78)*
  * **Bài 15.** Tạo màu cho chữ và nền *(Trang sách: 83 | Trang PDF: 85)*
  * **Bài 16.** Định dạng khung *(Trang sách: 89 | Trang PDF: 91)*
  * **Bài 17.** Các mức ưu tiên của bộ chọn *(Trang sách: 96 | Trang PDF: 98)*
  * **Bài 18.** Thực hành tổng hợp thiết kế trang web *(Trang sách: 102 | Trang PDF: 104)*
* **CHỦ ĐỀ 5. HƯỚNG NGHIỆP VỚI TIN HỌC**
  * **Bài 19.** Dịch vụ sửa chữa và bảo trì máy tính *(Trang sách: 106 | Trang PDF: 108)*
  * **Bài 20.** Nhóm nghề quản trị trong ngành Công nghệ thông tin *(Trang sách: 110 | Trang PDF: 112)*
  * **Bài 21.** Hội thảo hướng nghiệp *(Trang sách: 113 | Trang PDF: 115)*

#### B. Phần riêng Lớp 12 – Định hướng Khoa học máy tính (`SGK Tin học 12 - Định hướng Khoa học máy tính-mau.pdf`)
* **CHỦ ĐỀ 6. MẠNG MÁY TÍNH VÀ INTERNET**
  * **Bài 22.** Tìm hiểu thiết bị mạng *(Trang sách: 118 | Trang PDF: 120)*
  * **Bài 23.** Đường truyền mạng và ứng dụng *(Trang sách: 123 | Trang PDF: 125)*
  * **Bài 24.** Sơ bộ về thiết kế mạng *(Trang sách: 128 | Trang PDF: 130)*
* **CHỦ ĐỀ 7. GIẢI QUYẾT VẤN ĐỀ VỚI SỰ TRỢ GIÚP CỦA MÁY TÍNH**
  * **Bài 25.** Làm quen với Học máy *(Trang sách: 134 | Trang PDF: 136)*
  * **Bài 26.** Làm quen với Khoa học dữ liệu *(Trang sách: 140 | Trang PDF: 142)*
  * **Bài 27.** Máy tính và Khoa học dữ liệu *(Trang sách: 145 | Trang PDF: 147)*
  * **Bài 28.** Thực hành trải nghiệm trích rút thông tin và tri thức *(Trang sách: 149 | Trang PDF: 151)*
  * **Bài 29.** Mô phỏng trong giải quyết vấn đề *(Trang sách: 155 | Trang PDF: 157)*
  * **Bài 30.** Ứng dụng mô phỏng trong giáo dục *(Trang sách: 159 | Trang PDF: 161)*

#### C. Phần riêng Lớp 12 – Định hướng Tin học ứng dụng (`SGK Tin học 12 - Định hướng Tin học ứng dụng-mau.pdf`)
* **CHỦ ĐỀ 6. MÁY TÍNH VÀ XÃ HỘI TRI THỨC**
  * **Bài 22.** Thực hành kết nối các thiết bị số *(Trang sách: 118 | Trang PDF: 120)*
* **CHỦ ĐỀ 7. ỨNG DỤNG TIN HỌC** *(Tạo website bằng Google Sites)*
  * **Bài 23.** Chuẩn bị xây dựng trang web *(Trang sách: 126 | Trang PDF: 128)*
  * **Bài 24.** Xây dựng phần đầu trang web *(Trang sách: 134 | Trang PDF: 136)*
  * **Bài 25.** Xây dựng phần thân và chân trang web *(Trang sách: 139 | Trang PDF: 141)*
  * **Bài 26.** Liên kết và thanh điều hướng *(Trang sách: 144 | Trang PDF: 146)*
  * **Bài 27.** Biểu mẫu trên trang web *(Trang sách: 147 | Trang PDF: 149)*
  * **Bài 28.** Thực hành tổng hợp *(Trang sách: 151 | Trang PDF: 153)*

---

## 4. Quy tắc chuyển hóa Cấu trúc bài học SGK thành Hoạt động trong KHBD

Mỗi bài học trong SGK *Kết nối tri thức với cuộc sống* được thiết kế theo mô hình hoạt động thống nhất. Khi chuyển hóa thành KHBD (thường chia làm `2 tiết/bài` hoặc `1 tiết/bài` tùy phân phối chương trình), Agent cần ánh xạ như sau:

| Thành phần trong SGK KNTT | Cách xử lý và đưa vào Kế hoạch bài dạy (KHBD) |
| :--- | :--- |
| **Hộp "SAU BÀI NÀY EM SẼ:"** | Kết hợp với YCCĐ trong Ngăn 01 để viết mục `I. MỤC TIÊU -> 1. Về kiến thức` và mục `V. KẾT QUẢ DỰ KIẾN`. |
| **Mục Khởi động (Biểu tượng màn hình xanh)** | Chuyển thành **Hoạt động 1 (Mở đầu / Ôn tập và xác định nhiệm vụ)**: Tổ chức trò chơi nhanh, câu hỏi tình huống hoặc quan sát sản phẩm mẫu để dẫn dắt vào Tình huống học tập của tiết dạy. |
| **Các mục nội dung 1, 2... + Hộp "Hoạt động 1, 2..." + Hộp Kiến thức (Khung nét đứt vàng) + Câu hỏi củng cố (Biểu tượng dấu hỏi)** | Chuyển thành **Hoạt động 2, Hoạt động 3 (Tìm hiểu / Hình thành kiến thức mới)**: Yêu cầu học sinh nghiên cứu bảng/hình/đoạn mã trong SGK, trao đổi theo cặp trả lời bộ câu hỏi dẫn dắt `(1), (2), (3)...`, viết thử đoạn mã ngắn, sau đó GV chuẩn hóa theo Hộp kiến thức. |
| **Mục Thực hành + Luyện tập** | Chuyển thành **Hoạt động 4 (Thực hành thiết kế sản phẩm / Viết chương trình)**: Gắn với một **Tình huống học tập có ý nghĩa thực tiễn** (ví dụ: thiết kế trang báo điện tử cá nhân, viết chương trình quản lý quỹ lớp...), có yêu cầu kỹ thuật rõ ràng và đánh giá bằng bảng Rubric 10 điểm. |
| **Mục Vận dụng + Mở rộng AI** | Chuyển thành **Hoạt động 5 (Sử dụng AI kiểm tra mã / Phát hiện lỗi và Giao nhiệm vụ vận dụng sau giờ học)**. |
