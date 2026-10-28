# NGĂN 04: FORM MẪU KẾ HOẠCH BÀI DẠY (LESSON PLAN TEMPLATE)

> **Vai trò của ngăn này đối với AI Agent:** Đây là **biểu mẫu chuẩn mực tuyệt đối (Gold Standard Template)** do người dùng cung cấp. Mọi Kế hoạch bài dạy (Đề cương bài dạy) mà Agent biên soạn về sau **BẮT BUỘC** phải tuân thủ 100% bố cục 5 phần, cấu trúc các bảng biểu, phong cách diễn đạt, cách tích hợp Năng lực số / Năng lực AI và thông số định dạng Word (`.docx`) của file mẫu trong ngăn này.

---

## 1. Tài liệu mẫu trong ngăn

| Tên file | Thông tin bài mẫu | Đặc điểm nổi bật |
| :--- | :--- | :--- |
| `Đề cương bài dạy Tin12_Bai8_Tiet2 (sv Phan Trần Hồng Anh).docx` | **Bài 8. Định dạng văn bản** – *Tiết 2. Định dạng phông chữ và thực hành thiết kế trang báo điện tử* (Tin học 12, 45 phút, Dạy học thực hành tại phòng máy). | Đầy đủ 5 phần lớn (`I` $\rightarrow$ `V`), 4 bảng biểu chuẩn hóa, tích hợp **Năng lực đặc thù (`NLa`–`NLe`)**, **Năng lực chung**, **Năng lực số & Năng lực AI**, **Tình huống học tập xuyên suốt**, **Hoạt động sử dụng AI kiểm tra mã** và **Bảng tiêu chí đánh giá (Rubric 10 điểm)**. |

---

## 2. Thông số Định dạng Kỹ thuật File Word (`.docx`) Chuẩn

Qua phân tích XML gốc (`word/document.xml` và `word/styles.xml`) của file mẫu, khi Agent tạo hoặc xuất file `.docx` tại Ngăn 05, phải thiết lập đúng các thông số sau:

* **Khổ giấy (Page Size):** A4 dọc (`w:w="11906"`, `w:h="16838"` twips $\approx 21.0\text{ cm} \times 29.7\text{ cm}$).
* **Căn lề trang (Page Margins):**
  * Lề trên (`Top`): `2.0 cm` (`1134` twips)
  * Lề dưới (`Bottom`): `2.0 cm` (`1134` twips)
  * Lề trái (`Left`): `3.0 cm` (`1701` twips)
  * Lề phải (`Right`): `2.0 cm` (`1134` twips)
  * Header / Footer: `1.25 cm` (`708` twips)
* **Phông chữ và Cỡ chữ (Typography):**
  * **Phông chữ văn bản mặc định:** `Times New Roman`
  * **Cỡ chữ văn bản chính (Ngoài bảng):** `13 pt` (`w:sz val="26"`)
  * **Cỡ chữ quốc hiệu / tiêu đề lớn (`ĐỀ CƯƠNG BÀI DẠY`, `TÊN BÀI DẠY`):** `13 pt – 15 pt` (`w:sz val="26"` đến `"30"`), **IN HOA, ĐẬM**, căn giữa.
  * **Cỡ chữ nội dung bên trong Bảng biểu:** `12 pt` (`w:sz val="24"`).
  * **Phông chữ và cỡ chữ cho Đoạn mã (HTML / CSS / Python / SQL):** `Courier New`, cỡ `10 pt – 12 pt` (`w:sz val="20"` đến `"24"`).

---

## 3. Cấu trúc Chi tiết 5 Phần của Kế Hoạch Bài Dạy (Chuẩn Form Mẫu)

Dưới đây là cấu trúc khung chi tiết mà Agent phải tuân thủ nghiêm ngặt khi soạn bài:

### PHẦN ĐẦU TRANG (HEADER THÔNG TIN CHUNG)
```text
Trường: [Tên trường, VD: THPT Hai Bà Trưng]          Họ và tên: [Họ và tên giáo viên/sinh viên]
Tổ: Tin

                                  ĐỀ CƯƠNG BÀI DẠY
                        TÊN BÀI DẠY: BÀI [SỐ]. [TÊN BÀI HỌC IN HOA]
                     TIẾT [SỐ]. [TÊN NỘI DUNG TIẾT HỌC IN HOA]
                      Môn học: Tin học; lớp: [10 / 11 / 12]
                      Thời gian thực hiện: [01] tiết (45 phút)
                Hình thức tổ chức: [VD: Dạy học thực hành tại phòng máy]
```

---

### I. MỤC TIÊU

#### 1. Về kiến thức
 Mở đầu bằng cụm: *"Sau tiết học, học sinh:"*
* Gạch đầu dòng 3–4 ý cốt lõi, bắt đầu bằng động từ hành động đo lường được (*Nhận biết được..., Trình bày được..., Thực hiện được..., Viết được...*).
* Nếu tiết học có tích hợp AI, bổ sung 1 ý: *"Biết AI có thể hỗ trợ phát hiện lỗi và gợi ý điều chỉnh mã [HTML/Python/SQL...]; biết cần kiểm tra lại gợi ý trước khi sử dụng."*

#### 2. Về năng lực
Chia rõ 3 mục nhỏ `a)`, `b)`, `c)`:
* **a) Năng lực đặc thù môn Tin học:**
  Liệt kê các năng lực có trong bài kèm mã và tên đầy đủ theo chuẩn CT GDPT 2018:
  * `NLa – Sử dụng và quản lí các phương tiện công nghệ thông tin và truyền thông:` [Mô tả hành vi cụ thể trong bài].
  * `NLc – Giải quyết vấn đề với sự hỗ trợ của công nghệ thông tin và truyền thông:` [Mô tả hành vi cụ thể trong bài].
  * `NLd – Ứng dụng công nghệ thông tin và truyền thông trong học và tự học:` [Mô tả hành vi cụ thể trong bài].
  *(Bổ sung `NLb` hoặc `NLe` nếu nội dung bài học hướng tới ứng xử số hoặc hợp tác nhóm trên môi trường số).*
* **b) Năng lực chung:**
  * `Tự chủ và tự học:` Chủ động nghiên cứu SGK, thực hành viết mã/thao tác và tự điều chỉnh sản phẩm.
  * `Giao tiếp và hợp tác:` Trao đổi theo cặp/nhóm, trình bày và nhận xét kết quả học tập.
  * `Giải quyết vấn đề và sáng tạo:` Lựa chọn nội dung, công cụ, cấu trúc/thuật toán phù hợp để thiết kế sản phẩm.
* **c) Năng lực số và năng lực AI:**
  Ghi mã năng lực số (theo Khung năng lực số) và mã năng lực AI cụ thể gắn với bài học, ví dụ:
  * `3.1.NC1.a:` Phát triển / chỉnh sửa nội dung số (văn bản web, chương trình, cơ sở dữ liệu, đồ họa...).
  * `3.2.NC1.a:` Tinh chỉnh, tích hợp và hoàn thiện nội dung số hiển thị.
  * `[Lớp].C2.2` *(VD: `10.C2.2`, `11.C2.2`, `12.C2.2`)*: Vận dụng trợ lý AI để phát hiện lỗi cú pháp/logic, nhận gợi ý điều chỉnh mã; biết kiểm tra lại kết quả trước khi sử dụng.

#### 3. Về phẩm chất
* `Chăm chỉ:` Tích cực tìm hiểu kiến thức, thực hành và điều chỉnh sản phẩm.
* `Trách nhiệm:` Hoàn thành nhiệm vụ được giao; sử dụng thiết bị và công cụ hỗ trợ đúng mục đích.
* `Trung thực:` Trình bày đúng quá trình thực hiện và giải thích được việc sử dụng gợi ý của AI khi hoàn thiện sản phẩm.

---

### II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU

#### 1. Thiết bị dạy học và điều kiện tổ chức
* Máy tính, máy chiếu.
* Phòng máy có phần mềm phù hợp với bài học *(VD: phần mềm soạn thảo mã/văn bản và trình duyệt web; hoặc IDE Python/Thonny; hoặc hệ QTCSDL MySQL/HeidiSQL; hoặc Inkscape/GIMP)*.
* Kết nối Internet nếu học sinh trực tiếp sử dụng công cụ AI.

#### 2. Học liệu
* **Giáo viên chuẩn bị:**
  * SGK Tin học [Lớp] – Bài [Số]: [Tên bài].
  * Bài trình chiếu, trò chơi ôn tập / khởi động (nêu tên trò chơi cụ thể).
  * Đoạn mã / tệp dữ liệu minh họa các đơn vị kiến thức mới.
  * Sản phẩm mẫu, đoạn mã có lỗi và kết quả phản hồi của AI được chuẩn bị trước (phương án dự phòng khi phòng máy mất mạng).
  * Bảng tiêu chí đánh giá sản phẩm (Rubric).
* **Học sinh chuẩn bị:**
  * SGK, vở ghi và dụng cụ học tập.
  * Tệp sản phẩm đã thực hiện ở tiết trước (nếu là tiết 2 nối tiếp).

---

### III. TIẾN TRÌNH DẠY HỌC

#### 1. Nội dung đã học ở tiết trước (làm cơ sở cho tiết học hiện tại)
*(Nếu soạn Tiết 2 của bài học hoặc bài học kế thừa trực tiếp bài trước, lập Bảng 2 cột tóm tắt nền tảng kiến thức học sinh đã có)*:

| Nội dung | Kiến thức, kĩ năng học sinh đã học |
| :--- | :--- |
| [Đơn vị kiến thức 1] | [Mô tả ngắn gọn những gì HS đã nhận biết và làm được] |
| [Đơn vị kiến thức 2] | [Mô tả ngắn gọn những gì HS đã nhận biết và làm được] |
| Thực hành cơ bản | [Kỹ năng thao tác cơ bản HS đã thực hiện được ở tiết trước] |

* **Sản phẩm dự kiến cuối tiết trước:** Mô tả rõ tệp sản phẩm học sinh đã có để tiếp tục phát triển trong tiết này. *(Nếu là Tiết 1 mở đầu chủ đề mới, có thể tóm tắt kiến thức nền tảng từ chủ đề/lớp trước).*

#### 2. Tình huống học tập
* Thiết kế một bối cảnh thực tiễn, gần gũi với học sinh nhà trường phổ thông (3–5 câu văn), đóng vai trò "sợi chỉ đỏ" xuyên suốt toàn bộ các hoạt động trong tiết học.
* *Ví dụ mẫu:* *"Nhà trường tổ chức hoạt động 'Góc báo điện tử của học sinh'. Mỗi học sinh thiết kế một trang báo điện tử cá nhân giới thiệu một chủ đề gần gũi như môn học yêu thích, một cuốn sách hay, hoạt động của lớp hoặc thông điệp bảo vệ môi trường..."*

#### 3. Phân bố thời gian
Lập **Bảng 2 cột** tổng hợp thời lượng các hoạt động (Tổng cộng = `45 phút`):

| Hoạt động | Thời lượng |
| :--- | :--- |
| Hoạt động 1. [Ôn tập / Khởi động và xác định nhiệm vụ...] | 4 – 5 phút |
| Hoạt động 2. [Tìm hiểu đơn vị kiến thức mới 1...] | 7 – 8 phút |
| Hoạt động 3. [Tìm hiểu đơn vị kiến thức mới 2...] | 7 – 8 phút |
| Hoạt động 4. [Thực hành thiết kế sản phẩm / lập trình...] | 17 – 20 phút |
| Hoạt động 5. [Sử dụng AI kiểm tra mã và giao nhiệm vụ vận dụng] | 6 – 7 phút |
| **Tổng** | **45 phút** |

#### 4. Các hoạt động dạy học (BẢNG CHÍNH 4 CỘT)
Đây là phần trọng tâm nhất của Đề cương bài dạy. Bảng gồm **4 cột bắt buộc** với cấu trúc nội dung từng ô như sau:

| Hoạt động – thời lượng | Nội dung hoạt động và nhiệm vụ học tập | Năng lực, phẩm chất hướng tới | Sản phẩm và đánh giá |
| :--- | :--- | :--- | :--- |
| **HOẠT ĐỘNG [SỐ]. [TÊN HOẠT ĐỘNG IN HOA]**<br>([X] phút) | **Mục tiêu:** [Nêu ngắn gọn 1–2 câu mục tiêu của hoạt động].<br>**Nội dung:** [Mô tả nhiệm vụ cốt lõi HS thực hiện].<br>**Tổ chức thực hiện:**<br>**Giáo viên:**<br>- [Hành động trình chiếu / tổ chức trò chơi / nêu đoạn mã mẫu].<br>- [Đặt câu hỏi đánh số rõ ràng: (1) ...? (2) ...? (3) ...? hoặc giao nhiệm vụ cụ thể].<br>**Học sinh:**<br>- [Hoạt động quan sát, nghiên cứu SGK, trao đổi theo cặp / thực hành cá nhân].<br>- [Viết mã / thực hiện thao tác, trình bày câu trả lời, nhận xét và bổ sung].<br>**Giáo viên:**<br>- [Quan sát, hỗ trợ các điểm khó cụ thể].<br>- [Mời HS trình bày, nhận xét, chuẩn hóa kiến thức].<br>- [Câu dẫn dắt chuyển tiếp sang hoạt động tiếp theo]. | - [Mã NL đặc thù: `NLa`, `NLc`, `NLd`...]<br>- [Mã Năng lực số / Năng lực AI nếu có]<br>- [Năng lực chung: `Tự chủ và tự học`, `Giao tiếp và hợp tác`...]<br>- [Phẩm chất: `Chăm chỉ`, `Trách nhiệm`, `Trung thực`...] | **Câu trả lời dự kiến / Đoạn mã dự kiến / Sản phẩm dự kiến / Lỗi dự kiến:**<br>- [Ghi rõ đáp án từng câu hỏi (1), (2), (3)...]<br>- [Ghi rõ đoạn mã mẫu hoặc mô tả chi tiết sản phẩm hoàn chỉnh kèm ví dụ cụ thể].<br>**Đánh giá:**<br>- [Ghi rõ phương thức: Qua trò chơi, hỏi – đáp / Qua câu trả lời và đoạn mã học sinh viết / Qua thao tác, sản phẩm và bảng tiêu chí đánh giá...]. |

---

### IV. KIỂM TRA, ĐÁNH GIÁ

#### 1. Đánh giá quá trình
Liệt kê 4–5 điểm chạm đánh giá xuyên suốt tiết học:
* Kết quả trò chơi ôn tập / khởi động.
* Câu trả lời và đoạn mã học sinh viết trong hoạt động tìm hiểu kiến thức.
* Quá trình thực hành thiết kế sản phẩm tại phòng máy.
* Việc tìm lỗi, viết prompt và nhận xét gợi ý của AI.
* Phần trình bày, nhận xét và điều chỉnh sản phẩm của học sinh.

#### 2. Tiêu chí đánh giá sản phẩm (Bảng Rubric 3 cột)
Lập **Bảng 3 cột** chấm điểm sản phẩm thực hành (Tổng = `10 điểm`):

| Tiêu chí | Yêu cầu cần đạt | Điểm tối đa |
| :--- | :--- | :---: |
| Cấu trúc và lưu tệp | Tệp có cấu trúc đúng quy chuẩn, lưu đúng tên (`HoTen_Lop_TenBai.ext`) và mở/chạy được. | 2 |
| [Tiêu chí nội dung/kỹ thuật 1] | [Mô tả yêu cầu cụ thể 1] | 2 |
| [Tiêu chí kỹ thuật 2 (Kiến thức mới 1)] | [Mô tả yêu cầu cụ thể 2] | 1 |
| [Tiêu chí kỹ thuật 3 (Kiến thức mới 2)] | [Mô tả yêu cầu cụ thể 3] | 2 |
| [Tiêu chí kỹ thuật phụ / kết hợp] | [Mô tả yêu cầu cụ thể 4] | 1 |
| Nội dung và trình bày | Nội dung rõ ràng, phù hợp chủ đề; trình bày khoa học, dễ đọc / code mạch lạc. | 2 |
| **Tổng** | | **10** |

* **Cách cho điểm dự kiến:**
  * *Tiêu chí 2 điểm:* đạt đầy đủ yêu cầu được 2 điểm; đạt một phần được 1 điểm; chưa thực hiện được 0 điểm.
  * *Tiêu chí 1 điểm:* thực hiện phù hợp được 1 điểm; chưa thực hiện hoặc thực hiện chưa phù hợp được 0 điểm.
  * Giáo viên công bố tiêu chí trước khi học sinh thực hành. Nếu sử dụng để nhận xét thường xuyên, giáo viên có thể góp ý theo từng tiêu chí mà không nhất thiết cho điểm.

---

### V. KẾT QUẢ DỰ KIẾN
Mở đầu bằng cụm: *"Sau tiết học, học sinh:"*
* Tóm tắt 3–4 kết quả đầu ra thực chất mà học sinh làm được sau khi kết thúc 45 phút (nhận biết và sử dụng được công cụ/cú pháp gì, tạo và lưu được sản phẩm gì, viết được prompt nhờ AI hỗ trợ phát hiện lỗi và biết kiểm tra lại gợi ý trước khi sử dụng).
