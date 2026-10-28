# NGĂN 02: LÝ LUẬN VÀ PHƯƠNG PHÁP DẠY HỌC TIN HỌC

> **Vai trò của ngăn này đối với AI Agent:** Đây là **cơ sở lý luận sư phạm, phương pháp dạy học (PPDH), kỹ thuật dạy học và quy tắc thiết kế – tổ chức hoạt động** mà người dùng đã cung cấp (biên soạn bởi TS. Nguyễn Thế Dũng – Khoa Tin học, Trường ĐHSP – Đại học Huế). Mọi Kế hoạch bài dạy (KHBD) do Agent tạo ra **BẮT BUỘC** phải tuân thủ triết lý thiết kế, công thức mục tiêu, cấu trúc hoạt động, nguyên tắc thực hành và phương pháp tích hợp AI được quy định trong ngăn này.

---

## 1. Danh mục tài liệu trong ngăn

| STT | Tên file | Nội dung cốt lõi |
| :-: | :--- | :--- |
| 1 | `17_8_Giao trinh LL-PPDH Tin học - NTD.pdf` | Giáo trình đầy đủ *Lý luận và Phương pháp dạy học Tin học* (TS. Nguyễn Thế Dũng). |
| 2 | `Chương 1 - Lý luận và Phương pháp Dạy học Tin học.pptx` | Đối tượng, nhiệm vụ, "Bộ ba" thành phần cốt lõi (Mục tiêu – Nội dung & Phương pháp – Học sinh) và 4 câu hỏi cốt lõi (Dạy để làm gì? Dạy cái gì? Dạy cho ai? Dạy như thế nào?). |
| 3 | `Chuong2_DinhHuongDayHocMonTinHoc.pptx` | Định hướng dạy học theo CT GDPT 2018; 4 thành phần nội dung giáo dục Tin học; 4 tầng nhận thức (Tri Vật – Tri Nhân – Tri Sự – Tri Hành); Tư duy máy tính (*Computational Thinking*). |
| 4 | `Chuong3_PhuongPhapDayHocMonTinHoc.pptx` | Các PPDH cơ bản (Thuyết trình tích cực, Vấn đáp, Minh họa & Live Coding, Thực hành – Luyện tập) và PPDH tích cực (Giải quyết vấn đề PBL, Dạy học dự án PjBL, Dạy học hợp tác, Dạy học khám phá). |
| 5 | `Chuong4_HinhThucToChucVaKyThuatDayHocMonTinHoc.pptx` | Kỹ thuật dạy học tích cực, mô hình Lớp học đảo ngược (*Flipped Classroom*), Dạy học hỗn hợp (*Blended Learning*), 6 chức năng điều hành quá trình dạy học và quy tắc phối hợp PPDH theo loại bài. |
| 6 | `Chuong5_ThietKeKeHoachDayHocMonTinHoc.pptx` | **TRỌNG TÂM 1:** Quy trình thiết kế KHBD theo định hướng phát triển năng lực, công thức viết mục tiêu 4 thành phần, chuỗi 4 hoạt động học tập, nguyên tắc "50% thực hành", kỹ thuật *Scaffolding* và Rubric đánh giá KHBD. |
| 7 | `Chuong6_ToChucDayHocTinHoc.pptx` | **TRỌNG TÂM 2:** Tổ chức dạy thực hành Tin học (PP 4 bước, 3 bước, 6 bước), bài tập thực hành 3 tầng bậc, quản lý phòng máy, xử lý tình huống sư phạm, **Tổ chức dạy học Tin học với sự hỗ trợ của AI**, và thiết kế Rubric đánh giá sản phẩm số. |

---

## 2. Các Nguyên Tắc Sư Phạm Bắt Buộc Agent Phải Tuân Thủ Khi Soạn KHBD

### 2.1. Triết lý trung tâm của KHBD định hướng năng lực (Chương 5)
* **Chuyển dịch trọng tâm:** Từ *"Giáo viên giảng cái gì"* sang *"Học sinh làm được gì và kiến tạo tri thức thông qua hoạt động nào"*.
* **Nguyên tắc xuất phát:** Bắt đầu từ **Yêu cầu cần đạt (YCCĐ)** trong Chương trình GDPT 2018 $\rightarrow$ sau đó mới chọn ngữ liệu, ví dụ và hoạt động từ Sách giáo khoa hoặc thực tiễn (Tuyệt đối không làm ngược lại là chép lại từng trang SGK).

### 2.2. Công thức viết Mục tiêu bài học 4 thành phần (Chương 5 – Mục 5.2.2)
Mọi mục tiêu bài học và mục tiêu từng hoạt động phải tuân theo công thức:
$$\text{[Chủ thể: Học sinh]} + \text{[Động từ hành động quan sát/đo lường được]} + \text{[Nội dung cụ thể]} + \text{[Điều kiện / Tiêu chí đánh giá]}$$

* ❌ **KHÔNG ĐẠT (Cấm dùng đơn độc):** *"Học sinh hiểu về...", "Học sinh nắm được...", "Học sinh biết..."* mà không có hành vi/sản phẩm kiểm chứng.
* ✅ **ĐẠT / TỐT:**
  * *"Nhận biết được chức năng của một số thẻ định dạng kiểu chữ như `<strong>`, `<em>`, `<mark>`, `<sub>`, `<sup>`."*
  * *"Trình bày được cách sử dụng thuộc tính `style` để điều chỉnh màu chữ, phông chữ và cỡ chữ."*
  * *"Viết được chương trình Python sử dụng vòng lặp `for` để in bảng cửu chương của một số nguyên bất kỳ, code chạy đúng kết quả."*
* **Cụ thể hóa Phẩm chất thành hành vi:** Không viết chung chung *"Học sinh có tính chăm chỉ, trung thực"*, mà phải gắn với hành động trong tiết Tin học, ví dụ:
  * *Chăm chỉ:* Tích cực tìm hiểu kiến thức, thực hành viết mã và kiên trì điều chỉnh sản phẩm.
  * *Trách nhiệm:* Hoàn thành nhiệm vụ được giao; sử dụng thiết bị phòng máy và công cụ hỗ trợ đúng mục đích.
  * *Trung thực:* Trình bày đúng quá trình thực hiện và giải thích được việc sử dụng gợi ý của AI khi hoàn thiện sản phẩm.

### 2.3. Nguyên tắc thiết kế Chuỗi hoạt động học tập (Chương 5 – Mục 5.2.3 & 5.5)
Một KHBD phải đảm bảo **3 nguyên tắc vàng**:
1. **Tính liên kết:** Hoạt động sau kế thừa trực tiếp sản phẩm/kết quả của hoạt động trước (xuyên suốt qua một **Tình huống học tập** thống nhất).
2. **Tính tăng tiến (Scaffolding):** Đầu bài hỗ trợ nhiều (mẫu có sẵn, câu hỏi gợi dẫn, code khung); cuối bài học sinh tự lực thực hành tạo sản phẩm cá nhân và kiểm chứng lỗi.
3. **Tính mở:** Hoạt động cuối bài mở ra nhiệm vụ vận dụng thực tế hoặc kết nối với công cụ hiện đại (như sử dụng AI kiểm tra mã và giao nhiệm vụ sau giờ học).

**Mỗi hoạt động học tập trong bảng tiến trình BẮT BUỘC có đủ 4 thành phần:**
1. **Mục tiêu:** Học sinh làm được gì sau hoạt động này?
2. **Nội dung:** Nhiệm vụ học tập cụ thể học sinh phải thực hiện là gì? (Đọc bảng nào, phân tích đoạn mã nào, thảo luận câu hỏi nào, thiết kế sản phẩm gì?).
3. **Tổ chức thực hiện:** Chia rõ kịch bản tương tác 3 nhịp:
   * **Giáo viên (Chuyển giao nhiệm vụ):** Trình chiếu tình huống/đoạn mã mẫu, nêu yêu cầu và câu hỏi định hướng cụ thể `(1)... (2)... (3)...`.
   * **Học sinh (Thực hiện & Báo cáo):** Quan sát, nghiên cứu SGK, trao đổi cặp đôi / thực hành cá nhân trên máy, viết mã, trình bày và nhận xét chéo.
   * **Giáo viên (Quan sát, Hỗ trợ & Chuẩn hóa):** Quan sát, hỗ trợ các lỗi thao tác cụ thể, mời HS trình bày, nhận xét, chuẩn hóa kiến thức và **dẫn dắt chuyển tiếp** sang hoạt động tiếp theo.
4. **Sản phẩm và đánh giá:** Ghi rõ **Câu trả lời dự kiến / Đoạn mã dự kiến / Sản phẩm dự kiến / Lỗi dự kiến** và **Phương thức đánh giá** (Qua trò chơi, hỏi – đáp, đoạn mã, thao tác thực hành, bảng tiêu chí).

### 2.4. Nguyên tắc "Tối giản lý thuyết – Tối đa thực hành (≥ 50%)" (Chương 5 & Chương 6)
* Trong một tiết Tin học 45 phút (đặc biệt là tiết dạy tại phòng máy hoặc tiết thực hành):
  * Hoạt động Mở đầu / Ôn tập & xác định nhiệm vụ: **4 – 7 phút**.
  * Các hoạt động Hình thành kiến thức mới: Chia thành các đơn vị kiến thức nhỏ (**7 – 10 phút/hoạt động**), sử dụng chiến lược **Phân tích ví dụ mẫu** hoặc **Dự đoán – Chạy thử – Đối chiếu**, tuyệt đối không giảng lý thuyết suông quá 20 phút.
  * Hoạt động Thực hành / Luyện tập tạo sản phẩm trên máy: **15 – 22 phút** (chiếm tỷ trọng lớn nhất trong tiết học).
  * Hoạt động Vận dụng / Kiểm tra mã với AI & Giao nhiệm vụ: **5 – 8 phút**.

### 2.5. Phương pháp tổ chức Thực hành & Hỗ trợ học sinh trong phòng máy (Chương 6 – Mục 6.3)
Khi thiết kế hoạt động thực hành trong KHBD, Agent lựa chọn 1 trong 3 phương pháp phù hợp:
1. **Phương pháp 4 bước (Làm mẫu – Bắt chước):** Dành cho kỹ năng cơ bản, mới làm quen công cụ (*Bước 1: Nêu mục tiêu $\rightarrow$ Bước 2: GV làm mẫu $\rightarrow$ Bước 3: HS làm lại theo mẫu $\rightarrow$ Bước 4: HS luyện tập độc lập*).
2. **Phương pháp 3 bước (Thực hành theo quy trình):** Dành cho bài thực hành tổng hợp tạo sản phẩm số (*Bước 1: Nêu yêu cầu sản phẩm & công bố Rubric tiêu chí đánh giá $\rightarrow$ Bước 2: Hướng dẫn quy trình/các yêu cầu kỹ thuật $\rightarrow$ Bước 3: HS thực hành cá nhân/nhóm tạo sản phẩm, lưu tệp theo quy ước `HoTen_Lop_TenSanPham.ext`, GV quan sát hỗ trợ*).
3. **Phương pháp 6 bước (Định hướng hành động / Dự án):** *Thông tin nhiệm vụ $\rightarrow$ Lập kế hoạch $\rightarrow$ Ra quyết định $\rightarrow$ Thực hiện $\rightarrow$ Kiểm tra $\rightarrow$ Đánh giá*.

**Nguyên tắc GV hỗ trợ HS trong lúc thực hành:**
* Áp dụng hỗ trợ phân bậc: **Gợi mở (Đặt câu hỏi Socratic) $\rightarrow$ Chỉ dẫn vị trí lỗi $\rightarrow$ Làm mẫu (cuối cùng)**.
* Khi chọn sản phẩm để nhận xét trước lớp: Chọn **2–3 sản phẩm tiêu biểu** (1 sản phẩm tốt làm mẫu, 1 sản phẩm có lỗi điển hình để cả lớp phân tích và rút kinh nghiệm).

### 2.6. Nguyên tắc Tích hợp AI vào Kế hoạch bài dạy Tin học (Chương 6 – Mục 6.5)
Theo chuẩn của bộ tài liệu này và Form mẫu của người dùng, KHBD môn Tin học hiện đại khuyến khích tích hợp **Năng lực số và Năng lực AI** theo nguyên tắc:
1. **Học sinh dùng AI để phát hiện lỗi (debug) và giải thích lỗi, KHÔNG dùng AI làm thay toàn bộ bài tập:**
   * Giáo viên đưa ra một đoạn mã có lỗi cú pháp/logic điển hình (ví dụ: thiếu dấu `:` trong CSS `style="color blue"`, thiếu thẻ đóng `</strong>`, lỗi thụt lề Python, lỗi `IndexError`...).
   * Học sinh thảo luận tìm lỗi trước, sau đó viết **Prompt** yêu cầu trợ lý AI kiểm tra đoạn mã, giải thích ngắn gọn từng lỗi và đề xuất mã sửa.
2. **Cung cấp Prompt mẫu chuẩn mực trong KHBD:**
   * Ví dụ Prompt mẫu: *"Hãy kiểm tra đoạn mã [HTML/Python/SQL] dưới đây, chỉ ra lỗi liên quan đến [thẻ/cú pháp/thuộc tính], giải thích ngắn gọn từng lỗi và đề xuất mã đã sửa. Giữ nguyên nội dung văn bản của tôi: ..."*
3. **Bắt buộc rèn luyện Tư duy phản biện & Kiểm chứng kết quả AI:**
   * Học sinh phải **đối chiếu kết quả của AI với kiến thức đã học trong SGK** và **chạy thử trên trình duyệt / IDE** trước khi sử dụng, gắn liền với phẩm chất **Trung thực** và **Trách nhiệm**.

### 2.7. Nguyên tắc thiết kế Kiểm tra, Đánh giá & Rubric (Chương 5 – Mục 5.6 & Chương 6 – Mục 6.6)
* Mọi sản phẩm thực hành (trang web HTML, chương trình Python, bảng CSDL, hình vẽ Inkscape...) đều phải có **Bảng tiêu chí đánh giá (Rubric) thang điểm 10** được **công bố trước khi học sinh thực hành**.
* Bảng Rubric chia thành 5–6 tiêu chí rõ ràng (Cấu trúc & lưu tệp; Các yêu cầu kỹ thuật cốt lõi của bài học; Nội dung, thẩm mỹ và trình bày), với mức điểm từng tiêu chí là `1 điểm` hoặc `2 điểm` (tổng = `10 điểm`), kèm theo hướng dẫn cách cho điểm minh bạch.

---

## 3. Cách trích xuất nhanh nội dung các file Slide `.pptx` cho Agent

Nếu Agent cần đọc chi tiết bất kỳ slide nào trong 6 file `.pptx` của ngăn này, sử dụng đoạn mã PowerShell sau:
```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression.FileSystem
$pptx = "c:\Users\PC\khbd_docx\02_ly_luan_va_ppdh\Chuong5_ThietKeKeHoachDayHocMonTinHoc.pptx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($pptx)
$slides = $zip.Entries | Where-Object { $_.FullName -match "^ppt/slides/slide(\d+)\.xml$" } |
    Sort-Object { [int]([regex]::Match($_.FullName, "slide(\d+)\.xml").Groups[1].Value) }
foreach ($s in $slides) {
    $num = [regex]::Match($s.FullName, "slide(\d+)\.xml").Groups[1].Value
    $reader = New-Object System.IO.StreamReader($s.Open(), [System.Text.Encoding]::UTF8)
    [xml]$xml = $reader.ReadToEnd(); $reader.Close()
    $ns = New-Object System.Xml.XmlNamespaceManager(,$xml.NameTable)
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
    Write-Output "=== Slide $num ==="
    $xml.SelectNodes("//a:p", $ns) | ForEach-Object {
        $t = ($_.SelectNodes(".//a:t", $ns) | ForEach-Object { $_.InnerText }) -join ""
        if ($t.Trim() -ne "") { Write-Output $t.Trim() }
    }
}
$zip.Dispose()
```
