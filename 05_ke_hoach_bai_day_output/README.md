# NGĂN 05: ĐẦU RA KẾ HOẠCH BÀI DẠY (LESSON PLAN OUTPUT)

> **Vai trò của ngăn này đối với AI Agent:** Đây là nơi lưu trữ toàn bộ các **Kế hoạch bài dạy / Đề cương bài dạy** (dưới cả hai định dạng `.md` và `.docx`) do AI Agent biên soạn theo yêu cầu của người dùng.

---

## 1. Quy ước đặt tên file đầu ra

Mỗi Kế hoạch bài dạy khi hoàn thiện được lưu tại thư mục `05_ke_hoach_bai_day_output/` theo quy tắc đặt tên thống nhất:

* **Bản thảo / Bản xem nhanh Markdown:**
  `KHBD_Tin[Lop]_Bai[So]_Tiet[So]_[TenRutGonKhongDau].md`
  *(Ví dụ: `KHBD_Tin10_Bai20_Tiet1_CauLenhLapFor.md`)*
* **Bản Word hoàn chỉnh (`.docx`) chuẩn định dạng Form mẫu:**
  `Đề cương bài dạy Tin[Lop]_Bai[So]_Tiet[So].docx`
  *(Ví dụ: `Đề cương bài dạy Tin10_Bai20_Tiet1.docx`)*

---

## 2. Hướng dẫn Kỹ thuật cho Agent khi xuất file Word (`.docx`) chuẩn Form mẫu

Để đảm bảo file `.docx` xuất ra giữ nguyên 100% phông chữ (`Times New Roman` 13pt/12pt, `Courier New` cho code), khổ giấy A4, căn lề (`Trái 3cm, Trên/Dưới/Phải 2cm`), đường viền bảng (Table Borders) và Header/Footer như bản mẫu tại `04_form_mau_khbd/`, Agent thực hiện theo quy trình sau:

1. **Sao chép bản mẫu làm khung (`Clone Template`):**
   Sao chép file `04_form_mau_khbd\Đề cương bài dạy Tin12_Bai8_Tiet2 (sv Phan Trần Hồng Anh).docx` sang thư mục `05_ke_hoach_bai_day_output\` với tên bài mới.
2. **Cập nhật nội dung XML (`word/document.xml`) bằng PowerShell (`.NET System.IO.Compression.FileSystem` & `System.Xml`):**
   * Vì file `.docx` thực chất là gói nén ZIP chứa `word/document.xml`, `word/styles.xml`, Agent giữ nguyên toàn bộ `word/styles.xml`, `w:sectPr` (thông số trang và lề) cũng như `w:tblPr` / `w:tcPr` (thuộc tính bảng và ô) của file mẫu, chỉ thay thế văn bản và các hàng dữ liệu (`w:tr`) tương ứng với bài dạy mới.
   * Luôn sử dụng `[System.Text.Encoding]::UTF8` và escape các ký tự đặc biệt trong XML (`<` $\rightarrow$ `&lt;`, `>` $\rightarrow$ `&gt;`, `&` $\rightarrow$ `&amp;`) đặc biệt khi bài dạy có chứa thẻ HTML hoặc toán tử so sánh trong Python/SQL.

---

## 3. Checklist 10 Tiêu chí Kiểm duyệt Chất lượng (QA Checklist) trước khi bàn giao

Trước khi thông báo hoàn thành cho người dùng, Agent **BẮT BUỘC** tự rà soát KHBD theo 10 tiêu chí sau:

- [ ] **1. Đúng Yêu cầu cần đạt (Ngăn 01):** Mục tiêu kiến thức bám sát 100% YCCĐ của Chương trình GDPT 2018 môn Tin học ứng với bài học đó.
- [ ] **2. Đúng Công thức Mục tiêu (Ngăn 02):** Sử dụng động từ hành động quan sát và đo lường được; không dùng từ mơ hồ ("nắm được", "hiểu về" đơn thuần).
- [ ] **3. Đủ 3 nhóm Năng lực & Phẩm chất (Ngăn 01 & 04):** Có đủ Năng lực đặc thù (`NLa`–`NLe`), Năng lực chung, Năng lực số & Năng lực AI, và Phẩm chất cụ thể hóa thành hành vi.
- [ ] **4. Sát ngữ liệu SGK KNTT (Ngăn 03):** Trích dẫn chính xác tên Bảng, Hình, Hoạt động, Hộp kiến thức trong trang SGK của bài học tương ứng.
- [ ] **5. Có Tình huống học tập xuyên suốt (Ngăn 02 & 04):** Mục `III.2` nêu được một bối cảnh thực tiễn gần gũi với học sinh, kết nối liền mạch từ Hoạt động 1 đến Hoạt động 5.
- [ ] **6. Đảm bảo tỷ lệ Thực hành & Phân bố thời gian 45 phút (Ngăn 02 & 04):** Tổng thời gian các hoạt động trong Bảng `III.3` và `III.4` khớp đúng `45 phút`; hoạt động thực hành/thao tác của học sinh chiếm tỷ trọng trọng tâm.
- [ ] **7. Kịch bản Tổ chức thực hiện rõ ràng 3 nhịp (Ngăn 04):** Trong cột 2 của Bảng `III.4`, mỗi hoạt động đều trình bày rõ: `Mục tiêu:` $\rightarrow$ `Nội dung:` $\rightarrow$ `Tổ chức thực hiện:` (`Giáo viên:` giao nhiệm vụ/câu hỏi đánh số $\rightarrow$ `Học sinh:` thực hiện/báo cáo $\rightarrow$ `Giáo viên:` hỗ trợ/chuẩn hóa/dẫn dắt).
- [ ] **8. Có Sản phẩm dự kiến cụ thể (Ngăn 04):** Cột 4 của Bảng `III.4` có đầy đủ câu trả lời dự kiến cho từng câu hỏi `(1), (2), (3)...`, đoạn mã dự kiến chạy đúng cú pháp và mô tả sản phẩm mẫu.
- [ ] **9. Có hoạt động tích hợp AI kiểm tra mã & tư duy phản biện (Ngăn 02 & 04):** Có đoạn mã lỗi điển hình, Prompt mẫu gợi ý cho học sinh và yêu cầu học sinh đối chiếu, kiểm chứng lại gợi ý của AI trước khi dùng.
- [ ] **10. Có Bảng tiêu chí đánh giá (Rubric) thang điểm 10 (Ngăn 02 & 04):** Mục `IV.2` có bảng Rubric 3 cột (`Tiêu chí | Yêu cầu cần đạt | Điểm tối đa`) tổng cộng `10 điểm` kèm hướng dẫn cách cho điểm dự kiến.
