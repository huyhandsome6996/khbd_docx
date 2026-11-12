# -*- coding: utf-8 -*-
"""
Script tao file:
1. c:/Users/PC/khbd_docx/06_thanh_pham_docx/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
2. c:/Users/PC/khbd_docx/khoi12/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
3. c:/Users/PC/khbd_docx/08_goi_nop_du_gio/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
4. c:/Users/PC/khbd_docx/05_ke_hoach_bai_day_output/De_cuong_du_gio_SHCN_Tuan10_12-2_HoQuangHuy.md

Dam bao:
- Giu nguyen 100% form chuan cua Co Nguyen Thi Thu Phuong (Times New Roman, danh so trang tu 1, khong tach khung chu ky, 100% cau co day du Chu ngu - Vi ngu).
- Dien dat van phong va noi dung KHAC BIET HOAN TOAN so voi Tuan 9 (dac biet la Hoat dong 6: Tro choi tap the voi kich ban va 08 cau hoi - dap an moi 100%).
"""

import copy
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture")
ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")


def set_paragraph_runs(p_elem, segments):
    """
    segments: list of tuples (text, style_flags)
      where style_flags is a string containing any of 'B', 'I', 'U'
    Preserves pPr and the base rPr (font family, size, color) from the first existing run.
    """
    existing_runs = p_elem.findall("w:r", NS)
    base_rPr = None
    if existing_runs:
        first_rPr = existing_runs[0].find("w:rPr", NS)
        if first_rPr is not None:
            base_rPr = copy.deepcopy(first_rPr)
            for tag in ("w:b", "w:bCs", "w:i", "w:iCs", "w:u"):
                for child in list(base_rPr.findall(tag, NS)):
                    base_rPr.remove(child)

    for r in list(existing_runs):
        p_elem.remove(r)

    for text, flags in segments:
        r_elem = ET.SubElement(p_elem, f"{{{W_NS}}}r")
        rPr = copy.deepcopy(base_rPr) if base_rPr is not None else ET.Element(f"{{{W_NS}}}rPr")
        if "B" in flags:
            ET.SubElement(rPr, f"{{{W_NS}}}b")
            ET.SubElement(rPr, f"{{{W_NS}}}bCs")
        if "I" in flags:
            ET.SubElement(rPr, f"{{{W_NS}}}i")
            ET.SubElement(rPr, f"{{{W_NS}}}iCs")
        if "U" in flags:
            u_el = ET.SubElement(rPr, f"{{{W_NS}}}u")
            u_el.set(f"{{{W_NS}}}val", "single")
        r_elem.append(rPr)
        t_elem = ET.SubElement(r_elem, f"{{{W_NS}}}t")
        if text.startswith(" ") or text.endswith(" ") or "  " in text:
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_elem.text = text


def build_tuan10_docx():
    src_docx = r"c:\Users\PC\khbd_docx\06_thanh_pham_docx\Đề cương dự giờ SHCN_Tuan9_12-2_HoQuangHuy.docx"
    dst_docx = r"c:\Users\PC\khbd_docx\06_thanh_pham_docx\Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx"

    with zipfile.ZipFile(src_docx, "r") as zin:
        doc_xml_bytes = zin.read("word/document.xml")
        other_files = {item.filename: zin.read(item.filename) for item in zin.infolist() if item.filename != "word/document.xml"}

    root = ET.fromstring(doc_xml_bytes)
    body = root.find("w:body", NS)
    elems = list(body)

    # 1. Trang bia & Bang thong tin trang bia
    set_paragraph_runs(elems[6], [("SINH HOẠT CHỦ NHIỆM – TUẦN 10 (LỚP 12/2)", "B")])

    tbl_cover = elems[7]
    cover_r6_c2_p0 = tbl_cover.findall("w:tr", NS)[6].findall("w:tc", NS)[2].findall("w:p", NS)[0]
    set_paragraph_runs(cover_r6_c2_p0, [("Thứ Bảy, 14/11/2026 – Tiết 5 (Buổi sáng)", "")])

    # 2. Quoc hieu - Tieu ngu & Tieu de trang 1
    tbl_qh = elems[9]
    qh_r0_c1_p2 = tbl_qh.findall("w:tr", NS)[0].findall("w:tc", NS)[1].findall("w:p", NS)[2]
    set_paragraph_runs(qh_r0_c1_p2, [("Huế, ngày 14 tháng 11 năm 2026", "I")])

    set_paragraph_runs(elems[11], [("TUẦN 10 – HỌC KÌ I (Từ ngày 09/11/2026 đến ngày 14/11/2026)", "BI")])

    tbl_info = elems[12]
    info_r1_c0_p0 = tbl_info.findall("w:tr", NS)[1].findall("w:tc", NS)[0].findall("w:p", NS)[0]
    set_paragraph_runs(info_r1_c0_p0, [("Ngày dự giờ: ", "B"), ("Thứ Bảy, ngày 14/11/2026", "")])

    # 3. Phan I. MUC DICH - YEU CAU (Viet moi hoan toan so voi Tuan 9)
    set_paragraph_runs(elems[15], [
        ("• Tổng kết, đánh giá thực chất tình hình lớp 12/2 trong Tuần 10: ", "B"),
        ("Giờ sinh hoạt chủ nhiệm tạo điều kiện để cô giáo chủ nhiệm Nguyễn Thị Thu Phương cùng tập thể học sinh lớp 12/2 rà soát, đối chiếu và đánh giá toàn diện chất lượng các giờ học văn hóa trên Sổ đầu bài, ý thức chấp hành quy định chuyên cần – đúng giờ cũng như tác phong đồng phục, bảng tên của học sinh trong tuần học thứ 10.", "")
    ])
    set_paragraph_runs(elems[16], [
        ("• Biểu dương nỗ lực học tập và chấn chỉnh nghiêm các vi phạm nội quy: ", "B"),
        ("Qua buổi sinh hoạt, cô giáo chủ nhiệm kịp thời khen ngợi tinh thần học tập tích cực của tập thể lớp và biểu dương tập thể học sinh nữ khi đạt nhiều tiết học xếp loại T (Tốt) ở các môn Hóa học, Lịch sử, Công nghệ; mặt khác, cô phân tích rõ nguyên nhân và phê bình nghiêm khắc các trường hợp đi học muộn bị ghi sổ Cờ đỏ, nghỉ học hoặc vi phạm kỷ luật giờ học môn Giáo dục thể chất để học sinh tự giác sửa đổi.", "")
    ])
    set_paragraph_runs(elems[17], [
        ("• Quán triệt phương hướng và nhiệm vụ trọng tâm Tuần 11: ", "B"),
        ("Cô giáo chủ nhiệm định hướng cho tập thể lớp 12/2 thống nhất chương trình hành động cụ thể cho tuần học thứ 11 – tuần lễ cao điểm kỷ niệm 44 năm Ngày Nhà giáo Việt Nam (20/11/1982 – 20/11/2026), với mục tiêu cốt lõi là chấm dứt hoàn toàn tình trạng đi học trễ, đeo bảng tên 100% và giữ vững chất lượng các tiết học xếp loại Tốt để tri ân thầy cô giáo.", "")
    ])
    set_paragraph_runs(elems[18], [
        ("• Rèn luyện năng lực tự quản và gắn kết tập thể học sinh cuối cấp: ", "B"),
        ("Tiết sinh hoạt phát huy vai trò điều hành của đội ngũ Ban Cán sự lớp, Ban Chấp hành Chi đoàn và 04 Tổ trưởng; đồng thời buổi sinh hoạt mở ra không gian đối thoại cởi mở để học sinh lớp 12/2 bày tỏ ý kiến xây dựng lớp và nâng cao ý thức trách nhiệm đối với danh dự chung của tập thể.", "")
    ])

    # I.2. Yeu cau can dat (Dien dat moi hoan toan)
    set_paragraph_runs(elems[21], [
        ("- Toàn thể học sinh lớp 12/2 nắm bắt đầy đủ, chính xác số liệu đánh giá về học tập, chuyên cần và kỷ luật của lớp trong tuần 10; từ đó các em hiểu rõ và thống nhất quy trình thực hiện các nhiệm vụ trọng tâm trong tuần 11.", "")
    ])
    set_paragraph_runs(elems[22], [
        ("- Từng cá nhân học sinh tự soi chiếu lại thái độ học tập và tác phong rèn luyện của bản thân trong tuần qua (đặc biệt là việc tuân thủ giờ giấc đến trường, quy định đeo bảng tên và nội quy giờ học), qua đó mỗi em tự đề ra cam kết khắc phục triệt để các thiếu sót.", "")
    ])
    set_paragraph_runs(elems[24], [
        ("- Đội ngũ cán bộ lớp (Lớp trưởng, Lớp phó học tập, Bí thư Chi đoàn và 04 Tổ trưởng) nâng cao kĩ năng đối chiếu Sổ đầu bài, sổ Cờ đỏ, tổng hợp báo cáo ngắn gọn và điều hành thảo luận trước cô giáo chủ nhiệm cùng tập thể lớp.", "")
    ])
    set_paragraph_runs(elems[25], [
        ("- Học sinh lớp 12/2 phát triển kĩ năng lắng nghe tích cực, kĩ năng trình bày ý kiến phản hồi lễ phép, văn minh và kĩ năng hợp tác nhóm thông qua hoạt động giao lưu trí tuệ cuối giờ sinh hoạt.", "")
    ])
    set_paragraph_runs(elems[27], [
        ("- Các em học sinh bồi đắp lòng tự trọng, tinh thần trách nhiệm cá nhân trước tập thể và thái độ trung thực, cầu thị khi tiếp thu lời nhắc nhở của cô giáo chủ nhiệm Nguyễn Thị Thu Phương.", "")
    ])
    set_paragraph_runs(elems[28], [
        ("- Tập thể học sinh lớp 12/2 khắc sâu truyền thống “Tôn sư trọng đạo”, tự giác giữ gìn kỷ cương học đường tại Trường THPT Thuận Hoá và nỗ lực giành nhiều tiết học xếp loại T (Tốt) làm món quà ý nghĩa dâng tặng thầy cô nhân tuần lễ 20/11.", "")
    ])

    # I.3. Cong tac chuan bi truoc gio sinh hoat (Dien dat moi hoan toan)
    set_paragraph_runs(elems[31], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương: ", "B"),
        ("Cô kiểm tra, đối chiếu kết quả đánh giá từng tiết học trong Sổ đầu bài lớp 12/2 với số liệu ghi nhận của đội Cờ đỏ và báo cáo theo dõi hằng ngày từ nhóm giáo sinh thực tập; trên cơ sở đó, cô phân loại cụ thể các ưu điểm cần biểu dương, các trường hợp vi phạm cần chấn chỉnh và xây dựng dự thảo công tác tuần 11.", "")
    ])
    set_paragraph_runs(elems[32], [
        ("- Giáo sinh thực tập Hồ Quang Huy và nhóm giáo sinh kiến tập: ", "B"),
        ("Nhóm giáo sinh trực tiếp điểm danh, nắm bắt tình hình đi học muộn và tác phong của học sinh lớp 12/2 vào các buổi sáng trong tuần 10 để kịp thời báo cáo cô giáo chủ nhiệm; đồng thời giáo sinh hoàn thiện sổ dự giờ và chuẩn bị bộ câu hỏi, bảng điểm, phần quà cho trò chơi tập thể “Chuyến Tàu Tri Ân – Vững Bước Tuần Học Tốt”.", "")
    ])
    set_paragraph_runs(elems[34], [
        ("- Ban Cán sự lớp và 04 Tổ trưởng: ", "B"),
        ("Lớp trưởng Hồ Quốc Đạt, Lớp phó học tập Nguyễn Phương Linh phối hợp cùng 04 Tổ trưởng (Đoàn Anh Khoa, Lê Thị Hương Lan, Phạm Văn Bách, Lê Quốc Huy) rà soát sổ theo dõi của từng tổ, thống kê chính xác danh sách học sinh vắng học, đi học trễ hoặc vi phạm tác phong trong tuần 10.", "")
    ])
    set_paragraph_runs(elems[35], [
        ("- Ban Chấp hành Chi đoàn 12/2: ", "B"),
        ("Bí thư Chi đoàn Trần Văn Thiên Tâm đánh giá tiến độ thực hiện đợt thi đua “Tuần học tốt – ngàn lời tri ân” của đoàn viên chi đoàn 12/2 trong tuần 10 và chuẩn bị nội dung phổ biến kế hoạch tham gia Lễ kỷ niệm Ngày Nhà giáo Việt Nam 20/11 trong tuần 11.", "")
    ])
    set_paragraph_runs(elems[36], [
        ("- Toàn thể học sinh lớp 12/2: ", "B"),
        ("Mỗi học sinh tự kiểm điểm lại quá trình học tập, rèn luyện của bản thân từ đầu tuần 10 đến nay và chuẩn bị các đề xuất cụ thể nhằm giúp lớp chấm dứt tình trạng đi học muộn trong tuần tới.", "")
    ])

    # Phan III. PHUONG PHAP TO CHUC SINH HOAT LOP (Dien dat moi hoan toan)
    set_paragraph_runs(elems[46], [
        ("• Phương pháp phát huy vai trò tự quản của học sinh gắn với sự định hướng của GVCN: ", "B"),
        ("Cô giáo chủ nhiệm Nguyễn Thị Thu Phương giữ vai trò định hướng và điều phối trung tâm; dưới sự dẫn dắt của cô, đội ngũ Ban Cán sự lớp, Bí thư Chi đoàn và 04 Tổ trưởng chủ động trình bày báo cáo tổng hợp số liệu thực tế của lớp 12/2 trong tuần 10.", "")
    ])
    set_paragraph_runs(elems[47], [
        ("• Phương pháp đàm thoại, trao đổi dân chủ và đối thoại mang tính xây dựng: ", "B"),
        ("Học sinh toàn lớp 12/2 tập trung theo dõi báo cáo sơ kết, sau đó các em trực tiếp thảo luận, phân tích nguyên nhân dẫn đến tình trạng đi học muộn và cùng thống nhất giải pháp hỗ trợ nhau khắc phục.", "")
    ])
    set_paragraph_runs(elems[48], [
        ("• Phương pháp biểu dương gương người tốt – việc tốt kết hợp kỷ luật tích cực: ", "B"),
        ("Cô giáo chủ nhiệm phân tích công tâm các kết quả ghi nhận trên Sổ đầu bài và sổ Cờ đỏ; cô kịp thời ngợi khen các tập thể, cá nhân có ý thức xây dựng bài tốt, đồng thời cô nghiêm khắc nhắc nhở các học sinh vi phạm trên tinh thần giáo dục nhân văn để các em nhận lỗi và tiến bộ.", "")
    ])
    set_paragraph_runs(elems[49], [
        ("• Phương pháp giáo dục thông qua hoạt động trò chơi tập thể: ", "B"),
        ("Nhóm giáo sinh thực tập tổ chức sân chơi đố vui kiến thức kết hợp giáo dục truyền thống “Tôn sư trọng đạo” và nội quy nhà trường ở cuối tiết học nhằm mang lại bầu không khí sôi nổi, đoàn kết cho tập thể lớp 12/2.", "")
    ])

    # Phan IV. NOI DUNG CONG VIEC TRONG TAM (Dien dat moi hoan toan)
    set_paragraph_runs(elems[51], [
        ("1. Ổn định nề nếp và kiểm diện đầu giờ: ", "B"),
        ("Cô giáo chủ nhiệm Nguyễn Thị Thu Phương cùng Lớp trưởng Hồ Quốc Đạt kiểm tra sĩ số hiện diện, đánh giá công tác trực nhật vệ sinh và rà soát việc chấp hành đồng phục, đeo bảng tên của học sinh lớp 12/2.", "")
    ])
    set_paragraph_runs(elems[52], [
        ("2. Sơ kết, đánh giá toàn diện tình hình lớp 12/2 trong Tuần 10: ", "B"),
        ("Đội ngũ cán bộ lớp và 04 Tổ trưởng báo cáo chi tiết kết quả các tiết học trên Sổ đầu bài cùng tình hình chuyên cần, đi học đúng giờ của các tổ viên trong tuần 10; Cô giáo chủ nhiệm Nguyễn Thị Thu Phương tổng kết ưu điểm, chỉ rõ các hạn chế và quyết định biện pháp chấn chỉnh.", "")
    ])
    set_paragraph_runs(elems[53], [
        ("3. Triển khai kế hoạch hành động Tuần 11 hướng về Ngày Nhà giáo Việt Nam 20/11: ", "B"),
        ("Cô giáo chủ nhiệm phối hợp cùng Lớp trưởng và Bí thư Chi đoàn phổ biến các nhiệm vụ trọng tâm tuần 11 (từ ngày 16/11/2026 đến ngày 21/11/2026), yêu cầu toàn lớp chấm dứt tình trạng đi học trễ, quên bảng tên và duy trì 100% tiết học xếp loại T (Tốt) để tri ân thầy cô giáo.", "")
    ])
    set_paragraph_runs(elems[54], [
        ("4. Giao lưu sinh hoạt tập thể cuối giờ: ", "B"),
        ("Giáo sinh thực tập Hồ Quang Huy cùng nhóm giáo sinh tổ chức trò chơi tương tác “Chuyến Tàu Tri Ân – Vững Bước Tuần Học Tốt” dành cho 04 tổ của lớp 12/2.", "")
    ])

    # Phan V. TIEN TRINH TO CHUC SINH HOAT LOP (TBL[56])
    tbl_proc = elems[56]
    proc_rows = tbl_proc.findall("w:tr", NS)

    # Row 1: Hoat dong 1 (On dinh to chuc - Dien dat moi)
    r1_c1 = proc_rows[1].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r1_c1[0], [("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương cùng nhóm giáo sinh thực tập bước vào lớp, hướng dẫn học sinh lớp 12/2 ổn định vị trí ngồi và chuẩn bị tâm thế nghiêm túc cho tiết sinh hoạt.", "")])
    set_paragraph_runs(r1_c1[1], [("- Cô giáo chủ nhiệm đề nghị Lớp trưởng Hồ Quốc Đạt kiểm diện và báo cáo cụ thể sĩ số học sinh có mặt, vắng mặt trong buổi sinh hoạt cuối tuần.", "")])
    set_paragraph_runs(r1_c1[2], [("- Cô giáo chủ nhiệm bao quát lớp học, kiểm tra trực tiếp tình trạng vệ sinh phòng học cũng như tác phong đồng phục và việc đeo bảng tên của từng học sinh.", "")])

    r1_c2 = proc_rows[1].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r1_c2[0], [("- Lớp trưởng Hồ Quốc Đạt hô nhịp điều khiển tập thể học sinh lớp 12/2 đứng dậy chỉnh đốn trang phục và chào cô giáo chủ nhiệm cùng nhóm giáo sinh.", "")])
    set_paragraph_runs(r1_c2[1], [("- Lớp trưởng Hồ Quốc Đạt rà soát quân số các tổ và báo cáo chính xác sĩ số hiện diện của lớp 12/2 lên cô giáo chủ nhiệm.", "")])

    # Row 2: Hoat dong 2 (Bao cao tinh hinh Tuan 10 - Dien dat moi)
    r2_c0 = proc_rows[2].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r2_c0[1], [("Ban Cán sự lớp và các Tổ trưởng báo cáo tình hình Tuần 10", "B")])

    r2_c1 = proc_rows[2].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r2_c1[0], [("- Cô giáo chủ nhiệm mời Lớp trưởng Hồ Quốc Đạt trình bày báo cáo tổng hợp về nề nếp chuyên cần, tình hình đi học đúng giờ và ý thức kỷ luật của lớp 12/2 trong tuần 10.", "")])
    set_paragraph_runs(r2_c1[1], [("- Cô giáo chủ nhiệm mời Lớp phó học tập Nguyễn Phương Linh thông qua kết quả xếp loại các tiết học và nhận xét của giáo viên bộ môn trong Sổ đầu bài tuần 10.", "")])
    set_paragraph_runs(r2_c1[2], [("- Cô giáo chủ nhiệm lần lượt gọi 04 Tổ trưởng đứng dậy báo cáo cụ thể tình hình chuyên cần, tác phong bảng tên và thái độ học tập của từng thành viên trong tổ:", "")])

    r2_c2 = proc_rows[2].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r2_c2[0], [("- Lớp trưởng Hồ Quốc Đạt báo cáo chi tiết số liệu chuyên cần, danh sách các bạn vắng học và các trường hợp bị ghi tên trong sổ Cờ đỏ của lớp 12/2 trong tuần 10.", "")])
    set_paragraph_runs(r2_c2[1], [("- Lớp phó học tập Nguyễn Phương Linh tổng hợp số lượng tiết học đạt loại T (Tốt) cùng các lời khen, nhắc nhở của giáo viên bộ môn trên Sổ đầu bài tuần qua.", "")])
    set_paragraph_runs(r2_c2[2], [("- Bốn Tổ trưởng (Đoàn Anh Khoa, Lê Thị Hương Lan, Phạm Văn Bách, Lê Quốc Huy) lần lượt thông báo kết quả theo dõi ưu điểm và tồn tại của từng tổ viên trong tuần 10.", "")])
    set_paragraph_runs(r2_c2[3], [("- Toàn thể học sinh lớp 12/2 tập trung lắng nghe, đối chiếu với kết quả rèn luyện của bản thân và giơ tay đóng góp ý kiến bổ sung.", "")])

    # Row 3: Hoat dong 3 (GVCN nhan xet, danh gia tong ket Tuan 10)
    r3_c0 = proc_rows[3].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r3_c0[1], [("Cô giáo chủ nhiệm nhận xét, đánh giá tổng kết Tuần 10 & đưa ra biện pháp khắc phục", "B")])

    r3_c1 = proc_rows[3].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r3_c1[0], [
        ("- Trên cơ sở số liệu báo cáo của đội ngũ cán bộ lớp, thông tin phản ánh từ nhóm giáo sinh thực tập cùng kết quả kiểm tra Sổ đầu bài và sổ Cờ đỏ, cô giáo chủ nhiệm Nguyễn Thị Thu Phương đưa ra nhận xét tổng kết sâu sát về tình hình lớp 12/2 trong tuần 10:", "")
    ])
    set_paragraph_runs(r3_c1[1], [("• Về những mặt làm tốt (Ưu điểm):", "B")])
    set_paragraph_runs(r3_c1[2], [
        ("+ Tập thể học sinh lớp 12/2 duy trì thái độ học tập nghiêm túc trên lớp, tích cực hưởng ứng đợt thi đua cao điểm ", ""),
        ("“Tuần học tốt – ngàn lời tri ân”", "I"),
        (" hướng về Ngày Nhà giáo Việt Nam 20/11.", "")
    ])
    set_paragraph_runs(r3_c1[3], [
        ("+ Qua kiểm tra Sổ đầu bài tuần 10, cô giáo chủ nhiệm biểu dương tập thể lớp có nhiều buổi học giữ vững sĩ số tuyệt đối (vắng 0), trật tự và tích cực phát biểu xây dựng bài nên được giáo viên các bộ môn đánh giá rất tốt: tiêu biểu là các tiết học môn ", ""),
        ("Hóa học", "B"),
        (", môn ", ""),
        ("Lịch sử", "B"),
        (" (tiết PPCT 19, 20) và môn ", ""),
        ("Công nghệ", "B"),
        (" (tiết PPCT 19, 20) đều được giáo viên bộ môn ghi nhận xét ", ""),
        ("“Tốt”", "B"),
        (" và xếp loại tiết học đạt loại ", ""),
        ("T (Tốt)", "B"),
        (".", "")
    ])
    set_paragraph_runs(r3_c1[4], [
        ("+ Đặc biệt trong các tiết học môn ", ""),
        ("Lịch sử", "B"),
        (", tập thể học sinh nữ của lớp 12/2 đã phát huy tinh thần tự giác chuẩn bị bài, hăng hái tương tác và được giáo viên bộ môn ghi lời khen ngợi trực tiếp vào cột nhận xét của Sổ đầu bài.", "")
    ])
    set_paragraph_runs(r3_c1[5], [
        ("+ Bên cạnh đó, cô giáo chủ nhiệm ghi nhận sự chuyển biến tích cực của những học sinh từng mắc khuyết điểm ở tuần trước, khi bước sang tuần 10 các em đã biết lắng nghe lời khuyên của cô để tự giác điều chỉnh thái độ học tập trên lớp.", "")
    ])
    set_paragraph_runs(r3_c1[6], [("• Về những mặt còn tồn tại (Nhược điểm):", "B")])
    set_paragraph_runs(r3_c1[7], [
        ("+ Về chuyên cần và việc chấp hành giờ giấc đến trường: Cô giáo chủ nhiệm nghiêm khắc phê bình tình trạng học sinh nghỉ học và đi học muộn bị đội Cờ đỏ ghi tên vẫn diễn ra ở nhiều học sinh, ảnh hưởng trực tiếp đến điểm thi đua của lớp. Cụ thể: về chuyên cần, lớp vắng bạn ", ""),
        ("Nhật Anh", "B"),
        (" (sáng Thứ Hai, ngày 09/11/2026) và vắng bạn ", ""),
        ("Minh Nhật", "B"),
        (" (sáng Thứ Ba, ngày 10/11/2026); về nề nếp đi học đúng giờ, sổ Cờ đỏ ghi nhận 03 bạn đi muộn vào buổi học cuối tuần trước gồm bạn ", ""),
        ("Phương Linh", "B"),
        (", bạn ", ""),
        ("Nam Phong", "B"),
        (" và bạn ", ""),
        ("Anh Khoa", "B"),
        (", đặc biệt trong buổi sáng Thứ Tư (ngày 11/11/2026) lớp tiếp tục có tới 05 bạn đi học muộn gồm bạn ", ""),
        ("Văn Bách", "B"),
        (", bạn ", ""),
        ("Gia Khánh", "B"),
        (", bạn ", ""),
        ("Quốc Phong", "B"),
        (", bạn ", ""),
        ("Đức Nghĩa", "B"),
        (" và bạn ", ""),
        ("Quốc Huy", "B"),
        (".", "")
    ])
    set_paragraph_runs(r3_c1[8], [
        ("+ Về kỷ luật giờ học và tác phong đeo bảng tên: Cô giáo chủ nhiệm đặc biệt nhắc nhở và phê bình bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" vì đã liên tiếp vi phạm nội quy trong tuần 10: trong buổi học chiều Thứ Hai (ngày 09/11/2026), bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" bị giáo viên bộ môn ghi phê bình vào Sổ đầu bài do ", ""),
        ("tự ý ra vào lớp trong tiết môn Giáo dục thể chất (GDTC)", "B"),
        ("; đến buổi sáng Thứ Tư (ngày 11/11/2026), bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" lại tiếp tục mắc lỗi ", ""),
        ("không đeo bảng tên", "B"),
        (" nên khi ra ngoài hành lang đã bị đội Cờ đỏ phát hiện và lập biên bản ghi sổ.", "")
    ])
    set_paragraph_runs(r3_c1[9], [("• Về biện pháp xử lí và khắc phục:", "B")])
    set_paragraph_runs(r3_c1[10], [
        ("+ Cô giáo chủ nhiệm Nguyễn Thị Thu Phương yêu cầu bạn Khiêm Gia Bảo cùng các bạn đi học muộn nghiêm túc kiểm điểm và thực hiện hình thức rèn luyện kỷ luật theo quy định của lớp; đồng thời cô giao nhiệm vụ cho 04 Tổ trưởng kiểm tra việc đeo bảng tên của tổ viên ngay tại cửa lớp trước giờ truy bài 05 phút mỗi sáng.", "")
    ])

    r3_c2 = proc_rows[3].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r3_c2[0], [
        ("- Toàn thể học sinh lớp 12/2 giữ thái độ nghiêm túc, cầu thị lắng nghe từng nội dung biểu dương cũng như phê bình của cô giáo chủ nhiệm Nguyễn Thị Thu Phương.", "")
    ])
    set_paragraph_runs(r3_c2[1], [
        ("- Đại diện các tổ tích cực thảo luận và đề xuất giải pháp khắc phục thực tế (như các bạn nhà gần chủ động gọi điện đánh thức các bạn hay đi muộn, Tổ trưởng kiểm tra bảng tên đầu giờ và nhắc nhau giữ kỷ luật trong giờ học GDTC) để không làm ảnh hưởng đến thành tích chung của lớp.", "")
    ])
    set_paragraph_runs(r3_c2[2], [
        ("- Bạn Khiêm Gia Bảo và các học sinh có tên trong danh sách đi học muộn đứng dậy nhận lỗi trước cô giáo chủ nhiệm cùng tập thể lớp, đồng thời các em hứa sẽ chấn chỉnh ngay từ sáng Thứ Hai của tuần 11.", "")
    ])

    # Row 4: Hoat dong 4 (Pho bien ke hoach Tuan 11 - Dien dat moi)
    r4_c0 = proc_rows[4].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r4_c0[1], [("Phổ biến kế hoạch, triển khai công tác trọng tâm Tuần 11", "B")])

    r4_c1 = proc_rows[4].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r4_c1[0], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương phối hợp cùng Lớp trưởng ", ""),
        ("Hồ Quốc Đạt", "B"),
        (" và Bí thư Chi đoàn ", ""),
        ("Trần Văn Thiên Tâm", "B"),
        (" triển khai phương hướng và nhiệm vụ công tác trọng tâm của tuần 11 (từ ngày 16/11/2026 đến ngày 21/11/2026).", "")
    ])
    set_paragraph_runs(r4_c1[1], [
        ("- Cô giáo chủ nhiệm cùng Bí thư Chi đoàn phát động toàn thể đoàn viên, học sinh lớp 12/2 bước vào tuần lễ cao điểm lập thành tích chào mừng Ngày Nhà giáo Việt Nam 20/11 với chủ đề ", ""),
        ("“Tuần học tốt – ngàn lời tri ân”", "BI"),
        (".", "")
    ])
    set_paragraph_runs(r4_c1[2], [
        ("- Cô giáo chủ nhiệm giao 05 chỉ tiêu hành động cụ thể và yêu cầu tập thể học sinh lớp 12/2 thực hiện nghiêm chỉnh trong tuần 11:", "")
    ])
    set_paragraph_runs(r4_c1[3], [
        ("  1. Tập thể lớp 12/2 quyết tâm khắc phục triệt để hạn chế của tuần 10, chấm dứt hoàn toàn hiện tượng đi học muộn tập thể và lỗi quên đeo bảng tên bị ghi vào sổ Cờ đỏ.", "")
    ])
    set_paragraph_runs(r4_c1[4], [
        ("  2. Mỗi học sinh tự giác bố trí thời gian di chuyển hợp lí để có mặt tại lớp trước giờ trống truy bài, đảm bảo đồng phục chỉnh tề và đeo bảng tên 100% trong suốt thời gian ở trường.", "")
    ])
    set_paragraph_runs(r4_c1[5], [
        ("  3. Trong mọi tiết học chính khóa cả buổi sáng lẫn buổi chiều (đặc biệt là tiết môn ", ""),
        ("Giáo dục thể chất", "B"),
        ("), học sinh tuyệt đối tuân thủ sự điều hành của giáo viên bộ môn, không tự ý rời khỏi vị trí hoặc ra vào lớp khi chưa xin phép.", "")
    ])
    set_paragraph_runs(r4_c1[6], [
        ("  4. Toàn lớp phát huy khí thế học tập sôi nổi như các tiết môn Hóa học, Lịch sử, Công nghệ tuần qua; các tổ thi đua chuẩn bị bài kỹ lưỡng ở nhà để đạt 100% tiết học xếp loại ", ""),
        ("T (Tốt)", "B"),
        (" trên Sổ đầu bài trong tuần lễ tri ân thầy cô.", "")
    ])
    set_paragraph_runs(r4_c1[7], [
        ("  5. Chi đoàn 12/2 chuẩn bị chu đáo và tham gia đầy đủ, đúng giờ, trang nghiêm Lễ kỷ niệm ", ""),
        ("Ngày Nhà giáo Việt Nam 20/11", "B"),
        (" cùng các hoạt động tri ân do Trường THPT Thuận Hoá và Đoàn trường tổ chức.", "")
    ])

    r4_c2 = proc_rows[4].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r4_c2[0], [
        ("- Học sinh toàn lớp 12/2 chăm chú lắng nghe và ghi chép đầy đủ 05 nhiệm vụ trọng tâm của tuần 11 vào sổ tay cá nhân để chủ động thực hiện.", "")
    ])
    set_paragraph_runs(r4_c2[1], [
        ("- Đội ngũ Ban Cán sự lớp, BCH Chi đoàn và 04 Tổ trưởng tiếp nhận nhiệm vụ từ cô giáo chủ nhiệm để phân công kiểm tra chéo hằng ngày và đôn đốc cả lớp giành tuần học tốt dâng tặng thầy cô.", "")
    ])

    # Row 5: Hoat dong 5 (Tong ket phan sinh hoat hanh chinh - Dien dat moi)
    r5_c1 = proc_rows[5].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r5_c1[0], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương khái quát lại các kết luận quan trọng của phần sinh hoạt hành chính tuần 10.", "")
    ])
    set_paragraph_runs(r5_c1[1], [
        ("- Cô giáo chủ nhiệm nhắc nhở tập thể lớp 12/2 giữ vững kỷ luật nội quy và thực hiện nghiêm túc lời hứa khắc phục lỗi đi học trễ, quên bảng tên.", "")
    ])
    set_paragraph_runs(r5_c1[2], [
        ("- Cô giáo chủ nhiệm gửi lời động viên ân cần tới toàn thể học sinh lớp 12/2, chúc các em bước vào tuần học thứ 11 với tinh thần rèn luyện tốt nhất để chào mừng Ngày Nhà giáo Việt Nam 20/11.", "")
    ])

    r5_c2 = proc_rows[5].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r5_c2[0], [
        ("- Tập thể học sinh lớp 12/2 nghiêm túc lĩnh hội lời căn dặn của cô giáo chủ nhiệm và đồng thanh thể hiện quyết tâm hoàn thành tốt kế hoạch tuần 11.", "")
    ])

    # Row 6: Hoat dong 6 (Tro choi tap the - VIET MOI HOAN TOAN 100% CA TEN TRO CHOI, THE LE VA 08 CAU HOI + DAP AN)
    r6_c1 = proc_rows[6].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r6_c1[0], [
        ("• Trò chơi tập thể: “Chuyến Tàu Tri Ân – Vững Bước Tuần Học Tốt”", "B")
    ])
    set_paragraph_runs(r6_c1[1], [
        ("• Cách thức tổ chức và thể lệ: ", "BI"),
        ("Giáo sinh thực tập Hồ Quang Huy cùng nhóm giáo sinh thiết kế hành trình gồm ", ""),
        ("08 Trạm Dừng Tri Ân", "B"),
        (" ẩn chứa 08 câu hỏi về truyền thống hiếu học xứ Huế, các tấm gương người thầy trong lịch sử, kiến thức liên môn Lớp 12 và nội quy Trường THPT Thuận Hoá. Lớp 12/2 chia thành 04 đội theo 04 tổ. Ở mỗi trạm, các đội có 10 giây hội ý và giơ tay giành quyền trả lời (đúng được 15 điểm; nếu sai, 03 tổ còn lại được quyền bổ sung để nhận 10 điểm). Kết thúc 08 trạm, tổ có điểm cao nhất sẽ nhận phần quà tuyên dương từ nhóm giáo sinh.", "")
    ])
    set_paragraph_runs(r6_c1[2], [
        ("• Hệ thống 08 câu hỏi tại các Trạm Dừng Tri Ân:", "B")
    ])
    set_paragraph_runs(r6_c1[3], [
        ("  - Câu 1 (Trạm Cố đô Hiếu học): ", "B"),
        ("Ngôi trường trung học lâu đời nào tại Cố đô Huế (thành lập năm 1896) là nơi Chủ tịch Hồ Chí Minh từng theo học và là biểu tượng hiếu học của xứ Huế?", "")
    ])
    set_paragraph_runs(r6_c1[4], [
        ("  - Câu 2 (Trạm Người Thầy Lịch sử): ", "B"),
        ("Trước khi trở thành vị Đại tướng đầu tiên của Quân đội nhân dân Việt Nam, Đại tướng Võ Nguyên Giáp từng giảng dạy bộ môn nào tại Trường Tư thục Thăng Long?", "")
    ])
    set_paragraph_runs(r6_c1[5], [
        ("  - Câu 3 (Trạm Danh ngôn Giáo dục): ", "B"),
        ("Câu danh ngôn tôn vinh nghề giáo: “Dưới ánh mặt trời không có nghề nào cao quý hơn nghề dạy học” là của nhà giáo dục vĩ đại người Séc nào?", "")
    ])
    set_paragraph_runs(r6_c1[6], [
        ("  - Câu 4 (Trạm Thơ văn Tri ân): ", "B"),
        ("Trong bài thơ “Nghe thầy đọc thơ” của Trần Đăng Khoa, em hãy điền hai từ còn thiếu vào cuối câu thơ: “Em nghe thầy đọc bao ngày / Tiếng thơ đỏ nắng, xanh cây .........”?", "")
    ])
    set_paragraph_runs(r6_c1[7], [
        ("  - Câu 5 (Trạm Theo dấu chân Bác): ", "B"),
        ("Năm 1910, trước khi ra đi tìm đường cứu nước, người thầy giáo trẻ Nguyễn Tất Thành đã dừng chân dạy học tại ngôi trường nào ở Phan Thiết?", "")
    ])
    set_paragraph_runs(r6_c1[8], [
        ("  - Câu 6 (Trạm Giờ học Tốt): ", "B"),
        ("Trong Sổ đầu bài Trường THPT Thuận Hoá, khi lớp đảm bảo sĩ số (vắng 0), giữ trật tự và tích cực xây dựng bài như tiết Lịch sử, Hóa học, Công nghệ vừa qua, giáo viên bộ môn sẽ xếp loại tiết học ở mức cao nhất bằng chữ cái viết tắt nào?", "")
    ])
    set_paragraph_runs(r6_c1[9], [
        ("  - Câu 7 (Trạm Tác phong Học đường): ", "B"),
        ("Theo nội quy Trường THPT Thuận Hoá, khi bước qua cổng trường đầu mỗi buổi học, ngoài việc đi học đúng giờ và mặc đồng phục, học sinh bắt buộc phải đeo vật dụng nào trên ngực áo?", "")
    ])
    set_paragraph_runs(r6_c1[10], [
        ("  - Câu 8 (Trạm Kết nối Tri thức Tin học 12): ", "B"),
        ("Vận dụng Bài 10 môn Tin học 12 vừa học: Khi làm trang web báo tường chào mừng ngày 20/11, để tạo siêu liên kết (hyperlink) dẫn tới trang tri ân thầy cô, học sinh cần dùng thẻ HTML nào và thuộc tính nào?", "")
    ])

    r6_c2 = proc_rows[6].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r6_c2[0], [
        ("- Tập thể học sinh cả 04 tổ hào hứng lắng nghe giáo sinh thực tập phổ biến thể lệ hành trình và sôi nổi hội ý, giơ tay chinh phục 08 Trạm Dừng Tri Ân.", "")
    ])
    set_paragraph_runs(r6_c2[1], [
        ("• Đáp án câu trả lời của học sinh lớp 12/2:", "B")
    ])
    set_paragraph_runs(r6_c2[2], [
        ("  - Câu 1: ", "B"),
        ("Trường Quốc Học Huế (Trường THPT Chuyên Quốc Học – Huế).", "")
    ])
    set_paragraph_runs(r6_c2[3], [
        ("  - Câu 2: ", "B"),
        ("Bộ môn Lịch sử.", "")
    ])
    set_paragraph_runs(r6_c2[4], [
        ("  - Câu 3: ", "B"),
        ("Nhà giáo dục Comenxki (Jan Amos Comenius).", "")
    ])
    set_paragraph_runs(r6_c2[5], [
        ("  - Câu 4: ", "B"),
        ("Hai từ “quanh nhà” (“Tiếng thơ đỏ nắng, xanh cây quanh nhà”).", "")
    ])
    set_paragraph_runs(r6_c2[6], [
        ("  - Câu 5: ", "B"),
        ("Trường Dục Thanh (tại Phan Thiết, Bình Thuận).", "")
    ])
    set_paragraph_runs(r6_c2[7], [
        ("  - Câu 6: ", "B"),
        ("Chữ T (Xếp loại tiết học đạt loại Tốt).", "")
    ])
    set_paragraph_runs(r6_c2[8], [
        ("  - Câu 7: ", "B"),
        ("Bảng tên học sinh (thẻ học sinh).", "")
    ])
    set_paragraph_runs(r6_c2[9], [
        ("  - Câu 8: ", "B"),
        ("Thẻ <a> kết hợp với thuộc tính href (<a href=\"đường_dẫn\">...</a>).", "")
    ])

    # Phan VI. BAI HOC KINH NGHIEM SU PHAM RUT RA (Dien dat moi hoan toan so voi Tuan 9)
    set_paragraph_runs(elems[58], [
        ("Thông qua việc trực tiếp theo dõi nề nếp lớp trong suốt tuần học thứ 10 và dự giờ tiết Sinh hoạt chủ nhiệm tại lớp 12/2 do cô giáo chủ nhiệm Nguyễn Thị Thu Phương chủ trì, bản thân em đã rút ra được 04 bài học kinh nghiệm sư phạm thiết thực sau:", "")
    ])
    set_paragraph_runs(elems[59], [
        ("1. Nghệ thuật kết hợp giữa động viên, khích lệ kịp thời với sự nghiêm minh khi uốn nắn học sinh: ", "B"),
        ("Bản thân em học tập được từ cô Nguyễn Thị Thu Phương cách đánh giá học sinh hết sức công tâm và giàu tính giáo dục. Khi tập thể lớp 12/2 và đặc biệt là tập thể học sinh nữ được giáo viên các môn Lịch sử, Hóa học, Công nghệ khen ngợi và xếp loại tiết học đạt loại T (Tốt) trong Sổ đầu bài, cô dành lời biểu dương trang trọng để nhân rộng phong trào học tốt; đồng thời, trước tình trạng đi học muộn của một số cá nhân hay việc bạn Khiêm Gia Bảo tự ý ra vào lớp trong tiết GDTC và quên đeo bảng tên, cô phân tích thẳng thắn cái sai nhưng vẫn mở ra cơ hội để học sinh tự nhận thức và cam kết sửa đổi.", "")
    ])
    set_paragraph_runs(elems[60], [
        ("2. Phương pháp phối hợp chặt chẽ các kênh thông tin và phát huy vai trò tự quản của cán bộ lớp: ", "B"),
        ("Em nhận thấy cô giáo chủ nhiệm luôn đối chiếu chặt chẽ giữa Sổ đầu bài, sổ Cờ đỏ với báo cáo của nhóm giáo sinh thực tập và đội ngũ cán bộ lớp (Lớp trưởng Hồ Quốc Đạt, Lớp phó học tập Nguyễn Phương Linh, Bí thư Trần Văn Thiên Tâm cùng 04 Tổ trưởng Đoàn Anh Khoa, Lê Thị Hương Lan, Phạm Văn Bách, Lê Quốc Huy). Việc giao cho các Tổ trưởng trực tiếp kiểm tra chéo bảng tên và đôn đốc giờ giấc đầu giờ đã giúp học sinh rèn luyện ý thức trách nhiệm tập thể rất hiệu quả.", "")
    ])
    set_paragraph_runs(elems[61], [
        ("3. Kĩ năng thiết kế hoạt động giáo dục ngoài giờ lên lớp gắn kết với chủ đề tháng 11 và kiến thức bộ môn: ", "B"),
        ("Qua quá trình chuẩn bị và tổ chức trò chơi “Chuyến Tàu Tri Ân – Vững Bước Tuần Học Tốt”, bản thân em học được cách lồng ghép khéo léo các câu hỏi về truyền thống Tôn sư trọng đạo, lịch sử xứ Huế, nội quy nhà trường cùng kiến thức môn Tin học 12 vào hoạt động tập thể, giúp tiết sinh hoạt cuối tuần vừa nghiêm túc vừa sôi nổi, bổ ích.", "")
    ])
    set_paragraph_runs(elems[62], [
        ("4. Bài học về việc rèn luyện bản lĩnh đứng lớp, ngôn phong chuẩn mực và khả năng bao quát học sinh của giáo sinh: ", "B"),
        ("Khi trực tiếp đứng trước tập thể học sinh khối 12 để điều hành hoạt động giao lưu, bản thân em tự đúc kết rằng người giáo viên cần chuẩn bị kịch bản kỹ lưỡng, diễn đạt lưu loát bằng các câu văn trọn vẹn chủ ngữ – vị ngữ, giữ âm lượng giọng nói rõ ràng, dứt khoát và luôn bao quát đều khắp cả 04 tổ để khích lệ mọi học sinh cùng tham gia.", "")
    ])

    # Phan VII. Chu ky cuoi trang
    tbl_sig = elems[67]
    sig_r0_c1_p0 = tbl_sig.findall("w:tr", NS)[0].findall("w:tc", NS)[1].findall("w:p", NS)[0]
    set_paragraph_runs(sig_r0_c1_p0, [("Huế, ngày 14 tháng 11 năm 2026", "I")])

    # Write output docx
    new_doc_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    os.makedirs(os.path.dirname(dst_docx), exist_ok=True)
    with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("word/document.xml", new_doc_xml)
        for fname, data in other_files.items():
            zout.writestr(fname, data)

    # Copy to khoi12 and 08_goi_nop_du_gio for instant convenience
    copy_targets = [
        r"c:\Users\PC\khbd_docx\khoi12\Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx",
        r"c:\Users\PC\khbd_docx\08_goi_nop_du_gio\Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx",
    ]
    for target in copy_targets:
        os.makedirs(os.path.dirname(target), exist_ok=True)
        shutil.copy2(dst_docx, target)

    print("Successfully generated completely rewritten Week 10 SHCN:", dst_docx)


if __name__ == "__main__":
    build_tuan10_docx()
