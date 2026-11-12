# -*- coding: utf-8 -*-
"""
Script tao file:
1. c:/Users/PC/khbd_docx/06_thanh_pham_docx/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
2. c:/Users/PC/khbd_docx/khoi12/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
3. c:/Users/PC/khbd_docx/08_goi_nop_du_gio/Đề cương dự giờ SHCN_Tuan10_12-2_HoQuangHuy.docx
4. c:/Users/PC/khbd_docx/05_ke_hoach_bai_day_output/De_cuong_du_gio_SHCN_Tuan10_12-2_HoQuangHuy.md

Ke thua 100% cau truc XML chuan cua mau:
c:/Users/PC/khbd_docx/06_thanh_pham_docx/Đề cương dự giờ SHCN_Tuan9_12-2_HoQuangHuy.docx
va tuan thu tuyet doi 05 quy tac cua Co Nguyen Thi Thu Phuong.
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
            # Remove B, I, U from base_rPr so we can apply per-segment
            for tag in ("w:b", "w:bCs", "w:i", "w:iCs", "w:u"):
                for child in list(base_rPr.findall(tag, NS)):
                    base_rPr.remove(child)

    # Remove all existing w:r elements from p_elem
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

    # 3. Phan I. MUC DICH - YEU CAU
    set_paragraph_runs(elems[15], [
        ("• Đánh giá toàn diện kết quả hoạt động Tuần 10: ", "B"),
        ("Tiết sinh hoạt chủ nhiệm giúp cô giáo chủ nhiệm Nguyễn Thị Thu Phương và tập thể học sinh lớp 12/2 cùng nhìn nhận, phân tích và đánh giá khách quan kết quả học tập trong các tiết học trên lớp, tình hình chấp hành nề nếp chuyên cần – tác phong cũng như việc hưởng ứng phong trào thi đua chào mừng Ngày Nhà giáo Việt Nam 20/11 trong tuần học thứ 10.", "")
    ])
    set_paragraph_runs(elems[16], [
        ("• Tuyên dương gương tốt và chấn chỉnh kịp thời các tồn tại: ", "B"),
        ("Thông qua buổi sinh hoạt, cô giáo chủ nhiệm kịp thời biểu dương, khích lệ tinh thần học tập nghiêm túc của tập thể lớp và các tổ học sinh đạt nhiều tiết học xếp loại T (Tốt) trong Sổ đầu bài (tiêu biểu ở các môn Hóa học, Lịch sử, Công nghệ); đồng thời cô chỉ rõ các hạn chế còn tồn tại về tình trạng đi học muộn, vắng học, quên đeo bảng tên và tự ý ra vào lớp trong giờ học để nghiêm khắc nhắc nhở, phê bình và thống nhất giải pháp khắc phục triệt để.", "")
    ])
    set_paragraph_runs(elems[17], [
        ("• Xây dựng phương hướng và triển khai nhiệm vụ trọng tâm Tuần 11: ", "B"),
        ("Cô giáo chủ nhiệm định hướng cho tập thể lớp 12/2 xây dựng kế hoạch hành động cụ thể cho tuần học thứ 11 – tuần lễ cao điểm kỷ niệm Ngày Nhà giáo Việt Nam 20/11 (từ ngày 16/11/2026 đến ngày 21/11/2026), trọng tâm là chấn chỉnh dứt điểm nề nếp đi học đúng giờ, đeo bảng tên đầy đủ và phấn đấu đạt 100% tiết học xếp loại T (Tốt) để dâng tặng thầy cô giáo.", "")
    ])
    set_paragraph_runs(elems[18], [
        ("• Phát huy tinh thần dân chủ và ý thức tự quản của học sinh: ", "B"),
        ("Buổi sinh hoạt tạo diễn đàn để các em học sinh lớp 12/2 nghiêm túc tiếp thu ý kiến chỉ đạo của cô giáo chủ nhiệm, đồng thời các em mạnh dạn bày tỏ đề xuất, nguyện vọng chính đáng và rèn luyện tính tự giác, tinh thần trách nhiệm tập thể trong đợt thi đua cao điểm chào mừng ngày 20/11 của Trường THPT Thuận Hoá.", "")
    ])

    # 2. Yeu cau can dat
    set_paragraph_runs(elems[21], [
        ("- Học sinh toàn lớp 12/2 đánh giá đúng đắn kết quả học tập, nề nếp chuyên cần và hoạt động phong trào của tập thể lớp trong tuần 10; trên cơ sở đó, các em tích cực thảo luận và nắm vững kế hoạch hành động cụ thể cho tuần 11.", "")
    ])
    set_paragraph_runs(elems[22], [
        ("- Mỗi học sinh tự nhận thức rõ những ưu điểm và hạn chế của bản thân trong tuần học vừa qua (nhất là về ý thức đi học đúng giờ, đeo bảng tên và giữ kỷ luật giờ học), từ đó các em chủ động xác định hướng phát huy điểm mạnh và biện pháp khắc phục dứt điểm các thiếu sót cá nhân.", "")
    ])
    set_paragraph_runs(elems[28], [
        ("- Tập thể học sinh lớp 12/2 nêu cao ý thức đạo đức, tự giác chấp hành nghiêm chỉnh nội quy của lớp và Trường THPT Thuận Hoá, quyết tâm giành nhiều tiết học tốt và hoa điểm tốt để lập thành tích thiết thực tri ân thầy cô giáo nhân dịp Ngày Nhà giáo Việt Nam 20/11.", "")
    ])

    # 3. Cong tac chuan bi truoc gio sinh hoat
    set_paragraph_runs(elems[31], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương: ", "B"),
        ("Cô theo dõi sát sao bảng thông báo kế hoạch hoạt động của nhà trường và Đoàn trường; cô tổng hợp tình hình thực tế của lớp 12/2 thông qua Sổ đầu bài, sổ Cờ đỏ cùng báo cáo hằng ngày của nhóm giáo sinh thực tập và Ban Cán sự lớp trên các mặt chuyên cần, kỷ luật, học tập, vệ sinh; đồng thời cô chuẩn bị dự thảo kế hoạch công tác tuần 11 để triển khai và thảo luận cùng học sinh.", "")
    ])
    set_paragraph_runs(elems[32], [
        ("- Giáo sinh thực tập Hồ Quang Huy và nhóm giáo sinh kiến tập: ", "B"),
        ("Nhóm giáo sinh phối hợp theo dõi sát sao tình hình chuyên cần, nề nếp đầu giờ của lớp 12/2 trong suốt tuần 10 để báo cáo cô giáo chủ nhiệm; đồng thời giáo sinh chuẩn bị chu đáo sổ dự giờ, thiết kế kịch bản trò chơi, hệ thống câu hỏi và phần thưởng cho hoạt động giao lưu tập thể “Hoa Điểm Tốt – Ngàn Lời Tri Ân Dâng Thầy Cô” ở cuối tiết sinh hoạt.", "")
    ])
    set_paragraph_runs(elems[34], [
        ("- Ban Cán sự lớp và 04 Tổ trưởng: ", "B"),
        ("Lớp trưởng Hồ Quốc Đạt, Lớp phó học tập Nguyễn Phương Linh cùng 04 Tổ trưởng (Đoàn Anh Khoa, Lê Thị Hương Lan, Phạm Văn Bách, Lê Quốc Huy) tổng hợp đầy đủ số liệu chuyên cần, học tập và tình hình chấp hành nội quy của từng tổ viên trong tuần 10.", "")
    ])
    set_paragraph_runs(elems[35], [
        ("- Ban Chấp hành Chi đoàn 12/2: ", "B"),
        ("Bí thư Chi đoàn Trần Văn Thiên Tâm tổng hợp kết quả hưởng ứng phong trào “Tuần học tốt – ngàn lời tri ân” của chi đoàn trong tuần 10 và chuẩn bị triển khai kế hoạch tham gia các hoạt động kỷ niệm Ngày Nhà giáo Việt Nam 20/11 trong tuần 11.", "")
    ])
    set_paragraph_runs(elems[36], [
        ("- Toàn thể học sinh lớp 12/2: ", "B"),
        ("Các thành viên trong lớp trung thực, tự giác đánh giá kết quả rèn luyện cá nhân trong tuần 10 và chuẩn bị các ý kiến đóng góp thiết thực nhằm xây dựng tập thể lớp vững mạnh.", "")
    ])

    # Phan III & IV
    set_paragraph_runs(elems[49], [
        ("• Phương pháp tổ chức hoạt động giáo dục tập thể: ", "B"),
        ("Nhóm giáo sinh kiến tập trực tiếp điều phối trò chơi trí tuệ – giao lưu cuối giờ nhằm tạo tâm thế hào hứng, gắn kết tình bạn và bồi đắp tình cảm tri ân thầy cô giáo nhân dịp Ngày Nhà giáo Việt Nam 20/11 cho học sinh.", "")
    ])
    set_paragraph_runs(elems[52], [
        ("2. Tổng kết tình hình học tập, rèn luyện và hoạt động phong trào Tuần 10: ", "B"),
        ("Ban Cán sự lớp, BCH Chi đoàn và 04 Tổ trưởng báo cáo kết quả học tập trên Sổ đầu bài, tình hình chuyên cần và nề nếp tác phong của tập thể lớp trong tuần 10; Cô giáo chủ nhiệm Nguyễn Thị Thu Phương nhận xét toàn diện ưu điểm, nhược điểm và thống nhất biện pháp chấn chỉnh.", "")
    ])
    set_paragraph_runs(elems[53], [
        ("3. Phổ biến kế hoạch công tác Tuần 11 – Tuần lễ cao điểm kỷ niệm Ngày Nhà giáo Việt Nam 20/11: ", "B"),
        ("Cô giáo chủ nhiệm cùng Lớp trưởng và Bí thư Chi đoàn triển khai nhiệm vụ trọng tâm tuần 11, quán triệt toàn lớp chấm dứt triệt để việc đi học muộn, quên đeo bảng tên và tiếp tục đẩy mạnh đợt thi đua cao điểm “Tuần học tốt – ngàn lời tri ân” chào mừng ngày 20/11.", "")
    ])
    set_paragraph_runs(elems[54], [
        ("4. Tổ chức hoạt động sinh hoạt tập thể: ", "B"),
        ("Giáo sinh thực tập tổ chức trò chơi giao lưu “Hoa Điểm Tốt – Ngàn Lời Tri Ân Dâng Thầy Cô” dành cho tập thể học sinh lớp 12/2.", "")
    ])

    # Phan V. TIEN TRINH TO CHUC SINH HOAT LOP (TBL[56])
    tbl_proc = elems[56]
    proc_rows = tbl_proc.findall("w:tr", NS)

    # Row 2: Hoat dong 2 (Bao cao tinh hinh Tuan 10)
    r2_c0 = proc_rows[2].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r2_c0[1], [("Ban Cán sự lớp và các Tổ trưởng báo cáo tình hình Tuần 10", "B")])

    r2_c1 = proc_rows[2].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r2_c1[0], [("- Cô giáo chủ nhiệm mời Lớp trưởng Hồ Quốc Đạt báo cáo tình hình chấp hành nội quy, chuyên cần và nề nếp của lớp trong tuần 10.", "")])
    set_paragraph_runs(r2_c1[1], [("- Cô giáo chủ nhiệm mời Lớp phó học tập Nguyễn Phương Linh báo cáo tình hình học tập và kết quả đánh giá các tiết học trong Sổ đầu bài của lớp trong tuần qua.", "")])
    set_paragraph_runs(r2_c1[2], [("- Cô giáo chủ nhiệm lần lượt mời 04 Tổ trưởng báo cáo chi tiết tình hình học tập, chuyên cần và tác phong của các tổ viên trong tuần 10:", "")])

    r2_c2 = proc_rows[2].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r2_c2[0], [("- Lớp trưởng Hồ Quốc Đạt nhận xét chung về tình hình chuyên cần, việc đi học đúng giờ và đeo bảng tên của tập thể lớp 12/2 trong tuần 10.", "")])
    set_paragraph_runs(r2_c2[1], [("- Lớp phó học tập Nguyễn Phương Linh báo cáo cụ thể kết quả các tiết học trong Sổ đầu bài của lớp trong tuần qua.", "")])

    # Row 3: Hoat dong 3 (GVCN nhan xet, danh gia tong ket Tuan 10)
    r3_c0 = proc_rows[3].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r3_c0[1], [("Cô giáo chủ nhiệm nhận xét, đánh giá tổng kết Tuần 10 & đưa ra biện pháp khắc phục", "B")])

    r3_c1 = proc_rows[3].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r3_c1[0], [
        ("- Dựa trên báo cáo của Ban Cán sự lớp, Ban Chấp hành Chi đoàn cùng quá trình theo dõi Sổ đầu bài và sổ Cờ đỏ hằng ngày, cô giáo chủ nhiệm Nguyễn Thị Thu Phương đưa ra nhận xét, đánh giá toàn diện về tình hình lớp 12/2 trong tuần 10:", "")
    ])
    set_paragraph_runs(r3_c1[1], [("• Về những mặt làm tốt (Ưu điểm):", "B")])
    set_paragraph_runs(r3_c1[2], [
        ("+ Nhìn chung, tập thể học sinh lớp 12/2 duy trì tốt tinh thần học tập nghiêm túc, tích cực hưởng ứng phong trào thi đua ", ""),
        ("“Tuần học tốt – ngàn lời tri ân”", "I"),
        (" chào mừng Ngày Nhà giáo Việt Nam 20/11 do Đoàn trường phát động.", "")
    ])
    set_paragraph_runs(r3_c1[3], [
        ("+ Qua theo dõi Sổ đầu bài của lớp trong tuần 10, cô giáo chủ nhiệm biểu dương tập thể lớp có hầu hết các tiết học đạt sĩ số hiện diện đầy đủ (vắng 0), giữ trật tự và hăng hái xây dựng bài nên được các giáo viên bộ môn đánh giá cao: tiêu biểu như các tiết học môn ", ""),
        ("Hóa học", "B"),
        (", môn ", ""),
        ("Lịch sử", "B"),
        (" (tiết PPCT 19, 20) và môn ", ""),
        ("Công nghệ", "B"),
        (" (tiết PPCT 19, 20) đều được giáo viên bộ môn nhận xét ", ""),
        ("“Tốt”", "B"),
        (" và xếp loại tiết học đạt loại ", ""),
        ("T (Tốt)", "B"),
        (".", "")
    ])
    set_paragraph_runs(r3_c1[4], [
        ("+ Đặc biệt trong tiết học môn ", ""),
        ("Lịch sử", "B"),
        (", tập thể học sinh nữ của lớp 12/2 đã thể hiện tinh thần học tập tích cực, chuẩn bị bài chu đáo và được giáo viên bộ môn ghi nhận xét khen ngợi trực tiếp vào Sổ đầu bài.", "")
    ])
    set_paragraph_runs(r3_c1[5], [
        ("+ Cô giáo chủ nhiệm ghi nhận, tuyên dương và khích lệ những học sinh tuần 9 còn mắc lỗi nhưng sang tuần 10 đã biết lắng nghe, tiếp thu ý kiến của cô để chủ động sửa đổi và tiến bộ hơn trong học tập.", "")
    ])
    set_paragraph_runs(r3_c1[6], [("• Về những mặt còn tồn tại (Nhược điểm):", "B")])
    set_paragraph_runs(r3_c1[7], [
        ("+ Về chuyên cần và nề nếp đi học đúng giờ: Cô giáo chủ nhiệm nghiêm khắc nhắc nhở tình trạng học sinh vắng học và đi học muộn bị ghi tên vào sổ Cờ đỏ vẫn còn diễn ra ở nhiều cá nhân. Cụ thể: lớp vắng bạn ", ""),
        ("Nhật Anh", "B"),
        (" (buổi sáng Thứ Hai, ngày 09/11/2026) và vắng bạn ", ""),
        ("Minh Nhật", "B"),
        (" (buổi sáng Thứ Ba, ngày 10/11/2026); về lỗi đi học muộn, sổ Cờ đỏ ghi nhận 03 bạn đi học muộn vào buổi học cuối tuần trước gồm bạn ", ""),
        ("Phương Linh", "B"),
        (", bạn ", ""),
        ("Nam Phong", "B"),
        (" và bạn ", ""),
        ("Anh Khoa", "B"),
        (", đồng thời trong buổi sáng Thứ Tư (ngày 11/11/2026) lớp tiếp tục có tới 05 bạn đi học muộn gồm bạn ", ""),
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
        ("+ Về ý thức kỷ luật trong giờ học và tác phong bảng tên: Cô giáo chủ nhiệm đặc biệt phê bình bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" đã vi phạm nội quy nhiều lần trong tuần 10: vào buổi chiều Thứ Hai (ngày 09/11/2026), bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" bị giáo viên bộ môn phê bình vào Sổ đầu bài do ", ""),
        ("tự ý ra vào lớp trong tiết học môn Giáo dục thể chất (GDTC)", "B"),
        ("; tiếp đó vào buổi sáng Thứ Tư (ngày 11/11/2026), bạn ", ""),
        ("Khiêm Gia Bảo", "B"),
        (" tiếp tục vi phạm lỗi ", ""),
        ("không đeo bảng tên", "B"),
        (" nên khi đi ra ngoài đã bị đội Cờ đỏ kiểm tra và ghi nhận lỗi vào sổ Cờ đỏ.", "")
    ])
    set_paragraph_runs(r3_c1[9], [("• Về biện pháp xử lí và khắc phục:", "B")])
    set_paragraph_runs(r3_c1[10], [
        ("+ Cô giáo chủ nhiệm Nguyễn Thị Thu Phương đưa ra biện pháp xử lí kỷ luật tương ứng đối với bạn Khiêm Gia Bảo và nhóm học sinh đi học muộn để giữ nghiêm nề nếp của lớp; đồng thời cô giao các Tổ trưởng kiểm tra chéo bảng tên và đôn đốc tổ viên đi học đúng giờ ngay từ đầu giờ truy bài mỗi buổi sáng trong tuần 11.", "")
    ])

    r3_c2 = proc_rows[3].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r3_c2[0], [
        ("- Tập thể học sinh lớp 12/2 nghiêm túc, chăm chú lắng nghe những lời nhận xét, đánh giá sâu sát và công tâm của cô giáo chủ nhiệm.", "")
    ])
    set_paragraph_runs(r3_c2[1], [
        ("- Các thành viên trong lớp sôi nổi phát biểu ý kiến và đề xuất những giải pháp thiết thực (như lập nhóm nhắc nhau đi học sớm trước giờ trống truy bài 10 phút, Tổ trưởng kiểm tra bảng tên ngay tại cửa lớp và giữ nghiêm trật tự trong giờ GDTC) nhằm thúc đẩy sự tiến bộ của tập thể lớp trong tuần 11.", "")
    ])
    set_paragraph_runs(r3_c2[2], [
        ("- Bạn Khiêm Gia Bảo cùng các bạn đi học muộn đứng dậy lễ phép tự kiểm điểm bản thân trước cô giáo chủ nhiệm và tập thể lớp, đồng thời các em cam kết sẽ khắc phục triệt để những lỗi sai trong tuần 11.", "")
    ])

    # Row 4: Hoat dong 4 (Pho bien ke hoach Tuan 11)
    r4_c0 = proc_rows[4].findall("w:tc", NS)[0].findall("w:p", NS)
    set_paragraph_runs(r4_c0[1], [("Phổ biến kế hoạch, triển khai công tác trọng tâm Tuần 11", "B")])

    r4_c1 = proc_rows[4].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r4_c1[0], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương cùng với Lớp trưởng ", ""),
        ("Hồ Quốc Đạt", "B"),
        (" và Bí thư Chi đoàn ", ""),
        ("Trần Văn Thiên Tâm", "B"),
        (" phổ biến nội dung, kế hoạch hoạt động trọng tâm của tuần 11 (từ ngày 16/11/2026 đến ngày 21/11/2026).", "")
    ])
    set_paragraph_runs(r4_c1[1], [
        ("- Cô giáo chủ nhiệm và Bí thư Chi đoàn phát động tập thể lớp 12/2 bước vào tuần lễ cao điểm chào mừng Ngày Nhà giáo Việt Nam 20/11 với chủ đề ", ""),
        ("“Tuần học tốt – ngàn lời tri ân”", "BI"),
        (".", "")
    ])
    set_paragraph_runs(r4_c1[2], [
        ("- Cô giáo chủ nhiệm nhấn mạnh và quán triệt 05 công việc trọng tâm mà học sinh lớp 12/2 phải thực hiện nghiêm túc trong tuần 11:", "")
    ])
    set_paragraph_runs(r4_c1[3], [
        ("  1. Tập thể lớp quyết tâm khắc phục dứt điểm những tồn tại của tuần 10, tuyệt đối không để tái diễn tình trạng đi học muộn hàng loạt và lỗi quên đeo bảng tên làm trừ điểm thi đua của lớp.", "")
    ])
    set_paragraph_runs(r4_c1[4], [
        ("  2. Toàn thể học sinh giữ vững nề nếp chuyên cần, đi học đúng giờ quy định và đeo bảng tên đầy đủ 100% khi bước vào cổng trường.", "")
    ])
    set_paragraph_runs(r4_c1[5], [
        ("  3. Trong tất cả các tiết học văn hóa cũng như tiết học ", ""),
        ("Giáo dục thể chất (GDTC)", "B"),
        (", học sinh phải chấp hành nghiêm kỷ luật giờ học, tuyệt đối không tự ý ra vào lớp khi chưa có sự đồng ý của giáo viên bộ môn.", "")
    ])
    set_paragraph_runs(r4_c1[6], [
        ("  4. Toàn lớp tiếp tục phát huy tinh thần học tập tốt như ở các môn Hóa học, Lịch sử, Công nghệ tuần qua; học sinh chủ động chuẩn bị bài chu đáo ở nhà và phấn đấu đạt 100% tiết học xếp loại ", ""),
        ("T (Tốt)", "B"),
        (" trên Sổ đầu bài.", "")
    ])
    set_paragraph_runs(r4_c1[7], [
        ("  5. Toàn thể học sinh lớp 12/2 tham gia đầy đủ, trang nghiêm các hoạt động kỷ niệm ", ""),
        ("Ngày Nhà giáo Việt Nam 20/11", "B"),
        (" do Trường THPT Thuận Hoá và Đoàn trường tổ chức với tinh thần ", ""),
        ("“Tuần học tốt – ngàn lời tri ân”", "I"),
        (".", "")
    ])

    r4_c2 = proc_rows[4].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r4_c2[0], [
        ("- Tập thể học sinh lớp 12/2 nghiêm túc, chú ý lắng nghe và ghi chép đầy đủ những thông tin, nhiệm vụ trọng tâm tuần 11 mà cô giáo chủ nhiệm, Bí thư Chi đoàn và Lớp trưởng đã phổ biến.", "")
    ])
    set_paragraph_runs(r4_c2[1], [
        ("- Ban Cán sự lớp, BCH Chi đoàn và 04 Tổ trưởng nắm vững kế hoạch để bám sát, đôn đốc các thành viên đi học đúng giờ, đeo bảng tên đầy đủ và giành nhiều tiết học tốt dâng tặng thầy cô nhân dịp 20/11.", "")
    ])

    # Row 5: Hoat dong 5 (Tong ket phan sinh hoat hanh chinh)
    r5_c1 = proc_rows[5].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r5_c1[0], [
        ("- Cô giáo chủ nhiệm Nguyễn Thị Thu Phương tổng kết lại các nội dung chính của buổi sinh hoạt chủ nhiệm tuần 10.", "")
    ])
    set_paragraph_runs(r5_c1[1], [
        ("- Cô giáo chủ nhiệm yêu cầu cả lớp 12/2 tự giác thực hiện nghiêm túc nội quy của nhà trường và cam kết khắc phục lỗi đi học muộn, quên bảng tên.", "")
    ])
    set_paragraph_runs(r5_c1[2], [
        ("- Cô giáo chủ nhiệm ân cần động viên, khích lệ các em học sinh giữ gìn sức khỏe, phát huy các tiết học đạt loại T (Tốt) và lập nhiều thành tích xuất sắc trong tuần lễ tri ân thầy cô giáo 20/11.", "")
    ])

    # Row 6: Hoat dong 6 (Sinh hoat tap the cuoi gio)
    r6_c1 = proc_rows[6].findall("w:tc", NS)[1].findall("w:p", NS)
    set_paragraph_runs(r6_c1[0], [
        ("• Trò chơi tập thể: “Hoa Điểm Tốt – Ngàn Lời Tri Ân Dâng Thầy Cô”", "B")
    ])
    set_paragraph_runs(r6_c1[1], [
        ("• Cách thức tổ chức và thể lệ: ", "BI"),
        ("Giáo sinh thực tập Hồ Quang Huy cùng nhóm giáo sinh chia lớp 12/2 thành 04 đội thi tương ứng với 04 tổ. Giáo sinh chuẩn bị ", ""),
        ("08 Bông hoa Điểm Tốt", "B"),
        (" ẩn chứa các câu hỏi về truyền thống Tôn sư trọng đạo, ý nghĩa Ngày Nhà giáo Việt Nam 20/11 và nội quy rèn luyện của học sinh Trường THPT Thuận Hoá. Khi giáo sinh đọc câu hỏi, các tổ có 10 giây suy nghĩ và cử đại diện giơ tay giành quyền trả lời. Mỗi câu trả lời đúng ghi được 10 điểm; nếu sai, quyền trả lời nhường cho các tổ còn lại. Kết thúc 08 câu hỏi, tổ có tổng điểm cao nhất sẽ giành chiến thắng và nhận phần thưởng khích lệ từ nhóm giáo sinh thực tập.", "")
    ])
    set_paragraph_runs(r6_c1[2], [
        ("• Hệ thống 08 câu hỏi trong các Bông hoa Điểm Tốt:", "B")
    ])

    r6_c2 = proc_rows[6].findall("w:tc", NS)[2].findall("w:p", NS)
    set_paragraph_runs(r6_c2[0], [
        ("- Tập thể học sinh cả 04 tổ hào hứng lắng nghe giáo sinh thực tập phổ biến luật chơi và sôi nổi giơ tay hái các Bông hoa Điểm Tốt.", "")
    ])

    # Phan VI. BAI HOC KINH NGHIEM SU PHAM RUT RA
    set_paragraph_runs(elems[58], [
        ("Qua buổi dự giờ tiết Sinh hoạt chủ nhiệm Tuần 10 tại lớp 12/2 dưới sự chủ trì của cô giáo chủ nhiệm Nguyễn Thị Thu Phương, bản thân em đã quan sát, học hỏi và đúc kết được những bài học kinh nghiệm sư phạm quý báu sau:", "")
    ])
    set_paragraph_runs(elems[59], [
        ("1. Kinh nghiệm xử lí linh hoạt, khéo léo các tình huống sư phạm và giáo dục học sinh vi phạm: ", "B"),
        ("Bản thân em học hỏi được từ cô Nguyễn Thị Thu Phương phương pháp kết hợp hài hòa giữa sự nghiêm khắc và lòng bao dung. Khi tập thể lớp 12/2 và đặc biệt là tập thể học sinh nữ được giáo viên bộ môn Lịch sử, Hóa học, Công nghệ khen ngợi và xếp loại tiết học đạt loại T (Tốt) trong Sổ đầu bài, cô kịp thời tuyên dương để khích lệ tinh thần cả lớp; ngược lại, đối với tình trạng đi học muộn của một số cá nhân hay lỗi tự ý ra vào lớp trong giờ GDTC và không đeo bảng tên của bạn Khiêm Gia Bảo, cô phê bình rõ ràng, nghiêm túc nhưng luôn phân tích thấu đáo để học sinh nhận ra lỗi và tự giác sửa chữa.", "")
    ])
    set_paragraph_runs(elems[60], [
        ("2. Kĩ năng bồi dưỡng và phát huy năng lực tự quản của đội ngũ cán bộ lớp: ", "B"),
        ("Em nhận thấy việc cô giáo chủ nhiệm giao nhiệm vụ rõ ràng cho Lớp trưởng Hồ Quốc Đạt, Lớp phó học tập Nguyễn Phương Linh, Bí thư Trần Văn Thiên Tâm và 04 Tổ trưởng (Đoàn Anh Khoa, Lê Thị Hương Lan, Phạm Văn Bách, Lê Quốc Huy) trực tiếp theo dõi, báo cáo và kiểm tra chéo bảng tên, giờ giấc đi học đã giúp tiết sinh hoạt diễn ra khách quan, dân chủ và rèn luyện bản lĩnh điều hành cho học sinh.", "")
    ])
    set_paragraph_runs(elems[61], [
        ("3. Kĩ năng điều phối nhịp độ và phân bổ thời gian hợp lí trong tiết sinh hoạt chủ nhiệm: ", "B"),
        ("Qua tiết dự giờ, em tích lũy được kinh nghiệm phân chia thời lượng khoa học giữa phần sơ kết tuần 10, phần quán triệt kế hoạch chấn chỉnh nề nếp – thi đua chào mừng Ngày Nhà giáo Việt Nam 20/11 trong tuần 11 và phần tổ chức trò chơi tập thể, giúp buổi sinh hoạt vừa đảm bảo tính kỷ cương hành chính vừa tạo không khí ấm áp, hào hứng cho học sinh.", "")
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

    print("Successfully generated:", dst_docx)


if __name__ == "__main__":
    build_tuan10_docx()
