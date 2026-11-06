from pathlib import Path
import xml.etree.ElementTree as E
import zipfile, json, subprocess, hashlib
from xml.sax.saxutils import escape

ROOT=Path(__file__).resolve().parent.parent
QA=ROOT/'_kiem_tra'
OUT=ROOT/'Suon_thuyet_trinh_Ung_dung_AI_lam_hoc_lieu_HoQuangHuy.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
E.register_namespace('w',W);E.register_namespace('r',R)
def w(s):return '{'+W+'}'+s
def n(s,attrs=None):return E.Element(w(s),{w(k):str(v) for k,v in (attrs or {}).items()})
def rp(bold=False,italic=False):
    p=n('rPr');p.append(n('rFonts',{'ascii':'Times New Roman','hAnsi':'Times New Roman','eastAsia':'Times New Roman','cs':'Times New Roman'}))
    p.append(n('sz',{'val':26}));p.append(n('szCs',{'val':26}));p.append(n('color',{'val':'000000'}))
    if bold:p.append(n('b'))
    if italic:p.append(n('i'))
    return p
doc=n('document');body=n('body');doc.append(body)
pending=False;links=[]
def para(text='',kind='body',target=None):
    global pending
    p=n('p');pr=n('pPr');p.append(pr)
    if pending:pr.append(n('pageBreakBefore'));pending=False
    if kind=='title':pr.append(n('pStyle',{'val':'Title'}))
    elif kind=='h':pr.append(n('pStyle',{'val':'Heading1'}))
    if kind in ['title','h','label']:pr.append(n('keepNext'))
    pr.append(n('widowControl'))
    pr.append(n('spacing',{'before':0,'after':0,'line':288,'lineRule':'auto'}))
    pr.append(n('jc',{'val':'center' if kind=='title' else 'both' if kind=='body' else 'left'}))
    pr.append(n('ind',{'left':0,'right':0,'firstLine':567 if kind=='body' else 0}))
    for i,line in enumerate(text.split('\n')):
        r=n('r');r.append(rp(kind in ['title','h','label'],kind=='note'))
        if i:r.append(n('br'))
        t=n('t');t.set('{http://www.w3.org/XML/1998/namespace}space','preserve');t.text=line;r.append(t)
        if target:
            link=n('hyperlink');rid='rId'+str(len(links)+10);link.set('{'+R+'}id',rid);links.append((rid,target));link.append(r);p.append(link)
        else:p.append(r)
    body.append(p);return p
def h(t):para(t,'h')
def label(t):para(t,'label')
def b(t):para('• '+t,'list')
def br():
    global pending
    pending=True
def gap():para('','note')
def table(rows,widths):
    t=n('tbl');pr=n('tblPr');t.append(pr)
    pr.append(n('tblW',{'w':sum(widths),'type':'dxa'}));pr.append(n('tblLayout',{'type':'fixed'}))
    borders=n('tblBorders')
    for s in ['top','bottom','left','right','insideH','insideV']:borders.append(n(s,{'val':'single','sz':4,'color':'D9D9D9'}))
    pr.append(borders);m=n('tblCellMar')
    for s,v in [('top',65),('bottom',65),('left',95),('right',95)]:m.append(n(s,{'w':v,'type':'dxa'}))
    pr.append(m);grid=n('tblGrid');t.append(grid)
    for width in widths:grid.append(n('gridCol',{'w':width}))
    for i,rowdata in enumerate(rows):
        row=n('tr');rowpr=n('trPr');row.append(rowpr);rowpr.append(n('cantSplit'))
        if i==0:rowpr.append(n('tblHeader'))
        for j,value in enumerate(rowdata):
            cell=n('tc');cp=n('tcPr');cp.append(n('tcW',{'w':widths[j],'type':'dxa'}));cp.append(n('vAlign',{'val':'center'}))
            if i==0:cp.append(n('shd',{'fill':'E8EDF2'}))
            cell.append(cp)
            for line in value.split('\n'):
                p=n('p');pp=n('pPr');pp.append(n('spacing',{'before':0,'after':0,'line':288,'lineRule':'auto'}));pp.append(n('jc',{'val':'center' if j==0 else 'left'}));p.append(pp)
                r=n('r');r.append(rp(i==0));te=n('t');te.text=line;r.append(te);p.append(r);cell.append(p)
            row.append(cell)
        t.append(row)
    body.append(t)
def slide(num,title,time,items,speech,visual):
    h(f'Slide {num} {title}')
    para(f'Thời lượng {time} giây','note')
    label('Nội dung hiện trên slide')
    for s in items:b(s)
    label('Lời nói gợi ý')
    for p in speech:para(p)
    para('Hình minh hoạ: '+visual,'note')

para('Ứng dụng AI để làm tài liệu học tập','title')
para('Sườn 7 slide và lời nói trong 5 đến 8 phút','title')
para('Người trình bày Hồ Quang Huy','note')
gap()
para('Tôi tổ chức tài liệu và cách làm thành một thư mục có hướng dẫn, dùng AI hỗ trợ tạo học liệu, rồi tự kiểm chứng trước khi sử dụng. Người nghe có thể áp dụng ngay cho một bài học nhỏ, chưa cần biết lập trình hay thiết lập kết nối giữa các ứng dụng.')
h('Điều người nghe cần mang về')
b('Biết sắp tài liệu nguồn, yêu cầu và sản phẩm vào các ngăn dễ tìm.')
b('Biết giao cho AI việc cụ thể, có nguồn và tiêu chí kiểm tra.')
b('Biết tự giải thích, đối chiếu và quyết định thay vì chấp nhận bản AI viết.')
h('Nhịp trình bày mục tiêu 7 phút')
table([
    ['Slide','Nội dung','Thời gian'],
    ['1','Tôi dùng AI để giải quyết việc gì','0:00–0:40'],
    ['2','Mỗi công cụ làm một việc gì','0:40–1:30'],
    ['3','Chia ngăn và viết hướng dẫn cho AI','1:30–2:35'],
    ['4','Nhiều agent cùng làm như thế nào','2:35–3:35'],
    ['5','Ví dụ thật từ đề cương Bài 9','3:35–5:00'],
    ['6','Con người kiểm chứng và tự suy nghĩ','5:00–6:10'],
    ['7','Người nghe có thể bắt đầu ra sao','6:10–7:00'],
], [950,5350,2770])
gap()
para('Bản 7 phút gồm khoảng 30 giây mở xem tài liệu ở slide 5. Lời nói là gợi ý để tập, không cần đọc nguyên văn. Các trang cuối dành cho chuẩn bị và trả lời câu hỏi, không đưa hết lên slide.','note')

br()
slide(1,'Tôi dùng AI để làm tài liệu học tập',40,[
    'Đầu vào: sách, giáo trình, mẫu và yêu cầu.',
    'AI hỗ trợ: soạn nháp, tạo phiếu, định dạng.',
    'Tôi chọn nguồn, kiểm chứng và duyệt bản dùng.',
], [
    'Khi làm một bộ tài liệu học tập, tôi thường có sách, giáo trình, mẫu Word và yêu cầu của thầy. Nếu mỗi lần mở cuộc trò chuyện mới lại gửi và giải thích từ đầu thì khá mất công. Tôi gom chúng vào một thư mục có hướng dẫn để AI tiếp tục làm việc trên cùng bộ tài liệu. Hôm nay tôi chia sẻ cách sắp xếp và giao việc đó. Điều tôi muốn mọi người mang về là cách dùng AI có căn cứ, biết kiểm tra và vẫn hiểu sản phẩm mình tạo ra.',
], 'Ba nhóm Đầu vào → AI hỗ trợ → Tôi kiểm chứng. Chưa cần hiện tên nhiều công cụ.')
gap()
slide(2,'Mỗi công cụ có một vai trò',50,[
    'GitHub: giữ và chia sẻ cả bộ tài liệu cùng hướng dẫn.',
    'Antigravity: bàn làm việc để tôi và AI đọc, sửa file.',
    'OpenCode: thêm một cách gọi AI từ cửa sổ lệnh.',
], [
    'Có thể hình dung GitHub như nơi giữ một bộ hồ sơ có lịch sử sửa đổi. Antigravity là bàn làm việc trên máy: tôi mở thư mục, giao việc cho AI và xem kết quả. Khi cần, tôi mở OpenCode trong terminal, tức cửa sổ gõ lệnh, để dùng model đã cấu hình. Model là bộ máy tạo câu trả lời; agent là trợ lý dùng model cùng công cụ để làm việc với file. GitHub tự nó không đọc sách hay viết giáo án. Điểm tôi chia sẻ ở đây là cách nối tài liệu, hướng dẫn và công cụ thành một quy trình.',
], 'Ba ô có biểu tượng tủ hồ sơ, bàn làm việc và cửa sổ lệnh; dưới mỗi ô chỉ một câu giải thích.')

br()
slide(3,'Chia ngăn để AI biết đọc gì và làm gì',65,[
    'Nguồn: chương trình, giáo trình, sách giáo khoa.',
    'Hướng dẫn: bản đồ, quy định và quy trình.',
    'Kết quả: bản nháp, thành phẩm và phần kiểm tra.',
], [
    'Trong dự án của tôi, chương trình, giáo trình phương pháp, sách giáo khoa và mẫu đề cương nằm ở các ngăn riêng. Mỗi ngăn có file hướng dẫn đuôi chấm em đê. Đây chỉ là file văn bản dễ đọc, giống tờ giấy chỉ đường. README cho biết tài liệu ở đâu. Rule ghi điều phải tuân theo, ví dụ Word phải là A4. Skill ghi cách làm một việc, ví dụ đọc nguồn, soạn, xuất Word rồi kiểm tra từng trang. Nhờ vậy, tôi có thể chia sẻ cả tài liệu lẫn cách làm. Khi đổi công cụ, tôi vẫn yêu cầu agent xác nhận đã đọc hướng dẫn; không mặc định mọi công cụ tự nhận mọi file.',
], 'Ảnh cây thư mục thật của dự án, khoanh README và ba nhóm Nguồn, Hướng dẫn, Kết quả. Không chụp thông tin riêng tư.')
gap()
slide(4,'Chia việc cho nhiều agent',60,[
    'Agent A: soạn bản nháp trong ngăn riêng.',
    'Agent B: đối chiếu nguồn và ghi nhận xét riêng.',
    'Tôi quyết định sửa gì và hợp nhất bản cuối.',
], [
    'Tôi có thể mở AI trong Antigravity và thêm OpenCode trên cùng máy. Một cách phân công là cho bên A soạn phiếu, bên B kiểm tra phiếu với sách. Hai bên dùng chung tài liệu nguồn nhưng ghi kết quả vào hai ngăn khác nhau. Tôi xem nhận xét, yêu cầu sửa và duyệt bản cuối. Cùng thư mục không có nghĩa là hai AI tự biết cuộc trò chuyện của nhau. Nếu mỗi bên làm trên một bản sao thì phải cập nhật và hợp nhất thay đổi. Tôi tránh cho hai agent cùng sửa một file Word. Cách này giúp phân việc, còn kết quả đúng hay sai vẫn phải dựa vào bằng chứng.',
], 'Sơ đồ một nguồn chung tách sang A và B, rồi về người duyệt. Đây là cách phân công minh hoạ, không cần chạy hai AI trực tiếp khi thuyết trình.')

br()
slide(5,'Ví dụ thật từ đề cương Bài 9',85,[
    'Nguồn: SGK, mẫu đề cương và nhận xét của thầy.',
    'Sửa: từ 45 lên 90 phút và đổi cách trình bày hoạt động.',
    'Kiểm tra: nội dung theo nguồn và từng trang Word.',
], [
    'Ví dụ gần nhất là đề cương Bài 9 về tạo danh sách và bảng trong HTML. Bản trước dự kiến cả bài trong 45 phút. Thầy góp ý bài được bố trí hai tiết, tức 90 phút; năng lực chung cần đặt trước năng lực đặc thù; mục tiêu, nội dung và sản phẩm cần đưa ra ngoài bảng. Tôi dùng những nhận xét đó làm yêu cầu sửa, rồi cho AI đồng bộ đề cương, phiếu và học liệu. Tôi mở Word để xem lại và đối chiếu nguồn. Ví dụ này cho thấy một bản trình bày đẹp vẫn cần người có chuyên môn phản biện. Sau đây tôi mở nhanh một phần để mọi người thấy cách tôi kiểm tra.',
], 'Đặt cạnh nhau ảnh nhận xét của thầy và một đoạn sau khi sửa. Nếu dùng ảnh nhận xét, xin phép người liên quan hoặc ẩn tên.')
gap()
h('Minh hoạ trực tiếp trong khoảng 30 giây')
b('Mở cây thư mục có ngăn SGK, mẫu và thành phẩm, khoảng 5 giây.')
b('Mở đề cương ở phần hai tiết và một hoạt động có mục tiêu ngoài bảng, khoảng 15 giây.')
b('Chỉ trang SGK hoặc PDF nguồn và nói mình đang đối chiếu điều gì, khoảng 10 giây.')
para('Câu nói khi thao tác: “Trên máy tôi có extension để xem, sửa DOCX và xem PDF ngay trong IDE. Tôi đối chiếu được nhanh mà không phải rời chỗ làm. Xem được file mới là một bước; tôi còn kiểm tra nguồn, thời lượng và bố cục từng trang.”')
para('Chuẩn bị sẵn các tab. Nếu máy hoặc mạng chậm, dùng hai ảnh đã chụp; không chờ AI tạo tài liệu mới trước khán giả. Khi cần in hoặc nộp, kiểm tra thêm bản Word hoặc PDF xuất cuối vì các trình xem có thể hiển thị khác nhau.','note')

br()
slide(6,'Tôi làm gì để AI không suy nghĩ thay',70,[
    'Trước khi hỏi: tự xác định mục tiêu và viết ý chính.',
    'Khi nhận: hỏi căn cứ và thử một ví dụ cụ thể.',
    'Trước khi dùng: tự giải thích và chịu trách nhiệm.',
], [
    'Tôi giữ lại ba việc cho mình: chọn mục tiêu, đánh giá bằng chứng và quyết định bản dùng. Chẳng hạn, học sinh đang sửa bảng HTML có hai ô dữ liệu nằm ngoài hàng. Em nên tự dự đoán lỗi trước, rồi yêu cầu AI chỉ gợi ý, chưa viết đáp án. Sau đó em tự sửa, chạy thử và giải thích vì sao cần thẻ tr bao các ô. Nếu AI chỉ khuyên thêm viền, em phải nhận ra viền không sửa được lỗi cấu trúc. Cách kiểm tra rất đơn giản: tắt AI đi, mình có giải thích được lựa chọn và làm lại một ví dụ gần giống không? Nếu chưa, mình cần học tiếp trước khi dùng hoặc nộp.',
], 'Ba bước Tự làm trước → Nhận phản hồi → Tự giải thích. Giữ đoạn mã chi tiết ở phần dự phòng, không nhét lên slide.')
gap()
slide(7,'Bắt đầu từ một bài học nhỏ',50,[
    'Một bài học và vài tài liệu đáng tin.',
    'Một hướng dẫn rõ và một sản phẩm nhỏ.',
    'Đọc lại, thử lại rồi mới chia sẻ.',
], [
    'Mọi người có thể bắt đầu bằng một thư mục của một bài học: để tài liệu nguồn vào một ngăn, yêu cầu vào một file và kết quả vào ngăn khác. Giao AI làm một phiếu ngắn rồi tự kiểm tra. Khi dùng ổn, có thể đưa phần được phép chia sẻ lên GitHub để người khác lấy về và tiếp tục làm với công cụ của họ. Điểm mạnh của cách này là giữ được nguồn, quy trình và các lần sửa; không phải giải thích lại mọi thứ từ đầu. Cách làm cần công sắp xếp và kiểm chứng, nhưng áp dụng được cho phiếu học tập, đề cương ôn tập, câu hỏi tự kiểm tra và slide bài học. Tôi chọn việc và duyệt kết quả; AI hỗ trợ tôi thực hiện.',
], 'Một ảnh thư mục tối giản và bốn biểu tượng sản phẩm. Câu cuối dừng ở vai trò của người dùng.')

br()
h('Mẫu nhỏ để người nghe áp dụng')
para('Đây là cấu trúc gợi ý cho người mới, không phải yêu cầu phải đổi toàn bộ dự án hiện có. Bắt đầu bằng một môn và một bài, chưa cần nhiều agent.')
para('bai_hoc/\n  README.md          Bản đồ tài liệu và cách bắt đầu\n  AGENTS.md          Hướng dẫn chung cho công cụ hỗ trợ\n  01_nguon/          Tài liệu có tên và số trang\n  02_yeu_cau/        Mục tiêu và mẫu trình bày\n  03_ban_nhap/       Nơi AI tạo bản để xem\n  04_kiem_tra/       Nhận xét và đối chiếu nguồn\n  05_thanh_pham/     Chỉ bản đã được người dùng duyệt','list')
h('Năm dòng hướng dẫn ban đầu')
b('Đọc README và yêu cầu trước khi làm; nêu các file nguồn đã dùng.')
b('Ghi rõ tên tài liệu và trang cho kiến thức quan trọng; thiếu căn cứ thì hỏi hoặc đánh dấu để tôi kiểm tra.')
b('Không ghi đè tài liệu nguồn; bản mới lưu vào ngăn bản nháp.')
b('Word A4, Times New Roman 13, lề trái 3 cm, còn lại 2 cm; kiểm tra từng trang sau khi xuất.')
b('Góp ý mới được tôi duyệt phải cập nhật vào yêu cầu hiện hành; báo nếu mâu thuẫn với hướng dẫn cũ.')
h('Lệnh giao việc có thể dùng ngay')
para('“Đọc README, yêu cầu và tài liệu về bài đang học. Hãy tạo một phiếu 5 câu phù hợp với mục tiêu tôi ghi, kèm đáp án riêng. Nêu nguồn và trang cho từng nhóm câu hỏi. Lưu bản nháp để tôi đọc, chưa ghi vào ngăn thành phẩm. Chỗ nào chưa đủ căn cứ thì ghi rõ, không tự điền thông tin.”')
h('Mẫu lệnh để luyện tư duy thay vì chép đáp án')
para('“Tôi dự đoán đoạn bảng HTML này sai vì các ô td nằm ngoài tr. Hãy hỏi tôi một câu để kiểm tra lập luận, rồi cho một gợi ý nếu cần. Chưa viết mã hoàn chỉnh. Sau khi tôi sửa, hãy chỉ ra cách tự kiểm tra bằng cấu trúc và kết quả chạy.”')
para('Bài tự kiểm tra: đóng cửa sổ AI, giải thích vai trò của tr, th và td, rồi tự tạo bảng hai cột có một hàng dữ liệu. Việc trình duyệt vẫn hiển thị được không đủ để kết luận mã đúng.','note')

br()
h('Ghi chú để nói chính xác và trả lời câu hỏi')
label('Vì sao tôi chọn GitHub thay cho cách làm trong NotebookLM')
para('Nên nói: “Tôi chọn thư mục và GitHub vì muốn quản lý nhiều loại file, hướng dẫn và lịch sử sửa, rồi dùng chúng với các công cụ AI khác nhau.” NotebookLM có chức năng chia sẻ và xuất nội dung [4]. Vì vậy, tránh khẳng định NotebookLM chỉ cho Gemini dùng hoặc bắt buộc phải có MCP mới dùng được tài liệu ở công cụ khác. Với file đã có trong thư mục và công cụ có quyền đọc, không cần MCP chỉ để đọc những file đó.')
label('Clone repo có làm agent tự hiểu mọi thứ không')
para('Clone là lấy bản sao kho về máy, không sao chép cuộc trò chuyện hay quyền truy cập dịch vụ [1]. README là điểm bắt đầu. Việc tự nạp rule, skill và vị trí file phụ thuộc công cụ [2, 3]. Hãy yêu cầu agent nêu đã đọc gì. PDF scan hoặc file khó đọc có thể cần OCR hay công cụ trích xuất. Người nhận vẫn phải có công cụ và model phù hợp.')
label('Điểm mạnh và giới hạn cần nói cân bằng')
para('Có thể chia sẻ cả cách làm, đổi công cụ và rà lại các lần sửa. Đổi lại, phải tổ chức file và cập nhật yêu cầu. Nhiều model cùng đồng ý chưa phải bằng chứng đúng. Lịch sử Git dễ xem với văn bản .md; thay đổi trong Word hoặc PDF vẫn cần mở bản cũ và mới để đối chiếu. Chỉ chia sẻ học liệu có quyền sử dụng; loại thông tin học sinh, mật khẩu và khoá truy cập khỏi repo công khai.')
label('Nếu chỉ được 5 phút hoặc được nói đủ 8 phút')
para('Bản 5 phút: giữ đủ bảy ý, mỗi slide nói một đoạn ngắn; slide 4 chỉ nêu phân vai, slide 5 dùng ảnh thay thao tác, bỏ phần giới thiệu thuật ngữ dài. Bản 8 phút: giữ nhịp 7 phút rồi thêm khoảng 1 phút hỏi khán giả chọn cách kiểm tra lỗi HTML ở slide 6. Tập có bấm giờ một lần; không đọc các trang hướng dẫn dự phòng.')
h('Nguồn tra cứu khi chuẩn bị')
para('Các nguồn dưới đây hỗ trợ phần mô tả công cụ; ví dụ đề cương dựa trên hồ sơ và ba nhận xét của thầy trong dự án. Tra cứu ngày 04/11/2026.','note')
for title,url in [
    ('[1] GitHub Docs về clone và đồng bộ bản sao kho','https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository'),
    ('[2] Google Antigravity về rule và skill','https://antigravity.google/docs/rules/'),
    ('[2] Google Antigravity về quy trình trong SKILL md','https://antigravity.google/docs/skills'),
    ('[3] OpenCode về chỉ dẫn trong AGENTS md','https://docs.opencode.ai/docs/rules/'),
    ('[4] Google về tạo chia sẻ và xuất nội dung notebook','https://support.google.com/notebooklm/answer/16206563'),
]:para(title,'list',target=url)

sect=n('sectPr');fr=n('footerReference',{'type':'default'});fr.set('{'+R+'}id','rIdFooter');sect.append(fr)
sect.append(n('pgSz',{'w':11906,'h':16838}));sect.append(n('pgMar',{'top':1134,'bottom':1134,'left':1701,'right':1134,'header':500,'footer':567,'gutter':0}));body.append(sect)
styles=n('styles');defaults=n('docDefaults');rd=n('rPrDefault');rd.append(rp());defaults.append(rd);pd=n('pPrDefault');pp=n('pPr');pp.append(n('spacing',{'before':0,'after':0,'line':288,'lineRule':'auto'}));pd.append(pp);defaults.append(pd);styles.append(defaults)
for sid,name in [('Normal','Normal'),('Title','Title'),('Heading1','heading 1')]:
    s=n('style',{'type':'paragraph','styleId':sid});s.append(n('name',{'val':name}));s.append(rp(sid!='Normal'))
    if sid=='Normal':s.set(w('default'),'1')
    if sid=='Heading1':pp=n('pPr');pp.append(n('outlineLvl',{'val':0}));s.append(pp)
    styles.append(s)
footer=n('ftr');fp=n('p');pr=n('pPr');pr.append(n('jc',{'val':'center'}));pr.append(n('spacing',{'before':0,'after':0,'line':288,'lineRule':'auto'}));fp.append(pr)
field=n('fldSimple',{'instr':' PAGE '});r=n('r');r.append(rp());t=n('t');t.text='1';r.append(t);field.append(r);fp.append(field);footer.append(fp)
relsroot=E.Element('Relationships',xmlns='http://schemas.openxmlformats.org/package/2006/relationships')
for rid,typ,target in [('rIdStyles','styles','styles.xml'),('rIdFooter','footer','footer1.xml')]:E.SubElement(relsroot,'Relationship',Id=rid,Type=R+'/'+typ,Target=target)
for rid,target in links:E.SubElement(relsroot,'Relationship',Id=rid,Type=R+'/hyperlink',Target=target,TargetMode='External')
content='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'''
rootrels='''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',content);z.writestr('_rels/.rels',rootrels)
    for name,el in [('word/document.xml',doc),('word/styles.xml',styles),('word/footer1.xml',footer),('word/_rels/document.xml.rels',relsroot)]:z.writestr(name,E.tostring(el,encoding='utf-8',xml_declaration=True))
assert sum([40,50,65,60,85,70,50])==420
print(OUT)
