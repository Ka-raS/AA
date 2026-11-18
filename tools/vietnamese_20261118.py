"""Concise Vietnamese reading edition, with every project form represented."""
import sys,re,json
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree
from revise_20261118 import base_doc,table,paragraph,font_run,DATA,STAKES,TMP,ROOT,MINH,DAI,HIEU,A,THU,revised_risks
from dlib import q

NAMES={MINH:'Minh',DAI:'Đại',HIEU:'Hiếu',A:'A',THU:'Thu'}
PHASE=['Yêu cầu và lập kế hoạch','Thiết kế hệ thống và dữ liệu','Phát triển MVP','Phát triển đầy đủ chức năng','Kiểm thử và QA','Triển khai và bàn giao','Quản lý dự án']
MILESTONE_VI=[['Hoàn thành Charter','Rà soát SRS/RTM/phạm vi','Customer ký SRS/RTM'],['Có đủ thiết kế để review','Customer review thiết kế','Customer duyệt thiết kế'],['Hoàn thành code MVP','Customer xem demo MVP','Hoàn thiện MVP nội bộ'],['Hoàn thành code đầy đủ chức năng','Customer/Sponsor xem demo','Hoàn thiện đầy đủ chức năng'],['Có mẫu test case và test case','Xong test tải/bảo mật/thiết bị','RC sẵn sàng UAT'],['RC trên staging','Customer ký nghiệm thu UAT','Phát hành và bàn giao'],['Sponsor duyệt PMP','Sponsor review báo cáo cuối','Sponsor ký đóng dự án']]
WPNAMES='''Họp lấy yêu cầu và viết Charter|Viết SRS|Lập RTM và tiêu chí nghiệm thu|Viết Scope Statement, WBS và Dictionary|Nghiên cứu người dùng và ứng dụng tương tự|Rà soát SRS, RTM với Customer và Sponsor|Sửa SRS và RTM
Thiết kế kiến trúc|Thiết kế cơ sở dữ liệu|Thiết kế wireframe|Thiết kế mockup và prototype|Viết thiết kế chi tiết|Định nghĩa API chung|Thiết lập kho mã, CI/CD, staging và CDN|Rà soát thiết kế với Customer|Sửa mockup và prototype|Sửa thiết kế và API
Phát triển API xác thực và người dùng|Phát triển API phát nhạc và tích hợp CDN|Phát triển WebSocket Live Premiere|Xây màn hình phát nhạc và Live Premiere|Kiểm thử tích hợp MVP|Demo MVP với Customer|Sửa lỗi MVP và phản hồi demo
Phát triển API khám phá và thư viện|Phát triển API xã hội, thông báo, admin, hóa đơn|Xây màn hình tài khoản, khám phá, thư viện, xã hội|Xây màn hình tải, video, mua hàng, admin|Kiểm thử tích hợp đầy đủ chức năng|Demo đầy đủ với Customer và Sponsor|Sửa lỗi chức năng và phản hồi demo
Viết kế hoạch kiểm thử và chỉ số chất lượng|Lập mẫu test case và test case|Kiểm thử tải, bảo mật và thiết bị|Sửa lỗi tải, bảo mật và thiết bị|Viết báo cáo tổng hợp kiểm thử|Rà soát báo cáo với Customer|Sửa lỗi sau rà soát kiểm thử
Đưa release candidate lên staging|Sửa lỗi Customer báo trong UAT|Triển khai production|Hỗ trợ sau phát hành|Nộp ứng dụng lên App Store và Google Play|Đào tạo quản trị viên|Viết hướng dẫn người dùng, admin, release notes|Viết hướng dẫn triển khai và vận hành|Rà soát bàn giao với Customer|Sửa tài liệu bàn giao
Lập Project Management Plan|Mua cloud, CDN, thiết bị và công cụ test|Theo dõi tiến độ, báo cáo trạng thái và chi phí|Theo dõi rủi ro và yêu cầu thay đổi|Viết bài học và báo cáo cuối kỳ|Rà soát báo cáo cuối kỳ với Sponsor|Sửa báo cáo và đóng dự án'''
WPVI={}
for phase,line in enumerate(WPNAMES.splitlines(),1):
    src=[r for r in DATA['tables'][44+phase] if re.fullmatch(r'1\.\d\.\d+',r[0]) and r[2]!='-']
    vals=line.split('|');assert len(src)==len(vals),(phase,len(src),len(vals))
    WPVI.update({r[0]:v for r,v in zip(src,vals)})

FR=[
('Đăng ký và đăng nhập','Email/số điện thoại và mật khẩu.','Thông tin đúng: tạo tài khoản/đăng nhập được; sai: bị từ chối.','Kiểm thử chức năng + UAT'),
('Quản lý hồ sơ','Đổi ảnh, tiểu sử, mật khẩu; xóa tài khoản và dữ liệu cá nhân.','Lưu sau đăng nhập lại; tài khoản đã xóa không đăng nhập được, dữ liệu cá nhân bị xóa.','Kiểm thử chức năng + UAT'),
('Tìm kiếm','Theo bài hát, nghệ sĩ, album, thể loại; gợi ý khi nhập.','Kết quả khớp từ khóa ở trường tương ứng; API p90 < 200 ms.','Chức năng + hiệu năng'),
('Phát nhạc','Chọn chất lượng; phát, dừng, tua, chuyển bài, lặp.','Các thao tác chạy được ở mọi mức chất lượng hỗ trợ.','Chức năng + UAT'),
('Phát video','360p/720p/1080p; toàn màn hình ngang.','Phát được ở từng độ phân giải; xoay ngang toàn màn hình.','Kiểm thử chức năng'),
('Playlist','Tạo, sửa, xóa, chia sẻ; thêm/bỏ bài hát.','Các thay đổi được giữ sau đăng nhập lại.','Kiểm thử chức năng'),
('Thư viện','Lưu nghệ sĩ, album, bài hát và playlist của người khác.','Mục thêm xuất hiện và được lưu trong thư viện.','Kiểm thử chức năng'),
('Gợi ý nhạc','Dựa vào lịch sử nghe, thể loại/nghệ sĩ và người nghe tương tự.','Mỗi bài khớp thể loại/nghệ sĩ trong 50 lượt nghe gần nhất hoặc người nghe tương tự đã nghe; ghi lượt thích, bỏ qua, phát.','Chức năng + UAT'),
('Live Premiere','Vào buổi ra mắt, nghe đồng bộ và chat.','15,000 người đồng thời: lệch nhạc <= 2 s; tin nhắn tới trong <= 1 s.','Kiểm thử tải WebSocket + UAT'),
('Nghe offline','Tải bài hát về thiết bị.','Tắt mạng vẫn phát hết bài đã tải.','Kiểm thử chức năng'),
('Tương tác','Yêu thích, đánh giá, bình luận và lịch sử nghe.','Lưu các thao tác; lịch sử xếp mới nhất trước.','Kiểm thử chức năng'),
('Quản trị nội dung','Thêm, xem, sửa, xóa bài hát/album/nghệ sĩ; quản lý tài khoản.','Mỗi thao tác CRUD thành công; thay đổi hiển thị cho người dùng.','Chức năng + tích hợp'),
('Thông báo','Bản phát hành, gợi ý mới và trả lời bình luận.','Gửi thông báo cho từng sự kiện nêu trên.','Kiểm thử chức năng'),
('Mua tính năng trong ứng dụng','Store xử lý thanh toán; backend kiểm tra hóa đơn rồi mở tính năng.','Mua sandbox chỉ mở tính năng sau xác minh; không lưu thẻ.','Kiểm thử sandbox của store')]
NFR=[
('Tải đồng thời','Ra mắt >= 9,000 người; thử tới 15,000.','15,000 VU: lỗi < 1%, API p90 < 200 ms.','Kiểm thử tải'),
('Độ trễ API','90% truy vấn < 200 ms.','p90 < 200 ms khi có tải.','Kiểm thử tải/hiệu năng'),
('Mở ứng dụng','Cold start < 2 s; warm start < 1 s.','Trung vị 10 lần mở trên từng thiết bị trong 4 máy.','Kiểm thử hiệu năng'),
('Pin','Tiêu hao < 6%/giờ nghe qua Wi-Fi, tắt màn hình, âm lượng 50%.','Đo trên 4 máy; quy đổi mAh về pin 4,000 mAh.','Kiểm thử thiết bị'),
('Bảo mật','bcrypt cost >= 10; JWT access <= 1 h; TLS >= 1.2; refresh token trong Keychain/Keystore/httpOnly cookie.','Đúng cấu hình; không còn phát hiện critical/high.','Kiểm thử bảo mật + review mã'),
('Dễ dùng','Apple HIG / Material Design 3; UAT >= 4.2/5.','Review đạt guideline và khảo sát UAT >= 4.2/5.','Review thiết kế + UAT'),
('Mở rộng','Container; thiết kế mở rộng ngang tới 50,000 người ở giai đoạn sau.','Review xác nhận hướng mở rộng; không kiểm thử tải 50,000 trong phạm vi này.','Review kiến trúc'),
('Tài liệu API','OpenAPI 3.0; endpoint dưới /api/v1.','Có đặc tả và version cho mọi endpoint.','Review thiết kế'),
('Độ phủ test','Backend tự động >= 70%.','Báo cáo coverage >= 70%.','Báo cáo test tự động'),
('Sao lưu','PostgreSQL hằng ngày; RPO 24 h; RTO 4 h.','Có lịch backup; diễn tập phục hồi <= 4 h, mất dữ liệu <= 24 h.','Diễn tập backup/restore'),
('Vận hành','Log tập trung, giám sát uptime; cảnh báo lỗi/độ trễ.','Cấu hình và kiểm tra cảnh báo trước go-live.','Review sẵn sàng vận hành'),
('Ngôn ngữ','Việt/Anh cho mobile, web client và admin.','Mọi màn hình đủ hai ngôn ngữ; giữ lựa chọn sau khởi động lại.','Chức năng + UAT')]

def newpage(d,title):
    if d.paragraphs and any(p.text for p in d.paragraphs):d.add_page_break()
    d.add_heading(title,level=1)

def p(d,s):return paragraph(d,s,9)
def h(d,s):return d.add_heading(s,level=2)
def tb(d,headers,data,widths=None,size=8):
    if widths is None:widths=[int(9900/len(headers))]*len(headers)
    return table(d,headers,data,widths,size)
def fmt(n):return f'{int(n):,}'
def vi_priority(s):return {'Must-have':'Bắt buộc','Should-have':'Nên có','Could-have':'Có thể có'}.get(s,s)
def actor(s):return {'Customer':'Customer','Sponsor':'Sponsor','Project Manager':'PM'}.get(s,s)
def num(s):return int(s.replace(',',''))
def vi_dates(s):
    months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12,
    'January':1,'February':2,'March':3,'April':4,'June':6,'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}
    s=s.replace(' (End of Month 5)',' (cuối tháng 5)')
    s=re.sub(r'([A-Z][a-z]+) (\d+)-(\d+), (\d{4})',lambda m:f'{m[2]}–{m[3]}/{months[m[1]]:02d}/{m[4]}',s)
    s=re.sub(r'([A-Z][a-z]+) (\d+) - ([A-Z][a-z]+) (\d+), (\d{4})',lambda m:f'{m[2]}/{months[m[1]]:02d}–{m[4]}/{months[m[3]]:02d}/{m[5]}',s)
    return re.sub(r'([A-Z][a-z]+) (\d+), (\d{4})',lambda m:f'{m[2]}/{months[m[1]]:02d}/{m[3]}',s)

def build():
    d=base_doc();s=d.sections[0]
    s.left_margin=s.right_margin=Inches(.7)
    s.page_width=Inches(8.5);s.page_height=Inches(11)
    d.styles['Normal'].font.size=Pt(9)
    d.add_heading('Nền tảng nghe nhạc trực tuyến',0)
    p(d,'Bản tiếng Việt rút gọn của báo cáo nhóm 04. Các phần dưới giải thích nội dung biểu mẫu, giữ mã yêu cầu, mã công việc, số liệu và điều kiện nghiệm thu để đối chiếu bản nộp tiếng Anh.')
    tb(d,['Thông tin','Nội dung'],[
    ['Môn học','Software Project Management; Class 01; nhóm 04'],
    ['Thời gian','17/08/2026–15/01/2027; 108 ngày làm việc; 5 tháng'],
    ['Sản phẩm','Flutter cho iOS/Android/Web; Node.js/Express; PostgreSQL; WebSocket; object storage và CDN'],
    ['Quy mô','5 vai trò dự án; 55 WP; 2,464 giờ; không ai làm hai việc cùng lúc'],
    ['Ngân sách','757,000,000 VND; cost baseline 720,921,916 VND'],
    ['Customer','Đỗ Thị Bích Ngọc: yêu cầu, UAT và nghiệm thu sản phẩm'],
    ['Sponsor','Nguyễn Trung B, PTIT Innovation Fund: tài trợ, duyệt thay đổi vượt quyền PM và đóng dự án'],
    ['PM','Trần Vũ Tiến Minh: lập kế hoạch, điều phối, báo cáo và xử lý thay đổi trong thẩm quyền']], [2000,7900])
    h(d,'Cách đọc thuật ngữ')
    tb(d,['Thuật ngữ','Nghĩa ngắn'],[
    ['Charter / Scope Statement','Văn bản khởi động / tuyên bố phạm vi'],['SRS / RTM','Đặc tả yêu cầu / ma trận truy vết yêu cầu'],
    ['WBS / WP / Dictionary','Cấu trúc phân rã công việc / gói công việc / mô tả chi tiết WBS'],
    ['PMP / QMP','Kế hoạch quản lý dự án / chất lượng'],['Baseline / reserve','Đường cơ sở được duyệt / quỹ dự phòng'],
    ['RC / staging / production','Bản chờ phát hành / môi trường thử / môi trường vận hành'],['UAT / hypercare','Customer kiểm thử nghiệm thu / hỗ trợ sát sau phát hành']], [2400,7500])
    p(d,'Bốn sinh viên: Minh B23DCCE067; Đại B23DCVT071; Hiếu B23DCCE034; Thu B23DCCE091. Trần Văn A là nhân sự giả định đảm nhận QA/DevOps trong mô hình dự án.')
    newpage(d,'1 Khởi động và phạm vi dự án')
    h(d,'Mục tiêu kinh doanh')
    bos=[
    'BO-1: Nghe nhạc trên iOS/Android/Web; ra mắt >= 9,000 người đồng thời; UAT >= 4.2/5.',
    'BO-2: Tìm và tổ chức nhạc bằng tìm kiếm, gợi ý, playlist, thư viện.',
    'BO-3: Tương tác qua Live Premiere, chat, yêu thích, đánh giá, bình luận, thông báo.',
    'BO-4: Admin quản lý bài hát, album, nghệ sĩ, người dùng trên Flutter Web.',
    'BO-5: Bảo vệ tài khoản/dữ liệu; backup hằng ngày, RPO 24 h, RTO 4 h.',
    'BO-6: Thiết kế mở rộng tới 50,000 người ở giai đoạn sau; bảo hành 1 năm.']
    for x in bos:p(d,x)
    h(d,'Phạm vi và sản phẩm bàn giao')
    p(d,'Xây một codebase Flutter cho ứng dụng iOS, Android, web client và admin; giao diện Việt/Anh. Backend Express, PostgreSQL và WebSocket phục vụ xác thực, phát nhạc qua CDN, Live Premiere, khám phá, thư viện, tương tác và xác minh hóa đơn store.')
    p(d,'Bàn giao: tài liệu yêu cầu/kế hoạch/phạm vi; thiết kế kiến trúc, dữ liệu, API và UI; mã nguồn kèm unit test; báo cáo kiểm thử; ứng dụng đã phát hành; hướng dẫn người dùng, admin, triển khai và vận hành. Có hỗ trợ kỹ thuật/bảo trì 1 năm sau đóng dự án.')
    p(d,'Ngoài phạm vi: hai codebase native riêng; app desktop native; tự xử lý thanh toán/thẻ/ví; thiết kế gói thuê bao; ngôn ngữ ngoài Việt/Anh; marketing; bản quyền âm nhạc; Customer tự chạy UAT; hỗ trợ sau năm bảo hành. Nhóm vẫn chuẩn bị staging và sửa lỗi UAT trong WBS.')
    h(d,'Điều kiện thành công và quyền quyết định')
    p(d,'100% yêu cầu Must/Should phải qua phép kiểm trong RTM và được Customer chấp nhận. Could-have làm theo năng lực, chỉ cắt sau thỏa thuận/thay đổi được duyệt. RC 1.5.8 không còn lỗi S1/S2; crash-free >= 99% trong hypercare. Mỗi milestone chậm không quá 14 ngày so với baseline; tổng chi không vượt 757,000,000 VND.')
    p(d,'PM chỉ tự duyệt khi đồng thời: chi phí thay đổi <= 2% ngân sách (15,140,000 VND) VÀ chậm <= 3 ngày lịch. Vượt một ngưỡng phải xin Sponsor; đổi yêu cầu cần Customer đồng ý. Chỉ Sponsor giải ngân management reserve. Nhóm cố định 5 người; tuyển thêm cần Sponsor duyệt. PM chốt kỹ thuật sau tham vấn Technical Lead và thành viên kỹ thuật.')
    p(d,'Đóng thành công khi Must/Should đạt, app lên store, tài liệu được chấp nhận và Sponsor ký 1.7.9. Thất bại khi trễ bàn giao quá 14 ngày, Sponsor ngừng tài trợ, thay đổi vượt quỹ dự phòng/giới hạn giao hàng, hoặc tổng chi vượt ngân sách.')
    p(d,'Rủi ro tổng quát: chậm phê duyệt (R-01); WebSocket không đạt tải (R-02); vượt giờ phát triển (R-03); đổi yêu cầu (R-04). Chi tiết ở phần 7.')

    newpage(d,'Sổ giả định')
    assumptions=[
    ('A-01','Customer/Sponsor dự họp lấy yêu cầu 17–27/08/2026.','17/08/2026','Gửi lịch/agenda; đã họp đúng lịch.','Đóng'),
    ('A-02','Đọc SRS/RTM/phạm vi 10–11/09; review 14–15/09.','09/09/2026','Gửi tài liệu và xác nhận; đã review.','Đóng'),
    ('A-03','Đọc bản sửa và ký SRS/RTM 18/09.','17/09/2026','Gửi bản sửa để ký; đã ký đúng lịch.','Đóng'),
    ('A-04','Sponsor đọc PMP 30/09–01/10; duyệt 01/10.','29/09/2026','Gửi kế hoạch, nhắc đúng hạn.','Đang theo dõi'),
    ('A-05','Đọc thiết kế 07–08/10; review 09–12/10; duyệt 21/10.','06/10/2026','Gửi kiến trúc/API/mockup; đặt lịch.','Đang theo dõi'),
    ('A-06','Customer dự demo MVP 12/11.','05/11/2026','Mời trước 1 tuần; quay demo dự phòng.','Đang theo dõi'),
    ('A-07','Customer/Sponsor dự demo đủ tính năng 30/11.','23/11/2026','Mời trước 1 tuần; quay demo dự phòng.','Đang theo dõi'),
    ('A-08','Đọc báo cáo test 10–11/12; review 14/12.','09/12/2026','Gửi báo cáo và lời mời.','Đang theo dõi'),
    ('A-09','Customer UAT từ 23/12; báo lỗi để sửa 1.6.2; ký 04/01/2027.','18/12/2026','Chốt thời gian/kênh báo lỗi; UAT ngoài scope.','Đang theo dõi'),
    ('A-10','Admin Customer dự đào tạo 12–13/01/2027.','05/01/2027','Chốt người dự và lịch trước 1 tuần.','Đang theo dõi'),
    ('A-11','Đọc bàn giao 04–05/01; review 06–07/01; nhận 15/01/2027.','31/12/2026','Gửi tài liệu; đặt lịch bàn giao.','Đang theo dõi'),
    ('A-12','Sponsor review báo cáo cuối 08/01; ký đóng 15/01/2027.','21/12/2026','Gửi báo cáo; đặt lịch review/ký.','Đang theo dõi'),
    ('A-13','Store xét duyệt trong 3 ngày làm việc 12–14/01/2027.','11/01/2027','Theo checklist; thử TestFlight/internal test.','Đang theo dõi'),
    ('A-14','Cloud/CDN kích hoạt trong 2 ngày làm việc 09–12/10.','08/10/2026','Xác nhận trước mua.','Đang theo dõi')]
    p(d,'Người phụ trách A-01–A-12: Minh; A-13: Thu; A-14: A. Trạng thái là tại mốc lập kế hoạch, không phải báo cáo thực tế ngày hôm nay.')
    tb(d,['ID','Giả định','Hạn xử lý','Hành động / trạng thái'],[(a,b,c,e+' '+f) for a,b,c,e,f in assumptions],[650,4350,1250,3650])

    newpage(d,'Danh sách yêu cầu chức năng')
    for i in range(14):
        src=DATA['tables'][31][i+3];x=FR[i]
        h(d,f'{src[0]} {x[0]} — {vi_priority(src[4])}')
        p(d,f'{x[1]} Nghiệm thu: {x[2]} Kiểm tra: {x[3]}.')
    newpage(d,'Danh sách yêu cầu phi chức năng')
    tb(d,['ID / ưu tiên','Yêu cầu','Nghiệm thu và cách kiểm'],[
    [DATA['tables'][31][i+17][0]+'\n'+vi_priority(DATA['tables'][31][i+17][4]),x[0]+': '+x[1],x[2]+' '+x[3]+'.'] for i,x in enumerate(NFR)],[1250,4400,4250])
    p(d,'Nguồn FR-01–FR-13 và NFR-12: Customer; FR-14: Sponsor. NFR-01–04, 06–07: Sponsor; NFR-05, 08–11: PM. Must = bắt buộc; Should = nên có nhưng vẫn thuộc cam kết nghiệm thu của bài; Could = có thể cắt qua change control.')

    newpage(d,'2 Truy vết yêu cầu và WBS')
    p(d,'RTM nối yêu cầu với nguồn, ưu tiên, mục tiêu kinh doanh, công việc và phép kiểm. Tiêu chí nghiệm thu/cách kiểm ở phần 1 cũng là hai cột Verification/Validation của bản nộp.')
    rr=[r for n in [41,42] for r in DATA['tables'][n] if re.fullmatch(r'N?FR-\d\d',r[0])]
    tb(d,['ID','Nguồn / ưu tiên','Mục tiêu','Công việc bàn giao'],[[r[0],actor(r[2])+'; '+vi_priority(r[3]),', '.join(dict.fromkeys(re.findall(r'BO-\d',r[5]))),', '.join(dict.fromkeys(re.findall(r'1\.\d(?:\.\d+)?(?:-1\.\d(?:\.\d+)?)?',r[6])))] for r in rr],[1100,2600,1300,4900],7.5)
    h(d,'Cách đọc WBS Dictionary')
    p(d,'WBS có 7 nhóm giai đoạn 1.1–1.7 và 55 work package 1.x.y. Mỗi WP có một người thực hiện, 8–80 h. Hai WP 1.7.4/1.7.5 chia thành các phiên theo lịch. Unit test nằm trong WP phát triển; QA làm kiểm thử tích hợp, tải, bảo mật, thiết bị.')
    p(d,'Sách Book of Forms mục 2.10, trang in 52–53, cho phép Dictionary chi tiết từng WP hoặc tóm tắt ở cấp control account. Bài dùng 1.1–1.7 làm control account tổng hợp phạm vi, lịch và chi phí; bảng con liệt kê các WP. Vì vậy nhãn đã đổi thành Control Account Name. Nếu giảng viên yêu cầu nguyên mẫu một WP thì phải tách form theo từng 1.x.y.')
    p(d,'Milestone là sự kiện thời lượng 0, không phải WP có giờ công. Cây WBS bỏ các dòng milestone; lịch, milestone list và trường Milestones vẫn giữ chúng. Không xóa ID/mã lịch vì các liên kết và bảng khác đang tham chiếu; mã có khoảng trống là bình thường.')

    newpage(d,'Truy vết mục tiêu sang yêu cầu kỹ thuật')
    trvi=[
    'Một codebase Flutter: iOS 15.0+, Android 8.0+, Chrome/Edge 96+, Firefox 99+, Safari 15.6+.',
    'Audio/video lưu object storage và phân phối CDN; PostgreSQL chỉ lưu metadata.',
    '>= 9,000 người; kiểm thử 15,000 với lỗi < 1%, p90 < 200 ms.',
    'Cold < 2 s, warm < 1 s; pin < 6%/giờ.',
    'Theo Apple HIG / Material 3; UAT >= 4.2/5.',
    'Chuỗi Việt/Anh ở file ARB; lưu lựa chọn ngôn ngữ trên thiết bị.',
    'API search/playlist/library dưới /api/v1; gợi ý từ lịch sử, thể loại, người nghe tương tự.',
    'WebSocket cho Live Premiere; thử tải kết nối đồng thời.',
    'API yêu thích, đánh giá, bình luận, lịch sử nghe; lưu PostgreSQL.',
    'Push cho bản phát hành, gợi ý, trả lời bình luận.',
    'Admin Flutter Web; API CRUD có phân quyền bài hát/album/nghệ sĩ/người dùng.',
    'bcrypt >= 10; JWT access <= 1 h; TLS >= 1.2.',
    'Backup PostgreSQL hằng ngày; RPO 24 h, RTO 4 h.',
    'StoreKit/Google Play Billing xử lý mua; backend xác minh hóa đơn, không lưu thẻ.',
    'Backend container, thiết kế mở rộng ngang 50,000 người ở giai đoạn sau.',
    'OpenAPI 3.0 và /api/v1.',
    'Coverage backend tự động >= 70%.',
    'Log tập trung, giám sát uptime, cảnh báo lỗi/độ trễ.']
    src=[r for r in DATA['tables'][43] if len(r)>4 and re.fullmatch(r'TR-\d\d',r[4])]
    current='';tr=[]
    for row,txt in zip(src,trvi):
        if row[0]:current=row[0]
        tr.append([row[4],current,txt,vi_priority(row[6])])
    tb(d,['ID','Mục tiêu','Yêu cầu kỹ thuật','Ưu tiên'],tr,[750,950,6600,1600])

    desc=[
    ('Lấy yêu cầu; viết Charter, SRS, RTM và đường cơ sở phạm vi; review và sửa trước ký. Stakeholder Register được lập/cập nhật trong công việc lập kế hoạch.',
    'Customer/Sponsor tham dự A-01–A-03; đọc tài liệu trước review. PMP nằm ở 1.7.1. Nghỉ 01–02/09 và làm bù 05/09.',
    'Mỗi yêu cầu có nguồn, ưu tiên và phép kiểm đo được. Customer ký SRS/RTM; Sponsor duyệt Charter; 100% Must/Should có tiêu chí kiểm tra.',
    'SRS, RTM, Scope Statement, Charter; Book of Forms; điều khoản tài trợ PTIT Innovation Fund.'),
    ('Thiết kế kiến trúc, dữ liệu, module và API; mockup/prototype; thiết lập CI/CD, staging, CDN và backup.',
    'Thiết kế chi tiết/API sau ký SRS; màn hình sau duyệt thiết kế 21/10. Cloud mua 1.7.3, kích hoạt A-14; tính tiền 3 tháng ở 1.2.7.',
    'Review kiến trúc/API/ER trước coding; mockup được Customer duyệt. Schema >= 3NF, đủ 14 chức năng; API p90 < 200 ms ở tải mục tiêu; CI chạy mỗi commit.',
    'Node.js, PostgreSQL DDL, WebSocket, Docker, OpenAPI 3.0, Flutter, file thiết kế; hợp đồng cloud/CDN.'),
    ('Làm API tài khoản/phát nhạc, WebSocket và màn hình MVP; developer viết unit test. QA tích hợp, chạy thử sớm 1,000 kết nối, ZAP và đo mở app trong 1.3.5.',
    'API sau 1.2.6; CDN trước tích hợp streaming; UI sau 1.2.11. Trước CI hoạt động 20/10, chạy lint/test/coverage tại máy rồi chạy lại CI. 1.3.7 sau 1.4.1 để tránh trùng người.',
    'Coverage backend >= 70%; cold < 2 s, warm < 1 s, pin < 6%/giờ theo điều kiện test. FR-04/FR-09 qua tích hợp trên OS/browser hỗ trợ; xử lý phản hồi demo.',
    'Express, PostgreSQL, WebSocket, CDN, Flutter, Jest/flutter_test; hợp đồng CDN/hosting.'),
    ('Làm API khám phá/thư viện/xã hội/admin/hóa đơn; màn hình tài khoản, ngôn ngữ, tải offline, video, mua tính năng, admin; test và demo.',
    '1.4.2 chạy song song MVP từ tháng 3; code xong 18/11, hoàn thiện 07/12. Dashboard dùng chung API. Cắt Could-have trước khi cần, qua change control.',
    'Unit test từng chức năng; không hồi quy FR-04/09; CRUD admin có integration test; UI theo mockup Việt/Anh. Đủ 100% Must/Should; FR-12 đạt.',
    'FR-01–14, Flutter/Web, /api/v1, StoreKit/Play Billing; thỏa thuận developer của hai store.'),
    ('QA lập kế hoạch, test case theo RTM; kiểm thử tải/bảo mật/thiết bị; viết báo cáo; developer sửa lỗi; review với Customer.',
    'Unit test do developer làm trong WP phát triển; tích hợp ở 1.3.5/1.4.5. Customer tự UAT ngoài scope. Mua 4 máy và 1 tháng k6: 45,600,000 VND.',
    '15,000 VU: lỗi < 1%, p90 < 200 ms. RC 1.5.8 không còn S1/S2. Báo cáo đủ yêu cầu RTM; retest xác nhận sửa lỗi và không hồi quy; Customer review.',
    'Grafana Cloud k6, OWASP ASVS, framework API/end-to-end; 2 iOS, 2 Android; điều khoản công cụ test.'),
    ('Đưa RC lên staging, sửa lỗi UAT; sau Customer ký mới lên production; nộp store, đào tạo, hypercare và bàn giao tài liệu.',
    'UAT Customer ngoài scope; ký 04/01 trước go-live. Store review <= 3 ngày làm việc. Bảo hành ngoài WBS: 480 h, Hiếu/A mỗi người 240 h; 71,822,880 VND.',
    'Có log/uptime trước go-live; restore <= 4 h. App lên hai store; crash-free >= 99% trong hypercare; smoke test backend đạt; Customer nhận tài liệu, Sponsor ký đóng.',
    'Docker, CDN, hướng dẫn store; tài liệu triển khai/vận hành Việt/Anh; thỏa thuận Apple/Google/cloud.'),
    ('Lập PMP gồm lịch, chi phí, chất lượng, nguồn lực, giao tiếp, mua sắm, stakeholder và R-01–R-14; mua sắm; báo cáo; theo dõi rủi ro/thay đổi; đóng dự án.',
    'Sponsor duyệt PMP 01/10. Chi phí hàng mua đặt ở WP sử dụng. 1.7.4/1.7.5 là phiên lịch riêng; quyền PM đồng thời <= 2% ngân sách và <= 3 ngày lịch.',
    'Báo cáo có sai lệch lịch/chi phí; review rủi ro ít nhất 2 tuần/lần, score >= 15 đưa vào báo cáo tháng. PMP, 5 báo cáo tháng và báo cáo cuối được Sponsor chấp nhận.',
    'Charter, Book of Forms, Jira, MS Project; hợp đồng cloud/CDN, k6, tài khoản store.')]
    # Dictionary tables retain every WP's resource, hours, actual direct cost and scheduled window.
    res={r[0]:r for r in DATA['tables'][67] if re.fullmatch(r'1\.\d\.\d+',r[0])}
    for k in range(1,8):
        newpage(d,f'WBS Dictionary 1.{k} {PHASE[k-1]}')
        x=desc[k-1];p(d,'Nội dung: '+x[0]);p(d,'Giả định/ràng buộc: '+x[1])
        sr=DATA['tables'][44+k];out=[]
        for r in sr:
            if not re.fullmatch(r'1\.\d\.\d+',r[0]):continue
            if r[2]=='-':continue
            w=r[0];rr=res[w];dates=rr[3].split(', one task')[0].removeprefix('Available ')
            if w in ['1.7.4','1.7.5']:dates='Các phiên ở phần 3'
            direct=sum(num(xr[9]) for xr in sr if xr[0]==w)
            out.append([w,WPVI[w],NAMES[r[2]],r[3],fmt(direct),vi_dates(dates)])
        tb(d,['WP','Công việc','Người','Giờ','Chi phí trực tiếp VND','Lịch làm việc'],out,[850,3300,600,500,1600,3050],8)
        totals=next(r for r in sr if len(r)>1 and r[1]=='Total')
        extra={5:' 1.5.3 gồm 28,000,000 VND thiết bị + 17,600,000 VND k6.',6:' 1.6.6 gồm 2,600,000 VND Apple + 700,000 VND Google.'}.get(k,'')
        p(d,f'Tổng nhóm: {totals[3]} h; nhân công {totals[5]} VND; vật tư/dịch vụ {totals[8]} VND; trực tiếp {totals[9]} VND, chưa cộng contingency.'+extra)
        p(d,'Chất lượng/nghiệm thu: '+x[2]);p(d,'Kỹ thuật và thỏa thuận: '+x[3])
        ms=next(r for r in sr if r[0].startswith('Milestones:'))
        dates=[re.sub(r'^\d+\. ','',z) for z in ms[1].replace('Due Dates:\n','').splitlines()]
        p(d,'Mốc: '+'; '.join(name+' '+vi_dates(date) for name,date in zip(MILESTONE_VI[k-1],dates))+'.')

    newpage(d,'3 Lịch dự án')
    p(d,'Lịch gốc là 4.xml; Gantt đã xuất là schedule.pdf. 108 ngày làm việc, 8 h/ngày (08–12 h, 13–17 h), thứ Hai–Sáu. Nghỉ 01–02/09/2026 và 01/01/2027; làm bù thứ Bảy 05/09/2026. Mỗi người làm một việc tại một thời điểm.')
    milestone_names=['Customer ký SRS/RTM và phạm vi','Sponsor duyệt PMP/lịch','Customer duyệt thiết kế','Demo MVP','Hoàn thiện MVP nội bộ','Demo đủ chức năng','Hoàn thiện đủ chức năng','Review báo cáo test','RC sẵn sàng UAT','Customer ký nghiệm thu sau UAT','Phát hành app và bàn giao','Sponsor ký đóng dự án']
    mids=['1.1.8','1.7.2','1.2.11','1.3.6','1.3.8','1.4.6','1.4.8','1.5.6','1.5.8','1.6.3','1.6.12','1.7.9']
    mdates=['18/09/2026','01/10/2026','21/10/2026','12/11/2026','20/11/2026','30/11/2026','07/12/2026','14/12/2026','18/12/2026','04/01/2027','15/01/2027','15/01/2027']
    tb(d,['Mã','Mốc quản lý','Ngày baseline'],list(zip(mids,milestone_names,mdates)),[1000,6900,2000])
    p(d,'1.3.6, 1.4.6, 1.5.6 là công việc demo/review có giờ công; ngày kết thúc cũng được dùng làm mốc quản lý trong Charter. Các dòng milestone thực sự trong MS Project có duration/work = 0.')
    h(d,'Công việc định kỳ của PM')
    p(d,'1.7.4: 5 báo cáo tháng, tổng 72 h: 11/09; 15–16/10; 13 và 16/11; 10–11/12/2026; 14–15/01/2027.')
    p(d,'1.7.5: 8 phiên review rủi ro/thay đổi, mỗi phiên 8 h: 14/10, 28/10, 11/11, 25/11, 09/12, 23/12/2026, 05/01 và 13/01/2027.')
    p(d,'Tháng dự án: M1 17/08–16/09; M2 17/09–16/10; M3 17/10–16/11; M4 17/11–16/12; M5 17/12–15/01. Không dùng tháng dương lịch để tính bảng tải nguồn lực.')

    newpage(d,'4 Ước lượng chi phí')
    h(d,'Ngân sách và dự phòng')
    budget_names=['Nhân công trong WBS (2,464 h)','Cloud/database/storage/CDN (3 tháng)','Thiết bị, k6 và phí store','Contingency cho rủi ro đã nhận diện','Tổng ước lượng WP','Bảo hành 12 tháng (480 h), ngoài WBS','Cost baseline','Management reserve 5%, Sponsor quản lý','Ngân sách làm tròn']
    budget_vals=['406,794,048','55,500,000','48,900,000','137,904,988','649,099,036','71,822,880','720,921,916','36,046,096','757,000,000']
    tb(d,['Khoản','VND'],list(zip(budget_names,budget_vals)),[7300,2600])
    p(d,'Baseline + management reserve = 756,968,012 VND, làm tròn lên 757,000,000. Khoản làm tròn 31,988 VND không phải quỹ dự phòng mới. Warranty: Hiếu 240 h x 175,001 + A 240 h x 124,261 = 71,822,880 VND.')
    p(d,'Lương giờ = lương trung vị ITviec cho 1–2 năm kinh nghiệm x 1.215 / 176 h, làm tròn. Hệ số 1.215 gồm bảo hiểm phía doanh nghiệp trong giả định dự toán. Các thành viên được giả định có 2 năm kinh nghiệm; chỉ tính giờ được phân cho dự án.')
    rs=DATA['tables'][52][3:8]
    rates={MINH:206067,DAI:118566,HIEU:175001,A:124261,THU:198818}
    tb(d,['Người','Đơn giá VND/h','Giờ','Nhân công VND'],[[r[0],fmt(rates[r[0]]),r[7],r[9]] for r in rs],[3300,2200,1400,3000])
    p(d,'Không tính riêng indirect cost/travel: văn phòng, laptop, internet là overhead doanh nghiệp; bảo hiểm đã ở rate. Không cộng lại các khoản này.')
    h(d,'Công thức dùng trong các bảng')
    p(d,'Nhân công = giờ x rate; M = nhân công + chi phí vật tư/dịch vụ trực tiếp; contingency = P − M; ước lượng WP = M + contingency. Tổng trực tiếp M = 511,194,048 VND. O = 0.9M; P = M x (1 + tỷ lệ vượt chi theo nhóm). PERT = (O + 4M + P) / 6.')
    p(d,'Bài dùng P để dự phòng trong baseline, không dùng kết quả PERT làm baseline. PERT là một phép ước lượng so sánh. Tổng PERT 525,658,311 VND khác tổng P 649,099,036 VND vì cách chọn mức thận trọng khác nhau.')

    newpage(d,'Chi phí theo từng work package')
    ace=[r for r in DATA['tables'][56] if re.fullmatch(r'1\.\d\.\d+',r[0])]
    tb(d,['WP','Người / giờ','Nhân công','Vật tư/dịch vụ','Contingency','Ước lượng'],[[r[0],r[1],r[2],r[3],r[4],r[5]] for r in ace],[950,1250,1900,1900,1950,1950],7.5)
    p(d,'Đơn vị VND. Không tính chi phí cho milestone 0 h. Activity Cost Estimates gốc còn nêu phương pháp, khoảng ước lượng, độ tin cậy và cơ sở; các bảng dưới giải thích cách dùng.')
    newpage(d,'Phương pháp và cơ sở dự toán')
    h(d,'Ước lượng ba điểm theo nhóm')
    three=[r for r in DATA['tables'][57] if re.fullmatch(r'1\.\d',r[0]) and len(r)==6 and r[4]=='(O + 4M + P) / 6']
    tb(d,['Nhóm','O','M','P','PERT'],[[r[0],r[1],r[2],r[3],r[5]] for r in three],[650,2300,2300,2300,2350],7.5)
    p(d,'Tỷ lệ P vượt M: 1.1/1.7 = 15%; 1.2 = 20%; 1.3/1.4 = 40%; 1.5 = 35%; 1.6 = 30%. Contingency từng nhóm: 8,429,891; 26,802,470; 24,184,535; 24,999,699; 27,811,401; 15,785,776; 9,891,216 VND.')
    h(d,'Parametric và analogous')
    p(d,'Parametric: đơn giá x số lượng, dùng cho giờ công, cloud, máy test, k6, tài khoản store và bảo hành. Analogous lấy công việc tương tự rồi nhân tỷ lệ quy mô, kèm tỷ lệ rate nếu đổi người.')
    analog=[['1.4.5 từ 1.3.5','4/4 gói tích hợp; cùng 40 h','4,970,440'],['1.4.6 từ 1.3.6','1/1 buổi demo','1,648,536'],['1.4.7 từ 1.3.7','1 chu kỳ; đổi Hiếu sang Thu','7,952,720'],['1.5.4 từ 1.3.7','1 chu kỳ; đổi Hiếu sang Đại','4,742,640'],['1.6.10 từ 1.2.8','1 gói review trong 2 ngày','3,297,072'],['1.6.11 từ 1.1.7','3/2 tài liệu sửa','2,845,584'],['1.7.7 từ 1.5.6','1 báo cáo review trong 1 ngày','1,648,536']]
    tb(d,['So sánh','Cơ sở','Ước lượng VND'],analog,[2600,4900,2400])
    h(d,'Bottom-up và phân loại khoản mua')
    p(d,'Cộng từ 55 WP lên 7 nhóm rồi toàn dự án. Supplies = k6 + phí store = 20,900,000; Equipment = 4 máy test = 28,000,000; Other Direct Costs = hóa đơn cloud PR-01–04 = 55,500,000. Chúng cộng thành 104,400,000 VND chi phí phi nhân công. Không cộng lặp chi phí gộp trong hóa đơn cloud.')
    h(d,'Đơn giá tham chiếu')
    p(d,'Cloud mỗi tháng: 2 máy API 6,621,788; WebSocket 1,975,643; staging 945,100; PostgreSQL 3,060,000; storage 1,000,000; CDN 2,500,000 VND. Trước VAT 16,102,531; VAT 1,610,253; sau VAT 17,712,784; dự toán 18,500,000/tháng gồm biên cho backup, log và ngày phát sinh.')
    p(d,'4 máy: iPhone 13 12,890,000; iPhone 11 cũ 5,500,000; Galaxy A17 5,090,000; Galaxy A07 3,690,000; tổng 27,170,000, dự toán 28,000,000 VND. k6: 4,855 VUh, trừ 500 miễn phí, 4,355 x 0.15 USD + 19 USD = 672.25 USD; tỷ giá dự toán 26,170, dự toán 17,600,000 VND. Apple 99 USD/năm, Google 25 USD/lần; làm tròn lần lượt 2,600,000 và 700,000 VND.')
    p(d,'Đây là cơ sở giá đã chốt trong bài, không phải báo giá hiện tại. Nguồn/URL tham chiếu đầy đủ giữ trong Basis of Estimates and References của bản nộp tiếng Anh.')

    newpage(d,'5 Quản lý chất lượng')
    p(d,'Chuẩn nền: ISO/IEC 25010; Apple HIG/Material Design 3; OWASP ASVS/MASVS. Mục tiêu được kiểm bằng test/review do nhóm thực hiện; UAT là phần Customer tự làm.')
    quality=[
    ['1','Ra mắt >= 9,000 người; test 15,000, lỗi < 1%','k6 trong 1.5.3; đồng thời p90 < 200 ms'],['2','API p90 < 200 ms','k6 ở 1.5.3; dashboard độ trễ API hằng ngày trong hypercare 1.6.5'],
    ['3','Cold < 2 s; warm < 1 s','Trung vị 10 lần mở trên từng máy'],['4','Pin < 6%/giờ','1 h Wi-Fi, tắt màn hình, âm lượng 50%; quy đổi 4,000 mAh'],
    ['5','bcrypt >= 10; JWT <= 1 h; TLS >= 1.2; không còn lỗ hổng critical/high','ZAP và ASVS/MASVS ở 1.5.3; review từng PR'],
    ['6','Coverage backend >= 70%','Jest trên CI; chặn merge dưới ngưỡng'],['7','100% Must/Should có test đã đạt','Đối chiếu RTM trong báo cáo 1.5.5'],
    ['8','Không còn S1/S2 tại RC 1.5.8','QA kiểm defect log ngày 18/12'],['9','Crash-free >= 99%','Crashlytics trong hypercare'],
    ['10','Đúng guideline, Việt/Anh; UAT >= 4.2/5','Design review 1.2.8; đối chiếu mockup 1.3.5/1.4.5; Customer khảo sát trước 1.6.3'],
    ['11','Backup ngày; RPO 24 h, RTO 4 h','Diễn tập restore trước go-live 1.6.4'],['12','Đóng mọi nhận xét must fix trước gửi tài liệu','Peer-review checklist cho mỗi tài liệu']]
    tb(d,['Mục','Mục tiêu','Phép kiểm'],quality,[500,4800,4600])
    h(d,'Vai trò')
    p(d,'Minh sở hữu QMP, audit và báo cáo. A lập/chạy test, giữ defect log và CI gate. Hiếu đặt chuẩn code, review backend/kiến trúc và phân tích nguyên nhân. Đại duy trì yêu cầu/RTM và review hướng dẫn. Thu đối chiếu UI với mockup/guideline. Developer tự viết unit test; Customer review/UAT/nghiệm thu; Sponsor duyệt kế hoạch và thay đổi vượt quyền PM.')
    h(d,'Phòng ngừa và kiểm soát')
    p(d,'Definition of Done cho code: test pass, coverage backend >= 70%, lint không lỗi, một developer khác duyệt PR, endpoint khớp OpenAPI, cập nhật Jira. Với tài liệu: đúng form và đóng mọi must fix. Checklist yêu cầu, thiết kế, code, tài liệu, phát hành được soạn ở 1.5.1 và hoàn thiện trong PMP.')
    p(d,'CI chạy mỗi commit từ 20/10; trước đó developer chạy cùng kiểm tra ở máy. QA tích hợp sau từng giai đoạn; Must-have phải đạt trước demo. Kiểm thử phi chức năng ở 1.5.3, retest và regression sau mỗi lần sửa. Không đóng lỗi chỉ dựa vào lời báo của developer.')
    p(d,'Audit trước các review: 09/10, 12/11, 30/11, 14/12/2026 và 06/01/2027, PM + QA khoảng 2 h; phát hiện có người xử lý/hạn, đóng trong 5 ngày làm việc. Báo cáo tháng nêu coverage, pass rate, lỗi theo mức, lỗi mỗi giai đoạn và audit còn mở. Cuối mỗi giai đoạn họp bài học 30 phút; Pareto chỉ nhóm nguyên nhân gây nhiều rework.')
    h(d,'Xử lý lỗi và cải tiến')
    p(d,'S1 (crash/mất dữ liệu/lỗ hổng/chặn Must): sửa trong 1 ngày làm việc. S2: 3 ngày. S3: trước RC. S4: sửa trong WP còn giờ, nếu không đưa vào warranty backlog. Quy trình: ghi Jira → QA/Tech Lead phân loại → developer sửa kèm test → QA retest/regression → đóng.')
    p(d,'WP sửa lỗi: 1.3.7, 1.4.7, 1.5.4, 1.5.7, 1.6.2; vượt ước lượng dùng contingency được duyệt. S1, trượt chỉ số, >= 3 lỗi cùng nguyên nhân hoặc audit finding kích hoạt phân tích 5 Whys/fishbone trong 2 ngày làm việc. Có người/hạn xử lý; đổi scope/lịch/chi phí phải qua 1.7.5. PDCA cập nhật checklist và kiểm lại ở audit sau.')

    newpage(d,'6 Quản lý nguồn lực')
    h(d,'Tải công việc theo tháng dự án')
    tb(d,['Người','M1','M2','M3','M4','M5','Tổng giờ','Số WP'],[[NAMES[r[0]]]+r[2:9] for r in rs],[2000,1000,1000,1000,1000,1000,1500,1400])
    p(d,'Tổng nhóm: 496 / 640 / 592 / 352 / 384 h, cộng 2,464 h trên 4,320 h khả dụng. Mỗi người <= 176 h/tháng. 1,856 h chưa phân bổ là công suất tổng; không thể tự chuyển thành giờ của người đang nghẽn. Đổi người phải xét kỹ năng, lịch cá nhân, dependency và chi phí.')
    h(d,'Team Charter')
    p(d,'Giá trị: tôn trọng; quyết định theo dữ liệu; một người chịu trách nhiệm mỗi WP; báo vấn đề trong 1 ngày làm việc; chỉ coi xong khi đạt Definition of Done.')
    p(d,'Họp: stand-up 09:00, 15 phút; thứ Hai họp nhóm 60 phút. Gửi agenda trước 24 h; ghi quyết định, việc, người và hạn trong 24 h sau họp. PM và chủ WP chuẩn bị trước buổi Customer/Sponsor 1 ngày.')
    p(d,'Kênh: Jira cho task/lỗi; GitHub cho code/review; Drive cho tài liệu; chat cho câu hỏi; gọi ngay khi S1/sự cố production. Trả lời trong 4 giờ làm việc. Tài liệu chính thức tiếng Anh, nội bộ có thể tiếng Việt; PM là đầu mối ngoài nhóm; cập nhật Jira hằng ngày.')
    p(d,'Quyết định: tìm đồng thuận; nếu chưa thống nhất thì bỏ phiếu, PM phá hòa. Tech Lead nêu lựa chọn và đánh đổi, PM chốt sau tham vấn. Ghi decision log; chỉ xét lại khi có dữ liệu mới, thay đổi được duyệt hoặc đe dọa baseline.')
    p(d,'Xung đột: trao đổi trực tiếp trong 2 ngày làm việc → PM hỗ trợ giải quyết → PM chốt và ghi lại. Vượt quyền thì Sponsor giải quyết. Vấn đề không chặn việc được chờ họp tuần nhưng không quá 1 tuần.')
    p(d,'Làm 08–17 h, giờ chung 09–16 h. Xin nghỉ trước 3 ngày làm việc, báo bệnh trước 09:00. Log giờ theo WBS mỗi ngày. Review PR trong 1 ngày; mỗi WP có người dự phòng đủ kỹ năng. Không đưa dữ liệu cá nhân thật lên thiết bị cá nhân/kênh công khai.')
    h(d,'Yêu cầu kỹ năng và Resource Breakdown Structure')
    p(d,'Minh: điều phối, lập kế hoạch, báo cáo. Đại: phân tích yêu cầu, Node.js, Flutter. Hiếu: Node.js, PostgreSQL, WebSocket. A: test design, k6, ZAP, Docker/CI. Thu: Figma, Apple HIG/Material 3, Flutter. Lịch/giờ và người của từng WP đã nêu ở Dictionary; đây cũng là dữ liệu của Resource Requirements.')
    p(d,'RBS phân theo người; thiết bị thuê/mua; vật tư; supplies; địa điểm. Thiết bị cloud thuê tuy là tài nguyên equipment trong RBS nhưng hóa đơn dịch vụ được gộp Other Direct Costs trong bảng bottom-up; không mâu thuẫn vì một bảng phân loại nguồn lực, bảng kia phân loại chi phí.')
    newpage(d,'Nguồn lực vật chất')
    pr=[
    ['PR-01','4 cloud server x 3 tháng','2 API 12 vCPU/8 GB; 1 WebSocket 8 vCPU/6 GB; 1 staging 6 vCPU/4 GB; mua 1.7.3, dùng từ 1.2.7.'],
    ['PR-02','1 PostgreSQL x 3 tháng','4 vCPU/8 GB/200 GB; có backup ngày.'],['PR-03','Storage 1,000 GB/tháng x 3','VNG Cloud; danh mục audio/video trong giới hạn.'],
    ['PR-04','CDN 5,000 GB/tháng x 3','Khoảng 87,000 giờ nghe 128 kbps; sau launch tương đương khoảng 450 người nghe trung bình liên tục. 9,000–15,000 là tải đỉnh server, không phải mức nghe liên tục.'],
    ['PR-05','4 thiết bị: 2 iOS + 2 Android','iPhone 13, iPhone 11 cũ, Galaxy A17, Galaxy A07; sẵn từ 1.3.5.'],['PR-06','k6 Cloud 1 tháng','4,855 VUh, trả tiền 4,355; test 12–20/11; ramp 5 phút + tải đủ 15 phút.'],
    ['PR-07','Apple Developer 1 tài khoản/năm','Đăng ký/xác minh trước 1.6.6, dùng TestFlight/App Store.'],['PR-08','Google Play 1 tài khoản/lần','Xác minh danh tính trước 1.6.6.'],
    ['PR-09','5 laptop, ít nhất 1 macOS','Doanh nghiệp cấp; macOS/Xcode để build iOS; không tính vào dự án.'],['PR-10','GitHub/Jira/Figma/Firebase','Gói miễn phí cho 5 người theo giả định; không tính license.'],
    ['PR-11','Văn phòng 5 chỗ','17/08/2026–15/01/2027; doanh nghiệp cấp; cho làm từ xa.'],['PR-12','Phòng họp cho 9 hoạt động','Đặt trước 1 tuần, có màn hình; họp online dự phòng.']]
    tb(d,['ID','Số lượng','Điều kiện và cách dùng'],pr,[900,2800,6200])
    p(d,'PR-04: vượt 5 TB dùng contingency 1.2.7 là 12,292,906 VND, khoảng 22 TB ở 550 VND/GB gồm VAT. Chi phí đầy đủ đã nằm trong phần 4. Bảo hành ngoài WBS do Hiếu và A thực hiện tổng 480 h sau closeout.')

    risk_vi(d)
    stake_vi(d)
    footer=d.sections[0].footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    font_run(footer.add_run('Nhóm 04 • Bản tiếng Việt rút gọn • Trang '),8)
    f=etree.SubElement(footer._p,q('fldSimple'));f.set(q('instr'),'PAGE')
    d.save(ROOT/'4_Tieng_Viet.docx')
    print('Saved Vietnamese reading edition; 55 WPs, 26 requirements, 18 TRs, 14 risks.')

def risk_vi(d):
    newpage(d,'7 Sổ rủi ro')
    p(d,'Điểm = xác suất x tác động lớn nhất trong 4 mặt: phạm vi, chất lượng, lịch, chi phí. Điểm sau hành động là mục tiêu rủi ro còn lại, chưa phải kết quả đo. Cả 14 rủi ro ở trạng thái Open trong kế hoạch.')
    scale=[['1','<= 10%','Không đổi yêu cầu','Không quá S4','< 1 ngày','< 1%'],['2','11–30%','Đổi/cắt Could-have','S3','1–3 ngày','1–2%'],['3','31–50%','Đổi Should-have','S2 hoặc trượt chỉ số <= 10%','4–7 ngày','> 2–5%'],['4','51–70%','Đổi Must có Customer đồng ý','S1 hoặc trượt > 10%','8–14 ngày','> 5–10%'],['5','> 70%','Không giao Must','Chặn RC/go-live','> 14 ngày','> 10%']]
    tb(d,['Mức','Xác suất','Phạm vi','Chất lượng','Trễ lịch','Vượt chi'],scale,[650,1300,2500,2600,1450,1400],7.5)
    p(d,'Chi phí tính theo ngân sách 757,000,000: 1% = 7,570,000; 2% = 15,140,000; 5% = 37,850,000; 10% = 75,700,000 VND. PM phải đồng thời trong ngưỡng 2% và 3 ngày lịch; vượt một ngưỡng trình Sponsor. Score 15–25: xử lý trước phiên review tới, báo cáo tháng; 6–14: kiểm mỗi phiên 1.7.5; 1–5: chấp nhận/theo dõi tháng. Chỉ đóng khi không còn khả năng xảy ra hoặc cửa sổ rủi ro đã qua.')
    risks=[
    ('Chậm phản hồi/phê duyệt','Nếu Customer/Sponsor vắng hoặc ký muộn, WP sau và mốc 1.7.2/1.2.11/1.6.3/1.6.12 có thể trễ.','Đặt lịch trước 1 tuần, gửi tài liệu đúng A-04–A-12, chốt tham dự trước 2 ngày làm việc. Ký trễ 1 ngày thì gọi ngay; > 3 ngày làm change request.','Assumption Log, lịch/milestone list và thông tin liên hệ/khả dụng của stakeholder.'),
    ('WebSocket không đạt tải','Nếu 1.3.3 không mở rộng theo tải, test 1.5.3 có thể không đạt 15,000 người, lỗi < 1%, p90 < 200 ms; phải sửa kiến trúc.','Socket.IO + PostgreSQL adapter; test sớm 1,000 kết nối ở 1.3.5. Nếu test đủ tải trượt, đề xuất thêm server và chạy lại, xin Sponsor duyệt chi từ contingency 1.5. Trigger: lỗi >= 1% hoặc p90 >= 200 ms.','Thiết kế/API; Dictionary 1.3/1.5; PR-01; kế hoạch/kết quả test; lịch và dự báo chi phí.'),
    ('Vượt giờ phát triển','Nếu 1.3/1.4 vượt ước lượng tới 40%, mốc 1.4.8 có thể trễ, giảm thời gian test.','Làm Must trước; họp tuần đối chiếu giờ đã làm/còn lại. Trigger: giờ còn lại > ngày còn lại x 8 h. Chậm > 3 ngày đưa vào review 1.7.5; chỉ cắt FR-05/10/13 khi Customer đồng ý và thay đổi được duyệt.','WBS/Dictionary; SRS/RTM/Scope nếu cắt tính năng; lịch, nguồn lực, dự báo chi phí.'),
    ('Đổi yêu cầu sau ký','Nếu Customer thêm/đổi yêu cầu sau 18/09, có thể phát sinh thiết kế/code, làm tăng lịch và chi phí.','Log Jira; phân tích tác động trong 2 ngày làm việc; quyết định ở 1.7.5. PM chỉ duyệt đồng thời <= 2% và <= 3 ngày. Sau 07/12, đề nghị ngoài Must chuyển backlog bảo hành theo thỏa thuận, không tự cam kết làm miễn phí.','SRS, Requirements Documentation, RTM, Scope; WBS/Dictionary, lịch và cost baseline nếu được duyệt.'),
    ('Store duyệt chậm/từ chối','Nếu store xét quá 3 ngày làm việc hoặc từ chối, có thể không lên store trước 15/01.','Checklist privacy, xóa tài khoản, billing; thử TestFlight/internal. Nộp 08/01; ba ngày làm việc tính từ ngày kế tiếp kết thúc 13/01. Không có quyết định 13/01 thì escalates; sửa/nộp lại trong 1 ngày, đề nghị Apple ưu tiên nếu đủ điều kiện.','A-13, checklist nộp store, lịch phát hành/bàn giao.'),
    ('Cloud kích hoạt muộn','Nếu quá 2 ngày làm việc sau mua, 1.2.7 và tích hợp CDN/MVP có thể trễ.','Chốt thời gian kích hoạt trước mua 1.7.3. Chưa active 12/10 thì làm repo/CI/Docker trước, staging/CDN sau; báo tại review 14/10.','A-14, hồ sơ mua/kích hoạt dịch vụ, Resource Requirements và lịch.'),
    ('Vượt dung lượng CDN','Nếu lưu lượng vượt 5,000 GB/tháng sẽ tăng chi phí 1.2.7.','Chấp nhận có dự phòng; alert 4,000 GB; kiểm hằng ngày trong hypercare; báo cáo tháng. Dự phòng 12,292,906 VND, khoảng 22 TB ở 550 VND/GB gồm VAT; mặc định 128 kbps.','PR-04, dự báo chi phí/báo cáo tháng; chỉ đổi cost baseline sau phê duyệt.'),
    ('Lỗ hổng bảo mật','Nếu scan 1.5.3 thấy critical/high, sửa có thể vượt 5 ngày của 1.5.4, chậm RC và trượt NFR-05.','Review bảo mật mỗi PR; ZAP baseline 1.3.5; ghi lỗi. S1 sửa 1 ngày, S2 3 ngày làm việc; còn mở 27/11 thì phân tích nguyên nhân. Giờ vượt dùng contingency 1.5 theo phê duyệt.','Defect log, checklist/báo cáo security; QMP/Dictionary nếu đổi quy trình/tiêu chí; lịch và dự báo chi phí.'),
    ('Không đáp ứng bảo vệ dữ liệu','Nếu xử lý tài khoản/lịch sử nghe không đáp ứng yêu cầu bảo vệ dữ liệu của bài, có thể bị Customer từ chối nghiệm thu hoặc store từ chối.','Đối chiếu consent, thông báo Việt/Anh, xóa tài khoản/dữ liệu FR-02, chỉ dùng lịch sử cho FR-08. Map nghĩa vụ sang test trước review 1.2.8; mục thiếu làm change request 14/10; kiểm lại trong 1.5.3/1.6.6.','Requirements/SRS/RTM FR-01/02/08; thông báo riêng tư; checklist thiết kế/security/store.'),
    ('Vắng nhân sự','Nếu một người nghỉ > 3 ngày làm việc, WP có một chủ có thể dừng, trễ việc phụ thuộc.','Ghi backup đủ kỹ năng trong PMP/Team Charter; bàn giao và kiểm ngày trống ở 1.7.5. Trước đổi người phải kiểm kỹ năng, dependency, không chồng việc và <= 176 h/tháng; duyệt đổi lịch/chi phí.','Team Charter/backup; Resource Requirements; lịch/assignment; Dictionary và dự báo chi phí.'),
    ('App trượt chỉ tiêu thiết bị','Nếu cold/warm start hoặc pin không đạt trên 4 máy, cần sửa sau test và có thể chậm RC.','Đo sớm 10 cold + 10 warm start trên Galaxy A07/iPhone 11 ở 1.3.5; trượt median là trigger, ghi S2 và sửa 3 ngày; bật Crashlytics từ build đầu.','Thiết kế/code Flutter, báo cáo device/defect log, Dictionary 1.3/1.5, lịch và dự báo chi phí.'),
    ('Phục hồi không đạt RTO/RPO','Nếu restore > 4 h hoặc mất > 24 h dữ liệu, có thể hoãn go-live.','Bật backup ngày ở 1.2.7; viết/diễn tập hướng dẫn 1.6.9. Rehearsal > 3 h là trigger; đo lại 1.6.4, sửa và diễn tập lại trước go-live nếu > 4 h.','Hướng dẫn vận hành, cấu hình và báo cáo restore, checklist release, lịch.'),
    ('Nghẽn ở Technical Lead','Nếu API decision/review chờ Technical Lead > 1 ngày làm việc, development phụ thuộc có thể dừng, chậm 1.3.8/1.4.8.','Dùng API contract 1.2.6; có người review dự phòng ở PMP. Kiểm queue/giờ còn lại ở họp tuần, 1.7.4/1.7.5. Chờ > 1 ngày hoặc thiếu công suất trước hạn thì giao reviewer đủ kỹ năng có lịch trống; code của Tech Lead phải được người khác review.','Team Charter/review cover; Resource Requirements; lịch; Dictionary và dự báo chi phí.'),
    ('Thiếu kỹ năng làm giảm chất lượng','Nếu người được giao thiếu kỹ năng, code có thể trượt nghiệm thu/lặp lỗi, tăng rework và chậm RC.','Đối chiếu kỹ năng trước giao việc. PR cần peer duyệt, unit test pass, coverage backend >= 70%. S1/S2, quality gate trượt hoặc >= 3 lỗi cùng nguyên nhân thì phân tích theo QMP; đổi assignment/checklist, dùng WP sửa lỗi; phần vượt cần duyệt.','Resource Requirements/Team Charter; QMP/checklist; defect log/test; Dictionary, lịch và dự báo chi phí.')]
    u,l=revised_risks()
    p(d,'Cột P / I ghi xác suất và 4 tác động theo thứ tự phạm vi/chất lượng/lịch/chi phí; dấu → chỉ mục tiêu sau hành động.')
    tb(d,['ID','Rủi ro','Owner / xử lý','P / I ban đầu','P / I sau','Điểm đầu → sau'],[[r[0],risks[i][0],NAMES[r[2]]+' / '+NAMES[l[i][7]],r[3]+' / '+','.join(r[4:8]),l[i][1]+' / '+','.join(l[i][2:6]),r[8]+' → '+l[i][6]] for i,r in enumerate(u)],[850,2800,1700,1550,1550,1450],7.5)
    for i,x in enumerate(risks):
        h(d,f'R-{i+1:02d} {x[0]}')
        for j,txt in enumerate([x[1],'Ứng phó/trigger: '+x[2],'Tài liệu cập nhật nếu xảy ra: '+x[3]]):
            rp=p(d,txt.replace('thì escalates','thì báo PM'))
            rp.paragraph_format.keep_with_next=j<2
            rp.paragraph_format.keep_together=True
    h(d,'Ứng phó rủi ro có phải thêm WP không')
    p(d,'Mọi việc được duyệt phải nằm trong phạm vi và ước lượng. Phòng ngừa đã có như unit test/review, test sớm 1.3.5, backup 1.2.7, quản lý rủi ro 1.7.5 thì ghi rõ vào WP đó; không tạo WP trùng. Nếu thêm công việc mới vượt ước lượng/phạm vi, lập change request, cập nhật WBS/Dictionary, lịch, nguồn lực, chi phí sau duyệt. Có tiền reserve không tự làm tăng giờ trong lịch.')

def stake_vi(d):
    newpage(d,'8 Danh sách các bên liên quan')
    p(d,'Stakeholder Register theo Book of Forms 1.3: Name, Position/Role, Contact Information, Requirements, Expectations, Classification. Phân loại dưới đây là đánh giá quyền quyết định/mức quan tâm phục vụ kế hoạch. Địa chỉ/email cụ thể chưa có trong nguồn nên chỉ ghi kênh liên hệ; PM bổ sung trước dùng.')
    rows=[
    ['Nguyễn Trung B / Quỹ PTIT','Sponsor','Qua PM','Duyệt tài trợ/PMP, thay đổi vượt PM; ngân sách <= 757 triệu, ký đóng 15/01.','Ngoài; quyền cao / quan tâm cao'],
    ['Đỗ Thị Bích Ngọc','Customer','Qua PM','Yêu cầu/thiết kế, tự UAT; Must/Should đạt, UAT >= 4.2/5, ký 04/01.','Ngoài; cao / cao'],
    [MINH,'PM','Jira/chat; bổ sung email','Quản lý baseline, nguồn lực, rủi ro và giao tiếp; quyết định đúng hạn.','Trong; cao / cao'],
    [DAI,'BA / Full-stack','Jira/GitHub/chat','Giữ yêu cầu/RTM; code API/Flutter; tiêu chí đo được, API đã duyệt.','Trong; vừa / cao'],
    [HIEU,'Backend / Tech Lead','Jira/GitHub/chat','Kiến trúc/API/backend, review; p90 < 200 ms, coverage >= 70%.','Trong; vừa / cao'],
    [A,'QA / DevOps','Jira/GitHub/chat','Truy vết test, tải/security/device, CI/CD; 15,000 VU đạt, RC không S1/S2.','Trong; vừa / cao'],
    [THU,'UI/UX / Flutter','Thiết kế/Jira/chat','UI đúng mockup/guideline, Việt/Anh; một codebase; nộp store.','Trong; vừa / cao'],
    ['Người nghe','End-users','Customer/UAT','Khám phá, nghe, tổ chức nhạc; Live Premiere đúng FR-09; bảo vệ/xóa dữ liệu.','Ngoài; thấp / cao'],
    ['Admin Customer','Quản trị nội dung','Qua Customer/PM','CRUD FR-12 đạt; học vận hành 12–13/01; có hướng dẫn.','Ngoài; vừa / cao'],
    ['Bizfly Cloud','Cloud, DB, CDN','Portal qua PM/DevOps','Kích hoạt A-14 trong 2 ngày; đủ PR-01/02/04, dịch vụ và hóa đơn theo thỏa thuận.','Ngoài; vừa / vừa'],
    ['VNG Cloud','Object storage','Portal qua PM/DevOps','1,000 GB/tháng x 3; kích hoạt trước setup; rõ truy cập/phí.','Ngoài; vừa / vừa'],
    ['Apple App Store','Phân phối/billing iOS','App Store Connect qua Thu','Build đúng policy, privacy/xóa tài khoản/billing; TestFlight chạy được.','Ngoài; cao / thấp'],
    ['Google Play','Phân phối/billing Android','Play Console qua Thu','Build đúng policy/Data safety/xóa tài khoản; tài khoản xác minh, internal test chạy.','Ngoài; cao / thấp']]
    tb(d,['Tên','Vai trò','Kênh','Yêu cầu / kỳ vọng','Phân loại'],rows,[1600,1300,1500,3900,1600],7.5)
    p(d,'Đối chiếu Charter, Requirements, QMP, Resource Requirements và Risk Register. PM rà soát tại 1.7.5 và sau thay đổi được duyệt; cập nhật tên admin và liên hệ cụ thể khi Customer/nhà cung cấp xác nhận.')
    h(d,'Các chỉnh sửa từ note')
    p(d,'Date Prepared chỉ còn một ngày. Cây WBS bỏ milestone nhưng lịch giữ event/ID. Dictionary ghi rõ cấp control account. Risk Register đưa thang điểm trước, còn hai bảng dữ liệu xuyên trang, bổ sung R-13/R-14 và tài liệu chịu tác động. Stakeholder Register được thêm thành phần 8. Số giờ, ngày và ngân sách baseline giữ nguyên.')
    p(d,'Nguồn đối chiếu: A Project Manager’s Book of Forms, mục 1.3 (trang in 12–14), 2.9 (49–51), 2.10 (52–55); báo cáo và lịch của nhóm. Bản tiếng Anh giữ các bảng số liệu chi tiết và URL nguồn giá; bản này giảm phần diễn giải lặp để nhóm đọc chung.')

if __name__=='__main__':build()
