# BRIEF — Bài tập QLDA nhóm 04 (Music Streaming Platform)

Đọc file này trước khi làm. Trả lời user bằng **tiếng Việt**; nội dung trong tài liệu viết bằng **tiếng Anh**.

## 1. Bối cảnh
- Môn Software Project Management, Class 01 - Group 04. Giảng viên Dr. Đỗ Thị Bích Ngọc — trong tài liệu bà là **Customer**. Sponsor: Nguyễn Trung B (PTIT Innovation Fund).
- Nhóm 5 người (người thứ 5 "Trần Văn A" là nhân vật giả định):

| Người | Vai trò | Rate VND/h | Giờ WBS |
|---|---|---|---|
| Trần Vũ Tiến Minh | Project Manager | 206,067 | 512 |
| Nguyễn Phú Đại | Business Analyst / Full-Stack Developer | 118,566 | 504 |
| Trần Đình Hiếu (user) | Backend Developer / Technical Lead | 175,001 | 544 |
| Trần Văn A | QA Engineer / DevOps Engineer | 124,261 | 448 |
| Lê Quang Thu | UI/UX Designer / Flutter Developer | 198,818 | 456 |

- Dự án: "Development and Deployment of a Music Streaming Platform". Gồm Flutter (một codebase cho iOS/Android/Web), Node.js (Express.js), PostgreSQL, WebSocket cho Live Premiere, object storage + CDN.
- Giáo trình: `A_Project_Managers_Book_of_Forms-2.pdf` (mẫu biểu; trang vật lý = số trang in + 11) và `PMBOK-6-EN.pdf` (tham khảo).

## 2. File trong AA/
- `4.docx` — tài liệu **duy nhất**, 43 trang, 17 section. Cấu trúc ở mục 4.
- `4.xml` — lịch Microsoft Project (MSPDI), 86 task. Không xuất PDF được bằng code; **user tự mở bằng MS Project và xuất**.
- `schedule.pdf` — Gantt user đã xuất từ 4.xml (2 trang, khổ lớn 34x22 in, dùng tên task mới). **Chưa nhúng** vào ô placeholder ở mục 3 của 4.docx.
- `Setup_Schedule_View.bas` — macro MS Project dựng bảng WBS | Task Name | Duration | Start | Finish | WBS Predecessors | Resource Names. Phím F của user là phím âm lượng, nên chỉ hướng dẫn thao tác bằng menu (View → Macros…).
- `promt.txt` — yêu cầu hiện tại của user. `tools/` — script hỗ trợ (mục 7).

## 3. Quy tắc bắt buộc
- **Không dùng git** dưới mọi hình thức (user cấm từ 04/11/2026).
- Gặp điểm chưa rõ thì **hỏi user trước** (dùng AskUserQuestion, câu hỏi bằng tiếng Việt, đưa lựa chọn và ghi rõ phương án đề xuất).
- Chỉ làm phần được yêu cầu, không thêm biểu mẫu mới ngoài yêu cầu. Nội dung lấy từ dữ liệu đã có trong tài liệu và phải nhất quán với nó.
- Góp ý của giảng viên phải tuân theo:
  - không dùng tính từ mơ hồ (stable, quickly, relevant, minor, major…) → viết tiêu chí đo được;
  - tên task/WBS bắt đầu bằng động từ, ngắn gọn (tối đa khoảng 52 ký tự);
  - unit test do developer viết, nằm trong work package phát triển;
  - UAT do Customer làm, nằm ngoài scope.
- Bảng có tính toán thì "viết như Excel": header ghi công thức, ví dụ `Hours (A)`, `Total (C = A x B)`, `(SUM)`. Vẫn giữ đúng các cột của form trong giáo trình.
- Mọi con số, ngày, ID phải khớp chéo giữa các bảng và với 4.xml. Sau mỗi lần sửa phải chạy `tools/check2.py` và kết quả phải là `PROBLEMS: 0`.
- Định dạng:
  - font **Arial**; chữ trong bảng 7–7.5 pt (w:sz 14–15); title bar nền tối, chữ trắng đậm;
  - mỗi biểu mẫu là **một section**, footer căn giữa `Page {PAGE} of {SECTIONPAGES}`, đánh số trang lại từ 1 (dùng `tools/footers.py`);
  - tiêu đề nhóm dạng `N. CREATE ...` (Arial đậm 13 pt, outline level 1) — dùng `dlib.heading_par`.
- Sửa docx bằng python-docx + lxml (`tools/dlib.py`). Word chỉ dùng qua COM để xuất PDF kiểm tra bố cục. **Không kill WINWORD** (bị chặn); `topdf.ps1` tự đóng Word.

## 4. Cấu trúc 4.docx (thứ tự theo danh sách user đưa; trong ngoặc là số trang)
- Cover: không có footer.
- **1. CREATE PROJECT CHARTER**
  - Project Charter (4)
  - Assumption Log A-01..A-14 (1)
  - Requirements Documentation FR-01..FR-14, NFR-01..NFR-12 (2)
  - Project Scope Statement (2)
- **2. CREATE WBS**
  - Requirements Traceability Matrix + Inter-RTM BO-1..BO-6 / TR-01..TR-16 (3)
  - Work Breakdown Structure (2)
  - WBS Dictionary 1.1–1.7 + Resource Summary (8)
- **3. CREATE PROJECT SCHEDULE**
  - Ô placeholder "[Insert here the Gantt chart exported from 4.xml.]" (1)
- **4. CREATE COST ESTIMATION**
  - Activity Cost Estimates (4)
  - Cost Estimating Worksheet (1)
  - Bottom-up Cost Estimating Worksheet (2)
  - Basis of Estimates and References (2)
- **5. CREATE QUALITY MANAGEMENT PLAN**
  - QMP: ISO/IEC 25010, 12 quality objectives (3)
- **6. RESOURCE MANAGEMENT**
  - Team Charter (2)
  - Resource Requirements: người + PR-01..PR-12 (4)
  - Resource Breakdown Structure (1)
- **7. CREATE RISK REGISTER** (4) — xong ngày 04/11/2026. Khổ dọc, mỗi trang có bảng trên (ID..Response) và bảng dưới (ID, Revised..Comments) cho 3 rủi ro; trang 4 có thêm bảng "Probability and Impact Scale (1-5)".
  - R-01..R-12. Thang 1–5. Score C = A x MAX(B1:B4), Revised F = D x MAX(E1:E4). Date Prepared September 29, 2026; tất cả Open. Cao nhất là R-02 WebSocket (16).
  - Đã thêm tham chiếu ở: Charter Key Deliverables ("Risk Register"), Scope Statement mục 1.7, WBS Dictionary 1.7 (Description, Assumptions, Quality Requirements).

Tìm bảng bằng chữ ở ô đầu tiên (ví dụ `"ACTIVITY COST ESTIMATES"`), không dùng chỉ số.

## 5. Dữ liệu gốc (không được lệch)
- **Thời gian:**
  - Aug 17, 2026 – Jan 15, 2027 = 108 ngày làm việc (T2–T6, 8–12h và 13–17h).
  - Nghỉ Sep 1–2 (Quốc khánh, doanh nghiệp chọn 1–2/9) và Jan 1. Làm bù thứ Bảy Sep 5.
  - Tháng dự án: M1 Aug 17–Sep 16, M2 Sep 17–Oct 16, M3 Oct 17–Nov 16, M4 Nov 17–Dec 16, M5 Dec 17–Jan 15.
- **Quy tắc nguồn lực:** mỗi work package có đúng 1 người phụ trách; mỗi người làm 1 việc một lúc (không trùng lịch); mỗi WP 8–80 h; mỗi người tối đa 176 h/tháng.
  - Giờ theo tháng: Minh 152/152/40/64/104; Đại 72/128/160/96/48; Hiếu 64/160/168/48/104; A 88/96/80/88/96; Thu 120/104/144/56/32.
  - Cả nhóm: 496/640/592/352/384 = 2,464 h (trên tổng 4,320 h khả dụng).
- **Chi phí (VND):**
  - Labor 406,794,048.
  - Materials 104,400,000: 1.2.7 cloud/CDN 3 tháng × 18,500,000; 1.5.3 thiết bị test 28,000,000 + k6 17,600,000; 1.6.6 phí store 3,300,000.
  - Contingency reserve 137,904,988 = Σ(P − M) của từng WP. M = most likely; O = M − 10%; P = M + overrun % theo phase: 1.1 15%, 1.2 20%, 1.3 40%, 1.4 40%, 1.5 35%, 1.6 30%, 1.7 15%.
    - Contingency theo phase: 1.1 8,429,891; 1.2 26,802,470; 1.3 24,184,535; 1.4 24,999,699; 1.5 27,811,401; 1.6 15,785,776; 1.7 9,891,216.
  - Work package estimates 649,099,036. Warranty (ngoài WBS) 71,822,880 = Hiếu 240 h + A 240 h.
  - Cost baseline 720,921,916. Management reserve 5% = 36,046,096 (cho unknown risks, chỉ Sponsor được giải ngân). Budget 757,000,000.
  - PM được tự duyệt thay đổi dưới 2% ngân sách (15,140,000) hoặc dưới 3 ngày; lớn hơn phải trình Sponsor.
- **Milestone:** 1.1.8 Sep 18 · 1.7.2 Oct 1 (PMP được duyệt) · 1.2.11 Oct 21 · 1.3.6 MVP demo Nov 12 · 1.3.8 Nov 20 · 1.4.6 demo Nov 30 · 1.4.8 Dec 7 · 1.5.6 Dec 14 · 1.5.8 RC Dec 18 · 1.6.3 UAT sign-off Jan 4 · 1.6.12 và 1.7.9 Jan 15.
- **Công việc quản lý dự án:**
  - 1.7.1 PMP (Sep 16–29) gồm cả kế hoạch rủi ro; các form PMP (QMP, Team Charter, Assumption Log…) đều đề **Date Prepared: September 29, 2026**.
  - 1.7.4: 5 báo cáo cuối tháng vào Sep 11; Oct 15–16; Nov 13 và 16; Dec 10–11; Jan 14–15.
  - 1.7.5 "Monitor risks and manage change requests" (Minh, 64 h): 8 buổi review rủi ro vào Oct 14, Oct 28, Nov 11, Nov 25, Dec 9, Dec 23, Jan 5, Jan 13 (buổi cuối).

## 6. Đầu vào cho Risk Register (Book of Forms 2.32, tr. vật lý 139–141)
- **Cột của form:** ID | Risk Statement | Owner | Probability | Impact (Scope / Quality / Schedule / Cost) | Score | Response | Revised Probability | Revised Impact (Scope / Quality / Schedule / Cost) | Revised Score | Responsible Party | Actions | Status | Comments.
  - Câu risk statement có dạng "EVENT may occur, causing IMPACT" hoặc "If CONDITION exists, EVENT may occur, leading to EFFECT".
  - Score = Probability × Impact → ghi công thức ở header.
  - Status: open/closed.
- **Rủi ro đã nêu sẵn trong tài liệu** (register phải khớp với các chỗ này):
  - Charter → Overall Project Risk "Medium", gồm 4 rủi ro:
    1. Customer/Sponsor phản hồi, ký duyệt chậm hoặc vắng mặt (A-01..A-12);
    2. WebSocket của Live Premiere có thể không scale tuyến tính → phải load test thêm hoặc đổi kiến trúc;
    3. năng lực nguồn lực (2,464 h trên 4,320 h) → cắt Could-have (FR-05, FR-10, FR-13) trước;
    4. thay đổi yêu cầu giữa chừng.
  - Assumption Log: A-13 App Store/Google Play duyệt tối đa 3 ngày (Thu phụ trách); A-14 tài khoản cloud/CDN kích hoạt trong 2 ngày (A phụ trách).
  - Activity Cost Estimates:
    - 1.3.3 ghi "Live Premiere WebSocket scaling is the main technical risk";
    - phase 1.3 và 1.4 có overrun +40%, độ tin cậy Medium-Low (bất định cao nhất).
  - Resource Requirements PR-04: lưu lượng CDN vượt 5 TB/tháng được trả bằng contingency của 1.2.7 (12,292,906 VND, khoảng 24 TB).
  - QMP:
    - mức lỗi S1–S4 (S1 sửa trong 1 ngày, S2 trong 3 ngày);
    - corrective action và change control qua 1.7.5;
    - các chỉ tiêu: 9,000–15,000 user, p90 < 200 ms, crash-free ≥ 99%, coverage ≥ 70%, RPO 24 h / RTO 4 h;
    - OWASP ASVS/MASVS; Luật Bảo vệ dữ liệu cá nhân của Việt Nam (hiệu lực 1/1/2026); chính sách duyệt của App Store và Google Play.
- **Gợi ý:**
  - Owner, responsible party là người trong nhóm.
  - Ngày review khớp các buổi 1.7.5.
  - Response nào dùng tiền thì lấy từ contingency của phase tương ứng hoặc management reserve.
- **Vị trí:** sau RBS, section mới có tiêu đề `7. CREATE RISK REGISTER` và footer Page x of y. Form có khoảng 17 cột nên khổ dọc có thể chật — cân nhắc section khổ ngang hoặc hỏi user.
- Sau khi thêm register, rà các chỗ đã nhắc tới nó để cập nhật cho khớp: WBS Dictionary 1.7 (Technical Information, Quality Requirements "risk register is reviewed at least every two weeks"), mô tả 1.7.1/1.7.5, Charter Key Deliverables.

## 7. Công cụ (`tools/`) và quy trình
- `dlib.py`: các hàm hỗ trợ.
  - `Repl(root, name)(old, new, expect)`: thay chữ có xét nhiều run, kiểm tra số lần thay đúng `expect`.
  - Đọc/ghi ô: `set_cell_text`, `set_cell_lines`, `rows`, `cells`, `cell_text`, `find_row`.
  - Bố cục: `set_grid` (độ rộng cột), `heading_par`, `section_break_par`, `page_break_par`.
- `check2.py <xml> <docx>`: kiểm tra chéo WBS / Dictionary / Resource Requirements / ACE / Resource Summary / Requirements Documentation↔RTM với 4.xml, và tìm cụm từ cũ còn sót.
- `mspdi.py <xml>`: liệt kê task. `fwdpass.py <xml>`: tính lại lịch, phát hiện trùng việc, tính giờ theo tháng.
- `footers.py <in> <out>`: đặt footer Page x of y cho mọi section trừ trang bìa, đánh số lại từ 1.
- `topdf.ps1`: Word COM → PDF. Gọi bằng `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\topdf.ps1 -Files <docx> -OutDir <dir>` (dùng đường dẫn Windows).
- `montage.py <pdf> <png> <first> <last> <cols> <dpi>`: ghép ảnh nhiều trang để xem nhanh.
- `pdftext.py <pdf> 139-141`: lấy chữ từ giáo trình.
- `mpxjread.py <xml>`: kiểm tra 4.xml đọc được bằng MPXJ. Cần `JAVA_HOME="E:/Program Files/Java/jdk-17"`; pip `mpxj` và `jpype1` đã cài.
- `reference_edit4.py`: **chỉ để tham khảo**. Cho thấy cách clone title bar, header row, label + box của Charter, và clone bảng dữ liệu (Requirements Documentation được clone từ Assumption Log).
- **Quy trình:**
  1. Chép 4.docx ra scratchpad.
  2. Sửa bằng python.
  3. Chạy `footers.py`.
  4. Chạy `check2.py` (phải ra 0).
  5. Chạy `topdf.ps1` + `montage.py`, xem ảnh để chắc mỗi form không tràn trang.
  6. Chép lại vào AA/.
- **Lưu ý môi trường Windows:**
  - đặt `PYTHONIOENCODING=utf-8` (tên tiếng Việt);
  - trong Bash không dùng `cat > file` mà không có heredoc (lệnh sẽ treo).

## 8. Việc còn dở
1. ~~Risk Register~~ — xong (mục 4). Cách dựng xem `tools/build_rr.py` (chỉ để tham khảo, script dừng nếu register đã có); `check2.py` đã kiểm tra thêm score, ID, owner, status của register.
2. Nhúng `schedule.pdf` vào ô placeholder ở mục 3. Hỏi user muốn chèn dạng ảnh hay object; nếu thêm trang thì footer phải đếm đúng.
