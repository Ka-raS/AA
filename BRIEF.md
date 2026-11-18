# BRIEF — Bài tập QLDA nhóm 04 (Music Streaming Platform)

Đọc file này trước khi làm. Trả lời user bằng **tiếng Việt**; nội dung trong tài liệu viết bằng **tiếng Anh**.

## 1. Bối cảnh
- Môn Software Project Management, Class 01 - Group 04. Giảng viên Dr. Đỗ Thị Bích Ngọc — trong tài liệu bà là **Customer**. Sponsor: Nguyễn Trung B (PTIT Innovation Fund).
- Nhóm 5 người (người thứ 5 "Trần Văn A" là nhân vật giả định, không ghi trên bìa):

| Người | Vai trò | Rate VND/h | Giờ WBS | Mã SV |
|---|---|---|---|---|
| Trần Vũ Tiến Minh | Project Manager | 206,067 | 512 | B23DCCE067 |
| Nguyễn Phú Đại | Business Analyst / Full-Stack Developer | 118,566 | 504 | B23DCVT071 |
| Trần Đình Hiếu (user) | Backend Developer / Technical Lead | 175,001 | 544 | B23DCCE034 |
| Trần Văn A | QA Engineer / DevOps Engineer | 124,261 | 448 | — |
| Lê Quang Thu | UI/UX Designer / Flutter Developer | 198,818 | 456 | B23DCCE091 |

- Dự án: "Development and Deployment of a Music Streaming Platform". Gồm Flutter (một codebase cho iOS/Android/Web), Node.js (Express.js), PostgreSQL, WebSocket cho Live Premiere, object storage + CDN.
- Giáo trình: `A_Project_Managers_Book_of_Forms-2.pdf` (mẫu biểu; trang vật lý = số trang in + 11) và `PMBOK-6-EN.pdf` (tham khảo).

## 2. File trong AA/
- `4.docx` — tài liệu **duy nhất**, 49 trang (bìa + 48), 20 section (bìa + 19 form). Cấu trúc ở mục 4.
- `4_vi.docx` — bản tiếng Việt rút gọn của 4.docx để nhóm đọc (tạo 18/11/2026), cùng bố cục, đủ các form và bảng; **không phải bản nộp**, không chạy check2/colwidth trên file này. Khi 4.docx đổi nội dung thì hỏi user có cần cập nhật bản này không. **Cập nhật 18/11/2026** theo các sửa đổi cùng ngày của 4.docx (câu chữ khách quan, merge dainp, số trang lưu sẵn) bằng `tools/vi_update_20261118.py`; 49 trang, footer "Trang x / y".
- `4.xml` — lịch Microsoft Project (MSPDI), **86 task**, ID liên tiếp. Milestone **không có mã WBS**: cột WBS ghi M-01..M-09 (đúng như trong 4.docx), mỗi milestone là **dòng cuối của phase chứa nó** (promt 18/11/2026; không còn nhóm "Milestones" ở cuối lịch): 1.1: M-01 · 1.2: M-03 · 1.3: M-04 · 1.4: M-05 · 1.5: M-06 · 1.6: M-07, M-08 · 1.7: M-02, M-09 (OutlineNumber 1.1.8, 1.2.11, 1.3.8, 1.4.8, 1.5.8, 1.6.11-12, 1.7.8-9). Ngày: M-01 SRS/RTM sign-off Sep 18, M-02 PMP approved Oct 1, M-03 design approved Oct 21, M-04 MVP Nov 20, M-05 all functions Dec 7, M-06 RC Dec 18, M-07 UAT sign-off Jan 4, M-08 apps handed over Jan 15, M-09 project closed Jan 15. Liên kết và ngày không đổi; summary 1.1 kết thúc Sep 18, 1.2 kết thúc Oct 21 vì chứa milestone. OutlineNumber của mọi work package trùng mã WBS.
- `schedule(v3).pdf` — Gantt hiện hành (18/11/2026), xuất bằng `tools/export_gantt.py` (MS Project qua COM): 1 trang 34x22 in ngang, cùng bảng với macro, thang thời gian tháng/tuần; kiểu thanh, legend, gridline mặc định. `schedule.pdf`, `schedule(v2).pdf` là bản cũ (nhóm Milestones ở cuối). **Chưa nhúng** vào ô placeholder ở mục 3 của 4.docx.
- `Setup_Schedule_View.bas` — macro MS Project dựng bảng WBS | Task Name | Duration | Start | Finish | WBS Predecessors | Resource Names. Phím F của user là phím âm lượng, nên chỉ hướng dẫn thao tác bằng menu (View → Macros…).
- `promt.txt` — yêu cầu hiện tại của user. `tools/` — script hỗ trợ (mục 7).

## 3. Quy tắc bắt buộc
- **Không dùng git** dưới mọi hình thức (user cấm từ 04/11/2026).
- Gặp điểm chưa rõ thì **hỏi user trước** (dùng AskUserQuestion, câu hỏi bằng tiếng Việt, đưa lựa chọn và ghi rõ phương án đề xuất).
- Chỉ làm phần được yêu cầu, không thêm biểu mẫu mới ngoài yêu cầu. Nội dung lấy từ dữ liệu đã có trong tài liệu và phải nhất quán với nó.
- Góp ý của giảng viên phải tuân theo:
  - không dùng tính từ mơ hồ (stable, quickly, relevant, minor, major, "in real time", "without degradation"…) → viết tiêu chí đo được;
  - (user yêu cầu 18/11/2026) mọi form **ngoài Project Charter** viết cho kỹ sư trong nhóm đọc: khách quan, kỹ thuật, không tâng bốc. Thay "large/small changes" bằng ngưỡng 2% / 3 ngày, "as capacity allows" / "if the schedule slips" bằng ngưỡng của R-03 (trễ > 3 ngày → change request), "main/critical" bằng mã và điểm (R-02 score 16, S1). Mức S1–S4 có định nghĩa trong QMP. Script: `tools/objective_20261118.py`;
  - tên task/WBS bắt đầu bằng động từ, ngắn gọn (tối đa khoảng 52 ký tự);
  - unit test do developer viết, nằm trong work package phát triển;
  - UAT do Customer làm, nằm ngoài scope; vì vậy mọi yêu cầu Must/Should phải có thêm một phép kiểm do nhóm làm (Functional test, Design review…), không chỉ "UAT".
- **Tiêu đề cột ghi đúng như sách** (user chọn 04/11/2026, áp dụng cho mọi bảng): không dùng ký hiệu kiểu `Hours (A)`, `Total (C = A x B)`, `(SUM)`, `Score (C = A x MAX(B1:B4))`. Công thức ghi trong **chú thích dưới bảng** (ACE, Cost Estimating Worksheet, Bottom-up, Risk Register) hoặc một dòng chú thích in nghiêng ngay dưới bảng chi phí (WBS Dictionary). Vẫn giữ đúng các cột của form trong giáo trình.
- Mọi con số, ngày, ID phải khớp chéo giữa các bảng và với 4.xml. Sau mỗi lần sửa phải chạy `tools/check2.py` và kết quả phải là `PROBLEMS: 0`.
- **Đồng bộ 4.xml cùng lúc với tài liệu** (user yêu cầu 18/11/2026): sửa 4.docx mà động tới lịch (mã WBS, tên task, ngày, người phụ trách, giờ, mã được nhắc trong Notes) thì sửa luôn 4.xml (`<WBS>`, `<OutlineNumber>`, `<Name>`, `<Notes>`…), chạy `check2.py` + `mpxjread.py`, rồi nhắc user mở 4.xml bằng MS Project để xuất lại schedule.pdf.
- Định dạng:
  - font **Arial**; title bar nền tối, chữ trắng đậm 16 pt; dòng Project Title / Date Prepared **8 pt** (w:sz 16) ở mọi form; nội dung bảng 7–7.5 pt (w:sz 14–15), header tối thiểu 7 pt; cây WBS/RBS giữ 9 pt;
  - độ rộng cột: sau khi sửa nội dung chạy `tools/colwidth.py` (cột chữ dài được rộng, cột 1 ký tự hẹp; không để số hay từ ngắn bị gãy; tên người tối đa 2 dòng);
  - mỗi biểu mẫu là **một section**, footer căn giữa `Page {PAGE} of {SECTIONPAGES}`, đánh số trang lại từ 1 (`tools/footers.py`), sau đó ghi số trang thật vào giá trị lưu sẵn của trường (`tools/pagecount.py`) để trình xem không cập nhật trường vẫn hiện đúng "of N";
  - tiêu đề nhóm dạng `N. CREATE ...` (Arial đậm 13 pt, outline level 1) — dùng `dlib.heading_par`;
  - đoạn ngắt section sau bảng cuối của form cao 1 pt, không có đoạn trống thừa trước nó (tránh trang trắng).
- Bìa (trang 1, user yêu cầu 05/11/2026): font **Times New Roman** (ngoại lệ duy nhất so với quy tắc font Arial), gồm POSTS AND TELECOMMUNICATIONS INSTITUTE OF TECHNOLOGY / FACULTY OF INFORMATION TECHNOLOGY 1 / logo PTIT / ASSIGNMENT REPORT / tên dự án / COURSE: Software Project Management / bảng Class E23CNPM02, Group 04, Students (4 người + mã SV), Instructor / Hanoi – 2026. Khung viền và logo PTIT là hai ảnh trong 4.docx (rId26/rId27). Không đổi bìa trừ khi user yêu cầu. Bìa không có footer.
- Sửa docx bằng python-docx + lxml (`tools/dlib.py`). Word chỉ dùng qua COM để xuất PDF kiểm tra bố cục, và **chỉ khi user yêu cầu** (18/11/2026 user nói không cần xuất PDF). **Không kill WINWORD** (bị chặn); `topdf.ps1` tự đóng Word.

## 4. Cấu trúc 4.docx (trong ngoặc là số trang)
Thứ tự form theo số thứ tự trong Book of Forms (user chốt 18/11/2026): 1.1 Charter → 1.2 Assumption Log → 1.3 Stakeholder Register → 1.4 Stakeholder Analysis → 2.6 Requirements Documentation → 2.7 RTM → 2.8 Scope Statement → 2.9 WBS → 2.10 WBS Dictionary → 2.18 Schedule → 2.20/2.21 chi phí → 2.23 QMP → 2.27 Team Charter → 2.28 Resource Requirements → 2.29 RBS → 2.32 Risk Register → bảng thang Probability and Impact (2.34/2.35) ở cuối section Risk Register. Dời form thì dời nguyên khối section (footer riêng đi theo, giữ số trang lưu sẵn).
- Cover (1): không có footer.
- **1. CREATE PROJECT CHARTER**
  - Project Charter (4) — Date Prepared August 27, 2026; 757,000,000 VND là **pre-approved budget** (ước tính sơ bộ lúc khởi tạo, các khoản ghi "about ...", cộng đúng 757,000,000), Charter không nhắc "cost baseline". Approvals có tên Sponsor Nguyễn Trung B.
  - Assumption Log A-01..A-14 (1)
  - Stakeholder Register (1) và Stakeholder Analysis (1) — Book of Forms 1.3/1.4, Date Prepared August 27, 2026; 12 stakeholder (Customer, Sponsor, 5 thành viên, Customer administrators, end users, App Store/Google Play, cloud/CDN provider, Ministry of Public Security); email giả định dạng PTIT; Interest/Influence thang 1–5, Attitude Supportive/Neutral/Resistant; Classification = Internal/External + mức (Manage closely / Keep satisfied / Keep informed / Monitor) suy từ điểm.
  - Requirements Documentation FR-01..FR-14, NFR-01..NFR-12 (2) — Date Prepared September 17, 2026
- **2. CREATE WBS**
  - Requirements Traceability Matrix + Inter-RTM BO-1..BO-6 / **TR-01..TR-18** (3)
  - Project Scope Statement (2) — Date Prepared October 8, 2026
  - Work Breakdown Structure (2) — chỉ có phase và work package; milestone không nằm trong WBS (chú thích dưới cây nói rõ)
  - WBS Dictionary 1.1–1.7 + Resource Summary (8) — Date Prepared October 8, 2026; ở mức control account (chú thích dưới cây WBS nói rõ)
- **3. CREATE PROJECT SCHEDULE**
  - Ô placeholder "[Insert here the Gantt chart exported from 4.xml.]" (1)
- **4. CREATE COST ESTIMATION**
  - Activity Cost Estimates (5) — trang 5 chỉ có 3 dòng tổng, do các cột số phải đủ rộng để không gãy số
  - Cost Estimating Worksheet (1)
  - Bottom-up Cost Estimating Worksheet (2)
  - Basis of Estimates and References (2)
- **5. CREATE QUALITY MANAGEMENT PLAN**
  - QMP: ISO/IEC 25010, 12 quality objectives (3 từ 18/11/2026, sau khi thêm định nghĩa S2-S4 — trước đó 2)
- **6. RESOURCE MANAGEMENT**
  - Team Charter (2) — cột chữ ký giữ 50/50, không chạy colwidth
  - Resource Requirements: người + PR-01..PR-12 (3, từ 18/11/2026) — Date Prepared October 8, 2026 (RBS cũng vậy). "one task at a time" và kỹ năng theo vai trò chỉ ghi **một lần** ở dòng "Team resources (people)", không lặp ở từng dòng (lấy từ nhánh dainp 18/11/2026). WBS Dictionary: ô Assumptions and Constraints trỏ tới Responsibility Rule ở Resource Summary thay vì lặp lại quy tắc.
  - Resource Breakdown Structure (1)
- **7. CREATE RISK REGISTER** (4) — khổ dọc: bảng tiêu đề (RISK REGISTER + Project Title/Date) → đoạn 1 pt (keepNext) → bảng Risk Statement (ID..Response, R-01..R-14) + chú thích Score → bảng Actions (ID, Revised Probability..Comments) + chú thích Revised Score → bảng "Probability and Impact Scale (1-5)" (theo thứ tự sách, user chốt 18/11/2026; giữ trọn một trang). Hai bảng rủi ro lặp dòng tiêu đề mỗi trang, dòng không bị tách.
  - R-01..R-14 (R-13 thắt cổ chai, Owner Minh; R-14 chất lượng sản phẩm thành viên, Owner A), thang 1–5, Score = Probability x Impact cao nhất trong 4 loại. Date Prepared October 8, 2026. Response ghi mã WP chứa hành động phòng ngừa "(in 1.x.y)"; Comments kết thúc bằng "If it occurs: change request (1.7.4); update <tài liệu>". Tất cả Open. Cao nhất là R-02 WebSocket (16).

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
  - Vật tư 104,400,000: 1.2.7 cloud/CDN 3 tháng × 18,500,000 = 55,500,000; 1.5.3 thiết bị test 28,000,000 + k6 17,600,000; 1.6.5 phí store 3,300,000.
    - Bottom-up worksheet xếp cột: **Supplies** = k6 + phí store (20,900,000); **Equipment** = 4 thiết bị (28,000,000); **Other Direct Costs** = hóa đơn cloud 1.2.7 (55,500,000, gồm PR-01..PR-04). Khớp với RBS / Resource Requirements.
  - Contingency reserve 137,904,988 = Σ(P − M) của từng WP. M = most likely; O = M − 10%; P = M + overrun % theo phase: 1.1 15%, 1.2 20%, 1.3 40%, 1.4 40%, 1.5 35%, 1.6 30%, 1.7 15%.
    - Contingency theo phase: 1.1 8,429,891; 1.2 26,802,470; 1.3 24,184,535; 1.4 24,999,699; 1.5 27,811,401; 1.6 15,785,776; 1.7 9,891,216.
    - Contingency 1.2.7 = 12,292,906 ≈ 22 TB CDN ở 550 VND/GB đã gồm VAT (PR-04, R-07).
  - Work package estimates 649,099,036. Warranty (ngoài WBS) 71,822,880 = Hiếu 240 h + A 240 h.
  - Cost baseline 720,921,916. Management reserve 5% = 36,046,096 (cho unknown risks, chỉ Sponsor được giải ngân). Budget 757,000,000.
  - **Quyền PM:** tự duyệt thay đổi tốn **tối đa 2% ngân sách (15,140,000) VÀ tối đa 3 ngày lịch**; vượt một trong hai thì trình Sponsor. Viết thống nhất "up to 2% … and up to 3 days" (Charter, Scope Statement, QMP, Team Charter, WBS Dictionary 1.7, R-04, thang rủi ro).
- **Chất lượng:** tại release candidate (M-06) **không còn lỗi S1 (critical) hoặc S2 (high)** — Charter, Scope Statement, QMP objective 8, WBS Dictionary 1.5 ghi giống nhau. NFR-01: tới 15,000 user với error rate < 1% (và p90 < 200 ms ở NFR-02).
- **Yêu cầu:** FR-02 gồm cả **xóa tài khoản và dữ liệu cá nhân** (App Store bắt buộc; R-05, R-09 dẫn tới FR-02). FR-08 và FR-09 có tiêu chí đo được (50 lượt nghe gần nhất; lệch ≤ 2 s, chat ≤ 1 s, 15,000 user). FR-01, FR-08, NFR-12: "Functional test + UAT"; NFR-06: "Design review + UAT".
- **Milestone:** M-01 Sep 18 · M-02 Oct 1 (PMP được duyệt) · M-03 Oct 21 · 1.3.6 MVP demo Nov 12 · M-04 Nov 20 · 1.4.6 demo Nov 30 · M-05 Dec 7 · 1.5.6 Dec 14 · M-06 RC Dec 18 · M-07 UAT sign-off Jan 4 · M-08 và M-09 Jan 15. Charter: mỗi milestone đạt không trễ quá 14 ngày so với baseline.
- **Công việc quản lý dự án:**
  - 1.7.1 PMP (Sep 16–29) gồm cả kế hoạch rủi ro; các form PMP (QMP, Team Charter, Assumption Log…) đề **Date Prepared: September 29, 2026**. Ô Date Prepared chỉ ghi **một ngày**, không ghi "revised" hay chú thích (user chốt 18/11/2026; check2 kiểm tra); form có số của cost baseline ghi October 8, 2026; "All amounts are in VND" nằm trong chú thích dưới bảng chi phí.
  - 1.7.3: 5 báo cáo cuối tháng vào Sep 11; Oct 15–16; Nov 13 và 16; Dec 10–11; Jan 14–15.
  - 1.7.4 "Monitor risks and manage change requests" (Minh, 64 h): 8 buổi review rủi ro vào Oct 14, Oct 28, Nov 11, Nov 25, Dec 9, Dec 23, Jan 5, Jan 13 (buổi cuối).
  - 1.3.5 (A, Nov 5–11, 40 h) gồm cả 3 kiểm tra sớm của Risk Register: k6 1,000 kết nối WebSocket trên staging (R-02), quét OWASP ZAP baseline (R-08), đo khởi động trên Galaxy A07 / iPhone 11 (R-11); dùng công cụ miễn phí, không đổi giờ/chi phí.

## 6. Risk Register (Book of Forms 2.32, tr. vật lý 139–141) — đã xong
- Cột theo sách: ID | Risk Statement | Owner | Probability | Impact (Scope / Quality / Schedule / Cost) | Score | Response | Revised Probability | Revised Impact (…) | Revised Score | Responsible Party | Actions | Status | Comments. Câu risk statement dạng "If CONDITION, EVENT may occur, leading to EFFECT".
- Khớp với: Charter Overall Project Risk (4 rủi ro, ghi R-01..R-04 ở cuối mỗi ý), Assumption Log A-13/A-14 (R-05/R-06), ACE 1.3.3 + overrun 40% (R-02/R-03), PR-04 (R-07), QMP (R-08, R-11, R-12), Luật BVDLCN (R-09), Team Charter (R-10).
- Owner và Responsible Party là người trong nhóm; ngày review khớp các buổi 1.7.4; tiền lấy từ contingency của phase tương ứng hoặc management reserve.

## 7. Công cụ (`tools/`) và quy trình
- `dlib.py`: các hàm hỗ trợ.
  - `Repl(root, name)(old, new, expect)`: thay chữ có xét nhiều run, kiểm tra số lần thay đúng `expect`.
  - Đọc/ghi ô: `set_cell_text`, `set_cell_lines`, `rows`, `cells`, `cell_text`, `find_row`.
  - Bố cục: `set_grid` (độ rộng cột), `heading_par`, `section_break_par`, `page_break_par`.
- `check2.py <xml> <docx>`: kiểm tra chéo WBS / Dictionary / Resource Requirements / ACE / Resource Summary / Requirements Documentation↔RTM với 4.xml, số TR liên tiếp ở Inter-RTM, score/ID/owner/status của Risk Register, mọi mã 1.x.y / M-xx được nhắc trong docx (trừ RBS, có hệ số riêng) phải có trong 4.xml, cây WBS không có milestone, và các cụm từ cũ không được quay lại (ký hiệu `(A)`, `(SUM)`, "below 2%", "or 3 days", "favourite", "without degradation", "open major"…). Trả mã thoát 1 khi PROBLEMS > 0; báo trùng ID rủi ro.
- `colwidth.py <in> <out> [--dry]`: căn lại độ rộng cột mọi bảng ≥ 3 cột (và QMP) bằng số đo font Arial thật; các bảng cùng form dùng chung độ rộng (RTM, WBS Dictionary, Risk Register trên/dưới, các bảng con của Basis of Estimates); chọn lại số cột gộp của các dòng 2 ô (Project Title | Date Prepared…). Bỏ qua bìa và Team Charter.
- `footers.py <in> <out>`: đặt footer Page x of y cho mọi section trừ trang bìa, đánh số lại từ 1.
- `pagecount.py <docx> <pdf> <out>`: ghi số trang thật của từng section (đọc từ PDF) vào giá trị lưu sẵn của trường SECTIONPAGES. Dùng pymupdf, không có thì dùng pypdf. Nhận cả footer "Page x of y" (4.docx) và "Trang x / y" (4_vi.docx).
- `topdf.ps1`: Word COM → PDF. Gọi bằng `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\topdf.ps1 -Files <docx> -OutDir <dir>` (dùng đường dẫn Windows).
- `montage.py <pdf> <png> <first> <last> <cols> <dpi>`: ghép ảnh nhiều trang để xem nhanh.
- `mspdi.py <xml>`: liệt kê task. `fwdpass.py <xml>`: tính lại lịch, phát hiện trùng việc, tính giờ theo tháng.
- `pdftext.py <pdf> 139-141`: lấy chữ từ giáo trình.
- `mpxjread.py <xml>`: kiểm tra 4.xml đọc được bằng MPXJ. Cần `JAVA_HOME="E:/.jdks/ms-17.0.21.0.2"` (máy hiện tại); pip `mpxj` và `jpype1` đã cài.
- `export_gantt.py <4.xml> <out.pdf> <makepy dir>`: mở 4.xml bằng MS Project qua COM (pywin32), dựng bảng như macro và xuất PDF 1 trang; không lưu .mpp; thư mục makepy phải là thư mục riêng (cache mặc định của máy bị hỏng). **Không dùng `OptionsViewEx`**: đó là tùy chọn chung của MS Project trên máy user (định dạng ngày mặc định là 9 = "Mon 8/17/26").
- `xml_milestones_20261118.py`: script đã chuyển milestone từ nhóm "Milestones" vào cuối từng phase (tự dừng nếu đã làm).
- **Chỉ để tham khảo** (script tự dừng nếu thay đổi đã có): `reference_edit4.py` (clone title bar, header row, label + box của Charter, clone bảng dữ liệu), `build_rr.py` (dựng Risk Register), `edit5.py` (phiên 04/11/2026: thống nhất ngưỡng/mức lỗi, tiêu đề cột như sách + chú thích, cỡ chữ, bìa, Inter-RTM, keepNext cho bảng con, bỏ đoạn trống trước ngắt section).
- **Quy trình:**
  1. Chép 4.docx ra scratchpad.
  2. Sửa bằng python.
  3. Chạy `colwidth.py` (nếu nội dung bảng thay đổi), rồi `footers.py`.
  4. Chạy `check2.py` (phải ra 0).
  5. Chỉ khi user yêu cầu: chạy `topdf.ps1` + `montage.py` để xem bố cục, rồi `pagecount.py` với PDF vừa xuất.
  6. Chép kết quả vào AA/ (4.docx và 4.xml cùng lúc).
- **Lưu ý môi trường Windows:**
  - đặt `PYTHONIOENCODING=utf-8` (tên tiếng Việt);
  - trong Bash không dùng `cat > file` mà không có heredoc (lệnh sẽ treo);
  - lệnh xóa file (`rm`) trong AA/ có thể bị chặn; khi cần bỏ file thì hỏi user hoặc chuyển nó ra scratchpad.

## 8. Việc còn dở
1. Nhúng `schedule(v3).pdf` vào ô placeholder ở mục 3 khi user yêu cầu. Hỏi user muốn chèn dạng ảnh hay object; nếu thêm trang thì chạy lại `footers.py` + `pagecount.py`.
2. ~~Cập nhật `4_vi.docx`~~ — XONG 18/11/2026.
3. ~~Các quyết định đã chốt ngày 18/11/2026~~ — XONG HẾT 18/11/2026 (cả docx và xml) — chi tiết ở cuối `note.txt`, mục "ĐÃ CHỐT (18/11/2026)":
   - bỏ "revised" khỏi Date Prepared (Charter giữ Aug 27 và viết lại 757,000,000 thành pre-approved budget);
   - ~~bỏ 9 milestone khỏi WBS, đánh số lại work package~~ — XONG 18/11/2026 (docx + xml; user cần mở 4.xml bằng MS Project để xuất lại schedule.pdf);
   - Risk Register: thêm R-13 (Minh), R-14 (A); thang 1–5 lên đầu; tách thành 2 bảng; ghi Comments và mã WP cho hành động phòng ngừa;
   - thêm Stakeholder Register + Stakeholder Analysis (thang 1–5) sau Project Charter;
   - thêm câu giải thích mức control account cho WBS Dictionary.
   Làm xong mục nào thì sửa phần mô tả hiện trạng trong BRIEF (mục 4, 5, 6) cho khớp.
