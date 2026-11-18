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

- Dự án: "Development and Deployment of a Music Streaming Platform": Flutter (một codebase cho iOS/Android/Web), Node.js (Express.js), PostgreSQL, WebSocket cho Live Premiere, object storage + CDN.
- Giáo trình: `A_Project_Managers_Book_of_Forms-2.pdf` (mẫu biểu; trang vật lý = số trang in + 11) và `PMBOK-6-EN.pdf` (tham khảo).
- Git: user làm trên nhánh `hieu` (nhánh `origin/feature/dainp` của Đại đã được so sánh và lấy phần tốt ngày 18/11/2026, không merge).

## 2. File trong AA/
- `4.docx` — tài liệu nộp **duy nhất**, **70 trang** (bìa + mục lục + 68), 18 section (bìa, mục lục, 16 form). Cấu trúc ở mục 4. Chỉ gồm 13 mục cô yêu cầu (Charter, RTM, WBS, WBS Dictionary, Schedule, cost estimation + references, QMP, Team Charter, Resource Requirements, RBS, Risk Register, Stakeholder Register, Stakeholder Analysis); **Assumption Log, Requirements Documentation, Project Scope Statement đã bỏ** (user chốt 18/11/2026), mã A-01..A-14 đã gỡ khỏi mọi form (kể cả Charter). Bản trước khi bỏ còn trong git.
- `4.pdf` — bản PDF nộp kèm: Word xuất PDF, rồi trang Schedule (trang 32) được **đè bản vector của `schedule.pdf`** lên đúng khung ảnh Gantt (pymupdf: redact ảnh cũ rồi `show_pdf_page`) để phóng to vẫn nét. Mỗi khi sửa 4.docx thì xuất lại 4.pdf theo cách này.
- `5.pdf` — bản tách riêng Stakeholder Register + Stakeholder Analysis (5 trang): bìa y hệt 4.pdf, không mục lục, **không số trang** (mọi section trỏ footer rỗng của bìa). Dựng từ 4.docx (giữ section bìa + 2 form), Word xuất PDF. Khi 2 form này đổi trong 4.docx thì xuất lại 5.pdf.
- `4_vi.docx` — bản tiếng Việt để nhóm đọc, cùng bố cục (dòng nào khớp dòng nấy với 4.docx), **không phải bản nộp**, không chạy check2/colwidth trên file này. Footer "Trang x / y". Khi 4.docx đổi nội dung thì hỏi user có cần cập nhật bản này không. Đã đồng bộ với 4.docx ngày 18/11/2026.
- `4.xml` — lịch Microsoft Project (MSPDI), **86 task**, ID liên tiếp. Milestone **không có mã WBS**: trường WBS ghi **M-01..M-09** (đúng như trong 4.docx), không có nhóm "Milestones". Mỗi milestone nằm **ngay sau work package dẫn tới nó** (thứ tự giống lịch cũ, user chốt 18/11/2026): M-01 sau 1.1.7, M-03 sau 1.2.10, M-04 sau 1.3.7, M-05 sau 1.4.7, M-06 sau 1.5.7, **M-07 sau 1.6.2** và **M-02 sau 1.7.1** (hai cái này xen giữa phase), M-08 sau 1.6.10, M-09 sau 1.7.7. OutlineNumber của work package trùng mã WBS, **trừ các gói sau M-07 trong 1.6 và sau M-02 trong 1.7** (lệch 1, ví dụ 1.6.3 có OutlineNumber 1.6.4); bảng in dùng trường WBS nên vẫn đúng.
- `schedule.pdf` — Gantt hiện hành, xuất từ 4.xml: **1 trang 34 x 22 in ngang** (ANSI D); bảng WBS | Task Name | Duration | Start | Finish | Predecessors (trường WBS Predecessors) | Resource Names, không cột ID; thang thời gian tháng/tuần; kiểu thanh, legend, gridline mặc định; ngày dạng "Mon 8/17/26". Đã nhúng vào 4.docx dạng ảnh PNG 200 dpi rộng 6.8 in, **trang dọc** (user yêu cầu giữ dọc); trong 4.pdf ảnh được thay bằng bản vector.
- `Setup_Schedule_View.bas` — macro MS Project dựng đúng bố cục của `schedule.pdf` (độ rộng cột 9/60/10/15/15/62/15, đặt lại mã M-01..M-09 theo tên milestone, thang tháng/tuần, 1 trang 34x22). User mở 4.xml → View → Macros… → SetupScheduleView → Run, rồi File → Export → PDF. File phải giữ **CRLF**. Phím F của user là phím âm lượng: chỉ hướng dẫn thao tác bằng menu.
- Thư mục `tools/` **đã bị user xoá** (commit 78bf038). Script cũ lấy lại để chạy ở **scratchpad** (không tạo lại `tools/` trong AA/): `git show 78bf038~1:tools/<tên file> > <scratchpad>/tools/<tên file>`.

## 3. Quy tắc bắt buộc
- **Git: chỉ đọc** (`git log/show/diff/status`). Không dùng lệnh cập nhật (commit, merge, checkout, pull, push, reset…) — user tự commit.
- Gặp điểm chưa rõ thì **hỏi user trước** (AskUserQuestion, tiếng Việt, có phương án đề xuất).
- Chỉ làm phần được yêu cầu; không thêm biểu mẫu mới. Nội dung lấy từ dữ liệu đã có và phải nhất quán với nó. **Không đụng Project Charter** khi sửa các form khác.
- Cách viết (góp ý giảng viên + user 18/11/2026):
  - mọi form ngoài Charter viết cho kỹ sư trong nhóm: **khách quan, kỹ thuật, đo được, không tâng bốc**. Không dùng tính từ mơ hồ (stable, quickly, relevant, minor, major, main, large/small, "in real time", "as capacity allows", "if the schedule slips", "without degradation"…) → thay bằng ngưỡng/mã cụ thể (2% ngân sách / 3 ngày; R-03: gói trễ > 3 ngày → change request; R-02 score 16; S1–S4 định nghĩa trong QMP);
  - tên task/WBS bắt đầu bằng động từ, ngắn gọn (≤ khoảng 52 ký tự);
  - unit test do developer viết, nằm trong work package phát triển;
  - UAT do Customer làm, ngoài scope → mọi yêu cầu Must/Should có thêm một phép kiểm do nhóm làm (Functional test, Design review…).
- **Không dùng dấu `;`** ở bất kỳ đâu (user yêu cầu 18/11/2026): nối mệnh đề bằng dấu chấm, liệt kê bằng dấu phẩy.
- **Tiêu đề cột ghi đúng như sách**: không dùng `Hours (A)`, `Total (C = A x B)`, `(SUM)`…; công thức ghi ở chú thích dưới bảng. Giữ đúng các cột của form.
- Mọi con số, ngày, ID khớp chéo giữa các bảng và với 4.xml. Sau mỗi lần sửa chạy `check2.py 4.xml 4.docx` → phải `PROBLEMS: 0`.
- **Đồng bộ 4.xml, 4.docx (và 4_vi.docx, macro) cùng lúc**: sửa gì động tới lịch (mã, tên task, ngày, người, giờ, Notes) thì sửa cả hai phía, chạy `check2.py` + `mpxjread.py`, xuất lại `schedule.pdf`.
- Ô Date Prepared chỉ ghi **một ngày** (không "revised", không chú thích). Charter: August 27, 2026; form PMP (QMP, Team Charter, Assumption Log, Schedule…): September 29, 2026; form có số cost baseline: October 8, 2026.
- Định dạng docx:
  - (user yêu cầu 18/11/2026) font **Arial**, **mọi chữ 12 pt** (bảng, cây WBS/RBS, Project Title/Date Prepared, chú thích, footer); **tiêu đề 13 pt**: title bar (nền tối, chữ trắng đậm), tiêu đề nhóm, "Basis of Estimates and References"; mọi trang **dọc** Letter. Nội dung dài thì **tóm tắt/cắt bớt** (không giảm cỡ chữ), giữ đủ dòng, ID, số liệu; mục tiêu tổng 70–80 trang;
  - **màu**: chỉ có màu đường kẻ bảng và nền trắng mặc định — không tô nền ô/đoạn (`w:shd`), không chữ màu; title bar là chữ đậm trên nền trắng (user yêu cầu 18/11/2026, cho tài liệu bớt "kiểu AI"). **Ngoại lệ duy nhất**: ô nhóm chưa chắc đúng/sai tô **vàng nhạt FFF2CC** (yêu cầu của cô). Chỉ tô từng ô dễ sai, không tô cả cột. Hiện có 23 ô: Stakeholder Analysis (End users Interest, Trần Văn A Influence, Ministry Influence) + 3 ô Classification tương ứng ở Register; Risk R-02 Probability/Score, R-09 Cost/Score, R-07 Revised Probability/Revised Score (Accept mà vẫn hạ xác suất); ACE Range 1.3/1.4/1.5 total, Confidence 1.2.7, ô Basis của management reserve 5%, ô giả định bảo hành 40h/tháng; CEW Pessimistic 1.3/1.4/1.5; bảng 1.2.7 Cloud: dòng 2 x API server và dòng CDN traffic 5,000 GB;
  - **chú thích dưới bảng viết rất gọn** (chỉ đơn vị tiền + công thức chính) hoặc bỏ;
  - bảng không vừa 12 pt thì **tách thành 2–3 bảng liên tiếp** giữ thứ tự cột của form, lặp cột ID (Stakeholder Register, RTM, Inter-RTM (TR nhóm theo dòng BO-n), ACE, Bottom-up 2 bảng: ID..Supplies | ID, Equipment..Estimate, Risk Register 3 bảng: ID..Score | ID, Response, Revised..Responsible Party | ID, Actions, Status, Comments); cột ID lặp lại có **cùng độ rộng vừa khít** ở mọi phần của bảng;
  - **mọi tiêu đề cột viết ngang** (không còn btLr); cột phải đủ rộng cho từ dài nhất của tiêu đề (không gãy giữa từ);
  - **dòng tiêu đề bảng không lặp lại khi sang trang** (không `tblHeader`), chỉ xuất hiện 1 lần;
  - **ACE, Bottom-up, WBS Dictionary ghi tiền theo nghìn VND** (làm tròn; Rate vẫn VND/giờ; tổng tính từ số chưa làm tròn; ghi ở chú thích); các form khác vẫn VND đầy đủ. WBS Dictionary/ACE dùng tên ngắn (Minh, Đại, Hiếu, A, Thu); Resource Requirements cột Resource chỉ ghi họ tên, vai trò + kỹ năng ghi một lần ở dòng "Team resources (people)", cột Assumptions chỉ còn "Available …";
  - sửa nội dung bảng xong chạy bản colwidth cho 12 pt (mọi từ/số ≤ 200 pt không bị gãy, báo `NOT FITTING: 0`; tên người ≤ 2 dòng); Team Charter không chạy colwidth;
  - **mục lục** (section riêng sau bìa): trường TOC `\o "1-2" \u`, cấp 1 = tiêu đề nhóm đánh số La Mã **I. … VII.** (không dùng "I.1"), cấp 2 = tên form (title bar đầu tiên của form, outline level 2), **không đánh số form**; cập nhật trường bằng Word trước khi lưu;
  - mỗi form vẫn là một section; footer căn giữa **chỉ có số trang** `{PAGE}` (user yêu cầu, không "Page x of y"), **đánh số liên tục** (bìa tính là trang 1, không hiện số);
  - dòng tiêu đề bảng giữ liền dòng sau (keep with next), không để dòng tiêu đề lẻ cuối trang; bỏ title bar lặp giữa trang (Charter, Scope Statement, WBS Dictionary) và các ngắt trang ép buộc giữa các phần của một form.
- Bìa: font **Times New Roman**, tiêu đề 13 pt (ASSIGNMENT REPORT, tên dự án), dòng khác 12 pt; POSTS AND TELECOMMUNICATIONS INSTITUTE OF TECHNOLOGY / FACULTY OF INFORMATION TECHNOLOGY 1 / logo PTIT / ASSIGNMENT REPORT / tên dự án / COURSE: Software Project Management / bảng Class E23CNPM02, Group 04, Students (4 người + mã SV), Instructor / Hanoi – 2026; khung và logo là ảnh rId26/rId27; không có footer; không đổi trừ khi user yêu cầu.
- Sửa docx bằng python-docx + lxml (`dlib.py`), làm trên bản sao ở scratchpad rồi mới chép vào AA/ (giữ bản cũ ở scratchpad).

## 4. Cấu trúc 4.docx
Số trang (Word 18/11/2026, sau khi bỏ 3 form): bìa 1, mục lục 1, Charter 7, Stakeholder Register 3, Stakeholder Analysis 1, RTM + Inter-RTM 6, WBS 2, WBS Dictionary + Resource Summary 10, Schedule 1, ACE 7, Cost Estimating Worksheet 2, Bottom-up 4, References 3, QMP 4, Team Charter 2, Resource Requirements 6, RBS 1, Risk Register 11 = 72; sau khi rút gọn chú thích: Resource Requirements 5, Risk Register 10 = **70** (mục lục: Schedule 32, ACE 33, REFERENCES 46, QMP 49, Team Charter 53, RBS 60, Risk 61). Mục "Basis of Estimates and References" đã đổi tên thành **REFERENCES** (tiêu đề + mục lục; nội dung giữ nguyên). Các số trong ngoặc ở dưới là của bản 7 pt cũ; các mô tả form đã bỏ chỉ còn để tham khảo.
Thứ tự theo Book of Forms: 1.1 Charter → 1.2 Assumption Log → 1.3 Stakeholder Register → 1.4 Stakeholder Analysis → 2.6 Requirements Documentation → 2.7 RTM → 2.8 Scope Statement → 2.9 WBS → 2.10 WBS Dictionary → 2.18 Schedule → 2.20/2.21 chi phí → 2.23 QMP → 2.27 Team Charter → 2.28 Resource Requirements → 2.29 RBS → 2.32 Risk Register (thang Probability/Impact ở cuối). Dời form thì dời nguyên section.
- Cover (1)
- **1. CREATE PROJECT CHARTER**: Project Charter (4) — 757,000,000 VND là pre-approved budget (các khoản "about ...", cộng đúng 757,000,000), không nhắc "cost baseline"; Assumption Log A-01..A-14 (1); Stakeholder Register (1) + Stakeholder Analysis (1) — 12 stakeholder, email giả định dạng PTIT (user chọn), Interest/Influence 1–5, Attitude Supportive/Neutral/Resistant, Classification = Internal/External + mức suy từ điểm (check2 kiểm tra); Requirements Documentation FR-01..FR-14, NFR-01..NFR-12 (2).
- **2. CREATE WBS**: RTM + Inter-RTM BO-1..BO-6 / TR-01..TR-18 (3); Scope Statement (2); WBS (2) — chỉ phase + 55 work package, milestone không thuộc WBS; WBS Dictionary 1.1–1.7 mức control account + Resource Summary (8) — ô Assumptions and Constraints trỏ tới Responsibility Rule ở Resource Summary.
- **3. CREATE PROJECT SCHEDULE** (1): **chỉ có tiêu đề nhóm "III. CREATE PROJECT SCHEDULE" và ảnh Gantt** (từ `schedule.pdf`), không title bar, không Project Title/Date Prepared, không câu mô tả (user yêu cầu 18/11/2026).
- **4. CREATE COST ESTIMATION**: Activity Cost Estimates (5); Cost Estimating Worksheet (1); Bottom-up (2); Basis of Estimates and References (2).
- **5. QMP** (3): ISO/IEC 25010, 12 quality objectives; định nghĩa S1–S4 và thời hạn sửa (S1 1 ngày, S2 3 ngày, S3 trước M-06, S4 gói sửa lỗi sau / warranty).
- **6. RESOURCE MANAGEMENT**: Team Charter (2, cột chữ ký 50/50); Resource Requirements (3) — người + PR-01..PR-12, "one task at a time" và kỹ năng theo vai trò ghi **một lần** ở dòng "Team resources (people)"; RBS (1).
- **7. CREATE RISK REGISTER** (4): bảng tiêu đề → bảng Risk Statement (R-01..R-14) + chú thích Score → bảng Actions + chú thích Revised Score → bảng "Probability and Impact Scale (1-5)". Hai bảng rủi ro lặp header, dòng không bị tách. Response ghi WP chứa hành động phòng ngừa "(in 1.x.y)"; Comments kết thúc "If it occurs: change request (1.7.4); update <tài liệu>". Tất cả Open; cao nhất R-02 WebSocket (16); R-13 thắt cổ chai (Owner Minh), R-14 chất lượng công việc (Owner A).

Tìm bảng bằng chữ ở ô đầu tiên (ví dụ `"ACTIVITY COST ESTIMATES"`), không dùng chỉ số.

## 5. Dữ liệu gốc (không được lệch)
- **Thời gian:** Aug 17, 2026 – Jan 15, 2027 = 108 ngày làm việc (T2–T6, 8–12h, 13–17h). Nghỉ Sep 1–2 và Jan 1; làm bù thứ Bảy Sep 5. Tháng dự án: M1 Aug 17–Sep 16, M2 Sep 17–Oct 16, M3 Oct 17–Nov 16, M4 Nov 17–Dec 16, M5 Dec 17–Jan 15.
- **Nguồn lực:** mỗi WP đúng 1 người; mỗi người 1 việc một lúc; WP 8–80 h; ≤ 176 h/người/tháng. Giờ theo tháng: Minh 152/152/40/64/104; Đại 72/128/160/96/48; Hiếu 64/160/168/48/104; A 88/96/80/88/96; Thu 120/104/144/56/32; cả nhóm 496/640/592/352/384 = 2,464 h (trên 4,320 h khả dụng).
- **Chi phí (VND):** labor 406,794,048; vật tư 104,400,000 (1.2.7 cloud/CDN 3 × 18,500,000 = 55,500,000; 1.5.3 thiết bị 28,000,000 + k6 17,600,000; 1.6.5 phí store 3,300,000). Bottom-up: Supplies 20,900,000, Equipment 28,000,000, Other Direct Costs 55,500,000. Contingency 137,904,988 = Σ(P − M), overrun theo phase 1.1 15%, 1.2 20%, 1.3 40%, 1.4 40%, 1.5 35%, 1.6 30%, 1.7 15% (theo phase: 8,429,891 / 26,802,470 / 24,184,535 / 24,999,699 / 27,811,401 / 15,785,776 / 9,891,216; riêng 1.2.7 = 12,292,906 ≈ 22 TB CDN). WP estimates 649,099,036; warranty 71,822,880 (Hiếu 240 h + A 240 h, ngoài WBS); cost baseline 720,921,916; management reserve 5% = 36,046,096 (chỉ Sponsor giải ngân); budget 757,000,000.
- **Quyền PM:** tự duyệt thay đổi **tối đa 2% ngân sách (15,140,000) VÀ tối đa 3 ngày**; vượt một trong hai thì trình Sponsor. Viết thống nhất "up to 2% … and up to 3 days".
- **Chất lượng:** tại release candidate (M-06) không còn lỗi S1/S2. NFR-01: tới 15,000 user, error rate < 1%; NFR-02: p90 < 200 ms; backend coverage ≥ 70%; UAT ≥ 4.2/5; crash-free ≥ 99%.
- **Yêu cầu:** FR-02 gồm xóa tài khoản và dữ liệu cá nhân; FR-09: phát lệch ≤ 2 s so với đồng hồ công chiếu, chat ≤ 1 s, 15,000 user.
- **Milestone:** M-01 Sep 18 (SRS/RTM sign-off) · M-02 Oct 1 (PMP approved) · M-03 Oct 21 (design approved) · M-04 Nov 20 (MVP) · M-05 Dec 7 (all functions) · M-06 Dec 18 (RC) · M-07 Jan 4 (UAT sign-off) · M-08 Jan 15 (apps handed over) · M-09 Jan 15 (project closed). Charter: mỗi milestone không trễ quá 14 ngày.
- **Quản lý dự án:** 1.7.1 PMP Sep 16–29; 1.7.3 năm báo cáo cuối tháng Sep 11; Oct 15–16; Nov 13 và 16; Dec 10–11; Jan 14–15; 1.7.4 tám buổi review rủi ro Oct 14, Oct 28, Nov 11, Nov 25, Dec 9, Dec 23, Jan 5, Jan 13. 1.3.5 (A, Nov 5–11) gồm 3 kiểm tra sớm: k6 1,000 kết nối WebSocket (R-02), OWASP ZAP baseline (R-08), đo khởi động Galaxy A07 / iPhone 11 (R-11).

## 6. Công cụ (lấy từ git `78bf038~1:tools/`, chạy ở scratchpad)
- `dlib.py`: `Repl(root, name)(old, new, expect)` thay chữ xuyên run và kiểm số lần thay; `rows`, `cells`, `cell_text`, `set_cell_text`, `set_cell_lines`, `find_row`, `set_grid`, `heading_par`, `section_break_par`.
- `check2.py <xml> <docx>`: kiểm tra chéo docx ↔ 4.xml (WBS, Dictionary, Resource Requirements, ACE, Resource Summary, RTM, Inter-RTM, Risk Register, Stakeholder, Date Prepared một ngày, mã 1.x.y / M-xx phải có trong 4.xml, cụm từ cấm); mã thoát 1 khi có lỗi.
- `colwidth.py <in> <out> [--dry]`, `footers.py <in> <out>`, `pagecount.py <docx> <pdf> <out>` (bản trong git chỉ đọc footer "Page x of y"; với 4_vi.docx phải nhận thêm "Trang x / y"), `topdf.ps1` (Word COM → PDF), `mspdi.py <xml>` (liệt kê task), `mpxjread.py <xml>`, `fwdpass.py`, `montage.py`, `pdftext.py`.
- Xuất `schedule.pdf` qua COM (python + pywin32): `gencache.EnsureDispatch("MSProject.Application")` → `FileOpenEx(Name=4.xml, ReadOnly=True, FormatID="MSProject.XML")` → đặt lại mã M-xx theo tên milestone → `ViewApply("Gantt Chart")`, `OutlineShowAllTasks`, `TableEditEx` như macro, `TableApply` → `TimescaleEdit(TierCount=2, MajorUnits=2, MajorAlign=1, MinorUnits=3, MinorAlign=1)` → `FilePageSetupPage(Portrait=False, PagesTall=1, PagesWide=1, PaperSize=25)` → `FilePageSetupView(AllSheetColumns=True, BestPageFitTimescale=True)` → `DocumentExport(Filename, FileType=0)` → `FileCloseEx(0)`, `Quit(0)`. Xong thì render PDF ra ảnh (pymupdf) để tự kiểm tra 1 trang, không cắt chữ.
- **Lưu ý môi trường (máy hiện tại):**
  - đặt `PYTHONIOENCODING=utf-8`; `JAVA_HOME="E:/.jdks/ms-17.0.21.0.2"` cho mpxjread;
  - cache makepy mặc định của MS Project bị hỏng → đặt `win32com.__gen_path__` và `win32com.gen_py.__path__` sang một thư mục riêng ở scratchpad trước khi `EnsureDispatch`;
  - **không gọi `OptionsViewEx`** (đổi tùy chọn chung của MS Project trên máy user; định dạng ngày của user là 9 = "Mon 8/17/26");
  - `DocumentExport` bỏ qua FromDate/ToDate (bản in tự chừa ~2 tuần trước Aug 17); `FilePageSetupLegend TextWidth` không có tác dụng (ô legend ~2 in);
  - `topdf.ps1` gọi `Word.Quit()` → **không chạy khi user đang mở Word** (có file `~$...docx`); khi Word đang mở 4.docx thì không ghi đè được, nhờ user đóng (không lưu);
  - không kill WINWORD; trong Bash không dùng `cat > file` mà không có heredoc; xóa file trong AA/ thì chuyển bản sao ra scratchpad trước.

## 7. Việc còn dở
1. `4_vi.docx` **chưa** chuyển sang bố cục 12 pt / mục lục / bảng tách / nghìn VND (vẫn là bản 7 pt khớp với 4.docx cũ) — chờ user quyết định có cập nhật không.
2. Các bước dựng bản 12 pt (18/11/2026) chạy ở scratchpad, theo thứ tự: sửa nội dung + tách bảng → định dạng 12/13 pt, mục lục, footer liên tục, keep-with-next → chèn ảnh Gantt → colwidth 12 pt + chiều cao dòng tiêu đề dọc → Word cập nhật trường TOC/footer, lưu docx, xuất PDF → ghi NUMPAGES lưu sẵn + đè Gantt vector vào PDF. check2 đã được chỉnh theo bố cục mới (tách bảng, nghìn VND, Risk 3 bảng) và báo `PROBLEMS: 0`; bản `check2.py` trong git **không** chạy đúng trên bố cục mới (bản đã chỉnh không được lưu lại).
3. Lần sửa 18/11/2026 (bỏ 3 form, bỏ `;`, bỏ màu nền, tiêu đề ngang, không lặp header, tách Bottom-up, đổi tên REFERENCES, tạo 5.pdf) chỉ kiểm tra bằng so sánh tập số/ID từng section trước–sau (không mất số/ID nào). Độ rộng cột tính bằng `colwidth.py` sửa cho 12 pt (ngưỡng từ không gãy 200 pt); WBS Dictionary dùng lề ô 30 dxa để vừa tiêu đề ngang. Ảnh Gantt (MS Project) và logo bìa vẫn có màu — chưa đổi.
