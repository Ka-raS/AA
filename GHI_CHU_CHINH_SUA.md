# Các chỉnh sửa ngày 18 tháng 11 năm 2026

Đã đối chiếu `note.txt` tại GitHub Ka-raS/AA với bản trong thư mục; nội dung trùng nhau. Chỉnh sửa được lưu trong tài liệu cục bộ, chưa đưa lên GitHub.

## Những việc đã làm

- `4.docx`: xử lý ghi chú; thêm phần 8 Stakeholder Register; rút các câu quy tắc và kỹ năng bị lặp.
- `Stakeholder_Register.docx`: biểu mẫu riêng để nộp, 13 bản ghi, đúng 6 cột của Book of Forms 1.3. Kênh liên hệ được ghi theo dữ liệu có sẵn; không đặt email/số điện thoại giả. PM cần bổ sung liên hệ cụ thể khi được cung cấp.
- `4_Tieng_Viet.docx`: bản đọc chung rút gọn; gồm nội dung các nhóm biểu mẫu, 26 yêu cầu, 18 yêu cầu kỹ thuật, 55 WP, số liệu chi phí/nguồn lực và 14 rủi ro. Giữ mã để tra bản nộp, không lặp toàn bộ bảng số liệu ở nhiều form.
- Phần chia người và ôn vấn đáp để sau theo yêu cầu mới nhất.

## Kết luận cho từng note

**Date Prepared.** Bỏ phần “revised …”; dùng một ngày là ngày chuẩn bị/cập nhật mới nhất đã được ghi trong nguồn. Không đổi ngày ký/phê duyệt và không tự đánh dấu các sự kiện tương lai là đã xảy ra.

**Milestone trong WBS và lịch.** Bỏ 9 dòng sự kiện 0 giờ khỏi cây WBS, giữ 55 WP. Giữ các sự kiện trong lịch và giữ nguyên ID/UID/mã tham chiếu để không phá liên kết, RTM và các bảng liên quan. ID là định danh bản ghi lịch, không có nghĩa mọi bản ghi đều là công việc có công sức. Demo/review có giờ công vẫn là WP, dù ngày hoàn thành cũng được Charter dùng làm mốc quản lý. Không đánh lại mã chỉ để lấp chỗ trống.

**WBS Dictionary theo work package hay major deliverable.** Book of Forms 2.10, trang in 52–53 (PDF trang 63–64), cho phép thông tin chi tiết từng WP hoặc thông tin tóm tắt ở cấp control account. Mẫu trống ở trang in 55 (PDF trang 66) đặt nhãn “Work Package Name” vì minh họa trường hợp một WP. Bài nhóm dùng 1.1–1.7 làm control account theo giai đoạn để tổng hợp phạm vi, lịch và chi phí, mỗi bảng liệt kê các WP 1.x.y. Đã đổi nhãn thành **Control Account Name** và ghi rõ việc điều chỉnh mẫu. Không còn gọi 1.1–1.7 là từng work package. Nếu yêu cầu chấm bài bắt buộc mỗi form cho một WP thì cần tách từng 1.x.y; tài liệu hiện tại là bản tóm tắt có chủ đích, có cơ sở từ hướng dẫn của sách.

**Risk Register.** Đưa thang Probability/Impact lên đầu; còn hai bảng dữ liệu xuyên trang là Risk Statements và Actions, lặp header khi qua trang. Thêm R-13 về nghẽn Technical Lead và R-14 về thiếu kỹ năng/chất lượng công việc. Mỗi rủi ro chỉ ra tài liệu phải cập nhật khi trigger xảy ra. Baseline chỉ đổi sau phê duyệt; không tự cắt phạm vi hoặc chuyển chi phí chỉ vì đã ghi response.

**Response có phải tạo WP riêng.** Không cần tạo một WP mới cho từng response. Việc phòng ngừa đã có (unit test/review trong WP phát triển, kiểm tra sớm 1.3.5, backup 1.2.7, review 1.7.5) phải dẫn đúng WP và nằm trong ước lượng của nó. Việc mới vượt phạm vi/giờ/chi phí đã duyệt cần change request; sau duyệt cập nhật WBS/Dictionary, lịch, nguồn lực và cost baseline khi bị ảnh hưởng. Reserve là nguồn tiền, không thay thế công việc và năng lực trong lịch.

**Nhân sự.** R-10 xử lý vắng mặt; R-13 xử lý nghẽn người ra quyết định/review; R-14 xử lý kỹ năng không khớp công việc. Trước đổi người phải xét kỹ năng, lịch từng người, dependency, không làm chồng việc và giới hạn 176 h/tháng. Không dùng tổng 1,856 h công suất trống như thể mọi giờ đều có thể thay thế lẫn nhau.

## Dữ liệu được giữ

108 ngày làm việc; 2,464 giờ WBS; 406,794,048 VND nhân công; 137,904,988 VND contingency; 720,921,916 VND cost baseline; ngân sách 757,000,000 VND. Lịch `4.xml` và Gantt `schedule.pdf` không đổi. Quỹ dự phòng là khoản dùng chung theo baseline, không cộng thêm một lần nữa cho R-13/R-14. Các response mở rộng cần được duyệt trong quỹ còn lại và năng lực khả dụng.

## Kiểm tra

Chạy `tools/check2.py 4.xml 4.docx` và kiểm tra các mã, người, giờ, rate, tổng chi phí, truy vết yêu cầu và điểm rủi ro. Script được cập nhật để nhận hai bảng rủi ro xuyên trang và trả mã lỗi khác 0 nếu phát hiện vấn đề. Bản gốc được lưu tại `tmp/revision_20261118/4.original.docx` để có thể đối chiếu/khôi phục.

Tài liệu gốc vẫn giữ vị trí chèn Gantt trong phần 3; file Gantt riêng do nhóm đã xuất là `schedule.pdf`. Chưa thay hình hoặc tạo lại lịch ngoài phạm vi các note.
