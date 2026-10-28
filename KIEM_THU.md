# Kết quả kiểm thử bản 0.1

Ngày 28/10/2026, chạy trên Linux với Python 3.12. Chưa chạy trực tiếp Windows hoặc trên mạng LAN Windows thực tế.

- 4 nhóm kiểm thử tự động đều đạt: đăng nhập/trang ứng dụng; sửa/xuất/kiểm tra xung đột phiên bản; phân quyền/CSRF/sao lưu; tạo 3 loại quyết định và thay placeholder Word bị chia qua nhiều run.
- Máy chủ Waitress khởi động thật; kiểm tra HTTP cục bộ qua tạo quản trị viên, đăng nhập và truy vấn API có xác thực thành công.
- JavaScript qua kiểm tra cú pháp `node --check`; `pip check` không thấy phụ thuộc hỏng. Chưa chạy kiểm thử thao tác giao diện bằng trình duyệt tự động.
- Với tệp Excel được cung cấp: nhập và đọc được 14 trang tính; tìm được 763 dòng có họ tên (không phải kết luận số người đang làm); bản xuất giữ 24.210 công thức, 5 liên kết ngoài, giá trị ô và các vùng gộp. Ô chuỗi rỗng có thể chuẩn hóa thành ô trống trong bản xuất.
- Không kiểm chứng tính đúng của từng chỉ tiêu nghiệp vụ, không chạy bộ máy tính Excel, không xác minh được các tệp ngoài còn thiếu. Không khẳng định kết quả công thức đã được cập nhật sau chỉnh sửa.
- Không đưa tệp nhân sự, tên thật, số giấy tờ, địa chỉ hay cơ sở dữ liệu thử vào gói phần mềm. Tệp xuất dùng trong kiểm thử ở thư mục tạm đã bị xóa khi kiểm thử hoàn tất.

Trước khi dùng chính thức: kiểm tra trên máy Windows của đơn vị, đối chiếu mẫu quyết định và báo cáo với người phụ trách, hoàn thiện các chức năng chưa có liệt kê trong README.md.
