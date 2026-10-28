# Kết quả kiểm thử bản 0.1

Ngày 28/10/2026, chạy trên Linux với Python 3.12. Chưa chạy trực tiếp Windows hoặc trên mạng LAN Windows thực tế.

- 4 nhóm kiểm thử tự động đều đạt: đăng nhập/trang ứng dụng; sửa/xuất/kiểm tra xung đột phiên bản; phân quyền/CSRF/sao lưu; tạo 3 loại quyết định và thay placeholder Word bị chia qua nhiều run.
- Máy chủ Waitress khởi động thật; kiểm tra HTTP cục bộ qua tạo quản trị viên, đăng nhập và truy vấn API có xác thực thành công.
- JavaScript qua kiểm tra cú pháp `node --check`; `pip check` không thấy phụ thuộc hỏng. Chưa chạy kiểm thử thao tác giao diện bằng trình duyệt tự động.
- Với tệp Excel được cung cấp: nhập và đọc được 14 trang tính; tìm được 763 dòng có họ tên (không phải kết luận số người đang làm); bản xuất giữ 24.210 công thức, 5 liên kết ngoài, giá trị ô và các vùng gộp. Ô chuỗi rỗng có thể chuẩn hóa thành ô trống trong bản xuất.
- Không kiểm chứng tính đúng của từng chỉ tiêu nghiệp vụ, không chạy bộ máy tính Excel, không xác minh được các tệp ngoài còn thiếu. Không khẳng định kết quả công thức đã được cập nhật sau chỉnh sửa.
- Không đưa tệp nhân sự, tên thật, số giấy tờ, địa chỉ hay cơ sở dữ liệu thử vào gói phần mềm. Tệp xuất dùng trong kiểm thử ở thư mục tạm đã bị xóa khi kiểm thử hoàn tất.

Trước khi dùng chính thức: kiểm tra trên máy Windows của đơn vị, đối chiếu mẫu quyết định và báo cáo với người phụ trách, hoàn thiện các chức năng chưa có liệt kê trong README.md.

## Bản 0.2 — hồ sơ, tự phân nhóm, quyết định và báo cáo

- Bộ kiểm thử có thêm nhóm về tuổi ở đúng sinh nhật và ranh giới nhóm, ngày sinh 29/2, tự phân nhóm người mới, điều động tương lai, báo cáo lịch sử, chấm dứt, xung đột dự thảo, quy tắc có ngày hiệu lực, cập nhật trình độ theo lịch sử, nhập nguồn cần rà soát và chống trùng, quyền chỉ xem, cập nhật hàng loạt nguyên giao dịch và phục hồi sao lưu có hồ sơ chuẩn hóa/tài khoản cũ.
- Kiểm thử Chromium thực tế trên Linux: tạo quản trị; tạo hai vị trí/quy tắc; thêm người mới tự vào nhóm công nghệ/CNKT bậc 3; báo cáo tháng 10; tải Word điều động; áp dụng chuyển sang quản lý từ tháng 11; báo cáo tháng 11 tự đổi nhóm và tuổi; báo cáo tháng 10 giữ nhóm cũ; tải báo cáo Excel; không có lỗi JavaScript.
- Chỉ dùng dữ liệu giả trong kiểm thử tự động và trình duyệt. Chưa kiểm thử trực tiếp Windows/LAN Windows, chưa đối chiếu đầy đủ mẫu báo cáo chính thức và quy tắc nghiệp vụ đơn vị. Chưa đồng bộ ngược hồ sơ chuẩn hóa vào Excel gốc.
- Kiểm tra nhập nguồn thật trong thư mục tạm: xem trước 765 dòng có nội dung ở cột họ tên (bao gồm 2 ô công thức tổng); 694 dòng đủ điều kiện vào trạng thái Cần rà soát, 71 dòng bị bỏ qua do mã thiếu/trùng hoặc tên không hợp lệ. Không tự chọn một trong các dòng trùng mã. Nhập lại không sinh hồ sơ mới; báo cáo không đếm các hồ sơ chưa xác nhận. Đây không phải kết luận có 694 lao động đang làm.

## Bản 0.2.1 — báo cáo số 0 và tình trạng dữ liệu

- 17 kiểm thử API/nghiệp vụ đều đạt, gồm chẩn đoán chưa có hồ sơ dù đã có Excel, toàn bộ hồ sơ chờ xác nhận, kỳ trước ngày hiệu lực và tệp Excel xuất có ghi tình trạng dữ liệu.
- Trình duyệt Chromium: tái hiện Excel nguồn chưa có hồ sơ → nhập hồ sơ cần rà soát → báo cáo 0 với thông báo rõ → dùng nút tới hồ sơ, xác nhận một người → báo cáo thành 1. Kiểm tra kỳ trước ngày hiệu lực, nút đổi sang kỳ có dữ liệu, lỗi API giả lập hiển thị ngay tại báo cáo và tổng hợp lại thành công. Không có lỗi JavaScript.
- Luồng trước đây gồm thêm người, điều động tương lai, báo cáo trước/sau quyết định và xuất Excel vẫn qua kiểm thử trình duyệt. Chưa truy cập máy Windows hay cơ sở dữ liệu của người dùng; không khẳng định đã kiểm tra trực tiếp bản cài tại đơn vị.
