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

## Bản 0.3.0 — bộ lọc và trạng thái mã cuối kỳ

- Kiểm thử API: đọc GO/HE có tiêu đề, mã có/trống/0/lỗi, công thức thiếu cache, hai cột bất đồng; nhập trạng thái tự động; giữ người đã nghỉ và người thiếu/trùng mã bằng mã lưu trữ nội bộ; kết hợp bộ lọc hồ sơ; lọc nguồn và xuất XLSX; chuyển trang không lặp/bỏ sót; cập nhật trạng thái nguồn cho hồ sơ cũ có kiểm tra phiên bản và giữ lịch sử.
- Trên tệp thật, trong thư mục tạm: đọc được 692 dòng còn mã và 70 dòng hết mã; nhập giữ đủ 762 hồ sơ. Trong nhóm đã nghỉ có 38 dòng thiếu mã nhân sự được giữ bằng mã lưu trữ. Báo cáo tự tổng hợp đang làm đếm 692 dòng, không đếm 70 dòng đã nghỉ. Nhập lại không tạo thêm hồ sơ. Có dòng trùng mã nhân sự nguồn nên chưa kết luận số cá nhân duy nhất; không tự gộp/xóa.
- Kiểm thử Chromium: lọc người đã nghỉ ngay trên Excel nguồn; tải và đọc danh sách XLSX theo bộ lọc; lọc kết hợp nguồn/trạng thái/đơn vị; nhập hồ sơ theo mã cuối kỳ; lọc hồ sơ theo trạng thái/đơn vị/nhóm/trình độ/tuổi; chọn chỉ các hồ sơ đang lọc; cập nhật trạng thái từ nguồn cho hồ sơ đã nhập; kiểm tra báo cáo lịch sử 2 → 3 → 2 sau các mốc trạng thái. Không có lỗi JavaScript.
- Giữ nguyên tệp nguồn, không phát tán dữ liệu nhân sự trong mã nguồn. Ngày chốt nguồn không được coi là ngày nghỉ thực tế; chưa suy diễn lịch sử trước ngày có dữ liệu. Chưa kiểm thử trực tiếp Windows.

## Bản 0.3.1 — phục hồi pip và trình khởi động

- 32 kiểm thử đều đạt, gồm bỏ qua pip khi thư viện đã sẵn sàng; phát hiện pip hỏng và tạo lại một lần; giữ môi trường cũ; khôi phục đường dẫn cũ nếu venv/ensurepip thất bại; báo lỗi cài gói mà không lặp vô hạn; chặn tạo lại nếu EMGA_DATA_DIR đặt trong .venv; kiểm tra phiên bản/import thật.
- Chạy `bootstrap.py --check` trên môi trường ứng dụng hiện tại: thành công, không gọi pip hay tải lại thư viện.
- Kiểm tra tích hợp trong thư mục tạm có dấu/khoảng trắng: tạo venv riêng, làm thiếu rich.themes để tái hiện lỗi pip, chạy prepare để giữ môi trường lỗi và tạo mới từ ensurepip, cài lại phụ thuộc chốt phiên bản, kiểm tra import và tệp dữ liệu mẫu không đổi.
- Chưa chạy trực tiếp batch script hoặc kiểm chứng sửa lỗi trên chính máy Windows của người dùng. Không thay đổi luồng nghiệp vụ báo cáo/nhân sự trong bản này.

## Bản 0.3.2 — lỗi gốc và vị trí thư viện Windows ngắn

- Kiểm thử bổ sung: lỗi pip được hiện trên console và ghi log; lựa chọn cache Windows ngắn, ổn định qua các thư mục giải nén và tách cache khi phụ thuộc đổi; chế độ chẩn đoán không cài/sửa môi trường; kiểm tra cả môi trường cache tạo thất bại gần nhất.
- Giữ kết quả thử trước là bằng chứng Linux, không coi là xác nhận đã sửa được máy Windows của người dùng. Python người dùng báo 3.12.5 nằm trong dải được hỗ trợ; chưa có lỗi gốc để kết luận Python cài hỏng hoặc đường dẫn dài là nguyên nhân.
- Không đổi dữ liệu/tài khoản, không tự gửi log đi. Nhật ký chẩn đoán không đưa vào Git.
- 37 kiểm thử đều đạt. Chạy tích hợp thật trên Linux với bố trí cache ngắn giống quyết định đường dẫn Windows: cài/import đủ thư viện ngoài thư mục giải nén dài/có dấu, thư mục giải nén thứ hai dùng lại cache mà không chạy pip, tệp dữ liệu mẫu không thay đổi. Kiểm thử log chạy chẩn đoán nhiều lần giữ lỗi trước đó. Đây là kiểm tra logic/bố trí trên Linux, không phải chứng nhận hoạt động trên Windows.

## Bản 0.3.3 — sửa ngày chốt nhập nhầm

- Tái hiện với bản sao tệp thật: nhập 762 hồ sơ ngày 28/10/2026; báo cáo 01/01–30/09 trả 0. Xem trước sửa ngày chốt 30/09 có 762 hồ sơ đủ điều kiện; sau sửa báo cáo cuối kỳ đếm 692 đang làm và giữ 70 đã nghỉ, không sinh/xóa hồ sơ. Cảnh báo thiếu lịch sử trước ngày chốt vẫn giữ nguyên.
- Kiểm thử API: bảo toàn snapshot và nhật ký ngày cũ/mới; chặn sửa hồ sơ có biến động/quyết định, ngày chốt trước ngày sinh/ngày vào làm; kiểm tra xác nhận/lý do, xung đột phiên bản, thay đổi nguồn và rollback toàn giao dịch.
- Chromium: tái hiện ngày chốt sai → mở hộp thoại sửa từ báo cáo → xem trước/nhập lý do/xác nhận → tổng hợp và xuất được số liệu trong cùng kỳ đã chọn; thao tác lặp bị vô hiệu khi đã đúng ngày; ngày nhập mới để trống bắt buộc chọn; thông báo đỏ cũ được xóa. Không có lỗi JavaScript.
- Kiểm thử trên Linux, chưa trực tiếp trên cơ sở dữ liệu máy Windows người dùng. Giữ chức năng khởi động/chẩn đoán 0.3.2, chỉ đổi nhãn phiên bản.
