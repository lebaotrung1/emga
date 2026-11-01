# Từ nghiên cứu nguồn đến thiết kế bản thử

## Những vấn đề cần giải quyết

Bộ Excel có nhiều sheet đang dùng và sheet bản sao/lịch sử; tên tệp không bảo đảm mọi sheet cùng kỳ. Liên kết ngoài gồm bảng công, điểm, ăn ca, nghiệm thu sản xuất và BHXH, có cả đường dẫn tương đối lẫn đường dẫn tuyệt đối đến thư mục cũ/tệp tạm. Một số ô có lỗi công thức; một số ô có số cache nhưng không có đủ tệp liên kết để chạy lại.

Word là tập nhiều biên bản theo tháng/quý/lũy kế, không phải một mẫu đơn trống. Có dấu hiệu sao chép lẫn tên đơn vị/kỳ/năm và đơn vị tiền. Không nên thay đoạn văn theo vị trí rồi coi toàn bộ dữ liệu cũ là đầu vào đáng tin.

Tài liệu này không chứa số lương, tên người hoặc đường dẫn cá nhân. Chẩn đoán chi tiết được tạo trên máy người dùng từ chính tệp họ chọn, không nhúng nguồn thật trong chương trình.

## Quyết định thiết kế

- Excel là chứng cứ chỉ đọc: giữ ô/công thức/cache/SHA256, không sửa nguồn để làm kết quả khớp.
- Giá trị nguồn, giá trị làm việc và kết quả tính là ba thứ riêng. Thiếu/lỗi không phải 0. Sửa đầu vào làm mất trạng thái xác nhận trước.
- Quy tắc tiền thể hiện bằng thành phần, đơn giá, hệ số, mức làm tròn. Không nhập quỹ tổng rồi chia ngược để hợp thức hóa nguồn. Khoản ngoài mô hình chuẩn phải ghi rõ và giải trình.
- SP phụ, GNV, chế độ, điều chỉnh tách biệt. Không áp cấu trúc tấn/sản lượng của phân xưởng chính cho cơ sở quỹ của đơn vị phục vụ.
- Chốt là bản sao có dấu vết đầu vào/kết quả. Nháp mới không sửa bản đã chốt. Bản sửa cùng tháng thay bản cũ trong phép cộng lũy kế, nhưng vẫn giữ lịch sử.
- Dữ liệu lịch sử phải có từng kỳ đã đối chiếu. Thiếu kỳ hiển thị thiếu và giữ báo cáo ở dạng dự thảo, không nhân một tháng cho nhiều tháng.
- Word sinh mới theo bố cục hành chính tương ứng mẫu: quốc hiệu/đơn vị/kỳ, thành phần, lao động, sản lượng/GNV/quỹ, ăn ca, BDĐH, kiến nghị, ký. Không tự lấy tên/người ký cũ hoặc phá tập biên bản lịch sử.

## Việc cần đối chiếu khi thử thực tế

Kiểm tra trước một tháng của một phân xưởng: kỳ nguồn, giá L1/L2, quy tắc lễ gấp đôi, mức làm tròn, các dòng sản phẩm phụ, GNV có hệ số giảm trừ, chế độ và điều chỉnh. So quỹ thành phần với nguồn và giải trình mọi chênh. Sau đó mới chốt và mở rộng các tháng/đơn vị.

Các công thức đặc thù về điều tiết quỹ, sửa chữa lớn, ứng trước, tỷ lệ hưởng và phân bổ giữa đơn vị chưa thể được xác nhận chỉ từ một tệp có liên kết lỗi. Cần bổ sung tệp nguồn/quy chế hoặc định nghĩa quy tắc rồi kiểm thử riêng; không suy đoán hệ số để hết chênh. Nếu cần bước tiếp theo, ưu tiên chuẩn hóa bảng quy tắc/cơ sở quỹ từng đơn vị và nhập lịch sử đã xác nhận trước khi bổ sung chia lương cá nhân.
