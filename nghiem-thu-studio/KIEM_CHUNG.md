# Kiểm chứng bản thử 1.0

- 72 kiểm tra mô hình, đầu vào, tính quỹ, bản chốt, lưu/đọc, tổng hợp kỳ, nguồn Excel và DOCX đã đạt.
- 13 kiểm tra giao diện đã đạt, gồm sửa đầu vào, GNV tính lại, chốt/sửa và hủy đóng cửa sổ mà vẫn tiếp tục sửa được.
- Đọc thử năm sheet nghiệm thu của Excel và mẫu Word được cung cấp theo chế độ chỉ đọc; kiểm tra hash xác nhận nguồn không đổi.
- DOCX giả lập mở được bằng thư viện đọc DOCX độc lập: 25 đoạn, 6 bảng.
- EXE biên dịch với reference assemblies .NET Framework 4.5 chính thức. Các kiểm tra chạy trên Linux/Mono; chưa chạy trực tiếp trên Windows/Microsoft Word.

Cần thử đối chiếu một kỳ thực tế trên Windows trước khi dùng thanh toán. Gói chỉ có dữ liệu giả lập, không chứa tài liệu hay số liệu của người dùng.
