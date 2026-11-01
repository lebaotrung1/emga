# Ứng dụng Windows v5 — lương quản lý phân xưởng

Bản chương trình riêng, không thay đổi ứng dụng Emga ở thư mục gốc kho này.

## Tải và chạy

Tải `antigravity_v5_program_only.zip` bằng nút **Download raw** trên trang tệp GitHub. Giải nén vào thư mục mới trên Windows rồi chạy `Chay_Ung_Dung.bat` hoặc `CiscoLDTLApp_v5.exe`.

Gói này **không kèm tệp Excel, mẫu Word, hồ sơ nhân sự hoặc số lương mẫu**. Trong ứng dụng, bấm chọn tệp Excel tổng hợp hiện có trên máy của bạn. Nếu cần các mẫu Word riêng, sao chép thư mục `Templates` từ bản ZIP cũ vào thư mục chạy mới, giữ bản gốc để dự phòng. Không nạp dữ liệu thì các bảng để trống.

Chọn phân xưởng/tháng/năm, mở menu **TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5)**, kiểm tra quỹ và công/hệ số, bấm **Tính / xem trước** → **Áp dụng vào bảng nghiệm thu** → **Lưu Excel**. Nên dùng bản sao Excel để kiểm tra một kỳ đã duyệt trước.

Xem `KIEM_THU_V5.md` để biết công thức, giới hạn và các kiểm tra đã thực hiện. Đã bỏ tên người và dữ liệu nghiệp vụ được nhúng sẵn trong bộ mặc định của bản gốc; các ánh xạ sheet/dòng vẫn được giữ. Bộ kiểm thử dùng Excel cần bản dữ liệu riêng ở đường dẫn `Data/03.9.2026_NT_T9_TCHC.xlsx`, tệp này không đưa lên GitHub. Kiểm thử chỉ dùng dữ liệu giả có thể chạy bằng `Kiem_Thu_Logic_V5.ps1`.

Chưa chạy trực tiếp trên Windows; đã biên dịch và kiểm thử logic, Excel, giao diện trên Linux/Mono.
