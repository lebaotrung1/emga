# Ứng dụng Windows v5.1 — lương quản lý phân xưởng

Bản chương trình riêng, không thay đổi ứng dụng Emga ở thư mục gốc kho này.

## Sửa lỗi nạp file trên Windows (v5.1)

Sửa lỗi `Method not found: System.String.Split(Char, System.StringSplitOptions)` của bản EXE v5. Bản v5.1 gọi overload `Split(char[], StringSplitOptions)` tương thích .NET Framework và được biên dịch bằng bộ tham chiếu chính thức .NET Framework 4.5 của Microsoft. Tên EXE vẫn là `CiscoLDTLApp_v5.exe` để giữ tương thích với launcher; menu và tiêu đề cửa sổ có **v5.1**, thuộc tính phiên bản tệp là **5.1.0.0**.

Đóng ứng dụng cũ, tải ZIP v5.1, chép `CiscoLDTLApp_v5.exe` và `Chay_Ung_Dung.bat` từ thư mục `antigravity` trong ZIP vào thư mục đang sử dụng, chọn thay thế. Giữ nguyên `Data` và `Templates`; bản cập nhật không chứa hoặc thay thế dữ liệu. Mở lại, kiểm tra menu hiện **v5.1**, rồi nạp Excel/bảng công.

## Tải và chạy

Tải `antigravity_v5_1_program_only.zip` bằng nút **Download raw** trên trang tệp GitHub. Giải nén vào thư mục mới trên Windows rồi chạy `Chay_Ung_Dung.bat` hoặc `CiscoLDTLApp_v5.exe`.

Gói này **không kèm tệp Excel, mẫu Word, hồ sơ nhân sự hoặc số lương mẫu**. Trong ứng dụng, bấm chọn tệp Excel tổng hợp hiện có trên máy của bạn. Nếu cần các mẫu Word riêng, sao chép thư mục `Templates` từ bản ZIP cũ vào thư mục chạy mới, giữ bản gốc để dự phòng. Không nạp dữ liệu thì các bảng để trống.

Chọn phân xưởng/tháng/năm, mở menu **TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5)**, kiểm tra quỹ và công/hệ số, bấm **Tính / xem trước** → **Áp dụng vào bảng nghiệm thu** → **Lưu Excel**. Nên dùng bản sao Excel để kiểm tra một kỳ đã duyệt trước.

Xem `KIEM_THU_V5.md` để biết công thức, giới hạn và các kiểm tra đã thực hiện. Đã bỏ tên người và dữ liệu nghiệp vụ được nhúng sẵn trong bộ mặc định của bản gốc; các ánh xạ sheet/dòng vẫn được giữ. Bộ kiểm thử dùng Excel cần bản dữ liệu riêng ở đường dẫn `Data/03.9.2026_NT_T9_TCHC.xlsx`, tệp này không đưa lên GitHub. Kiểm thử chỉ dùng dữ liệu giả có thể chạy bằng `Kiem_Thu_Logic_V5.ps1`.

Chưa chạy trực tiếp trên Windows; đã biên dịch và kiểm thử logic, Excel, giao diện trên Linux/Mono.
