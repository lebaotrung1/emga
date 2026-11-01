# Ứng dụng Windows v5.2 — lương quản lý phân xưởng

Bản chương trình riêng, không thay đổi ứng dụng Emga ở thư mục gốc kho này.

## Sửa tính tiền sản lượng (v5.2)

Xóa hoặc nhập 0 ở sản lượng/đơn giá loại 1, loại 2 hay ngày lễ sẽ đưa thành tiền dòng đó về 0 ngay. Nhập lại thì tự tính thành tiền = sản lượng × đơn giá (làm tròn theo bảng nguồn: 100đ; ngày lễ Luyện gang 1.000đ). Các ô thành tiền này chỉ hiển thị kết quả tính. Bổ sung đọc/ghi các ô đơn giá loại 2 của Thiêu kết, Luyện thép và đơn giá ngày lễ Luyện thép vốn bị bỏ sót.

Quỹ sản phẩm trực tiếp cập nhật theo phần tăng/giảm của dòng vừa sửa; giữ các khoản khác trong quỹ. Lương cá nhân trực tiếp được phân bổ lại, quỹ về 0 không giữ lương cũ. Quỹ quản lý vẫn là khoản được duyệt riêng trong cửa sổ quản lý; chưa tự suy diễn các công thức sản lượng tính quỹ quản lý. Việc nhập dữ liệu từ Excel vẫn lấy giá trị đã lưu của nguồn; không thay thế bộ máy công thức Excel. Khi sửa sản lượng/đơn giá, chỉ tính lại dòng vừa sửa, không tự điều chỉnh sản lượng các dòng khác.

Đóng ứng dụng, tải ZIP v5.2, thay `CiscoLDTLApp_v5.exe` và `Chay_Ung_Dung.bat` trong thư mục cũ. Giữ nguyên Data/Templates. Menu và tiêu đề phải hiện **v5.2**, phiên bản EXE **5.2.0.0**. Kiểm tra trên bản sao Excel trước khi lưu đè.

## Sửa lỗi nạp file trên Windows (v5.1)

Sửa lỗi `Method not found: System.String.Split(Char, System.StringSplitOptions)` của bản EXE v5. Bản v5.1 gọi overload `Split(char[], StringSplitOptions)` tương thích .NET Framework và được biên dịch bằng bộ tham chiếu chính thức .NET Framework 4.5 của Microsoft. Tên EXE vẫn là `CiscoLDTLApp_v5.exe` để giữ tương thích với launcher; bản sửa v5.1 đã được giữ trong v5.2 hiện tại.

Đóng ứng dụng cũ, tải ZIP v5.1, chép `CiscoLDTLApp_v5.exe` và `Chay_Ung_Dung.bat` từ thư mục `antigravity` trong ZIP vào thư mục đang sử dụng, chọn thay thế. Giữ nguyên `Data` và `Templates`; bản cập nhật không chứa hoặc thay thế dữ liệu. Mở lại, kiểm tra menu hiện **v5.2**, rồi nạp Excel/bảng công.

## Tải và chạy

Tải `antigravity_v5_2_program_only.zip` bằng nút **Download raw** trên trang tệp GitHub. Giải nén vào thư mục mới trên Windows rồi chạy `Chay_Ung_Dung.bat` hoặc `CiscoLDTLApp_v5.exe`.

Gói này **không kèm tệp Excel, mẫu Word, hồ sơ nhân sự hoặc số lương mẫu**. Trong ứng dụng, bấm chọn tệp Excel tổng hợp hiện có trên máy của bạn. Nếu cần các mẫu Word riêng, sao chép thư mục `Templates` từ bản ZIP cũ vào thư mục chạy mới, giữ bản gốc để dự phòng. Không nạp dữ liệu thì các bảng để trống.

Chọn phân xưởng/tháng/năm, mở menu **TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5)**, kiểm tra quỹ và công/hệ số, bấm **Tính / xem trước** → **Áp dụng vào bảng nghiệm thu** → **Lưu Excel**. Nên dùng bản sao Excel để kiểm tra một kỳ đã duyệt trước.

Xem `KIEM_THU_V5.md` để biết công thức, giới hạn và các kiểm tra đã thực hiện. Đã bỏ tên người và dữ liệu nghiệp vụ được nhúng sẵn trong bộ mặc định của bản gốc; các ánh xạ sheet/dòng vẫn được giữ. Bộ kiểm thử dùng Excel cần bản dữ liệu riêng ở đường dẫn `Data/03.9.2026_NT_T9_TCHC.xlsx`, tệp này không đưa lên GitHub. Kiểm thử chỉ dùng dữ liệu giả có thể chạy bằng `Kiem_Thu_Logic_V5.ps1`.

Chưa chạy trực tiếp trên Windows; đã biên dịch và kiểm thử logic, Excel, giao diện trên Linux/Mono.
