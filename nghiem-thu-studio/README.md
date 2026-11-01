# Nghiệm Thu Studio 1.0 — bản thử độc lập cho nghiệm thu tổng hợp

Bản mới tập trung **nghiệm thu tổng hợp theo phân xưởng và biên bản Word**, theo ưu tiên người dùng. Không phải bản vá/chép đè antigravity và chưa chia lương từng người. Chạy trực tiếp trên Windows 10/11 có .NET Framework 4.5 trở lên; không cần Python hoặc Office để tính/lưu/xuất DOCX.

## Tải và chạy

Tải **NghiemThuStudio_1_0_Windows.zip** bằng nút **Download raw** trên trang GitHub. Giải nén vào **thư mục mới**, chạy `Mo_Nghiem_Thu.bat` hoặc `NghiemThuStudio.exe`. Không chép vào bản antigravity cũ.

Gói không chứa Excel, Word hay dữ liệu nhân sự người dùng. `Mau_GIA_LAP.ntproject` chỉ là **số giả để thử, KHÔNG dùng thanh toán**; có thể bấm Mở hồ sơ để thử trước. Excel/Word thật vẫn chọn từ máy của bạn, phần mềm không gửi tệp lên mạng.

## Luồng làm việc

1. **Nguồn & đường dẫn:** chọn đơn vị/tháng/năm, Nạp Excel nguồn. Chỉ đọc tệp `.xlsx` đã lưu. Kiểm tra sheet, kỳ, liên kết ngoài, ô lỗi. Không tự chọn sheet tên gần giống. Nếu nguồn ghi kỳ khác, có cảnh báo để chọn lại hoặc giải trình.
2. **Đầu vào & đơn giá:** xem giá trị nguồn/ô nguồn bên cạnh đầu vào. Nhập đơn giá trực tiếp L1/L2 và quản lý L1/L2; giá lễ tự bằng hai lần giá L1 tương ứng. Kiểm tra ngày công, tiền chế độ, ăn ca, BDĐH, điều chỉnh và mức làm tròn. **Ô trống = chưa biết, không phải 0.** Muốn bỏ một khoản, nhập `0` hoặc dùng nút ghi 0 cho các dòng đã chọn.
3. **GNV & sản phẩm phụ:** nhập khối lượng, đơn giá, hệ số và mức làm tròn. Thành tiền chỉ đọc, luôn tính lại theo đầu vào. Hệ số âm dùng cho khoản giảm trừ. Cột loại khoản tách **GNV** khỏi **SP phụ** (ví dụ cẩu phôi), tránh cộng vào hai quỹ. Số tiền nguồn chỉ để so sánh; chênh phải kiểm tra, không tự chép làm kết quả.
4. **Cuộc họp & giải trình:** điền ngày, địa điểm, ĐVT cơ sở sản lượng, thành phần, người ký, kiến nghị và lý do chênh lệch. Đánh dấu đã đối chiếu. Mỗi lần sửa số liệu sẽ bỏ xác nhận cũ, yêu cầu kiểm tra lại.
5. **Kết quả & biên bản:** kiểm tra quỹ thành phần và chênh nguồn. Có thể xuất Word dự thảo ngay; thiếu dữ liệu/kỳ chốt sẽ có nhãn rõ. Các khoản tiền Word dùng **triệu đồng, đúng một số sau dấu phẩy**; đơn giá theo công/ĐVT ghi đồng riêng. Tiền nội bộ vẫn tính bằng `decimal`, không làm tròn theo triệu.
6. **Lưu/chốt:** Lưu hồ sơ `.ntproject` ở thư mục bạn chọn. Hồ sơ giữ nhiều nháp và nhiều kỳ đã chốt. Chốt kỳ kiểm tra đầu vào, giải trình, thành phần/người ký và tệp nguồn còn nguyên. Bản chốt là bản sao riêng; sửa tiếp không sửa bản đã chốt. Mỗi lần lưu dùng tệp tạm và giữ bản trước trong `.bak`.

Có thể mở bản nháp/bản chốt ở **Hồ sơ & nhật ký** để xem hoặc tạo bản sửa. Quý/lũy kế chỉ cộng **bản chốt mới nhất của từng tháng** đúng đơn vị/năm; không cộng hai lần bản sửa. Thiếu tháng nào thì báo tháng đó, không coi là 0. Bình quân dùng tổng người-tháng, không chia thừa thêm một lần cho số tháng. Tháng hiện tại chưa chốt thì báo cáo gộp cũng là dự thảo.

Khi chuyển chỗ Excel, dùng **Gắn lại Excel đã di chuyển**. Chỉ gắn lại nếu SHA256 trùng tệp đã nhập, không thay chứng cứ bằng một tệp khác. Muốn đổi nội dung nguồn, nạp lại để đối chiếu. Mẫu Word được đọc để kiểm tra cấu trúc/kỳ/tên đơn vị; báo cáo xuất mới theo các phần I/II/III, không ghi đè mẫu hoặc chép tên/người ký/số cũ từ tập nhiều tháng.

## Công thức và phạm vi

- TT sản phẩm = L1 × ĐG TT1 + L2 × ĐG TT2 + Lễ × 2 × ĐG TT1 + SP phụ.
- QL sản phẩm = L1 × ĐG QL1 + L2 × ĐG QL2 + Lễ × 2 × ĐG QL1. Hai bộ đơn giá độc lập.
- Lương TT = TT sản phẩm + GNV + chế độ TT + điều chỉnh TT.
- Lương QL = QL sản phẩm + chế độ QL + điều chỉnh QL. Tổng quỹ = TT + QL.
- GNV/SP phụ = khối lượng × đơn giá × hệ số, làm tròn theo mức hiển thị. Các khoản cố định/khác phải nhập và giải trình rõ; không giấu trong quỹ tổng.
- Ăn ca, BDĐH = công đã đối chiếu × mức giá. Các mức khởi tạo và ngày công chuẩn cần kiểm tra theo quy chế thực tế.

Sơ đồ đọc có cho 5 sheet NT của bộ nguồn hiện tại, không phải trình đọc mọi mẫu Excel. Thiêu kết/Gang/Thép có ánh xạ ba loại sản lượng; CĐNL/Khai tuyển có cấu trúc/cơ sở quỹ đặc thù nên cần nhập rõ cơ sở và đơn giá đã xác nhận, **không tự biến quỹ tiền thành sản lượng tấn**. Chưa tự tái tạo tất cả công thức quỹ đặc thù, ứng trước, trung đại tu, điều tiết giữa đơn vị. Phải bổ sung khoản/cơ sở đúng và đối chiếu chênh trước chốt.

Đây **không có bộ máy tính Excel**: đọc giá trị đã lưu và công thức làm chứng cứ. Kiểm tra liên kết chỉ tìm tệp theo tên ở thư mục được chọn, không tính lại liên kết/INDEX/MATCH ngoài. Ô lỗi hoặc thiếu cache giữ là chưa đủ, không thay bằng 0. `.xls` cần chuyển `.xlsx` bằng Excel. Dữ liệu các tháng cũ trong Word/cột lũy kế Excel không tự trở thành lịch sử đã chốt; cần nhập/đối chiếu các kỳ tương ứng.

Không ghi ngược Excel gốc. Đầu ra gồm hồ sơ, CSV đối chiếu và Word mới. Bản này chưa tính bảo hiểm phải nộp, thuế hay thực nhận từng người; chưa có đăng nhập/phân quyền nhiều người hoặc chữ ký số. Lưu hồ sơ trên máy có quyền truy cập phù hợp và giữ bản sao dự phòng.

## Kiểm chứng

Đã biên dịch EXE bằng reference assemblies .NET Framework 4.5 chính thức của Microsoft. Kiểm thử mô hình/lưu hồ sơ/khóa bản chốt/các kỳ/đơn vị tiền/DOCX/nguồn và giao diện trên Linux/Mono; kiểm tra cả Excel và Word người dùng cung cấp, không đổi tệp gốc. Các nguồn thật không đưa vào GitHub. Chưa chạy trực tiếp trên Windows hoặc Microsoft Word; đây là bản thử để đối chiếu một kỳ trước khi dùng thanh toán.

`Build.ps1` biên dịch trên Windows; `Kiem_Thu.ps1` chạy kiểm thử giả lập không cần Excel người dùng; `Kiem_Thu.ps1 -Ui` chạy thêm thử giao diện trên máy có desktop. Mã nguồn tách Domain, Workbook, ProjectStore, Report, TemplateInspector và MainForm; phần tính là hàm thuần, không lấy ô tiền nguồn làm kết quả.
