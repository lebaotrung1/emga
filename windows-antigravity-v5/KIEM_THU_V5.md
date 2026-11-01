# Bản v5 — tính lương quản lý từng phân xưởng

## Cách dùng

1. Giải nén vào thư mục mới, giữ lại bản cũ và sao lưu tệp Excel đang sử dụng.
2. Chạy `Chay_Ung_Dung.bat` hoặc `CiscoLDTLApp_v5.exe`.
3. Chọn phân xưởng, tháng/năm và nạp tệp Excel cần tính.
4. Bấm menu **TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5)**. Màn hình chỉ hiển thị nhân sự được xác định thuộc bộ phận quản lý của phân xưởng đó.
5. Kiểm tra quỹ SP quản lý, công, hệ số cố định, xếp loại A/B/C, lương BHXH, công phép, học/họp, VHTT. Nhập số tiền làm lễ và phụ cấp đã được duyệt; nhập 0 để bỏ khoản cũ.
6. Bấm **Tính / xem trước**, đối chiếu tổng cá nhân, chênh làm tròn và tổng quản lý. Bấm **Áp dụng vào bảng nghiệm thu**, rồi bấm **Lưu Excel** ở màn hình chính. Chỉ xem trước/đóng cửa sổ không áp dụng các thay đổi trong cửa sổ.

## Phạm vi và công thức

Đây là phân bổ **quỹ lương quản lý phân xưởng**, không phải lương bộ phận quản lý Công ty, BHXH phải nộp, thuế TNCN hay thực nhận sau khấu trừ.

- Giữ cách phân bổ theo quỹ SP trong mã nguồn cũ: điểm = ROUND(công SP × hệ số cố định / 26; 2); hệ số hưởng = ROUND(điểm × tỷ lệ xếp loại; 2), A/B/C = 100%/90%/80%.
- Lương SP mỗi người = quỹ SP × hệ số hưởng / tổng hệ số hưởng, làm tròn 100đ theo quy tắc midpoint của Excel.
- Lương phép = công phép × mức lương BHXH / 26, làm tròn 100đ.
- Học/họp = 250.000đ/công; VHTT = 300.000đ/công theo quy tắc bảng NT/mã nguồn hiện có. Một số bảng chi tiết Excel có đơn giá khác; cần kiểm tra quy chế đơn vị trước khi dùng thanh toán.
- Tiền làm lễ và phụ cấp trong cửa sổ mới là số tiền nhập/đọc riêng. Không tự áp một công thức ngày lễ cho tất cả phân xưởng.
- Tổng cá nhân gồm SP + phép + học/họp/VHTT + làm lễ + phụ cấp. Dòng tổng các khoản ngoài SP được đưa vào `LuongCheDoQL` để đồng bộ tổng; không cộng lần thứ hai vào quỹ SP. Nếu quỹ SP đã bao gồm làm lễ/phụ cấp, cần loại khoản trùng khỏi quỹ hoặc khoản nhập trước khi áp dụng.
- Điều chỉnh chung chỉ cộng vào tổng quản lý của phân xưởng, chưa phân bổ cho cá nhân. Tổng NT dùng quỹ SP đã nhập; tổng chi tiết có thể chênh quỹ do làm tròn. Cửa sổ hiển thị chênh lệch, không tự gán chênh lệch cho một người.
- Quỹ Khai tuyển trong bản cũ là số mẫu hard-code, còn bảng nguồn có lỗi `#N/A`. Bản v5 để quỹ chưa xác định bằng 0 và cảnh báo phải nhập quỹ đã duyệt, thay vì âm thầm lấy số mẫu. Các PX khác cũng cảnh báo nếu ô quỹ không có giá trị số hợp lệ.
- Không tự tính quỹ SP từ các công thức sản lượng trong Excel: tệp nguồn có công thức `#REF!`, liên kết ngoài và các cách tính khác nhau. Quỹ cần được đối chiếu/phê duyệt trước khi phân bổ. Việc giữ công thức và cập nhật cache không chứng minh Excel sẽ tính lại cùng kết quả sau khi mở.

## Các sửa lỗi kèm theo

- Công hoặc quỹ SP về 0 không giữ lại lương quản lý cũ.
- Xóa khoản phép/học/VHTT sẽ cập nhật lại tổng, không giữ tổng cũ.
- Tổng quản lý tính đủ các khoản ngoài SP; không thay lương cá nhân của lao động trực tiếp khi dùng cửa sổ quản lý.
- Nhận diện tiêu đề `Tổ ...`, `Tổ: ...`, kể cả dòng có mã 0, ô lỗi hoặc chức danh, để tránh kéo tổ sản xuất vào nhóm quản lý. Khi snapshot v4 có cờ quản lý mâu thuẫn với nhóm đọc lại từ sheet gốc, ưu tiên nhóm của sheet gốc; các điều chuyển nhóm thủ công cũ cần đối chiếu lại.
- Khi ghi dòng quản lý trong bảng NT, ghép theo tên khớp duy nhất thay vì thứ tự. Không ghi vào dòng tên lỗi/không khớp/trùng; cần sửa mẫu nếu muốn các dòng đó tự cập nhật.
- Khi lưu Excel, thêm sheet **Luong quan ly v5** dễ đọc: tổng hợp quỹ/tổng từng PX và danh sách lương từng quản lý, tách chênh làm tròn/điều chỉnh chung. Đây là số liệu v5 đã lưu, không phụ thuộc Excel tính lại các công thức NT cũ.
- Lưu cấu hình và kết quả quản lý v5 trong sheet `Nhan su bo sung`, đọc lại khi mở. Giữ sheet này; các hồ sơ đã tính v5 ưu tiên dữ liệu lưu tại đây. Nếu chỉnh dữ liệu bằng Excel sau đó, cần chỉnh lại qua cửa sổ v5 để đồng bộ.

## Kiểm thử

Đã biên dịch bản EXE bằng Mono C# trên Linux. Kiểm thử tính lương và lưu–mở lại trên **cả 5 phân xưởng**: 64 kiểm tra đạt; kiểm thử hồi quy ImportTests đạt. Kiểm thử giao diện WinForms trên Linux/Mono: mở cửa sổ, xem trước không ghi vào dữ liệu chính, áp dụng, đóng/hủy, menu chính mở cửa sổ đều đạt.

Đã đối chiếu 72.757 công thức gốc trong tệp Excel kiểm thử: nội dung/vị trí công thức được giữ nguyên. Tệp Excel gốc trong Data và các mẫu trong Templates được giữ nguyên từng byte so với ZIP người dùng gửi. Kiểm thử chỉ ghi vào thư mục test-output của bản làm việc.

Chưa chạy trực tiếp trên Windows. Bản phát hành có EXE đã biên dịch, mã nguồn và kiểm thử; cần mở thử trên máy Windows và đối chiếu một kỳ lương đã duyệt trước khi dùng để thanh toán. Không dùng kết quả ở đây để kết luận toàn bộ công thức nghiệp vụ Excel đã được xác nhận.

Biên dịch trên Windows: chạy `Build.ps1` (giữ cách dùng csc .NET Framework của bản cũ). Script tạo `CiscoLDTLApp_v5.exe`. `Kiem_Thu_V5.ps1` biên dịch/chạy hai bộ kiểm thử logic và Excel trên Windows; không tự động chạy giao diện.

## Gói GitHub chỉ chứa chương trình

Bản đưa lên GitHub không kèm Data, Templates, báo cáo Word, tệp kiểm thử có dữ liệu hay bất kỳ Excel nào. Bộ mặc định trong mã nguồn/EXE chỉ giữ cấu trúc 5 phân xưởng, không chứa tên người hoặc các số liệu lương được nhúng ở bản gốc. Cần chọn Excel riêng trên máy; có thể sao chép Templates từ bản cũ nếu dùng mẫu riêng.

Bản đã làm sạch được biên dịch và kiểm tra lại: 18 kiểm tra bằng dữ liệu giả, toàn bộ 64 kiểm tra cùng vòng lưu–mở Excel 5 phân xưởng trên bản dữ liệu cục bộ ngoài kho Git, kiểm thử nạp nhân sự và kiểm thử giao diện khởi động không có Excel đều đạt. Không đưa dữ liệu/đầu ra của các kiểm tra đó lên GitHub.

## Bản sửa tương thích Windows v5.1

Ảnh lỗi Windows xác nhận EXE v5 gọi `String.Split(Char, StringSplitOptions)`, overload có ở thư viện Mono dùng khi biên dịch nhưng không có ở .NET Framework. Sửa lời gọi thành `Split(new char[] { ',' }, StringSplitOptions.None)`, biên dịch toàn bộ ứng dụng với `-noconfig -nostdlib` và các reference assemblies .NET Framework 4.5 chính thức từ gói NuGet `Microsoft.NETFramework.ReferenceAssemblies.net45` phiên bản `1.0.3`. Các tham chiếu được chỉ định tường minh: mscorlib, System, System.Core, System.Xml, System.IO.Compression, System.IO.Compression.FileSystem, System.Windows.Forms, System.Drawing. Không phân phối thư viện Mono trong gói ứng dụng. Build.ps1 trên Windows vẫn dùng trình biên dịch và thư viện .NET Framework của Windows.

Thêm FrameworkCompatibilityTests: đọc IL của EXE để kiểm tra chính xác chữ ký Split(char[], StringSplitOptions), kiểm tra phiên bản 5.1.0.0, nạp dữ liệu BDĐH giả có nhiều mã cách nhau bởi dấu phẩy và kiểm tra nạp lại không cộng trùng. Kiểm thử này từ chối EXE v5 cũ và chấp nhận EXE v5.1. Đã chạy lại 18 kiểm tra logic, 64 kiểm tra lương/Excel, hồi quy nạp nhân sự và smoke test giao diện trên Linux/Mono, đều đạt. Chưa chạy trực tiếp v5.1 trên máy Windows; không coi các kiểm tra Linux là bằng chứng đã chạy Windows.
