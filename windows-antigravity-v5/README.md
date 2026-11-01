# Antigravity v5.5 — lương cẩu phôi PX Thép

Tải `antigravity_v5_5_program_only.zip` trên nhánh `antigravity-payroll-v5`. Đóng chương trình, sao lưu Excel, giải nén và thay EXE/launcher; giữ nguyên Data và Templates. Tiêu đề là v5.5, phiên bản EXE 5.5.0.0. Gói chỉ chứa chương trình, mã nguồn, hướng dẫn và kiểm thử; không có dữ liệu người dùng.

Chọn PX Thép, nạp Excel như trước, mở **LƯƠNG CẨU PHÔI PX THÉP (v5.5)**. Có ba công đoạn:

1. Vận chuyển phôi từ SX đến xếp kiêu ngoài bãi.
2. Xuất bán phôi từ kiêu ngoài bãi.
3. Vệ sinh phôi, xuất bán hàng từ kho sản xuất.

Mỗi dòng nhập SL thường **không gồm SL lễ**, SL lễ và đơn giá riêng; dùng dấu phẩy cho phần thập phân. Thành tiền = SL thường × đơn giá + SL lễ × đơn giá × 2. Sản lượng lễ chỉ tính trong phần lễ; không tự cộng sang sản lượng sản xuất hoặc sản lượng quản lý. Sản lượng cẩu phôi là khối lượng từng công đoạn, không phải tổng tấn sản phẩm duy nhất.

Bấm **Áp dụng** để thay phần cẩu phôi cũ trong quỹ sản phẩm trực tiếp, cập nhật nghiệm thu và Word. Bấm **Lưu Excel** để lưu. Hủy cửa sổ không áp dụng đầu vào mới. Có thể sửa/xóa từng sản lượng hoặc đơn giá về 0; áp dụng lặp không cộng quỹ lần nữa.

Ánh xạ Excel: thường D50:D52, lễ D46:D48, đơn giá I50:I52; J49/J45 là tổng thường/lễ. Các khoản khác tại J55:J58 được giữ, cộng một lần trong J44; tổng này nằm trong J35. Công thức cũ ở đầu vào thường đã trừ lễ được thay bằng SL thường đã tách, tránh trừ lễ lần nữa. Sáu ô thành tiền dùng công thức nhân trực tiếp theo yêu cầu mới. Sheet **Cau phoi v5** lưu chi tiết và đầu vào. Dòng tổng SP phụ/cẩu phôi ở màn hình chính chỉ đọc sau khi áp dụng, tránh nhập đè quỹ tính tự động.

Word xem trước, xuất mới và cập nhật mẫu đều liệt kê ba công đoạn, SL thường/lễ, đơn giá đồng/ĐVT và tiền triệu đồng có đúng một số lẻ. Với kỳ quý/lũy kế, chi tiết cẩu phôi ghi rõ là **tháng**; tổng lương quý/lũy kế dùng cơ chế cộng nền hiện có.

Giữ chức năng nạp công, quản lý theo sản lượng với đơn giá riêng, GNV tính lại theo công/khối lượng × đơn giá và Word triệu đồng của bản trước. Chỉ sửa ứng dụng antigravity; không sửa Nghiệm Thu Studio. Xem **RA_SOAT_CDNL.md** để biết công thức từng bộ phận và giới hạn tính lại liên sheet của CĐNL; không áp công thức cẩu phôi chung cho các bộ phận này.

`Build.ps1` biên dịch trên Windows. `Kiem_Thu_Logic_V5.ps1` chạy các kiểm thử với dữ liệu giả. Kiểm thử Excel nguồn dùng bản sao ngoài repo; kiểm tra tính toán, lưu–mở lặp, giữ công thức CĐNL, nạp công, quản lý, GNV và xuất/cập nhật Word. Kiểm tra chạy trên Linux/Mono, chưa thử trực tiếp Windows.

---

# Ứng dụng Windows v5.4 — quỹ quản lý theo sản lượng và báo cáo Word triệu đồng

Ứng dụng Windows riêng trong nhánh `antigravity-payroll-v5`; không thay ứng dụng Emga trên nhánh chính.

## Tải và cập nhật

Tải `antigravity_v5_4_program_only.zip` bằng nút **Download raw** trên trang tệp GitHub. Đóng chương trình, giải nén rồi chép `CiscoLDTLApp_v5.exe` và `Chay_Ung_Dung.bat` trong thư mục `antigravity` vào thư mục bản đang dùng, chọn thay thế. Giữ nguyên **Data** và **Templates**, sao lưu Excel trước khi kiểm tra.

Chạy `Chay_Ung_Dung.bat`. Menu/tiêu đề phải hiện **v5.4**, thuộc tính EXE là **5.4.0.0**. Tên EXE vẫn là v5 để launcher cũ tương thích. Gói chỉ có chương trình/mã nguồn/hướng dẫn/kiểm thử, không chứa Excel, mẫu Word, hồ sơ nhân sự, tên thật hay số lương mẫu.

## Tính tiền công GNV (v5.4)

Dòng mới hoặc dòng có sửa khối lượng/số công hay đơn giá tự tính **thành tiền = khối lượng × đơn giá**, làm tròn đến đồng. Sửa công/giá về 0 hoặc xóa nội dung ô số sẽ đưa tiền về 0; cột thành tiền chỉ đọc. Sau khi kết thúc sửa ô, tổng GNV và lương phân bổ được cập nhật. Xóa dòng cuối không giữ tổng tiền cũ.

Dòng vừa nhập từ Excel nhưng chưa sửa vẫn giữ số tiền nguồn để không phá các công thức giảm trừ/làm tròn riêng. Sau khi sửa công/đơn giá, dòng chuyển sang tính theo phép nhân. Vẫn hỗ trợ số âm cho khoản điều chỉnh. CĐNL giữ cách tổng hợp quỹ riêng của bản gốc, không cộng trùng GNV vào quỹ đã có.

Bấm **Lưu Excel** để lưu thay đổi. Sheet **GNV v5** giữ danh sách/đầu vào đã chỉnh, dòng Excel gốc và trạng thái tự tính (kể cả danh sách đã xóa hết), ưu tiên khi ứng dụng đọc lại. Nếu sửa GNV trực tiếp bằng Excel sau đó, cần đối chiếu với dữ liệu ứng dụng đã lưu trong sheet này. Các công thức NT cũ vẫn có thể tính lại theo quy tắc riêng khi mở Excel; kiểm tra các biểu gốc trước khi dùng thanh toán.

## Quỹ lương quản lý tính từ sản lượng

1. Chọn phân xưởng, tháng/năm và nạp Excel của đơn vị. Kiểm tra sản lượng **loại 1, loại 2, ngày lễ** tại bảng sản xuất. Nếu chưa có dữ liệu nguồn, nhập sản lượng tương ứng ở đó.
2. Mở menu **TÍNH LƯƠNG QUẢN LÝ PHÂN XƯỞNG (v5.4)**. Nhập **đơn giá quản lý loại 1** và **đơn giá quản lý loại 2** (đồng/đơn vị sản lượng). Đây là đơn giá riêng, không dùng đơn giá lao động trực tiếp.
3. Đơn giá quản lý ngày lễ **tự bằng 2 × đơn giá quản lý loại 1**, chỉ hiển thị, không nhập riêng. Quỹ sản phẩm quản lý = SL loại 1 × ĐG QL loại 1 + SL loại 2 × ĐG QL loại 2 + SL lễ × (2 × ĐG QL loại 1). Mỗi dòng làm tròn đến đồng rồi cộng, không nhập tổng quỹ bằng tay.
4. Kiểm tra quỹ vừa tính, bấm **Tính / xem trước**, kiểm tra hệ số/công/xếp loại từng người rồi **Áp dụng vào bảng nghiệm thu**. Quỹ được phân bổ cho quản lý theo công/hệ số/xếp loại như bản trước; lễ phép/học/phụ cấp và điều chỉnh chung vẫn tách riêng để tránh cộng trùng.
5. Bấm **Lưu Excel**. Đơn giá/chế độ tính được lưu riêng cho từng PX trong sheet **Quy QL san luong**; giữ sheet này và **Nhan su bo sung**. Sheet **Luong quan ly v5** chứa tổng hợp và chi tiết lương.

Sau khi áp dụng, sửa/xóa sản lượng tại bảng sản xuất sẽ tính lại quỹ quản lý và lương phân bổ, kể cả khi sản lượng bằng 0. Sửa đơn giá QL loại 1 trong cửa sổ sẽ tự cập nhật đơn giá lễ và tổng quỹ; đóng/hủy không ghi các thay đổi chưa áp dụng. Dữ liệu cũ chưa có đơn giá QL vẫn giữ quỹ đã lưu cho đến khi nhập đơn giá và áp dụng v5.3, không tự suy đơn giá từ số tiền cũ.

Thiêu kết/Luyện gang/Luyện thép đọc lại sản lượng từ các dòng NT gốc. Cơ điện Năng lượng/Khai tuyển không có cùng ánh xạ ba dòng trong bản cũ, nên các sản lượng ba loại nhập qua ứng dụng được lưu trong sheet cấu hình trên. Không coi đại lượng quỹ lương bằng đồng của CĐNL là sản lượng tấn.

## Báo cáo nghiệm thu Word

Các khoản tiền trên báo cáo của **cả 5 PX**, kỳ tháng/quý/lũy kế, xem trước, xuất mới và phần báo cáo cập nhật trong mẫu có sẵn đều dùng **triệu đồng với đúng 1 chữ số sau dấu phẩy**: `123.456.789đ → 123,5 triệu đồng`, `1.000.000đ → 1,0 triệu đồng`. Tiền lương bình quân vốn lưu bằng nghìn đồng được chuyển đúng sang triệu đồng/người/tháng. Bảng GNV, thành tiền BDĐH và bảng lương từng người kèm báo cáo cũng đổi đơn vị.

Đơn giá theo công/đơn vị vẫn ghi rõ bằng đồng/công hoặc đồng/ĐVT để các mức nhỏ 13.000đ, 20.000đ không mất ý nghĩa khi làm tròn. Việc đổi đơn vị chỉ ở phần hiển thị Word; dữ liệu tính toán, nhập liệu và Excel vẫn là đồng. Các biên bản lịch sử nằm trước phần báo cáo đang cập nhật trong mẫu gốc được giữ nguyên.

## Kiểm thử và giới hạn

Bản EXE được biên dịch với bộ tham chiếu .NET Framework 4.5 chính thức của Microsoft, giữ bản sửa lỗi Split trên Windows. Đã chạy 382 kiểm tra mới cho quỹ theo sản lượng và Word (5 PX × 5 kỳ, xuất mới/cập nhật mẫu, lưu–mở lại Excel), kiểm thử giao diện nhập đơn giá/tự nhân đôi lễ/áp dụng/hủy, 64 kiểm tra lương/Excel, các kiểm thử nạp nhân sự và sửa/xóa sản lượng trước đó. Các kiểm tra đạt trên Linux/Mono; chưa chạy trực tiếp Windows.

`Build.ps1` biên dịch trên Windows. `Kiem_Thu_Logic_V5.ps1` chạy với dữ liệu giả, không cần Excel riêng; `Kiem_Thu_V5.ps1` cần tệp dữ liệu cục bộ `Data/03.9.2026_NT_T9_TCHC.xlsx` (không đưa lên GitHub). `KIEM_THU_V5.md` lưu lịch sử các thay đổi và chi tiết kiểm tra. Nên đối chiếu một kỳ đã duyệt trên bản sao Excel trước khi dùng thanh toán.
