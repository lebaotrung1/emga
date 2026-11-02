# v5.11 — nhận diện cột, không dùng thiếu cột làm công bằng 0

- `TimesheetLayoutTests` tái tạo tiêu đề gộp ô của mẫu Luyện gang: AJ công làm việc, AR tổng ngày lịch; tổng điểm/điểm chia lương; điểm chia lương đã xếp loại không bị nhân lại; điểm và công số 0 được lưu đúng, gồm điểm ở Tổng hợp công. Kiểm tra thiếu cột ở sheet sau không làm sheet trước được nạp dở, ô công lỗi giữ công đang có và báo nguồn.
- Đối chiếu trực tiếp file được tải lên qua đối số tùy chọn, tạo danh sách kiểm thử từ danh tính nguồn để cô lập lỗi đọc cột, so từng trường công/học/VHTT/phép/BHXH/K3, ăn ca nhà ăn/tiền mặt, M1/M2/M3, điểm sản xuất/chia lương. Tạo workbook tổng hợp tạm để thử hai vòng lưu–mở và tính lại. Không thay tệp nguồn hoặc mang tên/mã người dùng vào mã kiểm thử.
- Lệnh: `TimesheetLayoutTests.exe` dùng dữ liệu giả; `TimesheetLayoutTests.exe <file-cong-Luyen-gang.xlsx>` kiểm tra thêm mẫu thực. Script PowerShell đã thêm phần dữ liệu giả. Nguồn điểm dùng sheet riêng, không thêm public property làm lệch cột snapshot/quỹ quản lý của các bản trước.
- Giữ các kiểm thử tên/năm sinh, Framework/BDĐH, nạp hồ sơ, phân loại và giữ công, quản lý/quỹ/Word, GNV, cẩu phôi và CĐNL. Chạy Linux/Mono, chưa xác nhận trực tiếp Windows.

---

# v5.10 — nhận diện năm sinh, giữ đúng Tên + SDB

- `ImportNameMatchingTests`: 51 kiểm tra dữ liệu giả cho bốn loại sheet (công, ăn ca, BDĐH, điểm), năm sinh trong ngoặc hai chiều, khoảng trắng/NBSP/Unicode tổ hợp, SDB có số 0 đầu, không đổi tên/mã hồ sơ, không tạo người trùng, nạp lặp và các dòng công tách hợp lệ.
- Phân biệt hai hồ sơ cùng tên/mã bằng năm sinh rõ ràng; chặn năm sinh mâu thuẫn, trùng hồ sơ, nhiều năm sinh trong nguồn khi hồ sơ chưa phân biệt được, cùng tên khác mã, cùng mã khác người và chú thích trong ngoặc không phải năm. Không ghi công người sai mã vào người có cùng SDB hợp lệ. Không vượt kiểm tra mã của PX hiện tại bằng cách tự sao chép hồ sơ PX khác.
- Giữ luồng bổ sung hồ sơ người mới thực sự; sao chép hồ sơ PX khác dùng tên/mã chuẩn của hồ sơ nguồn và giữ nguyên công nguồn. Kiểm tra nạp–lưu–mở–nạp lặp và công/phép/K3 ở Tổng hợp công vẫn liên kết bằng danh tính gốc.
- Hai script PowerShell thêm bài kiểm thử mới. Giữ các bài kiểm thử GNV cả 5 PX, nạp nhân sự, BDĐH/Framework, quản lý và lưu công, CĐNL. Tình huống ảnh được tái tạo bằng dữ liệu giả, không chép tên/mã nhân sự trong ảnh vào repo. Chưa có tệp công tạo ra ảnh để xác nhận từng dòng nguồn; chưa chạy trực tiếp Windows.

---

# v5.9 — GNV cả 5 phân xưởng

- `AllWorkshopGnvExportTests`: 697 kiểm tra dữ liệu giả, gồm trực tiếp ô chi tiết/tổng trên 5 sheet NT và Word đủ danh sách; 454 kiểm tra khi chạy trên bản sao Excel nguồn. Các giới hạn và bảo vệ nhóm/vùng ghi xem `RA_SOAT_GNV.md`.
- `CdnlPayrollTests`: thêm kiểm tra âm ở cuối vùng GNV, không cộng lại vào I416, xóa cả giá trị live nguồn. Rà soát lỗi lưu–mở quỹ bằng bộ kiểm thử Excel nguồn; GNV chưa sửa không đè kết quả công thức bằng giá trị tạm từ model.
- Giữ các bài kiểm thử Thép v5.8, phân loại và giữ công v5.7, quỹ quản lý, nạp công, GNV, cẩu phôi, CĐNL, sản lượng và Word. Hai script PowerShell đã thêm bộ kiểm thử cả 5 PX.
- Chạy trên Linux/Mono, biên dịch tham chiếu Microsoft .NET Framework 4.5; chưa kiểm thử trực tiếp Windows. Excel nguồn chỉ được dùng trên bản sao và không đưa lên repo/gói.

---

# v5.8 — GNV Luyện thép không bị cắt ở dòng 80

`SteelGnvExportTests` kiểm tra trực tiếp biên bản NT Thép trong ZIP Excel, không dựa riêng vào snapshot GNV v5. Phủ nhập dòng cuối 97, ghi nguồn chưa sửa, dòng cao trước dòng thấp, giữ mốc cũ/cấp mốc mới không trùng, đủ 38 dòng, công thức tiền và tổng, số 0/điều chỉnh âm, xóa/xóa hết, vượt khuôn giữ nguyên SHA-256 tệp gốc và dọn tệp tạm. Bộ logic kiểm tra thêm nạp GNV đánh dấu danh sách để cập nhật mẫu đích. Chạy cả bản sao Excel nguồn ngoài repo; không sửa tệp đính kèm gốc.

Lệnh: `SteelGnvExportTests.exe` (dữ liệu giả), hoặc `SteelGnvExportTests.exe <bản-Excel-nguồn.xlsx>` (tự tạo bản sao tạm). Hai script kiểm thử PowerShell đã bổ sung bài kiểm tra này. Giữ các kiểm thử hồi quy GNV, nạp công, quản lý, CĐNL, cẩu phôi, Word và phân loại/lưu công của v5.7.

Giới hạn: mẫu Luyện thép hiện cố định 60–97; chưa tự chèn dòng và dịch công thức toàn workbook khi vượt 38 mục. Không coi các bản lỗi nguồn/link ngoài của Excel là kết quả đã duyệt. Biên dịch tham chiếu Microsoft .NET Framework 4.5, chạy Linux/Mono, chưa xác nhận trực tiếp trên Windows.

---

# v5.7 — quản lý theo chức danh, lưu công theo danh tính

- `ClassificationAndSaveTests`: 51 kiểm tra dữ liệu giả; 616 kiểm tra khi chạy thêm trên bản sao Excel cục bộ. Phủ chức danh viết tắt/đầy đủ, cờ quản lý cũ sai, nhóm có cả quản lý/trực tiếp, không ghi đè công giữa PX, lưu công K3/ăn ca/BDĐH, giữ P/Q, xung đột và tính nguyên vẹn của tệp, xóa công về 0, SDB trùng nhưng khác tên, ba vòng lưu–mở và công thức tra cứu/nghiệm thu. Excel nguồn không đưa vào repo. Những ô mã nguồn lỗi sẵn được cô lập trong bộ tính thử khi kiểm tra tra cứu công; dữ liệu trên đĩa không bị sửa để che lỗi.
- Kiểm thử quản lý Excel giữ công nguồn giữa các PX trùng người; tình huống ngày công khác nhau được kiểm tra với dữ liệu giả. Kiểm thử giữ công thức CĐNL loại trừ đúng hai ô tổng hợp số người/số công quản lý D20/D12 được sửa có chủ đích theo chức danh; toàn bộ công thức bộ phận khác vẫn được so sánh.
- Bộ kiểm thử giao diện kiểm tra trưởng ca có cờ quản lý cũ vẫn bị loại và không nhận phân bổ quỹ quản lý. Biên dịch với tham chiếu Microsoft .NET Framework 4.5, chạy Linux/Mono, chưa thử trực tiếp Windows.

---

# Bản hiện hành: v5.4

Hướng dẫn sử dụng/công thức hiện hành nằm trong **README.md**. Các mục v5/v5.1/v5.2 dưới đây là lịch sử; cách nhập tổng quỹ quản lý bằng tay được thay bằng sản lượng × đơn giá quản lý ở v5.3.

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

## Bản v5.2 — xóa sản lượng không giữ thành tiền cũ

Nguyên nhân: mã cũ chỉ tính lại khi cả sản lượng và đơn giá > 0; xóa đầu vào không xóa kết quả cũ. Luyện thép còn thiếu đọc đơn giá loại 2 và ngày lễ từ Excel. Đã thêm tính lại theo sự kiện sửa đầu vào (kể cả trống/0), khóa các ô thành tiền L1/L2/lễ, cập nhật phần chênh vào quỹ SP trực tiếp và phân bổ lại lương cá nhân. Giữ các khoản quỹ không thuộc dòng vừa sửa; không tự tính quỹ quản lý từ sản lượng. Không tự cân đối lại sản lượng các dòng khác khi sửa một dòng. Ngày lễ Luyện gang làm tròn 1.000đ, L1/L2 và ngày lễ Thiêu kết/Luyện thép làm tròn 100đ theo công thức Excel nguồn.

ProductionPayrollTests: 26 kiểm tra mô hình và lưu–mở Excel đạt; 24 kiểm tra mô hình/giao diện đạt, bao gồm xóa sản lượng lễ, khôi phục, đơn giá 0, thành tiền chỉ đọc, quỹ SP và lương từng người về 0, tổng quý và tránh trừ chênh nhiều lần. Các bộ 64 kiểm tra lương/Excel, hồi quy nạp nhân sự và kiểm tra chữ ký Split tương thích .NET Framework tiếp tục đạt. Các kiểm tra có Excel dùng dữ liệu riêng ngoài kho Git, không đưa lên GitHub. Bản v5.2 vẫn biên dịch bằng reference assemblies .NET Framework 4.5 chính thức; chưa chạy trực tiếp trên Windows.

## v5.3 — quỹ quản lý theo sản lượng, Word triệu đồng

Người dùng xác nhận: tách loại 1/loại 2/ngày lễ; nhập đơn giá quản lý L1/L2 và đơn giá lễ tự bằng 2 × đơn giá L1. Thêm công thức quỹ là tổng 3 dòng sản lượng × đơn giá, làm tròn từng dòng đến đồng. Form hiển thị sản lượng nguồn và thành tiền/quỹ chỉ đọc; giá quản lý tách khỏi giá trực tiếp. Áp dụng lưu cấu hình PX; sửa sản lượng sau đó tự cập nhật quỹ và phân bổ. Đơn giá được lưu trong sheet Quy QL san luong; đọc lại chính xác, kể cả 0 và phần thập phân. Nguồn TK/LG/LT giữ độ chính xác sản lượng; CDNL/KT dùng đầu vào ba dòng riêng do ánh xạ cũ không có các dòng tương đương. Quỹ cũ giữ nguyên cho đến khi người dùng nhập đơn giá và áp dụng, không tự suy đơn giá.

Word: chuyển mọi khoản thành tiền, quỹ tháng/quý/lũy kế, bình quân (đúng hệ số nghìn đồng), GNV, BDĐH, chi tiết nhân viên sang triệu đồng, đúng một chữ số thập phân vi-VN; số nguyên và số 0 vẫn có ,0. Đơn giá theo công/đơn vị có nhãn đồng riêng để không làm mất các mức nhỏ. Không đổi số liệu tính toán/Excel, không đổi định dạng số lượng/công, không chuyển đơn vị lần hai khi cập nhật mẫu. Các biên bản lịch sử trước phần đang cập nhật được giữ nguyên.

382 kiểm tra mới đạt: mô hình quỹ/đơn giá lễ tự gấp đôi, sản lượng/đơn giá 0, giá quản lý độc lập trực tiếp, lưu/mở lại hai lần trên 5 PX; xuất mới và cập nhật 50 DOCX giả (5 PX × 5 kỳ × 2 đường xuất), đọc XML kiểm tra đơn vị và giá trị tiền, bảng chi tiết, đổi bình quân từ nghìn đúng hệ số, độ chính xác cố định một số lẻ. Smoke test giao diện bổ sung nhập đơn giá, quỹ chỉ đọc, tự nhân đôi giá lễ, xem trước không ghi, áp dụng lưu, hủy không đổi; đạt. Các bộ kiểm thử cũ tiếp tục đạt. Chưa chạy trực tiếp trên Windows/Microsoft Word. Không đưa dữ liệu hoặc DOCX kiểm thử lên GitHub.

## v5.4 — tính lại thành tiền GNV khi sửa công/đơn giá

Bản cũ chỉ nhân khi thành tiền đang bằng 0. Dòng mới/đã sửa nay tự nhân lại công/khối lượng × đơn giá (làm tròn đồng), kể cả 0; ô thành tiền chỉ đọc. Dòng nguồn chưa sửa giữ số tiền nguồn vì có thể có làm tròn/giảm trừ riêng. Tính GNV trước khi phân bổ lương để không dùng tổng kỳ trước; xóa dòng cuối xóa tổng (CĐNL giữ nguyên quy tắc quỹ riêng, không cộng hai lần). Giữ dòng Excel/giai thích gốc khi sửa, lưu danh sách đã chỉnh trong GNV v5 và khôi phục đúng kể cả danh sách rỗng. Giữ nguyên số âm cho khoản điều chỉnh.

GnvCalculationTests: 18 kiểm tra mô hình/lưu–mở Excel của 5 PX, 15 kiểm tra mô hình/giao diện, gồm sửa tiền cũ, công/giá 0, khôi phục, khoản âm, giữ số nguồn chưa sửa, lưu dòng gốc/ghi chú, xóa dòng cuối và tiền lương phân bổ; đạt. Đã kiểm tra thêm đóng cửa sổ không kích hoạt tính lại trong lúc hủy các điều khiển. Kiểm thử trên Linux/Mono, chưa chạy trực tiếp Windows. Giữ thay đổi v5.3 về quỹ quản lý và tiền Word triệu đồng.

## v5.5 — cẩu phôi PX Thép, giữ riêng công thức CĐNL

- Ba công đoạn có SL thường/lễ và đơn giá riêng. Tiền = thường × giá + lễ × giá × 2; không đưa SL cẩu phôi vào sản lượng quản lý hoặc cộng lễ lần nữa.
- Đối chiếu Excel nguồn: thường 50–52, lễ 46–48. Thay đóng góp cũ của ba dòng trong quỹ, giữ khoản khác J55:J58; ghi lại chi tiết và tổng J45/J49/J44/J35. Đầu vào thường là số đã tách lễ; không giữ công thức trừ lễ cũ ở các ô này. Thành tiền giữ công thức nhân đúng theo đầu vào mới, không dùng ROUND hàng trăm của bản nguồn.
- Thêm kiểm thử CranePayrollTests: 3.446 kiểm tra model/Word/Excel trên bản sao nguồn, gồm ba vòng lưu–mở, xóa về 0, số lẻ, công thức CĐNL không đổi và bảo toàn các khoản ngoài ba công đoạn. 76 kiểm tra ở chế độ giao diện (có lặp model/Word): nhập số Việt Nam, chỉ đọc thành tiền, áp dụng/hủy, tổng nghiệm thu và không cộng quỹ lần nữa khi lưu UI.
- Hồi quy: 64 kiểm tra quản lý và lưu–mở Excel 5 PX; 26 kiểm tra sản lượng với Excel; 382 kiểm tra quản lý theo sản lượng/Word; 18 kiểm tra GNV với Excel; giao diện sản lượng 24, GNV 15; giao diện quản lý xem trước/áp dụng/hủy; kiểm thử nạp công nhân sự cũ/mới/bỏ qua/nạp lặp và lưu–mở Excel. Kiểm tra binary dùng Split tương thích Framework và nhập BDĐH giả đạt.
- Biên dịch EXE 5.5.0.0 bằng tham chiếu .NET Framework 4.5 chính thức từ Microsoft; chạy kiểm tra trên Linux/Mono. Chưa xác minh trực tiếp Windows. Không đưa bản sao Excel, Word sinh ra hoặc dữ liệu nhân sự vào checkout/gói/GitHub.
- `RA_SOAT_CDNL.md` ghi kết quả đọc công thức từng nhóm; không thay công thức CĐNL bằng một phép phân bổ chung. Việc tính lại toàn bộ liên kết Excel CĐNL vẫn cần Excel, như giới hạn được ghi trong tài liệu.

## v5.6 — màn hình và bộ tính CĐNL từng bộ phận

- Thêm menu CĐNL riêng với quỹ 5 nhóm trực tiếp và quản lý, chi tiết công thức/đầu vào, phân bổ từng người, chênh lệch để đối chiếu và danh sách lỗi nguồn. Mặc định đọc công thức AD/AB và đơn giá nhóm đúng từng dòng; quy tắc riêng có nhóm, điểm, đơn giá/ô đơn giá nhóm và lý do. Không suy nhóm từ tên, không phân bổ cả PX theo một công thức.
- CĐNL bỏ qua bộ phân bổ chung cũ. Đọc đúng trường hợp Z là đại lượng số của vận tải. Công/thuộc tính hồ sơ đang nạp được đưa vào bộ tính; giữ nguyên bộ nhận diện/khớp/nạp công. Quản lý theo sản lượng vẫn dùng đơn giá QL riêng; GNV sửa qua danh sách hiện có đi qua I371 một lần.
- Có thể chỉ cập nhật quỹ khi dữ liệu phân bổ chưa đủ. Không xuất phụ lục lương cũ như kết quả mới; khi phân bổ đã áp dụng mà dữ liệu thay đổi gây lỗi, chặn xuất/lưu kết quả cũ. Thiếu công thức AD của người có công không được âm thầm coi là lương 0. Lưu quỹ không biến tiền fallback đã đọc thành đầu vào AD được duyệt.
- Lưu đầu vào đối chiếu/lý do/công thức gốc, quy tắc và trạng thái vào ba sheet CDNL. Bỏ thay thế khôi phục nguồn. Khi thay anchor của công thức shared, mở rộng các công thức con tương đương trước để không mất tham chiếu. Giữ ô trống/chữ/lỗi nguồn chưa thay; ghi số theo định dạng roundtrip để đơn giá và khoản làm tròn không đổi sau lưu–mở.
- 261 kiểm tra CĐNL model/Word/Excel đạt: các quỹ tỷ lệ/tấn/lễ/cung độ/80%/giảm trừ/GNV khác nhau, nguồn AB và đơn giá nhóm, quản lý độc lập, hủy/xem trước/áp dụng, nguồn lỗi/vòng/chia 0, ô trống không được đếm là số 0, lưu–mở đầu vào và quy tắc AJ, khôi phục nguồn, tất cả tiền lương từng người sau lưu–mở và chế độ quản lý.
- Kiểm tra giao diện nguồn thật và dữ liệu giả đạt: mục CĐNL hiện trực tiếp trên menu, 6 quỹ hiện đủ, chọn nhóm gợi ý ô đơn giá tương ứng, nguồn thiếu/lỗi hiện trạng thái chưa tính được, không hiển thị lương 0 giả; cập nhật quỹ riêng, áp dụng/hủy và chỉ đọc tổng nghiệm thu.
- Hồi quy đạt: 64 kiểm tra quản lý/Excel, 382 quản lý theo sản lượng/Word 5 PX × 5 kỳ, GNV Excel, nạp nhân sự/công/lặp/lưu–mở, 3.446 kiểm tra cẩu phôi, giao diện quản lý/sản lượng/GNV/cẩu phôi và kiểm tra binary Framework/Split/BDĐH.
- Excel nguồn thực có tham chiếu lỗi, #N/A, các dòng có công nhưng AD trống và lỗi mã tra cứu. Kiểm thử phân bổ đầy đủ trên bản sao dùng các giá trị giả có nhãn TEST ONLY để kiểm tra luồng, không phải giá trị lương người dùng được xác nhận. Tệp nguồn và bản sao đọc ban đầu giữ nguyên; không đưa Excel hoặc báo cáo chứa nhân sự vào repo/gói. Người dùng phải đối chiếu nguồn trước khi chốt lương thật.
- Biên dịch EXE 5.6.0.0 bằng tham chiếu .NET Framework 4.5 chính thức, kiểm tra trên Linux/Mono. Chưa chạy trực tiếp Windows. Nghiệm Thu Studio và ứng dụng Emga gốc không thay đổi.
