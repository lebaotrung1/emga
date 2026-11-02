# Antigravity v5.13 — thêm người, bổ sung / giảm trừ quản lý có lý do

Tải [antigravity_v5_13_program_only.zip](antigravity_v5_13_program_only.zip). Đóng ứng dụng, giải nén và thay chương trình; giữ **Data**, **Templates** và tệp Excel đang dùng. Tiêu đề **v5.13**, EXE **5.13.0.0**.

- **Thêm người**: chọn phân xưởng, bấm **+ THÊM NGƯỜI** trên thanh menu. Nhập mã SDB, họ tên, tổ/bộ phận, chức danh, lương BHXH và hệ số. Sau đó nhập công tại bảng nhân sự hoặc nạp file công như trước, rồi **Lưu Excel**. Mã SDB trùng trong cùng PX (kể cả khác số 0 đầu mã) bị chặn. Hồ sơ mới lưu trong workbook, không cần tự chèn vào mẫu Excel; dòng XL = 0 là hồ sơ bổ sung và không ghi đè người cũ. Người thêm mới cần công/hệ số trước khi tính phân bổ lương. Cửa sổ lương quản lý cũng có **+ Thêm người quản lý**; người thêm tại đây chỉ ghi nhận khi **Áp dụng**.
- **Bổ sung / Giảm trừ quản lý**: mở **Tính lương quản lý**, nhập bảng **Nguyên nhân | Bổ sung (đ) | Giảm trừ (đ)** phía dưới. Mỗi khoản một dòng, số tiền không âm; khoản có tiền phải có lý do. Chọn dòng và nhấn Delete để xóa. **Tính / xem trước → Áp dụng vào bảng nghiệm thu → Lưu Excel**.
- Tổng quản lý = quỹ SP + ngoài SP + tổng bổ sung − tổng giảm trừ. Các khoản chung không tự chia cho từng cá nhân, không nhập lại vào phụ cấp hoặc tổng ngoài SP. Dòng 2.3 của bảng tổng hợp chuyển sang chỉ đọc sau khi có bảng điều chỉnh chi tiết; sửa qua cửa sổ quản lý để giữ nguyên nhân và tiền nhất quán. Điều chỉnh từ bản cũ được giữ thành dòng “Điều chỉnh đã lưu từ phiên bản trước”; hãy thay bằng lý do thực tế.
- **Excel**: từng khoản được ghi vào phần chi tiết bổ sung/giảm trừ ở cuối sheet NT của PX, có số tiền và lý do; vùng in hiện có được mở rộng để gồm phần chi tiết. Tổng NT đã gồm các khoản này đúng một lần. Sheet **Dieu chinh QL v5** giữ dữ liệu để mở lại, kể cả khi tổng chênh lệch bằng 0 hoặc đã xóa toàn bộ khoản. Giữ các sheet dữ liệu v5 trong workbook.
- **Word / xem trước**: ghi từng khoản bổ sung, giảm trừ và lý do dưới phần lương quản lý, cả báo cáo tháng và quý/lũy kế. Tiền tiếp tục ghi **triệu đồng, một số lẻ**. Sau khi áp dụng, xuất/cập nhật Word để ghi nhận thay đổi.

Kiểm thử công thức, lưu–mở lặp, thêm người không trùng, điều chỉnh dương/âm/bằng 0, xóa khoản, lý do chứa ký tự XML, Word mọi loại kỳ và giao diện trên Linux/Mono với tham chiếu .NET Framework 4.5; chưa chạy trực tiếp Windows. Giữ nạp công, quy tắc quản lý QĐ/PQĐ/KTV/TK, GNV, cẩu phôi và công thức riêng từng bộ phận CĐNL. Gói chỉ chứa chương trình/mã nguồn/tài liệu, không có dữ liệu người dùng.

---

# Antigravity v5.12 — bổ sung lương quản lý cập nhật nghiệm thu

Tải [antigravity_v5_12_program_only.zip](antigravity_v5_12_program_only.zip). Đóng ứng dụng, thay EXE/launcher và giữ **Data**, **Templates**. Tiêu đề hiện **v5.12**, EXE **5.12.0.0**. Kiểm tra trên bản sao Excel; nhập/áp dụng lại khoản chưa được lưu ở bản cũ rồi bấm **Lưu Excel**.

- **Sửa từng người**: mở **Tính lương quản lý**, nhập tiền làm lễ, phụ cấp/công/hệ số, bấm **Tính / xem trước → Áp dụng vào bảng nghiệm thu**. Tổng ngoài SP lấy từ chi tiết cá nhân. Chưa cấu hình đơn giá QL thì giữ quỹ đã nạp/duyệt khi chỉ sửa khoản cá nhân; nhập/đổi đơn giá riêng sẽ chuyển sang quỹ theo sản lượng. Không dùng đơn giá lao động trực tiếp.
- **Bổ sung/giảm trừ chung**: nhập dòng **2.3. Điều chỉnh tiền lương QL** trên bảng tổng hợp hoặc ô điều chỉnh trong cửa sổ quản lý. Khoản chung cộng/trừ một lần ở tổng, không tự phân bổ vào từng người.
- Dòng **2.2. Ngoài SP QL (chế độ/bổ sung)** nhận **tổng ngoài SP đã duyệt**, thay tổng tự cộng từ cá nhân; đây không phải khoản tăng thêm. Nhập số tăng/giảm riêng ở dòng 2.3. Nếu sau đó áp dụng lại chi tiết trong cửa sổ quản lý, tổng ngoài SP sẽ chuyển lại theo các khoản cá nhân; điều chỉnh chung vẫn giữ.
- Tổng quản lý là ô tự tính. Quỹ SP theo sản lượng sửa qua đơn giá riêng; bảng tổng hợp không cho ghi đè vào quỹ đang tự tính. Tổng tháng cập nhật ngay cả khi tắt tự động lũy kế quý/năm.

Bản cũ có thể ghi số mới vào bộ nhớ đệm Excel nhưng vẫn giữ công thức lương quản lý cũ, nên tính lại sẽ mất khoản bổ sung. Bản này liên kết các tổng NT và tiền lương cá nhân đã áp dụng với sheet **Luong quan ly v5**. Giữ sheet này cùng các cấu hình trong workbook; thay đổi lương bằng ứng dụng và lưu lại để đồng bộ. Các mẫu NT cũ có nhiều tiểu mục riêng, nên dùng sheet này để kiểm tra đầy đủ chi tiết khoản hiện đang áp dụng.

Đã kiểm tra tổng quản lý/tổng PX bằng tính lại công thức Excel trên cả 5 mẫu, khoản làm lễ/phụ cấp, điều chỉnh dương/âm/0, lưu–mở lặp, dữ liệu nhập ngay ở bảng tổng hợp và giữ công thức dùng chung của lao động trực tiếp. CĐNL giữ các công thức bộ phận riêng; điều chỉnh quản lý có trong cả kết quả tính và tổng, không cộng lại lần hai. Giữ nạp công, nhận diện năm sinh, GNV, cẩu phôi, điểm nguồn và Word triệu đồng một số lẻ. Gói không có dữ liệu người dùng. Kiểm thử Linux/Mono với tham chiếu .NET Framework 4.5, chưa chạy trực tiếp Windows.

---

# Antigravity v5.11 — sửa nạp công về 0 ở mẫu Luyện gang

Tải [antigravity_v5_11_program_only.zip](antigravity_v5_11_program_only.zip). Đóng ứng dụng, thay EXE/launcher và giữ **Data**, **Templates**. Tiêu đề hiện **v5.11**, EXE **5.11.0.0**. Chọn đúng phân xưởng/kỳ tháng, nạp lại file công gốc rồi **Lưu Excel** trên bản sao để kiểm tra. Dữ liệu công đọc sai ở bản cũ được khôi phục bằng cách nạp lại nguồn còn đầy đủ; không tự đoán công đã mất.

Mẫu Luyện gang dùng cột **AJ “Tổng công”** trong phần “Quy ra công” cho công làm việc, thay tên “Số công sản phẩm”. Bản cũ không xác định được cột này nhưng vẫn nạp giá trị 0. Bản mới nhận đúng cột, giữ riêng học/họp, VHTT, phép, BHXH, K3; cột tổng ngày lịch không bị coi là công sản phẩm. Nhận các tên **“Tổng điểm” / “Tổng điểm chia lương”**, đúng vùng ngày ở tiêu đề gộp ô và tên tổng BDĐH khác nhau. Chú giải mức độc hại dưới cuối bảng không được coi là nhân sự mới.

Nếu chưa nhận diện được cột công/điểm bắt buộc, chương trình báo trước khi nạp để giữ nguyên dữ liệu. Ô công thiếu giá trị số hoặc có lỗi được báo kèm địa chỉ, giữ công đang có của người đó; số **0 hợp lệ** vẫn được nạp để xóa công cũ.

Điểm chia lương trực tiếp đã có trong file được giữ theo nguồn, không nhân lại xếp loại chung. Sheet **Diem nap v5** giữ dữ liệu này qua lưu–mở và tính lại, bao gồm điểm bằng 0. Giữ sheet này cùng các sheet cấu hình khác. Muốn thay điểm đã nạp, sửa/nạp lại bằng ứng dụng để đồng bộ; chỉ sửa ô điểm trong Excel có thể bị dữ liệu đã lưu ưu tiên khi mở ứng dụng. Công thức quản lý và bộ tính CĐNL vẫn theo chế độ riêng; lựa chọn ăn ca Công thức/Đơn vị/Thủ công được giữ.

Đã đối chiếu file công thực tế theo từng người và từng trường, nạp lặp, chuyển sang bản Excel tổng hợp tạm rồi lưu–mở hai vòng; kiểm tra công và điểm không trở về số cũ/0. Dữ liệu thật chỉ dùng trong bản tạm ngoài repo, không coi bài kiểm thử là kết quả lương được duyệt. Giữ các sửa đổi tên có năm sinh, quản lý, GNV cả 5 PX, cẩu phôi, CĐNL và Word. Gói không có Excel/ảnh/dữ liệu người dùng. Kiểm thử Linux/Mono với tham chiếu .NET Framework 4.5, chưa chạy trực tiếp Windows.

---

# Antigravity v5.10 — nhận diện tên có năm sinh khi nạp công

Tải [antigravity_v5_10_program_only.zip](antigravity_v5_10_program_only.zip). Đóng ứng dụng, thay EXE/launcher, giữ **Data** và **Templates**. Tiêu đề hiện **v5.10**, EXE **5.10.0.0**. Chọn đúng PX, nạp lại bảng công/ăn ca/BDĐH/điểm rồi **Lưu Excel** trên bản sao để kiểm tra.

Bản cũ so nguyên văn họ tên nên tên có hậu tố `(1989)` trong file có thể không khớp hồ sơ không ghi năm sinh. Bản mới đối chiếu **tên + SDB**, hỗ trợ hậu tố năm sinh bốn chữ số ở cuối tên trong ngoặc, khoảng trắng và cách mã hóa dấu tiếng Việt. Chỉ bỏ qua phần năm sinh khi xác định được duy nhất một hồ sơ phù hợp; giữ nguyên tên/mã đã có trong hồ sơ, không thêm người trùng chỉ vì khác cách ghi năm sinh. Nếu hai bên đều có năm sinh thì phải tương thích. Hai hồ sơ cùng tên/mã cần năm sinh đủ để phân biệt; không chọn tùy ý hồ sơ đầu tiên.

**Khác SDB vẫn không tự ghép.** Nhật ký sẽ nói rõ “tên có trong danh sách nhưng SDB khác” hoặc “SDB đang thuộc tên khác” để đối chiếu. Cần sửa mã ở file nguồn hoặc danh sách gốc cho đúng và thống nhất rồi nạp lại. Không suy mã bị đảo chữ số, không bỏ dấu để đoán tên, không bỏ mọi nội dung trong ngoặc. Nếu nhiều sheet ghi các năm sinh khác nhau cho cùng tên/mã nhưng hồ sơ không đủ thông tin, chương trình báo để tránh nhập chung công của hai người.

Quy tắc áp dụng cho **công, ăn ca, BDĐH và điểm**, kể cả sao chép hồ sơ hợp lệ từ PX khác. Người mới thực sự vẫn có bước bổ sung hồ sơ như trước. Giữ các sửa đổi quản lý, lưu công, GNV cả 5 PX, cẩu phôi, CĐNL và Word. Đã kiểm tra bằng dữ liệu giả, gồm nạp lặp và lưu–mở Excel; ảnh lỗi dùng để xác định tình huống, chưa có tệp bảng công tương ứng để xác nhận từng dòng thực tế. Không đưa ảnh/danh sách nhân sự/Excel người dùng lên GitHub. Kiểm thử Linux/Mono với tham chiếu .NET Framework 4.5, chưa chạy trực tiếp Windows.

---

# Antigravity v5.9 — kiểm tra GNV cả 5 phân xưởng

Tải [antigravity_v5_9_program_only.zip](antigravity_v5_9_program_only.zip). Đóng ứng dụng, thay EXE/launcher, giữ **Data** và **Templates**. Tiêu đề ứng dụng hiện **v5.9**, EXE **5.9.0.0**. Mở bản sao Excel còn danh sách GNV đầy đủ hoặc nạp lại GNV rồi **Lưu Excel**.

Đã sửa giới hạn ghi GNV Thiêu kết/Luyện gang, bổ sung ghi chi tiết Khai tuyển/CĐNL, thống nhất vùng đọc–ghi theo từng mẫu. Không đọc dòng tổng phụ thành một khoản GNV rồi cộng trùng. Khoản chuyển thanh toán âm được giữ; thêm/xóa/sửa về 0 được ghi cả vào chi tiết và tổng trong biên bản Excel. Word giữ đủ các mục trong danh sách dài.

| Phân xưởng | Sức chứa chi tiết của mẫu hiện tại |
| --- | ---: |
| Thiêu kết | 11 dòng |
| Luyện gang | 38 dòng của PX + 2 dòng chuyển thanh toán PX khác |
| Luyện thép | 38 dòng |
| CĐNL | 22 dòng trong quỹ GNV sửa chữa nguồn |
| Khai tuyển | 36 dòng |

**Vượt khuôn sẽ báo rõ và giữ nguyên tệp gốc**, không ghi một phần rồi báo thành công. Không tự chèn dòng/dịch công thức toàn workbook. Mẫu chèn thêm dòng hoặc CĐNL có công/tiền ngoài vùng quỹ GNV gốc cần được đối chiếu trước; chương trình báo thay vì tự bỏ qua hoặc cộng vào quỹ khác. Xem [rà soát từng xưởng](RA_SOAT_GNV.md) để biết vùng và quy tắc tổng riêng.

Giữ các sửa đổi quản lý, công 171/177, nạp công/ăn ca/BDĐH, cẩu phôi, quỹ CĐNL, GNV và Word triệu đồng một số lẻ. Gói chỉ có chương trình/mã nguồn/hướng dẫn/kiểm thử; không có dữ liệu người dùng. Kiểm thử trên Linux/Mono với tham chiếu .NET Framework 4.5 chính thức; chưa chạy trực tiếp Windows.

---

# Antigravity v5.8 — đổ đầy đủ GNV Luyện thép vào Excel

Tải [antigravity_v5_8_program_only.zip](antigravity_v5_8_program_only.zip). Đóng ứng dụng, thay EXE/launcher trong thư mục đang dùng và giữ **Data**, **Templates**. Tiêu đề ứng dụng hiện **v5.8**, EXE **5.8.0.0**. Mở bản sao Excel, nạp lại danh sách GNV đầy đủ rồi lưu để cập nhật những dòng chưa được ghi bởi bản cũ.

Lỗi cũ: vùng GNV Luyện thép trong mẫu là **60–97**, nhưng hàm ghi chỉ chạy **61–80** và dùng `break` khi gặp mục ở dòng cao hơn. Vì vậy mục ở dòng 92 và cả các mục phía sau nó có thể không được ghi; bộ đọc còn bỏ dòng 97.

Bản mới đọc/ghi hết vùng này, giữ dòng gốc hợp lệ, cấp riêng ô trống cho dòng mới hoặc số dòng trùng/không hợp lệ. Khi danh sách được chỉnh sửa/nạp lại, các dòng đã xóa được xóa trong biên bản, dòng số 0 vẫn được lưu, và tổng GNV lấy từ toàn bộ vùng chi tiết. Dòng tự tính dùng khối lượng × đơn giá, làm tròn đến đồng; khoản nguồn đặc biệt chưa sửa vẫn giữ cách tính nguồn. Phần lương chế độ từ dòng 98 không được dùng làm chỗ ghi GNV.

Mẫu hiện hỗ trợ **38 dòng GNV Luyện thép**. Nếu vượt, chương trình báo số dòng và dừng toàn bộ lần lưu, giữ nguyên Excel gốc; không lưu một phần rồi báo thành công. Việc mở rộng/chèn dòng của mẫu vượt khuôn hiện tại chưa được hỗ trợ tự động.

Đã kiểm tra trực tiếp ô chi tiết và công thức tổng Excel, không chỉ sheet lưu nội bộ: đủ 38 mục, mục sau dòng 80, dòng 97, thứ tự đảo, trùng dòng, thêm/xóa/xóa hết, công bằng 0, điều chỉnh âm, nạp GNV, lưu–mở lại và lỗi vượt khuôn không thay tệp gốc. Kiểm thử dùng dữ liệu giả và bản sao Excel cục bộ; không đưa dữ liệu người dùng vào gói hoặc GitHub. Chạy kiểm thử Linux/Mono với tham chiếu .NET Framework 4.5, chưa chạy trực tiếp Windows. Giữ bản sửa quản lý, công 171/177, cẩu phôi, CĐNL, GNV và Word trước đây.

---

# Antigravity v5.7 — phân loại quản lý và giữ công khi lưu Excel

Tải [antigravity_v5_7_program_only.zip](antigravity_v5_7_program_only.zip). Đóng ứng dụng, giải nén và thay EXE/launcher trong thư mục đang dùng. Giữ **Data**, **Templates**; mở một bản sao Excel để đối chiếu trước. Tiêu đề ứng dụng hiện **v5.7**; EXE **5.7.0.0**.

- Quản lý chỉ gồm **quản đốc, phó quản đốc, KTV, TK** (hỗ trợ chữ viết tắt trong Excel). **Tổ trưởng, trưởng ca, phó ca, tổ phó** là lao động trực tiếp. Chương trình xét chức danh, không suy từ tên tổ hay cờ quản lý của bản cũ. Bảng nhân sự có cột **Phân loại LĐ**; cửa sổ tính lương quản lý chỉ nhận các chức danh hợp lệ. Số người và số công quản lý trong nghiệm thu được cập nhật cùng phân loại.
- Sửa lỗi dòng trống công của một PX ghi đè công đang có ở PX khác. Đối chiếu bằng **SDB + họ tên**, ghi Tổng hợp công một lần. Trường hợp trùng dòng tổng hợp dùng vùng INDEX/MATCH gốc để xác định, không chọn tùy ý dòng đầu. Nếu hai PX có công khác nhau hoặc không xác định được dòng nguồn, dừng lưu và báo để đối chiếu; không tự cộng công.
- Sửa công về 0 vẫn được lưu. Giữ các công độc lập ở cột P/Q trong tổng công khác; lưu công K3 vào cột riêng. Lưu trên tệp tạm rồi thay tệp chính sau khi hoàn tất, để lỗi giữa chừng không làm hỏng Excel gốc.
- Giữ nạp công/ăn ca/BDĐH, cẩu phôi ba công đoạn, đơn giá quản lý riêng, GNV và Word triệu đồng một số lẻ. Không thay Nghiệm Thu Studio. Gói không có Excel hoặc dữ liệu người dùng.

Đã kiểm tra phân loại, dữ liệu công trùng giữa PX, xung đột, sửa về 0, lưu–mở lặp và giao diện trên Linux/Mono với tham chiếu .NET Framework 4.5; chưa chạy trực tiếp trên Windows. Tệp Excel nguồn có một số ô lỗi sẵn: kiểm tra tra cứu công được cô lập khỏi các ô mã lỗi trong bộ tính thử, không sửa các ô lỗi đó trong dữ liệu của người dùng. Nếu tệp đã bị mất công ở bản cũ, cần mở lại bản sao còn đủ công hoặc nạp lại bảng công; bản mới không thể tự đoán lại số công đã mất.

---

# Antigravity v5.6 — tính lương CĐNL từng bộ phận

Tải **antigravity_v5_6_program_only.zip** trên nhánh `antigravity-payroll-v5`. Đóng bản cũ, sao lưu Excel, giải nén rồi thay EXE/launcher; giữ nguyên Data và Templates. Tiêu đề v5.6, phiên bản EXE 5.6.0.0. Không chứa dữ liệu người dùng.

## Vào đâu để tính CĐNL?

1. Chọn **Phân xưởng Cơ điện Năng lượng**, nạp Excel như trước.
2. Trên menu trên cùng, bấm **TÍNH LƯƠNG CĐNL TỪNG BỘ PHẬN (v5.6)**.
3. Tab **Quỹ từng bộ phận** hiển thị Trưởng ca, Vận hành năng lượng, Vận tải/cẩu/xúc lật, Sàng quặng, Sửa chữa cơ điện và Quản lý. Chọn một nhóm để xem các dòng thành phần, số tiền và công thức đang dùng.
4. Tab **Đầu vào / ô cần đối chiếu** hiển thị các đầu vào thực sự được công thức sử dụng. Nhập giá trị thay thế, lý do và đánh dấu áp dụng. Có thể thêm ô theo địa chỉ `Sheet!A1`. Bỏ chọn để trở về giá trị/công thức gốc. Công và mức BHXH lấy từ bảng công/hồ sơ đang nạp; sửa tại đó. Điểm hưởng theo AB nguồn hoặc giá trị đối chiếu nhập tại tab này, không tự đổi tất cả theo xếp loại.
5. Tab **Phân bổ từng người** xem lương SP và tổng lương. Mặc định đọc đúng công thức AD/AB của từng dòng và các ô đơn giá nhóm AJ146/AJ152/AJ158/AJ168. Một số dòng vận tải dùng Z là đại lượng số, không phải A/B/C; không áp xếp loại chung. Khi cần quy tắc riêng, chọn bộ phận, nhập điểm và lý do; ô đơn giá tự gợi ý AJ146/AJ152/AJ158/AJ168 theo nhóm khi có nguồn. Có thể dùng ô đơn giá nhóm này hoặc nhập đơn giá điểm riêng bằng số. Đây là lựa chọn chủ động, không tự đoán nhóm theo tên tổ. Chênh lệch quỹ/tiền đã phân bổ luôn được hiển thị để đối chiếu, không dồn vào người cuối.
6. Bấm **Tính / xem trước**, kiểm tra tab **Lỗi cần xử lý**, rồi **Áp dụng vào nghiệm thu**. Hủy không ghi thay đổi. Sau khi áp dụng, bấm **Lưu Excel** ở màn hình chính. Các số tiền tổng hợp tháng CĐNL đã tính là ô chỉ đọc; sửa đầu vào trong màn hình CĐNL để tính lại.

Bỏ chọn **Tính lại lương từng người** nếu chỉ muốn tính quỹ. Khi đó chưa tính lại phân bổ trực tiếp và ghi rõ trạng thái này; quản lý đã bật chế độ riêng vẫn theo chế độ đó; không cho đính kèm bảng lương cá nhân vào Word như kết quả mới. Hồ sơ mới chưa có công thức nguồn phải khai báo quy tắc phân bổ riêng; nếu có công chế độ/lễ/học nhưng chưa có dòng công thức tương ứng, phần mềm yêu cầu nạp công thức đầy đủ, không tự đoán mức tiền.

## Các cách tính được giữ riêng

Quỹ từng bộ phận được tính từ chính công thức `NT CĐNL`: tỷ lệ quỹ các PX, công/tấn/chuyến chia định mức, hệ số lễ/80%, khoản trừ/thu hồi và tổng các GNV theo đúng vùng tham chiếu. Chế độ, phụ cấp và tiền lễ của từng người cũng theo công thức dòng nguồn. Không lấy tổng quỹ CĐNL chia cho mọi nhân viên bằng cùng một công thức. Bộ phân bổ chung cũ được bỏ qua riêng cho CĐNL, kể cả khi chưa áp dụng công cụ mới.

Sửa quỹ/sản lượng Thép, Gang, Thiêu kết sẽ cập nhật cơ sở liên kết khi tính lại CĐNL, xem trước/xuất Word hoặc lưu Excel. GNV đã chỉnh qua danh sách hiện có chỉ đi vào quỹ sửa chữa qua I371 một lần. Chế độ quản lý tính từ sản lượng và đơn giá QL riêng vẫn dùng màn hình quản lý hiện có; khi bật chế độ đó, phần CĐNL dùng kết quả quản lý này, không kéo lại công thức quản lý cũ.

Word xem trước, xuất mới và cập nhật mẫu đều ghi quỹ từng nhóm theo **triệu đồng với đúng một số lẻ**. Chi tiết quỹ nhóm ghi rõ là tháng; tổng quý/lũy kế theo nền hiện có. Nạp công và tính cẩu phôi v5.5 được giữ.

## Ô nguồn lỗi trong Excel đã gửi

Tệp hiện có các điểm làm phân bổ đầy đủ chưa tính được: `THop (có tạm ứng)!I51`, `CĐNL!AH5` chứa tham chiếu mất; `CĐNL!AB98:AB101` có `#N/A`. Ngoài ra có các dòng có công nhưng AD trống và lỗi tra cứu mã ở nguồn liên kết. Không biến AD trống thành lương 0 hoặc tự lấy tiền fallback cũ làm lương mới khi lưu quỹ. Màn hình chỉ rõ từng ô/hồ sơ; có thể khai báo quy tắc riêng ở tab Phân bổ hoặc đối chiếu giá trị/công thức nguồn. Cần đối chiếu tài liệu nghiệp vụ để nhập giá trị thay thế cùng lý do hoặc nạp nguồn đã sửa. **Không tự điền 0, không lấy số cũ làm kết quả tính mới.** Chỉ tính quỹ vẫn thực hiện được. Đây là lỗi/thiếu dữ liệu nguồn, không phải một công thức chung có thể áp thay cho các ô đó.

Bộ tính hỗ trợ các hàm số đang dùng trong đường tính được kiểm tra (SUM/ROUND/IF/IFERROR/COUNT/COUNTA/LEN/MIN/MAX/ABS/SUMIF/COUNTIF/INDEX và MATCH chính xác), tham chiếu liên sheet và công thức dùng chung. Không chạy macro hoặc truy cập liên kết ngoài. Công thức ngoài phạm vi hỗ trợ, lỗi nguồn, chia 0 và tham chiếu vòng được báo rõ, không dùng cache để giả lập thành công. Đây không phải bộ máy Excel đầy đủ.

## Lưu và khôi phục

Các sheet **CDNL dau vao**, **CDNL phan bo**, **CDNL ket qua** lưu đầu vào được đối chiếu/lý do, quy tắc riêng, trạng thái áp dụng và quỹ nhóm. Giữ các sheet này khi sao chép dữ liệu. Khi thay một ô bằng giá trị đã duyệt, công thức/giá trị gốc được lưu lại; ô Excel tham chiếu tới đầu vào đã đối chiếu để không chạy lại tham chiếu lỗi khi mở Excel. Bỏ chọn thay thế rồi lưu sẽ khôi phục nguồn gốc. Quy tắc phân bổ riêng và chế độ quản lý mới cũng giữ lại công thức gốc trước khi thay. Dữ liệu này nằm trong Excel cục bộ, không nằm trong chương trình hay GitHub.

Không chỉnh trực tiếp các sheet cấu hình nếu chưa hiểu chế độ đang áp dụng; dùng màn hình tương ứng để tính lại/lưu. Excel có thể tính lại các phần ngoài phạm vi bộ tính ứng dụng; cần đối chiếu một kỳ đã duyệt trước khi thanh toán.

`Build.ps1` biên dịch Windows; `Kiem_Thu_Logic_V5.ps1` chạy dữ liệu giả. Kiểm thử model, giao diện, Word, đầu vào đối chiếu và lưu–mở Excel thực hiện trên Linux/Mono; chưa chạy trực tiếp Windows. Các kiểm thử cần Excel riêng dùng bản sao ngoài checkout, không đóng gói dữ liệu nguồn hoặc kết quả có nhân sự. Không sửa Nghiệm Thu Studio.

---

## Hướng dẫn các bản trước

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
