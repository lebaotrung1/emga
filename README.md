# Emga HR — bản chạy thử 0.2.1

Ứng dụng nhân sự tiếng Việt cho Windows, sử dụng chung qua trình duyệt trong LAN. Bản 0.2 bổ sung **hồ sơ nhân sự, phân nhóm tự động, quyết định có ngày hiệu lực và báo cáo tự tổng hợp**. Chưa thay thế đầy đủ các mẫu báo cáo chính thức của đơn vị và chưa tính tiền lương. Mã nguồn không chứa tệp Excel hoặc dữ liệu nhân sự người dùng.

## Sửa trường hợp báo cáo ra toàn số 0 (0.2.1)

Báo cáo mở ra sẽ tự tổng hợp kỳ đang chọn và hiển thị tình trạng dữ liệu: chưa chuyển Excel thành hồ sơ, hồ sơ còn chờ xác nhận, kỳ trước ngày hồ sơ có hiệu lực, hoặc thực sự không còn ai đang làm. Mỗi trường hợp có hướng dẫn/nút chuyển tới bước xử lý. Nếu API lỗi thì hiển thị lỗi tại khu vực báo cáo và ẩn kết quả cũ.

Nếu đã nhập hồ sơ nhưng báo cáo vẫn 0: vào **Hồ sơ nhân sự**, xem đúng ngày, chọn những người đã rà soát, đặt trạng thái **Đang làm việc**, chọn ngày hiệu lực đúng với thực tế và bấm **Ghi nhận các hồ sơ đã chọn**. Sau đó mở lại **Báo cáo tự tổng hợp** với ngày cuối kỳ từ ngày hiệu lực trở đi. Không chọn tất cả là đang làm nếu chưa kiểm tra người đã nghỉ. Nếu mới tải Excel nguồn, dùng nút **Mở bước nhập hồ sơ từ Excel** để thực hiện bước chuyển thành hồ sơ trước.

Khi nâng cấp, dừng máy chủ bản cũ, chạy bản mới rồi bấm Ctrl+F5 trên trình duyệt. Nhãn góc trên phải là **0.2.1**. Tệp JavaScript/CSS có mã phiên bản mới để tránh dùng nhầm bản lưu trong trình duyệt.

## Cài mới trên Windows

1. Cài Python 3.12 trở lên từ python.org (có Python Launcher `py`). Tải ZIP từ GitHub và giải nén vào thư mục có quyền ghi.
2. Bấm `Chay_Emga.bat`. Lần đầu cần Internet để tải thư viện.
3. Trên chính máy chủ, mở trình duyệt tới địa chỉ loopback `127.0.0.1`, cổng **8765**. Tự tạo tài khoản quản trị đầu tiên, mật khẩu từ 12 ký tự. Không có tài khoản/mật khẩu mặc định.
4. Tạo tài khoản cho đồng nghiệp trong **Quản trị & sao lưu**: chỉ xem/xuất, biên tập hoặc quản trị.
5. Khi dùng LAN, chạy `ipconfig` trên máy chủ để xem IPv4. Máy khác dùng địa chỉ IP máy chủ, cổng 8765. Nếu Firewall chặn, cho phép Python/cổng TCP 8765 trên mạng **Private**, giới hạn máy trong LAN. Không mở cổng router ra Internet. Máy chủ và cửa sổ chạy phần mềm phải hoạt động.

Không có bộ cài `.exe`. Đã kiểm thử bằng Chromium và Waitress trên Linux, chưa kiểm thử trực tiếp Windows hoặc mạng LAN Windows thực tế. Máy trạm chỉ cần trình duyệt, không cần cài Python.

## Nâng cấp từ 0.1 — giữ tài khoản và dữ liệu

1. Trong bản cũ, tải **Sao lưu** và cất ở nơi an toàn. Dừng máy chủ cũ bằng Ctrl+C.
2. Giải nén bản mới ra thư mục khác, bấm `Chay_Emga.bat`.
3. Nếu vẫn dùng cùng tài khoản Windows và cùng `EMGA_DATA_DIR` (nếu đã cấu hình), bản mới tự đọc lại tài khoản, Excel, chỉnh sửa và mẫu Word. Cơ sở dữ liệu được bổ sung các bảng mới, không xóa bảng cũ. Không chạy đồng thời hai bản cùng cổng/thư mục dữ liệu.
4. Hồ sơ quản lý mới chưa tự lấy từ các tệp Excel cũ: thực hiện bước nhập có xem trước ở dưới. Tài khoản cũ vẫn đăng nhập như trước.

## Quy trình sử dụng

### 1. Chuẩn bị dữ liệu ban đầu

- Vào **Excel nguồn**, nhập/chọn tệp `.xlsx`.
- Vào **Hồ sơ nhân sự**, bấm **Nhập từ Excel đang chọn**. Xem trước họ tên, mã, ngày sinh, trình độ dự kiến, chức danh và mã đơn vị. Chọn ngày chốt số liệu thực sự đã biết rồi bấm nhập.
- Tệp chuẩn cần trang `DU LIEU CO SO`: mã E, tên F, sinh G, vào làm M, chức danh B; mã đơn vị ưu tiên GX (cột 206), nếu trống dùng N. Cột đánh dấu trình độ: AY/51 đại học, BZ/78 cao đẳng, DA/105 trung cấp, EB/132 CNKT, FC/159 chưa đào tạo. Đây là ánh xạ riêng cho tệp đã cung cấp, không áp dụng tùy ý cho tệp có cấu trúc khác.
- Không tự suy giới tính từ họ tên hay trạng thái đang làm từ công thức cũ. Tất cả hồ sơ nhập lần đầu ở trạng thái **Cần rà soát**, chưa tính vào số đang làm. Dữ liệu ngày không đủ ngày/tháng/năm để trống; trình độ không xác định/đánh dấu nhiều nhóm ghi Chưa xác định. Cần kiểm tra lại trên bản gốc, đặc biệt phân biệt đại học/trên đại học.
- Mã trùng, thiếu mã, tên không hợp lệ hoặc dòng đã nhập bị bỏ qua; mọi dòng cùng mã trùng trong tệp đều cần rà soát, không tùy ý chọn dòng đầu; không tự gộp theo tên. Mã nhân sự là duy nhất trong toàn bộ hồ sơ quản lý. Tên nguồn không tự đồng nhất với hồ sơ khác.
- Sửa từng hồ sơ hoặc đánh dấu nhiều người đã rà soát để cập nhật trạng thái, trình độ/vị trí cùng ngày hiệu lực. Ghi nhận hàng loạt có kiểm tra phiên bản và thực hiện trọn giao dịch: có xung đột thì không ghi nửa chừng.

### 2. Khai báo danh mục phân nhóm

Trong **Danh mục phân nhóm**, tạo vị trí gồm tên chức danh và đơn vị; quy tắc của vị trí gồm nhóm **Quản lý / Công nghệ / Phụ trợ / Phục vụ**, có thuộc CNKT hay không, bậc thợ 1–7 nếu có và ngày hiệu lực. Vị trí nhập từ Excel ban đầu chưa có quy tắc; cần khai báo theo nghiệp vụ của đơn vị.

**Chọn vị trí khi thêm người mới hoặc điều động là đủ để tự nhận nhóm lao động/CNKT/bậc của vị trí.** Phần mềm không đoán nhóm từ chữ trong chức danh. Nếu cùng chức danh có nhiều bậc, khai báo vị trí riêng cho từng bậc (ví dụ “Công nhân vận hành bậc 3”). Bậc ở bản này là thuộc tính mặc định của vị trí, chưa có bậc cá nhân độc lập.

Quy tắc có lịch sử theo ngày. Quy tắc mới phải sau quy tắc gần nhất, không ghi đè quy tắc cũ. Nếu cần sửa nhầm ngày/quy tắc cũ, chưa có chức năng đảo hoặc hủy; hãy đối chiếu và sao lưu trước khi xác nhận. Đổi quy tắc sẽ tác động mọi người giữ vị trí đó kể từ ngày hiệu lực của quy tắc, bao gồm các bản báo cáo tạo lại cho khoảng thời gian đó.

### 3. Thêm, sửa hồ sơ nhân sự

Bấm **Thêm nhân sự**, nhập mã, họ tên, ngày sinh, vào làm, giới tính, trình độ, vị trí, trạng thái và ngày ghi nhận có hiệu lực. Mã là văn bản để giữ số 0 đầu, không đổi mã trực tiếp sau khi tạo.

- **Nhóm tuổi tự tính theo ngày xem/ngày cuối kỳ**, không lưu tuổi cố định; tự chuyển nhóm khi đến sinh nhật. Theo cách tính hiện tại, sinh 29/2 được tăng tuổi vào 1/3 trong năm không nhuận.
- **Trình độ** là một lựa chọn chuẩn hóa được nhập/cập nhật trong hồ sơ; báo cáo tự đếm theo lựa chọn đó. Không đọc/đoán trình độ từ tên văn bằng tự do.
- **Nhóm lao động/CNKT/bậc** lấy từ quy tắc vị trí có hiệu lực ở ngày xem.
- Thay đổi lưu thành sự kiện, không ghi đè ảnh chụp hồ sơ cũ. Lịch sử hiển thị ở nút **Lịch sử**.
- Chọn ngày xem hồ sơ ở đầu trang để xem phân nhóm tại ngày đó. Tải lại khi đồng nghiệp vừa chỉnh sửa. Không ghi biến động có ngày trước biến động gần nhất; cần chọn đúng trình tự ngày. Nếu đã có quyết định tương lai, không thể chèn biến động sớm hơn trong bản thử này.

Hồ sơ mới trạng thái Đang làm cần ngày sinh, ngày vào làm và vị trí. Hồ sơ còn thiếu thông tin nên để Cần rà soát. Khi xác nhận các hồ sơ nhập cũ dù còn thiếu thuộc tính, báo cáo vẫn cảnh báo các thông tin chưa rõ và có nhóm Chưa xác định.

### 4. Quyết định tự cập nhật hồ sơ

Dùng **Quyết định & biến động** với hồ sơ đã xác nhận đang làm:

1. Chọn nhân sự, loại quyết định, vị trí mới (điều động/đổi chức danh), ngày hiệu lực, số, căn cứ/nội dung và người ký. Bấm **Lưu & tải Word dự thảo**.
2. Kiểm tra văn bản, thực hiện phê duyệt theo quy trình đơn vị.
3. Bấm **Áp dụng quyết định đã duyệt** cho đúng dự thảo và xác nhận. Phần mềm lưu biến động một lần. Nếu hồ sơ đã đổi từ lúc lập dự thảo, phải lập lại để tránh ghi đè.
4. Từ ngày hiệu lực, hồ sơ và báo cáo tự dùng vị trí/nhóm mới; trước ngày đó dùng vị trí cũ. Ngày tương lai được ghi lịch, không tác động số liệu trước ngày hiệu lực. Chấm dứt loại khỏi danh sách đang làm từ chính ngày hiệu lực.

Tải Word không tự phê duyệt quyết định. Bấm áp dụng là xác nhận của người dùng rằng quyết định đã được duyệt; chưa có quy trình ký số hoặc xét duyệt nhiều cấp. Chưa có hủy/đảo quyết định đã áp dụng. Quyết định không tự đổi ngày sinh hay trình độ: các dữ kiện đó được cập nhật trong hồ sơ.

Mục **Word từ Excel cũ** vẫn giữ chức năng cũ, chỉ xuất Word từ tệp nguồn và không cập nhật hồ sơ quản lý. Muốn cập nhật tự động phải dùng **Quyết định & biến động**.

### 5. Báo cáo tự tổng hợp

Vào **Báo cáo tự tổng hợp**, chọn từ/đến ngày (tháng, quý hoặc khoảng tùy chọn), bấm **Tổng hợp**:

- Số đang làm, nữ, CNKT tại ngày cuối kỳ.
- Phân nhóm lao động, trình độ, tuổi, bậc CNKT.
- Đầu kỳ / tăng-vào / giảm-ra / cuối kỳ theo đơn vị. Điều động nội bộ chuyển giữa hai đơn vị, không làm tăng tổng người. Tăng/giảm gồm cả xác nhận và điều chỉnh trạng thái; chưa tách đầy đủ các mã lý do của biểu chính thức.
- Chọn nhóm tuổi CNKT (&lt;25; 25–35; 36–45; 46–55; 56–60; &gt;60) hoặc CNKT chủ yếu (&lt;30; 30–40; 41–50; 51–60; &gt;60). Bảng tuổi hiển thị toàn bộ người đang làm; bảng bậc chỉ đếm vị trí CNKT.
- Xuất Excel độc lập gồm Tổng hợp, Biến động đơn vị, Danh sách chốt kỳ, Lịch sử trong kỳ. Các số liệu này do phần mềm tính trực tiếp, không cần Excel tính công thức.

Hồ sơ Cần rà soát không nằm trong tổng đang làm; số bị loại được cảnh báo. Kỳ chạm/trước ngày chốt nhập nguồn sẽ cảnh báo thiếu lịch sử đầu kỳ. Ngày vào làm cũ không được dùng để bịa lại lịch sử chức danh/trạng thái trước ngày chốt. **Đây là báo cáo nội bộ, chưa điền vào đúng các mẫu chính thức TKV, chưa tổng hợp tiền lương hay đầy đủ chỉ tiêu từng biểu.**

### Excel nguồn và mẫu Word

Excel nguồn giữ bản gốc, chỉnh sửa ô riêng, nhật ký và xuất lại toàn bộ biểu. Phần mềm không có bộ máy tính công thức Excel: các ô ƒ chỉ hiển thị kết quả lưu lần cuối, bản xuất cần Excel tính lại và 5 tệp liên kết ngoài còn thiếu. Hồ sơ quản lý/báo cáo tự tổng hợp là nguồn dữ liệu mới riêng, **không tự ghi ngược thay đổi vào 14 biểu Excel gốc**; chỉnh ô Excel sau khi nhập cũng không đồng bộ ngược vào hồ sơ.

Quản trị viên có thể nạp mẫu `.docx` cho 3 loại quyết định, áp dụng cho cả hai mục tạo Word. Các dấu hỗ trợ trong đoạn văn, bảng, header/footer:

`{{ho_ten}}`, `{{ma_nv}}`, `{{so_quyet_dinh}}`, `{{ngay_hieu_luc}}`, `{{chuc_danh_cu}}`, `{{don_vi_cu}}`, `{{chuc_danh_moi}}`, `{{don_vi_moi}}`, `{{can_cu}}`, `{{nguoi_ky}}`, `{{loai_quyet_dinh}}`.

Chưa hỗ trợ placeholder trong text box/SmartArt/field Word. Chưa có mẫu chính thức từ người dùng: mẫu mặc định ghi rõ Dự thảo, không suy diễn căn cứ pháp lý hay quyền lợi. Cần xuất thử kiểm tra định dạng.

## Lưu trữ, sao lưu và bảo mật bản thử

Dữ liệu mặc định ở `%LOCALAPPDATA%\EmgaHR` trên Windows; đặt `EMGA_DATA_DIR` để chọn nơi khác. Chỉ chạy **một tiến trình máy chủ** cho thư mục dữ liệu. Không đặt SQLite trên ổ mạng dùng chung; máy trạm truy cập HTTP đến máy chủ.

Sao lưu trong mục quản trị chứa Excel gốc, mẫu Word, tài khoản băm mật khẩu, nhật ký và toàn bộ hồ sơ/quyết định/lịch sử. Phục hồi: tắt máy chủ, sao chép dự phòng thư mục hiện tại, giải nén sao lưu vào thư mục dữ liệu riêng, đặt `EMGA_DATA_DIR` tới đó rồi chạy lại. Khóa phiên mới yêu cầu đăng nhập lại. Bảo quản bản sao lưu như dữ liệu nhân sự.

Có đăng nhập, CSRF, mật khẩu băm, vai trò, kiểm tra phiên bản chống ghi đè và ghi lịch sử. Chưa có giới hạn đăng nhập sai, đổi mật khẩu/khóa tài khoản, ký số hay phân quyền từng phòng ban. HTTP LAN chưa mã hóa; chỉ dùng mạng nội bộ tin cậy cho bản thử. Triển khai thật cần HTTPS và kiểm thử theo chính sách đơn vị.

## Phát triển và kiểm thử

Mỗi cloud task đã cô lập; dùng checkout hiện có, không tạo Git worktree trừ khi người dùng yêu cầu.

```bash
cd /workspace/emga
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests -v
EMGA_DATA_DIR=/tmp/emga-demo .venv/bin/python app.py
```

Máy chủ dùng Waitress; mặc định app.py chỉ lắng nghe loopback, cổng 8765. `Chay_Emga.bat` đặt `EMGA_HOST=0.0.0.0` cho LAN. Có thể đổi `EMGA_PORT`. Không dùng Flask debug server cho LAN.
