# Emga HR — bản chạy thử 0.1

Ứng dụng tiếng Việt cho Windows, dùng trình duyệt để quản lý bảng tính nhân sự trên máy chủ nội bộ. Đây là bản thử nghiệm để chốt nghiệp vụ, **chưa phải bản thay thế đầy đủ Excel**. Không kèm dữ liệu nhân sự của người dùng trong mã nguồn hay gói tải xuống.

## Chạy trên Windows

1. Cài Python 3.12 trở lên từ python.org (có Python Launcher `py`). Giải nén gói ứng dụng vào thư mục có quyền ghi.
2. Bấm `Chay_Emga.bat`. Lần đầu cần Internet để tải thư viện; các lần sau pip kiểm tra các bản đã cài.
3. Trên chính máy chủ, mở trình duyệt tới cổng **8765** qua địa chỉ loopback `127.0.0.1`. Tạo tài khoản quản trị (mật khẩu tối thiểu 12 ký tự). Chỉ máy chủ được tạo tài khoản đầu tiên.
4. Nhập tệp Excel `.xlsx` trong mục **Dữ liệu & báo cáo**. Chọn biểu ở danh sách, tìm nội dung, chuyển nhóm 16 cột / 30 dòng. Nhập chữ cột như `A`, `AH`, `HE` để đi nhanh. Bấm ô để chỉnh sửa; chọn kiểu Văn bản để giữ số 0 đầu, Số để tính toán hoặc Ngày dạng `yyyy-mm-dd`.
5. Vào **Quản trị & sao lưu** để tạo tài khoản cho đồng nghiệp: chỉ xem/xuất, sửa hoặc quản trị.
6. Để dùng LAN, dùng `ipconfig` trên máy chủ để xác định IPv4 của card mạng đang sử dụng; máy khác mở địa chỉ IP đó với cổng 8765. Nếu Windows Firewall chặn, chỉ cho phép Python/cổng TCP 8765 ở hồ sơ mạng **Private**, giới hạn máy trong LAN. Không mở cổng router ra Internet. Máy chủ và cửa sổ chạy phần mềm phải hoạt động.

Chưa kiểm thử trực tiếp trên Windows trong phiên phát triển Linux này. Không có tệp `.exe` hay bộ cài Windows đã ký. Máy trạm không cần Python hay Excel để xem/sửa, nhưng cần Word/Excel hoặc phần mềm tương thích để mở tài liệu xuất.

## Dữ liệu và sao lưu

Dữ liệu mặc định nằm ngoài thư mục mã nguồn: `%LOCALAPPDATA%\EmgaHR` trên Windows. Có thể đặt `EMGA_DATA_DIR` để chọn nơi khác. Chỉ chạy **một tiến trình máy chủ** cho thư mục dữ liệu. Không đặt cơ sở dữ liệu SQLite trên ổ mạng dùng chung; máy trạm truy cập máy chủ bằng HTTP. Giữ máy chủ có phân quyền tài khoản Windows phù hợp.

Nút **Tải bản sao lưu** xuất SQLite, tệp gốc và mẫu Word. Phục hồi thủ công: tắt máy chủ, sao chép dự phòng thư mục hiện tại, giải nén bản sao lưu vào thư mục dữ liệu riêng, đặt `EMGA_DATA_DIR` tới thư mục đó rồi khởi động. Khóa phiên mới sẽ yêu cầu đăng nhập lại. Bản sao chứa thông tin cá nhân và mật khẩu đã băm; bảo quản như dữ liệu nhân sự. Sao lưu không phải bản Excel báo cáo và không nên gửi cho tài khoản chỉ xem.

Ứng dụng có đăng nhập, mật khẩu băm, CSRF, vai trò, kiểm tra phiên bản chống ghi đè và nhật ký sửa ô. HTTP LAN chưa mã hóa; chỉ dùng mạng nội bộ tin cậy cho bản thử. Khi triển khai thật, cần cấu hình HTTPS, chính sách tài khoản, phân quyền dữ liệu chi tiết và kiểm thử bảo mật. Hiện chưa có khóa tài khoản/đổi mật khẩu hay giới hạn đăng nhập sai.

## Excel và phạm vi hiện tại

- Giữ bản gốc bất biến; các chỉnh sửa được lưu riêng và gắn vào bản xuất. Lần nhập tiếp theo tạo bản độc lập, không gộp nhân sự và không ghi đè bản trước.
- Xem/sửa các ô dữ liệu trên toàn bộ 14 trang tính của tệp đã gửi. Giao diện bảng không mô phỏng chính xác các ô gộp và định dạng bản in; tệp xuất giữ các trang, định dạng và công thức qua openpyxl.
- Danh sách nhân sự lấy các dòng có họ tên ở cột F của `DU LIEU CO SO`, từ dòng 4. Mã ở E, chức danh B, mã đơn vị N. Đây là danh sách dòng, **không khẳng định số lao động đang làm**, không tự loại người nghỉ hoặc trùng mã.
- Công thức chỉ đọc, hiển thị kết quả Excel lưu lần cuối. Phần mềm **không có bộ máy tính công thức Excel**. Sau chỉnh sửa, báo cáo hiển thị có thể cũ; bản xuất yêu cầu Excel tính lại. Chưa tự thêm nhân sự mới, chèn dòng, tự chuyển mã tăng/giảm hoặc tự tổng hợp kỳ mới.
- Tệp người dùng có **5 liên kết ngoài**, công thức INDEX/MATCH, COUNTIF, DATEDIF… Khi thiếu tệp nguồn, không thể bảo đảm báo cáo tính đầy đủ. Bản xuất giữ liên kết; chỉ cho Excel cập nhật nguồn mà đơn vị tin cậy. Nhập lại tệp đã được Excel tính và lưu nếu cần xem kết quả mới trong phần mềm.
- Chưa có đối chiếu nghiệp vụ từng chỉ tiêu với người phụ trách, chưa kiểm chứng bản in chính thức. Không khẳng định ứng dụng thực hiện đủ mọi chức năng của Excel.

## Quyết định Word

Có 3 loại: điều động, chấm dứt HĐLĐ, thay đổi chức danh. Chọn nhân sự, nhập số, ngày hiệu lực, căn cứ/nội dung, đơn vị/chức danh mới và người ký để tải Word. Khi chưa nạp mẫu, tệp ghi rõ **BẢN DỰ THẢO**, không tự suy diễn căn cứ pháp lý, quyền lợi hay người có thẩm quyền. Xuất văn bản không tự thay đổi hồ sơ hoặc trạng thái hợp đồng.

Quản trị viên có thể nạp mẫu `.docx` cho từng loại. Trong đoạn văn, ô bảng, header/footer, chèn các dấu sau (cả dấu ngoặc):

| Dấu | Nội dung |
|---|---|
| `{{ho_ten}}`, `{{ma_nv}}` | Họ tên, mã nhân sự |
| `{{so_quyet_dinh}}`, `{{ngay_hieu_luc}}` | Số quyết định, ngày hiệu lực |
| `{{chuc_danh_cu}}`, `{{don_vi_cu}}` | Chức danh và mã đơn vị hiện tại |
| `{{chuc_danh_moi}}`, `{{don_vi_moi}}` | Thông tin mới do người lập nhập |
| `{{can_cu}}`, `{{nguoi_ky}}`, `{{loai_quyet_dinh}}` | Nội dung, người ký, loại văn bản |

Chưa hỗ trợ placeholder trong text box, SmartArt, field Word; cần xuất thử để kiểm tra. Chỉ nạp mẫu tin cậy. Chưa nhận được mẫu quyết định chính thức từ người dùng.

## Phát triển / kiểm tra

Mỗi cloud task đã được cô lập; dùng checkout hiện có, không tạo Git worktree trừ khi người dùng yêu cầu.

```bash
cd /workspace/emga
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests -v
EMGA_DATA_DIR=/tmp/emga-demo .venv/bin/python app.py
```

Mặc định `app.py` chỉ lắng nghe loopback; đặt `EMGA_HOST=0.0.0.0` để phục vụ LAN. Có thể đặt `EMGA_PORT` (mặc định 8765). Không dùng Flask debug server cho máy chủ LAN. Mã nguồn dùng Waitress.
