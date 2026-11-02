# Rà soát GNV cả 5 phân xưởng — v5.9

Đối chiếu cấu trúc của workbook nguồn đính kèm bằng bản sao cục bộ. Không đưa Excel, tên nhân sự hay giá trị sản lượng/lương người dùng vào tài liệu/gói.

| PX / sheet đang dùng | Dòng chi tiết hỗ trợ | Số dòng | Tổng và quy tắc riêng |
| --- | --- | ---: | --- |
| Thiêu kết / NT thiêu kết | 80–90 | 11 | D79/J79. Dòng 91 là lương chế độ; không dùng vùng của sheet lịch sử NT T kết. |
| Luyện gang / NT gang | 61–62; 64–93; 95–100; 102–103 | 40 | 38 dòng của PX và 2 dòng chuyển thanh toán PX CĐNL. Giữ quy tắc nguồn J59 theo J63; J60 tách riêng. Dòng mới mặc định vào 38 dòng của PX, không tự tràn sang nhóm chuyển thanh toán. |
| Luyện thép / NT Thép | 60–97 | 38 | D59/J59; giữ phần chế độ từ 98. |
| CĐNL / NT CĐNL | 372–393 | 22 | D371/I371 đi qua I368 vào quỹ sửa chữa I356 một lần. Giữ đơn giá H, tiền I; không ghi trùng công sang E. I416 là khoản riêng ngoài quỹ này. |
| Khai tuyển / NT. K. tuyển | 60–63; 65–79; 82–83; 85–87; 89–94; 96–101 | 36 | J52 tổng các dòng chi tiết, không cộng lại tổng cha 59/64/80/81/84/88/95. Cập nhật tổng phụ riêng; phần chế độ từ 102. Hỗ trợ tên sheet có/không có dấu cách cuối. |

Các dòng tổng phụ/tiêu đề không được nhập như GNV độc lập hoặc dùng làm ô trống: Luyện gang 60/63/94/101; Khai tuyển 59/64/80/81/84/88/95. Snapshot cũ chứa dòng tổng phụ có nội dung trùng với tiêu đề nguồn được loại khi đọc để tránh phục hồi khoản trùng; một mục thực có nội dung khác vẫn được cấp dòng chi tiết hợp lệ.

Danh sách đã chỉnh/nạp là nguồn ghi chi tiết. Lưu đủ tên, ĐVT, khối lượng, đơn giá, tiền; giữ mốc hợp lệ, cấp dòng trống không đè mục có mốc phía sau. Khoản âm được đọc, số 0 giữ qua snapshot; xóa một/xóa hết phải xóa cả ô chi tiết và tổng. Dòng tự tính dùng ROUND(khối lượng × đơn giá,0); khoản tiền nguồn đặc biệt chưa bật tự tính giữ nguyên giá trị. Không ghi lại GNV chưa chỉnh bằng giá trị tạm trong model để tránh đè kết quả của bộ tính CĐNL.

Vượt sức chứa/nhóm thanh toán, thiếu sheet, trùng mốc chuyển thanh toán Luyện gang hoặc mẫu đã chèn dòng làm chuyển mốc lương chế độ: dừng lưu với thông báo rõ, giữ nguyên tệp gốc. CĐNL có công/tiền ở 394–415 phải đối chiếu vùng tổng trước khi nhập; đây không phải vùng GNV đang nằm trong công thức I371. Không bỏ qua dữ liệu và không tự gộp vào quỹ sửa chữa hoặc khoản riêng I416. Chưa tự chèn thêm dòng và dịch toàn bộ công thức liên sheet.

`AllWorkshopGnvExportTests` kiểm tra ô NT thực, công thức tổng, tổng phụ, các tiêu đề được bảo vệ, lưu–mở, đủ sức chứa/vượt khuôn ở từng PX, dòng mới/trùng mốc, khoản âm/0/tiền đặc biệt, xóa, snapshot cũ, tên sheet có khoảng trắng, mẫu chèn dòng và dữ liệu ngoài vùng. Kiểm tra cả Word với danh sách đầy đủ của 5 PX. `CdnlPayrollTests` kiểm tra khoản chuyển thanh toán âm đi qua quỹ sửa chữa một lần, xóa GNV cũng xóa giá trị live cũ, và lưu–mở quỹ/phân bổ không đổi. Bộ kiểm thử nguồn chỉ chạy trên bản sao, không coi các sửa lỗi giả trong kiểm thử là số liệu được duyệt.
