# Rà soát công thức CĐNL cho v5.5

Đã đọc công thức của bản Excel nguồn người dùng gửi trong phiên sửa này. Tệp nguồn được giữ ngoài checkout và không đưa vào gói hoặc GitHub. Tài liệu chỉ ghi cấu trúc công thức, không chép nhân sự, sản lượng, đơn giá hay số lương thực tế.

## Quỹ trong NT CĐNL

- `D37 = SUM(D38:D40)`: tổng quỹ sản phẩm Thiêu kết, Luyện gang, Luyện thép. Đây là tiền, không phải sản lượng tấn.
- Trưởng ca: `I44 = ROUND(D44*H44,-2)`, cơ sở là tổng quỹ các PX.
- Vận hành năng lượng: `I45 = I46+I49+I52+I58+I75+I76+I77+I79+I78+I80`. Các thành phần khác nhau: tỷ lệ quỹ sản xuất; bổ sung cơ sở lễ bằng một nửa tiền sản phẩm lễ của ba PX; trạm 110kV và bơm NHR có khoản nền/khấu trừ riêng; quạt gió khí nén theo tấn gang thường và lễ; oxy có công thức tỷ lệ và khấu trừ nửa khoản lễ. Các khoản lễ 78–80 còn có hệ số riêng. Không thay tổng này bằng SL × một đơn giá chung.
- Vận tải, cẩu, xúc lật: `I81 = I82+I340+I339+I83+I84+I85`. Có phần tỷ lệ quỹ, điều chỉnh thiếu định biên, giao nhiệm vụ và công dừng lò nhân hệ số riêng. Các dòng cung độ dùng số chuyến/chia định mức trước khi nhân đơn giá, có dòng lễ và khoản giảm trừ; không cộng thêm toàn bộ các dòng cung độ vào `I81` ngoài công thức đang có.
- Sàng quặng: `I341 = SUM(I342:I351)`. Tách tấn/công, lễ, chuyển thanh toán và các giao nhiệm vụ; không tự mở rộng vùng tổng sang các dòng bên dưới không thuộc công thức.
- Sửa chữa cơ điện: `I356 = I361+I368+I363+I357+I360+I362+I364+I367+I365+I366`. Gồm tỷ lệ quỹ, sửa chữa theo giao việc, công lễ, công SCTX/TĐT và thu hồi ứng. `I368=I369+I371`, `I371=SUM(I372:I393)`; không cộng lại danh sách GNV vào quỹ đã chứa các khoản này.
- Quỹ sản phẩm trực tiếp: `I43=I45+I81+I356+I341+I44`. Chế độ và khoản khác nằm ngoài tổng này theo `I42`.
- Quản lý: nguồn cũ có các dòng sản phẩm, lễ, chênh lệch và thu hồi; chế độ riêng ở `I432`. Chế độ quản lý v5 theo sản lượng và đơn giá riêng chỉ áp dụng khi người dùng chủ động bật/áp dụng, giữ như v5.4.

## Phân bổ trong sheet CĐNL

`AD` không dùng một đơn giá chung: công thức từng dòng nhân điểm hưởng `AB` với ô đơn giá phân bổ của nhóm:

| Nhóm tham chiếu | Đơn giá | Cơ sở trong Excel |
| --- | --- | --- |
| Vận hành năng lượng | `AJ146=AD144/AD145` | Quỹ nhóm sau phụ cấp / tổng điểm của vùng nhóm |
| Sửa chữa cơ điện | `AJ152=AD150/AD151` | Quỹ nhóm sau phụ cấp / tổng điểm của vùng nhóm |
| Vận tải | `AJ158=AD156/AD157` | Quỹ nhóm sau phụ cấp / tổng điểm của vùng nhóm |
| Quản lý theo nguồn cũ | `AJ168=AD166/AD167` | Quỹ sau các khoản trừ riêng / tổng điểm quản lý |

Phải xem công thức từng dòng, không suy nhóm chỉ theo vị trí liên tục: có dòng ngoài vùng chính vẫn tham chiếu `AJ146` hoặc `AJ152`. `AB` của nhiều dòng trực tiếp là đầu vào, trong khi quản lý có hệ số xếp loại và điều chỉnh theo ngày công. Không được thay tất cả `AB` bằng một công thức hệ số chung.

## Phạm vi v5.5 và điểm cần đối chiếu

Thay đổi tính cẩu phôi chỉ gọi cho PX Thép. Không sửa bộ nhập công hoặc bộ phân bổ CĐNL trong phiên này. Kiểm thử so sánh tất cả thẻ công thức của `NT CĐNL` và `CĐNL` trước/sau ba vòng lưu–mở và xóa sản lượng cẩu phôi: giữ nguyên.

Công thức phụ ở nguồn có tham chiếu `#REF!`, ví dụ `NT CĐNL!Q46` và `CĐNL!AH5`. Không đoán ô thay thế hoặc sửa các tham chiếu này. Ứng dụng đọc giá trị Excel lưu sẵn và không có bộ máy tính lại mọi công thức liên sheet; cờ tính lại khi mở Excel được giữ. Khi quỹ Thép thay đổi, cần mở bản Excel đã lưu để tính lại các liên kết CĐNL và đối chiếu trước khi xuất báo cáo CĐNL cho kỳ đó. Việc giữ công thức không phải xác nhận toàn bộ giá trị cache hoặc công thức nguồn đều chính xác. Công thức phân bổ CĐNL hiện hữu trong ứng dụng chưa được thay thế bằng bộ tính Excel đầy đủ trong phạm vi bổ sung cẩu phôi.
