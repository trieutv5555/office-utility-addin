# Office Utility Add-in

Tiện ích tiếng Việt dành cho Microsoft Excel và Word, giúp xử lý dữ liệu, rà soát sai sót và chuẩn hóa văn bản nhanh hơn.

## Chức năng chính

- **Excel:** đọc số thành chữ, tìm dòng trùng, kiểm tra công thức và lỗi dữ liệu, tìm kiếm nhiều sheet, so sánh vùng, chuyển đổi dữ liệu, gộp sheet và báo cáo có liên kết.
- **Word:** chuẩn hóa văn bản, dọn khoảng trắng, thay thế hàng loạt, rà soát dấu câu, đọc số tiền thành chữ căn lề A4 và AutoFit bảng hàng loạt.

## GPS (Excel)

Cấu hình từ 1 đến 50 nhà máy trong khung tiện ích: tên, tọa độ và bán kính riêng. Điều chỉnh ngưỡng chuyến (mặc định 5 km), thời gian dừng tại nhà máy (mặc định ≥30 phút) và dừng nghỉ (mặc định >10 phút). Cấu hình đã lưu từ 1.0.7 được tự chuyển sang định dạng mới. Tọa độ không có giá trị mặc định; chỉ lưu trong bộ nhớ cục bộ của thiết bị khi chọn Lưu cài đặt GPS. Có thể xóa cài đặt hoặc chạy mà không lưu. Dữ liệu hành trình được đọc từ workbook, không gửi đến máy chủ của add-in.

Các điều kiện GPS có hướng dẫn và ví dụ ngay dưới ô nhập. Khôi phục mặc định chỉ đặt lại 5 km, 30 phút, 10 phút; không đổi danh sách nhà máy. Xóa cài đặt cần xác nhận và xóa cả tọa độ đã lưu.

## Sử dụng

Add-in dùng Office.js, chạy trên các phiên bản Office tương thích (không hỗ trợ Office 2010). Triển khai qua GitHub Pages, không cần VBA hay máy chủ cục bộ.

Tải [manifest](https://trieutv5555.github.io/office-utility-addin/manifest.xml) hoặc sử dụng `INSTALL_ADDIN.bat` để cài thử qua Trusted Add-in Catalogs. Nên sao lưu tài liệu trước khi sửa hàng loạt.

Kiểm thử: `node tests/regression.js`. Chưa thay thế kiểm thử trên Office thực tế.
