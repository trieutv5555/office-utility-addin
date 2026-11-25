# Office Utility Add-in

Tiện ích tiếng Việt dành cho Microsoft Excel và Word, giúp xử lý dữ liệu, rà soát sai sót và chuẩn hóa văn bản nhanh hơn.

## Chức năng chính

- **Excel:** đọc số thành chữ, tìm dòng trùng, kiểm tra công thức và lỗi dữ liệu, tìm kiếm nhiều sheet, so sánh vùng, chuyển đổi dữ liệu, gộp sheet và báo cáo có liên kết.
- **Word:** chuẩn hóa văn bản, dọn khoảng trắng, thay thế hàng loạt, rà soát dấu câu, đọc số tiền thành chữ căn lề A4 và AutoFit bảng hàng loạt.

## Sử dụng

Add-in dùng Office.js, chạy trên các phiên bản Office tương thích (không hỗ trợ Office 2010). Triển khai qua GitHub Pages, không cần VBA hay máy chủ cục bộ.

Tải [manifest](https://trieutv5555.github.io/office-utility-addin/manifest.xml) hoặc sử dụng `INSTALL_ADDIN.bat` để cài thử qua Trusted Add-in Catalogs. Nên sao lưu tài liệu trước khi sửa hàng loạt.

Kiểm thử: `node tests/regression.js`. Chưa thay thế kiểm thử trên Office thực tế.
