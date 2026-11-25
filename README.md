# Tiện ích văn phòng – Add-in cho Excel và Word

Một add-in (Office.js) dùng chung một mã nguồn cho cả Excel và Word.

## Tính năng

**Dùng cho cả hai**
- Xoá khoảng trắng thừa, CHỮ HOA, chữ thường, Viết Hoa Đầu Mỗi Từ, bỏ dấu tiếng Việt

**Excel**
- Đọc số thành chữ tiếng Việt (ghi vào cột bên phải)
- Tô vàng hoặc xoá dòng trùng
- Kiểm tra công thức trong cột: tìm ô có công thức khác với đa số công thức cùng cột, hoặc ô nhập tay chen giữa các công thức; bấm vào kết quả để nhảy tới ô, có thể tô đỏ nhạt

**Word**
- Dọn dẹp tài liệu (gộp khoảng trắng liên tiếp, bỏ đoạn trống thừa)
- Thay thế hàng loạt theo danh sách `cũ => mới`
- Thống kê số từ, ký tự, đoạn
- Chèn ngày hiện tại (`25/11/2026` hoặc `Ngày 25 tháng 11 năm 2026`)

## Triển khai GitHub Pages

Đưa toàn bộ nội dung thư mục lên repository `trieutv5555/office-utility-addin` (nhánh main). Settings → Pages → Deploy from a branch → main → /docs → Save.

Kiểm tra https://trieutv5555.github.io/office-utility-addin/taskpane.html và /manifest.xml trước khi cài.

## Cài Windows

Chạy INSTALL_ADDIN.bat để tải manifest về %LOCALAPPDATA%\OfficeUtilityAddin. Chia sẻ thư mục này cho tài khoản Windows của bạn (không cấp Everyone), lấy đường dẫn UNC. Trong Office: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → Add catalog, nhập UNC, bật Show in Menu. Đóng và mở lại Office. Home → Add-ins → Advanced → Shared Folder → Tiện ích văn phòng → Add. Làm cho Excel và Word. Đây là sideload để sử dụng thử; không bảo đảm tự hiện trong mọi tài liệu trên mọi phiên bản Office. Office 2010 không hỗ trợ.

Office web: tải manifest.xml rồi dùng Upload My Add-in nếu tài khoản hỗ trợ.

## Cập nhật và gỡ

Sửa mã trong docs, chạy UPLOAD_ADDIN.bat (cần Git và quyền push). Nếu đã có repository trên GitHub, clone trước khi sửa để tránh xung đột. Thay đổi manifest phải cập nhật bản cục bộ và có thể cần thêm lại add-in. UNINSTALL_ADDIN.bat chỉ xóa manifest đã tải; gỡ add-in và catalog trong Office theo hướng dẫn.

## Lưu ý

- Thay đổi từ Office.js trong Excel thường **không hoàn tác được bằng Ctrl+Z**. Hãy sao lưu file trước khi xoá dòng trùng hoặc sửa hàng loạt. Trong Word thì hoàn tác được.
- Các thao tác chữ hoa/thường và bỏ dấu trong Word thay lại toàn bộ đoạn đang chọn, nên định dạng riêng từng phần (đậm, nghiêng…) trong đoạn đó có thể mất.
- Các hàm xử lý chữ trong Excel bỏ qua ô chứa công thức.
- Kiểm tra công thức so sánh theo kiểu R1C1 và lấy công thức chiếm quá nửa số công thức của cột làm chuẩn. Cột chưa đủ 3 công thức, hoặc không có công thức nào chiếm đa số, sẽ được bỏ qua. Dòng tổng (ví dụ SUM ở cuối cột) cũng sẽ được liệt kê vì khác công thức chuẩn, bạn tự đánh giá là hợp lệ hay không.
- Đọc số thành chữ hỗ trợ tối đa 15 chữ số, làm tròn đến số nguyên.
