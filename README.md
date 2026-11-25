# Tiện ích văn phòng – Add-in cho Excel và Word

Một add-in (Office.js) dùng chung một mã nguồn cho cả Excel và Word.

## Tính năng

**Dùng cho cả hai**
- Xoá khoảng trắng thừa, CHỮ HOA, chữ thường, Viết Hoa Đầu Mỗi Từ, bỏ dấu tiếng Việt

**Excel**
- Đọc số thành chữ tiếng Việt (ghi vào cột bên phải)
- Tô vàng dòng trùng (không tự xoá hàng)
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

## Bản cải tiến
- Thêm tìm kiếm workbook, kiểm tra lỗi Excel, dòng trùng từng sheet, ô trống và công thức sai/thiếu; mặc định bỏ qua sheet ẩn. Báo cáo có thể nhảy đến ô và xuất CSV (tuỳ host có cho tải xuống).
- Thêm thống kê vùng chọn, danh sách sheet, Wrap Text/Middle Align/AutoFit; Word thêm đếm từ khoá và chèn mẫu.
- Tắt xoá cả hàng trùng để bảo vệ dữ liệu; số thành chữ từ chối ghi đè; xác nhận thao tác sửa dữ liệu; khoá nút trong lúc chạy.
- Giữ nguyên manifest để cập nhật giao diện mà không cần đăng ký lại. Không thay đổi phiên bản manifest.
- Rà soát theo vùng đã dùng hoặc vùng chọn; tối đa 100.000 ô/vùng, 10.000 kết quả, hiển thị 200 trong báo cáo chung. Kiểm tra công thức theo đa số chỉ là gợi ý, không tự sửa. Không hỗ trợ Office 2010.
- Kiểm thử: node --check cho JS và 3 kiểm thử hàm thuần. Chưa kiểm thử Office thật, CSV tải xuống và khả năng tương thích từng bản Office. Màu nền phục hồi gần đúng trong cùng phiên, không bảo đảm màu theo theme/pattern. Hãy dùng bản sao dữ liệu khi thử.

## Bản mở rộng tiện ích phổ thông

Bản này không chứa kiểm tra sản lượng, lương, đơn giá hay quy tắc công ty. Chép các file trong gói cập nhật vào repository hiện có, chạy UPLOAD_ADDIN.bat; giữ manifest cũ. Công cụ Excel có giới hạn 100.000 ô/vùng, tối đa 5.000 ô ghi riêng lẻ mỗi lần; báo cáo lưu tối đa 10.000 kết quả. Chuyển ngày văn bản sang ISO dạng text, không đổi hệ ngày Excel. Gộp sheet theo thứ tự cột của vùng dùng, giữ mọi dòng (kể cả tiêu đề) và chỉ sao chép giá trị. Tham chiếu được nhận diện bằng mẫu công thức trực tiếp, không phân tích INDIRECT, tên định nghĩa, liên kết ẩn hoặc mọi trường hợp 3D. Kiểm tra số thứ tự theo thứ tự hiện tại, báo khoảng nhảy, giảm và trùng. Lưu tùy chọn cục bộ, không lưu nội dung tài liệu. Mẫu hành chính chỉ định dạng font và đoạn, không phải xác nhận tuân thủ thể thức pháp lý; chưa thay đổi lề trang. Chức năng bảng áp dụng kiểu bảng có sẵn. Kiểm thử cú pháp JS và 7 trường hợp phân tích số/ngày; chưa kiểm thử Office thực tế. Nên dùng bản sao tài liệu, nhất là thao tác chèn hàng, đổi tên, gộp và chuyển công thức.

## Bản ổn định
- Kiểm tra kích thước trước khi tải vùng; gom các ô ghi liên tiếp thành khối; kiểm tra lại nội dung trước khi ghi sau xác nhận.
- Rà soát công thức bỏ qua cột không có công thức; kết quả kiểm tra cột cũng được đưa vào báo cáo chung. Ô trống khi quét vùng chọn bao gồm cả phần cuối vùng chọn.
- Số âm làm tròn đối xứng; từ chối số quá lớn. Tô dòng trùng yêu cầu xác nhận; màu nền cũ không được phục hồi bởi công cụ này.
- Giữ nguyên manifest và chức năng hiện có. Kiểm thử thực tế trong Office vẫn cần thực hiện trên bản sao dữ liệu.

## Tối ưu theo các dự án mã nguồn mở
- Rà soát nhiều sheet đọc theo khối (mục tiêu 20.000 ô/khối), giới hạn 1 triệu ô/sheet và 5 triệu ô/lượt; riêng công thức tối đa 100.000 dòng/cột, giữ nguyên toàn bộ cột khi so sánh mẫu R1C1. Các sheet vượt giới hạn được nêu tên. Báo cáo đếm đầy đủ nhưng chỉ lưu 10.000 kết quả, hiển thị 200.
- Dòng trùng giữ trạng thái so sánh giữa các khối; chuẩn hóa Unicode NFC, giữ kiểu dữ liệu, dấu câu và khoảng trắng bên trong. Không tự gộp những giá trị gần giống.
- Tách hàm bảo vệ văn bản ghi vào Excel và CSV để nhất quán; thống kê kích thước và gộp sheet tải thông tin kích thước theo lô. Giới hạn ghi/sửa và gộp vẫn giữ nguyên.
- Tham khảo phương pháp, không sao chép thư viện: https://github.com/OfficeDev/Office-Add-in-samples ; https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-large ; https://openrefine.org/docs/technical-reference/clustering-in-depth ; https://www.papaparse.com/docs . SheetJS SSF được tham khảo về hệ ngày Excel nhưng chưa thay đổi cách chuẩn hóa ngày.

## Bản ổn định 4
- ExcelApi 1.9+: nhận diện công thức bằng SpecialCells/RangeAreas; trên bản cũ kiểm tra chỉ dùng suy đoán chuỗi. Chuyển công thức thành giá trị cần ExcelApi 1.12 để kiểm tra vùng mảng/tràn. Trên bản hỗ trợ 1.12, các vùng sửa dữ liệu trực tiếp có mảng/tràn sẽ bị từ chối (kể cả vùng hỗn hợp). Không cam kết bảo vệ mọi thao tác cấu trúc sheet.
- WordApiDesktop 1.4+: chữ hoa/thường dùng chức năng đổi kiểu chữ gốc của Word, tránh thay toàn đoạn. Các host cũ và thao tác khác vẫn dùng cơ chế thay đoạn có xác nhận, có thể mất định dạng.
- Ngày tháng vẫn là ISO dạng văn bản; không dùng API Preview của hệ ngày 1900/1904. Khôi phục màu vẫn gần đúng, chưa bảo toàn theme/pattern.
- Kiểm thử hồi quy: `node tests/regression.js` (không cần cài thư viện). Đây là test JS và mô phỏng, không thay thế kiểm thử Office thật.
- Tài liệu: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-multiple-ranges ; https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-dynamic-arrays ; https://learn.microsoft.com/en-us/javascript/api/word/word.range .

- Tô vàng dòng trùng: hiển thị số dòng đã tô và dòng xuất hiện đầu tiên ngay dưới nút; bấm để tới ô, đồng thời cập nhật báo cáo chung để xuất CSV hoặc tạo sheet báo cáo. Hiển thị tối đa 200 dòng, lưu đủ tối đa 2.000 dòng đã tô.
