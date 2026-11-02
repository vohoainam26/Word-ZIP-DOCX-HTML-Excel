# Word ZIP / DOCX to HTML Excel

Ứng dụng web tĩnh xử lý trực tiếp trong trình duyệt:

## Summary Labels

- ZIP hàng loạt và DOCX lẻ
- EN, FR, DE, CZ, PL, HU, ES, PT, IT, RO, SK, NL, DK, FI, SE, EE
- File kép ES + PT / ES > PT
- Tiền tố phiên bản như v0_, v1_, v12_
- ZIP có nhiều thư mục ngôn ngữ; ngôn ngữ được nhận từ tên thư mục
- Nếu có đồng thời CZ 5x3 và CZ 7x5, ưu tiên CZ 7x5
- EN và CZ dùng định dạng chuẩn cũ cùng Storage bắt buộc
- Các ngôn ngữ còn lại giữ đầy đủ nội dung và định dạng cơ bản theo Word
- Theo dõi trạng thái và nguyên nhân lỗi
- Xem trước HTML
- Duyệt thủ công trước khi xuất Excel
- Excel gộp theo sản phẩm, gồm default_code, name và các cột ngôn ngữ; ô thiếu ghi X

## IMPORT LABELS

- Nhận nhiều ZIP hoặc DOCX lẻ; ZIP được quét toàn bộ thư mục con và chỉ lấy DOCX
- Mỗi DOCX có nhiều khối ngôn ngữ và một bảng NTF dùng chung
- Tách từng khối theo mã `(EN)`, `(FR)`, `(DE)`...
- `default_code` không bắt buộc; tên sản phẩm lấy từ tiêu đề đầu tiên trong Word
- EN/CZ chỉ gồm mô tả, thành phần, NTF xuống từng dòng và Storage bắt buộc; không thêm thông tin công ty
- Các ngôn ngữ khác viết NTF trong một đoạn, cách nhau bằng dấu phẩy, không tự thêm Storage
- Cho phép chọn Importer hoặc Seller cho từng lần quét; chỉ áp dụng cho các ngôn ngữ ngoài EN/CZ
- Số chỉ tiêu NTF được đọc động theo từng bảng Word; không giới hạn cố định 7 dòng
- Hỗ trợ NTF dùng chung nằm trong bảng hoặc một đoạn văn phân cách các chỉ tiêu bằng dấu chấm phẩy
- Khối ngôn ngữ viết HOA toàn bộ được tự chuyển về kiểu câu, giữ mã E-number và định dạng in đậm
- Giữ cả các dòng bổ sung như chất xơ, natri hoặc chỉ tiêu khác khi nhãn được xác định rõ
- Dòng không có giá trị hoặc nhãn không xác định được sẽ báo lỗi thay vì tự đoán
- Excel giữ cấu trúc `default_code`, `name` và các cột ngôn ngữ; ô thiếu ghi X

## ODOO FILL

- Chọn một file Excel Odoo đang làm việc bằng Chrome/Edge; trình duyệt nhớ quyền truy cập file cho các lượt sau
- Nhận đồng thời ZIP/DOCX theo cả hai dạng Summary Labels và IMPORT LABELS
- ZIP được quét toàn bộ thư mục con; các file Summary được gom khi có cùng mã và cùng tên sản phẩm
- Kiểm tra tên/code trước khi chuyển HTML: chỉ sản phẩm có trong Odoo mới được xử lý và hiển thị; sản phẩm ngoài Odoo được bỏ qua
- Nếu sản phẩm đã khớp với Odoo nhưng DOCX không chuyển được, lỗi của sản phẩm đó được giữ lại để kiểm tra
- Có thể tải log lỗi Excel gồm tên sản phẩm Odoo, Internal Reference, tên ZIP/DOCX, đường dẫn đầy đủ trong thư mục input và nguyên nhân lỗi
- Tự tìm sản phẩm trên toàn bộ file Odoo theo `Internal Reference` hoặc tên; kết quả gần đúng phải được người dùng kiểm tra trước khi duyệt
- Có nút `Duyệt tất cả` để chọn nhanh toàn bộ sản phẩm đang có ô HTML mới hợp lệ
- Cho phép lưu từng phần: Daily có thể ghi trước, lượt Import sau đọc lại các ô đã có và bổ sung phần còn thiếu đến đủ 15 ngôn ngữ
- Chỉ ghi vào ô ngôn ngữ đang trống; dữ liệu cũ được giữ nguyên và xung đột được báo trên giao diện
- EE được Fill nếu nguồn có; cột `ES > PT` được bỏ qua
- Có thể xem ảnh nhãn hoặc mã HTML theo từng ngôn ngữ trước khi Fill
- Sau khi tick duyệt, kết quả được ghi trực tiếp vào chính file Odoo đã chọn, không tạo thêm file `*_filled.xlsx`

## Chạy cục bộ

Mở `index.html` hoặc phục vụ thư mục bằng HTTP server.

## Triển khai Vercel

Chạy lệnh triển khai tại chính thư mục này. Không cần bước build.

