# Excel Add-in: Đọc Số Thành Chữ Tiếng Việt (VND)

Công cụ Add-in dành cho Microsoft Excel giúp chuyển đổi các con số (số tiền) thành chữ viết tiếng Việt một cách tự động và chính xác. Hỗ trợ đầy đủ các quy tắc đọc số tiền tệ (đồng), các trường hợp đặc biệt như "lẻ", "mốt", "lăm", "tư".

## ✨ Tính năng

-   **Chuyển đổi nhanh**: Chuyển đổi vùng dữ liệu được chọn chỉ với một cú nhấp chuột.
-   **Đúng chuẩn tiếng Việt**: Xử lý chính xác các hàng đơn vị, chục, trăm, nghìn, triệu, tỷ...
-   **Tự động điền**: Kết quả được ghi trực tiếp vào cột bên phải của vùng dữ liệu được chọn.
-   **Giao diện tích hợp**: Xuất hiện trực tiếp trong tab "Home" của Excel.

## 🚀 Hướng dẫn cài đặt (Sideloading)

Vì đây là Add-in phát triển riêng, bạn có thể cài đặt theo các bước sau:

1.  **Chuẩn bị mã nguồn**: Đảm bảo bạn đã tải toàn bộ thư mục code về máy.
2.  **Chỉnh sửa Manifest**: Mở file `manifest.xml` và cập nhật các URL tại thẻ `<SourceLocation>` và `<bt:Url>` trỏ về nơi bạn lưu trữ file (ví dụ: `https://localhost:3000` nếu chạy local hoặc URL GitHub Pages).
3.  **Chia sẻ thư mục**: 
    -   Tạo một thư mục trên máy tính và đặt file `manifest.xml` vào đó.
    -   Chuột phải vào thư mục -> **Properties** -> **Sharing** -> **Share** (Cấp quyền cho bản thân hoặc Everyone).
    -   Copy đường dẫn mạng (Network Path) của thư mục này.
4.  **Thêm vào Excel**:
    -   Mở Excel -> **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...**
    -   Chọn **Trusted Add-in Catalogs**.
    -   Dán đường dẫn mạng vào ô **Catalog Url** -> **Add catalog**.
    -   Tích vào ô **Show in Menu** -> **OK**.
    -   Khởi động lại Excel.
5.  **Sử dụng**: Vào tab **Insert** -> **My Add-ins** -> **Shared Folder** và chọn **Đọc Số VNĐ**.

## 📖 Hướng dẫn sử dụng

1.  Mở bảng tính Excel chứa dữ liệu số tiền.
2.  Bôi đen (chọn) cột hoặc vùng chứa các con số cần chuyển đổi.
3.  Vào tab **Home**, bấm nút **Đọc Số VNĐ** ở phía cuối thanh công cụ.
4.  Bảng điều khiển (Taskpane) hiện ra, bấm nút **Chuyển sang chữ**.
5.  Kết quả sẽ tự động xuất hiện ở cột bên phải tương ứng.

## 🛠 Cấu trúc dự án

-   `manifest.xml`: File cấu hình định danh Add-in với Excel.
-   `index.html`: Giao diện người dùng của Taskpane.
-   `Office.js`: Logic xử lý sự kiện Office và thuật toán chuyển đổi số.
-   `taskpane.css`: Định dạng giao diện.

## 📝 Lưu ý kỹ thuật

Thuật toán sử dụng phương pháp chia khối 3 chữ số để xử lý, đảm bảo tốc độ nhanh và chính xác cho cả những con số cực lớn (hàng triệu tỷ).

---
*Phát triển bởi IT Admin Dept*
