🛠️ Ribbon XML Builder (Web-based Tool)
Ribbon XML Builder là một công cụ web đơn giản, mạnh mẽ giúp các lập trình viên VBA và Office Add-in thiết kế giao diện Ribbon (Custom UI) cho Microsoft Office (Excel, Word, PowerPoint, Access) một cách trực quan.

Thay vì phải gõ thủ công từng dòng lệnh XML dễ gây lỗi, công cụ này cung cấp giao diện Kéo-Thả (Drag & Drop) và Xem trước thời gian thực (Live Preview), sau đó tự động sinh mã XML chuẩn.


✨ Tính năng nổi bật

⚡ Live Preview (Xem trước tức thì):

Giao diện mô phỏng chính xác thanh Ribbon của Office (Excel style).

Hiển thị icon, label và bố cục ngay khi bạn nhập liệu.

🖱️ Trình biên tập trực quan:

Quản lý cấu trúc theo cấp bậc: Tab > Group > Button.

Dễ dàng thay đổi thứ tự các nút/nhóm bằng các nút điều hướng (Lên/Xuống).

Hỗ trợ 2 kích thước nút: Large (lớn) và Normal (nhỏ - tự động xếp chồng cột 3 nút).

🔄 Chuyển đổi hai chiều (Bi-directional):

Export: Tự động sinh mã customUI XML chuẩn (namespace 2009/07).

Import: Dán đoạn mã XML có sẵn để chỉnh sửa lại giao diện (Reverse Engineering).

🧠 Thông minh & Tiện lợi:

Auto ID: Tự động tạo ID chuẩn dựa trên Label bạn nhập (Ví dụ: "Nhập Dữ Liệu" -> btnNhapDuLieu).

Icon Support: Hỗ trợ imageMso (icon có sẵn của Office) và image (icon tùy chỉnh).

Local Storage: Tự động lưu lại quá trình làm việc, không sợ mất dữ liệu khi tải lại trang.

🚀 Hướng dẫn cài đặt
Công cụ này được xây dựng hoàn toàn bằng Vanilla HTML/CSS/JS (không cần thư viện ngoài), vì vậy bạn có thể chạy nó ngay lập tức:

Sử dụng Online: https://dn2dn2.github.io/Ribbon-XML-Builder/.

📖 Hướng dẫn sử dụng
1. Cấu hình Tab
Tại panel bên trái, nhập ID và Label cho Tab chính của bạn.

ID: Tên định danh duy nhất (VD: tabMyTools).

Label: Tên hiển thị trên thanh menu (VD: Tiện Ích).

2. Thêm Group và Button
Nhấn + THÊM GROUP MỚI để tạo nhóm chức năng.

Trong mỗi Group, nhấn + Thêm Nút Bấm.

Điền thông tin cho từng nút:

Label: Tên hiển thị của nút.

Icon: Chọn loại Mso (nếu dùng icon Office) hoặc Img (nếu dùng ảnh ngoài). Nhập tên icon vào ô bên cạnh.

Action: Tên hàm Callback trong VBA (VD: SubMyMacro).

Size: Chọn Large (nút to) hoặc Normal (nút nhỏ).

3. Xuất mã XML
Chuyển sang tab XML Code ở panel bên phải.

Nhấn nút ⬇️ Cập Nhật Code.

Copy toàn bộ đoạn mã trong khung đen.

Paste vào file XML trong cấu trúc file Office của bạn (hoặc dùng Custom UI Editor).

4. Import mã cũ (Chỉnh sửa)
Nếu bạn đã có đoạn code XML và muốn sửa giao diện:

Dán code vào khung XML Code.

Nhấn nút ⬆️ Import XML.

Công cụ sẽ vẽ lại giao diện để bạn tiếp tục chỉnh sửa.

🛠️ Công nghệ sử dụng
HTML5: Cấu trúc ngữ nghĩa.

CSS3: Sử dụng biến (:root), Flexbox và CSS Grid cho layout hiện đại, responsive.

JavaScript (ES6): Xử lý logic, DOM manipulation và localStorage.

DOMParser: Dùng để phân tích cú pháp XML khi thực hiện chức năng Import.

🤝 Đóng góp (Contributing)
Mọi đóng góp đều được hoan nghênh! Nếu bạn muốn cải thiện công cụ này

📝 Credits
Ý tưởng và phát triển cốt lõi bởi: Nhất Nguyễn (ThietKeTuDien.vn).

Icon placeholder service: UI Avatars.

Tra cứu ImageMso: Bert Toolkit.
