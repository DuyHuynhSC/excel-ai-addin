# Hướng Dẫn Chi Tiết Các Bước Triển Khai (Deploy) Lên Vercel & IIS

Vercel và IIS là hai nền tảng lưu trữ tuyệt vời để xuất bản ứng dụng **Trợ Lý Excel AI** phục vụ cho người dùng cuối. 

Dưới đây là 3 phương pháp triển khai dự án từ thư mục gốc của bạn (`D:\Dev\Source Code\excel-ai-addin`).

---

## 🛠️ Chuẩn Bị Trước Khi Deploy
Hãy chắc chắn bạn đã cấu hình tệp cấu hình [manifest.xml](file:///D:/Dev/Source%20Code/excel-ai-addin/manifest.xml) trỏ về địa chỉ máy chủ thật của bạn:
- URL mặc định thử nghiệm đám mây của bạn hiện tại là: **`https://excel-ai-addin-drab.vercel.app/`**
- URL mặc định thử nghiệm IIS nội bộ là: **`https://localhost:4433/`**

---

## CÁCH 1: Triển Khai Nhanh Bằng Vercel CLI (Đám mây công cộng)
Đây là cách nhanh nhất để đưa Add-in lên internet tĩnh, chạy trực tiếp bằng câu lệnh trên cửa sổ dòng lệnh.

### Bước 1: Cài đặt công cụ Vercel CLI
Mở Terminal tại thư mục dự án và chạy lệnh cài đặt công cụ Vercel toàn cầu (nếu máy của bạn chưa cài):
```bash
npm install -g vercel
```

### Bước 2: Đăng nhập vào tài khoản Vercel
Chạy lệnh sau để liên kết terminal với tài khoản Vercel của bạn (chọn đăng nhập qua Email hoặc GitHub):
```bash
vercel login
```

### Bước 3: Liên kết dự án và deploy thử nghiệm (Development)
Chạy lệnh `vercel` tại thư mục gốc dự án để bắt đầu liên kết:
```bash
vercel
```
Hệ thống sẽ hỏi bạn một số câu hỏi, hãy trả lời như sau:
1. *Set up and deploy “D:\Dev\Source Code\excel-ai-addin”?* ➔ Gõ **`y`** (Yes) rồi nhấn Enter.
2. *Which scope do you want to deploy to?* ➔ Nhấn **Enter** chọn tài khoản cá nhân của bạn.
3. *Link to existing project?* ➔ Gõ **`n`** (No) rồi nhấn Enter để tạo dự án mới.
4. *What’s your project name?* ➔ Nhấn **Enter** để lấy tên mặc định (`excel-ai-addin`).
5. *In which directory is your code located?* ➔ Nhấn **Enter** để chọn thư mục hiện tại (`./`).
6. *Want to modify these settings? (Vite auto-detected)* ➔ Gõ **`n`** (No) để sử dụng cấu hình tự động nhận diện của Vite.

*Vercel sẽ tự động tải mã nguồn lên và trả về một link thử nghiệm.*

### Bước 4: Triển khai chính thức (Production Deploy)
Chạy lệnh sau để biên dịch bản chính thức và cập nhật lên địa chỉ thật của bạn:
```bash
vercel --prod
```
*Sau khi chạy xong, địa chỉ `https://excel-ai-addin-drab.vercel.app` sẽ được cập nhật bản mới nhất chứa proxy CORS và giao diện thuật ngữ.*

---

## CÁCH 2: Triển Khai Tự Động Bằng Git Integration (GitHub/GitLab)
Mỗi khi bạn đẩy code mới lên Git, Vercel sẽ tự động build và cập nhật bản mới lên trang web mà bạn không cần gõ lệnh.

### Bước 1: Đẩy mã nguồn lên kho chứa Git của bạn (ví dụ GitHub)
Tạo kho chứa (Repository) mới trên GitHub và đẩy mã nguồn hiện tại lên:
```bash
git init
git add .
git commit -m "Initial commit with glossary and vercel configuration"
git branch -M main
git remote add origin <đường-dẫn-kho-chứa-github-của-bạn>
git push -u origin main
```

### Bước 2: Nhập dự án vào Vercel
1. Truy cập trang quản lý [Vercel Dashboard](https://vercel.com/dashboard).
2. Nhấp chọn nút **Add New...** ➔ Chọn **Project**.
3. Tại danh sách kho chứa Git, tìm tên dự án của bạn và nhấn **Import**.

### Bước 3: Cấu hình build dự án trên Vercel
Vercel sẽ tự động nhận diện đây là dự án **Vite**. Bạn chỉ cần giữ nguyên các cài đặt mặc định:
- **Framework Preset:** `Vite`
- **Build Command:** `npm run build`
- **Output Directory:** `dist`
- Nhấp nút **Deploy** ở cuối trang.

*Quá trình build sẽ diễn ra trong khoảng 30 giây. Kể từ giờ, mỗi khi bạn đẩy code mới lên GitHub, Vercel sẽ tự động cập nhật phiên bản mới nhất cho bạn.*

---

## CÁCH 3: Triển Khai Nội Bộ Bằng IIS (Internet Information Services)
Đây là giải pháp tốt nhất cho môi trường doanh nghiệp (Intranet). Vì máy chủ IIS chạy trực tiếp trong mạng nội bộ của bạn, nó sẽ giải quyết được 100% lỗi DNS và lỗi IP chặn của Gateway mà không cần Vercel Proxy.

Tôi đã chuẩn bị sẵn một tệp script tự động cấu hình IIS mang tên **`deploy-iis.ps1`** ở thư mục gốc của dự án.

### Bước 1: Chạy Script Triển Khai Tự Động
1. Click chuột phải vào tệp **`deploy-iis.ps1`** và chọn **Run with PowerShell** với quyền **Administrator** (Chạy dưới quyền Admin).
2. Script sẽ tự động thực hiện:
   - Kích hoạt tính năng IIS của Windows (nếu chưa bật).
   - Tự động chạy biên dịch dự án (`npm run build`).
   - Tạo tự động chứng chỉ bảo mật SSL (Self-signed certificate) cho `localhost` và tự động import vào Trust Store của máy để trình duyệt/Excel tin cậy ngay lập tức.
   - Tạo mới một Website tên là `ExcelAIAddin` trên IIS chạy cổng bảo mật **`4433`** trỏ thẳng về thư mục `dist` của dự án.
   - Định cấu hình bảo mật phân quyền thư mục.

### Bước 2: Thiết Lập Tính Năng Reverse Proxy Cho IIS (Chỉ cần làm 1 lần)
Để IIS có thể hoạt động như một Proxy chuyển tiếp các yêu cầu tránh lỗi CORS (thay thế cho Vercel):
1. **Cài đặt URL Rewrite:** Tải và cài đặt [URL Rewrite IIS](https://www.iis.net/downloads/microsoft/url-rewrite).
2. **Cài đặt Application Request Routing (ARR):** Tải và cài đặt [Application Request Routing IIS](https://www.iis.net/downloads/microsoft/application-request-routing).
3. **Kích hoạt tính năng Proxy:**
   - Mở công cụ quản lý **IIS Manager** (gõ `iis` trên thanh tìm kiếm Windows).
   - Nhấp đúp vào **Tên Máy Chủ (Server Name)** ở cây thư mục bên trái.
   - Nhấp đúp mở biểu tượng **Application Request Routing Cache** ở khung giữa.
   - Ở cột **Actions** bên phải ngoài cùng, click chọn **Server Proxy Settings...**.
   - Tích chọn vào hộp kiểm **Enable proxy**, sau đó nhấn **Apply** ở trên cùng bên phải để lưu.

*Tệp cấu hình định tuyến thông minh [web.config](file:///D:/Dev/Source%20Code/excel-ai-addin/public/web.config) đã được tôi tạo sẵn trong thư mục dự án và tự động nạp vào IIS để xử lý CORS.*

---

## BƯỚC 5: Phân Phối Cho Máy Khác
Sau khi hoàn tất deploy bằng 1 trong các cách trên:
1. Bạn hãy nén (ZIP) tệp **`manifest.xml`** và tệp **`install.ps1`** thành một gói cài đặt.
2. Gửi cho đồng nghiệp khác.
3. Đồng nghiệp chỉ cần giải nén thư mục, click chuột phải chạy **`install.ps1` bằng PowerShell** là xong!
   *(Lưu ý: Nếu sử dụng cách 3 - IIS, hãy đảm bảo tệp manifest.xml đã được cập nhật URL trỏ về cổng máy chủ IIS của bạn thay vì Vercel).*
