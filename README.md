# 📋 Trợ Lý Excel AI (Ollama & OpenAI Integration)

Dự án này là một Office Excel Add-in hiệu năng cao, giao diện hiện đại (Dark Glassmorphism) được phát triển bằng TypeScript, đóng gói bằng Vite. Add-in cho phép người dùng chạy các mô hình ngôn ngữ lớn (LLM) cục bộ thông qua **Ollama** hoặc kết nối tới **OpenAI-compatible API Gateway** của doanh nghiệp trực tiếp từ bảng tính Excel.

Dự án sử dụng mô hình **Shared Runtime**, giúp giao diện Sidebar (Task Pane) và các hàm công thức tự chọn (Custom Functions) dùng chung một tiến trình, chia sẻ trạng thái kết nối và khóa API được mã hóa an toàn.

---

## ✨ Tính Năng Nổi Bật

1. **Dịch thuật (Translation)**: Dịch thuật nội dung ô đang chọn sang nhiều ngôn ngữ đích với 3 tùy chọn giọng văn (Trung lập, Trang trọng, Thân mật).
2. **Phân tích dữ liệu (Analysis)**: Hỗ trợ phân tích tổng quan, phân tích cảm xúc (Sentiment Analysis), trích xuất các thực thể quan trọng (tên người, địa danh, email...) và làm sạch dữ liệu.
3. **Tóm tắt (Summarization)**: Tóm tắt các đoạn văn bản dài thành đoạn ngắn hoặc gạch đầu dòng ý chính. Có tùy chọn lưu kết quả vào một trang tính riêng mang tên `Summaries` kèm dấu thời gian.
4. **Hàm Excel Tự Chọn (Custom Functions)**: Cung cấp các công thức trực tiếp trong ô bảng tính:
   - `=AI.TRANSLATE(A1, "Tiếng Anh")`
   - `=AI.SUMMARIZE(A1, "bullets")`
   - `=AI.ANALYZE(A1, "sentiment")`
   - `=AI.ASK("Viết lại ngắn gọn hơn", A1)`
5. **Cơ chế Dự phòng (Fallback & Retry)**: Tự động chuyển đổi thông minh giữa Custom API và Ollama cục bộ khi dịch vụ chính gặp lỗi hoặc mất kết nối. Tự động thử lại (retry) với thuật toán tăng thời gian chờ lũy thừa (exponential backoff).
6. **Mã hóa Bảo mật**: Mã hóa khóa API lưu trữ trong trình duyệt bằng thuật toán AES-GCM (256-bit) của Web Crypto API, bảo vệ dữ liệu khỏi các cuộc tấn công script chéo trang (XSS).

---

## 📁 Cấu Trúc Thư Mục

```
excel-ai-addin/
├── manifest.xml                     # Tệp cấu hình Add-in (Quyền hạn, Ribbon, Custom Functions)
├── taskpane.html                    # Giao diện chính của Sidebar
├── taskpane.css                     # Định dạng phong cách Dark Glassmorphism
├── taskpane.ts                      # Xử lý sự kiện giao diện và tương tác với Excel
├── functions.ts                     # Đăng ký và xử lý các hàm công thức Excel Custom Functions
├── src/
│   ├── api/
│   │   ├── ollamaClient.ts          # Tương tác với Ollama API (hỗ trợ streaming)
│   │   ├── customApiClient.ts       # Tương tác với API Gateway tương thích OpenAI
│   │   └── apiFactory.ts            # Quản lý gọi API, cơ chế tự động thử lại và dự phòng
│   ├── services/
│   │   ├── translationService.ts    # Logic định dạng prompt và gọi dịch thuật
│   │   ├── analysisService.ts       # Logic định dạng prompt phân tích dữ liệu
│   │   ├── summarizationService.ts  # Logic định dạng prompt tóm tắt
│   │   └── historyService.ts        # Quản lý nhật ký lịch sử bằng IndexedDB
│   ├── utils/
│   │   ├── excelHelpers.ts          # Đọc/Ghi dữ liệu Excel qua Office.js
│   │   ├── cryptoUtils.ts           # Mã hóa và giải mã khóa API bằng Web Crypto
│   │   ├── configLoader.ts          # Đọc/Ghi cấu hình lưu trữ
│   │   └── validators.ts            # Kiểm tra dữ liệu đầu vào và chống injection
│   └── types/
│       └── index.ts                 # Định nghĩa các kiểu dữ liệu TypeScript
├── tests/
│   ├── validators.test.ts           # Kiểm thử tính hợp lệ đầu vào
│   ├── cryptoUtils.test.ts          # Kiểm thử mã hóa/giải mã khóa
│   └── apiClients.test.ts           # Kiểm thử mock gọi API
├── .env.example                     # Mẫu biến môi trường cấu hình mặc định
├── .gitignore                       # Cấu hình bỏ qua các tệp không cần đẩy lên Git
├── package.json                     # Quản lý thư viện phụ thuộc và kịch bản chạy
└── tsconfig.json                    # Cấu hình trình biên dịch TypeScript
```

---

## 🛠️ Hướng Dẫn Cài Đặt

### 1. Yêu cầu hệ thống
- Máy tính đã cài đặt **Node.js** (Khuyên dùng phiên bản 18+).
- **Ollama** đã cài đặt và đang chạy cục bộ tại địa chỉ mặc định `http://localhost:11434`.
- **Microsoft Excel** phiên bản 2019+ trên Desktop hoặc Excel Online (Office 365).

### 2. Cài đặt các gói phụ thuộc (Dependencies)
Mở terminal tại thư mục dự án và chạy kịch bản cài đặt:
```bash
npm install
```

### 3. Thiết lập biến môi trường
Tạo tệp `.env` từ tệp mẫu:
```bash
cp .env.example .env
```
Thiết lập lại các mô hình mặc định phù hợp với môi trường Ollama của bạn (ví dụ: `phi:2b`, `mistral:7b`, `llama2:7b-chat`).

---

## 🏃 Các Kịch Bản Chạy

### Chạy máy chủ phát triển (Development Server)
Khởi chạy máy chủ phát triển cục bộ của Vite tại cổng `3000`:
```bash
npm run dev
```

### Chạy kiểm thử tự động (Unit Tests)
Khởi chạy bộ kiểm thử Vitest để kiểm tra tính ổn định của mã nguồn:
```bash
npm run test
```

### Biên dịch dự án (Build)
Biên dịch TypeScript và đóng gói mã nguồn tối ưu sang thư mục `dist`:
```bash
npm run build
```

---

## 📑 Hướng Dẫn Sideload Add-in vào Excel

### Cách 1: Tải lên Excel Online (Tiện lợi nhất để kiểm thử)
1. Truy cập [Office.com](https://office.com) và tạo một tệp Excel mới.
2. Trên thanh Ribbon, chọn **Insert (Chèn)** -> **Office Add-ins (Tiện ích bổ sung Office)**.
3. Trong hộp thoại hiện ra, nhấp chọn **Upload My Add-in (Tải lên tiện ích bổ sung của tôi)** ở góc trên bên phải.
4. Chọn tệp [manifest.xml](manifest.xml) trong thư mục gốc của dự án này.
5. Tab **Trợ Lý AI** sẽ xuất hiện trên thanh Ribbon chính.

### Cách 2: Sideload vào Excel Desktop (Windows)
1. Chia sẻ thư mục chứa tệp `manifest.xml` thành một thư mục mạng chung (Network Shared Folder).
2. Mở Excel Desktop, chọn **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...** -> **Trusted Add-in Catalogs**.
3. Thêm đường dẫn mạng của thư mục chung vào trường **Catalog URL**, chọn **Add Catalog** và tích chọn **Show in Menu**.
4. Khởi động lại Excel, chọn **Insert** -> **My Add-ins** -> tab **Shared Folder** và bấm thêm Add-in của bạn.

---

## 💡 Khắc Phục Sự Cố
- **Lỗi kết nối Ollama trên Excel Desktop**: Do chính sách bảo mật Sandbox của Windows WebView2, Excel Desktop có thể chặn kết nối Loopback tới `localhost`. Để giải quyết, hãy mở PowerShell/CMD dưới quyền Administrator và chạy lệnh:
  ```cmd
  CheckNetIsolation LoopbackExempt -a -n="Microsoft.Win32WebViewHost_cw5n1h2txyewy"
  ```
- **Hàm Custom Functions báo lỗi `#VALUE!`**: Hãy chắc chắn rằng bạn đã mở Sidebar Task Pane của Add-in ít nhất một lần để khởi tạo Shared Runtime và nạp các khóa bảo mật cần thiết từ LocalStorage.
