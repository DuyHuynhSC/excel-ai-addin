# Script tự động triển khai và cấu hình Excel Add-in lên máy chủ IIS (Windows)
# Yêu cầu chạy bằng quyền Administrator (Run as Administrator)

$ErrorActionPreference = "Stop"

# Thiết lập mã hóa UTF-8 để hiển thị tiếng Việt Unicode sắc nét trên console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 1. Kiểm tra quyền Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (!$isAdmin) {
    Write-Error "Script này cần chạy dưới quyền Administrator để cấu hình IIS. Vui lòng mở lại PowerShell với quyền Run as Administrator!"
    Exit
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "      TIẾN TRÌNH TRIỂN KHAI ADD-IN LÊN IIS (WINDOWS)" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

# 2. Kiểm tra và kích hoạt IIS nếu chưa bật
Write-Host "Kiểm tra tính năng IIS..."
$iisFeature = Get-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerRole"
if ($iisFeature.State -ne "Enabled") {
    Write-Host "Đang bật tính năng IIS Web Server..." -ForegroundColor Yellow
    Enable-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerRole" -All -NoRestart | Out-Null
}
# Bật thêm tính năng IIS Management Console
Get-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerManagementTools" | ForEach-Object {
    if ($_.State -ne "Enabled") {
        Enable-WindowsOptionalFeature -Online -FeatureName $_.FeatureName -All -NoRestart | Out-Null
    }
}

# 3. Biên dịch dự án (Vite Build)
Write-Host "Đang chạy biên dịch đóng gói dự án (npm run build)..."
$currentDir = Get-Location
npm run build

$physicalPath = Join-Path $currentDir "dist"
if (!(Test-Path $physicalPath)) {
    Write-Error "Không tìm thấy thư mục dist sau khi build. Vui lòng kiểm tra lại quá trình biên dịch!"
    Exit
}

# 4. Tạo hoặc liên kết chứng chỉ SSL tự ký cho HTTPS
Write-Host "Đang tạo chứng chỉ SSL tự ký cho địa chỉ localhost..."
$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "cert:\LocalMachine\My" -FriendlyName "Excel AI Addin Localhost"

# Thêm chứng chỉ vào Trusted Root CAs để trình duyệt/Excel tin cậy không báo lỗi bảo mật
$rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$rootStore.Add($cert)
$rootStore.Close()

# 5. Khởi tạo Website trên IIS
Import-Module WebAdministration
$siteName = "ExcelAIAddin"
$port = 4433  # Cổng HTTPS thử nghiệm tránh xung đột cổng 443 mặc định

# Kiểm tra nếu site cũ tồn tại thì xóa đi để làm mới
if (Get-Website -Name $siteName -ErrorAction SilentlyContinue) {
    Write-Host "Đang cập nhật lại Website $siteName cũ trên IIS..."
    Remove-Website -Name $siteName | Out-Null
}

Write-Host "Đang cấu hình website mới trên IIS (Cổng $port, đường dẫn: $physicalPath)..."
New-Website -Name $siteName -PhysicalPath $physicalPath -Port $port -Protocol "https" -Ssl -Thumbprint $cert.Thumbprint | Out-Null

# Cấu hình quyền truy cập thư mục cho IIS AppPool
$acl = Get-Acl $physicalPath
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit, ObjectInherit", "None", "Allow")
$acl.AddAccessRule($rule)
Set-Acl $physicalPath $acl

# 6. Hướng dẫn cài đặt cấu hình bổ trợ
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "TRIỂN KHAI THÀNH CÔNG LÊN IIS!" -ForegroundColor Green
Write-Host "Đường dẫn Add-in của bạn: https://localhost:$port/taskpane" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "⚠️ LƯU Ý QUAN TRỌNG ĐỂ REVERSE PROXY HOẠT ĐỘNG:" -ForegroundColor Yellow
Write-Host "Để tính năng điều hướng tránh lỗi CORS (Reverse Proxy) chạy được trên IIS, bạn cần:"
Write-Host "1. Tải và cài đặt URL REWRITE:"
Write-Host "   https://www.iis.net/downloads/microsoft/url-rewrite"
Write-Host "2. Tải và cài đặt APPLICATION REQUEST ROUTING (ARR):"
Write-Host "   https://www.iis.net/downloads/microsoft/application-request-routing"
Write-Host "3. Mở IIS Manager -> Chọn Server Name -> Vào mục 'Application Request Routing Cache'"
Write-Host "   -> Nhấp 'Server Proxy Settings' ở cột bên phải -> Tích chọn 'Enable proxy' -> Nhấn 'Apply'."
Write-Host "----------------------------------------------------------"
Write-Host "4. Sau đó, cập nhật tệp manifest.xml của bạn để trỏ về địa chỉ IIS mới:"
Write-Host "   Thay thế các link https://excel-ai-addin-drab.vercel.app thành https://localhost:$port"
Write-Host "==========================================================" -ForegroundColor Green
Read-Host "Nhấn phím Enter để hoàn tất..."
