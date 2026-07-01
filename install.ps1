# Script tự động cài đặt / đăng ký Excel Add-in vào Excel Desktop (Windows)
# Chạy script này bằng cách click chuột phải và chọn "Run with PowerShell"

$ErrorActionPreference = "Stop"

# 1. Xác định thư mục chứa manifest.xml
$currentDir = Get-Location
$manifestPath = Join-Path $currentDir "manifest.xml"

if (!(Test-Path $manifestPath)) {
    Write-Error "Không tìm thấy tệp manifest.xml trong thư mục hiện tại. Vui lòng đặt script này trong cùng thư mục với manifest.xml!"
    Exit
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "   TIẾN TRÌNH CÀI ĐẶT / ĐĂNG KÝ TRỢ LÝ EXCEL AI" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "Thư mục chứa manifest: $currentDir"

# 2. Đăng ký thư mục tin cậy (Trusted Catalog) trong Windows Registry cho Excel
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs"

# Tạo khóa WEF nếu chưa tồn tại
if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

# Tạo một GUID ngẫu nhiên để làm định danh cho Catalog này
$guid = [Guid]::NewGuid().ToString("B")
$newCatalogPath = Join-Path $registryPath $guid

Write-Host "Đang tạo đăng ký Registry cho Excel Trusted Catalog..."
New-Item -Path $newCatalogPath -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Url" -Value $currentDir -PropertyType "String" -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Flags" -Value 1 -PropertyType "DWord" -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Id" -Value $guid -PropertyType "String" -Force | Out-Null

Write-Host "Đăng ký thành công!" -ForegroundColor Green
Write-Host "----------------------------------------------------------"
Write-Host "ĐĂNG KÝ HOÀN TẤT. VUI LÒNG THỰC HIỆN CÁC BƯỚC SAU TRÊN EXCEL:" -ForegroundColor Yellow
Write-Host "1. Mở Microsoft Excel (nếu đang mở thì tắt đi mở lại)."
Write-Host "2. Vào tab Insert (Chèn) -> Chọn My Add-ins (Tiện ích bổ sung của tôi)."
Write-Host "3. Chọn tab SHARED FOLDER (Thư mục dùng chung) ở trên cùng."
Write-Host "4. Bạn sẽ thấy 'Trợ Lý Excel AI' xuất hiện tại đây. Nhấp chọn và bấm 'Add' (Thêm)."
Write-Host "5. Tab 'Trợ Lý AI' sẽ hiển thị trên thanh công cụ của bạn!"
Write-Host "==========================================================" -ForegroundColor Green
Read-Host "Nhấn phím Enter để kết thúc..."
