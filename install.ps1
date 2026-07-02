# Script tu dong cai dat / dang ky Excel Add-in vao Excel Desktop (Windows)
# Script to automatically install / register Excel Add-in to Excel Desktop (Windows)

$ErrorActionPreference = "Stop"

# 1. Xac dinh thu muc chua manifest.xml (Find manifest.xml)
$currentDir = Get-Location
$manifestPath = Join-Path $currentDir "manifest.xml"

if (!(Test-Path $manifestPath)) {
    Write-Error "Khong tim thay tep manifest.xml trong thu muc hien tai! (manifest.xml not found!)"
    Exit
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "   TIEN TRINH CAI DAT / DANG KY TRO LY EXCEL AI" -ForegroundColor Green
Write-Host "   EXCEL AI ADD-IN INSTALLATION PROCESS" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "Folder path: $currentDir"

# 2. Dang ky thu muc tin cay (Trusted Catalog) trong Windows Registry (Register Trusted Catalog)
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs"

if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

$guid = [Guid]::NewGuid().ToString("B")
$newCatalogPath = Join-Path $registryPath $guid

Write-Host "Dang tao dang ky Registry cho Excel... (Configuring Registry...)"
New-Item -Path $newCatalogPath -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Url" -Value $currentDir -PropertyType "String" -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Flags" -Value 1 -PropertyType "DWord" -Force | Out-Null
New-ItemProperty -Path $newCatalogPath -Name "Id" -Value $guid -PropertyType "String" -Force | Out-Null

Write-Host "Dang ky thanh cong! (Registration successful!)" -ForegroundColor Green
Write-Host "----------------------------------------------------------"
Write-Host "VUI LONG THUC HIEN CAC BUOC SAU TREN EXCEL:" -ForegroundColor Yellow
Write-Host "PLEASE FOLLOW THESE STEPS IN EXCEL:" -ForegroundColor Yellow
Write-Host "1. Khoi dong lai Microsoft Excel (Close and reopen Excel)."
Write-Host "2. Vao tab Insert -> My Add-ins (Tien ich bo sung cua toi)."
Write-Host "3. Chon tab SHARED FOLDER (Thu muc dung chung) o tren cung."
Write-Host "4. Chon 'Tro Ly Excel AI' va bam 'Add' (Them)."
Write-Host "==========================================================" -ForegroundColor Green
Read-Host "Nhan phim Enter de ket thuc... (Press Enter to finish...)"
