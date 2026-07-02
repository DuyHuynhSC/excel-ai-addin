# Script tu dong trien khai va cau hinh Excel Add-in len may chu IIS (Windows)
# Script to automatically deploy and configure Excel Add-in on IIS (Windows)
# Yeu cau chay bang quyen Administrator (Run as Administrator)

$ErrorActionPreference = "Stop"

# 1. Kiem tra quyen Administrator (Check Administrator privileges)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (!$isAdmin) {
    Write-Error "Script nay can chay duoi quyen Administrator de cau hinh IIS! (Please run as Administrator!)"
    Exit
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "      TIEN TRINH TRIEN KHAI ADD-IN LEN IIS (WINDOWS)" -ForegroundColor Green
Write-Host "      IIS DEPLOYMENT PROCESS" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green

# 2. Kiem tra va kich hoa IIS (Enable IIS features)
Write-Host "Kiem tra tinh nang IIS... (Checking IIS features...)"
$iisFeature = Get-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerRole"
if ($iisFeature.State -ne "Enabled") {
    Write-Host "Dang bat tinh nang IIS Web Server... (Enabling IIS Web Server...)" -ForegroundColor Yellow
    Enable-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerRole" -All -NoRestart | Out-Null
}
Get-WindowsOptionalFeature -Online -FeatureName "IIS-WebServerManagementTools" | ForEach-Object {
    if ($_.State -ne "Enabled") {
        Enable-WindowsOptionalFeature -Online -FeatureName $_.FeatureName -All -NoRestart | Out-Null
    }
}

# 3. Bien dich du cach (Vite Build)
Write-Host "Dang chay bien dich dong goi du an... (Running npm run build...)"
$currentDir = Get-Location
npm run build

$physicalPath = Join-Path $currentDir "dist"
if (!(Test-Path $physicalPath)) {
    Write-Error "Khong tim thay thu muc dist sau khi build! (dist folder not found!)"
    Exit
}

# 4. Tao hoac lien ket chung chi SSL tu ky (Create Self-Signed Cert)
Write-Host "Dang tao chung chi SSL tu ky cho localhost... (Creating SSL Certificate...)"

# Tao chung chi bao mat cho localhost
$cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "cert:\LocalMachine\My" -FriendlyName "Excel AI Addin Localhost"

# Them vao Trusted Root CAs cua LocalMachine (Trust for Machine)
$rootStore1 = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$rootStore1.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
$rootStore1.Add($cert)
$rootStore1.Close()

# 5. Khoi tao Website tren IIS (Initialize Website on IIS)
Import-Module WebAdministration
$siteName = "ExcelAIAddin"
$port = 4433  # Cong HTTPS thu nghiem
$tempHttpPort = 8084 # Cong HTTP tam thoi dung de khoi tao Website

# Kiem tra neu site cu ton tai thi xoa di (Delete old site if exists)
if (Get-Website -Name $siteName -ErrorAction SilentlyContinue) {
    Write-Host "Dang cap nhat lai Website $siteName cu tren IIS... (Replacing old website...)"
    Remove-Website -Name $siteName | Out-Null
}

Write-Host "Dang cau hinh website moi tren IIS... (Configuring IIS Website...)"
# Khoi tao website voi HTTP tren cong tam thoi de tranh xung dot va tranh loi tham so 'Protocol'
New-Website -Name $siteName -PhysicalPath $physicalPath -Port $tempHttpPort -Force | Out-Null

# Them lien ket HTTPS tren cong $port (4433)
New-WebBinding -Name $siteName -IPAddress "*" -Port $port -Protocol "https" -Force | Out-Null

# Xoa lien ket HTTP tam thoi
Get-WebBinding -Name $siteName -Port $tempHttpPort | Remove-WebBinding -Confirm:$false | Out-Null

# Gan chung chi SSL vao cong HTTPS bang netsh de tuong thich ca PowerShell 7 va 5.1
Write-Host "Gan chung chi SSL vao cong $port bang netsh... (Binding SSL certificate...)"
# Xoa binding cu tren cong nay neu co de tranh loi da ton tai
$null = netsh http delete sslcert ipport=0.0.0.0:$port 2>&1
# Them binding moi
$thumbprint = $cert.Thumbprint
$appid = "{d53fa652-32a8-4c28-98e3-85f02be8d120}"
$null = netsh http add sslcert ipport=0.0.0.0:$port certhash=$thumbprint appid=$appid
# Cau hinh IIS ARR khong giu lai Host Header cua client (tuong duong changeOrigin: true cua Vite) de tranh loi 403/404 tu Gateway
Write-Host "Cau hinh IIS ARR khong giu Host Header... (Configuring IIS ARR proxy...)"
try {
    $null = & "$env:windir\system32\inetsrv\appcmd.exe" set config -section:system.webServer/proxy /preserveHostHeader:"False" /commit:apphost 2>&1
} catch {
    Write-Host "Luu y: Khong the thiet lap preserveHostHeader tren IIS, bo qua... (Skipping ARR host header config...)"
}
# Cau hinh quyen truy cap thu muc cho IIS AppPool va User An danh (IUSR) de tranh loi 401.3
$acl = Get-Acl $physicalPath
$rule1 = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit, ObjectInherit", "None", "Allow")
$rule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("IUSR", "ReadAndExecute", "ContainerInherit, ObjectInherit", "None", "Allow")
$acl.AddAccessRule($rule1)
$acl.AddAccessRule($rule2)
Set-Acl $physicalPath $acl

# 6. Huong dan cau hinh bo tro (Next steps guide)
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "TRIEN KHAI THANH CONG LEN IIS! (IIS DEPLOYMENT SUCCESSFUL!)" -ForegroundColor Green
Write-Host "Add-in URL: https://localhost:$port/taskpane" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "LUU Y QUAN TRONG DE REVERSE PROXY HOAT DONG:" -ForegroundColor Yellow
Write-Host "IMPORTANT NOTES FOR REVERSE PROXY OPERATION:" -ForegroundColor Yellow
Write-Host "1. Tai va cai dat URL REWRITE (Install URL REWRITE):"
Write-Host "   https://www.iis.net/downloads/microsoft/url-rewrite"
Write-Host "2. Tai va cai dat APPLICATION REQUEST ROUTING (ARR) (Install ARR):"
Write-Host "   https://www.iis.net/downloads/microsoft/application-request-routing"
Write-Host "3. Mo IIS Manager -> Server Name -> Application Request Routing Cache"
Write-Host "   -> Server Proxy Settings -> Tich chon 'Enable proxy' -> Apply."
Write-Host "----------------------------------------------------------"
Write-Host "4. Cap nhat manifest.xml (Update manifest.xml):"
Write-Host "   Thay the https://excel-ai-addin-drab.vercel.app thanh https://localhost:$port"
Write-Host "==========================================================" -ForegroundColor Green
Read-Host "Nhan phim Enter de hoan tat... (Press Enter to complete...)"
