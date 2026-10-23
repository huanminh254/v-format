<#
  Cài đặt V-FORMAT cho Microsoft Word (per-user, không cần quyền Administrator).
  - Chép VFormat.AddIn.dll vào %LOCALAPPDATA%\VFormat\AddIn
  - Đăng ký COM (CLSID/ProgId) trong HKCU cho cả Office 32 bit và 64 bit
  - Khai báo add-in với Word (LoadBehavior = 3: tự nạp khi khởi động)
  Chạy:  powershell -ExecutionPolicy Bypass -File install.ps1
#>
param(
    [string]$SourceDir = ""
)
$ErrorActionPreference = "Stop"

if (-not $SourceDir) {
    $candidates = @(
        $PSScriptRoot,
        (Join-Path $PSScriptRoot "files"),
        (Join-Path $PSScriptRoot "..\src\VFormat.AddIn\bin\Release\net48"),
        (Join-Path $PSScriptRoot "..\src\VFormat.AddIn\bin\Debug\net48")
    )
    $SourceDir = $candidates | Where-Object { Test-Path (Join-Path $_ "VFormat.AddIn.dll") } | Select-Object -First 1
}
if (-not $SourceDir -or -not (Test-Path (Join-Path $SourceDir "VFormat.AddIn.dll"))) {
    throw "Không tìm thấy VFormat.AddIn.dll. Hãy build (dotnet build -c Release) hoặc chỉ định -SourceDir."
}

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host "Microsoft Word đang mở. Hãy lưu tài liệu và đóng Word rồi chạy lại bộ cài." -ForegroundColor Yellow
    exit 1
}

$target = Join-Path $env:LOCALAPPDATA "VFormat\AddIn"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item (Join-Path $SourceDir "VFormat.AddIn.dll") $target -Force
$pdb = Join-Path $SourceDir "VFormat.AddIn.pdb"
if (Test-Path $pdb) { Copy-Item $pdb $target -Force }
$dll = Join-Path $target "VFormat.AddIn.dll"

$asmName = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
$assembly = $asmName.FullName
$version = $asmName.Version.ToString()
$codeBase = ([Uri]$dll).AbsoluteUri

$classes = @(
    @{ Clsid = "{FA30749A-6CEA-4764-8B98-CA6CCD4F03C8}"; ProgId = "VFormat.Connect";       Type = "VFormat.AddIn.Connect";       Control = $false },
    @{ Clsid = "{70138569-6BD8-4684-A0EE-C3F1BB5AAAFD}"; ProgId = "VFormat.TaskPaneHost";  Type = "VFormat.AddIn.TaskPaneHost";  Control = $true }
)

function Set-Key([string]$path, [hashtable]$values) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($k in $values.Keys) {
        if ($k -eq "(default)") { Set-Item -Path $path -Value $values[$k] }
        else { New-ItemProperty -Path $path -Name $k -Value $values[$k] -PropertyType String -Force | Out-Null }
    }
}

foreach ($c in $classes) {
    # ProgId
    Set-Key "HKCU:\Software\Classes\$($c.ProgId)" @{ "(default)" = $c.Type }
    Set-Key "HKCU:\Software\Classes\$($c.ProgId)\CLSID" @{ "(default)" = $c.Clsid }

    # CLSID cho cả view 64 bit và 32 bit (Office 32 bit trên Windows 64 bit)
    foreach ($root in @("HKCU:\Software\Classes\CLSID", "HKCU:\Software\Classes\WOW6432Node\CLSID")) {
        $base = "$root\$($c.Clsid)"
        $inproc = @{
            "(default)"      = "mscoree.dll"
            "ThreadingModel" = "Both"
            "Class"          = $c.Type
            "Assembly"       = $assembly
            "RuntimeVersion" = "v4.0.30319"
            "CodeBase"       = $codeBase
        }
        Set-Key $base @{ "(default)" = $c.Type }
        Set-Key "$base\InprocServer32" $inproc
        $inproc.Remove("(default)"); $inproc.Remove("ThreadingModel")
        Set-Key "$base\InprocServer32\$version" $inproc
        Set-Key "$base\ProgId" @{ "(default)" = $c.ProgId }
        Set-Key "$base\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}" @{}
        if ($c.Control) { Set-Key "$base\Control" @{} }
    }
}

# Khai báo add-in với Word
$addinKey = "HKCU:\Software\Microsoft\Office\Word\Addins\VFormat.Connect"
if (-not (Test-Path $addinKey)) { New-Item -Path $addinKey -Force | Out-Null }
New-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "V-FORMAT" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "Description" -Value "Chuẩn hóa thể thức văn bản Word 1-Click (NĐ 30, TT 68, HD 05, HD 29, Pháp điển, CV 5512)" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

# Gỡ cờ "đã vô hiệu hóa" nếu Word từng tắt add-in do lỗi
foreach ($ver in @("15.0", "16.0")) {
    $res = "HKCU:\Software\Microsoft\Office\$ver\Word\Resiliency\DisabledItems"
    if (Test-Path $res) { Remove-Item $res -Recurse -Force -ErrorAction SilentlyContinue }
}

New-Item -ItemType Directory -Force -Path (Join-Path $env:APPDATA "VFormat\Templates") | Out-Null
Write-Host "Đã cài V-FORMAT $version vào $target" -ForegroundColor Green
Write-Host "Mở Microsoft Word, tab V-FORMAT sẽ xuất hiện trên thanh Ribbon."
