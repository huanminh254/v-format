# Gỡ V-FORMAT khỏi Microsoft Word (per-user). Dữ liệu người dùng (%APPDATA%\VFormat) được giữ lại.
$ErrorActionPreference = "SilentlyContinue"
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host "Hãy đóng Microsoft Word trước khi gỡ cài đặt." -ForegroundColor Yellow
    exit 1
}
Remove-Item "HKCU:\Software\Microsoft\Office\Word\Addins\VFormat.Connect" -Recurse -Force
foreach ($id in @("{FA30749A-6CEA-4764-8B98-CA6CCD4F03C8}", "{70138569-6BD8-4684-A0EE-C3F1BB5AAAFD}")) {
    Remove-Item "HKCU:\Software\Classes\CLSID\$id" -Recurse -Force
    Remove-Item "HKCU:\Software\Classes\WOW6432Node\CLSID\$id" -Recurse -Force
}
Remove-Item "HKCU:\Software\Classes\VFormat.Connect" -Recurse -Force
Remove-Item "HKCU:\Software\Classes\VFormat.TaskPaneHost" -Recurse -Force
Remove-Item (Join-Path $env:LOCALAPPDATA "VFormat\AddIn") -Recurse -Force
Write-Host "Đã gỡ V-FORMAT." -ForegroundColor Green
