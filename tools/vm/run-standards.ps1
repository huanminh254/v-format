# Chạy dưới quyền SYSTEM (UTM guest agent): giải nén add-in rồi chạy user-run-standards.ps1 trong phiên người dùng.
$root = "C:\VFormatTest"
Remove-Item "$root\addin", "$root\run.log" -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive "$root\VFormat.zip" -DestinationPath "$root\addin" -Force
icacls "$root" /grant "*S-1-1-0:(OI)(CI)F" /T /Q | Out-Null
$user = (Get-CimInstance Win32_ComputerSystem).UserName
$a = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File $root\user-run-standards.ps1"
$pr = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "VFormatRunStandards" -Action $a -Principal $pr -Force | Out-Null
Start-ScheduledTask -TaskName "VFormatRunStandards"
