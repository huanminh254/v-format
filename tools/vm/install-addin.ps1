# Chạy dưới quyền SYSTEM (UTM guest agent): giải nén VFormat.zip rồi chạy install.ps1 TRONG PHIÊN NGƯỜI DÙNG,
# vì add-in đăng ký ở HKCU của người dùng chứ không phải của SYSTEM.
$root = "C:\VFormatTest"
$log = "$root\addin.log"
"=== $(Get-Date) ===" | Out-File $log -Encoding utf8
Remove-Item "$root\addin", "$root\addin-install.log" -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive "$root\VFormat.zip" -DestinationPath "$root\addin" -Force
icacls "$root\addin" /grant "*S-1-1-0:(OI)(CI)F" /T /Q | Out-Null
$word = Get-ChildItem "C:\Program Files\Microsoft Office", "C:\Program Files (x86)\Microsoft Office" -Recurse -Filter WINWORD.EXE -ErrorAction SilentlyContinue | Select-Object -First 1
"Word: $(if ($word) { $word.FullName + ' ' + $word.VersionInfo.FileVersion } else { 'CHUA CO' })" | Out-File $log -Append -Encoding utf8
if (-not $word) { "DONE" | Out-File $log -Append -Encoding utf8; exit 1 }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { "Word dang mo - can dong Word truoc" | Out-File $log -Append -Encoding utf8; "DONE" | Out-File $log -Append -Encoding utf8; exit 2 }
$user = (Get-CimInstance Win32_ComputerSystem).UserName
$inner = "& 'C:\VFormatTest\addin\VFormat\install.ps1' -SourceDir 'C:\VFormatTest\addin\VFormat' *> 'C:\VFormatTest\addin-install.log'; " +
         "Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\VFormat.Connect' | Out-File 'C:\VFormatTest\addin-install.log' -Append; " +
         "'CLSID: ' + (Test-Path 'HKCU:\Software\Classes\CLSID\{FA30749A-6CEA-4764-8B98-CA6CCD4F03C8}\InprocServer32') | Out-File 'C:\VFormatTest\addin-install.log' -Append; " +
         "'FINISHED' | Out-File 'C:\VFormatTest\addin-install.log' -Append"
$enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($inner))
$a = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand $enc"
$pr = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "VFormatAddin" -Action $a -Principal $pr -Force | Out-Null
Start-ScheduledTask -TaskName "VFormatAddin"
"Da chay install.ps1 cho $user" | Out-File $log -Append -Encoding utf8
"DONE" | Out-File $log -Append -Encoding utf8
