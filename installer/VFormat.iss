; Bộ cài Inno Setup (tùy chọn) — biên dịch trên Windows bằng Inno Setup 6:  iscc installer\VFormat.iss
#define AppVersion "1.2.0"
[Setup]
AppId={{B3E1C0A2-5D7E-4C11-9A53-6F0D5B1E7A10}
AppName=V-FORMAT
AppVersion={#AppVersion}
AppPublisher=V-FORMAT
DefaultDirName={localappdata}\VFormat\Setup
PrivilegesRequired=lowest
OutputBaseFilename=VFormat-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayName=V-FORMAT cho Microsoft Word

[Files]
Source: "..\src\VFormat.AddIn\bin\Release\net48\VFormat.AddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "install.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "uninstall.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\install.ps1"" -SourceDir ""{app}"""; Flags: runhidden waituntilterminated; StatusMsg: "Đang đăng ký V-FORMAT với Microsoft Word..."

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\uninstall.ps1"""; Flags: runhidden; RunOnceId: "VFormatUnreg"
