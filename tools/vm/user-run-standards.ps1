# Chạy TRONG PHIÊN NGƯỜI DÙNG (do run-standards.ps1 tạo scheduled task):
# cài lại add-in, mở Word, chuẩn hóa tệp thử theo từng quy chuẩn qua cổng FormatActive của add-in, xuất PDF.
param([string]$InFile = "C:\VFormatTest\bao-cao-T9_1.docx", [string]$Standards = "auto,nd30,ca08,tt68,hd05,hd29")
$root = "C:\VFormatTest"
$out = "$root\word-out"
$log = "$root\run.log"
"=== $(Get-Date) ===" | Out-File $log -Encoding utf8
function L($m) { $m | Out-File $log -Append -Encoding utf8 }
try {
    try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { $w = $null }
    if ($w) {
        New-Item -ItemType Directory "$root\backup" -Force | Out-Null
        foreach ($d in @($w.Documents)) {
            if ($d.Saved -eq $false) { $d.SaveAs2("$root\backup\" + [IO.Path]::GetFileNameWithoutExtension($d.Name) + "-$(Get-Date -Format HHmmss).docx"); L "backup: $($d.Name)" }
        }
        $w.Quit(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w); $w = $null
        Start-Sleep 3
    }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    for ($i = 0; $i -lt 30 -and (Get-Process WINWORD -ErrorAction SilentlyContinue); $i++) { Start-Sleep 1 }
    & "$root\addin\VFormat\install.ps1" -SourceDir "$root\addin\VFormat" *> "$root\addin-install.log"
    $src = (Get-FileHash "$root\addin\VFormat\VFormat.AddIn.dll" -Algorithm MD5).Hash
    $dst = (Get-FileHash "$env:LOCALAPPDATA\VFormat\AddIn\VFormat.AddIn.dll" -Algorithm MD5).Hash
    if ($src -ne $dst) { throw "Add-in chua duoc cai ban moi (installed $dst, expected $src)" }
    L "installed $dst"

    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory $out -Force | Out-Null
    # ProgID Word.Application có thể trỏ sang bản Office cũ: chạy thẳng WINWORD của Office16 rồi bắt qua ROT
    Start-Process "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -ArgumentList "/q", "/w"
    $w = $null
    for ($i = 0; $i -lt 40 -and -not $w; $i++) {
        Start-Sleep 1
        try { $w = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application") } catch { }
    }
    if (-not $w) { throw "Khong bat duoc Word qua ROT" }
    L "word started $($w.Version) $($w.Path)"
    $w.Visible = $true
    $w.DisplayAlerts = 0
    Start-Sleep 3
    foreach ($x in @($w.COMAddIns)) { L "addin: $($x.ProgId) connect=$($x.Connect)" }
    $ai = $null
    foreach ($x in @($w.COMAddIns)) { if ($x.ProgId -eq "VFormat.Connect") { $ai = $x } }
    if (-not $ai) { throw "VFormat.Connect khong co trong COMAddIns" }
    L "addin connected: $($ai.Connect)"
    $vf = $ai.Object
    if (-not $vf) { throw "COMAddIn.Object = null (add-in not loaded)" }
    foreach ($s in $Standards.Split(",")) {
        $f = "$out\$s.docx"
        Copy-Item $InFile $f -Force
        $d = $w.Documents.Open($f)
        $d.Activate()
        L "----- $s -----"
        try {
            L ($vf.FormatActive($s))
            L "REVIEW:"; L ($vf.ReviewActive())
        } catch { L "LOI: $($_.Exception.Message)" }
        $d.Save()
        $d.ExportAsFixedFormat("$out\$s.pdf", 17)
        $d.Close(0)
    }
    $orig = $w.Documents.Open($InFile)
    L "DONE"
} catch { L "LOI: $($_.Exception.Message) @ $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"; L "DONE" }
