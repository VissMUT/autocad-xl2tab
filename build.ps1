# Сборка xl2tab.dll из исходника.
# Требует csc.exe (.NET Framework) и установленный AutoCAD 2019+.
# Если DLL уже загружена в AutoCAD - закройте AutoCAD или смените -out:.

$ErrorActionPreference = "Stop"

$acadDir = "C:\Program Files\Autodesk\AutoCAD 2019"
if (-not (Test-Path "$acadDir\acdbmgd.dll")) {
    Write-Host "Не найден $acadDir\acdbmgd.dll - поправьте переменную acadDir" -ForegroundColor Red
    exit 1
}

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$out = Join-Path $PSScriptRoot "xl2tab.dll"

& $csc /nologo /target:library /out:$out `
    /r:"$acadDir\AcCoreMgd.dll" `
    /r:"$acadDir\acdbmgd.dll" `
    /r:"$acadDir\acmgd.dll" `
    /r:"$acadDir\AdWindows.dll" `
    /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.dll `
    /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll /r:System.Xml.dll `
    /r:"$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\Microsoft.CSharp.dll" `
    /r:"$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\System.Xaml.dll" `
    /r:"$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\WPF\PresentationCore.dll" `
    /r:"$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\WPF\PresentationFramework.dll" `
    /r:"$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\WPF\WindowsBase.dll" `
    (Join-Path $PSScriptRoot "xl2tab.cs")

if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
Write-Host "OK: $out"
