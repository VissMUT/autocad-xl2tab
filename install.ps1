# Установка XL2TAB в автозагрузку AutoCAD (Autoloader).
# Копирует бандл в папку ApplicationPlugins текущего пользователя -
# плагин будет загружаться автоматически при каждом запуске AutoCAD.
# Права администратора не нужны. Повторный запуск обновляет установку.

$ErrorActionPreference = "Stop"

$bundleName = "XL2TAB.bundle"
$srcDir = $PSScriptRoot
$dstRoot = Join-Path $env:APPDATA "Autodesk\ApplicationPlugins\$bundleName"
$dstContents = Join-Path $dstRoot "Contents"

# берём самый свежий по номеру xl2tab_vN.dll из папки репозитория
$dll = Get-ChildItem -Path $srcDir -Filter "xl2tab_v*.dll" |
    ForEach-Object {
        if ($_.Name -match 'xl2tab_v(\d+)\.dll') {
            [pscustomobject]@{ File = $_; Num = [int]$matches[1] }
        }
    } |
    Sort-Object Num -Descending |
    Select-Object -First 1

if (-not $dll) { throw "Не найден файл xl2tab_v*.dll рядом со скриптом" }

$pkgXml = Join-Path $srcDir "bundle\PackageContents.xml"
if (-not (Test-Path $pkgXml)) { throw "Не найден bundle\PackageContents.xml" }

New-Item -ItemType Directory -Force -Path $dstContents | Out-Null
Copy-Item -Force $pkgXml (Join-Path $dstRoot "PackageContents.xml")
Copy-Item -Force $dll.File.FullName (Join-Path $dstContents "xl2tab.dll")

Write-Host "Установлено в: $dstRoot"
Write-Host "DLL: $($dll.File.Name) -> Contents\xl2tab.dll"
Write-Host ""
Write-Host "Перезапустите AutoCAD. Команды будут доступны без NETLOAD:"
Write-Host "  XL2TAB        - палитра настроек"
Write-Host "  XL2TABPROMPT  - пошаговый режим"
Write-Host "  XLTABDIAG2    - диагностика"
Write-Host ""
Write-Host "Обновление версии: замените DLL в папке, закройте AutoCAD, запустите install.ps1 снова."
