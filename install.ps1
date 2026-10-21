# Установка XL2TAB в автозагрузку AutoCAD (Autoloader).
# Копирует бандл в папку ApplicationPlugins текущего пользователя -
# плагин загружается автоматически при каждом запуске AutoCAD.
# Права администратора не нужны. Повторный запуск обновляет установку.

$ErrorActionPreference = "Stop"

$srcDir = $PSScriptRoot
$dstRoot = Join-Path $env:APPDATA "Autodesk\ApplicationPlugins\XL2TAB.bundle"
$dstContents = Join-Path $dstRoot "Contents"

$dll = Join-Path $srcDir "xl2tab.dll"
$cuix = Join-Path $srcDir "XL2TAB.cuix"
$pkgXml = Join-Path $srcDir "bundle\PackageContents.xml"

if (-not (Test-Path $dll)) { throw "Не найден xl2tab.dll рядом со скриптом (соберите build.ps1)" }
if (-not (Test-Path $pkgXml)) { throw "Не найден bundle\PackageContents.xml" }

New-Item -ItemType Directory -Force -Path $dstContents | Out-Null
Copy-Item -Force $pkgXml (Join-Path $dstRoot "PackageContents.xml")
Copy-Item -Force $dll (Join-Path $dstContents "xl2tab.dll")
if (Test-Path $cuix) { Copy-Item -Force $cuix (Join-Path $dstContents "XL2TAB.cuix") }
else { Write-Host "ВНИМАНИЕ: XL2TAB.cuix не найден - панель будет создана через COM (без иконок)" }

Write-Host "Установлено в: $dstRoot"
Write-Host "Версия: $((Get-Item $dll).VersionInfo.FileVersion)"
Write-Host ""
Write-Host "Перезапустите AutoCAD. Команды доступны без NETLOAD:"
Write-Host "  XL2TAB        - палитра настроек"
Write-Host "  XL2TABPROMPT  - пошаговый режим"
Write-Host "  XLTABDIAG2    - диагностика"
Write-Host ""
Write-Host "Обновление: закройте AutoCAD, замените файлы и запустите install.ps1 снова."
