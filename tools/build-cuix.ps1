# Сборка XL2TAB.cuix - частичного файла настроек AutoCAD:
# панель инструментов XL2TAB с тремя кнопками, макросами и иконками 16x16 (32bpp, альфа-канал).
# Запуск: powershell -ExecutionPolicy Bypass -File tools\build-cuix.ps1

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$repo = Split-Path -Parent $PSScriptRoot
$out = Join-Path $repo "XL2TAB.cuix"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$xsd = 'xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'

# ---------- иконки 16x16, 32bpp ARGB, прозрачный фон ----------
function MakeIcon([int]$kind) {
    $bmp = New-Object System.Drawing.Bitmap 16, 16, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(235, 235, 235)), 1
    $br = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(235, 235, 235))
    if ($kind -eq 0) {
        $g.DrawRectangle($pen, 2, 3, 11, 9)
        $g.DrawLine($pen, 2, 6, 13, 6)
        $g.DrawLine($pen, 6, 3, 6, 12)
        $g.DrawLine($pen, 9, 3, 9, 12)
    } elseif ($kind -eq 1) {
        foreach ($y in 4, 8, 12) {
            $g.FillRectangle($br, 2, ($y - 1), 2, 2)
            $g.DrawLine($pen, 6, $y, 13, $y)
        }
    } else {
        $g.DrawEllipse($pen, 2, 2, 7, 7)
        $g.DrawLine($pen, 9, 9, 13, 13)
    }
    $pen.Dispose()
    $br.Dispose()
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $bmp.Dispose()
    $bytes = $ms.ToArray()
    $ms.Dispose()
    return $bytes
}

# ---------- части пакета ----------
$header = @"
<?xml version="1.0"?>
<CustSection $xsd>
  <FileVersion MajorVersion="0" MinorVersion="6" IncrementalVersion="1" UserVersion="1" />
  <Header>
    <CommonConfiguration>
      <CommonItems>
        <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
      </CommonItems>
    </CommonConfiguration>
  </Header>
</CustSection>
"@

$menuGroup = @"
<?xml version="1.0"?>
<MenuGroup $xsd Name="XL2TAB" DisplayName="XL2TAB">
  <MacroGroup Name="XL2TAB">
    <MenuMacro UID="MMU_0002">
      <Macro type="Any">
        <Revision MajorVersion="16" MinorVersion="2" UserVersion="1" />
        <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
        <Name xlate="true" UID="XLS_0101">XL2TAB</Name>
        <Command>^C^C_XL2TAB </Command>
        <HelpString xlate="true" UID="XLS_0102">Excel в связанные таблицы AutoCAD - открыть палитру настроек</HelpString>
        <SmallImage Name="xl2tab_table.bmp" />
        <LargeImage Name="xl2tab_table.bmp" />
      </Macro>
    </MenuMacro>
    <MenuMacro UID="MMU_0003">
      <Macro type="Any">
        <Revision MajorVersion="16" MinorVersion="2" UserVersion="1" />
        <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
        <Name xlate="true" UID="XLS_0103">Пошагово</Name>
        <Command>^C^C_XL2TABPROMPT </Command>
        <HelpString xlate="true" UID="XLS_0104">Пошаговый режим без палитры (в командной строке)</HelpString>
        <SmallImage Name="xl2tab_steps.bmp" />
        <LargeImage Name="xl2tab_steps.bmp" />
      </Macro>
    </MenuMacro>
    <MenuMacro UID="MMU_0004">
      <Macro type="Any">
        <Revision MajorVersion="16" MinorVersion="2" UserVersion="1" />
        <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
        <Name xlate="true" UID="XLS_0105">Диагностика</Name>
        <Command>^C^C_XLTABDIAG2 </Command>
        <HelpString xlate="true" UID="XLS_0106">Диагностика DataLink-связей (лог: %TEMP%\xltabdiag_log.txt)</HelpString>
        <SmallImage Name="xl2tab_diag.bmp" />
        <LargeImage Name="xl2tab_diag.bmp" />
      </Macro>
    </MenuMacro>
  </MacroGroup>
</MenuGroup>
"@

$toolbar = @"
<?xml version="1.0"?>
<ToolbarRoot $xsd>
  <Toolbar ToolbarOrient="floating" ToolbarVisible="show" xval="150" yval="120" rows="1" UID="TBU_XL2TAB">
    <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
    <Alias>XL2TAB</Alias>
    <Name xlate="true" UID="XLS_0201">XL2TAB</Name>
    <ToolbarButton IsSeparator="false" UID="TBBU_XL2TAB_1">
      <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
      <Name xlate="true" UID="XLS_0202">XL2TAB</Name>
      <MenuItem>
        <MacroRef MenuMacroID="MMU_0002" />
      </MenuItem>
    </ToolbarButton>
    <ToolbarButton IsSeparator="false" UID="TBBU_XL2TAB_2">
      <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
      <Name xlate="true" UID="XLS_0203">Пошагово</Name>
      <MenuItem>
        <MacroRef MenuMacroID="MMU_0003" />
      </MenuItem>
    </ToolbarButton>
    <ToolbarButton IsSeparator="false" UID="TBBU_XL2TAB_3">
      <ModifiedRev MajorVersion="20" MinorVersion="1" UserVersion="1" />
      <Name xlate="true" UID="XLS_0204">Диагностика</Name>
      <MenuItem>
        <MacroRef MenuMacroID="MMU_0004" />
      </MenuItem>
    </ToolbarButton>
  </Toolbar>
</ToolbarRoot>
"@

$rels = @"
<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="CUI" Target="/Header.cui" Id="R01" /><Relationship Type="CUI" Target="/WorkspaceRoot.cui" Id="R02" /><Relationship Type="CUI" Target="/MenuGroup.cui" Id="R03" /><Relationship Type="CUI" Target="/AcceleratorRoot.cui" Id="R04" /><Relationship Type="CUI" Target="/OverrideRoot.cui" Id="R05" /><Relationship Type="CUI" Target="/MouseButtonRoot.cui" Id="R06" /><Relationship Type="CUI" Target="/PopMenuRoot.cui" Id="R07" /><Relationship Type="CUI" Target="/ToolbarRoot.cui" Id="R08" /><Relationship Type="CUI" Target="/DoubleClickRoot.cui" Id="R09" /><Relationship Type="CUI" Target="/QuickPropertiesRoot.cui" Id="R10" /><Relationship Type="CUI" Target="/RolloverTooltipRoot.cui" Id="R11" /><Relationship Type="CUI" Target="/RibbonRoot.cui" Id="R12" /><Relationship Type="CUI" Target="/QuickAccessToolbarRoot.cui" Id="R13" /><Relationship Type="CUI" Target="/ToolPanelRoot.cui" Id="R14" /><Relationship Type="CUI" Target="/PanelSetRoot.cui" Id="R15" /><Relationship Type="CUI" Target="/ScreenMenuRoot.cui" Id="R16" /><Relationship Type="CUI" Target="/ImageMenuRoot.cui" Id="R17" /><Relationship Type="CUI" Target="/TabletMenuRoot.cui" Id="R18" /><Relationship Type="CUI" Target="/DigitizerButtonRoot.cui" Id="R19" /><Relationship Type="CUI" Target="/LSPFiles.cui" Id="R20" /><Relationship Type="Image" Target="/xl2tab_table.bmp" Id="R21" /><Relationship Type="Image" Target="/xl2tab_steps.bmp" Id="R22" /><Relationship Type="Image" Target="/xl2tab_diag.bmp" Id="R23" /></Relationships>
"@

$types = @"
<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="cui" ContentType="text/xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Default Extension="xml" ContentType="text/xml" /><Default Extension="bmp" ContentType="image/bmp" /></Types>
"@

$stubNames = @(
    'WorkspaceRoot', 'AcceleratorRoot', 'OverrideRoot', 'MouseButtonRoot', 'PopMenuRoot',
    'DoubleClickRoot', 'QuickPropertiesRoot', 'RolloverTooltipRoot', 'RibbonRoot',
    'QuickAccessToolbarRoot', 'ToolPanelRoot', 'PanelSetRoot', 'ScreenMenuRoot',
    'ImageMenuRoot', 'TabletMenuRoot', 'DigitizerButtonRoot', 'LSPFiles'
)

$parts = @('Header.cui') + ($stubNames | ForEach-Object { "$_.cui" }) + @('MenuGroup.cui', 'ToolbarRoot.cui')
$stamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
$pkgInfo = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`r`n<MenuPackageParts>`r`n"
$pkgInfo += ($parts | ForEach-Object { "  <PartData PartData_Name=`"/$_`" PartData_Modified=`"$stamp`" />`r`n" }) -join ''
$pkgInfo += "  <PartData PartData_Name=`"/xl2tab_table.bmp`" PartData_Modified=`"$stamp`" />`r`n"
$pkgInfo += "  <PartData PartData_Name=`"/xl2tab_steps.bmp`" PartData_Modified=`"$stamp`" />`r`n"
$pkgInfo += "  <PartData PartData_Name=`"/xl2tab_diag.bmp`" PartData_Modified=`"$stamp`" />`r`n"
$pkgInfo += "</MenuPackageParts>"

if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }

$zip = [IO.Compression.ZipFile]::Open($out, [IO.Compression.ZipArchiveMode]::Create)
function AddText($name, $text) {
    $e = $zip.CreateEntry($name)
    $s = $e.Open()
    $b = $utf8.GetBytes($text)
    $s.Write($b, 0, $b.Length)
    $s.Close()
}
function AddBytes($name, $bytes) {
    $e = $zip.CreateEntry($name)
    $s = $e.Open()
    $s.Write($bytes, 0, $bytes.Length)
    $s.Close()
}

AddText '_rels/.rels' $rels
AddText '[Content_Types].xml' $types
AddText 'Header.cui' $header
foreach ($n in $stubNames) { AddText "$n.cui" "<$n $xsd />" }
AddText 'MenuGroup.cui' $menuGroup
AddText 'ToolbarRoot.cui' $toolbar
AddText 'Menu_Package_Info.xml' $pkgInfo
AddBytes 'xl2tab_table.bmp' (MakeIcon 0)
AddBytes 'xl2tab_steps.bmp' (MakeIcon 1)
AddBytes 'xl2tab_diag.bmp' (MakeIcon 2)
$zip.Dispose()

Write-Host "OK: $out"
