#requires -Version 5.1

# ============================================================
# xlwings 技能公共工具库（场景 C：宏/VBA 集成）
# 用途：进程追踪 / 编码 / AccessVBOM / WPS 信任 / 注册表注册
# 统一由 build_addin.ps1、apply_vba_to_workbook.ps1、
# import_existing_workbook.ps1 点源加载（. 引用本文件）
# 禁止在业务脚本中再次复制这些实现，修改只允许改本文件
# ============================================================

$ErrorActionPreference = 'Stop'
$OfficeVersions = @('16.0', '15.0', '14.0', '12.0')

# ------------------------------------------------------------
# 编码工具
# ------------------------------------------------------------

function Get-Utf8BomEncoding {
    return New-Object System.Text.UTF8Encoding($true)
}

function Get-Utf8NoBomEncoding {
    return New-Object System.Text.UTF8Encoding($false)
}

function Get-AsciiEncoding {
    return New-Object System.Text.ASCIIEncoding
}

function Remove-LeadingBomCharacters {
    param([string]$Text)

    if ($null -eq $Text) {
        return $Text
    }

    while ($Text.Length -gt 0 -and $Text[0] -eq [char]0xFEFF) {
        $Text = $Text.Substring(1)
    }

    return $Text
}

function Read-TextSmart {
    # 优先按 UTF-8 严格解码；失败时回退系统 ANSI，兼容 VBE 导出的 .frm（系统代码页）
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }

    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        return $strictUtf8.GetString($bytes)
    }
    catch {
        return [System.Text.Encoding]::Default.GetString($bytes)
    }
}

function Read-TextFileAuto {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Write-TextFileUtf8Bom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    [System.IO.File]::WriteAllText($Path, $Content, (Get-Utf8BomEncoding))
}

function Write-TextFileUtf8NoBom {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    [System.IO.File]::WriteAllText($Path, $Content, (Get-Utf8NoBomEncoding))
}

function Write-TextFileAscii {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    [System.IO.File]::WriteAllText($Path, $Content, (Get-AsciiEncoding))
}

function Write-ProjectTextFile {
    # 项目内文本资产统一 UTF-8 无 BOM（与 instantiate_blueprint.ps1 约定一致）
    param(
        [string]$Path,
        [string]$Content
    )

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not [System.IO.Directory]::Exists($parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }

    $normalized = Remove-LeadingBomCharacters -Text $Content
    [System.IO.File]::WriteAllText($Path, $normalized, (Get-Utf8NoBomEncoding))
}

# ------------------------------------------------------------
# Excel 进程追踪（Hwnd + GetWindowThreadProcessId 精确方案，
# 仅杀本次脚本拉起的进程，绝不动用户已有 Excel）
# ------------------------------------------------------------

function Initialize-ExcelProcessResolver {
    if ('ExcelProcessNativeMethods' -as [type]) {
        return
    }

    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class ExcelProcessNativeMethods
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
"@
}

function Resolve-ExcelProcessId {
    param(
        [object]$ExcelApplication,
        [int[]]$BeforeIds = @()
    )

    Initialize-ExcelProcessResolver

    $deadline = (Get-Date).AddSeconds(2)
    do {
        $windowHandle = [IntPtr]::Zero
        try {
            $windowHandle = [IntPtr]::new([int]$ExcelApplication.Hwnd)
        }
        catch {}

        if ($windowHandle -ne [IntPtr]::Zero) {
            $processId = 0
            [void][ExcelProcessNativeMethods]::GetWindowThreadProcessId($windowHandle, [ref]$processId)
            if ($processId -gt 0) {
                return [int]$processId
            }
        }

        Start-Sleep -Milliseconds 50
    } while ((Get-Date) -lt $deadline)

    $fallbackDeadline = (Get-Date).AddSeconds(5)
    do {
        $newProcess = @(
            Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $BeforeIds -notcontains $_.Id } |
            Where-Object {
                $_.ProcessName -match '^(EXCEL|ET|WPS)$' -or
                $_.Path -match '(?i)excel|wps|kingsoft|office6'
            } |
            Sort-Object StartTime -Descending
        ) | Select-Object -First 1

        if ($newProcess) {
            return [int]$newProcess.Id
        }

        Start-Sleep -Milliseconds 100
    } while ((Get-Date) -lt $fallbackDeadline)

    throw '无法定位本次启动的 Excel 进程。'
}

function Start-TrackedExcelApplication {
    # 先记录现有进程，再创建 COM 实例，返回本次新拉起的进程（含 ProcessPath）
    $beforeIds = @(
        Get-Process -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Id
    )

    $excel = New-Object -ComObject Excel.Application
    $processId = Resolve-ExcelProcessId -ExcelApplication $excel -BeforeIds $beforeIds

    $processPath = $null
    try {
        $processPath = (Get-Process -Id $processId -ErrorAction Stop).Path
    }
    catch {
    }

    return [pscustomobject]@{
        Application = $excel
        ProcessId   = $processId
        ProcessPath = $processPath
    }
}

function Stop-TrackedExcelApplication {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TrackedExcel
    )

    if ($null -eq $TrackedExcel) { return }

    $excel = $TrackedExcel.Application
    $processId = $TrackedExcel.ProcessId

    if ($excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        finally {
            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            catch {
            }
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($processId) {
        try {
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    Start-Sleep -Milliseconds 300
}

function Get-OfficeAutomationHostInfo {
    # 有些机器上 Excel.Application 的 COM 绑定会被 WPS 抢占，这里要把真实宿主识别出来。
    $trackedExcel = Start-TrackedExcelApplication
    try {
        $excel = $trackedExcel.Application
        $reportedPath = $null
        try {
            $reportedPath = [string]$excel.Path
        }
        catch {
        }

        $fullPath = $trackedExcel.ProcessPath
        if (-not $fullPath -and $reportedPath) {
            $fullPath = $reportedPath
        }

        $normalizedPath = [string]$fullPath
        $pathToCheck = $normalizedPath.ToLowerInvariant()
        $isWpsHost = $false

        if ($pathToCheck -match 'wps|kingsoft|office6|\\et\.exe$') {
            $isWpsHost = $true
        }

        $caption = ""
        try {
            $caption = [string]$excel.Caption
        }
        catch {
        }

        return [pscustomobject]@{
            Version    = [string]$excel.Version
            Caption    = $caption
            Path       = $normalizedPath
            ProcessId  = $trackedExcel.ProcessId
            IsWpsHost  = $isWpsHost
        }
    }
    finally {
        Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
    }
}

function Get-ExcelVersion {
    return (Get-OfficeAutomationHostInfo).Version
}

function Assert-ValidExcelComHost {
    param(
        [Parameter(Mandatory = $true)]
        [object]$HostInfo
    )

    if ($HostInfo.IsWpsHost) {
        throw ("Excel.Application 当前实际绑定到 WPS 宿主，无法安全用于本脚本。" +
            " 当前宿主路径: {0}。请检查 Excel 的 COM 注册是否被 WPS 占用。" -f $HostInfo.Path)
    }
}

# ------------------------------------------------------------
# AccessVBOM（Excel VBA 工程访问信任）
# ------------------------------------------------------------

function Test-AccessVBOMEnabled {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExcelVersion
    )

    $regPath = "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security"
    if (-not (Test-Path -LiteralPath $regPath)) {
        return $false
    }

    $item = Get-ItemProperty -LiteralPath $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
    return $null -ne $item -and [int]$item.AccessVBOM -eq 1
}

function Enable-AccessVBOM {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExcelVersion
    )

    $regPath = "HKCU:\Software\Microsoft\Office\$ExcelVersion\Excel\Security"
    if (-not (Test-Path -LiteralPath $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    New-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -PropertyType DWord -Force | Out-Null
}

function Test-AccessVBOMEnabledForVersions {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ExcelVersions
    )

    foreach ($excelVersion in $ExcelVersions) {
        if (-not (Test-AccessVBOMEnabled -ExcelVersion $excelVersion)) {
            return $false
        }
    }

    return $true
}

function Enable-AccessVBOMForVersions {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ExcelVersions
    )

    foreach ($excelVersion in $ExcelVersions) {
        Enable-AccessVBOM -ExcelVersion $excelVersion
    }
}

function Ensure-AccessVBOMEnabledForVersions {
    # 返回 $true 表示原本已启用；返回 $false 表示本次由脚本代为启用
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ExcelVersions
    )

    $alreadyEnabled = Test-AccessVBOMEnabledForVersions -ExcelVersions $ExcelVersions
    if ($alreadyEnabled) {
        return $true
    }

    Write-Host "未开启 AccessVBOM，正在写入注册表信任设置（仅对后续新启动的 Excel 生效）..."
    Enable-AccessVBOMForVersions -ExcelVersions $ExcelVersions
    return $false
}

# ------------------------------------------------------------
# Excel 加载项注册表注册（用于 xlam 安装/卸载）
# ------------------------------------------------------------

function Get-ExcelRegistryVersions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FallbackVersion
    )

    $versions = New-Object System.Collections.Generic.List[string]

    foreach ($version in $OfficeVersions) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        if (Test-Path -LiteralPath $optionsPath) {
            [void]$versions.Add($version)
        }
    }

    if ($versions.Count -eq 0) {
        [void]$versions.Add($FallbackVersion)
    }

    # 防管道吞 null：return ,$arr 强制保留数组（PowerShell 管道展开坑，见 build_addin.ps1 注释）
    return ,$versions.ToArray()
}

function Ensure-ExcelOpenEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OptionsPath,

        [Parameter(Mandatory = $true)]
        [string]$AddinPath
    )

    $addinName = [System.IO.Path]::GetFileName($AddinPath)
    $current = Get-ItemProperty -LiteralPath $OptionsPath -ErrorAction SilentlyContinue
    $openNames = @()

    if ($current) {
        $openNames = @(
            $current.PSObject.Properties |
            Where-Object { $_.Name -match "^OPEN\d*$" } |
            Sort-Object {
                if ($_.Name -eq "OPEN") {
                    0
                }
                else {
                    [int]$_.Name.Substring(4)
                }
            }
        )
    }

    foreach ($property in $openNames) {
        $valueText = [string]$property.Value
        if ($valueText -like "*$addinPath*" -or $valueText -like "*$addinName*") {
            return
        }
    }

    $slotName = "OPEN"
    $index = 0
    while (Get-ItemProperty -LiteralPath $OptionsPath -Name $slotName -ErrorAction SilentlyContinue) {
        $index++
        $slotName = "OPEN$index"
    }

    New-ItemProperty -LiteralPath $OptionsPath -Name $slotName -Value ('/R "' + $AddinPath + '"') -PropertyType String -Force | Out-Null
}

function Register-ExcelAddinRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AddinPath,

        [Parameter(Mandatory = $true)]
        [string]$FallbackVersion,

        [ValidateSet('AddinManager', 'Open', 'Both')]
        [string]$Mechanism = 'AddinManager'
    )

    foreach ($version in (Get-ExcelRegistryVersions -FallbackVersion $FallbackVersion)) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        $managerPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Add-in Manager"

        if (-not (Test-Path -LiteralPath $optionsPath)) {
            New-Item -Path $optionsPath -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $managerPath)) {
            New-Item -Path $managerPath -Force | Out-Null
        }

        # 单一机制：默认仅 Add-in Manager 注册，与 verify_addin_registered.ps1 的 AddIns2（Installed=True）检查对齐。
        # 不再默认叠加 OPEN（/R），避免同一插件被注册两次导致重复加载。
        if ($Mechanism -in ('AddinManager', 'Both')) {
            New-ItemProperty -LiteralPath $managerPath -Name $AddinPath -Value "" -PropertyType String -Force | Out-Null
        }
        if ($Mechanism -in ('Open', 'Both')) {
            Ensure-ExcelOpenEntry -OptionsPath $optionsPath -AddinPath $AddinPath
        }
    }
}

function Unregister-ExcelAddinRegistry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AddinPath,

        [Parameter(Mandatory = $true)]
        [string]$FallbackVersion
    )

    $addinName = [System.IO.Path]::GetFileName($AddinPath)

    foreach ($version in (Get-ExcelRegistryVersions -FallbackVersion $FallbackVersion)) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        $managerPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Add-in Manager"

        if (Test-Path -LiteralPath $optionsPath) {
            $current = Get-ItemProperty -LiteralPath $optionsPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -match "^OPEN\d*$" })) {
                $valueText = [string]$property.Value
                if ($valueText -like "*$addinPath*" -or $valueText -like "*$addinName*") {
                    Remove-ItemProperty -LiteralPath $optionsPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }

        if (Test-Path -LiteralPath $managerPath) {
            $current = Get-ItemProperty -LiteralPath $managerPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -notlike "PS*" })) {
                if ($property.Name -eq $AddinPath -or $property.Name -like "*\$addinName") {
                    Remove-ItemProperty -LiteralPath $managerPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

# ------------------------------------------------------------
# WPS 表格信任与加载注册
# ------------------------------------------------------------

function Get-WpsRegistryViews {
    $views = @([Microsoft.Win32.RegistryView]::Default)
    if ([Environment]::Is64BitOperatingSystem) {
        $views = @(
            [Microsoft.Win32.RegistryView]::Registry64,
            [Microsoft.Win32.RegistryView]::Registry32
        )
    }

    return ($views | Select-Object -Unique)
}

function Test-WpsProjectTrustEnabled {
    $hasStringTrust = $false
    $hasDwordTrust = $false

    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $settingsKey = $baseKey.OpenSubKey("Software\kingsoft\office\6.0\et\Application Settings", $false)
            if (-not $settingsKey) {
                continue
            }

            try {
                $stringValue = $settingsKey.GetValue("KDEVBProjectTrust", $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $dwordValue = $settingsKey.GetValue("JSIDEProjectTrust", $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $dwordKind = $settingsKey.GetValueKind("JSIDEProjectTrust")

                if ($null -ne $stringValue -and [string]$stringValue -eq "1") {
                    $hasStringTrust = $true
                }

                if ($null -ne $dwordValue -and $dwordKind -eq [Microsoft.Win32.RegistryValueKind]::DWord -and [int]$dwordValue -eq 1) {
                    $hasDwordTrust = $true
                }
            }
            finally {
                $settingsKey.Close()
            }
        }
        finally {
            $baseKey.Close()
        }
    }

    return $hasStringTrust -and $hasDwordTrust
}

function Set-WpsProjectTrust {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enable
    )

    $stringValue = if ($Enable) { "1" } else { "0" }
    $dwordValue = if ($Enable) { 1 } else { 0 }

    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $settingsKey = $baseKey.CreateSubKey("Software\kingsoft\office\6.0\et\Application Settings")
            try {
                # KDEVBProjectTrust 是字符串，JSIDEProjectTrust 必须保持 DWord 类型，不能误写成字符串。
                $settingsKey.SetValue("KDEVBProjectTrust", $stringValue, [Microsoft.Win32.RegistryValueKind]::String)
                $settingsKey.SetValue("JSIDEProjectTrust", $dwordValue, [Microsoft.Win32.RegistryValueKind]::DWord)
            }
            finally {
                $settingsKey.Close()
            }
        }
        finally {
            $baseKey.Close()
        }
    }
}

function Set-WpsLoadMacroRegistration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AddinPath,

        [Parameter(Mandatory = $true)]
        [bool]$Enable
    )

    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $loadMacrosKey = $baseKey.CreateSubKey("Software\kingsoft\office\6.0\et\LoadMacros")
            try {
                if ($Enable) {
                    $loadMacrosKey.SetValue($AddinPath, "1", [Microsoft.Win32.RegistryValueKind]::String)
                }
                else {
                    $loadMacrosKey.DeleteValue($AddinPath, $false)
                }
            }
            finally {
                $loadMacrosKey.Close()
            }
        }
        finally {
            $baseKey.Close()
        }
    }
}

# ------------------------------------------------------------
# OOXML / Zip 工具（Ribbon 注入共用）
# ------------------------------------------------------------

function Get-CustomUIPartInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CustomUIXml
    )

    [xml]$customUiDocument = $CustomUIXml
    $namespaceUri = [string]$customUiDocument.DocumentElement.NamespaceURI
    $isOffice2010OrLater = ($namespaceUri -eq "http://schemas.microsoft.com/office/2009/07/customui")

    return [pscustomobject]@{
        NamespaceUri                   = $namespaceUri
        IsOffice2010                   = $isOffice2010OrLater
        EntryName                      = if ($isOffice2010OrLater) { "customUI/customUI14.xml" } else { "customUI/customUI.xml" }
        PartName                       = if ($isOffice2010OrLater) { "/customUI/customUI14.xml" } else { "/customUI/customUI.xml" }
        RelationshipsEntryName         = if ($isOffice2010OrLater) { "customUI/_rels/customUI14.xml.rels" } else { "customUI/_rels/customUI.xml.rels" }
        AlternateEntryName             = if ($isOffice2010OrLater) { "customUI/customUI.xml" } else { "customUI/customUI14.xml" }
        AlternatePartName              = if ($isOffice2010OrLater) { "/customUI/customUI.xml" } else { "/customUI/customUI14.xml" }
        AlternateRelationshipsEntryName = if ($isOffice2010OrLater) { "customUI/_rels/customUI.xml.rels" } else { "customUI/_rels/customUI14.xml.rels" }
        RelationshipType               = if ($isOffice2010OrLater) { "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" } else { "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" }
    }
}

function Save-XmlDocument {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = Get-Utf8NoBomEncoding
    $settings.Indent = $true
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::Entitize

    $writer = [System.Xml.XmlWriter]::Create($Path, $settings)
    try {
        $Xml.Save($writer)
    }
    finally {
        $writer.Dispose()
    }
}

function Set-ZipTextEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName,

        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($existing) {
        $existing.Delete()
    }

    $entry = $Zip.CreateEntry($EntryName)
    $stream = $entry.Open()
    $writer = New-Object System.IO.StreamWriter($stream, (Get-Utf8NoBomEncoding))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Set-ZipBinaryEntry {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($existing) {
        $existing.Delete()
    }

    $entry = $Zip.CreateEntry($EntryName)
    $entryStream = $entry.Open()
    $fileStream = [System.IO.File]::OpenRead($SourcePath)
    try {
        $fileStream.CopyTo($entryStream)
    }
    finally {
        $fileStream.Dispose()
        $entryStream.Dispose()
    }
}

function Remove-ZipEntryIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,

        [Parameter(Mandatory = $true)]
        [string]$EntryName
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($existing) {
        $existing.Delete()
    }
}

function New-TemporaryDirectory {
    $directoryPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString("N"))
    [void][System.IO.Directory]::CreateDirectory($directoryPath)
    return $directoryPath
}

function Remove-DirectoryIfExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Directory]::Exists($Path)) {
        [System.IO.Directory]::Delete($Path, $true)
    }
}

function New-TemporaryFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $tempPath = [System.IO.Path]::GetTempFileName()
    $targetPath = [System.IO.Path]::ChangeExtension($tempPath, $Extension)
    if ([System.IO.File]::Exists($tempPath)) {
        [System.IO.File]::Delete($tempPath)
    }
    return $targetPath
}

# ------------------------------------------------------------
# COM 对象释放辅助
# ------------------------------------------------------------

function Set-ComPropertyIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return
    }

    try {
        $Target.$PropertyName = $Value
    }
    catch {
    }
}

function Try-SetComPropertyIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Target,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    try {
        $Target.$PropertyName = $Value
        return $true
    }
    catch {
        return $false
    }
}
