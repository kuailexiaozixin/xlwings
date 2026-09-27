#requires -Version 5.1

# 现有工作簿回写脚本（与 import_existing_workbook.ps1 配对使用）：
# - 基于 baseline/ 中的基线副本构建，绝不动原始文件
# - xlsx 底稿：Open 后另存为 xlsm（"先另存副本再改"）
# - VBA upsert：src/ 同名组件整体替换代码（改）、新文件名即新增（增）、
#   -RemoveMissingModules 显式删除基线里 src/ 已不存在的组件（删）
# - Ribbon：src/ribbon/customUI*.xml 存在则整体回注，不存在则保留基线原有
# - 产物：dist/<原名>.xlsm（xlam 底稿保持 xlam，不生成安装脚本）
# - 从零制作 xlam 加载项请用 build_addin.ps1，与本脚本职责互不重叠
# - 公共工具（进程追踪/编码/AccessVBOM/Zip）统一来自 ./lib/excel-com-common.ps1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    # 显式删除语义：基线中存在但 src/ 已无对应文件的标准模块/类模块/窗体将被删除
    [switch]$RemoveMissingModules
)

$ErrorActionPreference = 'Stop'

# --- 加载公共工具库 ---
$commonLibPath = Join-Path $PSScriptRoot 'lib\excel-com-common.ps1'
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

# --- VBA upsert 核心 ---

function Get-ExistingVbComponent {
    param(
        [object]$VBProject,
        [string]$Name
    )

    foreach ($component in $VBProject.VBComponents) {
        if ($component.Name -eq $Name) {
            return $component
        }
    }

    return $null
}

function Replace-ComponentCode {
    # 同名组件整体替换代码（改）
    param(
        [object]$Component,
        [string]$Content
    )

    $codeModule = $Component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }

    if (-not [string]::IsNullOrWhiteSpace($Content)) {
        $codeModule.AddFromString($Content)
    }
}

function Import-FormFromFrmFile {
    # 新窗体导入：临时目录拷贝 .frm(+.frx) 后 VBComponents.Import
    param(
        [object]$VBProject,
        [string]$FrmPath,
        [string]$FrxPath
    )

    $tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($tempDirectory)

    try {
        $tempFrmPath = Join-Path $tempDirectory ([System.IO.Path]::GetFileName($FrmPath))
        [System.IO.File]::WriteAllText($tempFrmPath, (Read-TextSmart -Path $FrmPath), (Get-Utf8NoBomEncoding))

        if ($FrxPath -and [System.IO.File]::Exists($FrxPath)) {
            [System.IO.File]::Copy($FrxPath, (Join-Path $tempDirectory ([System.IO.Path]::GetFileName($FrxPath))), $true)
        }

        return $VBProject.VBComponents.Import($tempFrmPath)
    }
    finally {
        try { [System.IO.Directory]::Delete($tempDirectory, $true) } catch {}
    }
}

# --- Ribbon zip 回注（复用公共库的 Save-XmlDocument / Get-CustomUIPartInfo / Set-ZipTextEntry 等）---

function Ensure-ContentTypes {
    param(
        [string]$TempFile,
        [string]$CustomUiPartName
    )

    [xml]$contentTypes = Read-TextSmart -Path $TempFile
    $ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    $staleOverrides = @($contentTypes.Types.Override | Where-Object {
            ($_.PartName -eq "/customUI/customUI.xml" -or $_.PartName -eq "/customUI/customUI14.xml") -and $_.PartName -ne $CustomUiPartName
        })
    foreach ($staleOverride in $staleOverrides) {
        [void]$contentTypes.Types.RemoveChild($staleOverride)
    }

    $customUiOverride = $contentTypes.Types.Override | Where-Object { $_.PartName -eq $CustomUiPartName } | Select-Object -First 1
    if (-not $customUiOverride) {
        $overrideNode = $contentTypes.CreateElement("Override", $ns)
        [void]$overrideNode.SetAttribute("PartName", $CustomUiPartName)
        [void]$overrideNode.SetAttribute("ContentType", "application/xml")
        [void]$contentTypes.Types.AppendChild($overrideNode)
    }
    else {
        [void]$customUiOverride.SetAttribute("ContentType", "application/xml")
    }

    Save-XmlDocument -Xml $contentTypes -Path $TempFile
}

function Ensure-RootRelationships {
    param(
        [string]$TempFile,
        [string]$RelationshipType,
        [string]$CustomUiTarget
    )

    [xml]$rels = Read-TextSmart -Path $TempFile
    $ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    $existingNodes = @($rels.Relationships.Relationship | Where-Object {
            $_.Type -eq "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" -or
            $_.Type -eq "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" -or
            $_.Id -eq "rIdCustomUI"
        })

    foreach ($node in $existingNodes) {
        [void]$rels.Relationships.RemoveChild($node)
    }

    $relationshipNode = $rels.CreateElement("Relationship", $ns)
    [void]$relationshipNode.SetAttribute("Id", "rIdCustomUI")
    [void]$relationshipNode.SetAttribute("Type", $RelationshipType)
    [void]$relationshipNode.SetAttribute("Target", $CustomUiTarget)
    [void]$rels.Relationships.AppendChild($relationshipNode)

    Save-XmlDocument -Xml $rels -Path $TempFile
}

function Update-RibbonInPackage {
    # 把 src/ribbon 的 customUI XML 回注到 dist 产物；基线已有的 customUI 关系会被替换
    param(
        [string]$PackagePath,
        [string]$CustomUIXml
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $contentTypesTemp = [System.IO.Path]::GetTempFileName()
    $relsTemp = [System.IO.Path]::GetTempFileName()
    $zip = $null

    try {
        $partInfo = Get-CustomUIPartInfo -CustomUIXml $CustomUIXml
        $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)

        $contentTypesEntry = $zip.GetEntry("[Content_Types].xml")
        if (-not $contentTypesEntry) {
            throw "产物缺少 [Content_Types].xml：$PackagePath"
        }

        $relsEntry = $zip.GetEntry("_rels/.rels")
        if (-not $relsEntry) {
            throw "产物缺少 _rels/.rels：$PackagePath"
        }

        foreach ($pair in @(
                @{ Entry = $contentTypesEntry; Temp = $contentTypesTemp },
                @{ Entry = $relsEntry; Temp = $relsTemp }
            )) {
            $stream = $pair.Entry.Open()
            $fileStream = [System.IO.File]::Create($pair.Temp)
            try {
                $stream.CopyTo($fileStream)
            }
            finally {
                $fileStream.Dispose()
                $stream.Dispose()
            }
        }

        Ensure-ContentTypes -TempFile $contentTypesTemp -CustomUiPartName $partInfo.PartName
        Ensure-RootRelationships -TempFile $relsTemp -RelationshipType $partInfo.RelationshipType -CustomUiTarget $partInfo.EntryName

        $contentTypesEntry.Delete()
        $relsEntry.Delete()
        Set-ZipTextEntry -Zip $zip -EntryName "[Content_Types].xml" -Content (Read-TextSmart -Path $contentTypesTemp)
        Set-ZipTextEntry -Zip $zip -EntryName "_rels/.rels" -Content (Read-TextSmart -Path $relsTemp)

        # 只保留一份 customUI（按命名空间决定 entry 名），另一份删除避免双 Ribbon
        Remove-ZipEntryIfExists -Zip $zip -EntryName $partInfo.AlternateEntryName
        Set-ZipTextEntry -Zip $zip -EntryName $partInfo.EntryName -Content $CustomUIXml
    }
    finally {
        if ($zip) {
            $zip.Dispose()
        }

        foreach ($tempFile in @($contentTypesTemp, $relsTemp)) {
            if ([System.IO.File]::Exists($tempFile)) {
                [System.IO.File]::Delete($tempFile)
            }
        }
    }
}

# --- 主流程 ---

$resolvedProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
$baselineDirectory = Join-Path $resolvedProjectRoot 'baseline'
$srcRoot = Join-Path $resolvedProjectRoot 'src'
$distDirectory = Join-Path $resolvedProjectRoot 'dist'

if (-not [System.IO.Directory]::Exists($baselineDirectory)) {
    throw "未找到 baseline/ 目录：$baselineDirectory。请先用 import_existing_workbook.ps1 生成项目。"
}

# 定位基线副本：优先 inventory.json，其次扫描 baseline/ 下的工作簿文件
$baselinePath = ''
$inventoryPath = Join-Path $baselineDirectory 'inventory.json'
if ([System.IO.File]::Exists($inventoryPath)) {
    $inventory = Read-TextSmart -Path $inventoryPath | ConvertFrom-Json
    if ($inventory.BaselineCopy) {
        $candidate = Join-Path $resolvedProjectRoot $inventory.BaselineCopy
        if ([System.IO.File]::Exists($candidate)) {
            $baselinePath = $candidate
        }
    }
}

if (-not $baselinePath) {
    $candidate = @(Get-ChildItem -Path $baselineDirectory -File | Where-Object { $_.Extension -in @('.xlsx', '.xlsm', '.xlam', '.xltm') } | Sort-Object LastWriteTime -Descending) | Select-Object -First 1
    if ($candidate) {
        $baselinePath = $candidate.FullName
    }
}

if (-not $baselinePath) {
    throw "baseline/ 中没有工作簿文件（xlsx/xlsm/xlam/xltm）：$baselineDirectory"
}

$baselineExtension = [System.IO.Path]::GetExtension($baselinePath).TrimStart('.').ToLowerInvariant()
$hasVba = ($baselineExtension -in @('xlsm', 'xlam', 'xltm'))

# 输出格式：xlsx 另存为 xlsm，其余保持原扩展名
$outputExtension = if ($baselineExtension -eq 'xlsx') { 'xlsm' } else { $baselineExtension }
$outputFileFormat = switch ($outputExtension) {
    'xlsm' { 52 }
    'xltm' { 53 }
    'xlam' { 55 }
    default { throw "不支持的输出格式：$outputExtension" }
}

$outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($baselinePath)
$outputPath = Join-Path $distDirectory "$outputBaseName.$outputExtension"
[void][System.IO.Directory]::CreateDirectory($distDirectory)

Write-Host "基线副本：$baselinePath"
Write-Host "产物输出：$outputPath"

# AccessVBOM 预检（改代码必须信任 VBA 工程访问）
Ensure-AccessVBOMEnabledForVersions -ExcelVersions $OfficeVersions | Out-Null

# 收集 src/ 源码文件
$moduleFiles = @{}
$modulesDirectory = Join-Path $srcRoot 'modules'
if ([System.IO.Directory]::Exists($modulesDirectory)) {
    foreach ($file in @(Get-ChildItem -Path $modulesDirectory -File | Where-Object { $_.Extension -in @('.vba', '.bas') })) {
        $moduleFiles[[System.IO.Path]::GetFileNameWithoutExtension($file.Name)] = $file.FullName
    }
}

$classFiles = @{}
$classesDirectory = Join-Path $srcRoot 'classes'
if ([System.IO.Directory]::Exists($classesDirectory)) {
    foreach ($file in @(Get-ChildItem -Path $classesDirectory -File | Where-Object { $_.Extension -in @('.vba', '.cls') })) {
        $classFiles[[System.IO.Path]::GetFileNameWithoutExtension($file.Name)] = $file.FullName
    }
}

$documentFiles = @{}
$documentsDirectory = Join-Path $srcRoot 'documents'
if ([System.IO.Directory]::Exists($documentsDirectory)) {
    foreach ($file in @(Get-ChildItem -Path $documentsDirectory -File | Where-Object { $_.Extension -eq '.vba' })) {
        $documentFiles[[System.IO.Path]::GetFileNameWithoutExtension($file.Name)] = $file.FullName
    }
}

$formFrmFiles = @{}
$formCodeFiles = @{}
$formsDirectory = Join-Path $srcRoot 'forms'
if ([System.IO.Directory]::Exists($formsDirectory)) {
    foreach ($file in @(Get-ChildItem -Path $formsDirectory -File)) {
        if ($file.Extension -ieq '.frm') {
            $formFrmFiles[[System.IO.Path]::GetFileNameWithoutExtension($file.Name)] = $file.FullName
        }
        elseif ($file.Name -like '*.code.vba') {
            $baseName = [System.IO.Path]::GetFileName($file.Name) -replace '\.code\.vba$', ''
            $formCodeFiles[$baseName] = $file.FullName
        }
    }
}

# Ribbon 源：src/ribbon/customUI14.xml 或 customUI.xml
$ribbonXml = ''
foreach ($ribbonCandidate in @('customUI14.xml', 'customUI.xml')) {
    $ribbonPath = Join-Path (Join-Path $srcRoot 'ribbon') $ribbonCandidate
    if ([System.IO.File]::Exists($ribbonPath)) {
        $ribbonXml = Read-TextSmart -Path $ribbonPath
        break
    }
}

if (-not $hasVba -and ($moduleFiles.Count -gt 0 -or $classFiles.Count -gt 0 -or $documentFiles.Count -gt 0 -or $formFrmFiles.Count -gt 0 -or $formCodeFiles.Count -gt 0)) {
    # xlsx 另存为 xlsm 后即有 VBA 工程，属正常场景，不拦截
    Write-Host 'xlsx 底稿将另存为 xlsm 后写入 VBA 工程。'
}

$trackedExcel = $null
$workbook = $null
$vbProject = $null

try {
    $trackedExcel = Start-TrackedExcelApplication
    $excel = $trackedExcel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    try { $excel.AutomationSecurity = 3 } catch { Write-Host '警告：AutomationSecurity 不可用，继续构建。' }

    Write-Host "打开基线副本..."
    $workbook = $excel.Workbooks.Open($baselinePath, 0)

    # --- VBA upsert（作用于内存中的工程；SaveAs 时一次性落盘到 dist）---
    # 注意：实测 SaveAs 之后再 Save() 会把修改写回原路径而非新路径，
    # 因此必须采用"先改内存、后 SaveAs"的顺序，且 SaveAs 之后禁止再调用 Save()
    $vbProject = $workbook.VBProject
    $vbaProtection = [int]$vbProject.Protection
    if ($vbaProtection -ne 0) {
        throw "基线 VBA 工程已加锁（Protection=$vbaProtection），无法回写。请先解除 VBA 工程保护。"
    }

    $changedCount = 0
    $addedCount = 0
    $removedCount = 0

    # 标准模块 / 类模块：同名替换，异名新增
    foreach ($entry in @($moduleFiles.GetEnumerator())) {
        $existing = Get-ExistingVbComponent -VBProject $vbProject -Name $entry.Key
        $content = Read-TextSmart -Path $entry.Value
        if ($existing -and [int]$existing.Type -eq 1) {
            Replace-ComponentCode -Component $existing -Content $content
            $changedCount++
            Write-Host "  [改] 模块 $($entry.Key)"
        }
        elseif (-not $existing) {
            $module = $vbProject.VBComponents.Add(1)
            $module.Name = $entry.Key
            $module.CodeModule.AddFromString($content)
            $addedCount++
            Write-Host "  [增] 模块 $($entry.Key)"
        }
        else {
            throw "组件名冲突：'$($entry.Key)' 在基线中已存在但不是标准模块（实际类型 $($existing.Type)）。请改名后重试。"
        }
    }

    foreach ($entry in @($classFiles.GetEnumerator())) {
        $existing = Get-ExistingVbComponent -VBProject $vbProject -Name $entry.Key
        $content = Read-TextSmart -Path $entry.Value
        if ($existing -and [int]$existing.Type -eq 2) {
            Replace-ComponentCode -Component $existing -Content $content
            $changedCount++
            Write-Host "  [改] 类模块 $($entry.Key)"
        }
        elseif (-not $existing) {
            $classComponent = $vbProject.VBComponents.Add(2)
            $classComponent.Name = $entry.Key
            $classComponent.CodeModule.AddFromString($content)
            $addedCount++
            Write-Host "  [增] 类模块 $($entry.Key)"
        }
        else {
            throw "组件名冲突：'$($entry.Key)' 在基线中已存在但不是类模块（实际类型 $($existing.Type)）。请改名后重试。"
        }
    }

    # 文档模块：只能改，不能增删
    foreach ($entry in @($documentFiles.GetEnumerator())) {
        $existing = Get-ExistingVbComponent -VBProject $vbProject -Name $entry.Key
        if ($existing -and [int]$existing.Type -eq 100) {
            $content = Read-TextSmart -Path $entry.Value
            if (-not [string]::IsNullOrWhiteSpace($content)) {
                Replace-ComponentCode -Component $existing -Content $content
                $changedCount++
                Write-Host "  [改] 文档模块 $($entry.Key)"
            }
        }
        else {
            Write-Warning "  [跳过] 文档模块 $($entry.Key) 在基线中不存在（文档模块不可新建）。"
        }
    }

    # 窗体：基线已有 → 应用 .code.vba（布局不动）；基线没有 → 导入 .frm(+.frx)
    foreach ($entry in @($formFrmFiles.GetEnumerator())) {
        $existing = Get-ExistingVbComponent -VBProject $vbProject -Name $entry.Key
        if ($existing) {
            Write-Warning "  [跳过] 窗体 $($entry.Key) 已存在于基线，布局不覆盖（代码请用 .code.vba 回写）。"
        }
        else {
            $frxPath = [System.IO.Path]::ChangeExtension($entry.Value, '.frx')
            [void](Import-FormFromFrmFile -VBProject $vbProject -FrmPath $entry.Value -FrxPath $frxPath)
            $addedCount++
            Write-Host "  [增] 窗体 $($entry.Key)"
        }
    }

    foreach ($entry in @($formCodeFiles.GetEnumerator())) {
        $existing = Get-ExistingVbComponent -VBProject $vbProject -Name $entry.Key
        if ($existing -and [int]$existing.Type -eq 3) {
            Replace-ComponentCode -Component $existing -Content (Read-TextSmart -Path $entry.Value)
            $changedCount++
            Write-Host "  [改] 窗体代码 $($entry.Key)"
        }
        else {
            Write-Warning "  [跳过] 窗体 $($entry.Key) 不在基线中，.code.vba 无处回写（新窗体请提供 .frm）。"
        }
    }

    # 删除语义：显式开关下，基线组件在 src/ 无对应文件则删（文档模块永不删）
    if ($RemoveMissingModules) {
        $componentNames = New-Object System.Collections.Generic.List[string]
        foreach ($component in $vbProject.VBComponents) {
            [void]$componentNames.Add([string]$component.Name)
        }

        foreach ($componentName in $componentNames) {
            $component = Get-ExistingVbComponent -VBProject $vbProject -Name $componentName
            if (-not $component) { continue }

            $componentType = [int]$component.Type
            if ($componentType -eq 100) { continue }

            $managed = $false
            if ($componentType -eq 1 -and $moduleFiles.ContainsKey($componentName)) { $managed = $true }
            if ($componentType -eq 2 -and $classFiles.ContainsKey($componentName)) { $managed = $true }
            if ($componentType -eq 3 -and ($formFrmFiles.ContainsKey($componentName) -or $formCodeFiles.ContainsKey($componentName))) { $managed = $true }

            if (-not $managed) {
                $vbProject.VBComponents.Remove($component)
                $removedCount++
                Write-Host "  [删] 组件 $componentName（类型 $componentType）"
            }
        }
    }

    # SaveAs 一次性把内存修改落盘到 dist（之后不要再 Save）
    if ([System.IO.File]::Exists($outputPath)) {
        [System.IO.File]::Delete($outputPath)
    }

    $workbook.SaveAs($outputPath, $outputFileFormat)
    Write-Host "VBA 回写完成：改 $changedCount、增 $addedCount、删 $removedCount。"
}
finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }

    if ($trackedExcel) {
        Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
    }
}

# Ribbon 回注：有源则注入，无源保留基线原样
if ($ribbonXml) {
    Write-Host '回注 Ribbon XML...'
    Update-RibbonInPackage -PackagePath $outputPath -CustomUIXml $ribbonXml
    Write-Host 'Ribbon 已回注（src/ribbon -> 产物 customUI）。'
}
else {
    Write-Host 'src/ribbon 无 customUI，保留基线原有 Ribbon。'
}

Write-Host "构建完成：$outputPath"
Write-Host '说明：本产物为工作簿/加载项文件本体，不生成安装脚本；如需 xlam 安装注册，沿用原加载项的安装方式。'
