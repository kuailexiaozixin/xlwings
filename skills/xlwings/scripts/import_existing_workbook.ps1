#requires -Version 5.1

# 现有工作簿导入脚本：
# - 把用户提供的 xlsx/xlsm/xlam/xltm 文件导入为标准 VBA 项目结构
# - xlsm/xlam/xltm 通过 Excel COM 提取全部 VBA 组件（标准模块/类模块/窗体/文档模块）
# - 同时从 OOXML 包中提取 Ribbon XML（customUI/customUI14）
# - 源文件只读不改；副本进入 baseline/，供 apply_vba_to_workbook.ps1 构建使用
# - 生成 baseline/inventory.json 结构清单，供 AI 完整分析，避免"只读一个模块就下结论"
# - AI 禁止绕过本脚本自行编写 COM 提取代码
# - 公共工具（进程追踪/编码/AccessVBOM/Zip）统一来自 ./lib/excel-com-common.ps1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceWorkbook,

    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [string]$ProjectName = '',

    [switch]$DryRun,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# --- 加载公共工具库 ---
$commonLibPath = Join-Path $PSScriptRoot 'lib\excel-com-common.ps1'
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

$SupportedExtensions = @('.xlsx', '.xlsm', '.xlam', '.xltm')

# --- Ribbon 提取（纯 zip 读取，不依赖 Excel）---

function Get-RibbonInventoryFromPackage {
    param([string]$PackagePath)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $result = $null
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
        # customUI14（2010+）优先于 customUI（2007）
        $selectedEntry = $null
        foreach ($entryName in @('customUI/customUI14.xml', 'customUI/customUI.xml')) {
            $entry = $zip.GetEntry($entryName)
            if ($entry) {
                $selectedEntry = $entry
                break
            }
        }

        if ($selectedEntry) {
            $content = $null
            $stream = $selectedEntry.Open()
            try {
                $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
                try {
                    $content = $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
            finally {
                $stream.Dispose()
            }

            # 收集 Ribbon 回调名（onXxx="..." 与 loadImage），供 AI 对照 VBA 侧实现
            $callbacks = New-Object System.Collections.Generic.List[string]
            foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($content, '(?:\bon[A-Za-z]+|\bloadImage)="([^"]+)"')) {
                $value = $match.Groups[1].Value
                if (-not $callbacks.Contains($value)) {
                    [void]$callbacks.Add($value)
                }
            }

            # 按根元素命名空间区分版本：2009/07 为 2010+ 版，其余按 2007 版处理
            $namespaceMatch = [System.Text.RegularExpressions.Regex]::Match($content, '<customUI[^>]*xmlns="([^"]+)"')
            $isOffice2010 = $namespaceMatch.Success -and $namespaceMatch.Groups[1].Value -eq 'http://schemas.microsoft.com/office/2009/07/customui'

            $result = [pscustomobject]@{
                EntryName    = $selectedEntry.FullName
                Content      = $content
                Callbacks    = $callbacks.ToArray()
                IsOffice2010 = $isOffice2010
            }
        }
    }
    finally {
        if ($zip) {
            $zip.Dispose()
        }
    }

    return $result
}

# --- VBA 提取核心 ---

function Get-ComponentTypeName {
    param([int]$ComponentType)

    switch ($ComponentType) {
        1 { return 'StandardModule' }
        2 { return 'ClassModule' }
        3 { return 'UserForm' }
        11 { return 'ActiveXDesigner' }
        100 { return 'DocumentModule' }
        default { return "Unknown($ComponentType)" }
    }
}

function Import-VbaProjectFromWorkbook {
    param(
        [string]$BaselinePath,
        [string]$ResolvedProjectRoot,
        [bool]$HasVba,
        [bool]$IsDryRun
    )

    $components = New-Object System.Collections.Generic.List[object]
    $sheets = New-Object System.Collections.Generic.List[object]
    $definedNamesCount = 0

    if ($HasVba) {
        Ensure-AccessVBOMEnabledForVersions -ExcelVersions $OfficeVersions | Out-Null
    }

    $trackedExcel = $null
    $workbook = $null
    $vbProject = $null
    $tempDirectory = $null

    try {
        $trackedExcel = Start-TrackedExcelApplication
        $excel = $trackedExcel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        # 打开前禁用宏自动执行，防止 Workbook_Open / Auto_Open 干扰提取；不更新外部链接
        try { $excel.AutomationSecurity = 3 } catch { Write-Host '警告：AutomationSecurity 不可用，继续提取。' }

        $workbook = $excel.Workbooks.Open($BaselinePath, 0)

        try { $definedNamesCount = [int]$workbook.Names.Count } catch { $definedNamesCount = 0 }

        $visibleMap = @{ -1 = 'visible'; 0 = 'hidden'; 2 = 'veryHidden' }
        foreach ($sheet in $workbook.Sheets) {
            $visibleValue = -1
            $codeName = ''
            $sheetName = ''
            try { $sheetName = [string]$sheet.Name } catch {}
            try { $codeName = [string]$sheet.CodeName } catch {}
            try { $visibleValue = [int]$sheet.Visible } catch {}

            [void]$sheets.Add([pscustomobject]@{
                Name     = $sheetName
                CodeName = $codeName
                Visible  = if ($visibleMap.ContainsKey($visibleValue)) { $visibleMap[$visibleValue] } else { [string]$visibleValue }
            })
        }

        if (-not $HasVba) {
            # xlsx 没有 VBA 工程，访问 VBProject 会抛 COM 异常，只收集结构清单
            return [pscustomobject]@{
                Components        = $components.ToArray()
                Sheets            = $sheets.ToArray()
                DefinedNamesCount = $definedNamesCount
            }
        }

        $vbProject = $workbook.VBProject
        $vbaProtection = [int]$vbProject.Protection
        if ($vbaProtection -ne 0) {
            throw "VBA 工程已加锁（Protection=$vbaProtection），无法提取代码。请先在 Excel 中解除 VBA 工程保护后重试。"
        }

        foreach ($component in $vbProject.VBComponents) {
            $componentName = [string]$component.Name
            $componentType = [int]$component.Type
            $typeName = Get-ComponentTypeName -ComponentType $componentType

            # 窗体也有 CodeModule（其事件代码），统一读取；.frm 布局另行 Export
            $lineCount = 0
            $codeText = ''
            $codeModule = $component.CodeModule
            $lineCount = [int]$codeModule.CountOfLines
            if ($lineCount -gt 0) {
                $codeText = [string]$codeModule.Lines(1, $lineCount)
            }

            $extractedTo = ''
            switch ($componentType) {
                1 {
                    # 标准模块
                    $extractedTo = "src/modules/$componentName.vba"
                    if (-not $IsDryRun) {
                        Write-ProjectTextFile -Path (Join-Path $ResolvedProjectRoot $extractedTo) -Content $codeText
                    }
                }
                2 {
                    # 类模块
                    $extractedTo = "src/classes/$componentName.vba"
                    if (-not $IsDryRun) {
                        Write-ProjectTextFile -Path (Join-Path $ResolvedProjectRoot $extractedTo) -Content $codeText
                    }
                }
                3 {
                    # 窗体：.frm(+.frx) 保留布局；.code.vba 单独存事件代码，供回写脚本更新
                    $extractedTo = "src/forms/$componentName.frm (+ .frx + .code.vba)"
                    if (-not $IsDryRun) {
                        if (-not $tempDirectory) {
                            $tempDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString('N'))
                            [void][System.IO.Directory]::CreateDirectory($tempDirectory)
                        }

                        $tempFrmPath = Join-Path $tempDirectory "$componentName.frm"
                        $component.Export($tempFrmPath)

                        $targetFrmPath = Join-Path $ResolvedProjectRoot "src/forms/$componentName.frm"
                        Write-ProjectTextFile -Path $targetFrmPath -Content (Read-TextSmart -Path $tempFrmPath)

                        $tempFrxPath = [System.IO.Path]::ChangeExtension($tempFrmPath, '.frx')
                        if ([System.IO.File]::Exists($tempFrxPath)) {
                            $targetFrxPath = Join-Path $ResolvedProjectRoot "src/forms/$componentName.frx"
                            [System.IO.File]::Copy($tempFrxPath, $targetFrxPath, $true)
                        }

                        if ($lineCount -gt 0) {
                            Write-ProjectTextFile -Path (Join-Path $ResolvedProjectRoot "src/forms/$componentName.code.vba") -Content $codeText
                        }
                    }
                }
                100 {
                    # 文档模块（Sheet*/ThisWorkbook）：仅保存代码，构建时按组件名回写
                    $extractedTo = "src/documents/$componentName.vba"
                    if (-not $IsDryRun) {
                        Write-ProjectTextFile -Path (Join-Path $ResolvedProjectRoot $extractedTo) -Content $codeText
                    }
                }
                default {
                    $extractedTo = ''
                    Write-Host "跳过不支持的组件类型：$typeName - $componentName"
                }
            }

            [void]$components.Add([pscustomobject]@{
                Name       = $componentName
                Type       = $typeName
                CodeLines  = $lineCount
                ExtractedTo = $extractedTo
            })
        }
    }
    finally {
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
        }

        if ($tempDirectory -and [System.IO.Directory]::Exists($tempDirectory)) {
            try { [System.IO.Directory]::Delete($tempDirectory, $true) } catch {}
        }

        if ($trackedExcel) {
            Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
        }
    }

    return [pscustomobject]@{
        Components        = $components.ToArray()
        Sheets            = $sheets.ToArray()
        DefinedNamesCount = $definedNamesCount
    }
}

# --- 主流程 ---

$sourceFullPath = [System.IO.Path]::GetFullPath($SourceWorkbook)
if (-not [System.IO.File]::Exists($sourceFullPath)) {
    throw "源文件不存在：$sourceFullPath"
}

$sourceExtension = [System.IO.Path]::GetExtension($sourceFullPath).ToLowerInvariant()
if ($sourceExtension -notin $SupportedExtensions) {
    throw "不支持的文件类型：$sourceExtension。仅支持：$($SupportedExtensions -join ' / ')"
}

$sourceKind = $sourceExtension.TrimStart('.')
$sourceFileName = [System.IO.Path]::GetFileName($sourceFullPath)

if ([string]::IsNullOrWhiteSpace($ProjectName)) {
    $ProjectName = [System.IO.Path]::GetFileNameWithoutExtension($sourceFullPath)
}

$resolvedProjectRoot = [System.IO.Path]::GetFullPath($ProjectRoot)
$skillScriptsDirectory = $PSScriptRoot

$hasVba = ($sourceExtension -in @('.xlsm', '.xlam', '.xltm'))

$plannedDirectories = @(
    'src/modules',
    'src/classes',
    'src/forms',
    'src/documents',
    'src/ribbon',
    'baseline',
    'dist',
    'logs'
)

if ($DryRun) {
    Write-Host "预演模式：将基于 $sourceFileName 创建项目：$resolvedProjectRoot"
    foreach ($relativeDirectory in $plannedDirectories) {
        Write-Host "  [目录] $relativeDirectory/"
    }
    Write-Host "  [文件] baseline/$sourceFileName（源文件副本，构建基线）"
    if ($hasVba) {
        Write-Host "  [提取] VBA 组件（标准模块/类模块/窗体/文档模块）via Excel COM"
    }
    else {
        Write-Host "  [跳过] xlsx 无 VBA 工程，不执行 COM 提取"
    }
    Write-Host "  [提取] Ribbon XML（如存在 customUI/customUI14）"
    Write-Host "  [文件] baseline/inventory.json 结构清单"
    Write-Host "  [文件] README.md"
    return
}

# 冲突检查：目标目录非空时要求 -Force
if ([System.IO.Directory]::Exists($resolvedProjectRoot)) {
    $existingFiles = @(Get-ChildItem -Path $resolvedProjectRoot -Recurse -File)
    if ($existingFiles.Count -gt 0 -and -not $Force) {
        throw "目标目录非空（已有 $($existingFiles.Count) 个文件）。请换空目录，或使用 -Force 覆盖提取产物。"
    }
}

foreach ($relativeDirectory in $plannedDirectories) {
    [void][System.IO.Directory]::CreateDirectory((Join-Path $resolvedProjectRoot $relativeDirectory))
}

# 源文件只读：复制副本到 baseline/，后续一切操作基于副本
$baselinePath = Join-Path $resolvedProjectRoot "baseline/$sourceFileName"
[System.IO.File]::Copy($sourceFullPath, $baselinePath, $true)

# Ribbon 提取（zip 级，xlsx 也检查以防带 customUI 的特殊包）
# 文件名约定：2010+ 命名空间 -> customUI14.xml（本 Skill 标准版本）；2007 命名空间 -> customUI.xml（保真，建议升级）
$ribbonInventory = Get-RibbonInventoryFromPackage -PackagePath $baselinePath
$ribbonRecord = $null
if ($ribbonInventory) {
    $ribbonFileName = if ($ribbonInventory.IsOffice2010) { 'customUI14.xml' } else { 'customUI.xml' }
    Write-ProjectTextFile -Path (Join-Path $resolvedProjectRoot "src/ribbon/$ribbonFileName") -Content $ribbonInventory.Content
    $ribbonRecord = [pscustomobject]@{
        EntryName    = $ribbonInventory.EntryName
        Callbacks    = $ribbonInventory.Callbacks
        ExtractedTo  = "src/ribbon/$ribbonFileName"
        IsOffice2010 = $ribbonInventory.IsOffice2010
    }
    Write-Host "已提取 Ribbon XML：$($ribbonInventory.EntryName) -> src/ribbon/$ribbonFileName（回调 $($ribbonInventory.Callbacks.Count) 个）"
    if (-not $ribbonInventory.IsOffice2010) {
        Write-Host '提示：该 Ribbon 为 2007 版命名空间。建议升级为 2010 版：把根元素 xmlns 改为 http://schemas.microsoft.com/office/2009/07/customui 并把文件重命名为 customUI14.xml，可解锁更多控件能力。'
    }
}
else {
    Write-Host '未发现 Ribbon XML（customUI/customUI14），跳过。'
}

# VBA 提取
$componentInventory = @()
$sheetInventory = @()
$definedNamesCount = 0

if ($hasVba) {
    Write-Host '正在通过 Excel COM 提取 VBA 工程...'
    $extractionResult = Import-VbaProjectFromWorkbook -BaselinePath $baselinePath -ResolvedProjectRoot $resolvedProjectRoot -HasVba:$true -IsDryRun:$false
    $componentInventory = $extractionResult.Components
    $sheetInventory = $extractionResult.Sheets
    $definedNamesCount = $extractionResult.DefinedNamesCount

    $moduleCount = @($componentInventory | Where-Object { $_.Type -eq 'StandardModule' }).Count
    $classCount = @($componentInventory | Where-Object { $_.Type -eq 'ClassModule' }).Count
    $formCount = @($componentInventory | Where-Object { $_.Type -eq 'UserForm' }).Count
    $documentCount = @($componentInventory | Where-Object { $_.Type -eq 'DocumentModule' }).Count
    Write-Host "VBA 提取完成：标准模块 $moduleCount、类模块 $classCount、窗体 $formCount、文档模块 $documentCount、工作表 $($sheetInventory.Count)。"
}
else {
    # xlsx 无 VBA：仍然收集工作表清单（只读枚举，不访问 VBProject）
    Write-Host 'xlsx 无 VBA 工程，仅收集工作簿结构清单...'
    $extractionResult = Import-VbaProjectFromWorkbook -BaselinePath $baselinePath -ResolvedProjectRoot $resolvedProjectRoot -HasVba:$false -IsDryRun:$false
    $sheetInventory = $extractionResult.Sheets
    $definedNamesCount = $extractionResult.DefinedNamesCount
    $componentInventory = @()
    Write-Host "结构清单收集完成：工作表 $($sheetInventory.Count)。"
}

# 结构清单：AI 分析的唯一入口，避免遗漏
$inventory = [pscustomobject]@{
    SourceFile         = $sourceFileName
    SourceKind         = $sourceKind
    SourceFullPath     = $sourceFullPath
    ImportedAtUtc      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    ProjectRoot        = $resolvedProjectRoot
    HasVbaProject      = $hasVba
    BaselineCopy       = "baseline/$sourceFileName"
    Sheets             = $sheetInventory
    DefinedNamesCount  = $definedNamesCount
    VbComponents       = $componentInventory
    Ribbon             = $ribbonRecord
}

$inventoryJson = $inventory | ConvertTo-Json -Depth 5
Write-ProjectTextFile -Path (Join-Path $resolvedProjectRoot 'baseline/inventory.json') -Content $inventoryJson

# README：固化后续构建命令，AI 与用户均按此执行
$applyScriptPath = [System.IO.Path]::GetFullPath((Join-Path $skillScriptsDirectory 'apply_vba_to_workbook.ps1'))
$buildCommand = "powershell -ExecutionPolicy Bypass -File `"$applyScriptPath`" -ProjectRoot `"$resolvedProjectRoot`""

$readmeContent = "# $ProjectName（基于现有文件导入）

## 来源

- 原始文件：$sourceFileName（$sourceKind）
- 基线副本：baseline/$sourceFileName（构建基于此副本，原始文件不会被修改）
- 结构清单：baseline/inventory.json（工作表/组件/回调全景，分析先看它）

## 目录结构

- src/modules/    标准模块
- src/classes/    类模块
- src/forms/      窗体（.frm 布局 + .frx 二进制 + .code.vba 事件代码）
- src/documents/  文档模块代码（Sheet*/ThisWorkbook）
- src/ribbon/     customUI14.xml（2010 版，本 Skill 标准）或 customUI.xml（2007 旧版保真；修改后构建时会回注）
- baseline/       基线副本与 inventory.json
- dist/           构建产物输出
- logs/           日志

## 构建（基于现有文件，产物进入 dist/）

````text
$buildCommand
````

xlsx 底稿会自动另存为 xlsm；xlsm/xlam 底稿保持原格式语义输出到 dist/。
如需从零制作 xlam 加载项（含安装脚本），请使用 skill 的 build_addin.ps1，与本流程无关。

## 修改规则（构建时的合并语义）

- 修改已有模块/类/文档模块：直接编辑 src/ 下对应 .vba，构建时同名组件整体替换代码
- 修改已有窗体的事件代码：编辑 src/forms/X.code.vba（布局 .frm 不动）
- 新增标准模块/类模块：在 src/modules/ 或 src/classes/ 新增 .vba 文件
- 新增窗体：src/forms/ 放 .frm（+同名 .frx）；与基线同名的窗体不会重复导入
- 删除基线组件：从 src/ 删掉对应文件，构建时加 -RemoveMissingModules 显式删除
- Ribbon：编辑 src/ribbon/customUI14.xml（或旧版 customUI.xml）后构建会整体回注；删除该文件则构建时保留基线原有 Ribbon
- 新写 Ribbon 一律用 2010 版：xmlns=http://schemas.microsoft.com/office/2009/07/customui + 文件名 customUI14.xml
- 禁止新建空白工作簿搬运数据；一切改动通过 src/ + baseline/ 组合构建产出
"

Write-ProjectTextFile -Path (Join-Path $resolvedProjectRoot 'README.md') -Content $readmeContent

Write-Host "导入完成：$resolvedProjectRoot"
Write-Host "下一步：阅读 baseline/inventory.json 与 src/ 提取产物，按 README.md 中的命令构建 dist。"
