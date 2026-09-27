#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$BlueprintName,

    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [Parameter(Mandatory = $true)]
    [string]$ProjectName,

    [string]$ProjectDescription = 'Excel VBA 加载项项目',

    [ValidateSet('xlam', 'xlsm')]
    [string]$OutputKind = 'xlam',

    [switch]$SkipSampleWorkbook,

    [switch]$DryRun,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$utf8Bom = [System.Text.UTF8Encoding]::new($true)

# --- 加载公共工具库（进程追踪/编码/Zip，统一来自 ./lib/excel-com-common.ps1）---
$commonLibPath = Join-Path $PSScriptRoot 'lib\excel-com-common.ps1'
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

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

function Read-Utf8File {
    param([string]$Path)
    $content = [System.IO.File]::ReadAllText($Path, $utf8NoBom)
    return Remove-LeadingBomCharacters -Text $content
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    $parent = Split-Path -Path $Path -Parent
    if ($parent -and -not [System.IO.Directory]::Exists($parent)) {
        [void][System.IO.Directory]::CreateDirectory($parent)
    }

    $extension = [System.IO.Path]::GetExtension($Path)
    $encoding = if ($extension -and $extension.Equals('.ps1', [System.StringComparison]::OrdinalIgnoreCase)) {
        $utf8Bom
    }
    else {
        $utf8NoBom
    }

    $normalizedContent = Remove-LeadingBomCharacters -Text $Content
    [System.IO.File]::WriteAllText($Path, $normalizedContent, $encoding)
}

function Expand-Template {
    param(
        [string]$Content,
        [hashtable]$Tokens
    )

    $result = $Content
    foreach ($key in $Tokens.Keys) {
        $result = $result.Replace("__${key}__", [string]$Tokens[$key])
    }
    return $result
}

function Get-SafeModuleName {
    param([string]$Name)

    $matches = [System.Text.RegularExpressions.Regex]::Matches($Name, '[A-Za-z0-9]+')
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('mod')
    foreach ($match in $matches) {
        $segment = $match.Value
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        [void]$builder.Append($segment.Substring(0, 1).ToUpperInvariant())
        if ($segment.Length -gt 1) {
            [void]$builder.Append($segment.Substring(1))
        }
    }

    if ($builder.Length -le 3) {
        [void]$builder.Append('ExcelProject')
    }

    return $builder.ToString()
}

function Get-BlueprintBuildTokens {
    param(
        [string]$ResolvedBlueprintName,
        [string]$ResolvedProjectName,
        [string]$ResolvedBuilderScriptPath
    )

    $config = switch ($ResolvedBlueprintName) {
        'userform-data-entry-addin' {
            @{
                MODULE_NAME          = 'modUserformDataEntryHelper'
                RIBBON_TAB_ID        = 'userformDataEntryTab'
                RIBBON_TAB_LABEL     = '数据录入助手'
                RIBBON_GROUP_ID      = 'userformDataEntryGroup'
                RIBBON_GROUP_LABEL   = '窗体交互'
                FIRST_BUTTON_ID      = 'showDataEntryWizardButton'
                FIRST_BUTTON_LABEL   = '打开录入窗体'
                FIRST_CALLBACK_NAME  = 'OnShowDataEntryWizardClick'
                SECOND_BUTTON_ID     = 'runDataEntrySmokeButton'
                SECOND_BUTTON_LABEL  = '执行快速示例'
                SECOND_CALLBACK_NAME = 'OnRunDataEntrySmokeClick'
            }
        }
        'listobject-workbench-addin' {
            @{
                MODULE_NAME          = 'modListObjectWorkbench'
                RIBBON_TAB_ID        = 'listObjectWorkbenchTab'
                RIBBON_TAB_LABEL     = '数据工作台'
                RIBBON_GROUP_ID      = 'listObjectWorkbenchGroup'
                RIBBON_GROUP_LABEL   = '导入与汇总'
                FIRST_BUTTON_ID      = 'importNormalizeButton'
                FIRST_BUTTON_LABEL   = '导入并规范化'
                FIRST_CALLBACK_NAME  = 'OnImportAndNormalizeClick'
                SECOND_BUTTON_ID     = 'refreshSummaryButton'
                SECOND_BUTTON_LABEL  = '刷新汇总结果'
                SECOND_CALLBACK_NAME = 'OnRefreshSummaryTableClick'
            }
        }
        default {
            @{
                MODULE_NAME          = 'modMain'
                RIBBON_TAB_ID        = 'mainTab'
                RIBBON_TAB_LABEL     = 'My Addin'
                RIBBON_GROUP_ID      = 'mainGroup'
                RIBBON_GROUP_LABEL   = 'Actions'
                FIRST_BUTTON_ID      = 'btnOne'
                FIRST_BUTTON_LABEL   = 'Run'
                FIRST_CALLBACK_NAME  = 'RunActionOne'
                SECOND_BUTTON_ID     = 'btnTwo'
                SECOND_BUTTON_LABEL  = 'Run 2'
                SECOND_CALLBACK_NAME = 'RunActionTwo'
            }
        }
    }

    $config.PROJECT_NAME = $ResolvedProjectName
    $config.EXCEL_BUILDER_SCRIPT_PATH = $ResolvedBuilderScriptPath

    if ([string]::IsNullOrWhiteSpace($config.MODULE_NAME)) {
        $config.MODULE_NAME = Get-SafeModuleName -Name $ResolvedProjectName
    }

    return $config
}

function Get-BlueprintPlan {
    param(
        [string[]]$BlueprintRoots,
        [string]$ResolvedTargetRoot,
        [hashtable]$Tokens
    )

    $planMap = @{}

    foreach ($blueprintRoot in $BlueprintRoots) {
        Get-ChildItem -Path $blueprintRoot -Recurse -File | ForEach-Object {
            $relativePath = $_.FullName.Substring($blueprintRoot.Length).TrimStart('\')
            $relativePath = $relativePath.Replace('project_name', $Tokens.PROJECT_NAME)

            if ($relativePath.EndsWith('.tmpl')) {
                $relativePath = $relativePath.Substring(0, $relativePath.Length - 5)
                $content = Expand-Template -Content (Read-Utf8File $_.FullName) -Tokens $Tokens
            }
            else {
                $content = Read-Utf8File $_.FullName
            }

            $targetPath = Join-Path $ResolvedTargetRoot $relativePath
            $planMap[$relativePath] = [pscustomobject]@{
                RelativePath = $relativePath
                TargetPath   = $targetPath
                Exists       = [System.IO.File]::Exists($targetPath)
                Content      = $content
            }
        }
    }

    $plan = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($planMap.GetEnumerator() | Sort-Object Key | ForEach-Object { $_.Value })) {
        $plan.Add($item) | Out-Null
    }
    return $plan
}

function Write-SampleWorkbook {
    param(
        [string]$ResolvedTargetRoot,
        [string]$ResolvedBlueprintName
    )

    $sampleDirectory = Join-Path $ResolvedTargetRoot 'dist'
    [void][System.IO.Directory]::CreateDirectory($sampleDirectory)

    $sampleFileName = switch ($ResolvedBlueprintName) {
        'userform-data-entry-addin' { '测试示例-窗体录入工作簿.xlsx' }
        'listobject-workbench-addin' { '测试示例-ListObject工作台.xlsx' }
        default { '测试示例-Excel工具.xlsx' }
    }

    $samplePath = Join-Path $sampleDirectory $sampleFileName
    $trackedExcel = $null
    $excel = $null
    $workbook = $null
    $worksheet = $null
    $summarySheet = $null

    try {
        $trackedExcel = Start-TrackedExcelApplication
        $excel = $trackedExcel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $workbook = $excel.Workbooks.Add()

        switch ($ResolvedBlueprintName) {
            'userform-data-entry-addin' {
                $worksheet = $workbook.Worksheets(1)
                $worksheet.Name = '录入结果'

                $headers = @('客户名称', '项目名称', '负责人', '金额', '优先级', '生成汇总')
                for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
                    $worksheet.Cells(1, $columnIndex + 1).Value2 = $headers[$columnIndex]
                }

                $worksheet.Range('A1:F1').Font.Bold = $true
                $worksheet.Columns('A:F').ColumnWidth = 16
                $worksheet.Range('A2').Value2 = '示例客户有限公司'
                $worksheet.Range('B2').Value2 = '年度经营分析'
                $worksheet.Range('C2').Value2 = '张三'
                $worksheet.Range('D2').Value2 = 128000
                $worksheet.Range('E2').Value2 = '高'
                $worksheet.Range('F2').Value2 = '是'
                $worksheet.Columns('D').NumberFormat = '#,##0.00'
                $worksheet.Range('H1').Value2 = '说明'
                $worksheet.Range('H2').Value2 = '加载 xlam 后，可点击“数据录入助手”中的“打开录入窗体”进行录入。'
                $worksheet.Columns('H').ColumnWidth = 42
            }
            'listobject-workbench-addin' {
                $worksheet = $workbook.Worksheets(1)
                $worksheet.Name = '导入源数据'

                $data = @(
                    @('订单号', '客户名称', '区域', '下单日期', '最近更新时间', '金额', '状态'),
                    @('SO1001', '示例客户A', '华南', '2026-07-01', '2026-07-01 09:30:00', '12880.5', '已完成'),
                    @('SO1002', '示例客户B', '华东', '2026-07-02', '2026-07-02 11:20:00', '8200', '处理中'),
                    @('SO1003', '示例客户A', '华南', '2026-07-03', '2026-07-03 15:40:00', '5600', '已完成'),
                    @('SO1004', '示例客户C', '华北', '2026-07-04', '2026-07-04 10:05:00', '19800', '待确认'),
                    @('SO1005', '示例客户B', '华东', '2026-07-05', '2026-07-05 18:15:00', '9600', '已完成')
                )

                for ($rowIndex = 0; $rowIndex -lt $data.Count; $rowIndex++) {
                    for ($columnIndex = 0; $columnIndex -lt $data[$rowIndex].Count; $columnIndex++) {
                        $worksheet.Cells($rowIndex + 1, $columnIndex + 1).Value2 = $data[$rowIndex][$columnIndex]
                    }
                }

                $worksheet.Columns('A').NumberFormat = '@'
                $worksheet.Columns('D').NumberFormat = 'yyyy-mm-dd'
                $worksheet.Columns('E').NumberFormat = 'yyyy-mm-dd hh:mm:ss'
                $worksheet.Columns('F').NumberFormat = '#,##0.00'
                $worksheet.Range('A1:G1').Font.Bold = $true
                $worksheet.Columns('A:G').AutoFit() | Out-Null

                $summarySheet = $workbook.Worksheets.Add()
                $summarySheet.Name = '使用说明'
                $summarySheet.Range('A1').Value2 = '1. 打开 xlam 后，先点“导入并规范化”'
                $summarySheet.Range('A2').Value2 = '2. 再点“刷新汇总结果”'
                $summarySheet.Range('A3').Value2 = '3. 观察导入结果和客户汇总两个工作表'
            }
            default {
                $worksheet = $workbook.Worksheets(1)
                $worksheet.Name = '数据'
                $worksheet.Range('A1').Value2 = '项目'
                $worksheet.Range('B1').Value2 = '值'
                $worksheet.Range('A2').Value2 = '示例'
                $worksheet.Range('B2').Value2 = '请在此基础上继续扩展业务逻辑'
                $worksheet.Columns('A:B').AutoFit() | Out-Null
            }
        }

        if ([System.IO.File]::Exists($samplePath)) {
            [System.IO.File]::Delete($samplePath)
        }

        $xlsxFormat = 51
        $workbook.SaveAs($samplePath, $xlsxFormat)
    }
    finally {
        foreach ($comObject in @($summarySheet, $worksheet, $workbook)) {
            if ($comObject) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
            }
        }

        if ($trackedExcel) {
            Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
        }
    }
}

$scriptRoot = Split-Path -Parent $PSScriptRoot
$blueprintRoot = Join-Path $scriptRoot ('templates\project-blueprints\' + $BlueprintName)
if (-not (Test-Path -LiteralPath $blueprintRoot -PathType Container)) {
    throw "未找到蓝图：$BlueprintName"
}

$blueprintRoots = New-Object System.Collections.Generic.List[string]
$baseBlueprintRoot = Join-Path $scriptRoot 'templates\internal-bases\excel-addin-core'
if (-not (Test-Path -LiteralPath $baseBlueprintRoot -PathType Container)) {
    throw "未找到内部基础骨架：$baseBlueprintRoot"
}
$blueprintRoots.Add($baseBlueprintRoot) | Out-Null
$blueprintRoots.Add($blueprintRoot) | Out-Null

$resolvedTargetRoot = [System.IO.Path]::GetFullPath($TargetRoot)
$tokens = @{
    PROJECT_NAME        = $ProjectName
    PROJECT_DESCRIPTION = $ProjectDescription
    OUTPUT_KIND         = $OutputKind
}
$builderScriptPath = [System.IO.Path]::GetFullPath((Join-Path $scriptRoot 'scripts\build_addin.ps1'))
$blueprintBuildTokens = Get-BlueprintBuildTokens -ResolvedBlueprintName $BlueprintName -ResolvedProjectName $ProjectName -ResolvedBuilderScriptPath $builderScriptPath
foreach ($key in $blueprintBuildTokens.Keys) {
    $tokens[$key] = $blueprintBuildTokens[$key]
}

$plan = Get-BlueprintPlan -BlueprintRoots $blueprintRoots.ToArray() -ResolvedTargetRoot $resolvedTargetRoot -Tokens $tokens
$conflicts = @($plan | Where-Object { $_.Exists })

if ($DryRun) {
    Write-Host "蓝图预演：$BlueprintName -> $resolvedTargetRoot"
    foreach ($item in $plan) {
        $status = if ($item.Exists) { '已存在' } else { '新建' }
        Write-Host ("[{0}] {1}" -f $status, $item.RelativePath)
    }
    if ($conflicts.Count -gt 0) {
        Write-Host ('检测到冲突文件数量：' + $conflicts.Count)
    }
    return
}

if ($conflicts.Count -gt 0 -and -not $Force) {
    $preview = ($conflicts | Select-Object -First 10 | ForEach-Object { $_.RelativePath }) -join ', '
    throw "目标目录中已存在文件。请先处理冲突，或使用 -Force 覆盖。示例冲突：$preview"
}

foreach ($item in $plan) {
    Write-Utf8File -Path $item.TargetPath -Content $item.Content
}

if (-not $SkipSampleWorkbook) {
    Write-Host '正在生成测试示例工作簿...'
    Write-SampleWorkbook -ResolvedTargetRoot $resolvedTargetRoot -ResolvedBlueprintName $BlueprintName
}
else {
    Write-Host '已跳过测试示例工作簿生成（批量快速实例化模式）。'
}

Write-Host "蓝图已实例化：$BlueprintName -> $resolvedTargetRoot"
if ($conflicts.Count -gt 0) {
    Write-Host ('已覆盖冲突文件数量：' + $conflicts.Count)
}
if (-not $SkipSampleWorkbook) {
    Write-Host 'dist 目录已附带测试示例工作簿，可直接用于冒烟验证。'
}
else {
    Write-Host '当前模式未生成测试示例工作簿；如需示例，请单独实例化蓝图或后续运行冒烟脚本。'
}
Write-Host '建议下一步：先在生成项目上补业务逻辑或直接构建 xlam/xlsm，再用示例工作簿做冒烟验证。'
