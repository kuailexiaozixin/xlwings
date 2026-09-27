#requires -Version 5.1

[CmdletBinding()]
param(
    [string]$BuilderRoot = '',

    [string]$BlueprintName = 'listobject-workbench-addin',

    [string]$ProjectName = '',

    [string]$OutputRoot = ''
)

$ErrorActionPreference = 'Stop'
$script:CachedScriptBaseDirectory = $null

# --- 加载公共工具库（进程追踪/编码/Zip，统一来自 ./lib/excel-com-common.ps1）---
$commonLibPath = Join-Path $PSScriptRoot 'lib\excel-com-common.ps1'
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

function Get-ScriptBaseDirectory {
    if (-not [string]::IsNullOrWhiteSpace($script:CachedScriptBaseDirectory)) {
        return $script:CachedScriptBaseDirectory
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        [void]$candidates.Add($PSScriptRoot)
    }
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        [void]$candidates.Add((Split-Path -Parent $PSCommandPath))
    }
    if ($script:MyInvocation -and -not [string]::IsNullOrWhiteSpace($script:MyInvocation.MyCommand.Path)) {
        [void]$candidates.Add((Split-Path -Parent $script:MyInvocation.MyCommand.Path))
    }

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and [System.IO.Directory]::Exists($candidate)) {
            $script:CachedScriptBaseDirectory = [System.IO.Path]::GetFullPath($candidate)
            return $script:CachedScriptBaseDirectory
        }
    }

    throw '无法确定脚本所在目录。'
}

function Resolve-FullPath {
    param([string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Remove-DirectoryIfExists {
    param([string]$Path)
    if ([System.IO.Directory]::Exists($Path)) {
        [System.IO.Directory]::Delete($Path, $true)
    }
}

$baseDirectory = Split-Path -Parent (Get-ScriptBaseDirectory)
$resolvedBuilderRoot = if ([string]::IsNullOrWhiteSpace($BuilderRoot)) {
    $baseDirectory
}
else {
    Resolve-FullPath -Path $BuilderRoot
}

$instantiateScript = [System.IO.Path]::Combine($resolvedBuilderRoot, 'scripts', 'instantiate_blueprint.ps1')
if (-not [System.IO.File]::Exists($instantiateScript)) {
    throw "未找到实例化脚本：$instantiateScript"
}

$resolvedOutputRoot = if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    [System.IO.Path]::Combine((Split-Path -Parent $resolvedBuilderRoot), '_smoke_test_projects')
}
else {
    Resolve-FullPath -Path $OutputRoot
}

if ([string]::IsNullOrWhiteSpace($ProjectName)) {
    $ProjectName = switch ($BlueprintName) {
        'userform-data-entry-addin' { 'SmokeUserformDataEntry' }
        default { 'SmokeListObjectWorkbench' }
    }
}

$projectRoot = [System.IO.Path]::Combine($resolvedOutputRoot, $BlueprintName)
Remove-DirectoryIfExists -Path $projectRoot

Write-Host "正在实例化蓝图：$BlueprintName"
& $instantiateScript `
    -BlueprintName $BlueprintName `
    -TargetRoot $projectRoot `
    -ProjectName $ProjectName `
    -OutputKind xlam `
    -Force

$projectBuildScript = [System.IO.Path]::Combine($projectRoot, 'build_addin.ps1')
if (-not [System.IO.File]::Exists($projectBuildScript)) {
    throw "未找到项目构建脚本：$projectBuildScript"
}

Write-Host "正在构建专题蓝图：$ProjectName"
& $projectBuildScript

$addinPath = [System.IO.Path]::Combine($projectRoot, 'dist', ($ProjectName + '.xlam'))
$sampleWorkbookFileName = switch ($BlueprintName) {
    'userform-data-entry-addin' { '测试示例-窗体录入工作簿.xlsx' }
    default { '测试示例-ListObject工作台.xlsx' }
}
$sampleWorkbookPath = [System.IO.Path]::Combine($projectRoot, 'dist', $sampleWorkbookFileName)

if (-not [System.IO.File]::Exists($addinPath)) {
    throw "未找到构建产物：$addinPath"
}

if (-not [System.IO.File]::Exists($sampleWorkbookPath)) {
    throw "未找到测试示例工作簿：$sampleWorkbookPath"
}

$trackedExcel = $null
$excel = $null
$addinWorkbook = $null
$sampleWorkbook = $null
$importSheet = $null
$summarySheet = $null
$importTable = $null
$summaryTable = $null

try {
    $trackedExcel = Start-TrackedExcelApplication
    $excel = $trackedExcel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $addinWorkbook = $excel.Workbooks.Open($addinPath)
    $sampleWorkbook = $excel.Workbooks.Open($sampleWorkbookPath)

    if ($BlueprintName -eq 'userform-data-entry-addin') {
        $excel.Run("'" + $addinWorkbook.Name + "'!SmokeSaveBusinessRecord")
        $importSheet = $sampleWorkbook.Worksheets('录入结果')
        if ($null -eq $importSheet) {
            throw '录入结果工作表未成功生成。'
        }

        if ([string]::IsNullOrWhiteSpace([string]$importSheet.Cells(3, 1).Value2)) {
            throw '窗体数据录入蓝图未成功写入演示记录。'
        }
    }
    else {
        $excel.Run("'" + $addinWorkbook.Name + "'!SmokeImportAndNormalizeCurrentWorkbook")
        $excel.Run("'" + $addinWorkbook.Name + "'!SmokeRefreshSummaryTable")

        $importSheet = $sampleWorkbook.Worksheets('导入结果')
        $summarySheet = $sampleWorkbook.Worksheets('客户汇总')
        $importTable = $importSheet.ListObjects('tblImportedData')
        $summaryTable = $summarySheet.ListObjects('tblCustomerSummary')

        if ($null -eq $importTable -or $importTable.ListRows.Count -lt 1) {
            throw '导入结果表未成功生成。'
        }

        if ($null -eq $summaryTable -or $summaryTable.ListRows.Count -lt 1) {
            throw '客户汇总表未成功生成。'
        }
    }

    $sampleWorkbook.Save()
    Write-Host "冒烟测试通过：$projectRoot"
    Write-Host "产物：$addinPath"
    Write-Host "测试工作簿：$sampleWorkbookPath"
}
finally {
    foreach ($comObject in @($summaryTable, $importTable, $summarySheet, $importSheet, $sampleWorkbook, $addinWorkbook)) {
        if ($comObject) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) } catch {}
        }
    }

    if ($trackedExcel) {
        Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
    }
}
