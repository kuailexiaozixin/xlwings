#requires -Version 5.1

<#
.SYNOPSIS
    xlwings customaddin 一键构建脚本
.DESCRIPTION
    桥接 xlwings quickstart 与 xlwings 技能（场景 C）的构建能力，全程自动化。
    流程：
    1. 运行 xlwings quickstart 生成项目骨架
    2. 复制 VBA 模块和 Ribbon XML 到源码目录
    3. 调用 build_addin.ps1 注入 VBA 和 Ribbon
    4. 通过 Python COM 激活官方引擎和配置表
    5. 输出最终 .xlam 加载项
.PARAMETER ProjectName
    项目名称（也是 Python 模块名）
.PARAMETER OutputDir
    输出目录（默认当前目录）
.PARAMETER AddinType
    加载项类型：addin（默认，白标 xlam）或 standard（普通 xlsm）
.PARAMETER WithRibbon
    是否包含自定义 Ribbon
.PARAMETER VbaModuleDir
    VBA 模块源码目录（.vba 或 .bas 文件），可选
.PARAMETER RibbonXmlPath
    Ribbon XML 文件路径（customUI14.xml 或 customUI.xml），可选
.PARAMETER InterpreterWin
    Windows 解释器路径，如 "C:\Python313\python.exe"
.PARAMETER PythonPath
    PYTHONPATH 设置，指向 Python 代码目录
.PARAMETER UdfModules
    UDF 模块名，多个用逗号分隔
.PARAMETER OutputPath
    输出文件路径（可选，默认 dist/<ProjectName>.xlam）
.EXAMPLE
    .\build_xlwings_addin.ps1 -ProjectName "MyTool" -WithRibbon -InterpreterWin "C:\Python313\python.exe"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectName,

    [string]$OutputDir = (Get-Location).Path,

    [ValidateSet("addin", "standard")]
    [string]$AddinType = "addin",

    [switch]$WithRibbon,

    [AllowEmptyString()]
    [string]$VbaModuleDir = "",

    [AllowEmptyString()]
    [string]$RibbonXmlPath = "",

    [AllowEmptyString()]
    [string]$InterpreterWin = "",

    [AllowEmptyString()]
    [string]$PythonPath = "",

    [AllowEmptyString()]
    [string]$UdfModules = "",

    [AllowEmptyString()]
    [string]$OutputPath = ""
)

$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $PSScriptRoot

# ============================================================
# 辅助函数
# ============================================================

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-Error {
    param([string]$Message)
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

# ============================================================
# 第一步：xlwings quickstart 生成骨架
# ============================================================

$projectDir = Join-Path $OutputDir $ProjectName

Write-Info "第一步：执行 xlwings quickstart..."

# 构建 quickstart 参数
$quickstartArgs = @($ProjectName)
if ($AddinType -eq "addin") {
    $quickstartArgs += "--addin"
}
if ($WithRibbon) {
    $quickstartArgs += "--ribbon"
}

# 如果项目目录已存在，先删除
if (Test-Path $projectDir) {
    Remove-Item -Path $projectDir -Recurse -Force
}

# 执行 quickstart
try {
    # 方法1：通过 xlwings CLI
    $xlwingsCmd = "xlwings quickstart $($quickstartArgs -join ' ')"
    Write-Info "  运行: $xlwingsCmd"
    Invoke-Expression $xlwingsCmd 2>&1 | Out-Null
} catch {
    # 方法2：通过 Python 直接调用
    Write-Info "  CLI 不可用，通过 Python 调用..."
    python -c "import xlwings.cli; import sys; sys.argv = ['xlwings', 'quickstart', $($quickstartArgs | ForEach-Object { "'$_'" }) -join ', ']; xlwings.cli.main()" 2>&1 | Out-Null
}

if (-not (Test-Path $projectDir)) {
    Write-Error "quickstart 失败，项目目录未创建"
    exit 1
}

Write-Success "quickstart 骨架已创建: $projectDir"

# 确定生成的 Excel 文件
$excelFile = Get-ChildItem -Path $projectDir -File | Where-Object { $_.Extension -in @('.xlam', '.xlsm') } | Select-Object -First 1
if (-not $excelFile) {
    Write-Error "未找到生成的 Excel 文件"
    exit 1
}
$excelFilePath = $excelFile.FullName
Write-Info "  Excel 文件: $excelFilePath"

# ============================================================
# 第二步：复制 VBA 模块和 Ribbon XML
# ============================================================

if ($VbaModuleDir -and (Test-Path $VbaModuleDir)) {
    Write-Info "第二步：复制 VBA 模块..."

    $srcModulesDir = Join-Path $projectDir "src" "modules"
    $srcRibbonDir = Join-Path $projectDir "src" "ribbon"

    if (-not (Test-Path $srcModulesDir)) {
        New-Item -ItemType Directory -Path $srcModulesDir -Force | Out-Null
    }

    # 复制 .vba 和 .bas 文件
    Get-ChildItem -Path $VbaModuleDir -File | Where-Object { $_.Extension -in @('.vba', '.bas') } | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination (Join-Path $srcModulesDir $_.Name) -Force
        Write-Info "  复制模块: $($_.Name)"
    }

    # 复制 Ribbon XML
    if ($RibbonXmlPath -and (Test-Path $RibbonXmlPath)) {
        if (-not (Test-Path $srcRibbonDir)) {
            New-Item -ItemType Directory -Path $srcRibbonDir -Force | Out-Null
        }
        Copy-Item -Path $RibbonXmlPath -Destination (Join-Path $srcRibbonDir (Split-Path -Leaf $RibbonXmlPath)) -Force
        Write-Info "  复制 Ribbon: $(Split-Path -Leaf $RibbonXmlPath)"
    }

    Write-Success "VBA 模块已复制"
}

# ============================================================
# 第三步：调用 build_addin.ps1 注入 VBA 和 Ribbon
# ============================================================

$buildAddinScript = Join-Path $scriptRoot "scripts" "build_addin.ps1"
if (Test-Path $buildAddinScript) {
    Write-Info "第三步：调用 build_addin.ps1 构建加载项..."

    $buildArgs = @{
        ProjectName = $ProjectName
        OutputKind = if ($AddinType -eq "addin") { "xlam" } else { "xlsm" }
        OutputPath = if ($OutputPath) { $OutputPath } else { Join-Path $projectDir "dist" "$ProjectName.$($excelFile.Extension.TrimStart('.'))" }
    }

    # 如果有 VBA 模块，设置 ProjectSourceRoot
    $srcDir = Join-Path $projectDir "src"
    if (Test-Path $srcDir) {
        $buildArgs.ProjectSourceRoot = $srcDir
    }

    # 构建参数字符串
    $buildArgString = @()
    foreach ($key in $buildArgs.Keys) {
        $buildArgString += "-$key `"$($buildArgs[$key])`""
    }

    $buildCommand = "& `"$buildAddinScript`" $($buildArgString -join ' ')"
    Write-Info "  运行构建脚本..."
    Invoke-Expression $buildCommand

    # 获取构建产物路径
    $distDir = Join-Path $projectDir "dist"
    $builtFile = Get-ChildItem -Path $distDir -File | Where-Object { $_.Extension -in @('.xlam', '.xlsm') } | Select-Object -First 1
    if ($builtFile) {
        $excelFilePath = $builtFile.FullName
        Write-Success "构建产物: $excelFilePath"
    }
} else {
    Write-Info "build_addin.ps1 不存在，跳过 VBA 注入步骤"
    Write-Info "请在 $projectDir 中手动添加 VBA 代码"
}

# ============================================================
# 第四步：通过 Python COM 激活官方引擎和配置表
# ============================================================

if ($AddinType -eq "addin" -and (Test-Path $excelFilePath)) {
    Write-Info "第四步：激活官方引擎和配置表..."

    # 生成 Python 激活脚本
    $activateScript = Join-Path $projectDir "activate_addin.py"
    $activateCode = @"
import xlwings as xw
import sys, os

addin_path = r'$excelFilePath'.replace('/', '\\')
interpreter = r'$InterpreterWin'.replace('/', '\\')
pythonpath = r'$PythonPath'.replace('/', '\\')
udf_modules = '$UdfModules'

app = xw.App(visible=False)
try:
    wb = app.books.open(addin_path)
    xl_app = app.api

    # 激活官方 xlwings.xlam 引擎
    try:
        addin = xl_app.AddIns2("xlwings")
        addin.Installed = True
    except:
        xlwings_dir = os.path.dirname(xw.__file__)
        official_xlam = os.path.join(xlwings_dir, 'addin', 'xlwings.xlam')
        if os.path.exists(official_xlam):
            xl_app.Workbooks.Open(official_xlam)
            xl_app.AddIns2("xlwings").Installed = True

    # 修改 IsAddin 使配置表可见
    wb.api.IsAddin = False

    # 重命名配置表
    for s in wb.sheets:
        if s.name.startswith('_') and s.name.endswith('.conf'):
            s.name = s.name.lstrip('_')
            break

    # 写入配置
    config_sheet = None
    for s in wb.sheets:
        if s.name.endswith('.conf'):
            config_sheet = s
            break

    if config_sheet:
        config = {}
        if interpreter:
            config['Interpreter_Win'] = interpreter
        if pythonpath:
            config['PYTHONPATH'] = pythonpath
        if udf_modules:
            config['UDF Modules'] = udf_modules

        # 整表清空：先清掉所有旧行，避免重复键（如 PYTHONPATH 写两遍会触发 xlwings 引擎 VBA 457 报错）
        config_sheet.range('A1:B50').clear_contents()

        for i, (key, value) in enumerate(config.items(), start=1):
            config_sheet.range(f'A{i}').value = key
            config_sheet.range(f'B{i}').value = value

        print(f'配置表已写入 {len(config)} 项（已先清空旧内容，杜绝重复键）')

    # 改回 IsAddin
    wb.api.IsAddin = True

    wb.save()
    print('加载项激活完成')
finally:
    wb.close()
    app.quit()
"@

    [System.IO.File]::WriteAllText($activateScript, $activateCode, [System.Text.Encoding]::UTF8)

    # 执行激活脚本
    try {
        $pythonOutput = python $activateScript 2>&1
        Write-Info "  Python 输出: $pythonOutput"
        Write-Success "引擎激活和配置表设置完成"
    } catch {
        Write-Error "Python 激活脚本执行失败: $_"
        Write-Info "可在 Excel 中手动打开 $excelFilePath 完成配置"
    }
}

# ============================================================
# 验证门禁：构建产物完整性检查
# ============================================================

Write-Info "验证门禁：构建产物完整性检查..."

$validationScript = Join-Path $projectDir "validate_build.py"
$validationCode = @"
import os, zipfile, sys, re
from collections import Counter

errors = []
xlam_path = r'$excelFilePath'.replace('/', '\\')

# 1. 文件存在性
if not os.path.exists(xlam_path):
    errors.append('产物文件不存在')
else:
    print(f'  文件: {xlam_path} ({os.path.getsize(xlam_path)} 字节)')

# 2. ZIP包完整性
with zipfile.ZipFile(xlam_path, 'r') as z:
    names = z.namelist()
    dupes = {k: v for k, v in Counter(names).items() if v > 1}
    if dupes:
        errors.append(f'ZIP包有重复条目: {dupes}')
    else:
        print(f'  ZIP: 无重复条目')

    # 3. customUI 存在（优先 2010 版 customUI14.xml，兼容 2007 版 customUI.xml）
    if 'customUI/customUI14.xml' in names:
        print(f'  customUI: 存在（2010 版 customUI14.xml）')
    elif 'customUI/customUI.xml' in names:
        print(f'  customUI: 存在（2007 版 customUI.xml）')
    else:
        errors.append('缺少 customUI/customUI14.xml 或 customUI/customUI.xml')

    # 4. _rels关系数
    rels = z.read('_rels/.rels').decode('utf-8')
    cui_count = rels.count('ui/extensibility')
    if cui_count != 1:
        errors.append(f'_rels中customUI关系数={cui_count}')
    else:
        print(f'  _rels: {cui_count}条customUI关系')

    # 5. Content_Types注册
    ct = z.read('[Content_Types].xml').decode('utf-8')
    if 'customUI' not in ct:
        errors.append('Content_Types缺少customUI')
    else:
        print(f'  Content_Types: 已注册')

    # 6. vbaProject存在
    if 'xl/vbaProject.bin' not in names:
        errors.append('缺少vbaProject.bin')
    else:
        print(f'  vbaProject: 存在')

    # 7. 回调名提取
    cui = z.read('customUI/customUI.xml').decode('utf-8')
    ribbon_callbacks = re.findall(r'onAction="([^"]+)"', cui)
    print(f'  Ribbon回调: {ribbon_callbacks}')

if errors:
    print(f'\n❌ 验证门禁失败 ({len(errors)}项):')
    for e in errors:
        print(f'  - {e}')
    sys.exit(1)
else:
    print(f'\n✅ 验证门禁通过')
"@

[System.IO.File]::WriteAllText($validationScript, $validationCode, [System.Text.Encoding]::UTF8)

try {
    $validationOutput = python $validationScript 2>&1
    Write-Host $validationOutput
    if ($LASTEXITCODE -ne 0) {
        Write-Error "验证门禁未通过，请修复后重新构建"
        exit 1
    }
    Write-Success "验证门禁通过"
} catch {
    Write-Error "验证脚本执行失败: $_"
    exit 1
}

# ============================================================
# 完成
# ============================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " 构建完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "项目目录: $projectDir"
Write-Host "加载项: $excelFilePath"
Write-Host ""
Write-Host "下一步："
Write-Host "1. 运行 xlwings addin install --file $excelFilePath 安装"
Write-Host "2. 运行验证门禁D (AddIns2注册验证)"
Write-Host "3. 运行验证门禁E (交付前最终验证)"
Write-Host "========================================"