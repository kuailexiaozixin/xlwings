#requires -Version 5.1

# AddIns2 注册验证脚本（收敛原 SKILL.md 中门禁 C / 门禁 D 的三段重复代码）
# 用途：按分发方式验证 .xlam 在 Excel 中的可用性。
# 两种分发模式（-DistMode，默认 Auto 自动识别）：
#   XLSTART  ：加载项文件放在 XLSTART 目录，Excel 启动时无条件自动加载。
#              真机铁律（references/09-testing-debugging-guidance.md）：COM 新起实例不加载 XLSTART，其 AddIns2 只枚举注册表注册过的
#              加载项，XLSTART 文件不出现在其中——用 AddIns2 验证 XLSTART 分发属于设计性误报；
#              且此分发方式无需任何注册表注册（Add-in Manager / OPEN）。
#              本模式验证：文件存在（真实加载由门禁 L 真实启动验证）。
#   Registry ：加载项通过注册表（Add-in Manager / OPEN）注册加载。
#              本模式验证：COM 新起实例的 AddIns2 列出且 Installed=True（原行为）。
# 用法：
#   powershell -ExecutionPolicy Bypass -File scripts\verify_addin_registered.ps1 `
#     -AddinName "yourproject" [-AddinPath "C:\path\to\yourproject.xlam"] [-DistMode Auto|XLSTART|Registry]
# 说明：
#   - AddinName 传项目名关键词（大小写不敏感，模糊匹配）
#   - AddinPath 可选；传入时同时校验该路径文件存在；Auto 模式下路径含 \XLSTART\ 即判为 XLSTART 模式
#   - 公共工具（进程追踪）来自 ./lib/excel-com-common.ps1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AddinName,

    [AllowEmptyString()]
    [string]$AddinPath = "",

    [ValidateSet('Auto', 'XLSTART', 'Registry')]
    [string]$DistMode = 'Auto'
)

$ErrorActionPreference = 'Stop'

# --- 加载公共工具库 ---
$commonLibPath = Join-Path $PSScriptRoot 'lib\excel-com-common.ps1'
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

# --- 分发模式判定 ---
$mode = $DistMode
if ($mode -eq 'Auto') {
    if (-not [string]::IsNullOrWhiteSpace($AddinPath) -and $AddinPath -match '\\XLSTART\\') {
        $mode = 'XLSTART'
    }
    else {
        $mode = 'Registry'
    }
}
Write-Host "分发模式：$mode"

if (-not [string]::IsNullOrWhiteSpace($AddinPath)) {
    if (-not (Test-Path -LiteralPath $AddinPath -PathType Leaf)) {
        throw "加载项文件不存在：$AddinPath"
    }
    Write-Host "加载项文件存在：$AddinPath"
}

# --- XLSTART 模式：不走 COM/AddIns2 验证 ---
# 真机铁律（references/09-testing-debugging-guidance.md）：COM 新起实例不加载 XLSTART，其 AddIns2 只枚举注册表注册过的加载项，
# 因此 XLSTART 文件永远不出现在新 COM 实例的 AddIns2 中——用 AddIns2 验证 XLSTART 分发属于设计性误报。
# XLSTART 分发的真实加载验证由门禁 L（真实启动 Excel）承担；本模式只做文件级验证 + 信息性检查。
if ($mode -eq 'XLSTART') {
    if ([string]::IsNullOrWhiteSpace($AddinPath)) {
        throw "XLSTART 模式必须传入 -AddinPath（XLSTART 下的文件路径）"
    }
    $fileItem = Get-Item -LiteralPath $AddinPath
    Write-Host ("  文件大小：{0} 字节 | 修改时间：{1}" -f $fileItem.Length, $fileItem.LastWriteTime)

    # 信息性检查（不作判定依据）：若当前有可连接的 Excel 实例且已加载该加载项工作簿，顺带报告
    try {
        $runningXl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        try {
            $null = $runningXl.Workbooks.Item($fileItem.Name)
            Write-Host "  信息：当前运行中的 Excel 实例已加载 $($fileItem.Name)（真实加载确认）"
        }
        catch {
            Write-Host "  信息：当前可连接的 Excel 实例未加载该工作簿（COM 实例或尚未启动新会话，不作判定依据）"
        }
    }
    catch {
        Write-Host "  信息：当前无可连接的 Excel 实例（XLSTART 自动加载由门禁 L 在无实例时真机验证）"
    }

    Write-Host "✅ XLSTART 分发验证通过（文件存在；自动加载验证见门禁 L）"
    exit 0
}

# --- Registry 模式：COM 新起实例 + AddIns2 + Installed=True ---
$trackedExcel = $null
try {
    $trackedExcel = Start-TrackedExcelApplication
    $excel = $trackedExcel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $xl = $excel.Application
    $addins = $xl.AddIns2
    $count = $addins.Count

    Write-Host "AddIns2 集合条目数：$count"

    # 第一阶段：仅枚举收集（单个条目 COM 访问失败只跳过，不上抛）
    $matched = @()
    for ($i = 1; $i -le $count; $i++) {
        try {
            $addin = $addins.Item($i)
            $name = [string]$addin.Name
            if ($name.ToLowerInvariant() -like "*$($AddinName.ToLowerInvariant())*") {
                $matched += [pscustomobject]@{
                    Name      = $name
                    FullName  = [string]$addin.FullName
                    Installed = [bool]$addin.Installed
                }
            }
        }
        catch {
            # 该条目访问失败，跳过继续枚举
        }
    }

    # 第二阶段：业务校验（此处的失败必须真实上抛，不得被枚举容错吞掉）
    if ($matched.Count -eq 0) {
        throw "加载项未在 AddIns2 集合中列出（关键词：$AddinName）"
    }

    foreach ($m in $matched) {
        Write-Host "  找到加载项：$($m.Name)"
        Write-Host "  FullName：$($m.FullName)"
        Write-Host "  Installed：$($m.Installed)"
        if (-not $m.Installed) {
            throw "加载项 $($m.Name) 未启用（Installed=False）"
        }
    }

    Write-Host "✅ AddIns2 注册验证通过（分发模式：Registry）"
}
finally {
    if ($trackedExcel) {
        Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
    }
}
