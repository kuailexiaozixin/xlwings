# ensure_uv_env.ps1 — 检查/安装 uv、配置镜像、安装 Python
# 任何新项目的第一步：检查/安装 uv、配置镜像、安装 Python
# 可选：-ExcelChain 验证解释器在 Excel 加载项链中能否导入 xlwings（venv 链不激活拦截）
# UTF-8 with BOM + CRLF

param(
    [string]$PythonVersion = "3.11",
    [switch]$Quiet,
    [switch]$ExcelChain,
    [string]$VerifyInterpreter = ""
)

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message)
    if (-not $Quiet) { Write-Host "[INFO] $Message" }
}

# 1. 检查 uv 是否已安装
$uvPath = Get-Command "uv" -ErrorAction SilentlyContinue
if (-not $uvPath) {
    Write-Log "正在安装 uv..."
    # 使用官方安装脚本
    $installScript = Invoke-RestMethod -Uri "https://astral.sh/uv/install.ps1"
    Invoke-Expression $installScript
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","User") + ";" + [System.Environment]::GetEnvironmentVariable("Path","Machine")
    Write-Log "uv 安装完成"
} else {
    Write-Log "uv 已安装：$($uvPath.Source)"
}

# 2. 设置中国镜像（uv 0.12+ 已移除 `uv config` 子命令，探测可用性后降级为环境变量）
Write-Log "配置镜像源..."
$uvConfigOk = $false
try { & uv config --help *> $null; if ($LASTEXITCODE -eq 0) { $uvConfigOk = $true } } catch { $uvConfigOk = $false }
if ($uvConfigOk) {
    & uv config set index-url "https://mirrors.aliyun.com/pypi/simple/" 2>$null
} else {
    $env:UV_DEFAULT_INDEX = "https://mirrors.aliyun.com/pypi/simple/"
    Write-Log "uv 0.12+ 无 config 子命令，改用环境变量 UV_DEFAULT_INDEX（本次会话生效）"
}

# 3. 安装/检查 Python
Write-Log "检查 Python $PythonVersion..."
$pythonCheck = & uv python list 2>$null | Select-String "$PythonVersion"
if (-not $pythonCheck) {
    Write-Log "正在安装 Python $PythonVersion..."
    & uv python install $PythonVersion
    Write-Log "Python $PythonVersion 安装完成"
} else {
    Write-Log "Python $PythonVersion 已就绪"
}

# 4. 可选：Excel 加载项链解释器身份验证
# 背景（2026-09-17 实测）：uv venv 在 Excel/cmd 链中可能不激活——显式调用 venv python.exe，
# 但环境存在 PYTHONHOME 指向 base 时，PREFIX 回落到 base、venv site-packages 不进 sys.path，
# 加载项/按钮报 'No module named xlwings'。本验证在真实 Excel 报错之前拦截。
if ($ExcelChain -or $VerifyInterpreter) {
    Write-Log "Excel 链解释器身份验证..."
    $target = $VerifyInterpreter
    if (-not $target) {
        # 未显式指定时，尝试定位 uv 管理的 python（优先 PythonVersion）
        $uvPy = & uv python find $PythonVersion 2>$null
        if ($LASTEXITCODE -eq 0 -and $uvPy) { $target = $uvPy }
        else {
            $cmdPy = Get-Command "python" -ErrorAction SilentlyContinue
            if ($cmdPy) { $target = $cmdPy.Source }
        }
    }
    if (-not $target -or -not (Test-Path $target)) {
        Write-Host "[WARN] 未找到待验证的解释器（用 -VerifyInterpreter <python.exe 或 .bat> 显式指定）" -ForegroundColor Yellow
    } else {
        $probe = @'
import sys, site
print('EXE=' + sys.executable)
print('PREFIX=' + sys.prefix)
print('BASE=' + sys.base_prefix)
print('SITE=' + ','.join(site.getsitepackages()))
try:
    import xlwings
    print('XLWINGS=' + xlwings.__version__)
except Exception as e:
    print('XLWINGS_ERR=' + repr(e))
'@
        $tmp = Join-Path $env:TEMP "xlwings_chain_probe_$PID.py"
        [System.IO.File]::WriteAllText($tmp, $probe, (New-Object System.Text.UTF8Encoding($false)))
        try {
            if ($target -like '*.exe') {
                & $target $tmp
            } else {
                # .bat 包装解释器：经 cmd 执行（bat 不能直接调用）
                cmd.exe /d /s /c "call `"$target`" `"$tmp`""
            }
            if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) {
                Write-Host "[WARN] 解释器链验证命令返回码 $LASTEXITCODE" -ForegroundColor Yellow
            }
        } finally {
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    Write-Log "Excel 链验证完成。判据：PREFIX==BASE 且 XLWINGS 缺失 → venv 未生效，需用 bat 包装解释器（set PYTHONHOME=<base> + set PYTHONPATH=<模块目录> + 调 base 解释器），详见 references/04-python-guidance.md 第五节。"
}

Write-Log "环境就绪。"
