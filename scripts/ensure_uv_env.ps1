# ensure_uv_env.ps1 — 检查/安装 uv、配置镜像、安装 Python
# 任何新项目的第一步：检查/安装 uv、配置镜像、安装 Python
# UTF-8 with BOM + CRLF

param(
    [string]$PythonVersion = "3.11",
    [switch]$Quiet
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

# 2. 设置中国镜像
Write-Log "配置镜像源..."
& uv config set index-url "https://mirrors.aliyun.com/pypi/simple/" 2>$null

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

Write-Log "环境就绪。"
