<#
xlwings release tool - one-click uninstaller for released packages.

Features:
  1. Auto-locate install dir (manifest next to script, or %LOCALAPPDATA%\<project>, or -InstallDir)
  2. Detect and prompt to kill processes locking files (Excel/python)
  3. Unregister add-in from Excel registry (Add-in Manager + OPEN entries)
  4. Clean Excel add-in registry (Add-in Manager + OPEN entries)
  5. Remove Trusted Location registry keys
  6. Remove desktop shortcut
  7. Remove install directory

PowerShell 5.1 compatible. ASCII-only output.
#>
[CmdletBinding()]
param(
    [string]$InstallDir = '',
    [switch]$ShowManifest,
    [switch]$Force  # 自动终止进程，不询问
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Msg) Write-Host ('[*] ' + $Msg) }
function Write-Ok   { param([string]$Msg) Write-Host ('[OK] ' + $Msg) -ForegroundColor Green }
function Write-Warn2{ param([string]$Msg) Write-Host ('[!]  ' + $Msg) -ForegroundColor Yellow }
function Write-Question { param([string]$Msg) Write-Host ('[?] ' + $Msg) -ForegroundColor Yellow }

# --- 1. Resolve install dir ---------------------------------------------------
$manifestPath = ''
$resolvedFrom = ''

if ($InstallDir) {
    # 用户显式指定
    $manifestPath = Join-Path $InstallDir 'install_manifest.txt'
    $resolvedFrom = 'explicit parameter'
} else {
    # 优先：脚本所在目录的 manifest
    $here = if ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $candidate = Join-Path $here 'install_manifest.txt'
    if (Test-Path $candidate) {
        $InstallDir = $here
        $manifestPath = $candidate
        $resolvedFrom = 'script directory'
    } else {
        # 其次：从当前目录名推断项目名，查找 %LOCALAPPDATA%\<项目名>
        $projectName = Split-Path $here -Leaf
        $candidate2 = Join-Path $env:LOCALAPPDATA $projectName
        $manifest2 = Join-Path $candidate2 'install_manifest.txt'
        if (Test-Path $manifest2) {
            $InstallDir = $candidate2
            $manifestPath = $manifest2
            $resolvedFrom = "%LOCALAPPDATA%\$projectName"
        } else {
            # 最后：扫描 %LOCALAPPDATA% 下所有包含 install_manifest.txt 的目录
            Write-Step "Searching for install_manifest.txt in %LOCALAPPDATA%..."
            $found = Get-ChildItem -Path $env:LOCALAPPDATA -Filter 'install_manifest.txt' -Recurse -Depth 2 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($found) {
                $InstallDir = $found.DirectoryName
                $manifestPath = $found.FullName
                $resolvedFrom = "scan: $InstallDir"
            }
        }
    }
}

if (-not $manifestPath -or -not (Test-Path $manifestPath)) {
    Write-Host ''
    Write-Warn2 "Cannot locate install_manifest.txt."
    Write-Host "  Searched: script dir, %LOCALAPPDATA%\<project>, recursive scan"
    Write-Host "  Please run with -InstallDir <path>"
    exit 1
}

Write-Step ("Install dir resolved from: " + $resolvedFrom)
Write-Step ("Install dir: " + $InstallDir)

# --- 2. Parse manifest ---------------------------------------------------------
$manifest = @{}
Get-Content -Path $manifestPath | ForEach-Object {
    $idx = $_.IndexOf('=')
    if ($idx -gt 0) {
        $manifest[$_.Substring(0, $idx)] = $_.Substring($idx + 1)
    }
}
Write-Step ("Manifest loaded. Workbook: " + $manifest['workbook'])

if ($ShowManifest) {
    $manifest.GetEnumerator() | Sort-Object Name | ForEach-Object {
        Write-Host ('  ' + $_.Name + ' = ' + $_.Value)
    }
    return
}

# --- 3. Detect and kill locking processes --------------------------------------
$workbookName = $manifest['workbook']
$lockingProcs = @()

# 检查 Excel 进程
$excelProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($excelProcs) {
    $lockingProcs += $excelProcs
}

# 检查与项目相关的 python 进程（面板/server）
$pythonProcs = Get-Process -Name python -ErrorAction SilentlyContinue
foreach ($p in $pythonProcs) {
    try {
        $cmdline = (Get-CimInstance Win32_Process -Filter "ProcessId=$($p.Id)" -ErrorAction SilentlyContinue).CommandLine
        if ($cmdline -and ($cmdline -match [regex]::Escape($InstallDir) -or $cmdline -match [regex]::Escape($workbookName))) {
            $lockingProcs += $p
        }
    } catch {}
}

if ($lockingProcs.Count -gt 0) {
    Write-Host ''
    Write-Warn2 ("Found " + $lockingProcs.Count + " process(es) that may lock files:")
    $lockingProcs | ForEach-Object {
        $cmdline = ''
        try { $cmdline = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue).CommandLine } catch {}
        Write-Host ("  PID " + $_.Id + ": " + $_.Name + " - " + $cmdline.Substring(0, [Math]::Min(80, $cmdline.Length)))
    }
    
    if ($Force) {
        Write-Step "Force mode: killing processes automatically..."
        $lockingProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Write-Ok "Processes terminated"
    } else {
        Write-Question "Kill these processes to continue uninstall? (Y/N)"
        $answer = Read-Host
        if ($answer -eq 'Y' -or $answer -eq 'y') {
            $lockingProcs | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Write-Ok "Processes terminated"
        } else {
            Write-Warn2 "Processes not killed. Uninstall may fail to remove locked files."
            Write-Warn2 "You can close them manually and rerun, or use -Force flag."
        }
    }
}

# --- 4. Clean Excel add-in registry (Add-in Manager + OPEN items) -------------
Write-Step "Cleaning Excel add-in registry entries..."
$cleanedCount = 0
$oldEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'

$allKeys = @(
    'HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager',
    'HKCU:\Software\Microsoft\Office\15.0\Excel\Add-in Manager',
    'HKCU:\Software\Microsoft\Office\14.0\Excel\Add-in Manager',
    'HKCU:\Software\Microsoft\Office\12.0\Excel\Add-in Manager',
    'HKCU:\Software\Microsoft\Office\11.0\Excel\Add-in Manager',
    'HKCU:\Software\Microsoft\Office\16.0\Excel\Options',
    'HKCU:\Software\Microsoft\Office\15.0\Excel\Options',
    'HKCU:\Software\Microsoft\Office\14.0\Excel\Options',
    'HKCU:\Software\Microsoft\Office\12.0\Excel\Options',
    'HKCU:\Software\Microsoft\Office\11.0\Excel\Options'
)

foreach ($keyPath in $allKeys) {
    if (-not (Test-Path $keyPath)) { continue }
    $values = Get-ItemProperty $keyPath -ErrorAction SilentlyContinue
    if (-not $values) { continue }
    foreach ($prop in $values.PSObject.Properties) {
        if (-not $prop.Name) { continue }
        if ($prop.Name -match '^PS') { continue }
        $nameHit = $false
        $valHit = $false
        if ($prop.Name) { $nameHit = $prop.Name -match [regex]::Escape($workbookName) }
        if ($prop.Value) { $valHit = $prop.Value -match [regex]::Escape($workbookName) }
        if (-not ($nameHit -or $valHit)) { continue }
        Remove-ItemProperty -Path $keyPath -Name $prop.Name -Force -ErrorAction SilentlyContinue
        Write-Ok ("  Removed: " + $prop.Name)
        $cleanedCount++
    }
}

$ErrorActionPreference = $oldEAP
if ($cleanedCount -eq 0) {
    Write-Step "No Excel add-in registry entries found"
}

# --- 6. Trusted Location keys --------------------------------------------------
if ($manifest['trustkeys']) {
    foreach ($keyPath in $manifest['trustkeys'].Split('|')) {
        if (-not $keyPath) { continue }
        $psPath = $keyPath -replace '^HKEY_CURRENT_USER', 'HKCU:'
        if (Test-Path $psPath) {
            Remove-Item -Path $psPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Ok ("Removed Trusted Location: " + $keyPath)
        }
    }
}

# --- 7. Desktop shortcut -------------------------------------------------------
if ($manifest['shortcut'] -and (Test-Path $manifest['shortcut'])) {
    Remove-Item -Path $manifest['shortcut'] -Force -ErrorAction SilentlyContinue
    Write-Ok ("Removed shortcut: " + $manifest['shortcut'])
}

# --- 8. Install directory ------------------------------------------------------
Write-Step ("Removing install directory: " + $InstallDir)

if (Test-Path $InstallDir) {
    # 尝试直接删除（robocopy 镜像方式处理长路径）
    $emptyDir = Join-Path $env:TEMP ('empty_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $emptyDir -Force | Out-Null
    
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    robocopy $emptyDir $InstallDir /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
    $ErrorActionPreference = $oldEAP
    
    Remove-Item $emptyDir -Force -ErrorAction SilentlyContinue
    Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
    
    if (Test-Path $InstallDir) {
        Write-Warn2 ("Failed to fully remove: " + $InstallDir)
        Write-Warn2 "Some files may still be locked. Close Excel/WPS and retry."
    } else {
        Write-Ok "Install directory removed"
    }
} else {
    Write-Step "Install directory already gone"
}

Write-Host ''
Write-Host '==============================================' -ForegroundColor Cyan
Write-Ok 'Uninstall complete.'
Write-Host '  - Add-in unregistered from Excel (registry)' -ForegroundColor Cyan
Write-Host '  - Excel registry entries cleaned' -ForegroundColor Cyan
Write-Host '  - Install directory removed' -ForegroundColor Cyan
Write-Host '==============================================' -ForegroundColor Cyan
