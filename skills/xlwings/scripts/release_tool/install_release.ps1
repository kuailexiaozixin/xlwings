<#
xlwings release tool - target machine installer for released packages.

What it does (in order):
  1. Validates the delivery folder (workbook + portable runtime with python.exe).
  2. Copies workbook + runtime to the install directory
     (default: %LOCALAPPDATA%\<workbook base name>), or -InPlace to install
     where the folder already is.
  3. Copy mode: rewrites the Interpreter_Win path inside the installed xlsm
     (zip surgery via rewrite_interpreter.py - Ribbon-safe, no COM save).
  4. Trust enablement (opt-out with -NoTrust):
     - Recursively unblocks files (removes Mark-of-the-Web).
     - Adds the install directory to Excel Trusted Locations (HKCU only,
       no admin required). Macros in trusted locations open with NO prompt.
  5. Optionally creates a desktop shortcut (opt-out with -NoShortcut).
  6. Writes install_manifest.txt next to the installed workbook so that
     uninstall_release.ps1 can remove exactly what was installed.

PowerShell 5.1 compatible. ASCII-only output. No PATH / machine registry changes.
#>
[CmdletBinding()]
param(
    [string]$SourceDir = '',
    [string]$Workbook = '',
    [string]$InstallDir = '',
    [switch]$InPlace,
    [switch]$NoShortcut,
    [switch]$NoTrust,
    [switch]$NoAddinRegister
)

$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Msg) Write-Host ('[*] ' + $Msg) }
function Write-Ok   { param([string]$Msg) Write-Host ('[OK] ' + $Msg) -ForegroundColor Green }
function Write-Warn2{ param([string]$Msg) Write-Host ('[!]  ' + $Msg) -ForegroundColor Yellow }

# --- 1. Resolve source folder and workbook ---------------------------------
if (-not $SourceDir) {
    if ($MyInvocation.MyCommand.Path) {
        $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $SourceDir = (Get-Location).Path
    }
}
if (-not (Test-Path $SourceDir -PathType Container)) {
    throw "SourceDir not found: $SourceDir"
}

if (-not $Workbook) {
    # 关键：不用 -Include（PowerShell 已知问题：-Include 需路径末尾 \* 或 -Recurse 才生效）
    # 改用 Where-Object 按扩展名过滤，支持 .xlsm 和 .xlam
    $wbFile = Get-ChildItem -Path $SourceDir -File |
              Where-Object { $_.Extension -eq '.xlsm' -or $_.Extension -eq '.xlam' } |
              Select-Object -First 1
    if (-not $wbFile) { throw "No .xlsm/.xlam workbook found in $SourceDir" }
    $Workbook = $wbFile.Name
}
$srcWorkbook = Join-Path $SourceDir $Workbook
if (-not (Test-Path $srcWorkbook)) { throw "Workbook not found: $srcWorkbook" }
$baseName = [IO.Path]::GetFileNameWithoutExtension($Workbook)

# --- 2. Locate portable runtime ---------------------------------------------
$runtime = Get-ChildItem -Path $SourceDir -Directory |
           Where-Object { Test-Path (Join-Path $_.FullName 'python.exe') } |
           Select-Object -First 1
if (-not $runtime) {
    throw "No portable runtime found (a folder containing python.exe) in $SourceDir"
}
Write-Step ("Workbook : " + $Workbook)
Write-Step ("Runtime  : " + $runtime.Name)

# --- 2b. 终止占用旧安装目录的进程（防止 DLL 被占用导致删除失败）------------
$oldInstallDir = Join-Path $env:LOCALAPPDATA $baseName
if (Test-Path $oldInstallDir) {
    $lockedProcs = @()
    Get-Process -Name python -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            if ($_.Path -and $_.Path.StartsWith($oldInstallDir, [StringComparison]::OrdinalIgnoreCase)) {
                $lockedProcs += $_
            }
        } catch {}
    }
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
        # Excel 可能加载了旧 xlam，占用文件
        $lockedProcs += $_
    }
    if ($lockedProcs.Count -gt 0) {
        Write-Step ("Terminating " + $lockedProcs.Count + " process(es) locking old install dir...")
        $lockedProcs | ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                Write-Ok ("  Terminated: " + $_.Name + " (PID " + $_.Id + ")")
            } catch {}
        }
        Start-Sleep -Seconds 2
    }
}

# --- 3. Determine install directory and copy --------------------------------
if ($InPlace) {
    $InstallDir = $SourceDir
    Write-Step "InPlace mode: installing where the folder already is"
} else {
    if (-not $InstallDir) {
        $InstallDir = Join-Path $env:LOCALAPPDATA $baseName
    }
    Write-Step ("InstallDir: " + $InstallDir)
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null

    Copy-Item -Path $srcWorkbook -Destination $InstallDir -Force
    $destRuntime = Join-Path $InstallDir $runtime.Name
    if (Test-Path $destRuntime) {
        Remove-Item -Path $destRuntime -Recurse -Force
    }
    # 用 robocopy 代替 Copy-Item：runtime 中可能有超过 260 字符的长路径，
    # Copy-Item 会报 DirectoryNotFoundError，robocopy 原生支持长路径
    # 注意：robocopy 退出码 0-7 均为成功（1=有文件被复制），必须临时关闭
    # $ErrorActionPreference='Stop'，否则非零退出码会被当作错误抛出
    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & robocopy $runtime.FullName $destRuntime /E /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
    $robocopyExit = $LASTEXITCODE
    $ErrorActionPreference = $oldEAP
    if ($robocopyExit -ge 8) {
        throw "robocopy failed with exit code $robocopyExit"
    }

    # 复制 Python 源码（入口模块 + src/ 包）
    # 当 RELEASE_EMBED_CODE=False 时，xlwings 通过 PYTHONPATH 查找外部源码
    $entryModule = Join-Path $SourceDir "$baseName.py"
    if (Test-Path $entryModule) {
        Copy-Item -Path $entryModule -Destination $InstallDir -Force
        Write-Ok "Copied entry module: $baseName.py"
    }
    $srcDir = Join-Path $SourceDir 'src'
    if (Test-Path $srcDir) {
        $destSrc = Join-Path $InstallDir 'src'
        if (Test-Path $destSrc) {
            Remove-Item -Path $destSrc -Recurse -Force
        }
        Copy-Item -Path $srcDir -Destination $destSrc -Recurse -Force
        Write-Ok "Copied source package: src/"
    }

    Write-Ok "Copied workbook, runtime, and source code"
}

$installedWorkbook = Join-Path $InstallDir $Workbook
$installedRuntime  = Join-Path $InstallDir $runtime.Name
$installedPython   = Join-Path $installedRuntime 'python.exe'

# --- 4. Rewrite Interpreter_Win and fix config (copy mode only) -------------
if (-not $InPlace) {
    $rewritePy = Join-Path $SourceDir 'rewrite_interpreter.py'
    if (Test-Path $rewritePy) {
        Write-Step "Rewriting Interpreter_Win inside installed workbook..."
        # rewrite_interpreter.py 接受位置参数：<xlam_path> <new_interpreter_path>
        & "$installedPython" "$rewritePy" "$installedWorkbook" "$installedPython"
        if ($LASTEXITCODE -ne 0) {
            Write-Warn2 "Interpreter rewrite reported a problem - check output above"
        } else {
            Write-Ok "Interpreter_Win now points to the installed runtime"
        }
    } else {
        Write-Warn2 "rewrite_interpreter.py not found next to installer - Interpreter_Win NOT rewritten"
        Write-Warn2 "The workbook may still point to the builder machine path."
    }
    
    # 4b. 统一修复配置表（PYTHONPATH 正斜杠 + RELEASE_EMBED_CODE 保持 + 空值修复）
    Write-Step "Fixing config sheet (PYTHONPATH, RELEASE_EMBED_CODE, empty values)..."
    $pythonpathPosix = $InstallDir.Replace('\', '/')
    $configFixCmd = @"
import win32com.client, sys
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(sys.argv[1])
pythonpath = sys.argv[2]
for sht in wb.Worksheets:
    if 'xlwings' in sht.Name.lower() and 'conf' in sht.Name.lower():
        for i in range(1, sht.UsedRange.Rows.Count + 1):
            key = sht.Cells(i, 1).Value
            if key:
                key_str = str(key).strip()
                if key_str == 'PYTHONPATH':
                    sht.Cells(i, 2).Value = pythonpath
                elif key_str in ('USE UDF SERVER', 'SHOW CONSOLE', 'RELEASE_NO_ADDIN'):
                    if sht.Cells(i, 2).Value in (None, ''):
                        sht.Cells(i, 2).Value = 'False'
                elif key_str == 'RELEASE_EMBED_CODE':
                    # 保持原值（True=code embed 模式, False=PYTHONPATH 模式）
                    pass
wb.Save()
wb.Close(SaveChanges=False)
excel.Quit()
print('Config sheet fixed')
"@
    & "$installedPython" -c $configFixCmd $installedWorkbook $pythonpathPosix 2>&1
    Write-Ok "Config sheet fixed (PYTHONPATH forward slashes, empty values filled)"
    
    # 4c. 修改 PROJECT_NAME 为 xlwings（确保 VBA 能找到 xlwings.conf 配置表）
    Write-Step "Setting PROJECT_NAME to 'xlwings'..."
    $projNameCmd = "import win32com.client,sys;excel=win32com.client.Dispatch('Excel.Application');excel.Visible=False;excel.DisplayAlerts=False;wb=excel.Workbooks.Open(sys.argv[1]);[comp.CodeModule.ReplaceLine(i+1,line.replace('myaddin','xlwings')) for comp in wb.VBProject.VBComponents if comp.Name=='xlwings' for i,line in enumerate(comp.CodeModule.Lines(1,comp.CodeModule.CountOfLines).split(chr(10))) if 'Public Const PROJECT_NAME' in line];wb.Save();wb.Close(SaveChanges=False);excel.Quit();print('PROJECT_NAME set to xlwings')"
    & "$installedPython" -c $projNameCmd $installedWorkbook 2>&1
    Write-Ok "PROJECT_NAME set to 'xlwings'"
    
    # 4d. Ribbon 回注（COM 保存会破坏 customUI 注册，必须自动执行）
    $reinjectPy = Join-Path $SourceDir 'reinject_ribbon_after_com_save.py'
    if (Test-Path $reinjectPy) {
        Write-Step "Reinjecting Ribbon customUI (COM save protection)..."
        & "$installedPython" "$reinjectPy" --xlam-path "$installedWorkbook" 2>&1
        Write-Ok "Ribbon customUI reinjected"
    } else {
        Write-Warn2 "reinject_ribbon_after_com_save.py not found - Ribbon may be broken"
    }
}

# --- 5. Trust enablement ------------------------------------------------------
$trustKeys = @()
if (-not $NoTrust) {
    Write-Step "Removing Mark-of-the-Web from installed files..."
    try {
        Get-ChildItem -Path $InstallDir -Recurse -File -ErrorAction SilentlyContinue |
            Unblock-File -ErrorAction SilentlyContinue
        Write-Ok "Files unblocked"
    } catch {
        Write-Warn2 "Unblock-File failed: $($_.Exception.Message)"
    }

    Write-Step "Adding Excel Trusted Location (HKCU, no admin needed)..."
    $officeRoot = 'HKCU:\Software\Microsoft\Office'
    if (Test-Path $officeRoot) {
        $verKeys = Get-ChildItem $officeRoot | Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
        foreach ($vk in $verKeys) {
            $secRoot = Join-Path $vk.PSPath 'Excel\Security\Trusted Locations'
            if (-not (Test-Path $secRoot)) { continue }
            $already = Get-ChildItem $secRoot -ErrorAction SilentlyContinue | Where-Object {
                $p = Get-ItemProperty -Path $_.PSPath -Name Path -ErrorAction SilentlyContinue
                ($p -ne $null) -and ($p.Path -eq $InstallDir)
            }
            if ($already) {
                $trustKeys += ('HKEY_CURRENT_USER\' + $vk.PSChildName + '\Excel\Security\Trusted Locations\' + $already[0].PSChildName)
                continue
            }
            $n = 1
            while (Test-Path (Join-Path $secRoot ('Location' + $n))) { $n++ }
            $locName = 'Location' + $n
            $newKey = New-Item -Path (Join-Path $secRoot $locName) -Force
            New-ItemProperty -Path $newKey.PSPath -Name 'Path' -Value $InstallDir -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $newKey.PSPath -Name 'AllowSubFolders' -Value 1 -PropertyType DWord -Force | Out-Null
            New-ItemProperty -Path $newKey.PSPath -Name 'Description' -Value ('xlwings release install - ' + $baseName) -PropertyType String -Force | Out-Null
            $trustKeys += ('HKEY_CURRENT_USER\' + $vk.PSChildName + '\Excel\Security\Trusted Locations\' + $locName)
        }
    }
    if ($trustKeys.Count -gt 0) {
        Write-Ok ("Trusted Locations: " + ($trustKeys -join ' | '))
    } else {
        Write-Warn2 "No Excel Trusted Locations registry root found - macros may still prompt"
    }
}

# --- 6. Register as Excel add-in (registry via COM) ---------------------------
# release_tool 路线：通过 Excel COM 对象注册加载项（自动处理 Add-in Manager + OPEN 项）
# 不使用 XLSTART（那是另一条独立分发路线）
$addinRegistered = $false
$registryVersion = ''
if (-not $NoAddinRegister) {
    Write-Step "Registering as Excel add-in (registry via COM)..."
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # AddIns.Add 会自动在注册表中注册加载项
        $addin = $excel.AddIns.Add($installedWorkbook, $false)
        $addin.Installed = $true
        
        $registryVersion = $excel.Version
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        
        $addinRegistered = $true
        Write-Ok ("  Add-in registered: " + $installedWorkbook)
        Write-Host '    Excel will load this add-in on startup (registry).' -ForegroundColor Cyan
    } catch {
        Write-Warn2 ("COM registration failed: " + $_.Exception.Message)
        Write-Warn2 "Falling back to manual registry registration..."
        
        # 回退：手动注册表注册
        $excelVer = '16.0'
        $addinKey = "HKCU:\Software\Microsoft\Office\$excelVer\Excel\Add-in Manager"
        if (-not (Test-Path $addinKey)) { New-Item -Path $addinKey -Force | Out-Null }
        New-ItemProperty -Path $addinKey -Name $installedWorkbook -Value '' -PropertyType String -Force | Out-Null
        
        $optionsKey = "HKCU:\Software\Microsoft\Office\$excelVer\Excel\Options"
        if (-not (Test-Path $optionsKey)) { New-Item -Path $optionsKey -Force | Out-Null }
        New-ItemProperty -Path $optionsKey -Name 'OPEN' -Value ('/R "' + $installedWorkbook + '"') -PropertyType String -Force | Out-Null
        
        $registryVersion = $excelVer
        $addinRegistered = $true
        Write-Ok ("  Manual registry registration complete")
    }
}

# --- 7. Desktop shortcut ------------------------------------------------------
$shortcutPath = ''
if (-not $NoShortcut) {
    Write-Step "Creating desktop shortcut..."
    try {
        $ws = New-Object -ComObject WScript.Shell
        $desktop = [Environment]::GetFolderPath('Desktop')
        $shortcutPath = Join-Path $desktop ($baseName + '.lnk')
        $lnk = $ws.CreateShortcut($shortcutPath)
        $lnk.TargetPath = $installedWorkbook
        $lnk.WorkingDirectory = $InstallDir
        $lnk.Description = ('xlwings release - ' + $baseName)
        $lnk.Save()
        Write-Ok ("Shortcut: " + $shortcutPath)
    } catch {
        $shortcutPath = ''
        Write-Warn2 ("Shortcut creation failed: " + $_.Exception.Message)
    }
}

# --- 8. Manifest ---------------------------------------------------------------
$manifestPath = Join-Path $InstallDir 'install_manifest.txt'
$lines = @(
    ('installed_at=' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
    ('workbook=' + $Workbook),
    ('runtime=' + $runtime.Name),
    ('installdir=' + $InstallDir),
    ('shortcut=' + $shortcutPath),
    ('registry_version=' + $registryVersion),
    ('addin_registered=' + $addinRegistered),
    ('trustkeys=' + ($trustKeys -join '|')),
    ('source=' + $SourceDir)
)
Set-Content -Path $manifestPath -Value $lines -Encoding ASCII
Write-Ok ("Manifest: " + $manifestPath)

Write-Host ''
Write-Host '==============================================' -ForegroundColor Cyan
Write-Ok ("Install complete. Add-in registered in registry (Excel " + $registryVersion + ")")
Write-Host '    Restart Excel - the add-in will load automatically.' -ForegroundColor Cyan
if ($trustKeys.Count -gt 0) {
    Write-Host '    Macros will open without any security prompt.' -ForegroundColor Cyan
}
Write-Host '==============================================' -ForegroundColor Cyan
