$ErrorActionPreference = 'Stop'
$officeVersions = @('16.0', '15.0', '14.0', '12.0')
$fallbackVersion = '16.0'
$installTarget = 'excel'
$targetHostName = 'Microsoft Excel'
$addinPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'myaddin.xlam'
$addinName = [System.IO.Path]::GetFileName($addinPath)

function Get-ExcelVersions {
    $versions = New-Object System.Collections.Generic.List[string]
    foreach ($version in $officeVersions) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        if (Test-Path -LiteralPath $optionsPath) {
            [void]$versions.Add($version)
        }
    }
    if ($versions.Count -eq 0) {
        [void]$versions.Add($fallbackVersion)
    }
    return $versions.ToArray()
}

function Ensure-ExcelOpenEntry {
    param([string]$OptionsPath)
    $current = Get-ItemProperty -LiteralPath $OptionsPath -ErrorAction SilentlyContinue
    $openNames = @()
    if ($current) {
        $openNames = @(
            $current.PSObject.Properties |
            Where-Object { $_.Name -match '^OPEN\d*$' } |
            Sort-Object {
                if ($_.Name -eq 'OPEN') { 0 } else { [int]$_.Name.Substring(4) }
            }
        )
    }
    foreach ($property in $openNames) {
        $valueText = [string]$property.Value
        if ($valueText -like "*$addinPath*" -or $valueText -like "*$addinName*") {
            return
        }
    }
    $slotName = 'OPEN'
    $index = 0
    while (Get-ItemProperty -LiteralPath $OptionsPath -Name $slotName -ErrorAction SilentlyContinue) {
        $index++
        $slotName = 'OPEN' + $index
    }
    New-ItemProperty -LiteralPath $OptionsPath -Name $slotName -Value ('/R "' + $addinPath + '"') -PropertyType String -Force | Out-Null
}

function Register-ExcelAddin {
    foreach ($version in (Get-ExcelVersions)) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        $managerPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Add-in Manager"
        if (-not (Test-Path -LiteralPath $optionsPath)) { New-Item -Path $optionsPath -Force | Out-Null }
        if (-not (Test-Path -LiteralPath $managerPath)) { New-Item -Path $managerPath -Force | Out-Null }
        New-ItemProperty -LiteralPath $managerPath -Name $addinPath -Value '' -PropertyType String -Force | Out-Null
        Ensure-ExcelOpenEntry -OptionsPath $optionsPath
    }
}

function Get-WpsRegistryViews {
    $views = @([Microsoft.Win32.RegistryView]::Default)
    if ([Environment]::Is64BitOperatingSystem) {
        $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    }
    return ($views | Select-Object -Unique)
}

function Set-WpsRegistration {
    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $loadMacrosKey = $baseKey.CreateSubKey('Software\kingsoft\office\6.0\et\LoadMacros')
            try {
                $loadMacrosKey.SetValue($addinPath, '1', [Microsoft.Win32.RegistryValueKind]::String)
            }
            finally {
                $loadMacrosKey.Close()
            }
        }
        finally {
            $baseKey.Close()
        }
    }
}

function Enable-WpsProjectTrust {
    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $settingsKey = $baseKey.CreateSubKey('Software\kingsoft\office\6.0\et\Application Settings')
            try {
                $settingsKey.SetValue('KDEVBProjectTrust', '1', [Microsoft.Win32.RegistryValueKind]::String)
                $settingsKey.SetValue('JSIDEProjectTrust', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
            }
            finally {
                $settingsKey.Close()
            }
        }
        finally {
            $baseKey.Close()
        }
    }
}

if (-not (Test-Path -LiteralPath $addinPath -PathType Leaf)) {
    throw "未找到加载项文件：$addinPath"
}

switch ($installTarget) {
    'excel' {
        Register-ExcelAddin
        Write-Host "已写入 $targetHostName 加载项注册：$addinPath"
    }
    'wps' {
        Set-WpsRegistration
        Enable-WpsProjectTrust
        Write-Host "已写入 $targetHostName 加载项注册：$addinPath"
    }
    default {
        throw "不支持的安装目标：$installTarget"
    }
}
