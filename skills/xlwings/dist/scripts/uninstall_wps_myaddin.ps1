$ErrorActionPreference = 'Stop'
$officeVersions = @('16.0', '15.0', '14.0', '12.0')
$fallbackVersion = '16.0'
$installTarget = 'wps'
$targetHostName = 'WPS 表格'
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

function Remove-ExcelAddin {
    foreach ($version in (Get-ExcelVersions)) {
        $optionsPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Options"
        $managerPath = "HKCU:\Software\Microsoft\Office\$version\Excel\Add-in Manager"
        if (Test-Path -LiteralPath $optionsPath) {
            $current = Get-ItemProperty -LiteralPath $optionsPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })) {
                $valueText = [string]$property.Value
                if ($valueText -like "*$addinPath*" -or $valueText -like "*$addinName*") {
                    Remove-ItemProperty -LiteralPath $optionsPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }
        if (Test-Path -LiteralPath $managerPath) {
            $current = Get-ItemProperty -LiteralPath $managerPath
            foreach ($property in @($current.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' })) {
                if ($property.Name -eq $addinPath -or $property.Name -like "*\$addinName") {
                    Remove-ItemProperty -LiteralPath $managerPath -Name $property.Name -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

function Get-WpsRegistryViews {
    $views = @([Microsoft.Win32.RegistryView]::Default)
    if ([Environment]::Is64BitOperatingSystem) {
        $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    }
    return ($views | Select-Object -Unique)
}

function Remove-WpsRegistration {
    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $loadMacrosKey = $baseKey.CreateSubKey('Software\kingsoft\office\6.0\et\LoadMacros')
            try {
                $loadMacrosKey.DeleteValue($addinPath, $false)
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

function Disable-WpsProjectTrust {
    foreach ($view in (Get-WpsRegistryViews)) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            $settingsKey = $baseKey.CreateSubKey('Software\kingsoft\office\6.0\et\Application Settings')
            try {
                $settingsKey.SetValue('KDEVBProjectTrust', '0', [Microsoft.Win32.RegistryValueKind]::String)
                $settingsKey.SetValue('JSIDEProjectTrust', 0, [Microsoft.Win32.RegistryValueKind]::DWord)
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

switch ($installTarget) {
    'excel' {
        Remove-ExcelAddin
        Write-Host "已移除 $targetHostName 加载项注册：$addinPath"
    }
    'wps' {
        Remove-WpsRegistration
        Disable-WpsProjectTrust
        Write-Host "已移除 $targetHostName 加载项注册：$addinPath"
    }
    default {
        throw "不支持的卸载目标：$installTarget"
    }
}
