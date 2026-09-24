#requires -Version 5.1

# 说明：
# - 这是加载项构建脚本，位于 skill 的 scripts/ 目录，由 AI 直接调用。
# - 推荐从 skill 根目录用相对路径引用：scripts\build_addin.ps1
# - 如果直接运行本脚本，默认输出会落到脚本所在目录下的 dist 子目录。
# - 两种起点模式：
#   1. 消费 quickstart 工程（推荐，xlwings 标准路径）：-QuickstartProjectPath <工程目录>
#      脚本会打开 quickstart 生成的 .xlam/.xlsm（保留 xlwings 引擎占位与配置表），
#      注入 VBA 模块与 Ribbon 后另存为产物。
#   2. 从零空白工作簿构建：不传 -QuickstartProjectPath（旧 VBA 时代路径，仅作后备）。
# - 本项目参数区里的默认值只是示例，不是业务固定值。ProjectName 必须为合法 Python
#   identifier（xlwings quickstart 同样要求），因为它会作为 RunPython import 的模块名。
# - 公共工具（进程追踪/编码/AccessVBOM/WPS/注册表/Zip）统一来自 ./lib/excel-com-common.ps1，
#   禁止在本脚本内复制实现。

[CmdletBinding()]
param(
    [string]$ProjectName = "myaddin",

    [ValidateSet("xlam", "xlsm")]
    [string]$OutputKind = "xlam",

    [AllowEmptyString()]
    [string]$OutputPath = "",

    # quickstart 生成的工程目录（内含 <ProjectName>.xlam/.xlsm 与 <ProjectName>.py）
    [AllowEmptyString()]
    [string]$QuickstartProjectPath = "",

    # 项目源码目录（含 src/modules、src/forms、src/ribbon 等），可选
    [AllowEmptyString()]
    [string]$ProjectSourceRoot = "",

    [string]$ModuleName = "modMain",

    [string]$RibbonTabId = "mainTab",
    [string]$RibbonTabLabel = "My Addin",
    [string]$RibbonGroupId = "mainGroup",
    [string]$RibbonGroupLabel = "Actions",

    [string]$FirstButtonId = "btnOne",
    [string]$FirstButtonLabel = "Run",
    [string]$FirstCallbackName = "RunActionOne",
    [string]$FirstButtonImageMso = "HappyFace",
    [AllowEmptyString()]
    [string]$FirstIconPath = "",

    [string]$SecondButtonId = "btnTwo",
    [string]$SecondButtonLabel = "Run 2",
    [string]$SecondCallbackName = "RunActionTwo",
    [string]$SecondButtonImageMso = "TableInsertDialog",
    [AllowEmptyString()]
    [string]$SecondIconPath = "",

    # 版本号与构建时间戳注入（替代旧水印系统，写入 VBA 模块头部注释）
    [string]$ProductVersion = "1.0.0",

    [string]$ModuleTemplatePath = "..\templates\internal-bases\excel-addin-core\src\modules\default_module.vba.tmpl",

    [string]$RibbonTemplatePath = "..\templates\internal-bases\excel-addin-core\ribbon\customUI14.xml.tmpl",

    [string]$InstallPs1TemplatePath = "..\templates\internal-bases\excel-addin-core\install\install_addin.ps1.tmpl",

    [string]$UninstallPs1TemplatePath = "..\templates\internal-bases\excel-addin-core\install\uninstall_addin.ps1.tmpl",

    [string]$InstallBatTemplatePath = "..\templates\internal-bases\excel-addin-core\install\install_addin.bat.tmpl",

    [string]$UninstallBatTemplatePath = "..\templates\internal-bases\excel-addin-core\install\uninstall_addin.bat.tmpl",

    [switch]$InstallAfterBuild,

    # 引擎激活后自动回注 Ribbon（实测确认：COM wb.save() 会覆盖丢失 customUI Content_Types 注册，
    # 本开关在构建完成后再次执行 zip 级 Ribbon 注入作为兜底，确保最终产物 Ribbon 注册完整）
    [switch]$ReinjectRibbonAfterActivation
)

$ErrorActionPreference = "Stop"

# ============================================================
# 加载公共工具库
# ============================================================
$commonLibPath = Join-Path $PSScriptRoot "lib\excel-com-common.ps1"
if (-not (Test-Path -LiteralPath $commonLibPath)) {
    throw "公共工具库不存在：$commonLibPath"
}
. $commonLibPath

# ============================================================
# 基础辅助
# ============================================================

function Get-ScriptTimestampText {
    return (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
}

function Get-ScriptBaseDirectory {
    $candidates = New-Object System.Collections.Generic.List[string]

    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        [void]$candidates.Add($PSScriptRoot)
    }

    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        [void]$candidates.Add((Split-Path -Parent $PSCommandPath))
    }

    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return [System.IO.Path]::GetFullPath($candidate)
        }
    }

    throw "Unable to determine script base directory."
}

function Resolve-ScriptRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-ScriptBaseDirectory) $Path))
}

function Get-ResolvedOptionalPath {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    return Resolve-ScriptRelativePath -Path $Path
}

function Test-HasCustomButtonImage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$IconPath
    )

    return -not [string]::IsNullOrWhiteSpace($IconPath)
}

function Assert-FileExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$DisplayName not found: $Path"
    }
}

function Expand-TemplateContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplateContent,

        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    $result = $TemplateContent
    foreach ($key in $Tokens.Keys) {
        $result = $result.Replace(("__{0}__" -f $key), [string]$Tokens[$key])
    }
    return $result
}

function Get-RenderedTemplateContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,

        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    Assert-FileExists -Path $TemplatePath -DisplayName "Template file"
    $template = Read-TextFileAuto -Path $TemplatePath
    return Expand-TemplateContent -TemplateContent $template -Tokens $Tokens
}

function Get-RenderedTextAsset {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $content = Read-TextFileAuto -Path $Path
    return Expand-TemplateContent -TemplateContent $content -Tokens (Get-TemplateTokens -Config $Config)
}

# ============================================================
# 构建配置
# ============================================================

function New-BuildConfig {
    $effectiveOutputPath = $OutputPath
    if ([string]::IsNullOrWhiteSpace($effectiveOutputPath)) {
        $effectiveOutputPath = Join-Path (Split-Path -Parent (Get-ScriptBaseDirectory)) "dist\$ProjectName.xlam"
    }

    if ([System.IO.Path]::IsPathRooted($effectiveOutputPath)) {
        $resolvedOutputPath = [System.IO.Path]::GetFullPath($effectiveOutputPath)
    }
    else {
        $resolvedOutputPath = Resolve-ScriptRelativePath -Path $effectiveOutputPath
    }

    $resolvedOutputKind = $OutputKind
    $outputExtension = [System.IO.Path]::GetExtension($resolvedOutputPath).TrimStart(".").ToLowerInvariant()
    $effectiveBuildTimestampUtc = Get-ScriptTimestampText

    if ($outputExtension) {
        if ($outputExtension -notin @("xlam", "xlsm")) {
            throw "OutputPath extension must be .xlam or .xlsm."
        }

        if ($OutputKind -ne $outputExtension) {
            $resolvedOutputKind = $outputExtension
        }
    }

    $resolvedQuickstartProjectPath = Get-ResolvedOptionalPath -Path $QuickstartProjectPath
    $resolvedProjectSourceRoot = Get-ResolvedOptionalPath -Path $ProjectSourceRoot

    # 如果传了 quickstart 工程但没传源码目录，默认指向工程内的 src/
    if (-not [string]::IsNullOrWhiteSpace($resolvedQuickstartProjectPath) -and [string]::IsNullOrWhiteSpace($resolvedProjectSourceRoot)) {
        $candidateSrc = Join-Path $resolvedQuickstartProjectPath "src"
        if ([System.IO.Directory]::Exists($candidateSrc)) {
            $resolvedProjectSourceRoot = $candidateSrc
        }
    }

    return [pscustomobject]@{
        OutputKind         = $resolvedOutputKind
        OutputPath         = $resolvedOutputPath
        ProjectName        = $ProjectName
        ProductVersion     = $ProductVersion
        BuildTimestampUtc  = $effectiveBuildTimestampUtc
        ModuleName         = $ModuleName
        RibbonTabId        = $RibbonTabId
        RibbonTabLabel     = $RibbonTabLabel
        RibbonGroupId      = $RibbonGroupId
        RibbonGroupLabel   = $RibbonGroupLabel
        FirstButtonId      = $FirstButtonId
        FirstButtonLabel   = $FirstButtonLabel
        FirstCallbackName  = $FirstCallbackName
        FirstButtonImageMso = $FirstButtonImageMso
        FirstIconPath      = (Get-ResolvedOptionalPath -Path $FirstIconPath)
        SecondButtonId     = $SecondButtonId
        SecondButtonLabel  = $SecondButtonLabel
        SecondCallbackName = $SecondCallbackName
        SecondButtonImageMso = $SecondButtonImageMso
        SecondIconPath     = (Get-ResolvedOptionalPath -Path $SecondIconPath)
        QuickstartProjectPath = $resolvedQuickstartProjectPath
        ProjectSourceRoot  = $resolvedProjectSourceRoot
        ModuleTemplatePath = (Resolve-ScriptRelativePath -Path $ModuleTemplatePath)
        RibbonTemplatePath = (Resolve-ScriptRelativePath -Path $RibbonTemplatePath)
        InstallPs1TemplatePath = (Resolve-ScriptRelativePath -Path $InstallPs1TemplatePath)
        UninstallPs1TemplatePath = (Resolve-ScriptRelativePath -Path $UninstallPs1TemplatePath)
        InstallBatTemplatePath = (Resolve-ScriptRelativePath -Path $InstallBatTemplatePath)
        UninstallBatTemplatePath = (Resolve-ScriptRelativePath -Path $UninstallBatTemplatePath)
        HelperBaseName     = ([System.IO.Path]::GetFileNameWithoutExtension($resolvedOutputPath) -replace '[^\w\-]+', '_')
        OutputFileName     = [System.IO.Path]::GetFileName($resolvedOutputPath)
        VbaFileFormat      = if ($resolvedOutputKind -eq "xlam") { 55 } else { 52 }
    }
}

function Get-TemplateTokens {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    return @{
        PROJECT_NAME        = $Config.ProjectName
        PRODUCT_VERSION     = $Config.ProductVersion
        BUILD_TIMESTAMP_UTC = $Config.BuildTimestampUtc
        FIRST_BUTTON_LABEL  = $Config.FirstButtonLabel
        FIRST_CALLBACK_NAME = $Config.FirstCallbackName
        SECOND_BUTTON_LABEL = $Config.SecondButtonLabel
        SECOND_CALLBACK_NAME = $Config.SecondCallbackName
    }
}

function Get-VbaModuleCode {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    Assert-FileExists -Path $Config.ModuleTemplatePath -DisplayName "Module template"

    $template = Read-TextFileAuto -Path $Config.ModuleTemplatePath
    return Expand-TemplateContent -TemplateContent $template -Tokens (Get-TemplateTokens -Config $Config)
}

function Get-ComponentBaseName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $baseName = $baseName -replace '\.(form|code)$', ''
    return $baseName
}

# ============================================================
# UserForm 构建（JSON 规格 -> COM UserForm）
# ============================================================

function Get-FormsProgId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ControlType
    )

    switch ($ControlType.ToLowerInvariant()) {
        'label' { return 'Forms.Label.1' }
        'textbox' { return 'Forms.TextBox.1' }
        'combobox' { return 'Forms.ComboBox.1' }
        'listbox' { return 'Forms.ListBox.1' }
        'checkbox' { return 'Forms.CheckBox.1' }
        'optionbutton' { return 'Forms.OptionButton.1' }
        'commandbutton' { return 'Forms.CommandButton.1' }
        'frame' { return 'Forms.Frame.1' }
        'togglebutton' { return 'Forms.ToggleButton.1' }
        default { throw ("Unsupported UserForm control type: {0}" -f $ControlType) }
    }
}

function Set-ControlFontProperties {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Control,

        [AllowNull()]
        [object]$FontSpec
    )

    if ($null -eq $FontSpec) {
        return
    }

    try {
        $font = $Control.Font
        try {
            if ($FontSpec.PSObject.Properties.Name -contains 'Name') { $font.Name = [string]$FontSpec.Name }
            if ($FontSpec.PSObject.Properties.Name -contains 'Size') { $font.Size = [single]$FontSpec.Size }
            if ($FontSpec.PSObject.Properties.Name -contains 'Bold') { $font.Bold = [bool]$FontSpec.Bold }
            if ($FontSpec.PSObject.Properties.Name -contains 'Italic') { $font.Italic = [bool]$FontSpec.Italic }
        }
        finally {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($font) } catch {}
        }
    }
    catch {
    }
}

function New-UserFormComponentFromSpec {
    param(
        [Parameter(Mandatory = $true)]
        [object]$VBProject,

        [Parameter(Mandatory = $true)]
        [object]$Spec,

        [Parameter(Mandatory = $true)]
        [string]$CodeText
    )

    $component = $null
    $designer = $null
    $controlMap = @{}
    $preferredInsideWidth = $null
    $preferredInsideHeight = $null
    $preferredOuterWidth = $null
    $preferredOuterHeight = $null
    $insideWidthApplied = $false
    $insideHeightApplied = $false

    try {
        $component = $VBProject.VBComponents.Add(3)
        $component.Name = [string]$Spec.name
        $designer = $component.Designer

        if ($Spec.PSObject.Properties.Name -contains 'insideWidth') {
            $preferredInsideWidth = $Spec.insideWidth
        }
        elseif ($Spec.PSObject.Properties.Name -contains 'width') {
            $preferredInsideWidth = $Spec.width
        }

        if ($Spec.PSObject.Properties.Name -contains 'insideHeight') {
            $preferredInsideHeight = $Spec.insideHeight
        }
        elseif ($Spec.PSObject.Properties.Name -contains 'height') {
            $preferredInsideHeight = $Spec.height
        }

        if ($Spec.PSObject.Properties.Name -contains 'outerWidth') {
            $preferredOuterWidth = $Spec.outerWidth
        }

        if ($Spec.PSObject.Properties.Name -contains 'outerHeight') {
            $preferredOuterHeight = $Spec.outerHeight
        }

        Set-ComPropertyIfPresent -Target $designer -PropertyName 'Caption' -Value $Spec.caption
        $insideWidthApplied = Try-SetComPropertyIfPresent -Target $designer -PropertyName 'InsideWidth' -Value $preferredInsideWidth
        $insideHeightApplied = Try-SetComPropertyIfPresent -Target $designer -PropertyName 'InsideHeight' -Value $preferredInsideHeight

        if ($preferredOuterWidth -ne $null) {
            Set-ComPropertyIfPresent -Target $designer -PropertyName 'Width' -Value $preferredOuterWidth
        }
        elseif (-not $insideWidthApplied) {
            Set-ComPropertyIfPresent -Target $designer -PropertyName 'Width' -Value $preferredInsideWidth
        }

        if ($preferredOuterHeight -ne $null) {
            Set-ComPropertyIfPresent -Target $designer -PropertyName 'Height' -Value $preferredOuterHeight
        }
        elseif (-not $insideHeightApplied) {
            Set-ComPropertyIfPresent -Target $designer -PropertyName 'Height' -Value $preferredInsideHeight
        }

        Set-ComPropertyIfPresent -Target $designer -PropertyName 'StartUpPosition' -Value $Spec.startUpPosition
        Set-ComPropertyIfPresent -Target $designer -PropertyName 'ScrollBars' -Value $Spec.scrollBars
        Set-ComPropertyIfPresent -Target $designer -PropertyName 'KeepScrollBarsVisible' -Value $Spec.keepScrollBarsVisible

        foreach ($controlSpec in @($Spec.controls)) {
            $parentContainer = $designer
            if ($controlSpec.PSObject.Properties.Name -contains 'parent' -and -not [string]::IsNullOrWhiteSpace([string]$controlSpec.parent)) {
                if ($controlMap.ContainsKey([string]$controlSpec.parent)) {
                    $parentContainer = $controlMap[[string]$controlSpec.parent]
                }
            }

            $control = $parentContainer.Controls.Add((Get-FormsProgId -ControlType ([string]$controlSpec.type)), [string]$controlSpec.name, $true)
            $controlMap[[string]$controlSpec.name] = $control

            Set-ComPropertyIfPresent -Target $control -PropertyName 'Caption' -Value $controlSpec.caption
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Left' -Value $controlSpec.left
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Top' -Value $controlSpec.top
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Width' -Value $controlSpec.width
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Height' -Value $controlSpec.height
            Set-ComPropertyIfPresent -Target $control -PropertyName 'TabIndex' -Value $controlSpec.tabIndex
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Visible' -Value $controlSpec.visible
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Enabled' -Value $controlSpec.enabled
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Value' -Value $controlSpec.value
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Text' -Value $controlSpec.text
            Set-ComPropertyIfPresent -Target $control -PropertyName 'ControlTipText' -Value $controlSpec.controlTipText
            Set-ComPropertyIfPresent -Target $control -PropertyName 'MultiLine' -Value $controlSpec.multiLine
            Set-ComPropertyIfPresent -Target $control -PropertyName 'WordWrap' -Value $controlSpec.wordWrap
            Set-ComPropertyIfPresent -Target $control -PropertyName 'EnterKeyBehavior' -Value $controlSpec.enterKeyBehavior
            Set-ComPropertyIfPresent -Target $control -PropertyName 'ScrollBars' -Value $controlSpec.scrollBars
            Set-ComPropertyIfPresent -Target $control -PropertyName 'AutoSize' -Value $controlSpec.autoSize
            Set-ComPropertyIfPresent -Target $control -PropertyName 'ColumnCount' -Value $controlSpec.columnCount
            Set-ComPropertyIfPresent -Target $control -PropertyName 'ListStyle' -Value $controlSpec.listStyle
            Set-ComPropertyIfPresent -Target $control -PropertyName 'Style' -Value $controlSpec.style

            if ($controlSpec.PSObject.Properties.Name -contains 'items') {
                foreach ($item in @($controlSpec.items)) {
                    try { [void]$control.AddItem([string]$item) } catch {}
                }
            }

            Set-ControlFontProperties -Control $control -FontSpec $controlSpec.font
        }

        if (-not [string]::IsNullOrWhiteSpace($CodeText)) {
            $component.CodeModule.AddFromString($CodeText)
        }

        return $component
    }
    finally {
        foreach ($controlName in $controlMap.Keys) {
            $controlObject = $controlMap[$controlName]
            if ($controlObject) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($controlObject) } catch {}
            }
        }

        if ($designer) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($designer) } catch {}
        }
    }
}

# ============================================================
# 项目源码扫描与导入
# ============================================================

function Get-ProjectSourceArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    # ⚠️ 防管道吞 null（2026-09-05 实测踩坑）：
    # PowerShell 中 `return $list.ToArray()` 会被管道展开，空数组/单元素数组在某些路径下变 $null，
    # 导致调用方 `-gt 0` 恒假、VBA 模块从不注入。必须用一元逗号 `return ,$list.ToArray()` 强制保留数组。
    $artifacts = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Config.ProjectSourceRoot)) {
        return ,$artifacts.ToArray()
    }

    if (-not [System.IO.Directory]::Exists($Config.ProjectSourceRoot)) {
        return ,$artifacts.ToArray()
    }

    $modulesDirectory = Join-Path $Config.ProjectSourceRoot 'modules'
    if ([System.IO.Directory]::Exists($modulesDirectory)) {
        foreach ($file in @(Get-ChildItem -Path $modulesDirectory -File | Where-Object { $_.Extension -in @('.vba', '.bas') } | Sort-Object Name)) {
            $content = Get-RenderedTextAsset -Path $file.FullName -Config $Config
            $artifacts.Add([pscustomobject]@{
                Type    = 'StandardModule'
                Name    = (Get-ComponentBaseName -FileName $file.Name)
                Content = $content
            }) | Out-Null
        }
    }

    $formsDirectory = Join-Path $Config.ProjectSourceRoot 'forms'
    if ([System.IO.Directory]::Exists($formsDirectory)) {
        foreach ($file in @(Get-ChildItem -Path $formsDirectory -File | Sort-Object Name)) {
            if ($file.Name -like '*.form.json') {
                $baseName = Get-ComponentBaseName -FileName $file.Name
                $codeFilePath = Join-Path $formsDirectory ($baseName + '.code.vba')
                $artifacts.Add([pscustomobject]@{
                    Type        = 'GeneratedUserForm'
                    Name        = $baseName
                    SpecContent = (Get-RenderedTextAsset -Path $file.FullName -Config $Config)
                    CodeContent = if ([System.IO.File]::Exists($codeFilePath)) { Get-RenderedTextAsset -Path $codeFilePath -Config $Config } else { '' }
                }) | Out-Null
            }
            elseif ($file.Extension -ieq '.frm') {
                $artifacts.Add([pscustomobject]@{
                    Type     = 'ImportedUserForm'
                    Name     = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    FormPath = $file.FullName
                    FrxPath  = [System.IO.Path]::ChangeExtension($file.FullName, '.frx')
                }) | Out-Null
            }
        }
    }

    return ,$artifacts.ToArray()
}

function Import-ProjectSourceArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$VBProject,

        [Parameter(Mandatory = $true)]
        [object]$Config,

        [Parameter(Mandatory = $true)]
        [object[]]$Artifacts
    )

    $temporaryDirectories = New-Object System.Collections.Generic.List[string]
    $importedComponents = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($artifact in $Artifacts) {
            switch ($artifact.Type) {
                'StandardModule' {
                    $module = $VBProject.VBComponents.Add(1)
                    $module.Name = $artifact.Name
                    $module.CodeModule.AddFromString([string]$artifact.Content)
                    $importedComponents.Add($module) | Out-Null
                }
                'GeneratedUserForm' {
                    $spec = $artifact.SpecContent | ConvertFrom-Json
                    $component = New-UserFormComponentFromSpec -VBProject $VBProject -Spec $spec -CodeText ([string]$artifact.CodeContent)
                    $importedComponents.Add($component) | Out-Null
                }
                'ImportedUserForm' {
                    $tempDirectory = New-TemporaryDirectory
                    $temporaryDirectories.Add($tempDirectory) | Out-Null

                    $tempFrmPath = Join-Path $tempDirectory ([System.IO.Path]::GetFileName($artifact.FormPath))
                    [System.IO.File]::WriteAllText($tempFrmPath, (Get-RenderedTextAsset -Path $artifact.FormPath -Config $Config), (Get-Utf8NoBomEncoding))

                    if ([System.IO.File]::Exists($artifact.FrxPath)) {
                        [System.IO.File]::Copy($artifact.FrxPath, (Join-Path $tempDirectory ([System.IO.Path]::GetFileName($artifact.FrxPath))), $true)
                    }

                    $component = $VBProject.VBComponents.Import($tempFrmPath)
                    $importedComponents.Add($component) | Out-Null
                }
            }
        }
    }
    finally {
        foreach ($component in $importedComponents) {
            if ($component) {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($component) } catch {}
            }
        }

        foreach ($tempDirectory in $temporaryDirectories) {
            Remove-DirectoryIfExists -Path $tempDirectory
        }
    }
}

# ============================================================
# Ribbon 注入（Zip 级操作 OOXML 包）
# ============================================================

function Get-ButtonImageXmlAttribute {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CustomImageId,

        [Parameter(Mandatory = $true)]
        [string]$ImageMsoName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$IconPath
    )

    if (Test-HasCustomButtonImage -IconPath $IconPath) {
        return ('image="{0}"' -f $CustomImageId)
    }

    return ('imageMso="{0}"' -f $ImageMsoName)
}

function Get-CustomUIXml {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    # 优先使用项目源码目录 src/ribbon/customUI14.xml（或 customUI.xml）：
    # 用户手动编排的 Ribbon 优先于模板生成的默认双按钮 Ribbon。
    $projectRibbonPath = $null
    if (-not [string]::IsNullOrWhiteSpace($Config.ProjectSourceRoot)) {
        foreach ($candidate in @('customUI14.xml', 'customUI.xml')) {
            $candidatePath = Join-Path (Join-Path $Config.ProjectSourceRoot 'ribbon') $candidate
            if ([System.IO.File]::Exists($candidatePath)) {
                $projectRibbonPath = $candidatePath
                break
            }
        }
    }

    if ($projectRibbonPath) {
        Write-Host "使用项目 Ribbon 源：$projectRibbonPath"
        $content = Read-TextFileAuto -Path $projectRibbonPath
        # 项目 Ribbon 若含模板占位符（__XXX__）则一并渲染，无占位符则原样使用
        return Expand-TemplateContent -TemplateContent $content -Tokens @{
            RIBBON_TAB_ID          = $Config.RibbonTabId
            RIBBON_TAB_LABEL       = $Config.RibbonTabLabel
            RIBBON_GROUP_ID        = $Config.RibbonGroupId
            RIBBON_GROUP_LABEL     = $Config.RibbonGroupLabel
            FIRST_BUTTON_ID        = $Config.FirstButtonId
            FIRST_BUTTON_LABEL     = $Config.FirstButtonLabel
            FIRST_CALLBACK_NAME    = $Config.FirstCallbackName
            FIRST_BUTTON_IMAGE_XML = (Get-ButtonImageXmlAttribute -CustomImageId "iconOne" -ImageMsoName $Config.FirstButtonImageMso -IconPath $Config.FirstIconPath)
            SECOND_BUTTON_ID       = $Config.SecondButtonId
            SECOND_BUTTON_LABEL    = $Config.SecondButtonLabel
            SECOND_CALLBACK_NAME   = $Config.SecondCallbackName
            SECOND_BUTTON_IMAGE_XML = (Get-ButtonImageXmlAttribute -CustomImageId "iconTwo" -ImageMsoName $Config.SecondButtonImageMso -IconPath $Config.SecondIconPath)
        }
    }

    Assert-FileExists -Path $Config.RibbonTemplatePath -DisplayName "Ribbon template"

    $template = Read-TextFileAuto -Path $Config.RibbonTemplatePath
    return Expand-TemplateContent -TemplateContent $template -Tokens @{
        RIBBON_TAB_ID         = $Config.RibbonTabId
        RIBBON_TAB_LABEL      = $Config.RibbonTabLabel
        RIBBON_GROUP_ID       = $Config.RibbonGroupId
        RIBBON_GROUP_LABEL    = $Config.RibbonGroupLabel
        FIRST_BUTTON_ID       = $Config.FirstButtonId
        FIRST_BUTTON_LABEL    = $Config.FirstButtonLabel
        FIRST_CALLBACK_NAME   = $Config.FirstCallbackName
        FIRST_BUTTON_IMAGE_XML = (Get-ButtonImageXmlAttribute -CustomImageId "iconOne" -ImageMsoName $Config.FirstButtonImageMso -IconPath $Config.FirstIconPath)
        SECOND_BUTTON_ID      = $Config.SecondButtonId
        SECOND_BUTTON_LABEL   = $Config.SecondButtonLabel
        SECOND_CALLBACK_NAME  = $Config.SecondCallbackName
        SECOND_BUTTON_IMAGE_XML = (Get-ButtonImageXmlAttribute -CustomImageId "iconTwo" -ImageMsoName $Config.SecondButtonImageMso -IconPath $Config.SecondIconPath)
    }
}

function Get-CustomUIRelationshipsXml {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $relationships = New-Object System.Collections.Generic.List[string]
    if (Test-HasCustomButtonImage -IconPath $Config.FirstIconPath) {
        [void]$relationships.Add('  <Relationship Id="iconOne" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/icon_one.png" />')
    }
    if (Test-HasCustomButtonImage -IconPath $Config.SecondIconPath) {
        [void]$relationships.Add('  <Relationship Id="iconTwo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/icon_two.png" />')
    }

    if ($relationships.Count -eq 0) {
        return ""
    }

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$(($relationships -join "`r`n"))
</Relationships>
"@
}

function Convert-ToWpsCompatibleRibbonPng {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )

    Add-Type -AssemblyName System.Drawing

    $fileStream = $null
    $sourceImage = $null
    $originalBitmap = $null
    $processedBitmap = $null
    $graphics = $null
    $outputPath = New-TemporaryFilePath -Extension ".png"

    try {
        $fileStream = [System.IO.File]::OpenRead($SourcePath)
        $sourceImage = [System.Drawing.Image]::FromStream($fileStream, $true, $true)
        $originalBitmap = New-Object System.Drawing.Bitmap($sourceImage)
        $processedBitmap = New-Object System.Drawing.Bitmap($originalBitmap.Width, $originalBitmap.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($processedBitmap)

        # 解决 WPS Ribbon 对透明 PNG 容易渲染出灰底的问题。
        $graphics.Clear([System.Drawing.Color]::FromArgb(1, 255, 255, 255))
        $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        [void]$graphics.DrawImage($originalBitmap, 0, 0, $originalBitmap.Width, $originalBitmap.Height)

        $processedBitmap.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host ("Prepared WPS-compatible ribbon PNG for {0}: {1}" -f $DisplayName, $SourcePath)
        return $outputPath
    }
    catch {
        if ([System.IO.File]::Exists($outputPath)) {
            [System.IO.File]::Delete($outputPath)
        }
        throw ("Failed to prepare WPS-compatible PNG for {0}: {1}" -f $DisplayName, $_.Exception.Message)
    }
    finally {
        foreach ($disposable in @($graphics, $processedBitmap, $originalBitmap, $sourceImage, $fileStream)) {
            if ($disposable) {
                try {
                    $disposable.Dispose()
                }
                catch {
                }
            }
        }
    }
}

function Ensure-ContentTypes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempFile,

        [Parameter(Mandatory = $true)]
        [bool]$NeedPngDefault,

        [Parameter(Mandatory = $true)]
        [string]$CustomUiPartName
    )

    [xml]$contentTypes = Read-TextFileAuto -Path $TempFile
    $ns = "http://schemas.openxmlformats.org/package/2006/content-types"

    if ($NeedPngDefault) {
        $pngDefault = $contentTypes.Types.Default | Where-Object { $_.Extension -eq "png" }
        if (-not $pngDefault) {
            $defaultNode = $contentTypes.CreateElement("Default", $ns)
            [void]$defaultNode.SetAttribute("Extension", "png")
            [void]$defaultNode.SetAttribute("ContentType", "image/png")
            [void]$contentTypes.Types.AppendChild($defaultNode)
        }
    }

    $staleOverrides = @($contentTypes.Types.Override | Where-Object {
            ($_.PartName -eq "/customUI/customUI.xml" -or $_.PartName -eq "/customUI/customUI14.xml") -and $_.PartName -ne $CustomUiPartName
        })
    foreach ($staleOverride in $staleOverrides) {
        [void]$contentTypes.Types.RemoveChild($staleOverride)
    }

    $customUiOverride = $contentTypes.Types.Override | Where-Object { $_.PartName -eq $CustomUiPartName } | Select-Object -First 1
    if (-not $customUiOverride) {
        $overrideNode = $contentTypes.CreateElement("Override", $ns)
        [void]$overrideNode.SetAttribute("PartName", $CustomUiPartName)
        [void]$overrideNode.SetAttribute("ContentType", "application/xml")
        [void]$contentTypes.Types.AppendChild($overrideNode)
    }
    else {
        [void]$customUiOverride.SetAttribute("ContentType", "application/xml")
    }

    Save-XmlDocument -Xml $contentTypes -Path $TempFile
}

function Ensure-RootRelationships {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TempFile,

        [Parameter(Mandatory = $true)]
        [string]$RelationshipType,

        [Parameter(Mandatory = $true)]
        [string]$CustomUiTarget
    )

    [xml]$rels = Read-TextFileAuto -Path $TempFile
    $ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    $existingNodes = @($rels.Relationships.Relationship | Where-Object {
            $_.Type -eq "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" -or
            $_.Type -eq "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" -or
            $_.Id -eq "rIdCustomUI"
        })

    foreach ($node in $existingNodes) {
        [void]$rels.Relationships.RemoveChild($node)
    }

    $relationshipNode = $rels.CreateElement("Relationship", $ns)
    [void]$relationshipNode.SetAttribute("Id", "rIdCustomUI")
    [void]$relationshipNode.SetAttribute("Type", $RelationshipType)
    [void]$relationshipNode.SetAttribute("Target", $CustomUiTarget)
    [void]$rels.Relationships.AppendChild($relationshipNode)

    Save-XmlDocument -Xml $rels -Path $TempFile
}

function Update-OfficePackage {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $contentTypesTemp = [System.IO.Path]::GetTempFileName()
    $relsTemp = [System.IO.Path]::GetTempFileName()
    $customUiRelsContent = Get-CustomUIRelationshipsXml -Config $Config
    $needCustomImageRelationships = -not [string]::IsNullOrWhiteSpace($customUiRelsContent)
    $needPngDefault = (Test-HasCustomButtonImage -IconPath $Config.FirstIconPath) -or (Test-HasCustomButtonImage -IconPath $Config.SecondIconPath)
    $temporaryIconPaths = New-Object System.Collections.Generic.List[string]
    $firstPackageIconPath = $Config.FirstIconPath
    $secondPackageIconPath = $Config.SecondIconPath
    $customUiXml = Get-CustomUIXml -Config $Config
    $customUiPartInfo = Get-CustomUIPartInfo -CustomUIXml $customUiXml

    if (Test-HasCustomButtonImage -IconPath $Config.FirstIconPath) {
        $firstPackageIconPath = Convert-ToWpsCompatibleRibbonPng -SourcePath $Config.FirstIconPath -DisplayName "First button icon"
        [void]$temporaryIconPaths.Add($firstPackageIconPath)
    }
    if (Test-HasCustomButtonImage -IconPath $Config.SecondIconPath) {
        $secondPackageIconPath = Convert-ToWpsCompatibleRibbonPng -SourcePath $Config.SecondIconPath -DisplayName "Second button icon"
        [void]$temporaryIconPaths.Add($secondPackageIconPath)
    }

    # 通过直接修改 OOXML 包，把 Ribbon 和所需资源塞进最终产物。
    $zip = [System.IO.Compression.ZipFile]::Open($Config.OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $contentTypesEntry = $zip.GetEntry("[Content_Types].xml")
        if (-not $contentTypesEntry) {
            throw "Missing [Content_Types].xml in package."
        }

        $relsEntry = $zip.GetEntry("_rels/.rels")
        if (-not $relsEntry) {
            throw "Missing _rels/.rels in package."
        }

        $contentTypesStream = $contentTypesEntry.Open()
        $contentTypesFileStream = [System.IO.File]::Create($contentTypesTemp)
        try {
            $contentTypesStream.CopyTo($contentTypesFileStream)
        }
        finally {
            $contentTypesFileStream.Dispose()
            $contentTypesStream.Dispose()
        }

        $relsStream = $relsEntry.Open()
        $relsFileStream = [System.IO.File]::Create($relsTemp)
        try {
            $relsStream.CopyTo($relsFileStream)
        }
        finally {
            $relsFileStream.Dispose()
            $relsStream.Dispose()
        }

        Ensure-ContentTypes -TempFile $contentTypesTemp -NeedPngDefault $needPngDefault -CustomUiPartName $customUiPartInfo.PartName
        Ensure-RootRelationships -TempFile $relsTemp -RelationshipType $customUiPartInfo.RelationshipType -CustomUiTarget $customUiPartInfo.EntryName

        $contentTypesEntry.Delete()
        $relsEntry.Delete()

        Set-ZipTextEntry -Zip $zip -EntryName "[Content_Types].xml" -Content (Read-TextFileAuto -Path $contentTypesTemp)
        Set-ZipTextEntry -Zip $zip -EntryName "_rels/.rels" -Content (Read-TextFileAuto -Path $relsTemp)
        Remove-ZipEntryIfExists -Zip $zip -EntryName $customUiPartInfo.AlternateEntryName
        Set-ZipTextEntry -Zip $zip -EntryName $customUiPartInfo.EntryName -Content $customUiXml
        if ($needCustomImageRelationships) {
            Remove-ZipEntryIfExists -Zip $zip -EntryName $customUiPartInfo.AlternateRelationshipsEntryName
            Set-ZipTextEntry -Zip $zip -EntryName $customUiPartInfo.RelationshipsEntryName -Content $customUiRelsContent
        }
        else {
            Remove-ZipEntryIfExists -Zip $zip -EntryName $customUiPartInfo.RelationshipsEntryName
            Remove-ZipEntryIfExists -Zip $zip -EntryName $customUiPartInfo.AlternateRelationshipsEntryName
        }

        foreach ($entryName in @("customUI/images/icon_one.png", "customUI/images/icon_two.png")) {
            $existingImageEntry = $zip.GetEntry($entryName)
            if ($existingImageEntry) {
                $existingImageEntry.Delete()
            }
        }

        if (Test-HasCustomButtonImage -IconPath $Config.FirstIconPath) {
            Set-ZipBinaryEntry -Zip $zip -EntryName "customUI/images/icon_one.png" -SourcePath $firstPackageIconPath
        }
        if (Test-HasCustomButtonImage -IconPath $Config.SecondIconPath) {
            Set-ZipBinaryEntry -Zip $zip -EntryName "customUI/images/icon_two.png" -SourcePath $secondPackageIconPath
        }
    }
    finally {
        $zip.Dispose()
        $pathsToDelete = @($contentTypesTemp, $relsTemp) + @($temporaryIconPaths.ToArray())
        foreach ($pathToDelete in $pathsToDelete) {
            if ([string]::IsNullOrWhiteSpace($pathToDelete)) {
                continue
            }

            if ([System.IO.File]::Exists($pathToDelete)) {
                [System.IO.File]::Delete($pathToDelete)
            }
        }
    }
}

# ============================================================
# 构建主流程
# ============================================================

function Get-QuickstartWorkbookPath {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    if ([string]::IsNullOrWhiteSpace($Config.QuickstartProjectPath)) {
        return ""
    }

    if (-not [System.IO.Directory]::Exists($Config.QuickstartProjectPath)) {
        throw "Quickstart 工程目录不存在：$($Config.QuickstartProjectPath)"
    }

    # 优先找与 ProjectName 同名的工作簿（xlam/xlsm），其次任意 xlam/xlsm
    $candidates = @()
    foreach ($ext in @('.xlam', '.xlsm')) {
        $named = Join-Path $Config.QuickstartProjectPath ($Config.ProjectName + $ext)
        if ([System.IO.File]::Exists($named)) {
            return $named
        }
    }

    $anyBook = @(Get-ChildItem -Path $Config.QuickstartProjectPath -File | Where-Object { $_.Extension -in @('.xlam', '.xlsm') }) | Select-Object -First 1
    if ($anyBook) {
        return $anyBook.FullName
    }

    throw "Quickstart 工程中未找到 .xlam/.xlsm 文件：$($Config.QuickstartProjectPath)"
}

function Build-Addin {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $trackedExcel = $null
    $excel = $null
    $workbook = $null
    $vbProject = $null
    $module = $null

    $quickstartWorkbookPath = Get-QuickstartWorkbookPath -Config $Config

    try {
        $trackedExcel = Start-TrackedExcelApplication
        $excel = $trackedExcel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        if (-not [string]::IsNullOrWhiteSpace($quickstartWorkbookPath)) {
            Write-Host "基于 quickstart 工作簿构建：$quickstartWorkbookPath"
            $workbook = $excel.Workbooks.Open($quickstartWorkbookPath)
            # 以工作簿实际格式为准；xlam 保持 IsAddin
            if ($Config.OutputKind -eq "xlam") {
                try { $workbook.IsAddin = $true } catch {}
            }
            else {
                try { $workbook.IsAddin = $false } catch {}
            }
        }
        else {
            Write-Host "从零空白工作簿构建（未提供 -QuickstartProjectPath）"
            $workbook = $excel.Workbooks.Add()
            $workbook.IsAddin = $true
        }

        $vbProject = $workbook.VBProject

        $projectSourceArtifacts = Get-ProjectSourceArtifacts -Config $Config
        if ($projectSourceArtifacts -and $projectSourceArtifacts.Count -gt 0) {
            Import-ProjectSourceArtifacts -VBProject $vbProject -Config $Config -Artifacts $projectSourceArtifacts
        }
        elseif ([string]::IsNullOrWhiteSpace($quickstartWorkbookPath)) {
            # 仅从零模式且无源码时才注入默认模块；quickstart 模式保留工程原有模块
            $module = $vbProject.VBComponents.Add(1)
            $module.Name = $Config.ModuleName
            $module.CodeModule.AddFromString((Get-VbaModuleCode -Config $Config))
        }

        if ([System.IO.File]::Exists($Config.OutputPath)) {
            [System.IO.File]::Delete($Config.OutputPath)
        }

        $workbook.SaveAs($Config.OutputPath, $Config.VbaFileFormat)
    }
    catch {
        if ($_.Exception.Message -match "Programmatic access to Visual Basic Project is not trusted") {
            throw "Excel still blocks VBProject access. Check AccessVBOM and local policy."
        }

        throw
    }
    finally {
        if ($workbook) {
            try {
                $workbook.Close($false)
            }
            catch {
            }
        }

        foreach ($comObject in @($module, $vbProject, $workbook)) {
            if ($comObject) {
                try {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
                }
                catch {
                }
            }
        }

        if ($trackedExcel) {
            Stop-TrackedExcelApplication -TrackedExcel $trackedExcel
        }
    }
}

# ============================================================
# 安装/卸载配套产物
# ============================================================

function Write-InstallArtifacts {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Config,

        [Parameter(Mandatory = $true)]
        [string]$FallbackVersion
    )

    if ($Config.OutputKind -ne "xlam") {
        return
    }

    $directory = Split-Path -Path $Config.OutputPath -Parent
    $scriptsDirectory = Join-Path $directory "scripts"
    $addinFileName = Split-Path -Path $Config.OutputPath -Leaf
    $baseName = $Config.HelperBaseName
    $installGuidePath = Join-Path $directory "安装说明.txt"

    [void][System.IO.Directory]::CreateDirectory($scriptsDirectory)

    $variants = @(
        @{
            Key            = 'excel'
            InstallPs1Path = (Join-Path $scriptsDirectory ("install_excel_{0}.ps1" -f $baseName))
            UninstallPs1Path = (Join-Path $scriptsDirectory ("uninstall_excel_{0}.ps1" -f $baseName))
            InstallBatPath = (Join-Path $directory ("install_excel_{0}.bat" -f $baseName))
            UninstallBatPath = (Join-Path $directory ("uninstall_excel_{0}.bat" -f $baseName))
            HostName       = 'Microsoft Excel'
            InstallMessage = 'Microsoft Excel add-in installed successfully.'
            UninstallMessage = 'Microsoft Excel add-in uninstalled successfully.'
        }
        @{
            Key            = 'wps'
            InstallPs1Path = (Join-Path $scriptsDirectory ("install_wps_{0}.ps1" -f $baseName))
            UninstallPs1Path = (Join-Path $scriptsDirectory ("uninstall_wps_{0}.ps1" -f $baseName))
            InstallBatPath = (Join-Path $directory ("install_wps_{0}.bat" -f $baseName))
            UninstallBatPath = (Join-Path $directory ("uninstall_wps_{0}.bat" -f $baseName))
            HostName       = 'WPS 表格'
            InstallMessage = 'WPS Spreadsheets add-in installed successfully.'
            UninstallMessage = 'WPS Spreadsheets add-in uninstalled successfully.'
        }
    )

    foreach ($variant in $variants) {
        $installPs1Content = Get-RenderedTemplateContent -TemplatePath $Config.InstallPs1TemplatePath -Tokens @{
            ADDIN_FILE_NAME  = $addinFileName
            FALLBACK_VERSION = $FallbackVersion
            INSTALL_TARGET   = $variant.Key
            TARGET_HOST_NAME = $variant.HostName
        }
        $uninstallPs1Content = Get-RenderedTemplateContent -TemplatePath $Config.UninstallPs1TemplatePath -Tokens @{
            ADDIN_FILE_NAME  = $addinFileName
            FALLBACK_VERSION = $FallbackVersion
            INSTALL_TARGET   = $variant.Key
            TARGET_HOST_NAME = $variant.HostName
        }
        $installBatContent = Get-RenderedTemplateContent -TemplatePath $Config.InstallBatTemplatePath -Tokens @{
            POWERSHELL_SCRIPT_NAME = [System.IO.Path]::GetFileName($variant.InstallPs1Path)
            SUCCESS_MESSAGE        = $variant.InstallMessage
        }
        $uninstallBatContent = Get-RenderedTemplateContent -TemplatePath $Config.UninstallBatTemplatePath -Tokens @{
            POWERSHELL_SCRIPT_NAME = [System.IO.Path]::GetFileName($variant.UninstallPs1Path)
            SUCCESS_MESSAGE        = $variant.UninstallMessage
        }

        Write-TextFileUtf8Bom -Path $variant.InstallPs1Path -Content $installPs1Content
        Write-TextFileUtf8Bom -Path $variant.UninstallPs1Path -Content $uninstallPs1Content
        Write-TextFileAscii -Path $variant.InstallBatPath -Content $installBatContent
        Write-TextFileAscii -Path $variant.UninstallBatPath -Content $uninstallBatContent
    }

    $installGuideContent = @"
$($Config.ProjectName) 插件安装说明

一、文件说明
1. $addinFileName
   Excel/WPS 加载项主文件。
2. install_excel_$baseName.bat / uninstall_excel_$baseName.bat
   只对 Microsoft Excel 写入或移除加载项注册，不影响 WPS 表格。
3. install_wps_$baseName.bat / uninstall_wps_$baseName.bat
   只对 WPS 表格写入或移除加载项注册，不影响 Microsoft Excel。
4. scripts\
   存放实际执行的 PowerShell 安装/卸载脚本，通常不需要手动打开编辑。

二、安装步骤
1. 保持 dist 文件夹结构完整，不要单独移动 $addinFileName 或 scripts 文件夹。
2. 如果只是临时使用，可直接双击 $addinFileName，或右键选择 Excel/WPS 打开该加载项。
3. 只想在 Excel 生效时，双击 install_excel_$baseName.bat。
4. 只想在 WPS 表格生效时，双击 install_wps_$baseName.bat。
5. 安装前请先关闭对应宿主，安装完成后重新打开。

三、卸载步骤
1. 关闭对应宿主程序。
2. Excel 版本运行 uninstall_excel_$baseName.bat。
3. WPS 版本运行 uninstall_wps_$baseName.bat。
4. 重新打开对应宿主，确认插件入口已移除。

四、常见说明
1. 若系统提示脚本权限问题，请优先使用 bat 文件启动，不要直接双击 ps1。
2. 各安装脚本只修改对应宿主的注册机制，适合同时安装 Microsoft Office 与 WPS Office 的场景。
3. 如果你要把插件发给别人，请将整个 dist 文件夹一起打包分发。
4. 已安装后不要随意移动 $addinFileName 或 scripts 文件夹；如果换了位置，建议重新运行对应安装脚本。
"@

    Write-TextFileUtf8Bom -Path $installGuidePath -Content $installGuideContent
}

# ============================================================
# 主流程
# ============================================================

Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)"
Write-Host "Loading shared library: $commonLibPath"

# 构建前杀残留 Excel/WPS 进程（防止目标文件被占用导致 PermissionError）
$residualProcs = @(Get-Process -Name "EXCEL","ET" -ErrorAction SilentlyContinue)
if ($residualProcs.Count -gt 0) {
    Write-Host "[WARN] 发现 $($residualProcs.Count) 个残留 Office 进程，构建前终止..."
    foreach ($p in $residualProcs) {
        try {
            Stop-Process -Id $p.Id -Force -ErrorAction Stop
            Write-Host "  已终止 PID=$($p.Id) ($($p.Name))"
        } catch {
            Write-Host "  终止 PID=$($p.Id) 失败: $_" -ForegroundColor Yellow
        }
    }
    Start-Sleep -Seconds 2
    Write-Host "[OK] 残留进程已清理"
} else {
    Write-Host "[OK] 无残留 Office 进程"
}

$wpsProjectTrustEnabled = Test-WpsProjectTrustEnabled
Write-Host "WPS VBA project trust enabled: $wpsProjectTrustEnabled"
if (-not $wpsProjectTrustEnabled) {
    Write-Host "Enabling WPS VBA project trust before any COM automation..."
    Set-WpsProjectTrust -Enable $true
}

$preflightExcelVersions = @($OfficeVersions)
Write-Host "Excel AccessVBOM preflight versions: $($preflightExcelVersions -join ', ')"
$preflightAccessVBOMEnabled = Ensure-AccessVBOMEnabledForVersions -ExcelVersions $preflightExcelVersions
Write-Host "AccessVBOM preflight already enabled: $preflightAccessVBOMEnabled"
if (-not $preflightAccessVBOMEnabled) {
    Write-Host "AccessVBOM registry updated. Existing user Excel processes are NOT affected."
}

$config = New-BuildConfig
$outputDirectory = Split-Path -Path $config.OutputPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void][System.IO.Directory]::CreateDirectory($outputDirectory)
}

if (Test-HasCustomButtonImage -IconPath $config.FirstIconPath) {
    Assert-FileExists -Path $config.FirstIconPath -DisplayName "First button icon"
}
if (Test-HasCustomButtonImage -IconPath $config.SecondIconPath) {
    Assert-FileExists -Path $config.SecondIconPath -DisplayName "Second button icon"
}

Write-Host "Detecting Office automation host..."
$hostInfo = Get-OfficeAutomationHostInfo
Assert-ValidExcelComHost -HostInfo $hostInfo

Write-Host "Office host version: $($hostInfo.Version)"
Write-Host "Office host path: $($hostInfo.Path)"

$excelVersion = $hostInfo.Version

$accessVBOMEnabled = Test-AccessVBOMEnabled -ExcelVersion $excelVersion
Write-Host "AccessVBOM enabled: $accessVBOMEnabled"

if (-not $accessVBOMEnabled) {
    throw "AccessVBOM should have been enabled before COM startup, but the detected host version is still disabled."
}

Write-Host "Building $($config.OutputKind) package..."
Build-Addin -Config $config

Write-Host "Injecting custom ribbon and icon settings..."
Update-OfficePackage -Config $config

if ($ReinjectRibbonAfterActivation) {
    Write-Host "Reinjecting ribbon after activation (COM save may have dropped customUI registration)..."
    if (Test-Path -LiteralPath $config.OutputPath -PathType Leaf) {
        Update-OfficePackage -Config $config
        Write-Host "Ribbon re-injection done."
    }
    else {
        Write-Warning "产物不存在，跳过 Ribbon 回注：$($config.OutputPath)"
    }
}

if ($config.OutputKind -eq "xlam") {
    Write-Host "Writing install and uninstall helpers..."
    Write-InstallArtifacts -Config $config -FallbackVersion $excelVersion
}

if ($InstallAfterBuild) {
    if ($config.OutputKind -ne "xlam") {
        throw "InstallAfterBuild only supports xlam output."
    }

    Write-Host "Registering add-in in Excel and WPS registry..."
    Register-ExcelAddinRegistry -AddinPath $config.OutputPath -FallbackVersion $excelVersion
    Set-WpsLoadMacroRegistration -AddinPath $config.OutputPath -Enable $true
    Set-WpsProjectTrust -Enable $true
}

Write-Host "Done."
Write-Host "Output: $($config.OutputPath)"
