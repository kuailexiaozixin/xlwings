# 编码与 COM 前置
这个文档保留 Excel COM、AccessVBOM、WPS 信任和 VBA / Ribbon 交付所需的前置约束。
如果问题更偏向 PowerShell `5.1` 的文件读写、模板实例化、中文乱码、BOM / 无 BOM 处理，优先看本文件末尾"PowerShell 5.1 推荐写法与排查"节。
## 编码固定规则

| 文件类型 | 推荐编码 | 说明 |
|----------|----------|------|
| `.ps1` / `.ps1.tmpl` | `UTF-8 with BOM` | 兼容 Windows PowerShell 5.x |
| `.vba.tmpl` | `UTF-8 with BOM` | 便于写中文注释、中文 MsgBox |
| `.xml.tmpl` | `UTF-8 with BOM` | 便于保留中文模板文本 |
| 包内 XML (`customUI.xml` / `customUI14.xml` / `.rels`) | `UTF-8 without BOM` | 写入 `xlam/xlsm` ZIP 包内时统一无 BOM |
| `.bat` | `ASCII` 优先 | 只做英文启动壳，不承担中文逻辑 |

## 先分三种场景

### 1. 内部可控文件

- 例如：`templates\*.ps1.tmpl`、`.vba.tmpl`、临时生成的中间文本
- 可以统一约定编码
- `.ps1` 默认 `UTF-8 with BOM`

### 2. 外部未知文件

- 例如：用户给的历史脚本、第三方导出的 XML / TXT
- 不要直接假定 `UTF-8`
- 不要把 `Get-Content -Raw -Encoding UTF8` 当成万能读取方式

### 3. 对外交付文件

- 例如：写入 `xlam/xlsm` ZIP 包内的 XML、发给用户或第三方系统导入的文本
- 是否带 BOM 要看目标消费者
- 包内 XML 继续保持 `UTF-8 without BOM`

## 普通 AI 的最简规则

1. 只要是 PowerShell 脚本，就用 `UTF-8 with BOM`
2. 只要是写进 `xlam/xlsm` ZIP 包里的 XML，就用 `UTF-8 without BOM`
3. 只要是 `.bat`，就尽量不要放中文
4. 不要用默认 `Get-Content` / `Set-Content` 读写中文模板
5. 不要把 `-Raw` 理解成自动识别编码
6. 做用户项目时，必须通过 `xlwings quickstart` 生成项目骨架（或基于现有文件走 `import_existing_workbook.ps1`），禁止手动复制 `templates/` 文件夹

## 典型乱码案例 (反面教材)

如果你看到类似下面的乱码，说明编码处理出了问题：

- **现象**：`label="Excel鍚堝苟宸ュ叿"` (本意是 "Excel合并工具")
- **原因**：UTF-8 字节流被错误地识别为了 GBK/ANSI 编码。
- **罪魁祸首**：
    1. 使用了 PowerShell 的 `>` 或 `Out-File` 而没有指定 `-Encoding utf8`。
    2. XML 文件带了 BOM 头，导致 Excel 识别引擎切换到了错误的编码模式。
    3. 注入 ZIP 包时，调用了会自动转换编码的中间件。

## 规避指南

1. **统一标准**：包内所有 XML 必须是 **UTF-8 无 BOM**。
2. **禁止使用 > 重定向**：在生成 XML 时，必须使用 `[System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding($false)))` 来确保无 BOM。
3. **自检**：如果 XML 中包含中文，生成后必须检查字节流，确保没有 `EF BB BF` 这三个字节出现在文件头。

## 模板安全规则

### 做用户项目时

- **禁止**直接修改 `templates\`（技能本体模板仓库），也**禁止**手动复制 `templates\` 目录到工作区
- 唯一合法方式：通过 `xlwings quickstart` 生成骨架（或 `../scripts/import_existing_workbook.ps1` 导入现有文件）、`../scripts/instantiate_blueprint.ps1` 读取模板、替换占位符、写入用户项目目录
- PS1 脚本执行后，在生成的用户项目文件中编写业务代码，不要回头修改原始 `templates/`

### 维护 skill 本体时

- 只有明确是在维护本技能（xlwings）本体时，才修改原始 `templates\`
- 修改 `.ps1.tmpl` 时，继续保持 `UTF-8 with BOM`
- 修改 ZIP 包内 XML 模板时，继续保持写入结果为 `UTF-8 without BOM`

## AccessVBOM 与 WPS 信任

在调用任何 `Excel.Application` / Office COM 之前，先处理这些前置条件：

### Excel

- `AccessVBOM`
- 普通用户电脑场景下，可直接覆盖常见版本位：`12.0`、`14.0`、`15.0`、`16.0`

### WPS

- 注册表路径：`HKCU\Software\kingsoft\office\6.0\et\Application Settings`
- **必须开启**以下两项以允许 COM 写入 VBA：
    - `KDEVBProjectTrust = "1"` (字符串类型)
    - `JSIDEProjectTrust = 1` (DWord 类型)

## .bat 特别说明

`.bat` 不要默认承载中文提示。更稳妥的方案：

- `.bat`：只保留英文、路径、启动命令
- `.ps1`：承载中文提示、中文报错、中文逻辑

## 环境预检快捷命令

在技能根目录下执行以下命令完成环境就绪预检。本段为可直接复制的完整脚本；`$skillRoot` 按技能实际位置替换——**技能外发后路径会变，禁止写死任何机器专属绝对路径**，解释器路径属于项目级配置（写入配置表 `Interpreter_Win`），不属于技能规则：

```powershell
$skillRoot = "C:\path\to\xlwings"   # 替换为技能实际根目录（外发后按新位置改）
. "$skillRoot\scripts\lib\excel-com-common.ps1"

# 1. 开发解释器身份（显式确认，勿依赖 PATH 命令名；见 docs/troubleshooting.md 坑 10）
python -c "import sys; print('exe:', sys.executable)"
python -c "import sys; print('ver:', sys.version.split()[0])"
python -c "import xlwings; print('xlwings', xlwings.__version__)"

# 2. Office 宿主与版本（均无参）
Get-OfficeAutomationHostInfo | Format-List Version, Caption, Path, IsWpsHost
Get-ExcelVersion

# 3. AccessVBOM（强制参数 -ExcelVersions，按本机 Office 版本传入，可多版本）
Test-AccessVBOMEnabledForVersions -ExcelVersions @("14.0", "15.0", "16.0")

# 4. WPS 信任（无参；仅目标含 WPS 时需要）
Test-WpsProjectTrustEnabled

# 5. 数据源连通性（按项目实际外部依赖调整；以下为通用示例，
#    替换为目标 API 的真实端点与参数）
python -c "import requests; print('HTTP', requests.get('https://api.example.com/endpoint', params={'key':'value'}, timeout=15).status_code)"
```

> **函数签名速查**：`Get-OfficeAutomationHostInfo` / `Get-ExcelVersion` / `Test-WpsProjectTrustEnabled` 均无参数；`Test-AccessVBOMEnabledForVersions` 强制 `-ExcelVersions <string[]>`；`Get-ExcelRegistryVersions` 强制 `-FallbackVersion <string>`。签名以 `scripts/lib/excel-com-common.ps1` 为准。

## 最低验证要求

- `.ps1` 头字节是否为 `EF BB BF`
- 包内 XML 首字节是否不是 `EF BB BF`
- Excel 实测 Ribbon 中文是否正常
- VBA 中文 `MsgBox` 是否正常

## PowerShell 5.1 推荐写法与排查（原 08-powershell 专题）

### 一句话原则

- `.ps1` / `.ps1.tmpl` 默认保存为 `UTF-8 with BOM`
- `-Raw` 只表示整文件读取，**不负责识别原始编码**
- `-Encoding UTF8` 只表示按 UTF-8 解码，**不表示自动识别或自动转码**
- 包内 XML（`customUI.xml` / `customUI14.xml` / `.rels`）默认写成 `UTF-8 without BOM`

### 推荐写法

```powershell
# 写 .ps1（UTF-8 with BOM）
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($scriptPath, $content, $utf8Bom)

# 写包内 XML（UTF-8 without BOM）
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($xmlPath, $content, $utf8NoBom)

# 读内部已知 UTF-8 文件（前提：确定文件本来就是 UTF-8）
$text = Get-Content $path -Raw -Encoding UTF8
```

### 禁止误解

- 不要把 `-Raw` 理解成自动识别编码
- 不要把 `-Encoding UTF8` 理解成自动转码
- 不要把"脚本能跑"误认为"对外交付兼容性没问题"
- 不要在用户项目开发时直接改全局 skill 里的原模板

### 出现问题时的排查顺序

1. 先确认是不是 Windows PowerShell `5.1`
2. 再确认当前是内部文件、外部未知文件，还是对外交付文件
3. 如果是 `.ps1`，先检查是否为 `UTF-8 with BOM`
4. 如果是包内 XML，检查是否误带了 BOM
5. 检查读写代码是否显式指定了正确编码
6. 最后再排查业务逻辑
