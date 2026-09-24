# 部署与交付工作流（6.5.10/6.5.11/6.5.12/6.5.13 扩展）

> 本文件是工作流 6.5.10（部署与交付物整备）、6.5.11（安装）、6.5.12（安装后集成验证）、6.5.13（分发）的扩展资源，把四个步骤编排为可执行工作流。
> **主线 = 自研 release_tool 便携运行时分发**（无需 PRO 许可证）。官方 xlwings PRO 部署机制（code embed / release）仅作附录 A 对照研读——理解其设计、借鉴其细节，但开发加载项一律按自研 release_tool 实施，不采用 PRO 机制。

## 目录

- [〇 工作流总览](#〇-工作流总览)
- [阶段一 分发方式选型](#阶段一-分发方式选型)
- [阶段二 便携运行时打包](#阶段二-便携运行时打包)
- [阶段三 代码嵌入——自研 code embed](#阶段三-代码嵌入自研-code-embed)
- [阶段四 配置表整备](#阶段四-配置表整备)
- [阶段五 门禁与真机验证](#阶段五-门禁与真机验证)
- [阶段六 交付清单](#阶段六-交付清单)
- [附录 A PRO 部署机制对照研读（迁移自 13 号 1.4）](#附录-a-pro-部署机制对照研读迁移自-13-号-14)
- [附录 B 官方文档映射](#附录-b-官方文档映射)
- [可配套阅读](#可配套阅读)

## 〇 工作流总览

| 阶段 | 目的 | 输出物 | 验收要点 |
|------|------|--------|---------|
| 一 分发方式选型 | 按目标机环境选分发路线 | 选型决策 | 目标机无 Python → release_tool 路线；排除项写明原因 |
| 二 便携运行时打包 | 构建目标机可独立运行的 Python 运行时 | `dist_release/runtime/` | `runtime/python.exe` 可启动；C 扩展包经包管理器安装；`import numpy/pandas` 成功 |
| 三 代码嵌入（自研 code embed） | 单文件分发，代码不外露 | xlam 含 .py sheet + MAP | .py sheet 存在；MAP 与文件一一对应；VBA 无 LICENSE_KEY 检查 |
| 四 配置表整备 | 让便携运行时真正生效 | 修复后的 `xlwings.conf` | 无重复键；Interpreter_Win 指向 runtime；PYTHONPATH 正斜杠；PROJECT_NAME="xlwings" |
| 五 门禁与真机验证 | 安装后确认可用 | 门禁 E/J/K/L 记录 + 冒烟结果 | 真机加载无报错；Ribbon 可见；点击触发 Python；业务端到端可用 |
| 六 交付清单 | 交付前逐项核对 | 核对记录 | 清单 15 项全部勾选；无开发机硬编码路径 |

**默认假设**：目标机未安装 Python——自研 release_tool 便携运行时分发是标配，不是例外分支。

**路由判据**：
- 开发加载项 / 对外分发 / 目标机无 Python → 走本工作流（自研 release_tool）
- 想理解官方 PRO 机制设计（LicenseHandler / Deploy Key / RELEASE_NO_ADDIN / 嵌入运行时）→ 读附录 A，仅借鉴不采用
- PRO 机制依赖 PRO 许可证（无许可调用抛 LicenseError），本技能不部署运行 PRO（见场景 D 主干规则）

## 阶段一 分发方式选型

- **任务**：按目标机环境与代码形态选分发路线。
- **要求**：默认假设目标机无 Python；逐行比对选型矩阵，排除项写明原因。
- **输出物**：选型决策（路线 + 代码分发模式）。
- **验收**：目标机无 Python → release_tool 路线；代码需易改 → PYTHONPATH 模式；单文件分发 / 代码不外露 → code embed 模式。

**选型矩阵**：

| 方式 | 适用场景 | 关键操作 | 目标机要求 |
|------|---------|---------|-----------|
| **XLSTART 自动加载** | 本机自用 / 固定机器 | xlam 放入 `%APPDATA%\Microsoft\Excel\XLSTART\`，Excel 启动自动加载 | 需 Python + xlwings |
| **手动加载项** | 分发他人 / 可选加载 | Excel → 文件 → 选项 → 加载项 → 浏览 → 选 xlam | 需 Python + xlwings |
| **release_tool - PYTHONPATH 模式** | 目标机无 Python，代码需易修改 | xlam + runtime/ + src/ + 入口模块，install_release.bat 一键安装 | 无需预装 Python |
| **release_tool - code embed 模式**（推荐） | 目标机无 Python，单文件分发，代码不外露 | xlam（含 .py 隐藏 sheet）+ runtime/，install_release.bat 一键安装 | 无需预装 Python |

**code embed vs PYTHONPATH 模式对比**：

| 维度 | code embed | PYTHONPATH |
|------|-----------|------------|
| 代码存储 | .py 隐藏 sheet（UUID 命名） | 外部 src/ 目录 |
| 配置表 | RELEASE_EMBED_CODE=True + RELEASE_EMBED_CODE_MAP | RELEASE_EMBED_CODE=False |
| 运行时 | VBA 检测 .py sheet → embedded_code_extractor 提取到 %TEMP% → 执行 | 直接从 PYTHONPATH 目录 import |
| 单文件分发 | 是（代码在 xlam 内） | 否（需附带 src/） |
| 代码修改 | 需重新嵌入 | 直接修改 src/ 文件 |
| 许可证 | 无需 PRO（自研 embedded_code_extractor） | 无需 PRO |

## 阶段二 便携运行时打包

- **任务**：构建目标机可独立运行的 Python 运行时。
- **要求**：C 扩展包必须经包管理器安装；禁止从开发机复制 numpy/pandas；构建后清理冗余。
- **输出物**：`dist_release/runtime/`（~120MB）+ 冒烟通过。
- **验收**：`runtime/python.exe` 可启动；`import numpy/pandas` 成功；`__pycache__` 等已清理。

**推荐方式**：`install_runtime.py install_minimal`（~120MB，uv 安装依赖）

```powershell
# 构建最小便携运行时
python scripts/release_tool/install_runtime.py install_minimal --target .\dist_release\runtime
```

**关键铁律**：
- 禁止从开发机复制 numpy/pandas（C 扩展 DLL 不完整，_multiarray_umath 加载失败）
- 必须用包管理器（uv/pip）安装 C 扩展包
- 安装后清理 `__pycache__`、tests、docs 等不必要文件
- 长路径删除用 robocopy 镜像方式（PowerShell Remove-Item -Recurse 无法处理 >260 字符路径）

**传统方式（备选）**：从 python.org 下载 `windows embeddable package`，解压后用 pip 安装依赖。

**目录结构（dist_release/ 内）**：

```
dist_release/
├── <项目名>.xlam              # 加载项（code embed 模式含 .py sheet）
├── <项目名>.py                # 入口薄壳模块（PYTHONPATH 模式必需）
├── src/                       # 业务包（PYTHONPATH 模式必需）
├── runtime/                   # 便携 Python（~120MB）
│   ├── python.exe
│   └── Lib/site-packages/     # 依赖 + embedded_code_extractor.py
├── install_release.bat/.ps1   # 一键安装
├── uninstall_release.bat/.ps1 # 一键卸载
├── smoke_test.py              # 5 项冒烟测试
├── reinject_ribbon_after_com_save.py  # COM 保存后 Ribbon 回注
├── rewrite_interpreter.py     # 安装时重写 Interpreter_Win
└── 安装说明.txt
```

**`__file__` 语义（ZIP 分发关键）**：
- 入口模块中 `os.path.dirname(os.path.abspath(__file__))` = 解压后的项目根目录
- 便携 Python 路径 = `os.path.join(PROJ_DIR, "runtime", "python.exe")`
- 配置表中 `Interpreter_Win` 必须指向便携 Python 的绝对路径（安装脚本自动修正）
- **禁止硬编码开发机路径**（如 `D:\Python\...`），分发后必失效

## 阶段三 代码嵌入——自研 code embed

- **任务**：把 Python 源码嵌入 xlam 隐藏 sheet，实现单文件分发。
- **要求**：复刻官方 PRO code embed 的嵌入/提取机制（附录 A.1 研读其原理），但移除许可证检查；与官方差异见下表。
- **输出物**：xlam（含 .py sheet + RELEASE_EMBED_CODE_MAP）+ embedded_code_extractor.py（随 runtime 分发到 site-packages）。
- **验收**：.py sheet 存在且 MAP 与文件一一对应；VBA 命令走 `embedded_code_extractor`；真机触发业务代码成功。

**核心机制**：复刻官方 xlwings PRO 的 code embed，但移除许可证检查。

**构建时**（embed_code.py）：
1. 删除旧 .py sheet
2. UUID 化 sheet 名（28 位十六进制 + .py）
3. A 列文本格式，代码按行写入
4. 单引号转义（`'` 开头双写，`'''` 替换为 `"""`）
5. 写入 `RELEASE_EMBED_CODE_MAP`（JSON：sheet名 → 相对路径）
6. 自动排除非业务目录（tests/、dist/、dist_release/、__pycache__/、.venv/、runtime/、site-packages/ 等）

**运行时**（embedded_code_extractor.py）：
1. VBA 检测到 .py sheet → `uses_embedded_code=True`
2. VBA 构建命令：`import embedded_code_extractor;embedded_code_extractor.runpython_embedded_code('command')`
3. Python 端 `dump_embedded_code()` 从 .py sheet 提取代码到 `%TEMP%\xlwings_embedded_xxx\`
4. `sys.path.insert(0, temp_dir)` 加入模块搜索路径
5. `exec(command)` 执行业务代码

**VBA 关键修改**（与官方的区别）：
- 移除 `LICENSE_KEY` 检查（官方为空直接 MsgBox 报错）
- 命令从 `import xlwings.pro;xlwings.pro.runpython_embedded_code(...)` 改为 `import embedded_code_extractor;embedded_code_extractor.runpython_embedded_code(...)`
- 官方导入 `xlwings.pro` 会触发许可证验证，自研版本不导入 pro 包

**embedded_code_extractor.py 位置**：随 runtime 分发到 `Lib/site-packages/`，VBA 调用 `import embedded_code_extractor` 时自动找到。

**自研 vs 官方 PRO 差异**（为何自研可行）：

| 环节 | 官方 PRO（附录 A 详研） | 自研 release_tool |
|------|----------------------|-------------------|
| 嵌入工具 | CLI `xlwings code embed`（`cli.py:code_embed()`） | `scripts/release_tool/embed_code.py` |
| 提取模块 | `xlwings.pro.embedded_code.dump_embedded_code()` | `embedded_code_extractor.py`（site-packages） |
| VBA 命令 | `import xlwings.pro;xlwings.pro.runpython_embedded_code(...)` | `import embedded_code_extractor;embedded_code_extractor.runpython_embedded_code(...)` |
| 许可证 | 无 LICENSE_KEY → MsgBox 报错并退出 | 已移除检查，不导入 pro 包 |
| 单文件分发 | 是 | 是 |
| 适用 | 需 PRO 许可证 | 无需许可证，本技能默认路线 |

## 阶段四 配置表整备

- **任务**：修复 xlwings.conf，让便携运行时真正生效。
- **要求**：逐键核对下方键表；处理 PROJECT_NAME / 激活机制 / 重复键三类高频坑。
- **输出物**：修复后的 `xlwings.conf`。
- **验收**：无重复键；Interpreter_Win 指向 runtime；PYTHONPATH 正斜杠；PROJECT_NAME="xlwings"；RELEASE_EMBED_CODE 与所选模式一致。

xlwings 加载项通过 `xlwings.conf`（PROJECT_NAME 常量决定前缀）配置 Python 解释器、模块路径等。配置表是分发后最容易出问题的环节。

**配置表键（权威见 `xlwings/docs/addin.md`）**：

| 键 | 作用 | 分发时注意 |
|----|------|-----------|
| `Interpreter_Win` | Python 解释器路径 | 指向 `runtime\python.exe`（安装脚本自动修正） |
| `PYTHONPATH` | 模块搜索路径 | **必须用正斜杠**（`C:/Users/...`），避免 `\U` 被 Python 解析为 unicode 转义 |
| `RELEASE_EMBED_CODE` | code embed 模式开关 | True=嵌入代码模式，False=PYTHONPATH 外部源码模式 |
| `RELEASE_EMBED_CODE_MAP` | .py sheet 名 → 文件路径映射 | code embed 模式必需，JSON 格式 |
| `RELEASE_HIDE_CONFIG_SHEET` | 隐藏配置表 | 分发时设 True |
| `RELEASE_HIDE_CODE_SHEETS` | 隐藏 .py sheet | code embed 模式分发时设 True |
| `RELEASE_NO_ADDIN` | 不依赖 xlwings.xlam 加载项 | 分发时设 False（白标 xlam 自带 VBA 模块） |
| `USE UDF SERVER` | UDF 服务器 | 本机用 False |
| `SHOW CONSOLE` | 是否显示控制台 | 分发时设 False（隐藏黑窗） |
| `LICENSE_KEY` | PRO 许可证 | 自研 code embed 模式不需要（VBA 已移除检查） |

**PROJECT_NAME 常量（高频坑）**：
- VBA 模块中 `Public Const PROJECT_NAME As String = "xlwings"`
- VBA 用 `PROJECT_NAME & ".conf"` 查找配置表（即 `xlwings.conf`）
- quickstart 默认 `PROJECT_NAME="myaddin"`，会导致查找 `myaddin.conf` 失败 → Interpreter_Win 回退到系统 `python` → 报错
- **必须修改为 `"xlwings"`**（安装脚本自动执行）

**配置表激活机制（高频坑）**：
- xlwings 读取的是 `xlwings.conf`（与 PROJECT_NAME 匹配，无下划线前缀）
- 开发模板中常生成 `_xlwings.conf`（带下划线，**未激活**）——必须重命名为 `xlwings.conf` 才生效
- **坑点**：`_xlwings.conf` 存在但 `xlwings.conf` 不存在时，xlwings 用默认配置（系统 Python），导致"明明配了便携运行时却不生效"

**重复键问题（门禁 K 捕获）**：
- 配置表中同一键出现多次时，xlwings 行为未定义（可能取第一个或最后一个）
- **门禁 K**（配置表检查）会扫描重复键并阻断
- 解决：用文本编辑器打开，搜索键名，保留一行

**安装脚本自动修复**：
- install_release.ps1 安装后自动修复配置表：PYTHONPATH 正斜杠、空值填充、RELEASE_EMBED_CODE 保持原值
- 自动执行 Ribbon 回注（COM 保存会破坏 customUI 注册，机制见附录 A.3）

## 阶段五 门禁与真机验证

- **任务**：构建/安装/验证各环节过门禁，并在无 Python 目标机真机验证。
- **要求**：门禁 E/J/K/L 逐项执行；真机验证在干净机器上做，不得跳过。
- **输出物**：门禁记录 + 冒烟结果 + 日志。
- **验收**：四门禁全过；目标机加载无"加载项损坏"；Ribbon 可见；点击触发 Python；业务端到端可用。

**门禁路由**：

| 门禁 | 触发时机 | 检查内容 | 失败处理 |
|------|---------|---------|---------|
| **门禁 E** | 6.5.10 交付物整备后 | xlam 包内 customUI14.xml 存在且 2010 命名空间、vbaProject.bin 存在 | 重新构建，检查 06-ribbon-guidance 注入结果 |
| **门禁 J** | 6.5.10 交付物整备后 | dist_release/ 内含 xlam + runtime/ + python.exe + 安装脚本；code embed 模式验证 .py sheet 存在 | 补全缺失文件 |
| **门禁 K** | 6.5.10 配置表 | xlwings.conf 无重复键、Interpreter_Win 指向有效路径、PYTHONPATH 用正斜杠 | 修正配置表 |
| **门禁 L** | 6.5.12 安装后集成验证 | 真机冒烟：Excel 加载 xlam → Ribbon 按钮可见 → 点击触发 Python → 结果正确 | 按真机验证 7 条实测结论排查 |

**目标机无 Python 的验证**（在**未安装 Python 的干净机器**上）：
1. 双击 install_release.bat 一键安装
2. 重启 Excel，加载项自动加载
3. 触发功能——若报"找不到 Python"或"找不到 xlwings"，说明 Interpreter_Win 或 PYTHONPATH 配置错误
4. 查看日志 `%LOCALAPPDATA%\<AddinName>\app.log` 定位具体错误

**常见失败**：
- `Interpreter_Win` 指向开发机路径 → 安装脚本应自动修正
- `PYTHONPATH` 用反斜杠 → `\U` 被 Python 解析为 unicode 转义，SyntaxError
- `PROJECT_NAME="myaddin"` → 找不到 xlwings.conf，回退到系统 Python
- 便携 Python 缺 `site-packages` → 重新构建 runtime
- WebView2 Runtime 缺失（pywebview 面板）→ 提示用户安装或改用 UserForm
- code embed 模式：RELEASE_EMBED_CODE_MAP 缺失或不匹配 → 重新运行 embed_code.py

## 阶段六 交付清单

- **任务**：交付前逐项核对。
- **要求**：清单 15 项全部勾选；不勾选不得交付。
- **输出物**：核对记录。
- **验收**：全部勾选；无开发机硬编码路径；无调试 print 残留。

- [ ] xlam 文件可在目标机 Excel 中加载（无"加载项损坏"提示）
- [ ] Ribbon 自定义选项卡/按钮可见，图标正常
- [ ] 点击按钮触发 Python 代码（无"找不到模块"错误）
- [ ] 业务功能端到端可用（查询/计算/写入/下载）
- [ ] UDF（如有）在单元格中可调用且返回正确
- [ ] 配置表 `xlwings.conf` 存在且 Interpreter_Win 指向便携 Python
- [ ] 便携运行时 `runtime/python.exe` 存在且依赖齐全
- [ ] code embed 模式：.py sheet 存在且 RELEASE_EMBED_CODE_MAP 正确
- [ ] PYTHONPATH 模式：src/ 和入口模块存在，PYTHONPATH 用正斜杠
- [ ] PROJECT_NAME 常量为 "xlwings"
- [ ] 日志目录 `%LOCALAPPDATA%\<AddinName>\` 可写
- [ ] 安装说明.txt 含安装/卸载/使用步骤
- [ ] 无开发机硬编码路径（`D:\`、`C:\Users\开发者名\` 等）
- [ ] 无调试 `print` 语句残留（入口层除外）
- [ ] 门禁 E/J/K/L 全部通过

## 附录 A PRO 部署机制对照研读（迁移自 13 号 1.4）

> **定位**：以下内容迁移自 `references/13-scenario-d-source-code.md` 1.4 节，源码级研读官方 PRO 部署机制（code embed / release）。
> **用途**：仅用于理解官方机制的设计（LicenseHandler 许可模型、Deploy Key、RELEASE_NO_ADDIN 独立运行、COM 保存坑），在阶段三/四的实施中**借鉴其设计、不采用其机制**——开发加载项部署走自研 release_tool（正文主线），不依赖 PRO 许可证。
> 官方文档：`xlwings/docs/pro/release.md`（PRO release 官方教程）、`docs/deployment.md`（部署总览）、`docs/addin.md`（配置表权威）。

### A.1 code embed 技术原理（官方源码研读）

`code embed`（`pro/embedded_code.py` + CLI `xlwings code embed`）将 Python 源码嵌入 Excel 工作簿的隐藏 sheet，实现「单文件分发」——用户只需一个 .xlsm/.xlam，无需附带 .py 文件。

**嵌入流程**（`cli.py:code_embed()`）：
1. **删除旧代码 sheet**：遍历工作簿所有 sheet，删除以 `.py` 结尾的 sheet（避免重复嵌入）
2. **UUID 化 sheet 名**：每个 .py 文件对应一个隐藏 sheet，sheet 名为 `uuid.uuid4().hex[:28] + ".py"`（28 位十六进制 + .py 后缀）。**不用原文件名作 sheet 名**——Excel sheet 名限 31 字符且禁止 `[]:*?/\`，UUID 规避所有限制
3. **单元格文本格式**：A 列设 `NumberFormat = "@"`（文本格式），列宽 65，代码按行写入 A1:A{n}
4. **单引号转义**：以 `'` 开头的行前补一个 `'`（Excel 会把前导单引号解释为「文本前缀标记」，需双写才能保留原文）；`'''` 三引号替换为 `"""`（避免 VBA 字符串解析冲突）
5. **写入 RELEASE_EMBED_CODE_MAP**：配置表中写入 JSON 映射 `{sheet名: 相对路径}`，如 `{"a1b2c3.py": "src/app.py"}`。此映射超 32767 字符时 `sys.exit("ERROR: The package structure is too complex to embed.")`——包结构过深会超限

**运行时提取流程**（`pro/embedded_code.py:dump_embedded_code()`）：
1. 从配置表读取 `RELEASE_EMBED_CODE_MAP`（JSON 解析为 sheet→路径映射）
2. 遍历所有以 `.py` 结尾的 sheet，读取 A 列值（`options(ndim=1)`），按行写入临时目录
3. `None` 行写为空行（`\n`），保留原始空行
4. `sys.path[0:0] = [target_dir]` 将临时目录插入模块搜索路径最前
5. `@lru_cache()` 缓存提取结果——同一工作簿只提取一次

**关键限制**：
- 单 sheet 单元格上限 32767 字符（Excel 单元格文本上限），超大型模块需拆分
- `RELEASE_EMBED_CODE_MAP` JSON 也受 32767 字符限制，包文件过多会超限
- 提取到 `%TEMP%` 临时目录，进程退出后不自动清理（多次运行会累积临时文件）

**VBA 端检测逻辑（`addin/Main.bas`，理解「为什么禁用 code embed」的关键）**：

VBA 的 `RunPython` 函数在构建 Python 命令前，先检测工作簿中是否存在 `.py` sheet：

```vba
' Check for embedded Python code
uses_embedded_code = False
For i = 1 To 2
    If i = 1 Then
        Set wb = ActiveWorkbook
    Else
        Set wb = ThisWorkbook
    End If
    For Each sht In wb.Worksheets
        If Right$(sht.Name, 3) = ".py" Then
            uses_embedded_code = True
            Exit For
        End If
    Next
Next i

If uses_embedded_code = True Then
    AddExcelDir = "false"
Else
    AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")
End If
```

**三个关键事实**：
1. **检测范围是 ActiveWorkbook + ThisWorkbook**：不仅检查当前加载项（ThisWorkbook），还检查用户打开的活动工作簿（ActiveWorkbook）。如果用户在任意工作簿中创建了名为 `xxx.py` 的 sheet，也会触发 embedded code 模式
2. **检测条件仅为 sheet 名后缀 `.py`**：不检查 `RELEASE_EMBED_CODE` 配置项，不检查 sheet 是否隐藏，不检查 sheet 是否有内容。只要 sheet 名以 `.py` 结尾就触发
3. **触发后走 `RunPythonEmbeddedCode` 分支**：该分支最终调用 Python 端 `xlwings.pro.embedded_code.runpython_embedded_code(command)`，而 `pro/embedded_code.py` 模块级调用 `LicenseHandler.validate_license("pro")`，无有效许可直接抛 `LicenseError`

**推论：配置表 `RELEASE_EMBED_CODE=False` 无法阻止许可证检查**。VBA 是通过检测 sheet 名判断的，不是通过读取配置项。因此，要彻底避免许可证检查，**必须确保工作簿中不存在任何以 `.py` 结尾的 sheet**，而不是仅设置配置项。

**VBA 端 LICENSE_KEY 检查（code embed 的核心障碍）**：

检测到 `.py` sheet 后，VBA `RunPython` 函数执行以下逻辑：

```vba
' PythonCommand with embedded code
If uses_embedded_code = True Then
    licenseKey = GetConfig("LICENSE_KEY")
    If licenseKey = "" Then
        MsgBox "Embedded code requires a valid LICENSE_KEY."
        Exit Function
    Else
        PythonCommand = "import xlwings.pro;xlwings.pro.runpython_embedded_code('" & SourcePythonCommand & "')"
    End If
End If
```

**官方机制的两个关键环节**：
1. **LICENSE_KEY 检查**：VBA 先检查配置表中的 `LICENSE_KEY`，为空直接 `MsgBox` 报错并 `Exit Function`
2. **命令构建**：VBA 构建的命令是 `import xlwings.pro;xlwings.pro.runpython_embedded_code('command')`，这会导入 `xlwings.pro` 包，触发 `pro/__init__.py` 中的许可证验证

**UDF 场景的特殊处理**：`GetUdfModules` 函数会自动把 `.py` sheet 名（去掉后缀）加入 UDF 模块列表，因此嵌入代码的 UDF 不需要在 `UDF Modules` 配置中显式声明。但修改代码后需要重新导入 UDF。

**对自研路线的启示**：官方检测仅看 sheet 名后缀，自研 `embedded_code_extractor` 复用同一嵌入格式（UUID sheet + MAP）；VBA 命令改为自研模块即避开 `xlwings.pro` 导入与许可证验证（阶段三已实施）。

### A.2 release 命令完整流程（官方源码研读）

`xlwings release`（`cli.py:release()`）在当前打开的工作簿上执行「一键发布」，**不区分 xlam 还是 xlsm**——用 `xw.apps.active.books.active` 获取当前活动工作簿。

**执行步骤**：
1. **创建 Deploy Key**：`LicenseHandler.create_deploy_key()` 生成部署密钥（试用密钥直接用 license_key）。Deploy Key 与开发者密钥不同——它绑定到发布的工作簿，目标机无需开发者许可
2. **写入 xlwings.conf 配置表**（首次 release 时交互询问）：
   - `Interpreter_Win`: `%LOCALAPPDATA%\{project_name}\python.exe`——便携运行时路径，`%LOCALAPPDATA%` 在运行时由 xlwings 展开
   - `Interpreter_Mac`: `$HOME/{project_name}/bin/python`
   - `PYTHONPATH`: None（嵌入代码模式不需要外部模块路径）
   - `LICENSE_KEY`: deploy_key
   - `RELEASE_EMBED_CODE`: True/False（是否嵌入代码）
   - `RELEASE_HIDE_CONFIG_SHEET`: True/False（隐藏配置表）
   - `RELEASE_HIDE_CODE_SHEETS`: True/False（隐藏代码 sheet）
   - `RELEASE_NO_ADDIN`: True/False（是否不需要 xlwings 加载项，独立运行模式）
   - `RELEASE_REMOTE_INTERPRETER`: True/False（是否支持 xlwings Server 远程解释器）
3. **RELEASE_NO_ADDIN 模式**（独立运行，目标机无需安装 xlwings.xlam）：
   - 移除 VBA 引用 `xlwings`（`VBProject.References.Remove`）
   - 移除 VBA 模块：`xlwings`、`Dictionary`、`IWebAuthenticator`、`WebClient`、`WebRequest`、`WebResponse`、`WebHelpers`
   - 导入独立运行模块：`xlwings.bas`（不含加载项依赖的 RunPython 实现）、`Dictionary.cls`
   - 远程解释器模式额外导入 `Remote.bas` 等 6 个模块
4. **嵌入代码**（RELEASE_EMBED_CODE=True 时）：调用 `code_embed(None)` 嵌入工作簿同目录下所有 .py 文件
5. **隐藏 sheet**：按配置隐藏 xlwings.conf 和 .py 代码 sheet
6. **版本兼容性检查**：用 `Interpreter_Win` 指向的解释器运行 `xlwings.__version__`，与工作簿中 VBA 模块版本比对

**关键设计决策**：
- release 是「在当前打开的工作簿上原地修改」，不是生成新文件——发布前应备份
- `%LOCALAPPDATA%\{project_name}\python.exe` 是约定路径，安装器负责把便携运行时解压到该路径
- RELEASE_NO_ADDIN 模式下，VBA 端的 RunPython 实现内嵌在 xlwings.bas 中，不依赖 xlwings.xlam 加载项

**对自研路线的启示**：release 的"原地修改 + 配置写入 + 隐藏 sheet + 版本比对"四步被自研 release_tool 拆解为构建脚本（embed_code.py / rewrite_interpreter.py / 门禁）——自研不依赖 Deploy Key 与 LICENSE_KEY，其余机制逐一对应。

### A.3 COM 保存与 Ribbon 回注（通用坑）

**「COM 保存会剥掉 customUI 注册」的含义**：
- 用 COM 自动化打开 .xlam/.xlsm 并执行 `wb.save()` 时，Excel 重新序列化 OOXML 包
- 在此过程中，`[Content_Types].xml` 中 customUI 的 `Override` 条目可能被 `Default` 条目覆盖，`_rels/.rels` 中 customUI 的 Relationship 可能丢失
- 结果：Ribbon 自定义消失，Excel 只显示默认 Ribbon
- 解决：用 zip 级操作重新注入 customUI（直接操作 OOXML 包，补回 Override 和 Relationship）

**对自研路线的启示**：`dist_release/reinject_ribbon_after_com_save.py` 即此机制的自研实现（门禁 E 验证注入结果）。

### A.4 部署方式对比（官方 PRO 视角 + 自研定位）

| 方式 | 许可 | 单文件 | UDF 支持 | 目标机 Python | 适用场景 |
|------|------|--------|---------|-------------|---------|
| ZIP 打包（xlam+src） | 免费 | 否 | 是 | 需要 | 团队内部、环境可控 |
| RunFrozenPython（xlam+exe） | 免费 | 否 | 否 | 不需要 | 简单宏、无 UDF |
| code embed（PRO） | PRO | 是 | 是 | 需要 | 单文件分发、目标机有 Python |
| release（PRO） | PRO | 是 | 是 | 不需要（嵌入运行时） | 对外分发、目标机无 Python |

**自研定位**：本技能开发的加载项对应"release（PRO）"的能力（单文件 + UDF + 目标机无 Python），但以自研 release_tool 实现（code embed 自研 + 便携运行时 + 一键安装），**无需 PRO 许可证**。官方 PRO 机制仅作对照研读。

## 附录 B 官方文档映射

| 官方文档（`xlwings/docs/`） | 内容 | 对应阶段 |
|------|------|---------|
| `addin.md` | 配置表全部键权威、PROJECT_NAME 语义 | 阶段四 |
| `deployment.md` | 官方部署选项（ZIP / RunFrozenPython / 代码嵌入 / 发布） | 阶段一 / 附录 A |
| `pro/release.md` | PRO release 官方教程（code embed / release / Deploy Key） | 附录 A |
| `pro/license_key.md` | PRO 许可证机制 | 附录 A |
| `troubleshooting.md` | 坑点沉淀（含配置表/便携运行时相关） | 阶段五 |

## 可配套阅读

- `references/13-scenario-d-source-code.md` 1.4（原 PRO 部署机制详研，已压缩为指针）
- `references/09-testing-debugging-guidance.md`（测试/调试/真机验证基础）
- `references/06-ribbon-guidance.md`（Ribbon 注入与回调签名）
- `xlwings/docs/addin.md`、`docs/deployment.md`、`docs/pro/release.md`
