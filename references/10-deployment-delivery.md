# 部署与交付指南（6.5.10/6.5.11/6.5.12/6.5.13 扩展）

> 本文件是 SKILL.md 6.5.10（部署与交付物整备）、6.5.11（安装）、6.5.12（安装后集成验证）、6.5.13（分发）的**扩展资源**——分发方式、便携运行时、配置表管理、门禁路由、交付清单。规则以 6.5.10/11/12/13 正文为准，本文件不重复。

## 一、分发方式选型

| 方式 | 适用场景 | 关键操作 | 目标机要求 |
|------|---------|---------|-----------|
| **XLSTART 自动加载** | 本机自用 / 固定机器 | xlam 放入 `%APPDATA%\Microsoft\Excel\XLSTART\`，Excel 启动自动加载 | 需 Python + xlwings |
| **手动加载项** | 分发他人 / 可选加载 | Excel → 文件 → 选项 → 加载项 → 浏览 → 选 xlam | 需 Python + xlwings |
| **release_tool - PYTHONPATH 模式** | 目标机无 Python，代码需易修改 | xlam + runtime/ + src/ + 入口模块，install_release.bat 一键安装 | 无需预装 Python |
| **release_tool - code embed 模式**（推荐） | 目标机无 Python，单文件分发，代码不外露 | xlam（含 .py 隐藏 sheet）+ runtime/，install_release.bat 一键安装 | 无需预装 Python |

**默认假设**：目标机未安装 Python——release_tool 便携运行时分发是标配，不是例外分支。

**code embed vs PYTHONPATH 模式对比**：

| 维度 | code embed | PYTHONPATH |
|------|-----------|------------|
| 代码存储 | .py 隐藏 sheet（UUID 命名） | 外部 src/ 目录 |
| 配置表 | RELEASE_EMBED_CODE=True + RELEASE_EMBED_CODE_MAP | RELEASE_EMBED_CODE=False |
| 运行时 | VBA 检测 .py sheet → embedded_code_extractor 提取到 %TEMP% → 执行 | 直接从 PYTHONPATH 目录 import |
| 单文件分发 | 是（代码在 xlam 内） | 否（需附带 src/） |
| 代码修改 | 需重新嵌入 | 直接修改 src/ 文件 |
| 许可证 | 无需 PRO（自研 embedded_code_extractor） | 无需 PRO |

## 二、便携运行时打包

**推荐方式：install_runtime.py install_minimal**（~120MB，uv 安装依赖）

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

## 三、配置表管理（高频坑集中区）

xlwings 加载项通过 `xlwings.conf`（PROJECT_NAME 常量决定前缀）配置 Python 解释器、模块路径等。配置表是分发后最容易出问题的环节。

**配置表键（权威见 `xlwings-0.37.2/docs/addin.md`）**：

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
- 自动执行 Ribbon 回注（COM 保存会破坏 customUI 注册）

## 四、code embed 模式详解（自研，无需 PRO 许可证）

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

## 五、门禁路由（部署与交付阶段）

| 门禁 | 触发时机 | 检查内容 | 失败处理 |
|------|---------|---------|---------|
| **门禁 E** | 6.5.10 交付物整备后 | xlam 包内 customUI14.xml 存在且 2010 命名空间、vbaProject.bin 存在 | 重新构建，检查 06-ribbon-guidance 注入结果 |
| **门禁 J** | 6.5.10 交付物整备后 | dist_release/ 内含 xlam + runtime/ + python.exe + 安装脚本；code embed 模式验证 .py sheet 存在 | 补全缺失文件 |
| **门禁 K** | 6.5.10 配置表 | xlwings.conf 无重复键、Interpreter_Win 指向有效路径、PYTHONPATH 用正斜杠 | 修正配置表 |
| **门禁 L** | 6.5.12 安装后集成验证 | 真机冒烟：Excel 加载 xlam → Ribbon 按钮可见 → 点击触发 Python → 结果正确 | 按真机验证 7 条实测结论排查 |

## 六、交付清单（交付前逐项核对）

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

## 七、目标机无 Python 的验证

在**未安装 Python 的干净机器**上验证：
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

## 可配套阅读

- `xlwings-0.37.2/docs/addin.md`（配置表全部键权威）
- `xlwings-0.37.2/docs/deployment.md`（官方部署选项）
- `docs/troubleshooting.md`（坑点沉淀，含配置表/便携运行时相关）


