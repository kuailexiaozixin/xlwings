---
name: xlwings
description: xlwings 全场景技能：自动化操作 Excel——Python 脚本与 MCP 两种方式（场景 A）、Web/服务化方案——xlwings Server / Lite / Office.js（场景 B）、开发 VBA 宏或 VBA 加载项（场景 C）、源码学习与二次开发（场景 D）。场景 C 提供 13 步完整工作流（需求→环境→设计→Python→VBA→Ribbon→UDF→配置→测试→交付→安装→集成验证→分发），含 12 个质量门禁、便携运行时分发、白标 xlam 构建。触发关键词：xlwings、Excel 自动化、Python 操作 Excel、Excel 加载项、xlam、UDF、RunPython、Ribbon、VBA 桥接、数据回填、面板交互、MCP、xlwings Server、xlwings Lite、Office.js。
---

# xlwings 全场景技能

## 1 技能定位与资源地图

本技能覆盖 xlwings 的四类使用场景，按「需求路由 → 场景执行 → 质量门禁 → 交付分发」组织。

**核心原则**：
- 场景 C（VBA 宏 / VBA 加载项开发）是本技能的主干，提供 13 步完整工作流与 12 个质量门禁
- 场景 A/B/D 为支线：A（自动化，含 MCP）、B（Web/服务化：Server/Lite/Office.js）、D（源码研读），提供主线步骤与资源导航
- 坑点优先沉淀 `docs/troubleshooting.md`，只有核心主干规则才进本 SKILL.md
- 禁止绝对路径；技能外发后所有路径基于技能根目录相对计算
- **references 文档为独立扩展参考，内容节点间不做内容级互引**（仅案例详展位置等路由指针例外）；SKILL.md 承载完成工作流所需的全部核心规则

**资源目录**：

| 目录 | 用途 | 加载时机 |
|------|------|---------|
| `references/` | 17 份文档（00 索引 + 16 篇扩展指南：01-10 工作流各阶段深度 + 11 场景A + 12 MCP 自动化 + 13 场景D + 14 Server + 15 Lite + 16 Office.js）。进入 references 前先读 `references/00-index.md` 决策索引 | 进入对应阶段需要深度扩展时按需读取 |
| `templates/` | 代码模板（VBA 模块、Ribbon XML、Python 桥接、UDF、面板入口、测试、安装脚本） | 编码阶段直接复制使用 |
| `scripts/` | 可执行脚本（门禁 12 个、构建脚本、配置激活、路由校验、release_tool 分发工具链） | 工作流各步门禁与构建时调用 |
| `docs/` | 故障排查、术语表、交付清单 | 遇到问题时查阅 |
| `dist/` | 示例交付物（myaddin.xlam + 安装脚本） | 参考交付文件夹结构 |
| `xlwings/` | xlwings 官方源码与文档（离线查阅）。目录名不含版本号，跟踪 `xlwings/xlwings` 的 main 分支，当前快照 commit 与最新正式版见 `manifest.json` 与 `SYNCLOG.md` | API 查证时读取 |
| `../VBA-Docs/` | 微软 VBA 官方文档（离线查阅） | VBA 编码时查证 |
| `examples/` | 官方示例与演示项目 | 学习与参考时读取 |

---

## 2 需求澄清与场景路由

### 2.1 三维判定模型

进入任何 xlwings 任务前，先按三个维度判定场景：

| 维度 | 判定问题 | 影响 |
|------|---------|------|
| ① 使用范围/载体 | 个人脚本 / 团队工具 / 对外分发产品？ | 决定产物形态（.py / .xlsm / .xlam） |
| ② 交互方式 | 命令行 / 工作表交互 / 弹窗窗体 / Ribbon 按钮？ | 决定是否需要 VBA、UserForm、面板 |
| ③ 目标机环境 | 开发机自用 / 团队内已装 Python / 对外无 Python？ | 决定分发路线（源码 / XLSTART / 便携运行时） |

### 2.2 场景路由决策树

```
用户需求
├─ 自动化操作 Excel（Python 脚本 / MCP）→ 场景 A
├─ Web / 服务化方案（Server / Lite / Office.js 加载项）→ 场景 B（指向 14/15/16 号文档）
├─ 开发 VBA 宏或 VBA 加载项 → 场景 C（主干）
└─ 源码研读 / 二次开发 / 教学演示 / 产物质检 → 场景 D
```

**场景 C 内部形态路由**（维度①+②+③组合）：

| 形态 | 维度① | 维度② | 维度③ | 产物 |
|------|-------|-------|-------|------|
| 纯 VBA 宏工作簿 | 个人/团队 | 工作表/按钮 | 任意 | .xlsm |
| VBA + RunPython 混合 | 团队 | 工作表/UserForm | 已装 Python | .xlsm |
| 白标 xlam 加载项 | 团队/对外 | Ribbon/面板 | 已装 Python | .xlam + XLSTART |
| 便携运行时分发 | 对外 | Ribbon/面板 | 无 Python | .xlsm/.xlam + runtime |

---

## 3 通用运行基础（所有场景共用）

### 3.1 本地源码注入（开发期）

使用本技能内置的 xlwings 源码而非 pip 安装版，做到与文档同一份代码、可直接下断点调试：

```python
import sys
sys.path.insert(0, r'<技能根目录>\xlwings')
import xlwings as xw
```

注入是否生效看 `xw.__file__`，应指向 `<技能根目录>\xlwings\xlwings\__init__.py`。
不要用 `xw.__version__` 判断：源码里该常量是构建期占位符，`xlwings/xlwings/__init__.py` 第 8 行写死
`"0.0.0"`，只有 pip 装出来的包才会显示真实版本号。快照对应哪个上游版本，查 `manifest.json`
的 `ref` 与 `pinned_sha`。

### 3.2 依赖与环境

- Windows 推荐：`scripts/ensure_uv_env.ps1` 创建 uv 虚拟环境
- 核心依赖：`xlwings==0.37.4`、`pywin32`（Windows COM）
- 面板依赖：`pywebview`、`python-fasthtml`、`uvicorn`
- 测试依赖：`pytest`、`pytest-cov`

### 3.3 代码化运行铁律

- **禁止手动操作 Excel**：所有构建、配置、测试动作必须通过脚本/COM 自动化完成
- **禁止单条 shell 命令批量造文件**：源码/配置逐文件 Write/Edit
- **控制台输出禁止 emoji**：用 `[OK]`/`[FAIL]`/`[INFO]`/`[WARN]`/`[ERROR]` 替代
- **PowerShell 脚本编码**：`.ps1`/`.ps1.tmpl` 必须 `UTF-8 with BOM + CRLF`

### 3.4 进程与清理收尾

- Excel 实例必须用 `with xw.App()` 上下文管理器或显式 `app.quit()` 退出
- **禁止裸 `taskkill /F /IM EXCEL.EXE`**：按映像名全局强杀会连用户正在工作的 Excel 一起终止，且反复强杀会触发 Excel 韧性机制（Resiliency）把加载项列入禁用列表（`HKCU\...\Excel\Resiliency\DisableItems`），下次启动加载项不再加载
- 终止残留进程必须**按 PID 精确操作**：`Get-Process EXCEL -ErrorAction SilentlyContinue | Select Id, SessionId, StartTime, MainWindowTitle` → 识别非用户会话/无窗口/本次自动化创建实例 → `Stop-Process -Id <PID> -Force`。诊断与清理脚本见 `docs/troubleshooting.md` 坑 12
- **跨会话僵尸进程**：其他登录会话残留的 EXCEL 无法被本会话终止（拒绝访问），且会污染 COM/ROT 导致自动化"拒绝访问"——COM 测试须容忍该事实，改用 `DispatchEx`/独立实例规避，勿把"杀不掉"当失败重试
- COM 对象必须彻底销毁，避免误杀用户正在工作的 Excel
- 构建/覆盖 xlam 前先终止残留 Excel 进程，避免 PermissionError

---

## 4 场景 A：使用 xlwings 自动化操作 Excel（Python 脚本 + MCP 两种方式）

**适用**：自动化操作 Excel 的两条通道——① 纯 Python 脚本读写 Excel（无需 VBA/加载项/UDF）；② MCP 自动化（见 4.4，将 Excel 操作能力暴露给 AI Agent）。

### 4.1 主线工作流

```
连接 Excel → 对象导航 → 读取数据 → 写入数据 → 格式化样式 → 图表/图片/形状/表格 → 清理收尾
```

1. **连接 Excel**：`xw.Book()` 打开或创建工作簿；新实例用 `with xw.App()` 管理生命周期；多实例用 `xw.apps.keys()` 获取 PID 后精确指定；`xw.books`/`xw.sheets` 为顶层快捷方式。连接方式与 OneDrive/SharePoint 云盘见 `xlwings/docs/connect_to_workbook.md`、`onedrive_sharepoint.md`
2. **对象导航**：`app → books → book → sheets → sheet → range/charts/shapes/pictures/tables`；反向导航 `rng.sheet → rng.sheet.book → rng.sheet.book.app`；Range 三种选择方式：A1 表示法（推荐 `sheet['A1']`）、1-based 元组、命名区域。完整 API 文档在 `xlwings/docs/api/`（按类一页一文件，共 40 篇，上游已不设总索引页），语法总览见 `xlwings/docs/syntax_overview.md`
3. **读取数据**：`.value` 读取，`.options()` 控制转换器，大数据用 `chunksize` 分块；数据结构与转换器见 `xlwings/docs/datastructures.md`、`converters.md`
4. **写入数据**：指定左上角自动填充；公式分 `.formula`（普通）、`.formula2`（dynamic array）、`.formula_array`（CSE 数组）三类；写入后需 `app.calculate()` 触发计算
5. **格式化样式**：`.number_format`/`.color`/`.font`/`.autofit()`；批注用 `.note`（创建须经 COM `api.AddComment`）；超链接用 `.add_hyperlink`；子字符串格式化用 `.characters`（macOS 不支持）
6. **图表图片形状表格**：`sheet.charts.add()`、`sheet.pictures.add()`（支持文件或 Matplotlib/Plotly 对象，Matplotlib 集成见 `xlwings/docs/matplotlib.md`）、形状通过 COM `sheet.shapes.api.AddShape()`（`sheet.shapes` 无 `add()` 方法）、`sheet.tables.add()`
7. **清理收尾**：保存退出，`tasklist | findstr EXCEL` 确认无残留进程。多线程/多进程注意事项见 `xlwings/docs/threading_and_multiprocessing.md`，Jupyter 交互见 `jupyternotebooks.md`

### 4.2 关键规则

- `chunksize` 读取返回 `list` 而非生成器，无需 for 循环逐块收集——其内部机制（分批 COM 调用后 extend 到同一 list）与大数据性能对比见 `references/11-scenario-a-python-automation.md` 第一节
- 多线程/多进程不传递 xlwings 对象，在子线程内重新连接——重连模式完整代码见 `references/11-scenario-a-python-automation.md` 第五节
- Jupyter 交互（`xw.view()`/`xw.load()`）仅限交互式工作，写脚本用标准 API
- 形状创建必须走 COM `AddShape(type, left, top, width, height)`，type: 1=矩形, 5=圆角矩形, 9=椭圆——COM 调用细节与类型常量表见 `references/11-scenario-a-python-automation.md` 第四节
- 批注创建必须先 `sheet['A1'].api.AddComment('内容')`，之后才能用 `.note.text` 读写——完整生命周期代码与常见错误见 `references/11-scenario-a-python-automation.md` 第三节
- 公式三类型（formula/formula2/formula_array）的区别与选择——溢出行为与版本兼容性见 `references/11-scenario-a-python-automation.md` 第二节

### 4.3 实战案例参考（examples/）

遇到具体任务时，优先参考内置案例的代码结构与实现模式：

| 任务类型 | 推荐案例 | 核心参考点 |
|---------|---------|-----------|
| 最小脚本+UDF | `xlwings-demo-master/basics/xwdemo.py` | 脚本骨架、Range 读写、@xw.sub 宏入口 |
| 大数据性能 | `xlwings-demo-master/performance/arrays.ipynb` | chunksize 实测、数组批量写入、COM 超时边界 |
| 图表生成 | `xlwings-demo-master/correlation/correlation.py` | DataFrame → 散点图、图表操作、图片导出 |
| 外部 API 接入 | `xlwings-eikon-master/` | 行情 API → DataFrame → Excel 完整链路，五种接入形态 |
| Excel/Web 双端 | `simulation-demo-master/` | 计算逻辑层与表现层分离，同一份代码驱动两端 |
| Power Query 管理 | `xl-pq-handler-master/src/xl_pq_handler/` | M 代码批量提取/注入、PQManager 编程接口 |
| 自动化测试 | `excel-automated-testing-master/` | pytest 驱动 xlsm 回归、COM 测试 fixture |
| 静态质检 | `static-excel-test-master/test_refs.py` | openpyxl 扫描坏引用，无需 Excel 实例 |
| UDF 完整工程 | `Excel_udf_itus-main/` | SQLite → @xw.func → 单元格公式端到端 |

依赖 xlwings PRO 的案例不属场景 A 实战参考，索引与研读见 `references/13-scenario-d-source-code.md` 四章案例总览与 PRO 案例研读。

官方 quickstart 迷你示例（`xlwings/examples/`：`database`/`mpl`/`simulation`/`udf`/`fibonacci` + `build_lite.py`，均社区版）的详细展开见 `references/11-scenario-a-python-automation.md` 七章（A 侧：database/mpl/simulation）与 `references/07-udf-guidance.md` 六章（C 侧：udf/fibonacci）。

`xlwings-demo-master` 社区版部分（basics/correlation/frozen/interactive/performance/simulation/testing）的文件级用途详展见 `references/11-scenario-a-python-automation.md` 七章 7.7（PRO 部分见 13-scenario-d 四章 PRO 案例研读）。

完整案例索引（11 个案例按工作流环节分类）见 `references/11-scenario-a-python-automation.md` 附录。

### 4.4 MCP 自动化（场景 A 的协议封装通道）

**适用**：需要将 Excel 操作能力暴露给 MCP 客户端（AI Agent、Claude Desktop、Claude Code、Roo Code 等），让 Agent 通过自然语言对话操作 Excel。

**工作流**：需求判定 → 环境准备 → 选型安装 → 客户端配置 → 安全加固 → 测试验证 → 故障排查

1. **需求判定**：如果只是 Python 脚本自动化，走 4.1 主线，避免 MCP 协议额外复杂度。需要 MCP 的典型特征：AI Agent 需通过对话操作 Excel、多个 Agent 共享 Excel 操作能力、需将 Excel 工具接入 MCP 生态。
2. **环境准备**：Windows + Microsoft Excel（4 个社区 Server 均通过 xlwings COM 驱动 Excel，不支持 headless）；Python 3.10+；非 Windows 仅 vohoailinh90 提供 MockOpenpyxlBackend 用于开发测试。
3. **选型安装**：根据需求特征选择社区 MCP Server（本技能 `MCP-Server/` 目录内置 4 个，各项目的依赖、入口命令、工具清单与源码模块逐一详解见 `references/12-scenario-a-mcp-automation.md` 第二节）：
   - **快速验证 / DRM 加密文件** → geniuskey/mcp-server-xlwings（v0.4.1，`uvx mcp-server-xlwings` 一键启动，支持 DRM 保护文件、读取用户当前选区）
   - **生产环境 / 功能最全** → hyunjae-labs/xlwings-mcp-server（v1.0.4，`pip install xlwings-mcp-server`，会话式架构 + 线程安全 + 图表支持，20 个源码模块）
   - **VBA 宏为核心** → prabodh-hydexcel/excel-mcp（`python server.py`，连接实时 Excel，支持宏执行与命名区域）
   - **宏安全 / 跨平台开发** → vohoailinh90/Excel_MCP_Server（`python server.py`，VBA 宏白名单 `EXCEL_MCP_ALLOWED_MACROS`，专用 COM 线程解决 STA，非 Windows 用 Mock backend 开发测试）
4. **客户端配置**：在 MCP 客户端配置文件中注册 Server（stdio 传输）。Claude Desktop 编辑 `%APPDATA%\Claude\claude_desktop_config.json`；Claude Code 用 `claude mcp add`；Roo Code 用 `.roo/mcp.json`。command 指向 Python 解释器或 `uvx`，args 指向入口模块或脚本。三种客户端的完整配置 JSON 模板见 `references/12-scenario-a-mcp-automation.md` 第五节。
5. **安全加固**：VBA 宏是任意代码执行，生产环境必须设置 `EXCEL_MCP_ALLOWED_MACROS` 白名单——白名单机制的源码分析（环境变量解析、set 查找、None 语义坑）与专用 COM 线程模型见 `references/12-scenario-a-mcp-automation.md` 第四节；只暴露任务所需的最小工具集；关键操作前提示用户确认。
6. **测试验证**：重启 MCP 客户端，确认 Server 已连接（`/mcp` 或连接状态）；在对话中调用基础工具（读取单元格、写入数据）验证链路；测试 VBA 宏调用（如使用）确认白名单生效。
7. **故障排查**：Server 未连接 → 检查 command/args 路径；"Application is busy" → COM STA 冲突，使用带专用 COM 线程的项目；Excel 进程残留 → 任务管理器结束 EXCEL.EXE；宏无反应 → 检查 VBA 项目是否解锁、白名单是否包含该宏。7 种常见故障的排查步骤见 `references/12-scenario-a-mcp-automation.md` 第六节。

**核心原则**：
1. **MCP Server 本质是场景 A 的协议封装**：内部调用 `xw.Book()`/`xw.App()`，场景 A 的所有规则（COM 单线程、进程清理、大块读写）均适用
2. **COM STA 冲突**：多工具并发调用会导致 "Application is busy"，需专用 COM 线程或每会话锁（hyunjae-labs/vohoailinh90 已内置解决）
3. **权限最小化**：只暴露任务所需的最小工具集，不暴露完整文件系统

---
## 5 场景 B：Web / 服务化 Excel 方案（xlwings Server / Lite / Office.js）

**适用**：需要 Excel 能力跑在**浏览器 / 服务器**而非本机 COM——多人在线协作、Excel on the web、企业级部署（SSO/RBAC）、目标机无 Python。三条 Web 通道的深度参考文档：

| 通道 | 定位 | 参考文档 | 关键资源 |
|------|------|---------|---------|
| xlwings Server | 服务端运行时（FastAPI + Socket.IO + Redis），Office.js 加载项后端 | `references/14-xlwings-server-guidance.md` | `xlwings-server/`（docs 61 篇 + xlwings_server 包 + deployment/nginx/scripts/tests） |
| xlwings Lite | 浏览器端运行时（Pyodide/WASM），免费自包含 | `references/15-xlwings-lite-guidance.md` | `xlwings-lite/`（官方文档 31 篇镜像）+ taxi-duckdb 案例 |
| Office.js 引擎 | 客户端引擎（值转换层 + UDF/脚本全链路），Server 与 Lite 共享的语义内核 | `references/16-xlwings-officejs.md` | `xlwings/xlwings/pro/_xlofficejs.py` + `udfs_officejs.py` |

**工作流**：

```
需求判定 → 选型（Server / Lite / 引擎研读）→ 按对应文档阶段推进 → 验证交付
```

1. **需求判定**：确认是否真的需要 Web/服务化路线。单机离线自动化走场景 A；桌面 VBA 生态走场景 C；需要浏览器端 / 服务器集中 / 多人在线 / SSO 才走本场景。
2. **选型**：
   - 需要自有服务器、企业认证、Python 集中 → **Server**（14 号文档，8 阶段工作流）
   - 需要免费、浏览器内自包含、无服务器 → **Lite**（15 号文档，7 阶段工作流）
   - 需要理解 Web 加载项底层机制、自定义函数/脚本协议 → **Office.js 引擎**（16 号文档，7 阶段工作流）
3. **推进**：按所选文档的阶段 → 步骤 → 输出物执行（14/15/16 号文档均为开发生命周期工作流编排，各阶段含任务内容 / 步骤 / 输出物）。
4. **验证**：三通道共用的 "加载项三件套" 检查（manifest / 端点 / 任务窗格）见 14 号文档 0 定位。

**核心原则**：
- Office.js 引擎是 Server 与 Lite 共享的语义内核（值转换层 + UDF/脚本管线）——先读 16 号理解机制，再进 14/15 应用；
- 本场景三条通道均不承载 runPython（VBA 桥接走场景 C）与 COM 桌面自动化（场景 A）；
- Web 通道的 "目标机" 是无状态的浏览器——一切状态在 Server（Redis）或浏览器（Pyodide）侧管理。
---

## 6 场景 C：使用 xlwings 开发 VBA 宏或 VBA 加载项（主干工作流）

### 6.1 场景能力与核心原则

场景 C 覆盖：VBA 宏、RunPython 桥接、UDF 自定义函数（桌面 UDF）、Ribbon 自定义、UserForm 窗体、桌面面板（pywebview+FastHTML）、白标 xlam 加载项、便携运行时分发。**专指桌面 VBA 通道**；Web/Office.js 加载项（Server/Lite）不属本场景，见场景 B（5 章）。

**架构约束（必须遵守）**：
1. **RunPython 不支持传参/返回值**：需参数时改用 UDF 或从单元格读写
2. **FastHTML 路由必须 `async def` + `await req.json()`**：禁止 `json.loads(req.body)`（req.body 是 coroutine）
3. **面板 uvicorn 必须独立 subprocess**：禁止 threading（GIL 阻塞）、禁止 multiprocessing（FastHTML app 不可 pickle）
4. **URL 一律 `127.0.0.1`**：禁止 `https://localhost`（IPv6 解析坑）
5. **目标机默认无 Python**：便携运行时是默认分发场景
6. **架构分层**：表示层（panel_html + app.py 路由）→ 业务层（api/downloader/cache）→ 数据层（外部 API/本地缓存），禁止面板路由直接调用外部 API
7. **修改面板代码后须重启面板**：面板进程是独立子进程，不热重载

### 6.2 13 步工作流总览

```
6.5.1 需求澄清与形态判定 → 6.5.2 环境准备 → 6.5.3 项目初始化与设计
→ 6.5.4 Python 代码构建 → 6.5.5 VBA 代码构建与引擎激活 → 6.5.6 Ribbon 构建
→ 6.5.7 UDF 开发与导入 → 6.5.8 插件配置与重命名 → 6.5.9 单元测试与静态检查
→ 6.5.10 部署与交付物整备 → 6.5.11 安装 → 6.5.12 安装后集成验证 → 6.5.13 分发
```

**门禁体系速查（12 个）**：

| 门禁 | 名称 | 脚本 | 执行时机 |
|------|------|------|---------|
| A | 构建完整性 | `gate_a_build_integrity.py` | 每次构建后 |
| B | 配置与分发 | `gate_b_config_and_dist.py` | 配置后 |
| C/D | 加载项注册 | `verify_addin_registered.ps1` | 安装后 |
| E | 交付前最终验证 | `gate_e_final_delivery.py` | 交付前 |
| F | RunPython 链路 | `gate_f_runpython_chain.py` | 安装后 |
| G | 引擎注入 | `gate_g_engine_injected.py` | 引擎激活后 |
| H | 语法检查 | `gate_h_syntax_check.py` | 每次改 .py 后 |
| I | 路径计算 | `gate_i_path_calculation.py` | ZIP 分发时 |
| J | 端到端运行时 | `gate_j_e2e_runtime.py` | 交付前 |
| K | 配置表重复键 | `gate_k_config_duplicate.py` | 配置后 |
| L | 真实 Excel 启动 | `gate_l_real_excel_launch.py` | 安装后 |

**0 级错误阻断进入下一步**。

### 6.5.1 需求澄清与形态判定

**目标**：明确功能需求、交互方式、目标机环境，判定产物形态与分发路线。

**必问问题（12 条）**：
1. 你现在是手工怎么做这件事的？
2. 最费时间的是哪一步？
3. 你希望最后看到什么结果？
4. 你希望结果是怎么产生的？点按钮运行（RunPython 宏），还是输入公式后自动计算（UDF）？
5. 如果原始数据变了，你希望结果自动跟着变吗？
6. 你的 Python 代码需要用到哪些第三方库？（pandas、numpy、requests 等）
7. 是改原表、生成新表、弹提示，还是放一个长期复用的入口？
8. 这是只给一个固定文件用，还是很多文件都要用？
9. 是你自己用，还是别人也要用？
10. 如果别人也要用，他们的电脑上有没有 Python 环境？
11. 你希望怎么触发，是打开文件自动跑，还是点按钮，还是写公式？
12. 你能不能给我样例表、截图、列名说明？

**必要截图要求**：带行号列标、尽量带公式栏、带工作表名称、框出当前处理区域、报错带完整错误弹窗、交互问题带按钮位置。

**操作**：
1. 按 2.1 三维判定模型确认维度①②③
2. 采集功能需求，按 MoSCoW 优先级排序：
   - **Must-have**：必须有的核心功能，第一轮实现
   - **Should-have**：重要但可以推迟，第二轮实现
   - **Could-have**：锦上添花，有时间再做
   - **Won't-have**：本次不做，明确告知用户
   - 产出：MVP 功能清单 = 所有 Must-have 功能
3. 判定交互载体：表单录入/简单弹窗 → VBA UserForm；列表+链接+富交互 → 桌面面板；无需窗口 → 纯工作表
4. 判定分发路线：团队内部/环境可控 → XLSTART 分发；对外/无 Python → 自研 release_tool 便携运行时分发
5. 需求复述确认后才能开工

**确认清单**（全部通过才进入下一步）：
- [ ] 工具用途已明确
- [ ] 输入数据格式已明确
- [ ] 输出结果格式已明确
- [ ] 用户群体已明确（自己用/给别人用；目标机是否有 Python）
- [ ] 触发方式已明确（Ribbon 按钮/表内按钮/公式/自动/面板）
- [ ] MVP 功能清单已确定（Must-have 项已确认）

**零基础用户沟通规范**：不要解释技术细节（装饰器、RunPython、配置表、pywebview）；用比喻说明（"我会给你做一个 Excel 里的按钮"）；只问业务问题；用户角色是提需求，AI 负责全部技术实现。需求采集话术与沟通示例见 `references/01-need-discovery.md`。

### 6.5.2 环境准备

**目标**：确保开发环境就绪（Python 解释器、依赖、COM 权限、WPS/Excel 信任中心）。

**操作**：
1. 创建虚拟环境：`scripts/ensure_uv_env.ps1`（Windows 推荐）
2. 安装依赖：`xlwings==0.37.4`、`pywin32`、面板/测试按需
3. 验证本地源码注入：`python -c "import sys; sys.path.insert(0,'xlwings'); import xlwings; print(xlwings.__file__)"`，输出的路径应落在技能根目录内（`__version__` 在源码里恒为 `0.0.0`，不能用它判断，见 3.1）
4. 配置 COM 权限：Excel 信任中心 → 启用「对 VBA 项目对象模型的访问」（AccessVBOM）——白标 xlam UDF 导入必需
5. 验证 COM 可用：`python -c "import xlwings as xw; app=xw.App(visible=False); print(app.api.Version); app.quit()"`

**环境就绪清单**（6 项全部通过才进入下一步）：
- [ ] Python 解释器可用（3.10+）
- [ ] xlwings 0.37.4 源码注入成功
- [ ] pywin32 安装，COM 可创建 Excel 实例
- [ ] AccessVBOM 已启用
- [ ] WPS/Excel 信任中心已配置（宏启用）
- [ ] 外部数据源可访问（如 API 需网络）

**PowerShell 编码规则**：`.ps1`/`.ps1.tmpl` 必须 `UTF-8 with BOM + CRLF`，否则中文乱码。环境预检快捷命令（dot-source + 6 个关键函数调用）、COM 权限配置步骤与 PowerShell 5.1 排查见 `references/02-encoding-com-prereqs.md`。

### 6.5.3 项目初始化与设计

**目标**：生成项目骨架，完成功能/界面/架构设计，产出设计六件套。

**操作 1——项目骨架生成**：

使用 xlwings CLI quickstart 生成骨架。注意：`python -m xlwings` 不可用（无 `__main__.py`），须用 `Scripts\xlwings.exe`：

```powershell
.\Scripts\xlwings.exe quickstart myproject --standalone
```

quickstart 已知问题：
- 生成的 `customUI.xml` 为 2006 版，构建时须升级为 `customUI14.xml`（2010 版）
- 回调默认调用 `module.main()`，若入口函数名不同需增加别名或修改 Ribbon VBA 模块
- 生成的自测入口指向 `.xlsm`，项目是 xlam 时需调整

**操作 2——设计六件套**：

| 产物 | 内容 |
|------|------|
| ① 功能清单 | MoSCoW 优先级、用户故事 |
| ② 数据设计 | 列 schema（字段名/类型/来源/必填/显示格式/写回位置），存在参考实现时必须与参考实现对齐 |
| ③ 界面设计 | Ribbon 按钮布局、面板/UserForm wireframe |
| ④ 架构设计 | 三层分层（表示/业务/数据）、模块划分、API 契约 |
| ⑤ 接口设计 | 外部 API 字段映射（集成前必须打印原始响应确认字段） |
| ⑥ 依赖清单 | `requirements.txt`（xlwings/pywin32/面板依赖/测试依赖），release_tool 分发必需产物 |
| ⑦ 决策记录 | 关键技术选型与取舍理由 |

**列 Schema 设计决策法**（设计②的核心方法）：

- **第一原则**：Schema 优先于当前显示值——先定义"每一列是什么类型、什么格式、什么来源"，再谈怎么写值
- **推荐最小 Schema 字段**：列名、列类型、数据来源、显示格式、写回位置（含起始行/列）
- **常见列类型与格式决策**：

| 列类型 | 写入策略 | 格式决策 |
|--------|---------|---------|
| 文本列（含前导零：股票代码/身份证） | 先设 `NumberFormat = "@"` 再写值 | 前导零归一化 |
| 日期列 | 写入日期值后补 `NumberFormat` | 显示格式与存储值分离 |
| 数值列 | 直接写数值，`Value2` 读取 | 长数字（15 位以上）先设文本再写 |
| 公式列 | 写公式字符串，配合 `.Formula`/`.Formula2` | ListObject 中禁止纯数据数组整块覆盖公式列 |

- **推荐决策顺序**：列类型 → 数据来源 → 显示格式 → 写回位置
- **二维数组写入前检查**：数组列数 == Schema 列数；文本列已是字符串；日期列已是 datetime；长数字列已是字符串
- **输出前自检**：每列类型已定、前导零列已设文本、长数字列已设文本、日期列写后补 NumberFormat、公式列不会被整块覆盖、存在参考实现时列 schema 与参考实现对齐（或以设计为准并在决策记录说明差异）

**操作 3——蓝图实例化**（需要完整场景骨架时）：
- UserForm 数据录入蓝图：`templates/project-blueprints/userform-data-entry-addin/`
- ListObject 工作台蓝图：`templates/project-blueprints/listobject-workbench-addin/`
- 实例化：`scripts/instantiate_blueprint.ps1`

**验证门禁 A**：`python scripts/gates/gate_a_build_integrity.py --xlam-path <产物>`

列 Schema 的深度设计方法（类型决策、格式分离、二维数组读写核验）与 ListObject 工作台设计见 `references/03-project-design-guidance.md`。

### 6.5.4 Python 代码构建

**目标**：编写 Python 业务代码（RunPython 入口、UDF、面板侧、业务逻辑）。

**操作**：
1. **模块设计**：按三层分层组织——`src/app.py`（表示层路由）、`src/api/`（业务层）、`src/cache/`（数据层）
2. **RunPython 入口**：无参数无返回值，通过 `xw.Book.caller()` 读写单元格。模板见 `templates/internal-bases/excel-addin-core/src/python/runpython_bridge.py.tmpl`
3. **UDF 函数**：`@xw.func` 装饰器，推荐 `Annotated` 类型提示写法。模板见 `templates/internal-bases/excel-addin-core/src/python/udf_example.py.tmpl`
4. **面板入口**：`open_panel()` 启动 uvicorn subprocess + pywebview。模板见 `templates/internal-bases/excel-addin-core/src/python/panel_entry.py.tmpl`；完整面板骨架（app.py/main.py/HTML）见 `references/08-form-guidance.md`「面板侧完整模板」节
5. **业务逻辑**：纯函数，可被 UDF 和 RunPython 复用；外部 API 调用前打印原始响应确认字段
6. **自测入口**：每个模块底部 `if __name__ == '__main__'` 提供无需 Excel 的自测路径

**关键规则**：
- RunPython 不支持传参/返回值，参数通过单元格读写
- `sys.path` 插入必须在所有 `src` 包 import 之前（直接运行子进程入口脚本时）
- 面板侧代码修改后须关闭面板窗口并重新从 Ribbon 启动（不热重载）
- 禁止面板路由直接调用外部 API，必须经业务层
- 转换器与数据结构：pandas DataFrame 进出、numpy 数组、二维数组语义（`ndim=2`）
- **数据源读取**：UDF/RunPython 读**调用方工作簿**必须 `xw.Book.caller()` COM（活会话）
- **面板进程健康检查（必须实现）**：`open_panel()` 启动子进程后，须在 3 秒内 HTTP 探测 `http://127.0.0.1:<port>/`，失败则弹窗提示"面板启动失败，请检查 Python 环境"，避免用户看到空白窗口或"按钮无响应"假象。探测成功后再显示 pywebview 窗口
- **pywebview 必须强制 EdgeChromium 后端（高频坑）**：`webview.start(gui="edgechromium")`。MSHTML（IE 内核）不支持 `fetch`/`Promise`，会导致所有 API 调用静默失败，用户看到"按钮无响应"。启动失败时须明确提示安装 WebView2 Runtime。安装脚本须检测 WebView2 Runtime 并告警

**窗体三形态选型**（承接 6.5.1 维度③）：

| 形态 | 适用场景 | 技术栈 |
|------|---------|--------|
| VBA UserForm | 表单录入/简单弹窗 | VBA 管窗体与对象模型，Python 管复杂计算/网络请求 |
| pywebview 桌面面板 | 列表+链接+局部刷新的富交互 | pywebview + FastHTML 独立进程 |
| tkinter | 简单跨平台弹窗 | Python 原生，无需 WebView2 |

**验证门禁 H**：`python scripts/gates/gate_h_syntax_check.py --module <file>`（每次修改 .py 后立即执行）

转换器与数据结构实战（pandas DataFrame 进出、numpy 数组、二维数组语义 `ndim=2`、自定义转换器）见 `references/04-python-guidance.md`；面板装配规范与三形态（UserForm/pywebview/tkinter）对比见 `references/08-form-guidance.md`。

### 6.5.5 VBA 代码构建与引擎激活

**目标**：编写 VBA 侧桥接代码，构建 .xlsm/.xlam 产物，激活引擎与配置表。

**VBA 编码总纲（最高优先级 9 条）**：
1. **读取区域优先一次读入数组**：`arr = Range.Value2`，内存中处理
2. **写回区域优先整块回写**：`Range.Value2 = arr`，禁止逐格
3. **大批量操作前统一管理 Excel 应用状态**：`Application.ScreenUpdating/Calculation/EnableEvents` 关闭并在 finally 恢复
4. **禁止默认使用 `Select`/`Activate`**：直接 `ws.Range(...)` 操作
5. **变量类型默认更严格**：`Dim` 显式类型、`Option Explicit`
6. **能复用边界值就不要重复计算**：`lastRow`/`lastCol` 算一次
7. **查找/去重/计数优先字典**：`Dictionary`/`Collection` 而非多重循环
8. **长任务给用户反馈，但不在循环里疯狂弹窗**：状态栏或一次性 MsgBox
9. **UDF 只做计算，不做副作用**：不写单元格、不弹窗、不操作文件

**落码前自检问题**（逐条过）：
- [ ] 区域读写是否用了数组整块（非逐格）？
- [ ] 大批量操作前是否关闭了 ScreenUpdating/Calculation/EnableEvents 并在出口恢复？
- [ ] 是否避免了 Select/Activate/ActiveSheet/Selection？
- [ ] 边界值（lastRow/lastCol）是否只算一次？
- [ ] 日期/长数字列是否补了 NumberFormat 或先设文本？
- [ ] Ribbon 回调是否只做分发、业务逻辑在独立过程？
- [ ] 错误处理是否用了统一出口而非 Resume Next/End？

**一句话决策口诀**：能整块不逐格、能显式不 Active、能字典不循环、能统一出口不硬退。

**VBA API 查证流程**（写代码前必做，禁止凭记忆写 API）：
1. 用索引库定位：`../vba_docs_index.db`（表 `api_ref(app,kind,name,title,rel_path,description)`），`kind` 包括 method/property/event/enumeration/object
2. 打开参考页取权威语法：`../VBA-Docs/api/` 下对应页（Syntax/Remarks/Example）
3. 校准语言基础：`../VBA-Docs/Language/`（Reference/Concepts/How-to）

**结构设计规则**：
- 推荐默认业务骨架：入口（Ribbon 回调/按钮）→ 业务过程 → 辅助函数，分层清晰
- **Ribbon 回调只做分发**：回调过程内只调业务过程，不写业务逻辑；回调与业务过程必须分离命名
- **加载项（xlam）中 `ThisWorkbook` 指加载项自身**，用户数据在 `ActiveWorkbook`——操作用户数据用 `ActiveWorkbook.Worksheets(...)` 或通过参数传入工作簿引用
- 普通宏工作簿（xlsm）中 `ThisWorkbook` 即用户工作簿

**操作**：
1. 离线查阅 VBA 文档：`../VBA-Docs/`（与 xlwings 平级），对象模型按类逐一落文件，枚举与常量见 `../VBA-Docs/api/Excel(enumerations).md`
2. 编写 VBA 代码：按上述编码总纲执行，区域数据处理默认「读入数组→内存计算→整块写回」
3. 激活配置表并注入引擎：推荐用一键脚本 `python scripts/activate_addin_config.py --xlam-path <xlam> --interpreter <python.exe> --pythonpath <proj-dir>`，自动完成 IsAddin 切换、配置表改名、写 Interpreter_Win/PYTHONPATH/UDF Modules
4. 构建产物：`scripts/build_addin.ps1 -QuickstartProjectPath <工程>` 注入 VBA + Ribbon；或 `scripts/build_xlwings_addin.ps1` 一键构建

**UserForm 数据过桥规则**（表单与 Python 的数据边界）：
- 表单输入 → 写入工作簿数据区（隐藏行/列或专用 sheet）→ `RunPython` 触发 → Python 经 `Book.caller()` 读取数据区
- Python 结果写回工作簿后，VBA 读回填充表单列表或直接留表呈现
- **禁止**在 VBA 事件中直接调用 `RunPython` 并期望返回值（RunPython 无返回值通道），也禁止用字符串拼接向 RunPython 传参
- UserForm 是交互层，不是数据通道

**关键顺序约束（2010 版 customUI14.xml 适用）**：
- 适用条件：本约束针对 2010 版 `customUI14.xml` 且在 `[Content_Types].xml` 中有单独 Override 的情形。若 Ribbon 为 2006 版且仅靠 Default xml 覆盖（无 Override），`wb.save()` 通常不会删除 `_rels` 中的 Relationship
- 引擎激活的 `wb.save()` 会覆盖 `[Content_Types].xml` 中 customUI 的 Override 注册
- 正确流程：`quickstart → build_addin.ps1（注入VBA+Ribbon）→ 引擎激活（wb.save()）→ 最后 zip 级 Ribbon 回注`
- 推荐构建时加 `-ReinjectRibbonAfterActivation` 开关自动兜底
- 验证：跑门禁 A 检查 Content_Types 含 customUI 且 _rels 恰好 1 条

**验证门禁**：A（构建完整性）、G（引擎注入）、H（语法检查）

VBA 对象模型速查、单元格引用 7 种方法、现有文件迭代链路与反模式清单见 `references/05-vba-guidance.md`。

### 6.5.6 Ribbon 构建

**目标**：构建 Ribbon XML 与回调，注入产物包。

**操作**：
1. **规划 Ribbon 结构**：需要常驻入口、分组/图标/动态状态时才用 Ribbon；纯表内按钮足够时不必用
2. **版本规范**：一律 2010 版——`customUI14.xml`，命名空间 `http://schemas.microsoft.com/office/2009/07/customui`。2006 版仅作导入兼容
3. **核对回调签名**：`onAction` 指向的 VBA 过程参数必须与控件签名完全匹配。签名不匹配时 Ribbon 静默不响应，无报错
4. **Ribbon XML 必须 UTF-8 编码**（中文标签乱码的根因）
5. **修改后重新构建**：通过 `build_addin.ps1` 回注，再跑门禁 A

**回调签名速查表**：

| 控件 | 回调属性 | 签名 |
|------|---------|------|
| `button` | `onAction` | `Sub OnAction(control As IRibbonControl)` |
| `button` | `getLabel`/`getEnabled`/`getVisible`/`getImage` | `Sub GetXxx(control As IRibbonControl, ByRef xxx)` |
| `checkBox`/`toggleButton` | `onAction` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` |
| `comboBox`/`editBox` | `onChange` | `Sub OnChange(control As IRibbonControl, text As String)` |
| `dropDown`/`gallery` | `onAction` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` |
| `dynamicMenu` | `getContent` | `Sub GetContent(control As IRibbonControl, ByRef ribbonXml)` |
| `customUI` | `onLoad` | `Sub OnLoad(ribbon As IRibbonUI)` |

**回调签名关键规则**：
- 只有被 Ribbon XML 属性直接绑定的过程才需要回调签名；普通宏/工具方法不需要 `IRibbonControl` 参数
- 回调与业务过程必须分离命名（`onDownload_Click` 回调 → `DownloadData` 业务），不得同名
- `ByRef` 返回型回调必须在过程内给参数赋值
- `getImage` 返回 `IPictureDisp`
- `onLoad` 需保存 `ribbon` 对象供 `Invalidate`/`InvalidateControl` 刷新

**Ribbon 图标原则**：
- 默认优先使用 Office `imageMso` 内置图标（免资源文件）
- 用户提供品牌 logo 时：原始文件放 `./assets/`，缩小为 16×16/32×32 后放项目 `src/ribbon/`，构建时嵌入
- 自定义图标用 `getImage` 返回 `IPictureDisp`，建议 PNG 格式
- 蓝图项目优先在 `build_addin.ps1` 中传入 `-FirstIconPath`/`-SecondIconPath`

**普通用户入口优先级**：Ribbon 按钮（全局可见）> 工作表内按钮（随文件走）> 自动事件（慎用，打扰用户）

**验证门禁 A**：构建后执行（检查 ZIP 结构/Content_Types/_rels/customUI 版本/回调核验）

Ribbon 官方文档导航与注入结果验证方法见 `references/06-ribbon-guidance.md`；Ribbon XML 模板见 `templates/internal-bases/excel-addin-core/ribbon/customUI14.xml.tmpl`。

### 6.5.7 UDF 开发与导入

**目标**：按规范开发 UDF，注册测试并导入目标产物。UDF 仅 Windows 支持，必须承载于 .xlsm 或 .xlam。

**官方示例**：`xlwings/examples/udf/udf.py`（`@xw.sub`/`@xw.func` 最小集）与 `fibonacci/fibonacci.py`（含 `build_standalone.py` PyInstaller 冻结 → zip 便携分发）——可对照模板与官方 quickstart 实现。

**UDF 编写规范**：
- 使用 `@xw.func` 装饰器标记函数，导入后可在 Excel 公式中使用
- 推荐类型提示写法（v0.32.0+）：`def hello(name: str) -> str:`
- 二维数组用 `Annotated[list[list[float]], {"ndim": 2}]` 替代 `@xw.arg(..., ndim=2)`
- 支持数组公式（`ndim=2`）、异步执行（`async_mode='threading'`）
- `doc` 参数用于函数向导显示；保留参数 `caller` 与 `vba`

**操作**：
1. 编写 UDF：模板见 `templates/internal-bases/excel-addin-core/src/python/udf_example.py.tmpl`
2. 开发期注册验证：
   - xlsm + 标准加载项：VBA 中运行 `import_udfs "module_name"`
   - 白标 xlam：xlam 的 xlwings 模块中运行 `ImportPythonUDFsToAddin`
   - 用 COM 自动化注册并验证：单元格输入公式 → `app.api.Calculate()` → 断言结果（注意 `wb` 无 `calculate()`，须用 `app.api.Calculate()`）
3. 导入白标 xlam：COM 调用 `ImportPythonUDFsToAddin`，前提是 AccessVBOM 已启用
4. 导入 .xlsm：UDF 以模块形式随构建脚本写入，无需额外步骤

**行为规则**：
- UDF 代码变更自动生效；函数签名变更需重新 `import_udfs()`
- 导入模块变更不自动生效，需点击 xlwings Ribbon 的 Restart UDF Server 或重启 Excel（CLI 无 udf 子命令）
- 调试：源码末尾 `xw.serve()` 启动调试服务器，VBA 模块设 `UDF_DEBUG_SERVER = True`

**持久化警告**：UDF 必须承载于 .xlsm/.xlam。保存为 .xlsx 会移除 VBA 模块，UDF 全部失效（显示 #NAME?）。RunPython 宏不受此限制。

`@xw.func` 全部参数详解、官方 udfs 文档导航与 import 要点见 `references/07-udf-guidance.md`。

### 6.5.8 插件配置与重命名

**目标**：白标 xlam 完成最终重命名与配置验证。配置表写入与激活已在 6.5.5 完成，本步只做重命名和最终验证。

**配置表管理（高频坑集中区）**：

配置表键名与语义：

| 键 | 作用 | 分发时注意 |
|----|------|-----------|
| `Interpreter_Win` | Python 解释器路径 | XLSTART 分发构建时写入；便携运行时指向 runtime\python.exe |
| `PYTHONPATH` | 模块搜索路径 | XLSTART 分发由安装脚本动态写入（指向交付文件夹根目录）；多个路径用 `;` 分隔 |
| `UDF Modules` | UDF 模块名 | 逗号分隔 |
| `UDF Server` | UDF 服务器地址 | 本机用默认 `127.0.0.1` |
| `Show Console` | 是否显示控制台 | 分发时设 `False` |
| `Log File` | 日志路径 | 指向 `%LOCALAPPDATA%\<项目名>\app.log` |

**配置表激活机制**：
- xlwings 按 VBA 模块中 `PROJECT_NAME` 常量查找配置表（`PROJECT_NAME & ".conf"`），**不是按 xlam 文件名查找**
- 开发模板中常生成 `_myaddin.conf`（带下划线，未激活）——必须重命名为 `myaddin.conf` 才生效
- `_myaddin.conf` 存在但 `myaddin.conf` 不存在时，xlwings 用默认配置（系统 Python），导致"明明配了便携运行时却不生效"
- 验证：Excel 加载后在 xlwings 功能区看 Interpreter 显示的路径

**重复键问题**：配置表中同一键出现多次时，xlwings 行为未定义。门禁 K 会扫描重复键并阻断。常见原因：手动编辑配置表时复制粘贴未删旧行。

**操作**：
1. **重命名 xlam 文件**：关闭 Excel 后文件系统重命名。RunPython 模块名自动从新文件名推导，但配置表名不会自动跟随
2. **重命名 Ribbon VBA 模块**（推荐）：`RibbonMyAddin` → `Ribbon<项目名>`，须同步修改 `customUI14.xml` 中所有 `onAction` 属性
3. **重命名 VBA 项目名**（可选）：VBA 编辑器 → 工具 → VBAProject 属性 → 工程名称
4. **修改 PROJECT_NAME 常量**（必须）：xlwings VBA 模块中 `Public Const PROJECT_NAME As String = "myaddin"` 改为项目名。此常量决定配置表名和用户级配置文件路径。不改会导致多个 xlwings 加载项共享同一配置文件路径
5. **配置表最终验证**：确认 `<项目名>.conf`（无下划线前缀）已激活，`Interpreter_Win`/`PYTHONPATH`/`UDF Modules` 三键非空
6. **Ribbon 回注**：若经 COM 修改并 `wb.save()`，须重新运行 `python scripts/tools/reinject_ribbon_after_com_save.py --xlam-path <产物>`

**验证门禁**：
- B：`python scripts/gates/gate_b_config_and_dist.py --xlam-path <产物> --proj-dir <Python目录> --module <Python入口模块名>`（`--module` 为 Python 入口模块名，用于验证 PYTHONPATH 下模块可导入）
- K：`python scripts/gates/gate_k_config_duplicate.py --xlam-path <产物>`（A 列重复键扫描，阻止 VBA 457）

### 6.5.9 单元测试与静态检查

**目标**：在不安装加载项的情况下，对 Python 源码进行单元测试和静态质量检查，确保业务逻辑正确、代码质量达标。

**测试纪律（Prove-It，必须遵守）**：
- 新业务模块必须 test-first（Red-Green-Refactor）：先写失败测试→实现→重构
- 修 bug 必须先写复现测试：先写能复现 bug 的测试（Red）→修复（Green）→跑全量回归。禁止「先修复后补测」或注释掉失败测试绕过
- 失败测试处理：失败→单独运行（`pytest -k <name>`）复现→读 traceback 定位→修复→全量回归。失败测试不得注释绕过，必须修复或标记 `@pytest.mark.skip(reason=...)` 并记录 issue

**操作**：
1. **纯 Python 单元测试**：业务逻辑、数据转换、API 调用（用 mock 不依赖真实网络）、错误处理。测试数据与代码分离（`tests/fixtures/`）
2. **面板 API 集成测试**（含面板的项目必须执行）：独立启动 `src/server.py`→用 `requests` 调用 `/api/query`、`/api/download`、`/api/fill` 等接口→断言返回值结构与业务正确性。测试模板见 `templates/internal-bases/excel-addin-core/tests/test_panel_api.py.tmpl`。此测试不依赖 Excel，可快速验证面板侧 HTTP 链路
3. **前端路由链路校验**（含面板的项目必须执行）：`python scripts/check_routes_linkage.py --panel-html src\panel_html.py --app src\app.py`，exit 1 = 阻断发布
4. **覆盖率检查**：`python -m pytest --cov=src --cov-report=term-missing`，Python 业务模块覆盖率基线 ≥ 80%；核心业务逻辑（数据转换、API 调用、错误处理）必须覆盖
5. **执行汇总**：`python -m pytest tests/unit/ tests/panel_api/ -v`

**测试运行策略**：
- 纯 Python 单元测试可并行（`pytest -n auto`，需 pytest-xdist）
- 增量开发跑 `pytest -k "not serial"`（纯 Python 快速反馈），提交前跑全量
- 外部 API 测试用 mock（responses 库或 unittest.mock），不依赖真实网络

**UserForm 业务方法测试**：窗体逻辑解耦为可单测的业务方法，在本步完成第一层测试；窗体打开冒烟和 UI 自动化在 6.5.12 安装后验证。

**本步规则——示例文件**：测试通过后的冒烟产物保留为交付示例文件，供 6.5.10 整备使用。

conftest.py fixture 定义、UDF 双路径测试（单元格公式断言 + 纯 Python 逻辑）与真实机器验证 7 条铁律见 `references/09-testing-debugging-guidance.md`。
### 6.5.10 部署与交付物整备

**目标**：整备交付物（产物、安装/卸载脚本、示例文件、安装说明），确保开箱即用。安装脚本在此步生成，6.5.11 用脚本执行安装，6.5.12 安装后集成验证，6.5.13 打包分发。

**前置规则——xlam 重新构建前须源码验证**：修改 Python 代码后不立即重构建 xlam。先用源码模式验证改动（直接运行 `src\main.py` 或 `python -c "import <模块>; <模块>.<入口函数>()"`），确认功能正常后才重新构建。

**操作 1——分发路线选型**：

| 维度 | XLSTART 分发 | 自研 release_tool 分发 |
|------|-------------|----------------------|
| 产物形态 | .xlam | .xlsm 或 .xlam |
| Python 运行时 | 目标机需已安装 Python（或便携运行时随包） | 交付包自带便携 Python 运行时 |
| 安装方式 | xlam 复制到 XLSTART，Excel 启动自动加载 | install_release.bat 一键安装 |
| 适用场景 | 团队内部、环境可控 | 对外分发、目标机无 Python |
| 配置表 | PYTHONPATH 指向模块目录 | Interpreter_Win 指向便携运行时 python.exe |

**操作 2——XLSTART 分发：整备交付文件夹**：

```
交付文件夹/
├── <项目名>.xlam              # 加载项产物
├── <项目名>.py                # 入口薄壳模块（必须包含）
├── src/                       # Python 业务包
├── install.ps1                # 一键安装脚本（从最新模板生成）
├── uninstall.ps1              # 一键卸载脚本（从最新模板生成）
├── 示例.xlsx                  # 示例文件（从 6.5.9 冒烟产物获取，必须包含）
└── 安装说明.txt               # 安装/卸载/使用说明
```

**交付物整备铁律**：
1. 安装/卸载脚本必须从最新模板重新生成：`templates/internal-bases/excel-addin-core/install/install_xlstart.ps1.tmpl`（卸载用 `uninstall_xlstart.ps1.tmpl`）。不得复用旧版本——模板更新后须重新整备
2. 入口模块必须包含：`<项目名>.py` 是 RunPython 调用的入口薄壳（含 `main()`），PYTHONPATH 指向交付文件夹根目录
3. 示例文件必须包含：从 6.5.9 冒烟产物获取，无示例文件不得交付
4. PYTHONPATH 由安装脚本动态写入：构建时只写 `Interpreter_Win`，不写 `PYTHONPATH`（或写占位符）；安装脚本在复制 xlam 到 XLSTART 后通过 COM 写入，值为交付文件夹根目录的绝对路径

安装脚本核心逻辑（全自动，禁止手动复制）：WebView2 检测→获取 XLSTART→清理旧注册→复制 xlam→动态写入 PYTHONPATH→Ribbon 回注→输出结果。

**操作 3——自研 release_tool 分发（便携运行时 + 一键安装）**：

自研 release_tool 分发路线提供两种代码分发模式，均无需 xlwings PRO 许可证：

| 模式 | 代码存储 | 配置表 | 适用场景 |
|------|---------|--------|---------|
| **code embed**（推荐） | 嵌入 .py 隐藏 sheet | RELEASE_EMBED_CODE=True | 单文件分发、代码不外露 |
| **PYTHONPATH** | 外部 src/ 目录 | RELEASE_EMBED_CODE=False | 调试方便、代码易修改 |

**code embed 模式核心机制**（自研，复刻官方 PRO 但移除许可证检查）：
1. 构建时：`embed_code.py` 将 Python 源码嵌入 xlam 的隐藏 sheet（UUID 命名 + RELEASE_EMBED_CODE_MAP 映射）
2. 运行时：VBA 检测到 .py sheet → 调用 `embedded_code_extractor.runpython_embedded_code()` → 提取代码到 `%TEMP%` → 执行
3. 关键修改：VBA 模块移除 LICENSE_KEY 检查，命令改为 `import embedded_code_extractor;...`（而非官方的 `import xlwings.pro;...`）
4. `embedded_code_extractor.py` 随 runtime 分发到 `Lib/site-packages/`

**整备交付文件夹**（默认路径 `dist_release/`，与 XLSTART 分发的 `dist/` 区分）：

```
dist_release/
├── <项目名>.xlam              # 加载项（含 .py sheet 或配置表指向外部源码）
├── runtime/                   # 便携 Python 运行时（~120MB，含 embedded_code_extractor.py）
├── <项目名>.py                # 入口薄壳模块（PYTHONPATH 模式必需，code embed 模式可选）
├── src/                       # Python 业务包（PYTHONPATH 模式必需，code embed 模式可选）
├── install_release.bat/.ps1   # 一键安装（自动 Ribbon 回注 + 配置表修复）
├── uninstall_release.bat/.ps1 # 一键卸载（自动检测并终止占用进程）
├── smoke_test.py              # 5 项冒烟测试
├── reinject_ribbon_after_com_save.py  # COM 保存后 Ribbon 回注
├── rewrite_interpreter.py     # 安装时重写 Interpreter_Win
└── 安装说明.txt
```

**release_tool 核心铁律**：
1. **便携运行时**：`install_runtime.py install_minimal` 构建（~120MB），用 uv 安装依赖，禁止从开发机复制 numpy/pandas（C 扩展 DLL 不完整）
2. **PYTHONPATH 必须正斜杠**：`C:/Users/...` 而非 `C:\Users\...`，避免 `\U` 被 Python 解析为 unicode 转义
3. **PROJECT_NAME 必须为 "xlwings"**：VBA 用 `PROJECT_NAME & ".conf"` 查找配置表，quickstart 默认 "myaddin" 会导致找不到配置表
4. **COM 保存后必须 Ribbon 回注**：`wb.Save()` 会破坏 customUI 注册，安装脚本自动执行 `reinject_ribbon_after_com_save.py`
5. **embed_code.py 排除非业务目录**：自动排除 tests/、dist/、dist_release/、__pycache__/、.venv/、runtime/、site-packages/ 等
6. **安装脚本自动修复配置表**：安装后自动修复 PYTHONPATH 正斜杠、填充空值、保持 RELEASE_EMBED_CODE 原值、自动 Ribbon 回注
**操作 4——交付物验证门禁**：E（配置表完整性）、J（运行时验证）、K（配置表重复键）。

**COM 散装修改后的 Ribbon 回注**：任何绕过构建脚本直接对 xlam 做 COM 保存都会丢失 customUI 注册，必须随后执行 `reinject_ribbon_after_com_save.py` 回注，再跑门禁 A。

便携运行时打包细节（embeddable zip 构建流程、体积控制 < 100MB）、ZIP 分发时 `__file__` 语义与交付清单见 `references/10-deployment-delivery.md`。

### 6.5.11 安装

**目标**：用 6.5.10 整备的安装脚本执行安装。安装必须用安装脚本，禁止手动复制文件或手动操作 Excel。

**操作——执行安装**：
- XLSTART 分发：双击 `install.ps1`，自动完成获取 XLSTART→禁用同名旧 AddIns 注册→复制 xlam→写配置表
- 自研 release_tool 分发：双击 `install_release.bat`，自动完成校验完整性→复制程序与运行时→改写 Interpreter_Win→配置表统一修复（PYTHONPATH 正斜杠+空值填充+RELEASE_EMBED_CODE 保持）→Ribbon 回注→移除 Mark-of-the-Web+受信任位置→创建桌面快捷方式

**XLSTART 分发 vs 注册表分发（关键区别）**：
- XLSTART 分发：xlam 复制到 `%APPDATA%\Microsoft\Excel\XLSTART\`，Excel 启动自动加载，不需要 AddIns.Add 注册
- 注册表分发：xlam 在非 XLSTART 路径，须 `AddIns.Add(path)` + `Installed=True` 注册
- 安装前清理是必须步骤：枚举 AddIns，禁用同名旧注册（防止旧路径残留与 XLSTART 新副本冲突——报"无法打开两个同名工作簿"）

**安装前必须终止残留进程**：旧面板进程（python.exe）会占用 runtime DLL，导致安装时 `Remove-Item` 报"访问被拒绝"。安装脚本应自动检测并终止占用进程，或提示用户关闭 Excel 和面板窗口。

### 6.5.12 安装后集成验证

**目标**：安装加载项后，在真实 Excel 环境中验证加载项端到端可用——包括加载项冒烟、RunPython 链路、UDF 调用、Ribbon 交互、门禁验证。

**前置条件**：6.5.11 安装完成，加载项已注册到 XLSTART 或注册表。

**操作 1——安装目录冒烟测试**：

release_tool 分发必须执行 `smoke_test.py --install-dir <安装目录>`，5 项测试全部通过才算安装成功：
1. **目录结构**：xlam、runtime/、python.exe 存在；PYTHONPATH 模式还需 src/ 和入口模块
2. **配置表**：Interpreter_Win 指向 runtime/python.exe、PYTHONPATH 用正斜杠；code embed 模式 RELEASE_EMBED_CODE=True 且存在 .py sheet，PYTHONPATH 模式 RELEASE_EMBED_CODE=False
3. **源码导入**：入口模块、src.app、src.main 等可正常 import（code embed 模式验证 embedded_code_extractor 可导入）
4. **面板启动**：模拟 open_panel() 调用，验证端口开放
5. **API 响应**：GET / 返回 200 OK

**操作 2——加载项冒烟测试**：

冒烟 = "能不能跑起来"，不是业务正确性验证。新实例（避免污染用户 Excel）→ 确认加载 → 触发回调 → 校验输出 → 退出。推荐最小测试数据：1 行正常 + 1 行边界 + 1 行错误。

**加载成功判据（高频误判坑）**：COM 附着实例的 `Workbooks` 集合**不枚举** XLSTART 加载的 xlam（`Workbooks.Count` 不含加载项），`Workbooks` 遍历看不到不代表未加载。正确判据二选一：
1. **VBE 工程存在**：`app.api.VBE.VBProjects` 中存在与加载项同名的 VBA 工程
2. **直接调用探针**：`app.api.Run("'<xlam文件名>'!RunPython", "import sys; print('loaded_ok')")` 不抛错即加载成功

按名访问 `xl.Workbooks('<xlam名>')` 也可（不抛错即存在），但**禁止用 Workbooks.Count 或遍历判断加载与否**。

**操作 3——RunPython 调用验证**：COM 启动 Excel→打开 xlam→`wb.set_mock_caller()`→import 业务模块→调用函数→断言返回值。`set_mock_caller` 用法：`xw.Book('file.xlam').set_mock_caller()`。

**操作 4——UDF 测试**（含 UDF 的项目必须执行）：
- **xlsm 与 xlam 测试模式不同**：xlam 须先加载加载项再创建测试工作簿
- **白标 xlam UDF 测试前提（高频坑）**：开发机同时安装标准 xlwings.xlam 和白标 xlam 时，两者 `xlwings` VBA 模块同名冲突，白标 xlam UDF 返回 #NAME?。测试前须在 COM 自动化中禁用标准加载项：
  ```python
  for a in app.api.AddIns:
      if 'xlwings.xlam' in a.Name and 'myproject' not in a.Name:
          a.Installed = False
  ```
  且白标 xlam 须以 `app.books.open()` 方式保持打开
- **UDF 双路径测试**：单元格公式断言（验转换器）+ 纯 Python 逻辑断言（验业务）

**操作 5——xlsm 宏工作簿回归**：真实 Excel 自动化回归，骨架 `examples/excel-automated-testing-master/`。覆盖 VBA 宏调用、公式联动、UDF 输出。测试模板见 `templates/internal-bases/excel-addin-core/tests/test_smoke.py.tmpl`

**操作 6——UserForm 验证**：
- 第二层：窗体打开冒烟（Show 不报错）
- 第三层：UI 自动化（可选增强，默认不做——脆弱）
- **模态对话框风险**：MsgBox 会阻塞自动化——测试时用参数化/回调替代弹窗，或确保弹窗可自动关闭（见 `VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/automatically-dismiss-a-message-box.md`）

**操作 7——门禁验证**：
- C/D：`verify_addin_registered.ps1 -AddinName <项目名> -AddinPath <路径>`
- F：`gate_f_runpython_chain.py`（RunPython 加载链路验证）
- J：`gate_j_e2e_runtime.py [--ports ...] [--api-url ...]`。Excel 集成验证（启动 Excel→加载 xlam→调用 RunPython→验证返回值）在交付前手动执行一次（默认不纳入自动门禁，因耗时较长）
- L：`gate_l_real_excel_launch.py --xlstart <目录> --xlam <名> [--force-kill]`（真实启动 Excel 验证自动加载）。运行前终止本次自动化产生的残留实例（按 PID，见 3.4；禁止裸 `taskkill /F /IM`）

**验证失败处理**：
- 冒烟测试失败 → 检查配置表（PYTHONPATH 正斜杠、PROJECT_NAME、Interpreter_Win）
- 门禁 L 失败（Excel 启动无加载项）→ 检查 XLSTART 路径或注册表注册是否正确
- RunPython 失败 → 检查 embedded_code_extractor 是否在 runtime/Lib/site-packages/ 中
- UDF 返回 #NAME? → 检查标准 xlwings.xlam 是否禁用、白标 xlam 是否保持打开
- 面板启动失败 → 检查端口是否被占用、pywebview 是否指定 edgechromium 后端
- COM 测试失败 → 优先检查本次自动化创建的 Excel 实例是否残留（按 PID 查 `Get-Process EXCEL`）；跨会话僵尸实例无法终止且污染 ROT 时，改用 `DispatchEx` 新实例规避（见 3.4）
### 6.5.13 分发

**目标**：将验证通过的交付物打包，准备分发给最终用户。

**操作——打包分发**：
- XLSTART 分发：交付文件夹打包为 zip
- 自研 release_tool 分发：`python scripts/release_tool/make_delivery_package.py --dist <交付文件夹>` → 自检分发 zip

**用户安装/卸载**（全程无 Python 概念）：解压→双击 install.bat→重启 Excel→使用；卸载双击 uninstall.bat 自动逐项清理（受信任位置注册表键→桌面快捷方式→XLSTART 中的 xlam→安装目录，移入回收站）。对偶保证：安装器写过的每一项都在清单里，卸载器逐项回收；系统 Python、PATH、机器级注册表永不动。

**交付说明文档要求**：`安装说明.txt` 至少覆盖文件清单、安装步骤、使用说明（Ribbon 入口/UDF 用法/配置表）、卸载步骤、常见问题（宏被禁用、解释器路径错误、"无法打开两个同名工作簿"=旧注册残留先运行卸载脚本）。

**本步规则——交付纪律**：
- 不要默认让用户自己打开终端运行 PowerShell（交付走一键安装，开发期脚本不出现在交付包）
- 普通用户优先窗口/按钮/文件选择，而非命令行参数
- COM 对象必须彻底销毁
- 交付产物必须包含示例文件
- 安装必须用安装脚本，禁止手动复制

---
## 7 场景 D：源码学习与二次开发

**适用**：xlwings 源码研读、二次开发参考、入门学习与教学演示、产物质检。

### 7.1 源码架构认知

源码研读入口为技能根目录下的 `xlwings/`（上游仓库根）：Python 包 `xlwings/xlwings/`、原生桥 `xlwings/xlwingsdll/`、测试 `xlwings/tests/`、文档 `xlwings/docs/`。**详细技术研读（xlwingsdll 原生 DLL、xlwings 包架构、构建与测试体系）见 `references/13-scenario-d-source-code.md` 第 2 章**。

- **对象模型**：`main.py` 定义全部公开类（App/Book/Sheet/Range/Chart 等）；`base_classes.py` 为平台无关基类；`constants.py` 为 Excel 常量枚举；`utils.py` 为工具函数
- **平台适配层**：Windows 用 `_xlwindows.py` + `_win32patch.py`（pywin32 COM + 行为补丁）；macOS 用 `_xlmac.py` + `mac_dict.py` + `xlwings-dev.applescript`
- **COM 桥**：`com_server.py`（XLPython COM 接口，VBA↔Python 对象操作协议）+ `xlwingsdll/`（C++ 原生 DLL：进程启动与数组维度转换）
- **类型转换**：`conversion/`（Converter/Accessor/Pipeline 框架 + standard/numpy/pandas/polars 转换器）
- **UDF 系统**：`udfs.py`（`@xw.func`/`@xw.arg`/`@xw.ret`/`@xw.script`，VBA 包装自动生成）
- **命令行与杂项**：`cli.py`（quickstart/addin install）、`server.py`（REST 服务端）、`reports.py`、`expansion.py`（Range 动态扩展）
- **PRO 功能模块**（`xlwings/pro/`，需许可）：四引擎、Reports、ObjectHandle、OfficeJS UDF、许可证管理

平台差异汇总见 `xlwings/docs/missing_features.md`。

Web 加载项（Lite）完整案例 `examples/taxi-duckdb-main/`（一个内嵌 Office.js 加载项与 Python 脚本的自包含 .xlsx，DuckDB 数据应用，不依赖 PRO/Server）——内嵌内容已提取至 `taxi-duckdb-main/taxi-duckdb-main/extracted/`（`extract_webextension.py` 一步到位：解包到 `full/` 且 `main.py` / `requirements.txt` 解码明文就地生成），案例研读见 `references/15-xlwings-lite-guidance.md` 阶段七，条目说明见 `examples/ReadMe.md` 第 10 节。

### 7.2 入门学习与教学演示

材料为内置 `examples/python-for-excel-course-main/`（2018 官方入门课程，0-Intro 至 7-Part7）与 `examples/xlwings-demo-master/`（13 场景综合演示）。演示文件按业务主题唯一选择：最小脚本+UDF 用 `basics/xwdemo.py`；数据读取与图表用 `correlation/correlation.py`；性能主题用 `performance/arrays.ipynb`；Jupyter 交互用 `interactive/interactive.ipynb`。执行前先注入本地源码（见 3.1），含可执行单元格的 notebook 顺序执行验证输出。

### 7.3 构建与测试体系

构建系统（pyproject.toml/Makefile/Rust Cargo.toml/C++ xlwingsdll 工程）、测试套件分类（对象层/转换层/异步与流式/引擎与专项/UDF/报告）、构建脚本与开发计划详见 `references/13-scenario-d-source-code.md` 2.3 节。

### 7.4 PRO 功能研读（仅源码研读，不部署运行）

xlwings PRO 为双许可（PolyForm Noncommercial 1.0.0 / 商业许可），无许可调用抛 `LicenseError`。核心能力包括四种引擎（excel/remote/calamine/officejs）、Reports 报告系统、Reader 只读引擎、code embed/release 部署、ObjectHandle 自定义对象传递。深度技术分析见 `references/13-scenario-d-source-code.md`：引擎原理（含 calamine 只读引擎 `pro/_xlcalamine.py` 的 Python 适配层与 Office.js 转换层）见 1.2 节，许可证机制见 1.1，Reports 架构见 1.3，部署机制见 1.4；Office.js 自定义函数与自定义脚本的完整源码研读（remote 引擎转换层、注入类型、流式推流、data types 协议、locale 日期归一）已独立为 `references/16-xlwings-officejs.md`（与官方教程对照的工作流编排）。**xlwings Lite（免费 Web 加载项）的独立参考手册（产品能力/运行机制/开发工作流/自托管/taxi-duckdb 案例 + BookAsync 等源码研读）见 `references/15-xlwings-lite-guidance.md`**。

### 7.5 服务端工程与部署机制

官方 xlwings Server 实现位于 `xlwings-server/`（与源码包独立，source-available 双许可），为 Excel 与 Google Sheets 提供 Python 支持，可自托管于裸机/Linux VM/Docker Compose/Kubernetes/Serverless。**其完整六块研读（docs 61 篇教程 / xlwings_server 包 / deployment / nginx / scripts / tests）与开发生命周期工作流（预研选型→初始化→扩展机制→自定义函数/脚本→认证安全→测试调试→部署运维升级）见 `references/14-xlwings-server-guidance.md`（Office.js Excel 加载项的后端运行时，8 阶段工作流，13-d 1.7 已压缩为指针）；Office.js 客户端引擎（值转换层 + UDF/脚本全链路，Server 与 Lite 共享的语义内核）见 `references/16-xlwings-officejs.md`**。Python 代码部署三种方式：ZIP 打包（与工作簿同名，xlwings 自动识别）、RunFrozenPython（PyInstaller 冻结 .exe，仅 Windows 不支持 UDF）、PRO 部署（`code embed`/`release`，需许可）。

### 7.6 产物质检

| 任务 | 骨架项目 | 要点 |
|------|---------|------|
| 静态坏引用扫描 | `examples/static-excel-test-master/` | 全程不需要 Excel 实例 |
| 报表交叉核验 | `examples/cross-check-reports-main/` | 两份报表对比 |

各任务执行步骤与仓库职责见 `examples/ReadMe.md` 质检小节，按该小节流程整体执行；详细展开见 `references/11-scenario-a-python-automation.md` 七章 7.3。

实战任务案例（Eikon 数据接入、Excel/Web 双端原型、Power Query 库管理）由场景 A 4.3 引用介绍，详细展开见 `references/11-scenario-a-python-automation.md` 七章 7.4–7.6。

xlwings PRO 深度技术分析（许可证机制、四引擎原理、Reports 架构、ObjectHandle）见 `references/13-scenario-d-source-code.md`；PRO 部署机制（code embed/release）对照研读见 `references/10-deployment-delivery.md` 附录 A（开发加载项部署走自研 release_tool 主线）；**xlwings Lite 独立参考手册（免费 Web 加载项，开发生命周期工作流编排）见 `references/15-xlwings-lite-guidance.md`**；Office.js 引擎（Server 与 Lite 共享的语义内核）见 `references/16-xlwings-officejs.md`；xlwings Server 服务端的独立深度参考（开发生命周期工作流编排，六块全覆盖，13-d 1.7 已指针化）见 `references/14-xlwings-server-guidance.md`。

---

## 8 调试与故障排除（所有场景通用）

**故障沉淀规则（技能维护约定）**：
- 开发过程中遇到的故障、BUG、踩坑，**优先沉淀**到 `docs/troubleshooting.md`（持续追加，按「症状/根因/修复/教训」组织）
- 只有**核心、主干、非常重要**的内容（必经流程、门禁、铁律、形态判定、关键顺序约束）才沉淀到本 SKILL.md
- 判断标准：能防止重复踩坑的经验与坑点 → `docs/troubleshooting.md`；决定工作流走向的规则与铁律 → 本 SKILL.md

**通用调试方法**：`xlwings/docs/debugging.md`；性能优化：`xlwings/docs/troubleshooting.md`。

**跨平台差异速查**：
- UDF 仅 Windows；`Characters` 对象与 `app.interactive` macOS 不支持；macOS 不支持多实例写同一文件
- Excel 版本 < 15 时 macOS 使用 `mac_latin2` 编码
- WPS 表格通过 COM 兼容，部分高级功能可能不可用

**其他参考**：缺失功能 `xlwings/docs/missing_features.md`；常见问题 `xlwings/docs/troubleshooting.md`。

> docs 目录中 `conf.py`/`Makefile`/`make.bat`/`requirements.txt`/`index_latex.md` 及 `locales/`/`_static/`/`_templates/`/`_ext/`/`images/` 为官方文档站构建产物、翻译与资源文件，非工作流引用对象；`pro/` 为需 PRO 许可的付费文档，本技能不依赖。

---

## 9 上游跟踪与更新机制（技能维护）

本地上游原文目录共 **17 路**，随官方仓库演进，由脚本统一跟踪，**禁止手改上游原文**（zip 快照）：

- `xlwings/`：xlwings 核心源码包，跟踪 `main`（开发线，提交数领先最新正式版），快照 commit 见 `manifest.json` 的 `pinned_sha`；目录名不含版本号
- `examples/`、`MCP-Server/` 与技能根目录下全部开源仓库：`xlwings-server`（跟 main，目录在技能根目录 `xlwings-server/`，由 examples/ 迁出）、`mcp-server-xlwings`、`xlwings-mcp-server`、`excel-mcp`、`Excel_MCP_Server`、`Excel_udf_itus`、`cross-check-reports`、`excel-automated-testing`、`python-for-excel-course`、`simulation-demo`、`static-excel-test`、`taxi-duckdb`、`xl-pq-handler`、`xlwings-demo`、`xlwings-eikon`、`xlwings-factsheet-demo`——后 15 个跟踪默认分支（main/master）

**组件**：
- `manifest.json`：上游锁定清单（repo / ref / local_dir / pinned_sha / last_synced，单一事实来源；ref 为 tag 或默认分支）
- `scripts/sync_upstream.py`：检测与同步——`--check` 仅检测；默认检测并同步有更新的上游；`--force <id>` 强制重拉指定来源。检测经 gh API（降级 git ls-remote），下载经 codeload tarball，原子替换并追加 SYNCLOG.md
- `SYNCLOG.md`：同步历史（每次同步留痕）

用法：`python scripts/sync_upstream.py --check` 先看是否有更新；有更新或确认同步则运行 `python scripts/sync_upstream.py`。

同步后注意：全部 `local_dir` 都是固定目录名，上游出新版不需要改自研文档里的路径；需要改的是本文档与 `references/13-scenario-d-source-code.md` 等自研文件里**引用上游源码处数**（行数、文件大小、行号锚点、被上游删改的文件名）——脚本不做这件事，同步后必须人工复核。zip 快照无法追溯确切 commit，登记基线取登记日上游 ref 最新 commit，若实际快照更旧，`--check` 会提示漂移，按需 `--force` 重拉。每次同步后向用户汇报变更摘要。









