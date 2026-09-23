# 场景 D 附录：xlwings Lite 深度技术分析与开发工作流

> 本文件是场景 D（源码学习与二次开发）的**独立深度扩展资源**，承接 `references/13-scenario-d-source-code.md` 中分离出的 xlwings Lite 内容（原 1.6 节、4.5 节，以及 1.8 的 Lite 特有适配分支），并按**开发生命周期工作流**重新编排：认知选型 → 环境准备 → 核心开发 → 测试调试 → 应用化分发 → 自托管部署 → 案例研读。
>
> **与 13/16 号文档的分工**：13 号文档保留 PRO 许可证机制、四引擎原理、Reports 架构、code embed/release 部署等 PRO 专有内容；Office.js 引擎**整体**源码研读（`pro/_xlofficejs.py` + `pro/udfs_officejs.py` 全链路）已独立为 `references/16-xlwings-officejs.md`（值转换层 / UDF 全链路 / 脚本全链路 / socket.io 会话，Server 与 Lite 共享的语义内核）；凡属 **Lite 特有**的适配分支（Pyodide 环境的 JsNull 空值归一、流式函数 `streaming_callback` 直推与 Lite 侧任务重启、`BookAsync` 注解的懒加载注入）由本文件对应小节承接（机制细节处指向 16 号），13 号 1.8 仅保留指针。本文件聚焦 xlwings Lite 的**产品能力、运行机制、开发工作流与全部 Lite 相关源码研读**。
>
> **官方教程的层次结构**（lite.xlwings.org 文档体系，本文件按此组织叙述顺序，各节均标注对应官方文档）：
> 1. **入门**：index → quickstart → custom-functions → custom-scripts → notebooks；
> 2. **进阶（数据与集成）**：dependencies → databases → web-requests → async-api → plotting；
> 3. **分发（应用化与共享）**：app-mode → personal-modules → files → security → limitations；
> 4. **自托管**：self-hosting → hosting-solutions → linux-vm / azure-container-apps / aws-ecs-express / migrating-to-self-hosted；
> 5. **参考**：editor → tests → configuration → environment-variables → troubleshooting → faq → changelog / about。
> 开发时按"入门 → 进阶 → 分发 → 自托管"的官方路径推进；本文件阶段一/二承载选型与准备，阶段三按"入门组 → 进阶组"展开，阶段四/五/六对应参考/分发/自托管组，阶段七以完整案例收束。
>
> **内容来源**：① `xlwings-lite/`（lite.xlwings.org 官方文档 31 篇离线镜像，本文件按阶段语义引用，不重复全文）；② `xlwings/xlwings/` 源码（ext 包、main.py、base_classes.py、__init__.py、udfs.py、pro/udfs_officejs.py 的 Lite 分支）——每节"源码验证"块给出官方教程结论对应的实现位置；③ `examples/taxi-duckdb-main/`（Lite 完整案例）。

## 目录

- [阶段一：认知与选型](#阶段一认知与选型)
  - [1.1 产品定位：Lite 是什么](#11-产品定位lite-是什么)
  - [1.2 能力全景：能做什么](#12-能力全景能做什么)
  - [1.3 适用场景与边界](#13-适用场景与边界)
  - [1.4 选型对比与升级路径](#14-选型对比与升级路径)
  - [1.5 源码探针：Lite 在 xlwings 包中的投影](#15-源码探针lite-在-xlwings-包中的投影)
- [阶段二：环境准备与运行机制](#阶段二环境准备与运行机制)
  - [2.1 安装与前置条件](#21-安装与前置条件)
  - [2.2 运行时架构：Pyodide 与浏览器沙箱](#22-运行时架构pyodide-与浏览器沙箱)
  - [2.3 代码载体与信任模型](#23-代码载体与信任模型)
  - [2.4 编辑器、配置与环境变量](#24-编辑器配置与环境变量)
- [阶段三：核心开发工作流（官方入门组 → 进阶组）](#阶段三核心开发工作流官方入门组--进阶组)
  - [3.1 自定义函数（Custom Functions）](#31-自定义函数custom-functions)
  - [3.2 自定义脚本（Custom Scripts）](#32-自定义脚本custom-scripts)
  - [3.3 Notebooks 单元格开发](#33-notebooks-单元格开发)
  - [3.4 依赖与包管理](#34-依赖与包管理)
  - [3.5 数据库](#35-数据库)
  - [3.6 Web 请求](#36-web-请求)
  - [3.7 异步 API：get_value / flush / load](#37-异步-apiget_value--flush--load)
  - [3.8 绘图](#38-绘图)
  - [3.9 个人模块与跨工作簿共享](#39-个人模块与跨工作簿共享)
- [阶段四：测试与调试](#阶段四测试与调试)
  - [4.1 内置 pytest 测试运行器](#41-内置-pytest-测试运行器)
  - [4.2 故障排查](#42-故障排查)
- [阶段五：应用化与分发](#阶段五应用化与分发)
  - [5.1 App Mode：面向终端用户的应用形态](#51-app-mode面向终端用户的应用形态)
  - [5.2 文件、安全与隐私](#52-文件安全与隐私)
  - [5.3 分发机制](#53-分发机制)
- [阶段六：自托管部署](#阶段六自托管部署)
  - [6.1 自托管概览与架构](#61-自托管概览与架构)
  - [6.2 托管解决方案横向对比](#62-托管解决方案横向对比)
  - [6.3 Linux VM 自托管](#63-linux-vm-自托管)
  - [6.4 云平台容器化自托管（Azure / AWS）](#64-云平台容器化自托管azure--aws)
  - [6.5 从云端迁移到自托管](#65-从云端迁移到自托管)
- [阶段七：案例研读——taxi-duckdb 一体化数据应用](#阶段七案例研读taxi-duckdb-一体化数据应用)
  - [7.1 案例定位与仓库构成](#71-案例定位与仓库构成)
  - [7.2 xlsx 内嵌结构解密](#72-xlsx-内嵌结构解密)
  - [7.3 Python 代码的嵌入与执行机制](#73-python-代码的嵌入与执行机制)
  - [7.4 main.py 双脚本分析](#74-mainpy-双脚本分析)
  - [7.5 DuckDB 在浏览器端（Pyodide）的使用模式](#75-duckdb-在浏览器端pyodide的使用模式)
  - [7.6 工程借鉴](#76-工程借鉴)
- [附录：版本演进要点（changelog 提炼）](#附录版本演进要点changelog-提炼)

---

## 阶段一：认知与选型

### 1.1 产品定位：Lite 是什么

**一句话**：xlwings Lite 是一个从 Office 加载项商店安装的 Excel 加载项，让你用 Python 自动化 Excel、构建原生自定义函数——**免费、跨平台、且无需安装 Python**（`xlwings-lite/index.md`）。

**官方自述逐句解读**（`index.md` 首屏）——"Install the free add-in from Excel's add-in store and you're done. It comes with a VS Code-like editor and Wingman, a powerful AI assistant."：
- **"Install the free add-in from Excel's add-in store"**：Lite **本身就是加载项**（Office Web Add-in，应用商店 ID `WA200008175`，见阶段七 7.2 的 OMEX 引用）——它是**成品**，不是让你开发的开发框架；安装动作发生在 Excel 的 `Add-ins` 入口内。
- **"and you're done"**：安装即完成——**无需安装 Python、无需配置解释器、无需命令行**。Python 运行时随加载项从 Pyodide CDN 拉取（阶段二 2.2），编辑器里 `main.py` 已含 `=HELLO("World")` 示例，安装后即可写出并运行第一个自定义函数。
- **"VS Code-like editor"**：加载项内置 Monaco 内核编辑器（`editor.md`），写 `.py` 模块的代码编辑体验与 VS Code 一致（快捷键/自动补全/输出窗格）。
- **"Wingman"**：内置 AI 助手（聊天 + `=WINGMAN()` 函数），自带 API key 接入 Claude/Gemini/OpenAI/自托管模型。

**Lite 与"开发 Excel 加载项"的关系——三种语境必须分清**：

| 语境 | 与 Lite 的关系 | 归属 |
|------|--------------|------|
| ① "开发一个 Office Web Add-in" | **不是 Lite 的职责**——Lite 本身就是一个 Web Add-in（闭源，从商店分发）；你不开发它，只安装它 | 官方/微软生态 |
| ② "用 Python 开发 Excel 自动化应用 / 自定义函数集" | **正是 Lite 的职责**——你开发的产物是"内嵌 Python 代码的工作簿"（`.xlsx`），App Mode 下是面向终端用户的应用形态；**分发形态是工作簿而非加载项** | Lite（本文件） |
| ③ "开发传统 Excel 加载项（.xlam VBA 宏加载项 / .xll 原生 XLL 扩展 / COM 加载项 .dll）" | **不是 Lite 的领域**——那是桌面版 xlwings + VBA 的领域（场景 A/C：Ribbon/VBA/COM 桥） | 场景 A/C |

**判断要点**：官方介绍里"add-in"指的是 **Lite 本体**（用户装的是加载项）；你在 Lite 里开发的是**应用内容**（Python 代码 + 工作簿），不是加载项。因此 Lite 与桌面版不冲突——桌面版开发 .xlam 加载项，Lite 开发"代码内嵌工作簿"的轻应用；二者共享同一套 Python API 语义（`@func`/`@script`/对象模型），代码可互相迁移（1.4 升级路径）。

**运行本质**：Python 不是运行在你的系统 Python 解释器中，而是运行在**浏览器运行时内的 Pyodide（WebAssembly 版 Python）**里。Office 加载项（含 Excel）基于浏览器内核渲染，Lite 把 Python 塞进了这个浏览器沙箱。因此：

- 目标机**零安装**：不需要 Python、不需要 conda、不需要 xlwings 桌面版，打开 Excel 加载加载项即可；
- 代码**内嵌工作簿**：Python 源码存储在 Excel 文件内部（详见 2.3 与阶段七 7.2），单文件即可分发；
- **本地优先**：Python 在用户自己的电脑上执行，不上传代码（详见 5.2）。

**愿景与许可**（`xlwings-lite/about.md`）：免费用于个人与商业用途，无需注册或登录；微软应用商店评分 5.0（43 条评价）。基于 xlwings 开源包（2014 年起）构建——Lite 与本地版共用同一套 API 语义（`xw.App/Book/Sheet/Range`、`@func`/`@script`），因此从 Lite 迁移到本地版通常**几乎零代码改动**（`index.md` "A path beyond WebAssembly"）。

### 1.2 能力全景：能做什么

Lite 提供六类核心能力（`index.md` "What you can do"，各自详展于对应文档，阶段三逐个展开）：

| 能力 | 入口 | 本质 | 对应源码/机制 |
|------|------|------|--------------|
| **自定义函数** | `=函数名(...)` 单元格公式 | `@func` 装饰器把 Python 函数注册为原生 UDF，pandas/Polars DataFrame 直进直出、Object handles 传自定义对象、绘图返回图片、流式函数实时推值 | `xlwings.udfs`（桌面同源）、`ObjectHandle`（`__init__.py`）、`pro/udfs_officejs.py`（officejs 引擎侧） |
| **自定义脚本** | 任务窗格按钮 / sheet 按钮 / F5 | `@script` 装饰器把整段 Python 变成按钮动作，参数由表单字段承载 | `WithScript`（`__init__.py`）、`pro/udfs_officejs.py` 脚本管线 |
| **Notebooks** | 任务窗格单元格式执行 | 类 Jupyter 的单元格执行，支持 Markdown 单元格与顶层 await | 加载项内置笔记本运行时 |
| **包安装** | requirements.txt | 用 PyPI/Pyodide 内置包（pandas/Polars/DuckDB 等），依赖存进工作簿 | Pyodide 包管理（`dependencies.md`） |
| **Web API 访问** | `httpx` 等 | 浏览器内发 HTTP 请求拉外部 API 数据 | 浏览器 fetch 通道（`web-requests.md`） |
| **App Mode** | Ctrl+Shift+M | 隐藏代码编辑器、脚本按钮化、自动生成表单字段，工作簿变成终端用户应用 | 加载项 UI 层（`app-mode.md`） |

**开发者体验**（`index.md` "Developer experience"）：VS Code 内核编辑器（快捷键/自动补全/输出窗格）、内置 pytest 运行器、个人模块（跨工作簿共享代码，类似 `Personal.xlsb`，可用 Git 管理）、Wingman AI 助手（自带 API key 接入 Claude/Gemini/OpenAI）。

### 1.3 适用场景与边界

**适用场景**（满足其一即值得考虑 Lite）：

1. **零安装分发**：目标用户不会/不能装 Python，但会用 Excel（Office 2021+ / Microsoft 365 / Excel for the web）——Lite 是最短路径；
2. **轻量数据应用**：Excel 内查询数据库、拉取 Web API、做 pandas/Polars 分析、Matplotlib 绘图，无需后端服务器；
3. **跨平台一致**：Windows / macOS / Web 三端同一套代码；
4. **教学与原型**：notebooks 单元格式开发 + 内置 pytest，适合渐进式学习与快速验证；
5. **企业内部受限环境**：代码需用户逐次批准执行（信任提示），自托管可完全内网化（阶段六）。

**边界与限制**（`xlwings-lite/limitations.md`，源自 Pyodide WASM 环境，立项前必须核对）：

| 限制 | 影响 | 绕过路径 |
|------|------|---------|
| `multiprocessing` / `threading` 不支持 | 依赖它们的包无法运行 | 用 asyncio 并发（Lite 原生支持 async/await） |
| TCP/IP 连接不支持 | 不能直连 PostgreSQL/MySQL 等数据库 | SQLite（内存/文件）可用；HTTP 类 API 可用 |
| Web API 受 CORS 限制 | 服务器须显式允许跨域；GitHub 可用 | 自建 CORS 代理（`web-requests.md`） |
| Pyodide 内存上限 2GB | 超大工作簿/数据集受限 | 异步 API 按需读取（阶段三 3.7）；超限转本地版/Server |
| 无调试器 | 无法断点调试 | print + Output pane、pytest 断言 |
| 部分非纯 Python 包需专为 Pyodide 构建 | 冷门 C 扩展包可能缺失 | 查 Pyodide 已构建包列表；纯 Python 包可直接安装 |

**性能认知**（`custom-scripts.md` "Performance"）：同步 API 会在脚本/单元格开始时**把整个工作簿的值全部加载进浏览器**——大工作簿慢且耗内存。两条官方优化路径：① `include`/`exclude` 参数限定只把需要的 sheet 传给 Python；② 异步 API 按需读取（3.7）。这决定了 Lite 开发中"数据边界"是第一设计考量。

### 1.4 选型对比与升级路径

Lite 在 xlwings 生态中的位置（与本地版 / Server 的关系，`index.md` "A path beyond WebAssembly" + `faq.md` "Comparisons"）：

| 维度 | xlwings Lite | 桌面 xlwings（社区版） | xlwings Server（PRO 生态） |
|------|-------------|----------------------|--------------------------|
| 运行位置 | 浏览器沙箱（Pyodide WASM） | 本地 Python 解释器（COM/appscript 桥） | 服务器端 Python（FastAPI + Socket.IO） |
| 目标机安装 | 只需 Office 加载项 | 需装 Python + xlwings 包 | 零安装（浏览器经 Web 访问） |
| 数据通道 | Office.js ↔ 浏览器 | COM/appscript ↔ 本地进程 | REST/socket.io ↔ 服务端 |
| 能力边界 | Pyodide 包生态、2GB 内存、无 TCP | 完整桌面能力（VBA/Ribbon/UDF/Excel 对象全量） | 服务端能力 + 多客户端（Excel/Google Sheets/Web） |
| 许可 | 免费 | BSD 开源 | source-available（Server）/ PRO 双许可 |
| 典型场景 | 零安装轻应用、教学、内部分发 | 深度 Excel 自动化、加载项开发 | 企业级服务化、多端共享 |

**升级路径**（官方明确支持的演进方向，`index.md`）：
1. **Lite → 桌面版**：代码 API 语义一致（`@func`/`@script`/对象模型），浏览器限制（2GB、无 TCP、无 threading）成为瓶颈时迁本地，通常只改少量导入与数据读取方式（同步 API 可用性恢复）；
2. **Lite → Server**：需要服务端能力、多用户共享、或把计算放服务器时迁移，前端概念（Office.js、任务窗格）一致。

**选型决策要点**：优先问三个问题——①目标用户能否安装 Python？（不能 → Lite 或 Server）；②是否需要本地文件系统/数据库直连/大内存？（需要 → 桌面版）；③是否需要服务端集中计算或多端接入？（需要 → Server）。Lite 是"零安装 + 无服务器"场景的默认答案；超出边界时按上述路径升级，而不是在 Lite 内硬扛。

**与场景 A/C 的文档关系**（本文件与 `references/11-scenario-a-python-automation.md` 的分工）：
- **同源 API，入口相反**：场景 A 教的是**从 Python 进程驱动 Excel**（`xw.App/Book` 脚本自动化、COM 桥、批量处理）；本文件教的是**在 Excel 内嵌代码、由 Excel 承载 Python**（Pyodide 沙箱）。两者共用同一套 `@func`/`@script`/对象模型语义（1.1、1.5），差别只在"谁主动"。
- **代码互迁**：场景 A 写的 UDF/脚本在 Lite 中大部分逐字可用（异步化例外，见 3.7）；Lite 应用迁到桌面版通常几乎零改动（官方升级路径，见上）。
- **功能互补、边界清晰**：场景 A 覆盖本机 Python + 完整 Excel 对象模型（VBA/Ribbon/图表/批注/性能调优）的深度自动化；Lite 覆盖"零安装 + 浏览器沙箱"的轻应用（Pyodide 依赖生态、≤2GB、无 TCP）。需要 VBA/Ribbon 或本机资源 → 场景 A/C；需要免安装分发 → 本文件。
- **文档定位**：场景 A 是 Python 自动化主线，加载项开发是场景 C 6.5 工作流；本文件是场景 D 的附录——用途是**理解与选型**，不是替代 6.5 开发流程：Lite 应用的交付形态是"内嵌 Python 的工作簿"，不经场景 C 的 xlam/Ribbon/安装脚本体系。

### 1.5 源码探针：Lite 在 xlwings 包中的投影

在 `xlwings/xlwings/` 源码中，Lite 并非独立实现一套对象模型，而是通过三种方式"投影"到既有架构上——理解这三处投影，就理解了 Lite 与桌面版的关系：

**① `BookAsync` 类型提示（`main.py` L1424-1441）——异步 API 的开关**

```python
class BookAsync(Book):
    """Type hint for the async API in xlwings Lite scripts.

    Annotate a script's `book` parameter with `xw.BookAsync` to opt into the
    async, on-demand API --- no cell values are pre-loaded, and you read them via
    `await myrange.get_value()`:
    ...
    At runtime it's a plain `Book`; the annotation only signals
    xlwings Lite to skip loading the values of the entire book up front.
    """
```

关键洞察：`BookAsync` **运行时就是普通 `Book`**，类型提示只是一个"信号"——Lite 加载器看到该注解就跳过整个工作簿的预加载。这就是 3.7 异步 API 的实现基础：不是新对象模型，而是"预加载行为的开关"。

**② `base_classes.py` 的异步方法族——能力边界在源码中的哨兵**

`base_classes.py` 定义了平台无关的异步读取方法（`get_value()` / `get_formula()` / `get_number_format()` / `get_color()` / `get_current_region()` 等 40+ 个 `get_*`），每个方法在非 Lite 环境抛：

```python
raise NotImplementedError("get_value() is only supported in xlwings Lite")
```

这是"接口契约完整、能力按需声明"的样板：所有平台（桌面/Server/Lite）共享同一基类签名，具体平台决定是否实现。Lite 侧实现这些 `get_*` 方法（经 Office.js 异步通道），桌面侧不实现（同步 `value` 属性已够用）。

**③ `ext/` 包（`sql.py` + `__init__.py`）——Lite/桌面共用的内置扩展**

`xlwings/ext/` 是 xlwings 包内独立扩展模块，当前只含 SQL 扩展：

```python
# ext/__init__.py
from .sql import sql, sql_dynamic

__all__ = "sql", "sql_dynamic"
```

`=SQL()` 是公式内嵌的数据库查询扩展（3.5 详研读），对 Lite 与桌面版同时可用——是"扩展点独立于引擎"的架构样例。

---

## 阶段二：环境准备与运行机制

### 2.1 安装与前置条件

**前置条件**（`xlwings-lite/quickstart.md`）：

| 平台 | 要求 |
|------|------|
| Windows | Microsoft 365，或 Office 2021+ |
| macOS | Microsoft 365，或 Office 2021+；macOS Ventura（13）或更新 |
| Excel for the web | 现代浏览器的最近版本；**免费版 Excel 也可用** |

**安装三步**（全部在 Excel 内完成，无需命令行）：
1. 点击 `Add-ins` 按钮（通常在 `Home` 选项卡；旧版在 `Insert` 选项卡）；
2. 搜索 `xlwings`；
3. 点击 `Add` 安装。

安装后 `Home` 选项卡右端出现 xlwings Lite 按钮，点击打开右侧任务窗格编辑器。`main.py` 自带示例代码：`=HELLO("World")` 调用自定义函数，绿色 Run Hello World 按钮（或 F5）运行脚本。

**企业环境**：加载项商店被封锁时，由 IT/Office 管理员从 Microsoft 365 管理中心安装，或自托管（阶段六）。**卸载**：`Add-ins` → `More Add-ins` → `Manage your Apps` → 三点菜单 → `Remove`。

### 2.2 运行时架构：Pyodide 与浏览器沙箱

**完整调用链**（理解 Lite 一切行为的基础）：

```
Excel 公式/按钮
   ↓ Office.js API（Excel 与加载项之间的官方桥）
Office 加载项 WebView（浏览器内核）
   ↓
xlwings Lite 加载项本体（office.js + 运行时，闭源，从商店拉取）
   ↓
Pyodide（WebAssembly 版 Python 解释器，从 CDN 加载）
   ↓
你的 main.py / 其他模块（在浏览器沙箱内执行）
```

**三层远程加载**（"无服务器"的真实边界——指无你自己的后端，而非零网络依赖，见阶段七 7.3）：
- 加载项本体：Office 加载项商店（OMEX）CDN；
- Python 运行时：Pyodide CDN（版本由工作簿内 `pyodideVersion` 指定）；
- 数据/依赖：PyPI 包（`requirements.txt`）、外部 API（浏览器 fetch）。

**计算全部在本地**（WebView 内 WASM）：SQL 查询、DataFrame 转换、绘图都在用户浏览器完成，无服务端计算。这也解释了为什么安全模型是"本地优先"（5.2）——代码和数据不出用户机器。

**与桌面版的本质差异**（决定开发习惯）：桌面版是"Python 进程 + COM 桥"（每次调用穿透进程边界）；Lite 是"Python 在浏览器里 + Office.js 桥"（异步、批量、按需）。因此 Lite 侧同步 API 需要预加载全簿值、写入是队列化的（阶段三 3.7 正是针对这两个差异的解法）。

### 2.3 代码载体与信任模型

**代码存储**：Python 源码以文本形式**内嵌在 Excel 文件内部**（Office Web 加载项机制，阶段七 7.2 有 zip 级解密）——分发只需一个 .xlsx/.xlsm，加载项是唯一依赖（`index.md` "Easy distribution"）。

**信任提示（Trust prompt）**（`security.md`）：打开含 Python 代码的工作簿时，Lite 先询问是否信任作者：
- **Trust and Continue**：加载并可能自动运行代码；选择被记住（除非代码变化 / Office 缓存清除 / Lite 更新）；
- **Don't Trust (View Only)**：只读编辑器查看代码，不执行任何 Python（App Mode 下需先按 Ctrl/Cmd+Shift+M 才显示此选项）。

**安全边界三层**（`security.md` 全篇）：
1. **沙箱隔离**：Python 在浏览器沙箱内，无直接操作系统访问（区别于 VBA 宏）；
2. **浏览器存储**：本地环境变量、设置（含 AI API key）、信任决策、Wingman 聊天记录存于浏览器本地存储（按源隔离，其他网站/加载项不可访问）；Files 导入导出走 IndexedDB；挂载的本地文件夹仅 Windows 桌面版支持且需逐次授权；
3. **残余风险**：被信任的工作簿代码可读取上述数据并可外发；XSS/被攻破的依赖/CDN/构建链风险由 CSP + SRI（Pyodide 加载校验）+ Dependabot 缓解但不可消除。

**工作簿级 vs 本地级配置**（`security.md` + `environment-variables.md`）：环境变量与设置可设为工作簿作用域（存进 Excel 文件内，随文件分发——**严禁存放密钥**）或本地作用域（存浏览器，每个用户各配各的）。非机密但属于工作簿的值（base URL、区域名）用工作簿级；密钥一律本地级。

### 2.4 编辑器、配置与环境变量

**编辑器**（`xlwings-lite/editor.md`）：VS Code 内核——自动保存、键盘快捷键（沿用 VS Code 习惯）、自动补全、输出窗格（print 结果显示处）、独立编辑器窗口（分离式多显示器开发）。

**配置**（`configuration.md` 四域）：
- **General**：本地/工作簿级设置、App Mode、信任与示例代码加载；
- **Python packages**：包安装源与 CORS 代理配置；
- **CORS proxy**：自建代理地址（应对 1.3 的 CORS 限制）；
- **Tooling**：编辑器外观与工具链。

**环境变量**（`environment-variables.md`）：本地级（浏览器存储）与工作簿级（内嵌文件）两作用域，命名与取值规则见官方文档；**密钥只放本地级**（5.2 已强调）。

---

## 阶段三：核心开发工作流（官方入门组 → 进阶组）

> 本阶段是 Lite 开发的主干，按 **lite.xlwings.org 官方教程的层次结构**推进：**入门组（3.1-3.3）**——自定义函数 → 自定义脚本 → Notebooks，从"单元格函数"到"整段脚本"再到"单元格式开发"，先验证最小闭环；**进阶组（3.4-3.8）**——依赖 → 数据库 → Web 请求 → 异步 API → 绘图，把应用接入真实数据与外部世界；**3.9 个人模块**收尾（官方归"分发/共享"语境，作为代码组织手段放在开发末端）。每小节统一为：官方文档要点（`xlwings-lite/` 对应篇）→ **源码验证**（`xlwings/` 实现位置），结论可追溯到实现，不凭文档印象。

### 3.1 自定义函数（Custom Functions）

**对应文档**：`xlwings-lite/custom-functions.md`（本技能镜像，53 个代码块全量保留）。

**核心语法**（`custom-functions.md` "Basic syntax"）：

```python
from xlwings import func

@func
def hello(name):
    return f"Hello {name}!"
```

单元格输入 `=HELLO("World")` 即得 `Hello World!`。这是"原生 Excel 自定义函数，用 Python 编写"的最小闭环。

**源码验证——`@func` 从哪来**：`@func`/`@arg`/`@ret`/`@script` 装饰器实现于 xlwings 包的 `xlwings/udfs.py`（与桌面版**完全同源**）——Lite 经 PyPI 安装 `xlwings==<版本>` 获得该模块，再由 Lite 加载项（闭源）经 Office.js 引擎桥接到 Excel（引擎侧转换层见 13 号文档 1.8）。因此桌面版学会的 `@func` 写法在 Lite 中逐字可用；差异只在"引擎通道"（桌面 COM vs Lite Office.js），不在装饰器语义。

**功能矩阵**（开发时按需取用）：

| 特性 | 用法 | 要点 |
|------|------|------|
| pandas DataFrame 参数/返回值 | `@arg("df", pd.DataFrame)` / `@ret(index=False, header=False)` | 单元格区域自动转 DataFrame；写回控制表头/索引（转换器实现见 13 号 2.2.1，Lite 复用同一 `conversion/` 管线） |
| Polars DataFrame | 类型提示 `df: pl.DataFrame` | 无 index 概念、header 仅 True/False |
| 类型提示替代装饰器 | `def f(df: pd.DataFrame) -> pd.DataFrame` | 返回类型可选，xlwings 自动探测 |
| 变长参数 | `def f(*args)` | 对应 Excel 区域展开 |
| 枚举参数 | `Literal["a","b"]` | 单元格下拉候选 |
| 日期时间 | 类型提示 `dt.date` / `dt.datetime` | 读写双向转换；datetime 精确到秒 |
| 自定义函数名 | `@func(name="MY_NAME")` | 覆盖默认（默认 = Python 函数名大写） |
| 命名空间 | `@func(namespace="NS")` | `=NS.FUNC()` 形式；支持子命名空间 |
| 数组维度 | 默认 1 维 list；`@arg(ndim=2)` 强制 2D | 单元格单值默认标量 |
| 错误处理 | 抛 `xw.XlwingsError` | 显示为 `#VALUE!` 等；NaN 写 `#NUM!`；Error 单元格对象 |
| 动态数组 | `@ret(expand="table")` | 函数结果自动溢出到邻格（Excel 365） |
| 易失函数 | `@func(volatile=True)` | 每次重算都调用 |
| **Object handles** | `-> object` 注解 | 大对象存句柄，不落单元格（下详） |
| **绘图** | 返回 matplotlib Figure | 自动上传图片缓存，Excel 单元格内显示（需联网） |
| **流式函数** | `async def` + `yield` | 实时推值到单元格（下详） |
| 调用格地址 | `Caller` 类型提示 | 拿到调用单元格引用（`custom-functions.md` "Accessing the calling cell"） |
| 函数后触发脚本 | 返回 `xw.WithScript(...)` | 计算边界后自动运行脚本（下详） |
| 异步函数 | `async def` | 内部 await；Lite 原生支持 |

**Object handles 深度**（`custom-functions.md` "Object handles" + 源码 `xlwings/__init__.py`）：

- 用途：把 pandas DataFrame 等**无法转成单元格**或**体量巨大**的 Python 对象返回给单个单元格；后续函数可把该句柄单元格当参数继续操作（链式处理，如 `=GetData() → FilterData(A1) → PlotData(B1)`）。
- 触发：返回值注解 `-> object`（或 `-> xw.ObjectHandle`）即返回句柄；单元格显示图标 + 类型，点击图标查看对象信息卡片。
- 定制展示（`ObjectHandle` 类，`__init__.py` L50-100）：`xw.ObjectHandle(df, text="120 rows", icon=..., properties={...})` 覆盖单元格文本、图标、卡片属性；三种注入方式——`@ret` 装饰器、Annotated 类型提示、直接返回 `xw.ObjectHandle` 实例。
- 缓存与失效：对象存于 Lite 对象缓存；`ObjectCacheMissError`（`__init__.py` L43-58）携带缓存 key，供调用方渲染"stale object handle"而非硬错误——句柄过期（LRU 驱逐/缓存重置）时的优雅降级路径。
- 版本要求：Object handles 需 `xlwings >= 0.36.5`（`requirements.txt` 中声明）。

**值语义的 Lite 特有细节——JsNull 空值归一**（引擎机制完整研读见 `references/16-xlwings-officejs.md` 阶段二·值转换层，Lite/Pyodide 环境才有）：引擎读侧 `clean_value_data` 逐元素 `_clean_value_data_element`——**Pyodide ≥ 0.28 的 `JsNull` 哨兵**（JS `null` 空单元格）与 `""` 均归一为 `empty_as`；data types 协议的 `dict`（`Error`/日期）取 `basicValue`；float 走 `number_builder`。**开发含义**：Lite 中空单元格在 Python 侧统一表现为 `empty_as`（默认 `None`），与桌面版一致，但实现路径不同（桌面是 COM 变体空值，Lite 是 JS 空值哨兵）——跨端迁移代码时空值语义不变。

**流式函数深度**（`custom-functions.md` "Streaming functions" + 源码 `xlwings/pro/udfs_officejs.py` L530、L634）：

- 形态：`async def` + `yield`（异步生成器），如 `streaming_clock()` 每秒 yield 一次当前时间；支持 1d/2d 返回值（numpy 数组）。
- 与传统 RTD 的区别：不用本地 COM 服务器；进程作为后台任务运行，**持续把更新推给 Excel**（而非 Excel 轮询）。
- 推送通道（源码验证，见 16 号文档阶段三·UDF 全链路 流式三形态）：`udfs_officejs.py` 的 `custom_functions_call` 对 asyncgen 走"三形态"分发——`background_tasks` 字典 + `task_key` 去重复用 + `on_task_done` 清理竞态保护；**推送端二选一：Server 走 socket.io `set-result-{task_key}` 推流，Lite/Pyodide 走 `streaming_callback` 直推**。Lite 侧流式任务**总是重启**（L634：任务去重逻辑在 Lite 侧简化，无 Server 的 sid 订阅计数）。
- 注意：流式函数不与 `WithScript` 组合（推送通道无法携带脚本负载）；流式函数带 `Caller` 注解时注册即报错（Office.js 中 stream 与 requiresAddress 互斥，注册期拒绝而非运行期注入 None）。

**绘图**（`custom-functions.md` "Plots" + `security.md` "Custom functions that return images"）：返回 matplotlib Figure 的自定义函数，因 Excel 只能经 URL 显示图片，Lite 自动把渲染 PNG 上传到图片缓存，Excel 取回并嵌入，取后即删。**此功能需要计算时联网**；想完全离线在 Settings > Local 关闭 "Allow custom functions to return images"。自托管时图片缓存随部署内网化（阶段六）。

### 3.2 自定义脚本（Custom Scripts）

**对应文档**：`xlwings-lite/custom-scripts.md`。

**核心语法**：

```python
from xlwings import script

@script
def hello_world(book: xw.Book):
    book.sheets.active["A1"].value = "Hello World!"
```

点击绿色 Run 按钮（或 F5）整段执行。`book` 参数是注入的当前工作簿对象（必须恰好一个）。

**源码验证——`@script` 与 `book` 注入**：`@script` 装饰器同样实现于 `xlwings/udfs.py`；引擎侧 `custom_scripts_call`（16 号文档阶段四·脚本全链路）负责运行时装配——current_user 前置 → 注入书对象（`_inject_value`）→ `_coerce_script_arg`（JSON 字符串按 `dt.date`/`dt.datetime` 提示 ISO 解析）→ keyword-only / `**kwargs` 参数拒绝 → 多余参数检查。**"book 必须恰一个"是引擎侧强制**（`_book_param_hint`：多 Book/BookAsync 注解 → 报错），不是文档约定。

**脚本参数与表单**（`custom-scripts.md` "Script arguments" + `app-mode.md` "Form fields"）：脚本可声明额外参数，App Mode 下自动渲染表单控件：

| 类型提示 | 表单元素 |
|---------|---------|
| `str` | 文本框 |
| `int` / `float` | 数字框 |
| `bool` | 复选框 |
| `Literal[...]` | 下拉框 |
| `datetime.date` | 日期选择器 |
| `datetime.datetime` | 日期时间选择器 |
| 无类型提示 | 文本框（传字符串） |

参数默认值预填；无默认值参数带红色星号必填，缺失则提示错误不运行。`typing.Annotated` 可加友好标签与帮助文本（`custom-scripts.md` "Labels and help text"）。**源码验证**：表单字段由引擎侧的元数据生成驱动——`custom_scripts_meta` 从函数签名提取参数类型与默认值，Lite 加载项据此渲染控件；与 `@func` 侧 `custom_functions_meta`（16 号文档阶段三·UDF 全链路：`stream`/`requiresAddress`/`volatile` 选项、枚举注册、重复 Excel 名检测）同一套"签名 → 清单 JSON"机制。

**按钮标签**：`@script(name="Run Hello World")` 覆盖按钮显示名；docstring 在 App Mode 渲染为 Markdown 说明（模块/脚本/函数三级均可写 docstring）。

**运行方式**：任务窗格按钮（`@script` 注册）、sheet 按钮（工作簿内插入按钮绑定脚本，`custom-scripts.md` "Running a script via sheet button"）、F5（当前选中脚本）。

**性能两法**（`custom-scripts.md` "Performance"）：
1. `@script(include=["Sheet1","Sheet2"])` / `@script(exclude=[...])`——限定传给 Python 的 sheet，避免整簿数据加载；
2. 异步 API 按需读取（3.7）。

**WithScript——函数完成后触发脚本**（源码 `xlwings/__init__.py` L95-150）：

```python
@xw.func
def hello_with_script(name):
    return xw.WithScript(f"Hello {name}!", "hello_args", args=[name, 42])
```

- 语义：自定义函数返回值正常写入单元格，`script` 参数指定的脚本在**下一个计算边界后**自动运行（ExcelApi < 1.8 时尽力而为、不保证单元格已提交）；函数执行期间禁止写网格，故脚本不能内联；
- 传函数对象优于传字符串：**记录定义模块**，让多模块运行的 Lite 能无歧义解析脚本（裸字符串只能按名搜索，两个模块同名脚本会歧义）；
- 陷阱：每次返回 WithScript 的调用都触发一次脚本——公式向下填充 N 行会运行 N 次；流式函数不支持 WithScript（3.1 已述）。

### 3.3 Notebooks 单元格开发

**对应文档**：`xlwings-lite/notebooks.md`。

**形态**：任务窗格内类 Jupyter 的单元格式开发——代码单元格（`# %%` 分隔）+ Markdown 单元格（`# %% [markdown]`）；运行按钮、单元格操作、键盘快捷键；**顶层 await 支持**（notebooks 中可不用 `async def` 直接 `await`，脚本则必须在 `async def` 内——`async-api.md` 明确区分）。

**绘图六库**（`notebooks.md` "Plots" + `plotting.md`）：Matplotlib / seaborn / Plotly / Bokeh / Altair / plotnine / xy——绘图结果直接渲染在 notebook 内。

**Excel 对象模型互操作**：notebooks 中获取当前工作簿用同步 `xw.books.active`（或异步 `await xw.books.get_active()`），读写语法与桌面版一致。

**迁移路径**（`notebooks.md` "Migrating from Python in Excel to xlwings Lite"）：官方为微软 Python in Excel 用户提供了迁移指南，对应概念映射可参考。

### 3.4 依赖与包管理

**对应文档**：`xlwings-lite/dependencies.md`。

**机制**：Lite 的 Python 是 Pyodide（WASM），包安装经 Pyodide 包工具（内置包 `loadPackage` / 非内置包 micropip 从 PyPI 解析）——因此**不是所有 PyPI 包都可用**（1.3 兼容性）；`requirements.txt` 编辑保存后需**重启 xlwings Lite** 才生效（官方明确提示）。

**双文件版本策略**（高频坑，必须理解）：
- `requirements.txt`（手写）：声明顶层依赖，**不要钉死精确版本**——建议只写上界（如 `xlwings<1.0.0`），精确版本由系统锁定；
- `requirements-pinned.txt`（自动生成）：Lite 记录实际安装的**精确版本（含子依赖）**，每次重启按它重装——**工作簿因此保持"构建时已验证"的版本组合，即使你从未手动钉版本**。文件头记录 `source sha256`；**只记包名+版本，不记 wheel URL/哈希**——保证版本稳定，但不是完整锁文件（安全缺口见 5.3）。
- **Pyodide 版本边界**：`requirements-pinned.txt` 的强制能力需要 **Pyodide 314.0.0+**；0.27.5（taxi-duckdb 案例所用）无法强制版本——每次启动从 `requirements.txt` 解析并刷新 pinned 文件。

**源码验证**：Lite 包安装链路不在开源 xlwings 包内（闭源加载项 + Pyodide 工具链）；开源侧可观测的是 `requirements.txt`/`requirements-pinned.txt` 作为工作簿 `we:property` 的存储形态（阶段七 7.2 zip 解密可见）。因此"依赖管理"在 Lite 中=文件管理 + 重启，无命令行。

### 3.5 数据库

**对应文档**：`xlwings-lite/databases.md`。

**硬约束**（Pyodide 无 TCP，1.3）：传统 SQL 数据库**直连不支持**。官方给出的三条出路：

1. **给数据库套 Web API 桥**（HTTPS 通道，浏览器可访问）：
   - PostgreSQL → **PostgREST**（把 PG 表自动映射为 REST 接口）；托管方案见下方 Supabase；
   - Oracle → **Oracle REST Data Services (ORDS)**；
   - MySQL → **MySQL REST Service (MRS)**；
   - 原生自带 REST 的数据库可直接用：**NocoDB** / **InfluxDB** / **CouchDB**。
2. **下载 SQLite 数据库文件**：从网络位置取 `.sqlite` 文件，浏览器内 WASM 版 SQLite 打开——**注意必须把 `sqlite3` 加进 `requirements.txt`**（Pyodide 内置模块并非默认加载）。
3. **Supabase（托管 Postgres，含 PostgREST）**：免费层可用、可自托管；**当前 Python 包兼容性坑**——`supabase`/`postgrest` 包在 Lite 下有问题，官方建议用 `requests` 或 `aiohttp` **直连 PostgREST 接口**（**httpx 当前也有问题**；用 `aiohttp` 必须同时把 `ssl` 加进 `requirements.txt`）。

**源码验证——`=SQL()` 内置扩展**（`xlwings/xlwings/ext/sql.py`，74 行，Lite 与桌面版共用）：公式内嵌的 SQLite 内存表查询——单元格区域即表（类型推断建表 → 值序列化 → 内存库 → 别名解析 → 查询返回，完整五步管线与 Excel 用法见 13 号 4.6 的 SQL 写法要点与下方代码）：

```python
@func
@arg("arg", expand="table", ndim=2)
@ret(expand="table")
def sql(query, *arg):
    return _sql(query, *arg)
```

Excel 用法：`=SQL("SELECT ...", ["alias1"], range1, ["alias2"], range2, ...)`——查询字符串 + 可选别名 + 数据区域；`expand="table"` 读取 + `expand="table"` 写回构成"区域↔表"双向映射。**开发含义**：不想引入数据库文件/服务的轻量场景，`=SQL()` 让"Excel 区域 = 可 SQL 查询的表"零依赖成立。

### 3.6 Web 请求

**对应文档**：`xlwings-lite/web-requests.md`。

**本质**：Python 在浏览器里发请求，受浏览器同源策略约束——**CORS 是 Lite Web 请求的第一约束**（不是 Python 网络库的问题，是浏览器安全模型）。

**三级应对**（按可控程度排序）：
1. **目标 API 已开 CORS**：直接用（如 GitHub raw——`pd.read_csv("https://raw.githubusercontent.com/.../penguins.csv")` 开箱即用）；
2. **你控制服务器**：在响应加 `Access-Control-Allow-Origin: https://addin.xlwings.org`（官方云版 origin；**自托管时用你自己的域名**，如 `https://xlwings-lite.mycompany.com`）；
3. **你不控制服务器**：走 **CORS 代理**——后端服务器转发请求并补 CORS 头；**推荐自建代理**（商业代理会把敏感数据送第三方）；自托管版内置 CORS 代理（`configuration.md` "CORS proxy" 配置）。

**源码验证**：浏览器内 HTTP 由 Pyodide 的 httpx 适配层实现（pyodide-http 对 `fetch` 的桥接），因此**httpx 兼容性取决于 Pyodide 版本**（3.5 中 Supabase 场景 httpx 有问题即是此类）；requests 在 Pyodide 中同样经该桥。**开发含义**：选 HTTP 库时以"目标环境实测"为准，不假设桌面行为。

### 3.7 异步 API：get_value / flush / load

**对应文档**：`xlwings-lite/async-api.md`。

**为什么需要**（Lite 与桌面版的三个差异 + 对策，`async-api.md` 开篇）：

| # | 同步 API 的局限 | 异步 API 对策 |
|---|----------------|--------------|
| 1 | 所有单元格值**一次性预加载**（大簿慢、可能 OOM） | `await myrange.get_value()` 按需只读你要的值 |
| 2 | 写入**队列化**（单元格/脚本结束时才统一应用） | `await book.flush()` 立即应用排队写入 |
| 3 | 读取**可能是旧状态**（反映单元格/脚本开始时的簿状态） | `await book.load()` 刷新全部 |

**三方法语义**：

- **`get_value()`**：`xw.BookAsync`（脚本）或 `await xw.books.get_active()`（notebook）获得异步 Book，之后用 `await myrange.get_value()` 读值；**支持与 `value` 相同的转换器与选项**——`await book.sheets["Sheet1"]["A1:C10"].options(pd.DataFrame, index=False).get_value()` 直接读 DataFrame。注意：同步 Book 上也能 `await get_value()` 拿当前值，但同步 Book 仍会预加载整簿。
- **`flush()`**：写值后立即可见——典型场景：写 A1 后在同一单元格/脚本内读它或读依赖单元格；print 后立即在 Output pane 看到输出；调用 `to_png()` 等写文件系统方法后在同一单元格/脚本内访问文件。自动 flush 仍发生在单元格/脚本结束时。
- **`load()`**：单元格/脚本运行前 Lite 自动执行一次，一般无需手动；中途刷新用 `await book.load()` 或 `await mysheet.load()`（sheet 级只刷该 sheet，更高效）。异步 Book 上 `load()` **不含值**（值走 `get_value()`）；同步 Book 上 `load()` 含值，可 `load(values=False)` 排除。

**源码研读——BookAsync 是"预加载开关"，懒加载注入在脚本管线**（引擎机制见 16 号文档阶段四·脚本全链路 懒加载注入，`xlwings/xlwings/main.py` L1424-1441 + `pro/udfs_officejs.py`）：

- `BookAsync` **运行时就是普通 `Book`**——类型提示只是信号：Lite 加载器看到该注解就**跳过整个工作簿的预加载**（`main.py` 注释原话："the annotation only signals xlwings Lite to skip loading the values of the entire book up front"）；
- 引擎侧 `@script` 的 `lazy=` 参数已弃用，**统一由 `book: xw.BookAsync` 注解表达**（内部发 `"lazy"` wire 键）；**book 参数必须恰一个**（`_book_param_hint`），`BookAsync` 注解与 `lazy=False` 显式冲突 → 报错（注解优先）；
- 运行时装配：`custom_scripts_call` 的 `_inject_value` 见 `BookAsync` 注解 → **`value.impl._lazy = True`**（懒加载开关落到 Range 对象），后续 `get_value()` 经 base_classes 异步方法族（40+ 个 `get_*`，非 Lite 环境抛 NotImplementedError）按需读取。

**工程借鉴**：接口契约统一（基类签名）+ 平台按需实现（Lite 实现、桌面不实现）+ 行为开关注解（BookAsync）——是多运行时共享 API 面但差异化执行的干净模板。

### 3.8 绘图

**对应文档**：`xlwings-lite/plotting.md`（与 `custom-functions.md` "Plots" / `notebooks.md` "Plots" 三处配合读）。

**三通道**（同一批 matplotlib/Plotly 代码，三个去处）：

| 通道 | 写法 | 产物去向 | 前置 |
|------|------|---------|------|
| 自定义函数 | `@func` 返回 Figure | 单元格内嵌图片（图片缓存上传/取回，3.1 绘图） | 需联网（关闭"Allow custom functions to return images"则禁用） |
| 脚本 | `sheet.pictures.add(fig)` | 写入 Excel 工作表（可定位 `anchor`） | 无网络依赖 |
| Notebooks | 单元格内 `plt.show()` 等 | 渲染在 notebook 内（六库支持，3.3） | 无网络依赖 |

**开发含义**：要"图进工作表"走脚本通道；要"图进单元格（随公式重算）"走自定义函数通道（接受联网）；要"开发期看图"走 notebooks。数据在浏览器本地渲染，无服务端。

### 3.9 个人模块与跨工作簿共享

**对应文档**：`xlwings-lite/personal-modules.md`（官方归"分发/共享"语境，本质是代码组织手段）。

- 用途：跨工作簿共享 Python 文件与包，类似 VBA 的 `Personal.xlsb`；可用 Git 管理（模块目录即仓库）；
- **Setup**：指定个人模块目录，Lite 将其中文件/包注入每个工作簿的 Python 环境；
- **Projects：工作簿专属模块**：`Projects` 目录下按工作簿隔离（工作簿级覆盖个人级）；
- **导入顺序**（`import order`）：工作簿内代码 → 项目专属模块 → 个人模块——同名时靠前者优先（与 `custom-functions.md` "workbook function takes precedence over personal module" 一致）；
- 限制：个人模块不可含非纯 Python 包（需按 Pyodide 规则构建，见 1.3）。

## 阶段四：测试与调试

### 4.1 内置 pytest 测试运行器

**对应文档**：`xlwings-lite/tests.md`。

Lite 内置 pytest 运行器，可测试 Python 代码与工作簿本身：

- **测试文件与函数**：任务窗格 Test 标签页发现并运行 `test_*.py` 中 `test_*` 函数；`pytest.main()` 可直接调用；
- **导入自己的代码**：`test_main.py` 中 `import main`（或测试文件与被测模块同目录），测试 `main.py` 的纯函数与转换逻辑；
- **异步测试**：pytest 原生支持 `async def test_...`（anyio/asyncio 插件），配合 Lite 异步 API 写 `await` 断言；
- **访问活动工作簿**：测试中 `xw.books.active`（或异步 `await xw.books.get_active()`）拿到当前簿做集成断言；
- **pytest 配置**：任务窗格可配置 pytest 参数（如 `-k` 选择、`--tb` 回溯格式）。

**与桌面版测试的分工**（对比 `09-testing-debugging-guidance.md`）：桌面版测试 = 本地 Python + COM 集成（`set_mock_caller`/`xw.serve`）；Lite 测试 = 浏览器内 pytest 运行器，天然集成工作簿。**共同点**：都是"纯函数逻辑断言 + 簿对象集成断言"两层；**差异点**：Lite 无 COM 进程、无需 mock caller（book 由加载器注入）、异步 API 需 `await`。

### 4.2 故障排查

**对应文档**：`xlwings-lite/troubleshooting.md` + `cleanup-missing-addin-refs.md`。

**高频故障与对策**（`troubleshooting.md`）：

| 现象 | 本质 | 对策 |
|------|------|------|
| `pyodide.ffi.JsException: NetworkError` | 网络/CDN 加载失败（Pyodide、包源、API 请求） | 检查网络、CORS、CDN 可达性；自建代理/自托管 |
| `Error initializing application: ... editorInstance.setScrollPosition` | 编辑器初始化竞态（缓存/版本冲突） | 清 Office 缓存（下）后重载 |
| `Unexpected error: SSL is not supported` | Pyodide 内 SSL/TLS 限制（部分场景） | 改用 http 源（内网）或支持的 HTTPS 通道 |
| 单元格显示 `#NAME?` | 自定义函数未注册/未导入 | 检查函数名、模块加载、requirements 版本（≥0.36.5 支持 Object handles 等） |
| 行为异常 | Office 缓存陈旧 | 清 Office 缓存（`cleanup-missing-addin-refs.md` 提供工作簿级清理指引） |

**清理缺失的加载项引用**（`cleanup-missing-addin-refs.md`）：工作簿可能残留已卸载加载项的引用；提供清理既有工作簿的步骤（移除失效的 `we:property`/引用），防止打开报错。

**无调试器的替代**（结合 1.3 限制）：print + Output pane（`flush()` 及时可见）、pytest 断言、逐步拆分脚本（notebooks 单元格粒度天然支持分步执行）。

---

## 阶段五：应用化与分发

### 5.1 App Mode：面向终端用户的应用形态

**对应文档**：`xlwings-lite/app-mode.md`。

**定位**：把任务窗格从"开发环境"变成"终端用户应用"——隐藏代码编辑器、每个脚本变成一个按钮、渲染 docstring 说明、列出可用自定义函数。

**核心机制**：
- **切换**：Ctrl+Shift+M（Win）/ Cmd+Shift+M（macOS），或 Settings > Workbook > App Mode；**设置保存在工作簿内**，下次打开仍为 App Mode；
- **表单字段**：脚本的额外参数自动渲染为表单（3.2 控件表）；字段用参数名作标签，`typing.Annotated` 可改标签与加帮助文本；无默认值参数红星必填；
- **Output pane**：默认隐藏；Settings > Workbook > App Mode 启用 "Show the Output pane"（同样随工作簿保存）；
- **docstring 渲染**：模块/`@script`/`@func` 的 docstring 在 App Mode 以 Markdown 显示（阶段七 7.4 的 `main.py` 展示了完整写法）；
- **信任提示**：App Mode 下打开工作簿只显示单一 **Trust and Continue** 按钮；想审查代码的进阶用户按 Ctrl/Cmd+Shift+M 可调出 **Don't Trust (View Only)**（只读编辑器，不执行）。

**工程价值**：一个 .xlsx + 加载项 = 完整的终端应用（表单输入 + 按钮动作 + 函数计算 + 输出），是"Excel 作为应用壳"的最轻交付形态。

### 5.2 文件、安全与隐私

**对应文档**：`xlwings-lite/files.md` + `security.md`。

**文件系统**（`files.md`）：
- **Import/Export**：从本地导入文件 / 导出文件到本地（浏览器 IndexedDB 中转）；**File system layout**：浏览器沙箱内的虚拟文件系统（`/data/` 等路径）；
- **Local Folders（仅 Windows 桌面版）**：挂载本地文件夹给受信任代码读写（每次会话可能重新授权）；挂载跨工作簿共享——**只挂最小所需文件夹**。

**安全要点汇总**（`security.md`，供分发前自查）：
1. 密钥（API key、密码、连接串）**绝不**放工作簿级环境变量（随文件分发）——只放本地级；
2. 工作簿级值随文件走：邮件、SharePoint、Git 都会带出——只放非机密值（base URL、区域名、报告名）；
3. 信任决策、本地环境变量、AI key、聊天记录存浏览器本地存储（未加密、源隔离）；
4. 互联网访问：请求从用户机器发出，走公司网络路径（代理/防火墙/日志照常生效）——受企业网络策略约束；
5. 包供应链：`requirements-pinned.txt` 提升可复现性但不验证下载字节；高保障部署用内部包镜像（Azure Artifacts）+ 禁公网 PyPI（需自托管）；
6. Wingman：提示词/上下文/源码上下文/API key 发往配置的 AI 提供方——审阅其数据处理条款，用低权限可轮换 key。

### 5.3 分发机制

**单文件分发**（Lite 的默认分发模型）：Python 代码内嵌工作簿（阶段二 2.3 / 阶段七 7.2），用户拿到 .xlsx 后：打开 → 信任提示 → Trust and Continue → 代码运行。**加载项本体是唯一外部依赖**（用户需先装 Lite，或企业预装）。

**分发前检查清单**（综合本文件各阶段）：
- [ ] 代码是否在浏览器限制内（1.3：无 threading/TCP、≤2GB、CORS 已测）？
- [ ] 依赖版本是否锁定（requirements-pinned.txt）？包是否均为 Pyodide 兼容？
- [ ] 密钥是否只存在于本地级配置（未入工作簿）？
- [ ] 是否需要 App Mode（终端用户无代码视角）？
- [ ] 是否需要自托管（内网化、内部包镜像、图片缓存不出网）？
- [ ] 无网络环境是否已关闭图片返回功能（3.1）？

**升级路径提醒**：Lite 分发受 Pyodide 边界限制；超出边界时按 1.4 升级到桌面版/Server（API 语义兼容，迁移成本低）。

---

## 阶段六：自托管部署

### 6.1 自托管概览与架构

**对应文档**：`xlwings-lite/self-hosting.md`。

**为什么自托管**（`security.md` 结尾）：把加载项本体、Pyodide 运行时、Python 包全部放到**你自己拥有和审计的基础设施**上——内网 PyPI 镜像（Azure Artifacts）降低供应链风险、图片缓存留在内网、加载项从内网分发、代码与数据不出企业边界。

**架构**（`self-hosting.md` "Architecture"）：自托管部署 = 静态文件服务器（托管加载项清单 + Pyodide 运行时 + 包镜像）+ 可选图片缓存 + 企业基础设施集成（TLS、身份、更新通道）。

### 6.2 托管解决方案横向对比

**对应文档**：`xlwings-lite/hosting-solutions.md` + `self-hosting.md` "Hosting Solutions"。

| 方案 | 形态 | 适用 | 文档 |
|------|------|------|------|
| **Linux VM** | 自管虚拟机，xlwings-lite CLI 安装 | 最大控制、常规运维 | `linux-vm.md` |
| **Azure Container Apps** | 托管容器 | 无服务器化、Azure 生态（Artifacts 镜像） | `azure-container-apps.md` + `azure-container-apps-cli.md` |
| **AWS ECS Express** | AWS 托管容器 | AWS 生态、Express 模式免集群管理 | `aws-ecs-express.md` |
| 其他 | 自有基础设施 | 定制 | `hosting-solutions.md` 总览 |

### 6.3 Linux VM 自托管

**对应文档**：`xlwings-lite/linux-vm.md`。

**安装五步**（`linux-vm.md` "Install"）：
1. **提供 TLS 证书**（HTTPS 必需：加载项商店要求安全源）；
2. **安装 xlwings-lite CLI**（官方命令行工具，负责部署与更新）；
3. **运行安装器**（CLI 拉取加载项 + Pyodide 运行时 + 包缓存到本机）；
4. **注册加载项到 Microsoft 365 管理中心**（企业内分发入口：管理员中心 Add-in 部署指向自托管 URL）；
5. 配置验证。

**维护**（`linux-vm.md` "Maintenance"）：更新 xlwings Lite（CLI 升级）、更新 xlwings-lite CLI、编辑配置、本地下载 PyPI 包（内网包缓存）、卸载、TLS 证书（Windows CA Server / Let's Encrypt certbot）、故障排查。

### 6.4 云平台容器化自托管（Azure / AWS）

**Azure Container Apps**（`azure-container-apps.md` + `azure-container-apps-cli.md`）：
- **部署容器**：CLI 或门户推送镜像 → Container Apps 托管运行（自动伸缩、免集群管理）；
- **连接 Azure Artifacts（可选）**：托管身份（managed identity，免密钥）或 PAT（Personal Access Token）两种方式接内部包镜像；
- **注册加载项**到 Microsoft 365 管理中心；
- **更新**：重建/滚动更新容器。

**AWS ECS Express**（`aws-ecs-express.md`）：
1. **创建 ECS Express Mode 服务**（Express 模式：无需管理集群/节点，AWS 托管）；
2. **配置**（可选）：环境变量、镜像、伸缩；
3. **加载项安装**：注册到 Microsoft 365 管理中心；
4. **更新**：服务滚动更新。

**共同要点**：容器化自托管 = 镜像（加载项静态资源 + Pyodide 运行时 + 包缓存）+ 平台托管（伸缩/更新/TLS）+ 企业注册（M365 管理中心分发）。

### 6.5 从云端迁移到自托管

**对应文档**：`xlwings-lite/migrating-to-self-hosted.md`。

- **转换工作簿**：把工作簿从"指向官方云（OMEX/CDN/PyPI）"改为"指向自托管源"——需要重新生成/修补工作簿内加载项引用（`xlwingsWorkbookId`、`pyodideVersion`、包源等字段，见阶段七 7.2 结构）；
- **Windows Protected View**：迁移后的工作簿可能触发受保护视图（来自网络/不可信位置）——按文档解除；
- **验证**：转换后打开工作簿，确认信任提示、依赖加载、图片缓存均走内网源。

---

## 阶段七：案例研读——taxi-duckdb 一体化数据应用

### 7.1 案例定位与仓库构成

**定位**：xlwings 官方"一个 .xlsx 自包含完整数据应用"样板——工作簿内嵌 Office Web 加载项（xlwings Lite / Pyodide 0.27.5 浏览器端 Python），任务窗格提供两个 `@script` 按钮，用 DuckDB 直接查询纽约出租车 Parquet 数据并回写 Excel 报告图表。**不依赖 PRO/Server**，是阶段一至阶段五全部概念的完整落地样例。

**仓库构成**（`examples/taxi-duckdb-main/taxi-duckdb-main/`）：
- `LICENSE`（MIT）；
- `taxi_local.xlsx`（11.6KB）——全部应用逻辑在此一个文件；
- `extracted/`——本技能提取的内嵌源码（`README.md` 目录说明）：`extract_webextension.py`（一步到位提取脚本：解包 + 解码，stdlib 无依赖）、`full/`（完整解包产物：12 个 zip 条目原样落盘 + `main.py` 98 行 / `requirements.txt` 解码明文就地生成于 `xl/webextensions/`——解包即解码，无第二目录）。

### 7.2 xlsx 内嵌结构解密

zip 解包 `taxi_local.xlsx`（xlsx 本质是 zip，`extract_webextension.py` 完整落盘 12 个条目）：

- `xl/webextensions/webextension1.xml`：加载项定义——AppSource 引用 `wa200008175`（xlwings 官方 OMEX 加载项）、`xlwingsWorkbookId`（UUID `a0403022-...`）、`pyodideVersion=0.27.5`、`addinVersion=1.0.0.0-35`；
- `xl/webextensions/taskpanes.xml`：任务窗格右侧停靠、宽 837、可见；`xlwingsSettingsWorkbook={"startupBehavior":"taskpane"}`——打开工作簿即启动窗格；
- 六个 `we:property` 承载全部应用内容：`main.py`（后端脚本全文，**HTML 实体 + JSON 字符串双重编码**）、`requirements.txt`（依赖清单）、运行参数与设置；
- `xl/worksheets/sheet1.xml`：`<sheetData/>` 空表——数据全部由脚本运行时写入，文件只是"壳"。

**双编码必然性**：XML 属性值只能是单行文本，整段 Python 必须两次转换才能安全落位（先转义成合法属性字符 `&quot;`/`&gt;`，再包成合法 JSON 字符串）；还原即反向 `html.unescape` + `json.loads`——`extract_webextension.py` 已工具化，实测还原 98 行 / 2528 字符与原文一致。

### 7.3 Python 代码的嵌入与执行机制

**嵌入载体链**（代码如何进入文件）：
```
明文 main.py → HTML 实体转义 → JSON 字符串包裹 → <we:property name="main.py" value="…"/> → webextension1.xml → 随 xlsx（zip）分发
```
磁盘上只有 `taxi_local.xlsx` 一个文件，代码"活在" zip 内清单的属性里。

**运行时执行链**（代码如何跑起来）：
1. 打开 xlsx → Excel 读清单，按 OMEX 引用 `wa200008175` 从 Office 加载项商店拉取 xlwings Lite 加载项本体（office.js + Lite 运行时，闭源）；
2. `startupBehavior=taskpane` → 右侧任务窗格自动打开；`xlwingsWorkbookId` 确认宿主工作簿；
3. 读 `pyodideVersion` → 从 Pyodide CDN 启动浏览器内 Python（WASM 0.27.5）；
4. 解析 `requirements.txt`：PyPI 安装 `xlwings==0.33.20` 等，pandas / matplotlib / duckdb 直接用 Pyodide 内置；
5. 还原 `main.py` → 交 Pyodide 编译 → `@script` 把 `explore` / `report` 注册为任务窗格按钮；
6. 点击按钮 → 整段脚本执行：`duckdb.read_parquet("/taxi/…")`（pyodide-http 虚拟路径取数）→ SQL → pandas → Lite 异步 API 写回 Excel（`options` / `tables` / `pictures`）。

**存储形态 ≠ 运行形态**：文件里是"XML 属性内双编码文本"，运行中是"Pyodide 内存中的 Python 模块"——清单只负责分发，执行在浏览器 WebView 内完成；不落盘、无后端。

**远程加载的三来源**（呼应阶段二 2.2）：加载项本体（OMEX CDN）、Python 运行时（Pyodide CDN）、数据文件（`/taxi/` 虚拟路径，经 pyodide-http 在 Python 内发 HTTP 获取）；计算全部在 WebView 本地。

### 7.4 main.py 双脚本分析

`from xlwings import script`，任务窗格按钮触发整段脚本，参数仅 `book`：

**`explore(book)`——数据探查**：
```python
duckdb.read_parquet("/taxi/yellow_tripdata_2025-01.parquet")
SELECT count(*) + DESCRIBE SELECT *   # 规模 + 结构
```
结果打印到控制台（规模 + schema），是"先侦察再分析"的标准第一步。

**`report(book)`——查询 → 分析 → 报告四层管线**：
1. **查询**：SQL 按 `hour(tpep_pickup_datetime)` 分组（行程数 `count(*)`、平均里程 `round(avg(trip_distance),2)`，过滤 `trip_distance BETWEEN 0.1 AND 50`）→ `duckdb.sql(query).df()` 转 pandas；
2. **写回**：新建 sheet → 标题（A1 加粗 24pt）→ 表格（A3 起 `options(index=False).value = df` + `sheet.tables.add(resize(len(df)+1, len(df.columns)))`）；
3. **可视化**：`create_plot(df)` 双子图（行程数柱状 + 平均里程折线，`sharex=True`，figsize 12×8）→ `sheet.pictures.add(fig, name="TaxiAnalysis", anchor="E3")`；
4. **性能度量**：`print(f"Processed {df['Number of Trips'].sum():,} records in {total_time:.2f} seconds")`。

### 7.5 DuckDB 在浏览器端（Pyodide）的使用模式

- **`duckdb.read_parquet(path)` 返回 `DuckDBPyRelation`（惰性关系）**——不把数据载入内存，仅建立"数据源 → 查询"的引用；同一文件可多次 `duckdb.sql()` 复用；
- **执行与物化**：`duckdb.sql(query)` 执行 SQL 返回 relation；`relation.df()` 物化为 pandas DataFrame（案例即 `duckdb.sql(query).df()`）；`relation.show()` 打印前若干行到控制台（`explore` 用）；
- **DESCRIBE**：`DESCRIBE SELECT * FROM taxi` 取表结构（列名 + 类型），是"先侦察 schema 再写查询"的标准探查步骤；
- **SQL 特性**：`hour()` 时间函数、`count(*)`/`avg()`/`round(...,2)` 聚合、`BETWEEN` 过滤、`GROUP BY "Hour" ORDER BY "Hour"`（**带引号别名**——DuckDB 双引号标识符保留大小写与空格）；
- **与 pandas 互操作边界**：`.df()` 得到 pandas DataFrame 后即脱离 DuckDB 上下文（后续 `set_index`/`plot` 均走 pandas）——查询层与展示层解耦；
- **环境注意**：Pyodide 内置 duckdb 版本不固定（requirements.txt 未锁版本），`hour()` 等函数在不同版本行为一致但性能特性不同；`/taxi/` 为加载项环境虚拟路径。


> **桌面端内容已迁出**：以下与 xlwings Lite 无关的桌面端数据层知识已移至 `references/13-scenario-d-source-code.md` 4.7（数据服务层章节）——桌面端 DuckDB API 扩充（多源混查 / 视图注册 / 导出通道 / 性能特性）、DuckDB 官方 excel 扩展（`INSTALL excel; LOAD excel;` 直读直写 .xlsx）、与 xlwings 桌面版的通道分工与组合示例。本小节只保留 **Pyodide 浏览器端**的 DuckDB 用法。

### 7.6 工程借鉴

- "空 xlsx 壳 + 内嵌加载项"自包含交付：应用配置、Python 脚本、依赖清单全打进 `webextensions`，用户拿单文件即用——零安装、零后端（对应阶段五 5.3 分发机制）；
- `@script` 函数 = 任务窗格按钮的原子动作：参数仅 `book`，数据经 DuckDB/pandas 在脚本内流转，Excel 只做展示层（对应 3.2）；
- 四层职责分离：DuckDB 只管 SQL 查询、pandas 管转换、matplotlib 管可视化、`options/tables/pictures` 管写回 Excel——与阶段一 1.2 的能力矩阵一一对应；
- **内嵌源码提取工具化**：`extracted/extract_webextension.py`——一次命令完成解包 + 解码，`we:property` 双编码文本经 `html.unescape` + `json.loads` 双层还原；默认提取本案例，换路径即可提取任意 xlwings Lite 工作簿（原理与用法见 `extracted/README.md`）——是审计/学习他人 Lite 工作簿的标准工具；
- **注意**：DuckDB 读取依赖 Pyodide 运行环境与 `/taxi/` 虚拟路径，脱离 xlwings Lite 无法直接复现；研读价值在"加载项应用的组织方式"与"查询 → 分析 → 图表 → 写回"管线，而非可离线执行。

---

## 附录：版本演进要点（changelog 提炼）

**对应文档**：`xlwings-lite/changelog.md`（完整逐版更新记录）。

Lite 持续演进中，开发与分发时关注以下维度（具体版本号见 changelog.md）：

- **运行时与包管理**：Pyodide 版本升级（内存、兼容包列表变化）、requirements-pinned.txt 机制引入（锁定可复现性）、包安装 UI 演进；
- **API 能力**：Object handles（≥0.36.5）、异步 API（`get_value`/`flush`/`load`）、流式函数、`WithScript`、绘图与图片缓存机制；
- **应用形态**：App Mode 表单字段扩展（更多类型提示 → 控件映射）、Notebooks 能力（Markdown 单元格、顶层 await）、测试运行器（pytest 集成）；
- **安全与合规**：信任提示策略、CSP/SRI 加固、Wingman 隐私边界、图片缓存自托管化；
- **自托管**：xlwings-lite CLI、容器化方案（Azure Container Apps / AWS ECS Express）、内网包镜像支持。

---

## 可配套阅读

- `xlwings-lite/`（31 篇官方文档离线镜像 + README 索引 + 34 张插图）——本文件的逐阶段语义引用源，官方教程层次（入门/进阶/分发/自托管/参考）见 README"建议阅读顺序"；
- `xlwings/xlwings/udfs.py`——`@func`/`@arg`/`@ret`/`@script` 装饰器实现（Lite 与桌面版同源，3.1/3.2 源码验证）；
- `xlwings/xlwings/ext/`（sql.py + __init__.py）——`=SQL()` 扩展源码（3.5）；
- `xlwings/xlwings/main.py`（BookAsync，L1424-1441）与 `base_classes.py`（异步 get_* 方法族）——异步 API 源码（3.7）；API 文档见 `xlwings/docs/api/book_async.md`；
- `xlwings/xlwings/__init__.py`（ObjectHandle/ObjectCacheMissError/WithScript）——对象句柄与函数后触发脚本（3.1/3.2）；
- `xlwings/xlwings/pro/udfs_officejs.py`（streaming_callback、Lite 流式任务分支、_inject_value 懒加载注入）——officejs 引擎侧的 Lite 适配（引擎整体研读见 `references/16-xlwings-officejs.md`，Lite 侧解读见本文件 3.1/3.7）；
- `examples/taxi-duckdb-main/`（taxi_local.xlsx + extracted/）——阶段七案例；
- `references/13-scenario-d-source-code.md`（PRO 深度分析：许可证/四引擎/Reports/部署）与 `references/16-xlwings-officejs.md`（Office.js 引擎全链路，Lite 适配分支指针在其中）与 `references/14-xlwings-server-guidance.md`（xlwings Server 六块研读）——Lite 的 PRO/Server 兄弟专册。
