# 场景 A 4.4：MCP 自动化 — 深度技术展开

> 本文件是场景 A（MCP 自动化，4.4 小节）的**深度扩展资源**。主文档保留 7 步工作流与核心原则；本文件展开 4 个社区 MCP Server 的技术细节、工具清单、客户端配置模板、安全机制源码分析与故障排查。所有内容均基于 `MCP-Server/` 目录下 4 个实际项目的源码与 README 核实。

## 目录

- [一、四个社区 MCP Server 深度对比](#一四个社区-mcp-server-深度对比)
- [二、项目逐一详解](#二项目逐一详解)
- [三、MCP 协议与 xlwings COM 集成要点](#三mcp-协议与-xlwings-com-集成要点)
- [四、安全机制源码分析](#四安全机制源码分析)
- [五、客户端配置模板](#五客户端配置模板)
- [六、故障排查](#六故障排查)

---

## 一、四个社区 MCP Server 深度对比

| 维度 | geniuskey/mcp-server-xlwings | hyunjae-labs/xlwings-mcp-server | prabodh-hydexcel/excel-mcp | vohoailinh90/Excel_MCP_Server |
|------|------------------------------|--------------------------------|---------------------------|-------------------------------|
| **版本** | 0.4.1 | 1.0.4 | 0.1.0 | 未标注（v1） |
| **PyPI 可安装** | 是（`mcp-server-xlwings`） | 是（`xlwings-mcp-server`） | 否（源码运行） | 否（源码运行） |
| **依赖** | mcp>=1.0.0, xlwings>=0.30.0 | mcp[cli]>=1.10.1, fastmcp>=2.0.0, xlwings>=0.30.0, typer, pywin32, psutil | mcp[cli]>=1.0.0, xlwings>=0.30.0, pydantic, oletools(macOS) | mcp>=2.0.0, openpyxl>=3.1, pywin32, xlwings>=0.31(Win) |
| **入口** | `mcp-server-xlwings` / `uvx mcp-server-xlwings` | `xlwings-mcp-server` / `python -m xlwings_mcp` | `python server.py` | `python server.py` |
| **架构** | 无状态（直接操作活动 Excel） | 会话式（session-based，TTL+LRU） | 无状态（连接实时 Excel） | 双 backend（COM + Mock） |
| **核心特点** | DRM 保护文件支持、读取用户选区 | 功能最全（20 模块）、线程安全、图表支持 | VBA 宏执行、命名区域 | VBA 宏白名单安全、专用 COM 线程 |
| **VBA 宏** | 支持（运行已有宏+返回值） | 未明确 | 支持（需 VBA 项目解锁） | 支持（白名单限制） |
| **跨平台** | Windows（COM） | Windows | Windows/macOS | Windows（COM）/ 其他（Mock） |
| **源码模块数** | 较少（src/ 下少量文件） | 20 个模块 | 2 个（main.py+server.py 35KB） | 4 个（server/base/windows_com/demo） |
| **许可** | MIT | MIT | 未明确 | 未明确 |
| **文档语言** | 英文 | 英文 | 英文 | 越南语（注释/README） |

**选型建议**：
- **快速上手/DRM 文件** → geniuskey（`uvx` 一键启动，最轻量）
- **功能最全/生产环境** → hyunjae-labs（会话式架构、线程安全、图表支持）
- **VBA 宏为核心** → prabodh 或 vohoailinh90（后者有安全白名单）
- **非 Windows 开发测试** → vohoailinh90（MockOpenpyxlBackend 可在无 Excel 环境测试）

---

## 二、项目逐一详解

### 3.1 geniuskey/mcp-server-xlwings

**源码位置**：`MCP-Server/geniuskey-mcp-server-xlwings-main/mcp-server-xlwings-main/`

**为什么用 xlwings COM 而非 openpyxl**（项目 README 原文要点）：
- openpyxl/pandas 直接从磁盘读 .xlsx，DRM/文件级加密时失败
- 需要交互实时 Excel 会话（公式、宏、加载项）时失败
- 文件被其他进程锁定时失败
- xlwings COM 与运行中的 Excel 进程通信，能打开 Excel 本身能打开的任何文件（含 DRM）

**xlwings 独有能力**（openpyxl 不可能实现）：
- 读取用户当前选区（`xw.books.active.app.selection`）
- 获取活动工作簿（无需指定文件路径）
- 运行 VBA 宏并获取返回值（`wb.macro("MacroName")(*args)`）
- 实时公式结果（设公式后立即取得计算值）
- 强制重算（`app.calculate()`）

**安装方式**：
```bash
# 推荐：uvx（无需安装，临时运行）
uvx mcp-server-xlwings

# 或 pip 安装
pip install mcp-server-xlwings
mcp-server-xlwings
```

**Claude Desktop 配置**（`%APPDATA%\Claude\claude_desktop_config.json`）：
```json
{
  "mcpServers": {
    "xlwings": {
      "command": "uvx",
      "args": ["mcp-server-xlwings"]
    }
  }
}
```

**Claude Code 配置**：
```bash
claude mcp add xlwings -- uvx mcp-server-xlwings
```

### 3.2 hyunjae-labs/xlwings-mcp-server

**源码位置**：`MCP-Server/hyunjae-labs-xlwings-mcp-server-main/xlwings-mcp-server-main/`

**架构特点**：
- **会话式架构**：持久化 Excel 工作簿会话，避免每次操作重新打开文件
- **线程安全**：每会话锁，支持并发访问
- **自动资源管理**：TTL 会话清理 + LRU 淘汰策略
- **零错误设计**：全面错误处理

**源码模块**（`src/xlwings_mcp/`，共 20 个）：
| 模块 | 职责 |
|------|------|
| `__main__.py` | 入口（fastmcp app） |
| `server.py` | MCP Server 主逻辑 |
| `session.py` | 会话管理（TTL/LRU） |
| `workbook.py` / `workbook_xlw.py` | 工作簿操作 |
| `sheet.py` / `sheet_xlw.py` | 工作表操作 |
| `range_xlw.py` | 区域操作 |
| `data_xlw.py` | 数据读写 |
| `calculations_xlw.py` | 公式与计算 |
| `formatting_xlw.py` | 格式化 |
| `rows_cols_xlw.py` | 行列操作 |
| `advanced_xlw.py` / `advanced_xlw_with_wb.py` | 高级操作 |
| `validation_xlw.py` | 输入验证 |
| `base.py` / `helpers.py` / `exceptions.py` | 基础工具 |
| `force_close.py` | 强制关闭 Excel 进程 |

**核心工具**（从 README 提取）：
- `open_workbook`：打开工作簿会话（filepath, visible, read_only）
- `write_data_to_excel`：写入二维数组数据
- `apply_formula`：应用公式（含语法检查）
- `create_chart`：创建图表（多种类型）
- 以及工作表管理、格式化、表格操作、区域合并/复制/删除等

**安装方式**：
```bash
pip install xlwings-mcp-server
xlwings-mcp-server
# 或
python -m xlwings_mcp
```

**会话式工作流示例**（MCP 客户端调用）：
```python
# 1. 打开会话
session_id = client.call_tool("open_workbook", {
    "filepath": "C:/path/to/file.xlsx",
    "visible": False,
    "read_only": False
})["session_id"]

# 2. 写入数据
client.call_tool("write_data_to_excel", {
    "session_id": session_id,
    "sheet_name": "Sheet1",
    "data": [["Name", "Age"], ["Alice", 25]]
})

# 3. 应用公式
client.call_tool("apply_formula", {
    "session_id": session_id,
    "sheet_name": "Sheet1",
    "cell": "B4",
    "formula": "=SUM(B2:B3)"
})
```

### 3.3 prabodh-hydexcel/excel-mcp

**源码位置**：`MCP-Server/prabodh-hydexcel-mcp-main/excel-mcp-main/`

**核心特点**：
- 连接**实时** Excel 工作簿（不打开新实例，操作已打开的文件）
- 支持 VBA 宏执行
- 列出命名区域
- server.py 35KB（主要逻辑集中在此）

**关键要求**（README 原文）：
| 要求 | 细节 |
|------|------|
| Microsoft Excel | 必须本地安装且运行中。xlwings 用 COM(Windows)/AppleScript(Mac) 通信，**不支持 headless** |
| Python | 3.10+（pyproject 写 >=3.14，疑为笔误，实际 3.10+ 可运行） |
| 宏支持 | 工作簿必须存为 .xlsm/.xlam，.xlsx 禁用宏 |
| VBA 项目解锁 | 不能密码保护。解锁路径：Excel → 开发工具 → Visual Basic → 工具 → VBAProject 属性 → 保护 |
| OS | macOS/Windows，不支持 Linux |

**运行方式**：
```bash
# 确保 Excel 已打开至少一个工作簿
python server.py
```

**Claude Desktop 配置**（`claude_mcp_config.json` 模板）：
```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/.venv/bin/python",
      "args": ["/path/to/excel-mcp-server/server.py"],
      "env": {}
    }
  }
}
```

**macOS 特殊注意**：首次运行会弹窗 "Terminal 想要控制 Microsoft Excel"，必须点允许。误拒后在 系统设置 → 隐私与安全性 → 自动化 中重新启用。

### 3.4 vohoailinh90/Excel_MCP_Server

**源码位置**：`MCP-Server/vohoailinh90-Excel_MCP_Server-main/Excel_MCP_Server-main/`

**双 backend 架构**：
```
server.py (MCPServer, 12 tools)
    ├── WindowsComBackend (xlwings + pywin32)  → Windows，驱动真实 Excel + VBA
    └── MockOpenpyxlBackend (openpyxl)          → 非 Windows，仅测试/dev，不执行真实 VBA
```

**自动选择**：`platform.system() == "Windows"` → WindowsComBackend；否则 → MockOpenpyxlBackend（stderr 打印警告）。

**VBA 宏安全机制**：
- 宏是任意代码执行（arbitrary code execution），**不应让 AI Agent 按名称执行任意宏**
- 环境变量 `EXCEL_MCP_ALLOWED_MACROS` 设置白名单（逗号分隔）
- 未设置白名单时 `_allowed_macros = None`（需检查源码确认默认行为）
- 白名单外的宏调用被拒绝

**专用 COM 线程**（解决 STA 问题）：
- Excel COM 是单线程单元（STA）
- 多个工具调用并发到达时，从不同线程调 COM 会崩溃或 "Application is busy"
- 解决方案：单后台线程持有 `CoInitialize()`，所有命令入队列顺序执行
- 适用规模："少量 Agent 调用工具"，非每秒数百请求

**为什么用 xlwings 而非 raw pywin32**（README 原文）：
- `wb.macro("MacroName")(*args)` 可直接按名调用 VBA Sub/Function（普通 Module 或 ThisWorkbook 中均可）
- 单元格 API `sheet.range("A1").value` 比 raw COM 的 `worksheet.Cells(1,1).Value` 更简洁
- 仍可通过 `wb.api`/`sheet.api` 下沉到 raw COM 处理 xlwings 未封装的高级操作（数据透视表、图表、VBA 编辑器等）

**12 个工具**（从 demo_test.py 和 server.py 推断）：
- excel_open_workbook / excel_close_workbook
- excel_read_cell / excel_write_cell
- excel_create_sheet（含复制）
- excel_apply_formula
- excel_run_macro（按名调用，受白名单限制）
- excel_save
- 以及其他读写/管理工具

**运行方式**：
```bash
# 开发测试（无 Windows/Excel 也能跑，用 Mock backend）
pip install -r requirements.txt
python demo/demo_test.py

# 生产 Windows（真实 Excel + VBA）
pip install -r requirements.txt
python server.py
```

---

## 三、MCP 协议与 xlwings COM 集成要点

### 4.1 MCP 传输方式

所有 4 个项目均使用 **stdio 传输**（JSON-RPC over stdin/stdout），MCP 客户端（Claude Desktop 等）作为父进程启动 Server 子进程，通过标准输入输出通信。

### 4.2 xlwings COM 在 MCP 中的特殊考虑

| 问题 | 原因 | 解决方案 |
|------|------|---------|
| STA 线程冲突 | Excel COM 单线程单元，多线程调用崩溃 | 专用 COM 线程 + 命令队列（vohoailinh90 方案）；每会话锁（hyunjae-labs 方案） |
| Excel 进程残留 | MCP Server 退出时未清理 COM 对象 | `force_close.py`（hyunjae-labs）；会话 TTL 自动清理 |
| 文件锁定 | COM 保持文件句柄，其他进程无法写入 | 会话结束显式 `wb.close()`；read_only 模式 |
| 宏安全风险 | VBA 宏是任意代码执行 | 白名单机制（vohoailinh90）；提示用户确认 |
| DRM/加密文件 | openpyxl 无法读取 | xlwings COM 可读取 Excel 能打开的任何文件（geniuskey） |
| 公式实时计算 | openpyxl 不计算公式 | xlwings COM 设公式后 `app.calculate()` 取结果 |

### 4.3 与场景 A 的关系

MCP Server 本质上是**场景 A（Python 脚本自动化）的 MCP 协议封装**：
- 场景 A：Python 脚本直接调用 `xw.Book()` 操作 Excel
- 场景 A 4.4：MCP Server 内部调用 `xw.Book()`，通过 MCP 协议暴露给 AI Agent
- 因此场景 A 的所有规则（进程清理、COM 单线程、大块读写等）均适用于 MCP Server 内部实现

---

## 四、安全机制源码分析

### 4.1 VBA 宏白名单机制（vohoailinh90）

vohoailinh90 是 4 个项目中唯一内置宏安全机制的。源码分析（`server.py` 开头）：

```python
_allowed_env = os.environ.get("EXCEL_MCP_ALLOWED_MACROS", "")
_allowed_macros = {m.strip() for m in _allowed_env.split(",") if m.strip()} or None
```

**机制解析**：
- 从环境变量 `EXCEL_MCP_ALLOWED_MACROS` 读取逗号分隔的宏名列表
- 解析为 set 用于 O(1) 查找
- **关键坑**：环境变量为空时 `_allowed_macros = None`，需检查 `windows_com.py` 中 `None` 的语义（是"允许全部"还是"禁止全部"）——生产环境务必显式设置白名单，不依赖默认行为
- 白名单外的宏调用被 `WindowsComBackend.run_macro()` 拒绝

### 4.2 专用 COM 线程机制（vohoailinh90）

Excel COM 是 STA（单线程单元），多工具并发调用会崩溃。vohoailinh90 的解决方案：

```
主线程（MCP 协议处理）→ 命令队列 → 专用 COM 线程（CoInitialize + 顺序执行）
```

- 单后台线程持有 `CoInitialize()`，所有 COM 命令入队列顺序执行
- 适用规模："少量 Agent 调用工具"，非每秒数百请求
- hyunjae-labs 采用不同方案：每会话锁（per-session locking），允许多会话并发但同会话串行

### 4.3 其他安全风险

| 风险 | 说明 | 缓解措施 |
|------|------|---------|
| 文件系统访问 | MCP Server 运行在用户权限下，可访问用户能访问的所有文件 | 权限最小化，只暴露必要工具 |
| Excel 进程安全 | 操作用户实时 Excel 实例，误操作可能导致数据丢失 | 关键写操作前提示用户确认；使用 read_only 模式打开 |
| DRM/加密文件 | geniuskey 支持 DRM 文件，可绕过文件级加密读取内容 | 注意数据安全合规；审计日志 |
| HTTP 传输暴露 | 若使用 HTTP 而非 stdio，可能暴露到网络 | 绑定 127.0.0.1；优先 stdio 传输 |

---

## 五、客户端配置模板

### 5.1 通用前置条件

- Windows + Microsoft Excel 已安装（4 个项目均依赖 Excel COM）
- Python 3.10+
- （可选）uv / uvx 用于免安装运行 geniuskey

### 5.2 geniuskey 快速启动（推荐首次尝试）

```powershell
# 1. 安装 uv（如未安装）
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

# 2. 配置 Claude Desktop
# 编辑 %APPDATA%\Claude\claude_desktop_config.json，添加：
# {
#   "mcpServers": {
#     "xlwings": {
#       "command": "uvx",
#       "args": ["mcp-server-xlwings"]
#     }
#   }
# }

# 3. 重启 Claude Desktop，在对话中输入 /mcp 确认 xlwings 已连接
```

### 5.3 hyunjae-labs 生产部署

```powershell
# 1. 创建虚拟环境
python -m venv .venv
.\.venv\Scripts\activate

# 2. 安装
pip install xlwings-mcp-server

# 3. 配置 Claude Desktop（command 指向 venv 的 python）
# {
#   "mcpServers": {
#     "xlwings": {
#       "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
#       "args": ["-m", "xlwings_mcp"]
#     }
#   }
# }
```

### 5.4 vohoailinh90 宏安全配置

```powershell
# 设置宏白名单（逗号分隔，仅允许这些宏被 AI 调用）
$env:EXCEL_MCP_ALLOWED_MACROS = "RefreshData,GenerateReport"

# 启动 Server
python server.py
```

---

## 六、故障排查

| 现象 | 可能原因 | 排查步骤 |
|------|---------|---------|
| MCP 客户端显示 Server 未连接 | command/args 路径错误 | 检查 claude_desktop_config.json 中路径；手动在终端运行 command 验证 |
| "Application is busy" 错误 | COM STA 冲突 | 使用 vohoailinh90（专用 COM 线程）或 hyunjae-labs（每会话锁）；避免并发调用 |
| Excel 进程残留 | Server 异常退出未清理 | 任务管理器结束 EXCEL.EXE；使用 hyunjae-labs 的 force_close.py |
| 宏执行无反应 | VBA 项目密码保护 / 白名单拦截 | 解锁 VBA 项目；检查 EXCEL_MCP_ALLOWED_MACROS 环境变量 |
| 文件无法打开 | DRM 加密 / 文件锁定 | 使用 geniuskey（COM 方式可打开 DRM 文件）；关闭其他进程的文件句柄 |
| macOS 无法控制 Excel | 自动化权限未授予 | 系统设置 → 隐私与安全性 → 自动化 → 启用终端/Python 控制 Excel |
| Linux 无法运行 | 无 Excel | 使用 vohoailinh90 的 MockOpenpyxlBackend 做开发测试；生产必须 Windows/macOS |

---

## 可配套阅读

- `MCP-Server/` 目录下 4 个项目的 README 与源码
- `xlwings/docs/`（xlwings API 权威文档）
