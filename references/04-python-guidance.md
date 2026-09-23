# Python 构建指南（6.5.4 扩展与导航）

> 本文件是工作流 6.5.4（自足模板：双形态×双入口、装饰器模板、分层铁律、业务核心规范、编码铁律、三阶段门禁）的**扩展资源**——官方文档导航 + 官方示例导航 + 进阶主题提炼。规则与模板以 6.5.4 正文为准，本文件不重复。

## 一、官方文档导航（`xlwings/docs/`，Python 侧权威）

| 文档 | 一句话要点 | 何时读 |
|------|-----------|--------|
| `udfs.md` | UDF 全部装饰器参数与语义（权威） | 写 UDF 前 |
| `converters.md` | 类型转换体系：内置转换器、`@arg`/`@ret`、自定义转换器 | 类型转换出问题时 |
| `datastructures.md` | DataFrame / NumPy 数组 / 列表进出 Excel 的规则（index/header 控制） | 写表/读表时 |
| `syntax_overview.md` | API 形态速览：`app.books[0]` vs `apps`、对象层级 | 不熟悉 xlwings API 时 |
| `connect_to_workbook.md` | `Book.caller()` 与连接工作簿的机制 | 理解 RunPython/UDF 上下文 |
| `threading_and_multiprocessing.md` | UDF 异步、多线程/多进程注意（Excel 单线程限制） | 异步 UDF / 长任务 |
| `command_line.md` | CLI 命令：`quickstart`/`udf`/`runpython`（`-m` 不可用，用 `Scripts\xlwings.exe`，见 `docs/troubleshooting.md` 坑 11） | 任何 CLI 调用前 |
| `debugging.md` | 调试手段：`xw.serve()`、断点、日志 | 调试 Python 侧 |
| `troubleshooting.md` | 官方常见问题 | 报错排查 |
| `deployment.md` | 部署选项概述 | 分发前 |
| `addin.md` | 配置表（Interpreter/PYTHONPATH 等）全部键 | 6.5.8 配置 |
| `customaddin.md` | 加载项打包与 Ribbon 集成（含 VBA 模块引用） | 6.5.6 Ribbon / 打包 |
| `matplotlib.md` / `jupyternotebooks.md` / `onedrive_sharepoint.md` / `other_office_apps.md` / `missing_features.md` / `installation.md` / `quickstart.md` | 图表 / 笔记本 / 云盘 / 其他 Office 宿主 / 缺失特性 / 安装 / 快速上手 | 按需 |

## 二、官方示例导航（`examples/`）

| 项目 | 可借鉴点 |
|------|---------|
| `xlwings-demo-master/` | 13 场景综合演示（文件级用途详展见 11-scenario-a 七章 7.7；PRO 部分——reader/restapi/Reports 四件套——见 13-scenario-d 四章） |
| `Excel_udf_itus-main/` | UDF 教程（类型提示/二维数组/DataFrame） |
| `excel-automated-testing-master/` | 自动化测试骨架（`test_mybook.py` 含 UDF 单测示例），配合安装后集成验证 |
| `python-for-excel-course-main/` | xlwings 官方课程（入门到进阶） |
| `xl-pq-handler-master/` | Power Query 联动处理 |
| `static-excel-test-master/` / `cross-check-reports-main/` | 静态检查 / 交叉核对报表 |
| `simulation-demo-master/` | 蒙特卡洛模拟（大量数组运算示范） |
| `xlwings-factsheet-demo-main/` | 基金数据表 |
| `xlwings-eikon-master/` | 金融数据终端集成（Refinitiv Eikon） |
| `xlwings-server/` | xlwings Server（服务端形态，399 文件，按需深读） |

## 三、进阶主题提炼（官方文档要点 + 实战）

### 转换器与数据结构实战（`converters.md` + `datastructures.md`）

**类型转换优先级**：类型提示/`Annotated`（PEP 484）是首选；`@xw.arg`/`@xw.ret` 装饰器优先级更高，覆盖类型提示。

**pandas DataFrame 进出 Excel（最常用）**：

```python
# 写入：控制 index/header 是否写入
import pandas as pd
from typing import Annotated

@xw.ret(index=False, header=True)  # 不写行索引，写列名
def get_data():
    return pd.DataFrame({"A": [1,2], "B": [3,4]})

# 读取：从单元格区域读入 DataFrame
df = xw.Book.caller().sheets[0].range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
```

**关键语义**：
- `header=1` 表示第 1 行是列名；`index=False` 不把首列当索引
- 空单元格 → `None`（DataFrame 中为 NaN）
- 日期列：读入为 `datetime`，写入时 Excel 自动识别为日期格式
- 二维数组 `ndim=2`：单行/单列结果注意维度塌缩（`[[1,2]]` vs `[1,2]`）

**numpy 数组**：

```python
@xw.ret(ndim=2)  # 强制二维，避免单行塌缩为一维
def get_array():
    return np.array([[1,2],[3,4]])
```

**自定义转换器**（复杂类型时实现 `converter` 协议）：

```python
class MyConverter:
    @staticmethod
    def read_value(value, options):
        return MyType.from_excel(value)
    @staticmethod
    def write_value(value, options):
        return value.to_excel()

# 注册后用 @xw.arg(converters=MyConverter) 或类型提示
```

**常见坑**：
- DataFrame 写入时 `index=True` 会多写一列行号（用户常困惑）→ 默认 `index=False`
- 大表写入用 `sheet.range("A1").value = df`（一次性写入），禁止逐单元格循环
- 读表前先 `expand("table")` 获取动态区域，禁止硬编码行号

### 线程与异步（`threading_and_multiprocessing.md`）
- UDF `async_mode='threading'` 不阻塞 Excel 重算
- Excel COM 是单线程单元（STA）：跨线程操作 Excel 对象受限
- 面板/服务线程（daemon）不得直接碰 `xw.books`，经 HTTP 或单元格中转

### 连接与语法（`connect_to_workbook.md` / `syntax_overview.md`）
- `Book.caller()` 必须放在被调用函数内部（非全局）
- 路径语义：`app.books` 打开的文件 vs 配置表连接的工作簿

## 四、模块设计理念（通用，源自实战）

- **职责单一**：入口层（装饰器薄壳）/ 业务核心层 / 基础设施层单向依赖
- **接口显式**：业务函数签名即契约（`(data, error)` 双元组返回，禁止吞异常）
- **延迟导入**：装饰器函数体内延迟 `import` 业务模块，保持入口轻、冷启动快
- **过桥成本意识**：模块划分粒度以"一次 RunPython 完成一个完整业务动作"为准，不过碎（6.5.3.2 原则）

### 错误处理与日志规范

**用户可见错误 vs 日志（严格分层）**：
- **用户可见错误**：入口层（`@xw.sub`/`@xw.func`）用 `_show_msg`（ctypes MessageBoxW）弹出，内容简短（≤50 字），告诉用户"发生了什么 + 怎么办"
- **日志**：业务核心层用 `logging` 记录完整堆栈和上下文，写入 `%LOCALAPPDATA%\<AddinName>\app.log`；禁止把堆栈弹给用户
- **RunPython 异常传播**：RunPython 不支持返回值，异常会被 xlwings 捕获并弹默认错误框——入口层必须 try/except 包裹，用 `_show_msg` 给出友好提示，禁止异常穿透到 xlwings 默认处理

**错误码约定（可选但推荐）**：
- 业务错误用 `(None, "E001: 股票代码格式错误")` 双元组返回，E 前缀 + 3 位数字
- 系统错误（网络/文件/COM）用 `(None, f"E999: {str(e)[:50]}")`，保留原始异常摘要
- 入口层根据 error 前缀决定弹什么消息

**日志配置模板**：

```python
import logging, os
log_dir = os.path.join(os.environ.get("LOCALAPPDATA", "."), "<AddinName>")
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    filename=os.path.join(log_dir, "app.log"),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)
```

## 五、Excel 启动链解释器身份与 venv 限制（实战铁证）

**适用**：加载项/RunPython/UDF 在真实 Excel 中报 `ModuleNotFoundError: No module named 'xlwings'`、`No module named '<业务模块>'`、或解释器版本与预期不符时的排查与修复。

### 5.1 现象与根因链

- **现象**：开发环境 `python -c "import xlwings"` 正常，但 Excel Ribbon 按钮 / UDF 报无模块
- **根因链**（按实测顺序）：
  1. Excel 进程环境存在 `HKCU\Environment` 级 `PYTHONHOME`/`PYTHONPATH`，指向另一份 Python（历史安装残留）——**劫持所有子进程解释器身份**
  2. 引擎（xlwingsdll / VBA RunPython）按配置表 `Interpreter_Win` 启动解释器，子进程继承 Excel 进程环境——即使配置表写的是 venv 的 `python.exe` 绝对路径
  3. **uv venv 在 Excel 启动链中无法激活**：Python 解释器身份由 EXE 决定，但当 `PYTHONHOME` 指向 base 解释器时，venv 的 site-packages **不进入 sys.path**——`sys.prefix` 回落到 base，venv 形同虚设

### 5.2 判定命令（先诊断，再修复）

```bat
:: 在 Excel 同款环境链中验证解释器身份（bat 显式调用目标 python.exe）
@echo off
"<配置表 Interpreter_Win 的完整路径>" -c "import sys, site; print('EXE=', sys.executable); print('PREFIX=', sys.prefix); print('BASE=', sys.base_prefix); print('SITE=', site.getsitepackages()); import xlwings; print('xlwings', xlwings.__version__)"
```

- `PREFIX == BASE` → **venv 未生效**（site-packages 不在搜索路径）→ 不能依赖 venv，见 5.3
- `PREFIX != BASE` 但仍无 xlwings → venv 激活正常，缺依赖 → 在 venv 内 `pip install xlwings==<版本>`
- 报 `sys` 模块错误 / 版本错乱 → PYTHONHOME 劫持 → 见 5.3 方案一

### 5.3 修复方案（按可靠度排序）

**方案一：bat 包装解释器（最可靠，推荐）**。配置表 `Interpreter_Win` 指向一个 `.bat`，bat 内显式钉死环境再调 base 解释器：

```bat
@echo off
set PYTHONHOME=D:\Python\cpython-3.13.14-windows-x86_64-none
set PYTHONPATH=D:\hermes
"D:\Python\cpython-3.13.14-windows-x86_64-none\python.exe" %*
```

- **必须** `set PYTHONHOME` 覆盖被劫持的值（清空 `set PYTHONHOME=` 亦可，取决于 base 解释器是否依赖它）
- base 解释器需自带 `xlwings`（或 PYTHONPATH 指向技能内 `xlwings/`）
- bat 文件要求：**纯 ASCII + CRLF 行尾**（中文路径/UTF-8 内容在 cmd 重解析下会被破坏）；路径含空格时整行加引号
- 配置表值示例：`Interpreter_Win = C:\...\.venv\Scripts\hermes_python.bat`（bat 自身路径可以含中文，bat 内部命令用纯 ASCII）

**方案二：junction 纯 ASCII 别名传 PYTHONPATH**。中文/长路径在 cmd 引号重解析与 GBK 代码页下会丢参数（实测 `prepare_sys_path` 参数被破坏），用 junction 建立纯 ASCII 短路径：

```bat
mklink /J D:\hermes "D:\含中文的长路径\项目工作目录"
:: 之后 PYTHONPATH 一律写 D:\hermes，模块内代码路径不受影响（物理路径未变）
```

junction 删除：`rmdir D:\hermes`（不是 rd /s，junction 不能递归删）

**方案三：不用 venv，直接 base 解释器 + 技能源码注入**（开发机场景）：

```python
# 入口模块顶部
import sys
sys.path.insert(0, r'<技能根目录>\xlwings')  # 版本一致 + 可调试
```

### 5.4 环境变量劫持排查

```bat
reg query HKCU\Environment   :: 查 PYTHONHOME / PYTHONPATH / PATH 残留
echo %PYTHONHOME%            :: cmd 当前值
```

- 存在 `PYTHONHOME` 指向非预期解释器 → 两种处理：① 修正该用户级变量（影响面大，需用户确认）；② 用方案一 bat 在加载项链内覆盖（**推荐，自包含不污染系统**）
- VBA 侧 Auto_Open 可临时清空 `PYTHONHOME`/`PYTHONPATH` 再启动引擎（白标加载项常见做法），但引擎子进程环境仍以 Excel 进程为准，最稳的还是 bat 包装


