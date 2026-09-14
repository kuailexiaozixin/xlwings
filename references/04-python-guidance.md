# Python 构建指南（6.5.4 扩展与导航）

> 本文件是 SKILL.md 6.5.4（自足模板：双形态×双入口、装饰器模板、分层铁律、业务核心规范、编码铁律、三阶段门禁）的**扩展资源**——官方文档导航 + 官方示例导航 + 进阶主题提炼。规则与模板以 6.5.4 正文为准，本文件不重复。

## 一、官方文档导航（`xlwings-0.37.2/docs/`，Python 侧权威）

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
| `xlwings-server-main/` | xlwings Server（服务端形态，393 文件，按需深读） |

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


