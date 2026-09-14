# UDF 构建指南（6.5.7 聚合）

> 本文件是 SKILL.md 6.5.7（UDF 开发与导入）的扩展资源。权威语义与全部装饰器参数见 `xlwings-0.37.2/docs/udfs.md`。UDF 侧开发以本文档为准。

## 一、存放位置选择：xlsm vs xlam

- **xlsm**：函数代码随文件发送，自包含，仅该文件内可用
- **xlam**：安装后全局可用（所有工作簿可调用）；分发给未装加载项的用户显示 `#NAME?`
- 判断逻辑与纯 VBA 相同；涉及"公式自动更新/结果随数据变动"需求时优先引导 UDF

## 二、xlwings UDF 编写规范

### 函数定义：类型提示优先（v0.32.0+，推荐）

```python
import xlwings as xw
from typing import Annotated
import pandas as pd

@xw.func
def hello(name: str) -> str:
    """返回问候语"""
    return f"Hello {name}!"

# 二维数组：Annotated 替代 @xw.arg(..., ndim=2)
List2d = Annotated[list[list[float]], {"ndim": 2}]

@xw.func
def add_one(data: List2d) -> List2d:
    return [[cell + 1 for cell in row] for row in data]

# DataFrame（index/header 控制）
@xw.func
def corr_matrix(df: Annotated[pd.DataFrame, {"index": False, "header": False}]) -> pd.DataFrame:
    return df.corr()
```

**为什么类型提示优先**：不重复参数名、IDE 自动补全、符合 PEP 484、`Annotated` 别名可全局复用。装饰器版本（`@xw.arg`/`@xw.ret`）功能等价，**装饰器优先级高于类型提示**（装饰器覆盖）。

### 能力要点

- 数组公式：`ndim=2`（二维数组进出）
- 异步：`async_mode='threading'`，不阻塞 Excel重算
- `doc` 参数：函数向导显示说明
- 保留参数 `caller`（调用工作簿）、`vba`（VBA 层信息）

### UDF 入参类型注意事项（实测坑）

- **Excel 自动类型转换**：单元格中输入 `"0101"` 会被 Excel 转为数字 `101` 传入 UDF，前导零丢失。若编码含前导零（如公告类型编码），须在函数内判断 `isinstance(code, (int, float))` 后 `zfill(4)` 还原，或要求用户将输入单元格格式设为文本。
- **空单元格传入 `None`**：UDF 中须显式判断 `if x is None or x == ''`，否则 `str(None)` 返回 `"None"`。
- **错误值传播**：输入单元格含 `#N/A` 等错误时，xlwings 传入 `None`（非错误对象），函数内无法区分"空"与"错误"。

### 白标 xlam 与标准 xlwings.xlam 冲突（实测坑）

- 当目标机同时安装了 xlwings 标准加载项（`xlwings.xlam`，通常在 XLSTART）和白标 xlam 时，两者的 `xlwings` VBA 模块同名，可能导致白标 xlam 的 UDF 返回 `#NAME?`。
- **解决方案**：分发白标 xlam 时，在安装脚本中检测并提示用户禁用标准 xlwings.xlam；或在白标 xlam 中重命名 xlwings 模块（需同步修改所有内部引用）。
- 开发机调试时，若 UDF 返回 `#NAME?`，先检查 `文件 → 选项 → 加载项` 中是否同时启用了两个 xlwings 加载项。

```python
@xw.func(async_mode='threading')
def long_running(n: int, m: int) -> int:
    import time; time.sleep(5)
    return n + m

@xw.func
def my_func(caller, vba):
    """caller 返回调用工作簿，vba 返回 VBA 层信息"""
    return xw.Book(caller).name
```

## 三、导入与测试

### 开发期注册

1. Excel 中运行 `import_udfs('module_name')` 注册
2. 函数签名变更（参数名/数量/装饰器）需重新 `import_udfs()`
3. 模块代码变更自动生效；模块文件增删需点击 xlwings Ribbon 的 **Restart UDF Server** 按钮，或重启 Excel（xlwings CLI 无 `udf` 子命令）

### 调试服务器

```python
if __name__ == "__main__":
    xw.serve()  # 默认 CLSID 自动生成；可传固定值 xw.serve(clsid='{...}')
```

Excel "Add-in & Settings" 勾选 "Debug UDFs"，`Ctrl+Alt+F9` 重算即在 IDE 断点停下。

### pytest / unittest 自动化回归（两条路径互补）

- **路径 A：单元格公式断言**（真实用户行为）：写入 `=func(A1)` 公式 → `calculate()` → 断言返回值。覆盖 xlwings 转换器行为。参考骨架：`examples/excel-automated-testing-master/excel-automated-testing-master/test_mybook.py`
- **路径 B：纯 Python 单测**（不依赖 Excel）：直接 import 函数断言。覆盖业务逻辑，速度快、CI 友好
- 只用一种不够：A 验转换器、B 验逻辑

## 四、交付时引导语

- 输入公式：`=函数名(` 自动弹出函数提示
- xlam：建议通过"插入函数 (fx)"对话框（"用户定义"类别下）查找
- 自动更新：UDF 随引用单元格变化自动重算；`F9`/`Ctrl+Alt+F9` 强制重算
- 依赖提示：目标机需 Python + xlwings + 业务依赖（见部署文档/分发规则）

## 五、常见 UDF 场景示例

1. 文本处理：身份证出生日期提取、规则文本合并
2. 复杂计算：梯度奖金、多条件加权（pandas 轻松实现）
3. 增强 VLOOKUP：查找所有匹配项并合并返回
4. 数组返回：动态数组多行多列

## 六、官方 UDF 案例详细展开

以下社区版案例（不依赖 PRO/Server）由场景 C 引用，详细展开供直接借鉴：

- **`xlwings-0.37.2/examples/udf/udf.py`**（最小 UDF 集）：`@xw.sub` 的 `get_workbook_name`（宏入口）+ `@xw.func` 函数（含 numpy/pandas 转换）——UDF 编写的最小对照；`udf.xlsm` 为已装配工作簿；
- **`xlwings-0.37.2/examples/fibonacci/`**（UDF + 便携分发）：`fibonacci.py`（`@xw.func` 斐波那契）+ `build_standalone.py`（PyInstaller 冻结独立可执行 → zip 分发）+ `setup_fibonacci.py`（冻结配置）——场景 C 6.5.10 便携运行时分发的官方样例；
- **`examples/Excel_udf_itus-main/`**（金融数据查询 UDF 完整工程）：`ebitda_margins_data_udf.py`（SQLite 查询 → `@xw.func` → 单元格公式端到端）、`config.ini`（数据源/参数配置）、`scema.sql`（表结构）、`example.xlsm`（已装配演示工作簿）——"把数据库查询封装为 Excel 公式"的完整参照；`readme.md` 含安装与信任中心设置指引。

## 七、官方导航

- **`xlwings-0.37.2/docs/udfs.md`（权威，全部装饰器参数与语义）**：读 UDF 前必读
- `xlwings-0.37.2/docs/converters.md`：类型转换体系（UDF 入参/返回转换的底层）
- `examples/Excel_udf_itus-main/`：UDF 教程（类型提示/二维数组/DataFrame 实战）
- `xlwings-0.37.2/docs/threading_and_multiprocessing.md`：异步 UDF 的线程语义


