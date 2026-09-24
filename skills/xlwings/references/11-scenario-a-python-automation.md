# 场景 A：Python 脚本自动化 — 深度技术展开

> 本文件是场景 A（Python 脚本自动化）的**深度技术扩展**。主文档保留主线工作流、关键规则与 API 文档导航；本文件展开容易踩坑的技术细节、内部机制、完整操作代码与性能对比。

## 目录

- [一、chunksize 内部机制与大数据读写](#一chunksize-内部机制与大数据读写)
- [二、公式三类型的区别与选择](#二公式三类型的区别与选择)
- [三、批注操作的完整生命周期](#三批注操作的完整生命周期)
- [四、形状创建的 COM 调用细节](#四形状创建的-com-调用细节)
- [五、多线程/多进程重连模式](#五多线程多进程重连模式)
- [六、格式化进阶与性能注意](#六格式化进阶与性能注意)

---

## 一、chunksize 内部机制与大数据读写

### 1.1 为什么需要 chunksize

xlwings 通过 COM 与 Excel 通信，单次读写的数据量过大时会触发 COM 超时（通常 30 秒）或内存溢出。`chunksize` 参数让 xlwings 内部分批拉取/推送数据，但**对外返回完整结果**。

### 1.2 读取行为（关键坑）

```python
# chunksize 读取返回的是 list，不是生成器
data = sheet['A1:A100000'].options(chunksize=10000).value
# data 是完整的 list，已包含全部 100000 行
# 不需要：for chunk in data: ... （这是错误用法）
```

**本质**：`chunksize` 只控制内部 COM 调用的批次大小，不改变返回类型。源码见 `xlwings/xlwings/conversion/standard.py` 的 `ReadValueFromRangeStage.__call__`（L181-189）：按行切块逐块取 `raw_value`，再 `parts.extend(...)` 汇总成一个 list 返回。

### 1.3 写入行为

```python
# 大数据分块写入，避免单次 COM 调用超时
sheet['A1'].options(chunksize=10000).value = large_data
```

写入时同样内部分批，但需要注意：如果目标区域已有数据，分块写入不会自动清除超出范围的旧数据。

### 1.4 expand 与 chunksize 配合

```python
# 从 A1 向下扩展到数据末尾，再分块读取
data = sheet['A1'].expand('down').options(chunksize=10000).value
```

`expand('down')` 先通过 COM 获取连续数据区域的边界，然后对该区域应用 chunksize。两步都是 COM 调用，超大区域时 expand 本身也可能慢。

### 1.5 性能对比参考

批量写性能实测对比见 `examples/xlwings-demo-master/xlwings-demo-master/performance/arrays.ipynb` 与 `raw.ipynb`。结论：单次大块写入通常优于多次小块写入，但受 COM 超时限制，chunksize=10000 是经验平衡点。

---

## 二、公式三类型的区别与选择

| 属性 | 适用场景 | Excel 版本 | 行为 |
|------|---------|-----------|------|
| `.formula` | 普通公式（`=SUM(A1:A10)`） | 所有版本 | 写入后单元格显示公式，需 `app.calculate()` 触发计算 |
| `.formula2` | Dynamic array 公式（`=SORT(A1:A10)`、`=FILTER(...)`） | Excel 365 / 2021+ | 自动溢出到相邻单元格，旧版本 Excel 不支持 |
| `.formula_array` | CSE 数组公式（Ctrl+Shift+Enter） | 所有版本 | 传统数组公式，需预先选中目标区域，已被 dynamic array 取代 |

### 2.1 写入后必须触发计算

```python
sheet['B1'].formula = '=SUM(A1:A10)'
app.calculate()  # 必须调用，否则 .value 读取到的可能是旧值或 None
print(sheet['B1'].value)  # 此时才能拿到计算结果
```

**坑**：在自动化脚本中，Excel 可能处于非交互模式，公式不会自动重算。必须显式 `app.calculate()`。

### 2.2 formula2 的溢出行为

```python
# formula2 会自动溢出，只需写入左上角
sheet['C1'].formula2 = '=SORT(A1:B10, 2)'
# C1:D10 会被自动填充
```

如果溢出区域有数据，Excel 会返回 `#SPILL!` 错误。脚本中应先清理目标区域。

---

## 三、批注操作的完整生命周期

### 3.1 创建批注（必须经 COM）

xlwings 的 `.note` 属性是**只读访问器**，不能直接赋值创建批注。创建必须通过底层 COM API：

```python
# 创建批注
sheet['A1'].api.AddComment('这是批注内容')

# 读取批注
print(sheet['A1'].note.text)

# 修改批注（note.text 可写，但前提是批注已存在）
sheet['A1'].note.text = '更新后的批注'

# 删除批注
sheet['A1'].note.delete()

# 无批注时 note 为 None
assert sheet['A1'].note is None
```

### 3.2 常见错误

```python
# 错误：直接赋值 .note.text 创建批注（会抛异常，因为 note 为 None）
sheet['A1'].note.text = '新批注'  # AttributeError: 'NoneType' object has no attribute 'text'

# 正确：先 AddComment，再修改
sheet['A1'].api.AddComment('临时')
sheet['A1'].note.text = '最终内容'
```

---

## 四、形状创建的 COM 调用细节

### 4.1 为什么 sheet.shapes 没有 add()

xlwings 的 `Shapes` 集合类只实现了索引访问（`__getitem__`）和 `count`，没有封装 `AddShape`。创建形状必须下沉到 COM：

```python
# AddShape(type, left, top, width, height)
# type 常量（MSO_SHAPE_TYPE）：
#   1 = msoShapeRectangle（矩形）
#   5 = msoShapeRoundedRectangle（圆角矩形）
#   9 = msoShapeOval（椭圆）
shape = sheet.shapes.api.AddShape(1, left=100, top=50, width=150, height=80)
```

### 4.2 形状的访问与修改

```python
s = sheet.shapes[0]
s.name           # 形状名称（可写）
s.left, s.top    # 位置（磅）
s.width, s.height # 尺寸
s.delete()       # 删除
```

### 4.3 形状类型常量

完整的 MSO_SHAPE_TYPE 枚举有 100+ 种，常用的：矩形(1)、圆角矩形(5)、椭圆(9)、三角形(7)、菱形(4)、箭头(33/34)、标注(51+)。需要时查 `xlwings/xlwings/constants.py` 或微软文档。

---

## 五、多线程/多进程重连模式

### 5.1 为什么不能传递 xlwings 对象

xlwings 对象底层持有 COM 指针。COM 对象是 STA（单线程单元）模型，跨线程传递会导致 "Application is busy" 或崩溃。

### 5.2 正确模式：子线程内重新连接

```python
import xlwings as xw
from concurrent.futures import ThreadPoolExecutor

def process_sheet(filepath, sheet_name):
    # 子线程内重新创建连接，不使用主线程的 xw.Book 对象
    with xw.App(visible=False) as app:
        wb = app.books.open(filepath)
        data = wb.sheets[sheet_name]['A1:Z1000'].value
        wb.close()
    return data

# 多进程仅 Windows 支持（macOS 的 AppleScript 不支持多进程）
with ThreadPoolExecutor(max_workers=4) as executor:
    results = list(executor.map(process_sheet, files, sheet_names))
```

### 5.3 多进程注意

- 仅 Windows 支持多进程 xlwings（macOS 因 AppleScript 限制不支持）
- 每个子进程启动独立 Excel 实例，资源开销大
- 大量小任务用多线程（每个线程独立连接），CPU 密集型任务才考虑多进程

详细说明见 `xlwings/docs/threading_and_multiprocessing.md`。

---

## 六、格式化进阶与性能注意

### 6.1 formatter 选项

写入时用 `formatter` 选项自动应用格式化函数，避免写入后逐单元格设置样式：

```python
def color_negative(val):
    return val  # formatter 函数接收值，返回格式化后的值

sheet['A1'].options(formatter=color_negative).value = data
```

### 6.2 批量格式化优于逐单元格

```python
# 好：一次性设置整个区域
sheet['A1:Z1000'].color = (255, 255, 200)
sheet['A1:Z1000'].number_format = '0.00'

# 差：逐单元格设置（COM 调用次数 = 单元格数，极慢）
for cell in sheet['A1:Z1000']:
    cell.color = (255, 255, 200)
```

### 6.3 autofit 的开销

`sheet['A:Z'].autofit()` 内部需要遍历每列测量内容宽度，大表时很慢。建议只对必要列调用，或在所有数据写入完成后一次性调用。

### 6.4 子字符串格式化

`.characters` 支持对单元格内部分文本设置字体样式，但 **macOS 不支持**：

```python
# Windows  only
sheet['A1'].value = '总计：1000'
sheet['A1'].characters[5:].font.bold = True  # "1000" 加粗
```

---

## 七、社区版案例详细展开（场景 A 侧）

以下社区版案例（不依赖 PRO/Server）由场景 A 引用并介绍，详细展开供直接借鉴：

### 7.1 官方 quickstart 迷你示例（xlwings/examples/，BSD）

- **`database/database.py`**（SQLite → Excel 回填）：`sqlite3` 连接 `chinook.sqlite` → 查询结果 → `xw.Range` 写入指定工作表——外部数据源接入 Excel 的最小闭环（场景 A 数据接入参考）；`database.xlsm` 为已装配演示工作簿；
- **`mpl/mpl.py`**（matplotlib → Excel）：matplotlib/seaborn 绘图 → `xw` 写入图表（`plots` 函数）——场景 A 图表输出样例；`mpl.xlsm` 演示工作簿；
- **`simulation/simulation.py`**（numpy 蒙特卡洛 → Excel）：纯计算（`simulation` 函数）→ 回填 Excel——场景 A 计算回填样例；与 `simulation-demo-master`（Excel/Web 双端）的计算层同构；
- **`build_lite.py`**：把 5 个示例目录打包为 `_build/*.zip`（xlwings Lite 运行时加载的示例包）——Lite/Server 场景的示例组织方式。

### 7.2 excel-automated-testing-master（自动化测试骨架）

- `mybook.py`：被测脚本（最小 xlwings 操作集，写值/读值）；
- `test_mybook.py` + `test_sample.py`：pytest 驱动 xlsm 回归——启动 Excel 实例 → 执行脚本 → 断言单元格值 → 收尾清理（COM 测试 fixture 模式）；
- `.gitlab-ci.yml`：CI 接入样例（无头环境跑 pytest）；
- 用途：场景 A 脚本的自动化测试参照（对应自动化测试环节的测试门禁）。

### 7.3 产物质检与报表交叉核验（static-excel-test + cross-check-reports）

- **`static-excel-test-master/test_refs.py`**：openpyxl 扫描 `#REF!`/`#DIV/0!`/`#VALUE!` 坏引用——交付前对 xlsx 静态产物做质量门禁，全程无需 Excel 实例；
- **`cross-check-reports-main/test_reports.py`**：两份同构报表逐单元格断言、差异定位与汇总——报表交叉核验模板；
- 执行流程（明确目标 → 静态扫描 → 交叉核验 → 汇总结论）与仓库职责见 `examples/ReadMe.md` 质检小节。

### 7.4 Excel/Web 双端原型（simulation-demo-master）

- `simulation.py`：纯计算逻辑层（蒙特卡洛模拟），无 Excel/Web 依赖；
- `xlwings_app.py`（Excel 端）与 `web_app.py`（Flask Web 端）：分别调用同一份 `simulation.py`——"计算逻辑层与表现层分离、同一份代码驱动两端"的架构样例（场景 A 架构设计参考实现）。

### 7.5 Power Query 库管理（xl-pq-handler-master）

- `xl_pq_handler` 包：从工作簿批量提取 M 代码为 `.pq` 脚本库、按依赖顺序注入回工作簿——`PQManager` 编程接口 + UI 两种形态，第三方依赖 `pip install xl-pq-handler`；
- 场景 A 中把"Power Query 查询"纳入自动化管线的参考（提取/注入两端均以 xlwings 驱动 Excel）。

### 7.6 Eikon 数据接入（社区版主链路）

- 行情 API → DataFrame → Excel 写入的完整链路；含相关性分析、蒙特卡洛模拟、实时数据流等五种接入形态；
- 前置：可用的 Eikon 环境（桌面终端或 API 授权）+ `eikon` 包；
- **PRO Reports 部分**（`_report_template.py`/`sample3.py` 的 `create_report`）属场景 D，见 `13-scenario-d-source-code.md` 四章 PRO 案例研读。

### 7.7 xlwings-demo-master 社区版实战参考（basics/correlation/frozen/interactive/performance/simulation/testing）

`xlwings-demo-master/` 的社区版部分由场景 A 4.3 引用介绍，文件级用途详展如下（与 `13-scenario-d-source-code.md` 四章教学演示选材流程互补；PRO Reports 部分见 13-d 四章 PRO 案例研读）：

- **`basics/xwdemo.py`**：最小脚本骨架 + Range 读写 + `@xw.sub` 宏入口——场景 A 脚本骨架参考（数据读写与转换环节）；
- **`correlation/correlation.py`**：DataFrame → 散点图、图表对象操作、图片导出——图表与可视化参考；
- **`performance/arrays.ipynb` + `raw.ipynb`**：chunksize 实测对比、数组批量写入 vs 逐单元格、COM 超时边界——大数据性能参考；
- **`performance/udfs.xlsm` + `arrayudf/`**：`@xw.func` UDF 定义、数组 UDF、公式重算触发——公式与计算参考；
- **`interactive/interactive.ipynb`**：Jupyter 交互（`xw.view()`/`xw.load()`）——交互式工作参考；
- **`frozen/demo.py` + `demo.spec` + `demo.xlsm`**：PyInstaller 冻结部署样例——`demo.spec` 冻结配置、`RunFrozenPython` 运行方式（PyInstaller 冻结部署机制的配套演示）；
- **`simulation/simulation.py` + `simulation.xlsm`**：蒙特卡洛模拟 UDF（numpy 批量计算回填）——数值计算 UDF 参考（与 7.4 simulation-demo-master 双端原型的单 Excel 版对应）；
- **`testing/mybook.py` + `test_mybook.py` + `test_sample.py` + `.gitlab-ci.yml`**：官方 pytest 测试骨架——`mybook.xlsm` 工作簿夹具、Excel 交互测试写法、CI 配置参考（关联 7.2 excel-automated-testing）。

## 可配套阅读

- `xlwings/docs/`（官方文档权威参考）

## 附：场景 A 实战案例索引（examples/）

以下案例均为内置可运行的完整工程，按场景 A 工作流环节分类，遇到对应任务时直接参考其代码结构与实现方式：

| 工作流环节 | 推荐案例 | 核心参考点 |
|-----------|---------|-----------|
| **数据读写与转换** | `xlwings-demo-master/basics/xwdemo.py` | 最小脚本骨架、Range 读写、@xw.sub 宏入口 |
| **大数据量性能** | `xlwings-demo-master/performance/arrays.ipynb` + `raw.ipynb` | chunksize 实测对比、数组批量写入 vs 逐单元格、COM 超时边界 |
| **图表与可视化** | `xlwings-demo-master/correlation/correlation.py` | 从 DataFrame 生成散点图、图表对象操作、图片导出 |
| **公式与计算** | `xlwings-demo-master/performance/udfs.xlsm` + `arrayudf/` | @xw.func UDF 定义、数组 UDF、公式重算触发 |
| **外部数据接入** | `xlwings-eikon-master/` | 行情 API 调用 → DataFrame → Excel 写入的完整链路；含相关性分析、蒙特卡洛模拟、实时数据流五种形态 |
| **Excel/Web 双端** | `simulation-demo-master/simulation.py` + `xlwings_app.py` + `web_app.py` | 纯计算逻辑层与表现层分离：同一份 `simulation.py` 同时被 Excel 端和 Flask Web 端调用 |
| **Power Query 管理** | `xl-pq-handler-master/src/xl_pq_handler/` | 从工作簿批量提取 M 代码、按依赖顺序注入回工作簿、PQManager 编程接口 |
| **自动化测试** | `excel-automated-testing-master/test_mybook.py` + `test_sample.py` | pytest 驱动 xlsm 工作簿回归测试、COM 测试 fixture 模式、断言单元格值 |
| **静态质检** | `static-excel-test-master/test_refs.py` | openpyxl 扫描 #REF!/#DIV/0!/#VALUE! 坏引用，全程无需 Excel 实例 |
| **报表交叉核验** | `cross-check-reports-main/test_reports.py` | 两份同构报表逐单元格断言、差异定位与汇总 |
| **UDF 完整工程** | `Excel_udf_itus-main/ebitda_margins_data_udf.py` + `example.xlsm` | SQLite 数据查询 → @xw.func 封装 → 单元格公式调用的端到端 UDF 工程 |
| **系统入门学习** | `python-for-excel-course-main/`（0-Intro 至 7-Part7） | 2018 官方入门课程配套 Jupyter 讲义，按主题分 Part，适合查漏补缺 |

依赖 xlwings PRO 的案例（`xlwings-factsheet-demo-main`、demo 的 `reader`/`restapi`/Reports 四件套、eikon 的 Reports 部分）索引见 `13-scenario-d-source-code.md` 四章案例总览与 PRO 案例研读，不在本索引。

**使用原则**：案例代码是参考实现，不是模板——提取其结构与关键模式，结合当前任务需求重新实现，不要直接复制粘贴（路径、依赖、业务逻辑均不同）。
