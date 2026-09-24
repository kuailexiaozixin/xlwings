# 场景 D：xlwings PRO 深度技术分析与源码扩展

> 本文件是场景 D（源码学习与二次开发）的**深度扩展资源**。主文档保留源码架构导航、子场景路由与核心索引；本文件聚焦 xlwings PRO 的深度技术分析（许可证机制、四种引擎原理、Reports 架构、部署机制）、**xlwings 核心源码研读（xlwingsdll 原生桥 + xlwings Python 包架构）**以及入门学习执行步骤与案例总览（社区版案例详展在场景 A/C references，本文件不重复）。**xlwings Lite（免费 Web 加载项）的独立参考手册见 `references/15-xlwings-lite-guidance.md`（原 1.6/4.5 已迁入并指针化）**。

## 目录

- [一、xlwings PRO 深度技术分析](#一xlwings-pro-深度技术分析)
  - [1.1 许可证机制](#11-许可证机制)
  - [1.2 四种引擎深度分析与数据通道选型](#12-四种引擎深度分析与数据通道选型)
  - [1.3 Reports 报告系统架构](#13-reports-报告系统架构)
  - [1.4 PRO 部署机制：code embed 与 release（指针 → 10 号附录 A）](#14-pro-部署机制code-embed-与-release指针--10-号附录-a)
  - [1.5 caller 机制与 ObjectHandle](#15-caller-机制与-objecthandle)
  - [1.6 xlwings Lite（指针）](#16-xlwings-lite指针)
  - [1.7 xlwings Server 官方工程技术研读](#17-xlwings-server-官方工程技术研读)
   - [1.8 Office.js 引擎（指针，详见 16 号文档）](#18-officejs-引擎指针详见-16-号文档)
- [二、xlwings 核心源码研读（xlwingsdll + xlwings 包）](#二xlwings-核心源码研读xlwingsdll--xlwings-包)
  - [2.1 xlwingsdll 原生 DLL](#21-xlwingsdll-原生-dll)
  - [2.2 xlwings Python 包核心架构](#22-xlwings-python-包核心架构)
  - [2.3 构建与测试体系（scripts / src / tests）](#23-构建与测试体系scripts--src--tests)
- [三、入门学习与教学演示（执行步骤）](#三入门学习与教学演示执行步骤)
  - [3.1 课程仓库技术研读（python-for-excel-course-main）](#31-课程仓库技术研读python-for-excel-course-main)
  - [3.2 澄清范围与选材](#32-澄清范围与选材)
  - [3.3 执行讲义](#33-执行讲义)
  - [3.4 执行演示并输出小结](#34-执行演示并输出小结)
- [四、案例总览与 PRO 案例研读](#四案例总览与-pro-案例研读)
  - [4.1 案例总览与依赖矩阵](#41-案例总览与依赖矩阵)
  - [4.2 xlwings-demo-master 子目录地图（13 项全览）](#42-xlwings-demo-master-子目录地图13-项全览)
  - [4.3 教学演示选材与执行](#43-教学演示选材与执行)
  - [4.4 xlwings PRO 案例研读（依赖 PRO Reports / Reader / REST API）](#44-xlwings-pro-案例研读依赖-pro-reports--reader--rest-api)
  - [4.5 taxi-duckdb（指针）](#45-taxi-duckdb指针)
  - [4.6 数据服务层案例研读：SQLite 数据访问模式（SQL）](#46-数据服务层案例研读sqlite-数据访问模式sql)
  - [4.7 DuckDB 桌面端数据访问模式（taxi-duckdb 桌面端分支）](#47-duckdb-桌面端数据访问模式taxi-duckdb-桌面端分支)

---

## 一、xlwings PRO 深度技术分析

xlwings PRO 为双许可（PolyForm Noncommercial 1.0.0 / 商业许可），源码位于 `xlwings/xlwings/pro/`。所有 PRO 模块在导入时调用 `LicenseHandler.validate_license("pro")`，无有效许可抛 `xlwings.LicenseError`。

### 1.1 许可证机制

**LicenseHandler**（`pro/utils.py`）是 PRO 许可的核心验证器，采用三级密钥查找 + Fernet 对称加密：

1. **三级密钥查找**（`get_license()` 按优先级）：
   - 环境变量 `XLWINGS_LICENSE_KEY`（UDF 调用时也走此路径，因为 UDF 无法访问配置表）
   - 工作簿配置表中的 `LICENSE_KEY` 键（仅 RunPython 调用时可用，通过 `read_config_sheet(Book.caller())` 读取）
   - 用户配置文件 `xlwings.conf`（`%USERPROFILE%\.xlwings\xlwings.conf`，CSV 格式，查找 `"LICENSE_KEY"` 行）
   - 三级均未找到则抛 `LicenseError("Couldn't find an xlwings license key.")`

2. **Fernet 加密验证**（`get_cipher()`）：
   - 从环境变量 `XLWINGS_LICENSE_KEY_SECRET` 获取密钥
   - 使用 `cryptography.fernet.Fernet` 解密许可证
   - 解密失败抛 `LicenseError("Couldn't validate xlwings license key.")`
   - `cryptography` 包未安装时 `Fernet = None`，验证必然失败

3. **缓存验证**（`@lru_cache()`）：`validate_license(product, license_type)` 结果被缓存，避免重复解密开销。验证内容包括产品类型（pro/reports 等）、许可证有效期、使用范围。

**关键设计决策**：
- UDF 与 RunPython 的密钥查找路径不同——UDF 运行在独立进程中无法访问 `Book.caller()`，因此只能用环境变量；RunPython 可访问配置表。这是配置表 `LICENSE_KEY` 键在 UDF 中不生效的根本原因。
- `@lru_cache()` 意味着运行中更换许可证需要重启 Python 进程才能生效。

### 1.2 四种引擎深度分析与数据通道选型

xlwings 0.37.4 通过 `xw.engines` 管理四种引擎，引擎选择决定 `Book()` 的连接方式与能力边界：

| 引擎 | 许可 | 平台 | 实现文件 | 核心原理 |
|------|------|------|---------|---------|
| `excel` | 免费 | Windows/macOS | `_xlwindows.py`/`_xlmac.py` | 本地 COM/appscript 驱动桌面 Excel |
| `remote` | PRO | 全平台 | `pro/_xlremote.py` | REST API 连接 xlwings Server |
| `calamine` | PRO | 全平台 | `pro/_xlcalamine.py` | Rust calamine 库只读解析，无需 Excel |
| `officejs` | PRO | 全平台 | `pro/_xlofficejs.py` | Office.js API，Web 加载项环境；完整研读见 `references/16-xlwings-officejs.md` |

#### 1.2.1 数据通道选型（读取侧决策）

场景 A"读取数据"步骤的通道决策：读写通道不必全用 COM，按数据来源与去向选：

| 通道 | 能力 | 适用 |
|------|------|------|
| COM（内置 pywin32） | 读写活会话 / 写回 / 公式重算 / 宏 | 一切与"活 Excel"交互的场景（唯一选择） |
| `python-calamine` | 只读，免 Excel，毫秒级 | 磁盘大 xlsx/xlsm/xlsb 纯读取、结构侦察 |

决策规则：

1. 目标是**活会话**（用户正打开、未保存更改、要跑宏）→ COM 唯一
2. 纯**读磁盘文件**且量大 → `python-calamine`（Rust 快通道，毫秒级提取）
3. **无 Office 环境**（CI/服务器/批量 ETL）→ 读取侧 calamine，写回另选（openpyxl/生成文件）
4. **写回 + 公式 + 格式 + 图表** → COM 唯一
5. 混合模式（推荐）：文件级侦察用 calamine 秒看结构/前 N 行 → 需写/活会话时再 COM

通道与引擎实现的对应：COM 通道 = excel 引擎（xlwingsdll/COM 直连，见 1.2.2）；只读快通道 = calamine 引擎（`pro/_xlcalamine.py` 适配层 + Rust `xlwingslib`，见 1.2.4）与社区版 `python-calamine` 包（同一 Rust 库的独立 PyPI 发行）。

#### 1.2.2 excel 引擎（本地桌面）

- **Windows**：`_xlwindows.py` 通过 pywin32 调用 Excel COM 接口；`_win32patch.py` 修补 COM 行为差异（如 `Range.Value` 的数组返回、日期处理）
- **macOS**：`_xlmac.py` 通过 appscript 调用 AppleScript；`mac_dict.py` 维护 AppleScript 字典映射
- **特点**：完整读写能力，支持 UDF/RunPython/Ribbon；依赖本地 Excel 安装；单线程限制

#### 1.2.3 remote 引擎（xlwings Server）

- **实现**：`pro/_xlremote.py`（5266 行完整客户端引擎），通过 HTTP REST API 与 xlwings Server 通信
- **细节去向（单点原则）**：客户端"如何连 Server"的完整机制（JSON 动作协议、lazy load、计算模式映射、颜色规范化、版本校验、对象覆盖）已迁移至 `references/14-xlwings-server-guidance.md` 阶段一 1.6，本处不再展开
- **对比要点**：与 excel/calamine/officejs 三引擎并列，属"服务端无 Excel 环境、Linux 部署、Google Sheets 支持"通道

#### 1.2.4 calamine 引擎（Reader，只读）

- **实现**：`pro/_xlcalamine.py`，通过 `xlwings.xlwingslib`（Rust 编译的 Python 扩展）调用 calamine 库
- **核心能力**：无需安装 Excel，直接解析 .xlsx/.xls/.xlsb/.xlsm 文件的单元格值
- **值清理**：`clean_value_data()` 处理日期类型转换（`datetime_builder`）、空值（`empty_as`）、数字类型（`number_builder`）、错误值转字符串（`err_to_str`）
- **限制**：只读（`mode='r'`），不支持公式计算、格式修改、UDF、RunPython；`MAX_ROWS=1_048_576`、`MAX_COLUMNS=16_384`
- **适用场景**：CI/CD 环境数据读取、大数据量快速解析、无 Excel 服务器

##### 1.2.4.1 `_xlcalamine.py` 引擎层深度研读（Python 适配层）

`pro/_xlcalamine.py`（561 行）是 calamine 引擎的 **Python 侧适配层**：不直接碰 Rust，而是继承 `base_classes` 的对象骨架，把 `xlwingslib`（Rust pyo3 绑定）的四个底层函数包装成完整的 xlwings 对象 API（Apps/App/Book/Sheets/Sheet/Range/Names/Name），实现"免 Excel 读文件但高层 API 不变"。

**引擎身份与值归一化钩子**
- `Engine` 单例（模块尾 `engine = Engine()`）：`name="calamine"`、`type="reader"`，暴露 `apps`；`App.engine` 返回该全局实例——应用层 `xw.App(engine="calamine")` 即路由到这里。
- `clean_value_data()`：空值→`empty_as`；datetime 按 `datetime_builder` 重建（字段逐项重传、`tzinfo=None` 清零）；float 按 `number_builder` 转换。`err_to_str` 因性能原因下沉到 `raw_value` 层处理，不在这一层扫全表。
- `clean_value_data()` 的**短路优化**：当 `empty_as`/`number_builder` 均为默认且 `datetime_builder` 即 `dt.datetime` 时，直接返回原数据、不逐格扫描——默认路径零开销；`_clean_value_data_element` 内 `number_builder` 只作用于 `float`（整数 / 其他类型原样过）。
- `prepare_xl_data_element(x, options)` **原样返回**——只读引擎无写入侧转换钩子，是"引擎接口契约完整、能力按需空实现"的样板（对比 excel 引擎该钩子做类型编码）。
- 常量 `MAX_ROWS=1_048_576`、`MAX_COLUMNS=16_384` 即 Excel 网格上限；`Sheet.cells` = `Range((1,1),(MAX_ROWS,MAX_COLUMNS))`，是"全表"判定的哨兵值。

**对象树适配（模拟 Excel 进程）**
- `Apps`：单元素列表包一个 `App`；`add()` 头部插入新 App。
- `App`：**虚拟 pid**（类变量 `_next_pid=-1` 递减为 `-1、-2…`），`visible=True` 恒真、`activate()` 空操作——纯模拟无真实进程；`kill()` 从 apps 列表移除并断开引用。`add_book=True` 时默认建一个 `path="dummy"`、仅含 `Sheet1` 的空 Book（模仿 Excel 新建空白工作簿）。
- `Books.open(filename)`：**打开即全量读元数据**——`get_sheet_names()` 取 sheet 列表、`get_defined_names()` 取定义名称；名称按 `ref.split("!")[0].strip("'")` 是否命中 sheet 名过滤出**工作表级名称**，组装成 `Book.api={"sheet_names":…,"names":[…]}` 注入字典；`names` 项带 `sheet_index`、`address`、`book_scope=True`（TODO 注释：calamine 不提供书级作用域信息，先硬编码）。
- `Book`：`name`=文件 basename、`fullname`=绝对路径；`close()` 用 `assert self.api is not None` 防重复关闭，随后置空引用。

**Sheet / Range 寻址与坐标换算**
- `Sheets(name_or_index)`：str → 查 `book.api["sheet_names"]`（不存在抛 `NoSuchObjectError`）；数字 → 直接作 1-based 索引。`Sheet.index` 全程 1-based。
- `Range` 构造把四种入参归一为 `arg1/arg2`（1-based 元组）：
  - 两个 Range → 并集边界（各自 coords 取 min/max）；
  - A1 字符串 → `utils.a1_to_tuples`；未命中时在书级 `names` 里按名称 + `sheet_index` 双重匹配解析命名区域，仍不中抛 `NoSuchObjectError`；
  - 4 元组 `(row, col, nrows, ncols)` → 展开为对角坐标；
  - 直接坐标元组。
- `raw_value`（读取核心）：
  - **全表**（arg2 命中 MAX_ROWS/MAX_COLUMNS，即 `sheet.cells`）：`get_sheet_values(fullname, sheet_index-1, err_to_str)` 一次读全表，**缓存**到 `sheet._api["values_err_to_str_{err_to_str}"]`——同表多个 Range / `end()` 复用只读一次；`err_to_str` 作为缓存键的一部分。
  - **局部范围**：`get_range_values(fullname, sheet_index-1, (r0,c0), (r1,c1), err_to_str)`，入参为 **0-based**（arg1/arg2 各减 1）——坐标换算只出现在这一个边界。
- `Range.api` property 即 `raw_value`——**读结果直接充当 api 句柄**：上层 `main.Range` 只认 `.api` 拿值，与 excel 引擎（COM 对象为 api）形态一致，是"免环境适配层"接口对缝的关键。
- `coords`=`(sheet名, row, col, nrows, ncols)`；`shape` 由对角坐标相减 +1；`address` 生成 `$A$1:$C$3` 绝对引用（单格省略冒号）。
- `end(direction)`：**Ctrl+方向键语义**——基于（缓存的）全表值沿 down/up/right/left 四方向扫描第一个空值或越界点，返回相邻单元格 Range；`IndexError` 兜底（已用区域外视为空）。因此 `Range("A1").end("down")` 不需 Excel 即可实现选区边界定位。
- `Range.__call__`：标量 → row-major 线性索引（1-based）映射到 (row,col)；二元 → 相对偏移 `(row+arg1-1, col+arg2-1)`。

**Name / Names**
- 仅实现书级（parent 传 Book；注释 "only implemented for Book, not Sheet"）；`refers_to` 从 `sheet_index`+`address` **重建公式文本** `=Sheet!$A$1`（sheet 名含空格自动加单引号）；`refers_to_range` 解析回 Range。

**对构建应用系统的工程借鉴**
- **api 字典注入模拟**：真实引擎句柄（此处是 Rust 返回值）统一装进 `api` dict，上层对象只依赖字段名——可插拔、可测试，是"免环境适配层"的通用模板。
- **惰性全表缓存**：一次读全表、按 `err_to_str` 分键缓存，多 Range / `end()` 复用，避免大数据重复 IO。
- **坐标 0-based/1-based 集中换算**：底层读写 0-based、对外 API 全程 1-based，换算只在 `raw_value` 边界出现。
- **选项经 `options` 字典透传**：`Range.options` 由主包 `main.Range` 赋值（`err_to_str` 等），`sheet.range()` 签名保持干净。
- **错误语义分层**：对象不存在抛 `NoSuchObjectError`、名称查找失败抛 `KeyError`、重复 close 用 assert 拦截。
- **只读适配的价值**：CI / ETL / 侦察场景用同一套高层 API 免 Excel 读盘，无 COM 启动开销、无平台依赖；代价是写入 / 公式 / 格式 / 图表 / UDF 全不支持（`has_array` 返回 `None`，仅 legacy CSE 数组的 expansion.py 依赖它）。

底层 Rust 层见 2.3.2 `src/`（`lib.rs`：`CellValue` 枚举逐类转 Python 对象、`get_values(used_range)` 读区域、`CalamineError` → `XlwingsError` 透传）。

#### 1.2.5 officejs 引擎（Web 加载项）
**定位**：`pro/_xlofficejs.py`（160 行，值转换层）+ `pro/udfs_officejs.py`（1254 行，UDF/脚本全链路），基于 Office.js API 的 Excel Web 加载项通道——**Server 与 Lite 共享的语义内核**。
- **细节去向（单点原则）**：完整源码研读（值转换层 / UDF / 脚本 / socket.io 会话 / 测试）与官方教程对照已独立为 `references/16-xlwings-officejs.md`，本处不再展开；与桌面 UDF（2.2.2）同源异路
- **细节去向（单点原则）**：完整源码研读（值转换层 / UDF / 脚本 / socket.io 会话 / 测试）与官方教程对照已独立为 `references/16-xlwings-officejs.md`，本处不再展开

### 1.3 Reports 报告系统架构

Reports 是 xlwings PRO 的核心增值功能，实现"Excel 模板 + 结构化数据 → 批量报告"的三段式流水线。源码位于 `pro/reports/`。

**模块组成**：

| 文件 | 职责 |
|------|------|
| `main.py` | 核心引擎：Jinja 模板解析、占位符替换、数据注入 |
| `filters.py` | 模板过滤器（排序、聚合、列裁剪等） |
| `markdown.py` | Markdown 渲染（`Markdown`/`MarkdownStyle` 类） |
| `image.py` | 图片处理（`Image` 类，支持 PIL/matplotlib/plotly） |
| `pdf.py` | PDF 导出 |

**核心机制**：

1. **Jinja 模板解析**（`parse_single_placeholder()`）：
   - 模板单元格内嵌 Jinja 变量语法（`{{ var }}`）与过滤器链（`{{ var | filter1 | filter2 }}`）
   - 使用 `jinja2.Environment` 与 `jinja2.nodes` 解析 AST，识别占位符位置
   - 支持在 Excel 单元格注释中定义模板变量（避免污染显示内容）

2. **数据注入流水线**：
   - 读取模板工作簿 → 解析占位符 → 从传入的 `**data` 中取值 → 应用过滤器链 → 写入对应单元格
   - 支持 DataFrame 自动扩展（行数动态增长）、图片插入、Markdown 渲染
   - `app.render_template(template=, output=, **data)` 是入口 API

3. **过滤器链**（`filters.py`）：
   - 排序：`{{ df | sort('col') }}`
   - 聚合：`{{ df | groupby('col') | sum }}`
   - 列裁剪：`{{ df | columns(['a','b']) }}`
   - 过滤器可链式组合，在模板单元格中声明

4. **Markdown 渲染**（`markdown.py`）：
   - `Markdown` 类将 Markdown 文本转为 Excel 富文本（加粗/斜体/列表/链接）
   - `MarkdownStyle` 控制字体、颜色、间距等样式
   - 适用于报告中的叙述性段落

5. **图片支持**（`image.py`）：
   - 支持 PIL Image、matplotlib Figure、plotly Figure 三种来源
   - 自动转换为 Excel 可嵌入的图片格式
   - 与模板占位符配合：`{{ chart }}` 传入 Figure 对象即自动插入

**三段式流水线**（以 factsheet 为例，见 `examples/xlwings-factsheet-demo-main/`）：
1. **模板准备**：Excel 模板中用 Jinja 语法声明变量位置与过滤器
2. **数据准备**：结构化数据（DataFrame/dict）作为 `**data` 传入
3. **渲染输出**：`render_template()` 批量生成独立报告文件

### 1.4 PRO 部署机制：code embed 与 release（指针 → 10 号附录 A）

> 已迁移整合至 `references/10-deployment-delivery.md` 附录 A（部署与交付工作流的 PRO 对照研读）。原内容（源码深度：`cli.py:code_embed()` / `pro/embedded_code.py:dump_embedded_code()` / `addin/Main.bas` 检测逻辑与 LICENSE_KEY 检查 / `cli.py:release()` / Deploy Key / RELEASE_NO_ADDIN / COM 保存 Ribbon 回注 / 部署方式对比，含官方文档映射）在 10 号附录 A 完整保留。
> **定位**：PRO 部署机制仅作对照研读、不采用；开发加载项部署走自研 release_tool（10 号正文主线，无需 PRO 许可证）。

### 1.5 caller 机制与 ObjectHandle

**Caller 机制**（`pro/caller.py`）：

- 自定义函数（custom functions）只接收参数，不接收工作簿数据；Excel 发送调用单元格地址字符串
- `Caller` 类型提示建模这个引用，解析正则支持 `[Book1.xlsx]Sheet1!B21`、`Sheet1!B21`、`'My Sheet'!A1`、`$A$1:$C$3` 等格式
- 放在 `xlwings.pro` 而非上层框架，确保每个 custom-functions 运行时（xlwings Server、xlwings Lite/Pyodide）都能访问

**ObjectHandle 机制**（`pro/object_handles.py`）：

- Excel 与 Python 间传递自定义对象引用（非基本类型）
- UDF 返回自定义对象时，ObjectHandle 生成句柄字符串存入单元格，后续 UDF 可通过句柄取回对象
- 解决 Excel 单元格只能存基本类型的限制，支持链式 UDF 调用（如 `=GetData() → FilterData(A1) → PlotData(B1)`）

### 1.6 xlwings Lite（指针）

**定位**：xlwings Lite 是嵌入工作簿的轻量版本（Office Web 加载项，运行在 Pyodide 浏览器端 Python 中），免费、跨平台、目标机零安装。

**细节去向（单点原则）**：Lite 的全部内容（产品定位/能力全景/适用边界、Pyodide 运行机制、信任模型、自定义函数/脚本/Notebooks/异步 API/数据与网络开发工作流、测试调试、App Mode 与分发、自托管部署、taxi-duckdb 完整案例研读、`BookAsync`/`base_classes.py` 异步方法族/`ext/sql.py` 等源码研读）已独立为以**开发生命周期工作流**编排的参考手册：

- 读 Lite 相关内容一律进入 **`references/15-xlwings-lite-guidance.md`**；
- 本处不再重复任何 Lite 细节。

**与本章的关系**：1.8（Office.js 引擎 `pro/_xlofficejs.py` + `udfs_officejs.py` 全链路）是 PRO 侧的 Web 通道实现，Lite 复用其转换层与 UDF/脚本管线（含 `streaming_callback` 的 Lite/Pyodide 直推分支）——该部分属 PRO 源码研读，已独立为 `references/16-xlwings-officejs.md`；Lite 产品侧如何用这些能力见 15 号手册。

### 1.7 xlwings Server 官方工程技术研读

**定位**：把 xlwings 能力服务化的官方开源参考工程——Excel 与 Google Sheets（含 Excel on the web）经浏览器调用服务端 Python（FastAPI + Socket.IO + Redis + Office.js + 可选 WASM/Pyodide），目标机无需安装 Python。与桌面版（场景 A/C）互补；source-available 双许可（PolyForm Noncommercial / PRO EULA，见 `LICENSE.md`），**仅供源码研读借鉴设计，不部署、不抄码**。

**细节去向（单点原则）**：xlwings Server 的完整六块研读（`docs/` 61 篇官方教程、`xlwings_server/` 服务端包、`deployment/`、`nginx/`、`scripts/`、`tests/`）已独立为以**开发生命周期工作流**编排的参考手册：

- 定位与选型、初始化与本地开发环境、应用装配与扩展机制、自定义函数、自定义脚本与任务窗格、认证授权与安全、测试调试与性能、部署运维与升级——**全部分阶段详展见 `references/14-xlwings-server-guidance.md`**
- 本处不再重复任何 server 细节；读 server 相关内容一律进入 14 号手册

**与 1.2.3 的关系**：1.2.3 的客户端侧 remote 引擎细节（`pro/_xlremote.py`，如何连 Server）已一并迁移至 14 号手册阶段一 1.6；14 号手册其余章节是 Server 服务端如何工作。
### 1.8 Office.js 引擎（指针，详见 16 号文档）

**定位**：`xlwings/pro/_xlofficejs.py`（160 行）+ `xlwings/pro/udfs_officejs.py`（1254 行）共同实现 xlwings PRO 的 **Office.js 自定义函数（Custom Functions）与自定义脚本（Custom Scripts）**——remote 类型引擎，走 socket.io 推流（文件头注释明确 "only used in connection with Office.js UDFs, not with runPython"）。

**要点**（完整源码研读 + 官方教程对照 + 工作流编排见 `references/16-xlwings-officejs.md`）：

- 值转换层（`_xlofficejs.py`）：读侧 `clean_value_data`（JsNull 哨兵 / 空单元格归一 / Error dict 取 basicValue）、写侧 `prepare_xl_data_element`（None→空、NaN→#NUM!、日期→FormattedNumber、`#`开头→Error 类型），data types 协议随 `runtime < 1.4` 降级裸值；
- UDF 全链路（`udfs_officejs.py`）：注入类型隐藏（CurrentUser/Book）、`xlfunc`/`xlret`/`xlarg` 装饰器（Literal→枚举、Annotated 选项、volatile/name/namespace/help_url/required_roles）、签名约束（无 keyword、vararg+optional 互斥）、22 locale 日期格式映射、调用管线三形态（流式/coroutine/同步）、WithScript 解包、对象句柄代际回收；
- 脚本全链路：book 注入与 `_lazy` 懒加载、参数强制与 ISO 解析、元数据生成（JS 包装 + `CustomFunctions.associate` + 注册期校验：重复名 / stream+requiresAddress 互斥）；
- socket.io 会话：`sio_connect`/`sio_disconnect`/`sio_custom_function_call`/`sio_cancel_task` 双字典管理流式任务生命周期。

## 二、xlwings 核心源码研读（xlwingsdll + xlwings 包）

xlwings 社区版（开源，BSD-3-Clause）由两部分组成：**C++ 原生 DLL**（`xlwingsdll/`，VBA↔Python 通信桥）与 **Python 包**（`xlwings/`，对象模型/COM 服务器/转换器/UDF）。本技能此前各章已按需引用其模式（§2.1 多实例枚举、§5.4 双向类型归一、§6.2 addin、§7 忙重试），本章做整体架构研读，供借鉴其工程实践。

### 2.1 xlwingsdll 原生 DLL

**定位**：xlwings 的 VBA 侧原生组件（C++，编译为 x86/x64 DLL），在 VBA 与 Python 进程之间建立并复用连接。源码工程 `xlwingsdll/`：`xlwings.sln` + `xlwingsdll.vcxproj`，源文件 `xlwingsdll.cpp`（导出实现）、`config.cpp`/`config.h`（配置与进程生命周期）、`dispatch.cpp`/`dispatch.h`（IDispatch 包装）、`utils.cpp`/`utils.h`（SAFEARRAY/变体工具）。

**导出面最小化**（`xlwingsdll.def`，仅 4 个函数）：

- `XLPyDLLVersion`：DLL 版本；
- `XLPyDLLNDims(VARIANT* src, int* dims, bool* transpose, VARIANT* dest)`：**VBA 数组维度转换**——1 维/2 维/标量互转 + 转置。校验源为 `VT_VARIANT|VT_ARRAY`；按 SafeArray `cDims`/`rgsabound` 判定行列；`dims=-1` 时按数据形状自动判定目标维度（标量/1 维/2 维）；2→1 维要求源为 (1×n) 或 (n×1)；转置在拷贝时用列主序索引变换实现（`srcIdx = transpose ? iDestRow + iDestCol*nDestRows : destIdx`）；错误文本回传 VARIANT 并返回 `E_FAIL`；
- `XLPyDLLActivate(VARIANT* result, const char* configFile, int mode)`：**配置驱动的连接建立**——按配置文件（`xlwings.conf`：`Command`=启动 Python 的命令行、超时等）取/建 `Config`（缓存复用）；`mode=-1` 杀掉 RPC 服务器；`mode=0` 仅查询已有接口（有则返回 IDispatch，无则返回空）；`mode=1` 不存在则启动。成功把接口 `VT_DISPATCH` 交回 VBA；
- `XLPyDLLActivateAuto(VARIANT* result, const char* command, int mode)`：**免配置入口**——直接把启动命令传给 DLL（xlwings 自身用此路径），其余设置走默认值。

**进程生命周期管理**（`config.cpp`）：`ActivateRPCServer()` 取 `Command` 拼接命令行 → `CreateProcessA` 启动 Python 进程 → **反复尝试创建接口对象，最多等 2 分钟** → 期间 Python 进程退出则报 "Python process exited before it was possible to create the interface object"；`Config` 持有 `HANDLE hJob`（**Job 对象**：Excel 退出时自动终止 Python 子进程，防止孤儿进程）。

**IDispatch 包装**（`dispatch.cpp`）：`CDispatchWrapper` 实现 `IUnknown` + `IDispatch`（`GetTypeInfo`/`GetIDsOfNames`/`Invoke`），VBA 拿到的 `XLPython` 接口即经此包装。

**对构建应用系统的借鉴**：原生桥 DLL 的"**导出面最小化 + 配置驱动进程生命周期 + 句柄复用**"模式——VBA 侧只暴露 4 个函数，复杂逻辑全在 Python 侧；数组维度/转置在原生层完成，避免 COM 往返；Job 对象挂接子进程，父进程退出自动清理。

**故障提示（2026-09-17 实战）**：`CreateProcessA` 以 ANSI 拼接命令行，**中文/长路径会破坏启动参数**——症状为 Excel 内 UDF server 报"拒绝访问"而 cmd 直调成功。规避：用 junction 纯 ASCII 别名作为模块/工作目录（`mklink /J D:\alias "中文长路径"`），勿尝试改引擎调用（引擎行为固定）。详见 `references/09-testing-debugging-guidance.md` 四·8。

### 2.2 xlwings Python 包核心架构

**公开 API**（`__init__.py`，13KB）：`xw.App`/`Book`/`Sheet`/`Range`/`Chart`/`Shape`/`Name`/`Table` 等对象、`xw.func`/`xw.sub`/`xw.arg`/`xw.ret`/`xw.script` 装饰器、`xw.Book.caller()`（UDF/脚本内定位调用方工作簿）、`xw.view`（交互查看）。动态导入平台层（Windows `_xlwindows` / macOS `_xlmac`）。

**对象模型**（`main.py`，165KB）：`App`（Excel 实例：`books`/`activate`/`quit`/`calculation`/`screen_updating`/`visible`）；`Book`（工作簿：`sheets`/`app`/`open`/`save`/`close`/`names`，UDF 模块注入 `book.set_mock_caller`）；`Sheet`（`cells`/`range`/`used_range`/`names`/`autofit`）；`Range`（`value`/`formula`/`options`/`expand`/`end`/`api`/`color`/`number_format`）；`Chart`/`Shape`/`Name`/`Table`。**`api` 属性直通 COM 原生对象**——xlwings 对象只是瘦封装，复杂操作可下探 COM。基类在 `base_classes.py`（平台无关抽象），`constants.py` 为 Excel 常量枚举（146KB）。

**COM 服务器**（`com_server.py`，10KB）——VBA↔Python 对象操作协议：

- `XLPython` COM 类（`_public_methods_` 26 个方法）：`Module`（导入 Python 模块）/`Tuple`/`TupleFromArray`/`Dict`/`DictFromArray`/`List`/`ListFromArray`（构造容器）/`Obj`（通用包装）/`Str`/`Var`/`Call`（调用可调用对象）/`GetItem`/`SetItem`/`DelItem`/`Contains`（容器操作）/`GetAttr`/`SetAttr`/`DelAttr`/`HasAttr`（属性操作）/`Eval`/`Exec`（执行代码）/`ShowConsole`/`Builtin`/`Len`/`Bool`/`CallUDF`（调用 UDF）；
- `XLPythonObject`：**任意 Python 对象包装成 COM**（`Item`/`Count`/`_NewEnum`）；`XLPythonEnumerator`：Python 迭代器 → `IEnumVARIANT`；
- `FromVariant`/`ToVariant`：COM 变体与 Python 对象双向（`unwrap`/`wrap`）；
- 依赖 pywin32（`win32com.server`）；`os.chdir(sys.exec_prefix)` 处理 pythoncom.dll 定位。
- **源码级研读**（启动引导 / XLPython 26 方法 / 对象包装 / 消息泵 / serve 注册）见 **2.2.3**

**转换器框架**（`conversion/`）——读写双向类型归一：

- `framework.py`：`Converter`（转换器）/`Accessor`（读写器注册）/`Pipeline`（阶段管线）/`ConversionContext`/`Options`（`@xw.arg`/`@xw.ret` 的 options 透传）/`accessors` 注册表；
- `standard.py`：**读取管线** `ReadValueFromRangeStage → CleanDataFromReadStage → AdjustDimensionsStage → TransposeStage → Ensure2DStage`；**写入管线** `WriteValueToRangeStage → CleanDataForWriteStage`；`RangeAccessor`/`RawValueAccessor`/`ValueAccessor`（`router` 按值类型选读写器）；`Dict`/`OrderedDict`/`Datetime`/`Date`/`Tuple`/`Json` 转换器；`_check_not_jagged` 拒绝锯齿数组；
- `numpy_conv.py`/`pandas_conv.py`/`polars_conv.py`：numpy/pandas/polars 转换器，**按 import 成功与否条件注册**；`__init__.py` 的 `read()`/`write()`/`async_read()`/`async_write()` 为统一入口（支持 `pipeline_overrides` 替换阶段）。

**UDF 系统**（`udfs.py`，29KB）：`@xw.func`（类别/异步/调用链/自动转置）、`@xw.sub`、`@xw.ret`（返回值转换）、`@xw.arg`（参数转换）；`get_udf_module`/`call_udf`（**按工作簿加载 UDF 模块并执行**）；`generate_vba_wrapper`（**自动生成 VBA 包装代码**——`import_udfs` 把 Python 函数注入 Excel 为可调 UDF）；`ComRange`（UDF 内的 Range 参数包装）；`has_dynamic_array`（动态数组检测）。**源码级研读**（装饰器/签名、ComRange 跨线程、call_udf 运行时、VBA 包装器生成、import_udfs 注入）见 **2.2.2**。

**平台适配层**：Windows `_xlwindows.py`（89KB，pywin32 COM 实现，见场景 A 章节与场景 D 主干导航）+ `_win32patch.py`（上游 COM 行为补丁）；macOS `_xlmac.py`（102KB）+ `mac_dict.py`（AppleScript 字典映射，258KB）+ `xlwings-dev.applescript`。平台差异汇总见 `xlwings/docs/missing_features.md`。

**对构建应用系统的借鉴**：COM 服务器"**最小方法集覆盖全部对象操作**"设计（27 个方法即完成 Python 任意对象的操纵）；转换器"**读写分管道 + 可插拔 Converter + 注册表**"模式（新增类型只需 Converter 子类 + 注册）；UDF"**Python 定义 → 自动生成 VBA 包装 → 注入工作簿**"链路。

#### 2.2.1 pandas / polars 转换器源码研读（DataFrame 数据通道细节）

`conversion/pandas_conv.py`（6.8KB）与 `conversion/polars_conv.py`（4.4KB）实现 DataFrame/Series 与 Excel 单元格区域的双向转换，是"结构化数据 → Excel 表"通道的核心。两者在模块层 `if pd:` / `if pl:` 条件注册（未安装则整块跳过），注册表位置 `conversion/__init__.py`。**pandas 注册了 `"df"` 别名**（`PandasDataFrameConverter.register(pd.DataFrame, "df")`），故 `options("df")` 与 `options(pd.DataFrame)` 等价；polars 只注册 `pl.DataFrame`/`pl.Series` 类型本身，无字符串别名。

**pandas DataFrame 写入（`write_value`）**：
- 选项：`index`（默认 True，写索引列）、`header`（默认 True，写表头）、`assign_empty_index_names`（False；True 时把 None 索引名替换为逐列递增的空格——Excel 表格仅允许一个空表头，多空格可让各列都有"名字"）
- 先处理索引：`index=True` 且 `value.index.name` 与某列同名时 `rename(None)` 防冲突，再 `reset_index()` 把索引平铺为最左列
- 类型兜底：`PeriodDtype` / `timedelta64[ns]` 列无对应 Excel 类型，统一 `astype(str)`
- 表头重组：单层索引时 `columns[0][:index_levels] = index_names` 把索引名写进表头行；`MultiIndex` 列时 `list(zip(*columns))` 转置为多层表头行，索引名只落在最底层
- 无表头（`header=False`）时直接 `value.values.tolist()` 纯数据

**pandas DataFrame 读取（`read_value`）**：
- `header=1` → 首行作列名；`header>1` → 前 N 行 `pd.MultiIndex.from_arrays` 作多层表头；`header=0` → 无表头
- `dtype` / `copy` 透传 `pd.DataFrame(...)`；`parse_dates`：`True` 等价 `[0]`，按列名或列号对目标列逐列 `.apply(xlserial_to_datetime)`（Excel 日期序列号 → datetime）
- 索引重建三步法：先给列临时整数名（`df.columns = pd.Index(range(...))`，规避重名列），`set_index` 前 N 列，再以"表头区最后一行前 N 格"为索引名——`value[header-1][:index]`
- **pandas Series**：读取走 `squeeze()`；写时若 `index.names` 全 None 且 `name` None 则默认无表头，否则转 `to_frame()` 复用 DataFrame 写入

**polars 写入（`write_value`）**：仅两行——`result = [df.columns]` + `result.extend([list(row) for row in df.rows()])`（polars 无索引概念，列名即表头）。Series 写为单列：有名字且要表头时 `[series.name] + series.to_list()`，每元素包一层 `[[item]]`。

**polars 读取（`read_value`）**：逐项映射 `pl.DataFrame` 构造参数——`has_header`（默认取 `header`，True 时首行作 `schema`）、`schema` / `schema_overrides`、`strict`（默认 True）、`infer_schema_length`（默认 100）、`nan_to_null`（默认 False）、`orient="row"`；`parse_dates` 与 pandas 不同，用 `pl.col(...).map_elements(xlserial_to_datetime, return_dtype=pl.Datetime)`。Series 读：`pl.Series(name=value[0], values=value[1:])`。

**行为契约（`tests/test_polars.py` 佐证）**：
- 含 None 的整型列往返后变 `Float64`——Excel 单元格无"整数空值"概念，None 只能以浮点 NaN 表示（"classic xlwings" 行为，非 pl.Int64）
- 纯日期列往返后为 `pl.Datetime` 而非 `pl.Date`——xlwings 默认按 datetime 交付
- 布尔 / 浮点列空值原样保持

**借鉴**：① 数据通道转换器统一 `read_value(value, options)` / `write_value(value, options)` 双方法签名，options 透传决定行为——新增数据格式只需按此契约实现并注册；② 类型兜底策略（Excel 无对应类型 → 字符串化）避免写回失败；③ 条件注册让可选依赖缺失时静默降级而非导入报错。

#### 2.2.2 UDF 系统源码研读（udfs.py，29KB / 838 行）

**定位**：桌面版 UDF 全链路（社区版，BSD-3-Clause）——`@xw.func`/`@xw.sub`/`@xw.ret`/`@xw.arg` 装饰器 → VBA 包装器生成 → COM 同步调用 → 异步线程与动态数组。与 1.8 的 officejs 版同源异路：桌面走 **COM 同步通道 + VBA 包装**（VBA `XLPy.CallUDF`），Web 走 socket.io + JS 包装。

**装饰器与签名系统**：

- `func_sig`：`inspect.signature` 解析——只接受位置参数（`POSITIONAL_OR_KEYWORD`/`VAR_POSITIONAL`），keyword 参数直接 `raise`（Excel UDF 无 keyword 语义）；返回 `args`/`defaults`/`vararg`
- `extract_type_and_annotations`：`Annotated` 顶层类型提取（`get_origin` 归一）——ValueAccessor 只需注册顶层类型，不必注册所有嵌套组合
- `xlfunc`（`@xw.func`）：类型提示 → `options["convert"]`（**datetime/date 除外**——pywin32 自动转换，GH #2571）；`Annotated` 的 `doc` 键进参数文档、其余键进 options；有默认值的参数记 `optional=默认值`；返回文档 = `__doc__` 或自动生成 `"Python function 'x' defined in 'file'"`；`category`（int 1–14 内置 Excel 类别 / str 自定义 ≤255 字符，默认 `"xlwings"`）；`call_in_wizard`（默认 True）/`volatile`（默认 False）/`async_mode`（仅 `"threading"`）
- `xlsub`（`@xw.sub`）：`xlfunc` + `sub=True`（生成 VBA `Sub` 而非 `Function`）
- `xlret`/`xlarg`：向 `__xlfunc__["ret"]["options"]` / `argmap[name]["options"]` 合并覆盖；`xlarg` 另接受 `vba`/`doc` 特殊键；参数名不存在 → 报错

**ComRange——跨线程 COM 序列化 Range**：

- 背景：COM 对象绑定创建线程（STA），UDF 后台线程不能直接触碰
- 序列化：`CoMarshalInterThreadInterfaceInStream(IID_IDispatch, rng.api)`；目标线程 `CoGetInterfaceAndReleaseStream` **单次解包**（解包后 `_ser=None`）；`__copy__` 重新序列化（每次 `_com` 调用传 `copy.copy(self)`）
- `_com` 重试：`loop.run_in_executor(com_executor, ...)`（线程池 initializer=`CoInitialize`）→ 可重试错误集：
  - `AttributeError`（Excel 忙、TypeInfo 未就绪）
  - `RPC_E_SERVERCALL_RETRYLATER`（-2147418111 / -2146777998）
  - `DISP_E_EXCEPTION` 且 scode ∈ `RETRYABLE_SCOPES`（None / -2146827284）
  - **随机指数退避**：max_backoff 1ms 起 ×2，封顶 512ms
- async 方法族：`get_formula(_array/2)` / `set_formula(_array/2)` / `get_shape` / `get_address` / `clear_contents`

**运行时调用 `call_udf`（核心）**：

1. `get_udf_module`：模块缓存 + **mtime 热重载**（`importlib.reload`）；嵌入式代码（`.py` sheet，仅 Excel）——社区版抛 `LicenseError`，PRO 走 `dump_embedded_code`
2. 保留参数 `caller`：按 `args_info` 定位名字，注入 `ComRange(xw_caller)`
3. **writing 保护**：函数声明 `writing` 且当前地址 == 写入地址 → 返回 `rval`（防递归重入）
4. 参数转换：COM missing（-2147352572）→ `optional` 默认值；Range 实例 → `conversion.read`；标量 → `conversion.read(None, arg, options)`
5. 设置 `BOOK_CALLER`（this_workbook）
6. **async_mode="threading"**：cache_key（func+args+pid+book+sheet+address 首格）命中 → 直接返回（非动态数组则删缓存）；未命中 → 填 `"#N/A waiting..."` 占位矩阵（按 caller 行列）→ `asyncio.run_coroutine_threadsafe(async_thread, loop)` → 立即返回占位；后台线程完成 `set_formula` 回写 → 触发重算（cache 命中）
7. 同步路径：动态数组 → cache 命中直返 / 未命中执行并写缓存（coroutine 转 loop）；普通 coroutine → `run_coroutine_threadsafe`；同步函数直接调用
8. `conversion.write` 写回；动态数组尺寸不匹配 → `xw_caller.resize(*result_size)` + `delayed_resize_dynamic_array_formula`（sleep 0.1s → 比较 caller/target 尺寸 → clear 超出的 caller → `set_formula_array` 重新触发，此时 cache 命中返回正确尺寸）；匹配 → 删缓存

**VBA 包装器生成 `generate_vba_wrapper`**（VBAWriter 模板输出）：

- 遍历模块 `dir()` 中所有 `__xlfunc__` → `Function`/`Sub` 签名（`Optional`/`ParamArray` 前缀、vararg 名 + `()`）
- Function 体：`call_in_wizard=False` 先查 `CommandBars("Standard").Controls(1).Enabled` 退出（函数向导场景）；`volatile` → `Application.Volatile`
- vararg 展平：`argsArray = Array(...)` → `ReDim Preserve` → For 循环逐项填充
- `#If App = "Microsoft Excel"` 分支：Excel 传 `ThisWorkbook`（`.xlam` 加载项用 `ActiveWorkbook`）+ `Application.Caller`；非 Excel 只传模块/函数/参数——统一 `XLPy.CallUDF` 调用；`If TypeOf Application.Caller Is Range Then On Error GoTo failed` → `failed:` 标签把 `Err.Description` 作为函数返回值

**导入 `import_udfs`**：

- 生成 `xlwings_udfs` 模块（`Attribute VB_Name` + "Autogenerated code" 注释 + `#Const App`）
- PRO 嵌入代码映射：读 config sheet 的 `RELEASE_EMBED_CODE_MAP` → `.py` sheet 名还原为真实模块名
- `VBComponents.Remove`（忽略 com_error）+ `VBComponents.Import`（com_error 440 → fallback `Application.Run("ImportXlwingsUdfsModule", path)`）
- **MacroOptions 注册**：Excel ≥14 时带 `Description`/`ArgumentDescriptions`/`Category`（函数向导显示参数说明）；文档截断 255 字符

**`has_dynamic_array`**：`lru_cache` 按 pid 检测（试调 `WorksheetFunction.Unique("dummy")`）——Excel 365 动态数组能力开关。

**借鉴**：跨线程 COM 的序列化 + 指数退避；"占位值回写触发重算"的异步完成模式；`writing` 防重入；mtime 热重载；VBA 包装器代码生成（模板 + `#If` 分支）——均为 Excel 自动化/异步化的可复用工程模式。

#### 2.2.3 XLPython COM 服务器源码研读（com_server.py，10KB / 385 行）

**定位**：VBA ↔ Python 的 **COM 服务器**（`XLPython` 对象）——Excel 内 VBA 调 `XLPy.CallUDF(...)` → COM 分派到本文件方法 → Python。与 2.1 xlwingsdll（VBA 侧桥）互为客户端/服务端。

**启动引导**：`os.chdir(sys.exec_prefix)` 后再 import pywin32（pythoncom.dll 定位 hack，`sys.frozen`/cx_Freeze 跳过）→ `os.chdir(cwd)` 还原。

**XLPython 类**（`_public_methods_` 26 个，覆盖"对象操纵全集"）：

- 容器构造：`Tuple`/`TupleFromArray`/`List`/`ListFromArray`（元素经 `FromVariant` 解包）/`Dict`/`DictFromArray`（交替键值，奇数参数报错）
- 通用：`Obj`（任意对象包装）/`Str`/`Var`（`lax=True`：dict→items 元组、ndarray→tolist、tuple→单元素元组）/`Len`/`Bool`/`Builtin`
- 调用：`Call`（参数中 tuple→位置参数、dict→kwargs、str→方法名）/`CallUDF`（解包 args → `call_udf` → 单元素/单行降维）
- 容器/属性：`GetItem`/`SetItem`/`DelItem`/`Contains`/`GetAttr`/`SetAttr`/`DelAttr`/`HasAttr`
- 动态执行：`Eval`/`Exec`（最多 2 个 dict：globals/locals）
- `ShowConsole`：`AllocConsole` + stdout/stderr 重定向到 `CONOUT$`

**对象包装**：`XLPythonObject`（`Item`/`Count`/`_NewEnum`）把任意 Python 对象包成 COM 集合；`XLPythonEnumerator`（`Next`/`Skip`/`Reset`/`Clone`，`_query_interface_` 认领 `IID_IEnumVARIANT`，Skip/Reset/Clone 抛 `E_NOTIMPL`）。

**转换**：`FromVariant`（`unwrap` 失败保留原值；`PyIDispatch` → `win32com.client.Dispatch` 并用 TypeInfo 文档名做 userName）/`ToVariant`（wrap 成 `XLPythonObject`）。

**异步事件循环**：模块级 `loop` + `_start_background_loop` 线程（`CoInitialize` + `run_forever`）——COM 服务器线程与 asyncio 线程分离，UDF 异步任务跑在后台 loop（2.2.2 的 `run_coroutine_threadsafe` 即投递到这里）。

**serve() 注册与消息泵**：

1. 覆写 `win32com.server.policy.DefaultPolicy._CreateInstance_`（请求 clsid == XLPython → `serverutil.wrap(XLPython())`，其余走基类）——win32com 服务器实例化钩子
2. `pythoncom.MakePyFactory(clsid)` + `CoRegisterClassObject(clsid, factory, CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE | REGCLS_SUSPENDED)`
3. `EnableQuitMessage` + `CoResumeClassObjects`
4. **消息泵**：`MsgWaitForMultipleObjects((), 0, INFINITE, QS_ALLEVENTS)` + `PumpWaitingMessages()`（收到 wm_quit → break 退出）
5. 收尾：`CoRevokeClassObject` + `CoUninitialize`

**借鉴**：**COM 方法集最小完备设计**（26 个方法即完成 Python 任意对象的操纵）；`Eval`/`Exec` 双字典 globals/locals 的动态执行入口；STA 服务器线程 + asyncio 后台线程分离；手写 COM 消息泵（`MsgWaitForMultipleObjects` + 消息泵循环）。

### 2.3 构建与测试体系（scripts / src / tests）

#### 2.3.1 scripts/（CI 构建与文档生成）

- `build_excel_files.py`（7.8KB）：CI 用 **Aspose.Cells（pythonnet/clr）** 生成测试工作簿与加载项产物——读 `xlwings.xlam`/`quickstart_*.xlsm`/`xlwings.bas`，按 `GITHUB_REF`/`GITHUB_SHA` 注入版本字符串，重命名 `xlwings32.dll`/`xlwings64.dll` 后重新打包 xlam/xlsm；展示"用代码重打 Office 二进制包"的完整 CI 写法；
- `build_rest_api_docs.py`（8.7KB）：**先启动 REST API 服务**（`xlwings.rest.api`，localhost:5000），再用真实 Excel（`xw.App` 建工作簿/写值/建图表/命名区域）生成样例响应，驱动 API 文档产出；展示"以真实 Excel 实例生成接口文档样例"的方法。

#### 2.3.2 src/（Rust calamine 只读引擎：xlwingslib）

`src/lib.rs`（9.6KB）+ `Cargo.toml` 构成 calamine 只读引擎的 **Rust 实现**，经 maturin 编译为 Python 扩展模块 `xlwings.xlwingslib`，由 PRO `pro/_xlcalamine.py`（1.2.4）`from xlwings import xlwingslib` 加载调用。

**Cargo.toml 工程定义**：
- `[package] name="xlwings"`、`version="0.0.0"`（版本号不参与发布，走 pyproject `dynamic=["version"]`）、`edition=2021`、`publish=false`
- `[package.metadata.maturin] name="xlwings.xlwingslib"` + `pyproject.toml [tool.maturin] module-name="xlwings.xlwingslib"`——编译产物的 Python 模块路径
- `[lib] name="xlwings"`、`crate-type=["cdylib"]`——Python 扩展的标准形态
- 依赖三件套：
  - `calamine`：**xlwings 自 fork 的 git 依赖**（`https://github.com/xlwings/calamine`，`rev ed0a4bd`），`features=["dates"]` 启用日期解析
  - `chrono 0.4.42`：`features=["serde"]`、`default-features=false`，支撑 datetime / duration 类型
  - `pyo3 0.29.0`：`features=["extension-module","chrono"]`
- 注释明确 **abi3 wheels 不支持**：`DateTime` 不属于 PEP 384 ABI spec——该扩展必须按 Python 版本分别编译，无法做通用 abi3 wheel

**lib.rs 导出面（4 个 `#[pyfunction]`，注册于 `#[pymodule] xlwingslib`）**：
- `get_sheet_values(path, sheet_index, err_to_str)`：`open_workbook_auto` 开文件 → `worksheet_range_at(sheet_index)` 取 used_range → `used_range.end()` 取右下角 → `get_values` 全表读；空表返回 `vec![vec![]]`
- `get_range_values(path, sheet_index, cell1, cell2, err_to_str)`：同上前半，空 used_range 用 `Range::new((0,0),(0,0))` 兜底避免 panic；`cell1/cell2` 为 0-based `(row, col)`
- `get_sheet_names(path)`：`book.sheet_names()` → `Vec<String>`
- `get_defined_names(path)`：`book.defined_names()` → `Vec<(name, ref)>`（ref 形如 `Sheet1!$A$1:$C$3`）

**`get_values` 值映射表（对齐 COM 语义的关键）**：
- `DataType::Int` → `CellValue::Float((*v) as f64)`——**整数统一转 Float**，注释明言 "Float to be in line with COM API"（xlwings 读 Excel 一律浮点）
- `Float` → Float；`String` → String；`Bool` → Bool；`Empty` → Empty
- `DateTime`：`as_datetime().is_none()`（日期溢出——如 `1e+20` 被格式化为日期单元格）→ Empty，否则 DateTime；`DateTimeIso` → 直接 `as_datetime().unwrap()`
- `Duration` → `Timedelta(as_duration())`；`DurationIso` → `Time(as_time())`
- `Error`（8 类 `CellErrorType`）：`Div0/NA/Name/Null/Num/Ref/Value/GettingData` 分别映射 `#DIV/0!` `#N/A` `#NAME?` `#NULL!` `#NUM!` `#REF!` `#VALUE!` `#DATA!`——`err_to_str=true` 转字符串、`false` 转 `Empty`

**`CellValue` → Python 对象（`IntoPyObject`）**：`Int→int`、`Float→float`、`String→str`、`Bool→bool`、`Time/DateTime/Timedelta→chrono`（经 pyo3 `chrono` 特性转 `datetime/time/timedelta`）、`Empty→py.None()`、残留 `Error` → "Error" 字符串兜底。

**错误透传**：
- `import_exception!(xlwings, XlwingsError)` 把 Python 侧 `xlwings.XlwingsError` 引入 Rust；`CalamineError`（包 `calamine::Error`）`From` 转 `PyErr`
- `get_sheet_values` / `get_range_values`：`Error::Io` → `PyIOError`，其余 → `XlwingsError`
- `get_sheet_names` / `get_defined_names`：直接 `Result<_, CalamineError>`（`?` 自动转）
- 注意：开文件 / 取工作表处大量 `.unwrap()`，源码 TODO 标注 "proper error handling" 未完成——文件不存在或 sheet 越界会 panic 而非抛 Python 异常，复用时须先自行校验路径与 sheet 序号

**与 python-calamine 的关系**：lib.rs 尾部注明基于 `dimastbk/python-calamine`（附 MIT 许可全文）——`CellValue` 枚举与 `get_values` 映射思路同源；但依赖的是 **xlwings 自 fork 的 calamine**（带 `dates` 特性），且模块名 / 异常体系（`XlwingsError`）为 xlwings 定制，属"同一 calamine 库的不同封装"（对应 1.2.1 通道选型表中的社区版 `python-calamine`）。

- 说明：社区版仓库即含 Rust 源，但引擎激活需 PRO 许可（对应 1.2.4 calamine 引擎）。

#### 2.3.3 tests/（pytest 测试体系）

- `pytest.ini`：`addopts = -v -p no:faulthandler -p no:warnings -p no:cacheprovider`；
- `common.py`：`TestBase(unittest.TestCase)` 基座——**双 App 实例**（`app1`/`app2`，`visible=False`，`tearDownClass` 用 `kill()` 收尾）、setUp 建两个 3-sheet 工作簿、**`NotImplementedError → skipTest`**（平台差异不判失败）；
- 顶层 40 个测试文件（约 380KB）+ 5 个子目录：`reports/`（PRO Reports）、`restapi/`（REST 接口）、`test_engines/`（含 `benchmark.py` 引擎基准）、`udfs/`（UDF 专项）；
- 对象层：`test_app.py`/`test_book.py`/`test_sheet.py`/`test_range.py`（32KB，最厚）/`test_shape.py`/`test_table.py`/`test_names.py`/`test_font.py`/`test_characters.py`/`test_active.py`；
- 转换层：`test_conversion.py`（23KB）/`test_ndim.py`/`test_chunking.py`/`test_converter_json.py`/`test_polars.py`；
- 引擎与专项：`test_remote_*.py`（3 个）/`test_custom_functions_officejs.py`/`test_custom_scripts_call.py`/`test_object_handles.py`（22KB）/`test_caller.py`/`test_e2e.py`/`test_jsnull.py`/`test_markdown.py`/`test_fileformats.py`/`test_streaming_*.py`（2 个）/`test_async_load.py`；
- 数据资产：`cell_errors.xlsx`/`tables.xlsx`/`test book.xlsx`/`macro book.xlsm`/`sample_picture.png`/`pandas_excel_files_quick_test.py`。

**构建系统与开发文档**：Python 包 `pyproject.toml`/`setup.py`/`MANIFEST.in`/`Makefile`；C++ DLL `xlwingsdll/`（`xlwings.sln`/`xlwingsdll.vcxproj`）；Rust `Cargo.toml`。路线图 `plans/issue-shortlist.md` 与 `plans/v1.0-breaking-changes.md`；开发者指南 `DEVELOPER_GUIDE.md`；版本历史 `xlwings/docs/whatsnew.md`。

---


## 三、入门学习与教学演示（执行步骤）

本任务为系统学习 xlwings 基础知识，或向用户演示 xlwings 能力。学习材料为本技能内置 `examples/python-for-excel-course-main/` 课程讲义（仓库详尽技术研读见 3.1）；教学演示的演示文件选材规则见 4.3「教学演示选材与执行」。

### 3.1 课程仓库技术研读（python-for-excel-course-main）

#### 3.1.1 仓库概览与学习路径

- **定位**：xlwings 官方 2018 年 YouTube 入门视频课程（"Python for Excel with xlwings"）配套 notebook 仓库，共 8 个目录（`0 - Intro` 至 `7 - Part7`）+ 7 个教程 notebook + 配套数据文件。README 明确标注课程偏旧（outdated）——**适合理解基础概念与经典模式，具体语法以 `xlwings/docs/` 与官方最新文档为准**。
- **目录结构**：
  - `0 - Intro/`：`Instructions.md`（学习路径建议）+ `Language and Regional Settings.md`（非英语 Excel 差异处理）
  - `1 - Part1/`：Tutorial 1 The Basics（+ `table_objects.xlsx`/`timeseries.xlsx`/`img` 架构图）
  - `2 - Part2/`：Tutorial 2 Converters and options
  - `3 - Part3/`：Tutorial 3 Add-In & RunPython
  - `4 - Part4/`：Module 4 UDFs（+ `sql.xlsx`/`timeseries.xlsm`）
  - `5 - Part5/`：Tutorial 5 Advanced xlwings topics
  - `6 - Part6/`：Tutorial 6 Unit Testing（+ `test_sample.py`）
  - `7 - Part7/`：Tutorial 7 App Development & Deployment
- **学习路径**（`Instructions.md`）：
  - 新手从 Intro（Anaconda + Jupyter 安装引导）开始；
  - 有基础按 Tutorial 1→4 顺序学习，再 Tutorial 5→7（高级部分可任意顺序）；
  - Tutorial 3 首次涉及 xlwings 加载项安装（课程大纲需要时才装）。
- **视频对应**：各 Part 主题与视频号对应关系见仓库 README（YouTube playlist 链接）。

#### 3.1.2 Tutorial 1 - The Basics（`1 - Part1/Tutorial 1 - The Basics.ipynb`，24.9KB，最厚）

主题：xlwings 对象模型与基础读写。技术要点：

- **`xw.view()`**：把 Excel 当表格查看器——`xw.view(data)` 开新簿显示；`xw.view(data, xw.sheets.active)` 复用活动 sheet（每次调用清空该 sheet）。
- **连接 Book 三种方式**：`xw.Book()` 新建；`xw.Book('Book2')` 连接未保存工作簿；`xw.Book(path)` 按文件名/路径连接（未打开则打开）。均跨所有 Excel 实例查找。
- **Range 读写**：写值/读值/多格同值（`range('A3:B4').value = 123`）；**Excel 数值格式是 float**；日期时间直接赋值；索引记号 `range((1,1))`（1-based 与 Excel 一致）；公式 `range('B1').formula`。
- **命名区域**：`range('B1').name = 'test'` 后可用 `range('test')` 读写——**读写参数/输出目标单元格的稳固方式，抗 sheet 重组**（课程明确推荐）。
- **2D 区域**：嵌套列表写左上角；`expand('table'/'down'/'right')` 对应 Ctrl-Shift+方向；`clear_contents()` 只清值 vs `clear()` 连格式。
- **1D 向量**：横向直接写列表；纵向 `options(transpose=True)`（等价 `[[5],[6],[7],[8]]`）。
- **`ndim` 选项**：`options(ndim=2)` 强制 2D 读取。
- **autofit**：单格 `range('A3').autofit()` / 区域列 `range('A1:C3').columns.autofit()` / 整列 `range('A:A').autofit()`。
- **Range 切片**：`rng[0,0]`/`rng[1]`/`rng[:,3:]`/`rng[1:3,1:3]`（对 Range 对象用方括号切片）。
- **完整限定（多实例）**：`xw.apps.keys()` 拿 PID；`xw.apps[pid].books[0].sheets[0].range('A1')`（方括号=Python 式）vs `xw.apps(pid).books(1).sheets(1)`（圆括号=Excel 式）；**同一文件在多个 Excel 实例打开时，`xw.Book('xxx')` 会报错，必须完整限定 `app1.books['timeseries.xlsx']`**。
- **Active 对象**：`xw.apps.active`/`xw.books.active`/`xw.sheets.active`/`xw.Range('A1')`——**仅供交互使用，脚本中不可靠，脚本一律走 `sheet.range(...)`**。
- **Sheets 操作**：`name`/`count`/`add(name='New', after='Sheet1')`；快捷记号 `sheet['A1']`、`sheet['A1:B5']`、`sheet[0,1]`、`sheet[:10,:10]`。
- **图表**：`sheet.charts.add()` + `set_source_data` + `chart_type`（'line'/'area'）+ `top`/`left` 定位；图表类型枚举 `xw.constants.chart_types`。
- **Matplotlib 图插入**：`sheet.pictures.add(fig, name='SwapRate', update=True)`——**`update=True` + 同名重调可改内容而不改变位置/尺寸；尺寸换算 `width = fig.size_inches × dpi`**；添加后可直接改 `plot.width/height`。
- **Excel Table 对象**（当时未官方支持，读取可用）：`range('Table1')`（表体）、`range('Table1[Symbol]')`（单列）、`range('Table1[[#All],[Last]]')`（含表头合计）、`[[#Headers],...]`/`[[#Totals],...]`（单取表头/合计行）、`[[Index]:[Last]]`（相邻多列）——结构化引用语法。
- **效率铁律**：最小化跨应用调用——**整块读写 2D 区域（`range('A1').value = np.arange(150).reshape((30,5))`）而非逐单元格循环**。
- **api 下探**：`range('A1').api` 拿到平台原生对象（Win=pywin32 COM / Mac=appscript）——补 xlwings 未实现功能，**代码平台相关**（如 `ClearFormats()` vs `clear_formats()` 按 `sys.platform` 分支）。
- **调用 VBA 宏**：`wb.macro('MySum')` 返回可调用对象，`my_sum(1, 2)` 直调。

#### 3.1.3 Tutorial 2 - Converters and options（`2 - Part2/Tutorial 2 - Converters and options.ipynb`，8.9KB）

主题：默认转换器（list of list + datetime/数字/字符串/空值自动转换）之上的可控选项。技术要点：

- **options 语法**：`range('A1').options(convert=None, **kwargs).value`；**选项在访问值时才求值**。
- **`numbers`（只读）**：`options(numbers=int)` 或传函数 `options(numbers=lambda x: round(x,1))`。
- **`ndim`（只读）**：强制 1D/2D 读取（单格/单列场景）。
- **Dict 转换器**：两列数据 `options(dict).value` 直接得字典。
- **NumPy**：写 ndarray 直接赋值；读 `options(np.array, expand='table').value`。
- **Pandas DataFrame**：写（带表头+索引）；读 `options(pd.DataFrame, index=False, expand='table')`；`index`/`header` 参数控制写回；**MultiIndex 读回用 `index=2`**。
- **Pandas Series**：带 `name`/`index.name` 往返；`header`/`index` 控制与 DataFrame 相同。

#### 3.1.4 Tutorial 3 - Add-In & RunPython（`3 - Part3/Tutorial 3 - Add-In & RunPython.ipynb`，7.1KB）

主题：Excel 加载项与 RunPython（从 Excel 调 Python）。技术要点：

- **架构**：`RunPython` 是 VBA 函数（Win/Mac 跨平台）——VBA 宏构建 `"import module;module.func()"` 命令串，Python 端用 `xw.Book.caller()` 引用调用方工作簿。
- **加载项安装**：CLI `xlwings addin install`；或 GitHub release 手动安装（`File > Options > Add-ins > Go...`）。
- **配置存储**：外部 `xlwings.conf`（Win：用户目录 `.xlwings\xlwings.conf`；Mac Excel 2016：`~/Library/Containers/com.microsoft.Excel/Data/xlwings.conf`）。
- **quickstart 骨架**：`xlwings quickstart hello` 生成 Excel + `hello.py` 工程（含 RunPython 与 UDF 双示例）。
- **VBA 引用**：`Tools > References` 勾选 xlwings——RunPython 调用与 Import UDFs 必需（运行已导入的 UDF 不需要）。
- **跨平台约束**：不用 ActiveX 控件，用表单控件按钮（ActiveX 仅 Windows）。
- **设置优先级**：`xlwings.conf` sheet > 工作簿同目录 `xlwings.conf` 文件 > 加载项全局配置 > 默认值；常用设置 `Interpreter`/`PYTHONPATH`。
- **Gotcha**：cwd 不一致——用全路径（`os.path.dirname(__file__)`）。

#### 3.1.5 Tutorial 4 - UDFs（`4 - Part4/Module 4 - UDFs.ipynb`，12.8KB）

主题：用户自定义函数（@xw.func 全链路）。技术要点：

- **位数组合**：Excel/Python 任意位数混搭可用（32/64 任意组合）。
- **一次性准备**：信任中心勾"Trust access to the VBA project object model" + 安装加载项。
- **首个 UDF**：`@xw.func` 装饰 → 加载项 `Import Functions` → 单元格 `=hello("world")`；**改函数体后 `Ctrl-Alt-F9` 重算即可，改参数/函数名才需重导入**。
- **数组公式**：`@xw.arg('x', ndim=2)` 保证参数是 list of list（单格/向量需显式）；**数组公式只跨界一次，效率高于大量单格公式**。
- **@xw.arg/@xw.ret 转换器**：NumPy（`np.array, ndim=2`）；DataFrame（默认 ndim=2，`index`/`header` 控制，或返回 `.values` 抑制表头索引）；`xw.Range`（"no-converter"：函数收 Range 对象操作公式）。
- **动态数组**：`@xw.ret(expand='table')` 让函数"写到公式外"（quandl 行情例）——**注意不得覆盖既有值、不得用易变公式（`=TODAY()`）**。
- **VBA 设置**：`UDF Modules`（分号分隔多模块；空=同目录同名 .py）；`PYTHONPATH` 保证模块可导入；`Restart UDF Server` 重载全部（改间接导入的模块时）。
- **内置扩展**：`=sql(SQL, table_a, ...)`——公式内嵌加载项，**无需 xlsm、无需 VBA 引用**（官方文档 `xlwings/docs/extensions.md`）。
- **异步 UDF**：`@xw.func(async_mode='threading')`——立即返回 `#N/A waiting...`，后台计算完成后更新单元格，不阻塞 Excel。
- **实战**：`end_of_month` UDF（`x.resample('M').last()` 重采样，配 `timeseries.xlsm`）。

#### 3.1.6 Tutorial 5 - Advanced xlwings topics（`5 - Part5/Tutorial 5 - Advanced xlwings topics.ipynb`，13.3KB）

主题：调试 / 实时数据 / 扩展 / 自定义转换器 / REST API。技术要点：

- **RunPython 调试**：文件尾加 `if __name__ == '__main__': xw.Book('hello.xlsm').set_mock_caller(); hello_xlwings()`——脱离 Excel 直接跑源码调试。
- **UDF 调试（Windows）**：VBA 设置勾 `Debug UDFs` + 文件尾 `if __name__ == '__main__': xw.serve()`——调试服务器挂起，IDE 断点/打印。
- **实时写**：`while True: sheet['A1'].value = dt.datetime.now(); time.sleep(0.5)`（轮询刷新 Excel 单元格）。
- **实时读**：轮询单元格值变化即打印（观察 Excel→Python 数据流）。
- **自定义扩展**：导入函数 → 粘贴进加载项 Extensions 模块（密码 `xlwings`）→ 代码中 `ThisWorkbook` 改 `ActiveWorkbook` → 源码放 `sys.path` 或加 `PYTHONPATH`——做成"装加载项即用"的内置公式。
- **自定义转换器**：继承 `xlwings.conversion.Converter`，实现静态 `read_value(value, options)`/`write_value`；可选 `base = 内建转换器`（如 PandasDataFrameConverter）；`register(类型)` 设为默认 / `register('别名')` 按名调用（DataFrameDropna 完整例）。
- **REST API**：xlwings 自带 Flask 服务（localhost:5000）——`/books`、`/book/timeseries.xlsx/sheets/sheet1/range/A1?expand=table` 返回 JSON，转 DataFrame 分析——Excel 数据对外暴露接口的雏形。

#### 3.1.7 Tutorial 6 - Unit Testing（`6 - Part6/Tutorial 6 - Unit Testing.ipynb`，3.5KB + `test_sample.py`）

主题：用 Python unittest 测 Excel 工具。技术要点：

- **定位**：Excel 无单元测试能力；Python 内置 unittest 可测——**无加载项、不改 VBA 代码、标准库 + xlwings 即可**；也是验证/调试存量 VBA 代码的方式。
- **可测对象**：带/不带 VBA 的表格（单元格逻辑）；VBA 代码 vs Python 替代实现（可引 SciPy/QuantLib 等被充分测试的库做交叉验证）。
- **运行**：`python -m unittest test_sample`（`-v` 详细）；单用例 `python -m unittest test_sample.TestStringMethods.test_upper -v`。
- **延伸**：官方 `xlwings-automated-testing` 仓库演示三种方式——VBA 单元测试 / 模型验证（替代实现）/ 单元格逻辑测试。

#### 3.1.8 Tutorial 7 - App Development & Deployment（`7 - Part7/Tutorial 7 - App Development & Deployment.ipynb`，9KB）

主题：应用开发与部署链路。技术要点：

- **Web 原型**：`simulation-demo` 展示 xlwings 做 web 应用原型（Excel 为前端）。
- **部署方式递进**：
  - **vanilla（Excel+py 文件）**：zip 分发，接收方须保持两文件在一起；
  - **Python 包**：默认 VBA 设置把 Excel 目录加入 `PYTHONPATH`，且 `site-packages` 也可访问——把源码打成包分发（`pip install`、版本号、代码对用户隐藏、Excel 可独立移动）；最小 `setup.py`（从源码正则提取 `__version__`）+ `python setup.py sdist` + `pip install dist/hello-1.0.0.zip`；
  - **分发渠道**：共享盘/内网（`pip install \\\\drive\\pkg.zip`）、公司 PyPI/Anaconda.org/Gemfury、`bdist_wininst` 自安装包、`pip install git+https://...`（按 tag/commit）；
  - **`--standalone`**：`xlwings quickstart hello --standalone` 把加载项内容作为标准 VBA 模块嵌入——**去掉加载项依赖**；
  - **RunFrozenPython**：`pyinstaller frozen.py` 冻结为 exe——`frozen.spec` 的 `excludes` 显式排除 scipy/numpy/pandas/matplotlib 等可选依赖瘦身；VBA 侧 `RunFrozenPython ThisWorkbook.Path & "\\dist\\frozen\\frozen.exe"`；Inno Setup/NSIS 打包 + 放 `%APPDATA%`/Program Files；**冻结可执行文件不支持 UDF，只支持 RunPython 函数**。

#### 3.1.9 配套数据文件用途

- `1 - Part1/table_objects.xlsx`：Tutorial 1 表对象结构化引用读取样例（含表头行+合计行的表）。
- `1 - Part1/timeseries.xlsx`：Tutorial 1 多实例完整限定 + Tutorial 5 REST API 读取样例。
- `4 - Part4/sql.xlsx`：Tutorial 4 内置 `=sql(...)` 扩展演示（普通 xlsx 即可）。
- `4 - Part4/timeseries.xlsm`：Tutorial 4 `end_of_month` 重采样 UDF 实战数据。
- `6 - Part6/test_sample.py`：Tutorial 6 unittest 最小样例（TestStringMethods）。
- `img/`：架构图（frompython / runpython / udf_architecture.png、array_formula.png）——理解 Python↔Excel 通信链路。

#### 3.1.10 对构建应用系统的借鉴

- **对象模型走查路径**：App→Book→Sheet→Range 的完整限定 vs Active 快捷——脚本一律完整限定（分层铁律）。
- **效率模式**：整块 2D 读写、数组公式、动态数组——跨界次数最小化（对应 04-python-guidance 性能主题）。
- **转换器体系**：内建转换器 + 自定义 Converter（base + register）——与第二章 2.2 转换器框架研读互证。
- **调试三件套**：`set_mock_caller` / `xw.serve()` / 实时轮询——开发期脱离 Excel 与直连调试。
- **测试思想**：Python 测试 Excel 工具的"无侵入"路径（09-testing-debugging-guidance）。
- **部署演进**：zip → pip 包 → standalone → 冻结 exe——按目标机环境选型（10-deployment-delivery）。

### 3.2 澄清范围与选材

先与用户确认本次目标并一次确定范围与演示文件，此后只按确认结果执行，不跳步、不换文件：

- **学习范围**：系统入门按 3.1 研读的课程结构顺序学习全部讲义；查漏补缺时由用户指明主题所属 Part，仅学习该部分及其前置依赖。课程为 2018 年官方入门视频课程（YouTube）配套 notebook，按 `0 - Intro` 至 `7 - Part7` 顺序分目录存放，各 Part 主题、视频号及对应关系见 `python-for-excel-course-main/README.md` 与目录名。

### 3.3 执行讲义

进入课程目录后，按确认的范围从首至末逐个打开 notebook 并理解内容。含可执行单元格的 notebook 用 Jupyter 内核或等价方式顺序执行，验证输出与讲义一致。执行前先注入本地 xlwings 源码，讲义涉及的第三方包按需安装。

### 3.4 执行演示并输出小结

若含教学演示环节，按 4.3「教学演示选材与执行」确定演示文件并运行验证输出；随后输出小结：覆盖了哪些概念、运行了哪些样例、用户可继续深入的方向（对应场景 A-C 及本场景源码部分）。

---

## 四、案例总览与 PRO 案例研读

### 4.1 案例总览与依赖矩阵

**技能内置案例**（`examples/`，12 项）：

| 案例 | 依赖 | 展开位置 |
|------|------|---------|
| `xlwings-factsheet-demo-main/` | **PRO**（`xlwings.pro.reports`：Markdown/MarkdownStyle/Image） | 本章 xlwings PRO 案例研读 |
| `xlwings-demo-master/` | PRO/社区版混合（13 场景） | 子目录地图见下方（唯一明细入口）；PRO 深读见本章 PRO 案例研读 |
| `xlwings-eikon-master/` | **PRO 部分**（`_report_template.py`/`sample3.py` 用 `create_report`） | 11-scenario-a 七章 7.6（主链路）+ 本章 PRO 案例研读 |
| `xlwings-server/` | **Server**（FastAPI 服务端工程） | `references/14-xlwings-server-guidance.md`（六块完整研读） |
| `xlwings-lite/` | Lite（官方文档 31 篇离线镜像 + README 索引） | `references/15-xlwings-lite-guidance.md`（逐阶段引用源） |
| `taxi-duckdb-main/` | xlwings Lite / Pyodide + DuckDB（社区版，不依赖 PRO/Server） | `references/15-xlwings-lite-guidance.md` 阶段七（完整案例研读） |
| `python-for-excel-course-main/` | 社区版 | 三章入门学习 |
| `simulation-demo-master/` | 社区版 | 11-scenario-a 七章 7.4 |
| `xl-pq-handler-master/` | 社区版 | 11-scenario-a 七章 7.5 |
| `static-excel-test-master/` | 社区版 | 11-scenario-a 七章 7.3 |
| `cross-check-reports-main/` | 社区版 | 11-scenario-a 七章 7.3 |
| `excel-automated-testing-master/` | 社区版 | 11-scenario-a 七章（自动化测试） |
| `Excel_udf_itus-main/` | 社区版 | 07-udf 六章（UDF 完整工程）；SQLite 数据服务模式见下方 4.6 |

**源码包内置案例**（`xlwings/examples/`，5 项 + `build_lite.py`，**全部社区版 BSD**，不依赖 PRO/Server）：

| 案例 | 用途 | 引用位置 |
|------|------|---------|
| `udf/udf.py` | `@xw.sub`/`@xw.func` 最小 UDF 集 | 07-udf 六章（C 侧） |
| `fibonacci/` | UDF + `build_standalone.py`（PyInstaller 冻结 → zip 便携分发） | 07-udf 六章（C 侧） |
| `database/database.py` | SQLite → Excel 数据回填 | 11-scenario-a 七章（A 侧）；SQLite 访问细节见下方 4.6 |
| `mpl/mpl.py` | matplotlib 图表 → Excel | 11-scenario-a 七章（A 侧） |
| `simulation/simulation.py` | numpy 蒙特卡洛模拟回填 | 11-scenario-a 七章（A 侧） |
| `build_lite.py` | 打包上述 5 例为 xlwings Lite 的 `_build/*.zip` | 1.6 / 2.2 |

**结论**：依赖 PRO 的为 `xlwings-factsheet-demo-main`、`xlwings-demo-master`（PRO 部分 6 个子目录：reader / restapi / reporting 系 4 项，见 4.2 子目录地图）、`xlwings-eikon-master`（Reports 部分）；依赖 Server 的为 `xlwings-server`；`taxi-duckdb-main` 为 **Web 加载项（Lite）案例**，不依赖 PRO/Server，完整研读见 `references/15-xlwings-lite-guidance.md` 阶段七（4.5 已指针化）；**其余全部社区版**由场景 A/场景 C 引用介绍，详细展开见 `references/11-scenario-a-python-automation.md` 七章与 `references/07-udf-guidance.md` 六章——本章不再详展社区版案例（唯二例外：SQLite 数据服务模式在下方 4.6 详展，因其同时是官方源码包案例与 UDF 工程的"数据层"样板）。

### 4.2 xlwings-demo-master 子目录地图（13 项全览）

`xlwings-demo-master/` 同时含 PRO 与社区版示例，是 13 场景综合演示（本技能内唯一"一个仓库两套依赖"的案例）。各子目录的归属、内容与详展位置：

| 子目录 | 依赖 | 内容 | 详展位置 |
|--------|------|------|---------|
| `basics/` | 社区版 | 最小脚本 + UDF（`xwdemo.py`/`xwdemo.xlsm`） | 11-a 七章 7.7；教学演示选材（见本节下） |
| `correlation/` | 社区版 | DataFrame → 图表、图片导出（`correlation.py`/`.xlsm`） | 11-a 七章 7.7；教学演示选材（见本节下） |
| `frozen/` | 社区版 | PyInstaller 冻结部署（`demo.py`/`demo.spec`/`demo.xlsm`） | 11-a 七章 7.7 |
| `interactive/` | 社区版 | Jupyter 交互（`interactive.ipynb`/`AAPL.xlsx`） | 11-a 七章 7.7；教学演示选材（见本节下） |
| `performance/` | 社区版 | 性能实测（`arrays.ipynb`/`raw.ipynb`/`udfs.py`/`arrayudf`） | 11-a 七章 7.7；教学演示选材（见本节下） |
| `simulation/` | 社区版 | 蒙特卡洛模拟 UDF（`simulation.py`/`.xlsm`） | 11-a 七章 7.7 |
| `testing/` | 社区版 | pytest 测试骨架（`test_mybook.py`/`test_sample.py`/`.gitlab-ci.yml`） | 11-a 七章 7.7 |
| `reader/` | **PRO** | File Reader 只读引擎（`xlwings File Reader.ipynb` + AAPL 系列 4 个文件） | 下方 4.4 PRO 案例研读 |
| `reporting/` | **PRO** | Reports 模板报告（`report_template.py`/`report_template.xlsx`/`reports101.ipynb`） | 下方 4.4 PRO 案例研读 |
| `reporting_bigmac/` | **PRO** | BigMac 指数报告（`_bigmac_index.py` + `big-mac-raw-index.csv` → `report_*.pdf`） | 下方 4.4 PRO 案例研读 |
| `reporting_filter_frames/` | **PRO** | 过滤帧报告（`filter_frames.py` + `holdings.csv`） | 下方 4.4 PRO 案例研读 |
| `reporting_fund/` | **PRO** | 基金报告（`fund_template.py` + `data.pickle`） | 下方 4.4 PRO 案例研读 |
| `restapi/` | **PRO** | REST API 示例数据（`timeseries.xlsx`，仅数据无代码） | 下方 4.4 PRO 案例研读 |

### 4.3 教学演示选材与执行

本技能教学演示的演示文件取自 `xlwings-demo-master/`（4.2 子目录地图）：按业务主题从地图中选取**唯一一份**演示文件；无法确定时以 `basics/xwdemo.py` 最小闭环为主。社区版演示文件在 `references/11-scenario-a-python-automation.md` 七章 7.7 有文件级用途详展。演示运行与小结输出见 3.4「执行演示并输出小结」。

### 4.4 xlwings PRO 案例研读（依赖 PRO Reports / Reader / REST API）

以下案例依赖 xlwings PRO（Reports 系统、File Reader、REST API，均需许可），研读其**报表生成模式 / 无 Excel 只读引擎 / REST 数据接口**。`xlwings-demo-master/` 的 PRO 部分共 6 个子目录（见 4.2 子目录地图），按三类研读：

- **报告系**（`reporting/` + `reporting_bigmac/` + `reporting_filter_frames/` + `reporting_fund/`）：`create_report` 模板报告 + `Markdown`/`MarkdownStyle` 样式——"数据 → 模板 → 多格式输出"的报告管线。各子目录按数据源类型示范：
  - `reporting/`：`report_template.py` + `report_template.xlsx` + `reports101.ipynb`——Reports 入门样板
  - `reporting_bigmac/`：`_bigmac_index.py` + `big-mac-raw-index.csv` → `report_2020-01-14.pdf`——外部 CSV 数据源 → 报告输出
  - `reporting_filter_frames/`：`filter_frames.py` + `holdings.csv`——模板内过滤帧 / 列裁剪
  - `reporting_fund/`：`fund_template.py` + `data.pickle`——序列化数据源报告
- **`reader/`**（`xlwings File Reader.ipynb` + AAPL 系列样例）：PRO **File Reader**（`xw.Book(..., mode="r")`，0.28.0+，无许可抛 `LicenseError`）——**不启动 Excel 直接读文件**；官方文档（`docs/pro/reader.md`）与案例实测（结果一致性 / 速度基准 / 语法对照 / 命名区域 / 转换器）的深度技术研读见 4.4.1「reader/ 深度研读」；
- **`restapi/`**（`timeseries.xlsx`）：PRO **REST API** 的示例数据——`xlwings.server`/`rest.api`（见 1.7 xlwings Server 工程）以时间序列数据文件作为接口输出的演示数据集；目录仅含数据无代码，作为 REST 服务端联调样例。

- **`xlwings-factsheet-demo-main/`**（`demo.py`）：完整基金说明书工程。要点：
  - 三段式流水线：数据清洗 → 模板填充 → `create_xlsx_and_pdf_reports` 出 xlsx/pdf
  - `Markdown` 富文本单元格、多基金批量循环、`run_sheet` 参数页驱动
  - `data/` 为多基金 CSV 数据，`requirements.txt` 声明 PRO 依赖；
- **`xlwings-eikon-master/`**（`_report_template.py`/`sample3.py`）：Eikon 行情数据经 `create_report` 出报表——PRO 报告层与外部 API 数据链路的组合；

#### 4.4.1 reader/ 深度研读（File Reader 只读引擎：官方文档 + 案例实测）

官方文档 `docs/pro/reader.md` 与案例 `xlwings-demo-master/reader/`（`xlwings File Reader.ipynb` + AAPL 系列 4 个数据文件）互为印证，构成 File Reader 的完整研读对象。

**官方文档要点（`docs/pro/reader.md`）**

- 定位：xlwings PRO ≥ v0.28.0 的免 Excel 只读引擎；官方口径单 sheet 读取比 `pandas.read_excel()` 快 **5–25 倍**（视内容 / 文件格式 / Python 版本）
- 格式：`xlsx` / `xlsm` / `xlam`、`xlsb`、`xls`；读文件不依赖 Excel 安装（但要求工作簿**先保存**才能读到改动）
- API 形态：
  - 打开 `xw.Book("file", mode="r")`；推荐 `with` 上下文管理器自动 close（不用则 `book.close()`）
  - 整表 `sheet.cells.value`；转换器 `sheet["A1:B2"].options("df", index=False).value`（DataFrame / NumPy / list / scalar / dict）
  - 命名区域 `sheet["myname"].value` / `.address`，或 `book.names` 遍历 `name.refers_to_range.value`
  - 动态区域 `sheet["A1"].expand().value` 或 `options(expand="table")`
  - 错误值默认读为 `None`；`options(err_to_str=True)` 读成 `"#N/A"` 等字符串
  - datetime 转换全格式支持（含 `xlsb`）
- 限制（构建应用系统前必须核对）：
  - 安装渠道仅 pip / conda-forge（官方 Anaconda 通道暂无）
  - `expand()` 当前低效，大动态区域明显变慢
  - sheet 级命名区域显示名不带 sheet 前缀、`names` 只能经 book 访问、公式 / 常量类定义名不支持
  - Excel 表格（table）不支持
  - 除 `err_to_str` 外的非默认 options（`dates` / `empty` / `numbers`）当前低效
  - 只支持单元格值、不支持公式；**只读、不支持写入**

**案例构成（`reader/` 目录）**

- `xlwings File Reader.ipynb`：官方演示 notebook（环境：Python 3.10.6 / xlwings 0.28.3 / OpenPyXL 3.0.10 / pyxlsb 1.0.10 / pandas 1.5.1）
- 数据文件 4 个：
  - `AAPL.xlsx` / `AAPL.xlsb`：1980 年代苹果日线 OHLC 历史行情（约 1 万行）
  - `AAPL2.xlsx`：新行情 + `statistics` 命名区域（`$B$4:$C$7`）+ 标题 / 表头布局
  - `AAPL.xls`：旧格式兼容样例

**实测演示（notebook 复现）**

- 交互模式 vs 读模式结果一致性：`xw.Book("AAPL.xlsx")`（交互，需 Excel）与 `xw.Book("AAPL.xlsx", mode="r")`（读模式）读同一区域，`df1.equals(df2)` → `True`
- 速度基准（全表 `cells.options("df")`）：
  - `AAPL.xlsx`：pandas / OpenPyXL 855 ms vs xlwings 101 ms（≈ 8.5×）
  - `AAPL.xlsb`：pandas / pyxlsb 591 ms vs xlwings 38.8 ms（≈ 15×）——**xlsb 源更快**
  - 相等性核对需先 `df_xlsx["Volume"] = df_xlsx["Volume"].astype(float)`；xlsb 侧再 `df_xlsb.index = df_xlsb.index.astype(float)`（日期序列读作 float）——两引擎数值类型默认存在差异，比对前需对齐
- 语法对照（`AAPL2.xlsx` 读 `B11:G13` 两行）：
  - pandas：`read_excel(sheet_name=0, skiprows=10, usecols="B:G", nrows=2, index_col="Date")`
  - OpenPyXL：`load_workbook(read_only=True, data_only=True)` + `iter_rows(min_row=11, max_row=13, min_col=2, max_col=7, values_only=True)` 手工拼 DataFrame
  - xlwings A1 记号：`book.sheets[0]["B11:G13"].options("df").value`（一行）
  - xlwings 切片记号：`book.sheets[0][10:13, 1:7].options("df").value`（0-based、上界排他，与 pandas 一致）
- 命名区域：
  - 直接访问：`sheet["statistics"].options("df").value` → DataFrame（Statistic / Value）
  - `book.names` 遍历：`name.name` / `name.refers_to_range`（绝对地址 `AAPL!$B$4:$C$7`）/ `name.refers_to_range.sheet.name`（`AAPL`）
  - 实证官方限制：sheet 级名称显示名不带 sheet 前缀（`statistics` 而非 `AAPL!statistics`），但 `refers_to_range` 可回推 sheet 与绝对地址
- 转换器：scalar（`sheet["B2"].value` → `'Apple Inc. (AAPL)'`）、list（`sheet["B4:C4"].value` → `['Statistic', 'Value']`）、dict（`sheet["B5"].expand().options(dict).value` → `{'Enterprise Value': '2.42T', 'Trailing P/E': '2.49T', 'Forward P/E': 24.86}`）

**工程借鉴（构建应用系统）**

- 只读大文件批处理（CI / 服务器 / 无 Office 环境）用 File Reader 替代 COM：毫秒级、免 Excel、结果与交互模式逐位一致
- 大数据量优先取 `.xlsb` 源（实测快约 4 倍于 `.xlsx` 读取差）；整表用 `sheet.cells`——对应 `_xlcalamine.py` 的 `raw_value` 全表缓存设计（见 1.2.4）
- 只读场景一律 `with xw.Book(..., mode="r")` 防句柄泄漏
- 需要公式 / 写回 / 活会话时仍走 COM 通道——reader 是纯读通道，通道决策见 1.2.1
- 错误值语义判断用 `err_to_str=True` 保留 `"#N/A"` 文本（默认读作 `None` 会丢语义）
- 命名区域读取统一走 `book.names` + `refers_to_range`（同时拿名称、绝对地址与所属 sheet）；大动态区域避免 `expand()`（低效），改用 `cells` 整读后自行裁剪

---

### 4.5 taxi-duckdb（指针）

**定位**：xlwings 官方"一个 .xlsx 自包含完整数据应用"样板——工作簿内嵌 Office Web 加载项（xlwings Lite / Pyodide 0.27.5 浏览器端 Python），任务窗格提供两个 `@script` 按钮，用 DuckDB 直接查询纽约出租车 Parquet 数据并回写 Excel 报告图表。**不依赖 PRO/Server**，是 1.2.5 officejs 引擎 / Lite 的完整实践案例。

**细节去向（单点原则）**：该案例的完整研读（xlsx 内嵌结构解密、Python 代码双编码嵌入与执行链、main.py 双脚本四层管线、DuckDB 浏览器端使用模式、桌面端 DuckDB API 扩充与 excel 扩展、工程借鉴）已迁至 **`references/15-xlwings-lite-guidance.md` 阶段七**；内嵌源码提取工具 `extracted/extract_webextension.py` 及其用法见该案例 `extracted/README.md`。本处不再重复任何细节。

### 4.6 数据服务层案例研读：SQLite 数据访问模式（SQL）

Excel 应用系统对接关系型数据的两种官方样例：源码包内置 `examples/database/database.py`（SQLite → Excel 回填，脚本式）与社区工程 `Excel_udf_itus-main/ebitda_margins_data_udf.py`（SQLite + pandas 的 UDF 数据服务，公式调用式）。两者构成"SQL 数据访问层"的完整写法。

**官方 `database/database.py`（源码包 `examples/database/`，配合 `chinook.sqlite`）**：
- 数据库文件放 **Excel 同目录**：`os.path.join(os.path.dirname(wb.fullname), "chinook.sqlite")`——`wb.fullname` 取调用方工作簿全路径
- 连接：`sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)`——声明类型自动转换
- **参数化查询**：`cursor.execute(sql, (playlist_id,))` 用 `?` 占位符，杜绝拼接注入
- **列名与数据分离写出**：`col_names = [col[0] for col in cursor.description]` 取表头 → `sht.range("A9").value = col_names`；`sht.range("A10").value = rows`（`fetchall()` 列表直接写 Excel，自动扩展区域）
- 清理旧数据：`sht.range("A9").expand().clear_contents()`
- **ComboBox 数据绑定**（`combobox()`）：数据先写隐藏 sheet，再经 COM 原生对象配置控件——`OLEObjects("ComboBox1").Object.ListFillRange = "Source!$A$1:..."`、`BoundColumn = 1`、`ColumnCount = 2`、`ColumnWidths = 0`（显示名 + 存 ID 的经典两列下拉）
- 资源释放：`cursor.close()` + `con.close()` 显式收尾

**`Excel_udf_itus` 的 `ebitda_margins_data_udf.py`（UDF 数据服务模式，10210B）**：
- **配置驱动可迁移**：`config.ini` 读取 `db_type`/`db_path`/`table_name`——`db_type=sqlite` 走 sqlite3，否则走 psycopg2（RDS），同一套代码切换 SQLite ↔ PostgreSQL
- 连接内建索引保障：首次连接执行 `CREATE INDEX IF NOT EXISTS idx_<table>_lookup ON <table> (index_name, accord_code, date)`（全局标志 `_index_checked` 只建一次）
- **查询缓存**：`@lru_cache(maxsize=64)` 包装 `_cached_query(sql, params_key)`，参数元组化（`tuple(str(p) for p in params)`）作缓存键；`clear_cache()` UDF 调用 `_cached_query.cache_clear()` 供数据更新后失效
- **pandas 读 SQL**：`pd.read_sql_query(sql, conn, params=params_key)`——返回 DataFrame，`df.empty` 判断空结果
- **日期规范化**：Excel 输入（"2024-03-31" 10 位 / 带时间 19 位）→ `datetime.strptime` → 统一 DB 格式；SQL 侧 `date = ? OR date LIKE ?` 用 `日期+"%"` 前缀匹配
- **UDF 返回契约**：`@xw.ret(expand='table')` + `return [df.columns.tolist()] + df.values.tolist()`——表头行 + 数据行直接进 Excel 公式区域
- **观测性**：`RotatingFileHandler` 日志（函数名 / 参数 / 耗时 ms / 成败状态）、`time.perf_counter()` 计时

**SQL 写法要点（两例共性）**：
- 一律参数化（`?` 占位），表名白名单来自 config，不拼接用户输入
- `ORDER BY weights DESC` / `BETWEEN` / `LIKE` 前缀匹配 / `INNER JOIN ... ON` 多表关联，均走标准 SQL 而非 ORM
- 返回前 `df.empty` 分支给出友好提示串而非空表

**工程借鉴**：① "数据库文件与工作簿同目录"免除路径配置；② `cursor.description` → 表头 + `rows` → 数据的两段式写出省去手写列名；③ 查询结果先缓存（lru_cache）后经 UDF 暴露，兼顾 Excel 公式高频重算与后端压力；④ config 驱动 DB 类型切换是"单机 SQLite 起步、线上 RDS 演进"的可迁移架构。

---


### 4.7 DuckDB 桌面端数据访问模式（taxi-duckdb 桌面端分支）

taxi-duckdb 案例的 **Lite/Pyodide 浏览器端**研读在 `references/15-xlwings-lite-guidance.md` 阶段七 7.5；本节承接其**桌面端分支**（脱离 Pyodide、本地 Python 场景，供升级到桌面版 / 自建数据层时参考）——属纯桌面端 Python 数据层知识，与 xlwings Lite 无涉，故集中于此（与 4.6 SQLite 模式并列，构成"关系型 / 文件型数据层"两种样板）。

**桌面端 DuckDB API 扩充**：
- **多源混查**：`SELECT ... FROM read_parquet('a.parquet') p JOIN read_xlsx('b.xlsx') x ON p.id = x.id JOIN read_csv_auto('c.csv') c ON ...`——跨格式关联免导入；
- **视图与内存注册**：`CREATE VIEW v AS SELECT ...` 复用复杂查询；`duckdb.register('tbl', df)` 注册 pandas DataFrame 为可 SQL 查询表；
- **导出通道**：`.df()`（pandas）/ `.pl()`（polars）/ `.arrow()`（PyArrow）/ `.fetchall()` / `.fetchone()`；
- **性能特性**：向量化执行 + **filter pushdown**（WHERE 下推到 Parquet 元数据层）——1 亿行级 `count(*)` 亚秒级。

**DuckDB 官方 excel 扩展**（桌面端直读直写 .xlsx，`INSTALL excel; LOAD excel;`）：
- **读取**：`SELECT * FROM 'test.xlsx'` 直接表名读（第一个 Sheet）；`read_xlsx('f.xlsx', header=true, sheet='Sheet2', range='A1:B2', ...)` 命名参数控制；`COPY tbl FROM 'f.xlsx' WITH (FORMAT xlsx, HEADER)` 导入已有表（不做类型推断）；
- **写入**：`COPY tbl TO 'out.xlsx' WITH (FORMAT xlsx, HEADER true, SHEET 'Sheet1')`；纯数据无样式/公式/合并单元格；
- **与 xlwings 通道分工**：DuckDB excel 扩展 = 无 Excel 进程的批量读/写（ETL/CI/大数据量）；xlwings COM / File Reader / calamine = 与 Excel 应用交互（样式、表格对象、图表、活会话）。组合示例：

```python
import duckdb, xlwings as xw
# ① DuckDB 直读 xlsx → SQL 分析 → xlwings 呈现
df = duckdb.sql("SELECT * FROM read_xlsx('data.xlsx', header=true) WHERE 金额 > 1000").df()
xw.Book().sheets[0]["A1"].options(index=False).value = df
# ② DuckDB 直写 xlsx（纯数据落盘）
duckdb.sql("COPY (SELECT * FROM read_xlsx('data.xlsx')) TO 'out.xlsx' WITH (FORMAT xlsx, HEADER true, SHEET '分析结果')")
```

**工程借鉴**：① "查询层（DuckDB SQL）与展示层（pandas/xlwings）解耦"——数据文件不载入内存，SQL 层物化再交给 Excel；② excel 扩展满足"无 Excel 进程的批量 ETL"，与 xlwings 活会话互补，按"是否需应用交互"选通道；③ 与 4.6 SQLite 模式对照：SQLite 适合结构化关系数据（事务/约束），DuckDB 适合文件型/分析型数据（Parquet/CSV/xlsx 直读），两者都是"config/代码驱动 DB 类型切换"的可迁移数据层样板。

---
## 可配套阅读

- `xlwings/docs/pro/`（PRO 官方文档：license_key.md、reader.md、release.md、reports/）
- `xlwings/xlwings/pro/`（PRO 源码）
- `examples/ReadMe.md`（各实战项目的详细说明）
