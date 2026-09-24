# VBA 构建指南（6.5.5 聚合）

> 本文件聚合 VBA 侧全部既有指南（编码规范 / 反模式清单 / 现有文件迭代链路），并新增 VBA-Docs 官方对象模型速查与单元格引用 7 法。VBA 侧编码以本文档为准。

## 一、VBA 编码补充规则（原 09）

### 范围处理建议

- 边界识别：明确 `lastRow` 口径（UsedRange vs End(xlUp) vs ListObject），避免误判
- 批量清理：`Range.ClearContents` 整块，不逐格
- `Value2` 读取后重新确认格式语义（日期序列值、文本转数值、前导零丢失）
- 输出列较多时先做列 Schema 设计
- 日期时间列写入后补 `NumberFormat`
- 长数字字段（15 位以上）先设文本再写入
- 结构化明细表优先 `ListObject`（自动扩展/样式/汇总）

### 结构设计建议

- 一个过程只做一件主要事情

### 错误处理与清理

推荐模式：`On Error GoTo` + 统一出口 + finally 式恢复应用状态；**禁止** `On Error Resume Next` 全程挂起、出错 `End` 硬退。

## VBA API 查证流程（写代码前必做）

写 VBA 代码前，禁止凭记忆写 API 调用。按以下三步查证：

### 第一步：用索引库定位 API（`../vba_docs_index.db`）

索引库是 `../VBA-Docs/` 的检索入口，表 `api_ref(app,kind,name,title,rel_path,description)`，覆盖 Office 各应用 API 记录，已建索引 `idx_app`/`idx_kind`/`idx_name`。`rel_path` 直接指向 `../VBA-Docs/` 下的对应参考页。

```bash
# 查具体 API（如 Excel.Workbook.Save）→ 得到参考页相对路径
python -c "import sqlite3;c=sqlite3.connect(r'../vba_docs_index.db');print(*c.execute(\"SELECT app,kind,name,rel_path FROM api_ref WHERE name='Workbook.Save'\").fetchall(),sep='\n')"

# 查某应用的某类 API（如 Excel 全部事件）
python -c "import sqlite3;c=sqlite3.connect(r'../vba_docs_index.db');print(*c.execute(\"SELECT name,rel_path FROM api_ref WHERE app='Excel' AND kind='event'\").fetchall(),sep='\n')"

# 按关键词模糊检索（如所有与 ListObject 相关的条目）
python -c "import sqlite3;c=sqlite3.connect(r'../vba_docs_index.db');print(*c.execute(\"SELECT name,rel_path FROM api_ref WHERE name LIKE '%ListObject%' LIMIT 30\").fetchall(),sep='\n')"
```

**kind 类型与查证重点**：
- `method` → 看调用语法与参数
- `property` → 看读写权限（Read/Write/Read-only）
- `event` → 看事件签名与触发时机（写类模块事件时）
- `enumeration` → 查枚举成员值与含义（赋给属性/参数前）
- `object` → 看对象层级关系（定位父对象引用）

⚠️ 索引覆盖 Office 各应用 API 参考，**不含 VBIDE 对象模型本体**（后者查 `references/vbide-object-model.md`，若存在）。

### 第二步：打开参考页取权威语法（`../VBA-Docs/api/`）

用第一步得到的 `rel_path` 打开 `../VBA-Docs/api/` 下对应参考页。每页结构固定：
- **Syntax**：权威签名，含参数语义
- **Remarks**：边界行为、坑点（如首次保存需用 `SaveAs`、`Saved` 属性标记等）
- **Example**：可直接借鉴/改写的官方示例代码

需要概览某应用对象模型时，参考 `../VBA-Docs/api/overview/` 下的应用级概览页。

### 第三步：校准语言基础（`../VBA-Docs/Language/`）

对语言层面（非 API 层面）写法不确定时查阅：
- **Reference/**：内置常量、关键字全集、函数签名、语言级事件、条件编译、错误码
- **Concepts/Getting-Started/**：变量声明、常量 vs 变量、数组声明与 ReDim、`Set` 语义、过程调用约定、命名冲突规避
- **How-to/**：按"我要做什么"查阅操作指南（28 篇）

### 查证后落笔

三步查证完成后，按本文件"一、VBA 编码总纲"的规范落笔。查证记录（API 名、参考页路径、关键参数）写入设计文档或代码注释，便于后续维护。

## 二、VBA 反模式清单（原 10，18 条速查）

| # | 反模式 | 推荐替代 |
|---|--------|---------|
| 1 | 逐格读取、逐格写入 | 数组整块读写 |
| 2 | 循环里反复改格式 | 写完后一次性设格式 |
| 3 | 到处 `Select`/`Activate` | 直接对象引用 |
| 4 | 无脑 `ActiveSheet`/`Selection` | 显式 `ThisWorkbook.Worksheets(...)` |
| 5 | 循环里反复计算边界 | 边界算一次存变量 |
| 6 | `UsedRange` 充当一切数据边界 | 按列 End(xlUp)/ListObject 定界 |
| 7 | 全程 `On Error Resume Next` | 局部错误处理 + 恢复 |
| 8 | 出错直接 `End` | 统一出口清理资源 |
| 9 | 长任务疯狂 `DoEvents` | 状态栏反馈/分批 |
| 10 | 每处理一点弹一次 `MsgBox` | 汇总一次性提示 |
| 11 | UDF 里做副作用 | UDF 纯计算 |
| 12 | 逻辑绑死固定工作表名/`ActiveSheet` | 显式工作簿归属 |
| 13 | 明细结果只输出普通区域不转 `ListObject` | 结构化表输出 |
| 14 | 非必要直接覆盖用户原始数据/原结果区 | 另区输出/先备份 |
| 15 | `ListObject` 扩容缩容前不检查周边 | 先检查再操作 |
| 16 | 15 位以上长数字按数值写入 | 文本格式写入 |
| 17 | 日期列写入后不补 `NumberFormat` | 写后补格式 |
| 18 | `ListObject` 穿插公式列仍整块覆盖 | 数据区/公式区分开写 |

## 三、Excel 对象模型速查（源自 VBA-Docs/api/Excel.*，官方权威）

> 完整逐成员参考在 `../VBA-Docs/api/Excel.*.md`；以下为加载项高频对象一行式。

| 对象 | 高频成员 | 用途 |
|------|---------|------|
| `Application` | `ActiveWorkbook`/`ScreenUpdating`/`Calculation`/`EnableEvents`/`Run` | 应用级状态与执行 |
| `Workbook` | `Worksheets`/`Names`/`SaveAs`/`Close`/`FullName` | 工作簿级操作 |
| `Worksheet` | `Range`/`Cells`/`ListObjects`/`UsedRange`/`Protect` | 表级操作 |
| `Range` | `.Value2`/`.Formula`/`.Formula2`/`.NumberFormat`/`.ClearContents`/`.End(xlUp)`/`.CurrentRegion` | 单元格区域读写 |
| `ListObject` | `.ListRows`/`.ListColumns`/`.HeaderRowRange`/`.DataBodyRange`/`.Resize` | 结构化表格 |
| `QueryTable` | `.Refresh`/`.Connection`/`.CommandText` | 外部数据查询表 |
| `Worksheet`/`Workbook` 事件 | `Worksheet_Change`/`Worksheet_BeforeDoubleClick`/`Workbook_Open` | 事件驱动 |
| `UserForm`/`Control` | `Controls`/`Show`/`Hide` | 窗体交互 |

**Formula vs Formula2**：`Formula2` 支持动态数组（溢出引用 `#` 运算符）；`Formula` 为传统公式。差异详见 `../VBA-Docs/excel/Concepts/Cells-and-Ranges/range-formula-vs-formula2.md`。

## 四、单元格与区域引用（源自 VBA-Docs excel/Concepts/Cells-and-Ranges/refer-to-* 系列，共 9 篇）

| 引用方式 | 示例 | 适用 |
|---------|------|------|
| A1 记号 | `Range("B2:D5")` | 最常见 |
| 索引数字 | `Cells(2, 4)` | 循环/动态定位 |
| 快捷记号 | `[A1]` | 快速原型（生产少用） |
| 相对引用 | `Range("A1").Offset(1, 0)` | 相对移动 |
| 命名区域 | `Range("MyName")` / `Names` | 语义化 |
| 行/列 | `Rows(2)`/`Columns("B")` | 整行整列 |
| 多区域 | `Union(Range("A1"), Range("C3"))` | 非连续区域 |

官方全文：`../VBA-Docs/excel/Concepts/Cells-and-Ranges/refer-to-cells-and-ranges-by-using-a1-notation.md` 等 9 篇。

## 五、现有文件迭代（原 04）

### 铁律：现有文件流程必须走脚本，禁止 AI 自由发挥

```
import_existing_workbook.ps1（只读分析）
        │  产出 baseline/ 副本 + inventory.json + src/ 全量提取
        ▼
AI 在 src/ 上做增/删/改（纯文件编辑，不碰 Excel）
        │
        ▼
apply_vba_to_workbook.ps1（回写交付）→ dist/<原名>.xlsm/.xlam
```

禁止：自己写 COM 读写旧 VBA/Ribbon、新建空白簿搬运、绕过 baseline 改原文件。

### 第一步：导入分析

```powershell
powershell -ExecutionPolicy Bypass -File "<skill根>\scripts\import_existing_workbook.ps1" `
  -SourceWorkbook "<用户提供的工作簿路径>.xlsx" -ProjectRoot "<项目根目录>"
```

产出：`baseline/` 只读副本、`src/modules|classes|forms|documents|ribbon/` 全量提取、`baseline/inventory.json`（工作表/组件/回调/命名区域全景）。AI 分析入口：**先读 inventory.json**。

### 第二步：AI 修改 src/（纯文件编辑）

| 意图 | 操作 |
|---|---|
| 改已有模块/类/文档模块 | 编辑 `src/` 对应 `.vba`，构建时同名整体替换 |
| 改窗体事件代码 | 编辑 `src/forms/X.code.vba`（`.frm` 布局不动） |
| 新增模块/类/窗体 | `src/modules|classes|forms/` 新增文件 |
| 删除基线组件 | 删对应文件，构建加 `-RemoveMissingModules` |
| 改/加 Ribbon | 编辑 `src/ribbon/customUI14.xml`（2010 版），删除文件则保留基线原 Ribbon |

### 第三步：回写交付

```powershell
powershell -ExecutionPolicy Bypass -File "<skill根>\scripts\apply_vba_to_workbook.ps1" -ProjectRoot "<项目根目录>"
```

要点：内存完成全部 VBA 增删改后一次性 `SaveAs` 落盘 dist（先改内存后 SaveAs，防污染 baseline）；VBA upsert（同名替换/异名新增/显式删除）；Ribbon zip 级整体回注；不生成安装脚本（那是 build_addin.ps1 的职责）。

### 与 build_addin.ps1 的分工

| | build_addin.ps1 | import + apply |
|---|---|---|
| 场景 | 从零 xlam（quickstart） | 现有文件增强 |
| 构建 | 注入 VBA + Ribbon + 安装脚本 | VBA upsert + Ribbon 回注 |

### 必须另存副本的场景

原文件是正式生产版且无备份、用户担心改坏、改动深度交叉、含不可恢复数据、AI 未吃透旧逻辑——一律走 baseline + dist 副本链路。

### 原文件里默认要保留的内容

- 现有 VBA 模块/类/窗体/文档模块：除非显式删除，否则原样保留（import 全量提取，apply 同名替换/异名新增）
- 现有 Ribbon XML：`src/ribbon/` 无文件时保留基线原 Ribbon；有文件则整体替换
- 工作表结构、命名区域、数据验证、条件格式：不主动改动
- 现有公式：不覆盖（除非业务明确要求）

### 蓝图实例化时的覆盖优先级

`project-blueprints/` 覆盖 `internal-bases/` 同名文件；实例化后用 `instantiate_blueprint.ps1` 生成项目，再遵循本步骤构建。蓝图与现有文件迭代是两条独立路径，不要混用（蓝图用于从零，现有文件迭代用于增强旧文件）。

## 六、官方 VBA 文档导航（VBA-Docs）

| 路径 | 内容 |
|------|------|
| `../VBA-Docs/api/Excel.*.md` | 全部 Excel 对象逐成员（Application/Workbook/Worksheet/Range/ListObject/...） |
| `../VBA-Docs/excel/Concepts/Cells-and-Ranges/refer-to-*.md` | 单元格/区域引用 9 篇 |
| `../VBA-Docs/excel/Concepts/Cells-and-Ranges/range-formula-vs-formula2.md` | Formula/Formula2 差异 |
| `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/create-a-user-form.md` 等 | UserForm 系列 |
| `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/activex-controls.md` | ActiveX 控件 |
| `../VBA-Docs/excel/Concepts/Events-WorksheetFunctions-Shapes/` | 事件驱动（Worksheet/Workbook/Application 事件）+ VBA 调用 WorksheetFunction + 形状操作 |
| `../VBA-Docs/excel/Concepts/Excel-Performance/` | 性能优化（计算性能/性能瓶颈/优化技巧），配合本文件"编码总纲"性能规则 |
| `../VBA-Docs/excel/Concepts/Workbooks-and-Worksheets/` | 工作簿/工作表操作（打开/创建/引用/保存） |
| `../VBA-Docs/Language/Reference/` | VBA 语言参考（语句/函数/错误代码/数据类型） |
| `../VBA-Docs/excel/Concepts/Miscellaneous/concepts-excel-vba-reference.md` | Excel VBA 概念总入口（从这里进入各主题） |

