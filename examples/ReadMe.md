---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 6789603449abf5a04d13cbbfc9bbf37e_ae1ccd11a83d11f1972e525400287e28
    ReservedCode1: vZwuZ8yLk4IM74QCRLwX7AEinrvhNhw4TJohGISLd1GK7K/7i4BMPxgytnBn912Ot69WaoL6UGxrjgrAAnv3EOAoZH+4Swde9uRIhedwI0LepBSFxre43+j4VmoIqqbIth925S81RMn7VGyTigsQUqRb5Mcid2tpNeadTnCwI/SH5s6gnIJcfSCTCMo=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 6789603449abf5a04d13cbbfc9bbf37e_ae1ccd11a83d11f1972e525400287e28
    ReservedCode2: vZwuZ8yLk4IM74QCRLwX7AEinrvhNhw4TJohGISLd1GK7K/7i4BMPxgytnBn912Ot69WaoL6UGxrjgrAAnv3EOAoZH+4Swde9uRIhedwI0LepBSFxre43+j4VmoIqqbIth925S81RMn7VGyTigsQUqRb5Mcid2tpNeadTnCwI/SH5s6gnIJcfSCTCMo=
---

# xlwings 技能案例集总览（examples/）
> 本目录下的示例仓库为官方/第三方开源项目的本地克隆，**未随本仓库分发**（以本文件占位并链接官方来源）；官方克隆方式：`git clone <url>`。
>
> | 子目录 | 官方来源 |
> |---|---|
> | `Excel_udf_itus-main` | <https://github.com/Kiran3002/Excel_udf_itus> |
> | `cross-check-reports-main` | <https://github.com/xlwings/cross-check-reports> |
> | `excel-automated-testing-master` | <https://github.com/xlwings/xlwings-automated-testing> |
> | `python-for-excel-course-main` | <https://github.com/xlwings/python-for-excel-course> |
> | `simulation-demo-master` | <https://github.com/xlwings/simulation-demo> |
> | `static-excel-test-master` | <https://github.com/xlwings/static-excel-test> |
> | `taxi-duckdb-main` | <https://github.com/xlwings/taxi-duckdb> |
> | `xl-pq-handler-master` | <https://pypi.org/project/xl-pq-handler/> |
> | `xlwings-demo-master` | <https://github.com/xlwings/xlwings-demo> |
> | `xlwings-eikon-master` | <https://github.com/xlwings/xlwings-eikon> |
> | `xlwings-factsheet-demo-main` | <https://github.com/xlwings/xlwings-factsheet-demo> |
> | `xlwings-server-main` | <https://github.com/xlwings/xlwings-server> |


本目录集中存放本技能内置的场景样例与演示仓库。SKILL.md 各场景任务块按主题引用本目录中的案例；**案例的目录结构、设计思路、运行方式与个性化改造说明一律以本文件（及仓库自带 README）为准**，SKILL.md 只做简要介绍与引用。

所有仓库均为 GitHub ZIP 下载结构，含一层同名嵌套目录（`examples/仓库名/仓库名/...`），本文件中的相对链接已按此写全。

## 案例总览

| 仓库 | 定位 | 引用场景 |
| --- | --- | --- |
| xlwings-factsheet-demo-main | Reports 模板报告完整工程（Excel 模板批量生成基金式 factsheet 报告） | 场景 D · 7.7 源码研读（需 PRO 许可，不执行） |
| xlwings-eikon-master | Eikon/Refinitiv 行情数据接入（相关性、Reports、博客报告、模拟、实时流五种形态） | 场景 A · 任务块 4.8 |
| simulation-demo-master | Excel/Web 双端快速原型（同一份蒙特卡洛计算逻辑同时驱动 Excel 与 Flask Web） | 场景 A · 任务块 4.9 |
| xl-pq-handler-master | Power Query（.pq）脚本库管理工具：从工作簿批量提取 M 代码、按依赖顺序注入回工作簿 | 场景 A · 任务块 4.10 |
| static-excel-test-master | 静态坏引用扫描（#REF!/#DIV/0!/#VALUE! 残留检查，openpyxl 实现，无 Excel 依赖） | 场景 D · 任务块 7.9 |
| excel-automated-testing-master | 自动化回归测试骨架（xlsm 工作簿的 pytest/unittest 驱动测试） | 场景 C · 工作流第 9 步、场景 B · 5.2 |
| cross-check-reports-main | 报表交叉核验（两份同构报表逐单元格断言） | 场景 D · 任务块 7.9 |
| Excel_udf_itus-main | 数据表查询驱动 UDF 完整工程（SQLite → @xw.func → example.xlsm 公式调用） | 场景 B · 5.2 |
| xlwings-demo-master | 官方综合演示集（basics/correlation/performance/frozen/reporting 等 13 个主题） | 场景 A/B/C/D 多处引用（A 整链、B·5.2、C 测试、D·7.5/7.7/7.8 等） |
| xlwings-server-main | 官方 xlwings Server 服务端工程（为 Excel/Google Sheets 提供 Python 支持，可自托管） | 场景 D · 7.8 源码研读（不部署） |
| python-for-excel-course-main | 官方入门视频课程（2018）配套 Jupyter 讲义，按 0-7 Part 组织 | 场景 D · 7.5 |

---

## 1. xlwings-factsheet-demo-main：Reports 模板报告完整工程

**一句话定位**：以单个 Excel 模板 + 结构化数据，批量生成基金式事实清单（xlsx/PDF）的完整 xlwings Reports 工程。

**目录**：`xlwings-factsheet-demo-main/xlwings-factsheet-demo-main/`

| 文件/目录 | 职责 |
| --- | --- |
| `demo.py` | 三段式流水线实现：pre-processing（读 Markdown/Word/CSV 素材）→ report generation（渲染 Excel 模板并输出 PDF）→ post-processing（上传 S3 等分发渠道） |
| `demo.xlsm` | Excel 端入口，经典 RunPython 调用 |
| `template/template.xlsx` | 报告模板：`{{ }}` 占位符 + Excel 表格 + 隐藏的 `# chart_data` 图表数据表；`layout.pdf` 为该模板打印版 |
| `data/common/` | 公共素材：`disclaimer.md`（Markdown 免责声明）、`intro.docx`（Word 简介） |
| `data/funds/<基金名>/` | 每只基金一个目录，内含该基金 CSV 数据（`history.csv`、`holdings.csv`） |
| `reports/` | 已生成的样例输出（Fund A/B/C 的 xlsx 与 pdf，可不开工程直接查看） |
| `requirements.txt` | 依赖清单；另需 xlwings add-in |

**前置条件**：xlwings Reports 属 xlwings PRO 功能，需要 license key（可申请免费试用 https://www.xlwings.org/trial）；`pip install -r requirements.txt`；`xlwings addin install`。

**运行方式**（二选一）：
- 打开 `demo.xlsm` 点 Run 按钮（经 add-in 配置）；
- 代码驱动：`xw.Book(r'...\demo.xlsm').set_mock_caller()` 后调用 `demo.py` 的 `main()`，内部以不可见 Excel 实例渲染全部报告，无需人工点 Run。

**模板占位符要点**：模板单元格中的 `{{ holdings | noheader | sortdesc(4) | aggsmall(0.04, 4, "Other") | mul(100, 4) | columns(0, None, None, None, None, 4) }}` 之类过滤器链可在 Excel 模板内完成排序、聚合（<4% 归入 "Other"）、乘 100、列裁剪等数据处理，无需在 Python 写逻辑；以 `template/template.xlsx` 实际内容与 `demo.py` 渲染调用为准。

**个性化替换三方面**：
1. 更换数据：按数据契约填充 `data/common/` 与 `data/funds/<基金名>/`，增加目录即增加报告；
2. 更换模板：以 `template/template.xlsx` 为蓝本重设 `{{ }}` 占位符；
3. 更换输出渠道：修改 `demo.py` 的 postprocess 段（默认演示上传 S3，可改为邮件、网页等）。

> 官方使用文档：xlwings Reports 详见 `../xlwings-0.37.2/docs/pro/reports/`。

---

## 2. xlwings-eikon-master：Eikon 数据接入演示集

**一句话定位**：把 Eikon（Refinitiv/LSEG）行情与数据接入 Excel 的官方演示集，覆盖五类典型接入形态。

**目录**：`xlwings-eikon-master/xlwings-eikon-master/`

顶层文件：`slides.pdf`（官方演示 slides）、`interactive.ipynb`（交互式讲解）。

**五类形态**（各形态对应一组完整演示文件，执行时按选定形态取用，不混用）：

| 形态 | 目录/文件 |
| --- | --- |
| 拉行情做相关性分析并回填图表 | `correlation/correlation.py` + `correlation/correlation.xlsm` |
| 用 Eikon 数据渲染 Reports 模板 | `reporting/_report_template.py` + `reporting/report_template.xlsx` |
| 博客式报告脚本 | `reporting_blog/sample1.py`、`sample2.py`、`sample3.py`（含样例输出 xlsx） |
| 蒙特卡洛模拟 | `simulation/simulation.py` + `simulation/simulation.xlsm` |
| 实时流（Streaming API，macOS 演示） | `streaming_api_macOS/realtime_eikon.py`、`realtime_rdp.py` 及对应 xlsx |

**前置条件**：可用的 Eikon 环境（桌面终端或 API 授权）；`eikon` 包已安装；本地 xlwings 源码注入正常。不确定目标形态时以 `correlation`（最小完整闭环）为主。

---

## 3. simulation-demo-master：Excel/Web 双端原型

**一句话定位**：演示"同一份计算逻辑同时驱动 Excel 与 Web"的双端快速原型——以蒙特卡洛模拟为例，从 Excel 原型平滑迁移到 Flask Web 应用。

**目录**：`simulation-demo-master/simulation-demo-master/`

| 文件/目录 | 职责 |
| --- | --- |
| `simulation.py` | 纯计算逻辑（仅依赖 numpy），被两端原样复用——这是原型可迁移设计的核心 |
| `xlwings_app.py` | Excel 端驱动器：从 `simulation.xlsm` 输入单元格读参数（E3 模拟次数、E4 时间跨度、E5 时间步数、E6 收益率、E7 波动率、E8 起始价），调用 `simulate()`，把 5/50/95 百分位与样本路径写回输出区域并更新图表 |
| `simulation.xlsm` | Excel 工作簿（含输入区与图表 Chart 5） |
| `web_app.py` | Flask Web 应用：首页渲染 + `POST /run-simulation` 接口接收同参 JSON 并返回模拟结果 JSON |
| `templates/`、`static/` | Web 页面（图表展示） |
| `requirements.txt`、`screenshot.png` | 依赖清单与效果截图 |

**跑通步骤**：
1. Excel 端：进入工程目录，注入本地 xlwings 源码后运行 `python xlwings_app.py`（模块 `__main__` 已内置 `xw.Book('simulation.xlsm').set_mock_caller()` 并执行 `run_simulation()`，全程代码驱动）；也可打开 `simulation.xlsm` 点 Run。
2. Web 端：`pip install -r requirements.txt` 后运行 `python web_app.py`，浏览器访问 `http://localhost:5002`；以 HTTP 请求验证首页与 `/run-simulation` 接口返回模拟 JSON。
3. 验证两端一致：Excel 端与 Web 端共用同一个 `simulate()`，对比输出可验证逻辑一致性。

**改造为自有业务**：把 `simulation.py` 替换为用户业务计算逻辑，同步维护 Excel 端（`xlwings_app.py` 的参数单元格映射）与 Web 端（`web_app.py` 的表单/JSON 契约、`templates/` 与 `static/` 页面），改造后重复两端跑通与一致性验证。

---

## 4. xl-pq-handler-master：Power Query 库管理工具

**一句话定位**：Pythonic 的 Power Query（.pq）脚本库管理工具，可把 Excel 工作簿中的 Power Query（M 代码）批量提取为 `.pq` 文件库，或将脚本库中的查询按依赖顺序注入回指定工作簿（底层以 xlwings 驱动 Excel）。

**目录**：`xl-pq-handler-master/xl-pq-handler-master/`（Python 包源码在 `src/`，工程元数据见 `pyproject.toml`；`README.md` 有完整功能说明）

**安装**：`pip install xl-pq-handler`（自带 customtkinter、xlwings、pydantic、pyyaml、pandas、filelock 等依赖）。

**两种使用形态**：
- **UI 形态**：`python -m xl_pq_handler <脚本库目录>` 或 `pqmagic <脚本库目录>`，启动桌面管理界面，支持浏览/搜索/过滤 .pq、智能提取（从文件或已打开工作簿，保存前可预览代码/参数/数据源）、依赖感知注入（自动把查询及其依赖按正确顺序写入指定已打开工作簿）、元数据编辑（改 category 自动移动文件）、语法高亮与依赖图谱。
- **编程形态（PQManager）**：
  ```python
  from xl_pq_handler import PQManager
  manager = PQManager(r"D:\...\PowerQuery_Repo")
  manager.build_index()
  # 注入到已打开的工作簿
  manager.insert_into_excel(names=["Calculate_KPIs"], workbook_name="Monthly_Report_WIP.xlsx")
  # 从文件提取全部查询
  manager.extract_from_excel(category="Downloaded", file_path=r"C:\...\NewDataSource.xlsx")
  ```

**脚本库组织约定**：`.pq` 文件按类别目录存放（`API/`、`Helpers/`、`Reports/` 等），每个 `.pq` 文件头部带 YAML frontmatter 记录元数据：

```yaml
---
name: Clean_RawSales          # 查询名
category: Staging             # 类别（与目录名保持同步）
tags: [cleaning, sales, raw]  # 搜索标签
dependencies:                 # 依赖的其它查询
  - fn_FormatDate
description: 清理原始销售数据
version: 2.1
---
let
    ...
in
    ...
```

根目录 `index.json` 由工具自动生成维护，不要手工编辑。

---

## 5. static-excel-test-master：静态坏引用扫描

**一句话定位**：对交付的 xlsx 等静态产物做坏引用扫描的质量门禁工具——遍历工作簿全部工作表，检查 `#REF!`、`#DIV/0!`、`#VALUE!` 等错误值残留；基于 openpyxl 实现，**全程不需要 Excel 实例**。

**目录**：`static-excel-test-master/static-excel-test-master/`

| 文件/目录 | 职责 |
| --- | --- |
| `test_refs.py` | 坏引用扫描骨架：遍历目标文件全部工作表，检查错误值残留；命中错误时输出 `文件: 工作表` 清单 |
| `sales_transactions.xlsx` | 自测样本 |
| `.github/workflows/main.yml` | 最小 CI 流水线 |

**执行步骤**（骨架为 `test_refs.py`）：
1. **明确目标**。确定待质检工作簿/报表的绝对路径；确认是否存在"基准版 vs 新版"两组可核对报表。扫描对象可以是单个文件，也可以是目录下全部 `.xls*` 文件。
2. **静态坏引用扫描**。以 `test_refs.py` 为骨架，读取方式 `openpyxl.load_workbook(path, read_only=True, data_only=True)`；错误类型清单（`#REF!`/`#DIV/0!`/`#VALUE!` 等）可按需扩展；命中错误时输出 `文件: 工作表` 清单并反馈修正。
3. **汇总结论**。输出通过/不通过结论与差异明细。全程不修改被测文件；若任务要求修复，回到场景 A 主线流程（4.2-4.7）打开文件修正后重新质检。

> 配套：报表交叉核验（两份同构报表逐单元格断言）骨架见 `examples/cross-check-reports-main/`（`test_reports.py`，用 `xw.Book(path, mode='r')` 只读打开；**注意 `mode='r'` 需 xlwings PRO 许可**，无许可时仅执行本小节的静态扫描）。

---

## 6. Excel_udf_itus-main：金融数据查询 UDF 完整工程

**一句话定位**：以 sqlite 数据源（Nifty500 指数成分股库）驱动的金融数据 UDF 完整工程——数据库 → `@xw.func` Python 函数 → `example.xlsm` 中以公式直接调用，模拟"指数成分/权重/行业一键查询"的集成体验。

**目录**：`Excel_udf_itus-main/Excel_udf_itus-main/`

| 文件 | 职责 |
| --- | --- |
| `ebitda_margins_data_udf.py` | UDF 模块：读取 `config.ini` 配置，从 SQLite（或 AWS RDS）查询指数成分股数据并返回格式化表格；`@lru_cache` 缓存 + RotatingFileHandler 日志（`query_log.txt`） |
| `equity_index_constituents - nifty500.db` | SQLite 数据库（字段：accord_code、company_name、sector、mcap_category、date、weights、index_name） |
| `scema.sql` | 数据库建表脚本 |
| `config.ini` | 数据库连接配置（`db_path` 需改为本地实际绝对路径；支持切换 rds） |
| `example.xlsm` | Excel 调用示例工作簿（已配置 UDF 模块） |
| `images/` | 配置截图 |

**可用公式**（点击 xlwings 选项卡 Import Functions 后可像内置函数一样使用）：

```excel
=get_monthly_data("nifty_500", "2023-04-30")   -- 指定日期指数成分
=get_series("nifty_50", "2020-03-31", "2025-09-30")  -- 区间内成分与权重序列
=get_matrix("2023-04-30", "nifty_500")         -- 指定日期全部成分
=get_all_data("nifty_500")                     -- 全量数据
```

**配置要点**（详细步骤见仓库 `readme.md`）：
1. 编辑 `config.ini` 的 `db_path` 指向仓库内 `.db` 文件绝对路径（反斜杠，不加引号）；
2. Excel 启用"信任对 VBA 项目对象模型的访问"；安装 xlwings add-in 并在 VBA References 勾选 xlwings；
3. xlwings 选项卡中设置 Interpreter（Python 路径）、Python Path（本工程目录）、UDF Modules（`ebitda_margins_data_udf`）；
4. 点击 Import Functions 后即可在单元格使用上述公式。

---

## 7. xlwings-demo-master：官方综合演示集

**一句话定位**：xlwings 官方综合演示仓库，覆盖从基础对象操作到 UDF、性能、Reports、冻结部署、测试等十余个主题；SKILL.md 各场景多处把它作为"拉数据-算指标-写图表"整链与部署/教学的样例源。仓库无独立 README，以目录名自解释。

**目录**：`xlwings-demo-master/xlwings-demo-master/`

| 主题目录 | 内容 | 引用场景 |
| --- | --- | --- |
| `basics/` | 最小脚本+UDF 骨架（`xwdemo.py` + `xwdemo.xlsm`，含 `xw.Book.caller()`/`set_mock_caller()` 写法） | A 通用、D·7.5 演示兜底 |
| `correlation/` | 拉历史行情做相关性分析并写图表（`correlation.py` + `correlation.xlsm`） | A 整链参考、D·7.5 |
| `interactive/` | Jupyter 交互示例（`interactive.ipynb`，AAPL.xlsx） | D·7.5 |
| `performance/` | 性能主题：数组读写实测（`arrays.ipynb`、`raw.ipynb`）、UDF 写法（`udfs.py` + `udfs.xlsm`）、数组 UDF（`arrayudf/arrayudf.py`） | A 性能参考、B·5.2 |
| `reader/` | Reader 引擎：无 Excel 读取 .xls/.xlsb/.xlsx（`xlwings File Reader.ipynb` + AAPL 系列样例文件） | D·7.7 Reader 引擎源码研读 |
| `reporting/` | Reports API 入门（`reports101.ipynb`、`report_template.py` + 模板 xlsx） | D·7.7 Reports 源码研读 |
| `reporting_bigmac/` | Big Mac 指数报告（`_bigmac_index.py` + 已生成 `report_*.pdf`） | D·7.7 Reports 源码参考 |
| `reporting_filter_frames/` | 模板过滤器（排序/聚合/列裁剪等，`filter_frames.py` + `holdings.csv`） | D·7.7 Reports 过滤器源码参考 |
| `reporting_fund/` | 基金报告场景（`fund_template.py` + `fund_template.xlsx` + data.pickle） | D·7.7 报告模板源码参考 |
| `frozen/` | PyInstaller 冻结部署（`demo.py`、`demo.spec`、`demo.xlsm`） | D·7.8 RunFrozenPython 机制源码研读 |
| `restapi/` | REST API 主题（`timeseries.xlsx` 演示输出） | D·7.7 REST API 源码研读 |
| `simulation/` | 蒙特卡洛模拟（`simulation.py` + `simulation.xlsm`，与 UDF 异步模式演示同源） | B·5.2 异步/模拟参考 |
| `testing/` | 自动化测试（`mybook.py` + `mybook.xlsm`、`test_mybook.py`、`test_sample.py`、`.gitlab-ci.yml`） | 场景 C 测试写法参考 |

## 8. xlwings-server-main：官方 xlwings Server 服务端工程

**一句话定位**：xlwings Server 官方服务端工程，为 Microsoft Excel 与 Google Sheets 提供 Python 支持（无需本地 Python 安装），可自托管于任意支持 Python/Docker 的平台；source-available 双许可（PolyForm Noncommercial License 1.0.0 非商业免费 / xlwings PRO EULA 商业需付费）。

**约束**：按本技能约定（SKILL.md 场景 D·7.8），本仓库仅作为源码在工作流 D 中研读（理解架构、启动编排、部署形态与二次开发参考），**不纳入工作流 A/B/C/D 执行实际部署运行**；需要部署能力时以仓库 `docs/` 与官方文档 https://server.xlwings.org 为准。

**目录**：`xlwings-server-main/xlwings-server-main/`

| 文件/目录 | 职责 |
| --- | --- |
| `run.py`、`Makefile`、`pyproject.toml`、`uv.lock` | 启动入口与工程构建/依赖编排（研读快速启动与开发环境搭建） |
| `xlwings_server/` | Server 核心包：`main.py`（FastAPI 应用）、`config.py`/`databases.py`/`dependencies.py`（配置与依赖）、`routers/`/`serializers/`/`models/`（REST 分层）、`auth/`（认证）、`custom_functions/`、`custom_scripts/`、`object_handles.py`、`templates.py`、`cli.py`、`wasm/` 等 |
| `tests/` | 服务端测试样例 |
| `docs/` | 完整文档源（quickstart、server_config、docker_compose、production、authentication、custom_functions、wasm 等 80 余篇） |
| `deployment/`、`nginx/`、`certs/`、`scripts/` | 部署编排样例、HTTPS 反向代理样例、证书与辅助脚本（研读部署形态） |
| `.devcontainer/`、`.github/`、`.vscode/` | 开发容器与 CI 配置 |
| `README.md`、`DEVELOPER_GUIDE.md`、`CLAUDE.md`、`LICENSE.md`、`SECURITY.md`、`TODO.md` | 工程说明与开发约定（研读前先读 README 与 DEVELOPER_GUIDE） |

**关键特性**：支持自托管于裸机、Linux VM、Docker Compose、Kubernetes 及 Azure Functions / AWS Lambda 等 Serverless；通过 Excel 自定义函数、Office JS 任务窗格、Google Sheets 集成使用；含认证、自定义函数/脚本、WASM、任务窗格等扩展体系。详细能力与限制见仓库 `docs/`。

**二次开发参考路径**：改认证/授权 → `xlwings_server/auth/` + `docs/authentication.md`；加自定义函数 → `xlwings_server/custom_functions/` + `docs/custom_functions.md`；接 Serverless → `xlwings_server/azure_functions_templates/` + `docs/azure_functions.md`；调整启动编排 → `run.py`/`Makefile` + `docs/docker_compose.md`。

---

## 9. python-for-excel-course-main：官方入门课程讲义（0-7 Part）

**一句话定位**：xlwings 官方 YouTube 入门视频课程（2018）配套 Jupyter notebook 讲义，按 `0 - Intro` 至 `7 - Part7` 共 8 个 Part 组织，系统覆盖 xlwings 基础到 UDF、单元测试与应用部署，用于场景 D·7.5 的系统入门学习与教学演示。

**注意**：课程发布于 2018 年，部分写法可能过时，但作为 xlwings 基础概念入门仍有效；执行时以本技能内置 xlwings-0.37.2 源码为准（见 SKILL.md 3.1 本地源码注入）。

**目录**：`python-for-excel-course-main/python-for-excel-course-main/`

| 目录 | 讲义内容 |
| --- | --- |
| `0 - Intro` | 课程说明、语言与区域设置（`Instructions.md`、`Language and Regional Settings.md`） |
| `1 - Part1` | 基础（`Tutorial 1 - The Basics.ipynb`，含 `table_objects.xlsx`、`timeseries.xlsx`） |
| `2 - Part2` | 转换器与选项（`Tutorial 2 - Converters and options.ipynb`） |
| `3 - Part3` | 加载项与 RunPython（`Tutorial 3 - Add-In & RunPython.ipynb`） |
| `4 - Part4` | UDF（`Module 4 - UDFs.ipynb`，含 `sql.xlsx`、`timeseries.xlsm`） |
| `5 - Part5` | 高级主题（`Tutorial 5 - Advanced xlwings topics.ipynb`） |
| `6 - Part6` | 单元测试（`Tutorial 6 - Unit Testing.ipynb` + `test_sample.py`） |
| `7 - Part7` | 应用开发与部署（`Tutorial 7 - App Development & Deployment.ipynb`） |
| `README.md` | 课程总览与 YouTube 播放列表链接 |

**使用方式**（见 SKILL.md 7.5）：系统入门按 0→7 顺序逐个打开并理解（含可执行单元格的 notebook 用 Jupyter 内核执行）；查漏补缺时由用户指明主题所属 Part，仅学习该部分及其前置依赖；教学演示可任选其中 notebook 作为演示素材。官方视频播放列表见仓库 README。

---

## 10. taxi-duckdb-main：Web 加载项 + DuckDB 一体化案例

**一句话定位**：xlwings 官方"一个 .xlsx 自包含完整数据应用"样板——工作簿内嵌 Office Web 加载项（xlwings Lite / Pyodide 0.27.5），任务窗格用 `@script` 触发 DuckDB 查询纽约出租车 Parquet 数据，并把统计表格与 matplotlib 图表回写 Excel；不依赖 PRO/Server，用于场景 D 的 Web 加载项（officejs / Lite）案例研读。

**注意**：`taxi_local.xlsx` 的工作表本身为空，全部"内容"在 `xl/webextensions/` 内嵌加载项中（main.py / requirements.txt / 运行参数）；DuckDB 数据读取依赖 Pyodide 运行环境与 `/taxi/` 虚拟路径，脱离 xlwings Lite 无法直接复现。

**目录**：`taxi-duckdb-main/taxi-duckdb-main/`

| 文件/目录 | 说明 |
| --- | --- |
| `LICENSE` | MIT 许可 |
| `taxi_local.xlsx` | 内嵌 Web 加载项的空白工作簿（全部应用逻辑在此） |
| `extracted/` | 本技能一步到位提取产物：`extract_webextension.py`（解包 + 解码脚本）、`full/`（12 个 zip 条目原样 + `main.py` 98 行 explore/report 两个 `@script`、`requirements.txt` 解码明文就地生成于 `full/xl/webextensions/`） |

**使用方式**（见 SKILL.md 7.1 与 13-scenario-d 4.5）：研读 `full/xl/webextensions/main.py` 的 `explore` / `report` 两个 `@script` 函数与依赖策略（可运行 `extract_webextension.py` 一步重新生成），借鉴"空 xlsx 壳 + 内嵌加载项"的自包含交付形态，以及"DuckDB 查询 → pandas 转换 → matplotlib 图表 → Excel 表格 / 图片回写"四层管线。
