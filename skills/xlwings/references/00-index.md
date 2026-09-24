# references 决策索引（先读这个）

本目录共 16 篇知识文档 + 本索引，**文档顺序 = 工作流顺序**（01→10 对应场景 C 6.5.1→6.5.13，11→13 对应场景 A/MCP/D，14→16 为 Web 三通道：Server / Lite / Office.js）。内容源自四个权威源：`xlwings/docs/`（官方 Python/加载项文档）、`../VBA-Docs/`（微软官方 VBA 文档）、`examples/`（xlwings 官方示例）、`MCP-Server/`（4 个社区 xlwings MCP Server 源码）。与 xlwings 无关的内容已全部去除。

## 文档地图（顺序 = 工作流）

| # | 文档 | 对应步骤 | 角色 |
|---|------|---------|------|
| 00 | 本索引 | 全流程 | 决策路由 + 三源导航 |
| 01 | `01-need-discovery.md` | 6.5.1 需求澄清与形态判定 | 问什么（必问问题引导）+ 怎么判 + 零基础沟通规范 |
| 02 | `02-encoding-com-prereqs.md` | 6.5.2 环境准备 | 编码规则 + COM/AccessVBOM/WPS 信任 + 预检命令 + PowerShell 专题 |
| 03 | `03-project-design-guidance.md` | 6.5.3 项目初始化与设计 | 列 Schema 决策法（列类型/列级策略/ListObject 映射） |
| 04 | `04-python-guidance.md` | 6.5.4 Python 代码构建 | 官方 docs 导航 + 示例导航 + 进阶主题（转换器/数据结构/线程/连接） |
| 05 | `05-vba-guidance.md` | 6.5.5 VBA 代码构建 | VBA 补充规则 + 反模式 + Excel 对象模型速查 + 单元格/区域引用 + 现有文件迭代 + API 查证流程 |
| 06 | `06-ribbon-guidance.md` | 6.5.6 Ribbon 构建 | Ribbon 版本/入口/回调签名速查 + 图标原则 |
| 07 | `07-udf-guidance.md` | 6.5.7 UDF | @xw.func 规范 + 导入测试 + 官方 udfs 导航 + 官方 UDF 案例详展（udf/fibonacci/itus） |
| 08 | `08-form-guidance.md` | 窗体域（6.5.1 判定 / 6.5.4 面板 / 6.5.5 UserForm） | UserForm / pywebview / tkinter 三形态 + 面板侧完整模板 |
| 09 | `09-testing-debugging-guidance.md` | 6.5.9 单元测试 + 6.5.12 集成验证 | TDD 理念 + COM 冒烟 + 真机验证 + 调试 |
| 10 | `10-deployment-delivery.md` | 6.5.10/11 构建打包与交付 | 部署与交付工作流（选型→打包→代码嵌入→配置表→门禁验证→交付清单）；主线自研 release_tool 便携运行时分发；附录 A = PRO 部署机制（code embed/release）对照研读（原 13-d 1.4 迁入）；附录 B 官方文档映射 |
| 11 | `11-scenario-a-python-automation.md` | 场景 A 主线 | 连接/导航/读写/格式化/图表图片形状表格 + API 参考导航 + 社区版案例详展（quickstart/自动化测试/质检/双端/PQ/Eikon） |
| 12 | `12-scenario-a-mcp-automation.md` | 场景 A 4.4（MCP 自动化）深度扩展 | 4 个社区 xlwings MCP Server 逐一详解 + 工具清单 + 安装配置 + 安全注意 + 故障排查 |
| 13 | `13-scenario-d-source-code.md` | 场景 D 深度扩展 | xlwings PRO 深度技术分析（许可证/四引擎/Reports/部署）、入门学习执行步骤、案例总览与 PRO 案例研读（社区版案例详展在 11 场景 A references；1.7 已指针化） |
| 14 | `14-xlwings-server-guidance.md` | 场景 B（Web/服务化）· xlwings Server | xlwings Server 独立参考手册：开发生命周期工作流（预研选型→初始化→扩展机制→自定义函数/脚本→认证安全→测试调试→部署运维升级），六块全覆盖（docs 61 篇教程 / xlwings_server 包 / deployment / nginx / scripts / tests） |
| 15 | `15-xlwings-lite-guidance.md` | 场景 B（Web/服务化）· xlwings Lite | xlwings Lite 独立参考手册：开发生命周期工作流（认知选型→环境准备→核心开发→测试调试→应用化分发→自托管部署→案例研读），开发节按官方教程层次编排（入门组：函数/脚本/Notebooks；进阶组：依赖/数据库/Web/异步/绘图；个人模块收尾），每节含"源码验证"块；覆盖官方文档镜像 xlwings-lite 语义引用 + Lite 相关源码研读（ext/sql.py、BookAsync、base_classes 异步方法族、ObjectHandle/WithScript、udfs_officejs Lite 分支：JsNull/streaming_callback/_lazy）+ taxi-duckdb 完整案例（原 13-d 1.6/4.5 与 1.8 Lite 适配分支均已迁入并指针化） |
| 16 | `16-xlwings-officejs.md` | 场景 B（Web/服务化）· Office.js 引擎 | Office.js 引擎深度技术分析与开发工作流：值转换层（`_xlofficejs.py`）+ UDF/脚本全链路（`udfs_officejs.py`）+ socket.io 会话 + 测试，官方教程对照（lite custom-functions/custom-scripts + whatsnew 演进），Server 与 Lite 共享的语义内核（原 13-d 1.2.5/1.8 独立而来） |

## 模板资产地图（templates/）

`templates/` 下有 **三个并列蓝图**（每个完整自包含：build 脚本 + Ribbon + VBA 模块 + install 脚本），通过 `scripts/instantiate_blueprint.ps1 -BlueprintName <名称>` 选择实例化。蓝图之间**无继承关系**，修改公共逻辑需同步三处（维护成本已知，后续可考虑模板继承机制）。

| 蓝图路径 | 名称 | 适用场景 | 关键资产 |
|---------|------|---------|---------|
| `internal-bases/excel-addin-core/` | excel-addin-core | 最小白标加载项骨架 | default_module.vba.tmpl + customUI14.xml.tmpl + install_xlstart.ps1.tmpl |
| `project-blueprints/listobject-workbench-addin/` | listobject-workbench-addin | ListObject 数据工作台 | bridge_example.py（xlwings 桥接示例） |
| `project-blueprints/userform-data-entry-addin/` | userform-data-entry-addin | UserForm 表单录入 | frmDataEntryWizard.form.json + .code.vba + userform_window_helper.vba |
| `shared/main.py` | — | 面板型加载项入口参考（非蓝图） | pywebview 面板启动骨架（完整模板见 08-form-guidance.md） |

**Python 代码模板**（`internal-bases/excel-addin-core/src/python/`）：

| 模板 | 用途 | 对应步骤 |
|------|------|---------|
| `udf_example.py.tmpl` | UDF 函数示例（@xw.func + Annotated 类型提示） | 6.5.7 |
| `runpython_bridge.py.tmpl` | RunPython 桥接入口（无参无返回值 + 单元格读写 + 自测入口） | 6.5.4 |
| `panel_entry.py.tmpl` | 面板入口（pywebview + uvicorn subprocess + 端口探测） | 6.5.4 |
| `tests/test_smoke.py.tmpl` | 冒烟测试模板（纯 Python 单元 + COM 集成 + xlam UDF + fixture） | 6.5.9 |

## 快速决策树

```
用户在哪个场景？
│
├─ 场景 A：自动化操作 Excel（Python 脚本 + MCP）→ 11（脚本）+ 12（MCP，4.4）+ SKILL.md 第 4 章
├─ 场景 B：Web / 服务化（Server / Lite / Office.js）→ 14 + 15 + 16 + SKILL.md 第 5 章
├─ 场景 C：开发 VBA 宏或 VBA 加载项（主干）→ 按 6.5 工作流：
│   ├─ 需求澄清 / 形态判定（6.5.1）     → 01 + SKILL.md 6.5.1 三维判定模型
│   ├─ 环境准备（6.5.2）               → 02（含环境预检快捷命令）
│   ├─ 项目初始化与设计（6.5.3）        → 03（列 Schema 决策法）
│   ├─ 写 Python 侧（6.5.4）           → templates/src/python/*.tmpl + 04 扩展导航
│   ├─ 写 VBA 侧（6.5.5）              → 05（补充规则/反模式/对象模型/API 查证）
│   ├─ 写 Ribbon（6.5.6）              → 06（版本/入口/回调签名）
│   ├─ 写 UDF（6.5.7）                 → 07（规范/导入测试）+ udf_example.py.tmpl
│   ├─ 做窗体（UserForm/pywebview/tkinter） → 08（三形态选型 + 面板完整模板）
│   ├─ 单元测试（6.5.9）+ 集成验证（6.5.12）     → 09（TDD/冒烟/真机/调试）+ test_smoke.py.tmpl
│   └─ 构建 / 打包 / 交付（6.5.10/11） → 10（分发/便携运行时/配置表/交付清单）
└─ 场景 D：源码学习/二次开发/产物质检 → 13（PRO 分析/案例总览）+ 14（Server）+ 15（Lite）+ 16（Office.js 引擎）+ SKILL.md 第 7 章
```

## 三源导航（官方权威文档路径）

| 权威源 | 用途 | 关键文件 |
|--------|------|---------|
| `xlwings/docs/` | xlwings 官方文档（Python 侧权威） | `udfs.md`（UDF 全参数）、`converters.md`（类型转换）、`datastructures.md`（数据结构）、`customaddin.md`（加载项/Ribbon）、`addin.md`（配置表）、`command_line.md`（CLI）、`connect_to_workbook.md`（连接）、`syntax_overview.md`（语法）、`threading_and_multiprocessing.md`（线程）、`debugging.md`（调试）、`troubleshooting.md`、`deployment.md`（部署）、`quickstart.md`、`vba.md`（VBA 侧）、`installation.md`（安装）、`matplotlib.md`、`jupyternotebooks.md`、`onedrive_sharepoint.md`、`other_office_apps.md`、`missing_features.md` |
| `../VBA-Docs/` | 微软官方 VBA 文档（VBA 侧权威） | `api/Excel.*.md`（对象模型逐成员）、`excel/Concepts/Cells-and-Ranges/`（单元格引用 9 篇+Formula2）、`excel/Concepts/Controls-DialogBoxes-Forms/`（UserForm+ActiveX）、`excel/Concepts/Events-WorksheetFunctions-Shapes/`（事件+WorksheetFunction）、`excel/Concepts/Excel-Performance/`（性能优化）、`excel/Concepts/Workbooks-and-Worksheets/`（工作簿操作）、`Language/Reference/`（VBA 语言参考） |
| `examples/` | xlwings 官方/社区示例 | `xlwings-demo-master/`（13 场景）、`Excel_udf_itus-main/`（UDF 教程）、`excel-automated-testing-master/`（自动化测试）、`python-for-excel-course-main/`（课程）、`xlwings-server/`（Server，已迁至技能根目录）、`xl-pq-handler-master/`、`simulation-demo-master/`、`static-excel-test-master/`、`cross-check-reports-main/`、`xlwings-factsheet-demo-main/`、`xlwings-eikon-master/` |

## 路由原则

- **先读主干，再看专题**：VBA 主干总纲在 6.5.5，专题细节再 05；Python 先 SKILL.md 6.5.4 模板；UDF 先 07；窗体先 08 选型表。
- **官方文档是权威**：凡与官方语义冲突处，以 `xlwings/docs/` 与 `../VBA-Docs/` 为准；本目录文档负责"要点提炼 + 导航 + 实战经验"，不复制官方全文。
- **示例按需消费**：需要完整场景时优先 `examples/` 对应项目，而不是从零拼片段。
- **文档不重复原则**：SKILL.md 正文承载"规则与铁律"，references 承载"要点与导航"；若发现重复，以 SKILL.md 为准并在 00 记录差异。

## 引用方清单（重构后逐层核查用）

> 每次 references 重构后，按此清单逐层核查：① SKILL.md ② docs/ ③ scripts/ ④ templates/ ⑤ references 内部互引。任何一篇被删除或改名，必须同步更新所有引用方。

| 文档 | 引用方（按层） |
|------|---------------|
| `00-index.md` | SKILL.md |
| `01-need-discovery.md` | SKILL.md、references/00-index.md |
| `02-encoding-com-prereqs.md` | SKILL.md、references/00-index.md |
| `03-project-design-guidance.md` | SKILL.md、references/00-index.md |
| `04-python-guidance.md` | SKILL.md、references/00-index.md、references/11-scenario-a（案例路由）、references/13-scenario-d（案例路由） |
| `05-vba-guidance.md` | SKILL.md、references/00-index.md、templates/internal-bases/excel-addin-core/src/modules/default_module.vba.tmpl |
| `06-ribbon-guidance.md` | SKILL.md、references/00-index.md |
| `07-udf-guidance.md` | SKILL.md、references/00-index.md、references/13-scenario-d（案例路由）、docs/troubleshooting.md |
| `08-form-guidance.md` | SKILL.md、references/00-index.md、templates/project-blueprints/userform-data-entry-addin/src/modules/default_module.vba.tmpl、templates/shared/main.py |
| `09-testing-debugging-guidance.md` | SKILL.md、references/00-index.md、docs/troubleshooting-python.md、scripts/verify_addin_registered.ps1 |
| `10-deployment-delivery.md` | SKILL.md、references/00-index.md |
| `11-scenario-a-python-automation.md` | SKILL.md、references/00-index.md、references/04-python-guidance（案例路由）、references/13-scenario-d（案例路由） |
| `12-scenario-a-mcp-automation.md` | SKILL.md、references/00-index.md |
| `13-scenario-d-source-code.md` | SKILL.md、references/00-index.md、references/04-python-guidance（案例路由）、references/07-udf-guidance（案例路由）、references/11-scenario-a（案例路由）、references/15-xlwings-lite-guidance（Lite 内容已迁出：1.6/4.5 指针 + 1.8 Lite 适配分支指针）、references/16-xlwings-officejs（Office.js 引擎内容已迁出：1.2.5/1.8 指针） |
| `14-xlwings-server-guidance.md` | SKILL.md（5 章场景 B、7.4/7.5/7.6）、references/00-index.md、references/13-scenario-d（1.7 指针）、references/15-xlwings-lite-guidance（导航） |
| `15-xlwings-lite-guidance.md` | SKILL.md（5 章场景 B、7.1 taxi 案例、7.4 PRO 研读、7.6 总览）、references/00-index.md、references/13-scenario-d（1.6/4.5 指针）、xlwings-lite（31 篇官方文档镜像） |
| `16-xlwings-officejs.md` | SKILL.md（5 章场景 B、7.4/7.5）、references/00-index.md、references/13-scenario-d（1.2.5/1.8 指针）、references/15-xlwings-lite-guidance（Lite 适配分支导航） |

**核查命令**（重构后执行）：
```powershell
# 断链扫描：全库 references/ 引用目标存在性
# 残留扫描：旧文件名/已删资产在全库（含 .tmpl/.ps1/.py）的命中
# 完整性：13 篇每篇至少被 SKILL.md 引用一次
```


