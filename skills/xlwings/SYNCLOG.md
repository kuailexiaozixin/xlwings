# xlwings 技能上游同步日志

> 2026-09-14 初始化登记：上游跟踪扩展到 17 路（xlwings 核心源码包 + examples/ 与 MCP-Server/ 下全部开源仓库），见 manifest.json。

| 日期 | 来源 | 方式 | 新 commit | 规模 |
|---|---|---|---|---|
| 2026-09-14 | `xlwings` | tarball | `509a336` | 402 个文件 |
| 2026-09-14 | `xlwings-server` | tarball | `76f5f1a` | 396 个文件 |
| 2026-09-15 | `xlwings` | tarball | `46301d0`（0.37.3） | 416 个文件 |
| 2026-09-15 | `xlwings-server` | tarball（main） | `5f3fccb` | 399 个文件 |
| 2026-09-15 | 目录迁移 | — | — | `examples/xlwings-server-main`→`xlwings-server/`、`examples/xlwings-lite`→`xlwings-lite/`（技能根目录） |
| 2026-09-17 | `xlwings-server` | tarball | `9a7807b` | 399 个文件 |
| 2026-09-22 | `xlwings` | tarball | `bbde5a7` | 416 个文件 |
| 2026-09-22 | `xlwings-server` | tarball | `00061e0` | 402 个文件 |
| 2026-09-23 | `xlwings` | tarball | `3dc83bb` | 424 个文件 |
| 2026-09-23 | `xlwings-server` | tarball | `8cbe5e1` | 403 个文件 |
| 2026-09-24 | `xlwings` | zip（ghfast 代理） | `f3862038` | 428 个文件 |
| 2026-09-24 | `xlwings-server` | zip（ghfast 代理） | `3b4450ba` | 405 个文件 |

> 2026-09-23 说明：`xlwings-0.37.4/` 改名为 `xlwings/`，`xlwings-server/` 名称不变，两路的 `local_dir` 从此不含版本号，上游出新版不再需要改自研文档里的路径。`xlwings` 一路的 ref 由 `0.37.4` 改为 `main`（当时最新正式版仍是 0.37.4，main 领先其 46 个提交，新增 Sheets.move、条件格式、数据验证、AutoFilter 等 API）。随本次替换刷新的事实：`pro/_xlremote.py` 4173→4814 行、`pro/_xlcalamine.py` 536→561 行、`pro/udfs_officejs.py` 1255→1254 行、`src/lib.rs` 9.4KB→9.6KB、main.py 的 BookAsync 锚点 L1389-1405→L1424-1441；上游删除了 `docs/api/index.md`，SKILL.md 第 120 行改指 `xlwings/docs/api/`（40 篇按类分页）。`cli.py` 的 13 处行号锚点全部未动。

> 2026-09-24 说明：本次两路均经 ghfast.top 代理下载 main 分支 zip（直连 GitHub 不可达），故方式记 `zip（ghfast 代理）`，commit 由 `git ls-remote` 经同一代理取得。`xlwings` 一路上游新增 Chart 轴/序列 API：`docs/api/` 新增 `chart_axis.md`、`chart_series.md`、`chart_series_collection.md` 三篇（40→43 篇，SKILL.md 第 120 行已同步改为 43 篇），`pro/_xlremote.py` 4814→5266 行（references/13 第 100 行、references/14 第 141 行已同步）；`xlwings-server` 一路上游新增 `range-colors.js` 自定义脚本与对应集成测试。`pro/_xlcalamine.py`（561 行）、`pro/udfs_officejs.py`（1254 行）、`_xlofficejs.py`（160 行）、`udfs.py`（838 行）、`com_server.py`（385 行）行数未变，`main.py` BookAsync 锚点 L1424-1441 仍有效，`cli.py` 行号锚点未动。
| 2026-09-27 | `xlwings` | codeload | `32c6453` | 447 个文件 |
| 2026-09-27 | `xlwings-server` | codeload | `4fcadb9` | 422 个文件 |
| 2026-10-05 | `xlwings` | codeload | `d5abd28` | 448 个文件 |
| 2026-10-05 | `xlwings-server` | codeload | `70c72bf` | 425 个文件 |

> 2026-10-05 说明：xlwings 上游 6 个提交（pivot 远程范围读取、远程透视表、mac/win 新增 comments API）：`docs/api/` 新增 `comment.md`、`comments.md`、`comment_reply.md`、`notes.md`、`table_column.md`、`table_columns.md`、`table_row.md`、`table_rows.md` 八篇（43→51 篇，SKILL.md 第 153 行枚举已同步）；`pro/_xlremote.py` 5856→5885 行（references/13 第 100 行、references/14 第 141 行已同步）；`main.py` BookAsync 锚点 L1449 未动；`pro/_xlcalamine.py`（561 行）、`pro/_xlofficejs.py`（160 行）、`pro/udfs_officejs.py`（1259 行）、`com_server.py`（385 行）行数未变。xlwings-server 上游 6 个提交（pivot 自定义脚本 pivot-identity/pivot-range-read/pivot-action-callbacks、htmx/socket.io/alpinejs vendor 升级），无自研锚点引用受影响。

## 2026-10-05 全量审计与内容补齐

> 审计方法：脚本化提取自研文档（SKILL.md + 16 篇 references）全部可验证锚点（行数声明 13 处、路径 884 处、行号锚点 4 处）与上游快照逐一比对；覆盖性扫描 xlwings/docs 全部 131 篇 md（43+8 api 篇除外）、pro/ 模块、main.py 全部类、xlwings-lite 31 篇、xlwings-server docs 61 篇。
>
> 修复清单：
> - SKILL.md L156/L166 + references/11 第三节：批注表述重写——桌面/remote 高层 `add_comment`、`book.comments`/`sheet.comments`、`CommentReply`（桌面 COM 不支持回复）；`.note` 仍限旧式备注（创建须 COM）
> - references/13：remote 引擎补透视表范围读取能力（report/data_body）；测试规模 40→52 个/380KB→540KB/5→4 子目录
> - references/14：新增 5.2b 节（pivot-identity/pivot-range-read/pivot-action-callbacks 三脚本与 wire 协议）；test_router_xlwings.py 586→650 行
> - references/12：hyunjae 模块表修正为快照实况（src/xlwings_mcp 7 文件；xlwings_impl 子包未随 OSS 发布，旧"20 个模块"描述失实）
> - references/10：reinject 脚本路径 `dist_release/`→`scripts/tools/`
> - SKILL.md 841：补 `xlwings/docs/opensource_licenses.md` 引用
>
> 验证：修复后复跑审计，上游锚点真阳性 0；修改前备份在 `D:\WPS灵犀过程文件\20261005-09-47-21-498\backup2\`。

> 2026-10-05 复审（批注条目二轮核验）：逐行核对 `_xlwindows.py`/`_xlmac.py`/`pro/_xlremote.py` 后修正首轮两处错误表述——① 旧式备注创建**不再必须 COM**：上游同步新增 `Range.add_note(text)` 高层方法与 `Notes` 集合（COM `AddComment` 仅为 Windows 底层实现）；② 线程回复 `replies`/`add_reply` **Windows 桌面可用**（COM `CommentThreaded.Replies`/`AddReply`），真正不可用的是 macOS（AppleScript 不支持线程批注整体）与桌面端的 `resolved` 状态（COM 不暴露）；remote 引擎同步读均不支持、异步 `get_*` 仅 Lite（emscripten）。SKILL.md L156/L166 与 references/11 第三节已按三引擎差异表重写，20 项源码断言全部通过。

> 2026-10-05 三审（存疑数量声明全量复核，KB 口径统一为 round(bytes/1024)）：
> - references/13 五处 KB 修正：constants 146→143、_xlwindows 89→137、_xlmac 102→152、mac_dict 258→255、test_range 32→39；test_object_handles 22KB 经独立复测（22549B）**正确未改**（二轮清单中"22→24"系误报）
> - references/14 四处"tests/ 41 项"→ 实况 59 个文件（Python 测试 21 + JS 测试 28 + 配置/数据 10；L36 表格两处、L530、L532 标题）；"41"在新旧快照任何口径下均不成立
> - SKILL.md L191 案例数 11→12（references/11 附录表格实 12 行）；L350 预检函数 6→4（references/02 预检脚本实际调用 4 个库函数：Get-OfficeAutomationHostInfo/Get-ExcelVersion/Test-AccessVBOMEnabledForVersions/Test-WpsProjectTrustEnabled）
> - references/00-index L111 审计注释"13 篇"→"17 篇"（references/ 共 17 篇 md 且已验证每篇至少被 SKILL.md 引用一次）
> - references/16 L232 `custom_functions_code.js`→`custom-functions-code.js`（上游真实模板文件名，xlwings_server/static/js/custom-functions/，由 routers/xlwings.py:72 custom_functions_code 路由读取）
> - 核实后未改：references/12 server.py 35KB（35772B，"36"系 du 块分配口径）；SKILL.md L37"后 15 个跟踪默认分支"（manifest 17 路 = xlwings + xlwings-server + 15）；SKILL.md 门禁"12 个"口径（A–L 12 门禁编号，C/D 共用 verify_addin_registered.ps1——曾误改 10 后回滚）；docs/troubleshooting-python.md 系技能本地 docs/；xlwings.bas 为上游 build_excel_files.py 运行时生成物（L26/L164 转述属实）；baseline/inventory.json 为导入工具用户项目运行时产物
> - 审计工具升级 audit_anchors.py v4（技能本地解析层：references 互引/scripts/gates/release_tool/tools/docs/baseline/VBA-Docs + 微软文档与示意名白名单）；终态复跑：行数声明 13 处、路径 765 处、行号锚点 4 处，**问题 0 处**；check_kb.py 17/17 OK
> - 备份：backup2/ 新增 round5-preKB（refs12/13/14）、round5-preCounts（SKILL.md、00-index.md）、16 号 round5-pre/post
| 2026-10-15 | `xlwings` | codeload | `882fb15` | 452 个文件 |
| 2026-10-15 | `xlwings-server` | codeload | `617541b` | 429 个文件 |

> 2026-10-15 说明：两路均经 codeload 直连（本机直连 GitHub 已恢复，代理 7897 不再需要）。xlwings 10 天窗口内 `pro/_xlremote.py` 5885→5986 行（+101）、`pro/udfs_officejs.py` 1259→1373 行（+114）、`main.py` 320KB→337KB（BookAsync 锚点 L1449→L1465、`custom_functions_call` L524→L605，references/13:100/158/251/287、14:141、15:161/333/476/753、16:3/66 已同步）；`pro/_xlcalamine.py`（561）、`pro/_xlofficejs.py`（160）、`com_server.py`（385）、`udfs.py`（838）行数未变，`docs/api/` 仍 51 篇。xlwings-server `tests/test_router_xlwings.py`（650 行）与 `static/js/custom-scripts/` 21 个脚本清单未变。track_upstream.py 复跑线② 13 条全部 OK。第三方 `Excel_MCP_Server`（vohoailinh90）检测到漂移 `c03ef06→e06b66e`，不在本次两路同步范围，未动。
| 2026-10-20 | `xlwings` | ghfast | `f73213c` | 454 个文件 |
| 2026-10-20 | `xlwings-server` | codeload | `26be458` | 424 个文件 |

> 2026-10-20 锚点复核（线②）：上游 f73213c / 26be458 重拉后 3 条 STALE 已修——xlremote-lines 5986→6117（references/13:100、14:141）；main-size 337KB→341KB（references/13:287）；bookasync-anchor L1465→L1470（references/15:161/476/753）。台账 upstream-facts.json 同步：updated 2026-10-20、snapshot f73213c/26be458、三条 value 修正。线②复跑 14 条全绿，线③ __version__=="0.0.0" 符合预期（zip 快照特征）。第三方 Excel_MCP_Server 漂移 ee31ecd 不在本次两路范围，未动。
| 2026-10-24 | `xlwings` | requests/codeload | `5b38a61` | 454 个文件 |
| 2026-10-24 | `xlwings-server` | requests/codeload | `627d244` | 425 个文件 |

> 2026-10-24 说明：本会话沙箱内 `curl.exe`/`git` 经系统 schannel 建立 TLS 时被拒（`SEC_E_NO_CREDENTIALS`，直连与 Clash 代理 7897 均失败），`sync_upstream.py` 的 `latest_sha`（gh/git ls-remote）与 `curl` 下载两环均不可用。改用 Python requests（certifi CA + 显式 `127.0.0.1:7897` 代理）下载 codeload tar.gz、经 `commits/main.atom` 取 head（绕开 GitHub API 匿名限流 403），并复用 `sync_upstream.replace_dir`/`count_files` 的原子替换 + 文件数校验 + 回滚逻辑，只替换两路 local_dir，不触碰 SKILL.md/references；manifest 已写回。
>
> 锚点复核（线②③）：线③ `__version__=="0.0.0"` 符合预期。线② 14 条锚点中 1 条 STALE——`xlwings/pro/_xlremote.py` 6117→6137 行（区间内 Lite 透视表/Range.move/get_expanded_range/autofit 等提交），已修 references/13:100、references/14:141 及台账 `xlremote-lines`。其余（`_xlcalamine.py`561、`_xlofficejs.py`160、`udfs_officejs.py`1373、`udfs.py`838、`com_server.py`385、main.py 341KB、docs/api 51 篇、BookAsync L1470、cfc L605、server cli 64KB）均未变；台账 snapshot 更新为 5b38a61/627d244、updated 2026-10-24。复跑线②③全绿。
>
> 存量缺陷登记（非本次引入，旧快照 26be458 即已悬空，raw 接口对 26be458 返回 404 核实，待用户决定是否清理）：references/14 仍引用上游已移除的 `docs/repo_setup.md`（L31/158/732，上游并入 quickstart.md）、`docs/upgrade.md`（L63/674/740，现仅存 migration.md）、`docs/dependencies.md`（L740，快照 docs 无此页），及 `python run.py deps compile`（L690，run.py 现存子命令仅 init/wasm）。这些均不在 `upstream-facts.json` 台账覆盖范围（线②只核行数/KB/行号/文件数/存在性，不核正文 docs 链接），故 track_upstream 未报。

> 2026-10-24 内容改写（更新内容·新增内容·移除过时内容）：承接同日快照同步，对自研文档做语义层审计改写（区别于仅数值锚点）。
>
> A. 移除/改指过时内容——根因：xlwings-server 自 v1.0 起改为标准 Python 包分发，废除"仓库即项目"工作流（`docs/migration.md` 实证：*"instead of requiring you to work directly in the xlwings-server Git repository"*）；本次区间 commit #445/#446 进一步删 `docs/repo_setup.md`、`docs/upgrade.md` 与 `run.py deps` 子命令（raw 对旧快照 26be458 即 404，属历史存量，台账线②不核正文链接故未报）。逐项去向：① `repo_setup.md`（ref14 L31/158/732）→ 拆入 `quickstart.md`「5. Set up Version Control」+ `local_development.md`，已改指；② `upgrade.md`（L63/674/740，含补漏的 L37 阶段表）→ 去向 `migration.md`，已删引用；③ `dependencies.md`（L740）→ 快照 docs 无此页（pre-1.0 遗留），已删；④ `python run.py deps compile`（L690）→ 现 run.py 子命令仅 init/wasm，升级改 `uv sync`，2.1 与 8.7 两节按包形态重写（去 fork/merge/merge.ours 等仓库形态流程）。全局复核：references 全部 docs/*.md 悬空引用 0。
>
> B. 新增内容——本次上游新 API，三引擎支持经源码实证（main.py 定义 + 五引擎实现 grep）：`Range.move(dest)`(0.37.7,main.py L3512) 与 `Sheet.move(before/after)`(L1846) 为**同步**，桌面 COM(Win/mac)+Lite remote 支持、Office.js/Calamine 不支持 → 补入 ref13 对象模型 L287 + 区域操作节新 bullet；`Range.get_expanded_range()`(L3341) 与 `Chart.get_title()`/`ChartAxis.get_title()`(0.37.8,L6596/7048) 为**异步仅 Lite**（与既有透视表/PivotTable.get_range 同类，docstring "Requires xlwings Lite"）→ 补入 ref15 §base_classes get_\* 族列举 + ref13 L100 remote 能力语义延续。上游 docs/api 未为此新增专页（仍 51 篇，api-doc-count 锚点不变）。
>
> C. 台账 upstream-facts.json——新增 3 条 symbol_line 锚点（range-move-line L3512 / get-expanded-range-line L3341 / chart-get-title-line L6596），纳入线②跟踪，锚点总数 14→17。版本基准 ref14 L63 修正：tag 1.14.0 → 跟踪 main、对应 changelog 1.20.0。
>
> 验证：线② 17 条全 OK、线③ PASS；悬空引用 0。改写前二次备份 `D:\WPS灵犀过程文件\20261024-08-46-44-248\backup2-content\`（含 6137 修正态）。
