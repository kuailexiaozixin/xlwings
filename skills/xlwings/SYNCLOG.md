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
