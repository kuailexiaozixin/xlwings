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
