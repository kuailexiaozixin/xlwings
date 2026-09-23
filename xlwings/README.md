# xlwings（占位）

本目录**未随本仓库分发**，仅作占位说明。

## 内容
xlwings 官方源码包的完整本地副本：Python 包源码（`xlwings/`，含 PRO 模块）、`docs/`、`examples/`、Rust 组件（`xlwingsdll`）、构建与测试配置。

目录名不含版本号：技能内所有文档都按 `xlwings/...` 引用快照，上游出新版时整体替换目录内容，读者的引用路径不变。本仓库 `manifest.json` 的 `xlwings` 条目记录该快照跟踪的上游分支（`ref`）与提交（`pinned_sha`），`SYNCLOG.md` 记录每次替换的日期与文件数。

## 官方来源
- 官方仓库：<https://github.com/xlwings/xlwings>
- 官网：<https://www.xlwings.org/>
- 文档：<https://docs.xlwings.org/>

## 在技能中的用途
技能源码研读对象（xlwings 包核心架构 / PRO 机制 / xlwingsdll / reader 与 calamine 数据通道）；构建应用系统时借鉴其工程实践。
