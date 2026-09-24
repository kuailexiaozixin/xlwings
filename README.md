# xlwings（xlwings 开发技能）

本仓库为 xlwings 开发技能的开源发布：使用 Python（xlwings）开发 Excel 应用的全生命周期技能——场景 A（Python 自动化操作 Excel）、场景 C（UDF 自定义函数）、场景 D（源码研读与工程借鉴），含构建、测试、部署与交付门禁。

技能按 Agent Skills 规范以虚拟树结构发布：`skills/xlwings/` 为一个完整技能。

## 目录结构

- `skills/xlwings/SKILL.md` — 技能主文档（工作流）
- `skills/xlwings/references/` — 场景 A/C/D 参考文档（案例路由、源码研读、部署交付）
- `skills/xlwings/scripts/` — 构建/测试/门禁脚本
- `skills/xlwings/docs/` — 补充文档
- `skills/xlwings/templates/` — 工程模板
- `skills/xlwings/dist/` — 构建产物
- `skills/xlwings/manifest.json` / `SYNCLOG.md` — 第三方上游快照的跟踪清单（`ref` + `pinned_sha`）与同步历史，配合 `scripts/sync_upstream.py` 使用

## 第三方资源（占位说明）

以下目录包含第三方源码包的本地副本，**未随本仓库分发**，以 `README.md` 占位并链接官方来源：

| 目录 | 官方来源 |
|---|---|
| `skills/xlwings/examples/` | 11 个官方/第三方示例仓库（见 `examples/ReadMe.md`） |
| `skills/xlwings/MCP-Server/` | 4 个 xlwings MCP 服务器实现（见 `MCP-Server/README.md`） |
| `skills/xlwings/xlwings/` | <https://github.com/xlwings/xlwings> · <https://www.xlwings.org/> |
| `skills/xlwings/xlwings-server/` | <https://github.com/xlwings/xlwings-server> · <https://server.xlwings.org/> |
| `skills/xlwings/xlwings-lite/` | <https://github.com/xlwings/xlwings-lite> · <https://docs.xlwings.org/en/latest/lite.html> |

## 许可

本仓库内容为技能文档与脚本；引用的第三方仓库/源码版权归其各自所有者，请遵循其各自许可证。
