# 场景 D 附录：xlwings Server 参考手册（开发生命周期工作流）

> 本手册由 `references/13-scenario-d-source-code.md` 1.7 节独立并大幅扩展而来：以开发生命周期工作流重排，深度研读官方仓库 `xlwings-server/` 的六块内容——`docs/`（61 篇官方教程）、`xlwings_server/`（服务端 Python 包）、`deployment/`、`nginx/`、`scripts/`、`tests/`。单点原则：xlwings Server 的一切细节只在本手册展开，13-d 1.7 只保留指针，本手册与 13-d 之间为单向导航。

## 目录

- [0 本手册的定位与使用方式](#0-本手册的定位与使用方式)
- [阶段一 预研与选型](#阶段一-预研与选型)
- [阶段二 初始化与本地开发环境](#阶段二-初始化与本地开发环境)
- [阶段三 应用装配与扩展机制（服务骨架）](#阶段三-应用装配与扩展机制服务骨架)
- [阶段四 核心能力开发——自定义函数](#阶段四-核心能力开发自定义函数)
- [阶段五 核心能力开发——自定义脚本与任务窗格](#阶段五-核心能力开发自定义脚本与任务窗格)
- [阶段六 认证、授权与安全](#阶段六-认证授权与安全)
- [阶段七 测试、调试与性能](#阶段七-测试调试与性能)
- [阶段八 部署、运维与升级](#阶段八-部署运维与升级)
- [附录 A 源码级工程借鉴](#附录-a-源码级工程借鉴)
- [附录 B 官方文档族地图（61 篇 → 阶段）](#附录-b-官方文档族地图61-篇-阶段)
- [附录 C 与 13-d 的分工](#附录-c-与-13-d-的分工)

## 0 本手册的定位与使用方式

**本手册回答什么**：xlwings Server 是什么、何时用、如何初始化、如何写自定义函数与脚本、如何做认证安全、如何测试调试、如何部署升级——按开发一个 Server 服务的真实顺序编排。

**加载项视角**：本手册的 8 阶段即 **Office.js Excel 加载项开发的完整工作流**——阶段一二选型与初始化（加载项骨架与环境）、阶段三四五开发（函数 / 脚本 / 任务窗格 / manifest）、阶段六安全、阶段七八测试与发布。若你的目标是“开发一个 Excel 加载项”，本手册就是端到端路线图；若只关心服务端运维，可跳过阶段四五的 UI 细节。

**工作流总览（8 阶段）**

| 阶段 | 目标 | 关键产出 | 对应仓库位置 | 验收要点 |
|------|------|---------|--------------|------|
| 一 预研与选型 | 判断是否该用 Server、选哪种集成方式 | 选型决策 | `docs/integrations.md`、`docs/index_*.md` | 选型决策表完整（方式+理由+排除项）；Server vs Wasm 取舍有结论 |
| 二 初始化与本地开发环境 | 跑起可热重载的本地 HTTPS 开发服务器 | 仓库 + .env + 本地 manifest | `docs/repo_setup.md`、`docs/quickstart.md`、`docs/server_config.md`、`docs/dev_certifi | 仓库可提交；HTTPS 开发服务器可启动；manifest 可 sideload；配置分工表明确 |cates.md`、`docs/dev_docker.md`、`docs/devcontainers.md`、`docs/github_codespaces.md`、`run.py` |
| 三 应用装配与扩展机制（服务骨架） | 理解服务骨架与用户扩展三通道，挂自己的路由/lifespan/静态资源 | 骨架认知 + 扩展文件 | `xlwings_server/main.py`、`config.py`、`cli.py`、`hotreload.py`、`templates.py`、 | 骨架装配图正确；自定义路由/lifespan/静态覆盖真实生效 |`utils.py`、`databases.py` |
| 四 核心能力开发——自定义函数 | 把 Python 函数暴露为 Excel 公式（转换/命名空间/流式/句柄/后续脚本） | 函数集 + 元数据 | `docs/custom_functions.md`、`xlwings_server/routers/xlwings.py`、`serializ | 函数经真机公式调用返回预期值；流式/句柄（若有）验证通过；元数据生成正确 |ers/`、`object_handles.py`、`custom_functions/` |
| 五 核心能力开发——自定义脚本与任务窗格 | 写整簿操作脚本 + 构建任务窗格 UI + 产出并安装 manifest | 脚本 + 任务窗格 + manifest | `docs/custom_scripts.md`、`docs/taskpane_*.md`、`docs/jinja.md`、 | 脚本三通道至少一条真机跑通；任务窗格可渲染；manifest 安装验证记录 |`docs/htmx.md`、`docs/alpinejs.md`、`docs/bootstrap.md`、`docs/manifest.md`、`docs/install_officejs_addin.md`、`custom_scripts/`、`templates/` |
| 六 认证、授权与安全 | 让函数/脚本受保护，分级授权，锁定任务窗格 | 认证配置 + RBAC + 安全头 | `docs/authentication.md`、`docs/authorization.md`、`docs/auth_entraid.md`、`docs/auth_provider | 未登录返回 Access Denied；角色分级生效；安全头生产生效 |s.md`、`docs/auth_taskpane.md`、`auth/`、`dependencies.py`、`security_headers.py` |
| 七 测试、调试与性能 | 建立质量门禁，掌握调试手段，应用性能配方 | 三层测试 + 验证清单 | `tests/`（41 项）、`docs/debugging.md`、`docs/performance.md`、`docs/troubleshooting.md` | 三层测试全绿（对齐官方 41 项结构）；调试记录 + 性能基线 |
| 八 部署、运维与升级 | 上线并持续演进 | 部署配置 + 生产环境 | `deployment/`、`nginx/`、`scripts/`、`docs/index_hosting.md` 及各托管族、`docs/migration.md`、`docs/upgrade.md`、`docs/was | 部署配置就绪；生产清单核查通过；迁移/升级执行记录 |m_*.md` |

**加载项三件套 ↔ 阶段速查**（Office.js 加载项 = 壳 + 核）：

| 加载项构件 | 说明 | 对应阶段 |
|-----------|------|---------|
| manifest.xml | 声明加载项身份、UI 命令、自定义函数端点、权限 | 阶段二（生成）→ 阶段五 5.5（完善 / 安装） |
| 任务窗格 | 前端 UI（htmx/Alpine/Bootstrap），经 `taskpane.py` 注入认证上下文 | 阶段五 5.3 / 5.4 |
| 自定义函数 + 脚本 | 业务核心：Python 函数 / 脚本 → Office.js 端点 | 阶段四 + 阶段五 5.1 / 5.2 |

**阅读方式**：每个阶段 = 任务（做什么）→ 要求（做到什么程度）→ 步骤（怎么做）→ 输出物（交付什么）→ 验收（如何判定达标）；步骤内以无序列表给出要点与源码/文档出处；进阶读者可沿附录 B 的文档族地图回到官方教程原文。

> **执行原则**：开工前先产出 `task_plan.md`（逐条列出：工作步骤、该步任务、手段、产出物、可量化的验收标准）；**逐步推进、逐步验收**——只有完成当前阶段全部目标并留证（产物/日志/门禁输出）才能进入下一阶段，任何一步未达标不得声称完成。

**与场景 C 的分工**：两条不同的加载项开发路径。

| 维度 | 场景 C（VBA 宏 / VBA 加载项开发） | 本手册（Server / Office.js） |
|------|--------------------------------|------------------------------|
| 路径 | 桌面本地加载项 | Web / Office.js 加载项 |
| 载体 | .xlam / .xlsm + XLSTART / 便携运行时 | manifest.xml + 任务窗格 |
| Python 位置 | 本机或随包（便携运行时免装） | 自己的服务器 |
| 安装 | XLSTART / 便携分发 | sideload / M365 管理中心 / AppSource |
| 认证 | 无或简单 | SSO / RBAC |

选型判据：目标机无 Python 且多人 / 浏览器 / 在线协作 → Server 路线（本手册）；单机离线、VBA 生态、需要桌面 UDF/Ribbon 体验 → 场景 C。二者都涉及 UDF（本地 vs Server 自定义函数）、Ribbon（本地 vs manifest 命令）、安装（XLSTART vs 加载项分发），但实现通道完全不同，互不替代。

**版本基准**：本手册对应上游 tag `1.14.0`（`docs/changelog.md` 最新条目 2026-09-09；技能 17 路上游 manifest 锁定该 tag）。`xlwings_server/__init__.py` 从 `_version` 读取版本号，`docs/changelog.md`、`docs/upgrade.md` 是版本事实的唯一权威。

**许可提示**：source-available 双许可（PolyForm Noncommercial / PRO EULA，见 `LICENSE.md`）——仅供源码研读借鉴设计，不部署、不抄码。
**加载项三件套就绪检查（发布前门禁）**：

| 检查 | 判据 |
|------|------|
| manifest 完整 | 三元素齐全：`<Hosts>`、`<FunctionFile>` / `<Action>`、`<VersionOverrides>`；`Id` 唯一 GUID |
| 端点可达 | 本地 HTTPS 端点可用；函数 / 脚本经 sideload 真机调用返回预期值 |
| 任务窗格可渲染 | 注入认证上下文后正常显示；无控制台错误 |
| 认证闭环 | 函数 / 脚本受保护；未登录返回 Access Denied；token 刷新无感 |


## 阶段一 预研与选型

- **任务**：判断项目是否该用 xlwings Server；在四种集成方式中选型；在 Server 与 Wasm 之间取舍。
- **要求**：逐条对照 1.1-1.4 的判据做决策；每个排除项必须写明原因，不留"待定"。
- **输出物**：见 1.5。
- **验收**：决策表四列齐全（方式+理由+排除项+原因）；结论与需求匹配，能一句话说清"为什么是它"。

### 1.1 xlwings Server 是什么

- **定义**：把 xlwings 能力服务化的官方开源参考工程——Excel 与 Google Sheets（含 Excel on the web）经浏览器或脚本调用服务端 Python，目标机无需安装 Python、无需加载项之外的前置依赖。
- **技术栈**：FastAPI（异步 Web 框架）+ Socket.IO（流式/实时）+ Redis（对象缓存与消息队列）+ Office.js（Excel 前端）+ 可选 WASM/Pyodide（浏览器内 Python）。
- **形态**：可自托管于裸机 / Linux VM / Docker Compose / Kubernetes / Serverless（见 `docs/index_hosting.md`）。
- **与桌面版的关系**：互补。桌面版（场景 A/C）在用户本机驱动 Excel；Server 把计算集中到服务端，天然适合多人共用、浏览器端用户、Linux/容器部署。
- **版本演进**：v1.0 起为标准 Python 包 `xlwings-server`（uv 管理，`uv run xlwings-server`）；pre-1.0 是"仓库即项目"形态。两代差异与迁移见阶段八 8.8。
- **与现代 Excel 加载项的关系**（官方定位 server.xlwings.org：“用 Python 而非 JavaScript 创建你自己的现代化 Excel 加载项”）：xlwings Server 就是 Office.js Excel 加载项的**后端运行时**——manifest 定义加载项 UI 与端点、任务窗格是前端，Server 提供 Python 业务后端、自定义函数/脚本端点、认证与前端集成；前端（任务窗格/按钮）由 Python 触发，不再写 JavaScript 业务逻辑。manifest 只负责把 Excel 指向后端（阶段五 5.5）。

### 1.2 适用与不适用

**适用场景**：

- 团队/组织级 Excel 自动化：多人共用一套 Python 能力，免去每人安装 Python 与依赖
- Excel on the web / Google Sheets：浏览器端没有 COM，只有服务端通道可用
- 服务器端批处理与集成：Docker/K8s/Serverless 上跑 Excel 计算，与现有系统对接
- 需要集中管控认证与安全：Entra ID 登录、角色、审计日志、敏感数据不出服务端

**不适用场景**：

- 单机离线、强交互式调试（仍用场景 A/C 桌面版）
- 对 xlwings API 全覆盖有硬需求（`docs/limitations.md` 列出的未覆盖成员，如 `App.calculate`、`Book.save`、`Range.formula` 等一批）
- 需要浏览器内离线运行的公开工具（那是 Wasm 的领域，见 1.4）

### 1.3 四种集成方式选型（`docs/integrations.md`）

| 方式 | 能力 | 优点 | 代价与限制 |
|------|------|------|-----------|
| Office.js 加载项（推荐） | 自定义函数 + 自定义脚本 + 流式函数 + 对象句柄 + 任务窗格 | 跨 Win/macOS/Web；集中部署（M365 管理中心/AppSource）；SSO + RBAC | 首次开发需 HTTPS 证书；生产部署需 M365 管理员 |
| VBA | 仅自定义脚本 | 独立 .xlsm/.xlsb 免加载项；可自建加载项 | 无函数/无认证/无 Web；VBA 面临禁用与病毒扫描误报压力 |
| Office Scripts | 仅自定义脚本 | 工作簿按钮；跨 Win/macOS/Web；单文件多簿复用 | 需 M365 + OneDrive/SharePoint；无认证；开发需隧道；不符合隐私标准 |
| Google Apps Script | 仅自定义脚本 | Google Sheets 内置调度器（cron 化）；Google SSO；菜单项 | 不是 Excel；无函数（计划中）；需隧道；内容托管于 Google |

**选型要点**：

- 函数（公式）、流式、句柄、认证是硬需求 → 只有 Office.js
- 已有 VBA 团队、只需脚本、要免管理员 → VBA
- 各集成可同时指向同一 Server（一个后端多通道接入）

### 1.4 xlwings Wasm 与 Server 的取舍

- **Wasm = 浏览器内 Python**（Pyodide），xlwings Server 同时充当其开发环境（`XLWINGS_ENABLE_WASM=true`，见 `docs/wasm_development.md`）。
- **Wasm 优势**：静态托管（Cloudflare Pages / GitHub Pages）、无后端、零运维。
- **Wasm 限制**（`docs/wasm_limitations.md`）：
  - 一切公开：Python 代码与 `app/wasm/.env` 对最终用户可见
  - 无认证（无 SSO/无自定义认证）
  - Pyodide 内存上限 2GB；无 TCP/IP（不能直连 PostgreSQL，SQLite 可用）；无调试器
  - 自定义函数/脚本只能全 Server 或全 Wasm 二选一，不能混跑
  - 仅支持 Office.js 集成
- **决策**：需要保密、认证、大数据、数据库直连 → Server；纯演示、公开轻量工具 → Wasm。

### 1.5 输出物

- 集成方式决策：表格记录（选定方式 + 理由 + 排除项及原因）
- 部署形态初判：自托管 / 容器 / Serverless 的初步倾向（阶段八细化）

### 1.6 客户端如何连接 Server：remote 引擎（`pro/_xlremote.py`）

- **定位**：xlwings 客户端侧连接 Server 的引擎（`xlwings/xlwings/pro/_xlremote.py`，5266 行完整实现），与 excel/calamine/officejs 并列为 PRO 四引擎；13-d 1.2.3 已压缩为指针，细节在此展开
- **通信方式**：HTTP REST API；对象操作经 `append_json_action` 打包为 JSON 动作（`func`/`args`/`values`/`sheet_position` 等）累积在 `_json["actions"]`，随请求一并发送服务端执行
- **对象覆盖**：Apps/App/Books/Book/Sheets/Sheet/Range/Names/Name/Pictures/Shapes/Charts/Characters/Note/PageSetup/Border/Borders/Font/FreezePanes/Table/Tables——与桌面引擎同构的高层 API，脚本代码无需感知通道差异
- **lazy load**：`_SHEET_VALUES_LOADED_KEY = "_xlwings_values_loaded"` 标记 sheet 值是否已加载（元数据先加载、值按需获取）；该标记仅本地读取，不序列化回 JS
- **元数据刷新**：`_update_api_in_place` 就地更新 API dict 并保留嵌套 dict 引用（Sheet/Name 对象持有引用，刷新后可见新值）
- **Office.js 协议映射**：`_CALCULATION_PY2JS`/`_CALCULATION_JS2JS`（xlwings `semiautomatic` → Office.js `AutomaticExceptTables`）；`_color_to_hex()` 把 RGB 元组/hex 串/整数统一为 `#RRGGBB`（`None` 透传 = 无填充）
- **版本校验**：客户端 `api["version"]` 与服务端 `__version__` 不一致时报错——两端必须同版本部署
- **适用场景**：服务器端无 Excel 环境、Linux 部署、Google Sheets 支持
- **与后续阶段的关系**：本阶段理解" + E + "客户端如何把请求发给 Server" + Q + "；阶段三起进入 Server 服务端如何接收与执行

## 阶段二 初始化与本地开发环境

- **任务**：建立可提交的开发仓库；跑通最小 Hello World；理解配置体系与环境分级。
- **要求**：按 2.1-2.5 顺序执行；证书、配置按环境分级（`.env` 敏感项不入库）；示例与开关按 2.5 处理。
- **输出物**：见 2.6。
- **验收**：仓库可提交；本地 HTTPS 开发服务器可启动且热重载生效；manifest 可 sideload；配置分工表明确。

### 2.1 步骤：仓库设置与项目初始化（`docs/repo_setup.md`、`docs/quickstart.md`）

- 克隆官方仓库改名项目；`git remote rename origin upstream`，另加自己的 `origin`（保留上游便于升级）
- 配 `.gitattributes` 并执行 `git config --local merge.ours.driver true` 减少日后 merge 冲突
- 安装 uv（`pip install uv`），同步依赖：仓库形态 `uv pip sync requirements-dev.txt`；v1.0+ 包形态 `uv sync`
- 初始化：仓库形态 `python run.py init`；v1.0+ `uv run xlwings-server init .`——生成 `.env` 并创建各环境 UUID
- 在 `.env` 顶部填 `XLWINGS_LICENSE_KEY`（试用 key 取自 xlwings.org/trial）；`.env` 不入 Git，需备份到安全处
- 只想本地玩可不建 origin（repo_setup 第 7 步注释）

### 2.2 步骤：开发证书（`docs/dev_certificates.md`）

- Office.js 加载项强制 HTTPS：用 mkcert 安装受信任自签证书
  - `./mkcert -install`
  - `./mkcert localhost 127.0.0.1 ::1` → 生成 `localhost+2.pem` / `localhost+2-key.pem` → 移入 `certs/`
- 替代路径：GitHub Codespaces（免证书）或隧道（`docs/tunneling.md`）
- 仅用 VBA / Office Scripts / Google Apps Script 集成时不需要证书（可跳过本步）

### 2.3 步骤：启动开发服务器

- 本地直接跑：`python run.py`（uvicorn 热重载；自动检测云端环境；注入 WASM 设置）→ 浏览器开 `https://127.0.0.1:8000` 应见 `{"status": "ok"}`
- Docker 开发环境（`docs/dev_docker.md`）：`docker compose up`——免本机装 Python，容器内 watchfiles 热重载
- GitHub Codespaces（`docs/github_codespaces.md`）：Create codespace → `python run.py init` → 填 key → `python run.py` → Ports 面板 Make Public → 得到公开 URL
- 热重载机制（`xlwings_server/hotreload.py`）：watchfiles 监视项目目录；dev 环境经 Socket.IO 向任务窗格 emit `xlwings:taskpane-reload`——改 Python/HTML 即时生效，无需手动刷新

### 2.4 步骤：配置体系（`docs/server_config.md` ↔ `xlwings_server/config.py`）

**源码事实**：`config.py` 用 pydantic-settings 定义 `Settings`（46 个字段），`settings_customise_sources` 实现优先级，`_create_settings()` 支持项目级覆盖，`jsconfig` computed 字段向前端注入运行配置。

- **配置源优先级**（源码顺序，高→低）：`Settings()` 初始化参数 > 环境变量（`XLWINGS_*`）> `.env` 文件 > `pyproject.toml [tool.xlwings_server]` > secret 文件
- **pyproject 键大小写不敏感**：`CaseInsensitivePyprojectTomlSource` 归一化键名并可带/不带前缀（`environment`/`ENVIRONMENT`/`XLWINGS_ENVIRONMENT` 均有效）
- **46 个设置字段按功能分组**（默认值以源码为准）：
  - 环境与身份：`environment`（默认 `prod`）、`license_key`、`hostname`、`project_name`、`functions_namespace`（默认 `XLWINGS`）、`xlwings_version`
  - 对象缓存：`object_cache_url`（生产必须）、`object_cache_expire_at`（默认 `0 12 * * sat`）、`object_cache_enable_compression`（默认 true）、`object_cache_maxsize`（默认 1000，仅内存 LRU 用）、`object_cache_partition_by_user`（默认 false）
  - 认证：`auth_providers`（默认空 = 匿名）、`auth_required_roles`、`auth_entraid_client_id` / `tenant_id` / `client_secret` / `multitenant` / `token_scopes`
  - 函数重试：`custom_functions_max_retries`（默认 3）、`custom_functions_retry_codes`（默认 `[500,502,503,504]`）
  - Socket.IO：`enable_socketio`（默认 true）、`socketio_message_queue_url`、`socketio_server_app`
  - 前端组件开关：`enable_htmx` / `enable_alpinejs_csp` / `enable_bootstrap`（默认均 true）、`taskpane_html`（默认 `taskpane.html`）、`taskpane_reload_attempts`（默认 3）
  - 安全：`add_security_headers`（默认 true）、`custom_headers`、`cors_allow_origins`（默认空）、`enable_excel_online`（默认 true）、`log_level`（默认 INFO）
  - 静态与 CDN：`cdn_officejs`（默认 false）、`pyodide_base_url`、`static_url_path`、`app_path`
  - 示例与开发：`enable_examples`（默认 false）、`enable_hotreload`（默认 true）、`enable_tests`、`enable_wasm`
  - manifest：`manifest_id_{dev,qa,uat,staging,prod}`（5 个环境 UUID，init 时生成）
  - 其他：`request_timeout`（默认 300s）、`secret_key`、`date_format`、`is_official_lite_addin`、`base_dir`
- **jsconfig 前端注入**：`jsconfig` computed 字段打包 appPath / authProviders / customFunctionsMaxRetries / customFunctionsRetryCodes / environment / isOfficialLiteAddin / pyodideBaseUrl / taskpaneReloadAttempts / onWasm / requestTimeout / xlwingsVersion——前端 `xlwings.js` 的运行配置即来自此（对应教程" + E + "模板中 `settings` 自动注入" + Q + "）
- **启动前序**：模块导入即写 `os.environ["XLWINGS_ON_SERVER"]="true"`；`PROJECT_DIR` 从 `XLWINGS_PROJECT_DIR` 或 cwd 解析并前置进 `sys.path`（config.py 与 main.py 双处执行——config 可能先于 main.py 被 custom_functions 导入）
- **init 默认**：`.env.template`（`xlwings_server/.env.template`）含 `XLWINGS_ENVIRONMENT="dev"`（未注释）——init 生成的 .env 默认 dev 环境，覆盖源码默认 `prod`；该文件也是 init 生成 .env 的唯一模板来源（不含 EXAMPLES 行，见 2.5）
- **环境分级语义**（与源码对齐）：
  - 每个环境用独立 UUID（`pyproject.toml [tool.xlwings_server]` 中生成，init 时写入）
  - 非 prod 环境 Ribbon 标签加环境后缀、函数命名空间加环境后缀（避免多环境同装冲突）；prod 不加后缀
  - dev 环境自动热重载 + 在 Excel 中显示全部 Python 异常详情；prod 只显示 `XlwingsError` 详情
### 2.5 步骤：示例与开关

- 内置示例**默认关闭**：`enable_examples=False` 是源码默认且 `.env.template` 不设该行（init 后即默认 false）；需要演练时在 `.env` 显式 `XLWINGS_ENABLE_EXAMPLES=true`（`docs/examples.md` 的演练步骤以显式开启为前提）
- 演练方式：单元格输入 `=XLWINGS.HELLO("world")`（prod）或 `=XLWINGS_DEV.HELLO("world")`（dev）；任务窗格点 Hello World 按钮（见 `docs/examples.md`）
- 上线前关闭：`XLWINGS_ENABLE_EXAMPLES=false`

### 2.6 输出物

- 可提交的仓库 + 已填 license 的 `.env`（备份到密码管理器）
- 本地 HTTPS 开发服务器 + 可 sideload 的 manifest
- 配置基线：`.env`（敏感）/ `pyproject.toml`（非敏感同环境项）分工表

## 阶段三 应用装配与扩展机制（服务骨架）

- **任务**：理解 `xlwings_server` 包的服务骨架（main.py 装配）与"用户扩展三通道"（目录注入 / 钩子 / 覆盖），学会挂载自定义路由、lifespan、静态资源与用户模型。
- **要求**：先读 main.py 装配序再动手挂载；扩展文件与官方模板结构对齐，不破坏默认装配。
- **输出物**：见 3.4。
- **验收**：每个挂载项真实生效（自定义路由可访问 / lifespan 触发 / 静态覆盖生效），且重启后仍在。

### 3.1 应用装配（`xlwings_server/main.py`）

- **项目目录注入**：`XLWINGS_PROJECT_DIR` → `sys.path` 前置（在 import 用户模块之前）；config.py 导入期同样执行（custom_functions 可能先于 main.py 导入 config）——用户代码放项目目录即被服务自动发现
- **中间件与 CORS**：`add_security_headers`（安全头，见阶段六 6.7）；CORS 仅 `allow_methods=["POST"]`、`allow_credentials=False`（对齐 Excel 前端调用语义）
- **Socket.IO 嵌套**：ASGIApp 与 FastAPI 同进程挂载（流式函数通道）
- **异常处理器分层**（可重试 vs 不可重试的语义设计）：
  - `XlwingsOperationalError` → **503**：瞬时/操作失败（如对象缓存后端不可达）→ 客户端应重试（对齐 `custom_functions_retry_codes`）
  - `XlwingsError` → **400**：确定性错误（参数错 / 角色缺失 / 非对象句柄）→ 客户端不应重试
  - `Exception` → **500**：未知错误，prod 环境隐藏细节
- **lifespan 别名清理**：`_cleanup_lifespan_alias` / `_cleanup_lifespan_module` 卸载用户 lifespan 模块并清理 `sys.modules` 别名，防热重载/重复导入污染
- **url_for 辅助**：`url_for(name, **path_params)` 统一生成应用内 URL（兼容 `app_path` 前缀）
- 借鉴：远程调用/自定义函数的错误语义用状态码区分"重试"与"放弃"，避免客户端盲目重试幂等错误

### 3.2 用户扩展三通道（对应 `xlwings_server/cli.py` 的 add 命令）

**通道一：lifespan.py（启动钩子）**：

- `uv run xlwings-server add lifespan` 生成
- 加载机制：`importlib.util.spec_from_file_location` 以**私有模块名** `_xlwings_server_user_lifespan` 加载（防与第三方 `lifespan` 包冲突）
- 铁律：文件存在但无 async lifespan 上下文管理器 → **fail fast**（宁可起不来也不静默无启动代码）——用于数据库连接池、缓存预热等

**通道二：routers/custom.py（自定义路由）**：

- `uv run xlwings-server add router` 生成；`importlib.util.find_spec` 探测后注册
- 端点可注入三类能力：`dep.User`（当前认证用户）、`dep.Book`（Excel 对象模型，htmx 场景）、`TemplateResponse`（Jinja 渲染）

**通道三：静态资源覆盖**：

- `OverridableStaticFiles`：用户目录优先、包目录兜底——同路径用户资源覆盖默认资源（如 `css/style.css`、`js/main.js`、`images/`）

**附加扩展点**：

- `models/user.py`：`add model user`——自定义 User 模型（阶段六 6.4）
- `auth/custom/`：`add auth custom`（阶段六 6.5）
- `auth/entraid/jwks.py`：`add auth entraid`——气隙 JWKS（阶段六 6.3）
- `config.py`：`add config`（阶段二 2.4）

### 3.3 CLI 命令全集（`xlwings_server/cli.py`，64KB，最大模块）

- **run**：`run`（开发服务器，等价 `python run.py`）
- **add 系**：`init` / `config` / `model user` / `auth custom` / `auth entraid` / `router` / `lifespan` / `azure functions` / `docker` / `devcontainer` / `iis`
- **wasm 系**：`wasm settings` / `wasm build`（阶段八 8.7）
- **migrate**：`xlwings-server migrate <old-project>`（pre-1.0 → v1.0，阶段八 8.8）

### 3.4 输出物

- 服务骨架认知图：装配（目录注入 → 中间件 → CORS → Socket.IO 挂载 → 异常分层）→ 扩展三通道
- 已挂载的自定义路由 / lifespan / 静态覆盖文件（每件标注用途与触发时机）


## 阶段四 核心能力开发——自定义函数

- **任务**：用 `@func` 把 Python 函数暴露为 Excel 公式；掌握转换器、命名空间、文档字符串、流式函数、对象句柄与后续脚本；理解服务端端点与序列化框架的支撑。
- **要求**：函数先过转换器/命名空间/文档字符串规范（4.1-4.4）；流式/句柄仅在确有场景时引入（4.5-4.7）。
- **输出物**：见 4.9。
- **验收**：每个函数经真机公式调用返回预期值；元数据无重复 Excel 名；流式/句柄（若有）验证通过。

### 4.1 语法与注册（`docs/custom_functions.md` + `custom_functions/examples.py`）

- 最小函数：`@func def hello(name): return f"Hello {name}!"` → 单元格 `=HELLO("world")`
- **模块注册铁律**：新增模块必须 `from .myfunctions import *` 回导出到 `custom_functions/__init__.py` 包级——只有能从包直接 import 的函数才会注册进 Excel；仅定义在子模块不够
- 开发期函数体修改自动重载（热重载通道）；增删函数或改签名需重启 Excel
- 显示名：`@func(name="helloName")` 保留大小写（0.36.12+；仅 Lite/Server 支持，经典加载项忽略）

### 4.2 转换器（`@arg`/`@ret`、类型提示、Annotated）

- pandas DataFrame：`@arg("df", pd.DataFrame, index=False, header=False)` + `@ret(index=False, header=False)`
- 类型提示等价写法：`df: Annotated[pd.DataFrame, {"index": False}]`；可提取为类型别名 `Df = Annotated[...]` 复用
- 优先级：装饰器与类型提示并用时，**装饰器覆盖类型提示**
- 变长参数：`@arg("*args", pd.DataFrame, index=False)`——转换器作用于 `*args` 全部实参
- 日期时间：
  - 读侧必须用转换器（xlwings 不自动识别日期格式单元格）：单值 `@arg("date", dt.datetime)`；批量 `xw.to_datetime()`；DataFrame 用 `parse_dates`（同 pandas.read_csv 语义）
  - 写侧自动按 Excel 文化格式；可用 `@ret(date_format="yyyy-m-d")` 或 `XLWINGS_DATE_FORMAT` 覆盖
- 完整转换器清单见官方 `docs/converters`（链接在 custom_functions.md）

### 4.3 命名空间与子命名空间

- `@func(namespace="numpy")` → 显示为 `NUMPY.STANDARD_NORMAL`
- 模块级 `__xlwings_func_namespace__ = "numpy"` → 该模块所有函数进同一命名空间；`@func` 的 namespace 参数优先
- 排除个别函数：`@func(namespace="")`；manifest 配置的全局命名空间仍生效
- 子命名空间：`namespace="numpy.linalg"` 式层级（custom_functions.md 续文）

### 4.4 文档字符串与函数向导

- 函数 docstring 描述函数；`@arg("name", doc='A name such as "World"')` 或 `Annotated[str, {"doc": ...}]` 描述参数
- 二者进入 Excel 公式向导 / formula builder；参数名自动进入单元格输入联想

### 4.5 流式函数（RTD，实时数据）

- **场景**：市场数据、传感器、任务进度等持续推送——函数先返回占位值，随后数据到达更新单元格
- **服务端支撑**（`xlwings_server/routers/socketio.py`）：`sio_function_call` / `sio_cancel_task` 事件；`AsyncRedisManager` 作为 Socket.IO 消息队列
- **部署要求**：流式函数依赖 Socket.IO 服务（有状态单 worker + Redis 桥接），精简部署形态不支持（阶段八 8.2）
- 性能注意：Socket.IO 服务不应管理 CPU 密集的长任务（阻塞事件循环）；流式函数适合外部数据源的 async 拉取（httpx/aiohttp）

### 4.6 对象句柄（跨调用保持 Python 对象）

- **场景**：链式 UDF `=GetData() → FilterData(A1) → PlotData(B1)`——单元格只能存基本类型，句柄字符串指向服务端缓存对象，解决"Excel 单元格 ↔ 服务端对象"的跨调用保持
- **服务端支撑**（`xlwings_server/object_handles.py`）：
  - `RedisObjectCache`（生产）：zlib 可选压缩；croniter 到期（默认每周六 12:00）；按用户分区（`XLWINGS_OBJECT_CACHE_PARTITION_BY_USER`；无认证用户时 **fail loud**——隔离控制不静默降级）；`evict_superseded` 把生产者映射存 Redis，**跨 worker 清理旧代**（并发 get/set 非原子 → 最坏一代泄漏至过期，自愈）
  - `_WarningLRUObjectCache`（开发期）：内存 LRU + **每次写都告警**（提醒生产须 Redis）+ **与 Redis 同一序列化路径**（dev/prod 行为一致：不支持的类型在开发期就在产生单元格报错）
- **过期处理**：`ObjectCacheMissError` → 返回 `stale_object_handle` 可操作卡片（用户可刷新/重算），而非 `#VALUE!`（`routers/xlwings.py`）
- 借鉴：缓存后端的"开发期告警 + 生产一致性"——开发内存、生产 Redis，同一序列化路径保证行为一致

### 4.7 高级能力

- `xw.Caller`：函数内访问调用单元格（地址/行列/表名）——v1.12+
- `xw.WithScript`：函数返回后自动执行指定 `@script`（`routers` 校验 `@script` 标记并解析 include/exclude/lazy 元数据）——v1.12+
- `CachedObject`：函数间共享耗时对象
- 错误语义：函数内抛 `XlwingsError` → 400（不可重试）；`XlwingsOperationalError` → 503（可重试，客户端配 retry codes）

### 4.8 服务端端点与序列化框架（源码支撑）

**`xlwings_server/routers/xlwings.py` 端点族**：

- `GET /xlwings/custom-functions-meta(.json)`：函数清单元数据（`.json` 为兼容别名路由）
- `GET /xlwings/custom-functions-code.js`：**动态生成客户端 JS**——`inspect.getmembers` 枚举 `__xlfunc__` 属性 → 生成 async 包装 + `CustomFunctions.associate`
- `POST /xlwings/custom-functions-call`：执行端点——`typehint_to_value={CurrentUser, Caller}` **类型提示注入**；`ObjectCacheMissError → stale_object_handle`；`CustomFunctionResult` + `WithScript` 后续脚本
- `POST /xlwings/custom-scripts-call/{name}`：脚本端点（阶段五 5.2）
- `GET /xlwings/custom-scripts-meta(.json)`：脚本清单（`.json` 为兼容别名路由）
- `POST /xlwings/alert`：前端弹窗通道
- `GET /pyodide.json`：WASM 模式动态包清单
- 日志注入防护：`sanitize_log_input`（换行 → `\n` 字面量）——服务端日志处理不可信输入

**`xlwings_server/serializers/` 注册表模式**：

- `serializers/__init__.py`：`serialize(obj, serializer_name=”default”)` / `deserialize(payload, serializer_name)` 总入口，按 `serializer` 键路由到注册表；`framework.py`：`Serializer` 基类 + `register(*types)` 注册表（名称与类型双键）；`custom_encoder/custom_decoder` 接入 json（datetime → isoformat）
- 四类内置：`pandas_serializer.py`（DataFrame/Series ↔ JSON：`to_json(date_format="iso")` + dtypes 记录 + **MultiIndex 临时列名映射还原** + **DatetimeIndex freq 保留**——pandas 对象跨进程往返的完整样板）、`numpy_serializer.py`、`dictionary_serializer.py`（datetime 键转 `[k,v]` 对）、`default_serializer.py`
- 借鉴：新类型接入只需 `Serializer` 子类 + `register` 一行注册

### 4.9 输出物

- 函数集（含转换器/命名空间/文档字符串规范）
- 若有流式/句柄/WithScript 场景：对应实现 + 服务端缓存配置（Redis URL/分区/压缩）与部署要求记录

## 阶段五 核心能力开发——自定义脚本与任务窗格

- **任务**：用 `@script` 写整簿操作脚本，从任务窗格 / 工作表 / 功能区三处触发；构建任务窗格 UI；产出并安装 manifest。
- **要求**：脚本恰一个 book 参数（对齐 5.1）；三通道按 5.2 逐个验证；任务窗格按 5.3-5.4 模板构建。
- **输出物**：见 5.6。
- **验收**：三通道至少一条真机跑通；任务窗格渲染无控制台错误；manifest 安装验证记录完整。

### 5.1 `@script` 语法与配置（`docs/custom_scripts.md`）

- 最小脚本：`@script def hello_world(book: xw.Book)`——**必须有 `xw.Book` 类型提示参数**（脚本与 UDF 的签名差异点）
- 配置项：
  - `include`/`exclude`：控制发送到后端的表内容。默认**整簿全发**；含大数据表时必须 exclude 整表（当前仅支持表名列表），否则调用慢或超时
  - `required_roles`：RBAC（阶段六 6.4）
  - `button` + `show_taskpane`：工作表按钮（5.3）
  - 任意参数（v1.7+）：脚本可带 str/int/bool/list/dict 参数，任务窗格经 `xw-args` JSON 数组传入

### 5.2 三种触发通道

**任务窗格按钮（推荐）**：

- `<button xw-click="hello_world" class="btn btn-primary btn-sm">Hello World</button>`——xlwings 前端集成自动关联到服务端脚本
- 带参：`xw-click="write_value" xw-args='["Hello!", "B2"]'`

**工作表按钮（shapes + hyperlink 变通）**：

- 步骤：Insert > Shapes 画圆角矩形 → 名称框命名（如 `xlwings_button`）→ 右键 Link/Hyperlink → 本工作簿位置填入被覆盖单元格（如 `B4`）
- 装饰器：`@script(button="[xlwings_button]Sheet1!B4", show_taskpane=True)`
- 让工作簿打开即自动加载加载项：浏览器控制台 `await Office.addin.setStartupBehavior(Office.StartupBehavior.load)`，随后保存工作簿；取消用 `Office.StartupBehavior.none`
- 原理：加载项注册选中事件监听——点击超链接选中 B4 → 触发脚本 → 立即选中下方单元格备用
- 局限：
  - 脚本依赖当前选区时不可用（点击会改变选中单元格）
  - 名称与引用必须一致：`button=[button_name]Sheet1!A1`
  - Excel on the web 不能给形状加超链接（桌面设置好的簿在 Web 可用）

**功能区按钮（`static/js/ribbon.js` 方案）**：

- `templates/manifest.xml` 定义 `ExecuteFunction` 按钮（`FunctionName: hello-ribbon`）
- `static/js/ribbon.js`：`helloRibbon(event)` 取 auth（`globalThis.getAuth`）→ `xlwings.runPython({auth, scriptName, headers: {"Auth-Provider": provider}, exclude})` → `event.completed()`
- **注意**：Ribbon 调用不尊重 `@script` 的 include/exclude 配置，须在 `runPython` 调用里显式传（`exclude: "SomeSheet,SomeOtherSheet"`）
- 可选：`Office.ribbon.requestUpdate` 在请求期间禁用按钮（finally 恢复）
- 官方 TODO：Ribbon 开发体验待改进（issue #102）

### 5.3 任务窗格前端栈（`docs/taskpane_intro.md` + jinja/htmx/alpinejs/bootstrap）

- **定位**：任务窗格 = VBA UserForm 的现代替代（纯 Web 应用，最大灵活性）
- **开箱四件套**（共同点：免 JS 构建工具、兼容最严格 CSP、库非框架可逐个替换、易上手）：
  - **Jinja**：服务端模板引擎。`TemplateResponse` 渲染；`settings` 变量自动注入（如 `settings.environment`）；`jinja2-fragments` 支持返回局部片段供 htmx；partial 约定文件名前导下划线（`_xxx.html`），`{% include "_book.html" %}` 复用
  - **htmx**：客户端-服务端交互无整页刷新。`hx-post`/`hx-target`/`hx-swap`；集成点：`xw-book="true"` 属性 → 后端端点注入 `book: dep.Book` → 模板 `{% include "_book.html" %}` 渲染 Excel 对象模型——业务逻辑与状态保持在 Python 后端
  - **Alpine.js CSP 版**：客户端交互。`x-data`/`x-show`/`@click`/`x-text`；CSP 版要求代码全部放在 HTML 外；`registerAlpineComponent("name", obj)` 连接 HTML 与 JS
  - **Bootstrap 5 + bootstrap-xlwings**：Excel 外观的 UI 工具包（响应式，任务窗格可扩到半个 Excel 网格）
- **安全**：HTML 响应一律走 `TemplateResponse`（自动转义用户输入）；htmx 安全配置在 `static/js/integrations/htmx.js`
- **结构建议**：复杂任务窗格按 `app/static/js/components/{auth,dashboard,shared}/` 组织；按需 `{% block extra_head %}` 加载；`base.html` 先载 `integrations/alpinejs-csp.js` → 自研 → `vendor/@alpinejs/csp/dist/cdn.min.js`

- **任务窗格路由**（`routers/taskpane.py`）：`GET /taskpane` 与 `GET /taskpane.html` 渲染 `taskpane_html` 设置指定的模板——前端入口是独立路由而非静态文件，因此可注入 `settings` 与认证上下文

### 5.4 任务窗格示例（`xlwings_server/templates/examples/` 9 个）

- `hello_world`：最小按钮 + 脚本
- `htmx_form`：表单 POST + 局部渲染（`_greeting.html`）
- `excel_object_model`：htmx + `xw-book` 访问 Excel 对象模型（`_book.html`）
- `auth`：受保护页 + htmx 认证头（阶段六 6.6）
- `navigation`：多页导航；`multi_app`：多应用切换；`pictures`：图片操作；`alpine`：Alpine 组件；`live_form_validation`：实时表单校验
- 每个目录自带 README；路由与页面写法见 `docs/auth_taskpane.md`

### 5.5 manifest 与安装（`docs/manifest.md` + `docs/install_officejs_addin.md`）

- **manifest.xml 是 Jinja 模板**（`xlwings_server/templates/manifest.xml`）：
  - 环境相关 Id（UUID 存 `pyproject.toml`，init 时生成）
  - 非 prod 环境 Ribbon 标签/命名空间加环境后缀 → 同一加载项多环境并存不冲突
  - URL 不对时设 `XLWINGS_HOSTNAME`
- 常见编辑点：`Version`/`ProviderName`/图标 URL/Ribbon 组与按钮/`OfficeTab` 定位（默认自定义选项卡，可并入 Home）
- **安装三通道**：
  - 开发 sideload：复制 `https://SERVER/manifest` 内容存为 xml → 按平台 sideload（Windows 网络共享目录 / macOS Developer add-ins / Web 直接上传）
  - 内部生产：M365 管理中心 > Integrated Apps > Upload custom apps > 链接 manifest（须公网可访问）
  - 公开分发：AppSource（需 Microsoft Partner；**必须** `XLWINGS_CDN_OFFICEJS=true` 按微软要求走 CDN）
- **清缓存**（加载项行为异常/更新不生效时）：Windows 删 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`（及 Win32WebViewHost INetCache）；macOS 执行 `scripts/clear_office_cache_macos.sh`；Web 刷新即可
- 从 sideload 切到管理中心/AppSource 前先移除 sideload 清单并清缓存

### 5.6 输出物

- 脚本集 + 触发方式（任务窗格按钮 / 工作表按钮 / ribbon.js 三选或组合）
- 任务窗格页面（Jinja 模板 + JS 组件 + 资源目录结构）
- 定制 manifest + 安装验证记录（三通道中至少一条跑通）

## 阶段六 认证、授权与安全

- **任务**：用 `XLWINGS_AUTH_PROVIDERS` 开启认证；用 RBAC / `is_authorized` 分级授权；锁定任务窗格受保护页；配置安全头与日志防护。
- **要求**：认证先开基础再上 SSO（6.1→6.2）；RBAC 角色先定义后实现（6.3）；任务窗格按 6.5 受保护。
- **输出物**：见 6.7。
- **验收**：未登录返回 Access Denied；函数/脚本按角色放行与拒绝正确；token 刷新无感；安全头生产生效。

### 6.1 认证基础（`docs/authentication.md`）

- 默认无 provider：所有人可调用函数/脚本，用户对象为 `User(id='n/a', name='Anonymous', email=None, domain=None, roles=[])`
- 开启：`XLWINGS_AUTH_PROVIDERS=["entraid"]`（或 `["custom"]`）→ 函数/脚本端点立即要求登录
- **关键认知**：认证只保护函数/脚本端点，**不自动锁定任务窗格**——任务窗格各页是否公开由你自己决定（6.6）
- 当前用户注入：函数/脚本参数用 `CurrentUser` 类型提示（`from xlwings_server.models import CurrentUser`）

### 6.2 Entra ID SSO（`docs/auth_entraid.md`）

- **前提**：在 Azure 门户把加载项注册为 Microsoft Identity Platform 应用（learn.microsoft 官方指南）
- **配置**：`XLWINGS_AUTH_PROVIDERS=["entraid"]` + `CLIENT_ID` + `TENANT_ID`；外部组织加 `MULTITENANT=true`
- 更新 manifest：从 `/manifest` 端点取新清单重新 sideload（重启 Excel）
- **角色（RBAC 数据源）**：Azure 门户 App roles 创建角色（如 `Task.Write`）→ Enterprise applications 分配用户/组 → 角色出现在 `User.roles`，供 `required_roles` 消费
- **OBO 流（代表用户调 Microsoft Graph）**：
  - 配置：`CLIENT_SECRET` + API permissions（如 `User.Read`、`Sites.Read.All`、`Files.Read.All`）+ **Grant admin consent**（OBO 无法交互式授权，缺 consent 得 403）
  - 使用：`async with current_user.get_graph_client() as graph:` → `graph.get("/me")` 等；`GraphClient` 提供 get/post/patch/delete，路径相对 `https://graph.microsoft.com/v1.0`，kwargs 透传 httpx（params/json/headers/timeout），返回 `httpx.Response`
  - 安全：OBO token 只存服务端，**永不回传 Excel 客户端**；仅返回 Graph 调用结果
- **JWKS 气隙部署**：air-gapped 服务器无法直连微软 → `uv run xlwings-server add auth entraid` 生成 `auth/entraid/jwks.py`，自定义 `get_jwks_json` 读 cron 每日下载的 `https://login.microsoftonline.com/common/discovery/v2.0/keys`
- **源码实现**（`auth/entraid/`）：`get_jwks_json_default`（默认从微软 discovery 端点拉取）→ `get_key_set`（缓存键集）→ `validate_token`（JWT 签名校验 + 过期检查 + `auth_required_roles` 消费）；`obo.py` 的 `acquire_obo_token`（内存缓存 token，提前 5 分钟刷新缓冲）与 `GraphClient`（get/post/patch/delete 封装 httpx）——教程中的 SSO/OBO 每一步都有对应源码锚点
- **依赖注入链**：`dependencies.py` 的 `authenticate` → `get_user`；端点参数声明 `dep.User` 即获得当前用户，`dep.Book` 即获得 Excel 对象模型（htmx 场景）

### 6.3 授权与 RBAC（`docs/authorization.md`）

- 函数/脚本级：`@script(required_roles=["xlwings.admin", "xlwings.user"])` / `@func(required_roles=[...])`——用户缺角色则拒绝
- 全局级：自定义 User 模型实现 `is_authorized()`——如仅允许 `mydomain.com` 域名：`return self.domain == "mydomain.com" if self.domain else False`；或 RBAC 组合：`return await self.has_required_roles(["xlwings.admin", "xlwings.user"])`
- 自定义 User 模型：`uv run xlwings-server add model user` → `models/user.py`（Pydantic / dataclass / SQLAlchemy 均可）
- 多 provider：请求需带 `Auth-Provider` header 指明走哪个 provider

### 6.4 自定义认证（`docs/auth_providers.md`）

- `uv run xlwings-server add auth custom` → 生成 `auth/custom/__init__.py` + `static/js/auth.js`，并激活 `.env` 的 `XLWINGS_AUTH_PROVIDERS=["custom"]`
- 后端：实现 `validate_token` 函数
- 前端：编辑 `globalThis.getAuth` 返回 `{token, provider}`：
  ```js
  globalThis.getAuth = async function () {
    return { token: "test-token", provider: "custom" };
  };
  ```
- 其他集成（VBA/Office Scripts/Google Apps Script）：经 `runPython` 的 `auth` 参数或 `Authorization` 头传 token

### 6.5 任务窗格认证（`docs/auth_taskpane.md` + `templates/examples/auth`）

- 原则：landing 页保持公开；其余页加 `dep.User` 依赖注入；每请求带 `Authorization` 头
- 示例：自定义路由中 `current_user: dep.User` 依赖 + `TemplateResponse` 渲染 `examples/auth/protected.html`，context 注入 `current_user`
- htmx 场景参考 `templates/examples/auth`（受保护页 + 认证头传递样板）

### 6.6 安全配置（`xlwings_server/security_headers.py` + `dependencies.py`）

- `add_security_headers`：安全头中间件（`security_headers.py` 的 `resolve_security_headers` 决定最终响应头集合）——CSP 兼容（前端栈设计为最严格 CSP 可用）；Excel Online/CDN 感知；图片扩展豁免 Cache-Control；`enable_excel_online=false` 时安全头更严格（生产清单建议）
- `custom_headers`：自定义响应头注入
- CORS：`allow_methods=["POST"]`、`allow_credentials=False`
- 日志：`sanitize_log_input`（换行 → `\n` 字面量）防日志注入
- RBAC 判据：`dependencies.py` 的 `has_required_roles` + `is_authorized()` 全局授权钩子（授权入口单点）

### 6.7 输出物

- 认证 provider 配置（`.env`/`pyproject.toml`）+ 更新后的 manifest
- 授权策略：角色定义 + 函数/脚本级与全局 `is_authorized` 实现
- 任务窗格受保护路由 + 安全头基线（生产 `XLWINGS_ADD_SECURITY_HEADERS` 生效验证）

## 阶段七 测试、调试与性能

> 质量心智图：**Check 管规约、Test 管行为、Debug 管失败归因**——7.1-7.2 建立测试门禁（Check/Test），7.3 是失败归因闭环（Debug），7.5 是排障沉淀；每轮改动后跑门禁，全绿才进阶段八。

- **任务**：建立三层测试（单元/集成/端到端）；以官方 `tests/` 套件为模板；掌握服务器与 Office.js 双端调试；应用性能配方；沉淀故障排查清单。
- **要求**：按 7.1 官方结构建测试；每轮改动后跑门禁；任何失败进入 7.3 调试闭环（复现→定位→修复→回归→沉淀）。
- **输出物**：见 7.6。
- **验收**：三层测试全绿（对齐官方 41 项结构）；性能基线有改造前后对比；故障排查清单已沉淀。

### 7.1 官方测试体系（`tests/` 41 项）

**单元测试层（TestClient 打真实端点，mock 外部依赖如 Redis/Graph）**：

- `test_router_xlwings.py`（28KB / 586 行）：custom-functions / custom-scripts 端点全行为
- `test_object_handles.py`（14KB / 355 行）：缓存后端行为（Redis/LRU 两实现）
- `test_serializers.py`：pandas / numpy / dict / default 四类往返
- `test_auth_entraid.py` + `test_obo.py`：JWT 校验与 OBO 流
- `test_security_headers.py` + `test_custom_headers.py`：安全头
- `test_log_sanitization.py`：日志注入防护
- `test_dependencies.py` / `test_user_model.py`：依赖注入与 User 模型
- `test_lifespan.py`（8.8KB）：用户 lifespan 钩子（**子进程隔离**验证 fail-fast 语义）
- `test_router_manifest.py` / `test_router_taskpane.py`：manifest 与任务窗格路由
- `test_static_files.py` / `test_static_file_hasher.py`：静态资源与指纹缓存
- `test_examples.py` / `test_custom_functions.py` / `test_custom_scripts_args.py` / `test_env2.py`：示例、参数与多环境

**集成测试层（部署变体矩阵）**：

- `tests/` 下 `docker-compose.apppath.yaml` / `docker-compose.redis.yaml` / `docker-compose.socketio.yaml`：分别验证子路径挂载、Redis 有无、Socket.IO 有无三种组合下的端到端行为

**端到端测试层（真实 Excel 运行）**：

- `e2e_custom_functions.py`（23KB / 963 行）+ `e2e_custom_functions.xlsx`：真实 Excel 实例运行自定义函数并断言
- `e2e_custom_scripts.py`：真实脚本执行
- `tests/object_handles.md`：**手工验证清单**——对象句柄无法全自动覆盖的 7 类行为兜底（卡片隐藏 UUID、`=A1` 复制解析、过期卡片等）
- 环境文件：`.env.test` / `.env.test2` / `.env.testwasm`（多环境配置隔离）

### 7.2 测试方法借鉴（为自己的 Server 服务写测试）

- 单元：TestClient + 打真实端点（不 mock 框架，mock 外部服务）
- 集成：docker-compose 变体矩阵——逐个开关（redis / socketio / apppath）验证组合
- e2e：真实 Excel 运行 + xlsx 断言（不可自动化部分用手工清单双轨）
- 多环境：`.env.test*` 系列隔离测试配置

### 7.3 服务器调试（`docs/debugging.md`）

- VS Code：`.vscode/launch.json` 配 debugpy（`"module": "xlwings_server.cli"`、`PYTHONPATH` 项目根、`.venv/bin/python`）→ F5 断点调试
- Office.js 加载项双端：
  - Windows 桌面：右键任务窗格 `Inspect` → DevTools
  - macOS 桌面：先执行一次 `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true` → 重启 Excel → 右键 `Inspect Element`（事后可改回 false）
  - Excel on the web：浏览器 F12（噪音多，注意过滤无关报错）
- 调试关注点：Console（错误）、Network（资源加载失败/服务器 URL/请求头）
- WASM：无调试器 → `print()` 打到浏览器控制台；或临时 `XLWINGS_ENABLE_WASM=false` 回 Server 断点调试

### 7.4 性能配方（`docs/performance.md`）

- 脚本：**include/exclude 控制发送表内容**——整簿全发是默认行为，大数据表是首要瓶颈
- 函数：**动态数组代替多单元格函数**（100 次调用 → 1 次）；减少函数数量
- 异步：FastAPI 是 async 框架 → IO 用 `httpx`/`aiohttp`（替代 requests）、`asyncpg`/`psycopg3`（替代 psycopg2）
- 缓存：`functools.cache` 装饰慢函数（注意：缓存按 app worker 独立，非共享）；客户端侧缓存未实现（issue #86）
- 流式：Socket.IO 服务不跑 CPU 密集长任务（阻塞事件循环）；流式函数适合外部数据源 async 拉取

### 7.5 常见故障排查（`docs/troubleshooting.md`）

- 加载项行为异常 → 先清 Office 缓存：Trust Center > Trusted Add-in Catalogs 勾选"下次启动清除"；或删 `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`（Windows）/ 执行 `scripts/clear_office_cache_macos.sh`（macOS）
- `Addin xlwings Server failed to download a required resource` → 服务必须 **https + 域名**（不能用 IP）
- 单元格 `#NAME?` / `#BUSY!` → 清缓存
- 工作表按钮无反应 → 检查按钮名称与 `button=[name]Sheet!Cell` 一致、未预选中引用单元格、重载加载项

### 7.6 输出物

- 测试套件（单元 / 集成矩阵 / e2e / 手工清单四层，对齐官方结构）
- 调试记录 + 性能基线（改造前后对比）

## 阶段八 部署、运维与升级

- **任务**：选择并落地部署形态；以官方 `deployment/`、`nginx/`、`scripts/` 为起点；落实生产配置清单；规划迁移与升级；明确平台限制。
- **要求**：按 8.1-8.7 顺序执行（选形态→配置→核查→迁移规划）；部署物以官方模板为起点再裁剪。
- **输出物**：见 8.9。
- **验收**：部署配置真实可起；生产清单逐项核查通过；迁移/升级有执行记录与版本基线。

### 8.1 部署形态总览（`docs/index_hosting.md` + 各托管族）

| 形态 | 文件 / 文档 | 要点 |
|------|------------|------|
| Docker Compose 全量 | `deployment/docker-compose.prod.yaml` | nginx + app（2 workers）+ socketio + redis 四容器 |
| Docker Compose 精简 | `deployment/docker-compose.prod-min.yaml` | 单 app 容器；无流式函数；对象句柄需外部 Redis（`XLWINGS_OBJECT_CACHE_URL`） |
| Windows Server IIS | `docs/windows_server.md` + `add iis` | HttpPlatformHandler；app pool 无托管代码；NetworkService 权限；PFX 证书 |
| Azure Container Apps | `docs/azure_container_apps.md` | 容器镜像 + gunicorn 命令覆盖；2 CPU/4Gi；ingress 8000 |
| Azure Functions | `docs/azure_functions.md` + `add azure functions` | 无服务器；`XLWINGS_ENABLE_SOCKETIO=false`（无 WebSocket → 无流式）；`uv export requirements.txt` |
| AWS App Runner | `docs/aws_app_runner.md` + `apprunner.yaml` | 不支持 WebSockets → 无流式 |
| Render | `docs/render.md` | 免费层测试；空闲冷启动约 30 秒；生产建议托管 Redis + 双容器 |
| 裸机 / Linux VM / K8s | README / DEVELOPER_GUIDE | 自托管，gunicorn + nginx 手动编排 |

### 8.2 Docker Compose 详解（`deployment/`）

- `docker-compose.prod.yaml` 四容器职责：
  - `nginx`：反向代理 + TLS 终止（证书放 `certs/privkey.pem` + `certs/fullchain.pem`）
  - `app`：xlwings Server 应用（gunicorn + uvicorn worker，2 worker）
  - `socketio`：流式函数专用服务（有状态，单 worker）
  - `redis`：两大用途——Socket.IO 消息队列 + 对象句柄缓存
- `docker-compose.prod-min.yaml`：仅 app 容器；无流式；对象句柄指向外部 Redis
- 操作命令：
  - 构建：`docker compose -f deployment/docker-compose.prod.yaml build`
  - 启动：`docker compose -f deployment/docker-compose.prod.yaml up -d`
  - 日志：`docker compose -f deployment/docker-compose.prod.yaml logs -f`

### 8.3 nginx 反向代理（`nginx/` 3 conf）

- `nginx-prod.conf`：主站 TLS 终止 + 静态指纹缓存
- `nginx-apppath.conf`：子路径挂载变体（`XLWINGS_APP_PATH` 非根路径部署）
- `nginx-socketio.conf`：`/socket.io/` 路径的 **WebSocket Upgrade**（Upgrade/Connection 头正确转发）
- 三 conf 分离的意义：流式函数依赖 WebSocket，反向代理必须正确转发 Upgrade；HTTP 缓存指纹与 WASM 禁缓存清单分路径控制

### 8.4 构建与镜像脚本（`scripts/`）

- `mirror_officejs.py`（13.6KB）：从 CDN 拉取 office.js，按 build version 目录化（离线/自托管 office.js 用）
- `officejs_inventory.py`：office.js 资源清单盘点
- `postinstall.py`：npm install 后 → 前端资产版本化并入 `vendor/` 并回写引用
- `start.sh`：容器启动入口（uvicorn/gunicorn 启动参数）
- `clear_office_cache_macos.sh`：macOS 清 Office 加载项缓存
- 工程借鉴：前端资产"版本化目录 + 引用回写 + 指纹缓存破坏"三步处理

### 8.5 生产配置清单（`docs/production.md` ↔ 源码/部署物）

- **任务窗格**：编辑 `templates/taskpane.html`（无需可用则删除该文件）
- **项目目录解析**：从仓库根启动 ASGI 进程，或设 `XLWINGS_PROJECT_DIR` 指向仓库根（`config.py` 的 `PROJECT_DIR` 解析逻辑）
- **环境**：`XLWINGS_ENVIRONMENT="prod"`（或 qa/uat/staging）——禁用热重载、隐藏非 `XlwingsError` 异常细节
- **品牌化**：manifest 与设置双处改——`XLWINGS_FUNCTIONS_NAMESPACE="YOUR_NAME"`、`XLWINGS_PROJECT_NAME="Your Name"`，并替换 `templates/manifest.xml` 中的 xlwings 引用与图标
- **关闭不用的库**：`XLWINGS_ENABLE_ALPINEJS_CSP=false`、`XLWINGS_ENABLE_HTMX=false`、`XLWINGS_ENABLE_SOCKETIO=false`（对应 config.py 的 `enable_*` 字段）
- **公开分发**：发布 Excel AppSource（Marketplace）须 `XLWINGS_CDN_OFFICEJS=true`（按微软要求从 CDN 加载 Office.js）
- **日志**：`XLWINGS_LOG_LEVEL="INFO"`（DEBUG 会记录敏感 token）
- **安全头**：`XLWINGS_ADD_SECURITY_HEADERS=true`；启用认证 `XLWINGS_AUTH_PROVIDERS=["entraid"]`
- **CORS**：不用 Office Scripts 且不用 Excel on the web 的自定义函数 → `XLWINGS_CORS_ALLOW_ORIGINS=[]`；用 Excel on the web 自定义函数或 Office Scripts → `["*"]`
- **Excel on the web**：不用则 `XLWINGS_ENABLE_EXCEL_ONLINE=false` 换取更严格安全头
- **workers**：每 CPU 核心 2-4 个；gunicorn 命令见 `deployment/docker-compose.prod.yaml`（`--workers 2 --timeout 30 --worker-class uvicorn.workers.UvicornWorker`）；Serverless 平台无需关心 worker 数
- **timeout 两层**：gunicorn（默认 30s，`--timeout` 调高）+ 反向代理/负载均衡（nginx `proxy_read_timeout 60s;` 等，须 ≥ gunicorn）
- **Redis 两大用途**：`XLWINGS_SOCKETIO_MESSAGE_QUEUE_URL`（pub-sub 连接 app workers 与 socketio 服务）+ `XLWINGS_OBJECT_CACHE_URL`（跨 worker 共享对象缓存）——生产必须 Redis
- **socketio**：有状态协议，必须恰好 1 worker 独立进程，经 Redis 与 app workers 连接；可支撑数千并发连接；不应承载 CPU 密集长任务（避免阻塞事件循环）——外部数据用 async HTTP（httpx/aiohttp）查询
### 8.6 xlwings Wasm 静态部署（`docs/wasm_deployment.md`）

- 构建：`python run.py wasm https://domain`（v1.0+ `uv run xlwings-server wasm build`）→ 输出 `dist/`；选项 `-z`（附 zip）、`-c`（先清输出）、`-e ENV`（覆盖环境）
- 托管：
  - Cloudflare Pages：zip 手动上传 / Git 仓库连接 / CI（wrangler-action）
  - GitHub Pages：`python run.py wasm https://username.github.io/reponame -o <repo子目录>`
- CDN 策略：
  - 首次构建建议 `XLWINGS_CDN_OFFICEJS=true` + `XLWINGS_CDN_PYODIDE=true`（dist 文件少且小；office.js 与 Pyodide 从 CDN 加载）
  - 离网/内网：两者设 false，`.whl` 拷入 `xlwings_server/static/vendor/pyodide/`，`app/wasm/requirements.txt` 以路径 `/static/vendor/pyodide/mypackage.whl` 引用
- 安装：`https://domain/manifest.xml` sideload（`docs/install_officejs_addin.md`）
- 开发侧（`docs/wasm_development.md`）：`XLWINGS_ENABLE_WASM=true` 用 Server 作开发环境；`app/wasm/requirements.txt` 独立依赖（Pyodide 先查 PyPI 纯 Python wheel，再查 Pyodide 官方包库）；Python 版本由 Pyodide 决定；`app/wasm/.env` 自动从 `.env` 同步但**会随产物公开**（敏感项勿放）
- 限制复盘：全公开、无认证、2GB 内存、无 TCP（SQLite 可用）、无调试器、函数/脚本 Server/Wasm 二选一、仅 Office.js（`docs/wasm_limitations.md`）

### 8.7 迁移与升级（`docs/migration.md` + `docs/upgrade.md`）

**pre-1.0 → v1.0 迁移**：

- 变化本质：标准 Python 包分发（uv 管理）替代"仓库即项目"；Python 3.10 最低；`uv run xlwings-server` 替代 `python run.py`
- 迁移命令：新目录 `uv init` + `uv add xlwings-server watchfiles` + `uv run xlwings-server init .` + `uv run xlwings-server migrate C:\path\to\old-project`
- 逐项核对：
  - 依赖：`requirements.in` → `pyproject.toml dependencies`（`uv add`）
  - 自定义静态文件（除 css/style.css、js/main.js、images/ 外）拷贝到新项目 `static/`
  - Azure Functions：重新 `add azure functions`，拷贝 `local.settings.json`；部署前 `uv export --format requirements.txt -o requirements.txt`
  - User 模型：重新 `add model user`，迁回逻辑（不再继承 BaseUser，模型已简化）
  - 自定义 config / auth / router / lifespan / jwks：分别 `add` 后迁回
  - 导入统一：`from xlwings_server.models import CurrentUser`、`from xlwings_server.dependencies import User`、`from xlwings_server.templates import TemplateResponse`、`from xlwings_server import settings`

**升级（仓库形态）**：

- `git fetch upstream` + `git merge --no-edit <VERSION>`；冲突解决后 `python run.py deps compile`；`uv pip sync requirements-dev.txt` 或 `docker compose build`
- 官方推荐 merge 而非 rebase（冲突单步解决、无需 force push、保留自己的提交历史）

**版本事实**：`docs/changelog.md` 为版本权威（本手册基准 1.14.0，2026-09-09）；技能侧 17 路上游 manifest 锁定 tag `1.14.0`，升级时同步刷新本手册与 13-d 中的版本引用。

**平台维护责任**：Azure Functions 的 host 运行时版本与 Python EOL 需人工跟踪（`docs/azure_functions.md` 明确列出跟踪项）。

### 8.8 限制与替代（`docs/limitations.md` + `docs/missing_features.md` + `docs/alternatives.md`）

- API 覆盖缺口（选列）：
  - `App`：calculate / status_bar / screen_updating / interactive / enable_events / calculation / quit 等
  - `Book`：save / to_pdf
  - `Range`：formula / formula2 / merge / unmerge / paste / autofill / note / hyperlink / font / current_region 等一批
  - `Shape` / `Chart`：部分属性与删除
  - `Font`：只写不读（size/italic/color/name/bold）
- 对策：`docs/missing_features.md` 提供 workaround 写法（如缺失的 `Book.save` 可用宏/脚本通道实现）
- 替代路径：功能缺口不可接受 → 桌面版（场景 A/C）；需求是公开轻量 → Wasm

### 8.9 输出物

- 选定形态的部署配置（docker-compose / nginx / IIS / 云平台 / Wasm）
- 生产清单核查记录（workers/timeout/Redis/socketio/缓存/示例关闭）
- 迁移或升级执行记录（含版本基线）

## 附录 A 源码级工程借鉴

按主题横切六块内容，提炼可复用的工程模式（均为源码事实，非建议空谈）：

- **服务化骨架**：目录注入 + 钩子探测 + 覆盖优先级（用户扩展三通道）——`xlwings_server/main.py`
- **错误语义**：503（可重试）/ 400（不重试）/ 500（未知）三分法，客户端配 retry codes——`main.py` 异常处理器
- **分布式对象缓存**：压缩 / 定时过期 / 按用户分区 / 跨 worker 清理 / 开发期告警 + 生产一致性——`object_handles.py`
- **序列化注册表**：`Serializer` 子类 + `register()`，新类型一行接入——`serializers/framework.py`
- **认证栈**：JWT + jwks 缓存 + OBO 委托 + RBAC 钩子 + 自定义 provider 双端（validate_token / getAuth）——`auth/`、`dependencies.py`
- **前端集成**：htmx 服务端状态保持 + `xw-book` 对象模型透传 + CSP 兼容栈（Jinja/htmx/Alpine-CSP/Bootstrap）——`templates/`、`static/js/integrations/`
- **构建脚本**：前端资产版本化目录 + 引用回写 + 指纹缓存破坏——`scripts/`
- **测试分层**：单元（TestClient 打真实端点）/ 集成（docker-compose 变体矩阵）/ e2e（真实 Excel + xlsx 断言）/ 手工清单（对象句柄 7 类行为）——`tests/`

## 附录 B 官方文档族地图（61 篇 → 阶段）

| 文档族 | 文档（docs/） | 对应阶段 |
|--------|--------------|---------|
| 导航族（11 篇） | `index.md`、`index_authentication.md`、`index_dev_prerequisites.md`、`index_development_environment.md`、`index_hosting.md`、`index_officejs_addins.md`、`index_other_integrations.md`、`index_server.md`、`index_taskpane.md`、`index_tutorials.md`、`index_wasm.md` | 全阶段入口 |
| 快速上手 | `quickstart.md`、`repo_setup.md`、`local_development.md`、`server_config.md` | 阶段二 |
| 功能 | `custom_functions.md`、`custom_scripts.md`、`manifest.md`、`install_officejs_addin.md`、`examples.md` | 阶段四/五 |
| 任务窗格 | `taskpane_intro.md`、`jinja.md`、`bootstrap.md`、`htmx.md`、`alpinejs.md` | 阶段五 |
| 认证 | `authentication.md`、`authorization.md`、`auth_entraid.md`、`auth_providers.md`、`auth_taskpane.md` | 阶段六 |
| 集成 | `integrations.md`、`vba_integration.md`、`officescripts_integration.md`、`googleappsscript_integration.md`、`tunneling.md` | 阶段一 |
| 性能调试 | `performance.md`、`debugging.md`、`troubleshooting.md` | 阶段七 |
| 部署托管 | `docker_compose.md`、`dev_docker.md`、`devcontainers.md`、`dev_certificates.md`、`github_codespaces.md`、`aws_app_runner.md`、`azure_container_apps.md`、`azure_functions.md`、`render.md`、`windows_server.md` | 阶段二/八 |
| WASM | `wasm_development.md`、`wasm_deployment.md`、`wasm_limitations.md` | 阶段一/八 |
| 迁移限制 | `migration.md`、`upgrade.md`、`limitations.md`、`missing_features.md`、`alternatives.md`、`dependencies.md`、`changelog.md`、`versioning-implementation-plan.md`、`license.md` | 阶段八 |

## 附录 C 与 13-d 的分工

- `references/13-scenario-d-source-code.md` 1.2.3（remote 引擎，客户端侧 `pro/_xlremote.py`）：细节已迁移至本手册阶段一 1.6（13-d 保留对比要点 + 指针）——客户端如何连 Server 与本手册服务端如何工作，构成同一协议的两端
- `references/13-scenario-d-source-code.md` 1.7（本手册前身）：已压缩为指针，细节全部上移至此
- `references/13-scenario-d-source-code.md` 4.6（SQLite 数据访问模式）：通用写法，保留在 13-d；Server 后端写数据访问可借鉴其参数化查询、lru_cache 缓存、config 驱动 DB 类型切换的做法
- 语义区分：13-d 回答"xlwings 的引擎与源码是什么"，本手册回答"如何用 xlwings Server 构建并运维一个服务"
