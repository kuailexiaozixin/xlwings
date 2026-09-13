# 排障指南

> 聚焦 FastHTML + pywebview 桌面应用常见的组合场景排障，
> 以及本技能特有的打包与运行问题。

---

## 一、环境与项目初始化

### `uv` 命令找不到

确保已运行 `./scripts/ensure_uv_env.ps1`。安装后需重启终端或手动刷新 PATH。

### `python -m venv .venv` 失败

确保 Python 已正确安装。运行 `python --version` 确认。

### `uv add` 慢或失败

设置中国镜像：
```bash
uv config set index-url https://mirrors.aliyun.com/pypi/simple/
```

---

## 二、编码与运行

### `SyntaxError: positional argument follows keyword argument`

FastHTML 组件中位置参数（子元素）必须在关键字参数（HTML 属性）之前：
```python
# 错误
Div(cls="box", H1("标题"))

# 正确
Div(H1("标题"), cls="box")
```

### 静态文件 404

**原因**：`static_path` 在打包后路径映射失效。

**修复**：所有 CSS/JS 内联到 `Style()` / `Script()`：
```python
app, rt = fast_app(
    hdrs=(Style("body { margin:0; }"),)
)
```

### pywebview 窗口白屏

1. 检查 FastHTML 服务是否在运行：浏览器访问 `http://127.0.0.1:5001`
2. 如果浏览器能访问但窗口白屏 → 服务启动时序问题，添加 `wait_for_server()`
3. 如果都访问不了 → 端口被占用，使用 `find_free_port()` 自动探测

### HTMX 请求返回完整页面而非片段

**原因**：请求缺少 `HX-Request` 头。

**修复**：确认 HTMX 属性拼写正确（用下划线）：
```python
Button("加载", hx_get="/data", hx_target="#result")
```

---

## 三、打包

### 打包后 `ModuleNotFoundError: No module named 'xxx'`

**原因**：PyInstaller 漏扫了依赖模块。

**修复**：使用 `--hidden-import xxx` 显式声明。

常见缺失模块：
| 模块 | 修复 |
|------|------|
| `clr` | `--hidden-import clr` |
| `webview.platforms.edgechromium` | `--hidden-import webview.platforms.edgechromium` |
| `webview.platforms.winforms` | `--hidden-import webview.platforms.winforms` |

### 打包后 `serve()` 卡死或无响应

**原因**：`serve()` 内部使用 `reload=True` + import-string，冻结后入口模块路径不可用。

**修复**：改用 `uvicorn.run(app, reload=False)`。

### 打包后 EXE 体积过大（>100 MB）

**原因**：PyInstaller 扫描了系统环境中的所有包（torch、scipy 等）。

**修复**：使用最小 venv 打包（仅含面板运行所需依赖）。

### 打包后 EXE 启动立即崩溃（无任何输出）

**原因**：系统 DLL 依赖缺失。pywebview + pythonnet 组合依赖 `libffi-8.dll`（注意名称是 libffi-8.dll，不是 ffi-8.dll）、`libcrypto-3-x64.dll`、`libssl-3-x64.dll`，
PyInstaller 的二进制依赖分析器不会自动追踪这些 `ctypes.CDLL()` 动态加载的 DLL。

**修复**：
1. 打开 CMD 或 PowerShell，拖入 EXE 查看错误输出
2. 或从 CMD 运行：`cd dist && MyApp.exe`
3. 根据报错信息对照下表：

| 错误信息 | 缺失 DLL | 修复（onefile 口径） |
|---------|---------|------|
| `DLL load failed while importing _ctypes` | `libffi-8.dll`（旧文档误写作 `ffi-8.dll`） | 用 `--add-binary "<py>/DLLs/libffi-8.dll;."`（路径以 sys.base_prefix 递归查找为准；标准安装为 `<py>/Library/bin/libffi-8.dll;.`） 显式打包 |
| `No module named \'_ssl\'` | `libcrypto-3-x64.dll` + `libssl-3-x64.dll` | 同时用 `--add-binary` 打包两个 openssl DLL |
| `DLL load failed: 找不到指定的模块` | 依赖链断裂 | 用 Dependency Walker 或 `dumpbin /dependents` 分析，逐一对缺失 DLL 加 `--add-binary` |

> **为何不能"复制到 dist/_internal/"**：本技能禁用 `--onedir`，只产出 onefile 单文件 EXE，运行时不存 `dist/_internal/` 目录（`_internal/` 是 onedir 产物）。DLL 必须随 EXE 用 `--add-binary` 一并打包，运行时由 PyInstaller 解包到临时目录加载。`build_windows_exe.ps1` 已自动对 `libffi-8.dll / libcrypto-3-x64.dll / libssl-3-x64.dll` 执行 `--add-binary`，一般无需手动处理。

完整 DLL 诊断：确认 `libffi-8.dll / libcrypto-3-x64.dll / libssl-3-x64.dll` 已 `--add-binary` 打包。

### 打包后 EXE 启动提示 `ModuleNotFoundError: No module named 'app'`

**原因**：`--add-data "src;src"` 未生效。可能原因：
1. 路径含空格导致参数被拆分
2. 在 Bash 中执行而非 PowerShell
3. `--add-data` 分隔符错误（Windows 用 `;`，Linux 用 `:`）

**修复**：
- 切换到 PowerShell 执行打包命令
- 路径含空格时使用 Python subprocess 脚本执行打包命令
- onefile 下资源解包到临时目录 `_MEIPASS`，不在 `dist/_internal/`；请确认打包命令包含 `--add-data "src;src"`

### 路由 404 / 页面导航全部失效（开发态或打包后）

**原因**：多文件 `APIRouter()` 路由用函数名自动生成（下划线转连字符），且类型注解参数被当作**查询参数**而非路径参数。页面里手写的 RESTful href（如 `/projects/1`）与自动路由（`/product-detail?id=1`）对不上，导致全站 404。

**修复**：给每个路由显式声明路径，路径参数用 `{name}`：

```python
@ar("/projects/{pid}")        # 显式路径
def project_detail(pid: int): ...
```

导航/模板中的 href 必须与显式路径完全一致（APIRouter 默认路由陷阱：根路径与子路径叠加时须用显式 `@apiroute("/path")`）。

### 打包后 EXE 窗口白屏，但 HTTP 在浏览器中可访问

**原因**：服务启动时序问题。pywebview 创建窗口时 FastHTML 服务尚未就绪。

**修复**：确认 `src/main.py` 中在 `webview.start()` 前等待服务就绪：
```python
# main.py 中应包含类似逻辑
import requests, webview
for _ in range(30):
    try:
        requests.get(f"http://127.0.0.1:{PORT}", timeout=1)
        break
    except Exception:
        import time; time.sleep(1)
webview.start()  # 服务就绪后再启动窗口
```

### 打包后 EXE 启动时报 `Failed to execute script 'main'`

**原因**：打包时入口脚本路径错误、依赖缺失或 Python 版本不兼容。

**修复步骤**：
1. 确认使用 `--console` 模式（能看到具体错误堆栈）
2. 从 CMD 运行 EXE 捕获完整回溯信息
3. 根据回溯中缺失的模块添加 `--hidden-import`
4. 检查是否在正确的虚拟环境中打包
5. 如果回溯指向 `serve()` 或 `reload=True`，改为 `uvicorn.run(app, reload=False)`

### `sqlite3.OperationalError: unable to open database file`

**原因**：打包后 `__file__` 指向临时解压目录 `_MEIxxxx`，无写权限。

**修复**：使用 `sys.frozen` 检测：
```python
BASE_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
```

---

## 四、组合场景排障（本技能特有）

### 窗口白屏但 HTTP 服务正常

**根因**：pywebview 加载 URL 时 FastHTML 服务尚未就绪。

**修复**：在开窗口前等待服务就绪：
```python
import requests
for _ in range(30):
    try:
        requests.get(f"http://127.0.0.1:{PORT}", timeout=1)
        break
    except:
        import time; time.sleep(1)
```

### 关闭窗口后进程不退出

**根因**：uvicorn 线程不是 daemon 线程。

**修复**：`threading.Thread(target=run_server, daemon=True)`

### 端口被占用

**根因**：固定端口冲突。

**修复**：使用 `find_free_port()` 自动探测空闲端口。

### 打包后窗口能开但页面加载失败

**根因**：使用 `localhost` 而非 `127.0.0.1`，某些环境 localhost 解析被拦截。

**修复**：URL 中一律用 `http://127.0.0.1:{PORT}`。

### 打包后控制台中文乱码

**根因**：Windows 控制台默认 GBK 编码。

**修复**：
- 代码输出中不要使用 emoji
- 使用 `[OK]` / `[FAIL]` 替代 ✅ / ❌
- 必要时在入口设置 `sys.stdout.reconfigure(encoding='utf-8')`

### WebView2 运行时缺失

**根因**：用户电脑未安装 WebView2 Runtime（Windows 10/11 通常已内置）。

**修复**：提示用户下载安装 [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)。

### `run_command` 120s 超时导致构建中断

**现象**：Agent 用单条 shell 命令（heredoc / `echo >`）一次性创建十几个骨架文件，命令执行超过 `run_command` 工具默认的 120 秒超时 → 被强制中断 → 项目半成品、进程「已停止」，构建彻底失败。

**根因**：`run_command` 工具默认超时 **120 秒**。单条 shell 命令批量造文件（`cat <<'EOF' > file`、`echo ... > file` 连写多个文件）耗时易超限，且超时后无法部分恢复。

**修复**：
1. **创建源码/配置/资源文件必须逐文件使用 Write/Edit 工具**（每个文件独立、可超时隔离、可并行）。严禁 `cat <<'EOF' > file`、`echo ... > file` 在一个命令里连写多个文件。
2. **`run_command` 只做「真正需要 shell」的事**：启动服务、跑 `uv`/`pip`/`pytest`/`pyinstaller`、调构建脚本等。且任何可能 >120s 的操作必须显式传大 `timeout`（如 `timeout=600`），或后台运行 + 轮询确认完成。
3. **骨架用 bootstrap 脚本**：新建项目一律跑 `bootstrap_project.ps1`（或 `.sh`）生成骨架，禁止为「并行建骨架」而把多文件塞进一条终端命令。
4. **构建（PyInstaller `--onefile`）必须带大超时或后台**：常 >120s，务必 `run_command(..., timeout=600)` 或后台启动并在 `dist/*.exe` 出现 + 冒烟通过后再继续。
5. **失败即拆步**：一旦某命令超时，把它拆成更小的步骤（如先建目录 → 再逐文件 Write → 再单独 `uv sync` → 再单独 `pyinstaller`），每步独立可验证。

### HTTP 200 但后台业务子进程崩溃（假绿）

**现象**：打包后的 EXE 启动后，FastHTML 主服务返回 HTTP 200（冒烟测试通过），但应用有后台业务子进程（如 Hermes 网关、第三方 LLM 代理等），该子进程实际已崩溃或未就绪，导致用户操作时功能不可用——冒烟测试显示「绿色」但实际是「假绿」。

**根因**：冒烟测试只验证了 FastHTML 主服务的 HTTP 200，未验证后台子进程的业务健康端点。子进程崩溃不影响主服务响应，但业务功能依赖子进程。

**修复**：
1. **冒烟测试须验证所有业务健康端点**：用 `src/health_endpoints.txt` 或构建脚本的 `-HealthCheckUrls` 参数声明关键端点（如 `http://127.0.0.1:8642/health`），全部返回 200 才放行。
2. **从 CMD 运行 EXE** 查看完整控制台输出（`--console=True`），确认子进程启动日志无报错。
3. **检查子进程端口**：确认子进程监听端口是否可达（`netstat -ano | findstr <port>`）。

---

### UI 视觉质检常见误报与修复

> 以下案例来自 `examples/01-hermes-desktop` 的 `ui_window_verify.py` 实际运行，2026-08-08 / 08-09 验证。

**现象**：脚本报告"几何重叠"或"对比度不达标"的缺陷，但人工视觉检查页面显示正常。

**根因 1：被裁剪元素产生假重叠**

可滚动容器（`overflow:auto/hidden`）内被滚出可视区的子元素，其 `getBoundingClientRect` 仍落在容器坐标内，会与底部固定元素算出「假重叠」。

**修复**：脚本已加 `isClipped()` 守卫（沿祖先链检查是否越出任一 `overflow` 非 `visible` 祖先的可视边界），裁剪元素不参与重叠与对比度计算。

**根因 2：浅色主题文字对比度不达标**

在 `examples/01-hermes-desktop` 上，浅色主题的 CSS 变量 `--text-faint`、`--text-dim`、`--accent` 在浅色背景上的对比度比值分别为 2.39、2.97、4.16，均低于 WCAG 2.1 AA 标准要求的 4.5:1。脚本正确检出这些缺陷，但人工视觉检查时因颜色"看起来还行"而容易忽略。

**修复**：将 `--text-faint` 从浅灰调深、`--text-dim` 从灰调深、`--accent` 从浅蓝调深，确保所有文字颜色与背景对比度 ≥ 4.5:1。共修复 **6 处**对比度问题。

**根因 3：侧边栏纵向溢出导致按钮不可达**

侧边栏容器 `overflow:visible`（默认值），内部 `conv-list` 设置了 `min-height` 导致内容超出容器高度，底部按钮被裁切且无法滚动到达。脚本正确检出了横向/纵向溢出。

**修复**：将侧边栏设为 `overflow-y:auto` 并移除 `conv-list` 的 `min-height`。

**验证结果**：修复后 `ui_window_verify.py` 复跑 **0 缺陷通过**。

**经验总结**：
- `isClipped()` 守卫是重叠检测的必备组件，否则可滚动容器内元素会产生大量假阳性。
- 浅色主题的文字颜色必须用工具验证对比度，人工视觉不可靠（人眼对低对比度不敏感）。
- 可滚动容器必须显式设置 `overflow:auto` 或 `overflow-y:auto`，禁止依赖 `visible` 默认值。

---

## 五、测试夹具与 db schema 对齐（防止 fixture rot）

> 由「测试夹具与 db schema 对齐」实战复盘沉淀。测试夹具的 `insert` 字段若与真实表 schema 不一致，会导致整批测试静默失败，且错误定位成本高。

### fixture 引用不存在的列，整批测试静默失败

**现象**：pytest fixture 凭印象写 `insert` 字典，引用了 db 模型**不存在的列**，例如：

- `Expense.subject_code` / `Expense.cost_element` / `Expense.batch_no` —— 真实模型只有 `subject_id / person_id / description / amount / phase / caliber ...`
- `Organization.code` / `Organization.is_rd_unit` / `Organization.district` —— 真实表只有 `id / level / name / parent_id`，上述字段均为凭记忆臆造

**后果**：整批测试在收集/运行时静默失败，且错误信息是"列不存在"而非"测试逻辑错"，排错成本高（曾因此导致 13 个新测试全挂）。

**根因**：测试用例不是由真实 db 模型**生成**的，而是 AI/人凭记忆构造；技能此前缺少"fixture 字段必须 1:1 对齐真实 schema"的约束，也没有任何自动化护栏。

**修复（强制规则）**：
1. 任何 fixture 的 `insert` 字典键，**必须**来自对应表真实列；禁止想象字段。
2. 新增/修改 db 模型列后，必须同步更新所有引用该表的 fixture（反之同理）。
3. 项目**必须**接入 `fixture_schema_helper.assert_fixture_matches_schema`，在 pytest 收集或测试内做契约校验，漂移即失败。

### 用 fixture_schema_helper 做契约校验

`scripts/fixture_schema_helper.py`（纯标准库，可复制到任意项目）：

```python
from fixture_schema_helper import assert_fixture_matches_schema

def test_expense_fixture_align(fresh_db):
    rows = [{"subject_id": 1, "person_id": 1, "amount": 100.0,
             "phase": "研究", "caliber": "高企认定"}]
    # 若 rows 出现 subject_code 等不存在的列，立即 AssertionError
    assert_fixture_matches_schema(fresh_db.t.expense, rows, "expense")
```

- 兼容 fastlite `Table.__columns__`、`dataclass.__dataclass_fields__`、`pydantic.model_fields`。
- 脚本自带自测：`python fixture_schema_helper.py` 能抓到故意构造的漂移并退出码 1。

**落地位置**：
- 在 `conftest.py` 或各 `test_*.py` 的顶部，用一条参数化测试覆盖所有表的 fixture 对齐，作为**第一道门禁**先于业务逻辑测试运行。
- 与 `references/09-testing-debugging-guidance.md` 的"TDD 理念"互为补充：TDD 管逻辑，本规则管**数据契约**。

**关联门禁**（与 `release_gate.py` 叠加）：
- `release_gate.py` → 流程级门禁（pytest 全绿才走到打包）
- `fixture_schema_helper` → 测试内契约（从根上不让漂移的 fixture 进绿）
