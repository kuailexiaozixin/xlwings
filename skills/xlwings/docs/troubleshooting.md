# Release 复刻流水线踩坑实录（2026-09-08，AnnouncementDownloader 实战）

本次用 AnnouncementDownloader 项目把 release 复刻链（convert → release_clone → reinject_ribbon → gate_a）从"静默失败"修到"真机 E2E 全 PASS"，共挖出 4 个坑。全部修复已进 `../scripts/release_tool/` 源码。

## 坑 1：LF 换行符的 Dictionary.cls → VBE Import 编译错误 → 模态框挂死（最隐蔽）

- **症状**：release 后的 xlsm 里任何宏（包括无辜的 Ping 测试函数）一运行就挂死，隐藏 Excel 实例弹模态编译错误框，自动化进程全部卡死。
- **根因**：技能源码目录里的 `Dictionary.cls` 是 **LF 换行**（15093 字节），VBE 的 `VBComponents.Import` 对 LF-only 文件按损坏内容导入 → 全工程编译失败。同内容 CRLF 版（15567 字节，差值=行数）导入一切正常。
- **诊断手法**：分阶段二分（每步保存→重开→Ping），锁定"导入 Dictionary.cls"这一步。
- **修复**：`release_clone.py` 导入任何 VBA 组件前强制 CRLF 规范化。
- **教训**：**VBE 只认 CRLF**。所有要 Import 的 .bas/.cls/.frm 必须先规范化换行符。

## 坑 2：用户级 `xlwings.conf` 的 `USE UDF SERVER=True` 劫持所有工作簿

- **症状**：RunPython 弹 `运行时错误 '1000': Could not activate Python COM server`。
- **根因**：`GetConfig` 回退链是 工作簿配置表 → 目录级 → **用户级 `%USERPROFILE%\.xlwings\xlwings.conf`** → 默认值。官方 CLI 曾往用户级写入 `"USE UDF SERVER","True"`，让所有工作簿走 XLPy COM 服务器分支；而便携运行时没有 `xlwings32-*.dll` → COM 激活失败。
- **诊断手法**：可见 Excel 复现 + Pillow 截屏 + win32 枚举模态对话框读文本；再对照 bas 源码反查 `OPTIMIZED_CONNECTION = GetConfig("USE UDF SERVER", False)`。
- **修复**：release 写配置时显式写 `USE UDF SERVER=False`（工作簿配置表优先级最高，覆盖用户级）。
- **教训**：用户级 conf 是隐形全局变量，交付工作簿必须在工作簿配置表里显式钉死关键键。

## 坑 3：便携运行时被注册表劫持，site-packages 完全没启用（交付杀手）

- **症状**：真机 E2E 里嵌入代码从未执行（工作簿一个单元格都没被写过），手动带 PYTHONPATH 复现又一切正常。
- **根因**：`install_runtime.py` "复制已有 Python"方案产出的目录里没有 `python3xx._pth`，`python.exe` 启动时通过**注册表**（`HKCU\Software\Python\PythonCore\3.13`）解析到源机器 prefix → sys.path 全指向源机器（`D:\Python\...`），便携自己的 `Lib\site-packages` 反而不在 sys.path：
  - `import release_tool` → ModuleNotFoundError（真机没这个目录）
  - `import xlwings` 解析到源机器路径 → 目标机上直接崩
- **诊断手法**：对比"带/不带 PYTHONPATH"的 `sys.path`，发现源机器路径全在里面而便携 site-packages 缺席。
- **修复**：创建 `python313._pth`（隔离模式：忽略注册表与 PYTHONPATH）：
  ```
  python313.zip
  .
  Lib
  DLLs
  Lib/site-packages
  import site
  ```
  同时 `install_runtime.py` 新增：`_make_pth_file`（自动创建 ._pth）、`_bundle_runtime_embedded`（打包 release_tool.runtime_embedded 进 site-packages）、`_purge_machine_specific_pth`（清理复制残留的源机器绝对路径 .pth，如 `__editable__*.pth`、marvis 残留）。
- **教训**：**复制式便携 Python 必须配 ._pth**，否则它只是源机器 Python 的"影子"，交付即翻车。

## 坑 4：Excel 把 '000001' 存成 1.0，股票代码丢失前导零

- **症状**：`未找到股票代码 1.0 的 orgId`。
- **修复**：两道防线——① 业务代码归一化（去 `.0` 后缀 + `zfill(6)`）；② 测试/前端写 B1 前设 `number_format='@'`（文本格式）。

## 坑 5：便携运行时瘦身重建的三连坑（2026-09-08 下午）

瘦身（清空源 site-packages + requirements 重建）后连续踩三个静默坑：

1. **`python3._pth` 文件名 bug**：`glob("python3*.dll")` 会先命中 `python3.dll`（转发 DLL），导致创建了无效的 `python3._pth`，隔离模式没生效。必须用正则 `python3\d+\.dll` 过滤取带版本号的。
2. **numpy "source directory" 假错误**：numpy 2.x 的 OpenBLAS DLL 在 site-packages 顶层的 `numpy.libs/` 目录（不在 numpy/ 里），只复制 numpy/ 目录必崩。pandas.libs 同理。凡 `*.libs` 顶层目录都要带上。
3. **单文件 stub 模块是复制式部署的经典漏项**：`six.py`、`typing_extensions.py`、`clr.py`、`pythoncom.py`、`pywintypes.py` 都是单文件（不是目录），按目录名复制清单会漏。pywebview 的 Windows 后端启动时 `import clr`（pythonnet stub），漏了就报 "You must have pythonnet installed" 但真实原因是缺 clr.py。

## 坑 6：xlwings 引擎静默加载失败（`has no attribute '_xlwindows'`）

- **症状**：`Book.caller()` 报 `module 'xlwings' has no attribute '_xlwindows'`，但 `import xlwings` 完全正常。
- **根因**：xlwings 0.37 `__init__.py` 里 `from . import _xlwindows` 包在 `try/except ImportError: pass` 中——pywin32 组件不全（漏 `win32com/`、`pythoncom.py` 等）时引擎加载**静默失败**，错误延迟到第一次用 COM 时才爆。
- **修复**：`install_runtime.py::_copy_pywin32` 完整复制 pywin32 组件；`_verify_engine` 硬门禁 `assert hasattr(xlwings, '_xlwindows')`。
- **教训**：**import 成功 ≠ 引擎加载成功**。验证断言必须打到最后一级实际使用的属性/对象，不能用"import 不报错"糊弄。

## 坑 7：`._pth` 隔离模式吃掉 cwd —— `-m` 找不到包

- **症状**：`python -m src.main` 报 `No module named 'src'`，但 src 就在 cwd 下。
- **根因**：`python3xx._pth` 隔离模式下 `sys.path` 完全由 ._pth 决定，cwd（''）不再自动进 path；`-m` 靠 cwd 找包。
- **修复**：改用**脚本方式启动**（`python panel_boot.py`，脚本内显式 `sys.path.insert(0, dirname(__file__))`），不依赖 cwd。
- **教训**：给便携运行时写启动命令时，一律脚本方式 + 显式 sys.path，禁用 `-m`。

## 坑 8：交付形态的功能降级必须有"用户旅程"级验证

- **症状**：`open_panel` 交付形态降级为"激活说明表"，产物级断言全绿（不报错、状态正确），但用户点按钮"没有任何响应"——真正期望是面板弹出。
- **教训**：回归断言必须写**用户预期行为**（面板 HTTP 200、窗口出现），禁止用"无异常/状态正确"糊弄；断言冻结原则：实现变更不得反向修改断言。已把面板旅程断言（A1b HTTP 200）加入回归脚本。
- 附带坑：fasthtml/starlette 会往 cwd 写 `.sesskey`，从 `C:\` 等只读目录启动面板会 PermissionError——面板进程必须以可写目录为 cwd。

## 坑 8b：面板进程的持久化文件不能放 `__file__` 目录（临时目录每次都是新的）

- **症状**：面板报"无法获取外部数据列表（网络异常）"，且本地缓存兜底永远不生效。
- **根因**：面板代码被解包到 `%TEMP%\xlwings_panel\<时间戳>\` 运行，`CACHE_FILE = os.path.join(BASE_DIR, ...)` 里的 BASE_DIR 是**临时目录**——每次打开面板都是全新目录，缓存文件写了也活不过本次会话，兜底机制形同虚设。
- **修复**：持久化文件（缓存/配置）一律放 `%LOCALAPPDATA%\<应用名>\`；`__file__` 目录只当"本次会话临时区"用。
- **教训**：凡是"进程在临时目录里跑"的交付形态（解包式面板、嵌入脚本），判断持久化路径时要问一句：这个目录下次还在吗？

## 坑 8c：上游接口不稳定不是偶发，要按"约三成失败率"设计

- **症状**：外部 API 的静态数据文件（数百 KB JSON）实测时好时坏——同机同代码，一次 ReadTimeout、一次数秒后 200；逐请求头对照实验所有组合都 200，排除 WAF 拒头，纯属服务端网关抖动（慢/504）。
- **修复**：重试提到 5 次（Retry backoff_factor=2）+ 共享缓存兜底（面板路径与 RunPython 路径共用 `%LOCALAPPDATA%` 缓存，命中后 1 小时内不联网）。
- **教训**：诊断网络类报错先做**单变量对照实验**（裸 requests vs Session、逐请求头拆解），区分"请求被拒"和"接口随机抖动"——两者修法完全不同；缓存兜底的位置必须先于重试参数检查，否则兜底是假的。

## 坑 8d："失败才读缓存"顺序反了 = 用户干等到放弃（查询空白）

- **症状**：面板查询按钮点了没任何结果（不报错、不转圈）——真实原因是股票列表走"先联网 5 连重试（最坏 170s）、全败才读缓存"，浏览器/用户早就放弃了。数据层直接调函数完全正常（148 条），盲区在 web 层 `/search`。
- **修复**：低价值高延迟数据（代码→orgId 映射，几乎不变）用**缓存立即可用 + 后台静默刷新**（stale-while-revalidate）：本地有副本就零等待返回，副本超过 1 小时由后台线程刷新写回；只有全新机器（无副本）才同步联网（快败 2 次 × 15s）。另加 htmx `hx-indicator`"查询中"提示。
- **教训**：
  1. "缓存兜底""缓存优先（7天死数据）""缓存+后台刷新"是三种设计；接口不稳定时最终形态是第三种——既要零等待，又不能让数据变成 7 天旧账。
  2. 回归断言要覆盖**完整用户请求路径**（HTTP 端点级），函数级单测全绿≠端点能用——本次 A2（RunPython 下载）PASS 但 `/search` 卡死，盲区正是没测的端点。回归已补 A1c：`/search?stock_code=...` 断言 HTTP 200 + 表格行数≥2。
  3. 端点级复现方法：面板目录里 `uvicorn.Config(app, port=8899)` 起线程 + `urllib.request.urlopen` 打真请求，超时即可复现"用户空白"。
  4. 缓存类设计要**分场景实测**：新鲜副本（应秒回）、过期副本（应秒回+后台刷新写回）、无副本全新机器（应联网拉取）——三个场景路径完全不同，只测一个不算验证。

## 坑 9：全量复制式运行时的体积炸弹

- **症状**：分发 zip 670 MB，且因 jupyter labextensions 超长文件名导致 Windows 解压失败。
- **根因**：`install_from_existing_python` 全量复制系统 Python（含 anaconda/jupyter/PyInstaller 等与项目无关的巨型包）。
- **修复**：复制后清空源 site-packages，依赖由 requirements 重新安装（或离线复制清单 + import 探测迭代补漏）；`_verify_engine` 防止瘦身把必需件删漏。瘦身后 526 MB（压缩后大幅减小）。
- **教训**：运行时必须"最小化 + 硬门禁验证"，两头都要硬。

## 坑 10：解释器身份被 PATH / PYTHONHOME 劫持（开发环境混杂）

- **症状**：`python --version` 显示非预期解释器（如 IDE 沙箱运行时 3.14.7）；`import xlwings` 报 `SyntaxWarning: 'return' in a 'finally' block`，stderr 指向另一份 stdlib 的 `subprocess.py`；`sys.executable` 与 `python --version` 版本不一致。
- **根因**：PATH 被其他运行时（IDE 沙箱 / 虚拟环境 / WindowsApps）劫持，且存在系统级 `PYTHONHOME`/`PYTHONPATH` 指向另一份 Python——解释器与 stdlib 版本错配。
- **诊断**：`python -c "import sys; print(sys.executable)"` 与 `python --version` 对照；`echo $env:PYTHONHOME` / `echo $env:PYTHONPATH` 检查残留；`where.exe python` 查看 PATH 命中顺序。
- **修复**：开发/构建一律用显式解释器绝对路径（`sys.executable` 确认后写入配置表 `Interpreter_Win` / 构建参数 `-InterpreterWin` / release_tool），禁止依赖 PATH 命令名；环境变量错配时临时清理或修正。
- **升级案例（实测铁证）**：**显式绝对路径并不能免疫 PYTHONHOME 劫持**。加载项链中 `Interpreter_Win` 指向 uv venv 的 `python.exe`（显式路径），但 Excel 进程环境带 `PYTHONHOME=<base解释器>` 时，实际启动 EXE=3.13、PREFIX/BASE 均回落 base、**venv site-packages 不进 sys.path** → 报 `No module named 'xlwings'`。判据：`python -c "import sys, site; print(sys.prefix, sys.base_prefix, site.getsitepackages())"`，`prefix == base_prefix` 即 venv 未生效。**修复用 bat 包装解释器**（`set PYTHONHOME=<base>` + `set PYTHONPATH=<模块目录>` + 调 base 解释器）或 junction 纯 ASCII 路径，完整方案见 `references/04-python-guidance.md` 第五节。
- **教训**：解释器身份以 `sys.executable` 为准，不以 `python` 命令名解析为准；**venv 在 Excel/cmd 启动链中不可靠，加载项分发默认走"bat 包装 base 解释器 + PYTHONPATH 指模块目录"**；技能外发后绝对路径不可写死进规则，规则只保留"显式确认方法论"，具体路径属于项目级配置。

## 坑 11：xlwings CLI 入口是独立脚本，`python -m xlwings` 不可用

- **症状**：`python -m xlwings quickstart myproject` 报 `No module named xlwings.__main__; 'xlwings' is a package and cannot be directly executed`。
- **根因**：xlwings 包没有 `__main__.py`，不支持 `-m` 执行；CLI 入口为安装后 Scripts 目录下的 `xlwings.exe`（`python -c "from xlwings.cli import main; main()"` 传参方式在 PowerShell 下参数透传不可靠，勿用）。
- **诊断**：`<解释器目录>\Scripts\xlwings.exe --version` 确认入口可用。
- **修复**：用显式路径 `Scripts\xlwings.exe quickstart <project> [-addin] [-ribbon]`；解释器路径按坑 10 规则显式确认，勿依赖 PATH。
- **教训**：CLI 工具先 `--help` 核实用法再执行，勿假定 `-m` 可用。

## 坑 12：技能维护——新旧同名文件替换时，批量删除会误删新文件（2026-09-09 references 重构）

- **现象**：重构 references 时先新建 `07-udf-guidance.md`（聚合新内容），随后批量删除旧聚合源，删除清单里也含 `07-udf-guidance.md`（旧文件同名），结果新文件被一并删除。
- **根因**：新旧文档同名（`07-udf-guidance.md` 新旧同名），删除命令按文件名匹配，不区分新旧。
- **修复**：立即用刚写入的新内容重写恢复；教训是执行删除前先 `Get-ChildItem` 核对现存文件名与删除清单的交集。
- **教训**：**"先写新、后删旧"且新旧同名时，删除清单必须排除新文件**——改名策略（新名加 `-new` 后缀再 rename）可彻底规避。

## 坑 13：pywebview + uvicorn 必须用 subprocess，threading 模式下所有 HTTP 请求静默超时（2026-09-09 面板实战）

- **现象**：面板窗口正常打开，`wait_for_server` 也通过（webview.start() 前 GET / 正常），但 webview 启动后所有 HTTP 请求（GET /、POST /api/query、POST /api/download）全部超时，无任何响应。
- **根因**：pywebview 的 `webview.start()` 在主线程运行 GUI 消息循环，阻塞 Python GIL；若 uvicorn 用 `threading.Thread(daemon=True)` 启动，webview 启动后 daemon 线程无法获得 GIL，所有请求挂起。`wait_for_server` 在 webview.start() 之前调用，此时 GIL 未被阻塞，故能通过。
- **诊断**：对比"webview.start() 前/后"的 GET / 响应；前正常后超时 → 锁定 GIL 阻塞；直接运行 `server.py`（无 webview）一切正常 → 排除 uvicorn/FastHTML 本身问题。
- **修复**：uvicorn 必须在独立 subprocess 中运行（`subprocess.Popen([sys.executable, "server.py", "--port", ...])`），与 webview 进程隔离；`server.py` 自行 `from src.app import app` 并 `uvicorn.run(app, ...)`；FastHTML app 对象不可 pickle，不能通过 multiprocessing 传递。
- **教训**：**GUI 消息循环与 ASGI 服务必须进程隔离**。任何"主线程跑 GUI、子线程跑 HTTP 服务"的架构在 CPython GIL 下都不可靠；面板侧模板必须默认 subprocess 模式。

## 坑 14：XLSTART 目录下的 `.conf` 文件被 Excel 当工作簿打开（2026-09-17 Hermes 实战）

- **现象**：安装脚本往 `XLSTART\` 写了 `xlwings.conf`（目录级配置）后，每次启动 Excel 多出一个**未隐藏的 "xlwings" 工作表**（A1 显示 INTERPRETER_WIN 等键值）；Ribbon 点击报错。删除该 conf 后工作表消失。
- **根因**：`XLSTART\` 目录下任何文件都会在 Excel 启动时被当作工作簿/加载项处理——目录级 `xlwings.conf` 不是"配置"，而是被打开的文本工作簿。
- **修复**：**XLSTART 目录永远不要放 `.conf` 文件**。配置只写 xlam 内嵌配置表（构建期写入）+ 安装脚本动态改配置表（如 PYTHONPATH），或写用户级/项目级 `xlwings.conf`（坑 2 场景）。安装脚本模板 `templates/internal-bases/excel-addin-core/install/install_xlstart.ps1.tmpl` 的做法正确（只改配置表、不写 conf）。
- **清理**：`Remove-Item "$env:APPDATA\Microsoft\Excel\XLSTART\*.conf" -Force`（仅当确认无用户自建 conf 时；更稳妥是只删自己写入的那个文件）。
- **教训**：**安装脚本写入任何 XLSTART 文件前，先问"Excel 启动时会怎么处理它"**——只允许 .xlam/.xll（加载项）与 .xlsm（模板宏）类文件。

## 坑 15：Excel Resiliency（韧性机制）禁用加载项——反复强杀 Excel 的代价（2026-09-17 Hermes 实战）

- **现象**：多次 `taskkill /F` 强杀 Excel（构建/测试循环）后，加载项在启动时"消失"——Ribbon 无选项卡、加载项管理器不显示；但文件与注册都正常。
- **根因**：Excel 检测到加载项导致反复崩溃/被强杀，会把加载项记入**禁用列表**（`HKCU\Software\Microsoft\Office\<ver>\Excel\Resiliency\DisabledItems` / `DisableItems`），后续启动自动跳过该加载项。
- **诊断**：
  ```bat
  reg query "HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency" /s
  ```
  存在 DisabledItems 且含目标加载项名 → 命中本坑。
- **修复**：删除 DisabledItems 中对应条目（先备份该项）；随后**避免用 `taskkill /F /IM EXCEL.EXE` 裸杀**——按 PID 精确终止（见 SKILL.md 3.4），并确保自动化用 `app.quit()` 正常退出。
- **教训**：**加载项开发循环中，进程终止必须按 PID、正常退出优先**；裸杀不仅误伤用户工作簿，还会触发 Resiliency 让加载项"神秘消失"，排查成本极高。

## 真机 E2E 验证脚本（复用）

`../dist/_e2e_test.py`：隐藏 Excel 打开产物 → 写输入 → `xl.Run("...xlwings.RunPython", "import ...;...")` → 断言 B3 状态 / A5 表头 / A6 数据 → 清空再断言。21s 全 PASS（下载 10 条真实公告）。

## 诊断方法论（可复用）

1. **模态框是自动化的头号杀手**：隐藏实例里任何 Err.Raise 都会挂死整条链。用 `DisplayAlerts=False` + 子进程隔离 + 超时 taskkill 保命（**仅模态框挂死、无法正常退出时的应急手段**；日常清理一律按 PID 精确终止，见 SKILL.md 3.4）。
2. **看不见就让它现形**：可见 Excel 复现 → Pillow 截屏 → win32 EnumWindows 读模态框文本，一步拿到真实错误（比猜快 10 倍）。
3. **分阶段二分**：每步保存→重开→Ping，用最小可运行探针（PingTest 函数）判断工程是否还能编译执行。
4. **手动复现子进程命令**：从 bas 源码拼出 VBA 将执行的确切命令行，手动跑拿 stderr——注意补全 `--wb/--from_xl/--hwnd` 参数（`Book.caller()` 依赖它们）。
