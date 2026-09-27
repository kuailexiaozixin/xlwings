# 窗体构建指南（UserForm / pywebview / tkinter）

> 本文件聚合窗体侧全部既有指南（UserForm UX 规范 / 布局与尺寸），并新增 pywebview 桌面面板（xlwings 语境重写）与 tkinter 形态。窗体形态统一以本文档为准。

## 一、三形态选型表（6.5.1 维度③交互判定用）

| 形态 | 技术栈 | 优点 | 限制 | 适用 |
|------|--------|------|------|------|
| **UserForm** | VBA 原生 | 与 Excel 深度集成（宿主内模态）、零额外依赖、随 xlam/xlsm 分发 | UI 能力有限（无现代 Web 组件）、开发体验弱 | 轻量录入/参数面板、VBA 管对象模型的场景 |
| **pywebview 面板** | Python + FastHTML + pywebview（子进程） | 富 UI（HTML/CSS/JS）、可远程数据、Python 生态 | 需便携运行时（目标机无 Python 时自带）、子进程与 Excel 隔离（经 HTTP 回写） | 查询/搜索/报表类富交互、复用参考实现模式 |
| **tkinter** | Python 标准库 | 零依赖（随 Python 自带）、简单 | UI 简陋、同样需 Python 运行时、无 Web 能力 | 极简弹窗/批处理参数、不想引入 pywebview 时 |

**选型铁律**：
- 目标机无 Python（默认分发场景）→ UserForm 优先；要富 UI 则 pywebview + 自带便携运行时
- 需要远程 API/复杂逻辑 → Python 形态（pywebview/tkinter）
- 仅参数录入、VBA 已够 → UserForm，不过度设计

## 二、UserForm 设计规范（原 05）

### 适用边界（先读）
UserForm 适合"普通用户录入/参数交互"；不适合需要现代 UI 或大量远程数据展示的场景（后者走 pywebview）。

### 总原则
- 面向普通用户：**可被看懂 > 花哨**
- 排版目标：分区清晰、宽度克制、提示可见
- **推荐实现方式**：优先用蓝图骨架（`templates/project-blueprints/userform-data-entry-addin/`）实例化，再按业务改控件与业务方法；不要从空白 UserForm 手搓（布局/事件/桥接容易漏）

### 什么时候优先用 UserForm
- 目标机无 Python（默认分发场景）、轻量录入/参数交互、VBA 管对象模型
- 需要富 UI/远程数据/大量列表展示 → 走 pywebview（见第三节）

### 宽度策略
- 普通用户更容易接受的宽度：窗体整体偏窄（单列字段为主）
- 哪些字段适合同行：短输入（代码/数量）可同行，长文本/日期单行

### 提示文本也是交互的一部分
- 每个输入框配提示（Label/ToolTip/占位符）
- 用户扫描顺序默认：从上到下、从左到右

### Excel 场景常见输入项
代码/日期区间/文件路径/下拉选择/多选/密码——按业务定

### 与测试的配合
- 业务逻辑与窗体解耦（业务方法可脱离窗体单测）
- 窗体打开冒烟：Show 不报错即可

### 混合架构：UserForm → Python 桥接（数据过桥）

UserForm 是 VBA 侧，调 Python 走 RunPython——**RunPython 不支持传参/返回值**，数据必须经单元格中转：

```vb
' VBA 侧：写参数到单元格 → RunPython → 读结果
Range("B1").Value = Me.txtCode.Text
RunPython "import myproj; myproj.fetch()"
Me.lblResult.Caption = Range("B2").Value
```

```python
# Python 侧：从 caller 工作簿单元格读参数、写结果
def fetch():
    wb = xw.Book.caller()
    code = wb.sheets[0].range('B1').value
    result = biz_fetch(code)
    wb.sheets[0].range('B2').value = result
```

**数据过桥规则**：一切参数/结果经工作簿单元格中转；**禁止**期望 RunPython 返回值、禁止字符串拼接传参。

### 用户窗体不要做的事
- 不要用 UserForm 做需要网页级展示的内容
- 交付前测试禁忌：不能只测"窗体能开"，要测业务动作（见 09）

### 蓝图组织规范（原 05 精华）

- **为什么要做成蓝图**：UserForm 的布局/事件/桥接模式高度可复用，蓝图把"可被普通用户看懂的窗体 + 可自动化测试的业务层 + 示例工作簿"固化为骨架，避免每次从零
- **合格蓝图必须含**：明确入口（Ribbon 按钮/工作表按钮）、可被普通用户看懂的窗体、可自动化测试的业务层（业务方法与窗体解耦）、示例工作簿
- **推荐蓝图模式**：多字段录入窗体 / 向导式选项窗体 / 批处理参数面板
- **蓝图里显式写**：宽度策略（单列为主、短字段可同行）、提示文本（每个输入框配 Label/ToolTip）
- **对后续 AI 的强约束**：实例化蓝图后只改控件名/业务方法/提示文本，不改蓝图的入口分发结构与业务层解耦模式
- **最低验收**：窗体能打开、业务方法可脱离窗体单测、入口按钮能触发窗体
- **推荐蓝图**：`templates/project-blueprints/userform-data-entry-addin/`（UserForm 数据录入）、`templates/project-blueprints/listobject-workbench-addin/`（ListObject 工作台）
- **参考实现**：蓝图 `src/forms/` 下的 `.frm` 布局 + `.code.vba` 事件代码 + `src/modules/default_module.vba.tmpl` 入口分发

## 三、pywebview 桌面面板（xlwings 语境）

> 面板 = Excel 内按钮拉起 Python 子进程（FastHTML 服务 + pywebview 窗口），HTTP 回写 Excel。完整骨架见本文档第三节。

### 装配规范（实战验证）

- **端口冲突自检**：`find_free_port` 从 8765 起探测，禁用固定端口
- **服务线程 `daemon=True`**：随主进程退出，不残留
- **就绪等待**：`webview.start()` 前等服务 HTTP 就绪
- **一律 `127.0.0.1`**：禁用 `localhost`（IPv6 解析坑）
- **WebView2 Runtime 硬依赖**：Windows 面板默认 Edge WebView2 后端，目标机缺 Runtime 会静默无窗口——检测与安装指引见 `docs/troubleshooting-python.md`
- **子进程拉起**：`subprocess.Popen([sys.executable, os.path.join(PROJ_DIR, 'src', 'main.py')], cwd=PROJ_DIR, creationflags=subprocess.CREATE_NO_WINDOW)`

### 回写 Excel

面板进程无 caller 上下文——写回走 HTTP 路由 → 业务核心写回函数（`xw.books.active` 获取 Excel 实例）；禁止面板线程直接操作 `xw.books`（COM STA 限制）。

### 与 Excel 的生命周期

- 面板由 Ribbon 按钮拉起（`@xw.sub` 回调）
- 关闭 Excel 时面板进程需一并退出（daemon 线程 + 窗口关闭事件）
- 分发：面板随加载项打包，目标机自带便携运行时

### 面板侧完整模板（main.py + server.py + app.py + HTML，可直接套用）

> 以下为最小可运行骨架，按业务改路由与 HTML。入口侧 `open_panel` 模板见本文档第三节 `src/main.py`（面板子进程入口）。

**`src/main.py`（面板子进程入口，接收 --workbook 参数定位目标工作簿）**：

> **铁律 1：sys.path 插入必须在所有 `src` 包 import 之前**——直接运行 `python src/main.py` 时 `sys.path[0]` 是 `src/` 目录，不含项目根，若先 `from src.app import` 会报 `ModuleNotFoundError: No module named 'src'`。
>
> **铁律 2：uvicorn 必须在独立 subprocess 中运行，禁止用 threading**——pywebview 的 `webview.start()` 在主线程运行 GUI 消息循环，会阻塞 Python GIL，若用 threading 启动 uvicorn，webview 启动后所有 HTTP 请求静默超时（GET/POST 均无响应）。须用 subprocess 启动独立的 `server.py` 进程。

```python
import os, sys, socket, subprocess, time, urllib.request, argparse
import webview

# sys.path 插入必须在 src 包 import 之前（直接运行 src/main.py 时 sys.path[0]=src/）
PROJ_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJ_DIR not in sys.path: sys.path.insert(0, PROJ_DIR)

from src.app import set_target_workbook  # noqa: E402

def find_free_port(start=8765, max_tries=50):
    for port in range(start, start + max_tries):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", port)); return port
        except OSError: continue
    raise RuntimeError("无可用端口")

def wait_for_server(url, timeout=10.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try: urllib.request.urlopen(url, timeout=1).status; return True
        except Exception: time.sleep(0.2)
    return False

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", default=None, help="目标工作簿名（由 open_panel 传入，多实例时定位）")
    args = parser.parse_args()
    if args.workbook:
        set_target_workbook(args.workbook)
    port = find_free_port()
    url = f"http://127.0.0.1:{port}"

    # uvicorn 必须在独立 subprocess（webview.start() 阻塞 GIL，threading 模式下所有请求超时）
    server_py = os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.py")
    server_proc = subprocess.Popen(
        [sys.executable, server_py, "--port", str(port), "--host", "127.0.0.1"],
        cwd=PROJ_DIR, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if not wait_for_server(url):
        print("[ERROR] 服务启动超时", file=sys.stderr)
        server_proc.terminate(); return
    webview.create_window("面板标题", url, width=900, height=600)
    webview.start()
    server_proc.terminate()
    try: server_proc.wait(timeout=3)
    except subprocess.TimeoutExpired: server_proc.kill()

if __name__ == "__main__":
    main()
```

**`src/server.py`（独立 uvicorn 服务进程，由 main.py 通过 subprocess 启动）**：

```python
import os, sys, argparse
import uvicorn

PROJ_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJ_DIR not in sys.path: sys.path.insert(0, PROJ_DIR)

from src.app import app

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()
    uvicorn.run(app, host=args.host, port=args.port, log_level="warning")
```

**`src/app.py`（FastHTTP 路由，用传入的工作簿名定位，禁止默认 xw.books.active）**：

> **铁律：FastHTML 路由必须 `async def` + `await req.json()`，禁止 `json.loads(req.body)`**——FastHTML/Starlette 的 `req.body` 是 coroutine 方法（非 bytes），同步函数中 `json.loads(req.body)` 会静默挂起或报类型错误，所有 POST 请求无响应。这是高频坑，修改路由后须立即用 `curl`/`urllib` 验证端点响应。

```python
import json
import xlwings as xw
from fasthtml.common import FastHTML, JSONResponse
from src import biz_module  # 业务核心

app = FastHTML()
_TARGET_WB = None  # 由 main.py set_target_workbook() 设置

def set_target_workbook(name):
    global _TARGET_WB
    _TARGET_WB = name

def _get_book():
    """多实例时用传入的工作簿名定位；未传入时回退 active（仅单实例安全）"""
    if _TARGET_WB:
        for b in xw.books:
            if b.name == _TARGET_WB:
                return b
    return xw.books.active

@app.post("/api/query")
async def api_query(req):
    body = await req.json()  # FastHTML/Starlette 请求体须 await，禁止 json.loads(req.body)
    data, error = biz_module.query(**body)
    return JSONResponse({"data": data, "error": error})

@app.post("/api/fill")
async def api_fill(req):
    body = await req.json()
    wb = _get_book()
    sheet = wb.sheets[0]
    sheet.range("A5").expand("table").clear_contents()
    sheet.range("A5").value = body.get("headers", [])
    sheet.range("A6").value = body.get("records", [])
    return JSONResponse({"ok": True})
```

**基础 HTML（赋值给 `PANEL_HTML` 变量，放在 main.py 顶部）**：

```python
PANEL_HTML = r"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{font-family:system-ui;margin:12px}
input{padding:4px}button{padding:6px 14px;margin-right:6px;cursor:pointer}
table{border-collapse:collapse;width:100%;margin-top:8px}
th,td{border:1px solid #ccc;padding:4px 8px;text-align:left}
th{background:#f0f0f0}#status{margin-top:8px;color:#666}</style></head><body>
<div>输入:<input id="kw"> <button onclick="query()">查询</button></div>
<div><button onclick="fillSheet()">填入表格</button></div>
<table id="result"><thead><tr><th></th><th>列1</th></tr></thead><tbody></tbody></table>
<div id="status">就绪</div>
<script>
let records=[];
async function api(p,b){const r=await fetch(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});return r.json();}
async function query(){
  status.textContent='查询中...';
  const r=await api('/api/query',{kw:kw.value});
  if(r.error){status.textContent=r.error;return;}
  records=r.data;
  const tb=document.querySelector('#result tbody');tb.innerHTML='';
  records.forEach((item,i)=>{
    const tr=document.createElement('tr');
    tr.innerHTML=`<td><input type="radio" name="sel" value="${i}"></td><td>${item.col1}</td>`;
    tb.appendChild(tr);
  });
  status.textContent=`完成：${records.length} 条`;
}
async function fillSheet(){if(!records.length){status.textContent='无数据';return;}
  const r=await api('/api/fill',{records});status.textContent=r.error||'已填入';}
</script></body></html>"""
```

**模板铁律**：端口探测/daemon 线程/就绪等待/127.0.0.1 四项缺一不可；面板回写 Excel 必须走 HTTP 路由→业务核心→`xw.books.active`，禁止面板线程直接操作 `xw.books`。

## 四、tkinter 形态

- **零依赖**：Python 标准库自带，无需 pywebview/WebView2
- **适用**：极简确认框、批处理参数、进度提示
- **限制**：UI 能力弱、无 HTML、跨 DPI 需适配
- **与 Excel 交互**：同样经单元格过桥（tkinter 侧读/写 `xw.books.active` 或 caller）
- 嵌入方式：Ribbon 回调 `subprocess` 或同进程 `tkinter.Tk()` 主循环（同进程需注意与 Excel COM 消息循环共存）

## 五、官方文档导航

- `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/create-a-user-form.md`：UserForm 创建官方步骤
- `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/add-controls-to-a-user-form.md`：控件添加
- `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/initializing-control-properties.md`：控件属性初始化
- `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/displaying-a-custom-dialog-box.md`：对话框显示
- `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/automatically-dismiss-a-message-box.md`：消息框自动关闭

