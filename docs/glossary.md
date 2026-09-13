# 术语表

> 面向零基础用户，解释本技能中出现的专业术语。

## A

### APP
Application 的缩写，即"应用程序"。

## C

### CLI
Command-Line Interface，命令行界面。需要用户在终端（黑窗口）中输入命令来操作。

## D

### 桌面应用
在电脑桌面上以窗口形式运行的程序，有菜单、按钮、输入框等界面元素。

## E

### EXE
Windows 可执行文件（`.exe`），双击即可运行，不需要安装 Python。

## F

### FastHTML
Python Web 框架。用 Python 写 HTML 页面，不需要写 JavaScript。

### Fastlite
FastHTML 自带的 SQLite 数据库工具，让数据存取变得简单。

### FT 组件
FastTags 的简称。用 Python 函数生成 HTML 标签的方式，如 `H1("标题")` 生成 `<h1>标题</h1>`。

## H

### HTMX
前端库。让 HTML 标签通过属性实现动态交互（局部刷新），无需写 JavaScript。

## I

### 内联
把 CSS 样式或 JavaScript 代码直接写在 HTML 中，而不是放在外部文件中。

## P

### 打包（Package）
将 Python 源代码、依赖库和资源文件打包成独立的可执行文件。用户拿到后可以直接运行，无需安装 Python。

### 打包术语

#### --add-data
PyInstaller 参数，指定需要包含但不会被自动扫描到的数据文件/目录。
格式为 `来源路径;目标路径`（Windows）或 `来源路径:目标路径`（Linux/macOS）。

```bash
--add-data "src;src"         # 包含 src/ 目录
--add-data "assets;assets"   # 包含 assets/ 目录
```

#### --hidden-import
PyInstaller 参数，显式声明 PyInstaller 静态分析漏扫的模块。
适用于动态导入（`importlib.import_module()`）或通过 `__getattr__` 延迟加载的模块。

```bash
--hidden-import clr
--hidden-import webview.platforms.edgechromium
```

#### --onefile
PyInstaller 打包模式。所有内容（代码、依赖、资源）打包为单个 `.exe` 文件。
启动时自解压到临时目录。

#### --onedir
PyInstaller 打包模式（默认）。生成一个目录（`dist/MyApp/`），内含 EXE 和依赖文件。
启动速度比 --onefile 快，便于调试。

#### _internal 目录
`--onedir` 模式下，依赖库和资源文件存放的子目录（`dist/MyApp/_internal/`）。
打包后 EXE 从该目录加载模块。

#### .spec 文件
PyInstaller 的构建配置文件（`MyApp.spec`），记录打包参数、hook 路径、数据文件等。
可通过 `python -m PyInstaller MyApp.spec` 复用。

#### UPX
可执行文件压缩工具。PyInstaller 可选集成，但**本技能禁用 UPX**（杀毒误报率高、压缩收益有限）。

#### 冒烟测试（Smoke Test）
打包后最基本的验证：确认 EXE 能启动、HTTP 服务可达、窗口句柄存在。

### pywebview
Python 库。把 Web 页面包装成原生桌面窗口（类似浏览器但没有地址栏）。

### PyInstaller
Python 打包工具。把 Python 程序变成独立的 EXE 文件，对方不需要装 Python。

### pip
Python 的包安装工具，用于安装第三方库。

## S

### SQLite
轻量级本地数据库，数据存在一个文件中，无需安装数据库服务器。

### SSE
Server-Sent Events，服务器主动向前端推送数据的技术（如进度更新、实时消息）。

### static_path
FastHTML 中配置静态文件（CSS/JS/图片）目录的参数。桌面打包场景中不推荐使用。

## U

### uv
新一代 Python 包管理器，比 pip 更快。本技能的默认包管理工具。

### uvicorn
Python 的 ASGI 服务器，用于运行 FastHTML 应用。

## V

### venv
Virtual Environment，虚拟环境。隔离项目依赖，避免不同项目依赖冲突。

## W

### WebView2
微软 Edge 浏览器的渲染引擎，pywebview 用它来显示 Web 页面。Windows 10/11 系统自带。
