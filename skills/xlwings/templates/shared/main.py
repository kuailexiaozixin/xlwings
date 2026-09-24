"""main.py — 入口文件（pywebview + uvicorn）
生成后通常不需要修改。

功能特性：
  - wait_for_server：等待 HTTP 服务就绪再打开窗口（避免白屏）
  - signal/SIGINT 优雅退出：关闭窗口或 Ctrl+C 自动清理
  - 双包结构支持：自动检测扁平 src/app.py 或嵌套 src/<pkg>/app.py
  - 打包路径适配：区分只读资源目录(_MEIPASS) 与可写数据目录(EXE 同级)
  - 自动端口探测：find_free_port 避免端口冲突

来源：templates/shared/main.py（与 templates/.../main.py.tmpl 同步）
"""
import sys, os, uvicorn, webview, threading, socket, signal
from pathlib import Path

if getattr(sys, 'frozen', False):
    # 只读资源基目录：--add-data 解包到 _MEIPASS（onefile 临时目录，不可写）
    RESOURCE_DIR = Path(sys._MEIPASS)
    # 可写数据基目录：EXE 同级目录（onefile 下 _MEIPASS 为只读临时目录，绝不能写数据到此处）
    DATA_DIR = Path(sys.executable).parent
    # 业务代码在资源基目录的 src/ 下（--add-data "src;src" → _MEIPASS/src）
    _src = str(RESOURCE_DIR / "src")
else:
    RESOURCE_DIR = Path(__file__).parent
    DATA_DIR = Path(__file__).parent
    # 开发态：main.py 即位于 src/ 下，src 本身即为包根（避免拼出 src/src 导致导入失败）
    _src = str(Path(__file__).parent)

if _src not in sys.path:
    sys.path.insert(0, _src)


def wait_for_server(url: str, timeout: int = 30) -> bool:
    """等待 HTTP 服务就绪（应用层响应才算真正就绪）"""
    import requests
    for _ in range(timeout):
        try:
            requests.get(url, timeout=1)
            return True
        except requests.RequestException:
            import time
            time.sleep(1)
    return False


def find_free_port(start: int = 5001, end: int = 6000) -> int:
    """自动探测空闲端口"""
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    raise RuntimeError("未找到可用端口")


def start() -> None:
    """供 pyproject.toml [project.scripts] 注册的入口函数"""
    PORT = int(os.environ.get("PORT", 0)) or find_free_port()
    print(f"[OK] 服务启动：http://127.0.0.1:{PORT}")
    print("[INFO] 关闭窗口或按 Ctrl+C 退出")

    # 优雅退出：捕获 Ctrl+C 和窗口关闭信号
    _server_instance = None
    def _cleanup(*args):
        if _server_instance:
            _server_instance.should_exit = True
        print("[INFO] 正在退出...")
        os._exit(0)
    signal.signal(signal.SIGINT, _cleanup)

    # 支持两种包结构：扁平 src/app.py 或嵌套 src/<pkg>/app.py
    try:
        from app import app
    except ImportError:
        import importlib
        # 查找包含 app.py 的子包（而非取第一个字母序子目录）
        _src_dir = Path(_src)
        _pkg = None
        for _p in os.listdir(_src_dir):
            _pkg_path = _src_dir / _p
            if _pkg_path.is_dir() and not _p.startswith("_"):
                if (_pkg_path / "app.py").exists():
                    _pkg = _p
                    break
        if _pkg:
            _mod = importlib.import_module(f"{_pkg}.app")
            app = getattr(_mod, "app")
        else:
            raise ImportError("未找到 app 模块：检查 src/ 下是否存在 app.py 或 <pkg>/app.py")

    # 启动 uvicorn 服务
    _config = uvicorn.Config(app, host="127.0.0.1", port=PORT, reload=False)
    _server_instance = uvicorn.Server(_config)

    # 仅服务模式（SERVER_ONLY=1）：不创建桌面窗口，直接以前台方式运行 uvicorn。
    # 用于无显示环境（服务器/CI/冒烟测试）或仅需要 HTTP 服务的场景。
    if os.environ.get("SERVER_ONLY") == "1":
        print("[INFO] SERVER_ONLY 模式：不创建桌面窗口，直接运行 HTTP 服务")
        _server_instance.run()
        return

    threading.Thread(target=_server_instance.run, daemon=True).start()

    # 等待 HTTP 服务就绪再打开窗口（避免白屏）
    if not wait_for_server(f"http://127.0.0.1:{PORT}"):
        print("[WARN] 服务启动超时，仍尝试打开窗口")

    # 打开 pywebview 桌面窗口
    webview.create_window("__APP_TITLE__", f"http://127.0.0.1:{PORT}")
    webview.start()


if __name__ == "__main__":
    start()
