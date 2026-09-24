"""前端路由链路校验：扫描面板 HTML 中 fetch 的路径，验证后端是否注册对应路由。

防运行时 404：前端引用了但后端未注册的 API 路由。
用法：
    python check_routes_linkage.py --panel-html <panel_html.py> --app <app.py>
退出码：0=全部匹配，1=有未匹配路由（阻断发布）
"""
import argparse
import ast
import re
import sys
from pathlib import Path


def extract_fetch_paths(panel_html_path: Path) -> set[str]:
    """从 panel_html.py 的 PANEL_HTML 字符串中提取 fetch('/api/...') 路径"""
    text = panel_html_path.read_text(encoding="utf-8")
    # 匹配 fetch('/api/xxx') 或 fetch("/api/xxx") 或 fetch(`/api/xxx`)
    paths = set(re.findall(r"""fetch\(['"`](/api/[^'"`]+)['"`]""", text))
    # 也匹配 api('/api/xxx') 封装调用
    paths |= set(re.findall(r"""api\(['"`](/api/[^'"`]+)['"`]""", text))
    return paths


def extract_registered_routes(app_path: Path) -> set[str]:
    """从 app.py 中提取 @app.get/@app.post 注册的路由路径"""
    tree = ast.parse(app_path.read_text(encoding="utf-8"))
    routes = set()
    for node in ast.walk(tree):
        # 同时匹配 def 和 async def
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for dec in node.decorator_list:
                if isinstance(dec, ast.Call) and isinstance(dec.func, ast.Attribute):
                    if dec.func.attr in ("get", "post", "put", "delete", "patch"):
                        if dec.args and isinstance(dec.args[0], ast.Constant):
                            routes.add(dec.args[0].value)
    return routes


def main():
    parser = argparse.ArgumentParser(description="前端路由链路校验")
    parser.add_argument("--panel-html", required=True, help="panel_html.py 路径")
    parser.add_argument("--app", required=True, help="app.py 路径")
    args = parser.parse_args()

    panel_path = Path(args.panel_html)
    app_path = Path(args.app)

    if not panel_path.exists():
        print(f"[FAIL] panel_html 不存在: {panel_path}")
        sys.exit(1)
    if not app_path.exists():
        print(f"[FAIL] app.py 不存在: {app_path}")
        sys.exit(1)

    fetch_paths = extract_fetch_paths(panel_path)
    registered = extract_registered_routes(app_path)

    print(f"前端 fetch 路径 ({len(fetch_paths)}): {sorted(fetch_paths)}")
    print(f"后端注册路由 ({len(registered)}): {sorted(registered)}")

    missing = fetch_paths - registered
    if missing:
        print(f"\n[FAIL] 前端引用但后端未注册的路由 ({len(missing)}):")
        for p in sorted(missing):
            print(f"  - {p}")
        print("\n阻断发布：请在 app.py 中注册上述路由，或修正前端 fetch 路径。")
        sys.exit(1)

    print("\n[OK] 全部前端 fetch 路径均已在后端注册")
    sys.exit(0)


if __name__ == "__main__":
    main()
