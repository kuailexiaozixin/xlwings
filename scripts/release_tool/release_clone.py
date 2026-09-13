"""
xlwings release 复刻工具 - release 主命令

复刻官方 `xlwings release`，去掉 PRO/deploy key 依赖：
1. 在 xlam/xlsm 中写入/更新 xlwings.conf（Interpreter_Win / RELEASE_*）
2. 嵌入 Python 代码（可选）
3. 导入无加载项 VBA 模块（可选，RELEASE_NO_ADDIN）
4. 隐藏配置表与代码 sheet

用法：
    python release_clone.py <xlam或xlsm路径> --interpreter <python.exe>
        [--proj-dir <源码目录>] [--embed-file <py路径> ...]
        [--embed] [--hide] [--no-addin]

实测关键点（2026-09-08）：
1. 官方 xlwings.bas 嵌入式分支有 LICENSE_KEY 门禁：LICENSE_KEY 为空时
   弹窗 "Embedded code requires a valid LICENSE_KEY." 并退出。本工具在
   --embed 时自动写入 LICENSE_KEY=CLONE（占位值，仅用于通过 VBA 门禁；
   复刻运行时 release_tool.runtime_embedded 不做任何许可校验）。
2. 白标 xlam 转换来的工作簿配置表名为 myaddin.conf，独立运行时读取的是
   xlwings.conf —— 本工具自动把 myaddin.conf 改名为 xlwings.conf。
3. embed 必须显式指定源文件（--embed-file），官方按"工作簿所在目录递归
   找 .py"的策略在 dist 交付目录下会找不到任何文件。
"""
import argparse
import json
import os
import sys
from pathlib import Path

# 定位本工具目录与 xlwings 源码（VBA 模块源）
TOOL_DIR = Path(__file__).resolve().parent
# release_tool 包的可导入性：把 release_tool 的父目录加入 sys.path，
# 这样无论从哪个解释器/工作目录运行都能 `from release_tool.embed_code import ...`
sys.path.insert(0, str(TOOL_DIR.parent))

# 技能根目录（VBA\xlwings\）：TOOL_DIR=scripts/release_tool/ → 上两级
SKILL_ROOT = TOOL_DIR.parent.parent
# xlwings 官方源码包（随技能分发）：用于 Dictionary.cls 等 VBA 模块
XLWINGS_SRC = SKILL_ROOT / "xlwings-0.37.0"
# 安装版 xlwings.bas（含 xlwings.bas 生成产物）：运行时从当前解释器的 site-packages 查找
def _find_installed_xlwings_bas() -> Path | None:
    """从当前解释器的 xlwings 包中定位 xlwings.bas"""
    try:
        import xlwings as _xw
        return Path(_xw.__file__).resolve().parent / "xlwings.bas"
    except Exception:
        return None
INSTALLED_XLWINGS_BAS = _find_installed_xlwings_bas()


def log(msg):
    print(msg, flush=True)


def _import_vba_component(vb, src: Path):
    """把 .bas/.cls 规范化为 CRLF 后再导入 VBA 工程

    实测关键（2026-09-08）：技能源码目录（git 仓库）里的 .cls/.bas 是 LF 换行，
    VBE 的 VBComponents.Import 对 LF-only 文件导入后内容损坏（编译错误），
    之后任何 Application.Run 都会挂死在隐藏实例的模态编译错误框上——
    症状是 release 后宏完全无法运行且无任何报错输出。
    统一先转 CRLF 写入临时文件再导入。
    """
    import tempfile
    raw = src.read_bytes()
    text = raw.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    tmp_dir = Path(tempfile.gettempdir()) / "release_tool_vba"
    tmp_dir.mkdir(exist_ok=True)
    tmp = tmp_dir / src.name
    tmp.write_bytes(text)
    vb.VBComponents.Import(str(tmp))
    return tmp


def normalize_conf_sheet(wb):
    """白标 xlam 配置表 → 独立工作簿配置表：myaddin.conf 改名 xlwings.conf"""
    names = list(wb.sheet_names)
    if "xlwings.conf" in names:
        return
    if "myaddin.conf" in names:
        wb.sheets["myaddin.conf"].name = "xlwings.conf"
        log("  [OK] 配置表改名: myaddin.conf -> xlwings.conf")


def write_config(wb, config):
    """写入 xlwings.conf 配置表（两列 键/值）

    关键：配置表已存在时删除重建，而非 clear()。
    clear() 只清值不格式，A 列若残留"全部大写"格式会把
    Interpreter_Win 写成 INTERPRETER_WIN，导致 xlwings
    read_config_sheet（大小写敏感）找不到键。
    """
    if "xlwings.conf" in wb.sheet_names:
        wb.sheets["xlwings.conf"].delete()
    config_sheet = wb.sheets.add("xlwings.conf")
    # 写入键值对（A 列键、B 列值）
    # 关键：逐个单元格写入，而非二维数组批量写入。
    # 二维数组写入会触发 Excel 自动更正，把 Interpreter_Win 变成 INTERPRETER_WIN，
    # 导致 xlwings read_config_sheet（大小写敏感）找不到键。
    rows = [(k, v) for k, v in config.items() if v is not None]
    for i, (k, v) in enumerate(rows, start=1):
        config_sheet.range((i, 1)).value = k
        config_sheet.range((i, 2)).value = v
    config_sheet["A:A"].autofit()
    return config_sheet


def embed_code(wb, source_files, hide=False):
    """调用 embed_code.py 的逻辑（release_tool 已通过 sys.path 可导入）"""
    from release_tool.embed_code import code_embed

    code_embed(wb, source_files=source_files, hide=hide)


def inject_standalone_vba(wb):
    """注入无加载项 VBA 模块（复刻 release 的 RELEASE_NO_ADDIN 分支）

    1. 移除 xlwings VBA 引用
    2. 删除官方 xlwings 加载项模块
    3. 导入 xlwings.bas（若源码包中有）与 Dictionary.cls
    """
    vb = wb.api.VBProject

    # 移除引用
    try:
        refs = [r.Name for r in vb.References]
        if "xlwings" in refs:
            vb.References.Remove(vb.References("xlwings"))
            log("  [OK] 移除 VBA 引用 xlwings")
    except Exception as e:
        log(f"  [WARN] 移除引用失败: {e}")

    # 删除官方模块
    for mod_name in ["xlwings", "Dictionary", "IWebAuthenticator", "WebClient",
                     "WebRequest", "WebResponse", "WebHelpers"]:
        try:
            comp = vb.VBComponents(mod_name)
            vb.VBComponents.Remove(comp)
            log(f"  [OK] 删除 VBA 模块 {mod_name}")
        except Exception:
            pass

    # 导入独立版 xlwings.bas（若存在；优先安装版 CRLF，其次源码包 LF→CRLF 规范化）
    installed_addin_dir = INSTALLED_XLWINGS_BAS.parent / "addin"
    bas_candidates = [INSTALLED_XLWINGS_BAS,
                      installed_addin_dir / "xlwings.bas" if installed_addin_dir.exists() else None,
                      XLWINGS_SRC / "xlwings" / "xlwings.bas"]
    bas_path = next((p for p in bas_candidates if p and p.exists()), None)
    cls_candidates = [installed_addin_dir / "Dictionary.cls" if installed_addin_dir.exists() else None,
                      XLWINGS_SRC / "xlwings" / "addin" / "Dictionary.cls"]
    cls_path = next((p for p in cls_candidates if p and p.exists()), None)
    if bas_path:
        _import_vba_component(vb, bas_path)
        log(f"  [OK] 导入 {bas_path.name}（CRLF 规范化）")
        # 关键：把嵌入式分支从 PRO 调用替换为复刻运行时
        _patch_embedded_branch(vb, bas_path)
    else:
        log(f"  [WARN] 未找到 xlwings.bas，跳过无加载项注入")
        log("    提示：xlwings.bas 由 scripts/build_excel_files.py 生成，")
        log("    也可手动从安装版 xlwings 包获取")
    if cls_path:
        _import_vba_component(vb, cls_path)
        log(f"  [OK] 导入 {cls_path.name}（CRLF 规范化）")


def _patch_embedded_branch(vb, bas_path):
    """把 xlwings.bas 中嵌入式分支的 PRO 调用替换为复刻运行时

    官方: import xlwings.pro;xlwings.pro.runpython_embedded_code('...')
    复刻: import release_tool.runtime_embedded;release_tool.runtime_embedded.run('...')
    """
    comp = vb.VBComponents("xlwings")
    cm = comp.CodeModule
    code = cm.Lines(1, cm.CountOfLines)
    old = "import xlwings.pro;xlwings.pro.runpython_embedded_code('"
    new = "import release_tool.runtime_embedded;release_tool.runtime_embedded.run('"
    if old in code:
        # 逐行替换（VBA 代码按行操作）
        replaced = False
        for i in range(1, cm.CountOfLines + 1):
            line = cm.Lines(i, 1)
            if old in line:
                cm.ReplaceLine(i, line.replace(old, new))
                replaced = True
        if replaced:
            log("  [OK] 嵌入式调用已替换为复刻运行时（release_tool.runtime_embedded）")
        else:
            log("  [WARN] 找到 PRO 调用但替换失败")
    else:
        log("  [OK] 未发现 PRO 嵌入式调用（无嵌入需求或已替换）")


def main():
    parser = argparse.ArgumentParser(description="xlwings release 复刻")
    parser.add_argument("workbook", help="目标 .xlam 或 .xlsm 路径")
    parser.add_argument("--interpreter", required=True, help="便携 Python 解释器路径")
    parser.add_argument("--proj-dir", default=None, help="Python 源码目录（PYTHONPATH；--embed 未指定源文件时递归查找 .py）")
    parser.add_argument("--embed-file", action="append", default=None,
                        help="要嵌入的 .py 文件路径（可重复；指定后忽略 --proj-dir 递归查找）")
    parser.add_argument("--embed", action="store_true", help="嵌入 Python 代码到 sheet")
    parser.add_argument("--hide", action="store_true", help="隐藏配置表与代码 sheet")
    parser.add_argument("--no-addin", action="store_true", help="注入无加载项 VBA（RELEASE_NO_ADDIN）")
    args = parser.parse_args()

    # 构建前杀残留 Excel 进程（防止目标文件被占用导致 PermissionError）
    target_name = Path(args.workbook).name.lower()
    try:
        import psutil
        for proc in psutil.process_iter(["name", "pid"]):
            if proc.info["name"] and proc.info["name"].lower() in ("excel.exe", "et.exe"):
                try:
                    # 检查该进程是否打开了目标文件
                    for open_file in proc.open_files():
                        if target_name in open_file.path.lower():
                            log(f"[WARN] 终止占用目标文件的进程: PID={proc.info['pid']}")
                            proc.terminate()
                            proc.wait(timeout=5)
                            break
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
                    pass
    except ImportError:
        # psutil 不可用时回退：仅警告，不强制杀进程
        log("[WARN] psutil 未安装，跳过残留进程检测（若构建失败请手动关闭 Excel）")

    import xlwings as xw

    app = xw.App(visible=False)
    try:
        wb = app.books.open(args.workbook)
        active = wb.sheets.active

        # 0. 白标配置表规范化
        normalize_conf_sheet(wb)

        # 1. 写入配置
        config = {
            "Interpreter_Win": args.interpreter,
            # 嵌入模式下 PYTHONPATH 指向开发机路径毫无意义（目标机不存在），不写入
            "PYTHONPATH": None if args.embed else args.proj_dir,
            # 嵌入模式必须写入占位 LICENSE_KEY 以通过 xlwings.bas 的 VBA 门禁
            # （官方分支：LICENSE_KEY 为空 → 弹窗退出，嵌入式命令根本不会被改写）
            "LICENSE_KEY": "CLONE" if args.embed else None,
            # 强制关闭 UDF 服务器模式：用户级全局配置（~/.xlwings/xlwings.conf）可能
            # 写了 USE UDF SERVER=True，那会让 RunPython 走 XLPy COM 服务器分支；
            # 便携运行时通常没有 xlwings32-*.dll → Err.Raise 1000 模态框挂死。
            # 工作簿级配置覆盖用户级配置，这里显式写 False 走标准子进程路径。
            "USE UDF SERVER": False,
            # 同理压制用户级 SHOW CONSOLE=True：否则每次 RunPython 弹 cmd 黑窗，
            # 跑完即关，用户看到"闪退的黑窗"。工作簿级覆盖用户级。
            "SHOW CONSOLE": False,
            "RELEASE_EMBED_CODE": args.embed,
            "RELEASE_HIDE_CONFIG_SHEET": args.hide,
            "RELEASE_HIDE_CODE_SHEETS": args.hide and args.embed,
            "RELEASE_NO_ADDIN": args.no_addin,
        }
        write_config(wb, config)
        log("[OK] 配置表写入完成")
        # xlam 加载项的工作表默认隐藏，wb.sheets.active 可能为 None
        if active is not None:
            try:
                active.activate()
            except Exception:
                pass  # 加载项模式下激活工作表可能失败，忽略

        # 2. 嵌入代码
        if args.embed:
            if args.embed_file:
                source_files = [Path(p) for p in args.embed_file]
                for p in source_files:
                    if not p.exists():
                        raise FileNotFoundError(f"--embed-file 不存在: {p}")
            elif args.proj_dir:
                # 排除标准库/运行时/测试/构建产物目录，避免嵌入几百个文件导致
                # RELEASE_EMBED_CODE_MAP 超 Excel 单元格 32767 字符上限 → sys.exit 跳过 wb.save()
                _EXCLUDE_DIRS = {"Lib", "site-packages", "runtime", "tests",
                                 "dist", "dist_release", ".venv", "__pycache__", "build", ".git"}
                source_files = [
                    p for p in Path(args.proj_dir).rglob("*.py")
                    if not any(part in _EXCLUDE_DIRS for part in p.parts)
                ]
                if len(source_files) > 50:
                    raise ValueError(
                        f"可嵌入 .py 文件过多（{len(source_files)} > 50），"
                        f"请用 --embed-file 显式指定，或检查 --proj-dir 是否指向了含标准库的目录"
                    )
            else:
                raise ValueError("--embed 需要指定 --embed-file 或 --proj-dir")
            if not source_files:
                raise ValueError("未找到任何可嵌入的 .py 文件")
            log(f"  嵌入源文件: {[str(p) for p in source_files]}")
            embed_code(wb, source_files, hide=args.hide)
            log("[OK] 代码嵌入完成")

        # 3. 无加载项注入
        if args.no_addin:
            inject_standalone_vba(wb)
            log("[OK] 无加载项 VBA 注入完成")

        # 4. 隐藏配置表
        if args.hide and "xlwings.conf" in wb.sheet_names:
            # Excel 不允许隐藏全部可见工作表（"不能设置类 Worksheet 的 Visible 属性"）。
            # 隐藏后若无任何可见表（如白标 xlam 转换来的工作簿只有配置表+代码表），
            # 自动补一张"说明"封面表，兼作交付使用说明。
            remaining_visible = [s for s in wb.sheets
                                 if s.visible and s.name != "xlwings.conf"]
            if not remaining_visible and "说明" not in wb.sheet_names:
                usage = wb.sheets.add("说明", after=wb.sheets[len(wb.sheets) - 1])
                usage["A1"].value = "使用说明"
                usage["A2"].value = "本工作簿已内置便携运行时所需的全部配置与代码（免装 Python）。"
                usage["A3"].value = "使用：功能区对应选项卡的按钮触发功能；或按交付说明文档操作。"
                usage["A:A"].autofit()
                log("  [OK] 已补建可见封面表「说明」（避免全部工作表被隐藏）")
            wb.sheets["xlwings.conf"].visible = False
            log("[OK] 配置表已隐藏")

        wb.save()
        wb.close()
        log(f"\n[OK] release 复刻完成: {args.workbook}")
        log("[WARN] 注意: wb.save() 会剥掉 customUI 注册（Ribbon 丢失），")
        log("   请随后执行 reinject_ribbon_after_com_save.py 回注并跑门禁 A。")
    finally:
        app.quit()


if __name__ == "__main__":
    main()
