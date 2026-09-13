"""
xlwings release 复刻工具 - 运行时嵌入式代码模块（Python 侧）

替代官方 xlwings.pro.embedded_code.runpython_embedded_code，
无 PRO 许可校验，供复刻版 VBA 调用。

用法（VBA 中）：
    RunPython "import release_tool.runtime_embedded;runtime_embedded.run('import mymodule;mymodule.main()')"

或 Python 侧直接：
    from release_tool import runtime_embedded
    runtime_embedded.run("import mymodule;mymodule.main()")
"""
import json
import os
import sys
import tempfile
import shutil
import atexit
import time
import glob
from functools import lru_cache
from pathlib import Path

# 临时目录根：%TEMP%\xlwings_clone\ （与官方 %TEMP%\xlwings 区分，避免冲突）
TEMP_BASE = os.path.join(tempfile.gettempdir(), "xlwings_clone")


@lru_cache()
def get_embedded_code_temp_dir():
    """创建并返回嵌入式代码解出目录（进程内缓存）"""
    os.makedirs(TEMP_BASE, exist_ok=True)
    # 清理超过 30 天的残留目录
    try:
        for subdir in glob.glob(TEMP_BASE + "/*/"):
            if os.path.getmtime(subdir) < time.time() - 30 * 86400:
                shutil.rmtree(subdir, ignore_errors=True)
    except Exception:
        pass
    tempdir = tempfile.mkdtemp(dir=TEMP_BASE)
    # RunPython 子进程退出时自动清理
    atexit.register(shutil.rmtree, tempdir, ignore_errors=True)
    return tempdir


def dump_embedded_code(book, target_dir):
    """把工作簿中所有 .py 结尾的 sheet 解出到 target_dir

    参数：
        book: xlwings Book 对象（工作簿）
        target_dir: 目标目录
    """
    # 读 RELEASE_EMBED_CODE_MAP（sheetname -> 相对路径）
    code_map = {}
    try:
        if "xlwings.conf" in [s.name for s in book.sheets]:
            config = read_config_sheet(book)
            raw = config.get("RELEASE_EMBED_CODE_MAP", "{}")
            if raw:
                code_map = json.loads(raw)
    except Exception:
        code_map = {}

    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = (
                sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value
            )
            # 目标文件名：优先用 map 里的相对路径，否则直接用 sheet 名
            rel_path = code_map.get(sheet.name, sheet.name)
            target_path = os.path.join(target_dir, rel_path)
            os.makedirs(os.path.dirname(target_path) or target_dir, exist_ok=True)
            with open(target_path, "w", encoding="utf-8", newline="\n") as f:
                if sheet_content is None:
                    continue
                # sheet_content 可能是单个值或列表
                if not isinstance(sheet_content, (list, tuple)):
                    sheet_content = [sheet_content]
                for row in sheet_content:
                    if row is None:
                        f.write("\n")
                    else:
                        f.write(str(row) + "\n")

    # 把 target_dir 插入 sys.path 最前
    sys.path[0:0] = [target_dir]


def read_config_sheet(book):
    """读取 xlwings.conf 配置表为 dict（两列键值）"""
    if "xlwings.conf" not in [s.name for s in book.sheets]:
        return {}
    sheet = book.sheets["xlwings.conf"]
    try:
        data = sheet.used_range.value
    except Exception:
        return {}
    config = {}
    if data is None:
        return config
    if not isinstance(data, (list, tuple)):
        data = [data]
    for row in data:
        if isinstance(row, (list, tuple)) and len(row) >= 2 and row[0] is not None:
            config[str(row[0]).strip().upper()] = row[1]
    return config


def run(command):
    """执行嵌入的 Python 命令（替代官方 runpython_embedded_code）

    流程：
    1. 定位调用方工作簿（xw.Book.caller()）
    2. 解出嵌入式代码到临时目录
    3. exec(command)
    """
    import xlwings as xw

    book = xw.Book.caller()
    dump_embedded_code(book, get_embedded_code_temp_dir())
    exec(compile(command, "<embedded>", "exec"))


if __name__ == "__main__":
    # 自测：无 Excel 时无法调用 Book.caller()，仅验证模块可导入
    print("runtime_embedded 模块导入成功")
    print("TEMP_BASE =", TEMP_BASE)
