"""
embedded_code_extractor.py — 自研 code embed 运行时提取逻辑

复刻自 xlwings PRO 的 pro/embedded_code.py，移除许可证检查，
实现从工作簿 .py 隐藏 sheet 提取代码到临时目录并执行。

与官方版本的区别：
- 移除 LicenseHandler.validate_license("pro") 调用
- get_embedded_code_temp_dir() 自行实现（不依赖 pro/utils.py）
- 其余逻辑与官方完全一致

使用方式（VBA 端修改后调用）：
    import embedded_code_extractor
    embedded_code_extractor.runpython_embedded_code("import mymodule; mymodule.func()")
"""

import json
import os
import sys
import tempfile
from functools import lru_cache
from pathlib import Path

import xlwings as xw
from xlwings.utils import read_config_sheet


def get_embedded_code_temp_dir():
    """创建并返回嵌入代码的临时目录。

    官方版本从 pro/utils.py 导入，此处自行实现。
    每次调用创建新目录（与官方行为一致，官方用 tempfile.mkdtemp）。
    """
    return tempfile.mkdtemp(prefix="xlwings_embedded_")


@lru_cache()
def dump_embedded_code(book, target_dir):
    """从工作簿的 .py sheet 提取代码到 target_dir，并加入 sys.path。

    Args:
        book: xlwings Book 对象
        target_dir: 目标目录路径

    逻辑与官方 pro/embedded_code.py 完全一致：
    1. 读取配置表 RELEASE_EMBED_CODE_MAP（sheet名 → 相对路径的 JSON 映射）
    2. 遍历所有 .py 结尾的 sheet，读取 A 列内容
    3. 按映射关系写入 target_dir（无子目录映射时直接用 sheet 名）
    4. sys.path[0:0] = [target_dir] 插入模块搜索路径最前
    """
    code_map = read_config_sheet(book).get("RELEASE_EMBED_CODE_MAP", "{}")
    sheetname_to_path = json.loads(code_map)

    for sheet in book.sheets:
        if sheet.name.endswith(".py"):
            last_cell = sheet.used_range.last_cell
            sheet_content = (
                sheet.range((1, 1), (last_cell.row, 1)).options(ndim=1).value
            )
            if sheetname_to_path:
                (Path(target_dir) / sheetname_to_path[sheet.name]).parent.mkdir(
                    parents=True,
                    exist_ok=True,
                )
            out_path = os.path.join(
                target_dir,
                sheetname_to_path[sheet.name] if sheetname_to_path else sheet.name,
            )
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                for row in sheet_content:
                    if row is None:
                        f.write("\n")
                    else:
                        f.write(row + "\n")

    sys.path[0:0] = [target_dir]


def runpython_embedded_code(command):
    """提取嵌入代码并执行 command。

    与官方 runpython_embedded_code 签名一致，供 VBA 端调用。

    Args:
        command: 要执行的 Python 代码字符串，如 "import mymodule; mymodule.func()"
    """
    dump_embedded_code(xw.Book.caller(), get_embedded_code_temp_dir())
    exec(command)
