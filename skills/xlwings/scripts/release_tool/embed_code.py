"""
xlwings release 复刻工具 - 代码嵌入器

把工作簿所在目录的所有 .py 文件嵌入 Excel sheet（复刻 cli.py::code_embed，
去掉 PRO 依赖）。运行于有 Excel 的环境。

用法：
    python embed_code.py <xlsm路径> [--file 单文件] [--hide]
"""
import json
import sys
import uuid
from pathlib import Path


def code_embed(wb, source_files=None, single_file=None, hide=False):
    """把 Python 文件嵌入工作簿（复刻 cli.py::code_embed）"""
    import xlwings as xw

    # 1. 确定源文件
    if single_file:
        source_files = [Path(single_file)]
        single_file = True
    elif source_files is None:
        # 递归查找 .py 文件，但排除非业务目录
        EXCLUDE_DIRS = {'tests', 'dist', 'dist_release', '__pycache__', '.venv',
                        'runtime', 'site-packages', 'Lib', 'build', '.pytest_cache'}
        all_files = list(Path(wb.fullname).resolve().parent.rglob("*.py"))
        source_files = [
            f for f in all_files
            if not any(part in EXCLUDE_DIRS for part in f.parts)
        ]
    else:
        source_files = [Path(f) for f in source_files]
        single_file = False

    if not source_files:
        print("WARNING: Couldn't find any Python files in the workbook's directory!")
        return

    # 2. 删除已有 .py sheet（多文件模式）
    if not single_file:
        for sheetname in [s.name for s in wb.sheets]:
            if sheetname.endswith(".py"):
                wb.sheets[sheetname].delete()

    # 3. 写入源码
    sheetname_to_path = {}
    for source_file in source_files:
        if not single_file:
            sheetname = uuid.uuid4().hex[:28] + ".py"
            try:
                rel_path = source_file.relative_to(Path(wb.fullname).parent)
            except ValueError:
                # 源文件不在工作簿目录下（如 --embed-file 指定外部文件）：
                # 回退为平铺文件名，保证 dump 时能还原到临时目录根
                rel_path = Path(source_file.name)
            sheetname_to_path[sheetname] = str(rel_path)
        else:
            sheetname = source_file.name

        with open(source_file, "r", encoding="utf-8") as f:
            content = []
            for line in f.read().splitlines():
                # docstring 引号转义
                line = line.replace("'''", '"""')
                # 单引号开头行加 ' 前缀（Excel 字符串转义）
                content.append(["'" + line if line.startswith("'") else line])

        if single_file and source_file.name not in wb.sheet_names:
            try:
                sheet = wb.sheets.add(source_file.name, after=wb.sheets[len(wb.sheets) - 1])
            except Exception:
                # xlam 加载项工作表全隐藏时 Move 无效，回退默认位置
                sheet = wb.sheets.add(source_file.name)
        elif single_file:
            sheet = wb.sheets[source_file.name]
        else:
            try:
                sheet = wb.sheets.add(sheetname, after=wb.sheets[len(wb.sheets) - 1])
            except Exception:
                # xlam 加载项工作表全隐藏时 Move 无效，回退默认位置
                sheet = wb.sheets.add(sheetname)

        sheet["A1"].resize(
            row_size=len(content) if content else 1
        ).number_format = "@"
        sheet["A1"].value = content
        sheet["A:A"].column_width = 65

    # 4. 写 RELEASE_EMBED_CODE_MAP 到 xlwings.conf
    if "xlwings.conf" in wb.sheet_names and not single_file:
        config = read_config_sheet(wb)
        sheetname_to_path_str = json.dumps(sheetname_to_path)
        if len(sheetname_to_path_str) > 32_767:
            raise ValueError(
                "嵌入文件映射超过 Excel 单元格 32767 字符上限，"
                "请减少嵌入文件数量或用 --embed-file 显式指定"
            )
        config["RELEASE_EMBED_CODE_MAP"] = sheetname_to_path_str
        # 用二维数组写入，确保格式正确（A列=键，B列=值）
        config_array = [[k, v] for k, v in config.items()]
        wb.sheets["xlwings.conf"].range("A1").value = config_array
    elif not single_file:
        try:
            config_sheet = wb.sheets.add("xlwings.conf", after=wb.sheets[len(wb.sheets) - 1])
        except Exception:
            config_sheet = wb.sheets.add("xlwings.conf")
        config_sheet["A1"].value = {
            "RELEASE_EMBED_CODE_MAP": json.dumps(sheetname_to_path)
        }

    # 5. 可选隐藏 .py sheet
    if hide:
        for sheet in wb.sheets:
            if sheet.name.endswith(".py"):
                sheet.visible = False

    print(f"[OK] 已嵌入 {len(source_files)} 个文件")
    return sheetname_to_path


def read_config_sheet(wb):
    """读取 xlwings.conf 为 dict"""
    if "xlwings.conf" not in wb.sheet_names:
        return {}
    sheet = wb.sheets["xlwings.conf"]
    data = sheet.used_range.value
    config = {}
    if data is None:
        return config
    if not isinstance(data, (list, tuple)):
        data = [data]
    for row in data:
        if isinstance(row, (list, tuple)) and len(row) >= 2 and row[0] is not None:
            # 关键：保留原始键名大小写，禁止 .upper()
            # xlwings read_config_sheet 大小写敏感，Interpreter_Win 被
            # 转成 INTERPRETER_WIN 后会导致运行时找不到解释器路径
            config[str(row[0]).strip()] = row[1]
    return config


if __name__ == "__main__":
    import xlwings as xw

    if len(sys.argv) < 2:
        print("用法: python embed_code.py <xlsm路径> [--file 单文件] [--hide]")
        sys.exit(1)

    xlsm_path = sys.argv[1]
    single_file = None
    hide = False
    if "--file" in sys.argv:
        single_file = sys.argv[sys.argv.index("--file") + 1]
    if "--hide" in sys.argv:
        hide = True

    app = xw.App(visible=False)
    try:
        wb = app.books.open(xlsm_path)
        code_embed(wb, single_file=single_file, hide=hide)
        wb.save()
        wb.close()
        print(f"[OK] 代码嵌入完成并保存: {xlsm_path}")
    finally:
        app.quit()
