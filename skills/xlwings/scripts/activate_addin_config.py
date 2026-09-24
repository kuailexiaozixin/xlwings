#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置表激活一键脚本：将 xlam 的 _myaddin.conf 激活为 myaddin.conf，
写入 Interpreter_Win / PYTHONPATH / UDF Modules 等配置。

用法:
    python activate_addin_config.py --xlam-path path/to/myaddin.xlam
    python activate_addin_config.py --xlam-path path/to/myaddin.xlam --interpreter C:\\Python313\\python.exe
    python activate_addin_config.py --xlam-path path/to/myaddin.xlam --pythonpath C:\\myproject --udf-modules mymodule

注意:
    - 本脚本经 Excel COM 保存产物，若产物含 2010 版 customUI14.xml 且 Content_Types 有 Override，
      保存后须重新执行 zip 级 Ribbon 回注（见 SKILL.md 6.5.5 关键顺序约束）。
    - 依赖 xlwings 和 Excel COM；AccessVBOM 必须为 True。
"""
import argparse
import os
import sys


def activate(xlam_path, interpreter=None, pythonpath=None, udf_modules=None):
    """激活配置表并写入配置。返回 (success, message)。"""
    if not os.path.exists(xlam_path):
        return False, f"文件不存在: {xlam_path}"

    try:
        import xlwings as xw
    except ImportError:
        return False, "无法导入 xlwings，请先安装"

    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(os.path.abspath(xlam_path))

        # 1. IsAddin=False 使配置表可见可编辑
        wb.api.IsAddin = False

        # 2. 改配置表名（_myaddin.conf → myaddin.conf）
        sheet = wb.sheets[0]
        old_name = sheet.name
        if old_name.startswith("_"):
            new_name = old_name.lstrip("_")
            sheet.name = new_name
            renamed = f"{old_name} → {new_name}"
        else:
            renamed = f"已激活: {old_name}"

        # 3. 写配置（A 列键名已存在，写 B 列值）
        config = {}
        if interpreter:
            config["Interpreter_Win"] = interpreter
        if pythonpath:
            config["PYTHONPATH"] = pythonpath
        if udf_modules:
            config["UDF Modules"] = udf_modules

        written = []
        for row in range(1, 20):
            key = sheet.range(f"A{row}").value
            if key is None:
                break
            if key in config:
                sheet.range(f"B{row}").value = config[key]
                written.append(key)

        # 4. IsAddin=True 并保存
        wb.api.IsAddin = True
        wb.save()
        wb.close()

        msg = f"配置表激活成功: {renamed}; 写入: {written if written else '无（仅改名）'}"
        return True, msg

    except Exception as e:
        return False, f"激活失败: {e}"
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(description="xlam 配置表激活一键脚本")
    parser.add_argument("--xlam-path", required=True, help="产物 .xlam 绝对路径")
    parser.add_argument("--interpreter", default=None, help="Interpreter_Win 值（Python 解释器路径）")
    parser.add_argument("--pythonpath", default=None, help="PYTHONPATH 值（项目根目录）")
    parser.add_argument("--udf-modules", default=None, help="UDF Modules 值（逗号分隔）")
    args = parser.parse_args()

    success, msg = activate(
        args.xlam_path,
        interpreter=args.interpreter,
        pythonpath=args.pythonpath,
        udf_modules=args.udf_modules,
    )
    print(msg)
    if not success:
        sys.exit(1)
    print("✅ 配置表激活完成")
    print("⚠️  若产物含 2010 版 customUI14.xml，请在激活后重新执行 zip 级 Ribbon 回注（见 SKILL.md 6.5.5）")


if __name__ == "__main__":
    main()
