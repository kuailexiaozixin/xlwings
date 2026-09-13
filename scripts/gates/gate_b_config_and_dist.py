#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 B：配置表与分发验证
- B1. 配置表激活验证：白标 xlam 的隐藏配置表必须为 _myaddin.conf（激活后为 myaddin.conf，引擎按 PROJECT_NAME.conf 查找）
- B2. Python 模块可导入验证（模拟 PYTHONPATH）
- B3. ZIP 分发文件存在性验证（若使用 ZIP 分发）

用法:
    python gate_b_config_and_dist.py --xlam-path path/to/your.xlam \
        --proj-dir path/to/python/code [--module your_module]
"""
import argparse
import os
import sys
import time


def verify_config_sheet(xlam_path, module=''):
    """通过 COM 打开 xlam，按"键名(A列)=值(B列)"读取配置表并校验 Interpreter_Win/PYTHONPATH。

    注意（2026-09-05 修复）：原实现把 B1（Interpreter_Win）的值同时当作 PYTHONPATH 打印和校验，
    造成"PYTHONPATH=解释器路径"的误导性输出。现改为按键名查找，并新增语义校验：
    PYTHONPATH 必须是目录（不能是 .exe/.py 文件），且目录下必须能找到目标模块。
    """
    import xlwings as xw

    app = None
    try:
        app = xw.App(visible=False)
        time.sleep(2)
        wb = app.books.open(xlam_path)
        wb.api.IsAddin = False
        time.sleep(0.5)

        config_found = False
        for s in wb.sheets:
            # 白标 xlam 配置表：_myaddin.conf（隐藏，未激活）或 myaddin.conf（已激活）
            if s.name in ('_myaddin.conf', 'myaddin.conf'):
                config_found = True
                print(f'  配置表: {s.name}')
                conf = {}
                for r in range(1, 51):
                    key = str(s.range(f'A{r}').value or '').strip()
                    if key:
                        conf[key] = str(s.range(f'B{r}').value or '').strip()
                print(f'  配置键值: {conf}')
                interp = conf.get('Interpreter_Win', '')
                py_path = conf.get('PYTHONPATH', '')
                print(f'  Interpreter_Win: {interp}')
                print(f'  PYTHONPATH: {py_path}')
                if not interp:
                    raise AssertionError('Interpreter_Win 未设置')
                if not py_path:
                    raise AssertionError('PYTHONPATH 未设置')
                if py_path.lower().endswith(('.exe', '.py', '.pyd', '.dll')):
                    raise AssertionError(f'PYTHONPATH 应为模块所在目录，不能是文件: {py_path}')
                if module:
                    # PYTHONPATH 允许分号分隔多个目录；任一目录能找到模块即通过
                    found = False
                    checked = []
                    for d in [p for p in py_path.split(';') if p.strip()]:
                        mod_file = os.path.join(d, module + '.py')
                        mod_pkg = os.path.join(d, module, '__init__.py')
                        checked += [mod_file, mod_pkg]
                        if os.path.exists(mod_file) or os.path.exists(mod_pkg):
                            found = True
                            break
                    if not found:
                        raise AssertionError(
                            f'PYTHONPATH 目录下找不到模块 {module}（检查过: {checked}）')
                break

        if not config_found:
            raise AssertionError('配置表未找到（应为 _myaddin.conf 或 myaddin.conf）')

        wb.api.IsAddin = True
        wb.close()
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
            time.sleep(1)


def main():
    parser = argparse.ArgumentParser(description='门禁 B：配置表与分发验证')
    parser.add_argument('--xlam-path', required=True, help='产物 .xlam 绝对路径')
    parser.add_argument('--proj-dir', required=True, help='Python 代码目录')
    parser.add_argument('--module', default='', help='Python 模块名（B2 导入验证用）')
    args = parser.parse_args()

    errors = []

    # B1. 配置表激活验证
    try:
        verify_config_sheet(args.xlam_path, module=args.module)
    except Exception as exc:  # noqa: BLE001
        errors.append(f'B1 配置表验证失败: {exc}')

    # B2. Python 模块可导入验证
    if args.module:
        sys.path.insert(0, args.proj_dir)
        try:
            __import__(args.module)
            print(f'  Python 模块导入: {args.module} ✅')
        except ImportError as exc:
            errors.append(f'B2 模块导入失败: {exc}')

    # B3. ZIP 分发文件存在性验证
    dist_dir = os.path.dirname(args.xlam_path)
    zip_path = os.path.join(dist_dir, os.path.splitext(os.path.basename(args.xlam_path))[0] + '.zip')
    if os.path.exists(zip_path):
        import zipfile
        with zipfile.ZipFile(zip_path, 'r') as z:
            py_files = [n for n in z.namelist() if n.endswith('.py')]
            print(f'  ZIP 中的 Python 文件: {py_files}')
            if not py_files:
                errors.append('B3 ZIP 包中缺少 Python 文件')
    else:
        print('  B3 未使用 ZIP 分发（依赖 PYTHONPATH 配置）')

    if errors:
        print(f'门禁 B 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 B 通过')


if __name__ == '__main__':
    main()
