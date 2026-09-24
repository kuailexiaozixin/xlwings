#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 G：引擎注入验证（白标 xlam 专用，构建后执行）
- G1. 打开 xlam，确认 VBA 工程中存在 xlwings 模块（引擎占位或已注入）
- G2. 确认配置表名称正确（_myaddin.conf 隐藏态 / myaddin.conf 激活态）

用法:
    python gate_g_engine_injected.py --xlam-path path/to/your.xlam
"""
import argparse
import sys
import time


def main():
    parser = argparse.ArgumentParser(description='门禁 G：引擎注入验证')
    parser.add_argument('--xlam-path', required=True, help='产物 .xlam 绝对路径')
    args = parser.parse_args()

    import xlwings as xw

    errors = []
    app = None
    try:
        app = xw.App(visible=False)
        time.sleep(2)
        wb = app.books.open(args.xlam_path)
        vb = wb.api.VBProject

        # G1. xlwings 引擎模块存在
        component_names = []
        for comp in vb.VBComponents:
            component_names.append(comp.Name)

        engine_found = any('xlwings' in name.lower() for name in component_names)
        if not engine_found:
            errors.append(f'VBA 工程中未找到 xlwings 模块（现有组件: {component_names}）')
        else:
            print(f'  xlwings 模块存在: {[n for n in component_names if "xlwings" in n.lower()]} ✅')

        # G2. 配置表名称
        sheet_names = [s.Name for s in wb.api.Worksheets]
        config_found = any(name in ('_myaddin.conf', 'myaddin.conf') for name in sheet_names)
        if not config_found:
            errors.append(f'配置表未找到（现有工作表: {sheet_names}）')
        else:
            print(f'  配置表存在: {[n for n in sheet_names if n in ("_myaddin.conf", "myaddin.conf")]} ✅')

        wb.close()
    except Exception as exc:  # noqa: BLE001
        errors.append(f'引擎注入验证异常: {exc}')
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
            time.sleep(1)

    if errors:
        print(f'门禁 G 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 G 通过：引擎模块与配置表均已注入')


if __name__ == '__main__':
    main()
