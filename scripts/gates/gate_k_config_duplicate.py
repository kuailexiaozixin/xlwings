#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 K：配置表重复键扫描（P0 防护，从源头阻止 VBA 457 报错）
- xlwings 引擎在 RunPython 时按 PROJECT_NAME.conf 读取配置表，遇到重复键会抛 VBA 457。
- 本门禁在交付前静态扫描配置表 A 列，发现重复键即阻断，早于用户打开 Excel。

用法:
    python gate_k_config_duplicate.py --xlam-path path/to/your.xlam
"""
import argparse
import sys
import time


def scan_config_sheet(xlam_path):
    import xlwings as xw

    app = None
    try:
        app = xw.App(visible=False)
        time.sleep(2)
        wb = app.books.open(xlam_path)
        wb.api.IsAddin = False
        time.sleep(0.5)

        target = None
        for s in wb.sheets:
            if s.name in ('_myaddin.conf', 'myaddin.conf'):
                target = s
                break
        if target is None:
            raise AssertionError('配置表未找到（应为 _myaddin.conf 或 myaddin.conf）')

        # 读 A 列前 50 行
        keys_raw = target.range('A1:A50').value
        keys = [str(k).strip() for k in keys_raw if k is not None and str(k).strip() != '']

        print(f'  配置表: {target.name}')
        print(f'  非空键: {keys}')

        # 1. 重复键检测（直接触发 457 的根因）
        seen = set()
        dups = []
        for k in keys:
            if k in seen:
                dups.append(k)
            else:
                seen.add(k)
        if dups:
            raise AssertionError(f'配置表存在重复键（将触发 VBA 457）：{sorted(set(dups))}')

        # 2. 关键键存在性
        if 'PYTHONPATH' not in keys:
            raise AssertionError('配置表缺少 PYTHONPATH 键')

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
    parser = argparse.ArgumentParser(description='门禁 K：配置表重复键扫描')
    parser.add_argument('--xlam-path', required=True, help='产物 .xlam 绝对路径')
    args = parser.parse_args()

    try:
        scan_config_sheet(args.xlam_path)
    except Exception as exc:  # noqa: BLE001
        print(f'门禁 K 失败: {exc}')
        sys.exit(1)

    print('门禁 K 通过：配置表无重复键、PYTHONPATH 存在')


if __name__ == '__main__':
    main()
