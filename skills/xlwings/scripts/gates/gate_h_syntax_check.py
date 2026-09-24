#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 H：语法检查门禁（每次修改 .py 后执行）
- IndentationError 是 0 级错误，不允许出现。
- 可传入多个模块文件，全部通过才返回 0。

用法:
    python gate_h_syntax_check.py --module path/to/your_module.py [--module ...]
"""
import argparse
import py_compile
import sys


def main():
    parser = argparse.ArgumentParser(description='门禁 H：语法检查')
    parser.add_argument('--module', action='append', required=True, help='待检查的 .py 文件（可多次传入）')
    args = parser.parse_args()

    errors = []
    for module_path in args.module:
        try:
            py_compile.compile(module_path, doraise=True)
            print(f'  ✅ 语法通过: {module_path}')
        except py_compile.PyCompileError as exc:
            errors.append(f'{module_path}: {exc}')
        except FileNotFoundError:
            errors.append(f'{module_path}: 文件不存在')

    if errors:
        print(f'门禁 H 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 H 通过')


if __name__ == '__main__':
    main()
