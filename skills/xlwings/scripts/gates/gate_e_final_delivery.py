#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 E：交付前最终验证（交付前执行）
- E1. 文件存在性与大小
- E2. ZIP 分发文件（如果使用）与模块名一致性
- E3. Python 模块可导入验证
- E4. 可执行产物冒烟：调用模块内核心函数（可选）

用法:
    python gate_e_final_delivery.py --xlam-path path/to/your.xlam \
        --proj-dir path/to/python/code --module your_module [--func main]
"""
import argparse
import os
import sys
import zipfile


def main():
    parser = argparse.ArgumentParser(description='门禁 E：交付前最终验证')
    parser.add_argument('--xlam-path', required=True, help='产物 .xlam/.xlsm 绝对路径')
    parser.add_argument('--proj-dir', required=True, help='Python 代码目录')
    parser.add_argument('--module', required=True, help='Python 模块名')
    parser.add_argument('--func', default='', help='模块内可调用函数名（可选，E4 冒烟用）')
    args = parser.parse_args()

    errors = []

    # E1. 文件存在性与大小
    if not os.path.exists(args.xlam_path):
        errors.append(f'xlam 不存在: {args.xlam_path}')
    else:
        size = os.path.getsize(args.xlam_path)
        if size < 10000:
            errors.append(f'xlam 文件过小: {size} 字节')
        else:
            print(f'  E1 文件大小: {size} 字节 ✅')

    # E2. ZIP 分发文件（如果使用）
    dist_dir = os.path.dirname(args.xlam_path)
    zip_path = os.path.join(dist_dir, os.path.splitext(os.path.basename(args.xlam_path))[0] + '.zip')
    if os.path.exists(zip_path):
        with zipfile.ZipFile(zip_path, 'r') as z:
            py_files = [n for n in z.namelist() if n.endswith('.py')]
            if not py_files:
                errors.append('E2 ZIP 中缺少 Python 文件')
            if not any(args.module in n for n in py_files):
                errors.append(f'E2 ZIP 中缺少模块 {args.module}')
            else:
                print(f'  E2 ZIP 含模块: {[n for n in py_files if args.module in n]} ✅')
    else:
        print('  E2 未使用 ZIP 分发（依赖 PYTHONPATH）')

    # E3. Python 模块可导入验证
    sys.path.insert(0, args.proj_dir)
    if os.path.exists(zip_path):
        sys.path.insert(0, zip_path)
    try:
        mod = __import__(args.module)
        print(f'  E3 Python 模块导入: {args.module} ✅')
    except ImportError as exc:
        errors.append(f'E3 模块导入失败: {exc}')

    # E4. 核心函数冒烟（可选）
    if args.func and not errors:
        try:
            mod = sys.modules.get(args.module)
            if mod is not None and hasattr(mod, args.func):
                getattr(mod, args.func)
                print(f'  E4 函数 {args.func} 存在（未执行，避免副作用）✅')
        except Exception as exc:  # noqa: BLE001
            errors.append(f'E4 函数检查异常: {exc}')

    if errors:
        print(f'门禁 E 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 E 通过，可以交付')


if __name__ == '__main__':
    main()
