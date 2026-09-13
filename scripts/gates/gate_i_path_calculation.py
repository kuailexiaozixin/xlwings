#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 I：路径计算验证（ZIP 分发时执行）
- 当模块从 ZIP 加载时，__file__ 是 zip_path/module.py，
  os.path.dirname(__file__) 是 ZIP 所在目录，不是项目目录。
- 验证所有涉及路径计算的变量不指向 XLSTART / 非预期目录。

用法:
    python gate_i_path_calculation.py --zip path/to/your.zip --module your_module
"""
import argparse
import os
import sys


def main():
    parser = argparse.ArgumentParser(description='门禁 I：路径计算验证')
    parser.add_argument('--zip', required=True, help='模块 ZIP 包绝对路径')
    parser.add_argument('--module', required=True, help='模块名（将被 import）')
    parser.add_argument('--forbidden-dir', default='XLSTART', help='路径不得指向的目录关键词（默认 XLSTART）')
    args = parser.parse_args()

    errors = []

    zip_path = os.path.abspath(args.zip)
    if not os.path.exists(zip_path):
        errors.append(f'ZIP 不存在: {zip_path}')
    else:
        sys.path.insert(0, zip_path)
        try:
            mod = __import__(args.module)
            module_file = getattr(mod, '__file__', '')
            module_dir = os.path.dirname(os.path.abspath(module_file))
            print(f'  __file__: {module_file}')
            print(f'  dirname : {module_dir}')

            forbidden = args.forbidden_dir.lower()
            if forbidden in module_dir.lower():
                errors.append(
                    f'路径计算错误: 模块目录指向 {forbidden} ({module_dir})，请硬编码项目路径'
                )
        except ImportError as exc:
            errors.append(f'模块导入失败: {exc}')

    if errors:
        print(f'门禁 I 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('✅ 门禁 I 通过：路径计算在 ZIP 上下文中正确')


if __name__ == '__main__':
    main()
