#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 F：端到端 RunPython 加载链路验证
- 模拟 xlwings 引擎从 xlam 所在目录（+ 同名 ZIP，若使用）加载模块的实际过程
- 这是唯一能发现"模块导入路径错误""src 包依赖缺失""sys.modules 缓存冲突"的门禁
- ZIP 分发可选：仅用 PYTHONPATH 分发（无 ZIP）时，本门禁自动跳过 ZIP 相关检查，只验证 xlam + 模块导入

用法:
    python gate_f_runpython_chain.py --xlstart C:/path/to/XLSTART \
        --xlam yourproject.xlam [--zip yourproject.zip] --module yourproject \
        [--module-dir D:/path/to/module]
    # 说明: PYTHONPATH 分发时 Python 模块目录与 XLSTART 不同，用 --module-dir 单独指定
"""
import argparse
import os
import sys
import zipfile


def main():
    parser = argparse.ArgumentParser(description='门禁 F：端到端 RunPython 加载链路验证')
    parser.add_argument('--xlstart', required=True, help='XLSTART 目录（xlam 所在）')
    parser.add_argument('--xlam', required=True, help='xlam 文件名')
    parser.add_argument('--zip', default='', help='同名 ZIP 文件名（可选，仅 ZIP 分发时提供；与 xlam 同目录）')
    parser.add_argument('--module', required=True, help='Python 模块名')
    parser.add_argument('--module-dir', default='', help='Python 模块所在目录（PYTHONPATH 分发时与 XLSTART 不同；默认等于 XLSTART）')
    args = parser.parse_args()

    errors = []

    xlam_path = os.path.join(args.xlstart, args.xlam)
    zip_path = os.path.join(args.xlstart, args.zip) if args.zip else ''
    module_dir = args.module_dir if args.module_dir else args.xlstart

    # F1. 文件存在性
    if not os.path.exists(xlam_path):
        errors.append(f'xlam 缺失: {xlam_path}')
    if args.zip and not os.path.exists(zip_path):
        errors.append(f'ZIP 缺失: {zip_path}')

    if errors:
        print(f'门禁 F 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    # F2. ZIP 内容检查（确认无 src./../ 等可能无法从 ZIP 解析的相对导入）
    if zip_path and os.path.exists(zip_path):
        with zipfile.ZipFile(zip_path, 'r') as z:
            for name in z.namelist():
                if name.endswith('.py'):
                    content = z.read(name).decode('utf-8', errors='replace')
                    for line in content.split('\n'):
                        stripped = line.strip()
                        if stripped.startswith('from ') or stripped.startswith('import '):
                            if 'src.' in stripped or '../' in stripped:
                                errors.append(f'{name} 中包含可能无法解析的导入: {stripped}')

    # F3. 模拟 xlwings 加载路径（与引擎实际行为一致）
    #     模块目录（PYTHONPATH 分发时为独立目录）放在最前，确保能导入到对的包
    sys.path.insert(0, module_dir)
    if zip_path and os.path.exists(zip_path):
        sys.path.insert(0, zip_path)

    # F4. 清除模块缓存（模拟新 Python 进程）
    for key in list(sys.modules.keys()):
        if args.module in key:
            del sys.modules[key]

    # F5. 导入模块并验证来源
    try:
        mod = __import__(args.module)
        module_file = getattr(mod, '__file__', '')
        if zip_path and os.path.exists(zip_path):
            if args.zip not in module_file:
                errors.append(f'模块未从 ZIP 加载，来源: {module_file}')
            else:
                print(f'  模块从 ZIP 加载: {module_file}')
        else:
            print(f'  模块从 PYTHONPATH({module_dir}) 加载: {module_file}')
    except ImportError as exc:
        errors.append(f'模块导入失败: {exc}')
    except Exception as exc:  # noqa: BLE001
        errors.append(f'模块验证异常: {exc}')

    if errors:
        print(f'门禁 F 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('门禁 F 通过：模块可正常导入（来源与分发方式一致）')


if __name__ == '__main__':
    main()
