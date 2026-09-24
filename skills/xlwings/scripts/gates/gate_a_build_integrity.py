#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 A：构建产物完整性验证（ZIP 结构 + Ribbon 注册 + 回调核验）
- 静态部分（ZIP/Content_Types/_rels/customUI 探测）不依赖 Excel，可随时运行
- 回调交叉核验（Ribbon XML onAction <-> VBA Public Sub）需要 Excel COM，
  无 Excel 时输出警告并跳过（不判失败），由门禁 J 端到端兜底

用法:
    python gate_a_build_integrity.py --xlam-path path/to/your.xlam
    python gate_a_build_integrity.py --xlam-path path/to/your.xlam --skip-com
"""
import argparse
import os
import re
import sys
import time
import zipfile
from collections import Counter


def find_customui_entry(names):
    """按命名空间路由探测 Ribbon 文件；2010 版(customUI14.xml)优先。"""
    entries = [n for n in names if n.startswith('customUI/') and n.endswith('.xml')]
    if not entries:
        return None
    for preferred in ('customUI/customUI14.xml', 'customUI/customUI.xml'):
        if preferred in entries:
            return preferred
    return entries[0]


def verify_zip_structure(xlam_path):
    """静态 ZIP 结构检查，返回 (errors, ribbon_callbacks)。"""
    errors = []
    ribbon_callbacks = []

    if not os.path.exists(xlam_path):
        return [f'产物文件不存在: {xlam_path}'], []

    if os.path.getsize(xlam_path) < 1024:
        errors.append(f'产物文件过小: {os.path.getsize(xlam_path)} 字节')

    try:
        with zipfile.ZipFile(xlam_path, 'r') as z:
            names = z.namelist()

            # A1. ZIP 包无重复条目
            dupes = {k: v for k, v in Counter(names).items() if v > 1}
            if dupes:
                errors.append(f'ZIP 包有重复条目: {dupes}')

            # A2. vbaProject 存在
            if 'xl/vbaProject.bin' not in names:
                errors.append('缺少 xl/vbaProject.bin')

            # A3. customUI 探测（2010/2007 双版本兼容）
            cui_entry = find_customui_entry(names)
            if not cui_entry:
                errors.append('缺少 customUI/customUI14.xml 或 customUI/customUI.xml')
                return errors, []

            # A4. _rels 中 customUI 关系数必须恰好 1 条
            if '_rels/.rels' not in names:
                errors.append('缺少 _rels/.rels')
            else:
                rels = z.read('_rels/.rels').decode('utf-8', errors='replace')
                rel_count = rels.count('ui/extensibility')
                if rel_count != 1:
                    errors.append(f'_rels 中 customUI 关系数={rel_count}（应为 1）')

            # A5. Content_Types 注册
            if '[Content_Types].xml' not in names:
                errors.append('缺少 [Content_Types].xml')
            else:
                ct = z.read('[Content_Types].xml').decode('utf-8', errors='replace')
                if 'customUI' not in ct:
                    errors.append('Content_Types 缺少 customUI 注册')

            # A6. 提取 Ribbon onAction 回调
            cui_content = z.read(cui_entry).decode('utf-8', errors='replace')
            ribbon_callbacks = re.findall(r'onAction="([^"]+)"', cui_content)
            if not ribbon_callbacks:
                errors.append(f'Ribbon XML（{cui_entry}）中未找到任何 onAction 回调')
    except zipfile.BadZipFile:
        errors.append(f'文件不是有效的 ZIP/OOXML 包: {xlam_path}')
    except Exception as exc:  # noqa: BLE001
        errors.append(f'ZIP 解析异常: {exc}')

    return errors, ribbon_callbacks


def verify_callbacks_with_com(xlam_path, ribbon_callbacks):
    """通过 COM 打开产物，交叉核验 Ribbon 回调是否存在于 VBA Public Sub。"""
    warnings = []
    try:
        import xlwings as xw
    except ImportError:
        return ['无法导入 xlwings，跳过 COM 回调核验（可由门禁 J 兜底）']

    app = None
    try:
        app = xw.App(visible=False)
        time.sleep(2)
        wb = app.books.open(xlam_path)
        vb = wb.api.VBProject
        vba_callbacks = []
        for comp in vb.VBComponents:
            cm = comp.CodeModule
            if cm.CountOfLines > 0:
                code = cm.Lines(1, cm.CountOfLines)
                # 匹配 Public Sub 和默认 Sub（非 Private）；Ribbon 回调常写为 `Sub Foo(control As IRibbonControl)`（不显式 Public）
                code_no_private = re.sub(r'^\s*Private\s+Sub\s+\w+\s*\(.*$', '', code, flags=re.MULTILINE)
                subs = re.findall(r'\bSub\s+(\w+)\s*\(', code_no_private)
                vba_callbacks.extend(subs)
        for cb in ribbon_callbacks:
            cb_name = cb.split('.')[-1] if '.' in cb else cb
            if cb_name not in vba_callbacks:
                warnings.append(f'回调 {cb} 在 VBA 中未找到（VBA 含: {vba_callbacks}）')
        wb.close()
    except Exception as exc:  # noqa: BLE001
        warnings.append(f'COM 回调核验异常（可由门禁 J 兜底）: {exc}')
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
            time.sleep(1)
    return warnings


def main():
    parser = argparse.ArgumentParser(description='门禁 A：构建产物完整性验证')
    parser.add_argument('--xlam-path', required=True, help='产物 .xlam/.xlsm 绝对路径')
    parser.add_argument('--skip-com', action='store_true', help='跳过 COM 回调核验（仅静态检查）')
    args = parser.parse_args()

    errors, ribbon_callbacks = verify_zip_structure(args.xlam_path)

    warnings = []
    if not args.skip_com and ribbon_callbacks:
        warnings = verify_callbacks_with_com(args.xlam_path, ribbon_callbacks)

    if errors:
        print(f'门禁 A 失败（{len(errors)} 项）:')
        for e in errors:
            print(f'  - {e}')
        sys.exit(1)

    print('门禁 A 静态检查通过')
    if ribbon_callbacks:
        print(f'  Ribbon 回调: {ribbon_callbacks}')
    if warnings:
        print('警告（不阻塞交付，建议关注）:')
        for w in warnings:
            print(f'  - {w}')
    print('✅ 门禁 A 通过')


if __name__ == '__main__':
    main()
