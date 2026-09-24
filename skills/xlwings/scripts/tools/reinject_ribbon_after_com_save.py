#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Ribbon 回注工具：修复"COM 保存后 xlam 丢失 customUI 注册"。

背景（真机铁律，见 references/18）：
    对白标 xlam 做任何 COM 保存（wb.Save()），Excel 会重写 OOXML 包并丢弃
    [Content_Types].xml 中的 customUI Override 注册，导致 Ribbon 消失。
    build_addin.ps1 的 -ReinjectRibbonAfterActivation 在构建流程内兜底；
    本工具用于构建流程之外的"散装 COM 修改"场景（如手工改配置表后保存），一键自愈。

行为（幂等，可重复运行）：
    1. 读取包内已有 customUI 部件（优先 customUI/customUI14.xml，回退 customUI/customUI.xml）；
       两者都不存在时视为"本就不带 Ribbon"，直接通过。
    2. 按 Get-CustomUIPartInfo 的权威语义修复 [Content_Types].xml：
       为实际存在的部件补 Override（ContentType=application/xml），移除另一版本的陈旧 Override。
    3. 按 Get-CustomUIPartInfo 的权威语义修复 _rels/.rels：
       统一为 Id=rIdCustomUI、类型 2007/2006 ui/extensibility（按部件命名空间）、Target=实际部件。
       已一致则不改动。
    4. 其余条目逐字节原样保留。

用法:
    python reinject_ribbon_after_com_save.py --xlam-path path/to/your.xlam
"""
import argparse
import os
import shutil
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPE_2010 = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_TYPE_2007 = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"


def analyze(root_xml: bytes):
    """从 customUI XML 命名空间推导权威部件信息（与 Get-CustomUIPartInfo 一致）。"""
    root = ET.fromstring(root_xml)
    ns_uri = root.tag.split('}')[0].lstrip('{')
    is_2010 = (ns_uri == "http://schemas.microsoft.com/office/2009/07/customui")
    if is_2010:
        return {
            'entry': 'customUI/customUI14.xml',
            'part': '/customUI/customUI14.xml',
            'alternate_part': '/customUI/customUI.xml',
            'rel_type': REL_TYPE_2010,
        }
    return {
        'entry': 'customUI/customUI.xml',
        'part': '/customUI/customUI.xml',
        'alternate_part': '/customUI/customUI14.xml',
        'rel_type': REL_TYPE_2007,
    }


def main():
    parser = argparse.ArgumentParser(description='COM 保存后 Ribbon 回注（幂等）')
    parser.add_argument('--xlam-path', required=True, help='xlam/xlsm 绝对路径')
    args = parser.parse_args()

    if not os.path.exists(args.xlam_path):
        print(f'文件不存在: {args.xlam_path}')
        sys.exit(1)

    with zipfile.ZipFile(args.xlam_path) as z:
        entries = [(i.filename, z.read(i.filename)) for i in z.infolist()]

    names = {n for n, _ in entries}
    if 'customUI/customUI14.xml' in names:
        entry = 'customUI/customUI14.xml'
    elif 'customUI/customUI.xml' in names:
        entry = 'customUI/customUI.xml'
    else:
        print('包内无 customUI 部件（本就不带 Ribbon），无需回注')
        sys.exit(0)

    info = analyze(dict(entries)[entry])
    print(f'customUI 部件: {entry}（{"2010版" if info["rel_type"] == REL_TYPE_2010 else "2007版"}）')
    changed = False

    # --- 修复 [Content_Types].xml ---
    ct_bytes = dict(entries)['[Content_Types].xml']
    ct_root = ET.fromstring(ct_bytes)
    overrides = {o.get('PartName'): o for o in ct_root.findall(f'{{{NS_CT}}}Override')}
    changed_ct = False
    if info['part'] not in overrides:
        ov = ET.SubElement(ct_root, f'{{{NS_CT}}}Override')
        ov.set('PartName', info['part'])
        ov.set('ContentType', 'application/xml')
        changed_ct = True
    else:
        if overrides[info['part']].get('ContentType') != 'application/xml':
            overrides[info['part']].set('ContentType', 'application/xml')
            changed_ct = True
    for stale in [p for p in overrides if p in (info['alternate_part'],) and p != info['part']]:
        ct_root.remove(overrides[stale])
        changed_ct = True
    if changed_ct:
        ET.register_namespace('', NS_CT)
        new_ct = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(ct_root)
        entries = [(n, new_ct if n == '[Content_Types].xml' else d) for n, d in entries]
        changed = True
        print('  [Content_Types].xml: 已补 customUI Override')
    else:
        print('  [Content_Types].xml: 已是正确状态')

    # --- 修复 _rels/.rels ---
    rels_bytes = dict(entries)['_rels/.rels']
    rels_root = ET.fromstring(rels_bytes)
    target_rels = [r for r in rels_root.findall(f'{{{NS_RELS}}}Relationship')
                   if r.get('Id') == 'rIdCustomUI'
                   or r.get('Type') in (REL_TYPE_2010, REL_TYPE_2007)]
    needs_fix = (len(target_rels) != 1
                 or target_rels[0].get('Id') != 'rIdCustomUI'
                 or target_rels[0].get('Type') != info['rel_type']
                 or target_rels[0].get('Target') != info['entry'])
    if needs_fix:
        for r in target_rels:
            rels_root.remove(r)
        rel_el = ET.SubElement(rels_root, f'{{{NS_RELS}}}Relationship')
        rel_el.set('Id', 'rIdCustomUI')
        rel_el.set('Type', info['rel_type'])
        rel_el.set('Target', info['entry'])
        ET.register_namespace('', NS_RELS)
        new_rels = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + ET.tostring(rels_root)
        entries = [(n, new_rels if n == '_rels/.rels' else d) for n, d in entries]
        changed = True
        print('  _rels/.rels: 已按权威语义修复')
    else:
        print('  _rels/.rels: 已是正确状态')

    if not changed:
        print('无需改动（幂等通过）')
        sys.exit(0)

    # --- 写回（临时文件 + 替换，避免写一半损坏） ---
    fd, tmp = tempfile.mkstemp(suffix='.tmp', dir=os.path.dirname(args.xlam_path) or '.')
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
            for n, d in entries:
                z.writestr(n, d)
        try:
            os.replace(tmp, args.xlam_path)
        except PermissionError:
            print('文件被占用（可能正被 Excel 打开），未写回；请关闭占用后重试')
            sys.exit(2)
        print('回注完成')
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


if __name__ == '__main__':
    main()
