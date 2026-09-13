"""
xlwings release 复刻工具 - xlam 转 xlsm
将 xlam 加载项转换为 xlsm 宏工作簿（步骤 11.1）
"""
import zipfile
import os
from pathlib import Path


def xlam_to_xlsm(xlam_path, xlsm_path):
    """将 .xlam 转换为 .xlsm"""
    xlam_path = Path(xlam_path)
    xlsm_path = Path(xlsm_path)

    with zipfile.ZipFile(xlam_path, 'r') as zin:
        entries = {}
        for name in zin.namelist():
            entries[name] = zin.read(name)

    # 修改 Content_Types：将 addin 类型改为 workbook 类型
    ct = entries['[Content_Types].xml'].decode('utf-8')
    ct = ct.replace(
        'application/vnd.ms-excel.addin.macroEnabled.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
    )
    entries['[Content_Types].xml'] = ct.encode('utf-8')

    # 修改 workbook.xml：移除 addin 属性
    wb_xml = entries['xl/workbook.xml'].decode('utf-8')
    wb_xml = wb_xml.replace('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
                           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xlrt="http://schemas.microsoft.com/office/excel/2006/ribbon"')
    # 移除 isAddin 属性
    wb_xml = wb_xml.replace(' isAddin="1"', '')
    entries['xl/workbook.xml'] = wb_xml.encode('utf-8')

    # 写入 xlsm
    with zipfile.ZipFile(xlsm_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    print(f"[OK] 已转换: {xlam_path} -> {xlsm_path}")
    return xlsm_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("用法: python convert_xlam_to_xlsm.py <xlam路径> <xlsm路径>")
        sys.exit(1)
    xlam_to_xlsm(sys.argv[1], sys.argv[2])
