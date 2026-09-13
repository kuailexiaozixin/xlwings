"""
xlwings release 复刻工具 - 卸载脚本

根据 release_clone.py 的安装逻辑，反向移除 MyAddin VBAProject。

卸载步骤：
1. 移除 xlwings VBA 引用（如果存在）
2. 删除所有注入的 VBA 组件：xlwings, Dictionary, RibbonMyAddin, modAnnouncement 等
3. 删除嵌入的代码 sheet（.py 结尾）
4. 删除配置表 xlwings.conf
5. 保存并关闭工作簿

用法：
    python uninstall_myaddin.py <xlsm路径>
    或
    python uninstall_myaddin.py --current  # 处理当前活跃 Excel 的工作簿
"""
import argparse
import sys
from pathlib import Path

# 要移除的 VBA 组件列表（根据 release_clone.py 的安装逻辑）
COMPONENTS_TO_REMOVE = [
    "xlwings",           # 无加载项 VBA 模块
    "Dictionary",        # VBA 字典类
    "IWebAuthenticator", # Web 认证（可选）
    "WebClient",         # Web 客户端（可选）
    "WebRequest",        # Web 请求（可选）
    "WebResponse",       # Web 响应（可选）
    "WebHelpers",        # Web 工具（可选）
    "RibbonMyAddin",     # Ribbon 回调
    "modAnnouncement",   # 业务逻辑模块
    "MyAddin",           # 通用 MyAddin 名称
]


def uninstall_from_workbook(wb, verbose=True):
    """从工作簿中移除 MyAddin VBAProject"""
    print(f"\n处理工作簿: {wb.name}")
    print(f"路径: {wb.fullname if wb.fullname else '(未保存)'}")

    removed_count = 0
    errors = []

    try:
        vb = wb.api.VBProject
    except Exception as e:
        print(f"  [WARN] 无法访问 VBA 项目: {e}")
        print("  提示: 可能需要启用 Trust Access to VBA Project Object Model")
        return False

    # 1. 移除 xlwings VBA 引用
    print("\n  [1/4] 检查 VBA 引用...")
    try:
        refs = vb.References
        ref_names = [r.Name for r in refs]
        if "xlwings" in ref_names:
            vb.References.Remove(vb.References("xlwings"))
            if verbose:
                print("    [OK] 已移除 xlwings 引用")
            removed_count += 1
        else:
            if verbose:
                print("    ○ xlwings 引用不存在，跳过")
    except Exception as e:
        errors.append(f"移除引用失败: {e}")
        if verbose:
            print(f"    [WARN] 移除引用失败: {e}")

    # 2. 删除注入的 VBA 组件
    print("\n  [2/4] 检查 VBA 组件...")
    components_found = []
    for comp_name in COMPONENTS_TO_REMOVE:
        try:
            comp = vb.VBComponents(comp_name)
            components_found.append(comp_name)
        except Exception:
            pass  # 组件不存在，跳过

    if components_found:
        for comp_name in components_found:
            try:
                comp = vb.VBComponents(comp_name)
                vb.VBComponents.Remove(comp)
                if verbose:
                    print(f"    [OK] 已移除组件: {comp_name}")
                removed_count += 1
            except Exception as e:
                errors.append(f"移除组件 {comp_name} 失败: {e}")
                if verbose:
                    print(f"    [WARN] 移除 {comp_name} 失败: {e}")
    else:
        if verbose:
            print("    ○ 未找到已知组件，跳过")

    # 3. 删除嵌入的代码 sheet（.py 结尾）
    print("\n  [3/4] 检查嵌入的代码 sheet...")
    sheets_to_remove = []
    for sheet in wb.sheets:
        if sheet.name.endswith(".py"):
            sheets_to_remove.append(sheet.name)

    if sheets_to_remove:
        for sheet_name in sheets_to_remove:
            try:
                wb.sheets[sheet_name].delete()
                if verbose:
                    print(f"    [OK] 已删除代码 sheet: {sheet_name}")
                removed_count += 1
            except Exception as e:
                errors.append(f"删除 sheet {sheet_name} 失败: {e}")
                if verbose:
                    print(f"    [WARN] 删除 {sheet_name} 失败: {e}")
    else:
        if verbose:
            print("    ○ 未找到嵌入的代码 sheet，跳过")

    # 4. 删除配置表 xlwings.conf
    print("\n  [4/4] 检查配置表...")
    if "xlwings.conf" in wb.sheet_names:
        try:
            wb.sheets["xlwings.conf"].delete()
            if verbose:
                print("    [OK] 已删除配置表 xlwings.conf")
            removed_count += 1
        except Exception as e:
            errors.append(f"删除配置表失败: {e}")
            if verbose:
                print(f"    [WARN] 删除配置表失败: {e}")
    else:
        if verbose:
            print("    ○ 配置表不存在，跳过")

    # 汇总结果
    print("\n" + "=" * 50)
    if removed_count > 0:
        print(f"[OK] 成功移除 {removed_count} 个项目")
    else:
        print("○ 未找到需要移除的项目（可能已经清理过）")

    if errors:
        print(f"\n[WARN] 发生 {len(errors)} 个错误:")
        for err in errors:
            print(f"  - {err}")

    return len(errors) == 0


def main():
    parser = argparse.ArgumentParser(description="卸载 MyAddin VBAProject")
    parser.add_argument("xlsm", nargs="?", help="目标 .xlsm 文件路径")
    parser.add_argument("--current", action="store_true", help="处理当前活跃 Excel 的工作簿")
    parser.add_argument("--verbose", "-v", action="store_true", help="详细输出")
    args = parser.parse_args()

    import xlwings as xw

    app = xw.App(visible=False)
    try:
        if args.current:
            # 处理当前活跃的工作簿
            if not app.books:
                print("错误: 当前没有打开的工作簿")
                sys.exit(1)
            wb = app.books.active
            success = uninstall_from_workbook(wb, verbose=args.verbose)
            if success and args.verbose:
                print("\n提示: 请保存并重新打开工作簿以应用更改")
        elif args.xlsm:
            # 处理指定文件
            xlsm_path = Path(args.xlsm)
            if not xlsm_path.exists():
                print(f"错误: 文件不存在: {xlsm_path}")
                sys.exit(1)
            wb = app.books.open(str(xlsm_path))
            success = uninstall_from_workbook(wb, verbose=args.verbose)
            if success:
                wb.save()
                print(f"\n[OK] 已保存更改: {xlsm_path}")
            wb.close()
        else:
            print("用法:")
            print("  python uninstall_myaddin.py <xlsm路径>      # 卸载指定文件")
            print("  python uninstall_myaddin.py --current        # 卸载当前活跃工作簿")
            sys.exit(1)

        if success:
            print("\n[OK] 卸载完成")
        else:
            print("\n[FAIL] 卸载完成但有错误，请检查上方输出")
            sys.exit(1)

    finally:
        app.quit()


if __name__ == "__main__":
    main()
