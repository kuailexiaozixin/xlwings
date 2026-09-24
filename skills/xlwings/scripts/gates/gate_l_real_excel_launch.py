#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
门禁 L：真实启动 Excel 验证 XLSTART 加载项（攻克"质检假绿"）
- 已知：COM 方式启动的 Excel 不会加载 XLSTART 文件夹，导致门禁 B/E 在"未加载插件"的 Excel 上跑，
  全部通过，用户一打开真实 Excel 就报错。
- 本门禁用 Popen 真实启动 EXCEL.EXE，再 GetActiveObject 接管，验证 XLSTART 中的 xlam 已被自动加载、
  IsAddin 状态正确；可选执行一次 RunPython 链路口保活。

用法:
    python gate_l_real_excel_launch.py --xlstart <XLSTART目录> --xlam <名> [--addin-name <项目名>] [--runpython] [--force-kill]
注意:
    - 若检测到已有 Excel 实例在运行，默认跳过真实启动（避免与用户正在使用的 Excel 冲突），并给出提示。
    - 使用 --force-kill 可自动终止所有残留 Excel 进程后继续验证（开发/CI 环境推荐）。
    - 需 pywin32（xlwings 已依赖）。
"""
import argparse
import os
import subprocess
import sys
import tempfile
import time


def _excel_running():
    try:
        import win32com.client
        win32com.client.GetActiveObject('Excel.Application')
        return True
    except Exception:
        return False


def _kill_all_excel():
    """终止所有 Excel 进程（--force-kill 时使用）；按 PID 精确操作，禁止 /IM 裸杀"""
    try:
        import psutil
        killed = []
        for proc in psutil.process_iter(['name', 'pid']):
            if proc.info['name'] and proc.info['name'].lower() in ('excel.exe', 'et.exe'):
                try:
                    proc.terminate()
                    proc.wait(timeout=5)
                    killed.append(proc.info['pid'])
                except Exception:
                    pass
        if killed:
            print(f'  --force-kill: 已终止 {len(killed)} 个 Excel 进程: {killed}')
            time.sleep(2)
        return True
    except ImportError:
        # psutil 不可用时回退：tasklist 解析 PID → 按 PID 终止（禁止按映像名全局裸杀，防误杀用户正在工作的 Excel）
        try:
            out = subprocess.run(
                ['tasklist', '/FO', 'CSV', '/FI', 'IMAGENAME eq EXCEL.EXE'],
                capture_output=True, text=True, timeout=10).stdout
            pids = []
            for line in out.splitlines()[1:]:
                parts = line.split('","')
                if len(parts) > 1:
                    pid = parts[1].strip('"')
                    if pid.isdigit():
                        pids.append(pid)
            for pid in pids:
                subprocess.run(['taskkill', '/PID', pid, '/F'],
                               capture_output=True, timeout=10)
            if pids:
                print(f'  --force-kill: 已按 PID 终止 {len(pids)} 个 Excel 进程: {pids}')
            time.sleep(2)
            return True
        except Exception as e:
            print(f'  --force-kill 失败: {e}')
            return False


def _verify_interpreter_chain(interp):
    """验证配置表 Interpreter_Win 指向的解释器在加载项链可导入 xlwings。

    背景（2026-09-17 实测）：uv venv 在 Excel/cmd 链中可能不激活——
    显式调用 venv python.exe，但 PYTHONHOME 指向 base 时 PREFIX 回落到 base，
    venv site-packages 不进 sys.path → 'No module named xlwings'。
    本子步在 Excel 内报错之前用同款命令行链拦截。返回 True/False。
    """
    if not interp:
        print('  [跳过] 未提供 --interpreter，解释器链验证跳过（提供配置表 Interpreter_Win 值可启用）')
        return True
    if not os.path.exists(interp):
        print(f'  [FAIL] --interpreter 路径不存在: {interp}')
        return False
    probe_src = (
        "import sys, site\n"
        "print('EXE=' + sys.executable)\n"
        "print('PREFIX=' + sys.prefix)\n"
        "print('BASE=' + sys.base_prefix)\n"
        "print('SITE=' + ','.join(site.getsitepackages()))\n"
        "try:\n"
        "    import xlwings\n"
        "    print('XLWINGS=' + xlwings.__version__)\n"
        "except Exception as e:\n"
        "    print('XLWINGS_ERR=' + repr(e))\n"
    )
    # 探针代码写临时文件再执行：cmd/bat 对 -c 内联长代码的引号重解析会破坏参数
    #（2026-09-17 实测），传文件路径最可靠
    fd, tmp = tempfile.mkstemp(suffix='.py', prefix='xlwings_chain_probe_')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(probe_src)
        if interp.lower().endswith('.exe'):
            r = subprocess.run([interp, tmp], capture_output=True,
                               text=True, timeout=60, encoding='utf-8', errors='replace')
        else:
            # .bat 包装解释器：经 cmd 执行（bat 不能直接 CreateProcess）。
            # call、bat 路径、脚本路径必须分开传参——subprocess 的 list2cmdline
            # 会为含空格路径自动加引号；若拼成单个字符串，cmd 会在首个空格截断
            # 路径（2026-09-17 三组实验实测结论）。
            r = subprocess.run(
                ['cmd.exe', '/d', '/s', '/c', 'call', interp, tmp],
                capture_output=True, text=True, timeout=60,
                encoding='utf-8', errors='replace')
    except Exception as exc:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        print(f'  [FAIL] 解释器链验证执行异常: {exc}')
        return False
    try:
        os.unlink(tmp)
    except OSError:
        pass
    out = (r.stdout or '') + '\n' + (r.stderr or '')
    for line in out.splitlines():
        if line.startswith(('EXE=', 'PREFIX=', 'BASE=', 'SITE=', 'XLWINGS=')):
            print(f'    {line}')
    if 'XLWINGS=' in out and 'XLWINGS_ERR' not in out:
        print('  [PASS] 解释器链可导入 xlwings')
        return True
    print('  [FAIL] 解释器链无法导入 xlwings')
    if 'PREFIX=' in out and 'BASE=' in out:
        import re
        pm = re.search(r'PREFIX=(\S+)', out)
        bm = re.search(r'BASE=(\S+)', out)
        if pm and bm and pm.group(1) == bm.group(1):
            print('    判据: PREFIX == BASE → venv 未生效（site-packages 不进 sys.path）')
            print('    修复: 用 bat 包装解释器（set PYTHONHOME=<base> + set PYTHONPATH=<模块目录> + 调 base 解释器），'
                  '见 references/04-python-guidance.md 第五节')
        else:
            print('    判据: PREFIX != BASE → venv 正常，需在 venv 内补装依赖或核对 PYTHONPATH')
    return False


def _find_excel_exe():
    # 权威来源：注册表 App Paths（覆盖即点即用 Root 布局与传统布局）
    try:
        import winreg
        for hive, view in ((winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY),
                           (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY)):
            try:
                with winreg.OpenKey(hive,
                                    r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
                                    0, winreg.KEY_READ | view) as k:
                    val, _ = winreg.QueryValueEx(k, '')
                    if val and os.path.exists(val):
                        return val
            except OSError:
                continue
    except Exception:
        pass
    # 回退：常见目录扫描（含即点即用 Root 布局）
    for root in (os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)')):
        if not root:
            continue
        for sub in ('Microsoft Office\\Root', 'Microsoft Office'):
            for ver in ('Office16', 'Office15', 'Office14'):
                cand = os.path.join(root, sub, ver, 'EXCEL.EXE')
                if os.path.exists(cand):
                    return cand
    return None


def main():
    parser = argparse.ArgumentParser(description='门禁 L：真实启动 Excel 验证 XLSTART 加载项')
    parser.add_argument('--xlstart', required=True, help='XLSTART 目录')
    parser.add_argument('--xlam', required=True, help='xlam 文件名')
    parser.add_argument('--addin-name', default='', help='项目名关键词（模糊匹配工作簿名，可选）')
    parser.add_argument('--runpython', action='store_true', help='额外执行一次 RunPython 链路口保活')
    parser.add_argument('--force-kill', action='store_true', help='检测到 Excel 实例时自动终止残留进程后继续验证（按 PID，不 /IM 裸杀）')
    parser.add_argument('--interpreter', default='', help='配置表 Interpreter_Win 的完整路径（python.exe 或 .bat 包装），验证加载项链可导入 xlwings')
    args = parser.parse_args()

    xlam_path = os.path.join(args.xlstart, args.xlam)
    if not os.path.exists(xlam_path):
        print(f'门禁 L 失败: XLSTART 中未找到 {xlam_path}')
        sys.exit(1)

    # 子步 0：解释器链验证（在真实 Excel 报错之前拦截 'No module named xlwings'）
    if args.interpreter:
        print('  子步 0——解释器链验证...')
        if not _verify_interpreter_chain(args.interpreter):
            print('门禁 L 失败: 解释器链无法导入 xlwings（见上方修复指引）')
            sys.exit(1)

    if _excel_running():
        if args.force_kill:
            print('检测到已有 Excel 实例，--force-kill 模式下自动终止...')
            if not _kill_all_excel():
                print('门禁 L 失败: 无法终止 Excel 残留进程')
                sys.exit(1)
        else:
            print('检测到已有 Excel 实例在运行，跳过真实启动（避免冲突）。')
            print('请在无 Excel 实例时运行本门禁以完成真机验证，或使用 --force-kill 自动清理。')
            sys.exit(0)

    import win32com.client

    exe = _find_excel_exe()
    if exe is None:
        print('门禁 L 失败: 未找到 EXCEL.EXE')
        sys.exit(1)

    print(f'  真实启动: {exe}')
    proc = subprocess.Popen([exe])
    xl = None
    try:
        for _ in range(30):
            try:
                xl = win32com.client.GetActiveObject('Excel.Application')
                if xl is not None:
                    break
            except Exception:
                time.sleep(1)
        if xl is None:
            print('门禁 L 失败: 无法接管 Excel 实例')
            sys.exit(1)

        # Excel 启动期间 XLSTART 加载项初始化，COM 调用可能被拒绝（0x800AC472 忙碌），须重试
        for _ in range(30):
            try:
                xl.Visible = False
                break
            except Exception:
                time.sleep(1)
        else:
            print('门禁 L 失败: Excel 启动后 30s 内未就绪（Visible=False 持续失败）')
            sys.exit(1)
        xl.DisplayAlerts = False
        time.sleep(2)

        # XLSTART 加载项按名字访问（Workbooks.Count 不枚举 XLSTART）
        wb = None
        try:
            wb = xl.Workbooks(args.xlam)
        except Exception:
            wb = None
        if wb is None:
            print(f'门禁 L 失败: XLSTART 加载项 {args.xlam} 未被自动加载')
            sys.exit(1)

        print(f'  XLSTART 加载项已自动加载: {args.xlam} (IsAddin={wb.IsAddin})')

        if args.runpython:
            try:
                xl.Run(f"'{args.xlam}'!RunPython", 'import sys; print("gate_l_runpython_ok")')
                print('  RunPython 链路口保活: OK')
            except Exception as exc:  # noqa: BLE001
                print(f'  RunPython 链路异常（非阻断）: {exc}')

        try:
            wb.Close(False)
        except Exception:
            pass
        print('门禁 L 通过：真实启动 Excel 后 XLSTART 加载项已自动加载')
    finally:
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass
        try:
            proc.terminate()
        except Exception:
            pass


if __name__ == '__main__':
    main()
