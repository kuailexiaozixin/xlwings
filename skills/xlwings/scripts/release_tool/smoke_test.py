"""
xlwings release 安装后冒烟测试

验证内容：
  1. 安装目录结构（xlam + runtime/ + src/ + 入口模块）
  2. 配置表（Interpreter_Win / RELEASE_EMBED_CODE / PYTHONPATH）
  3. Python 源码可导入
  4. 模拟按钮点击（open_panel）
  5. 面板 API 响应

用法：
  python smoke_test.py --install-dir <path>
"""

import argparse
import os
import sys
import time
import socket
import subprocess

def check_install_dir(install_dir):
    """检查安装目录结构"""
    print("\n[1/5] 检查安装目录结构")
    required = ["announcement_downloader.xlam", "runtime", "src", "announcement_downloader.py"]
    all_ok = True
    for item in required:
        path = os.path.join(install_dir, item)
        exists = os.path.exists(path)
        status = "OK" if exists else "FAIL"
        if not exists:
            all_ok = False
        print(f"  [{status}] {item}")
    
    # 检查 python.exe
    python_exe = os.path.join(install_dir, "runtime", "python.exe")
    if os.path.exists(python_exe):
        print(f"  [OK] runtime/python.exe")
    else:
        print(f"  [FAIL] runtime/python.exe")
        all_ok = False
    
    return all_ok


def check_config(install_dir):
    """检查配置表"""
    print("\n[2/5] 检查配置表")
    python_exe = os.path.join(install_dir, "runtime", "python.exe")
    xlam_path = os.path.join(install_dir, "announcement_downloader.xlam")
    
    check_code = rf'''
import win32com.client
import os
import sys

xlam_path = r"{xlam_path}"
install_dir = r"{install_dir}"
expected_interpreter = os.path.join(install_dir, "runtime", "python.exe")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(xlam_path)
    config = {{}}
    for sheet in wb.Sheets:
        if "xlwings" in sheet.Name.lower() and "conf" in sheet.Name.lower():
            used = sheet.UsedRange
            for i in range(1, used.Rows.Count + 1):
                key = sheet.Cells(i, 1).Value
                val = sheet.Cells(i, 2).Value
                if key:
                    # 处理布尔值：False 也是有效值，不能用 if val 判断
                    if val is None:
                        config[str(key).strip()] = ""
                    elif isinstance(val, bool):
                        config[str(key).strip()] = str(val)
                    else:
                        config[str(key).strip()] = str(val).strip()
    wb.Close(SaveChanges=False)
finally:
    excel.Quit()

# 验证关键配置
checks = [
    ("Interpreter_Win", expected_interpreter),
    ("RELEASE_EMBED_CODE", "False"),
    ("PYTHONPATH", install_dir),
]
all_ok = True
for key, expected in checks:
    actual = config.get(key, "")
    # PYTHONPATH 可能用正斜杠（避免 \U unicode 转义问题）
    if key == "PYTHONPATH":
        actual_norm = actual.replace("\\", "/").lower()
        expected_norm = expected.replace("\\", "/").lower()
        ok = (actual_norm == expected_norm)
    else:
        ok = (actual == expected)
    status = "OK" if ok else "FAIL"
    if not ok:
        all_ok = False
    print(f"  [{{status}}] {{key}}: {{actual}} (期望: {{expected}})")

sys.exit(0 if all_ok else 1)
'''
    
    script_path = os.path.join(os.environ.get('TEMP', '.'), '_smoke_config.py')
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(check_code)
    
    try:
        result = subprocess.run([python_exe, script_path], capture_output=True, 
                               text=True, encoding='utf-8', errors='replace', timeout=60)
        print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr[-500:])
        return result.returncode == 0
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)


def check_imports(install_dir):
    """检查 Python 源码可导入"""
    print("\n[3/5] 检查 Python 源码导入")
    python_exe = os.path.join(install_dir, "runtime", "python.exe")
    
    check_code = rf'''
import sys
sys.path.insert(0, r"{install_dir}")
try:
    import announcement_downloader
    print("  [OK] announcement_downloader")
    import src.app
    print("  [OK] src.app")
    import src.main
    print("  [OK] src.main")
    import src.cninfo_api
    print("  [OK] src.cninfo_api")
    sys.exit(0)
except Exception as e:
    print(f"  [FAIL] {{e}}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
'''
    
    script_path = os.path.join(os.environ.get('TEMP', '.'), '_smoke_import.py')
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(check_code)
    
    try:
        result = subprocess.run([python_exe, script_path], capture_output=True,
                               text=True, encoding='utf-8', errors='replace', timeout=30)
        print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr[-500:])
        return result.returncode == 0
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)


def check_panel(install_dir):
    """模拟按钮点击并验证面板启动"""
    print("\n[4/5] 模拟按钮点击（open_panel）")
    python_exe = os.path.join(install_dir, "runtime", "python.exe")
    
    check_code = rf'''
import sys
import os
import time
import socket
import subprocess

install_dir = r"{install_dir}"
sys.path.insert(0, install_dir)
os.chdir(install_dir)

import xlwings as xw
import win32com.client

# 启动 Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(os.path.join(install_dir, "announcement_downloader.xlam"))
time.sleep(1)

# 设置 mock_caller 并调用 open_panel
xw.Book("announcement_downloader.xlam").set_mock_caller()
import announcement_downloader

try:
    announcement_downloader.open_panel()
    print("  [OK] open_panel() 调用成功")
except Exception as e:
    print(f"  [FAIL] open_panel() 失败: {{e}}")
    wb.Close(SaveChanges=False)
    excel.Quit()
    sys.exit(1)

time.sleep(3)

# 检查端口
s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
s.settimeout(2)
port_open = (s.connect_ex(("127.0.0.1", 8766)) == 0)
s.close()

if port_open:
    print("  [OK] 面板端口 8766 已开放")
else:
    print("  [FAIL] 面板端口 8766 未开放")

wb.Close(SaveChanges=False)
excel.Quit()
sys.exit(0 if port_open else 1)
'''
    
    script_path = os.path.join(os.environ.get('TEMP', '.'), '_smoke_panel.py')
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(check_code)
    
    try:
        result = subprocess.run([python_exe, script_path], capture_output=True,
                               text=True, encoding='utf-8', errors='replace', timeout=60)
        print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr[-500:])
        return result.returncode == 0
    finally:
        if os.path.exists(script_path):
            os.remove(script_path)


def check_api():
    """验证面板 API 响应"""
    print("\n[5/5] 验证面板 API 响应")
    try:
        import urllib.request
        resp = urllib.request.urlopen("http://127.0.0.1:8766/", timeout=5)
        if resp.status == 200:
            content = resp.read().decode('utf-8', errors='replace')
            print(f"  [OK] GET / : {resp.status} OK, 内容长度: {len(content)}")
            return True
        else:
            print(f"  [FAIL] GET / : {resp.status}")
            return False
    except Exception as e:
        print(f"  [FAIL] GET / : {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description='xlwings release 安装后冒烟测试')
    parser.add_argument('--install-dir', required=True, help='安装目录路径')
    args = parser.parse_args()
    
    install_dir = os.path.abspath(args.install_dir)
    print(f"冒烟测试 - 安装目录: {install_dir}")
    
    results = []
    results.append(("目录结构", check_install_dir(install_dir)))
    results.append(("配置表", check_config(install_dir)))
    results.append(("源码导入", check_imports(install_dir)))
    results.append(("面板启动", check_panel(install_dir)))
    results.append(("API 响应", check_api()))
    
    print("\n" + "=" * 50)
    print("冒烟测试汇总:")
    all_pass = True
    for name, ok in results:
        status = "PASS" if ok else "FAIL"
        if not ok:
            all_pass = False
        print(f"  [{status}] {name}")
    
    print("=" * 50)
    if all_pass:
        print("全部通过！")
        sys.exit(0)
    else:
        print("存在失败项，请检查上方输出")
        sys.exit(1)


if __name__ == "__main__":
    main()
