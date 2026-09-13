"""
便携 Python 运行时构建工具

支持三种模式：
  1. --source minimal     从开发机 Python 提取最小运行时（推荐，体积 ~150MB）
  2. --source embeddable  从 Python embeddable zip 构建
  3. --source existing    从现有 Python 完整复制（体积大，不推荐）

用法：
  python install_runtime.py --target D:/path/to/runtime --source minimal
  python install_runtime.py --target D:/path/to/runtime --source embeddable --zip python-3.13.14-embed-amd64.zip
  python install_runtime.py --target D:/path/to/runtime --source existing --python D:/Python/cpython-3.13.14
"""

import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


# ─── 工具函数 ───────────────────────────────────────────────────────────────

def get_size(path):
    """计算目录总大小（字节）"""
    total = 0
    for root, dirs, files in os.walk(path):
        for f in files:
            fp = os.path.join(root, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def format_size(size_bytes):
    """格式化文件大小"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def ensure_dir(path):
    """确保目录存在"""
    Path(path).mkdir(parents=True, exist_ok=True)


def find_uv():
    """查找 uv 可执行文件，返回路径或 None"""
    # 1. 环境变量
    uv_path = shutil.which("uv")
    if uv_path:
        return uv_path
    # 2. 常见位置
    candidates = [
        os.path.join(os.environ.get("USERPROFILE", ""), ".local", "bin", "uv.exe"),
        os.path.join(sys.prefix, "Scripts", "uv.exe"),
        r"D:\Python\cpython-3.13.14-windows-x86_64-none\Scripts\uv.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def find_pip(python_exe):
    """查找 pip 模块（通过 python -m pip）"""
    return python_exe  # 用 python -m pip


def validate_requirements(req_file):
    """
    验证 requirements.txt 中的包名，修正常见错误
    返回 (修正后的包列表, 警告列表)
    """
    warnings = []
    packages = []

    # 常见包名映射（错误名 → 正确名）
    NAME_FIXES = {
        "fasthtml": "python-fasthtml",
        "fast-html": "python-fasthtml",
        "xlwings-pro": "xlwings",
        "pillow": "Pillow",
        "opencv": "opencv-python",
        "sklearn": "scikit-learn",
    }

    if not os.path.exists(req_file):
        return packages, [f"requirements.txt 不存在: {req_file}"]

    with open(req_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            # 提取包名（去掉版本号）
            pkg_name = line.split("==")[0].split(">=")[0].split("<=")[0].strip()
            if pkg_name.lower() in NAME_FIXES:
                correct = NAME_FIXES[pkg_name.lower()]
                warnings.append(f"包名修正: {pkg_name} → {correct}")
                # 保留版本约束
                if "==" in line:
                    packages.append(f"{correct}=={line.split('==')[1]}")
                elif ">=" in line:
                    packages.append(f"{correct}>={line.split('>=')[1]}")
                else:
                    packages.append(correct)
            else:
                packages.append(line)

    return packages, warnings


def install_dependencies(target_dir, req_file, use_uv=True):
    """
    安装依赖到目标运行时的 site-packages

    Args:
        target_dir: 运行时目录
        req_file: requirements.txt 路径
        use_uv: 是否使用 uv（更快，推荐）

    Returns:
        bool: 是否成功
    """
    python_exe = os.path.join(target_dir, "python.exe")
    site_packages = os.path.join(target_dir, "Lib", "site-packages")
    ensure_dir(site_packages)

    # 验证 requirements.txt
    packages, warnings = validate_requirements(req_file)
    for w in warnings:
        print(f"  [WARN] {w}")

    if not packages:
        print("  [INFO] 无依赖需要安装")
        return True

    # 写入临时 requirements 文件（修正后的包名）
    temp_req = os.path.join(target_dir, "_requirements_fixed.txt")
    with open(temp_req, "w", encoding="utf-8") as f:
        f.write("\n".join(packages) + "\n")

    try:
        if use_uv:
            uv_path = find_uv()
            if uv_path:
                print(f"  [INFO] 使用 uv 安装依赖（{uv_path}）")
                cmd = [
                    uv_path, "pip", "install",
                    "-r", temp_req,
                    "--target", site_packages,
                    "--no-cache",
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
                if result.returncode != 0:
                    print(f"  [ERROR] uv 安装失败:")
                    print(result.stderr[-2000:] if result.stderr else "无错误输出")
                    # 回退到 pip
                    print("  [INFO] 回退到 pip...")
                    use_uv = False
                else:
                    print("  [OK] uv 安装完成")
                    if result.stdout:
                        # 只打印最后几行
                        lines = result.stdout.strip().split("\n")
                        for line in lines[-5:]:
                            print(f"    {line}")
                    return True
            else:
                print("  [WARN] 未找到 uv，使用 pip")
                use_uv = False

        if not use_uv:
            print("  [INFO] 使用 pip 安装依赖")
            cmd = [
                python_exe, "-m", "pip", "install",
                "-r", temp_req,
                "--target", site_packages,
                "--no-cache-dir",
                "--no-warn-script-location",
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
            if result.returncode != 0:
                print(f"  [ERROR] pip 安装失败:")
                print(result.stderr[-2000:] if result.stderr else "无错误输出")
                return False
            print("  [OK] pip 安装完成")
            return True

    finally:
        # 清理临时文件
        if os.path.exists(temp_req):
            os.remove(temp_req)

    return False


def cleanup_runtime(target_dir):
    """
    安装后清理：删除不必要的文件，减小体积

    删除内容：
    - __pycache__ 目录
    - *.pyc 文件
    - tests / test 目录
    - docs / doc 目录
    - *.dist-info 中的 RECORD（保留元数据）
    - 未使用的标准库模块（可选，保守起见不删）
    """
    removed = 0
    freed_bytes = 0

    for root, dirs, files in os.walk(target_dir):
        # 删除 __pycache__
        if "__pycache__" in dirs:
            d = os.path.join(root, "__pycache__")
            try:
                size = get_size(d)
                shutil.rmtree(d, ignore_errors=True)
                removed += 1
                freed_bytes += size
            except Exception:
                pass
            dirs.remove("__pycache__")

        # 删除 tests 目录（仅在 site-packages 中）
        if "site-packages" in root:
            for test_dir in ["tests", "test", "testing"]:
                if test_dir in dirs:
                    d = os.path.join(root, test_dir)
                    try:
                        size = get_size(d)
                        shutil.rmtree(d, ignore_errors=True)
                        removed += 1
                        freed_bytes += size
                    except Exception:
                        pass
                    dirs.remove(test_dir)

        # 删除 .pyc 文件
        for f in files:
            if f.endswith(".pyc"):
                fp = os.path.join(root, f)
                try:
                    size = os.path.getsize(fp)
                    os.remove(fp)
                    removed += 1
                    freed_bytes += size
                except Exception:
                    pass

    print(f"  [CLEAN] 删除 {removed} 项，释放 {format_size(freed_bytes)}")


def verify_runtime(target_dir, extra_imports=None):
    """
    验证运行时是否可用

    Args:
        target_dir: 运行时目录
        extra_imports: 额外需要验证的导入列表

    Returns:
        bool: 是否全部通过
    """
    python_exe = os.path.join(target_dir, "python.exe")
    if not os.path.exists(python_exe):
        print(f"  [ERROR] python.exe 不存在: {python_exe}")
        return False

    default_imports = [
        "sys", "os", "json", "re", "datetime",
        "xlwings", "win32com.client", "pythoncom",
        "requests",
    ]
    if extra_imports:
        default_imports.extend(extra_imports)

    # 构建验证脚本（清理 BOM 和不可见字符）
    clean_imports = []
    for m in default_imports:
        clean = m.strip().lstrip("\ufeff").lstrip("\u200b")
        if clean:
            clean_imports.append(clean)
    import_lines = "\n".join([f"import {m}" for m in clean_imports])
    verify_script = f"""import sys
print(f"Python: {{sys.version.split()[0]}}")
print(f"sys.path[0]: {{sys.path[0]}}")
print("--- 验证导入 ---")
{import_lines}
print("RUNTIME_OK")
"""

    temp_script = os.path.join(target_dir, "_verify.py")
    with open(temp_script, "w", encoding="utf-8") as f:
        f.write(verify_script)

    try:
        result = subprocess.run(
            [python_exe, temp_script],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=60,
        )
        print(result.stdout)
        if result.stderr:
            print("STDERR:", result.stderr[-1000:])
        return "RUNTIME_OK" in result.stdout
    finally:
        if os.path.exists(temp_script):
            os.remove(temp_script)


# ─── 模式 1: 最小运行时（推荐） ──────────────────────────────────────────────

def install_minimal(target_dir, source_python=None, req_file=None, use_uv=True):
    """
    从开发机 Python 提取最小运行时

    只复制必要文件：
    - 核心文件（python.exe、python3XX.dll、vcruntime140.dll 等）
    - DLLs 目录（C 扩展依赖）
    - Lib 标准库精简版（排除 site-packages、test、idlelib、tkinter 等）
    - 创建 python3XX._pth 隔离模式文件

    然后用 uv/pip 安装依赖到 Lib/site-packages

    Args:
        target_dir: 目标运行时目录
        source_python: 源 Python 目录（默认使用当前 Python）
        req_file: requirements.txt 路径
        use_uv: 是否使用 uv 安装依赖

    Returns:
        bool: 是否成功
    """
    print(f"\n{'='*60}")
    print(f"[MODE] 最小运行时构建")
    print(f"{'='*60}")

    # 确定源 Python
    if source_python is None:
        source_python = sys.prefix
    source_python = os.path.abspath(source_python)
    print(f"[INFO] 源 Python: {source_python}")
    print(f"[INFO] 目标目录: {target_dir}")

    if not os.path.exists(os.path.join(source_python, "python.exe")):
        print(f"[ERROR] 源 Python 不存在: {source_python}")
        return False

    # 获取 Python 版本
    version_cmd = [os.path.join(source_python, "python.exe"), "-c", "import sys; print(sys.version_info.major, sys.version_info.minor, sys.version_info.micro)"]
    result = subprocess.run(version_cmd, capture_output=True, text=True)
    major, minor, micro = map(int, result.stdout.strip().split())
    pth_file = f"python{major}{minor}._pth"
    print(f"[INFO] Python 版本: {major}.{minor}.{micro}")

    # 清空目标目录
    if os.path.exists(target_dir):
        print(f"[INFO] 清空目标目录...")
        shutil.rmtree(target_dir, ignore_errors=True)
    ensure_dir(target_dir)

    # ── 1. 复制核心文件 ──
    print(f"\n[STEP 1/4] 复制核心文件")
    core_files = [
        "python.exe",
        "pythonw.exe",
        f"python{major}{minor}.dll",
        f"python{major}.dll",
        "vcruntime140.dll",
        "vcruntime140_1.dll",
        "msvcp140.dll",
        "concrt140.dll",
    ]
    copied = 0
    for f in core_files:
        src = os.path.join(source_python, f)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(target_dir, f))
            copied += 1
            print(f"  [OK] {f}")
        else:
            print(f"  [SKIP] {f}（不存在）")
    print(f"  共复制 {copied} 个核心文件")

    # ── 2. 复制 DLLs 目录 ──
    print(f"\n[STEP 2/4] 复制 DLLs 目录")
    src_dlls = os.path.join(source_python, "DLLs")
    dst_dlls = os.path.join(target_dir, "DLLs")
    if os.path.exists(src_dlls):
        shutil.copytree(src_dlls, dst_dlls)
        size = get_size(dst_dlls)
        print(f"  [OK] DLLs 复制完成: {format_size(size)}")
    else:
        print(f"  [WARN] DLLs 目录不存在: {src_dlls}")

    # ── 3. 复制精简标准库 ──
    print(f"\n[STEP 3/4] 复制精简标准库")
    src_lib = os.path.join(source_python, "Lib")
    dst_lib = os.path.join(target_dir, "Lib")
    ensure_dir(dst_lib)

    # 需要排除的目录
    EXCLUDE_DIRS = {
        "site-packages", "test", "tests", "idlelib", "tkinter",
        "turtledemo", "lib2to3", "ensurepip", "distutils",
        "unittest", "pydoc_data", "msilib", "win32comext",
    }
    # 需要排除的文件扩展名
    EXCLUDE_EXTS = {".pyc", ".pyo"}

    copied_files = 0
    skipped_dirs = set()
    for root, dirs, files in os.walk(src_lib):
        # 过滤排除目录
        rel = os.path.relpath(root, src_lib)
        parts = rel.split(os.sep) if rel != "." else []

        # 检查是否在排除路径中
        skip = False
        for p in parts:
            if p in EXCLUDE_DIRS:
                skip = True
                skipped_dirs.add(p)
                break
        if skip:
            dirs[:] = []  # 不遍历子目录
            continue

        # 创建目标目录
        rel_path = os.path.relpath(root, src_lib)
        if rel_path == ".":
            dst_root = dst_lib
        else:
            dst_root = os.path.join(dst_lib, rel_path)
            ensure_dir(dst_root)

        # 复制文件
        for f in files:
            if os.path.splitext(f)[1] in EXCLUDE_EXTS:
                continue
            src_file = os.path.join(root, f)
            dst_file = os.path.join(dst_root, f)
            try:
                shutil.copy2(src_file, dst_file)
                copied_files += 1
            except Exception as e:
                print(f"  [WARN] 复制失败 {f}: {e}")

    print(f"  [OK] 复制 {copied_files} 个标准库文件")
    print(f"  [SKIP] 排除目录: {', '.join(sorted(skipped_dirs))}")

    # 创建 site-packages 目录
    ensure_dir(os.path.join(dst_lib, "site-packages"))

    # ── 4. 创建 ._pth 隔离模式文件 ──
    print(f"\n[STEP 4/4] 创建 {pth_file} 隔离模式文件")
    pth_content = f"""python{major}{minor}.zip
.
Lib
Lib\\site-packages
DLLs
import site
"""
    with open(os.path.join(target_dir, pth_file), "w", encoding="utf-8") as f:
        f.write(pth_content)
    print(f"  [OK] {pth_file} 已创建")

    # 统计基础运行时大小
    base_size = get_size(target_dir)
    print(f"\n[INFO] 基础运行时大小: {format_size(base_size)}")

    # ── 5. 安装依赖 ──
    if req_file and os.path.exists(req_file):
        print(f"\n[STEP 5] 安装依赖")
        print(f"  requirements: {req_file}")
        if not install_dependencies(target_dir, req_file, use_uv=use_uv):
            print("  [ERROR] 依赖安装失败")
            return False
    else:
        print(f"\n[INFO] 未指定 requirements.txt，跳过依赖安装")

    # ── 6. 清理 ──
    print(f"\n[STEP 6] 清理不必要文件")
    cleanup_runtime(target_dir)

    # ── 7. 验证 ──
    print(f"\n[STEP 7] 验证运行时")
    extra = []
    if req_file:
        # 从 requirements.txt 提取需要验证的包
        with open(req_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    pkg = line.split("==")[0].split(">=")[0].strip().lower()
                    # 映射导入名
                    import_map = {
                        "python-fasthtml": "fasthtml",
                        "pyyaml": "yaml",
                        "pillow": "PIL",
                        "opencv-python": "cv2",
                        "scikit-learn": "sklearn",
                        "pywin32": "win32com",
                    }
                    extra.append(import_map.get(pkg, pkg))

    if not verify_runtime(target_dir, extra_imports=extra):
        print("  [ERROR] 运行时验证失败")
        return False

    # 最终大小
    final_size = get_size(target_dir)
    print(f"\n{'='*60}")
    print(f"[DONE] 最小运行时构建完成")
    print(f"  目录: {target_dir}")
    print(f"  大小: {format_size(final_size)}")
    print(f"{'='*60}")
    return True


# ─── 模式 2: embeddable zip ─────────────────────────────────────────────────

def install_from_embeddable_zip(target_dir, zip_path, req_file=None, use_uv=True):
    """
    从 Python embeddable zip 构建运行时

    Args:
        target_dir: 目标运行时目录
        zip_path: embeddable zip 文件路径
        req_file: requirements.txt 路径
        use_uv: 是否使用 uv 安装依赖

    Returns:
        bool: 是否成功
    """
    print(f"\n{'='*60}")
    print(f"[MODE] Embeddable Zip 构建")
    print(f"{'='*60}")
    print(f"[INFO] Zip: {zip_path}")
    print(f"[INFO] 目标: {target_dir}")

    if not os.path.exists(zip_path):
        print(f"[ERROR] Zip 文件不存在: {zip_path}")
        return False

    # 清空目标目录
    if os.path.exists(target_dir):
        shutil.rmtree(target_dir, ignore_errors=True)
    ensure_dir(target_dir)

    # 解压
    print(f"\n[STEP 1] 解压 embeddable zip")
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(target_dir)
    print(f"  [OK] 解压完成")

    # 找到 ._pth 文件并修改（启用 site-packages）
    pth_files = list(Path(target_dir).glob("python*._pth"))
    if pth_files:
        pth_file = pth_files[0]
        print(f"\n[STEP 2] 修改 {pth_file.name} 启用 site-packages")
        content = pth_file.read_text(encoding="utf-8")
        if "import site" not in content:
            content += "\nimport site\n"
        if "Lib\\site-packages" not in content:
            content = content.replace("import site", "Lib\\site-packages\nimport site")
        pth_file.write_text(content, encoding="utf-8")
        print(f"  [OK] {pth_file.name} 已更新")
        print(content)

    # 创建 Lib/site-packages
    ensure_dir(os.path.join(target_dir, "Lib", "site-packages"))

    # 安装 pip（embeddable 版本默认没有 pip）
    print(f"\n[STEP 3] 安装 pip")
    python_exe = os.path.join(target_dir, "python.exe")
    get_pip = os.path.join(target_dir, "get-pip.py")
    # 尝试用 ensurepip
    result = subprocess.run([python_exe, "-m", "ensurepip", "--upgrade"],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.returncode != 0:
        print(f"  [WARN] ensurepip 失败，尝试下载 get-pip.py")
        # 这里可以下载 get-pip.py，但为了简化，直接报错
        print("  [ERROR] 无法安装 pip，请手动安装")
        return False
    print("  [OK] pip 安装完成")

    # 安装依赖
    if req_file and os.path.exists(req_file):
        print(f"\n[STEP 4] 安装依赖")
        if not install_dependencies(target_dir, req_file, use_uv=use_uv):
            print("  [ERROR] 依赖安装失败")
            return False

    # 清理
    print(f"\n[STEP 5] 清理")
    cleanup_runtime(target_dir)

    # 验证
    print(f"\n[STEP 6] 验证运行时")
    if not verify_runtime(target_dir):
        print("  [ERROR] 运行时验证失败")
        return False

    final_size = get_size(target_dir)
    print(f"\n[DONE] Embeddable 运行时构建完成: {format_size(final_size)}")
    return True


# ─── 模式 3: 完整复制（不推荐） ──────────────────────────────────────────────

def install_from_existing_python(target_dir, source_python=None, req_file=None):
    """
    从现有 Python 完整复制运行时（体积大，不推荐）

    注意：此模式会复制整个 Python 目录（可能数 GB），
    建议使用 install_minimal 替代。

    Args:
        target_dir: 目标运行时目录
        source_python: 源 Python 目录
        req_file: requirements.txt 路径（可选，用于额外安装依赖）

    Returns:
        bool: 是否成功
    """
    print(f"\n{'='*60}")
    print(f"[MODE] 完整复制（不推荐，体积大）")
    print(f"{'='*60}")
    print(f"[WARN] 此模式会复制整个 Python 目录，可能数 GB")
    print(f"[WARN] 建议使用 --source minimal 替代")

    if source_python is None:
        source_python = sys.prefix
    source_python = os.path.abspath(source_python)

    if not os.path.exists(os.path.join(source_python, "python.exe")):
        print(f"[ERROR] 源 Python 不存在: {source_python}")
        return False

    # 清空目标目录
    if os.path.exists(target_dir):
        shutil.rmtree(target_dir, ignore_errors=True)

    # 复制
    print(f"\n[STEP 1] 复制 Python 目录（可能需要几分钟）...")
    shutil.copytree(source_python, target_dir)
    print(f"  [OK] 复制完成")

    # 清空 site-packages（避免携带开发机的无关包）
    site_packages = os.path.join(target_dir, "Lib", "site-packages")
    if os.path.exists(site_packages):
        print(f"\n[STEP 2] 清空 site-packages")
        for item in os.listdir(site_packages):
            item_path = os.path.join(site_packages, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path, ignore_errors=True)
                else:
                    os.remove(item_path)
            except Exception as e:
                print(f"  [WARN] 删除失败 {item}: {e}")
        print(f"  [OK] site-packages 已清空")

    # 安装依赖
    if req_file and os.path.exists(req_file):
        print(f"\n[STEP 3] 安装依赖")
        if not install_dependencies(target_dir, req_file, use_uv=True):
            print("  [ERROR] 依赖安装失败")
            return False

    # 验证
    print(f"\n[STEP 4] 验证运行时")
    if not verify_runtime(target_dir):
        print("  [ERROR] 运行时验证失败")
        return False

    final_size = get_size(target_dir)
    print(f"\n[DONE] 运行时构建完成: {format_size(final_size)}")
    return True


# ─── 主入口 ─────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="便携 Python 运行时构建工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 最小运行时（推荐，~150MB）
  python install_runtime.py --target D:/runtime --source minimal

  # 指定源 Python 和 requirements
  python install_runtime.py --target D:/runtime --source minimal --python D:/Python/3.13 --req requirements.txt

  # embeddable zip
  python install_runtime.py --target D:/runtime --source embeddable --zip python-3.13-embed-amd64.zip

  # 完整复制（不推荐）
  python install_runtime.py --target D:/runtime --source existing
        """,
    )
    parser.add_argument("--target", required=True, help="目标运行时目录")
    parser.add_argument("--source", choices=["minimal", "embeddable", "existing"],
                       default="minimal", help="构建模式（默认 minimal）")
    parser.add_argument("--python", help="源 Python 目录（minimal/existing 模式）")
    parser.add_argument("--zip", help="embeddable zip 文件路径（embeddable 模式）")
    parser.add_argument("--req", help="requirements.txt 路径")
    parser.add_argument("--no-uv", action="store_true", help="不使用 uv，强制用 pip")

    args = parser.parse_args()

    target_dir = os.path.abspath(args.target)
    req_file = os.path.abspath(args.req) if args.req else None
    use_uv = not args.no_uv

    success = False
    if args.source == "minimal":
        success = install_minimal(target_dir, source_python=args.python,
                                  req_file=req_file, use_uv=use_uv)
    elif args.source == "embeddable":
        if not args.zip:
            print("[ERROR] embeddable 模式需要 --zip 参数")
            sys.exit(1)
        success = install_from_embeddable_zip(target_dir, args.zip,
                                              req_file=req_file, use_uv=use_uv)
    elif args.source == "existing":
        success = install_from_existing_python(target_dir, source_python=args.python,
                                               req_file=req_file)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
