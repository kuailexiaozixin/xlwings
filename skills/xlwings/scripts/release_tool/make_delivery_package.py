"""
xlwings release tool - make a distributable delivery package (zip) from a
release dist folder.

Step 5 of the delivery workflow:
    build runtime -> release_clone -> reinject ribbon -> gates -> PACKAGE

What goes into the package:
    - the released .xlsm / .xlam
    - the portable runtime folder (contains python.exe)
    - installer pair:  install_release.bat/.ps1 + rewrite_interpreter.py
    - uninstaller pair: uninstall_release.bat/.ps1
    - docs: any *.txt in the dist folder (e.g. 安装说明.txt)

What is excluded:
    - developer regression scripts (_e2e_test.py), __pycache__, ~$ lock
      files, previous *.zip packages

Usage:
    python make_delivery_package.py --dist <dist folder> [--output <zip path>]
"""
import argparse
import datetime
import sys
import zipfile
from pathlib import Path

INCLUDE_DOCS = ("*.txt",)
EXCLUDE_NAMES = {"_e2e_test.py"}
INSTALLER_FILES = [
    "install_release.bat", "install_release.ps1",
    "uninstall_release.bat", "uninstall_release.ps1",
    "rewrite_interpreter.py",
]
INSTALLER_SRC = Path(__file__).parent


def find_runtime(dist: Path):
    for d in sorted(dist.iterdir()):
        if d.is_dir() and (d / "python.exe").exists():
            return d
    return None


def find_workbook(dist: Path):
    for pattern in ("*.xlsm", "*.xlam"):
        files = [f for f in sorted(dist.glob(pattern)) if not f.name.startswith("~$")]
        if files:
            return files[0]
    return None


def zip_dir(zf: zipfile.ZipFile, folder: Path, arc_prefix: str):
    for p in sorted(folder.rglob("*")):
        if "__pycache__" in p.parts:
            continue
        arcname = arc_prefix / p.relative_to(folder)
        zf.write(p, arcname)


def main():
    ap = argparse.ArgumentParser(description="Package a release dist folder for distribution")
    ap.add_argument("--dist", required=True, help="release dist folder")
    ap.add_argument("--output", default="", help="output zip path (default: <dist parent>\\<name>_v<date>.zip)")
    args = ap.parse_args()

    dist = Path(args.dist).resolve()
    if not dist.is_dir():
        print("ERROR: dist folder not found:", dist)
        return 1

    workbook = find_workbook(dist)
    runtime = find_runtime(dist)
    if not workbook:
        print("ERROR: no .xlsm/.xlam found in", dist)
        return 1
    if not runtime:
        print("ERROR: no portable runtime (folder with python.exe) found in", dist)
        return 1

    base = dist.name
    output = Path(args.output) if args.output else dist.parent / f"{base}_v{datetime.date.today():%Y%m%d}.zip"

    missing = [f for f in INSTALLER_FILES if not (INSTALLER_SRC / f).exists()]
    if missing:
        print("ERROR: installer components missing:", ", ".join(missing))
        return 1

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.write(workbook, workbook.name)
        zip_dir(zf, runtime, Path(runtime.name))
        for doc in sorted(dist.glob("*.txt")):
            zf.write(doc, doc.name)
        for f in INSTALLER_FILES:
            zf.write(INSTALLER_SRC / f, f)

    size_mb = output.stat().st_size / 1048576
    print("=" * 50)
    print("PACKAGE OK:", output)
    print(f"  size: {size_mb:.1f} MB")
    print("  contents: workbook + runtime + installer pair + uninstaller pair + docs")
    print("  user flow: unzip -> double-click install_release.bat -> done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
