"""
xlwings release tool - rewrite Interpreter_Win path inside a released .xlsm.

Zip surgery on xl/sharedStrings.xml only: replaces the old absolute
interpreter path stored in the hidden xlwings.conf sheet with the new one.
Never touches vbaProject.bin / customUI, so the Ribbon stays intact
(unlike COM save, which strips customUI registration).

Usage:
    python rewrite_interpreter.py -x <file.xlsm> -i <new python.exe path>

Exit codes: 0 = rewritten (or already correct), 1 = nothing to rewrite / error.
"""
import argparse
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

SS_PART = "xl/sharedStrings.xml"
INTERP_RE = re.compile(r"<t[^>]*>([^<]*python\.exe[^<]*)</t>", re.IGNORECASE)


def find_old_paths(ss_text):
    return sorted({m.group(1) for m in INTERP_RE.finditer(ss_text)})


def main():
    ap = argparse.ArgumentParser(description="Rewrite Interpreter_Win path inside released xlsm")
    ap.add_argument("-x", "--xlsm", required=True)
    ap.add_argument("-i", "--interpreter", required=True)
    ap.add_argument("--dry-run", action="store_true", help="Show what would change, do not write")
    args = ap.parse_args()

    xlsm = Path(args.xlsm)
    new_path = str(Path(args.interpreter))
    if not xlsm.exists():
        print("ERROR: file not found:", xlsm)
        return 1
    if not new_path.lower().endswith("python.exe") or not new_path.lower().startswith(("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")):
        # light sanity: must look like an absolute windows path to python.exe
        if not re.match(r"^[A-Za-z]:[\\/]", new_path):
            print("ERROR: interpreter must be an absolute Windows path to python.exe")
            return 1

    with zipfile.ZipFile(xlsm) as zin:
        if SS_PART not in zin.namelist():
            print("ERROR:", SS_PART, "not found in workbook (inline strings not supported)")
            return 1
        ss_text = zin.read(SS_PART).decode("utf-8")
        old_paths = find_old_paths(ss_text)
        if not old_paths:
            print("ERROR: no interpreter path (python.exe) found in", SS_PART)
            return 1

        changed = [p for p in old_paths if p != new_path]
        if args.dry_run:
            print("WOULD REWRITE:")
            for p in old_paths:
                print("  ", p, "->", new_path)
            return 0

        for p in changed:
            ss_text = ss_text.replace(">" + p + "<", ">" + new_path + "<")

        fd, tmp_name = tempfile.mkstemp(suffix=".xlsm", dir=str(xlsm.parent))
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == SS_PART:
                        data = ss_text.encode("utf-8")
                    zout.writestr(item, data)
        except Exception:
            if os.path.exists(tmp_name):
                os.remove(tmp_name)
            raise

    os.replace(tmp_name, xlsm)
    print("REWRITTEN", xlsm.name)
    for p in changed:
        print("  old:", p)
    print("  new:", new_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
