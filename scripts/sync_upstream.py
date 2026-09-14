#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlwings 技能上游同步脚本。

用法:
    python scripts/sync_upstream.py --check          # 仅检测更新，不替换
    python scripts/sync_upstream.py                  # 检测并同步有更新的上游
    python scripts/sync_upstream.py --force <id>     # 强制重拉指定 source（id: xlwings | xlwings-server）

机制:
  1. 读 manifest.json 中每路上游的 repo/ref/subpath/local_dir/pinned_sha
  2. 经 GitHub API（gh，失败降级 git ls-remote）查询 ref 最新 commit；无变化则跳过
  3. curl 下载 codeload tarball 到临时目录，tar 解压，校验非空
  4. 原子替换 local_dir（先备份 .old → 新内容就位 → 文件数校验 → 删备份），更新 manifest 并追加 SYNCLOG.md

约束:
  - 本脚本只写 local_dir、manifest.json、SYNCLOG.md，绝不触碰 SKILL.md / references/ 等自研文件。
  - 上游更新后若涉及版本号目录名（xlwings-<ver>），脚本自动以新 tag 命名；自研文档中的版本引用需人工核对刷新。
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
MANIFEST = SKILL_ROOT / "manifest.json"
SYNCLOG = SKILL_ROOT / "SYNCLOG.md"
GH = r"C:\Users\贺新\AppData\Local\Programs\gh\bin\gh.exe"
TIMEOUT = 120


def run(cmd: list[str], **kw) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=kw.pop("timeout", TIMEOUT), **kw)


def latest_sha(repo: str, ref: str) -> str | None:
    """ref 最新 commit SHA。优先 gh API，降级 git ls-remote（tags 取 peeled sha）。"""
    if os.path.isfile(GH):
        r = run([GH, "api", f"repos/{repo}/commits/{ref}", "--jq", ".sha"])
        if r.returncode == 0 and r.stdout.strip():
            return r.stdout.strip()
    try:
        r = run(["git", "ls-remote", f"https://github.com/{repo}.git",
                 f"refs/tags/{ref}", f"refs/tags/{ref}^{{}}"])
        for ln in r.stdout.splitlines():
            parts = ln.split()
            if len(parts) == 2 and parts[1].endswith("^{}"):
                return parts[0]
        if r.stdout.strip():
            return r.stdout.strip().split()[0]
    except Exception:
        pass
    return None


def download_tarball(repo: str, ref: str) -> Path | None:
    """curl 下载 codeload tarball 并 tar 解压，返回解压根目录（含顶层目录）或 None。"""
    tmp = Path(tempfile.mkdtemp(prefix="xlwings_up_"))
    tarball = tmp / "src.tar.gz"
    url = f"https://codeload.github.com/{repo}/tar.gz/refs/tags/{ref}"
    r = run(["curl.exe", "-sL", "-o", str(tarball), url, "-w", "%{http_code}"])
    if r.returncode != 0 or not tarball.exists() or r.stdout.strip() != "200":
        return None
    r = run(["tar", "-xzf", str(tarball), "-C", str(tmp)])
    if r.returncode != 0:
        return None
    # 解压根目录：唯一顶层目录
    dirs = [d for d in tmp.iterdir() if d.is_dir()]
    if len(dirs) != 1:
        return None
    return dirs[0]


def count_files(p: Path) -> int:
    return sum(1 for _ in p.rglob("*") if _.is_file())


def replace_dir(stage: Path, target: Path) -> bool:
    """原子替换：备份 .old → 复制 → 文件数校验 → 删备份。失败回滚。"""
    if target.exists():
        bak = target.with_name(target.name + ".old")
        shutil.rmtree(bak, ignore_errors=True)
        try:
            target.rename(bak)
        except OSError:
            # 目录被占用（云同步等）时逐文件删除后重建
            shutil.rmtree(target, ignore_errors=True)
        if target.exists():
            return False
    else:
        bak = None
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copytree(stage, target)
    except OSError:
        if bak and bak.exists():
            shutil.rmtree(target, ignore_errors=True)
            bak.rename(target)
        return False
    if count_files(target) != count_files(stage):
        shutil.rmtree(target, ignore_errors=True)
        if bak and bak.exists():
            bak.rename(target)
        return False
    if bak and bak.exists():
        shutil.rmtree(bak, ignore_errors=True)
    return True


def sync_one(src: dict, force: bool = False) -> str:
    """同步单路上游，返回 updated / unchanged / failed。"""
    vid, repo, ref = src["id"], src["repo"], src["ref"]
    head = latest_sha(repo, ref)
    if head is None:
        print(f"  [{vid}] 无法获取远端 commit，跳过")
        return "failed"
    if not force and head == src.get("pinned_sha"):
        print(f"  [{vid}] 无更新（{head[:7]}）")
        return "unchanged"
    print(f"  [{vid}] 检测到更新 {src.get('pinned_sha', '?')[:7]} → {head[:7]}")

    stage = download_tarball(repo, ref)
    if stage is None:
        print(f"  [{vid}] 下载/解压失败")
        return "failed"
    if count_files(stage) == 0:
        print(f"  [{vid}] 下载结果为空，放弃替换")
        return "failed"

    if vid == "xlwings":
        # 目录名跟随版本：xlwings-<tag>
        target = SKILL_ROOT / f"xlwings-{ref}"
    else:
        target = SKILL_ROOT / src["local_dir"]
    ok = replace_dir(stage, target)
    if not ok:
        print(f"  [{vid}] 替换失败（可能被占用，已回滚）")
        return "failed"

    n_files = count_files(target)
    src["pinned_sha"] = head
    src["last_synced"] = date.today().isoformat()
    if vid == "xlwings":
        src["local_dir"] = target.relative_to(SKILL_ROOT).as_posix()
    MANIFEST.write_text(json.dumps(MANIFEST_CFG, ensure_ascii=False, indent=2), encoding="utf-8")

    with open(SYNCLOG, "a", encoding="utf-8") as f:
        if f.tell() == 0:
            f.write("# xlwings 技能上游同步日志\n\n| 日期 | 来源 | 方式 | 新 commit | 规模 |\n|---|---|---|---|---|\n")
        f.write(f"| {date.today().isoformat()} | `{vid}` | tarball | `{head[:7]}` | {n_files} 个文件 |\n")
    print(f"  [{vid}] 已更新（tarball，{n_files} 文件）→ {target}")
    return "updated"


def main() -> int:
    args = sys.argv[1:]
    check_only = "--check" in args
    force_ids = [args[i + 1] for i, a in enumerate(args) if a == "--force"]
    global MANIFEST_CFG
    MANIFEST_CFG = json.loads(MANIFEST.read_text(encoding="utf-8"))
    stats: dict[str, list[str]] = {"updated": [], "unchanged": [], "failed": []}
    for src in MANIFEST_CFG["sources"]:
        if check_only:
            head = latest_sha(src["repo"], src["ref"])
            has_new = head is not None and head != src.get("pinned_sha")
            state = "有更新" if has_new else ("远端不可达" if head is None else "已是最新")
            print(f"[check] {src['id']} ({src.get('local_dir')}): {state}")
            stats["updated" if has_new else "failed" if head is None else "unchanged"].append(src["id"])
            continue
        if src["id"] in force_ids:
            result = sync_one(src, force=True)
        else:
            head = latest_sha(src["repo"], src["ref"])
            if head and head != src.get("pinned_sha"):
                result = sync_one(src)
            elif head:
                result = "unchanged"
            else:
                result = "failed"
        stats[result].append(src["id"])
    if not check_only:
        print("\n完成：" + "; ".join(f"{k}: {'、'.join(v) if v else '无'}" for k, v in stats.items()))
    return 1 if stats["failed"] else 0


MANIFEST_CFG = None
if __name__ == "__main__":
    sys.exit(main())
