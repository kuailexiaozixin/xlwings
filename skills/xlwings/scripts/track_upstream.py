#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlwings 技能上游漂移跟踪（三线检测，参照 hermes-business-agent 四线与 fastapi vendor 机制）。

三线:
  ① 上游 commit 线 —— manifest.json 各路 pinned_sha vs 远端 ref 最新 commit。
     DRIFT 为提示级（同步即消，跑 sync_upstream.py），--strict 时计为失败。
  ② 事实锚点线 —— references/upstream-facts.json 逐条对照本地快照实测
     （行数/KB/行号/文件数）。STALE 为硬漂移：自研文档断言已与上游不符，
     必须先修文档再引用（带病作业 = 断言失真）。
  ③ 版本一致性线 —— 快照特征校验：xlwings/__init__.py 的 __version__ 必须为
     "0.0.0"（源码 zip 快照无构建版本，见 SKILL.md §1）；异常说明目录被 pip
     产物覆盖，属硬问题。

用法:
    python scripts/track_upstream.py             # 三线全查
    python scripts/track_upstream.py --quick     # 线①只查核心两路（xlwings/xlwings-server）
    python scripts/track_upstream.py --line 2    # 只跑第②线（锚点）
    python scripts/track_upstream.py --strict    # 线① commit 漂移也计为失败

退出码:
    0 = 全绿（或仅线①提示级漂移且未加 --strict）
    1 = 线②锚点 STALE / 线③版本异常 / --strict 下线①有漂移 / 检测本身失败

分工纪律（SKILL.md §9）:
    - 同步上游后：必须跑本脚本 → 修 STALE 锚点指向的自研文档 → 重跑至 ②③ 全绿。
    - 引用上游行数/行号/KB 前可先跑 --quick 确认锚点未破。
    - 新增上游事实断言时：把锚点登记进 upstream-facts.json（含 docs 指针），
      否则漂移无人知晓。
"""
from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_ROOT / "scripts"))

import sync_upstream as sync  # noqa: E402  复用 latest_sha / manifest 读取

FACTS = SKILL_ROOT / "references" / "upstream-facts.json"
CORE_SOURCES = ("xlwings", "xlwings-server")


# ---------------------------------------------------------------- 线① commit
def line1_commit(quick: bool) -> tuple[list[str], int]:
    """返回 (漂移报告行列表, 漂移路数)。"""
    manifest = json.loads(sync.MANIFEST.read_text(encoding="utf-8-sig"))
    lines: list[str] = []
    drifted = 0
    for src in manifest["sources"]:
        if quick and src["id"] not in CORE_SOURCES:
            continue
        vid, ref = src["id"], src["ref"]
        head = sync.latest_sha(src["repo"], ref)
        pinned = src.get("pinned_sha", "")
        if head is None:
            lines.append(f"    [{vid}] UNREACHABLE（远端不可达，网络间歇，稍后重试）")
        elif head != pinned:
            drifted += 1
            lines.append(f"    [{vid}] DRIFT {pinned[:7]} → {head[:7]}（跑 sync_upstream.py 同步）")
        else:
            lines.append(f"    [{vid}] OK {head[:7]}")
    return lines, drifted


# ---------------------------------------------------------------- 线② 锚点

# ---------------------------------------------------------------- 线②b docs 正文链接悬空
DOCS_INDEX_DIRS = ("xlwings/docs", "xlwings-server/docs", "xlwings-lite",
                   "docs", "examples", "MCP-Server")

def _collect_docs_index() -> set:
    """候选上游/自研 docs 页面文件名并集（basename，递归含子目录）。"""
    known: set[str] = set()
    for d in DOCS_INDEX_DIRS:
        dd = SKILL_ROOT / d
        if dd.is_dir():
            for p in dd.rglob("*.md"):
                known.add(p.name)
    return known

def line2_docs_links() -> tuple[list[str], int]:
    """docs/<单段>.md 正文链接悬空检测（硬漂移）。

    自研文档（SKILL.md + references/*.md）里引用的 `docs/xxx.md` 顶层页面，
    若在本技能快照候选 docs 目录中已查无同名文件，判为 STALE——典型成因是
    上游重构删除/更名了文档页而自研文档未跟随（如 xlwings-server v1.0 移除
    repo_setup.md、upgrade.md）。口径与手工审计一致：仅核 `docs/<单段>.md`
    顶层链接，靠 basename 存在性判定；带通配（docs/wasm_*.md）与子目录引用
    （docs/api/）天然不匹配，子目录页面另由 filecount 类锚点覆盖。
    """
    known = _collect_docs_index()
    pat = re.compile(r'docs/([A-Za-z0-9_\-]+\.md)')
    lines: list[str] = []
    bad = 0
    docs = [SKILL_ROOT / "SKILL.md"] + sorted((SKILL_ROOT / "references").glob("*.md"))
    for doc in docs:
        if not doc.is_file():
            continue
        text = doc.read_text(encoding="utf-8", errors="replace")
        rel = doc.relative_to(SKILL_ROOT).as_posix()
        seen: set[str] = set()
        for m in pat.finditer(text):
            name = m.group(1)
            if name in seen:
                continue
            seen.add(name)
            if name not in known:
                ln = text[:m.start()].count("\n") + 1
                lines.append(f"    [{rel}:{ln}] STALE 悬空 docs 链接 `docs/{name}`（快照已无该页 → 查上游去向并改指）")
                bad += 1
    return lines, bad

def _check_fact(root: Path, fact: dict) -> str | None:
    """单条锚点检测。返回 None=OK，否则返回问题描述。"""
    src_dir = SKILL_ROOT / fact["source"]
    target = src_dir / fact["file"]
    kind = fact["kind"]
    if kind == "filecount":
        if not target.is_dir():
            return f"目录缺失 {target.name}"
        n = len(list(target.glob("*.md")))
        return None if n == fact["value"] else f"篇数 {n} ≠ 台账 {fact['value']}"
    if not target.is_file():
        return "文件缺失"
    if kind == "exists":
        ok = target.is_file()
        return None if ok == fact["value"] else f"存在性 {ok} ≠ 台账 {fact['value']}"
    if kind == "lines":
        n = len(target.read_text(encoding="utf-8", errors="replace").splitlines())
        return None if n == fact["value"] else f"行数 {n} ≠ 台账 {fact['value']}"
    if kind == "sizekb":
        kb = target.stat().st_size // 1024
        return None if kb == fact["value"] else f"大小 {kb}KB ≠ 台账 {fact['value']}KB"
    if kind == "symbol_line":
        want = fact["value"]
        text_lines = target.read_text(encoding="utf-8", errors="replace").splitlines()
        actual = text_lines[want["line"] - 1].strip() if want["line"] <= len(text_lines) else ""
        if actual.startswith(want["symbol"]):
            return None
        # 行号漂移：全文件搜符号真实位置，给出新行号便于修文档
        for i, l in enumerate(text_lines, 1):
            if l.strip().startswith(want["symbol"]):
                return f"锚点漂移：{want['symbol']!r} 实际在 L{i}（台账 L{want['line']}）"
        return f"符号 {want['symbol']!r} 已不存在于该文件"
    return f"未知 kind: {kind}"


def line2_facts() -> tuple[list[str], int]:
    lines: list[str] = []
    stale = 0
    data = json.loads(FACTS.read_text(encoding="utf-8-sig"))
    for fact in data["facts"]:
        err = _check_fact(SKILL_ROOT, fact)
        if err is None:
            lines.append(f"    [{fact['id']}] OK")
        else:
            stale += 1
            docs = ", ".join(fact.get("docs", [])) or "（无文档指针）"
            lines.append(f"    [{fact['id']}] STALE {err} → 修: {docs}")
    dl, dbad = line2_docs_links()
    lines += dl
    stale += dbad
    return lines, stale


# ---------------------------------------------------------------- 线③ 版本
def line3_version() -> tuple[list[str], int]:
    lines: list[str] = []
    broken = 0
    init = SKILL_ROOT / "xlwings" / "xlwings" / "__init__.py"
    if not init.is_file():
        return [f"    [snapshot] 快照文件缺失: {init.name}"], 1
    text = init.read_text(encoding="utf-8", errors="replace")
    m = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', text)
    ver = m.group(1) if m else "(未找到)"
    if ver == "0.0.0":
        lines.append("    [snapshot] __version__ == \"0.0.0\"（zip 快照特征，符合预期；真实版本查 manifest）")
    else:
        broken += 1
        lines.append(f'    [snapshot] __version__ == "{ver}" ≠ "0.0.0"（目录疑似被 pip 产物覆盖，硬问题）')
    return lines, broken


# ---------------------------------------------------------------- main
def main() -> int:
    args = sys.argv[1:]
    quick = "--quick" in args
    strict = "--strict" in args
    only_line = None
    if "--line" in args:
        only_line = args[args.index("--line") + 1]

    hard_fail = 0
    print("== 线① 上游 commit（提示级" + ("，--strict 计失败" if strict else "") + "）==")
    if only_line in (None, "1"):
        l1, drift = line1_commit(quick)
        print("\n".join(l1) or "    （无跟踪路）")
        if drift and strict:
            hard_fail += 1
        if drift:
            print("    ↳ 提示：上游有更新，同步后需跑线②复核锚点")
    else:
        print("    （跳过）")

    print("== 线② 事实锚点（硬漂移，STALE 必修）==")
    if only_line in (None, "2"):
        l2, stale = line2_facts()
        print("\n".join(l2))
        hard_fail += stale
        print(f"    ↳ {stale} 条 STALE" if stale else "    ↳ 全部一致")
    else:
        print("    （跳过）")

    print("== 线③ 版本一致性（硬）==")
    if only_line in (None, "3"):
        l3, broken = line3_version()
        print("\n".join(l3))
        hard_fail += broken
    else:
        print("    （跳过）")

    print()
    print("结论: " + ("FAIL（硬漂移，先修再引用）" if hard_fail else "PASS（可安全引用自研文档中的上游断言）"))
    return 1 if hard_fail else 0


if __name__ == "__main__":
    sys.exit(main())
