#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlwings-lite 官网镜像生成器（纯逻辑层）。

背景（为什么需要独立模块）:
  lite.xlwings.org 的文档源是**私有 Sphinx 构建**，公开仓库 `xlwings/xlwings-lite`
  仅做 issue/discussion 跟踪（只有 README + FUNDING，无 docs 源码）。站点唯一可机读的
  上游工件是它暴露的两个文件：
    - https://lite.xlwings.org/llms.txt       页面索引（name → 标题/摘要，可枚举全量页面）
    - https://lite.xlwings.org/llms-full.txt  全部正文，单文件，按 `# <name>.md` 分页
  本地 `xlwings-lite/` 目录就是这份 llms-full 的**离线 Markdown 镜像**（首版 2026-09-15）。

本模块把「镜像是否有实质更新」压缩成两个可跟踪量，并给出可复现的重建算法：
  1. top_version —— changelog 页顶层版本号（`## x.x.x.x` 最新一条）
  2. page_names  —— llms.txt 枚举出的页面集合（增删页即漂移）
上游出新版 → 二者之一变化 → 触发重建。重建规则与首版镜像逐页等价的排版约定：
  - 每页正文首行为其自带 H1（`# <Title>`），在 H1 后插入 `> 来源：<url>` 行；
  - 图片引用 `images/<f>` 归一为 `_images/<f>`（镜像 `_images/` 存放）；
  - 图片下载 URL 规则：`<base_url>/_images/<f>`（Sphinx furo 约定，已实测 200）。

纯逻辑层约定：本模块不碰网络、不落盘（写入/下载/原子替换由调用方 sync_upstream.py 负责），
便于单测与复用。
"""
from __future__ import annotations

import re
from collections import OrderedDict

# 分页标记：llms-full.txt 中每页以独立一行 `# <name>.md` 起始。
_PAGE_MARKER = re.compile(r"^# ([A-Za-z0-9_\-]+\.md)$")
# llms.txt 页面条目：- [Title](name.md): summary
_INDEX_ITEM = re.compile(r"^- \[([^\]]+)\]\(([A-Za-z0-9_\-]+\.md)\)")
# changelog 版本标题：## 1.0.0.105 (date)
_VERSION_HEAD = re.compile(r"^## (\d+\.\d+\.\d+\.\d+)\s*\(", re.M)
# 图片引用：_images/<f>（归一后）
_IMG_REF = re.compile(r"\(_images/([A-Za-z0-9_.\-]+)\)")


def split_pages(llms_full_text: str) -> "OrderedDict[str, list[str]]":
    """按 `# <name>.md` 分页，返回 {name: [正文行…]}（去掉首尾空行与分页标记本身）。"""
    pages: "OrderedDict[str, list[str]]" = OrderedDict()
    cur = None
    for ln in llms_full_text.splitlines():
        m = _PAGE_MARKER.match(ln)
        if m:
            cur = m.group(1)
            pages[cur] = []
        elif cur is not None:
            pages[cur].append(ln)
    for body in pages.values():
        while body and not body[0].strip():
            body.pop(0)
        while body and not body[-1].strip():
            body.pop()
    return pages


def parse_index(llms_txt: str) -> "OrderedDict[str, dict]":
    """解析 llms.txt 页面索引，返回 {name: {title, summary}}（保持站点顺序）。"""
    idx: "OrderedDict[str, dict]" = OrderedDict()
    for ln in llms_txt.splitlines():
        m = _INDEX_ITEM.match(ln.strip())
        if m:
            idx[m.group(2)] = {"title": m.group(1), "summary": ln[len(m.group(0)):].lstrip(": ").strip()}
    return idx


def cure_page(name: str, body_lines: list[str], base_url: str) -> str:
    """套用镜像排版约定：正文首行 H1 后插 `> 来源：<url>`，图片归一为 `_images/`。

    若正文首行不是 H1（异常情况），退化为直接在顶部插入来源行。
    """
    text = "\n".join(body_lines)
    text = text.replace("images/", "_images/")  # images/<f> -> _images/<f>
    url = f"{base_url.rstrip('/')}/{name[:-3]}.html"
    lines = text.split("\n")
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and lines[i].startswith("# "):
        i += 1  # 跳到 H1 之后
        while i < len(lines) and not lines[i].strip():
            i += 1  # 跳过 H1 与正文之间的空行
        lines = lines[:i] + [f"> 来源：{url}", ""] + lines[i:]
    else:
        lines = [f"> 来源：{url}", ""] + lines
    out = "\n".join(lines)
    while out.startswith("\n"):
        out = out[1:]
    return out + "\n"


def top_changelog_version(changelog_body: str) -> str:
    """changelog 页顶层版本号（最新一条 `## x.x.x.x`）；无则返回 ''。"""
    m = _VERSION_HEAD.findall(changelog_body)
    return m[0] if m else ""


def collect_image_refs(pages_content: dict[str, str]) -> set:
    """归一后 `_images/<f>` 引用集合（用于下载缺失图片）。"""
    refs: set = set()
    for content in pages_content.values():
        for m in _IMG_REF.finditer(content):
            refs.add(m.group(1))
    return refs


def build_mirror(llms_full_text: str, llms_txt: str, base_url: str) -> dict:
    """一次性产出重建所需全部结构。

    返回 dict:
      pages      : OrderedDict{name: markdown 全文}（保持站点顺序，仅含上游页，不含 README）
      top_version: changelog 顶层版本号（镜像实质更新判据之一）
      page_names : set(页名)（页面增删判据）
      images     : set(需存在于 _images/ 的图片名)
    """
    pages_raw = split_pages(llms_full_text)
    pages: "OrderedDict[str, str]" = OrderedDict()
    for name, body in pages_raw.items():
        pages[name] = cure_page(name, body, base_url)
    changelog = pages.get("changelog.md", "")
    return {
        "pages": pages,
        "top_version": top_changelog_version(changelog),
        "page_names": set(pages.keys()),
        "images": collect_image_refs(pages),
    }


if __name__ == "__main__":
    # 自检：给定已下载的 llms-full/llms 文件，打印重建摘要（不写盘）。
    import sys
    from pathlib import Path
    if len(sys.argv) >= 3:
        full = Path(sys.argv[1]).read_text(encoding="utf-8")
        idx = Path(sys.argv[2]).read_text(encoding="utf-8")
        r = build_mirror(full, idx, "https://lite.xlwings.org/")
        print("pages:", len(r["pages"]))
        print("top_version:", r["top_version"])
        print("images_needed:", len(r["images"]))
