#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlwings 技能上游同步脚本。

用法:
    python scripts/sync_upstream.py --check          # 仅检测更新，不替换
    python scripts/sync_upstream.py                  # 检测并同步有更新的上游
    python scripts/sync_upstream.py --force <id> [<id> ...]   # 只同步所列 source，且忽略 sha 比对强制重拉

机制:
  1. 读 manifest.json 中每路上游的 repo/ref/subpath/local_dir/pinned_sha
  2. 经 gh API（多候选探测 gh.exe，降级 git ls-remote）查询 ref 最新 commit；带重试；无变化则跳过
  3. 按通道链下载 tarball：codeload 直连 → archive 直连 → ghfast 代理，tar 解压，校验非空
  4. 原子替换 local_dir（先备份 .old → 新内容就位 → 文件数校验 → 删备份），更新 manifest 并追加 SYNCLOG.md

约束:
  - 本脚本只写 local_dir、manifest.json、SYNCLOG.md，绝不触碰 SKILL.md / references/ 等自研文件。
  - local_dir 一律是固定目录名（xlwings、xlwings-server），不随上游版本号改变。要切换跟踪点只改
    manifest.json 的 ref：跟最新正式版填 tag（如 0.37.4），跟开发线填 main。
  - 上游内容变化后，自研文档里引用的行号与代码片段需人工复核（脚本不做这件事）。
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
sys.path.insert(0, str(SKILL_ROOT / "scripts"))
import web_mirror  # noqa: E402  （xlwings-lite 官网镜像重建的纯逻辑层）

MANIFEST = SKILL_ROOT / "manifest.json"
SYNCLOG = SKILL_ROOT / "SYNCLOG.md"
TIMEOUT = 120
DOWNLOAD_TIMEOUT = 900
RETRIES = 3
RETRY_SLEEP = 3

# 下载通道降级链：codeload 直连 → archive 直连 → ghfast 代理 archive。
# 本机到 GitHub 直连时通时断（见 SYNCLOG 2026-09-24），通道链保证弱网下仍可同步。
DOWNLOAD_CHANNELS = (
    "codeload",   # https://codeload.github.com/{repo}/tar.gz/{ref_path}
    "archive",    # https://github.com/{repo}/archive/{ref_path}.tar.gz
    "ghfast",     # https://ghfast.top/https://github.com/{repo}/archive/{ref_path}.tar.gz
)


def _find_gh() -> str:
    """多候选探测 gh.exe：PATH 命中且真实存在 → Program Files → 各用户 LOCALAPPDATA。"""
    candidates = [shutil.which("gh")]
    candidates.append(str(Path(r"C:\Program Files\GitHub CLI\gh.exe")))
    drive = os.environ.get("SystemDrive", "C:")
    candidates += [str(p) for p in Path(f"{drive}\\Users").glob(
        r"*\AppData\Local\Programs\gh\bin\gh.exe")]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return ""


GH = _find_gh()


def run(cmd: list[str], **kw) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=kw.pop("timeout", TIMEOUT), **kw)


def latest_sha(repo: str, ref: str) -> str | None:
    """ref 最新 commit SHA。优先 gh API，降级 git ls-remote（tags 取 peeled sha）。
    网络间歇性中断：带 RETRIES 次重试。"""
    for attempt in range(RETRIES):
        if GH:
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
        if attempt < RETRIES - 1:
            import time
            time.sleep(RETRY_SLEEP)
    return None


def download_tarball(repo: str, ref: str) -> tuple[Path | None, str]:
    """按通道链下载 tarball 并解压。返回 (解压根目录, 通道名) 或 (None, '')。"""
    tmp = Path(tempfile.mkdtemp(prefix="xlwings_up_"))
    tarball = tmp / "src.tar.gz"
    # tag 用 refs/tags/<ref>，分支（如 main）用 refs/heads/<ref>
    for ref_path in (f"refs/tags/{ref}", f"refs/heads/{ref}", ref):
        urls = {
            "codeload": f"https://codeload.github.com/{repo}/tar.gz/{ref_path}",
            "archive": f"https://github.com/{repo}/archive/{ref_path}.tar.gz",
            "ghfast": f"https://ghfast.top/https://github.com/{repo}/archive/{ref_path}.tar.gz",
        }
        for channel in DOWNLOAD_CHANNELS:
            try:
                r = run(["curl.exe", "-sL", "-o", str(tarball), urls[channel],
                         "-w", "%{http_code}"], timeout=DOWNLOAD_TIMEOUT)
            except subprocess.TimeoutExpired:
                print(f"  下载超时（{DOWNLOAD_TIMEOUT}s）：{urls[channel]}")
                tarball.unlink(missing_ok=True)
                continue
            if r.returncode == 0 and tarball.exists() and r.stdout.strip() == "200":
                r2 = run(["tar", "-xzf", str(tarball), "-C", str(tmp)])
                if r2.returncode != 0:
                    tarball.unlink(missing_ok=True)
                    continue
                dirs = [d for d in tmp.iterdir() if d.is_dir()]
                if len(dirs) == 1:
                    return dirs[0], channel
            tarball.unlink(missing_ok=True)
    return None, ""


# web 源抓取：用 urllib 而非 curl.exe。本机 schannel curl 会因 CRL 吊销检查失败
# （CRYPT_E_REVOCATION，代理/直连都绕不过）；urllib 保留默认证书校验（不走那条失败的
# CRL 路径），既可靠又不像 curl --insecure 那样放弃证书校验（镜像要进公开仓库，完整性重要）。
_UA = {"User-Agent": "Mozilla/5.0 (xlwings-skill sync)"}

def _http_get(url: str, timeout: int) -> bytes | None:
    import ssl, urllib.request, time
    for _ in range(RETRIES):
        try:
            req = urllib.request.Request(url, headers=_UA)
            data = urllib.request.urlopen(req, timeout=timeout,
                                          context=ssl.create_default_context()).read()
            if data:
                return data
        except Exception:
            time.sleep(RETRY_SLEEP)
    return None

def _curl_text(url: str, timeout: int = 120) -> str | None:
    """抓取文本内容（web 源专用）。"""
    data = _http_get(url, timeout)
    return data.decode("utf-8", errors="replace") if data else None

def _curl_bytes(url: str, out: Path, timeout: int = 120) -> bool:
    """抓取二进制（图片）到 out。"""
    data = _http_get(url, timeout)
    if data:
        out.write_bytes(data)
        return True
    out.unlink(missing_ok=True)
    return False

def count_files(p: Path) -> int:
    return sum(1 for _ in p.rglob("*") if _.is_file())


def _keep_rel(target: Path, keep) -> set:
    """展开 keep 白名单（相对 local_dir 的路径片段）为 target 下需保留的文件相对路径集合。

    keep 项按「路径任一层级名命中」判定，故 `extracted` 会保留其整棵子树。
    """
    keep = set(keep or ())
    if not keep or not target.exists():
        return set()
    return {p.relative_to(target) for p in target.rglob("*")
            if p.is_file() and any(part in keep for part in p.relative_to(target).parts)}


def _restore_keep(bak: Path, target: Path, keep_rel) -> None:
    for rel in keep_rel:
        srcf = bak / rel
        if srcf.is_file():
            dst = target / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(srcf, dst)


def _run_extractor(target: Path) -> None:
    """local_dir 内若有 extract_webextension.py，同步后运行它刷新 extracted/ 明文产物。"""
    exe = target / "extract_webextension.py"
    if exe.is_file():
        try:
            r = subprocess.run([sys.executable, str(exe)], cwd=str(target),
                                capture_output=True, text=True, timeout=TIMEOUT)
            tag = "已刷新 extracted/" if r.returncode == 0 else f"退出码 {r.returncode}"
            print(f"  [{target.name}] extract_webextension.py {tag}")
            if r.returncode != 0:
                print(f"    stderr: {r.stderr.strip()[:200]}")
        except Exception as e:
            print(f"  [{target.name}] extract_webextension.py 运行失败（不影响同步）：{e}")

# ---------------------------------------------------------------- web 源（官网镜像）
def web_latest(src: dict) -> dict | None:
    """拉取 web 源的 llms-full.txt，算出 top_version / page 集合 / 全文 sha256。失败返回 None。"""
    import hashlib
    full = _curl_text(src["llms_url"])
    if not full:
        return None
    idx = _curl_text(src.get("index_url", "")) or ""   # llms.txt（仅用于页面顺序，缺失也能重建）
    mirror = web_mirror.build_mirror(full, idx, src["base_url"])
    return {
        "pages": mirror["pages"],          # OrderedDict{name: 归一后的 markdown 全文}
        "top_version": mirror["top_version"],
        "page_names": mirror["page_names"],
        "images": mirror["images"],
        "full_sha": hashlib.sha256(full.encode("utf-8")).hexdigest(),
    }

def sync_web_one(src: dict, force: bool = False) -> str:
    """同步 web 源（官网镜像）：重建 local_dir 下的上游页，保留 README.md 与 _images/，合并新图片。"""
    info = web_latest(src)
    if info is None:
        print(f"  [{src['id']}] 无法抓取 {src['llms_url']}，跳过")
        return "failed"

    pinned_ver = src.get("pinned_version", "")
    pinned_sha = src.get("pinned_full_sha", "")
    changed = force or info["top_version"] != pinned_ver or info["full_sha"] != pinned_sha
    if not changed:
        print(f"  [{src['id']}] 无更新（{info['top_version']}）")
        return "unchanged"

    base_url = src["base_url"]
    target = SKILL_ROOT / src["local_dir"]
    target.mkdir(parents=True, exist_ok=True)

    # 保留项：README.md（自研索引）与 _images/（图片素材）
    keep_readme = target / "README.md"
    readme_bak = keep_readme.read_bytes() if keep_readme.is_file() else None
    img_dir = target / "_images"
    img_dir.mkdir(parents=True, exist_ok=True)

    # 1) 合并缺失图片（从 <base_url>/_images/<f> 下载；已有则跳过）
    missing_imgs = []
    for f in sorted(info["images"]):
        dst = img_dir / f
        if dst.is_file() and dst.stat().st_size > 0:
            continue
        missing_imgs.append(f)
    for f in missing_imgs:
        ok = _curl_bytes(f"{base_url.rstrip('/')}/_images/{f}", img_dir / f)
        if not ok:
            print(f"  [{src['id']}] 图片下载失败: {f}")

    # 2) 写入/覆盖上游页（保留 README.md）
    for name, content in info["pages"].items():
        (target / name).write_text(content, encoding="utf-8", newline="\n")

    # 3) 清理上游已删除的旧页（top 层 .md 且不属于当前页集合，且非 README.md）
    for old in target.glob("*.md"):
        if old.name != "README.md" and old.name not in info["page_names"]:
            old.unlink()

    # 4) 恢复 README.md（自研，不被上游覆盖）
    if readme_bak is not None:
        keep_readme.write_bytes(readme_bak)

    src["pinned_version"] = info["top_version"]
    src["pinned_full_sha"] = info["full_sha"]
    src["pinned_pages"] = sorted(info["page_names"])
    src["last_synced"] = date.today().isoformat()
    MANIFEST.write_text(json.dumps(MANIFEST_CFG, ensure_ascii=False, indent=2), encoding="utf-8-sig")
    n_pages = len(info["page_names"])
    with open(SYNCLOG, "a", encoding="utf-8") as f:
        if f.tell() == 0:
            f.write("# xlwings 技能上游同步日志\n\n| 日期 | 来源 | 方式 | 新 commit | 规模 |\n|---|---|---|---|---|\n")
        n_imgs = sum(1 for _ in img_dir.iterdir())
        f.write(f"| {date.today().isoformat()} | `{src['id']}` | web(新图 {len(missing_imgs)}) | `{info['top_version']}` | {n_pages} 页 / {n_pages + 1 + n_imgs} 文件 |\n")
    print(f"  [{src['id']}] 已更新（{info['top_version']}，{n_pages} 页，新图 {len(missing_imgs)}）→ {target}")
    return "updated"


# ---------------------------------------------------------------- 增量同步（git 源）
# 全量 tarball 动辄 17~18MB，而一次上游提交常只改 1~2 个文件；本机到 GitHub 直连约
# 26KB/s，一次全量要 10 分钟以上。增量路径改走 API + raw 单文件：
#   - compare API（~1.4s）拿到变更清单（added/modified/removed/renamed）
#   - trees API（1 次请求）拿到全树 blob size，用于逐文件字节数校验
#   - raw 单文件（~1s/个）仅下载真正变动的文件
# 任一步异常（非 git 源、compare 截断、变更过多、下载或校验不符）自动回退全量 tarball，
# 绝不留下半新半旧的快照。

INCR_MAX_FILES = 120          # 变更文件超此值不建议增量，直接全量（省往返）
RAW_BASE = "https://raw.githubusercontent.com"


def _gh_api(path: str, timeout: int = 60):
    """调用 GitHub API（复用 gh CLI 凭据，带重试）。返回 dict 或 None。"""
    import time
    for attempt in range(RETRIES):
        if GH:
            r = run([GH, "api", path], timeout=timeout)
            if r.returncode == 0 and r.stdout.strip():
                try:
                    return json.loads(r.stdout)
                except json.JSONDecodeError:
                    pass
        if attempt < RETRIES - 1:
            time.sleep(RETRY_SLEEP)
    return None


def index_tree(repo: str, sha: str):
    """仓库全树 {path: size}。trees 返回 truncated 时返回 None（触发回退全量）。"""
    data = _gh_api(f"repos/{repo}/git/trees/{sha}?recursive=1")
    if not data or data.get("truncated"):
        return None
    return {n["path"]: n.get("size", 0)
            for n in data.get("tree", []) if n.get("type") == "blob"}


def compare_changes(repo: str, old: str, new: str):
    """compare old...new 的变更清单。清单不完整（300 条上限）或失败返回 None。"""
    data = _gh_api(f"repos/{repo}/compare/{old}...{new}")
    if not data or data.get("files") is None:
        return None
    files = data["files"]
    if len(files) >= 300:          # GitHub compare 的文件清单上限，达到即视为不完整
        return None
    return [{"path": f.get("filename", ""), "status": f.get("status", ""),
             "previous_path": f.get("previous_filename", "")}
            for f in files if f.get("filename")]


def fetch_raw(repo: str, sha: str, path: str, timeout: int = 60):
    """从 raw.githubusercontent 取单文件（公开仓库免 token，保留证书校验）。"""
    import ssl, time, urllib.parse, urllib.request
    url = f"{RAW_BASE}/{repo}/{sha}/{urllib.parse.quote(path)}"
    for attempt in range(RETRIES):
        try:
            req = urllib.request.Request(url, headers=_UA)
            return urllib.request.urlopen(req, timeout=timeout,
                                          context=ssl.create_default_context()).read()
        except Exception:
            if attempt < RETRIES - 1:
                time.sleep(RETRY_SLEEP)
    return None


def sync_incremental(src, old_sha, new_sha):
    """增量更新 local_dir。返回 (是否成功, 描述, 变更文件数)。

    只落真正变化的文件；成功条件 = 每个变更文件字节数与 trees 一致。先全部下载并校验，
    再统一写盘，任一步失败即返回 False 交给调用方回退全量，不产生中间态快照。
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    repo, subpath = src["repo"], src.get("subpath", "")
    changes = compare_changes(repo, old_sha, new_sha)
    if changes is None:
        return False, "compare 不可用或清单截断", 0
    pre = subpath.rstrip("/") + "/" if subpath else ""
    if pre:
        changes = [c for c in changes if c["path"].startswith(pre)]
    if not changes:
        return True, "无文件变化", 0
    if len(changes) > INCR_MAX_FILES:
        return False, f"变更 {len(changes)} 文件超过阈值 {INCR_MAX_FILES}", len(changes)

    sizes = index_tree(repo, new_sha)
    if sizes is None:
        return False, "trees 截断", len(changes)

    # 完整性护栏：增量只在「本地 == old_sha 的完整快照」时安全。用 old_sha 的 trees 做基线，
    # 本地文件数与它不一致（残缺 / 过旧 / 被手动改过）就回退全量重建，避免增量只补 diff
    # 而留下缺文件或多出的旧文件（实测 excel-mcp 本地缺 37、taxi-duckdb 多 16，均需全量）。
    target = SKILL_ROOT / src["local_dir"]
    keep = src.get("keep", ())
    old_trees = index_tree(repo, old_sha)
    if old_trees is None:
        return False, f"old {old_sha[:7]} 不可比对（可能 force-push），基线未知", len(changes)
    if target.exists():
        kn = len(_keep_rel(target, keep))
        local_n = count_files(target) - kn
        if local_n != len(old_trees):
            return False, f"本地(除 keep {kn} 项) {local_n} ≠ 基线 {len(old_trees)}，快照与 {old_sha[:7]} 不符", len(changes)

    downloads = []      # [(新路径, 期望字节数)]
    deletes = []        # 待删路径
    for c in changes:
        st, path = c["status"], c["path"]
        if st == "renamed" and c["previous_path"]:
            old_p = c["previous_path"]
            if not pre or old_p.startswith(pre):
                deletes.append(old_p)
        if st == "removed":
            deletes.append(path)
        else:
            downloads.append((path, sizes.get(path)))

    missing = [p for p, sz in downloads if sz is None]
    if missing:
        return False, f"trees 缺少 {missing[0]}", len(changes)

    blobs = {}
    with ThreadPoolExecutor(max_workers=6) as ex:
        futs = {ex.submit(fetch_raw, repo, new_sha, p): (p, sz) for p, sz in downloads}
        for fut in as_completed(futs):
            path, want = futs[fut]
            data = fut.result()
            if data is None:
                return False, f"下载失败 {path}", len(changes)
            if want is not None and len(data) != want:
                return False, f"字节数不符 {path}（期望 {want} 实得 {len(data)}）", len(changes)
            blobs[path] = data

    target.mkdir(parents=True, exist_ok=True)

    def _local(p):
        return target / (p[len(pre):] if pre else p)

    for p in deletes:
        fp = _local(p)
        if fp.is_file():
            fp.unlink()
        elif fp.is_dir():
            shutil.rmtree(fp, ignore_errors=True)
    for p, data in blobs.items():
        fp = _local(p)
        fp.parent.mkdir(parents=True, exist_ok=True)
        fp.write_bytes(data)
    return True, f"增量 {len(blobs)} 改 / {len(deletes)} 删", len(changes)

def replace_dir(stage: Path, target: Path, keep=()) -> bool:
    """原子替换：备份 .old → 复制 → 恢复 keep 白名单 → 文件数校验 → 删备份。失败回滚。

    keep 项（如本地提取脚本/产物）不在上游 tarball 里，替换后从 .old 拷回，避免被清掉。
    """
    keep_rel = _keep_rel(target, keep)
    if target.exists():
        bak = target.with_name(target.name + ".old")
        shutil.rmtree(bak, ignore_errors=True)
        try:
            target.rename(bak)
        except OSError:
            # 目录被占用（云同步等）时逐文件删除后重建
            shutil.rmtree(target, ignore_errors=True)
            keep_rel = set()
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
    if bak and bak.exists() and keep_rel:
        _restore_keep(bak, target, keep_rel)
    if count_files(target) != count_files(stage) + len(keep_rel):
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

    # local_dir 是固定目录名，不随上游版本变化：上游出新版只改 manifest 的 ref，
    # 自研文档里引用的路径（如 xlwings/xlwings/cli.py）因此长期有效。
    target = SKILL_ROOT / src["local_dir"]
    old_sha = src.get("pinned_sha", "")
    keep = src.get("keep", ())
    channel, n_changed = "incremental", 0

    # 优先增量：只下载 compare 出的变更文件，避免为 1~2 个文件重拉 17MB 全量包。
    # force / 首次同步（无旧 sha 或本地无快照）时无从比对，直接走全量。
    inc_ok = False
    if not force and old_sha and target.exists():
        inc_ok, desc, n_changed = sync_incremental(src, old_sha, head)
        print(f"  [{vid}] {'增量成功：' if inc_ok else '增量不适用（回退全量）：'}{desc}")

    if not inc_ok:
        stage, channel = download_tarball(repo, ref)
        if stage is None:
            print(f"  [{vid}] 全量下载/解压失败（三条通道均不可达）")
            return "failed"
        if count_files(stage) == 0:
            print(f"  [{vid}] 全量下载结果为空，放弃替换")
            return "failed"
        if not replace_dir(stage, target, keep):
            print(f"  [{vid}] 替换失败（可能被占用，已回滚）")
            return "failed"

    if target.exists() and (target / "extract_webextension.py").is_file():
        _run_extractor(target)
    n_files = count_files(target)
    src["pinned_sha"] = head
    src["last_synced"] = date.today().isoformat()
    MANIFEST.write_text(json.dumps(MANIFEST_CFG, ensure_ascii=False, indent=2), encoding="utf-8-sig")

    with open(SYNCLOG, "a", encoding="utf-8") as f:
        if f.tell() == 0:
            f.write("# xlwings 技能上游同步日志\n\n| 日期 | 来源 | 方式 | 新 commit | 规模 |\n|---|---|---|---|---|\n")
        chan = f"{channel}（{n_changed} 处变更）" if channel == "incremental" else channel
        f.write(f"| {date.today().isoformat()} | `{vid}` | {chan} | `{head[:7]}` | {n_files} 个文件 |\n")
    print(f"  [{vid}] 已更新（{channel}，{n_files} 文件）→ {target}")
    return "updated"


def main() -> int:
    args = sys.argv[1:]
    check_only = "--check" in args
    force_ids = [a for a in args[args.index("--force") + 1:] if not a.startswith("--")] \
        if "--force" in args else []
    global MANIFEST_CFG
    MANIFEST_CFG = json.loads(MANIFEST.read_text(encoding="utf-8-sig"))
    stats: dict[str, list[str]] = {"updated": [], "unchanged": [], "failed": []}
    for src in MANIFEST_CFG["sources"]:
        if force_ids and src["id"] not in force_ids:
            continue
        if src.get("type") == "web":
            # web 源（官网镜像）：版本/全文哈希漂移判据，非 git sha
            if check_only:
                info = web_latest(src)
                pinned_ver = src.get("pinned_version", "")
                if info is None:
                    state, res = "远端不可达", "failed"
                elif info["top_version"] != pinned_ver or info["full_sha"] != src.get("pinned_full_sha", ""):
                    state, res = f"有更新（{pinned_ver} → {info['top_version']}）", "updated"
                else:
                    state, res = f"已是最新（{pinned_ver}）", "unchanged"
                print(f"[check] {src['id']} ({src.get('local_dir')}, web): {state}")
            else:
                res = sync_web_one(src, force=(src["id"] in force_ids))
            stats[res].append(src["id"])
            continue
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
