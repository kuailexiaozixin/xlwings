#!/bin/sh
# check_refs.sh — 检查技能文件的引用完整性与模板一致性
# 用法：bash -c "sh scripts/check_refs.sh"
# 退出码：0 = 全部通过；1 = 发现死链/模板不一致
# 注意：所有计数通过临时文件累加，避免管道 subshell 吞掉变量。

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
SHARED="$SKILL_DIR/templates/shared/main.py"
# 模板一致性检查：对比 shared/main.py 与蓝图中的 Python 入口
# 注意：当前无面板型蓝图模板，此处用 userform 蓝图的 bridge_example.py 作占位；新增面板蓝图后应更新此路径
TMPL="$SKILL_DIR/templates/project-blueprints/userform-data-entry-addin/src/python/bridge_example.py"
EXAMPLES_DIR="$SKILL_DIR/examples"

# 用临时文件记录错误数（管道 subshell 无法回写变量）
ERRFILE="$(mktemp)"
echo 0 > "$ERRFILE"
inc_err() { echo $(( $(cat "$ERRFILE") + 1 )) > "$ERRFILE"; }

echo "=== 引用完整性检查 ==="

# 1. 检查所有 .md 文件中的引用路径是否有效
find "$SKILL_DIR" -name "*.md" -not -path "*/.venv/*" -not -path "*/node_modules/*" | while read -r md; do
  refs=$(grep -oP '\]\(\./[^)]+\)' "$md" 2>/dev/null | sed 's/](\.\///; s/)$//')
  for ref in $refs; do
    [ -z "$ref" ] && continue
    target="$(dirname "$md")/$ref"
    if [ ! -f "$target" ] && [ ! -d "$target" ]; then
      echo "  DEAD: $md -> $ref"
      inc_err
    fi
  done
done

echo ""
echo "=== 模板一致性检查 ==="

check_feature() {
  name="$1"; pattern="$2"
  f1=$(grep -c "$pattern" "$SHARED" 2>/dev/null || true)
  f2=$(grep -c "$pattern" "$TMPL" 2>/dev/null || true)
  if [ "$f1" -gt 0 ] && [ "$f2" -gt 0 ]; then
    echo "  OK $name ($f1 / $f2)"
  else
    echo "  MISMATCH $name (shared=$f1, tmpl=$f2)"
    inc_err
  fi
}

echo "  特性                     shared  tmpl"
check_feature "wait_for_server"    "wait_for_server"
check_feature "signal.SIGINT"      "signal.SIGINT"
check_feature "双包导入try/except" "except ImportError"
check_feature "_src路径注入"       'sys.path.insert.*_src'

# 3. 检查所有示例的 main.py 是否与 shared/main.py 一致（比较关键行数）
SHARED_LEN=$(grep -c "def \|import " "$SHARED" 2>/dev/null || true)
for ex_dir in "$EXAMPLES_DIR"/*/; do
  ex_name=$(basename "$ex_dir")
  ex_main="$ex_dir/src/main.py"
  if [ -f "$ex_main" ]; then
    EX_LEN=$(grep -c "def \|import " "$ex_main" 2>/dev/null || true)
    if [ "$EX_LEN" -ge $((SHARED_LEN - 2)) ]; then
      echo "  OK $ex_name ($EX_LEN)"
    else
      echo "  STALE $ex_name (got $EX_LEN, expected ~$SHARED_LEN)"
      inc_err
    fi
  fi
done

echo ""
ERR=$(cat "$ERRFILE")
rm -f "$ERRFILE"
if [ "$ERR" -eq 0 ]; then
  echo "=== 全部通过 ==="
  exit 0
else
  echo "=== $ERR 个问题 ==="
  exit 1
fi
