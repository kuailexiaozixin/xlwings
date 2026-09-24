#!/bin/sh
# verify_imports.sh — 导入完整性扫描
# 检查项目中的所有 .py 文件能否无错误导入
# 用法：cd <project_dir> && dash -c "sh /path/to/verify_imports.sh"

PROJ_DIR="${1:-.}"
cd "$PROJ_DIR" || exit 1
ERR=0

echo "=== 导入完整性扫描: $PROJ_DIR ==="
echo ""

# 检查 .venv 是否存在
if [ ! -d ".venv" ]; then
  echo "WARN: 未找到 .venv，跳过导入测试"
  exit 0
fi

# 激活 venv 并扫描 src/ 目录
SRC_DIR="$PROJ_DIR/src"
if [ ! -d "$SRC_DIR" ]; then
  echo "WARN: 未找到 src/ 目录，跳过"
  exit 0
fi

# 扫描所有 .py 文件
find "$SRC_DIR" -name "*.py" -not -path "*/__pycache__/*" | while read -r pyfile; do
  relpath=$(echo "$pyfile" | sed "s|$PROJ_DIR/||")
  
  # 转换为模块路径
  modname=$(echo "$relpath" | sed 's|/|.|g' | sed 's|\.py$||' | sed 's|\.__init__$||')
  
  # 跳过 main.py（需要 GUI 环境）
  echo "$modname" | grep -q "\.main$" && continue
  
  # 尝试导入
  result=$(python -c "import $modname" 2>&1)
  if [ $? -eq 0 ]; then
    echo "  OK $modname"
  else
    echo "  FAIL $modname: $(echo "$result" | head -1)"
    ERR=$((ERR+1))
  fi
done

echo ""
echo "=== 结果: $ERR 个导入错误 ==="
exit $ERR
