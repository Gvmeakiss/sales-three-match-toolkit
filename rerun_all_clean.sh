#!/bin/bash
# 删除所有三单匹配及差异分析文件，按新标准重新运行
# 用法: ./rerun_all_clean.sh
# 或:   bash rerun_all_clean.sh

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

OUTPUT_ROOT="${OUTPUT_ROOT:-$(python3 -c "import sys; sys.path.insert(0,'.'); from config import OUTPUT_ROOT; print(OUTPUT_ROOT)")}"
OUTPUT_ROOT="$(cd "$OUTPUT_ROOT" 2>/dev/null && pwd || echo "$OUTPUT_ROOT")"

echo "============================================================"
echo "步骤1: 删除三单匹配及差异分析文件"
echo "============================================================"
echo "输出根目录: $OUTPUT_ROOT"

deleted=0
while IFS= read -r -d '' f; do
  rm -f "$f"
  echo "  删除: ${f#$OUTPUT_ROOT/}"
  ((deleted++)) || true
done < <(find "$OUTPUT_ROOT" -maxdepth 2 -name "SalesThreeMatchResult*.xlsx" -print0 2>/dev/null || true)
while IFS= read -r -d '' f; do
  rm -f "$f"
  echo "  删除: ${f#$OUTPUT_ROOT/}"
  ((deleted++)) || true
done < <(find "$OUTPUT_ROOT" -maxdepth 2 -name '~\$SalesThreeMatchResult*.xlsx' -print0 2>/dev/null || true)
echo "已删除 $deleted 个文件"

echo ""
echo "============================================================"
echo "步骤2: 按新标准运行（三单清单->三单匹配->差异分析）"
echo "============================================================"
python3 run_all_by_company.py

echo ""
echo "============================================================"
echo "步骤3: 汇总各公司差异分析"
echo "============================================================"
python3 aggregate_difference_summary.py

echo ""
echo "============================================================"
echo "全部完成!"
echo "============================================================"
