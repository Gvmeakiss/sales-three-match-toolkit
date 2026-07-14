#!/bin/bash
# 4390 三单清单完成后，执行三单匹配及差异分析（使用新代码：SHKZG X 支持、订单发票金额差异等）
# 使用方法：4390 销售发票清单生成完毕后，在 SalesThreeMatchToolkit 目录下执行：
#   bash run_4390_match_diff.sh

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
INPUT_4390="$(cd "$SCRIPT_DIR/../InPut/4390" 2>/dev/null && pwd || echo '')"
if [ -z "$INPUT_4390" ] || [ ! -d "$INPUT_4390" ]; then
    echo "[ERROR] InPut/4390 不存在"
    exit 1
fi

export SALES_DATA_FOLDER="$INPUT_4390"
export SALES_COMPANY_CODE="4390"

echo "=============================================="
echo "4390 三单匹配 + 差异分析（新代码）"
echo "数据目录: $SALES_DATA_FOLDER"
echo "=============================================="

echo ""
echo "[1/2] 三单匹配..."
python3 sales_three_match.py
echo ""

echo "[2/2] 差异分析..."
python3 difference_analysis.py
echo ""

echo "[OK] 4390 处理完成"
ls -la "$SCRIPT_DIR/../OutPut/4390/" 2>/dev/null | head -20
