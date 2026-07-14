#!/bin/bash
# 执行全年流程：三单清单 -> 三单匹配 -> 差异分析 -> 汇总
# 默认处理所有公司；可选 --start CODE 从指定公司起执行
# 用法：bash run_full_year_pipeline.sh [WORKERS] [START_CODE]
#   例：bash run_full_year_pipeline.sh 1        # 所有公司，1 进程
#       bash run_full_year_pipeline.sh 4 4010   # 从 4010 起，4 进程

set -euo pipefail

ROOT="/Users/aatrox/Desktop/NewHope"
TOOLKIT="${ROOT}/SalesThreeMatchToolkit"
TS="$(date +%Y%m%d_%H%M%S)"
LOG="/tmp/full_year_pipeline_${TS}.log"
WORKERS="${1:-1}"
START_CODE="${2:-}"

echo "============================================================"
echo "[全年流程] 开始"
echo "TOOLKIT   : ${TOOLKIT}"
echo "WORKERS   : ${WORKERS}"
echo "START_CODE: ${START_CODE:-（全部公司）}"
echo "LOG       : ${LOG}"
echo "============================================================"

cd "${TOOLKIT}"

if [ -n "${START_CODE}" ]; then
  python3 run_all_by_company.py --start "${START_CODE}" --workers "${WORKERS}" | tee "${LOG}"
else
  python3 run_all_by_company.py --workers "${WORKERS}" | tee "${LOG}"
fi
python3 aggregate_difference_summary.py | tee -a "${LOG}"

echo "[OK] 全年流程完成。日志: ${LOG}"
