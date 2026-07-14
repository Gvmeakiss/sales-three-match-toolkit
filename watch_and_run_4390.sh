#!/bin/bash
# 监控 4390 三单清单生成，完成时终止 run_all/three_lists，使用新代码执行 三单匹配 + 差异分析
# 使用方法：bash watch_and_run_4390.sh

set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
LISTS_DIR="/Users/aatrox/Desktop/NewHope/OutPut/4390/销售三单清单"
INPUT_4390="/Users/aatrox/Desktop/NewHope/InPut/4390"
POLL_INTERVAL=30  # 秒
STABLE_SEC=90     # 销售发票文件数无新增后视为完成（秒）

echo "=============================================="
echo "监控 4390 三单清单，完成后用新代码执行匹配"
echo "=============================================="

# 等待 三单清单 完成：销售发票存在且 three_lists 退出，或文件数稳定 STABLE_SEC 秒
wait_for_completion() {
    local last_count=0
    local stable_since=$(date +%s)
    while true; do
        if [ ! -d "$LISTS_DIR" ]; then
            echo "[$(date '+%H:%M:%S')] 等待 销售三单清单 目录..."
            sleep $POLL_INTERVAL
            continue
        fi
        local inv_count=$(ls "$LISTS_DIR"/销售发票清单_4390_*.xlsx 2>/dev/null | wc -l | tr -d ' ')
        local order_count=$(ls "$LISTS_DIR"/销售订单清单_4390_*.xlsx 2>/dev/null | wc -l | tr -d ' ')
        local dlv_count=$(ls "$LISTS_DIR"/交货单清单_4390_*.xlsx 2>/dev/null | wc -l | tr -d ' ')
        local three_lists_running=$(pgrep -f "three_lists.py" | wc -l | tr -d ' ')
        echo "[$(date '+%H:%M:%S')] 销售订单:$order_count 交货:$dlv_count 销售发票:$inv_count three_lists:$three_lists_running"
        if [ "$inv_count" -gt 0 ] && [ "$order_count" -gt 0 ] && [ "$dlv_count" -gt 0 ]; then
            if [ "$three_lists_running" -eq 0 ]; then
                echo "[OK] three_lists 已退出，三单清单完成"
                return 0
            fi
            if [ "$inv_count" -eq "$last_count" ]; then
                local now=$(date +%s)
                if [ $((now - stable_since)) -ge $STABLE_SEC ]; then
                    echo "[OK] 销售发票 ${STABLE_SEC}s 无新增，视为完成"
                    return 0
                fi
            else
                last_count=$inv_count
                stable_since=$(date +%s)
            fi
        else
            last_count=0
            stable_since=$(date +%s)
        fi
        sleep $POLL_INTERVAL
    done
}

# 终止 run_all / three_lists
kill_old_processes() {
    echo ""
    echo "[INFO] 终止 run_all_by_company 与 three_lists 进程..."
    pkill -f "run_all_by_company.py --defer 4390" 2>/dev/null || true
    pkill -f "three_lists.py" 2>/dev/null || true
    sleep 3
    if pgrep -f "three_lists.py" >/dev/null 2>&1; then
        echo "[WARN] three_lists 仍在运行，尝试强制终止"
        pkill -9 -f "three_lists.py" 2>/dev/null || true
        sleep 2
    fi
    echo "[OK] 已终止"
}

# 执行 三单匹配 + 差异分析（新代码）
run_match_diff() {
    echo ""
    echo "=============================================="
    echo "4390 三单匹配 + 差异分析（新代码：SHKZG X 等）"
    echo "=============================================="
    export SALES_DATA_FOLDER="$INPUT_4390"
    export SALES_COMPANY_CODE="4390"
    echo "[1/2] 三单匹配..."
    python3 sales_three_match.py
    echo ""
    echo "[2/2] 差异分析..."
    python3 difference_analysis.py
    echo ""
    echo "[OK] 4390 处理完成"
}

wait_for_completion
kill_old_processes
run_match_diff
