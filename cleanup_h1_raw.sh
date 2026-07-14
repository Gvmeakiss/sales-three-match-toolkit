#!/bin/bash
# 清理 InPut_H1_RAW：删除空文件夹、已解压的 ZIP

set -euo pipefail

RAW="${1:-/Users/aatrox/Desktop/NewHope/InPut_H1_RAW}"

echo "============================================================"
echo "清理目录: ${RAW}"
echo "============================================================"

# 1. 删除空文件夹（自底向上，多次遍历直至无空目录）
empty_count=0
while true; do
  found=0
  while IFS= read -r -d '' d; do
    rmdir "$d" 2>/dev/null && ((found++)) || true
  done < <(find "$RAW" -type d -empty -print0 2>/dev/null)
  (( empty_count += found )) || true
  [[ $found -eq 0 ]] && break
done
echo "[OK] 删除空文件夹: ${empty_count} 个"

# 2. 删除已解压的 ZIP（存在同名目录且目录非空时删除 ZIP）
zip_count=0
while IFS= read -r zip; do
  base=$(basename "$zip" .ZIP)
  base=$(basename "$base" .zip)
  dir="${RAW}/${base}"
  if [[ -d "$dir" ]] && [[ -n "$(ls -A "$dir" 2>/dev/null)" ]]; then
    rm -f "$zip"
    ((zip_count++)) || true
    echo "  已删除(已解压): $(basename "$zip")"
  fi
done < <(find "$RAW" -maxdepth 1 \( -name "*.ZIP" -o -name "*.zip" \) -type f 2>/dev/null)
echo "[OK] 删除已解压压缩包: ${zip_count} 个"

echo "清理完成。"
