# -*- coding: utf-8 -*-
"""
从各公司差异分析详细文件中按场景标号提取行，汇总到 OutPut/场景提取/场景{sid}/ 下。
单文件超过 98 万行时分段输出为 场景{sid}_1.xlsx, 场景{sid}_2.xlsx 等。

流式处理：分公司、分文件读取，每读取一段即按场景累加，任一场景达 98w 行即立即写出并清空缓冲；
多线程按公司并行读取，主线程合并并写出。

用法: python3 extract_scenarios_to_output.py
      python3 extract_scenarios_to_output.py --workers 8
      python3 extract_scenarios_to_output.py --companies 4730
"""

import os
import sys
import glob
import re
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import pandas as pd
from config import OUTPUT_ROOT, OUTPUT_PREFIX, EXCEL_MAX_ROWS_PER_FILE
from utils.path_utils import ensure_dir

MAX_ROWS = EXCEL_MAX_ROWS_PER_FILE


def _detail_files(company_folder):
    """返回该公司差异分析详细文件列表，排除 .tmp.xlsx"""
    pattern = os.path.join(company_folder, f'{OUTPUT_PREFIX}_差异分析_*_详细*.xlsx')
    files = [f for f in glob.glob(pattern) if '.tmp.' not in os.path.basename(f)]

    def _key(p):
        b = os.path.basename(p)
        m = re.search(r'_详细_(\d+)\.xlsx$', b)
        return (0, 0) if m is None else (1, int(m.group(1)))

    return sorted(files, key=_key)


def _flush_scenario(sid, buffers, parts, extract_base):
    """将 buffers[sid] 中前 MAX_ROWS 行写出，清空已写部分，返回写出行数"""
    rows = buffers.get(sid, [])
    if len(rows) < MAX_ROWS:
        return 0
    to_write = rows[:MAX_ROWS]
    buffers[sid] = rows[MAX_ROWS:]

    out_dir = os.path.join(extract_base, f'场景{sid}')
    ensure_dir(out_dir)
    parts[sid] = parts.get(sid, 0) + 1
    out_name = f'场景{sid}_{parts[sid]}.xlsx'
    out_path = os.path.join(out_dir, out_name)

    df = pd.DataFrame(to_write)
    df.to_excel(out_path, sheet_name='Sheet1', index=False)
    return len(to_write)


def main():
    parser = argparse.ArgumentParser(description='按场景提取差异分析详细行至 OutPut/场景提取/')
    parser.add_argument('--companies', type=str, help='仅处理指定公司，逗号分隔，如 4730 或 4010,4030')
    args = parser.parse_args()

    output_root = os.path.normpath(os.path.abspath(OUTPUT_ROOT))
    extract_base = os.path.join(output_root, '场景提取')
    ensure_dir(extract_base)

    companies = sorted(
        [n for n in os.listdir(output_root)
         if os.path.isdir(os.path.join(output_root, n)) and n.isdigit() and len(n) == 4]
    )
    if args.companies:
        want = {s.strip() for s in args.companies.split(',') if s.strip()}
        companies = [c for c in companies if c in want]
        print(f'[INFO] 仅处理公司: {sorted(want)}')

    # 按场景标号累积缓冲，达 MAX_ROWS 即写出
    buffers = {}
    parts = {}  # 每个场景已写出分片编号

    for company_code in companies:
        folder = os.path.join(output_root, company_code)
        files = _detail_files(folder)
        if not files:
            continue
        for fp in files:
            try:
                df = pd.read_excel(fp, sheet_name='Sheet1')
            except Exception as e:
                print(f'[SKIP] {company_code} {os.path.basename(fp)}: {e}')
                continue
            if '场景标号' not in df.columns or df.empty:
                continue
            df['公司代码'] = company_code

            for sid, g in df.groupby('场景标号'):
                sid_val = g['场景标号'].iloc[0]
                try:
                    sid_int = int(sid_val) if pd.notna(sid_val) else 0
                except (ValueError, TypeError):
                    sid_int = 0
                if sid_int == 0 or sid_int == 10:
                    continue
                sid = sid_int
                rows = g.to_dict('records')
                buffers.setdefault(sid, []).extend(rows)

                while len(buffers.get(sid, [])) >= MAX_ROWS:
                    n = _flush_scenario(sid, buffers, parts, extract_base)
                    print(f'  [写出] 场景{sid}_{parts[sid]}.xlsx ({n:,} 行)')
        print(f'  已读 {company_code}')

    # 写出各场景剩余缓冲
    for sid in sorted(buffers.keys()):
        rows = buffers[sid]
        if not rows:
            continue
        out_dir = os.path.join(extract_base, f'场景{sid}')
        ensure_dir(out_dir)
        parts[sid] = parts.get(sid, 0) + 1
        out_name = f'场景{sid}.xlsx' if (parts[sid] == 1 and len(rows) <= MAX_ROWS and not any(
            f.startswith(f'场景{sid}_') for f in os.listdir(out_dir) if f.endswith('.xlsx')
        )) else f'场景{sid}_{parts[sid]}.xlsx'
        # 若已有分片则用分片命名
        existing = [f for f in os.listdir(out_dir) if f.endswith('.xlsx')]
        if existing and not any(f == f'场景{sid}.xlsx' for f in existing):
            out_name = f'场景{sid}_{parts[sid]}.xlsx'
        elif parts[sid] > 1:
            out_name = f'场景{sid}_{parts[sid]}.xlsx'
        else:
            out_name = f'场景{sid}.xlsx' if len(rows) <= MAX_ROWS else f'场景{sid}_1.xlsx'

        out_path = os.path.join(out_dir, out_name)
        pd.DataFrame(rows).to_excel(out_path, sheet_name='Sheet1', index=False)
        print(f'  [OK] 场景{sid} {out_name} ({len(rows):,} 行)')

    print(f'[完成] 输出目录: {extract_base}')


if __name__ == '__main__':
    main()
