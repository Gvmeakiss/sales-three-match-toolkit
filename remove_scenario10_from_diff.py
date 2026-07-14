# -*- coding: utf-8 -*-
"""
从现有差异分析结果中剔除场景10 的行，提高后续运行效率。

- 差异分析统计：删除 场景明细 中 场景标号==10 的行
- 差异分析详细：删除 场景标号==10 的行，空分片则删除

多线程并行处理各公司。

用法:
  python3 remove_scenario10_from_diff.py
  python3 remove_scenario10_from_diff.py --companies 4010,4030
  python3 remove_scenario10_from_diff.py --workers 8
"""

import os
import sys
import glob
import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import pandas as pd
from config import OUTPUT_ROOT, OUTPUT_PREFIX
from utils.path_utils import ensure_dir


def _parse_val(v):
    """解析可能的数值（含千分位）"""
    if v is None or v == '' or (isinstance(v, float) and (v != v or v == 0)):
        return 0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0


def _process_company(company_code, output_root):
    """处理单家公司的差异分析：统计 + 详细，剔除场景10"""
    folder = os.path.join(output_root, company_code)
    if not os.path.isdir(folder):
        return 0, 0, f'[SKIP] {company_code}: 目录不存在'

    prefix = OUTPUT_PREFIX
    stats_path = os.path.join(folder, f'{prefix}_差异分析_{company_code}.xlsx')
    if not os.path.exists(stats_path):
        return 0, 0, f'[SKIP] {company_code}: 无差异分析统计文件'

    modified_stats = 0
    modified_detail = 0

    try:
        # 1. 统计文件：删除场景10 行
        df = pd.read_excel(stats_path, sheet_name='场景明细')
        if '场景标号' not in df.columns:
            return 0, 0, f'[SKIP] {company_code}: 场景明细缺 场景标号 列'
        before = len(df)
        df = df[df['场景标号'] != 10]
        if len(df) < before:
            df.to_excel(stats_path, sheet_name='场景明细', index=False)
            modified_stats = 1
    except Exception as e:
        return 0, 0, f'[ERROR] {company_code} 统计: {e}'

    # 2. 详细文件：过滤 场景标号!=10
    detail_pattern = os.path.join(folder, f'{prefix}_差异分析_{company_code}_详细*.xlsx')
    detail_files = sorted(glob.glob(detail_pattern))

    for fp in detail_files:
        try:
            df = pd.read_excel(fp, sheet_name='Sheet1')
            if '场景标号' not in df.columns:
                continue
            before = len(df)
            df = df[df['场景标号'] != 10]
            if len(df) < before:
                if df.empty:
                    os.remove(fp)
                    modified_detail += 1
                else:
                    with pd.ExcelWriter(fp, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='Sheet1', index=False)
                    modified_detail += 1
        except Exception as e:
            pass

    msg = f'[OK] {company_code}' + (f' 统计+{modified_detail}详细' if modified_stats or modified_detail else ' 无场景10')
    return modified_stats, modified_detail, msg


def main():
    parser = argparse.ArgumentParser(description='剔除差异分析中场景10的行')
    parser.add_argument('--companies', type=str, help='仅处理指定公司，逗号分隔，如 4010,4030')
    parser.add_argument('--workers', type=int, default=None, help='线程数，默认 min(16, 公司数)')
    args = parser.parse_args()

    output_root = os.path.normpath(os.path.abspath(OUTPUT_ROOT))
    if not os.path.exists(output_root):
        print(f'[ERROR] 输出根目录不存在: {output_root}')
        sys.exit(1)

    companies = []
    for name in sorted(os.listdir(output_root)):
        path = os.path.join(output_root, name)
        if os.path.isdir(path) and name.isdigit() and len(name) == 4:
            companies.append(name)

    if args.companies:
        want = {s.strip() for s in args.companies.split(',') if s.strip()}
        companies = [c for c in companies if c in want]
        print(f'[INFO] --companies 仅处理: {sorted(want)}')

    if not companies:
        print('[WARNING] 无公司目录可处理')
        return

    n_workers = args.workers if args.workers is not None else min(16, len(companies))
    n_workers = max(1, n_workers)
    print(f'共 {len(companies)} 家公司，{n_workers} 线程并行')

    total_stats = 0
    total_detail = 0
    if n_workers <= 1:
        for c in companies:
            s, d, msg = _process_company(c, output_root)
            total_stats += s
            total_detail += d
            print(msg)
    else:
        with ThreadPoolExecutor(max_workers=n_workers) as ex:
            futures = {ex.submit(_process_company, c, output_root): c for c in companies}
            for fut in as_completed(futures):
                s, d, msg = fut.result()
                total_stats += s
                total_detail += d
                print(msg)

    print(f'[OK] 完成: 修改 {total_stats} 个统计文件, {total_detail} 个详细文件')


if __name__ == '__main__':
    main()
