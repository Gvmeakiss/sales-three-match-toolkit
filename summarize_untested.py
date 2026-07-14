#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
汇总各公司 Untested 中「仅订单」「仅订单及发货单」的条数及订单金额。

金额口径：订单行项目金额（VBAP.NETWR 折算本币），按 (VKORG,VBELN,POSNR) 逐行汇总。
过滤条件（与 sales_three_match / three_lists 对齐）：
  - 审计期间：ORDER_YEAR 年 ORDER_MONTH_START~END 月（VBAK.AUDAT/ERDAT）
  - 订单类型：ORDER_TYPE 非空时仅保留 AUART==ORDER_TYPE；ORDER_TYPE_EXCLUDE 中的类型一律剔除（如 AB 取消）
  - 内部交易：售达方 KUNNR 在 get_exclude_sold_to_codes() 内的订单剔除（与三单匹配一致）
"""

import glob
import os
import sys

import pandas as pd

# 使用项目 config
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import (
    OUTPUT_ROOT,
    INPUT_ROOT,
    ORDER_YEAR,
    ORDER_MONTH_START,
    ORDER_MONTH_END,
    ORDER_TYPE,
    ORDER_TYPE_EXCLUDE,
    get_exclude_sold_to_codes,
)

AMT_COL = '订单金额_本币'
SHEETS = ['仅订单', '仅订单及发货单']


def _get_vbeln_filtered(input_root, company):
    """
    从 VBAK 读取，返回通过所有过滤条件的 VBELN 集合。
    条件：审计期间 + ORDER_TYPE/ORDER_TYPE_EXCLUDE + 售达方 KUNNR 不在排除名单（内部交易剔除）
    """
    folder = os.path.join(input_root, company)
    if not os.path.isdir(folder):
        return set()
    vbak_files = glob.glob(os.path.join(folder, 'VBAK_[0-9]*.TXT'))
    if not vbak_files:
        return set()

    keep_cols = ('VBELN', 'AUDAT', 'ERDAT', 'AUART', 'KUNNR')
    dfs = []
    for fp in vbak_files:
        try:
            for enc in ('utf-8', 'gbk', 'gb2312', 'latin-1'):
                try:
                    df = pd.read_csv(
                        fp, sep=r'#\|#', engine='python',
                        encoding=enc, on_bad_lines='skip',
                        usecols=lambda c: c.strip().upper() in keep_cols,
                        dtype=str
                    )
                    break
                except UnicodeDecodeError:
                    continue
            else:
                continue
            df.columns = df.columns.str.strip().str.upper()
            if 'VBELN' not in df.columns:
                continue
            dc = 'AUDAT' if 'AUDAT' in df.columns else ('ERDAT' if 'ERDAT' in df.columns else None)
            if not dc:
                continue
            cols = ['VBELN', dc]
            if 'AUART' in df.columns:
                cols.append('AUART')
            if 'KUNNR' in df.columns:
                cols.append('KUNNR')
            df = df[[c for c in cols if c in df.columns]].copy()
            df.columns = [c if c != dc else '_dt_str' for c in df.columns]
            dfs.append(df)
        except Exception:
            continue
    if not dfs:
        return set()

    df_vbak = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=['VBELN'])
    df_vbak['_dt'] = pd.to_datetime(df_vbak['_dt_str'], format='%Y%m%d', errors='coerce')
    mask = (
        (df_vbak['_dt'].dt.year == ORDER_YEAR) &
        (df_vbak['_dt'].dt.month >= ORDER_MONTH_START) &
        (df_vbak['_dt'].dt.month <= ORDER_MONTH_END)
    )
    df_vbak = df_vbak[mask]
    if 'AUART' in df_vbak.columns:
        exclude = set(x.upper() for x in (ORDER_TYPE_EXCLUDE or []))
        if exclude:
            auart_s = df_vbak['AUART'].fillna('').astype(str).str.strip().str.upper()
            df_vbak = df_vbak[~auart_s.isin(exclude)]
        if ORDER_TYPE is not None:
            auart_match = df_vbak['AUART'].fillna('').astype(str).str.strip().str.upper() == str(ORDER_TYPE).upper()
            df_vbak = df_vbak[auart_match]
    # 剔除内部交易：售达方 KUNNR 在排除名单中的订单（与 sales_three_match / three_lists 一致）
    exclude_codes = get_exclude_sold_to_codes()
    if 'KUNNR' in df_vbak.columns and exclude_codes:
        kunnr_s = df_vbak['KUNNR'].fillna('').astype(str).str.strip()
        kunnr_normalized = kunnr_s.str.lstrip('0').replace('', '0')
        df_vbak = df_vbak[~kunnr_normalized.isin(exclude_codes)]
    vbeln_set = set(df_vbak['VBELN'].astype(str).str.strip().dropna().unique())
    return vbeln_set


def _get_untested_for_company(output_root, company):
    """获取该公司最新的 Untested 文件路径。
    支持 4390 分片：存在 *_分片_1.xlsx 时优先返回（含 仅订单/仅订单及发货单）。"""
    folder = os.path.join(output_root, company)
    if not os.path.isdir(folder):
        return None
    pattern = os.path.join(folder, f'SalesThreeMatchResult_Untested_{company}_*.xlsx')
    matches = [f for f in glob.glob(pattern) if os.path.isfile(f)]
    if not matches:
        return None
    split_1 = [f for f in matches if '_分片_1.xlsx' in f]
    if split_1:
        return max(split_1, key=os.path.getmtime)
    return max(matches, key=os.path.getmtime)


def summarize_untested(output_root=None, input_root=None):
    output_root = output_root or OUTPUT_ROOT
    input_root = input_root or INPUT_ROOT
    # 扫描 OutPut 下所有公司目录
    companies = []
    try:
        for name in sorted(os.listdir(output_root)):
            path = os.path.join(output_root, name)
            if os.path.isdir(path) and len(name) == 4 and name.isdigit():
                companies.append(name)
    except OSError:
        pass

    rows = []
    for company in companies:
        fp = _get_untested_for_company(output_root, company)
        if not fp or not os.path.exists(fp):
            rows.append({
                '公司': company,
                '仅订单_条数': 0,
                '仅订单_订单金额': 0,
                '仅订单及发运单_条数': 0,
                '仅订单及发运单_订单金额': 0,
                '合计_条数': 0,
                '合计_订单金额': 0,
            })
            continue

        # 审计期间 + 订单类型过滤后的 VBELN
        vbeln_filtered = _get_vbeln_filtered(input_root, company)

        counts = {}
        amounts = {}
        for sheet in SHEETS:
            try:
                df = pd.read_excel(fp, sheet_name=sheet)
                # 按审计期间 + 订单类型筛选
                if vbeln_filtered and 'VBELN' in df.columns:
                    vbeln_s = df['VBELN'].astype(str).str.strip()
                    mask = vbeln_s.isin(vbeln_filtered)
                    df = df[mask]
                cnt = len(df)
                amt = 0
                if AMT_COL in df.columns:
                    amt = pd.to_numeric(df[AMT_COL], errors='coerce').fillna(0).sum()
                else:
                    alt = '发票金额_本币' if '发票金额_本币' in df.columns else None
                    if alt:
                        amt = pd.to_numeric(df[alt], errors='coerce').fillna(0).sum()
                counts[sheet] = cnt
                amounts[sheet] = round(float(amt), 2)
            except (ValueError, KeyError):
                counts[sheet] = 0
                amounts[sheet] = 0

        c1 = counts.get('仅订单', 0)
        a1 = amounts.get('仅订单', 0)
        c2 = counts.get('仅订单及发货单', 0)
        a2 = amounts.get('仅订单及发货单', 0)

        rows.append({
            '公司': company,
            '仅订单_条数': c1,
            '仅订单_订单金额': a1,
            '仅订单及发运单_条数': c2,
            '仅订单及发运单_订单金额': a2,
            '合计_条数': c1 + c2,
            '合计_订单金额': round(a1 + a2, 2),
        })

    df_out = pd.DataFrame(rows)
    # 全公司合计行
    total = {
        '公司': '【全公司合计】',
        '仅订单_条数': df_out['仅订单_条数'].sum(),
        '仅订单_订单金额': round(df_out['仅订单_订单金额'].sum(), 2),
        '仅订单及发运单_条数': df_out['仅订单及发运单_条数'].sum(),
        '仅订单及发运单_订单金额': round(df_out['仅订单及发运单_订单金额'].sum(), 2),
        '合计_条数': df_out['合计_条数'].sum(),
        '合计_订单金额': round(df_out['合计_订单金额'].sum(), 2),
    }
    df_out = pd.concat([df_out, pd.DataFrame([total])], ignore_index=True)
    return df_out


if __name__ == '__main__':
    import argparse
    from utils.path_utils import ensure_dir

    parser = argparse.ArgumentParser(description='汇总 Untested 仅订单/仅订单及发运单')
    parser.add_argument('--debug', action='store_true', help='输出 4390/4510 的过滤诊断（VBELN 数、是否含 KUNNR）')
    args = parser.parse_args()

    exclude_str = f', 排除 AUART in {ORDER_TYPE_EXCLUDE}' if ORDER_TYPE_EXCLUDE else ''
    order_str = f', 仅 AUART={ORDER_TYPE}' if ORDER_TYPE else ''
    print(f'金额口径: 订单行项目金额（VBAP.NETWR 本币）')
    print(f'审计期间: {ORDER_YEAR}年{ORDER_MONTH_START}-{ORDER_MONTH_END}月（VBAK.AUDAT/ERDAT）{order_str}{exclude_str}')
    print(f'内部交易: 售达方 KUNNR 在 get_exclude_sold_to_codes() 内的订单已剔除')
    print(f'数据来源: InPut/{{公司}}/VBAK_*.TXT, OutPut/{{公司}}/SalesThreeMatchResult_Untested_*.xlsx')
    if args.debug:
        exc = get_exclude_sold_to_codes()
        print(f'  [DEBUG] 排除售达方数量: {len(exc)}, 示例: {sorted(exc)[:8]}...')
        for comp in ['4390', '4510']:
            v = _get_vbeln_filtered(INPUT_ROOT, comp)
            untested_fp = _get_untested_for_company(OUTPUT_ROOT, comp)
            if untested_fp and v and comp == '4390':
                df = pd.read_excel(untested_fp, sheet_name='仅订单')
                kept = df['VBELN'].astype(str).str.strip().isin(v).sum() if 'VBELN' in df.columns else 0
                print(f'  [DEBUG] {comp} 仅订单: 总行数={len(df)}, 过滤后保留={kept}, vbeln_filtered={len(v)}')
            else:
                print(f'  [DEBUG] {comp} vbeln_filtered={len(v)} (空=无VBAK或全部被筛掉，汇总时不作VBELN过滤)')
    print()
    df = summarize_untested()
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    pd.set_option('display.float_format', lambda x: f'{x:,.2f}')
    print(df.to_string(index=False))

    # 导出 Excel
    out_path = os.path.join(OUTPUT_ROOT, 'Untested_仅订单及发运单_汇总.xlsx')
    ensure_dir(os.path.dirname(out_path))
    df.to_excel(out_path, index=False, sheet_name='Untested汇总')
    print(f'\n已导出: {out_path}')
