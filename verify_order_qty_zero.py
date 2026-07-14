# -*- coding: utf-8 -*-
"""
核查「均有差异」且订单数量为0的行：订单数量是否真的为0

逻辑说明：
- 匹配结果中的 VBELN/POSNR 来自发票的 AUBEL/AUPOS（订单号）或 VBELV/POSNV（交货号，当 AUBEL 空时）
- 当使用 VBELV/POSNV（交货号）时，与订单表 merge 会失败，导致 订单数量=0
- 本脚本从源数据 VBAP、LIPS、VBRP 核对这些行的真实订单数量

用法：
    SALES_DATA_FOLDER=/path/to/InPut/4150 python verify_order_qty_zero.py [详细文件路径]
    不传路径时，自动查找 OutPut/4150 下最新的 差异分析_详细_*.xlsx
"""

import os
import sys
import glob
import pandas as pd
import numpy as np

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from config import (
    DATA_FOLDER,
    PICKLE_FOLDER,
    OUTPUT_ROOT,
    get_company_code,
    get_output_folder,
    KEY_FIELDS_VBAP,
    KEY_FIELDS_LIPS,
    KEY_FIELDS_VBRP,
)
from sales_three_match import read_sd_data, _get_company_from_folder
from config import VBAK_FILE, VBAP_FILE, LIKP_FILE, LIPS_FILE, VBRK_FILE, VBRP_FILE


def _norm_vbeln(s):
    return s.fillna('').astype(str).str.strip()


def _posnr_variants(posnr):
    """SAP POSNR 常为6位，生成多种格式用于查找"""
    s = str(posnr).strip() if pd.notna(posnr) else ''
    if not s or s == 'nan':
        return ['']
    variants = [s]
    try:
        n = int(float(s))
        variants.extend([f'{n:06d}', str(n)])
    except (ValueError, TypeError):
        if s.isdigit():
            variants.append(s.zfill(6))
    return list(dict.fromkeys(variants))


def main():
    detail_path = None
    if len(sys.argv) > 1:
        detail_path = sys.argv[1]
    else:
        output_folder = get_output_folder()
        pattern = os.path.join(output_folder, 'SalesThreeMatchResult_差异分析_*_详细_*.xlsx')
        files = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)
        if not files:
            print('[ERROR] 未找到差异分析详细文件')
            return
        detail_path = files[0]

    if not os.path.exists(detail_path):
        print(f'[ERROR] 文件不存在: {detail_path}')
        return

    company_code = get_company_code()
    print(f'公司: {company_code}')
    print(f'读取: {os.path.basename(detail_path)}')
    df = pd.read_excel(detail_path, sheet_name='Sheet1')
    print(f'总行数: {len(df):,}')

    # 去除 MATNR 前导零宽字符
    if 'MATNR' in df.columns:
        df['MATNR'] = df['MATNR'].fillna('').astype(str).str.replace('\u200b', '', regex=False).str.strip()

    # 筛选：均有差异 且 订单数量≈0
    if '均有差异' not in df.columns:
        df['均有差异'] = False
    ord_qty = pd.to_numeric(df['订单数量'], errors='coerce').fillna(0)
    mask = df['均有差异'] & (ord_qty.abs() < 0.01)
    df_zero = df[mask].copy()
    n_zero = len(df_zero)
    print(f'\n【均有差异 且 订单数量≈0】共 {n_zero:,} 行')

    if n_zero == 0:
        print('无此类行，无需核查')
        return

    # 抽样：最多核查 500 个不同的 (VKORG, VBELN, POSNR)
    keys = df_zero[['VKORG', 'VBELN', 'POSNR']].drop_duplicates()
    sample_size = min(500, len(keys))
    keys_sample = keys.sample(n=sample_size, random_state=42) if len(keys) > sample_size else keys

    # 加载源数据
    data_folder = DATA_FOLDER
    print(f'\n加载源数据 (data_folder={data_folder})...')
    df_vbap = read_sd_data(VBAP_FILE, 'VBAP', data_folder=data_folder, key_fields=KEY_FIELDS_VBAP)
    df_lips = read_sd_data(LIPS_FILE, 'LIPS', data_folder=data_folder, key_fields=KEY_FIELDS_LIPS)
    df_vbrp = read_sd_data(VBRP_FILE, 'VBRP', data_folder=data_folder, key_fields=KEY_FIELDS_VBRP)

    # 规范化
    df_vbap['VBELN'] = _norm_vbeln(df_vbap['VBELN'])
    df_vbap['POSNR'] = df_vbap['POSNR'].fillna('').astype(str).str.strip()
    df_lips['VBELN'] = _norm_vbeln(df_lips['VBELN'])  # LIPS.VBELN = 交货单号
    df_lips['VGBEL'] = _norm_vbeln(df_lips['VGBEL']) if 'VGBEL' in df_lips.columns else ''
    df_lips['VBELV'] = _norm_vbeln(df_lips['VBELV']) if 'VBELV' in df_lips.columns else ''
    df_vbrp['AUBEL'] = _norm_vbeln(df_vbrp['AUBEL']) if 'AUBEL' in df_vbrp.columns else ''
    df_vbrp['VBELV'] = _norm_vbeln(df_vbrp['VBELV']) if 'VBELV' in df_vbrp.columns else ''

    # VBAP: (VBELN, POSNR) -> KLMENG（同一订单行可能有多条物料/批次，需汇总）
    df_vbap['KLMENG_num'] = pd.to_numeric(df_vbap['KLMENG'], errors='coerce').fillna(0)
    vbap_agg = df_vbap.groupby(['VBELN', 'POSNR'])['KLMENG_num'].sum().reset_index()
    vbap_lookup = df_vbap.set_index(['VBELN', 'POSNR'])
    vbap_sum_lookup = vbap_agg.set_index(['VBELN', 'POSNR'])

    # LIPS: VBELN(交货单号) -> VGBEL/VGPOS(订单号/行)
    lips_as_delivery = df_lips.copy()
    lips_as_delivery['VBELN_dlv'] = lips_as_delivery['VBELN']
    lips_as_delivery['POSNR_dlv'] = lips_as_delivery['POSNR'].fillna('').astype(str).str.strip()
    ord_from_dlv = lips_as_delivery['VGBEL'].fillna(lips_as_delivery['VBELV'])
    lips_as_delivery['ORD_VBELN'] = _norm_vbeln(ord_from_dlv)
    lips_as_delivery['ORD_POSNR'] = lips_as_delivery['VGPOS'].fillna(lips_as_delivery['POSNV']).fillna('').astype(str).str.strip()
    lips_dlv_lookup = lips_as_delivery.set_index(['VBELN_dlv', 'POSNR_dlv'])[['ORD_VBELN', 'ORD_POSNR']].drop_duplicates()

    results = []
    for _, row in keys_sample.iterrows():
        vkorg = str(row.get('VKORG', '')).strip()
        vbeln = str(row.get('VBELN', '')).strip()
        posnr = str(row.get('POSNR', '')).strip()
        if not vbeln:
            continue

        # 1) 直接查 VBAP：VBELN 作为订单号，尝试多种 POSNR 格式，用汇总数量判断
        vbap_row = None
        klmeng_sum = 0.0
        for p in _posnr_variants(posnr):
            try:
                klmeng_sum = float(vbap_sum_lookup.loc[(vbeln, p), 'KLMENG_num'])
                vbap_row = vbap_lookup.loc[(vbeln, p)]
                break
            except (KeyError, TypeError):
                pass
        if vbap_row is not None and not isinstance(vbap_row, pd.DataFrame):
            vbap_row = vbap_row.to_frame().T
        if vbap_row is not None and len(vbap_row) > 0:
            if abs(klmeng_sum) >= 1e-6:
                results.append({
                    'VKORG': vkorg, 'VBELN': vbeln, 'POSNR': posnr,
                    '结论': 'VBAP中订单数量非0',
                    'KLMENG_汇总': klmeng_sum,
                    '说明': 'VBAP 中 KLMENG 汇总非0，匹配结果中订单数量为0可能有误',
                })
                continue
            else:
                results.append({
                    'VKORG': vkorg, 'VBELN': vbeln, 'POSNR': posnr,
                    '结论': 'VBAP中数量确实为0',
                    'KLMENG_汇总': klmeng_sum,
                    '说明': 'VBAP 中 KLMENG 汇总为0，订单数量为0正确',
                })
                continue

        # 2) VBELN 可能是交货单号，通过 LIPS 反查订单，尝试多种 POSNR 格式
        ord_info = None
        for p in _posnr_variants(posnr):
            try:
                ord_info = lips_dlv_lookup.loc[(vbeln, p)]
                break
            except KeyError:
                pass
        if ord_info is not None:
            if isinstance(ord_info, pd.Series):
                ord_vbeln = str(ord_info.get('ORD_VBELN', '')).strip()
                ord_posnr = str(ord_info.get('ORD_POSNR', '')).strip()
            else:
                ord_vbeln = str(ord_info['ORD_VBELN'].iloc[0]).strip()
                ord_posnr = str(ord_info['ORD_POSNR'].iloc[0]).strip()
            if ord_vbeln:
                try:
                    vbap_ord = vbap_lookup.loc[(ord_vbeln, ord_posnr)]
                except KeyError:
                    vbap_ord = None
                if vbap_ord is not None:
                    if not isinstance(vbap_ord, pd.DataFrame):
                        vbap_ord = vbap_ord.to_frame().T
                    klmeng = pd.to_numeric(vbap_ord['KLMENG'].iloc[0], errors='coerce')
                    if pd.notna(klmeng) and abs(klmeng) >= 1e-6:
                        results.append({
                            'VKORG': vkorg, 'VBELN': vbeln, 'POSNR': posnr,
                            '结论': 'VBELN为交货号，真实订单数量非0',
                            'KLMENG': klmeng, 'ORD_VBELN': ord_vbeln, 'ORD_POSNR': ord_posnr,
                            '说明': f'匹配键 VBELN={vbeln} 实为交货单号，对应订单 {ord_vbeln}/{ord_posnr} 的 KLMENG={klmeng}',
                        })
                        continue

        # 3) 查 VBRP：该发票行 AUBEL 是否为空（导致用了 VBELV 交货号）
        vbrp_match = df_vbrp[
            (df_vbrp['VBELV'] == vbeln) | (df_vbrp['AUBEL'] == vbeln)
        ]
        aubel_empty = False
        if len(vbrp_match) > 0:
            aubel_vals = vbrp_match['AUBEL'].fillna('').str.strip()
            aubel_empty = (aubel_vals == '').all() or (aubel_vals == '0').all()

        results.append({
            'VKORG': vkorg, 'VBELN': vbeln, 'POSNR': posnr,
            '结论': '未在VBAP/LIPS中找到或数量为0',
            'AUBEL空导致用交货号': aubel_empty,
            '说明': 'VBELN 既不是有效订单号，也无法通过 LIPS 解析到有数量的订单；或发票 AUBEL 为空导致用了交货号匹配',
        })

    # 输出
    res_df = pd.DataFrame(results)
    out_folder = get_output_folder()
    out_path = os.path.join(out_folder, f'订单数量核查_{company_code}.xlsx')
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        res_df.to_excel(w, sheet_name='核查结果', index=False)
        # 汇总
        summary = res_df['结论'].value_counts()
        pd.DataFrame({'结论': summary.index, '数量': summary.values}).to_excel(
            w, sheet_name='汇总', index=False
        )
    print(f'\n[OK] 核查结果已保存: {out_path}')
    print('\n【汇总】')
    for k, v in res_df['结论'].value_counts().items():
        print(f'  {k}: {v} 条')


if __name__ == '__main__':
    main()
