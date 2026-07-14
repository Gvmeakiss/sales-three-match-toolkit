"""
销售三单清单生成
分别生成：
1. 销售订单清单
2. 交货单清单
3. 销售发票清单

使用方法：
1. 修改config.py中的配置参数
2. 运行此脚本生成三个清单
   python three_lists.py              # 生成全部（订单+交货+发票）
   python three_lists.py --invoice-only   # 仅生成销售发票清单
"""

import argparse
import os
import pickle
import pandas as pd
import numpy as np
import warnings
import random
import glob
import time
from datetime import datetime
from config import (
    DATA_FOLDER,
    PICKLE_FOLDER,
    LISTS_PICKLE_FOLDER,
    get_exclude_sold_to_codes,
    SAVE_LIST_PKL,
    USE_LIST_PKL,
    ORDER_YEAR,
    ORDER_MONTH_START,
    ORDER_MONTH_END,
    DELIVERY_YEAR,
    DELIVERY_MONTH_START,
    DELIVERY_MONTH_END,
    INVOICE_YEAR,
    INVOICE_MONTH_START,
    INVOICE_MONTH_END,
    ORDER_TYPE,
    MATNR_MAX,
    EXPORT_ORDER_LIST,
    EXPORT_DELIVERY_LIST,
    EXPORT_INVOICE_LIST,
    EXCEL_MAX_ROWS_PER_FILE,
    USE_RANDOM_SUFFIX,
    VBAK_FILE,
    VBAP_FILE,
    LIKP_FILE,
    LIPS_FILE,
    VBRK_FILE,
    VBRP_FILE,
    get_output_folder,
    get_company_code,
    QUICK_TEST_ONE_MONTH,
    BATCH_DEDUP_ENABLED,
)
from utils.DylanTools import kpmg_txt_to_df
from utils.path_utils import ensure_dir

warnings.filterwarnings('ignore')


def print_progress(message, step=None, total=None, start_time=None):
    """打印带时间戳和进度的消息"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if step is not None and total is not None:
        message = f"[{step}/{total}] ({step/total*100:.1f}%) {message}"
    if start_time is not None:
        elapsed = time.time() - start_time
        elapsed_str = f"{elapsed:.1f}秒" if elapsed < 60 else f"{elapsed/60:.1f}分钟"
        message = f"{message} [耗时: {elapsed_str}]"
    print(f"[{timestamp}] {message}")


def print_section(title):
    """打印章节标题"""
    print('\n' + '=' * 60)
    print(title)
    print('=' * 60)


def _get_pkl_path(txt_path, pickle_folder, company_code=None):
    """按公司代码命名 pickle，避免多公司混用"""
    base = os.path.basename(txt_path).replace('.TXT', '.pkl').replace('.txt', '.pkl')
    if company_code:
        base = base.replace('.pkl', f'_{company_code}.pkl')
    return os.path.join(pickle_folder, base)


def _is_pkl_valid(pkl_path, txt_path=None):
    if not os.path.exists(pkl_path):
        return False
    try:
        test_df = pd.read_pickle(pkl_path)
        if test_df.empty:
            return False
        if txt_path and os.path.exists(txt_path):
            if os.path.getmtime(txt_path) > os.path.getmtime(pkl_path):
                return False
        return True
    except Exception:
        return False


def _get_list_pkl_path(list_name, period_key):
    """获取清单 pkl 文件路径。period_key 含公司代码，如 4010_2025_1-12"""
    folder = LISTS_PICKLE_FOLDER
    return os.path.join(folder, f'{list_name}_{period_key}.pkl')


def _is_list_pkl_valid(pkl_path):
    """验证清单 pkl 是否存在且有效"""
    if not os.path.exists(pkl_path):
        return False
    try:
        test_df = pd.read_pickle(pkl_path)
        return not test_df.empty
    except Exception:
        return False


def _save_list_pkl(df, list_name, period_key):
    """保存清单为 pkl 文件。period_key 含公司代码"""
    if df.empty or not SAVE_LIST_PKL:
        return
    ensure_dir(LISTS_PICKLE_FOLDER)
    pkl_path = os.path.join(LISTS_PICKLE_FOLDER, f'{list_name}_{period_key}.pkl')
    df.to_pickle(pkl_path)
    print_progress(f'[OK] 已保存 pkl: {pkl_path} ({len(df):,} 行)')


def _get_company_from_folder(data_folder):
    """从 data_folder 解析公司代码"""
    folder = os.path.normpath(data_folder or DATA_FOLDER)
    company = os.path.basename(folder)
    if company and company.replace(' ', '').isdigit():
        return company
    return get_company_code()


def read_sd_data(filename_pattern, table_name, data_folder=DATA_FOLDER, pickle_folder=PICKLE_FOLDER,
                 vbeln_filter=None, batch_concat_size=5, key_fields=None, all_columns=False):
    """
    读取SD数据文件。pickle 按公司代码命名，避免多公司混用。
    vbeln_filter: 可选，set of VBELN，仅保留 VBELN/VBELV 在该集合中的行（用于 VBAP/LIPS/VBRP 以节省内存）
    batch_concat_size: 多文件时每批合并的文件数，避免一次性 concat 过多导致 OOM
    key_fields: 可选，仅保留关键字段保存/返回。未指定时从 config 按 table_name 获取。
    all_columns: True 时读取 TXT 全部列，不使用 key_fields 裁剪。
    """
    if all_columns:
        key_fields_for_read = None  # kpmg 会保留全部列
        key_fields_for_keep = []    # 不做列裁剪
    else:
        key_fields_for_read = key_fields or getattr(__import__('config'), f'KEY_FIELDS_{table_name}', None)
        key_fields_for_keep = key_fields_for_read
    file_paths = sorted(glob.glob(os.path.join(data_folder, filename_pattern)))
    if not file_paths:
        print(f'[WARNING] 未找到匹配 {filename_pattern} 的文件')
        return pd.DataFrame()

    company_code = _get_company_from_folder(data_folder)
    filter_col = 'VBELN' if table_name in ('VBAK', 'VBAP') else 'VBELV' if table_name in ('LIPS', 'VBRP') else None
    filter_col_alt = 'AUBEL' if table_name == 'VBRP' else ('VGBEL' if table_name == 'LIPS' else None)  # VBRP用AUBEL; LIPS用VGBEL(参考单据=订单号)
    apply_filter = vbeln_filter is not None and len(vbeln_filter) > 0

    def _load_and_filter(file_path):
        pkl_path = _get_pkl_path(file_path, pickle_folder, company_code)
        if _is_pkl_valid(pkl_path, file_path) and not all_columns:
            df = pd.read_pickle(pkl_path)
        else:
            df = kpmg_txt_to_df(file_path, to_pickle=not all_columns, output_folder=pickle_folder, company_code=company_code, key_fields=key_fields_for_read)
        if apply_filter:
            masks = []
            if filter_col and filter_col in df.columns:
                masks.append(df[filter_col].astype(str).str.strip().isin(vbeln_filter))
            if filter_col_alt and filter_col_alt in df.columns:
                masks.append(df[filter_col_alt].astype(str).str.strip().isin(vbeln_filter))
            if masks:
                mask = masks[0]
                for m in masks[1:]:
                    mask = mask | m
                df = df[mask].copy()
        if key_fields_for_keep and len(key_fields_for_keep) > 0:
            keep = [c for c in key_fields_for_keep if c in df.columns]
            if keep:
                df = df[keep].copy()
        return df

    if len(file_paths) == 1:
        df = _load_and_filter(file_paths[0])
        print_progress(f'{table_name} 读取成功! 行数: {len(df):,}')
        return df

    # 多文件：分批读取，使用磁盘临时文件，按规模最小的两个优先合并，控制内存峰值
    import tempfile
    import shutil
    temp_dir = tempfile.mkdtemp(prefix=f'{table_name}_merge_')
    try:
        # (行数, 文件路径) 按行数排序，总是合并最小的两个
        part_sizes = []
        for i in range(0, len(file_paths), batch_concat_size):
            batch = file_paths[i:i + batch_concat_size]
            batch_dfs = [_load_and_filter(fp) for fp in batch]
            part = pd.concat(batch_dfs, ignore_index=True)
            del batch_dfs
            n = len(part)
            tf = os.path.join(temp_dir, f'part_{i}.pkl')
            part.to_pickle(tf)
            part_sizes.append((n, tf))
            del part
        # 归并：每次取行数最小的两个合并
        import heapq
        heapq.heapify(part_sizes)
        merge_idx = 0
        while len(part_sizes) >= 2:
            (n1, p1), (n2, p2) = heapq.heappop(part_sizes), heapq.heappop(part_sizes)
            df1 = pd.read_pickle(p1)
            df2 = pd.read_pickle(p2)
            merged = pd.concat([df1, df2], ignore_index=True)
            del df1, df2
            for pp in [p1, p2]:
                try:
                    os.remove(pp)
                except Exception:
                    pass
            m_path = os.path.join(temp_dir, f'm_{merge_idx}.pkl')
            merged.to_pickle(m_path)
            heapq.heappush(part_sizes, (len(merged), m_path))
            del merged
            merge_idx += 1
        df = pd.read_pickle(part_sizes[0][1]) if part_sizes else pd.DataFrame()
        print_progress(f'{table_name} 合并完成! 总行数: {len(df):,}')
    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass
    return df


def generate_sales_order_list(df_vbak, df_vbap):
    """生成销售订单清单"""
    print_section('步骤1: 生成销售订单清单')
    step_start = time.time()

    if df_vbak.empty or df_vbap.empty:
        print_progress('[WARNING] 订单数据为空', start_time=step_start)
        return pd.DataFrame()

    common = [c for c in ['MANDT', 'VBELN'] if c in df_vbak.columns and c in df_vbap.columns]
    df_order = df_vbap.merge(df_vbak, on=common, how='left')
    df_order.drop(columns=[c for c in df_order.columns if c.endswith('_y')], inplace=True, errors='ignore')
    df_order.columns = [c.replace('_x', '') for c in df_order.columns]

    # 订单金额与数量不再按 SHKZG/RETPO 或发票类型取负，仅用原始数值
    df_order['订单金额_外币'] = df_order['NETWR'].fillna(0)
    wkurs = df_order['WKURS'].fillna(1) if 'WKURS' in df_order.columns else 1
    df_order['订单金额_本币'] = df_order['订单金额_外币'] * wkurs
    # 订单数量：优先 KLMENG（基本单位，与交货/发票一致），否则按 UMVKZ/UMVKN 换算 KWMENG
    # 导出含 订单数量_KWMENG(销售单位)、订单数量_KLMENG(基本单位)、订单数量_计算(基本单位)
    kwmeng = pd.to_numeric(df_order['KWMENG'], errors='coerce').fillna(0)
    klmeng = pd.to_numeric(df_order['KLMENG'], errors='coerce') if 'KLMENG' in df_order.columns else pd.Series([np.nan] * len(df_order))
    umvkz = pd.to_numeric(df_order['UMVKZ'], errors='coerce').fillna(1) if 'UMVKZ' in df_order.columns else 1
    umvkn = pd.to_numeric(df_order['UMVKN'], errors='coerce').replace(0, np.nan).fillna(1) if 'UMVKN' in df_order.columns else 1
    base_qty = np.where(klmeng.notna() & (klmeng.abs() >= 1e-6), klmeng, kwmeng * umvkz / umvkn)
    df_order['订单数量'] = base_qty

    df_order['MATNR'] = df_order['MATNR'].fillna('').astype(str).str.strip()
    mask = (df_order['MATNR'] == '') | (df_order['MATNR'] == '0') | (df_order['MATNR'].str.match(r'^0+$', na=False))
    df_order.loc[mask, 'MATNR'] = ''

    date_col = 'AUDAT' if 'AUDAT' in df_order.columns else 'ERDAT'
    df_order['记录建立日期'] = df_order[date_col]
    df_order['记录建立日期'] = pd.to_datetime(df_order['记录建立日期'].astype(str), format='%Y%m%d', errors='coerce')
    # 不再对销售订单做日期/订单类型/物料号/MATNR_MAX 限制，仅保留售达方排除
    df_order_filtered = df_order.copy()

    # 剔除售达方(KUNNR)在排除名单中的订单（KUNNR 可能带前导零如 000004210，需规范为 4210 再匹配）
    exclude_codes = get_exclude_sold_to_codes()
    if 'KUNNR' in df_order_filtered.columns and exclude_codes:
        kunnr_s = df_order_filtered['KUNNR'].fillna('').astype(str).str.strip()
        kunnr_normalized = kunnr_s.str.lstrip('0').replace('', '0')
        mask_exclude = kunnr_normalized.isin(exclude_codes)
        excluded_count = mask_exclude.sum()
        if excluded_count > 0:
            df_order_filtered = df_order_filtered[~mask_exclude].copy()
            print_progress(f'  [INFO] 已剔除售达方在排除名单中的订单: {excluded_count:,} 行', start_time=step_start)

    fields = [
        'VKORG', 'VBELN', 'POSNR', 'AUART', 'MATNR', 'ARKTX', 'KUNNR', 'VKGRP', 'VKBUR',
        '记录建立日期', 'ERDAT', 'AUDAT', 'MEINS', 'KWMENG', 'KLMENG', 'VRKME', 'NETWR', 'NETPR',
        '订单数量', '订单金额_外币', '订单金额_本币', 'WAERK', 'WKURS', 'WERKS', 'LGORT', 'VSTEL'
    ]
    existing = [f for f in fields if f in df_order_filtered.columns]
    order_list = df_order_filtered[existing].copy()

    rename_order = {
        'VKORG': '销售组织_VKORG', 'VBELN': '销售订单号_VBELN', 'POSNR': '订单行号_POSNR',
        'AUART': '订单类型_AUART', 'MATNR': '物料号_MATNR', 'ARKTX': '物料描述_ARKTX',
        'KUNNR': '售达方_KUNNR', 'VKGRP': '销售组_VKGRP', 'VKBUR': '销售办公室_VKBUR',
        '记录建立日期': '订单创建日期', 'ERDAT': '记录建立日期_ERDAT', 'AUDAT': '订单日期_AUDAT',
        'KWMENG': '订单数量_KWMENG', 'KLMENG': '订单数量_KLMENG', 'VRKME': '销售单位_VRKME', 'NETWR': '净值_NETWR', 'NETPR': '净价_NETPR',
        '订单数量': '订单数量_计算', '订单金额_外币': '订单金额_外币_计算', '订单金额_本币': '订单金额_本币_计算',
        'WAERK': '货币_WAERK', 'WKURS': '汇率_WKURS', 'WERKS': '工厂_WERKS', 'LGORT': '库存地点_LGORT', 'VSTEL': '装运点_VSTEL'
    }
    existing_rename = {k: v for k, v in rename_order.items() if k in order_list.columns}
    order_list = order_list.rename(columns=existing_rename)

    print_progress(f'[OK] 销售订单清单: {len(order_list):,} 行', start_time=step_start)
    return order_list


def generate_delivery_list(df_likp, df_lips):
    """生成交货单清单"""
    print_section('步骤2: 生成交货单清单')
    step_start = time.time()

    if df_lips.empty:
        print_progress('[WARNING] LIPS 数据为空', start_time=step_start)
        return pd.DataFrame()

    # LIKP 为空时直接使用 LIPS，避免 merge(on=[]) 导致结果异常
    if df_likp.empty:
        df_delivery = df_lips.copy()
        df_delivery['LFART'] = ''
    else:
        common = [c for c in ['MANDT', 'VBELN'] if c in df_likp.columns and c in df_lips.columns]
        df_likp_dedup = df_likp.drop_duplicates(subset=common, keep='first')
        lips_key = [c for c in ['MANDT', 'VBELN', 'POSNR'] if c in df_lips.columns]
        df_lips_dedup = df_lips.drop_duplicates(subset=lips_key, keep='first') if len(lips_key) == 3 else df_lips
        df_delivery = df_lips_dedup.merge(df_likp_dedup, on=common, how='left')
        df_delivery.drop(columns=[c for c in df_delivery.columns if c.endswith('_y')], inplace=True, errors='ignore')
        df_delivery.columns = [c.replace('_x', '') for c in df_delivery.columns]

    # 源订单号/行：实际取数中 VBELV 常为空，用 VGBEL/VGPOS 补充（SAP：VGBEL=参考单据号，VGPOS=参考行）
    def _is_empty_or_zero(s):
        sv = s.fillna('').astype(str).str.strip().str.upper()
        return (sv == '') | (sv == 'NAN') | (sv == '0') | (sv == '000') | (sv == '000000')
    if 'VGBEL' in df_delivery.columns:
        if 'VBELV' not in df_delivery.columns:
            df_delivery['VBELV'] = df_delivery['VGBEL']
        else:
            empty = _is_empty_or_zero(df_delivery['VBELV'])
            df_delivery.loc[empty, 'VBELV'] = df_delivery.loc[empty, 'VGBEL']
    if 'VGPOS' in df_delivery.columns:
        if 'POSNV' not in df_delivery.columns:
            df_delivery['POSNV'] = df_delivery['VGPOS']
        else:
            empty = _is_empty_or_zero(df_delivery['POSNV'])
            df_delivery.loc[empty, 'POSNV'] = df_delivery.loc[empty, 'VGPOS']

    if 'VBELV' not in df_delivery.columns or 'POSNV' not in df_delivery.columns:
        print_progress('[WARNING] LIPS 缺少 VBELV/POSNV 且无 VGBEL/VGPOS 可补充', start_time=step_start)
        return pd.DataFrame()

    date_col = 'LFDAT' if 'LFDAT' in df_delivery.columns else 'ERDAT'
    if date_col in df_delivery.columns:
        df_delivery[date_col] = pd.to_datetime(df_delivery[date_col].astype(str), format='%Y%m%d', errors='coerce')
    # 不再对交货单做日期限制

    # SHKZG_VA: LIPS 中为退货项标识，X=退货(数量取负)，S=正常；部分实现沿用 S/H 借贷
    if 'SHKZG' in df_delivery.columns:
        shkzg_upper = df_delivery['SHKZG'].fillna('S').astype(str).str.strip().str.upper()
        shkzg_map = shkzg_upper.map({'S': 1, 'H': -1, 'X': -1}).fillna(1)
    else:
        shkzg_map = 1
    df_delivery['交货数量'] = df_delivery['LFIMG'].fillna(0) * shkzg_map

    # 筛选：仅保留有数量的行（排除批次拆分主行，主行 POSNR 0 开头 LFIMG=0，实际数量在 9 开头的批次子行）
    _lfimg = pd.to_numeric(df_delivery['LFIMG'], errors='coerce').fillna(0)
    df_delivery = df_delivery[(_lfimg.abs() >= 1e-6)].copy()

    # 批次去重：主行有数量时排除其批次子行（config.BATCH_DEDUP_ENABLED=True 时生效；先保留 False 供核验批次结构）
    uepos_col = 'UEPOS' if 'UEPOS' in df_delivery.columns else None
    uecha_col = 'UECHA' if 'UECHA' in df_delivery.columns else None
    if BATCH_DEDUP_ENABLED and (uepos_col or uecha_col):
        def _is_empty_ref(s):
            sv = s.fillna('').astype(str).str.strip()
            return (sv == '') | (sv == '0') | (sv == '000000')
        is_parent = pd.Series(True, index=df_delivery.index)
        if uepos_col:
            is_parent = is_parent & _is_empty_ref(df_delivery[uepos_col])
        if uecha_col:
            is_parent = is_parent & _is_empty_ref(df_delivery[uecha_col])
        parents_with_qty = set(
            zip(df_delivery.loc[is_parent, 'VBELN'].astype(str).str.strip(),
                df_delivery.loc[is_parent, 'POSNR'].astype(str).str.strip())
        )
        if parents_with_qty:
            drop_mask = pd.Series(False, index=df_delivery.index)
            for col in [c for c in [uepos_col, uecha_col] if c]:
                refs = df_delivery[col].fillna('').astype(str).str.strip()
                vbeln_s = df_delivery['VBELN'].astype(str).str.strip()
                drop_mask = drop_mask | (
                    ~_is_empty_ref(df_delivery[col]) &
                    pd.Series([(v, r) in parents_with_qty for v, r in zip(vbeln_s.tolist(), refs.tolist())], index=df_delivery.index)
                )
            n_drop = drop_mask.sum()
            if n_drop > 0:
                df_delivery = df_delivery[~drop_mask].copy()
                print_progress(f'  [INFO] 批次去重：排除 {n_drop:,} 行（主行有数量时的批次子行）', start_time=step_start)

    # 交货数量单位：LFIMG 按 SAP 标准对应 VRKME（销售单位），VRKME 为空时回退到 MEINS（基本单位），均空则保留空字符串
    vrkme_s = df_delivery['VRKME'].fillna('').astype(str).str.strip()
    meins_s = df_delivery['MEINS'].fillna('').astype(str).str.strip()
    df_delivery['交货数量单位'] = np.where(vrkme_s != '', vrkme_s, meins_s)

    fields = [
        'VKORG', 'VBELN', 'POSNR', 'LFART', 'UEPOS', 'UECHA', 'VBELV', 'POSNV', 'MATNR', 'ARKTX', 'KUNNR',
        'LFIMG', '交货数量单位', '交货数量', 'MEINS', 'VRKME', 'SHKZG',
        'WERKS', 'LGORT', 'ERDAT', 'LFDAT', 'CHARG'
    ]
    # 若取数无 UEPOS/UECHA 则从 fields 中移除
    fields = [f for f in fields if f in df_delivery.columns]
    # 4730 单家增加 LGMNG（交货数量基本单位），便于与订单 KLMENG 口径对齐
    if get_company_code() == '4730' and 'LGMNG' in df_delivery.columns:
        fields.insert(fields.index('LFIMG') + 1, 'LGMNG')
    existing = [f for f in fields if f in df_delivery.columns]
    delivery_list = df_delivery[existing].copy()
    # LFART 空值：LIKP 无或为空时导出显示为空字符串，保持一致
    if 'LFART' in delivery_list.columns:
        delivery_list['LFART'] = delivery_list['LFART'].fillna('').astype(str).str.strip()

    rename_dlv = {
        'VKORG': '销售组织_VKORG', 'VBELN': '交货单号_VBELN', 'POSNR': '交货行号_POSNR',
        'LFART': '交货类型_LFART',  # LF=标准交货 RL=退货 等
        'UEPOS': '批次上级行号_UEPOS', 'UECHA': '批次拆分上级项_UECHA',  # 子发运单/批次子行指向主行
        'VBELV': '源订单号_VBELV', 'POSNV': '源订单行号_POSNV', 'MATNR': '物料号_MATNR', 'ARKTX': '物料描述_ARKTX',
        'KUNNR': '客户号_KUNNR',
        'LFIMG': '交货数量_LFIMG',
        'LGMNG': '交货数量_LGMNG',  # 4730 单家输出，基本单位交货数量
        '交货数量单位': '交货数量单位',  # LFIMG 的单位，优先 VRKME，空时回退 MEINS
        '交货数量': '交货数量_计算',
        'MEINS': '基本单位_MEINS', 'VRKME': '销售单位_VRKME',
        'SHKZG': '借贷标识_SHKZG', 'WERKS': '工厂_WERKS', 'LGORT': '库存地点_LGORT', 'ERDAT': '创建日期_ERDAT',
        'LFDAT': '交货日期_LFDAT', 'CHARG': '批次_CHARG'
    }
    existing_rename = {k: v for k, v in rename_dlv.items() if k in delivery_list.columns}
    delivery_list = delivery_list.rename(columns=existing_rename)

    print_progress(f'[OK] 交货单清单: {len(delivery_list):,} 行', start_time=step_start)
    return delivery_list


def generate_delivery_list_full(df_likp, df_lips):
    """生成发运单清单（全字段、无筛选）：LIPS+LIKP 全量合并，保留所有列，不做 vbeln/LFIMG/批次 等筛选"""
    print_section('步骤: 生成发运单清单（全字段、无筛选）')
    step_start = time.time()

    if df_lips.empty:
        print_progress('[WARNING] LIPS 数据为空', start_time=step_start)
        return pd.DataFrame()

    if df_likp.empty:
        df_delivery = df_lips.copy()
        df_delivery['LFART'] = ''
    else:
        common = [c for c in ['MANDT', 'VBELN'] if c in df_likp.columns and c in df_lips.columns]
        df_likp_dedup = df_likp.drop_duplicates(subset=common, keep='first')
        lips_key = [c for c in ['MANDT', 'VBELN', 'POSNR'] if c in df_lips.columns]
        df_lips_dedup = df_lips.drop_duplicates(subset=lips_key, keep='first') if len(lips_key) == 3 else df_lips
        df_delivery = df_lips_dedup.merge(df_likp_dedup, on=common, how='left')
        df_delivery.drop(columns=[c for c in df_delivery.columns if c.endswith('_y')], inplace=True, errors='ignore')
        df_delivery.columns = [c.replace('_x', '') for c in df_delivery.columns]

    def _is_empty_or_zero(s):
        sv = s.fillna('').astype(str).str.strip().str.upper()
        return (sv == '') | (sv == 'NAN') | (sv == '0') | (sv == '000') | (sv == '000000')
    if 'VGBEL' in df_delivery.columns:
        if 'VBELV' not in df_delivery.columns:
            df_delivery['VBELV'] = df_delivery['VGBEL']
        else:
            empty = _is_empty_or_zero(df_delivery['VBELV'])
            df_delivery.loc[empty, 'VBELV'] = df_delivery.loc[empty, 'VGBEL']
    if 'VGPOS' in df_delivery.columns:
        if 'POSNV' not in df_delivery.columns:
            df_delivery['POSNV'] = df_delivery['VGPOS']
        else:
            empty = _is_empty_or_zero(df_delivery['POSNV'])
            df_delivery.loc[empty, 'POSNV'] = df_delivery.loc[empty, 'VGPOS']

    date_col = 'LFDAT' if 'LFDAT' in df_delivery.columns else 'ERDAT'
    if date_col in df_delivery.columns:
        df_delivery[date_col] = pd.to_datetime(df_delivery[date_col].astype(str), format='%Y%m%d', errors='coerce')

    if 'SHKZG' in df_delivery.columns:
        shkzg_map = df_delivery['SHKZG'].fillna('S').astype(str).str.strip().str.upper().map({'S': 1, 'H': -1, 'X': -1}).fillna(1)
    else:
        shkzg_map = 1
    df_delivery['交货数量'] = df_delivery['LFIMG'].fillna(0) * shkzg_map

    vrkme_s = df_delivery['VRKME'].fillna('').astype(str).str.strip()
    meins_s = df_delivery['MEINS'].fillna('').astype(str).str.strip()
    df_delivery['交货数量单位'] = np.where(vrkme_s != '', vrkme_s, meins_s)

    fields = [
        'VKORG', 'VBELN', 'POSNR', 'LFART', 'UEPOS', 'UECHA', 'VBELV', 'POSNV', 'MATNR', 'ARKTX', 'KUNNR',
        'LFIMG', '交货数量单位', '交货数量', 'MEINS', 'VRKME', 'SHKZG',
        'WERKS', 'LGORT', 'ERDAT', 'LFDAT', 'CHARG'
    ]
    fields = [f for f in fields if f in df_delivery.columns]
    if get_company_code() == '4730' and 'LGMNG' in df_delivery.columns:
        fields.insert(fields.index('LFIMG') + 1, 'LGMNG')
    existing = [f for f in fields if f in df_delivery.columns]
    delivery_list = df_delivery[existing].copy()
    if 'LFART' in delivery_list.columns:
        delivery_list['LFART'] = delivery_list['LFART'].fillna('').astype(str).str.strip()

    rename_dlv = {
        'VKORG': '销售组织_VKORG', 'VBELN': '交货单号_VBELN', 'POSNR': '交货行号_POSNR',
        'LFART': '交货类型_LFART', 'UEPOS': '批次上级行号_UEPOS', 'UECHA': '批次拆分上级项_UECHA',
        'VBELV': '源订单号_VBELV', 'POSNV': '源订单行号_POSNV', 'MATNR': '物料号_MATNR', 'ARKTX': '物料描述_ARKTX',
        'KUNNR': '客户号_KUNNR', 'LFIMG': '交货数量_LFIMG', 'LGMNG': '交货数量_LGMNG',
        '交货数量单位': '交货数量单位', '交货数量': '交货数量_计算',
        'MEINS': '基本单位_MEINS', 'VRKME': '销售单位_VRKME',
        'SHKZG': '借贷标识_SHKZG', 'WERKS': '工厂_WERKS', 'LGORT': '库存地点_LGORT', 'ERDAT': '创建日期_ERDAT',
        'LFDAT': '交货日期_LFDAT', 'CHARG': '批次_CHARG'
    }
    existing_rename = {k: v for k, v in rename_dlv.items() if k in delivery_list.columns}
    delivery_list = delivery_list.rename(columns=existing_rename)

    print_progress(f'[OK] 发运单清单: {len(delivery_list):,} 行', start_time=step_start)
    return delivery_list


def generate_invoice_list(df_vbrk, df_vbrp, df_vbap=None):
    """生成销售发票清单"""
    print_section('步骤3: 生成销售发票清单')
    step_start = time.time()

    if df_vbrp.empty:
        print_progress('[WARNING] VBRP 数据为空', start_time=step_start)
        return pd.DataFrame()

    common = [c for c in ['MANDT', 'VBELN'] if c in df_vbrk.columns and c in df_vbrp.columns]
    df_invoice = df_vbrp.merge(df_vbrk, on=common, how='left')
    df_invoice.drop(columns=[c for c in df_invoice.columns if c.endswith('_y')], inplace=True, errors='ignore')
    df_invoice.columns = [c.replace('_x', '') for c in df_invoice.columns]

    # 源订单号/行：实际取数中 VBELV 常为 NaN、POSNV 常为 0，用 AUBEL/AUPOS 补充（SAP：AUBEL=销售订单号，AUPOS=订单行）
    def _is_empty_or_zero(s):
        sv = s.fillna('').astype(str).str.strip().str.upper()
        return (sv == '') | (sv == 'NAN') | (sv == '0') | (sv == '000') | (sv == '000000')
    if 'AUBEL' in df_invoice.columns:
        if 'VBELV' not in df_invoice.columns:
            df_invoice['VBELV'] = df_invoice['AUBEL']
        else:
            empty = _is_empty_or_zero(df_invoice['VBELV'])
            df_invoice.loc[empty, 'VBELV'] = df_invoice.loc[empty, 'AUBEL']
    if 'AUPOS' in df_invoice.columns:
        if 'POSNV' not in df_invoice.columns:
            df_invoice['POSNV'] = df_invoice['AUPOS']
        else:
            empty = _is_empty_or_zero(df_invoice['POSNV'])
            df_invoice.loc[empty, 'POSNV'] = df_invoice.loc[empty, 'AUPOS']

    if 'FKDAT' in df_invoice.columns:
        df_invoice['FKDAT'] = pd.to_datetime(df_invoice['FKDAT'].astype(str), format='%Y%m%d', errors='coerce')
        df_invoice = df_invoice[
            (df_invoice['FKDAT'].dt.year == INVOICE_YEAR) &
            (df_invoice['FKDAT'].dt.month >= INVOICE_MONTH_START) &
            (df_invoice['FKDAT'].dt.month <= INVOICE_MONTH_END)
        ]

    # 剔除售达方(KUNAG)在排除名单中的发票（KUNAG 可能带前导零如 000004210，需规范为 4210 再匹配）
    exclude_codes = get_exclude_sold_to_codes()
    if 'KUNAG' in df_invoice.columns and exclude_codes:
        kunag_s = df_invoice['KUNAG'].fillna('').astype(str).str.strip()
        kunag_normalized = kunag_s.str.lstrip('0').replace('', '0')  # 000004210->4210, 0000->0
        mask_exclude = kunag_normalized.isin(exclude_codes)
        excluded_count = mask_exclude.sum()
        if excluded_count > 0:
            df_invoice = df_invoice[~mask_exclude].copy()
            print_progress(f'  [INFO] 已剔除售达方在排除名单中的发票: {excluded_count:,} 行')

    # SHKZG: S=借方(正), H=贷方(负), X=冲账/负值(部分SAP实现，需对NETWR取负)
    # 未映射值默认按借方处理；冲账发票(X)需取负才能正确反映财务冲销
    if 'SHKZG' in df_invoice.columns:
        shkzg_upper = df_invoice['SHKZG'].fillna('S').astype(str).str.strip().str.upper()
        shkzg_map = shkzg_upper.map({'S': 1, 'H': -1, 'X': -1}).fillna(1)
    else:
        shkzg_map = 1
    df_invoice['发票金额_本币'] = df_invoice['NETWR'].fillna(0) * shkzg_map
    df_invoice['发票数量'] = df_invoice['FKIMG'].fillna(0) * shkzg_map

    # 筛选：仅保留有数量或金额的行（排除 POSNR 尾非0 的批次子行，其 FKIMG/NETWR 均为 0）
    _fkimg = pd.to_numeric(df_invoice['FKIMG'], errors='coerce').fillna(0)
    _netwr = pd.to_numeric(df_invoice['NETWR'], errors='coerce').fillna(0)
    df_invoice = df_invoice[(_fkimg.abs() >= 1e-6) | (_netwr.abs() >= 1e-6)].copy()

    fields = [
        'VKORG', 'VBELN', 'POSNR', 'VBELV', 'POSNV', 'MATNR', 'ARKTX',
        'KUNRG', 'KUNAG', 'KUNNR', 'BUKRS',  # KUNAG=售达方(VBRK), KUNRG=收货方(VBRK)
        'FKART',  # 发票类型（来自 VBRK）
        'FKIMG', 'VRKME', 'NETWR', 'AKKUR', 'MWSBP', '发票数量', '发票金额_本币', 'SHKZG',
        'FKDAT', 'WAERK', 'ERDAT'
    ]
    existing = [f for f in fields if f in df_invoice.columns]
    invoice_list = df_invoice[existing].copy()

    rename_inv = {
        'VKORG': '销售组织_VKORG', 'VBELN': '发票号_VBELN', 'POSNR': '发票行号_POSNR',
        'VBELV': '源订单号_VBELV', 'POSNV': '源订单行号_POSNV', 'MATNR': '物料号_MATNR', 'ARKTX': '物料描述_ARKTX',
        'KUNRG': '收货方_KUNRG', 'KUNAG': '售达方_KUNAG', 'KUNNR': '客户号_KUNNR', 'BUKRS': '公司代码_BUKRS',
        'FKART': '发票类型_FKART',
        'FKIMG': '开票数量_FKIMG', 'VRKME': '销售单位_VRKME', 'NETWR': '净值_NETWR', 'AKKUR': '统计金额_AKKUR',
        'MWSBP': '税额_MWSBP', '发票数量': '发票数量_计算', '发票金额_本币': '发票金额_本币_计算', 'SHKZG': '借贷标识_SHKZG',
        'FKDAT': '开票日期_FKDAT',
        'WAERK': '货币_WAERK', 'ERDAT': '创建日期_ERDAT'
    }
    existing_rename = {k: v for k, v in rename_inv.items() if k in invoice_list.columns}
    invoice_list = invoice_list.rename(columns=existing_rename)

    print_progress(f'[OK] 销售发票清单: {len(invoice_list):,} 行', start_time=step_start)
    return invoice_list


def _get_vbeln_set_and_save(df_vbak, period_key):
    """从 VBAK 提取 VBELN 集合并保存到 pickle 供后续阶段使用（销售订单/交货单不做日期限制，取全部）"""
    if df_vbak is None or df_vbak.empty:
        return None
    if 'VBELN' not in df_vbak.columns:
        return None
    vbeln_set = set(df_vbak['VBELN'].astype(str).str.strip().dropna().unique())
    ensure_dir(LISTS_PICKLE_FOLDER)
    pkl_path = os.path.join(LISTS_PICKLE_FOLDER, f'vbeln_set_{period_key}.pkl')
    with open(pkl_path, 'wb') as f:
        pickle.dump(vbeln_set, f)
    print_progress(f'订单日期范围内 VBELN 数量: {len(vbeln_set):,}，已保存供后续阶段使用')
    return vbeln_set


def _load_vbeln_set(period_key):
    """加载已保存的 vbeln_set，若不存在则返回 None"""
    pkl_path = os.path.join(LISTS_PICKLE_FOLDER, f'vbeln_set_{period_key}.pkl')
    if not os.path.exists(pkl_path):
        return None
    try:
        with open(pkl_path, 'rb') as f:
            vbeln_set = pickle.load(f)
        print_progress(f'已加载 vbeln_set: {len(vbeln_set):,} 个订单号')
        return vbeln_set
    except Exception:
        return None


def main():
    """主函数 - 分阶段执行：每阶段只加载所需表，完成后释放内存"""
    parser = argparse.ArgumentParser(description='销售三单清单生成')
    parser.add_argument('--invoice-only', action='store_true', help='仅生成销售发票清单，跳过订单及交货单')
    parser.add_argument('--delivery-full', action='store_true',
                        help='仅生成发运单清单：全字段、无筛选（无vbeln_filter、无LFIMG过滤、无批次去重）')
    args = parser.parse_args()
    invoice_only = args.invoice_only
    delivery_full = args.delivery_full

    main_start = time.time()
    company_code = get_company_code()
    _quick = QUICK_TEST_ONE_MONTH
    _mo_st, _mo_end = (1, 1) if _quick else (ORDER_MONTH_START, ORDER_MONTH_END)
    period_str = f'{ORDER_YEAR}_{ORDER_MONTH_START}-{ORDER_MONTH_END}月'
    period_key = f'{company_code}_{ORDER_YEAR}_{ORDER_MONTH_START}-{ORDER_MONTH_END}'
    mode_hint = '仅发票' if invoice_only else ('发运单全字段无筛选' if delivery_full else '分阶段执行')
    print_section(f'销售三单清单生成 - {ORDER_YEAR}年{_mo_st}-{_mo_end}月（{company_code}，{mode_hint}）')

    lists_output_folder = os.path.join(get_output_folder(), '销售三单清单')
    ensure_dir(lists_output_folder)

    def _get_or_generate_list(list_name, sheet_name, generator_fn):
        pkl_path = _get_list_pkl_path(list_name, period_key)
        if USE_LIST_PKL and _is_list_pkl_valid(pkl_path):
            print_progress(f'从 pkl 加载 {list_name}...')
            lst_df = pd.read_pickle(pkl_path)
            print_progress(f'[OK] {list_name} 加载完成: {len(lst_df):,} 行')
            return lst_df
        lst_df = generator_fn()
        if not lst_df.empty and SAVE_LIST_PKL:
            _save_list_pkl(lst_df, list_name, period_key)
        return lst_df

    export_tasks = []
    do_order = EXPORT_ORDER_LIST and not invoice_only and not delivery_full
    do_delivery = EXPORT_DELIVERY_LIST and not invoice_only and not delivery_full
    do_invoice = EXPORT_INVOICE_LIST and not delivery_full

    # ---------- delivery-full: 仅发运单，全字段无筛选 ----------
    if delivery_full:
        print_section('发运单清单（全字段、无筛选）')
        df_likp = read_sd_data(LIKP_FILE, 'LIKP', all_columns=True)
        df_lips = read_sd_data(LIPS_FILE, 'LIPS', vbeln_filter=None, all_columns=True, batch_concat_size=3)
        if df_lips is not None and not df_lips.empty:
            lst = generate_delivery_list_full(
                df_likp if df_likp is not None else pd.DataFrame(),
                df_lips if df_lips is not None else pd.DataFrame()
            )
            if not lst.empty:
                export_tasks.append(('交货单清单', '交货单', lst))
        del df_likp, df_lips

    # ---------- 阶段1: 销售订单 ----------
    if do_order:
        print_section('阶段1: 合并销售订单（VBAK + VBAP）')
        df_vbak = read_sd_data(VBAK_FILE, 'VBAK', batch_concat_size=5)
        vbeln_set = _get_vbeln_set_and_save(df_vbak, period_key) if not df_vbak.empty else None
        df_vbap = read_sd_data(VBAP_FILE, 'VBAP', vbeln_filter=vbeln_set, batch_concat_size=2)
        if df_vbak is not None and df_vbap is not None:
            lst = _get_or_generate_list('销售订单清单', '销售订单', lambda: generate_sales_order_list(df_vbak, df_vbap))
            export_tasks.append(('销售订单清单', '销售订单', lst))
        del df_vbak, df_vbap, vbeln_set

    # ---------- 阶段2: 交货单 ----------
    if do_delivery:
        print_section('阶段2: 合并交货单（LIKP + LIPS）')
        vbeln_set = _load_vbeln_set(period_key)
        if vbeln_set is None:
            df_vbak = read_sd_data(VBAK_FILE, 'VBAK', batch_concat_size=5)
            vbeln_set = _get_vbeln_set_and_save(df_vbak, period_key) if not df_vbak.empty else None
            del df_vbak
        df_likp = read_sd_data(LIKP_FILE, 'LIKP')
        df_lips = read_sd_data(LIPS_FILE, 'LIPS', vbeln_filter=vbeln_set, batch_concat_size=3)
        if df_likp is not None or df_lips is not None:
            lst = _get_or_generate_list('交货单清单', '交货单',
                lambda: generate_delivery_list(df_likp if df_likp is not None else pd.DataFrame(), df_lips if df_lips is not None else pd.DataFrame()))
            export_tasks.append(('交货单清单', '交货单', lst))
        del df_likp, df_lips, vbeln_set

    # ---------- 阶段3: 销售发票 ----------
    if do_invoice:
        print_section('阶段3: 合并销售发票（VBRK + VBRP）')
        # 发票清单仅用 VBRK+VBRP，无需 vbeln_set
        df_vbrk = read_sd_data(VBRK_FILE, 'VBRK')
        df_vbrp = read_sd_data(VBRP_FILE, 'VBRP', batch_concat_size=5)
        if df_vbrk is not None and df_vbrp is not None:
            lst = _get_or_generate_list('销售发票清单', '销售发票', lambda: generate_invoice_list(df_vbrk, df_vbrp))
            export_tasks.append(('销售发票清单', '销售发票', lst))
        del df_vbrk, df_vbrp

    if not export_tasks:
        print_progress('[WARNING] 未生成任何清单')
    else:
        print_section('步骤4: 导出清单')
        ran_num = random.randint(1, 1000) if USE_RANDOM_SUFFIX else ''
        suffix = f'_{ran_num}' if ran_num else ''
        ZWSP = '\u200B'

        def add_matnr_zwsp(df_in):
            for c in df_in.columns:
                if 'MATNR' in str(c) or '物料号' in str(c):
                    s = df_in[c].fillna('').astype(str)
                    df_in[c] = s.apply(lambda x: ZWSP + x.lstrip(ZWSP) if str(x).strip() else x)
            return df_in

        def export_safe(df, fpath, sheet_name, desc):
            if df.empty:
                print_progress(f'[WARNING] {desc}为空，跳过')
                return
            ensure_dir(os.path.dirname(fpath))  # Mac 适配：写入前确保目录存在
            add_matnr_zwsp(df.copy())
            total = len(df)
            if total <= EXCEL_MAX_ROWS_PER_FILE:
                df_out = add_matnr_zwsp(df.copy())
                df_out.to_excel(fpath, index=False, sheet_name=sheet_name)
                print_progress(f'[OK] {desc}已导出: {fpath} ({total:,} 行)')
            else:
                n = (total + EXCEL_MAX_ROWS_PER_FILE - 1) // EXCEL_MAX_ROWS_PER_FILE
                base = fpath.replace('.xlsx', '')
                for i in range(n):
                    start, end = i * EXCEL_MAX_ROWS_PER_FILE, min((i + 1) * EXCEL_MAX_ROWS_PER_FILE, total)
                    chunk = add_matnr_zwsp(df.iloc[start:end].copy())
                    chunk.to_excel(f'{base}_{i+1}.xlsx', index=False, sheet_name=sheet_name)
                print_progress(f'[OK] {desc}已导出 {n} 个文件 ({total:,} 行)')

        for list_name, sheet_name, lst_df in export_tasks:
            fpath = os.path.join(lists_output_folder, f'{list_name}_{company_code}_{period_str}{suffix}.xlsx')
            export_safe(lst_df, fpath, sheet_name, list_name)

    print_section('全部完成!')
    print_progress(f'输出文件夹: {lists_output_folder}/', start_time=main_start)
    print('=' * 60)


if __name__ == "__main__":
    main()
