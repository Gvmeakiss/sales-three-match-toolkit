"""
销售三单匹配 - 核心代码
用于匹配销售订单、交货单、销售发票三单数据

使用方法：
1. 修改config.py中的配置参数
2. 运行此脚本进行三单匹配
3. 结果保存在输出文件夹中
"""

import os
import pandas as pd
import numpy as np
import warnings
import random
import glob
from config import (
    DATA_FOLDER,
    PICKLE_FOLDER,
    get_exclude_sold_to_codes,
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
    MEMORY_SAVE_MODE,
    LARGE_FOLDER_THRESHOLD_GB,
    LARGE_FILE_MB,
    CHUNK_SIZE_FOR_LARGE,
    EXCEL_MAX_ROWS_PER_FILE,
    USE_RANDOM_SUFFIX,
    OUTPUT_PREFIX,
    VBAK_FILE,
    VBAP_FILE,
    LIKP_FILE,
    LIPS_FILE,
    VBRK_FILE,
    VBRP_FILE,
    get_output_folder,
    get_company_code,
    get_effective_memory_save_mode,
)
from utils.DylanTools import kpmg_txt_to_df
from utils.path_utils import ensure_dir

warnings.filterwarnings('ignore')


def _get_pkl_path(txt_path, pickle_folder, period_suffix=None, company_code=None):
    """获取对应的 pkl 文件路径；company_code 按公司命名避免混用；period_suffix 如 '_2025_1-12' 用于日期过滤后的缓存"""
    base = os.path.basename(txt_path).replace('.TXT', '.pkl').replace('.txt', '.pkl')
    if company_code:
        base = base.replace('.pkl', f'_{company_code}.pkl')
    if period_suffix:
        base = base.replace('.pkl', f'{period_suffix}.pkl')
    return os.path.join(pickle_folder, base)


def _is_pkl_valid(pkl_path, txt_path=None):
    """验证 pkl 文件是否存在且有效"""
    if not os.path.exists(pkl_path):
        return False
    try:
        test_df = pd.read_pickle(pkl_path)
        if test_df.empty:
            print(f'        [WARNING] pkl 文件为空，将重新生成')
            return False
        if txt_path and os.path.exists(txt_path):
            pkl_mtime = os.path.getmtime(pkl_path)
            txt_mtime = os.path.getmtime(txt_path)
            if txt_mtime > pkl_mtime:
                print(f'        [INFO] TXT 文件比 pkl 新，将重新生成 pkl')
                return False
        return True
    except Exception as e:
        print(f'        [WARNING] pkl 文件损坏 ({e})，将重新生成')
        return False


def _get_company_from_folder(data_folder):
    """从 data_folder 解析公司代码（如 InPut/4390 -> 4390），非数字目录时回退 get_company_code"""
    folder = os.path.normpath(data_folder or DATA_FOLDER)
    company = os.path.basename(folder)
    if company and company.replace(' ', '').isdigit():
        return company
    return get_company_code()


def read_sd_data(filename_pattern, table_name, data_folder=DATA_FOLDER, pickle_folder=PICKLE_FOLDER,
                 key_fields=None, filter_date_col=None, chunksize=None):
    """
    读取SD数据文件（支持通配符匹配）
    key_fields: 仅保留指定列，减少内存
    filter_date_col: 分块读取时按该日期列过滤（如 LFDAT、FKDAT）
    chunksize: 大表分块大小；None 时用 config.CHUNK_SIZE_FOR_LARGE
    pickle 按公司代码命名（如 VBAK_001_4390.pkl），避免多公司混用
    """
    file_paths = glob.glob(os.path.join(data_folder, filename_pattern))

    if not file_paths:
        print(f'[WARNING] 未找到匹配 {filename_pattern} 的文件')
        return pd.DataFrame()

    company_code = _get_company_from_folder(data_folder)
    mem_save = get_effective_memory_save_mode(data_folder)
    large_mb = LARGE_FILE_MB
    chunk_sz = chunksize or CHUNK_SIZE_FOR_LARGE
    kf = key_fields
    if kf is None:
        import config as _cfg
        kf = getattr(_cfg, f'KEY_FIELDS_{table_name}', None)
    yr = DELIVERY_YEAR if filter_date_col == 'LFDAT' else INVOICE_YEAR
    ms = DELIVERY_MONTH_START if filter_date_col == 'LFDAT' else INVOICE_MONTH_START
    me = DELIVERY_MONTH_END if filter_date_col == 'LFDAT' else INVOICE_MONTH_END

    def _read_one(fp):
        force_chunked = mem_save or (large_mb and os.path.getsize(fp) > large_mb * 1024 * 1024)
        period_suffix = f'_{yr}_{ms}-{me}' if (filter_date_col and force_chunked and yr) else None
        pkl_path = _get_pkl_path(fp, pickle_folder, period_suffix, company_code)
        if _is_pkl_valid(pkl_path, fp):
            print(f'读取 {table_name} pickle 文件...')
            return pd.read_pickle(pkl_path)
        return kpmg_txt_to_df(
            fp, to_pickle=True, output_folder=pickle_folder, company_code=company_code,
            chunksize=chunk_sz, key_fields=kf,
            filter_date_col=filter_date_col if force_chunked else None,
            filter_year=yr if filter_date_col and force_chunked else None,
            filter_month_start=ms if filter_date_col and force_chunked else None,
            filter_month_end=me if filter_date_col and force_chunked else None,
            force_chunked=force_chunked
        )

    if len(file_paths) == 1:
        df = _read_one(file_paths[0])
        print(f'[OK] {table_name} 读取成功! 行数: {len(df):,}')
        return df

    print(f'找到 {len(file_paths)} 个 {table_name} 文件，正在合并...')
    df_list = []
    for fp in file_paths:
        df_part = _read_one(fp)
        df_list.append(df_part)
        print(f'        已读取: {os.path.basename(fp)} ({len(df_part):,} 行)')

    df = pd.concat(df_list, ignore_index=True)
    print(f'[OK] {table_name} 合并完成! 总行数: {len(df):,}')
    return df


def process_order_data(df_vbak, df_vbap):
    """处理销售订单数据"""
    print('\n正在合并订单数据...')
    if df_vbak.empty or df_vbap.empty:
        print('[WARNING] 订单数据为空')
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 与 three_lists 一致：不做订单日期筛选，取全部订单
    merge_on = ['MANDT', 'VBELN']
    common = [c for c in merge_on if c in df_vbak.columns and c in df_vbap.columns]
    if not common:
        common = ['VBELN'] if 'VBELN' in df_vbak.columns and 'VBELN' in df_vbap.columns else ['MANDT', 'VBELN']

    df_order = df_vbap.merge(df_vbak, on=common, how='left')

    # 处理列名冲突
    df_order.drop(columns=[col for col in df_order.columns if col.endswith('_y')], inplace=True, errors='ignore')
    df_order.columns = [col.replace('_x', '') for col in df_order.columns]

    # 订单金额与数量不再按 SHKZG/RETPO 或发票类型取负，仅用原始数值
    df_order['订单金额_外币'] = df_order['NETWR'].fillna(0)
    wkurs = df_order['WKURS'].fillna(1) if 'WKURS' in df_order.columns else 1
    df_order['订单金额_本币'] = df_order['订单金额_外币'] * wkurs
    # 订单数量：优先 KLMENG（基本单位，与交货/发票一致），否则按 UMVKZ/UMVKN 换算 KWMENG
    kwmeng = pd.to_numeric(df_order['KWMENG'], errors='coerce').fillna(0)
    klmeng = pd.to_numeric(df_order['KLMENG'], errors='coerce') if 'KLMENG' in df_order.columns else pd.Series([np.nan] * len(df_order))
    umvkz = pd.to_numeric(df_order['UMVKZ'], errors='coerce').fillna(1) if 'UMVKZ' in df_order.columns else 1
    umvkn = pd.to_numeric(df_order['UMVKN'], errors='coerce').replace(0, np.nan).fillna(1) if 'UMVKN' in df_order.columns else 1
    base_qty = np.where(klmeng.notna() & (klmeng.abs() >= 1e-6), klmeng, kwmeng * umvkz / umvkn)
    df_order['订单数量'] = base_qty

    # 物料号保持字符串
    df_order['MATNR'] = df_order['MATNR'].fillna('').astype(str).str.strip()
    mask_zero = (df_order['MATNR'] == '') | (df_order['MATNR'] == '0') | (df_order['MATNR'].str.match(r'^0+$', na=False))
    df_order.loc[mask_zero, 'MATNR'] = ''

    # 填充空值
    for col in ['VBELN', 'POSNR', 'VKORG']:
        if col in df_order.columns:
            df_order[col] = df_order[col].fillna('')

    # 日期字段（优先AUDAT，否则ERDAT）
    date_col = 'AUDAT' if 'AUDAT' in df_order.columns else 'ERDAT'
    df_order['记录建立日期'] = df_order[date_col]

    # 不再对订单做日期限制，与 three_lists 清单生成逻辑一致
    df_order['记录建立日期'] = pd.to_datetime(df_order['记录建立日期'].astype(str), format='%Y%m%d', errors='coerce')
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
            print(f'  [INFO] 已剔除售达方在排除名单中的订单: {excluded_count:,} 行')

    print(f'[OK] 筛选后订单数据: {len(df_order_filtered):,} 行')

    # 物料映射
    matnr_cols = ['MATNR', 'VKORG', 'VBELN', 'POSNR']
    arktx_col = 'ARKTX' if 'ARKTX' in df_order_filtered.columns else None
    if arktx_col:
        matnr_cols.append(arktx_col)
    matnr_mapper = df_order_filtered[[c for c in matnr_cols if c in df_order_filtered.columns]].drop_duplicates(
        [c for c in ['MATNR', 'VKORG', 'VBELN', 'POSNR'] if c in df_order_filtered.columns]
    )

    # 订单汇总（按订单行）
    pivot_order = df_order_filtered.pivot_table(
        index=['VKORG', 'VBELN', 'POSNR', 'MATNR'],
        values=['订单金额_外币', '订单金额_本币', '订单数量'],
        aggfunc='sum'
    ).reset_index()
    # 补充 AUART 供导出显示订单类型，与 three_lists 一致
    if 'AUART' in df_order_filtered.columns:
        auart_map = df_order_filtered[['VKORG', 'VBELN', 'POSNR', 'AUART']].drop_duplicates()
        pivot_order = pivot_order.merge(auart_map, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    # 不再对订单做 MATNR 非空/ORDER_TYPE/MATNR_MAX 限制，与 three_lists 清单生成逻辑一致
    print(f'[OK] 订单汇总表: {len(pivot_order):,} 行')
    return pivot_order, matnr_mapper, df_order_filtered


def process_delivery_data(df_likp, df_lips, date_prefiltered=False, vbeln_filter=None):
    """处理交货单数据。date_prefiltered=True 时跳过日期筛选（LIPS 已在读取时按 LFDAT 过滤）。
    vbeln_filter: 与发票清单一致，仅保留 VBELV 或 VGBEL 在订单号集合中的行。"""
    print('\n正在合并交货单数据...')
    if df_lips.empty:
        print('[WARNING] LIPS 数据为空，无法处理交货单')
        return pd.DataFrame(), pd.DataFrame()
    if df_likp.empty:
        print('[WARNING] LIKP 数据为空，无法获取 VKORG 等抬头信息，交货单无法正确关联到订单。'
              '请检查 DATA_FOLDER 中是否存在 *LIKP*.TXT 文件。')
        return pd.DataFrame(), pd.DataFrame()

    merge_on = ['MANDT', 'VBELN']
    common = [c for c in merge_on if c in df_likp.columns and c in df_lips.columns]
    if not common:
        common = ['VBELN'] if 'VBELN' in df_likp.columns and 'VBELN' in df_lips.columns else ['MANDT', 'VBELN']

    df_likp_dedup = df_likp.drop_duplicates(subset=common, keep='first')
    lips_key = [c for c in ['MANDT', 'VBELN', 'POSNR'] if c in df_lips.columns]
    df_lips_dedup = df_lips.drop_duplicates(subset=lips_key, keep='first') if len(lips_key) == 3 else df_lips
    df_delivery = df_lips_dedup.merge(df_likp_dedup, on=common, how='left')

    df_delivery.drop(columns=[col for col in df_delivery.columns if col.endswith('_y')], inplace=True, errors='ignore')
    df_delivery.columns = [col.replace('_x', '') for col in df_delivery.columns]

    # vbeln_filter：与发票清单一致，仅保留 VBELV 或 VGBEL 在订单号集合中的行
    if vbeln_filter and len(vbeln_filter) > 0:
        mask = pd.Series([False] * len(df_delivery), index=df_delivery.index)
        if 'VBELV' in df_delivery.columns:
            mask = mask | df_delivery['VBELV'].astype(str).str.strip().isin(vbeln_filter)
        if 'VGBEL' in df_delivery.columns:
            mask = mask | df_delivery['VGBEL'].astype(str).str.strip().isin(vbeln_filter)
        before = len(df_delivery)
        df_delivery = df_delivery[mask].copy()
        print(f'  [INFO] vbeln_filter 后 LIPS: {before:,} -> {len(df_delivery):,} 行')

    # 源订单号/行：实际取数中 LIPS.VBELV 常为空，用 VGBEL/VGPOS 补充（SAP：VGBEL=参考单据号）
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

    # 必须有 VBELV, POSNV 关联到订单
    if 'VBELV' not in df_delivery.columns or 'POSNV' not in df_delivery.columns:
        print('[WARNING] LIPS 缺少 VBELV/POSNV 且无 VGBEL/VGPOS 可补充')
        return pd.DataFrame(), pd.DataFrame()

    # 不再对交货单做日期限制，与 three_lists 清单生成逻辑一致
    date_col = 'LFDAT' if 'LFDAT' in df_delivery.columns else 'ERDAT'
    if date_col in df_delivery.columns and not date_prefiltered:
        df_delivery[date_col] = pd.to_datetime(df_delivery[date_col].astype(str), format='%Y%m%d', errors='coerce')

    # SHKZG_VA: LIPS 中为退货项标识，X=退货(数量取负)，S=正常；部分实现沿用 S/H 借贷
    if 'SHKZG' in df_delivery.columns:
        shkzg_upper = df_delivery['SHKZG'].fillna('S').astype(str).str.strip().str.upper()
        shkzg_map = shkzg_upper.map({'S': 1, 'H': -1, 'X': -1}).fillna(1)
    else:
        shkzg_map = 1
    # 交货数量：统一用基本单位(MEINS)，与订单KLMENG一致。优先 LGMNG；LGMNG 空/0 时用 LFIMG×UMVKZ/UMVKN 换算
    # 当销售单位≠基本单位时 LFIMG 会偏大（如 1基本单位=4销售单位 则 LFIMG=4×LGMNG），导致场景3误判
    lfimg = pd.to_numeric(df_delivery['LFIMG'], errors='coerce').fillna(0)
    if 'LGMNG' in df_delivery.columns:
        lgmng = pd.to_numeric(df_delivery['LGMNG'], errors='coerce').fillna(0)
        use_lgmng = lgmng.abs() >= 1e-6
        if ('UMVKZ' in df_delivery.columns and 'UMVKN' in df_delivery.columns) and not use_lgmng.all():
            umvkz = pd.to_numeric(df_delivery['UMVKZ'], errors='coerce').replace(0, np.nan).fillna(1)
            umvkn = pd.to_numeric(df_delivery['UMVKN'], errors='coerce').replace(0, np.nan).fillna(1)
            # 换算方向：Sales = Base × (UMVKZ/UMVKN)，故 Base = Sales × UMVKN/UMVKZ
            lfimg_to_base = lfimg * umvkn / umvkz
            qty_raw = np.where(use_lgmng, lgmng, lfimg_to_base)
        else:
            qty_raw = np.where(use_lgmng, lgmng, lfimg)
    else:
        qty_raw = lfimg
        print(f'  [INFO] 交货数量使用 LFIMG(销售单位)；取数无 LGMNG，建议 SAP 取数增加 LIPS.LGMNG 以与订单基本单位一致')
    df_delivery['交货数量'] = qty_raw * shkzg_map

    # 与 three_lists 一致：仅保留有数量的行（排除批次拆分主行等 LFIMG=0 的行）
    _lfimg = pd.to_numeric(df_delivery['LFIMG'], errors='coerce').fillna(0)
    df_delivery = df_delivery[(_lfimg.abs() >= 1e-6)].copy()

    # 批次去重：主行有数量时排除其批次子行，避免 pivot 汇总时重复计算（与 three_lists 一致）
    uepos_col = 'UEPOS' if 'UEPOS' in df_delivery.columns else None
    uecha_col = 'UECHA' if 'UECHA' in df_delivery.columns else None
    if uepos_col or uecha_col:
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
                print(f'  [INFO] 批次去重：排除 {n_drop:,} 行（主行有数量时的批次子行）')

    # 注：LIPS 交货单通常不维护金额，故不汇总交货金额，避免误导

    # 按订单行 (VKORG, VBELV, POSNV) 汇总，映射为 VBELN, POSNR
    df_delivery['VBELN_ORD'] = df_delivery['VBELV']
    df_delivery['POSNR_ORD'] = df_delivery['POSNV']

    pivot_delivery = df_delivery.pivot_table(
        index=['VKORG', 'VBELN_ORD', 'POSNR_ORD'],
        values=['交货数量'],
        aggfunc='sum'
    ).reset_index()
    pivot_delivery = pivot_delivery.rename(columns={'VBELN_ORD': 'VBELN', 'POSNR_ORD': 'POSNR'})

    # 聚合交货单号（LIPS.VBELN 为交货单号）
    dlv_num_col = [c for c in ['VBELN'] if c in df_delivery.columns][:1]
    if dlv_num_col:
        agg_dlv = df_delivery.groupby(['VKORG', 'VBELV', 'POSNV'])[dlv_num_col[0]].agg(
            lambda x: ', '.join(sorted(set(str(v).strip() for v in x if pd.notna(v) and str(v).strip())))
        ).reset_index().rename(columns={dlv_num_col[0]: '聚合交货单号', 'VBELV': 'VBELN', 'POSNV': 'POSNR'})
        pivot_delivery = pivot_delivery.merge(agg_dlv, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    print(f'[OK] 交货汇总表: {len(pivot_delivery):,} 行')
    return pivot_delivery, df_delivery


def process_invoice_data(df_vbrk, df_vbrp, date_prefiltered=False, vbeln_filter=None):
    """处理销售发票数据。date_prefiltered=True 时跳过日期筛选（VBRP 已在读取时按 FKDAT 过滤）。
    vbeln_filter: 与发票清单一致，仅保留 VBELV 或 AUBEL 在订单号集合中的行。"""
    print('\n正在合并发票数据...')
    if df_vbrp.empty:
        print('[WARNING] VBRP 数据为空，无法处理销售发票')
        return pd.DataFrame()
    if df_vbrk.empty:
        print('[WARNING] VBRK 数据为空，无法获取 VKORG 等抬头信息，发票无法正确关联到订单。'
              '请检查 DATA_FOLDER 中是否存在 *VBRK*.TXT 文件。')
        return pd.DataFrame()

    merge_on = ['MANDT', 'VBELN']
    common = [c for c in merge_on if c in df_vbrk.columns and c in df_vbrp.columns]
    if not common:
        common = ['VBELN'] if 'VBELN' in df_vbrk.columns and 'VBELN' in df_vbrp.columns else ['MANDT', 'VBELN']

    df_invoice = df_vbrp.merge(df_vbrk, on=common, how='left')

    df_invoice.drop(columns=[col for col in df_invoice.columns if col.endswith('_y')], inplace=True, errors='ignore')
    df_invoice.columns = [col.replace('_x', '') for col in df_invoice.columns]

    # vbeln_filter：与发票清单一致，仅保留 VBELV 或 AUBEL 在订单号集合中的行
    if vbeln_filter and len(vbeln_filter) > 0:
        mask = pd.Series([False] * len(df_invoice), index=df_invoice.index)
        if 'VBELV' in df_invoice.columns:
            mask = mask | df_invoice['VBELV'].astype(str).str.strip().isin(vbeln_filter)
        if 'AUBEL' in df_invoice.columns:
            mask = mask | df_invoice['AUBEL'].astype(str).str.strip().isin(vbeln_filter)
        before = len(df_invoice)
        df_invoice = df_invoice[mask].copy()
        print(f'  [INFO] vbeln_filter 后 VBRP: {before:,} -> {len(df_invoice):,} 行')

    # 订单关联键：VBRP.VBELV/POSNV=交货号，AUBEL/AUPOS=订单号。与订单/交货匹配必须用订单号，故优先 AUBEL/AUPOS
    def _is_empty_or_zero(s):
        sv = s.fillna('').astype(str).str.strip().str.upper()
        return (sv == '') | (sv == 'NAN') | (sv == '0') | (sv == '000') | (sv == '000000')
    if 'AUBEL' in df_invoice.columns and 'AUPOS' in df_invoice.columns:
        # 优先用 AUBEL/AUPOS（订单号），空时才用 VBELV/POSNV（交货号），避免用交货号匹配订单导致订单数量/金额为 0
        df_invoice['VBELN_ORD'] = df_invoice['AUBEL']
        df_invoice['POSNR_ORD'] = df_invoice['AUPOS']
        empty_ord = _is_empty_or_zero(df_invoice['VBELN_ORD']) | _is_empty_or_zero(df_invoice['POSNR_ORD'])
        if 'VBELV' in df_invoice.columns and 'POSNV' in df_invoice.columns:
            df_invoice.loc[empty_ord, 'VBELN_ORD'] = df_invoice.loc[empty_ord, 'VBELV']
            df_invoice.loc[empty_ord, 'POSNR_ORD'] = df_invoice.loc[empty_ord, 'POSNV']
    elif 'VBELV' in df_invoice.columns and 'POSNV' in df_invoice.columns:
        df_invoice['VBELN_ORD'] = df_invoice['VBELV']
        df_invoice['POSNR_ORD'] = df_invoice['POSNV']
    else:
        print('[WARNING] VBRP 缺少 AUBEL/AUPOS 且无 VBELV/POSNV 可补充')
        return pd.DataFrame()

    # 筛选日期（date_prefiltered 时 VBRP 已在读取时过滤，避免 26M 行二次筛选导致 7+ GiB OOM）
    if not date_prefiltered and 'FKDAT' in df_invoice.columns:
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
            print(f'  [INFO] 已剔除售达方在排除名单中的发票: {excluded_count:,} 行')

    # SHKZG: S=借方(正), H=贷方(负), X=冲账/负值(部分SAP实现，需对NETWR取负)
    if 'SHKZG' in df_invoice.columns:
        shkzg_upper = df_invoice['SHKZG'].fillna('S').astype(str).str.strip().str.upper()
        shkzg_map = shkzg_upper.map({'S': 1, 'H': -1, 'X': -1}).fillna(1)
    else:
        shkzg_map = 1
    df_invoice['发票金额_本币'] = df_invoice['NETWR'].fillna(0) * shkzg_map
    # 发票数量：统一用基本单位(MEINS)。优先 FKLMG；FKLMG 空/0 时用 FKIMG×UMVKZ/UMVKN 换算
    fkimg = pd.to_numeric(df_invoice['FKIMG'], errors='coerce').fillna(0)
    if 'FKLMG' in df_invoice.columns:
        fklmg = pd.to_numeric(df_invoice['FKLMG'], errors='coerce').fillna(0)
        use_fklmg = fklmg.abs() >= 1e-6
        if ('UMVKZ' in df_invoice.columns and 'UMVKN' in df_invoice.columns) and not use_fklmg.all():
            umvkz = pd.to_numeric(df_invoice['UMVKZ'], errors='coerce').replace(0, np.nan).fillna(1)
            umvkn = pd.to_numeric(df_invoice['UMVKN'], errors='coerce').replace(0, np.nan).fillna(1)
            # 换算方向：Sales = Base × (UMVKZ/UMVKN)，故 Base = Sales × UMVKN/UMVKZ
            fkimg_to_base = fkimg * umvkn / umvkz
            inv_qty_raw = np.where(use_fklmg, fklmg, fkimg_to_base)
        else:
            inv_qty_raw = np.where(use_fklmg, fklmg, fkimg)
    else:
        inv_qty_raw = fkimg
        print(f'  [INFO] 发票数量使用 FKIMG(销售单位)；取数无 FKLMG，建议 SAP 取数增加 VBRP.FKLMG 以与订单基本单位一致')
    df_invoice['发票数量'] = inv_qty_raw * shkzg_map

    # 与发票清单保持一致：仅保留有数量或金额的行（排除 FKIMG/NETWR 均为 0 的批次子行等）
    _fkimg = pd.to_numeric(df_invoice['FKIMG'], errors='coerce').fillna(0)
    _netwr = pd.to_numeric(df_invoice['NETWR'], errors='coerce').fillna(0)
    df_invoice = df_invoice[(_fkimg.abs() >= 1e-6) | (_netwr.abs() >= 1e-6)].copy()

    pivot_invoice = df_invoice.pivot_table(
        index=['VKORG', 'VBELN_ORD', 'POSNR_ORD'],
        values=['发票金额_本币', '发票数量'],
        aggfunc='sum'
    ).reset_index()
    pivot_invoice = pivot_invoice.rename(columns={'VBELN_ORD': 'VBELN', 'POSNR_ORD': 'POSNR'})

    # 聚合发票类型（FKART，多发票时取第一个）
    if 'FKART' in df_invoice.columns:
        agg_fkart = df_invoice.groupby(['VKORG', 'VBELN_ORD', 'POSNR_ORD'])['FKART'].first().reset_index()
        agg_fkart = agg_fkart.rename(columns={'VBELN_ORD': 'VBELN', 'POSNR_ORD': 'POSNR', 'FKART': '发票类型'})
        pivot_invoice = pivot_invoice.merge(agg_fkart, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    # 聚合发票号（VBRP.VBELN 为发票号）
    if 'VBELN' in df_invoice.columns:
        agg_inv = df_invoice.groupby(['VKORG', 'VBELN_ORD', 'POSNR_ORD'])['VBELN'].agg(
            lambda x: ', '.join(sorted(set(str(v).strip() for v in x if pd.notna(v) and str(v).strip())))
        ).reset_index().rename(columns={'VBELN': '发票号', 'VBELN_ORD': 'VBELN', 'POSNR_ORD': 'POSNR'})
        pivot_invoice = pivot_invoice.merge(agg_inv, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    print(f'[OK] 发票汇总表: {len(pivot_invoice):,} 行')
    return pivot_invoice


def match_three_documents(pivot_order, pivot_delivery, pivot_invoice, matnr_mapper):
    """执行三单匹配（以发票为基础）"""
    print('\n正在执行三单匹配...')
    if pivot_invoice.empty:
        print('[WARNING] 发票汇总为空，无法匹配')
        return pd.DataFrame()

    # 防止 pivot_order 中 (VKORG, VBELN, POSNR) 因 MATNR 分多行导致合并时发票金额重复计算
    key_cols = ['VKORG', 'VBELN', 'POSNR']
    dup_count = pivot_order.duplicated(subset=key_cols, keep=False).sum()
    if dup_count > 0:
        n_dup_keys = int((pivot_order.groupby(key_cols).size() > 1).sum())
        print(f'  [INFO] 订单表存在 {n_dup_keys:,} 个订单行对应多物料，将按订单行聚合以避免重复计算')
        sum_cols = [c for c in ['订单金额_外币', '订单金额_本币', '订单数量'] if c in pivot_order.columns]
        first_cols = [c for c in ['MATNR', 'AUART'] if c in pivot_order.columns]
        agg_dict = {c: 'sum' for c in sum_cols}
        agg_dict.update({c: 'first' for c in first_cols})
        pivot_order = pivot_order.groupby(key_cols, as_index=False).agg(agg_dict)

    df_join = pivot_invoice.merge(pivot_delivery, on=['VKORG', 'VBELN', 'POSNR'], how='left')
    df_join = df_join.merge(pivot_order, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    # 不再按 ORDER_TYPE 过滤，与 three_lists 一致

    # 物料描述
    merge_cols = [c for c in ['MATNR', 'VKORG', 'VBELN', 'POSNR'] if c in matnr_mapper.columns and c in df_join.columns]
    if merge_cols and 'ARKTX' in matnr_mapper.columns:
        mapper_sub = matnr_mapper[merge_cols + ['ARKTX']].drop_duplicates(merge_cols)
        df_join = df_join.merge(mapper_sub, on=merge_cols, how='left')

    for col in ['发票数量', '发票金额_本币', '交货数量', '订单数量', '订单金额_本币', '订单金额_外币']:
        if col in df_join.columns:
            df_join[col] = df_join[col].fillna(0).round(2)

    # 差异计算（交货单不维护金额，仅比较订单-发票金额；交货单只参与数量比对）
    df_join['订单-金额'] = df_join['订单金额_本币']
    df_join['发票-金额'] = df_join['发票金额_本币']
    df_join['订单-发票金额差异'] = df_join['订单-金额'] - df_join['发票-金额']
    df_join['订单-交货数量差异'] = df_join['订单数量'] - df_join['交货数量'].fillna(0)
    df_join['订单-发票数量差异'] = df_join['订单数量'] - df_join['发票数量']
    df_join['交货单-发票数量差异'] = df_join['交货数量'].fillna(0) - df_join['发票数量']

    inv_col = '发票金额_本币' if '发票金额_本币' in df_join.columns else '发票-金额'
    total_inv = df_join[inv_col].sum() if inv_col in df_join.columns else 0
    # 重复检测：行数不应超过 pivot_invoice（合并若产生重复则说明存在重复计算）
    if len(df_join) > len(pivot_invoice):
        dup_cnt = len(df_join) - len(pivot_invoice)
        print(f'  [WARNING] 合并后行数({len(df_join):,})>发票行数({len(pivot_invoice):,})，存在 {dup_cnt:,} 行重复计算')
    print(f'[OK] 三单匹配完成! 匹配行数: {len(df_join):,}, 发票金额合计: {total_inv:,.2f}')
    return df_join


def export_results(df_join, pivot_order, pivot_delivery, pivot_invoice, df_order_filtered, df_delivery):
    """导出结果"""
    output_folder = get_output_folder()
    ensure_dir(output_folder)

    # 发票为空时 match 返回空 DataFrame（无列），需提前处理
    if df_join.empty or '订单数量' not in df_join.columns:
        df_neg_inv = pd.DataFrame()
        if df_order_filtered.empty and len(df_neg_inv) == 0:
            print('[INFO] 无匹配结果且无订单数据，跳过导出')
            return
        # 仅处理 Untested（负开票冲帐等）
        print('\n正在分析未测试单据...')
        output_file_notest = os.path.join(output_folder, f'{OUTPUT_PREFIX}_Untested_{get_company_code()}_{ORDER_YEAR}_{ORDER_MONTH_START}-{ORDER_MONTH_END}.xlsx')
        if not df_order_filtered.empty:
            # 有订单无发票，需走完整 Untested 逻辑（在下方）
            pass
        else:
            print('[INFO] 无订单数据且无负开票冲帐，跳过未测试单据')
            return

    company_code = get_company_code()
    ran_num = random.randint(1, 1000) if USE_RANDOM_SUFFIX else ''
    suffix = f'_{ran_num}' if ran_num else ''
    period_str = f'{ORDER_YEAR}_{ORDER_MONTH_START}-{ORDER_MONTH_END}'

    _ZWSP = '\u200B'
    df_export = df_join.copy()
    if 'MATNR' in df_export.columns:
        df_export['MATNR'] = _ZWSP + df_export['MATNR'].fillna('').astype(str)
    if 'AUART' in df_export.columns:
        df_export = df_export.rename(columns={'AUART': '订单类型'})

    front_cols = ['VKORG', 'VBELN', 'POSNR', '订单类型', '发票号', '发票类型', '聚合交货单号']
    existing_front = [c for c in front_cols if c in df_export.columns]
    other_cols = [c for c in df_export.columns if c not in existing_front]
    df_export = df_export[existing_front + other_cols]

    # 负开票冲帐：订单数量≈0 且 (交货数量≈0 或未匹配) 且 发票金额<0，移入 Untested
    df_neg_inv = pd.DataFrame()
    if not df_export.empty and '订单数量' in df_export.columns:
        inv_col = '发票金额_本币' if '发票金额_本币' in df_export.columns else '发票-金额'
        ord_near_zero = pd.to_numeric(df_export['订单数量'], errors='coerce').fillna(0).abs() < 0.01
        dlv_missing = df_export['交货数量'].isna() if '交货数量' in df_export.columns else pd.Series([True] * len(df_export), index=df_export.index)
        dlv_near_zero = pd.to_numeric(df_export['交货数量'], errors='coerce').fillna(0).abs() < 0.01 if '交货数量' in df_export.columns else pd.Series([False] * len(df_export), index=df_export.index)
        inv_negative = pd.to_numeric(df_export[inv_col], errors='coerce').fillna(0) < -0.01 if inv_col in df_export.columns else False
        mask_neg_inv = ord_near_zero & (dlv_missing | dlv_near_zero) & inv_negative
        df_neg_inv = df_export[mask_neg_inv].copy()
        df_export = df_export[~mask_neg_inv]
        if len(df_neg_inv) > 0:
            print(f'  [INFO] 负开票冲帐 {len(df_neg_inv):,} 行移入 Untested')

    EXCEL_MAX = EXCEL_MAX_ROWS_PER_FILE
    total = len(df_export)
    base_path = os.path.join(output_folder, f'{OUTPUT_PREFIX}_{company_code}_{period_str}{suffix}')
    if total <= EXCEL_MAX:
        output_file = base_path + '.xlsx'
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']
            if 'MATNR' in df_export.columns:
                matnr_col_idx = df_export.columns.get_loc('MATNR') + 1
                for row in range(2, len(df_export) + 2):
                    cell = ws.cell(row=row, column=matnr_col_idx)
                    cell.number_format = '@'
                    v = cell.value
                    if v is not None and str(v).strip() and not str(v).startswith(_ZWSP):
                        cell.value = _ZWSP + str(v).lstrip(_ZWSP)
        print(f'[OK] 主要匹配结果已导出: {output_file}')
    else:
        n_parts = (total + EXCEL_MAX - 1) // EXCEL_MAX
        for p in range(n_parts):
            start, end = p * EXCEL_MAX, min((p + 1) * EXCEL_MAX, total)
            chunk = df_export.iloc[start:end]
            out_path = f'{base_path}.xlsx' if p == 0 else f'{base_path}_{p + 1}.xlsx'
            with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                chunk.to_excel(writer, index=False, sheet_name='Sheet1')
                ws = writer.sheets['Sheet1']
                if 'MATNR' in chunk.columns:
                    matnr_col_idx = chunk.columns.get_loc('MATNR') + 1
                    for row in range(2, len(chunk) + 2):
                        cell = ws.cell(row=row, column=matnr_col_idx)
                        cell.number_format = '@'
                        v = cell.value
                        if v is not None and str(v).strip() and not str(v).startswith(_ZWSP):
                            cell.value = _ZWSP + str(v).lstrip(_ZWSP)
        print(f'[OK] 主要匹配结果已导出 {n_parts} 个文件 ({total:,} 行): {base_path}.xlsx 等')

    # 未测试单据（有订单+交货但无发票 + 负开票冲帐）
    print('\n正在分析未测试单据...')
    output_file_notest = os.path.join(output_folder, f'{OUTPUT_PREFIX}_Untested_{company_code}_{period_str}{suffix}.xlsx')

    def _write_chunked(df_part, base_name, writer):
        """单 sheet 超 Excel 行限时拆成多个 sheet"""
        n = len(df_part)
        if n == 0:
            return
        if n <= EXCEL_MAX:
            df_part.to_excel(writer, sheet_name=base_name[:31], index=False)
            return
        n_parts = (n + EXCEL_MAX - 1) // EXCEL_MAX
        for i in range(n_parts):
            start, end = i * EXCEL_MAX, min((i + 1) * EXCEL_MAX, n)
            chunk = df_part.iloc[start:end]
            sheet_name = f'{base_name}_{i+1}'[:31]  # Excel sheet 名最长 31 字符
            chunk.to_excel(writer, sheet_name=sheet_name, index=False)

    if df_order_filtered.empty:
        if len(df_neg_inv) == 0:
            print('[INFO] 无订单数据且无负开票冲帐，跳过未测试单据')
            return
        with pd.ExcelWriter(output_file_notest) as writer:
            _write_chunked(df_neg_inv, '负开票冲帐', writer)
        print(f'[OK] 未测试单据已导出（仅负开票冲帐）: {output_file_notest}')
        return

    pivot_order_2 = df_order_filtered.pivot_table(
        index=['VKORG', 'VBELN', 'POSNR', 'MATNR'],
        values=['订单金额_外币', '订单金额_本币', '订单数量'],
        aggfunc='sum'
    ).reset_index()
    # 不再对 pivot_order_2 做 MATNR/ORDER_TYPE/MATNR_MAX 限制，与 three_lists 一致
    # 补充 AUART 供 Untested 导出显示订单类型
    if 'AUART' in df_order_filtered.columns:
        auart_map = df_order_filtered[['VKORG', 'VBELN', 'POSNR', 'AUART']].drop_duplicates()
        pivot_order_2 = pivot_order_2.merge(auart_map, on=['VKORG', 'VBELN', 'POSNR'], how='left')

    pivot_delivery_2 = pd.DataFrame()
    if not df_delivery.empty and 'VBELV' in df_delivery.columns:
        pivot_delivery_2 = df_delivery.pivot_table(
            index=['VKORG', 'VBELV', 'POSNV'],
            values=['交货数量'],
            aggfunc='sum'
        ).reset_index().rename(columns={'VBELV': 'VBELN', 'POSNV': 'POSNR'})

    if pivot_delivery_2.empty:
        df_notest = pivot_order_2.copy()
        df_notest['交货数量'] = np.nan
    else:
        df_notest = pivot_order_2.merge(pivot_delivery_2, on=['VKORG', 'VBELN', 'POSNR'], how='outer')
    if not pivot_invoice.empty and 'VKORG' in pivot_invoice.columns:
        df_notest = df_notest.merge(pivot_invoice, on=['VKORG', 'VBELN', 'POSNR'], how='left')
    if '发票数量' in df_notest.columns:
        df_notest = df_notest[df_notest['发票数量'].isna()]
    else:
        df_notest['发票数量'] = np.nan

    if 'AUART' in df_notest.columns:
        df_notest = df_notest.rename(columns={'AUART': '订单类型'})
    if 'MATNR' in df_notest.columns:
        df_notest = df_notest.copy()
        df_notest['MATNR'] = _ZWSP + df_notest['MATNR'].fillna('').astype(str)

    with pd.ExcelWriter(output_file_notest) as writer:
        has_dlv = '交货数量' in df_notest.columns
        has_ord = '订单数量' in df_notest.columns
        mask_ord_only = df_notest['交货数量'].isna() if has_dlv else pd.Series([True] * len(df_notest), index=df_notest.index)
        mask_ord_dlv = (df_notest['交货数量'].notna() if has_dlv else False) & (df_notest['订单数量'].notna() if has_ord else False)
        mask_dlv_only = df_notest['订单数量'].isna() if has_ord else pd.Series([True] * len(df_notest), index=df_notest.index)
        _write_chunked(df_notest[mask_ord_only], '仅订单', writer)
        _write_chunked(df_notest[mask_ord_dlv], '仅订单及发货单', writer)
        _write_chunked(df_notest[mask_dlv_only], '仅发货单', writer)
        _write_chunked(df_neg_inv, '负开票冲帐', writer)
    print(f'[OK] 未测试单据已导出: {output_file_notest}')


def main():
    """主函数"""
    print('=' * 60)
    print(f'销售三单匹配 - {ORDER_YEAR}年{ORDER_MONTH_START}-{ORDER_MONTH_END}月')
    print('=' * 60)

    import config as _cfg
    mem_save = get_effective_memory_save_mode()
    kf = lambda t: getattr(_cfg, f'KEY_FIELDS_{t}', None)
    if mem_save and not MEMORY_SAVE_MODE:
        print(f'[INFO] 数据文件夹 > {LARGE_FOLDER_THRESHOLD_GB}GB，已自动启用大表优化（分块+日期过滤）')

    print('\n步骤1: 读取订单数据')
    df_vbak = read_sd_data(VBAK_FILE, 'VBAK', key_fields=kf('VBAK'))
    df_vbap = read_sd_data(VBAP_FILE, 'VBAP', key_fields=kf('VBAP'))

    print('\n步骤2: 处理订单数据')
    pivot_order, matnr_mapper, df_order_filtered = process_order_data(df_vbak, df_vbap)
    # 与 three_lists 一致：从 VBAK 提取订单号集合，供 LIPS vbeln_filter 使用（交货单仅保留关联订单的行）
    vbeln_set = set(df_vbak['VBELN'].astype(str).str.strip().dropna().unique()) if not df_vbak.empty and 'VBELN' in df_vbak.columns else set()
    if mem_save:
        del df_vbak, df_vbap

    print('\n步骤3: 读取交货数据')
    df_likp = read_sd_data(LIKP_FILE, 'LIKP', key_fields=kf('LIKP'))
    df_lips = read_sd_data(LIPS_FILE, 'LIPS', key_fields=kf('LIPS'))  # 与 three_lists 一致：不做 LFDAT 日期过滤
    pivot_delivery, df_delivery = process_delivery_data(df_likp, df_lips, date_prefiltered=False, vbeln_filter=vbeln_set)
    if mem_save:
        del df_likp, df_lips

    print('\n步骤4: 读取发票数据')
    df_vbrk = read_sd_data(VBRK_FILE, 'VBRK', key_fields=kf('VBRK'))
    df_vbrp = read_sd_data(VBRP_FILE, 'VBRP', key_fields=kf('VBRP'), filter_date_col='FKDAT')
    # 与 three_lists 一致：发票不按 vbeln_filter 过滤，包含全部 FKDAT 范围内发票
    pivot_invoice = process_invoice_data(df_vbrk, df_vbrp, date_prefiltered=mem_save, vbeln_filter=None)
    if mem_save:
        del df_vbrk, df_vbrp

    print('\n步骤5: 三单匹配')
    df_join = match_three_documents(pivot_order, pivot_delivery, pivot_invoice, matnr_mapper)

    print('\n步骤6: 导出结果')
    export_results(df_join, pivot_order, pivot_delivery, pivot_invoice, df_order_filtered, df_delivery)

    print('\n' + '=' * 60)
    print('全部完成!')
    print(f'输出文件夹: {get_output_folder()}/')
    if not df_join.empty:
        print(f'匹配结果行数: {len(df_join):,}')
    print(f'期间范围: {ORDER_YEAR}年{ORDER_MONTH_START}月 - {ORDER_MONTH_END}月')
    print('=' * 60)


if __name__ == "__main__":
    main()
