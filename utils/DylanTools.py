"""
DylanTools - 数据工具库
用于处理KPMG数据提取工具导出的TXT文件
适配销售SD模块字段（VBELN, VBELV, POSNR, POSNV等）
"""

import pandas as pd
import numpy as np
import os


def get_encoding(file):
    """自动检测文件编码 - 增强版"""
    try:
        import chardet
        with open(file, 'rb') as f:
            sample = f.read(100000)
            result = chardet.detect(sample)
            encoding = result['encoding']
            confidence = result['confidence']
            if confidence > 0.7 and encoding:
                return encoding
    except Exception:
        pass

    common_encodings = ['gbk', 'gb2312', 'utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
    for encoding in common_encodings:
        try:
            with open(file, 'r', encoding=encoding) as f:
                for i in range(5):
                    line = f.readline()
                    if not line:
                        break
                return encoding
        except (UnicodeDecodeError, UnicodeError):
            continue

    return 'gbk'


def kpmg_txt_to_df(filename, to_pickle=True, output_folder='pickle', chunksize=200000, key_fields=None,
                   filter_date_col=None, filter_year=None, filter_month_start=None, filter_month_end=None,
                   force_chunked=False, company_code=None):
    """
    读取 KPMG 数据提取工具导出的 TXT 文件
    分隔符: #|#
    大文件（>50MB 或 force_chunked）时分块读取，避免内存溢出
    key_fields: 可选，仅保留指定字段保存 pkl，减少存储与内存
    filter_date_col, filter_year, filter_month_start, filter_month_end: 分块读取时每块先按日期过滤再合并
    company_code: 可选，公司代码，用于 pkl 文件名（如 VBAK_001_4390.pkl），避免多公司 pickle 混用
    """
    str_preserve_cols = [
        'MATNR', 'LIFNR', 'KUNNR', 'KUNRG', 'KUNAG', 'BANFN', 'BNFPO', 'CHARG', 'BELNR', 'MBLNR',
        'VBELN', 'VBELV', 'POSNR', 'POSNV', 'AUBEL', 'AUPOS', 'VGBEL', 'VGPOS', 'ANLN1', 'ANLN2'
    ]
    dtype_arg = {c: str for c in str_preserve_cols}
    enc = get_encoding(filename)
    key_fields_upper = [str(c).strip().upper() for c in (key_fields or []) if str(c).strip()]
    key_fields_set = set(key_fields_upper)

    file_size = os.path.getsize(filename) if os.path.exists(filename) else 0
    use_chunks = force_chunked or file_size > 50 * 1024 * 1024  # 50MB
    do_date_filter = use_chunks and filter_date_col and filter_year is not None
    filter_col_upper = str(filter_date_col).strip().upper() if filter_date_col else None

    if use_chunks:
        chunk_list = []
        for i, chunk in enumerate(pd.read_csv(filename, sep='#\\|#', engine='python',
                encoding=enc, dtype=dtype_arg, chunksize=chunksize, on_bad_lines='skip')):
            chunk.columns = chunk.columns.str.strip().str.upper()
            keep_cols = None
            if key_fields_set:
                keep_cols = [c for c in chunk.columns if c in key_fields_set]
                # 分块按日期过滤需要先保留日期列，过滤后再按 key_fields 输出
                if do_date_filter and filter_col_upper and filter_col_upper in chunk.columns and filter_col_upper not in keep_cols:
                    keep_cols.append(filter_col_upper)
                if keep_cols:
                    chunk = chunk[keep_cols]
            for col in chunk.columns:
                if chunk[col].dtype == 'object':
                    chunk[col] = chunk[col].str.strip()
            for col in str_preserve_cols:
                if col in chunk.columns:
                    chunk[col] = chunk[col].astype(str)
            # 分块内先按日期过滤，避免全量合并后再过滤导致 OOM
            if do_date_filter and filter_col_upper in chunk.columns:
                chunk[filter_col_upper] = pd.to_datetime(chunk[filter_col_upper].astype(str), format='%Y%m%d', errors='coerce')
                chunk = chunk[
                    (chunk[filter_col_upper].dt.year == filter_year) &
                    (chunk[filter_col_upper].dt.month >= (filter_month_start or 1)) &
                    (chunk[filter_col_upper].dt.month <= (filter_month_end or 12))
                ]
            if key_fields_set:
                out_keep = [c for c in key_fields_upper if c in chunk.columns]
                if out_keep:
                    chunk = chunk[out_keep]
            if not chunk.empty:
                chunk_list.append(chunk)
            if (i + 1) % 5 == 0:
                print(f'        已读取 {(i+1)*chunksize:,} 行...')
            if len(chunk_list) >= 3:
                df_part = pd.concat(chunk_list, ignore_index=True)
                chunk_list = [df_part]
        df = pd.concat(chunk_list, ignore_index=True) if chunk_list else pd.DataFrame()
        del chunk_list
    else:
        df = pd.read_csv(filename, sep='#\\|#', engine='python',
                         encoding=enc, dtype=dtype_arg, on_bad_lines='skip')

    df.columns = df.columns.str.strip().str.upper()
    if key_fields_set:
        keep = [c for c in key_fields_upper if c in df.columns]
        if keep:
            df = df[keep]

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()

    for col in str_preserve_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # 整数字段
    int_fields = ['GJAHR', 'BLDAT', 'BUDAT', 'MONAT', 'CPUDT', 'BUZEI', 'MANDT',
                  'EBELN', 'EBELP', 'MJAHR']
    for col in int_fields:
        if col in df.columns:
            try:
                # 使用更小整数类型以降低内存峰值
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(np.int32)
            except Exception:
                pass

    # 浮点字段（销售SD补充：FKIMG, LFIMG, AKKUR, KWMENG, MWSBP等）
    float_fields = [
        'WRBTR', 'DMBTR', 'WSL', 'HSL', 'HSLVT', 'NETWR', 'WKURS', 'KURSF',
        'MENGE', 'BPMNG', 'BNBTR', 'STOCK_POSTING',
        'FKIMG', 'LFIMG', 'AKKUR', 'KWMENG', 'MWSBP', 'NTGEW', 'BRGEW', 'VOLUM'
    ]
    for col in float_fields:
        if col in df.columns:
            try:
                # 业务上金额/数量精度对 float32 已足够，显著降低大表内存
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(np.float32)
            except Exception:
                pass

    if to_pickle:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        base_filename = os.path.basename(filename)
        pkl_base = base_filename.replace('.TXT', '.pkl').replace('.txt', '.pkl')
        if company_code:
            pkl_base = pkl_base.replace('.pkl', f'_{company_code}.pkl')
        if do_date_filter and filter_year:
            period = f'{filter_year}_{filter_month_start or 1}-{filter_month_end or 12}'
            output_path = os.path.join(output_folder, pkl_base.replace('.pkl', f'_{period}.pkl'))
        else:
            output_path = os.path.join(output_folder, pkl_base)
        df_out = df
        if key_fields_upper:
            keep = [c for c in key_fields_upper if c in df.columns]
            if keep:
                df_out = df[keep].copy()
        df_out.to_pickle(output_path)
        print(f'已保存: {output_path}' + (f' (仅 {len(df_out.columns)} 个关键字段)' if key_fields else ''))

    return df
