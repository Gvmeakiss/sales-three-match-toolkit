"""
按 KPMG_SD_SAP_ECC6.xml 导出逻辑 + 实际导出文件 拆分 InPut 原始数据

【导出逻辑】XML ConstraintTable：VBAK/LIKP 按 VKORG，VBRK 按 BUKRS(CoCode)
【实际文件】VBAK 含 VKORG+BUKRS_VF，LIKP 仅 VKORG，VBRK 含 VKORG+BUKRS

为与 XML 及"按公司处理"一致，默认按 BUKRS（公司代码）拆分，每文件夹含 6 张表。

使用方法：
    python split_by_vkorg.py --list
    python split_by_vkorg.py --code 4390      # 仅拆分指定公司
    python split_by_vkorg.py --by vkorg       # 按 VKORG 拆分（备选）
    python split_by_vkorg.py                  # 按 BUKRS 拆分（默认）
"""

import os
import re
import glob
import sys
import argparse
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd

try:
    from config import INPUT_ROOT
    from utils.DylanTools import get_encoding
    from utils.path_utils import ensure_parent_and_open, resolve_path
except ImportError:
    _dir = os.path.dirname(os.path.abspath(__file__))
    INPUT_ROOT = os.path.normpath(os.path.join(_dir, '..', 'InPut'))

    def ensure_parent_and_open(filepath, mode='w', encoding=None, **kw):
        if encoding:
            kw['encoding'] = encoding
        abs_path = os.path.abspath(os.path.normpath(filepath))
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        return abs_path, open(abs_path, mode, **kw)

    def resolve_path(p):
        return os.path.abspath(os.path.normpath(p))

SEP = '#|#'
CHUNKSIZE = 200000
VALID_COL_PATTERN = re.compile(r'^[A-Z][A-Z0-9_]*$')


def _get_files_simple(data_folder, table):
    pattern = os.path.join(data_folder, f'{table}_*.TXT')
    files = sorted(glob.glob(pattern))
    exclude = (f'{table}_VBRP', f'{table}_VBAK', f'{table}_VBRK')
    return [f for f in files if not any(ex in os.path.basename(f) for ex in exclude)]


def _clean_columns(df):
    cols = [c for c in df.columns if isinstance(c, str) and VALID_COL_PATTERN.match(c.upper())]
    return df[cols].copy() if cols else df


def _write_df_to_file(df, f, enc):
    df = _clean_columns(df)
    arr = df.fillna('').astype(str).replace('nan', '').values
    for row in arr:
        f.write(SEP.join(row) + '\n')


def build_vkorg_to_bukrs(data_folder):
    """从 VBAK 和 VBRK 构建 VKORG->BUKRS 映射（LIKP 无 BUKRS 需此映射）"""
    vkorg_bukrs = {}
    for table, vk_col, bukrs_col in [('VBAK', 'VKORG', 'BUKRS_VF'), ('VBRK', 'VKORG', 'BUKRS')]:
        files = _get_files_simple(data_folder, table)
        if not files:
            continue
        enc = get_encoding(files[0])
        # 读取全部分片文件，避免仅采样前几片导致 VKORG->BUKRS 映射不全
        for fp in files:
            for chunk in pd.read_csv(fp, sep=SEP, engine='python', encoding=enc,
                                    chunksize=CHUNKSIZE, dtype=str, on_bad_lines='skip'):
                chunk.columns = chunk.columns.str.strip().str.upper()
                if vk_col not in chunk.columns or bukrs_col not in chunk.columns:
                    continue
                for _, row in chunk.iterrows():
                    vk = str(row.get(vk_col, '')).strip()
                    bu = str(row.get(bukrs_col, '')).strip()
                    if vk and bu and vk.upper() != 'NAN' and bu.upper() != 'NAN':
                        vkorg_bukrs[vk] = bu
    return vkorg_bukrs


def split_header_table(table_name, data_folder, output_base, split_col, target_code=None,
                       vkorg_to_bukrs=None, append_existing=False):
    """
    split_col: 直接用于拆分的列
    vkorg_to_bukrs: 若提供且 split_col=='VKORG'，则按映射写入 BUKRS 文件夹（用于 LIKP）
    """
    output_base = Path(output_base).resolve()
    files = _get_files_simple(data_folder, table_name)
    if not files:
        print(f'  [SKIP] {table_name}: 无匹配文件')
        return {}
    enc = get_encoding(files[0])
    file_handles = {}
    header_written = set()
    vbeln_map = {}
    total_rows = 0
    tgt = str(target_code).strip() if target_code else None
    use_map = vkorg_to_bukrs and split_col == 'VKORG'
    for fp in files:
        for chunk in pd.read_csv(fp, sep=SEP, engine='python', encoding=enc,
                                chunksize=CHUNKSIZE, dtype=str, on_bad_lines='skip'):
            chunk.columns = chunk.columns.str.strip().str.upper()
            if split_col not in chunk.columns:
                print(f'  [WARN] {table_name} 无 {split_col} 列')
                return {}
            chunk['_split'] = chunk[split_col].fillna('').astype(str).str.strip()
            chunk.loc[chunk['_split'].isin(['', 'NAN', 'nan']), '_split'] = 'UNKNOWN'
            if use_map:
                chunk['_split'] = chunk['_split'].map(lambda x: vkorg_to_bukrs.get(x, 'UNMAPPED') if x != 'UNKNOWN' else x)
            for code, grp in chunk.groupby('_split', sort=False):
                code = str(code).strip()
                if code in ('UNKNOWN', 'UNMAPPED'):
                    continue
                if tgt and code != tgt:
                    continue
                out_path_str = resolve_path(str(output_base / code / f'{table_name}_001.TXT'))
                grp = grp.drop(columns=['_split'])
                grp = _clean_columns(grp)
                if code not in header_written:
                    header_written.add(code)
                    exists_with_content = os.path.exists(out_path_str) and os.path.getsize(out_path_str) > 0
                    if append_existing and exists_with_content:
                        _, fh = ensure_parent_and_open(out_path_str, 'a', encoding=enc)
                        _write_df_to_file(grp, fh, enc)
                    else:
                        _, fh = ensure_parent_and_open(out_path_str, 'w', encoding=enc)
                        fh.write(SEP.join(grp.columns) + '\n')
                        _write_df_to_file(grp, fh, enc)
                    file_handles[out_path_str] = fh
                else:
                    if out_path_str not in file_handles:
                        _, fh = ensure_parent_and_open(out_path_str, 'a', encoding=enc)
                        file_handles[out_path_str] = fh
                    _write_df_to_file(grp, file_handles[out_path_str], enc)
                total_rows += len(grp)
                if 'VBELN' in grp.columns:
                    for v in grp['VBELN'].astype(str).str.strip():
                        vbeln_map[v] = code
    for fh in file_handles.values():
        fh.close()
    label = 'BUKRS' if use_map or split_col in ('BUKRS', 'BUKRS_VF') else 'VKORG'
    print(f'  [OK] {table_name}: {total_rows:,} 行 -> {len(header_written)} 个 {label}, 映射 {len(vbeln_map):,} VBELN')
    return vbeln_map


def split_line_table(table_name, data_folder, output_base, vbeln_map, target_code=None, append_existing=False):
    output_base = Path(output_base).resolve()
    files = _get_files_simple(data_folder, table_name)
    if not files:
        print(f'  [SKIP] {table_name}: 无匹配文件')
        return
    if not vbeln_map:
        print(f'  [WARN] {table_name}: 无 VBELN 映射')
        return
    enc = get_encoding(files[0])
    file_handles = {}
    header_written = set()
    total_rows = 0
    unmapped = 0
    tgt = str(target_code).strip() if target_code else None
    for fp in files:
        for chunk in pd.read_csv(fp, sep=SEP, engine='python', encoding=enc,
                                chunksize=CHUNKSIZE, dtype=str, on_bad_lines='skip'):
            chunk.columns = chunk.columns.str.strip().str.upper()
            if 'VBELN' not in chunk.columns:
                print(f'  [WARN] {table_name} 无 VBELN')
                return
            chunk['_split'] = chunk['VBELN'].astype(str).str.strip().map(
                lambda x: vbeln_map.get(x, 'UNMAPPED'))
            unmapped += (chunk['_split'] == 'UNMAPPED').sum()
            if tgt:
                chunk = chunk[chunk['_split'] == tgt].copy()
                if chunk.empty:
                    continue
            for code, grp in chunk.groupby('_split', sort=False):
                if code == 'UNMAPPED':
                    continue
                code = str(code).strip()
                out_path_str = resolve_path(str(output_base / code / f'{table_name}_001.TXT'))
                grp = grp.drop(columns=['_split'])
                grp = _clean_columns(grp)
                if code not in header_written:
                    header_written.add(code)
                    exists_with_content = os.path.exists(out_path_str) and os.path.getsize(out_path_str) > 0
                    if append_existing and exists_with_content:
                        _, fh = ensure_parent_and_open(out_path_str, 'a', encoding=enc)
                        _write_df_to_file(grp, fh, enc)
                    else:
                        _, fh = ensure_parent_and_open(out_path_str, 'w', encoding=enc)
                        fh.write(SEP.join(grp.columns) + '\n')
                        _write_df_to_file(grp, fh, enc)
                    file_handles[out_path_str] = fh
                else:
                    if out_path_str not in file_handles:
                        _, fh = ensure_parent_and_open(out_path_str, 'a', encoding=enc)
                        file_handles[out_path_str] = fh
                    _write_df_to_file(grp, file_handles[out_path_str], enc)
                total_rows += len(grp)
    for fh in file_handles.values():
        fh.close()
    ustr = f', 未映射 {unmapped:,} 行' if unmapped else ''
    print(f'  [OK] {table_name}: {total_rows:,} 行 -> {len(header_written)} 个文件夹{ustr}')


def list_codes(data_folder):
    """列出 VBAK 的 BUKRS_VF/VKORG、VBRK 的 BUKRS"""
    bukrss, vkorgs = set(), set()
    for table, cols in [('VBAK', ['BUKRS_VF', 'VKORG']), ('VBRK', ['BUKRS'])]:
        files = _get_files_simple(data_folder, table)
        if not files:
            continue
        enc = get_encoding(files[0])
        for fp in files[:2]:
            for chunk in pd.read_csv(fp, sep=SEP, engine='python', encoding=enc,
                                    chunksize=50000, dtype=str, on_bad_lines='skip'):
                chunk.columns = chunk.columns.str.strip().str.upper()
                for c in cols:
                    if c not in chunk.columns:
                        continue
                    vals = chunk[c].astype(str).str.strip().replace('', pd.NA).dropna().unique()
                    (bukrss if c in ('BUKRS', 'BUKRS_VF') else vkorgs).update(vals)
    bukrss = sorted(v for v in bukrss if v and str(v).upper() != 'NAN')
    vkorgs = sorted(v for v in vkorgs if v and str(v).upper() != 'NAN')
    return bukrss, vkorgs


def main():
    parser = argparse.ArgumentParser(description='按 KPMG 导出逻辑+实际文件拆分')
    parser.add_argument('--code', type=str, help='仅处理指定 BUKRS 或 VKORG')
    parser.add_argument('--by', choices=['bukrs', 'vkorg'], default='bukrs',
                        help='拆分维度：bukrs=公司代码(默认), vkorg=销售组织')
    parser.add_argument('--source', type=str, default=None, help='原始导出文件目录（默认使用 INPUT_ROOT）')
    parser.add_argument('--output', type=str, default=None, help='拆分输出目录（默认使用 INPUT_ROOT）')
    parser.add_argument('--append', action='store_true', help='增量追加模式：若目标文件已存在则追加，不覆盖')
    parser.add_argument('--list', action='store_true', help='列出 BUKRS/VKORG 后退出')
    args = parser.parse_args()
    target = args.code.strip() if args.code else None
    by_bukrs = args.by == 'bukrs'

    # 默认从 INPUT_ROOT 读取并写回；可通过 --source / --output 解耦
    source_base = Path(args.source).resolve() if args.source else Path(INPUT_ROOT).resolve()
    output_base = Path(args.output).resolve() if args.output else Path(INPUT_ROOT).resolve()
    if args.list:
        print(f'扫描目录: {source_base}')
        bukrss, vkorgs = list_codes(str(source_base))
        print(f'BUKRS（公司代码）: {len(bukrss)} 个 - {", ".join(bukrss[:25])}{"..." if len(bukrss) > 25 else ""}')
        print(f'VKORG（销售组织）: {len(vkorgs)} 个 - {", ".join(vkorgs[:25])}{"..." if len(vkorgs) > 25 else ""}')
        return

    print('=' * 60)
    print(f'按 KPMG 导出逻辑拆分 - {datetime.now():%Y-%m-%d %H:%M:%S}')
    print(f'源数据目录: {source_base}')
    print(f'输出目录: {output_base}')
    print(f'拆分维度: {"BUKRS(公司代码)" if by_bukrs else "VKORG(销售组织)"}')
    print(f'写入模式: {"追加(不覆盖)" if args.append else "覆盖(默认)"}')
    if target:
        print(f'仅处理: {target}')
    print('=' * 60)

    # 启动自检：确认可写入 output_base，否则提前报错
    output_base_str = resolve_path(str(output_base))
    try:
        test_path = os.path.join(output_base_str, '.write_test')
        _, tf = ensure_parent_and_open(test_path, 'w', encoding='utf-8')
        tf.close()
        os.unlink(test_path)
    except Exception as e:
        print(f'[ERROR] 无法在 {output_base_str} 创建文件，请检查目录权限。错误: {e}')
        print(f'  当前工作目录: {os.getcwd()}')
        return

    source_base_str = resolve_path(str(source_base))
    vkorg_to_bukrs = build_vkorg_to_bukrs(source_base_str) if by_bukrs else None
    if by_bukrs and vkorg_to_bukrs:
        print(f'  VKORG->BUKRS 映射: {len(vkorg_to_bukrs)} 条')

    bukrss, vkorgs = list_codes(source_base_str)
    codes_to_create = bukrss if by_bukrs else vkorgs
    for c in codes_to_create:
        os.makedirs(os.path.join(output_base_str, c), exist_ok=True)
    print(f'  已预创建 {len(codes_to_create)} 个公司目录\n')

    if by_bukrs:
        vbak_files = _get_files_simple(source_base_str, 'VBAK')
        df0 = pd.read_csv(vbak_files[0], sep=SEP, nrows=1, dtype=str, engine='python', on_bad_lines='skip') if vbak_files else pd.DataFrame()
        df0.columns = df0.columns.astype(str).str.strip().str.upper()
        split_col = 'BUKRS_VF' if 'BUKRS_VF' in df0.columns else 'VKORG'
        print(f'\n[1/6] VBAK (按 {split_col})')
        vbak_map = split_header_table('VBAK', source_base_str, output_base_str, split_col, target, append_existing=args.append)
    else:
        print('\n[1/6] VBAK (按 VKORG)')
        vbak_map = split_header_table('VBAK', source_base_str, output_base_str, 'VKORG', target, append_existing=args.append)

    print('[2/6] VBAP')
    split_line_table('VBAP', source_base_str, output_base_str, vbak_map, target, append_existing=args.append)

    if by_bukrs and vkorg_to_bukrs:
        print('\n[3/6] LIKP (VKORG->BUKRS 映射)')
        likp_map = split_header_table('LIKP', source_base_str, output_base_str, 'VKORG', target, vkorg_to_bukrs, append_existing=args.append)
    else:
        print('\n[3/6] LIKP (按 VKORG)')
        likp_map = split_header_table('LIKP', source_base_str, output_base_str, 'VKORG', target, append_existing=args.append)

    print('[4/6] LIPS')
    split_line_table('LIPS', source_base_str, output_base_str, likp_map, target, append_existing=args.append)

    if by_bukrs:
        print('\n[5/6] VBRK (按 BUKRS)')
        vbrk_map = split_header_table('VBRK', source_base_str, output_base_str, 'BUKRS', target, append_existing=args.append)
    else:
        print('\n[5/6] VBRK (按 VKORG)')
        vbrk_map = split_header_table('VBRK', source_base_str, output_base_str, 'VKORG', target, append_existing=args.append)

    print('[6/6] VBRP')
    split_line_table('VBRP', source_base_str, output_base_str, vbrk_map, target, append_existing=args.append)

    print('\n' + '=' * 60)
    code_name = 'BUKRS' if by_bukrs else 'VKORG'
    print(f'拆分完成。将 config.DATA_FOLDER 指向 InPut/{code_name} 后运行 three_lists.py。')
    print('=' * 60)


if __name__ == '__main__':
    main()
