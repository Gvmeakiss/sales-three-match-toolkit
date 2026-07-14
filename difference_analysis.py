"""
销售三单匹配差异分析 - 核心代码
用于分析销售三单匹配结果中的差异情况
按场景标号输出每场景的记录数、发票金额，用户可自行汇入四大类

场景定义参见 docs/场景判断条件_模块条件.md

大表优化：分片≥3 时采用分块处理，逐文件读取/统计/写出，避免全量加载导致 OOM。
"""

import gc
import re
import sys
import pandas as pd
import numpy as np
import os
import glob
from collections import defaultdict

from config import OUTPUT_PREFIX, get_output_folder, EXCEL_MAX_ROWS_PER_FILE
from utils.path_utils import ensure_dir


def _get_match_base_path(path):
    """分片文件 ..._123_2.xlsx 归组到 ..._123.xlsx；主文件 ..._123.xlsx 保持不变"""
    b = os.path.basename(path)
    if re.search(r'_\d+_\d+\.xlsx$', b):
        return os.path.join(os.path.dirname(path), re.sub(r'_\d+\.xlsx$', '.xlsx', b))
    return path


def _part_order(path):
    """主文件 0，分片 _2/_3 返回 2/3"""
    b = os.path.basename(path)
    m = re.search(r'_\d+_(\d+)\.xlsx$', b)
    return int(m.group(1)) if m else 0


# 场景标号与识别场景描述（按输出顺序 1,2,3,4,5,6,8,9,10,11,12,13）
_SCENARIO_DESC = {
    1: '订单数量=发票≠交货单,订单金额<发票金额',
    2: '订单数量=发票≠交货单,订单金额>发票金额',
    3: '订单数量=发票≠交货单,无差异',
    4: '订单数量≠发票≠交货单,订单金额<发票金额',
    5: '订单数量≠发票≠交货单,订单金额>发票金额',
    6: '订单数量≠发票≠交货单,无差异',
    7: '缺失发票 (无发票无开票金额), N/A',
    8: '无差异,订单金额<发票金额',
    9: '无差异,订单金额>发票金额',
    10: '无差异,无差异',
    11: '订单数量=交货单≠发票,无差异',
    12: '订单数量=交货单≠发票,订单金额<发票金额',
    13: '订单数量=交货单≠发票,订单金额>发票金额',
}

def _compute_order_inv_diff(df):
    """
    计算订单金额与发票金额的差异。
    订单/发票金额已由 SHKZG 等借贷标识决定正负，不再按发票类型二次取负。
    返回 Series：每行 diff = 订单金额 - 发票金额
    """
    ord_col = '订单金额_本币' if '订单金额_本币' in df.columns else ('订单-金额' if '订单-金额' in df.columns else None)
    inv_col = '发票金额_本币' if '发票金额_本币' in df.columns else ('发票-金额' if '发票-金额' in df.columns else None)
    if ord_col is None or inv_col is None:
        return pd.Series(0.0, index=df.index)
    ord_amt = pd.to_numeric(df[ord_col], errors='coerce').fillna(0)
    inv_amt = pd.to_numeric(df[inv_col], errors='coerce').fillna(0)
    return ord_amt - inv_amt


def assign_scenario_per_row(df):
    """为每行分配场景标号（1-13），写入 场景标号 列"""
    print("按场景标号分类...")
    TOL = 0.01
    amt_col = '订单-发票金额差异'
    qty_col = '订单-发票数量差异'
    ord_col = '订单数量'
    dlv_col = '交货数量'
    if amt_col not in df.columns or qty_col not in df.columns:
        print("[WARNING] 缺少 订单-发票金额差异 或 订单-发票数量差异 列")
        return df

    amt_diff = pd.to_numeric(df[amt_col], errors='coerce').fillna(0)
    qty_diff = pd.to_numeric(df[qty_col], errors='coerce').fillna(0)
    ord_qty = pd.to_numeric(df[ord_col], errors='coerce').fillna(0) if ord_col in df.columns else pd.Series(0.0, index=df.index)
    dlv_qty = pd.to_numeric(df[dlv_col], errors='coerce') if dlv_col in df.columns else pd.Series(np.nan, index=df.index)

    ord_inv_eq = (qty_diff.abs() < TOL)
    amt_eq = (amt_diff.abs() < TOL)
    amt_lt = amt_diff < -TOL
    amt_gt = amt_diff > TOL
    dlv_valid = dlv_qty.notna()
    ord_dlv_eq = dlv_valid & ((ord_qty - dlv_qty.fillna(0)).abs() < TOL)

    # 兜底：订单=发票(数量+金额) 且 订单=0.25×发运单（单位换算遗留）时，修正发运单数量=订单数量，仅影响原场景3→场景10
    RATIO_TOL = 0.05
    cand = ord_inv_eq & amt_eq & ~ord_dlv_eq & dlv_valid
    ord_safe = np.where(ord_qty.abs() < 1e-9, 1, ord_qty)
    ratio_ok = np.abs((dlv_qty.fillna(0) / ord_safe) - 4) < RATIO_TOL
    unit_fix_mask = cand & ratio_ok
    if unit_fix_mask.any():
        n_fix = int(unit_fix_mask.sum())
        # 交货数量若为 int64，需将订单数量（可能为 float）转为兼容类型再赋值，避免 LossySetitemError
        vals = pd.to_numeric(df.loc[unit_fix_mask, ord_col], errors='coerce').fillna(0)
        if pd.api.types.is_integer_dtype(df[dlv_col]):
            vals = np.round(vals).astype(df[dlv_col].dtype)
        else:
            vals = vals.astype(df[dlv_col].dtype)
        df.loc[unit_fix_mask, dlv_col] = vals.values
        dlv_qty = pd.to_numeric(df[dlv_col], errors='coerce') if dlv_col in df.columns else dlv_qty
        ord_dlv_eq = ord_dlv_eq | unit_fix_mask
        print(f"  [INFO] 订单=发票且发运单≈4×订单（单位换算）时修正发运单数量，{n_fix:,} 行归入完全匹配")

    # 按优先级赋值（互斥）
    scenario = pd.Series(0, index=df.index, dtype=int)
    scenario[ord_dlv_eq & ord_inv_eq & amt_eq] = 10
    scenario[ord_dlv_eq & ord_inv_eq & amt_lt] = 8
    scenario[ord_dlv_eq & ord_inv_eq & amt_gt] = 9
    scenario[ord_inv_eq & ~ord_dlv_eq & amt_eq] = 3
    scenario[ord_dlv_eq & ~ord_inv_eq & amt_eq] = 11
    scenario[~ord_dlv_eq & ~ord_inv_eq & amt_eq] = 6
    scenario[ord_inv_eq & ~ord_dlv_eq & amt_lt] = 1
    scenario[ord_inv_eq & ~ord_dlv_eq & amt_gt] = 2
    scenario[ord_dlv_eq & ~ord_inv_eq & amt_lt] = 12
    scenario[ord_dlv_eq & ~ord_inv_eq & amt_gt] = 13
    scenario[~ord_dlv_eq & ~ord_inv_eq & amt_lt] = 4
    scenario[~ord_dlv_eq & ~ord_inv_eq & amt_gt] = 5

    df['场景标号'] = scenario
    return df


def _pct_str(count, total):
    """安全计算占比，避免 len(df)==0 除零"""
    if total == 0:
        return 'N/A'
    return f"{(count / total * 100):.1f}%"


def _amt_pct_str(amt, total_amt, total_amt_positive=None):
    """金额占比。存在冲销(负数)时，分母用正数发票合计，避免正数类别占比>100%"""
    denom = total_amt_positive if (total_amt_positive is not None and total_amt_positive > 0) else total_amt
    if denom == 0 or denom != denom:
        return 'N/A'
    return f"{(amt / denom * 100):.1f}%"


def generate_scenario_report(df, total_rows=None, total_amt=None, amt_denom=None):
    """按场景标号生成统计：场景标号、识别场景、记录数、占比、发票金额、发票金额占比、订单发票金额差异"""
    if total_rows is None:
        total_rows = len(df)
    inv_col = '发票金额_本币' if '发票金额_本币' in df.columns else '发票-金额'
    if total_amt is None and inv_col in df.columns:
        total_amt = df[inv_col].fillna(0).sum()
    total_amt = total_amt or 0
    if amt_denom is None:
        amt_denom = total_amt
    # 金额占比分母用正数合计
    if inv_col in df.columns:
        pos_amt = df[inv_col].fillna(0)
        amt_denom = pos_amt[pos_amt > 0].sum() if (pos_amt > 0).any() else total_amt

    diff_series = _compute_order_inv_diff(df)
    # 场景10（数量金额均相等）不输出到统计与详细，此处排除
    scenario_order = [1, 2, 3, 4, 5, 6, 8, 9, 11, 12, 13]
    rows = []
    for sid in scenario_order:
        mask = df['场景标号'] == sid if '场景标号' in df.columns else pd.Series(False, index=df.index)
        cnt = int(mask.sum())
        amt = float(df.loc[mask, inv_col].sum()) if inv_col in df.columns and mask.any() else 0.0
        order_inv_diff = float(diff_series[mask].sum()) if mask.any() else 0.0
        rows.append({
            '场景标号': sid,
            '识别场景': _SCENARIO_DESC.get(sid, ''),
            '记录数': cnt,
            '占比': _pct_str(cnt, total_rows),
            '发票金额': round(amt, 2),
            '发票金额占比': _amt_pct_str(amt, total_amt, amt_denom),
            '订单发票金额差异': round(order_inv_diff, 2),
        })
    return rows


def _get_untested_path(output_folder, prefix, newest_base):
    """获取同批次 Untested 文件路径，找不到则返回 None。
    支持 4390 分片模式：当存在 *_分片_1.xlsx 时优先返回主分片（含 仅订单/仅订单及发货单）。"""
    base_name = os.path.basename(newest_base)
    after_prefix = base_name.replace(f'{prefix}_', '', 1).replace('.xlsx', '')
    split_path = os.path.join(output_folder, f'{prefix}_Untested_{after_prefix}_分片_1.xlsx')
    if os.path.exists(split_path):
        return split_path
    untested_path = os.path.join(output_folder, f'{prefix}_Untested_{after_prefix}.xlsx')
    if os.path.exists(untested_path):
        return untested_path
    parts = after_prefix.split('_')
    company_code = parts[0] if parts else ''
    pattern = os.path.join(output_folder, f'{prefix}_Untested_{company_code}_*.xlsx')
    matches = [f for f in glob.glob(pattern) if os.path.exists(f) and '_分片_' not in os.path.basename(f)]
    if matches:
        return max(matches, key=os.path.getmtime)
    matches_split = [f for f in glob.glob(os.path.join(output_folder, f'{prefix}_Untested_{company_code}_*_分片_1.xlsx')) if os.path.exists(f)]
    return max(matches_split, key=os.path.getmtime) if matches_split else None


def _load_neg_inv_stats(untested_path):
    """从 Untested 的「负开票冲帐」sheet(s) 读取 (记录数, 发票金额)。
    支持分片：单文件内 负开票冲帐_1、_2 等；4390 拆分多文件时从 _分片_1 ~ _分片_N 依次读取并累加。
    仅读取发票金额列以节省内存。"""
    if not untested_path or not os.path.exists(untested_path):
        return 0, 0.0
    total_cnt = 0
    total_amt = 0.0

    def _add_from_file(fp):
        nonlocal total_cnt, total_amt
        try:
            xl = pd.ExcelFile(fp)
            neg_sheets = [s for s in xl.sheet_names if s == '负开票冲帐'
                          or (s.startswith('负开票冲帐_') and s[len('负开票冲帐_'):].replace('_', '').isdigit())]
            key = lambda x: (0 if x == '负开票冲帐' else int(x.split('_')[1]) if '_' in x and x.split('_')[1].isdigit() else 999)
            for sheet in sorted(neg_sheets, key=key):
                header = pd.read_excel(fp, sheet_name=sheet, nrows=0)
                cols = set(header.columns)
                inv_c = '发票金额_本币' if '发票金额_本币' in cols else ('发票-金额' if '发票-金额' in cols else None)
                usecols = [inv_c] if inv_c else None
                df_neg = pd.read_excel(fp, sheet_name=sheet, usecols=usecols)
                if df_neg.empty:
                    continue
                total_cnt += len(df_neg)
                if inv_c and inv_c in df_neg.columns:
                    total_amt += float(df_neg[inv_c].fillna(0).sum())
            xl.close()
        except (ValueError, KeyError, OSError):
            pass

    bn = os.path.basename(untested_path)
    if '_分片_' in bn:
        base = untested_path.rsplit('_分片_', 1)[0]
        pat = base + '_分片_*.xlsx'
        def _part_num(p):
            try:
                return int(os.path.basename(p).rsplit('_分片_', 1)[1].replace('.xlsx', ''))
            except (ValueError, IndexError):
                return 0
        files = sorted(glob.glob(pat), key=_part_num)
        for fp in files:
            _add_from_file(fp)
    else:
        _add_from_file(untested_path)
    return total_cnt, total_amt


def _load_untested_scenario7(untested_path):
    """从 Untested 的「仅订单」「仅订单及发货单」sheet 读取 场景7（缺失发票）记录数。
    仅读取首列以节省内存。"""
    if not untested_path:
        return 0
    cnt = 0
    for sheet in ['仅订单', '仅订单及发货单']:
        try:
            df = pd.read_excel(untested_path, sheet_name=sheet, usecols=[0])
            cnt += len(df) if not df.empty else 0
        except (ValueError, KeyError):
            pass
    return cnt


def _write_detail_chunk(chunk, path, apply_matnr_format=True):
    """写出一个详细数据块；大块时跳过逐格 MATNR 格式化以提速"""
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        chunk.to_excel(writer, sheet_name='Sheet1', index=False)
        if apply_matnr_format and 'MATNR' in chunk.columns and len(chunk) <= 50000:
            ws = writer.sheets['Sheet1']
            idx = chunk.columns.get_loc('MATNR') + 1
            for row in range(2, len(chunk) + 2):
                ws.cell(row=row, column=idx).number_format = '@'


def main():
    """主函数"""
    print("开始分析销售三单匹配结果...", flush=True)
    prefix = OUTPUT_PREFIX or 'SalesThreeMatchResult'
    output_folder = get_output_folder()
    match_files = glob.glob(os.path.join(output_folder, f'{prefix}_*.xlsx'))
    match_files = [f for f in match_files if 'Untested' not in f and '差异分析' not in f]

    if not match_files:
        print("错误: 未找到匹配结果文件，请先运行 sales_three_match.py")
        return

    # 按逻辑组聚合（主文件 + 分片 _2/_3）
    groups = {}
    for f in match_files:
        base = _get_match_base_path(f)
        if base not in groups:
            groups[base] = []
        groups[base].append(f)
    for base in groups:
        groups[base].sort(key=lambda p: (_part_order(p), p))

    # 选最新的那组（按主文件 mtime）
    primary_candidates = [b for b in groups if os.path.exists(b)]
    if not primary_candidates:
        print("错误: 未找到有效匹配结果文件")
        return
    newest_base = max(primary_candidates, key=os.path.getmtime)
    to_read = groups[newest_base]

    # 从主文件名提取公司代码
    basename = os.path.basename(newest_base)
    parts_name = basename.replace('.xlsx', '').split('_')
    company_code = parts_name[1] if len(parts_name) >= 2 else 'ALL'

    print(f'使用最新的匹配结果: {newest_base}' + (f' 等共 {len(to_read)} 个分片' if len(to_read) > 1 else ''), flush=True)

    use_chunked = len(to_read) >= 3
    if use_chunked:
        print('[INFO] 大表模式：分块处理，逐文件读取/统计/写出，降低内存占用', flush=True)

    ensure_dir(output_folder)
    EXCEL_MAX = EXCEL_MAX_ROWS_PER_FILE
    base_name = f'{prefix}_差异分析_{company_code}'
    stats_file = os.path.join(output_folder, f'{base_name}.xlsx')
    _ZWSP = '\u200B'
    DROP_COLS = ['交货金额', '交货单-金额', '订单-交货金额差异', '交货单-发票金额差异']
    inv_col = '发票金额_本币'  # 或 发票-金额，由列存在性决定

    try:
        # 删除该公司旧的差异分析文件
        for old_f in glob.glob(os.path.join(output_folder, f'{base_name}*.xlsx')):
            try:
                os.remove(old_f)
                print(f'  [删除旧文件] {os.path.basename(old_f)}')
            except OSError:
                pass

        untested_path = _get_untested_path(output_folder, prefix, newest_base)
        neg_inv_cnt, neg_inv_amt = _load_neg_inv_stats(untested_path)
        scenario7_cnt = _load_untested_scenario7(untested_path)

        if use_chunked:
            # 分块模式：逐文件处理，避免 OOM
            cnt_by_scenario = defaultdict(float)
            amt_by_scenario = defaultdict(float)
            diff_by_scenario = defaultdict(float)
            total_rows = 0
            total_amt = 0.0
            detail_files = []
            current_detail = None
            current_detail_rows = 0
            detail_idx = 0

            for i, fp in enumerate(to_read, 1):
                print(f'  读取分片 {i}/{len(to_read)}: {os.path.basename(fp)}...', flush=True)
                df = pd.read_excel(fp, sheet_name='Sheet1')
                print(f'    -> {len(df):,} 行', flush=True)

                if 'MATNR' in df.columns:
                    s = df['MATNR'].fillna('').astype(str).str.replace(r'\.0+$', '', regex=True)
                    df['MATNR'] = s.str.lstrip(_ZWSP)

                df = assign_scenario_per_row(df)

                inv_c = '发票金额_本币' if '发票金额_本币' in df.columns else '发票-金额'
                total_rows += len(df)
                if inv_c in df.columns:
                    total_amt += df[inv_c].fillna(0).sum()

                diff_series = _compute_order_inv_diff(df)
                for sid in range(1, 14):
                    if sid == 7:
                        continue
                    mask = df['场景标号'] == sid
                    cnt_by_scenario[sid] += mask.sum()
                    if inv_c in df.columns and mask.any():
                        amt_by_scenario[sid] += df.loc[mask, inv_c].sum()
                    if mask.any():
                        diff_by_scenario[sid] += float(diff_series[mask].sum())

                df_export = df.copy()
                for dc in DROP_COLS:
                    if dc in df_export.columns:
                        df_export = df_export.drop(columns=[dc])
                if 'MATNR' in df_export.columns:
                    df_export['MATNR'] = _ZWSP + df_export['MATNR'].fillna('').astype(str)
                # 不输出场景0、场景10（数量金额均相等）到详细
                df_export = df_export[~df_export['场景标号'].isin([0, 10])].reset_index(drop=True)

                pos = 0
                while pos < len(df_export):
                    need_new = current_detail is None or current_detail_rows >= EXCEL_MAX
                    if need_new and current_detail is not None and current_detail_rows > 0:
                        path = os.path.join(output_folder, f'{base_name}_详细_{detail_idx + 1}.xlsx')
                        _write_detail_chunk(current_detail, path, apply_matnr_format=(len(current_detail) <= 50000))
                        detail_files.append(path)
                        print(f'    [写出] {os.path.basename(path)} ({len(current_detail):,} 行)', flush=True)
                        current_detail = None
                        current_detail_rows = 0
                        detail_idx += 1

                    if current_detail is None:
                        current_detail = pd.DataFrame()
                        current_detail_rows = 0

                    space = EXCEL_MAX - current_detail_rows
                    take = min(space, len(df_export) - pos)
                    chunk = df_export.iloc[pos:pos + take]
                    pos += take
                    current_detail = pd.concat([current_detail, chunk], ignore_index=True) if len(current_detail) > 0 else chunk
                    current_detail_rows = len(current_detail)

                del df, df_export
                gc.collect()

            if current_detail is not None and len(current_detail) > 0:
                if detail_idx == 0 and len(detail_files) == 0:
                    path = os.path.join(output_folder, f'{base_name}_详细.xlsx')
                else:
                    path = os.path.join(output_folder, f'{base_name}_详细_{detail_idx + 1}.xlsx')
                _write_detail_chunk(current_detail, path, apply_matnr_format=(len(current_detail) <= 50000))
                detail_files.append(path)
                print(f'    [写出] {os.path.basename(path)} ({len(current_detail):,} 行)', flush=True)

            if neg_inv_cnt > 0:
                print(f'  [INFO] 负开票冲帐: {neg_inv_cnt:,} 条, 金额 {neg_inv_amt:,.2f}', flush=True)
            if scenario7_cnt > 0:
                print(f'  [INFO] 场景7（缺失发票）: {scenario7_cnt:,} 条', flush=True)

            # 构建场景明细（含分块累加结果）；场景10 不输出
            inv_c = '发票金额_本币'
            pos_amt = sum(amt_by_scenario[s] for s in cnt_by_scenario if amt_by_scenario.get(s, 0) > 0)
            amt_denom = pos_amt if pos_amt > 0 else total_amt
            scenario_order = [1, 2, 3, 4, 5, 6, 8, 9, 11, 12, 13]
            scenario_stats = []
            for sid in scenario_order:
                cnt = int(cnt_by_scenario.get(sid, 0))
                amt = float(amt_by_scenario.get(sid, 0))
                order_inv_diff = round(float(diff_by_scenario.get(sid, 0)), 2)
                scenario_stats.append({
                    '场景标号': sid,
                    '识别场景': _SCENARIO_DESC.get(sid, ''),
                    '记录数': cnt,
                    '占比': _pct_str(cnt, total_rows),
                    '发票金额': round(amt, 2),
                    '发票金额占比': _amt_pct_str(amt, total_amt, amt_denom),
                    '订单发票金额差异': order_inv_diff,
                })
            # 插入场景7（在场景6与8之间）
            scenario_stats.insert(6, {
                '场景标号': 7,
                '识别场景': _SCENARIO_DESC.get(7, ''),
                '记录数': scenario7_cnt,
                '占比': _pct_str(scenario7_cnt, total_rows + scenario7_cnt) if (total_rows + scenario7_cnt) > 0 else 'N/A',
                '发票金额': 0.0,
                '发票金额占比': 'N/A',
                '订单发票金额差异': 0.0,
            })
            if neg_inv_cnt > 0 or abs(neg_inv_amt) >= 1e-6:
                scenario_stats.append({
                    '场景标号': '负开票',
                    '识别场景': '有发票、订单或发运单缺失,Not Test',
                    '记录数': neg_inv_cnt,
                    '占比': _pct_str(neg_inv_cnt, total_rows + neg_inv_cnt),
                    '发票金额': round(float(neg_inv_amt), 2),
                    '发票金额占比': _amt_pct_str(neg_inv_amt, total_amt, amt_denom),
                    '订单发票金额差异': 0.0,
                })
            if len(detail_files) == 1 and '_详细_1.xlsx' in detail_files[0]:
                new_path = os.path.join(output_folder, f'{base_name}_详细.xlsx')
                try:
                    os.rename(detail_files[0], new_path)
                    detail_files[0] = new_path
                except OSError:
                    pass

        else:
            # 常规模式：全量加载
            dfs = []
            for i, fp in enumerate(to_read, 1):
                print(f'  读取分片 {i}/{len(to_read)}: {os.path.basename(fp)}...', flush=True)
                dfs.append(pd.read_excel(fp, sheet_name='Sheet1'))
                print(f'    -> {len(dfs[-1]):,} 行', flush=True)
            print('合并数据...', flush=True)
            df = pd.concat(dfs, ignore_index=True)
            del dfs
            print(f"成功读取数据，共 {len(df):,} 条记录", flush=True)

            if df.empty or '订单-发票金额差异' not in df.columns:
                print("[INFO] 匹配结果为空或缺少必要列，输出空统计报告")
                scenario_stats = [{'场景标号': sid, '识别场景': _SCENARIO_DESC.get(sid, ''),
                                  '记录数': 0, '占比': 'N/A', '发票金额': 0.0, '发票金额占比': 'N/A', '订单发票金额差异': 0.0}
                                 for sid in [1, 2, 3, 4, 5, 6, 8, 9, 11, 12, 13]]
                with pd.ExcelWriter(stats_file, engine='openpyxl') as writer:
                    pd.DataFrame(scenario_stats).to_excel(writer, sheet_name='场景明细', index=False)
                print(f"[OK] 统计页已保存: {stats_file} (空结果)")
                return

            if 'MATNR' in df.columns:
                s = df['MATNR'].fillna('').astype(str).str.replace(r'\.0+$', '', regex=True)
                df['MATNR'] = s.str.lstrip(_ZWSP)

            df = assign_scenario_per_row(df)

            scenario_stats = generate_scenario_report(df)
            # 插入场景7（在场景6与8之间，来自 Untested 仅订单/仅订单及发货单）
            scenario_stats.insert(6, {
                '场景标号': 7,
                '识别场景': _SCENARIO_DESC.get(7, ''),
                '记录数': scenario7_cnt,
                '占比': _pct_str(scenario7_cnt, len(df) + scenario7_cnt) if (len(df) + scenario7_cnt) > 0 else 'N/A',
                '发票金额': 0.0,
                '发票金额占比': 'N/A',
                '订单发票金额差异': 0.0,
            })
            if neg_inv_cnt > 0 or abs(neg_inv_amt) >= 1e-6:
                inv_col = '发票金额_本币' if '发票金额_本币' in df.columns else '发票-金额'
                total_amt = df[inv_col].fillna(0).sum() if inv_col in df.columns else 0
                total_amt_positive = sum(s['发票金额'] for s in scenario_stats if s.get('发票金额', 0) > 0)
                amt_denom = total_amt_positive if total_amt_positive > 0 else total_amt
                scenario_stats.append({
                    '场景标号': '负开票',
                    '识别场景': '有发票、订单或发运单缺失,Not Test',
                    '记录数': neg_inv_cnt,
                    '占比': _pct_str(neg_inv_cnt, len(df) + neg_inv_cnt),
                    '发票金额': round(float(neg_inv_amt), 2),
                    '发票金额占比': _amt_pct_str(neg_inv_amt, total_amt, amt_denom),
                    '订单发票金额差异': 0.0,
                })
                if neg_inv_cnt > 0:
                    print(f'  [INFO] 负开票冲帐: {neg_inv_cnt:,} 条, 金额 {neg_inv_amt:,.2f}', flush=True)
            if scenario7_cnt > 0:
                print(f'  [INFO] 场景7（缺失发票）: {scenario7_cnt:,} 条', flush=True)

            df_export = df.copy()
            for dc in DROP_COLS:
                if dc in df_export.columns:
                    df_export = df_export.drop(columns=[dc])
            if 'MATNR' in df_export.columns:
                df_export['MATNR'] = _ZWSP + df_export['MATNR'].fillna('').astype(str)
            # 不输出场景0、场景10（数量金额均相等）到详细
            df_export = df_export[~df_export['场景标号'].isin([0, 10])].reset_index(drop=True)

            total = len(df_export)
            n_parts = max(1, (total + EXCEL_MAX - 1) // EXCEL_MAX)
            detail_files = []
            for p in range(n_parts):
                start, end = p * EXCEL_MAX, min((p + 1) * EXCEL_MAX, total)
                chunk = df_export.iloc[start:end]
                if n_parts == 1:
                    detail_path = os.path.join(output_folder, f'{base_name}_详细.xlsx')
                else:
                    detail_path = os.path.join(output_folder, f'{base_name}_详细_{p + 1}.xlsx')
                detail_files.append(detail_path)
                _write_detail_chunk(chunk, detail_path, apply_matnr_format=(len(chunk) <= 50000))

        print("\n=== 场景明细 ===")
        for stat in scenario_stats:
            print(f"  场景{stat['场景标号']}: {stat['记录数']} 条 ({stat['占比']}), 发票金额 {stat['发票金额']:,.2f} ({stat['发票金额占比']})")

        with pd.ExcelWriter(stats_file, engine='openpyxl') as writer:
            pd.DataFrame(scenario_stats).to_excel(writer, sheet_name='场景明细', index=False)
        print(f"[OK] 统计页已保存: {stats_file}")

        if len(detail_files) > 1:
            print(f"[OK] 详细数据已拆分为 {len(detail_files)} 个文件")
        else:
            print(f"[OK] 详细数据已保存: {detail_files[0]}")

        print(f"\n结果汇总: 统计页 {stats_file}, 详细数据 {len(detail_files)} 个文件")

    except Exception as e:
        print(f"处理过程中出错: {e}", flush=True)
        import traceback
        traceback.print_exc()
        sys.stdout.flush()
        sys.stderr.flush()


if __name__ == "__main__":
    main()
