"""
按公司运行：三单清单 -> 三单匹配 -> 差异分析

针对 InPut 下按公司代码命名的文件夹，依次执行：
1. 生成原始三单清单
2. 三单匹配
3. 差异分析

使用方法：
    python run_all_by_company.py
    python run_all_by_company.py --workers 4   # 4 进程并行

前置条件：
    已运行 split_by_vkorg.py 完成数据拆分，InPut 下存在 4010、4030 等公司子目录。
"""

import os
import sys
import subprocess
import argparse
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

try:
    from config import PICKLE_FOLDER, LISTS_PICKLE_FOLDER, INPUT_ROOT, OUTPUT_ROOT
except ImportError:
    INPUT_ROOT = os.path.normpath(os.path.join(script_dir, '..', 'InPut'))
    OUTPUT_ROOT = os.path.normpath(os.path.join(script_dir, '..', 'OutPut'))
    PICKLE_FOLDER = os.path.join(script_dir, 'pickle')
    LISTS_PICKLE_FOLDER = os.path.join(script_dir, 'pickle', 'lists')

BASE_INPUT = os.path.normpath(os.path.abspath(INPUT_ROOT))
BASE_OUTPUT = os.path.normpath(os.path.abspath(OUTPUT_ROOT))

REQUIRED_TABLES = ['VBAK', 'VBAP', 'LIKP', 'LIPS', 'VBRK', 'VBRP']
MIN_TABLES = ['VBAK', 'VBAP']  # 至少需此二者才能运行


def _check_company_tables(company_folder):
    """检查公司目录下的表文件，返回 (完整, 缺失列表)"""
    missing = []
    for t in REQUIRED_TABLES:
        if not os.path.exists(os.path.join(company_folder, f'{t}_001.TXT')):
            missing.append(t)
    return len(missing) == 0, missing


def _has_three_lists(code):
    """检查该公司是否已有三单清单（订单、交货、发票三种均存在则跳过生成）"""
    lists_dir = os.path.join(BASE_OUTPUT, str(code), '销售三单清单')
    if not os.path.isdir(lists_dir):
        return False
    files = os.listdir(lists_dir)
    has_order = any(f.startswith('销售订单清单_') and f.endswith('.xlsx') for f in files)
    has_delivery = any(f.startswith('交货单清单_') and f.endswith('.xlsx') for f in files)
    has_invoice = any(f.startswith('销售发票清单_') and f.endswith('.xlsx') for f in files)
    return has_order and has_delivery and has_invoice


def _clear_pickle_cache(tables=None):
    """清理 pickle 缓存。tables=None 时清空全部；tables 为列表时仅删除指定表的 pkl（如 ['LIPS','VBRP']）"""
    for folder in [PICKLE_FOLDER, LISTS_PICKLE_FOLDER]:
        if not os.path.exists(folder):
            continue
        for f in os.listdir(folder):
            if not f.endswith('.pkl'):
                continue
            if tables is not None:
                if not any(t in f for t in tables):
                    continue
            try:
                os.remove(os.path.join(folder, f))
            except OSError:
                pass


def _process_one_company(code, company_folder, scripts, clear_pickle):
    """单公司处理：三单清单→三单匹配→差异分析。通过 env 指定数据目录，支持多进程并行。"""
    env = os.environ.copy()
    env['SALES_DATA_FOLDER'] = os.path.normpath(os.path.abspath(company_folder))
    env['SALES_COMPANY_CODE'] = str(code)
    if clear_pickle:
        _clear_pickle_cache()
    first_fail = None
    for script, desc in scripts:
        result = subprocess.run(
            [sys.executable, os.path.join(script_dir, script)],
            cwd=script_dir,
            capture_output=False,
            env=env
        )
        if result.returncode != 0:
            if first_fail is None:
                first_fail = desc
            break
    return code, first_fail


def _get_companies_for_diff_only():
    """仅差异分析模式：扫描 OutPut 下存在匹配结果文件的公司"""
    import glob
    prefix = 'SalesThreeMatchResult'
    seen = set()
    folders = []
    if not os.path.exists(BASE_OUTPUT):
        return folders
    for name in sorted(os.listdir(BASE_OUTPUT)):
        if name in seen or not name.strip().isdigit():
            continue
        path = os.path.join(BASE_OUTPUT, name)
        if not os.path.isdir(path):
            continue
        match_files = glob.glob(os.path.join(path, f'{prefix}_*.xlsx'))
        match_files = [f for f in match_files if 'Untested' not in f and '差异分析' not in f]
        if match_files:
            seen.add(name)
            in_path = os.path.join(BASE_INPUT, name)
            folders.append((name, in_path if os.path.exists(in_path) else path))
        else:
            print(f'[SKIP] {name}: 无匹配结果文件，跳过')
    return folders


def main(args=None):
    args = args or argparse.Namespace(start=None, skip_list=None)
    diff_only = getattr(args, 'diff_only', False)
    match_diff_only = getattr(args, 'match_diff_only', False)
    print(f'[开始] 自动化任务 - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    if diff_only:
        print('[模式] 仅执行差异分析（需已有三单匹配结果）')
    elif match_diff_only:
        print('[模式] 仅执行三单匹配+差异分析（跳过三单清单生成）')
    print(f'扫描目录: {BASE_INPUT}')

    # 启动时清空 pickle 缓存（diff-only 时删除此前运行产生的多余 pkl 文件）
    if getattr(args, 'clear_pickle_at_start', True):
        print('[INFO] 清空 pickle 缓存...')
        _clear_pickle_cache()
        print('[OK] 已清空')

    if not os.path.exists(BASE_INPUT):
        print(f'[ERROR] 数据根目录不存在: {BASE_INPUT}')
        return

    if diff_only:
        folders = _get_companies_for_diff_only()
        print(f'找到 {len(folders)} 个公司（具备匹配结果文件）')
    else:
        folders = []
        for name in os.listdir(BASE_INPUT):
            path = os.path.join(BASE_INPUT, name)
            if os.path.isdir(path) and name.isdigit():
                vbak = os.path.join(path, 'VBAK_001.TXT')
                vbap = os.path.join(path, 'VBAP_001.TXT')
                if os.path.exists(vbak) and os.path.exists(vbap):
                    folders.append((name, path))
                else:
                    print(f'[SKIP] {name}: 缺少 {"+".join(MIN_TABLES)}，跳过')
        folders = sorted(folders, key=lambda x: x[0])
        if match_diff_only:
            folders = [(c, p) for c, p in folders if _has_three_lists(c)]
    start_from = getattr(args, 'start', None)
    skip_list = getattr(args, 'skip_list', None) or []
    defer_list = getattr(args, 'defer_list', None) or []  # 延后执行的公司（如 4390 耗时长，最后跑）
    if start_from:
        folders = [(c, p) for c, p in folders if c >= start_from]
        print(f'从公司 {start_from} 起，共 {len(folders)} 个公司')
    else:
        hint = '已存在三单清单' if match_diff_only else '具备 VBAK+VBAP'
        print(f'找到 {len(folders)} 个公司（{hint}）')
    if skip_list:
        folders = [(c, p) for c, p in folders if c not in skip_list]
        print(f'跳过公司: {", ".join(skip_list)}，剩余 {len(folders)} 个')
    # 延后执行：主批次排除 defer 公司，defer 公司放在最后
    if defer_list:
        main_batch = [(c, p) for c, p in folders if c not in defer_list]
        deferred = [(c, p) for c, p in folders if c in defer_list]
        folders = main_batch + deferred
        if deferred:
            print(f'延后执行: {", ".join(c for c, _ in deferred)}（最后处理）')
    if not folders:
        print('[WARNING] 无待处理公司')
        return
    full_scripts = [
        ('three_lists.py', '三单清单'),
        ('sales_three_match.py', '三单匹配'),
        ('difference_analysis.py', '差异分析'),
    ] if not diff_only else [('difference_analysis.py', '差异分析')]
    match_diff_scripts = [
        ('sales_three_match.py', '三单匹配'),
        ('difference_analysis.py', '差异分析'),
    ] if not diff_only else [('difference_analysis.py', '差异分析')]
    workers = getattr(args, 'workers', 1) or 1
    clear_pickle_per_company = getattr(args, 'clear_pickle', False) and (workers <= 1)  # 多进程时不清理，避免误删他进程缓存

    if workers <= 1:
        # 单进程顺序执行
        for code, company_folder in folders:
            if not diff_only:
                ok, missing = _check_company_tables(company_folder)
                if not ok:
                    print(f'\n[INFO] 公司 {code} 缺少: {", ".join(missing)}（将产生空清单/匹配结果）')
            print(f'\n{"="*60}', flush=True)
            print(f'处理公司: {code}', flush=True)
            print(f'数据目录: {company_folder}', flush=True)
            print('='*60, flush=True)
            if diff_only or match_diff_only or _has_three_lists(code):
                if not diff_only and _has_three_lists(code) and not match_diff_only:
                    print(f'[INFO] 公司 {code} 已有三单清单，跳过生成，直接执行匹配与差异分析', flush=True)
                scripts_to_run = match_diff_scripts
            else:
                scripts_to_run = full_scripts
            c, fail = _process_one_company(code, company_folder, scripts_to_run, clear_pickle_per_company)
            if fail:
                print(f'[WARNING] 公司 {c} {fail} 返回码非0', flush=True)
    else:
        # 多进程并行执行
        print(f'[INFO] 多进程模式: {workers} 个 worker 并行', flush=True)
        failed = []
        def _scripts_for(code):
            if diff_only or match_diff_only or _has_three_lists(code):
                return match_diff_scripts
            return full_scripts
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(_process_one_company, code, folder, _scripts_for(code), clear_pickle_per_company): code
                for code, folder in folders
            }
            for future in as_completed(futures):
                code = futures[future]
                try:
                    c, fail = future.result()
                    if fail:
                        failed.append((c, fail))
                        print(f'[WARNING] 公司 {c} {fail} 返回码非0', flush=True)
                    else:
                        print(f'[OK] 公司 {c} 完成', flush=True)
                except Exception as e:
                    print(f'[ERROR] 公司 {code} 异常: {e}', flush=True)
                    failed.append((code, str(e)))
        if failed:
            print(f'\n[汇总] {len(failed)} 家公司存在问题: {failed}', flush=True)
    print('\n' + '='*60)
    print(f'[完成] 全部公司处理完成 - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    print('='*60)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='按公司依次执行三单清单→三单匹配→差异分析')
    parser.add_argument('--start', type=str, default=None, metavar='CODE', help='从指定公司代码起执行，如 --start 4100')
    parser.add_argument('--skip', type=str, default=None, metavar='CODE', help='跳过指定公司，如 --skip 4390；多个用逗号分隔 --skip 4390,4391')
    parser.add_argument('--defer', type=str, default=None, metavar='CODE', help='延后执行的公司（最后处理），如 --defer 4390；多个用逗号分隔')
    parser.add_argument('--workers', '-w', type=int, default=1, metavar='N', help='并行进程数，如 --workers 4')
    parser.add_argument('--clear-pickle', action='store_true', help='每公司处理前清理 pickle 缓存（仅单进程时生效）')
    parser.add_argument('--clear-pickle-at-start', action='store_true', help='启动时清空 pickle 缓存（默认不清空，建议先手动 rm）')
    parser.add_argument('--diff-only', action='store_true', help='仅执行差异分析（需已有三单匹配结果；会清空 pickle 多余文件）')
    parser.add_argument('--match-diff-only', action='store_true', help='仅执行三单匹配+差异分析，不生成/不保存三单清单；数据从 TXT 或已有 pickle 读取，可加速后续公司')
    args = parser.parse_args()
    args.diff_only = getattr(args, 'diff_only', False)
    args.clear_pickle_at_start = getattr(args, 'clear_pickle_at_start', False)
    args.skip_list = [s.strip() for s in (args.skip or '').split(',') if s.strip()]
    args.defer_list = [s.strip() for s in (args.defer or '').split(',') if s.strip()]
    main(args)
