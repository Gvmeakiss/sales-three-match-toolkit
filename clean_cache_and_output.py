#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
清理缓存与除 4390 外各公司生成结果，便于重跑全流程（保留 4390）。
- 清空 PICKLE_FOLDER、LISTS_PICKLE_FOLDER 下所有 .pkl
- 除公司 4390 外，删除各公司目录下的 销售三单清单、SalesThreeMatchResult*.xlsx
- 删除 OutPut 根目录下的 汇总_差异分析_全公司.xlsx
用法: python clean_cache_and_output.py
"""

import os
import sys
import glob

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from config import PICKLE_FOLDER, LISTS_PICKLE_FOLDER, OUTPUT_ROOT

KEEP_COMPANY = '4390'


def _rm_pkl_in(folder):
    if not os.path.isdir(folder):
        return 0
    n = 0
    for f in os.listdir(folder):
        if f.endswith('.pkl'):
            try:
                os.remove(os.path.join(folder, f))
                n += 1
            except OSError:
                pass
    return n


def main():
    output_root = os.path.normpath(os.path.abspath(OUTPUT_ROOT))
    pickle_folder = os.path.normpath(os.path.abspath(PICKLE_FOLDER))
    lists_folder = os.path.normpath(os.path.abspath(LISTS_PICKLE_FOLDER))

    print('步骤1: 清空 pickle 缓存')
    n1 = _rm_pkl_in(pickle_folder)
    n2 = _rm_pkl_in(lists_folder)
    print(f'  已删除 {n1 + n2} 个 .pkl 文件')

    print('步骤2: 删除除 4390 外各公司生成结果')
    if not os.path.isdir(output_root):
        print(f'  OutPut 目录不存在: {output_root}')
    else:
        for name in os.listdir(output_root):
            path = os.path.join(output_root, name)
            if not os.path.isdir(path):
                continue
            if name == KEEP_COMPANY:
                print(f'  保留: {name}/')
                continue
            deleted = 0
            lists_dir = os.path.join(path, '销售三单清单')
            if os.path.isdir(lists_dir):
                for f in os.listdir(lists_dir):
                    try:
                        os.remove(os.path.join(lists_dir, f))
                        deleted += 1
                    except OSError:
                        pass
                try:
                    os.rmdir(lists_dir)
                except OSError:
                    pass
            for f in glob.glob(os.path.join(path, 'SalesThreeMatchResult*.xlsx')):
                try:
                    os.remove(f)
                    deleted += 1
                except OSError:
                    pass
            if deleted:
                print(f'  已清理 {name}/: {deleted} 项')

    print('步骤3: 删除汇总文件')
    summary = os.path.join(output_root, '汇总_差异分析_全公司.xlsx')
    if os.path.isfile(summary):
        try:
            os.remove(summary)
            print('  已删除 汇总_差异分析_全公司.xlsx')
        except OSError as e:
            print(f'  删除失败: {e}')
    else:
        print('  汇总文件不存在，跳过')

    print('[OK] 清理完成')


if __name__ == '__main__':
    main()
