"""
按公司逐个运行 three_lists.py（仅生成三单清单）

本脚本只执行「三单清单」一步，不包含三单匹配与差异分析。
完整流程（清单 → 匹配 → 差异分析）请使用：python run_all_by_company.py

使用方法：
    python run_by_company.py

前置条件：
    1. 已运行 split_by_vkorg.py 完成数据拆分
    2. InPut 下存在按公司代码命名的子文件夹（如 4010、4390 等）
"""

import os
import sys
import subprocess

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)
try:
    from config import INPUT_ROOT, DATA_FOLDER
except ImportError:
    INPUT_ROOT = os.path.normpath(os.path.join(script_dir, '..', 'InPut'))
    DATA_FOLDER = INPUT_ROOT

# InPut 根目录（拆分后各公司子文件夹的父目录）
BASE_INPUT = os.path.normpath(os.path.abspath(INPUT_ROOT))


def main():
    base_input = BASE_INPUT
    if not os.path.exists(base_input):
        print(f'[ERROR] 数据目录不存在: {base_input}')
        return
    folders = []
    for name in os.listdir(base_input):
        path = os.path.join(base_input, name)
        if os.path.isdir(path) and name.isdigit():
            vbak = os.path.join(path, 'VBAK_001.TXT')
            if os.path.exists(vbak):
                folders.append(name)
    folders = sorted(folders)
    if not folders:
        print('[WARNING] 未找到按 VKORG 拆分的子文件夹，请先运行 split_by_vkorg.py')
        return
    print(f'找到 {len(folders)} 个公司: {", ".join(folders)}')
    for vkorg in folders:
        company_folder = os.path.join(base_input, vkorg)
        print(f'\n{"="*50}')
        print(f'处理公司: {vkorg}')
        print(f'数据目录: {company_folder}')
        print('='*50)
        env = os.environ.copy()
        env['SALES_DATA_FOLDER'] = os.path.normpath(os.path.abspath(company_folder))
        env['SALES_COMPANY_CODE'] = str(vkorg)
        result = subprocess.run(
            [sys.executable, os.path.join(script_dir, 'three_lists.py')],
            cwd=script_dir,
            capture_output=False,
            env=env
        )
        if result.returncode != 0:
            print(f'[WARNING] 公司 {vkorg} 处理返回码: {result.returncode}')
    print('\n全部公司处理完成。')


if __name__ == '__main__':
    main()
