# -*- coding: utf-8 -*-
"""
配置 - 销售三单匹配工具包
集中配置项，供 three_lists、sales_three_match、difference_analysis、run_all_by_company、aggregate_difference_summary 等脚本使用。
"""

import os

# 脚本所在目录（SalesThreeMatchToolkit），用于派生 pickle 等路径
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据根目录：优先环境变量 SALES_DATA_ROOT（便于换电脑/换环境），否则为脚本上一级目录（如 NewHope）
# Mac 示例：/Users/aatrox/Desktop/NewHope
_DATA_ROOT = os.environ.get('SALES_DATA_ROOT', os.path.join(_SCRIPT_DIR, '..'))
DATA_ROOT = os.path.normpath(os.path.abspath(_DATA_ROOT))

# ============================================================
# 数据路径
# ============================================================
# 单公司数据目录，指向某公司拆分后的 InPut 子目录。
# 环境变量 SALES_DATA_FOLDER 可覆盖（用于多进程并行时每进程指定不同目录）
# 使用于：three_lists, sales_three_match, difference_analysis（通过 get_output_folder）
_DATA_FOLDER_DEFAULT = os.path.join(DATA_ROOT, 'InPut')
DATA_FOLDER = os.path.normpath(os.path.abspath(
    os.environ.get('SALES_DATA_FOLDER') or _DATA_FOLDER_DEFAULT
))

# 自动化扫描根目录，run_all_by_company 从此目录下查找公司子文件夹（如 4010、4030）
# 使用于：run_all_by_company
INPUT_ROOT = os.path.join(DATA_ROOT, 'InPut')

# 输出根目录，各公司结果输出至 OutPut/<公司代码>
# 使用于：three_lists, sales_three_match, difference_analysis, aggregate_difference_summary
OUTPUT_ROOT = os.path.join(DATA_ROOT, 'OutPut')

# TXT 转 DataFrame 的 pickle 缓存目录，按公司代码命名（如 VBAK_001_4390.pkl），避免多公司混用
# 使用于：sales_three_match, three_lists；run_all_by_company 不再自动清理，便于修改代码后重跑
PICKLE_FOLDER = os.path.join(_SCRIPT_DIR, 'pickle')

# ============================================================
# 日期范围（用于按期间筛选订单、交货、发票）
# ============================================================
# 订单筛选：按 VBAK.AUDAT 或 ERDAT（记录建立日期）的年月范围
# 使用于：three_lists, sales_three_match
ORDER_YEAR = 2025
ORDER_MONTH_START = 1
ORDER_MONTH_END = 12

# 快速测试：True 时仅处理 1 个月，False 时按 ORDER_MONTH_START/END 全范围
# 使用于：three_lists
QUICK_TEST_ONE_MONTH = False

# 交货筛选：按 LIPS.LFDAT（交货日期）
# 使用于：three_lists, sales_three_match
DELIVERY_YEAR = 2025
DELIVERY_MONTH_START = 1
DELIVERY_MONTH_END = 12

# 发票筛选：按 VBRP.FKDAT（开票日期）
# 使用于：three_lists, sales_three_match
INVOICE_YEAR = 2025
INVOICE_MONTH_START = 1
INVOICE_MONTH_END = 12

# ============================================================
# 筛选条件（可选，用于缩小匹配范围）
# ============================================================
# 订单类型：VBAK.AUART（销售凭证类型），仅保留 AUART == 该值的订单；None 表示不过滤
# 示例：'ZOR' 表示仅标准订单
# 使用于：three_lists, sales_three_match, summarize_untested
ORDER_TYPE = None

# 订单类型排除：AUART 在此列表中的订单一律剔除，同时作用于 sales_three_match（Untested 生成）和 summarize_untested（汇总）
# 分类说明：
#   AB   - 取消单（Cancellation）：已取消，无业务意义，明确剔除
#   ZTR  - 内部调拨单（Transfer Order）：集团内部货物调拨，不产生对外收入，明确剔除
#   ZFT  - 免费赠品单（Free of Charge）：不产生收入，明确剔除
#   ZRE  - 退货单（Returns）：属于逆向流程，是否纳入视审计目的而定，暂剔除
#   ZCR  - 贷项凭证申请（Credit Memo Request）：价格/数量调整，暂剔除
#   ZDR  - 借项凭证申请（Debit Memo Request）：价格/数量调整，暂剔除
#   ZFD  - 框架交货协议（Frame Delivery）：框架订单，暂剔除（与实际发货无直接对应）
# 如需包含退货/贷项等，将对应类型从列表中移除即可
ORDER_TYPE_EXCLUDE = ['AB', 'ZTR', 'ZFT', 'ZRE', 'ZCR', 'ZDR', 'ZFD']

# 物料号上限：MATNR（物料号）小于此值的行才参与；None 表示不限制，用于测试时缩小数据量
# 使用于：three_lists, sales_three_match
MATNR_MAX = None

# ============================================================
# 大表内存优化（源数据 30GB+、本机 48GB 内存时建议开启）
# ============================================================
# True=强制开启，False=按数据文件夹大小自动判断（>LARGE_FOLDER_THRESHOLD_GB 时开启）
# 使用于：sales_three_match
MEMORY_SAVE_MODE = False

# 数据文件夹（InPut/公司代码）总大小超过此 GB 时自动开启大表优化；仅当 MEMORY_SAVE_MODE=False 时生效
LARGE_FOLDER_THRESHOLD_GB = 10

# 大表分块大小（行）；生效时使用
CHUNK_SIZE_FOR_LARGE = 150000

# 单文件超过此 MB 时强制分块+日期过滤；0 表示仅依赖 MEMORY_SAVE_MODE 或文件夹阈值
LARGE_FILE_MB = 500

# ============================================================
# 文件匹配（KPMG 取数导出的 TXT 命名规则）
# ============================================================
# 通配符匹配主表，排除 VBPA_VBAK、VBAK_VBRP 等辅助表（KPMG 导出会带 _数字 后缀）
# 使用于：sales_three_match, three_lists
VBAK_FILE = 'VBAK_[0-9]*.TXT'   # 排除 VBPA_VBAK_*, VBAK_VBRP_*
VBAP_FILE = 'VBAP_[0-9]*.TXT'   # 排除 VBAP_VBRP_*
LIKP_FILE = 'LIKP_[0-9]*.TXT'
LIPS_FILE = 'LIPS_[0-9]*.TXT'   # 排除 LIPS_VBRP_*
VBRK_FILE = 'VBRK_[0-9]*.TXT'   # 排除 VBPA_VBRK_*
VBRP_FILE = 'VBRP_[0-9]*.TXT'   # 排除 *_VBRP_* 等

# ============================================================
# 输出
# ============================================================
# 是否在输出文件名末尾加随机数，避免覆盖
# 使用于：sales_three_match
USE_RANDOM_SUFFIX = True

# 三单匹配 / 差异分析输出文件前缀
# 使用于：sales_three_match, difference_analysis, aggregate_difference_summary
OUTPUT_PREFIX = 'SalesThreeMatchResult'

# 公司代码：None 时从 DATA_FOLDER 末段自动解析（如 InPut/4010 -> 4010）
# 使用于：get_company_code，进而被 get_output_folder、sales_three_match、difference_analysis 等引用
COMPANY_CODE = None

# ============================================================
# 三单清单相关（three_lists.py）
# ============================================================
# 单 Excel 最大行数，超出则分片导出；Excel 上限 1048576，此处略小以留余量
# 使用于：three_lists, sales_three_match, difference_analysis
EXCEL_MAX_ROWS_PER_FILE = 980000

# 是否导出订单/交货/发票清单，False 时可跳过以加速测试
# 使用于：three_lists
EXPORT_ORDER_LIST = True
EXPORT_DELIVERY_LIST = True
EXPORT_INVOICE_LIST = True

# 批次去重：False=保留主行+子行供核验批次结构；True=主行有数量时排除其子行避免重复
BATCH_DEDUP_ENABLED = False

# 清单 PKL 缓存目录
# 使用于：three_lists
LISTS_PICKLE_FOLDER = os.path.join(_SCRIPT_DIR, 'pickle', 'lists')
SAVE_LIST_PKL = True
USE_LIST_PKL = False

# 售达方排除：订单(KUNNR)、发票(KUNAG)售达方在以下代码中的记录均剔除（这些代码对应的公司不纳入三单）
# get_exclude_sold_to_codes() = 本名单 ∪ InPut 下所有四位数公司代码（动态扩展），再减去 EXCLUDE_SOLD_TO_CODES_REMOVE
# 使用于：three_lists、sales_three_match
INVOICE_EXCLUDE_SOLD_TO_CODES = {
    '4000', '4010', '4040', '4090', '4100', '4110', '4120', '4130', '4140', '4150',
    '4170', '4180', '4220', '4294', '4300', '4310', '4350', '4390', '4410', '4460',
    '4470', '4510', '4520', '4530', '4540', '4580', '4610', '4620', '4670', '4680',
    '4730', '4740', '4750', '4770', '4780', '4790', '4800', '4810', '4820', '4830',
    '4840', '4870', '4900', '4990', '5010', '5020', '5030', '5050', '5060', '5080',
    '5100', '5140',
}
# 从关联交易（售达方排除）清单中剔除的公司代码，这些售达方将纳入三单处理
EXCLUDE_SOLD_TO_CODES_REMOVE = {'4480'}


def get_exclude_sold_to_codes():
    """售达方排除集合 = INVOICE_EXCLUDE_SOLD_TO_CODES ∪ InPut 下四位数公司代码，再减去 EXCLUDE_SOLD_TO_CODES_REMOVE"""
    codes = set(INVOICE_EXCLUDE_SOLD_TO_CODES)
    try:
        if os.path.exists(INPUT_ROOT):
            for name in os.listdir(INPUT_ROOT):
                path = os.path.join(INPUT_ROOT, name)
                if os.path.isdir(path) and len(name) == 4 and name.isdigit():
                    codes.add(name)
    except OSError:
        pass
    codes -= EXCLUDE_SOLD_TO_CODES_REMOVE
    return codes

# ============================================================
# 匹配键与字段定义（各表 pickle 读取时保留的字段）
# ============================================================
# 用于 three_lists、sales_three_match 读 TXT 时仅保留以下字段，减小内存
# 字段释义见下方各表注释

KEY_FIELDS_VBAK = [
    'MANDT', 'VBELN', 'VKORG', 'AUDAT', 'ERDAT', 'AUART', 'KUNNR', 'VKGRP', 'VKBUR', 'NETWR', 'WAERK'
]
# VBAK 销售订单抬头：MANDT客户端 VBELN销售凭证号 VKORG销售组织 AUDAT请求交货日期 ERDAT创建日期
# AUART订单类型 KUNNR售达方 VKGRP销售组 VKBUR销售办公室 NETWR净价值 WAERK货币

KEY_FIELDS_VBAP = [
    'MANDT', 'VBELN', 'POSNR', 'MATNR', 'ARKTX', 'KWMENG', 'KLMENG', 'UMVKZ', 'UMVKN', 'VRKME', 'NETWR', 'NETPR', 'MEINS',
    'SHKZG', 'RETPO', 'WKURS', 'WAERK', 'WERKS', 'LGORT', 'VSTEL', 'ERDAT', 'AUDAT'
]
# VBAP 销售订单行项：POSNR行号 MATNR物料号 ARKTX物料描述 KWMENG订单数量(销售单位) KLMENG订单数量(基本单位)
# UMVKZ/UMVKN换算因子 VRKME销售单位 NETWR行净价值 NETPR净价 MEINS基本单位 SHKZG借贷 RETPO退货
# WKURS汇率 WERKS工厂 LGORT库位 VSTEL装运点

KEY_FIELDS_LIKP = ['MANDT', 'VBELN', 'VKORG', 'LFART']
# LIKP 交货单抬头：VBELN交货单号 VKORG销售组织 LFART交货类型(LF标准/RL退货等)

KEY_FIELDS_LIPS = [
    'MANDT', 'VBELN', 'POSNR', 'UEPOS', 'UECHA', 'VBELV', 'POSNV', 'VGBEL', 'VGPOS', 'VKORG', 'MATNR', 'ARKTX', 'KUNNR',
    'LFIMG', 'LGMNG', 'UMVKZ', 'UMVKN', 'MEINS', 'VRKME', 'NETWR', 'NETPR', 'SHKZG', 'WERKS', 'LGORT', 'ERDAT', 'LFDAT', 'CHARG'
]
# LIPS 交货单行项：UEPOS/UECHA批次拆分上级行号 VBELV/POSNV参考订单号/行号 VGBEL/VGPOS参考凭证
# LFIMG交货数量(销售单位VRKME) LGMNG交货数量(基本单位MEINS) CHARG批次

KEY_FIELDS_VBRK = ['MANDT', 'VBELN', 'FKART', 'VKORG', 'BUKRS', 'KUNRG', 'KUNAG', 'FKDAT', 'NETWR', 'WAERK', 'BELNR', 'GJAHR', 'ERDAT']
# VBRK 发票抬头：FKART发票类型 BUKRS公司代码 KUNRG收货方 KUNAG售达方 FKDAT开票日期 BELNR会计凭证号 GJAHR会计年度

KEY_FIELDS_VBRP = [
    'MANDT', 'VBELN', 'POSNR', 'VBELV', 'POSNV', 'AUBEL', 'AUPOS',
    'MATNR', 'ARKTX', 'KUNNR', 'BUKRS', 'FKIMG', 'FKLMG', 'UMVKZ', 'UMVKN', 'VRKME', 'NETWR', 'AKKUR', 'MWSBP', 'SHKZG',
    'FKDAT', 'BELNR', 'GJAHR', 'WAERK', 'ERDAT'
]
# VBRP 发票行项：VBELV/POSNV参考交货号/行号 AUBEL/AUPOS销售订单号/行号 FKIMG开票数量(销售单位) FKLMG开票数量(基本单位)
# AKKUR累计汇率 MWSBP税额


def _get_folder_size_gb(folder):
    """计算目录下所有文件总大小（GB）"""
    total = 0
    try:
        for entry in os.scandir(folder):
            if entry.is_file():
                total += entry.stat().st_size
    except OSError:
        pass
    return total / (1024 ** 3)


def get_effective_memory_save_mode(data_folder=None):
    """
    获取是否启用大表优化：MEMORY_SAVE_MODE=True 时强制开启；
    False 时当 data_folder 大小 > LARGE_FOLDER_THRESHOLD_GB 则自动开启。
    """
    if MEMORY_SAVE_MODE:
        return True
    folder = os.path.normpath(data_folder or DATA_FOLDER)
    size_gb = _get_folder_size_gb(folder)
    return size_gb >= LARGE_FOLDER_THRESHOLD_GB


def get_company_code():
    """从 COMPANY_CODE、环境变量 SALES_COMPANY_CODE 或 DATA_FOLDER 末段解析公司代码"""
    env_code = os.environ.get('SALES_COMPANY_CODE', '').strip()
    if env_code and env_code.isdigit():
        return env_code
    if COMPANY_CODE is not None and str(COMPANY_CODE).strip():
        return str(COMPANY_CODE).strip()
    base = os.path.basename(os.path.normpath(DATA_FOLDER))
    return base if (base and base.replace(' ', '').isdigit()) else 'ALL'


def get_output_folder():
    """返回当前公司的输出目录：OUTPUT_ROOT/公司代码"""
    company = get_company_code()
    return os.path.join(OUTPUT_ROOT, company)


# 兼容旧代码引用
OUTPUT_FOLDER = get_output_folder()
