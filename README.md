# 销售三单匹配工具包 🧰

> 面向 SAP SD 的销售订单-交货单-发票三单匹配与差异分析流水线，按 `(VKORG, VBELN, POSNR)` 关联三张单据，自动定位金额/数量差异并输出审计底稿。

[![Language](https://img.shields.io/badge/language-Python-blue)](https://github.com/Gvmeakiss/sales-three-match-toolkit) [![License](https://img.shields.io/badge/license-MIT-green)](https://github.com/Gvmeakiss/sales-three-match-toolkit/blob/main/LICENSE) [![Domain](https://img.shields.io/badge/domain-Audit%20Analytics-orange)](https://github.com/Gvmeakiss/sales-three-match-toolkit)

## 📌 项目简介
面向 NewHope 客户的 SAP 销售业务，把 **销售订单（VBAK/VBAP）、交货单（LIKP/LIPS）、销售发票（VBRK/VBRP）** 三类 SD 单据按公司代码拆分后逐单匹配，核对订单-交货-开票在金额与数量上是否一致。工具包以配置驱动（见 `config.py`），支持 30GB+ 源数据下的内存优化，输出逐公司匹配结果、未匹配（Untested）清单与按业务场景归类的差异分析，供审计人员快速锁定差异。

## ✨ 功能特性
- **完整流水线**：`split_by_vkorg` → `three_lists` → `sales_three_match` → `difference_analysis` → `aggregate_difference_summary`，由 `run_full_year_pipeline.sh` / `run_all_by_company.py` 编排。
- **三单匹配**：以发票为基准 `left join` 订单与交货，关联键 `(VKORG, VBELN, POSNR)`（见 `sales_three_match.py`）。
- **负开票冲账处理**：订单数量≈0 且（交货≈0 或未匹配）且发票金额<0 的行，从主结果剔除并归入 Untested 的「负开票冲账」表（见 `sales_three_match.py` 注释与 `adjust_scenario5_diff.py`）。
- **Untested 四表**：未匹配结果拆为 4 个 sheet —— 仅订单 / 仅订单及发货单 / 仅发货单 / 负开票冲账。
- **场景化差异分析**：`difference_analysis.py` 按业务场景（场景定义见 `docs/场景判断条件_模块条件.md`）分类差异，输出场景汇总与各场景逐行明细，可进一步汇入四大类。
- **配置驱动**：`config.py` 集中管理数据路径、期间范围、订单类型排除、内存优化与大文件分块等参数。
- **大表内存优化**：按文件夹体积自动或强制开启分块读取（`CHUNK_SIZE_FOR_LARGE`、按日期列 `LFDAT`/`FKDAT` 过滤），适配 48GB 内存处理 30GB+ 源数据。
- **TXT 解析缓存**：`utils/DylanTools.kpmg_txt_to_df` 读取 KPMG 取数导出的 TXT，结果按公司代码缓存为 PKL（`VBAK_001_4390.pkl`），重跑免重复解析。
- **辅助工具**：`split_by_vkorg.py`（按 BUKRS 拆分）、`split_untested_4390.py`、`summarize_untested.py`、`verify_order_qty_zero.py`、`extract_scenarios_to_output.py` 等。

## 📂 目录结构
```
sales-three-match-toolkit/
├── README.md                         # 项目说明
├── LICENSE                           # MIT License
├── requirements.txt                  # pandas/numpy/openpyxl/chardet
├── config.py                         # 集中配置（路径/期间/排除/内存）
├── three_lists.py                    # 生成订单/交货/发票三单清单
├── sales_three_match.py              # 三单匹配与导出（核心）
├── difference_analysis.py            # 场景分类与差异统计
├── aggregate_difference_summary.py  # 全公司差异汇总
├── split_by_vkorg.py                 # 按公司代码拆分原始 TXT
├── run_all_by_company.py             # 逐公司编排流水
├── run_full_year_pipeline.sh         # 全年完整流程入口
├── rerun_all_clean.sh                # 清理后按新标准重跑
├── cleanup_h1_raw.sh / watch_and_run_4390.sh  # H1 数据整理与 4390 监控
├── adjust_scenario5_diff.py / remove_scenario10_from_diff.py / extract_scenarios_to_output.py  # 场景修正/抽取
├── split_untested_4390.py / summarize_untested.py / verify_order_qty_zero.py  # Untested 处理
├── utils/
│   ├── DylanTools.py                 # kpmg_txt_to_df（TXT→DataFrame）
│   └── path_utils.py                # 路径/目录工具
└── docs/                             # 源数据逻辑、场景条件、分类逻辑等核对文档
```

## 🔧 环境要求
- Python 3（建议 3.8+）
- 依赖（来自 `requirements.txt`）：`pandas>=1.3.0`、`numpy>=1.20.0`、`openpyxl>=3.0.0`、`chardet>=4.0.0`

## 🚀 安装
```bash
git clone https://github.com/Gvmeakiss/sales-three-match-toolkit.git
cd sales-three-match-toolkit
pip install -r requirements.txt
```

## 💡 快速开始 / 使用示例
```bash
# 全年完整流程（三单清单→匹配→差异分析→汇总，处理全部公司）
bash run_full_year_pipeline.sh 1

# 从指定公司（如 4390）起执行
bash run_full_year_pipeline.sh 1 4390

# 清理后按新标准全流程重跑并汇总
bash rerun_all_clean.sh

# 仅三单匹配 + 差异分析（需已有三单清单）
python run_all_by_company.py --match-diff-only

# 单公司手动串联：三单清单 → 匹配 → 差异分析
SALES_DATA_FOLDER=InPut/4390 python three_lists.py
SALES_DATA_FOLDER=InPut/4390 python sales_three_match.py
SALES_DATA_FOLDER=InPut/4390 python difference_analysis.py
```

## 🧠 核心逻辑（方法论）
1. **拆分**：`split_by_vkorg.py` 按 `BUKRS` 把原始 TXT 拆到 `InPut/{公司代码}`（如 4010、4030、4390）。
2. **三单清单**：`three_lists.py` 分别从 VBAK/VBAP、LIKP/LIPS、VBRK/VBRP 生成订单/交货/发票清单，按月期间（`ORDER_*`/`DELIVERY_*`/`INVOICE_*`）筛选，剔除 `ORDER_TYPE_EXCLUDE` 中的单据类型（AB 取消、ZTR 调拨、ZFT 赠品、ZRE 退货、ZCR/ZDR 贷/借项、ZFD 框架）。
3. **匹配**：`sales_three_match.py` 以发票为基准按 `(VKORG, VBELN, POSNR)` 左连订单与交货，计算金额/数量差额；负开票冲账行剔除至 Untested。
4. **差异分析**：`difference_analysis.py` 按业务场景（定义见 `docs/场景判断条件_模块条件.md`）归类差异，输出场景明细汇总与逐行清单。
5. **汇总**：`aggregate_difference_summary.py` 把各公司差异分析合并为单 Excel `OutPut/汇总_差异分析_全公司.xlsx`。

## 📋 输入与输出
- **输入**：KPMG 取数导出的 SAP SD TXT 文件（VBAK/VBAP、LIKP/LIPS、VBRK/VBRP 等，命名规则在 `config.py` 通配匹配，排除 `VBPA_VBAK`、`VBAK_VBRP` 等辅助表）；置放于 `InPut/` 或由 `SALES_DATA_ROOT` / `SALES_DATA_FOLDER` 指定。
- **输出**：逐公司 `OutPut/<公司代码>/`，含匹配结果、Untested（4 sheet）、差异分析；全公司汇总 `OutPut/汇总_差异分析_全公司.xlsx`。Excel 单文件超过 `EXCEL_MAX_ROWS_PER_FILE` 时自动分文件。

## ⚙️ 配置说明
`config.py` 主要开关：
- **路径**：`SALES_DATA_ROOT`（默认脚本上级目录）、`SALES_DATA_FOLDER`（多进程并行时指定公司目录）、`INPUT_ROOT`/`OUTPUT_ROOT`（InPut / OutPut）、`PICKLE_FOLDER`。
- **期间范围**：`ORDER_YEAR/MONTH_START/END`、`DELIVERY_*`、`INVOICE_*`（默认 2025 年 1–12 月）。
- **筛选**：`ORDER_TYPE`（仅保留指定 AUART，默认 None）、`ORDER_TYPE_EXCLUDE`、可选 `MATNR_MAX`。
- **内存优化**：`MEMORY_SAVE_MODE`、`LARGE_FOLDER_THRESHOLD_GB`、`LARGE_FILE_MB`、`CHUNK_SIZE_FOR_LARGE`（默认 150000）。
- **导出**：`EXCEL_MAX_ROWS_PER_FILE`、`USE_RANDOM_SUFFIX`、`OUTPUT_PREFIX`。

## ⚠️ 注意事项
- 数据脱敏：不含真实客户业务数据，目录示例中的客户名 NewHope 为脱敏化名；实际运行需用户提供自有 SAP 取数文件。
- 口径说明：订单类型排除、负开票冲账、场景分类口径以 `config.py` 与各 `docs/` 文档为准，审计归类前请先核对 `docs/差异分析_分类逻辑审阅.md`。
- 多公司并行：通过 `SALES_DATA_FOLDER` 为不同进程指定不同公司目录，PKL 按公司代码命名避免混用。

## 🔗 相关仓库
- https://github.com/Gvmeakiss/sales-oms-dms-match
- https://github.com/Gvmeakiss/sales-three-match-newhope

## 📄 License
MIT（Copyright © 2026 Gvmeakiss (James Li)）。

---

<div align="center">

*Disclaimer: Personal project and personal views. Not affiliated with or endorsed by KPMG or any client.*<br>
*本仓库为个人项目与个人观点，与任何前/现雇主及客户无关。*

</div>
