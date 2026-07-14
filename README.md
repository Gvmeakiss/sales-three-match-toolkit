# 销售三单匹配工具包

> **项目类型**：SAP 销售订单-交货单-发票三单匹配与差异分析  
> **主要语言**：Python 3  
> **依赖**：pandas, numpy, openpyxl, chardet  
> **平台**：Mac（已针对 macOS 优化）  
> **阅读指引**：新读者与 AI 建议先看 §0 快速开始，再按 §2 数据流理解执行顺序；详细脚本说明见 §4。

<!--
AI 上下文（供 AI 快速理解）：
- 核心流程：split_by_vkorg → three_lists → sales_three_match → difference_analysis → aggregate_difference_summary
- 主匹配：以发票为基准 left join 订单/交货，按 (VKORG, VBELN, POSNR) 关联
- 负开票冲帐：订单数量≈0 且 (交货≈0 或未匹配) 且 发票金额<0 → 从主结果剔除，归入 Untested 的「负开票冲帐」sheet
- Untested 文件：4 个 sheet = 仅订单 | 仅订单及发货单 | 仅发货单 | 负开票冲帐
- 差异分析：按场景标号(1-13)分类，输出场景明细汇总 + 逐行明细清单；用户可自行将场景汇入四大类
- 关键脚本：sales_three_match.py（匹配+导出）、difference_analysis.py（场景分类+统计）
- 场景定义：docs/场景判断条件_模块条件.md
-->

---

## 0. 快速开始（运行命令速查）

### 0.1 运行命令一览

| 场景 | 命令 | 说明 |
|------|------|------|
| **全年完整流程** | `bash run_full_year_pipeline.sh 1` | 三单清单→三单匹配→差异分析→汇总，处理所有公司，1 进程 |
| **从某公司起** | `bash run_full_year_pipeline.sh 1 4010` | 从 4010 起执行，1 进程 |
| **清理重跑** | `bash rerun_all_clean.sh` | 删除所有匹配/差异分析文件，按新标准全流程重跑并汇总 |
| **仅匹配+差异分析** | `python run_all_by_company.py --match-diff-only` | 跳过三单清单生成，仅执行三单匹配与差异分析（需已有三单清单） |
| **仅匹配+差异分析+汇总** | `python run_all_by_company.py --match-diff-only && python aggregate_difference_summary.py` | 同上，并执行全公司汇总 |
| **仅汇总** | `python aggregate_difference_summary.py` | 汇总各公司差异分析至单 Excel |
| **交叉汇总表** | `python summarize_three_match_crosstab.py --company 4150` | 按数量×金额差异类型生成交叉汇总（需已有匹配结果） |
| **H1 追加拆分** | `bash run_h1_append_split.sh` | 展平→备份→上半年追加到 InPut |
| **清理 H1 原始** | `bash cleanup_h1_raw.sh` | 删除空目录、已解压 ZIP |
| **单公司处理** | `SALES_DATA_FOLDER=InPut/4010 python three_lists.py` | 需依次执行 three_lists → sales_three_match → difference_analysis |

### 0.2 执行流程图（数据流）

```
原始 TXT (InPut 或 InPut_H1_RAW)
        │
        ▼
┌───────────────────┐
│ split_by_vkorg.py │  按 BUKRS 拆分 → InPut/{4010,4030,...}
└─────────┬─────────┘
          │
          ▼
┌───────────────────────┐
│ run_all_by_company.py │  对每公司依次执行：
│  (内部调用)            │  three_lists → sales_three_match → difference_analysis
└─────────┬─────────────┘
          │
          ▼
┌─────────────────────────────────┐
│ aggregate_difference_summary.py │  汇总 → OutPut/汇总_差异分析_全公司.xlsx
└─────────────────────────────────┘
```

### 0.3 脚本依赖关系

```
run_full_year_pipeline.sh
    ├── run_all_by_company.py
    │       ├── three_lists.py      (读 InPut/{公司}/)
    │       ├── sales_three_match.py
    │       └── difference_analysis.py
    └── aggregate_difference_summary.py  (读 OutPut/{公司}/差异分析)

rerun_all_clean.sh
    ├── 删除 OutPut/*/SalesThreeMatchResult*.xlsx（含匹配结果、Untested、差异分析）
    ├── run_all_by_company.py      (三单清单→匹配→差异分析)
    └── aggregate_difference_summary.py

run_h1_append_split.sh
    ├── flatten_h1_to_root.sh      (InPut_H1_RAW 子目录→根目录链接)
    └── split_by_vkorg.py --append  (读 InPut_H1_RAW，写 InPut/{公司}/)
```

### 0.4 所有可执行脚本速查

| 类型 | 脚本 | 用途 | 典型命令 |
|------|------|------|----------|
| Shell | `run_full_year_pipeline.sh` | 全年流程（三单清单→匹配→差异分析→汇总） | `bash run_full_year_pipeline.sh 1` |
| Shell | `rerun_all_clean.sh` | 删除旧结果后按新标准全流程重跑并汇总 | `bash rerun_all_clean.sh` |
| Shell | `run_h1_append_split.sh` | H1 追加拆分（展平→备份→split --append） | `bash run_h1_append_split.sh` |
| Shell | `flatten_h1_to_root.sh` | H1 展平：子目录 TXT 链接到根目录 | 由 run_h1_append_split 调用 |
| Shell | `cleanup_h1_raw.sh` | 清理 InPut_H1_RAW 空目录、已解压 ZIP | `bash cleanup_h1_raw.sh` |
| Python | `split_by_vkorg.py` | 按 BUKRS 拆分原始 TXT | `python split_by_vkorg.py` / `--append` |
| Python | `run_all_by_company.py` | 按公司批量执行三单流程 | `python run_all_by_company.py` / `--start 4010` |
| Python | `three_lists.py` | 生成三单清单 | `SALES_DATA_FOLDER=InPut/4010 python three_lists.py` |
| Python | `sales_three_match.py` | 三单匹配 | 同上，需先 three_lists |
| Python | `difference_analysis.py` | 差异分析 | 同上，需先 sales_three_match |
| Python | `aggregate_difference_summary.py` | 汇总各公司差异分析 | `python aggregate_difference_summary.py` |
| Python | `summarize_three_match_crosstab.py` | 三单匹配按数量×金额差异类型交叉汇总 | `python summarize_three_match_crosstab.py --company 4150` |

**前置**：`cd SalesThreeMatchToolkit`，数据已按公司拆分在 `InPut/{公司代码}/`。

---

## 1. 概述

本工具包用于匹配销售订单(VBAK/VBAP)、交货单(LIKP/LIPS)、销售发票(VBRK/VBRP)三单数据，并输出差异分析汇总。

| 功能 | 说明 |
|------|------|
| 数据拆分 | 按公司(BUKRS)拆分原始 TXT，便于分公司独立处理 |
| 三单清单 | 生成订单、交货、发票三类清单 |
| 三单匹配 | 按订单行关联交货与发票，输出匹配结果 |
| 差异分析 | 按场景标号(1-13)分类，输出每场景记录数/金额，及逐行明细；用户可自行汇入四大类 |
| 汇总差异分析 | 合并各公司差异分析至单 Excel，含场景明细、四大类重算、分析报告 |

---

## 2. 数据流与依赖

| 步骤 | 输入 | 输出 | 依赖 |
|------|------|------|------|
| split_by_vkorg | InPut/*.TXT 或 InPut_H1_RAW/*.TXT | InPut/{公司}/六张表 | 无 |
| three_lists | InPut/{公司}/六张表 | OutPut/{公司}/销售三单清单/ | split 完成 |
| sales_three_match | InPut/{公司}/六张表 | OutPut/{公司}/SalesThreeMatchResult_*.xlsx | three_lists（可跳过） |
| difference_analysis | SalesThreeMatchResult_*.xlsx | SalesThreeMatchResult_差异分析_*.xlsx | sales_three_match |
| aggregate_difference_summary | 各公司差异分析文件 | 汇总_差异分析_全公司.xlsx | 所有公司 difference_analysis 完成 |
| summarize_three_match_crosstab | SalesThreeMatchResult_*.xlsx + Untested | 三单匹配交叉汇总_*.xlsx | sales_three_match（可独立于差异分析运行） |

---

## 3. 组成结构

### 3.1 代码目录结构（SalesThreeMatchToolkit）

```
SalesThreeMatchToolkit/
├── config.py                      # 配置入口
├── split_by_vkorg.py              # 按 BUKRS/VKORG 拆分（支持 --append）
├── three_lists.py                 # 三单清单
├── sales_three_match.py           # 三单匹配
├── difference_analysis.py         # 差异分析
├── run_all_by_company.py          # 按公司批量运行（核心流程）
├── aggregate_difference_summary.py # 汇总差异分析
├── summarize_three_match_crosstab.py # 三单匹配按数量×金额差异类型交叉汇总
├── flatten_h1_to_root.sh          # H1 展平：子目录 TXT 链接到根目录
├── run_h1_append_split.sh         # H1 追加拆分：展平→备份→split --append
├── run_full_year_pipeline.sh       # 全年流程：三单清单→匹配→差异分析→汇总
├── rerun_all_clean.sh              # 清理旧结果后全流程重跑并汇总
├── cleanup_h1_raw.sh               # 清理 InPut_H1_RAW：空目录、已解压 ZIP
├── utils/DylanTools.py            # TXT 读取、编码检测
├── utils/path_utils.py            # Mac 路径工具（ensure_dir、ensure_parent_and_open）
├── pickle/                        # TXT→DataFrame 缓存
├── pickle/lists/                  # 清单缓存
└── requirements.txt               # pandas, numpy, openpyxl, chardet
```

### 3.2 输入目录结构（InPut）

**拆分前**（原始数据，由 split_by_vkorg 读取）：

```
InPut/                             # 或 config.INPUT_ROOT 指向的根目录
├── VBAK_001.TXT                   # 销售订单抬头
├── VBAP_001.TXT                   # 销售订单行项目
├── LIKP_001.TXT                   # 交货单抬头
├── LIPS_001.TXT                   # 交货单行项目
├── VBRK_001.TXT                   # 发票抬头
├── VBRP_001.TXT                   # 发票行项目
└── ...                            # 可含 VBPA_VBAK、VBAK_VBRP 等辅助表（会被排除）
```

**拆分后**（按公司子目录，供 three_lists、sales_three_match 使用）：

```
InPut/
├── 4010/                          # 公司代码（BUKRS）
│   ├── VBAK_001.TXT
│   ├── VBAP_001.TXT
│   ├── LIKP_001.TXT
│   ├── LIPS_001.TXT
│   ├── VBRK_001.TXT
│   └── VBRP_001.TXT
├── 4030/
├── 4100/
└── ...
```

- **数据格式**：KPMG ABAP 取数工具导出的 TXT，分隔符 `#|#`
- **config.DATA_FOLDER** 指向某公司目录（如 `InPut\4010`）时，单公司脚本从此读取

### 3.3 输出目录结构（OutPut）

```
OutPut/
├── 4010/                                    # 按公司代码分目录
│   ├── 销售三单清单/
│   │   ├── 销售订单清单_4010_2025_1-12月_*.xlsx
│   │   ├── 交货单清单_4010_2025_1-12月_*.xlsx
│   │   └── 销售发票清单_4010_2025_1-12月_*.xlsx
│   ├── SalesThreeMatchResult_4010_2025_1-12_*.xlsx     # 主匹配结果
│   ├── SalesThreeMatchResult_4010_2025_1-12_*_2.xlsx    # 分片（若超 98 万行）
│   ├── SalesThreeMatchResult_Untested_4010_*.xlsx       # 未测试单据
│   ├── SalesThreeMatchResult_差异分析_4010.xlsx          # 场景明细（汇总）
│   ├── SalesThreeMatchResult_差异分析_4010_详细.xlsx     # 明细清单（行数超限时分 _详细_1、_详细_2 等）
│   └── 三单匹配交叉汇总_4010.xlsx                       # 交叉汇总（summarize_three_match_crosstab 输出）
├── 4030/
├── ...
├── 汇总_差异分析_全公司.xlsx                              # 汇总各公司差异分析
└── 三单匹配交叉汇总_全公司.xlsx                            # 全公司交叉汇总（--all --consolidated）
```

### 3.4 未测试单据 (Untested) 文件结构

`SalesThreeMatchResult_Untested_{公司}_*.xlsx` 含 4 个 sheet，按以下顺序与含义：

| Sheet 名称 | 含义 | 数据来源 |
|------------|------|----------|
| **仅订单** | 有订单、无交货、无发票 | 订单 outer join 交货，过滤无发票 |
| **仅订单及发货单** | 有订单+交货、无发票 | 同上 |
| **仅发货单** | 有交货、无订单、无发票 | 同上 |

- **负开票冲帐**：满足「订单数量≈0 且 (交货数量≈0 或未匹配到交货单) 且 发票金额<0」的行，从主匹配结果中剔除，单独归入本 sheet，不参与差异分析。

---

## 4. 使用顺序及代码解析

### 4.1 整体流程

```
1. split_by_vkorg.py          拆分 InPut 根目录 → 按公司输出 InPut/{公司}
2. run_all_by_company.py      对每家执行：three_lists → sales_three_match → difference_analysis
3. aggregate_difference_summary.py  汇总各公司差异分析 → 单 Excel
```

或单公司手动运行：设置环境变量 `SALES_DATA_FOLDER=InPut/4010` 或修改 `config.DATA_FOLDER` 后，依次执行 three_lists → sales_three_match → difference_analysis。

### 4.2 split_by_vkorg.py — 按公司拆分

**用途**：将原始 TXT 按 BUKRS（公司代码）或 VKORG 拆分为公司子目录。

**输入**：`config.DATA_FOLDER` 下的 VBAK/VBAP/LIKP/LIPS/VBRK/VBRP（*_001.TXT 等）

**输出**：`InPut\{BUKRS}\` 下每公司 6 张表

**运行逻辑**：
1. 从 VBAK/VBRK 构建 VKORG→BUKRS 映射
2. 按 BUKRS 过滤各表行，分批写入对应公司目录
3. 列名校验、编码自动检测

**命令**：
```bash
python split_by_vkorg.py --list          # 列出 BUKRS
python split_by_vkorg.py --code 4010     # 仅拆分 4010
python split_by_vkorg.py                 # 拆分所有公司（默认按 BUKRS）
python split_by_vkorg.py --append        # 追加模式：不覆盖已有数据，追加到现有公司目录
python split_by_vkorg.py --by vkorg      # 按 VKORG 拆分
```

---

### 4.3 three_lists.py — 三单清单

**用途**：生成销售订单、交货单、销售发票三类清单。

**输入**：`config.DATA_FOLDER` 下的 VBAK/VBAP/LIKP/LIPS/VBRK/VBRP（TXT 或 pickle）

**输出**：`OutPut\{公司代码}\销售三单清单\`，含订单/交货/发票清单 Excel

**运行逻辑**：
1. 读取 TXT（优先使用 pickle 缓存），按日期筛选
2. **售达方排除**：剔除订单(KUNNR)、发票(KUNAG) 售达方在排除名单中的记录（名单 = 固定配置 ∪ InPut 下四位数公司代码）
3. 分别聚合 VBAK+VBAP、LIKP+LIPS、VBRK+VBRP
4. 过滤无效行（如 LFIMG=0、FKIMG=NETWR=0 等），导出 Excel

**命令**：
```bash
python three_lists.py
```

---

### 4.4 sales_three_match.py — 三单匹配

**用途**：将订单、交货、发票按订单行关联，输出主匹配结果及未测试单据。

**输入**：`config.DATA_FOLDER` 下的六张表（TXT/pickle）

**输出**：
- `SalesThreeMatchResult_{公司}_{年}_{月范围}_{随机}.xlsx`：主匹配结果（行数 >98 万时分片为 `_2.xlsx`、`_3.xlsx`…）
- `SalesThreeMatchResult_Untested_{公司}_*.xlsx`：未测试单据，含 4 个 sheet：仅订单、仅订单及发货单、仅发货单、负开票冲帐

**运行逻辑**：
1. 读取并合并 VBAK/VBAP、LIKP/LIPS、VBRK/VBRP；VBRP 空 VBELV/POSNV 时用 AUBEL/AUPOS 填充
2. **售达方排除**：剔除订单(KUNNR)、发票(KUNAG) 售达方在排除名单中的记录（名单 = 固定配置 ∪ InPut 下四位数公司代码）
3. 按 (VKORG, VBELN, POSNR) 为订单键，与交货、发票 left join
4. 计算订单-发票金额差异、订单-发票数量差异
5. 导出：单文件 ≤98 万行；超出则分片
6. **未测试单据**：单 sheet 超 Excel 行限(98 万)时，自动拆成多 sheet（如 仅发货单_1、仅发货单_2）
7. **大表优化**：数据文件夹 >10GB 或单文件 >500MB 时，自动分块读取 LIPS/VBRP，块内按日期过滤（参见 §5.1）

**命令**：
```bash
python sales_three_match.py
```

---

### 4.5 difference_analysis.py — 差异分析

**用途**：按场景标号(1-13)对匹配结果分类，输出每场景的汇总及逐行明细；用户可自行将场景汇入四大类。

**输入**：`OutPut\{公司代码}\` 下最新 `SalesThreeMatchResult_*.xlsx`（支持分片 `_2`、`_3` 合并读取）

**输出**：
- `SalesThreeMatchResult_差异分析_{公司}.xlsx`：统计页，含单 sheet「**场景明细**」，列：场景标号、识别场景、记录数、占比、发票金额、发票金额占比；按场景 1→13→负开票 顺序
- `SalesThreeMatchResult_差异分析_{公司}_详细.xlsx`：逐行明细，含场景标号列；行数超 98 万时拆分为 `_详细_1.xlsx`、`_详细_2.xlsx` 等

**场景定义**：见 [docs/场景判断条件_模块条件.md](docs/场景判断条件_模块条件.md)

**运行逻辑**：
1. 在输出目录查找最新匹配文件，按 base 归组分片
2. **空结果容错**：匹配结果为 0 行或缺少必要列时，输出空统计报告，不崩溃
3. 为每行分配场景标号(1-13)，场景 7（缺失发票）从 Untested 的「仅订单及发货单」统计
4. 生成场景明细汇总并导出；明细数据超 98 万行时分文件（大表模式分块处理，降低内存）

**命令**：
```bash
python difference_analysis.py
```

---

### 4.6 run_all_by_company.py — 按公司批量运行（推荐）

**用途**：对 `InPut` 下每家公司依次执行 三单清单 → 三单匹配 → 差异分析。

**输入**：`config.INPUT_ROOT`（默认 `InPut`）下按公司代码命名的子目录

**输出**：`OutPut\{公司代码}\` 下各公司的清单、匹配、差异分析

**运行逻辑**：
1. 扫描 INPUT_ROOT，筛选 `{数字}` 目录且含 VBAK/VBAP
2. 对每家：通过环境变量 `SALES_DATA_FOLDER` 指定数据目录（不再写 config）
3. **已有三单清单则跳过生成**：若 `OutPut/{公司}/销售三单清单/` 已含订单、交货、发票三类清单，则直接执行三单匹配与差异分析
4. **`--match-diff-only`**：仅处理已存在三单清单的公司，跳过 three_lists，只执行三单匹配与差异分析
5. **`--diff-only`**：仅执行差异分析，扫描 OutPut 下已有匹配结果文件的公司
6. 支持 `--start CODE` 从指定公司起执行、`--workers N` 多进程并行
7. pickle 按公司命名，默认不清空；`--clear-pickle` 可强制清理（仅单进程时生效）

**命令**：
```bash
python run_all_by_company.py                    # 处理所有公司（单进程）
python run_all_by_company.py --start 4100       # 从 4100 起执行
python run_all_by_company.py --workers 4       # 4 进程并行（注意内存）
python run_all_by_company.py --match-diff-only  # 仅三单匹配+差异分析，跳过三单清单生成（需已有三单清单）
python run_all_by_company.py --diff-only       # 仅执行差异分析（需已有匹配结果）
```

---

### 4.6a rerun_all_clean.sh — 清理重跑

**用途**：删除所有三单匹配及差异分析结果后，按新标准重新执行全流程并汇总。

**步骤**：
1. 删除 `OutPut/*/` 下所有 `SalesThreeMatchResult*.xlsx`（含主匹配、Untested、差异分析）
2. 执行 `run_all_by_company.py`（三单清单→三单匹配→差异分析）
3. 执行 `aggregate_difference_summary.py`

**命令**：
```bash
bash rerun_all_clean.sh
```

---

### 4.7 aggregate_difference_summary.py — 汇总差异分析

**用途**：在所有公司完成差异分析后，将各公司的汇总表合并至单一 Excel，并生成审计分析报告。

**输入**：`OutPut\{公司代码}\` 下的 `SalesThreeMatchResult_差异分析_*.xlsx`（读取「场景明细」sheet；若无则兼容旧版「四个大类」）

**输出**：`OutPut/汇总_差异分析_全公司.xlsx`，结构如下：

| Sheet | 内容 |
|-------|------|
| **全公司汇总** | 场景明细(1-10)、四个大类（由场景重算）、缺失发票/负开票 |
| **分析报告** | 汇总表一（全量）、汇总表二（剔除匹配率<80%公司）、剔除公司标注、分层汇总（≥80%、60–80%、<60%）、各公司明细、数据说明 |
| **{公司代码}** | 每公司一个 sheet：场景明细(1-10)、四个大类（由场景重算）、场景明细（差异分析原文） |

**运行逻辑**：
1. 扫描 OutPut 下公司子目录，查找差异分析文件（取最新）
2. 读取「场景明细」sheet；若无则兼容旧版「四个大类」
3. 基于各公司匹配结果重算场景、四大类
4. 按公司写入 sheet，每 sheet 含：场景明细、四大类（由场景重算）、缺失发票/负开票、场景明细（差异分析原文）
5. **无三单匹配结果的公司**（场景 1-13 记录数均为 0，如仅内部采购）：sheet 标题标注「（仅内部采购）」

**命令**：
```bash
python aggregate_difference_summary.py
```

---

### 4.8 summarize_three_match_crosstab.py — 三单匹配交叉汇总

**用途**：将三单匹配结果按「数量差异类型 × 金额差异类型」交叉汇总，输出 4×3 网格表。

**输入**：
- `OutPut/{公司}/` 下最新 `SalesThreeMatchResult_*.xlsx`（主匹配，排除 Untested、差异分析）
- `SalesThreeMatchResult_Untested_*.xlsx` 的 sheet「仅订单」「仅订单及发货单」（缺失发票）

**输出**：
- 单公司：`OutPut/{公司}/三单匹配交叉汇总_{公司}.xlsx`
- 全公司汇总：`OutPut/三单匹配交叉汇总_全公司.xlsx`（与 `--all --consolidated` 联用）

**分类维度**：
- **数量差异类型**：无差异、订单数量=发票≠交货单、订单数量≠发票≠交货单、缺失发票
- **金额差异类型**：订单金额<发票金额、订单金额>发票金额、无差异

**输出列**：数量差异类型、金额差异类型、差异笔数、发票金额、差异金额（差异金额 = Σ(订单-发票金额差异)，负值用括号表示）

**命令**：
```bash
python summarize_three_match_crosstab.py --company 4150   # 单公司
python summarize_three_match_crosstab.py --all            # 全公司分别输出
python summarize_three_match_crosstab.py --all --consolidated   # 全公司汇总至单 Excel
python summarize_three_match_crosstab.py --company 4150 --compact   # 仅输出有数据的行
```

**详细说明**：匹配方式、分类规则、结果解读见 [docs/交叉汇总_匹配方式及结果说明.md](docs/交叉汇总_匹配方式及结果说明.md)。

---

## 5. 配置 (config.py)

| 参数 | 类型 | 说明 |
|------|------|------|
| `DATA_FOLDER` | str | 单公司数据目录；可被环境变量 `SALES_DATA_FOLDER` 覆盖（多进程/自动化时用） |
| `INPUT_ROOT` | str | 自动化扫描根目录，run_all_by_company 从此下找公司子目录 |
| `OUTPUT_ROOT` | str | 输出根目录，各公司结果输出至 `OutPut\{公司代码}\` |
| `COMPANY_CODE` | str/None | 公司代码，留空则从 DATA_FOLDER 末段解析 |
| `ORDER_YEAR` / `ORDER_MONTH_START` / `ORDER_MONTH_END` | int | 订单、交货、发票的日期范围 |
| `EXCEL_MAX_ROWS_PER_FILE` | int | 单 Excel 最大行数，超出则分片导出（默认 980000） |
| `INVOICE_EXCLUDE_SOLD_TO_CODES` | set | 售达方排除固定名单；`get_exclude_sold_to_codes()` = 本名单 ∪ InPut 下四位数公司代码 |
| `MEMORY_SAVE_MODE` | bool | 大表优化：`True` 强制开启，`False` 按文件夹大小自动判断 |
| `LARGE_FOLDER_THRESHOLD_GB` | float | 数据文件夹（`InPut\{公司}`）超过此 GB 时自动开启大表优化（默认 10） |
| `CHUNK_SIZE_FOR_LARGE` | int | 大表分块行数（默认 150000） |
| `LARGE_FILE_MB` | int | 单文件超过此 MB 时强制分块+日期过滤（默认 500）；0 表示仅依赖文件夹阈值 |

### 5.1 大表内存优化（30GB+ 源数据、48GB 内存场景）

当单公司数据量较大（如 4390 等源数据 30GB+）时，可启用大表优化以降低内存占用：

1. **自动模式**（推荐）：`MEMORY_SAVE_MODE = False`  
   - 启动时检测当前 `DATA_FOLDER` 总大小  
   - 超过 `LARGE_FOLDER_THRESHOLD_GB`（默认 10GB）则自动开启，并打印 `[INFO] 数据文件夹 > {阈值}GB，已自动启用大表优化（分块+日期过滤）`

2. **强制模式**：`MEMORY_SAVE_MODE = True`  
   - 无论文件夹大小，始终启用分块读取与日期过滤

3. **优化措施**：  
   - LIPS/VBRP 分块读取，块内按 LFDAT/FKDAT 过滤后再合并  
   - 使用 `KEY_FIELDS_*` 仅保留匹配所需列  
   - 各表处理完毕后立即 `del` 释放内存  
   - pickle 缓存按公司代码+日期命名（如 `VBAK_001_4390.pkl`、`LIPS_001_4390_2025_1-12.pkl`），避免多公司混用；默认不清理便于重跑

### 5.2 售达方排除（订单与发票）

订单(KUNNR)、发票(KUNAG) 中售达方在排除名单内的记录会被剔除（视为内部采购场景，不纳入三单匹配）。

- **排除名单**：`get_exclude_sold_to_codes()` = `INVOICE_EXCLUDE_SOLD_TO_CODES` ∪ InPut 下所有四位数公司代码
- **应用范围**：three_lists、sales_three_match 的订单与发票处理环节
- **汇总表**：无三单匹配结果的公司，在 `aggregate_difference_summary` 的 sheet 标题会标注「（仅内部采购）」

### 5.3 发运单/发票数量单位与场景3 排查

**现象**：差异分析中「场景3（订单=发票≠交货、金额无差异）」占比约 50%，且抽样显示发运单数量为订单/发票数量的**四倍**，明显异常。

**原因**：SAP 中数量字段单位不一致，三单比较时混用导致误判。

| 来源 | 字段 | 单位 | 说明 |
|------|------|------|------|
| VBAP | KLMENG | **基本单位** (MEINS) | 订单数量（脚本已用） |
| LIPS | LFIMG | **销售单位** (VRKME) | 交货数量；当 1 基本单位=4 销售单位时 LFIMG≈4×LGMNG |
| LIPS | LGMNG | **基本单位** (MEINS) | 交货数量（基本单位） |
| VBRP | FKIMG | **销售单位** (VRKME) | 开票数量 |
| VBRP | FKLMG | **基本单位** (MEINS) | 开票数量（基本单位） |

**修正**（`sales_three_match.py` + `config.py` + `difference_analysis.py`）：

1. **取数**：`KEY_FIELDS_LIPS` 增加 `LGMNG`、`UMVKZ`、`UMVKN`；`KEY_FIELDS_VBRP` 增加 `FKLMG`、`UMVKZ`、`UMVKN`。
2. **交货数量**：优先使用 LGMNG（基本单位）；LGMNG 空/0 时用 LFIMG 按 UMVKZ/UMVKN 换算为基本单位。换算方向：SAP 约定 Sales = Base × (UMVKZ/UMVKN)，故 **Base = Sales × UMVKN/UMVKZ**。
3. **发票数量**：优先使用 FKLMG；FKLMG 空/0 时用 FKIMG × UMVKN/UMVKZ 换算为基本单位。
4. **兜底**（`difference_analysis.py`）：当 订单=发票(数量+金额) 且 发运单数量≈4×订单数量 时，视为单位换算遗留，修正发运单数量=订单数量并归入场景10；仅影响原场景3，不影响其余场景。

**验证**：修正后需删除该公司 LIPS/VBRP 的 pickle 缓存并重新执行三单匹配与差异分析，再查看场景3 占比是否下降。若 SAP 取数中 LGMNG/FKLMG 或 UMVKZ/UMVKN 未维护，兜底会将符合 4× 比例的行归入完全匹配；可参考 问题排查及解决记录.md 做进一步核对。

---

## 6. 匹配键与术语

| 术语 | 说明 |
|------|------|
| 匹配键 | 订单：VKORG + VBELN + POSNR；交货/发票通过 VBELV+POSNV 或 AUBEL+AUPOS 关联订单行 |
| BUKRS | 公司代码，拆分维度 |
| VKORG | 销售组织 |
| VBELN/POSNR | 订单号/行号 |
| VBELV/POSNV | 交货单号/行号；VBRP 空时用 AUBEL/AUPOS |

---

## 7. 依赖

```
pandas>=1.3.0, numpy>=1.20.0, openpyxl>=3.0.0, chardet>=4.0.0
```

---

## 8. 相关文档

| 文档 | 说明 |
|------|------|
| docs/交叉汇总_匹配方式及结果说明.md | 三单匹配交叉汇总的匹配逻辑、分类规则、输出含义 |
| docs/交叉汇总场景审阅.md | 交叉汇总场景定义与代码逻辑核对 |
| docs/场景判断条件_模块条件.md | 14 个场景标号的数量/金额判断条件及模块对应关系 |
| H1_增量追加执行说明.md | 2025 上半年数据增量追加到下半年、全年流程执行顺序 |
| 问题排查及解决记录.md | 发票/交货逻辑、取数核对、拆分审阅、4010 核对、Excel 行数超限修复等；场景3 发运单数量单位排查见 §5.3 |

---

## 9. 2025 全年数据合并进度

**完成时间**：2026-03-06

| 步骤 | 状态 | 说明 |
|------|------|------|
| H1 原始数据 | ✅ | 放入 `InPut_H1_RAW`（子目录含 TXT） |
| 展平 | ✅ | `flatten_h1_to_root.sh` 将子目录 TXT 符号链接到根目录 |
| 清理 | ✅ | `cleanup_h1_raw.sh` 删除空目录、已解压 ZIP（可选） |
| 备份 InPut | ✅ | 追加前自动备份至 `InPut_backup_before_H1_append_*` |
| 追加拆分 | ✅ | `run_h1_append_split.sh` → split_by_vkorg --append |
| 全年流程 | 待执行 | `run_full_year_pipeline.sh 1` |

**拆分结果**（2026-03-06 日志）：
- 43 个公司目录，VKORG→BUKRS 映射 43 条
- VBAK 4,118,476 行 / VBAP 25,549,678 行 / LIKP 8,805,745 行
- LIPS 102,629,885 行 / VBRK 3,875,092 行 / VBRP 45,865,984 行

**下一步**：执行 `bash run_full_year_pipeline.sh 1` 完成三单匹配与差异分析（处理所有公司）。

---

## 10. 最近更新

| 变更 | 说明 |
|------|------|
| 场景3 兜底逻辑 | 差异分析中，当 订单=发票(数量+金额) 且 发运单≈4×订单 时，修正发运单数量为订单数量并归入场景10；仅影响原场景3，不改变其余场景 |
| 发运单/发票数量单位统一 | 场景3 占比约 50%、发运单数量=4×订单 的排查：订单用 KLMENG(基本单位)，交货/发票原用 LFIMG/FKIMG(销售单位)。现优先 LGMNG/FKLMG(基本单位)，LGMNG/FKLMG 空时用 LFIMG/FKIMG×UMVKN/UMVKZ 换算；取数增加 LIPS.LGMNG/UMVKZ/UMVKN、VBRP.FKLMG/UMVKZ/UMVKN。详见 §5.3 |
| 场景优先输出 | 差异分析仅输出「场景明细」sheet，按场景 1→13→负开票 顺序；同时输出 `*_详细.xlsx` 明细清单，用户可自行将场景汇入四大类 |
| aggregate 读取场景明细 | 汇总脚本读取「场景明细」替代四个大类/细分；兼容旧版格式；各公司 sheet 展示场景明细（差异分析原文） |
| H1 增量追加 | split_by_vkorg `--append`、flatten_h1_to_root.sh、run_h1_append_split.sh、cleanup_h1_raw.sh |
| 售达方排除扩展 | 订单(KUNNR) 与发票(KUNAG) 均按售达方剔除；排除名单动态包含 InPut 下四位数公司代码 |
| 已有三单则跳过 | run_all_by_company 检测到公司已有三类清单时，跳过 three_lists，直接执行匹配与差异分析 |
| 多进程支持 | `--workers N` 支持多进程并行；通过 `SALES_DATA_FOLDER` 环境变量指定数据目录 |
| 空结果容错 | 差异分析对 0 行匹配结果输出空统计，不崩溃 |
| 未测试单据分片 | 单 sheet 超 Excel 行限时拆成多 sheet（如 仅发货单_1、_2） |
| 汇总表标注 | 无三单匹配结果的公司 sheet 标题加「（仅内部采购）」 |
| 负开票冲帐归入 Untested | 订单≈0、交货≈0或未匹配、发票<0 的行从主结果剔除，归入 Untested「负开票冲帐」sheet；Untested 前三个 sheet 重命名为：仅订单、仅订单及发货单、仅发货单 |
| `--match-diff-only` | run_all_by_company 新增参数，仅执行三单匹配+差异分析，跳过三单清单生成；仅处理已存在三单清单的公司 |
| rerun_all_clean.sh | 删除 OutPut 下所有匹配/差异分析文件后，按新标准全流程重跑并汇总 |
