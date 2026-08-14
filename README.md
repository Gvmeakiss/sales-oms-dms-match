# 销售三单匹配系统（OMS / DMS 双数据源） 🔍

> 对 ToB 销售订单、发货单与 SAP 开票三张单据按「订单-物料」键做三单匹配，自动计算金额/数量差异并分类，支持 OMS 与 DMS 两套数据源，用于对账与差异审计。

[![Language](https://img.shields.io/badge/language-Python-blue)](https://github.com/Gvmeakiss/sales-oms-dms-match) [![License](https://img.shields.io/badge/license-MIT-green)](https://github.com/Gvmeakiss/sales-oms-dms-match/blob/main/LICENSE) [![Domain](https://img.shields.io/badge/domain-Audit%20Analytics-orange)](https://github.com/Gvmeakiss/sales-oms-dms-match)

## 📌 项目简介
面向 Miaoke（妙可）ToB 销售业务，把**销售订单、发货单、SAP 开票**三套数据按 `order-item`（订单号 + 物料编码）对齐，逐行核对金额与数量是否一致。系统在 OMS（订单管理系统）与 DMS（经销商管理系统）两套数据源上分别实现全年三单匹配，输出汇总表与 5 类差异明细，辅助审计人员快速定位「金额不一致 / 数量不一致 / 数据缺失」等问题。

## ✨ 功能特性
- **双数据源匹配**：`match_oms_full_year.py` 基于 OMS 订单/发货/SAP 开票做三单匹配；`match_dms_full_year_from_sql.py` 直接解析 SQL 与发票 Excel 做 DMS 三单匹配。
- **三单差异计算**：逐行计算「订单-开票数量」「发货-开票数量」「订单-发票金额」（见 `match_oms_full_year.py` 第 5 节），DMS 则计算「SAP-DMS订单金额」「SAP-DMS发货数量」等（见 `match_dms_full_year_from_sql.py` 第 4 节）。
- **自动分类**：输出 5 个类别 —— `2.Not test`（关键字段缺失）、`1.1 完全匹配`、`1.2 金额不一致`、`1.3 数量不一致`、`1.4 均不一致`，阈值金额 1.0、数量 0.01（以代码为准）。
- **PKL 缓存加速**：首次解析后缓存到 `PKL/`（如 `oms_order_full_year.pkl`），后续直接加载，大幅提升重复运行速度（`preprocess_full_year.py` 产出，匹配脚本复用）。
- **按销售组织分组导出**：把三家销售组织（1240 / 1250 / 1260）与其余主体分开导出，各自含「汇总 + 全部数据 + 5 类 sheet」。
- **大文件分表**：当明细超过 1,048,575 行时自动拆解（`df_nottested` 取未匹配行；DMS 超限时先落 CSV 全量、再分 sheet 写 xlsx）。
- **数据纠错**：DMS 发货 7 列 SQL 含订单号/料号错位纠错逻辑；发票列名兼容（`OMS销售单号`/`OMS订单号`/`销售单号` 等多候选列自动匹配）。
- **辅助工具**：`launch_all.py` 串联 OMS/DMS 1–9 月匹配；`analyze_nottest.py` 分析未匹配项；`sap_invoice_processor.py` / `sql_to_accounting_excel.py` 处理 SAP 发票与生成会计 Excel。

## 📂 目录结构
```
sales-oms-dms-match/
├── README.md                          # 项目说明
├── LICENSE                            # MIT License
├── requirements.txt                   # pandas/numpy/openpyxl/chardet
├── preprocess_full_year.py           # OMS 全年 SQL→PKL 预处理
├── match_oms_full_year.py            # OMS 全年三单匹配（主入口）
├── match_oms_1_9_months.py           # OMS 1–9 月三单匹配
├── match_dms_full_year_from_sql.py   # DMS 全年三单匹配（从 SQL 读取）
├── match_dms_1_9_months_from_sql.py  # DMS 1–9 月三单匹配
├── launch_all.py                      # 串联 OMS/DMS 1–9 月匹配
├── analyze_nottest.py                # 未匹配项（Not test）分析
├── check_fullyear_pkl.py             # 全年 PKL 校验
├── check_全年pkl_合规.py             # PKL 合规性检查
├── sap_invoice_processor.py          # SAP 发票处理
└── sql_to_accounting_excel.py        # SQL 转会计 Excel
```

## 🔧 环境要求
- Python >= 3.8
- 依赖（来自 `requirements.txt`）：`pandas>=2.0`、`numpy>=1.24`、`openpyxl>=3.1`、`chardet>=5.0`

## 🚀 安装
```bash
git clone https://github.com/Gvmeakiss/sales-oms-dms-match.git
cd sales-oms-dms-match
pip install -r requirements.txt
```

## 💡 快速开始 / 使用示例
OMS 全年匹配（需先用 `preprocess_full_year.py` 将 SQL 解析为 PKL）：
```bash
# 1. 预处理：把 1–9 月 / 10–12 月 SQL 与 SAP 发票 XLSX 合并为 PKL
python preprocess_full_year.py
# 产出 2025年全年OMS订单.pkl / 2025年全年OMS发货.pkl / 2025年全年SAP原始数据.pkl

# 2. 匹配并导出
python match_oms_full_year.py
# 产出 2025年全年匹配结果-销售（toB OMS）明细（剔除三家）.xlsx 等

# DMS 全年匹配（直接读 SQL + 发票 Excel）
python match_dms_full_year_from_sql.py
# 产出 2025年全年匹配结果-销售（toB DMS）明细-从SQL.xlsx 等
```
也可一键串联 1–9 月匹配：`python launch_all.py`。

## 🧠 核心逻辑（方法论）
1. **建键与兼容**：`_norm_code()` 把物料/料号转字符串并去掉 `.0`；发票列名在 `['OMS销售单号','OMS订单号','销售单号']`、`['物料编码','料号','品号']` 间自动匹配（`match_oms_full_year.py`）。
2. **SQL 解析**：`_parse_values_line()` / `_parse_sql_values()` 解析 `VALUES (...)` 行并兼容 `NULL`、引号；`get_encoding()` 用 `chardet` 自动探测编码（`preprocess_full_year.py`、`match_dms_full_year_from_sql.py`）。
3. **聚合**：各单据按 `order-item`（OMS）或 `[DMS订单, 物料编码]`（DMS）`groupby.agg`，金额 `sum`、数量 `sum`、其余字段 `first`。
4. **匹配**：OMS 以发票为基准 `merge` 发货与订单；DMS 以 `[DMS订单, 物料编码]` 对齐。
5. **差异与分类**：计算金额/数量差额，按阈值（金额 1.0、数量 0.01）与关键字段缺失情况打标 `1.1`–`1.4` 与 `2.Not test`（见两匹配脚本第 4–5 节分类代码）。

## 📋 输入与输出
- **输入**：
  - OMS：由 `preprocess_full_year.py` 生成的 PKL（`2025年全年OMS订单.pkl`、`2025年全年OMS发货.pkl`、`2025年全年SAP原始数据.pkl`），原始来自 1–9 / 10–12 月 SQL 与 妙可 SAP 发票 XLSX。
  - DMS：1–9 与 10–12 月 SQL 文件 + `2025-01～12.XLSX` 发票 Excel。
- **输出**：每个匹配文件含「汇总表 + 全部数据 + `2.Not test` + `1.1 完全匹配` + `1.2 金额不一致` + `1.3 数量不一致` + `1.4 均不一致`」多个 sheet；DMS 超限时额外生成 `.csv` 全量与分 sheet `.xlsx`。

## ⚙️ 配置说明
- **PKL 缓存目录**：`PKL_DIR = Path('PKL')`，命中缓存则跳过 SQL 重新解析（提速关键）。
- **分类阈值**：金额阈值 `1.0`（四舍五入误差容忍）、数量阈值 `0.01`（小数精度容忍），均为代码内常量，可按需调整。
- **销售组织分组**：三家主体代码 `1240` / `1250` / `1260` 与其余按 `销售组织` 字段拆分组导出。

## ⚠️ 注意事项
- 数据脱敏：仓库不含真实客户业务数据，示例与文件名中的客户名为脱敏化名（Miaoke / 妙可），实际运行需用户提供自有数据。
- 口径说明：匹配口径、分类阈值与字段筛选（如订单状态剔除 `OBSOLETE`/`CANCEL`、订单类型 `liquid_milk_order`/`factory_item`/`oem_order`）以代码与配置为准。
- 运行前需自行准备对应年份的 SQL 与 SAP 发票文件，并先执行 `preprocess_full_year.py` 生成 OMS PKL。

## 🔗 相关仓库
- https://github.com/Gvmeakiss/sales-three-match-toolkit
- https://github.com/Gvmeakiss/sales-three-match-newhope

## 📄 License
MIT（Copyright © 2026 Gvmeakiss (James Li)）。

---

<div align="center">

*Disclaimer: Personal project and personal views. Not affiliated with or endorsed by KPMG or any client.*<br>
*本仓库为个人项目与个人观点，与任何前/现雇主及客户无关。*

</div>
