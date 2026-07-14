# -*- coding: utf-8 -*-
"""
OMS 全年销售三单匹配（订单 + 发货 + 发票）

依赖 preprocess_full_year.py 产出的 pkl：
  - 2025年全年OMS订单.pkl
  - 2025年全年OMS发货.pkl
  - 2025年全年SAP原始数据.pkl

功能：
  - PKL缓存机制：首次运行后保存到 PKL/ 目录，后续直接加载
  - 增加输出字段：订单、发货、发票的详细字段（包括销售组织、发票类型等）
  - 按销售组织分组导出：剔除三家（1240、1250、1260）vs 三家公司单独导出
  - 按分类拆分：每个文件包含汇总表 + 全部数据 + 5个分类sheet

输出：
  - 剔除三家：
    - 2025年全年匹配结果-销售（toB OMS）明细（剔除三家）.xlsx
    - 2025年全年匹配结果-销售（toB OMS）明细-其他未匹配（剔除三家）.xlsx
  - 三家公司（1240、1250、1260）：
    - 2025年全年匹配结果-销售（toB OMS）明细（1240、1250、1260）.xlsx
    - 2025年全年匹配结果-销售（toB OMS）明细-其他未匹配（1240、1250、1260）.xlsx
  - 每个文件包含：汇总表 + 全部数据 + 2.Not test + 1.1完全匹配 + 1.2金额不一致 + 1.3数量不一致 + 1.4均不一致
"""

import os
import pandas as pd
import warnings
from pathlib import Path
import pickle

warnings.filterwarnings('ignore')

# PKL缓存目录
PKL_DIR = Path('PKL')
PKL_DIR.mkdir(exist_ok=True)

def _norm_code(ser):
    """物料/料号/item 转字符串并去掉因 float 产生的 '.0'。"""
    return ser.astype(str).str.replace(r'\.0$', '', regex=True)

# ============================================================================
# 1. 数据导入
# ============================================================================

print("开始加载 2025 年全年数据...")

# 检查PKL缓存
pkl_order = PKL_DIR / 'oms_order_full_year.pkl'
pkl_delivery = PKL_DIR / 'oms_delivery_full_year.pkl'
pkl_invoice = PKL_DIR / 'oms_invoice_full_year.pkl'

if pkl_order.exists() and pkl_delivery.exists() and pkl_invoice.exists():
    print('  从PKL缓存加载数据...')
    df_order = pd.read_pickle(pkl_order)
    df_delivery = pd.read_pickle(pkl_delivery)
    df_invoice = pd.read_pickle(pkl_invoice)
    print(f"  订单: {len(df_order):,} 行, 列: {list(df_order.columns)}")
    print(f"  发货: {len(df_delivery):,} 行, 列: {list(df_delivery.columns)}")
    print(f"  发票: {len(df_invoice):,} 行, 列(前10): {list(df_invoice.columns[:10])}")
else:
    print('  从PKL文件读取原始数据...')
    order_pkl = '2025年全年OMS订单.pkl'
    delivery_pkl = '2025年全年OMS发货.pkl'
    invoice_pkl = '2025年全年SAP原始数据.pkl'

    for n, p in [('订单', order_pkl), ('发货', delivery_pkl), ('发票', invoice_pkl)]:
        if not os.path.exists(p):
            raise FileNotFoundError(f"未找到{n} pkl: {p}\n请先运行 preprocess_full_year.py")

    df_order = pd.read_pickle(order_pkl)
    df_delivery = pd.read_pickle(delivery_pkl)
    df_invoice = pd.read_pickle(invoice_pkl)
    print(f"  订单: {len(df_order):,} 行, 列: {list(df_order.columns)}")
    print(f"  发货: {len(df_delivery):,} 行, 列: {list(df_delivery.columns)}")
    print(f"  发票: {len(df_invoice):,} 行, 列(前10): {list(df_invoice.columns[:10])}")
    
    # 保存到PKL缓存
    print('\n  保存数据到PKL缓存...')
    df_order.to_pickle(pkl_order)
    df_delivery.to_pickle(pkl_delivery)
    df_invoice.to_pickle(pkl_invoice)
    print(f"  已保存: {pkl_order}")
    print(f"  已保存: {pkl_delivery}")
    print(f"  已保存: {pkl_invoice}")

# 发票列名兼容（1-9 与 10-12 列名可能略有差异）
oms_col = next((c for c in ['OMS销售单号', 'OMS订单号', '销售单号'] if c in df_invoice.columns), 'OMS销售单号')
mat_col = next((c for c in ['物料编码', '料号', '品号', '物料编号'] if c in df_invoice.columns), '物料编码')
if oms_col != 'OMS销售单号' or mat_col != '物料编码':
    print(f"  发票列名兼容: OMS→{oms_col}, 物料→{mat_col}")

# ============================================================================
# 2. 数据预处理
# ============================================================================

print("\n数据预处理...")

# 订单：状态、类型
df_order = df_order[df_order['order_status'] != 'OBSOLETE']
df_order = df_order[df_order['order_status'] != 'CANCEL']
df_order = df_order[df_order['order_type'].isin(['liquid_milk_order', 'factory_item', 'oem_order'])]
if 'create_time' in df_order.columns:
    df_order['create_time'] = pd.to_datetime(df_order['create_time'], errors='coerce')
    df_order = df_order[df_order['create_time'].dt.year == 2025]
print(f"  订单过滤后: {len(df_order):,} 行")

# 发货：2025 年 1–12 月
if '业务时间' in df_delivery.columns:
    df_delivery['业务时间'] = pd.to_datetime(df_delivery['业务时间'], format='mixed', errors='coerce')
    df_delivery = df_delivery[(df_delivery['业务时间'].dt.year == 2025) & (df_delivery['业务时间'].dt.month >= 1) & (df_delivery['业务时间'].dt.month <= 12)]
print(f"  发货 2025 年 1–12 月: {len(df_delivery):,} 行")

# 发票：类型过滤（To B 2B / 标准退货等）
invoice_type_col = None
for col in df_invoice.columns:
    if '发票类型' in str(col):
        invoice_type_col = col
        break
if not invoice_type_col and len(df_invoice.columns) > 14:
    invoice_type_col = df_invoice.columns[14]

if invoice_type_col:
    ser = df_invoice[invoice_type_col].astype(str)
    is_code = ser.str.fullmatch(r'[A-Z]{2}\d{2}', na=False).mean() > 0.5
    if is_code:
        keep = ser.isin({'ZA01','ZB02','ZB06'})
    else:
        keep = (ser.str.contains(r'标准发票\s*[（(]?\s*2B', regex=True) |
                ser.str.contains('标准退货发票', regex=False) |
                ser.str.contains(r'取消标准发票\s*[（(]?\s*2B', regex=True) |
                ser.str.contains('取消标准退货发票', regex=False))
    df_invoice = df_invoice[keep]
else:
    print("  警告: 未找到发票类型列，未做类型过滤")
print(f"  发票 ToB 过滤后: {len(df_invoice):,} 行")

# 发票 → order-item：OMS 直接 + DMS 映射
inv_with_oms = pd.DataFrame()
if oms_col in df_invoice.columns and mat_col in df_invoice.columns:
    inv_with_oms = df_invoice[df_invoice[oms_col].notna()].copy()
    inv_with_oms['order-item'] = inv_with_oms[oms_col].astype(str) + _norm_code(inv_with_oms[mat_col])

inv_need_map = pd.DataFrame()
if 'DMS销售单号' in df_invoice.columns and mat_col in df_invoice.columns:
    inv_need_map = df_invoice[df_invoice[oms_col].isna() & df_invoice['DMS销售单号'].notna()].copy()

mapped_df = pd.DataFrame()
if not inv_need_map.empty and {'external_order_no','item_code','sale_order_no'}.issubset(df_order.columns):
    order_map = df_order[['external_order_no','item_code','sale_order_no']].dropna(subset=['external_order_no'])
    order_map = order_map.copy()
    order_map['dms-oi'] = order_map['external_order_no'].astype(str) + _norm_code(order_map['item_code'])
    inv_need_map = inv_need_map.copy()
    inv_need_map['dms-oi'] = inv_need_map['DMS销售单号'].astype(str) + _norm_code(inv_need_map[mat_col])
    m = inv_need_map.merge(order_map[['dms-oi','sale_order_no']], on='dms-oi', how='left')
    m = m[m['sale_order_no'].notna()].copy()
    if not m.empty:
        m['order-item'] = m['sale_order_no'].astype(str) + _norm_code(m[mat_col])
        mapped_df = m

df_invoice_used = pd.concat([inv_with_oms, mapped_df], ignore_index=True)
print(f"  用于聚合的发票条数: {len(df_invoice_used):,}")

# 仅 OMS（DMS 销售单号为空）
if 'DMS销售单号' in df_invoice.columns:
    df_invoice = df_invoice[df_invoice['DMS销售单号'].isna()]

# ============================================================================
# 3. 匹配键 order-item
# ============================================================================

print("\n创建匹配键 order-item...")

df_order['order-item'] = df_order['sale_order_no'].astype(str) + _norm_code(df_order['item_code'])

if '主单号' in df_delivery.columns and '订单号' in df_delivery.columns and '料号' in df_delivery.columns:
    o = df_delivery.apply(lambda r: str(r['主单号']) if pd.notna(r['主单号']) and str(r['主单号']).startswith('DD') else str(r['订单号']), axis=1)
    df_delivery['order-item'] = o.astype(str) + _norm_code(df_delivery['料号'])
elif '订单号' in df_delivery.columns and '料号' in df_delivery.columns:
    df_delivery['order-item'] = df_delivery['订单号'].astype(str) + _norm_code(df_delivery['料号'])
else:
    raise ValueError("发货缺少 订单号 或 料号")

if oms_col in df_invoice.columns and mat_col in df_invoice.columns:
    df_invoice['order-item'] = df_invoice[oms_col].astype(str) + _norm_code(df_invoice[mat_col])

# ============================================================================
# 4. 聚合（保留更多字段）
# ============================================================================

df_order['pay_amount'] = pd.to_numeric(df_order['pay_amount'], errors='coerce').round(2)
df_order['item_num'] = pd.to_numeric(df_order['item_num'], errors='coerce').round(2)

# 订单聚合：金额和数量用sum，其他字段用first保留
# 注意：groupby key是'order-item'（计算列），不是原始列，所以可以安全地添加sale_order_no和item_code到agg_dict
order_agg_dict = {'pay_amount': 'sum', 'item_num': 'sum'}
order_first_cols = ['sale_order_no', 'platform_order_no', 'main_order_no', 'channel_name', 
                    'order_type', 'order_status', 'create_time', 'update_time', 'item_code']
for col in order_first_cols:
    if col in df_order.columns:
        order_agg_dict[col] = 'first'

pivot_order = df_order.groupby('order-item', as_index=False).agg(order_agg_dict)
pivot_order.rename(columns={
    'pay_amount': '订单金额',
    'item_num': '订单数量',
    'sale_order_no': '订单-销售订单号',
    'platform_order_no': '订单-平台订单号',
    'main_order_no': '订单-主订单号',
    'channel_name': '订单-渠道名称',
    'order_type': '订单-订单类型',
    'order_status': '订单-订单状态',
    'create_time': '订单-创建时间',
    'update_time': '订单-更新时间',
    'item_code': '订单-商品代码'
}, inplace=True)

df_delivery['已发货数量'] = pd.to_numeric(df_delivery['已发货数量'], errors='coerce').round(2)

# 发货聚合：数量用sum，其他字段用first保留
# 注意：groupby key是'order-item'（计算列），不是原始列，所以可以安全地添加订单号和料号到agg_dict
delivery_agg_dict = {'已发货数量': 'sum'}
delivery_first_cols = ['订单号', '主单号', 'external_order_no', '业务时间', '料号', '名称', 'business_type']
for col in delivery_first_cols:
    if col in df_delivery.columns:
        delivery_agg_dict[col] = 'first'

pivot_delivery = df_delivery.groupby('order-item', as_index=False).agg(delivery_agg_dict)
pivot_delivery.rename(columns={
    '已发货数量': '发货数量',
    '订单号': '发货-订单号',
    '主单号': '发货-主单号',
    'external_order_no': '发货-外部订单号',
    '业务时间': '发货-业务时间',
    '料号': '发货-料号',
    '名称': '发货-商品名称',
    'business_type': '发货-业务类型'
}, inplace=True)

amount_col = quantity_col = None
for c in df_invoice_used.columns:
    if '实际金额' in str(c) or ('金额' in str(c) and 'ZFN1' in str(c)):
        amount_col = c
    if '开票数量' in str(c) or ('数量' in str(c) and '基本单位' in str(c)):
        quantity_col = c

if df_invoice_used.empty or not amount_col or not quantity_col:
    pivot_invoice = pd.DataFrame({'order-item':[], '开票金额':[], '开票数量':[]})
else:
    df_invoice_used[amount_col] = pd.to_numeric(df_invoice_used[amount_col], errors='coerce').round(2)
    df_invoice_used[quantity_col] = pd.to_numeric(df_invoice_used[quantity_col], errors='coerce').round(2)
    
    # 发票聚合：金额和数量用sum，其他字段用first保留
    invoice_agg_dict = {amount_col: 'sum', quantity_col: 'sum'}
    
    # 查找发票相关字段（按优先级匹配）
    invoice_extra_cols = []
    invoice_field_candidates = [
        (['SAP发票号', 'SAP发票编号', 'SAP账单号', '发票号', '发票编号'], '发票-SAP发票号'),
        (['公司代码', '公司'], '发票-公司代码'),
        (['发票备注', '备注', '发票说明', '说明'], '发票-发票备注'),
        (['发票类型', '发票类型.1'], '发票-发票类型'),
        (['销售组织', '销售组织代码'], '发票-销售组织'),
        (['开票日期', '发票日期'], '发票-开票日期'),
        (['客户名称', '客户'], '发票-客户名称'),
        (['物料名称', '名称'], '发票-物料名称'),
        (['数据源文件'], '发票-数据源文件'),
        (['DMS销售单号'], '发票-DMS销售单号')
    ]
    
    for candidates, label in invoice_field_candidates:
        for col in df_invoice_used.columns:
            col_str = str(col)
            if any(cand in col_str for cand in candidates) and col not in invoice_agg_dict:
                invoice_agg_dict[col] = 'first'
                invoice_extra_cols.append((col, label))
                break
    
    # 保留OMS销售单号和物料编码用于匹配
    # 注意：groupby key是'order-item'（计算列 = oms_col + mat_col），不是原始列组合
    # 所以可以安全地添加oms_col和mat_col到agg_dict，不会导致重复错误
    if oms_col in df_invoice_used.columns and oms_col not in invoice_agg_dict:
        invoice_agg_dict[oms_col] = 'first'
    if mat_col in df_invoice_used.columns and mat_col not in invoice_agg_dict:
        invoice_agg_dict[mat_col] = 'first'
    
    pivot_invoice = df_invoice_used.groupby('order-item', as_index=False).agg(invoice_agg_dict)
    pivot_invoice.rename(columns={amount_col: '开票金额', quantity_col: '开票数量'}, inplace=True)
    
    # 重命名额外字段
    for old_col, new_label in invoice_extra_cols:
        if old_col in pivot_invoice.columns:
            pivot_invoice.rename(columns={old_col: new_label}, inplace=True)
    
for c in ['开票金额','开票数量']:
    if c not in pivot_invoice.columns:
        pivot_invoice[c] = pd.Series(dtype=float)

# ============================================================================
# 5. 匹配、差异、分类、导出
# ============================================================================

df_join = pivot_invoice.merge(pivot_delivery, on='order-item', how='left')
df_join = df_join.merge(pivot_order, on='order-item', how='left')

df_nottested = pivot_order.merge(pivot_delivery, on='order-item', how='outer')
df_nottested = df_nottested.merge(pivot_invoice, on='order-item', how='outer')
df_nottested = df_nottested[df_nottested['开票数量'].isna()]

df_join['订单-发货数量'] = pd.to_numeric(df_join['订单数量'], errors='coerce') - pd.to_numeric(df_join['发货数量'], errors='coerce')
df_join['订单-开票数量'] = pd.to_numeric(df_join['订单数量'], errors='coerce') - pd.to_numeric(df_join['开票数量'], errors='coerce')
df_join['发货-开票数量'] = pd.to_numeric(df_join['发货数量'], errors='coerce') - pd.to_numeric(df_join['开票数量'], errors='coerce')
df_join['订单-发票金额'] = (pd.to_numeric(df_join['订单金额'], errors='coerce') - pd.to_numeric(df_join['开票金额'], errors='coerce')).round(2)

# 只检查关键字段是否有缺失，而不是所有字段（避免额外字段的NaN影响判断）
key_cols = ['订单金额', '订单数量', '发货数量', '开票金额', '开票数量']
df_join['2.Not test'] = df_join[key_cols].isna().any(axis=1)
df_join['1.1完全匹配'] = ((df_join['订单-开票数量'] == 0) & (df_join['发货-开票数量'] == 0) & (df_join['订单-发票金额'].abs() < 1) & ~df_join['2.Not test'])
df_join['1.2金额不一致'] = ((df_join['订单-发票金额'].abs() >= 1) & (df_join['订单-开票数量'] == 0) & (df_join['发货-开票数量'] == 0) & ~df_join['2.Not test'])
df_join['1.3数量不一致'] = (((df_join['订单-开票数量'] != 0) | (df_join['发货-开票数量'] != 0)) & (df_join['订单-发票金额'].abs() < 1) & ~df_join['2.Not test'])
df_join['1.4均不一致'] = (((df_join['订单-开票数量'] != 0) | (df_join['发货-开票数量'] != 0)) & (df_join['订单-发票金额'].abs() >= 1) & ~df_join['2.Not test'])

print("\n分类统计:")
print(f"  1.1 完全匹配: {df_join['1.1完全匹配'].sum()}")
print(f"  1.2 金额不一致: {df_join['1.2金额不一致'].sum()}")
print(f"  1.3 数量不一致: {df_join['1.3数量不一致'].sum()}")
print(f"  1.4 均不一致: {df_join['1.4均不一致'].sum()}")
print(f"  2. 未测试(有缺失): {df_join['2.Not test'].sum()}")

# ============================================================================
# 6. 按销售组织分组（1240、1250、1260 与 其他）
# ============================================================================
print('\n6. 按销售组织分组...')

# 查找销售组织字段
sales_org_col = None
for col in df_join.columns:
    if '销售组织' in str(col):
        sales_org_col = col
        break

if sales_org_col:
    # 销售组织值（支持字符串和数字格式）
    sales_org_values = ['1240', '1250', '1260', 1240, 1250, 1260]
    sales_org_str_values = [str(v) for v in sales_org_values]
    
    # 筛选出1240、1250、1260的数据
    df_sales_org_123 = df_join[df_join[sales_org_col].astype(str).isin(sales_org_str_values)].copy()
    # 剔除1240、1250、1260的数据（保留其他公司）
    df_sales_org_other = df_join[~df_join[sales_org_col].astype(str).isin(sales_org_str_values)].copy()
    
    print(f"  销售组织（1240、1250、1260）数据: {len(df_sales_org_123):,} 条")
    print(f"  其他销售组织数据: {len(df_sales_org_other):,} 条")
else:
    print(f"  警告: 未找到销售组织字段，使用全部数据")
    df_sales_org_123 = pd.DataFrame()
    df_sales_org_other = df_join.copy()

# ============================================================================
# 7. 导出（按分类拆分到独立sheet + 汇总表）
# ============================================================================
print('\n7. 导出（按分类拆分）...')

EXCEL_MAX = 1_048_575

def export_with_classification(df_data, output_file, file_label=""):
    """导出数据，按分类拆分到独立sheet，并创建汇总表"""
    if len(df_data) == 0:
        print(f"  {file_label}: 无数据，跳过导出")
        return
    
    # 查找开票金额字段
    amount_col = None
    for col in df_data.columns:
        if '开票金额' in str(col):
            amount_col = col
            break
    
    # 准备分类数据
    categories = {
        '全部数据': df_data,
        '2.Not test': df_data[df_data['2.Not test']].copy() if '2.Not test' in df_data.columns else pd.DataFrame(),
        '1.1完全匹配': df_data[df_data['1.1完全匹配']].copy() if '1.1完全匹配' in df_data.columns else pd.DataFrame(),
        '1.2金额不一致': df_data[df_data['1.2金额不一致']].copy() if '1.2金额不一致' in df_data.columns else pd.DataFrame(),
        '1.3数量不一致': df_data[df_data['1.3数量不一致']].copy() if '1.3数量不一致' in df_data.columns else pd.DataFrame(),
        '1.4均不一致': df_data[df_data['1.4均不一致']].copy() if '1.4均不一致' in df_data.columns else pd.DataFrame(),
    }
    
    # 创建汇总表
    summary_data = []
    for cat_name, cat_df in categories.items():
        if len(cat_df) > 0:
            row_count = len(cat_df)
            if amount_col and amount_col in cat_df.columns:
                total_amount = pd.to_numeric(cat_df[amount_col], errors='coerce').sum()
                summary_data.append({
                    '分类': cat_name,
                    '记录数': row_count,
                    '开票金额': round(total_amount, 2) if not pd.isna(total_amount) else 0
                })
            else:
                summary_data.append({
                    '分类': cat_name,
                    '记录数': row_count,
                    '开票金额': 0
                })
    
    df_summary = pd.DataFrame(summary_data)
    
    # 添加总计行
    if len(df_summary) > 0:
        total_row = {
            '分类': '总计',
            '记录数': df_summary['记录数'].sum(),
            '开票金额': df_summary['开票金额'].sum()
        }
        df_summary = pd.concat([df_summary, pd.DataFrame([total_row])], ignore_index=True)
    
    # 导出到Excel（超过行数限制时自动分多个sheet）
    with pd.ExcelWriter(output_file, engine='openpyxl') as w:
        # 汇总表（第一个sheet）
        df_summary.to_excel(w, sheet_name='汇总表', index=False)
        
        # 各分类数据
        for cat_name, cat_df in categories.items():
            if len(cat_df) == 0:
                continue
            
            # 调整列顺序：将"发票-DMS销售单号"移到最后一列
            if '发票-DMS销售单号' in cat_df.columns:
                cols = [c for c in cat_df.columns if c != '发票-DMS销售单号']
                cols.append('发票-DMS销售单号')
                cat_df = cat_df[cols]
            
            # Excel sheet名称限制31个字符，需要截断
            sheet_name = cat_name[:31]
            
            if len(cat_df) <= EXCEL_MAX:
                cat_df.to_excel(w, sheet_name=sheet_name, index=False, na_rep='N/A')
            else:
                # 如果单个分类也超限，分多个sheet
                part_count = (len(cat_df) + EXCEL_MAX - 1) // EXCEL_MAX
                for i, start in enumerate(range(0, len(cat_df), EXCEL_MAX)):
                    part_sheet_name = f"{sheet_name[:25]}_P{i+1}"[:31]
                    cat_df.iloc[start:start + EXCEL_MAX].to_excel(w, sheet_name=part_sheet_name, index=False, na_rep='N/A')
                print(f"    {cat_name}: 数据量 {len(cat_df):,} 行，已分割为 {part_count} 个sheet")
        
        sheet_count = len([c for c in categories.values() if len(c) > 0])
        # 统计实际sheet数量（包括分割的sheet）
        total_sheets = 1  # 汇总表
        for cat_name, cat_df in categories.items():
            if len(cat_df) > 0:
                if len(cat_df) > EXCEL_MAX:
                    total_sheets += (len(cat_df) + EXCEL_MAX - 1) // EXCEL_MAX
                else:
                    total_sheets += 1
        print(f"  {file_label}: 已保存 {output_file} (共 {total_sheets} 个sheet，包含汇总表和 {sheet_count} 个分类)")

# 7.1 导出剔除三家公司的主明细和未匹配文件
if sales_org_col:
    out1 = '2025年全年匹配结果-销售（toB OMS）明细（剔除三家）.xlsx'
    out2 = '2025年全年匹配结果-销售（toB OMS）明细-其他未匹配（剔除三家）.xlsx'
    df_export = df_sales_org_other
else:
    out1 = '2025年全年匹配结果-销售（toB OMS）明细.xlsx'
    out2 = '2025年全年匹配结果-销售（toB OMS）明细-其他未匹配.xlsx'
    df_export = df_join

export_with_classification(df_export, out1, "剔除三家-主明细")

df_not_tested_other = df_export[df_export['2.Not test']].copy()
export_with_classification(df_not_tested_other, out2, "剔除三家-未匹配")

# 7.2 导出三家公司的明细和未匹配文件
if sales_org_col and len(df_sales_org_123) > 0:
    print('\n8. 导出三家公司的数据（1240、1250、1260）...')
    
    out3 = '2025年全年匹配结果-销售（toB OMS）明细（1240、1250、1260）.xlsx'
    out4 = '2025年全年匹配结果-销售（toB OMS）明细-其他未匹配（1240、1250、1260）.xlsx'
    
    export_with_classification(df_sales_org_123, out3, "三家公司-主明细")
    
    df_not_tested_123 = df_sales_org_123[df_sales_org_123['2.Not test']].copy()
    export_with_classification(df_not_tested_123, out4, "三家公司-未匹配")
elif sales_org_col:
    print('\n8. 警告: 未找到销售组织为1240、1250、1260的数据，跳过三家公司的导出')

print('\n全年 OMS 三单匹配完成。')
