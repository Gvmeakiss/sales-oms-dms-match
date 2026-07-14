#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2025年全年订单、发货、SAP 发票数据预处理脚本

- 订单：OMS25年1-12月订单及发货数据 下的 1-9 与 10-12 月 SQL
  - 1-9 月 11 列：补 line_amount=NaN、channel_name2=NaN 统一为 13 列
  - 10-12 月 13 列：直接解析
  - 合并 → 2025年全年OMS订单.pkl

- 发货：同目录下 1-9 与 10-12 月 SQL，均为 7 列（10-12 含 7 列纠错逻辑），合并
  - → 2025年全年OMS发货.pkl

- 发票：妙可 SAP发票拆分月份 下 2025-01～12.XLSX，concat(join='outer') 做列对齐
  - → 2025年全年SAP原始数据.pkl
"""

import pandas as pd
import numpy as np
import os
from pathlib import Path
import chardet

def get_encoding(file_path):
    with open(file_path, 'rb') as f:
        return chardet.detect(f.read())['encoding']

def _parse_values_line(line):
    """从 VALUES 行解析出值列表，处理引号与 NULL。"""
    line = line.strip()
    if not line.upper().startswith('VALUES'):
        return None
    part = line[6:].strip().strip(';')
    start = part.find('(')
    end = part.rfind(')')
    if start < 0 or end <= start:
        return None
    part = part[start + 1 : end]
    if not part.strip():
        return None
    parts = []
    cur = ''
    in_q = False
    for c in part:
        if c == "'":
            in_q = not in_q
            cur += c
        elif c == ',' and not in_q:
            if cur.strip():
                parts.append(cur.strip())
            cur = ''
        else:
            cur += c
    if cur.strip():
        parts.append(cur.strip())
    out = []
    for p in parts:
        p = p.strip()
        if p.upper() == 'NULL':
            out.append(None)
        else:
            if p.startswith("'") and p.endswith("'"):
                p = p[1:-1]
            out.append(p)
    return out

def _parse_sql_values(sql_path, expected_cols, min_cols=None):
    """只解析 VALUES 行，返回 DataFrame。"""
    enc = get_encoding(sql_path)
    with open(sql_path, 'r', encoding=enc, errors='ignore') as f:
        lines = f.readlines()
    rows = []
    for line in lines:
        vals = _parse_values_line(line)
        if vals is None:
            continue
        need = min_cols if min_cols is not None else len(expected_cols)
        if len(vals) < need:
            continue
        rows.append(vals[: len(expected_cols)])
    df = pd.DataFrame(rows)
    if len(df.columns) != len(expected_cols):
        df = df.iloc[:, : len(expected_cols)]
    df.columns = expected_cols
    return df, enc

# ---------------------------------------------------------------------------
# 订单：11 列 / 13 列分支，统一为 13 列后合并
# ---------------------------------------------------------------------------

COLS_11 = [
    'platform_order_no', 'main_order_no', 'sale_order_no', 'order_type', 'order_status',
    'create_time', 'update_time', 'channel_name', 'item_code', 'pay_amount', 'item_num'
]
COLS_13 = [
    'platform_order_no', 'main_order_no', 'sale_order_no', 'order_type', 'order_status',
    'create_time', 'update_time', 'channel_name', 'item_code', 'line_amount', 'pay_amount', 'item_num', 'channel_name2'
]

def process_order_sql_11cols(sql_path):
    """解析 11 列订单 SQL，补 line_amount、channel_name2 为 13 列。"""
    df, enc = _parse_sql_values(sql_path, COLS_11, min_cols=11)
    df.insert(9, 'line_amount', np.nan)
    df['channel_name2'] = np.nan
    df = df[COLS_13]
    return df, enc

def process_order_sql_13cols(sql_path):
    """解析 13 列订单 SQL。"""
    return _parse_sql_values(sql_path, COLS_13, min_cols=13)

def process_orders_full_year(data_dir, output_pkl):
    """合并 1-9（11 列）与 10-12（13 列）订单为全年 pkl。"""
    data_dir = Path(data_dir)
    out = Path(output_pkl)

    # 1-9 月：11 列
    sql_19 = data_dir / "25年1-9月订单数据.sql"
    if not sql_19.exists():
        raise FileNotFoundError(f"订单 SQL 不存在: {sql_19}")
    df_19, enc_19 = process_order_sql_11cols(str(sql_19))
    print(f"  1-9 月订单: 编码={enc_19}, 行数={len(df_19):,}, 已补 line_amount/channel_name2 为 13 列")

    # 10-12 月：13 列
    sql_1012 = data_dir / "25年10-12月订单数据.sql"
    if not sql_1012.exists():
        raise FileNotFoundError(f"订单 SQL 不存在: {sql_1012}")
    df_1012, enc_1012 = process_order_sql_13cols(str(sql_1012))
    print(f"  10-12 月订单: 编码={enc_1012}, 行数={len(df_1012):,}")

    combined = pd.concat([df_19, df_1012], ignore_index=True)
    combined.to_pickle(out)
    print(f"  已保存: {out} (总行数: {len(combined):,})")
    return combined

# ---------------------------------------------------------------------------
# 发货：7 列（或 8 列），对 7 列做错位纠错，合并 1-9 与 10-12
# ---------------------------------------------------------------------------

COLS_7 = ['business_type', '订单号', 'external_order_no', '业务时间', '料号', '名称', '已发货数量']
COLS_8 = ['business_type', '主单号', '订单号', 'external_order_no', '业务时间', '料号', '名称', '已发货数量']

def _apply_delivery_7col_fix(df):
    """7 列时：若第 2 列多为数字、第 5 列多为 DD 开头，则对调第 2 与 第 5 列。"""
    if df.shape[1] < 5:
        return
    s1 = df.iloc[:, 1].astype(str)
    s4 = df.iloc[:, 4].astype(str)
    n = len(df)
    dd1 = s1.str.match(r'^DD', na=False).sum()
    dd4 = s4.str.match(r'^DD', na=False).sum()
    dg1 = s1.str.match(r'^\d+$', na=False).sum()
    dg4 = s4.str.match(r'^\d+$', na=False).sum()
    if dg1 > n * 0.5 and dd4 > n * 0.5 and (dd1 <= n * 0.5 or dg4 <= n * 0.5):
        df.iloc[:, [1, 4]] = df.iloc[:, [4, 1]].values

def process_delivery_sql_one(sql_path):
    """处理单个发货 SQL，返回 DataFrame。"""
    enc = get_encoding(sql_path)
    with open(sql_path, 'r', encoding=enc, errors='ignore') as f:
        lines = f.readlines()
    rows = []
    for line in lines:
        vals = _parse_values_line(line)
        if vals is None:
            continue
        if len(vals) >= 8:
            rows.append(vals[:8])
        elif len(vals) >= 7:
            rows.append(vals[:7])
    if not rows:
        return pd.DataFrame(), enc
    n = len(rows[0])
    cols = COLS_8 if n == 8 else COLS_7
    df = pd.DataFrame(rows).iloc[:, : len(cols)]
    if len(cols) == 7:
        _apply_delivery_7col_fix(df)
    df.columns = cols
    return df, enc

def process_delivery_full_year(data_dir, output_pkl):
    """合并 1-9 与 10-12 发货为全年 pkl。"""
    data_dir = Path(data_dir)
    out = Path(output_pkl)

    d19 = data_dir / "25年1-9月发货数据.sql"
    d1012 = data_dir / "25年10-12月发货数据.sql"
    if not d19.exists():
        raise FileNotFoundError(f"发货 SQL 不存在: {d19}")
    if not d1012.exists():
        raise FileNotFoundError(f"发货 SQL 不存在: {d1012}")

    df_19, e19 = process_delivery_sql_one(str(d19))
    print(f"  1-9 月发货: 编码={e19}, 行数={len(df_19):,}, 列数={len(df_19.columns)}")
    df_1012, e1012 = process_delivery_sql_one(str(d1012))
    print(f"  10-12 月发货: 编码={e1012}, 行数={len(df_1012):,}, 列数={len(df_1012.columns)}")

    combined = pd.concat([df_19, df_1012], ignore_index=True)
    combined.to_pickle(out)
    print(f"  已保存: {out} (总行数: {len(combined):,})")
    return combined

# ---------------------------------------------------------------------------
# 发票：妙可 SAP发票拆分月份 下 2025-01～12，concat(join='outer') 列对齐
# ---------------------------------------------------------------------------

def process_invoice_full_year(invoice_dir, output_pkl):
    """读取 2025-01～12 月 Excel，concat(join='outer') 后保存为 pkl。"""
    invoice_dir = Path(invoice_dir)
    out = Path(output_pkl)
    if not invoice_dir.exists():
        raise FileNotFoundError(f"发票目录不存在: {invoice_dir}")

    months = [f"{i:02d}" for i in range(1, 13)]
    all_dfs = []
    for m in months:
        for ext in ('XLSX', 'xlsx'):
            p = invoice_dir / f"2025-{m}.{ext}"
            if p.exists():
                try:
                    df = pd.read_excel(p, engine="openpyxl")
                    df["数据源文件"] = p.name
                    all_dfs.append(df)
                    print(f"  读取: {p.name} 行数={len(df):,} 列数={len(df.columns)}")
                except Exception as e:
                    print(f"  读取失败 {p.name}: {e}")
                break
    if not all_dfs:
        raise ValueError("未成功读取任何 1–12 月发票 Excel")

    combined = pd.concat(all_dfs, ignore_index=True, join="outer")
    combined.to_pickle(out)
    print(f"  已保存: {out} (总行数: {len(combined):,}, 列数: {len(combined.columns)})")
    return combined

# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("2025年全年 订单、发货、SAP 发票 预处理")
    print("=" * 60)

    base = Path.cwd()
    data_dir = base / "OMS25年1-12月订单及发货数据"
    invoice_dir = base / "妙可 SAP发票拆分月份"
    if not invoice_dir.exists():
        invoice_dir = base / "SAP开票数据"

    if not data_dir.exists():
        raise FileNotFoundError(f"目录不存在: {data_dir}")

    # 1. 订单
    print("\n【1. 订单】")
    order_pkl = base / "2025年全年OMS订单.pkl"
    process_orders_full_year(str(data_dir), str(order_pkl))

    # 2. 发货
    print("\n【2. 发货】")
    delivery_pkl = base / "2025年全年OMS发货.pkl"
    process_delivery_full_year(str(data_dir), str(delivery_pkl))

    # 3. 发票
    print("\n【3. 发票】")
    invoice_pkl = base / "2025年全年SAP原始数据.pkl"
    if invoice_dir.exists():
        process_invoice_full_year(str(invoice_dir), str(invoice_pkl))
    else:
        raise FileNotFoundError("未找到发票目录: 妙可 SAP发票拆分月份 或 SAP开票数据")

    print("\n" + "=" * 60)
    print("全年预处理完成，产出：")
    for p in (order_pkl, delivery_pkl, invoice_pkl):
        print(f"  - {p}")
    print("=" * 60)

if __name__ == "__main__":
    main()
