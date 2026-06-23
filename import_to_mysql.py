#!/usr/bin/env python3
"""增量导入 Excel 到 MySQL smile_ke_bi

流程：
  stores / store_cost / commission_rate → 全量覆盖
  daily_profit / daily_promo → 按日期：DELETE 该日旧数据 → INSERT 新数据
"""

import pandas as pd
import pymysql
import os, numpy as np

BASE = r'E:\Desktop\工作文件（月度）\claw制作BI'

conn = pymysql.connect(host='localhost', user='root', password='', database='smile_ke_bi', charset='utf8mb4')
cursor = conn.cursor()

def safe_val(v, cast=str):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None if cast == str else 0
    return cast(v)

def bulk_insert(table, columns, rows, batch_size=500):
    cols_sql = ', '.join(f'`{c}`' for c in columns)
    ph = ', '.join(['%s'] * len(columns))
    sql = f'INSERT INTO {table} ({cols_sql}) VALUES ({ph})'
    total = 0
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i+batch_size]
        try:
            cursor.executemany(sql, batch)
            total += len(batch)
        except Exception as e:
            for r in batch:
                try: cursor.execute(sql, r); total += 1
                except: pass
    conn.commit()
    return total

def upsert_table(table, pk_cols, columns, rows, batch_size=300):
    """先按主键 DELETE 再 INSERT（多行批量优化版）"""
    if not rows: return 0
    
    # 收集所有唯一键值，按日期范围删除
    if pk_cols == ['dt']:
        dates = sorted(set(r[0] for r in rows))
        if len(dates) <= 10:
            for d in dates:
                cursor.execute(f'DELETE FROM {table} WHERE dt=%s', (d,))
        else:
            min_d, max_d = min(dates), max(dates)
            cursor.execute(f'DELETE FROM {table} WHERE dt>=%s AND dt<=%s', (min_d, max_d))
    else:
        # 复合主键：逐行删除（数据量小，可接受）
        pk_sql = ' AND '.join(f'`{c}`=%s' for c in pk_cols)
        for r in rows:
            vals = [r[columns.index(c)] for c in pk_cols]
            cursor.execute(f'DELETE FROM {table} WHERE {pk_sql}', vals)
    conn.commit()
    
    # INSERT
    return bulk_insert(table, columns, rows, batch_size)

# ============================================================
# === 1. 门店基本信息（全量覆盖）===
print('1/5 门店基本信息...')
stores_df = pd.read_excel(os.path.join(BASE, '数据表', '基础信息表', '门店基本信息.xlsx'))
cursor.execute('TRUNCATE TABLE stores')
rows = []
for _, r in stores_df.iterrows():
    sn = safe_val(r.get('门店名称'))
    rows.append((sn, sn, safe_val(r.get('门店ID')), safe_val(r.get('所属地区')), safe_val(r.get('渠道名'))))
n = bulk_insert('stores', ['store_name','short_name','qn_store_id','region','channels'], rows)
print(f'  -> {n} 行')

# === 2. 门店运营成本（全量覆盖）===
print('2/5 门店运营成本...')
cost_df = pd.read_excel(os.path.join(BASE, '数据表', '基础信息表', '门店运营成本.xlsx'))
cursor.execute('TRUNCATE TABLE store_cost')
rows = []
for _, r in cost_df.iterrows():
    cost = safe_val(r.get('门店运营成本'), float)
    if cost > 0:
        rows.append((safe_val(r.get('牵牛花门店ID')), safe_val(r.get('门店')), cost))
n = bulk_insert('store_cost', ['qn_store_id','store_name','monthly_cost'], rows)
print(f'  -> {n} 行')

# === 3. 门店抽佣点数（全量覆盖）===
print('3/5 门店抽佣点数...')
rate_df = pd.read_excel(os.path.join(BASE, '数据表', '基础信息表', '门店抽佣点数明细表.xlsx'))
cursor.execute('TRUNCATE TABLE commission_rate')
rows, seen = [], set()
for _, r in rate_df.iterrows():
    rate = safe_val(r.get('抽佣点数'), float)
    if rate > 1: rate = rate / 100
    sn = safe_val(r.get('门店名称'))
    if sn not in seen:
        seen.add(sn)
        rows.append((safe_val(r.get('门店id')), sn, rate))
n = bulk_insert('commission_rate', ['qn_store_id','store_name','rate'], rows)
print(f'  -> {n} 行')

# === 4. 每日门店渠道利润（增量：按日期覆盖）===
print('4/5 每日门店渠道利润...')
df = pd.read_excel(os.path.join(BASE, 'warehouse', 'daily_store_channel_profit.xlsx'))
cols_db = ['dt','store_name','qn_store_id','channel','order_cnt','revenue',
    'gross_profit','income','cost','goods_cost','delivery_fee','commission',
    'delivery_income','neg_cnt','neg_pct','delivery_order_cnt','avg_delivery_cost',
    'gross_margin_rate','promo_fee','real_profit','real_margin_rate',
    'commission_rate','commission_fee','commission_profit','commission_margin',
    'total_cost','total_qty']
cols_excel = cols_db[1:]
rows = []
for _, r in df.iterrows():
    dt = pd.to_datetime(r['日期']).strftime('%Y-%m-%d')
    row = [dt]
    for c in cols_excel:
        v = r.get(c, 0)
        if v is None or (isinstance(v, float) and np.isnan(v)): v = 0
        row.append(v)
    rows.append(row)

# 按日期 DELETE 旧数据
dates_in_excel = sorted(df['日期'].dropna().unique())
if len(dates_in_excel) > 0:
    date_strs = [pd.to_datetime(d).strftime('%Y-%m-%d') for d in dates_in_excel]
    if len(date_strs) <= 10:
        for d in date_strs:
            cursor.execute('DELETE FROM daily_profit WHERE dt=%s', (d,))
    else:
        min_d, max_d = min(date_strs), max(date_strs)
        cursor.execute('DELETE FROM daily_profit WHERE dt>=%s AND dt<=%s', (min_d, max_d))
conn.commit()
n = bulk_insert('daily_profit', cols_db, rows, 300)
print(f'  Excel有 {len(date_strs)} 个日期，覆盖后 -> {n} 行')

# === 5. 每日推广费（增量：按日期覆盖）===
print('5/5 每日推广费...')
promo_df = pd.read_excel(os.path.join(BASE, 'warehouse', 'promo_daily.xlsx'))
rows = []
for _, r in promo_df.iterrows():
    rows.append([
        pd.to_datetime(r['日期']).strftime('%Y-%m-%d'),
        safe_val(r.get('store_name')),
        safe_val(r.get('qn_store_id')),
        safe_val(r.get('channel')),
        safe_val(r.get('promo_fee'), float)
    ])

dates_in_promo = sorted(promo_df['日期'].dropna().unique())
if len(dates_in_promo) > 0:
    date_strs = [pd.to_datetime(d).strftime('%Y-%m-%d') for d in dates_in_promo]
    if len(date_strs) <= 10:
        for d in date_strs:
            cursor.execute('DELETE FROM daily_promo WHERE dt=%s', (d,))
    else:
        min_d, max_d = min(date_strs), max(date_strs)
        cursor.execute('DELETE FROM daily_promo WHERE dt>=%s AND dt<=%s', (min_d, max_d))
conn.commit()
n = bulk_insert('daily_promo', ['dt','store_name','qn_store_id','channel','promo_fee'], rows)
print(f'  Excel有 {len(date_strs)} 个日期，覆盖后 -> {n} 行')

# === 验证 ===
print('\n===== 汇总 =====')
for table in ['stores', 'store_cost', 'commission_rate', 'daily_profit', 'daily_promo']:
    cursor.execute(f'SELECT COUNT(*) FROM {table}')
    print(f'  {table}: {cursor.fetchone()[0]:,} 行')
cursor.execute("SELECT MIN(dt), MAX(dt), COUNT(DISTINCT dt), COUNT(DISTINCT store_name) FROM daily_profit")
r = cursor.fetchone()
print(f'  daily_profit: {r[0]} ~ {r[1]} | {r[2]}天 | {r[3]}门店')

cursor.close()
conn.close()
print('\n✅ 完成')
