# -*- coding: utf-8 -*-
"""
BI ETL Pipeline - Phase 1: Process promotions & calculate real profit
"""
import pandas as pd
import numpy as np
import os
from datetime import date

BASE = r'E:\Desktop\工作文件（月度）\claw制作BI'
DATA = os.path.join(BASE, '数据表')
WAREHOUSE = os.path.join(BASE, 'warehouse')
os.makedirs(WAREHOUSE, exist_ok=True)

# =============================================
# 1. Load channel-store mapping
# =============================================
mapping = pd.read_excel(os.path.join(BASE, 'channel_store_mapping.xlsx'))
print('=== Channel-Store Mapping ===')
print(f'Total: {len(mapping)} rows')

# Build lookup: (channel, channel_store_id) -> qn_store_name
store_lookup = {}
for _, row in mapping.iterrows():
    key = (row['channel'], str(row['channel_store_id']))
    store_lookup[key] = row['qn_store_name']

# Also build: (channel, channel_store_id) -> qn_store_id
id_lookup = {}
for _, row in mapping.iterrows():
    key = (row['channel'], str(row['channel_store_id']))
    id_lookup[key] = row['qn_store_id']

# =============================================
# 2. Process Eleme Promotion
# =============================================
print('\n=== Processing Eleme Promotion ===')
df_el = pd.read_excel(os.path.join(DATA, '饿了么推广.xlsx'))
df_el['日期'] = pd.to_datetime(df_el['日期']).dt.date

# Aggregate by store + date (same store may have multiple plans)
el_promo = df_el.groupby(['日期', '门店ID']).agg(
    promo_cash=('推广现金消费(元)', 'sum'),
    promo_total=('推广消费(元)', 'sum'),
).reset_index()

# Map store ID to store name
el_promo['channel'] = '饿了么'
el_promo['channel_store_id'] = el_promo['门店ID'].astype(str)
el_promo['store_name'] = el_promo.apply(
    lambda r: store_lookup.get(('饿了么', r['channel_store_id']), f"UNKNOWN_{r['channel_store_id']}"), axis=1)
el_promo['qn_store_id'] = el_promo.apply(
    lambda r: id_lookup.get(('饿了么', r['channel_store_id']), ''), axis=1)

print(f'Rows: {len(el_promo)}')
print(f'Date: {el_promo["日期"].unique()}')
print(f'Total Eleme promo: {el_promo["promo_cash"].sum():.2f}')
print(el_promo[['日期', 'store_name', 'promo_cash']].to_string(index=False))

# =============================================
# 3. Process Meituan Promotion
# =============================================
print('\n=== Processing Meituan Promotion ===')
df_mt = pd.read_excel(os.path.join(DATA, '美团推广.xlsx'), sheet_name='推广费流水')

# Filter: only "推广订单扣款"
df_mt = df_mt[df_mt['类型'] == '推广订单扣款'].copy()

# Convert date & shift -1 day (Meituan settles at midnight for previous day)
df_mt['日期'] = pd.to_datetime(df_mt['日期']).dt.date - pd.Timedelta(days=1)

# Convert amount to positive
df_mt['promo_amount'] = df_mt['金额(元)'].abs()

# Aggregate by store + date
mt_promo = df_mt.groupby(['日期', '门店id']).agg(
    promo_cash=('promo_amount', 'sum'),
).reset_index()

# Map store ID to store name
mt_promo['channel'] = '美团闪购'
mt_promo['channel_store_id'] = mt_promo['门店id'].astype(str)
mt_promo['store_name'] = mt_promo.apply(
    lambda r: store_lookup.get(('美团闪购', r['channel_store_id']), f"UNKNOWN_{r['channel_store_id']}"), axis=1)
mt_promo['qn_store_id'] = mt_promo.apply(
    lambda r: id_lookup.get(('美团闪购', r['channel_store_id']), ''), axis=1)

print(f'Rows (after filter): {len(mt_promo)}')
print(f'Date: {mt_promo["日期"].unique()}')
print(f'Total Meituan promo: {mt_promo["promo_cash"].sum():.2f}')
print(mt_promo[['日期', 'store_name', 'promo_cash']].to_string(index=False))

# =============================================
# 3. Load Commission Rates
# =============================================
print('\n=== Loading Commission Rates ===')
df_comm = pd.read_excel(os.path.join(DATA, '门店抽佣点数明细表.xlsx'))
comm_lookup = {}
for _, row in df_comm.iterrows():
    comm_lookup[int(row['门店id'])] = row['抽佣点数']  # Keep as int to match qn_store_id
print(f'Loaded {len(comm_lookup)} store commission rates')

# =============================================
# 4. Combine all promotions
# =============================================
print('\n=== Combined Promotion Summary ===')
el_out = el_promo[['日期', 'store_name', 'qn_store_id', 'channel', 'promo_cash']].copy()
el_out.columns = ['日期', 'store_name', 'qn_store_id', 'channel', 'promo_fee']

mt_out = mt_promo[['日期', 'store_name', 'qn_store_id', 'channel', 'promo_cash']].copy()
mt_out.columns = ['日期', 'store_name', 'qn_store_id', 'channel', 'promo_fee']

all_promo = pd.concat([el_out, mt_out], ignore_index=True)
print(f'Total promo records: {len(all_promo)}')
print(f'Total promo amount: {all_promo["promo_fee"].sum():.2f}')

promo_path = os.path.join(WAREHOUSE, 'promo_daily.xlsx')
all_promo.to_excel(promo_path, index=False)
print(f'Saved: {promo_path}')

# =============================================
# 5. Process Order Detail - Calculate Real Profit
# =============================================
print('\n=== Processing Order Detail ===')
df_order = pd.read_excel(os.path.join(DATA, '实时订单明细.xlsx'), header=1)
df_order = df_order.rename(columns={'Unnamed: 4': 'order_time'})
df_order['order_date'] = pd.to_datetime(df_order['order_time']).dt.date

# Add qn_store_id from mapping (via store name)
name_to_id = {}
for _, row in mapping.drop_duplicates(subset='qn_store_name').iterrows():
    name_to_id[row['qn_store_name']] = row['qn_store_id']

df_order['qn_store_id'] = df_order['门店'].map(name_to_id)

# =============================================
# 5b. Supplement Meituan delivery fees from 美团财务明细
#     (For Meituan orders where 三方配送费=0, check 美团财务明细 for actual fee)
# =============================================
print('\n=== Supplementing Meituan Delivery Fees ===')
df_mt_finance = pd.read_excel(os.path.join(DATA, '美团财务明细.xlsx'), header=1)

# Filter: only 配送费用 rows
df_mt_fee = df_mt_finance[df_mt_finance['交易类型'] == '配送费用'].copy()
print(f'配送费用 rows: {len(df_mt_fee)}')

# Build lookup: order_id → delivery_fee (positive)
mt_delivery_lookup = {}
for _, row in df_mt_fee.iterrows():
    order_id = str(row['订单号']).strip()
    fee = abs(row['商家应收款（结算金额）'])
    if fee > 0:
        mt_delivery_lookup[order_id] = fee

print(f'Unique orders with delivery fee in 美团财务明细: {len(mt_delivery_lookup)}')

# Find Meituan orders with 三方配送费 = 0
mt_zero_delivery = df_order[
    (df_order['渠道名称'] == '美团闪购') & 
    (df_order['三方配送费'] == 0)
].copy()

print(f'Meituan orders with 三方配送费=0: {len(mt_zero_delivery)}')

# Try to supplement
supplemented = 0
for idx, row in mt_zero_delivery.iterrows():
    order_id = str(row['订单号']).strip()
    if order_id in mt_delivery_lookup:
        df_order.at[idx, '三方配送费'] = mt_delivery_lookup[order_id]
        supplemented += 1

print(f'Supplemented delivery fees: {supplemented} orders')

# Delivery order count (orders with delivery_fee > 0)
delivery_ord = df_order[df_order['三方配送费'] > 0].groupby(
    ['order_date', '门店', '渠道名称']
).agg(delivery_order_cnt=('订单号', 'count')).reset_index()
delivery_ord.columns = ['日期', 'store_name', 'channel', 'delivery_order_cnt']

# Aggregate by date + store + channel
daily = df_order.groupby(['order_date', '门店', 'qn_store_id', '渠道名称']).agg(
    order_cnt=('订单号', 'count'),
    revenue=('实付金额', 'sum'),
    gross_profit=('线上毛利', 'sum'),
    income=('收入', 'sum'),
    cost=('成本', 'sum'),
    goods_cost=('商品成本', 'sum'),
    delivery_fee=('三方配送费', 'sum'),
    commission=('佣金', 'sum'),
    delivery_income=('配送收入', 'sum'),
).reset_index()

daily.columns = ['日期', 'store_name', 'qn_store_id', 'channel', 'order_cnt',
                 'revenue', 'gross_profit', 'income', 'cost', 'goods_cost',
                 'delivery_fee', 'commission', 'delivery_income']

# Negative profit count
neg_daily = df_order[df_order['线上毛利'] < 0].groupby(
    ['order_date', '门店', '渠道名称']
).agg(neg_cnt=('订单号', 'count')).reset_index()
neg_daily.columns = ['日期', 'store_name', 'channel', 'neg_cnt']

daily = daily.merge(neg_daily, on=['日期', 'store_name', 'channel'], how='left')
daily['neg_cnt'] = daily['neg_cnt'].fillna(0).astype(int)
daily['neg_pct'] = (daily['neg_cnt'] / daily['order_cnt'] * 100).round(1)

# Merge delivery order count
daily = daily.merge(delivery_ord, on=['日期', 'store_name', 'channel'], how='left')
daily['delivery_order_cnt'] = daily['delivery_order_cnt'].fillna(0).astype(int)
daily['avg_delivery_cost'] = daily.apply(
    lambda r: round(r['delivery_fee'] / r['delivery_order_cnt'], 2) if r['delivery_order_cnt'] > 0 else 0, axis=1
)

# Calculate gross margin rate
daily['gross_margin_rate'] = (daily['gross_profit'] / daily['revenue'] * 100).round(2)

print(f'Daily aggregation rows: {len(daily)}')
print(f'Date: {daily["日期"].unique()}')

# =============================================
# 6. Join with promotion to get real profit
# =============================================
# Build promo lookup: (date, store_name, channel) -> promo_fee
promo_lookup = {}
for _, row in all_promo.iterrows():
    key = (str(row['日期']), row['store_name'], row['channel'])
    promo_lookup[key] = row['promo_fee']

daily['promo_fee'] = daily.apply(
    lambda r: promo_lookup.get((str(r['日期']), r['store_name'], r['channel']), 0), axis=1)

daily['real_profit'] = (daily['gross_profit'] - daily['promo_fee']).round(2)
daily['real_margin_rate'] = np.where(
    daily['revenue'] > 0,
    (daily['real_profit'] / daily['revenue'] * 100).round(2),
    0
)

# =============================================
# 7. Add Commission Cost & Delivery Metrics
# =============================================
daily['commission_rate'] = daily['qn_store_id'].map(comm_lookup).fillna(0)
daily['commission_cost'] = (daily['revenue'] * daily['commission_rate']).round(2)  # 抽佣毛利
daily['commission_margin'] = np.where(
    daily['revenue'] > 0,
    (daily['commission_cost'] / daily['revenue'] * 100).round(2),
    0
)

print('\n=== Sample with real profit ===')
sample = daily[['日期', 'store_name', 'channel', 'order_cnt', 'revenue',
                'gross_profit', 'promo_fee', 'real_profit', 'neg_pct']].sort_values('revenue', ascending=False).head(15)
print(sample.to_string(index=False))

# =============================================
# 8. Save warehouse
# =============================================
wh_path = os.path.join(WAREHOUSE, 'daily_store_channel_profit.xlsx')
daily.to_excel(wh_path, index=False)
print(f'\nSaved: {wh_path} ({len(daily)} rows)')

# Summary stats
print('\n=== FINAL SUMMARY ===')
print(f'Orders: {daily["order_cnt"].sum()}')
print(f'Revenue: {daily["revenue"].sum():.2f}')
print(f'Gross profit (before promo): {daily["gross_profit"].sum():.2f}')
print(f'Total promo: {daily["promo_fee"].sum():.2f}')
print(f'Real profit (after promo): {daily["real_profit"].sum():.2f}')
print(f'Commission cost: {daily["commission_cost"].sum():.2f}')
print(f'Total delivery fee: {daily["delivery_fee"].sum():.2f}')
print(f'Total delivery orders: {daily["delivery_order_cnt"].sum()}')
print(f'Avg delivery cost: {daily["avg_delivery_cost"].mean():.2f}')
print(f'Overall margin rate: {daily["gross_profit"].sum()/daily["revenue"].sum()*100:.1f}%')
print(f'Real margin rate: {daily["real_profit"].sum()/daily["revenue"].sum()*100:.1f}%')
