# -*- coding: utf-8 -*-
"""
BI ETL Pipeline v3
 - 自动识别文件名（无需用户改名）
 - 去重：同日同店多文件不重复计算
 - 盈亏分析优先覆盖实时订单
 - 推广数据自动合并
 - 配送费矫正可选
"""
import pandas as pd
import numpy as np
import os, shutil, glob
from datetime import date, datetime

BASE = r'E:\Desktop\工作文件（月度）\claw制作BI'
DATA = os.path.join(BASE, '数据表')
CALIB_DIR = os.path.join(DATA, '数据校准')
ARCHIVE = os.path.join(DATA, '已处理')
WAREHOUSE = os.path.join(BASE, 'warehouse')
os.makedirs(WAREHOUSE, exist_ok=True)
os.makedirs(ARCHIVE, exist_ok=True)

# =============================================
# 0. 自动发现并分类文件
# =============================================
def discover_files():
    all_dirs = [d for d in [DATA, CALIB_DIR] if os.path.exists(d)]
    # Also scan month subdirectories (e.g. 数据表/202605/)
    expanded_dirs = list(all_dirs)
    for d in list(all_dirs):
        for sub in os.listdir(d):
            subp = os.path.join(d, sub)
            if os.path.isdir(subp) and not sub.startswith('.') and not sub.startswith('~'):
                expanded_dirs.append(subp)
    result = {'eleme_promo':[], 'mt_promo':[], 'profit_analysis':[], 'order_detail':[], 'mt_finance':[]}
    for d in expanded_dirs:
        for f in os.listdir(d):
            if not f.endswith('.xlsx') or f.startswith('~$'): continue
            fp = os.path.join(d, f)
            # 饿了么推广（原始名: 门店维度分日期 / 规范名: 饿了么推广）
            if '门店维度分日期' in f or ('饿了么' in f and '推广' in f):
                result['eleme_promo'].append(fp)
                continue
            # 美团推广（原始名: 全部门店...订单 / 规范名: 美团推广）
            if ('全部门店' in f and '订单' in f and '结算' not in f and '计费' not in f) or ('美团' in f and '推广' in f):
                result['mt_promo'].append(fp)
                continue
            # 美团财务明细（原始名: 结算账单/订单计费明细 / 规范名: 美团财务明细）
            if ('美团' in f and '财务明细' in f) or '结算账单' in f or '订单计费明细' in f:
                result['mt_finance'].append(fp)
                continue
            # 盈亏分析（规范名: 盈亏分析 / 原始名: 导出财务分析订单数据）
            if '盈亏分析' in f or '导出财务分析订单数据' in f:
                result['profit_analysis'].append(fp)
                continue
            # 实时订单明细（原始名: 导出订单明细/实时订单明细）
            if '导出订单明细' in f or '实时订单明细' in f:
                result['order_detail'].append(fp)
                continue

    for k in result:
        result[k] = sorted(list(set(result[k])), key=lambda p: os.path.getmtime(p))

    labels = {'eleme_promo':'饿了么推广','mt_promo':'美团推广','profit_analysis':'盈亏分析',
              'order_detail':'实时订单明细','mt_finance':'美团财务明细'}
    print('=== Discovered Files ===')
    for k, v in result.items():
        if v:
            print(f'  [{labels[k]}] {len(v)} files')
            for fp in v:
                print(f'    {os.path.basename(fp)} ({os.path.getsize(fp)//1024}KB)')
        else:
            print(f'  [{labels[k]}] (none)')
    return result

def archive(fp_list):
    for fp in fp_list:
        try:
            target = os.path.join(ARCHIVE, os.path.basename(fp))
            if os.path.exists(target):
                n, e = os.path.splitext(os.path.basename(fp))
                target = os.path.join(ARCHIVE, f'{n}_{datetime.now().strftime("%H%M%S")}{e}')
            shutil.move(fp, target)
        except Exception as ex:
            print(f'  Archive warn: {os.path.basename(fp)} - {ex}')

# =============================================
# Run discovery
# =============================================
files = discover_files()

# =============================================
# 1. Load channel-store mapping
# =============================================
mapping = pd.read_excel(os.path.join(BASE, 'channel_store_mapping.xlsx'))
store_lookup, id_lookup = {}, {}
for _, row in mapping.iterrows():
    key = (row['channel'], str(row['channel_store_id']))
    store_lookup[key] = row['qn_store_name']
    id_lookup[key] = row['qn_store_id']
print(f'Channel mapping: {len(mapping)} rows')

# Commission rates
df_comm = pd.read_excel(os.path.join(DATA, '基础信息表', '门店抽佣点数明细表.xlsx'))
comm_lookup = {int(row['门店id']): row['抽佣点数'] for _, row in df_comm.iterrows()}
print(f'Commission rates: {len(comm_lookup)} stores')

# =============================================
# 丰派线下渠道处理
# =============================================
def process_fengpai(fp_file, online_store_names, cost_map, upc_to_cost):
    """处理丰派商品销售流水表，返回按门店+日期汇总的DataFrame"""
    import re
    df = pd.read_excel(fp_file)
    if '交易类型' not in df.columns:
        print(f'  [丰派] 跳过（无交易类型列）')
        return None
    
    df = df[df['交易类型'] == '销售'].copy()
    if len(df) == 0:
        return None
    # 只算订单来源=收银机的（线下POS流水）
    if '订单来源' in df.columns:
        before = len(df)
        df = df[df['订单来源'] == '收银机'].copy()
        print(f'  [丰派] 过滤收银机: {len(df)}/{before} 行')
    
    df['交易时间'] = pd.to_datetime(df['交易时间'], errors='coerce')
    df = df.dropna(subset=['交易时间'])
    months = df['交易时间'].dt.to_period('M').unique()
    print(f'  [丰派] {len(df)} 行销售, 月份: {sorted(str(m) for m in months)}')
    
    # 门店映射
    SPECIAL_FP_MAP = {
        'B019-微笑客(惠州中园一路店)': 'Q019-微笑客（博罗中园路店）',
        'B019-微笑客（惠州中园一路店）': 'Q019-微笑客（博罗中园路店）',
    }
    code_to_full = {}
    for s in online_store_names:
        m = re.match(r'([A-Z]+\d+)\s*-\s*(.+)', s)
        if m:
            code_to_full[m.group(1).strip()] = s
    
    def match_fp(fp_name):
        fp_name = str(fp_name).strip()
        if fp_name in SPECIAL_FP_MAP:
            return SPECIAL_FP_MAP[fp_name]
        fp_cn = fp_name.replace('(','（').replace(')','）')
        if fp_cn in SPECIAL_FP_MAP:
            return SPECIAL_FP_MAP[fp_cn]
        if fp_cn in online_store_names:
            return fp_cn
        m = re.match(r'([A-Z]+\d+)\s*-\s*(.+)', fp_name)
        if m and m.group(1).strip() in code_to_full:
            return code_to_full[m.group(1).strip()]
        return None
    
    df['store_name'] = df['门店名称'].apply(match_fp)
    before = len(df)
    df = df[df['store_name'].notna()].copy()
    print(f'  [丰派] 门店匹配: {len(df)}/{before} 行')
    
    # 商品成本匹配
    sku_col = df['商品货号'].astype(str).str.strip()
    upc_col = df['商品条码'].astype(str).str.strip()
    retail_col = pd.to_numeric(df['零售价'], errors='coerce').fillna(0)
    
    cost1 = sku_col.map(cost_map)
    m1 = cost1.notna()
    cost2 = upc_col.map(upc_to_cost)
    m2 = (~m1) & cost2.notna()
    cost3 = (retail_col * 0.75).round(2)
    m3 = (~m1) & (~m2)
    
    df['unit_cost'] = cost1.fillna(0)
    df.loc[m2, 'unit_cost'] = cost2[m2]
    df.loc[m3, 'unit_cost'] = cost3[m3]
    print(f'  [丰派] 成本: SKU={m1.sum()}, UPC={m2.sum()}, 兜底={m3.sum()}')
    
    # 计算
    df['qty'] = pd.to_numeric(df['商品数量'], errors='coerce').fillna(0)
    df['revenue'] = pd.to_numeric(df['应收金额'], errors='coerce').fillna(0)
    df['total_cost'] = df['unit_cost'] * df['qty']
    df['gross_profit'] = df['revenue'] - df['total_cost']
    df['日期'] = df['交易时间'].dt.date
    
    daily_fp = df.groupby(['日期','store_name']).agg(
        order_cnt=('订单号','nunique'),
        revenue=('revenue','sum'),
        gross_profit=('gross_profit','sum'),
        total_cost=('total_cost','sum'),
        total_qty=('qty','sum'),
    ).reset_index()
    daily_fp['channel'] = '线下'
    daily_fp['promo_fee'] = 0
    daily_fp['commission_rate'] = 0
    daily_fp['commission_fee'] = 0
    daily_fp['commission_profit'] = daily_fp['gross_profit']
    daily_fp['commission_margin'] = np.where(
        daily_fp['revenue']>0, (daily_fp['gross_profit']/daily_fp['revenue']*100).round(2), 0)
    daily_fp['delivery_fee'] = 0
    daily_fp['delivery_order_cnt'] = 0
    daily_fp['avg_delivery_cost'] = 0
    daily_fp['neg_cnt'] = 0
    daily_fp['neg_pct'] = 0
    daily_fp['gross_margin_rate'] = np.where(
        daily_fp['revenue']>0, (daily_fp['gross_profit']/daily_fp['revenue']*100).round(2), 0)
    daily_fp['income'] = daily_fp['revenue']
    daily_fp['cost'] = daily_fp['total_cost']
    daily_fp['goods_cost'] = daily_fp['total_cost']
    daily_fp['commission'] = 0
    daily_fp['delivery_income'] = 0
    daily_fp['real_margin_rate'] = daily_fp['gross_margin_rate']
    daily_fp['qn_store_id'] = None  # 将在外面填充
    daily_fp['日期'] = pd.to_datetime(daily_fp['日期'])
    
    return daily_fp

# Store name normalization
def norm_name(n):
    if pd.isna(n): return n
    n = str(n).replace('(', '（').replace(')', '）')
    fixes = {'微笑客（深圳北站店）': 'B075-微笑客（深圳北站店）'}
    return fixes.get(n, n)

# =============================================
# 2. Process ALL promotions (merge + dedup)
# =============================================
print('\n=== Processing Promotions ===')
all_promo_parts = []

# --- Eleme ---
for fp in files['eleme_promo']:
    df = pd.read_excel(fp)
    df['日期'] = pd.to_datetime(df['日期']).dt.date
    el = df.groupby(['日期', '门店ID']).agg(promo_cash=('推广现金消费(元)', 'sum')).reset_index()
    el['channel'] = '饿了么'
    el['channel_store_id'] = el['门店ID'].astype(str)
    el['store_name'] = el.apply(lambda r: store_lookup.get(('饿了么',r['channel_store_id']),f"UNK_{r['channel_store_id']}"), axis=1)
    el['qn_store_id'] = el.apply(lambda r: id_lookup.get(('饿了么',r['channel_store_id']),''), axis=1)
    el['promo_fee'] = el['promo_cash']
    all_promo_parts.append(el[['日期','store_name','qn_store_id','channel','promo_fee']])
    print(f'  Eleme {os.path.basename(fp)}: {len(el)} rows, {el["日期"].nunique()}d, ¥{el["promo_fee"].sum():.0f}')

# --- Meituan ---
for fp in files['mt_promo']:
    try:
        df = pd.read_excel(fp, sheet_name='推广费流水')
    except:
        df = pd.read_excel(fp)  # fallback if sheet name wrong
    if '类型' in df.columns:
        df = df[df['类型'] == '推广订单扣款'].copy()
    # Handle column names: 门店id or 门店ID
    mt_id_col = '门店id' if '门店id' in df.columns else '门店ID'
    df['日期'] = pd.to_datetime(df['日期']).dt.date - pd.Timedelta(days=1)
    df['金额(元)'] = pd.to_numeric(df['金额(元)'], errors='coerce')
    df['promo_amount'] = df['金额(元)'].abs()
    mt = df.groupby(['日期', mt_id_col]).agg(promo_cash=('promo_amount','sum')).reset_index()
    mt['channel'] = '美团闪购'
    mt['channel_store_id'] = mt[mt_id_col].astype(str)
    mt['store_name'] = mt.apply(lambda r: store_lookup.get(('美团闪购',r['channel_store_id']),f"UNK_{r['channel_store_id']}"), axis=1)
    mt['qn_store_id'] = mt.apply(lambda r: id_lookup.get(('美团闪购',r['channel_store_id']),''), axis=1)
    mt['promo_fee'] = mt['promo_cash']
    all_promo_parts.append(mt[['日期','store_name','qn_store_id','channel','promo_fee']])
    print(f'  MT {os.path.basename(fp)}: {len(mt)} rows, {mt["日期"].nunique()}d, ¥{mt["promo_fee"].sum():.0f}')

if all_promo_parts:
    all_promo = pd.concat(all_promo_parts, ignore_index=True)
    # Dedup: keep last (newest) for duplicate (date,store,channel)
    all_promo = all_promo.drop_duplicates(subset=['日期','store_name','channel'], keep='last')
    print(f'Promo total (dedup): {len(all_promo)} records, ¥{all_promo["promo_fee"].sum():.0f}')
else:
    # Load existing promo data from warehouse
    promo_path = os.path.join(WAREHOUSE, 'promo_daily.xlsx')
    if os.path.exists(promo_path):
        all_promo = pd.read_excel(promo_path)
        print(f'Promo: loaded {len(all_promo)} existing records from warehouse')
    else:
        all_promo = pd.DataFrame(columns=['日期','store_name','qn_store_id','channel','promo_fee'])
        print('Promo: no data')

# Build promo lookup
promo_lookup = {}
for _, row in all_promo.iterrows():
    promo_lookup[(str(row['日期']), row['store_name'], row['channel'])] = row['promo_fee']

# Save promo
all_promo.to_excel(os.path.join(WAREHOUSE, 'promo_daily.xlsx'), index=False)

# =============================================
# 3. Process Order Detail (实时) then Calibrate (盈亏)
# =============================================
def process_orders(order_paths, revenue_col='预计收入'):
    """处理订单文件，返回按日期-门店-渠道的汇总DataFrame"""
    all_daily = []
    for fp in order_paths:
        df = pd.read_excel(fp, header=1)
        df.columns = [str(c).strip() for c in df.columns]
        # Find date column (Unnamed or 时间 column at position 4)
        cols = list(df.columns)
        for i, c in enumerate(cols):
            if i == 4 and ('Unnamed' in c or '时间' in c):
                df = df.rename(columns={c: 'order_time'})
                break
        else:
            # Fallback: try to find 下单时间 or similar
            for c in cols:
                if '时间' in c:
                    df = df.rename(columns={c: 'order_time'})
                    break
        # 分开处理：有效订单和无效订单(毛利率='-')
        df['_invalid'] = df['线上毛利率'] == '-'
        df_invalid = df[df['_invalid']].copy()
        df = df[~df['_invalid']].copy()
        # 提前解析日期（无效订单也需要）
        df_invalid['order_date'] = pd.to_datetime(df_invalid['order_time']).dt.date
        df_invalid['门店'] = df_invalid['门店'].astype(str)
        
        df['order_date'] = pd.to_datetime(df['order_time']).dt.date
        df['门店'] = df['门店'].astype(str)

        name_to_id = mapping.drop_duplicates(subset='qn_store_name').set_index('qn_store_name')['qn_store_id'].to_dict()

        # Delivery fee supplement from Meituan finance (SKIP for speed)
        mt_delivery = {}
        # for mf_fp in files['mt_finance']:
        #     mf = pd.read_excel(mf_fp, header=1)
        #     if '交易类型' not in mf.columns: continue
        #     mf_del = mf[mf['交易类型'].isin(['配送费用', '配送小费'])].copy()
        #     for _, r in mf_del.iterrows():
        #         oid = str(r['订单号']).strip()
        #         fee = abs(r['商家应收款（结算金额）'])
        #         if fee > 0:
        #             mt_delivery[oid] = mt_delivery.get(oid, 0) + fee

        # (mt_delivery 已跳过，配送费只用牵牛花原始数据)
        # if mt_delivery:
        #     no_fee = df[(df['渠道名称']=='美团闪购') & (df['三方配送费']==0)]
        #     supp = 0
        #     for idx, row in no_fee.iterrows():
        #         oid = str(row['订单号']).strip()
        #         if oid in mt_delivery:
        #             df.at[idx, '三方配送费'] = mt_delivery[oid]
        #             supp += 1
        #     if len(df_invalid) > 0:
        #         inv_no_fee = df_invalid[(df_invalid['渠道名称']=='美团闪购') & (df_invalid['三方配送费']==0)]
        #         supp2 = 0
        #         for idx, row in inv_no_fee.iterrows():
        #             oid = str(row['订单号']).strip()
        #             if oid in mt_delivery:
        #                 df_invalid.at[idx, '三方配送费'] = mt_delivery[oid]
        #                 supp2 += 1
        #         supp += supp2
        #     print(f'  Delivery supplement: {supp}/{len(no_fee)+len(inv_no_fee) if len(df_invalid)>0 else len(no_fee)} orders')

        groups = df.groupby(['order_date','门店','渠道名称'])

        daily = groups.agg(
            order_cnt=('订单号','count'),
            revenue=(revenue_col, 'sum'),
            gross_profit=('线上毛利','sum'),
            income=('收入','sum'), cost=('成本','sum'),
            goods_cost=('商品成本','sum'), delivery_fee=('三方配送费','sum'),
            commission=('佣金','sum'), delivery_income=('配送收入','sum'),
        ).reset_index()
        daily.columns = ['日期','store_name','channel','order_cnt','revenue','gross_profit',
                         'income','cost','goods_cost','delivery_fee','commission','delivery_income']

        # 无效订单：不计实收/单量/配送费，只统计毛利
        if len(df_invalid) > 0:
            inv_grp = df_invalid.groupby(['order_date','门店','渠道名称'])
            inv_agg = inv_grp.agg(invalid_profit=('线上毛利','sum')).reset_index()
            inv_agg.columns = ['日期','store_name','channel','invalid_profit']
            daily = daily.merge(inv_agg, on=['日期','store_name','channel'], how='left')
            daily['invalid_profit'] = daily['invalid_profit'].fillna(0)
            daily['gross_profit'] = daily['gross_profit'] + daily['invalid_profit']
            daily = daily.drop(columns=['invalid_profit'])

        # Neg profit
        neg = groups.apply(lambda g: (g['线上毛利']<0).sum()).reset_index()
        neg.columns = ['日期','store_name','channel','neg_cnt']
        daily = daily.merge(neg, on=['日期','store_name','channel'], how='left')
        daily['neg_cnt'] = daily['neg_cnt'].fillna(0).astype(int)

        # Delivery
        del_ord = df[df['三方配送费']>0].groupby(['order_date','门店','渠道名称']).agg(
            delivery_order_cnt=('订单号','count')).reset_index()
        del_ord.columns = ['日期','store_name','channel','delivery_order_cnt']
        daily = daily.merge(del_ord, on=['日期','store_name','channel'], how='left')
        daily['delivery_order_cnt'] = daily['delivery_order_cnt'].fillna(0).astype(int)
        daily['avg_delivery_cost'] = daily.apply(
            lambda r: round(r['delivery_fee']/r['delivery_order_cnt'],2) if r['delivery_order_cnt']>0 else 0, axis=1)

        daily['neg_pct'] = (daily['neg_cnt']/daily['order_cnt']*100).round(1)
        daily['gross_margin_rate'] = (daily['gross_profit']/daily['revenue']*100).round(2)
        daily['qn_store_id'] = daily['store_name'].map(name_to_id)
        daily['store_name'] = daily['store_name'].apply(norm_name)

        all_daily.append(daily)
        nd = daily['日期'].nunique()
        print(f'  {os.path.basename(fp)}: {len(daily)} rows, {nd}d, {daily["order_cnt"].sum()} orders, rev ¥{daily["revenue"].sum():.0f}')

    if not all_daily: return None
    result = pd.concat(all_daily, ignore_index=True)
    result = result.drop_duplicates(subset=['日期','store_name','channel'], keep='last')
    return result

# =============================================
# 3a. Process 盈亏分析 (priority data) first
# =============================================
print('\n=== Profit Analysis (T-2 accurate) ===')
cal_data = None
if files['profit_analysis']:
    for fp in files['profit_analysis']:
        df = pd.read_excel(fp, header=1)
        # Normalize column names (strip spaces from Unnamed cols)
        df.columns = [str(c).strip() for c in df.columns]
        # Find unnamed columns by position
        cols = list(df.columns)
        unnamed_map = {}
        for i, c in enumerate(cols):
            if i == 0 and ('Unnamed' in c or '订单' not in c and '渠道' not in c):
                unnamed_map[c] = '订单号'
            elif i == 1 and ('Unnamed' in c or len(str(c)) < 3):
                unnamed_map[c] = '渠道名称'
            elif i == 2 and ('Unnamed' in c or '时间' in str(c) or len(str(c)) < 3):
                unnamed_map[c] = 'order_time'
            elif i == 3 and ('Unnamed' in c or '门店' in str(c)):
                unnamed_map[c] = '门店'
            elif i == 4 and ('Unnamed' in c or 'ID' in str(c).upper() or str(c).isdigit()):
                unnamed_map[c] = 'qn_store_id_raw'
        df = df.rename(columns=unnamed_map)
        # 分开有效/无效订单
        df['_invalid'] = df['线上毛利率'] == '-'
        df_invalid_cal = df[df['_invalid']].copy()
        df = df[~df['_invalid']].copy()
        
        df['订单号'] = df['订单号'].astype(str)
        df['order_date'] = pd.to_datetime(df['order_time']).dt.date
        df['门店'] = df['门店'].astype(str).apply(norm_name)

        name_to_id = mapping.drop_duplicates(subset='qn_store_name').set_index('qn_store_name')['qn_store_id'].to_dict()
        cal_name_map = {}
        for _, row in df.iterrows():
            qid = row['qn_store_id_raw']
            if pd.isna(qid): continue
            cal_name_map[int(qid)] = row['门店']
        for nm, qid in name_to_id.items():
            if qid not in cal_name_map: cal_name_map[qid] = norm_name(nm)
        df['qn_store_id'] = df['qn_store_id_raw'].apply(lambda x: int(x) if not pd.isna(x) else None)
        df['门店'] = df.apply(lambda r: cal_name_map.get(r['qn_store_id'], r['门店']), axis=1)

        groups = df.groupby(['order_date','门店','qn_store_id','渠道名称'])
        daily = groups.agg(
            order_cnt=('订单号','count'), revenue=('实收交易额','sum'),
            gross_profit=('线上毛利','sum'), income=('收入','sum'),
            cost=('成本','sum'), goods_cost=('商品成本','sum'),
            delivery_fee=('三方配送费','sum'), commission=('佣金','sum'),
            delivery_income=('配送收入','sum'),
        ).reset_index()
        daily.columns = ['日期','store_name','qn_store_id','channel','order_cnt','revenue',
                         'gross_profit','income','cost','goods_cost','delivery_fee','commission','delivery_income']

        # 无效订单：不计实收/单量/配送费，只统计毛利
        if len(df_invalid_cal) > 0:
            df_invalid_cal['order_date'] = pd.to_datetime(df_invalid_cal['order_time']).dt.date
            df_invalid_cal['门店'] = df_invalid_cal['门店'].astype(str).apply(norm_name)
            df_invalid_cal['qn_store_id_raw_num'] = pd.to_numeric(df_invalid_cal['qn_store_id_raw'], errors='coerce')
            df_invalid_cal['qn_store_id'] = df_invalid_cal['qn_store_id_raw_num'].apply(lambda x: int(x) if not pd.isna(x) else None)
            df_invalid_cal['门店'] = df_invalid_cal.apply(lambda r: cal_name_map.get(r['qn_store_id'], r['门店']), axis=1)
            inv_grp = df_invalid_cal.groupby(['order_date','门店','qn_store_id','渠道名称'])
            inv_agg = inv_grp.agg(invalid_profit=('线上毛利','sum')).reset_index()
            inv_agg.columns = ['日期','store_name','qn_store_id','channel','invalid_profit']
            daily = daily.merge(inv_agg, on=['日期','store_name','qn_store_id','channel'], how='left')
            daily['invalid_profit'] = daily['invalid_profit'].fillna(0)
            daily['gross_profit'] = daily['gross_profit'] + daily['invalid_profit']
            daily = daily.drop(columns=['invalid_profit'])

        neg = df[df['线上毛利']<0].groupby(['order_date','门店','qn_store_id','渠道名称']).agg(
            neg_cnt=('订单号','count')).reset_index()
        neg.columns = ['日期','store_name','qn_store_id','channel','neg_cnt']
        daily = daily.merge(neg, on=['日期','store_name','qn_store_id','channel'], how='left')
        daily['neg_cnt'] = daily['neg_cnt'].fillna(0).astype(int)

        del_ord = df[df['三方配送费']>0].groupby(['order_date','门店','qn_store_id','渠道名称']).agg(
            delivery_order_cnt=('订单号','count')).reset_index()
        del_ord.columns = ['日期','store_name','qn_store_id','channel','delivery_order_cnt']
        daily = daily.merge(del_ord, on=['日期','store_name','qn_store_id','channel'], how='left')
        daily['delivery_order_cnt'] = daily['delivery_order_cnt'].fillna(0).astype(int)
        daily['avg_delivery_cost'] = daily.apply(
            lambda r: round(r['delivery_fee']/r['delivery_order_cnt'],2) if r['delivery_order_cnt']>0 else 0, axis=1)
        daily['neg_pct'] = (daily['neg_cnt']/daily['order_cnt']*100).round(1)
        daily['gross_margin_rate'] = (daily['gross_profit']/daily['revenue']*100).round(2)
        daily['store_name'] = daily['store_name'].apply(norm_name)

        cal_data = pd.concat([cal_data, daily]) if cal_data is not None else daily
        print(f'  {os.path.basename(fp)}: {len(daily)} rows, {daily["日期"].nunique()}d, {daily["order_cnt"].sum()} orders, rev ¥{daily["revenue"].sum():.0f}')

# =============================================
# 3b. Process 实时订单明细
# =============================================
print('\n=== Real-time Orders ===')
rt_data = process_orders(files['order_detail'], revenue_col='预计收入')

# =============================================
# 3c. Merge: 盈亏优先覆盖同日，其余保留实时
# =============================================
if cal_data is not None:
    cal_dates = set(cal_data['日期'].unique())
    if rt_data is not None:
        rt_data = rt_data[~rt_data['日期'].isin(cal_dates)]
    daily = pd.concat([d for d in [rt_data, cal_data] if d is not None], ignore_index=True)
    print(f'Merged: {len(daily)} rows ({cal_data["日期"].nunique()}d cal + {0 if rt_data is None else rt_data["日期"].nunique()}d rt)')
else:
    daily = rt_data

if daily is None:
    # No new order data; load existing warehouse
    wh_path = os.path.join(WAREHOUSE, 'daily_store_channel_profit.xlsx')
    if os.path.exists(wh_path):
        daily = pd.read_excel(wh_path)
        print(f'\nNo new orders; loaded {len(daily)} existing warehouse rows')
    else:
        print('\nERROR: No order data at all!')
        exit(1)

# =============================================
# 3d. 丰派线下渠道
# =============================================
fp_files = glob.glob(os.path.join(DATA, '**/商品销售流水信息*.xlsx'), recursive=True)
if fp_files:
    # 预加载补货参考表（只读一次）
    cost_map, upc_to_cost = {}, {}
    for month_dir in sorted(os.listdir(DATA), reverse=True):
        if not month_dir.isdigit() or len(month_dir) != 6: continue
        buo_path = os.path.join(DATA, month_dir, f'补货参考_{month_dir}.xlsx')
        if os.path.exists(buo_path):
            buo = pd.read_excel(buo_path, header=1)
            buo_sku = buo.iloc[:,4].astype(str).str.strip()
            buo_cost = pd.to_numeric(buo.iloc[:,7], errors='coerce').fillna(0)
            buo_upc = buo.iloc[:,19].astype(str).str.strip()
            for sku, c in zip(buo_sku, buo_cost):
                sku = str(sku).strip()
                if sku and sku != 'nan' and sku not in cost_map:
                    cost_map[sku] = float(c)
            for upc_str, sku in zip(buo_upc, buo_sku):
                sku = str(sku).strip()
                if not sku or sku == 'nan': continue
                upc_str = str(upc_str).strip()
                if not upc_str or upc_str == 'nan': continue
                if sku in cost_map:
                    for upc in [u.strip() for u in upc_str.split(',') if u.strip()]:
                        if upc not in upc_to_cost:
                            upc_to_cost[upc] = cost_map[sku]
            print(f'  [丰派] 补货参考: {os.path.basename(buo_path)}, SKU:{len(cost_map)}, UPC:{len(upc_to_cost)}')
            break

    online_names = set(daily[daily['channel'].isin(['美团闪购','饿了么','京东到家'])]['store_name'].unique())
    all_fp = []
    for fp_file in fp_files:
        fp_daily = process_fengpai(fp_file, online_names, cost_map, upc_to_cost)
        if fp_daily is not None:
            all_fp.append(fp_daily)
    if all_fp:
        fp_daily = pd.concat(all_fp, ignore_index=True)
        # 去重（同日同店取last）
        fp_daily = fp_daily.drop_duplicates(subset=['日期','store_name'], keep='last')
        # 补充qn_store_id
        name_id = daily[['store_name','qn_store_id']].dropna(subset=['qn_store_id']).drop_duplicates()
        name_id_dict = dict(zip(name_id['store_name'], name_id['qn_store_id']))
        fp_daily['qn_store_id'] = fp_daily['store_name'].map(name_id_dict)
        daily = pd.concat([daily, fp_daily], ignore_index=True)
        print(f'[丰派] Added {len(fp_daily)} rows to daily from {len(fp_files)} files')

# =============================================
# 4. Join promotions, commission, margins
# =============================================
daily['promo_fee'] = daily.apply(
    lambda r: promo_lookup.get((str(r['日期']), r['store_name'], r['channel']), 0), axis=1)
daily['real_profit'] = (daily['gross_profit'] - daily['promo_fee'] - daily['commission_fee']).round(2)
daily['real_margin_rate'] = np.where(daily['revenue']>0, (daily['real_profit']/daily['revenue']*100).round(2), 0)
daily['commission_rate'] = daily['qn_store_id'].map(comm_lookup).fillna(0)
daily['commission_fee'] = (daily['revenue'] * daily['commission_rate']).round(2)
# 线下渠道（客无忧POS + 丰派）不抽佣
offline_mask = daily['channel'].isin(['客无忧POS', '线下'])
daily.loc[offline_mask, 'commission_rate'] = 0
daily.loc[offline_mask, 'commission_fee'] = 0
# 客无忧POS 改名为 线下
daily.loc[daily['channel'] == '客无忧POS', 'channel'] = '线下'
# 门店毛利 = 线上毛利 - 推广费
daily['store_profit'] = (daily['gross_profit'] - daily['promo_fee']).round(2)
# 抽佣毛利 = 门店毛利 - 平台抽佣
daily['commission_profit'] = (daily['store_profit'] - daily['commission_fee']).round(2)
daily['commission_margin'] = np.where(daily['revenue']>0, (daily['commission_profit']/daily['revenue']*100).round(2), 0)

print(f'\n=== Final Daily ===')
print(f'Orders: {daily["order_cnt"].sum()}, Revenue: ¥{daily["revenue"].sum():.0f}')
print(f'Gross profit: ¥{daily["gross_profit"].sum():.0f}, Promo: ¥{daily["promo_fee"].sum():.0f}')
print(f'Real profit: ¥{daily["real_profit"].sum():.0f}, Commission: ¥{daily["commission_fee"].sum():.0f}')

# =============================================
# 5. Delivery fee update from finance files (when available)
# =============================================
if files['mt_finance'] and daily is not None:
    print('\n=== Delivery Fee Update ===')
    # Build delivery lookup from all finance files
    mt_del_lookup = {}
    for fp in files['mt_finance']:
        mf = pd.read_excel(fp, header=1)
        if '交易类型' not in mf.columns or '商家应收款（结算金额）' not in mf.columns: continue
        mf_fee = mf[mf['交易类型'] == '配送费用']
        for _, r in mf_fee.iterrows():
            oid = str(r['订单号']).strip()
            fee = abs(r['商家应收款（结算金额）'])
            if fee > 0: mt_del_lookup[oid] = fee
        print(f'  {os.path.basename(fp)}: {len(mf_fee)} delivery rows, {len(mf_fee["订单号"].unique())} unique orders')
    print(f'  Total delivery lookup: {len(mt_del_lookup)} orders')
    # Note: For calibrated data (盈亏分析), delivery fees are already accurate T-2 data.
    # This update only matters when real-time orders have missing delivery fees.
    print('  (Delivery fees already included in order data; finance files archived for reference)')

# =============================================
# 6. Save warehouse (merge with existing)
# =============================================
wh_path = os.path.join(WAREHOUSE, 'daily_store_channel_profit.xlsx')
daily['日期'] = pd.to_datetime(daily['日期'])
if os.path.exists(wh_path):
    existing = pd.read_excel(wh_path)
    existing['日期'] = pd.to_datetime(existing['日期'])
    kept = existing[~existing['日期'].isin(daily['日期'].unique())]
    merged = pd.concat([kept, daily], ignore_index=True)
    merged = merged.sort_values(['日期','store_name','channel']).reset_index(drop=True)
else:
    merged = daily
# 去重：同日期+门店+渠道只保留最后一条
n_before = len(merged)
merged = merged.drop_duplicates(subset=['日期','store_name','channel'], keep='last')
if len(merged) < n_before:
    print(f'  去重: {n_before} -> {len(merged)}行 (移除{n_before-len(merged)}行)')
merged.to_excel(wh_path, index=False)
print(f'\nWarehouse: {wh_path} ({len(merged)} rows, {len(daily)} new)')

print('\n=== ETL Complete ===')
