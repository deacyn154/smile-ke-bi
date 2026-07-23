# -*- coding: utf-8 -*-
"""
Generate self-contained HTML BI dashboard v3:
  - Data source: MySQL smile_ke_bi (fallback: Excel warehouse)
  - Calendar date picker + quick buttons
  - Store search + chips (short names)
"""
import pandas as pd
import json, os, re, calendar
from sqlalchemy import create_engine, text

BASE = r'E:\Desktop\工作文件（月度）\claw制作BI'
WAREHOUSE = os.path.join(BASE, 'warehouse')
OUTPUT = os.path.join(BASE, 'dashboard.html')

# --- Data Source: MySQL (primary) or Excel (fallback) ---
MYSQL_DSN = 'mysql+pymysql://root:@localhost:3306/smile_ke_bi?charset=utf8mb4'
USE_MYSQL = False
try:
    engine = create_engine(MYSQL_DSN)
    with engine.connect() as conn:
        conn.execute(text('SELECT 1'))
    USE_MYSQL = True
    print('Data source: MySQL')
except Exception as e:
    print(f'MySQL unavailable ({e}), fallback to Excel')

if USE_MYSQL:
    query = """
        SELECT dt AS `日期`, store_name, qn_store_id, channel,
               order_cnt, revenue, real_profit, commission_fee, commission_profit,
               neg_cnt, delivery_fee, delivery_order_cnt, promo_fee, store_profit
        FROM daily_profit
    """
    df = pd.read_sql(query, engine)
    df['日期'] = df['日期'].astype(str)
    promo = pd.read_sql("SELECT dt AS `日期`, store_name, qn_store_id, channel, promo_fee FROM daily_promo", engine)
    promo['日期'] = promo['日期'].astype(str)
else:
    df = pd.read_excel(os.path.join(WAREHOUSE, 'daily_store_channel_profit.xlsx'))
    df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
    promo = None
    promo_path = os.path.join(WAREHOUSE, 'promo_daily.xlsx')
    if os.path.exists(promo_path):
        promo = pd.read_excel(promo_path)
        promo['日期'] = pd.to_datetime(promo['日期']).dt.strftime('%Y-%m-%d')

# 仪表盘只取最近365天
df['_date_sort'] = pd.to_datetime(df['日期'])
max_date = df['_date_sort'].max()
cutoff = max_date - pd.Timedelta(days=365)
df = df[df['_date_sort'] >= cutoff].drop(columns=['_date_sort'])
if promo is not None and len(promo) > 0:
    promo['_date_sort'] = pd.to_datetime(promo['日期'])
    promo = promo[promo['_date_sort'] >= cutoff].drop(columns=['_date_sort'])

def clean(records):
    out = []
    for r in records:
        row = {}
        for k, v in r.items():
            if pd.isna(v): row[k] = 0
            else:
                try:
                    fv = float(v)
                    if fv == int(fv): row[k] = int(fv)
                    else: row[k] = round(fv, 2)
                except: row[k] = v
        out.append(row)
    return out

# 数据压缩: 数组格式 (无key名), 大幅减小体积
ch_index = {'美团闪购':0, '饿了么':1, '京东到家':2, '线下':3}
raw = df[['日期','qn_store_id','channel','order_cnt','revenue','store_profit',
           'commission_profit','neg_cnt','delivery_fee','delivery_order_cnt']]
data_compact = []
for _, r in raw.iterrows():
    qid = int(r['qn_store_id']) if pd.notna(r['qn_store_id']) else 0
    d = str(r['日期'])[:10].replace('-','')
    data_compact.append([
        d, qid,
        ch_index.get(r['channel'], 3),
        int(r['order_cnt'] or 0),
        round(float(r['revenue'] or 0)),
        round(float(r['store_profit'] or 0)),
        round(float(r['commission_profit'] or 0)),
        int(r['neg_cnt'] or 0),
        round(float(r['delivery_fee'] or 0)),
        int(r['delivery_order_cnt'] or 0),
    ])
data_json = json.dumps(data_compact, ensure_ascii=False)
# 额外输出 data.json 给 GitHub Pages 在线加载
data_json_path = os.path.join(BASE, 'data.json')
with open(data_json_path, 'w', encoding='utf-8') as f:
    f.write(data_json)

# promo 保持原样（小数据）
promo_records = clean(promo.to_dict(orient='records')) if promo is not None else []
promo_json = json.dumps(promo_records, ensure_ascii=False)

# 加载 channel_store_mapping (后续门店标准化需要)
mapping_path = os.path.join(BASE, 'channel_store_mapping.xlsx')
if os.path.exists(mapping_path):
    mapping = pd.read_excel(mapping_path)
    mapping['channel_store_id'] = mapping['channel_store_id'].astype(str)
    mapping['qn_store_id'] = pd.to_numeric(mapping['qn_store_id'], errors='coerce').fillna(0).astype(int)
else:
    mapping = pd.DataFrame(columns=['qn_store_id','qn_store_name','channel','channel_store_id'])

# 门店标准化: 以 qn_store_id 为唯一键, 名称只来自 mapping
store_id_to_name = {}
for _, r in mapping.iterrows():
    qid = str(int(r['qn_store_id']))
    name = r['qn_store_name']
    if qid not in store_id_to_name:
        store_id_to_name[qid] = name

# promo 门店名替换
for r in promo_records:
    qid = str(r.get('qn_store_id', '')).strip() if r.get('qn_store_id') is not None else ''
    if qid in store_id_to_name:
        r['store_name'] = store_id_to_name[qid]

# allStoresFull: 从 df 中提取(去重 qn_store_id)
all_store_ids = sorted(set(int(r['qn_store_id']) for _, r in df.iterrows() if pd.notna(r['qn_store_id'])))
stores_full = [str(qid) for qid in all_store_ids]

# --- Short store name extraction ---
def short_name(full):
    m = re.search(r'[（(]([^）)]+)[）)]', full)
    return m.group(1) if m else full

# allStoresFull: 用 qn_store_id 列表(去重)
all_store_ids = sorted(set(int(r['qn_store_id']) for _, r in df.iterrows() if pd.notna(r['qn_store_id'])))
stores_full = [str(int(qid)) for qid in all_store_ids]
store_name_map = {qid: store_id_to_name.get(qid, f'门店{qid}') for qid in stores_full}
short_map = {qid: short_name(name) for qid, name in store_name_map.items()}
search_map = {}
for qid in stores_full:
    name = store_name_map[qid]
    sn = short_name(name)
    sid = name.split('-')[0] if '-' in name else qid
    search_map[qid] = json.dumps([sn, sid, name, qid], ensure_ascii=False)
stores_json = json.dumps(stores_full, ensure_ascii=False)
store_name_map_json = json.dumps(store_name_map, ensure_ascii=False)
search_map_json = json.dumps(search_map, ensure_ascii=False)

# 门店基本信息（省市映射）
store_info_path = os.path.join(BASE, '数据表', '基础信息表', '门店基本信息.xlsx')
store_region_map = {}
if os.path.exists(store_info_path):
    try:
        info_df = pd.read_excel(store_info_path)
        for _, row in info_df.iterrows():
            name = str(row.get('门店名称', '')).strip()
            region = str(row.get('所属地区', '')).strip()
            if name and region:
                store_region_map[name] = region
        print(f'Store region map: {len(store_region_map)} stores loaded')
    except Exception as e:
        print(f'Store info load warning: {e}')
store_region_json = json.dumps(store_region_map, ensure_ascii=False)

# 门店运营成本
cost_path = os.path.join(BASE, '数据表', '基础信息表', '门店运营成本.xlsx')
store_cost_map = {}
if os.path.exists(cost_path):
    try:
        cost_df = pd.read_excel(cost_path)
        cost_list = []
        for _, row in cost_df.iterrows():
            sname = str(row.get('门店', '')).strip()
            cost = float(row.get('门店运营成本', 0) or 0)
            if sname and cost > 0:
                cost_list.append({'name': sname, 'cost': cost})
        # 精确匹配到完整门店名
        for s in stores_full:
            sn = short_name(s)
            matched = None
            # 先精确匹配省市区名（取门店名核心部分）
            for c in cost_list:
                # 检查cost名是否完整出现在store名中
                if c['name'] in s: matched = c; break
                if c['name'] in sn: matched = c; break
            if not matched:
                # 模糊匹配：至少4个连续汉字，避免短词误匹配
                for c in cost_list:
                    for i in range(len(c['name'])-3):
                        frag = c['name'][i:i+4]
                        if frag in s or frag in sn: matched = c; break
                    if matched: break
            if matched:
                store_cost_map[s] = {'cost': matched['cost'], 'name': matched['name']}
        print(f'Store cost map: {len(store_cost_map)} stores matched out of {len(cost_list)} costs')
    except Exception as e:
        print(f'Store cost load warning: {e}')
store_cost_json = json.dumps(store_cost_map, ensure_ascii=False)

# 门店→美团渠道门店ID 映射（导出Excel用）
mapping_path = os.path.join(BASE, 'channel_store_mapping.xlsx')
store_meituan_map = {}
if os.path.exists(mapping_path):
    try:
        mapping_df = pd.read_excel(mapping_path)
        for _, row in mapping_df.iterrows():
            if str(row.get('channel', '')) == '美团闪购':
                qn_name = str(row.get('qn_store_name', '')).strip()
                mt_id = str(row.get('channel_store_id', '')).strip()
                if qn_name and mt_id and mt_id not in ('nan', 'None', ''):
                    store_meituan_map[qn_name] = mt_id
        print(f'Meituan store map: {len(store_meituan_map)} stores loaded')
    except Exception as e:
        print(f'Meituan store map load warning: {e}')
store_meituan_json = json.dumps(store_meituan_map, ensure_ascii=False)

# 当前月绩效进度数据
perf_data = []
try:
    fp_duty = os.path.join(BASE, '数据表', '基础信息表', '门店分工.xlsx')
    if os.path.exists(fp_duty):
        duty = pd.read_excel(fp_duty, sheet_name='门店').dropna(subset=['负责人'])
        duty['牵牛花id'] = duty['牵牛花id'].astype(int).astype(str)
        targets = pd.read_excel(fp_duty, sheet_name='目标', header=1).dropna(subset=['姓名'])
        target_map = {}
        for _, r in targets.iterrows():
            target_map[r['姓名']] = {'orders': int(r['订单量']), 'profit': float(r['门店毛利（去推广不去抽佣）'])}
        # 只取最新月份数据（绩效按月考核）
        perf_month = max_date.month
        perf_df = df[pd.to_datetime(df['日期']).dt.month == perf_month].copy()
        # 当月新店不参与绩效（7月无新开/更名店，空集合）
        NEW_STORE_IDS = set()
        perf_df = perf_df[~perf_df['qn_store_id'].isin(NEW_STORE_IDS)]
        perf_df['qn_sid'] = perf_df['qn_store_id'].apply(lambda x: str(int(float(x))) if pd.notna(x) else '')
        perf_days = perf_df['日期'].nunique()
        # Build per-(channel, store) owner map to avoid collision across channels
        owner_map = {}
        for _, r in duty.iterrows():
            owner_map[(str(r['牵牛花id']), str(r['渠道']))] = str(r['负责人'])
        EXCLUDE_QN_SID = '1150787'
        for ch in ['美团闪购', '饿了么', '京东到家']:
            ch_df = perf_df[perf_df['channel'] == ch].copy()
            ch_df = ch_df[ch_df['qn_sid'] != EXCLUDE_QN_SID]
            ch_df['负责人'] = ch_df['qn_sid'].apply(lambda x: owner_map.get((x, ch), '未分配'))
            ch_df = ch_df[ch_df['负责人'] != '未分配']
            grp = ch_df.groupby('负责人').agg(o=('order_cnt','sum'),p=('store_profit','sum')).round(2).reset_index()
            for _, r in grp.iterrows():
                name = r['负责人']
                t = target_map.get(name, {})
                perf_data.append({
                    'name': name, 'orders': int(r['o']), 'profit': round(float(r['p']), 2),
                    'target_orders': t.get('orders',0), 'target_profit': t.get('profit',0)
                })
        # dedup same name
        seen = {}
        perf_data = [seen.setdefault(p['name'], p) for p in perf_data if p['name'] not in seen]
except Exception as e:
    print(f'Perf data warning: {e}')
perf_json = json.dumps({
    'month': max_date.month, 'year': max_date.year,
    'days': perf_days if 'perf_days' in dir() else 0,
    'total_days': calendar.monthrange(max_date.year, max_date.month)[1],
    'data': perf_data,
    'summary': {'total_orders': sum(p['orders'] for p in perf_data),
                'total_target_orders': sum(p['target_orders'] for p in perf_data),
                'total_profit': round(sum(p['profit'] for p in perf_data), 2),
                'total_target_profit': round(sum(p['target_profit'] for p in perf_data), 2)}
}, ensure_ascii=False)

channels = sorted(df['channel'].unique().tolist())
dates_all = sorted(df['日期'].unique().tolist())

short_map_json = json.dumps(short_map, ensure_ascii=False)

channels_json = json.dumps(channels, ensure_ascii=False)
dates_json = json.dumps(dates_all, ensure_ascii=False)

# --- Product data ---
import pandas as pd2
prod_path = os.path.join(WAREHOUSE, 'product', 'product_daily.xlsx')
product_json = '[]'
if os.path.exists(prod_path):
    product_df = pd2.read_excel(prod_path)
    product_df = product_df.fillna(0)
    # Use latest period as default and embed all periods in data
    product_df['period'] = product_df['period'].astype(str)
    periods = sorted(product_df['period'].unique())
    latest_period = periods[-1] if len(periods) > 0 else '202605'
    all_product_json = json.dumps(product_df.to_dict(orient='records'), ensure_ascii=False)
    product_json_all = json.dumps(periods, ensure_ascii=False)
    # Filter to latest period for dashboard display
    product_df_display = product_df[product_df['period'] == latest_period].copy()
    product_records = product_df_display.to_dict(orient='records')
    product_json = json.dumps(product_records, ensure_ascii=False)
    print(f'Product data: {len(product_records)} rows loaded (period={latest_period}, total={len(product_df)}, periods={periods})')

# --- Raw category data: combine all period files ---
rawcat_json = '[]'
rawcat_combined_json = '[]'
all_rawcat_turnover = []
all_rawcat_combined = []
for p in periods:
    rtp = os.path.join(WAREHOUSE, 'product', f'rawcat_turnover_{p}.xlsx')
    if os.path.exists(rtp):
        all_rawcat_turnover.append(pd2.read_excel(rtp).fillna(0))
    rcp_path = os.path.join(WAREHOUSE, 'product', 'rawcat_combined.xlsx')
# Also try legacy single file
if os.path.exists(os.path.join(WAREHOUSE, 'product', 'rawcat_turnover.xlsx')):
    all_rawcat_turnover.append(pd2.read_excel(os.path.join(WAREHOUSE, 'product', 'rawcat_turnover.xlsx')).fillna(0))
if all_rawcat_turnover:
    rawcat_df = pd.concat(all_rawcat_turnover, ignore_index=True)
    # Filter to latest period
    if 'period' in rawcat_df.columns:
        rawcat_df = rawcat_df[rawcat_df['period'].astype(str) == latest_period]
    rawcat_records = rawcat_df.to_dict(orient='records')
    rawcat_json = json.dumps(rawcat_records, ensure_ascii=False)
    print(f'Raw cat data: {len(rawcat_records)} rows loaded')
if os.path.exists(rawcat_combined_path := os.path.join(WAREHOUSE, 'product', 'rawcat_combined.xlsx')):
    rawcat_combined_df = pd2.read_excel(rawcat_combined_path)
    rawcat_combined_df = rawcat_combined_df.fillna(0)
    if 'period' in rawcat_combined_df.columns:
        rawcat_combined_df = rawcat_combined_df[rawcat_combined_df['period'].astype(str) == latest_period]
    rawcat_combined_records = rawcat_combined_df.to_dict(orient='records')
    rawcat_combined_json = json.dumps(rawcat_combined_records, ensure_ascii=False)
    print(f'Raw cat combined: {len(rawcat_combined_records)} rows loaded')

# =============================================
# 预渲染静态chip HTML（不依赖JS创建）
# =============================================
store_chips_html = ''.join(
    f'<span class="chip active" data-full="{qid}" title="{store_name_map.get(qid, qid)}">{short_map.get(qid, store_name_map.get(qid, qid))}</span>'
    for qid in stores_full
)
channel_chips_html = ''.join(
    f'<span class="chip active" data-full="{c}">{c}</span>'
    for c in channels
)

# =============================================
# HTML template
# =============================================
html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>微笑客经营看板</title>
<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist@2.32.0/plotly.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<style>
:root {
    --bg: #f5f6f8; --card: #ffffff; --border: #e5e7eb;
    --text: #1f2937; --text-dim: #9ca3af;
    --accent: #4f6ef7; --green: #10b981; --yellow: #f59e0b; --red: #ef4444; --blue: #4f6ef7;
    --orange: #f97316;
    --meituan: #FFD100; --eleme: #00AAFF; --jd: #E2231A; --pos: #4f6ef7;
    --card-hover: #f9fafb;
}
/* Scrollbar */
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:var(--bg); }
::-webkit-scrollbar-thumb { background:var(--border); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:var(--text-dim); }
* { margin:0; padding:0; box-sizing:border-box; -webkit-tap-highlight-color:transparent; }
body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif; background:var(--bg); color:var(--text); min-height:100vh; overflow-x:hidden; }
.dashboard { max-width:1440px; margin:0 auto; padding:24px; }
.header { display:flex; align-items:center; justify-content:space-between; margin-bottom:20px; padding-bottom:12px; border-bottom:1px solid var(--border); }
.header h1 { font-size:22px; font-weight:700; background:linear-gradient(135deg,var(--blue),var(--accent)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
.update-time { color:var(--text-dim); font-size:12px; }
/* === Filter Bar === */
.filter-bar { background:var(--card); border:1px solid var(--border); border-radius:12px; padding:12px 16px; margin-bottom:20px; }
.filter-row { display:flex; flex-wrap:wrap; align-items:center; gap:6px; margin-bottom:8px; }
.filter-row:last-child { margin-bottom:0; }
.filter-label { color:var(--text-dim); font-size:12px; font-weight:500; white-space:nowrap; min-width:36px; }
/* Quick date buttons */
.date-quick { display:flex; gap:4px; flex-wrap:wrap; }
.date-btn { padding:4px 10px; border-radius:14px; font-size:11px; font-weight:500; cursor:pointer; border:1px solid var(--border); background:var(--bg); color:var(--text-dim); transition:all .15s; touch-action:manipulation; }
.date-btn.active { background:var(--accent); border-color:var(--accent); color:#fff; }
.date-btn:hover:not(.active) { border-color:var(--accent); color:var(--text); }
/* Calendar */
.calendar-wrap { margin-top:8px; border-top:1px solid var(--border); padding-top:8px; display:none; }
.calendar-wrap.show { display:block; }
.cal-header { display:flex; align-items:center; justify-content:space-between; margin-bottom:6px; }
.cal-header .cal-nav { cursor:pointer; color:var(--text-dim); font-size:13px; padding:4px 8px; border-radius:6px; background:var(--bg); border:1px solid var(--border); user-select:none; }
.cal-header .cal-nav:hover { border-color:var(--accent); color:var(--text); }
.cal-header .cal-title { font-size:14px; font-weight:600; }
.cal-grid { display:grid; grid-template-columns:repeat(7,1fr); gap:2px; }
.cal-grid .cal-weekday { text-align:center; font-size:10px; color:var(--text-dim); padding:4px 0; }
.cal-grid .cal-day { text-align:center; font-size:12px; padding:4px 0; border-radius:6px; cursor:pointer; color:var(--text-dim); transition:all .1s; }
.cal-grid .cal-day:hover { background:rgba(79,110,247,0.1); }
.cal-grid .cal-day.in-range { background:rgba(79,110,247,0.15); color:var(--text); border-radius:0; }
.cal-grid .cal-day.range-start { background:var(--accent); color:#fff; border-radius:6px 0 0 6px; }
.cal-grid .cal-day.range-end { background:var(--accent); color:#fff; border-radius:0 6px 6px 0; }
.cal-grid .cal-day.range-start.range-end { border-radius:6px; }
.cal-grid .cal-day.other-month { color:rgba(139,143,163,0.3); }
.cal-grid .cal-day.has-data { color:#1f2937; }
.cal-grid .cal-day.no-data { color:rgba(139,143,163,0.2); cursor:default; pointer-events:none; }
.cal-toggle { color:var(--accent); font-size:12px; cursor:pointer; padding:3px 8px; border-radius:10px; border:1px solid var(--border); background:transparent; touch-action:manipulation; }
.cal-toggle:hover { background:rgba(79,110,247,0.08); }
/* Store search */
.store-search { width:100%; padding:6px 10px; border-radius:16px; border:1px solid var(--border); background:var(--bg); color:var(--text); font-size:12px; outline:none; margin-bottom:6px; }
.store-search:focus { border-color:var(--accent); }
.store-search::placeholder { color:var(--text-dim); }
.chip-group { display:flex; flex-wrap:wrap; gap:4px; flex:1; }
.chip { display:inline-flex; align-items:center; gap:3px; padding:4px 10px; border-radius:16px; font-size:11px; font-weight:500; cursor:pointer; border:1px solid var(--border); background:var(--bg); color:var(--text-dim); transition:all .12s; user-select:none; touch-action:manipulation; }
.chip.hidden { display:none; }
.chip:hover { border-color:var(--accent); color:var(--text); }
.chip.active { background:var(--accent); border-color:var(--accent); color:#fff; }
.chip-actions { display:flex; gap:4px; margin-left:auto; }
.chip-action { padding:4px 10px; border-radius:14px; font-size:11px; font-weight:500; cursor:pointer; border:1px solid var(--border); background:transparent; color:var(--text-dim); transition:all .12s; touch-action:manipulation; }
.chip-action:hover { border-color:var(--accent); color:var(--text); }
/* KPI */
.kpi-grid { display:grid; gap:12px; margin-bottom:20px; }
.kpi-card { background:var(--card); border:1px solid var(--border); border-radius:12px; padding:18px; position:relative; overflow:hidden; transition:all .25s; box-shadow:0 1px 3px rgba(0,0,0,0.04); }
.kpi-card:hover { background:var(--card-hover); transform:translateY(-2px); box-shadow:0 4px 16px rgba(0,0,0,0.08); }
.kpi-label { color:#1f2937; font-size:14px; margin-bottom:4px; font-weight:500; }
.kpi-label.important { color:var(--red); }
.kpi-value { font-size:24px; font-weight:700; letter-spacing:-0.3px; color:var(--blue) !important; }
.kpi-sub { font-size:11px; margin-top:4px; }
.kpi-mom { font-size:12px; margin-top:3px; font-weight:500; }
/* 名词解释 hint */
.kpi-hint { font-size:10px; cursor:help; position:relative; color:var(--text-dim); display:inline-block; }
.kpi-hint:hover { color:var(--accent); }
.kpi-tip { display:none; position:absolute; top:0; left:100%; margin-left:6px; background:var(--bg); border:1px solid var(--border); border-radius:6px; padding:4px 10px; font-size:11px; white-space:nowrap; z-index:10; color:var(--text); font-weight:400; box-shadow:0 2px 8px rgba(0,0,0,0.1); }
.kpi-hint:hover .kpi-tip { display:block; }
.accent .kpi-value { color:var(--accent) !important; }
.green .kpi-value { color:var(--green); }
.yellow .kpi-value { color:var(--yellow); }
.red .kpi-value { color:var(--red); }
.blue .kpi-value { color:var(--blue); }
.orange .kpi-value { color:var(--orange); }
/* Tabs */
.tab-nav-wrap { overflow-x:auto; -webkit-overflow-scrolling:touch; margin-bottom:16px; scrollbar-width:none; }
.tab-nav-wrap::-webkit-scrollbar { display:none; }
.tab-nav { display:inline-flex; gap:4px; background:var(--card); border:1px solid var(--border); border-radius:10px; padding:3px; min-width:100%; }
.tab-btn { flex-shrink:0; padding:8px 16px; border:none; background:transparent; color:var(--text-dim); font-size:13px; font-weight:500; border-radius:7px; cursor:pointer; transition:all .15s; white-space:nowrap; touch-action:manipulation; }
.tab-btn.active { background:var(--accent); color:#fff; }
.tab-content { display:none; }
.tab-content.active { display:block; }
.chart-section { background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; margin-bottom:12px; }
.chart-section h3 { font-size:13px; margin-bottom:12px; color:var(--text-dim); }
.chart-section [id^="chart"] { width:100% !important; max-width:100%; }
.chart-section .js-plotly-plot { width:100% !important; }
.row { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.table-scroll { overflow-x:auto; -webkit-overflow-scrolling:touch; border-radius:8px; border:1px solid var(--border); }
.table-scroll table { min-width:700px; width:100%; border-collapse:collapse; font-size:12px; }
th { text-align:left; padding:8px 10px; background:var(--bg); color:var(--text-dim); font-weight:600; border-bottom:1px solid var(--border); position:sticky; top:0; white-space:nowrap; }
td { padding:6px 10px; border-bottom:1px solid var(--border); white-space:nowrap; color:#1f2937; }
tr:hover td { background:rgba(79,110,247,0.04); }
.badge { display:inline-block; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:500; }
.badge-warn { background:rgba(239,68,68,0.1); color:var(--red); }
.badge-ok { background:rgba(16,185,129,0.1); color:var(--green); }
/* Mobile */
@media (max-width:768px) {
    .dashboard { padding:10px; }
    .header { flex-direction:column; align-items:flex-start; gap:2px; }
    .header h1 { font-size:18px; }
    .filter-bar { padding:10px; }
    .filter-row { flex-direction:column; align-items:flex-start; }
    .chip-actions { margin-left:0; }
    .kpi-grid { grid-template-columns:repeat(2,1fr); gap:8px; }
    .kpi-card { padding:10px; }
    .kpi-value { font-size:20px; }
    .kpi-label { font-size:11px; }
    .row { grid-template-columns:1fr; gap:10px; }
    .table-scroll table { min-width:580px; font-size:11px; }
    th,td { padding:4px 6px; }
}
@media (max-width:480px) {
    .kpi-value { font-size:16px; }
    .kpi-card { padding:8px; }
}
@media (min-width:1024px) { .kpi-grid { grid-template-columns:repeat(5,1fr); } }
@media (min-width:769px) and (max-width:1023px) { .kpi-grid { grid-template-columns:repeat(3,1fr); } }
.export-btn { padding:4px 12px; border-radius:6px; border:1px solid var(--border); background:var(--bg); color:var(--text-dim); font-size:11px; cursor:pointer; float:right; }
.export-btn:hover { border-color:var(--accent); color:var(--text); }
/* === Smart Summary === */
.summary-panel { background:linear-gradient(135deg, rgba(79,110,247,0.04), rgba(16,185,129,0.03)); border:1px solid rgba(79,110,247,0.15); border-radius:12px; padding:14px 18px; margin-bottom:16px; }
.summary-panel h3 { font-size:13px; color:var(--accent); margin-bottom:8px; display:flex; align-items:center; gap:6px; }
.summary-text { font-size:13px; line-height:1.7; color:var(--text); }
.summary-text .highlight-red { color:var(--red); font-weight:600; }
.summary-text .highlight-green { color:var(--green); font-weight:600; }
.summary-text .highlight-yellow { color:var(--yellow); font-weight:600; }
.summary-text .highlight-blue { color:var(--blue); font-weight:600; }
/* === Alert Banner === */
.alert-banner { display:none; background:rgba(232,100,82,0.08); border:1px solid rgba(232,100,82,0.25); border-radius:10px; padding:10px 14px; margin-bottom:12px; }
.alert-banner.show { display:flex; align-items:center; gap:8px; }
.alert-banner .alert-icon { font-size:16px; }
.alert-banner .alert-msg { font-size:12px; color:var(--red); flex:1; }
.alert-banner .alert-dismiss { cursor:pointer; color:var(--text-dim); font-size:16px; }
/* === KPI Upgrade === */
.kpi-card::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; opacity:0.5; }
.kpi-card.accent::before { background:var(--accent); }
.kpi-card.green::before { background:var(--green); }
.kpi-card.yellow::before { background:var(--yellow); }
.kpi-card.red::before { background:var(--red); }
.kpi-card.blue::before { background:var(--blue); }
.kpi-card.orange::before { background:var(--orange); }
.kpi-sparkline { margin-top:6px; height:28px; }
.kpi-sparkline svg { width:100%; height:100%; }
.kpi-sparkline path { fill:none; stroke-width:1.5; stroke-linecap:round; stroke-linejoin:round; }
.kpi-sparkline .area { opacity:0.15; }
/* === Sortable Tables === */
th.sortable { cursor:pointer; user-select:none; position:relative; padding-right:20px !important; }
th.sortable:hover { color:var(--accent); }
th.sortable::after { content:'↕'; position:absolute; right:6px; opacity:0.3; font-size:11px; }
th.sortable.asc::after { content:'▲'; opacity:1; color:var(--accent); }
th.sortable.desc::after { content:'▼'; opacity:1; color:var(--accent); }
/* Conditional formatting */
.cell-good { color:var(--green) !important; }
.cell-warn { color:var(--yellow) !important; }
.cell-bad { color:var(--red) !important; }
.cell-accent { color:var(--accent); }
/* In-cell bar */
.cell-bar-wrap { display:flex; align-items:center; gap:6px; }
.cell-bar { height:6px; border-radius:3px; min-width:2px; transition:width .3s; }
.cell-bar.green { background:var(--green); }
.cell-bar.blue { background:var(--blue); }
.cell-bar.orange { background:var(--orange); }
.cell-bar.red { background:var(--red); }
/* === Anomaly Badges === */
.anomaly-dot { display:inline-block; width:6px; height:6px; border-radius:50%; margin-right:3px; vertical-align:middle; }
.anomaly-dot.red { background:var(--red); animation:blink 2s infinite; }
.anomaly-dot.yellow { background:var(--yellow); }
@keyframes blink { 0%,100%{opacity:1;} 50%{opacity:0.3;} }
.anomaly-tag { display:inline-block; padding:1px 6px; border-radius:3px; font-size:10px; font-weight:600; margin-left:4px; }
.anomaly-tag.warn { background:rgba(239,68,68,0.1); color:var(--red); }
.anomaly-tag.info { background:rgba(79,110,247,0.1); color:var(--blue); }
/* === Page Nav Upgrade === */
.page-btn { transition:all .2s; }
.page-btn:hover { opacity:0.85; }
/* === Channel Badge === */
.channel-badge { display:inline-block; padding:1px 6px; border-radius:3px; font-size:10px; font-weight:600; }
.channel-badge.meituan { background:rgba(255,209,0,0.15); color:#FFD100; }
.channel-badge.eleme { background:rgba(0,170,255,0.15); color:#00AAFF; }
.channel-badge.jd { background:rgba(226,35,26,0.15); color:#E86452; }
.channel-badge.pos { background:rgba(79,110,247,0.1); color:#4f6ef7; }
/* === Profit Modal === */
.profit-kpi-row { display:grid; grid-template-columns:repeat(5,1fr); gap:10px; margin-bottom:16px; }
.profit-kpi { background:var(--card); border:1px solid var(--border); border-radius:8px; padding:12px 14px; text-align:center; }
.pk-label { font-size:12px; color:var(--text-dim); margin-bottom:4px; }
.pk-value { font-size:22px; font-weight:700; color:#1f2937; }
.pk-sub { font-size:11px; color:var(--text-dim); margin-top:3px; }
#profitModal .table-scroll { max-height:60vh; overflow-y:auto; }
#profitModal th { position:sticky; top:0; z-index:2; }
.profit-grp td { font-weight:600 !important; }
.profit-grp + tr + tr td { background:initial; }
@media (max-width:768px) { .profit-kpi-row { grid-template-columns:repeat(2,1fr); } }
/* === Skeleton Loading === */
@keyframes shimmer { 0%{background-position:-200% 0;} 100%{background-position:200% 0;} }
.skeleton { background:linear-gradient(90deg,var(--card) 25%,var(--card-hover) 50%,var(--card) 75%); background-size:200% 100%; animation:shimmer 1.5s infinite; border-radius:6px; }
/* === Pulse for live data === */
@keyframes pulse { 0%,100%{box-shadow:0 0 0 0 rgba(79,110,247,0.3);} 50%{box-shadow:0 0 0 6px rgba(79,110,247,0);} }
.pulse { animation:pulse 2s infinite; }
/* === Tooltip === */
.tooltip-wrap { position:relative; cursor:help; border-bottom:1px dotted var(--text-dim); }
.tooltip-wrap:hover .tooltip-content { display:block; }
.tooltip-content { display:none; position:absolute; bottom:120%; left:50%; transform:translateX(-50%); background:var(--card); border:1px solid var(--border); border-radius:8px; padding:6px 10px; font-size:11px; white-space:nowrap; z-index:100; box-shadow:0 4px 16px rgba(0,0,0,0.1); }
/* === KPI Clickable === */
.kpi-card.clickable { cursor:pointer; position:relative; }
.kpi-card.clickable:hover { border-color:var(--accent); }
.kpi-card.clickable::after { content:'\\1F50D'; position:absolute; top:8px; right:10px; font-size:11px; opacity:0; transition:opacity .2s; }
.kpi-card.clickable:hover::after { opacity:0.5; }
/* === Modal === */
.modal-overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,0.7); z-index:9998; justify-content:center; align-items:flex-start; padding:20px; overflow-y:auto; }
.modal-overlay.show { display:flex; }
.modal-box { background:var(--card); border:1px solid var(--border); border-radius:16px; width:100%; max-width:820px; max-height:85vh; display:flex; flex-direction:column; animation:slideUp .25s ease; }
@keyframes slideUp { from{opacity:0;transform:translateY(20px);} to{opacity:1;transform:translateY(0);} }
.modal-header { display:flex; align-items:center; justify-content:space-between; padding:16px 20px; border-bottom:1px solid var(--border); flex-shrink:0; }
.modal-header h2 { font-size:16px; }
.modal-close { cursor:pointer; font-size:20px; color:var(--text-dim); width:32px; height:32px; display:flex; align-items:center; justify-content:center; border-radius:8px; border:none; background:transparent; }
.modal-close:hover { background:rgba(255,255,255,0.05); color:var(--text); }
.modal-body { padding:16px 20px; overflow-y:auto; flex:1; }
.modal-chart { min-height:260px; margin-bottom:12px; }
.perf-table { width:100%; border-collapse:collapse; font-size:14px; }
.perf-table th { background:#1a1d2e; color:#f3f4f6; padding:10px 12px; text-align:right; border-bottom:2px solid var(--border); font-weight:700; white-space:nowrap; }
.perf-table th:first-child { text-align:left; }
.perf-table td { padding:10px 12px; text-align:right; border-bottom:1px solid var(--border); color:#1f2937; font-weight:600; }
.perf-table td:first-child { text-align:left; font-weight:700; color:#111827; }
.progress-good { color:#5AD8A6 !important; }
.progress-warn { color:#F6BD16 !important; }
.progress-bad { color:#E86452 !important; }
</style>
</head>
<body>
<div class="modal-overlay" id="kpiModal">
    <div class="modal-box">
        <div class="modal-header">
            <h2 id="modalTitle">门店明细</h2>
            <button class="modal-close" onclick="closeModal()">✕</button>
        </div>
        <div class="modal-body" id="modalBody"></div>
    </div>
</div>
<div class="modal-overlay" id="profitModal">
    <div class="modal-box" style="max-width:1100px;">
        <div class="modal-header">
            <h2 id="profitModalTitle">门店盈利分析</h2>
            <button class="modal-close" onclick="closeProfitModal()">✕</button>
        </div>
        <div id="profitModalBody"></div>
    </div>
</div>
<div class="dashboard" id="dashboard">
    <div class="header">
        <h1>微笑客经营看板</h1>
        <div id="loadingMsg" style="text-align:center;padding:40px;color:#6b7280;font-size:14px;">⏳ 正在加载数据...</div>
        <div style="display:flex;align-items:center;gap:8px;">
            <button class="export-btn" onclick="openProfitModal()" style="font-size:13px;padding:6px 16px;">💰 门店盈利分析</button>
            <button class="export-btn" onclick="exportExcel()">📥 导出</button>
            <span class="update-time" id="updateTime"></span>
            <span class="update-time" id="dateRangeLabel" style="font-size:13px;color:var(--accent)"></span>
        </div>
    </div>
    <div class="page-nav" style="display:flex;gap:0;margin-bottom:16px;">
        <button class="page-btn active" onclick="switchPage('ops',this)" style="padding:8px 24px;border:1px solid var(--border);background:var(--card);color:var(--text);border-radius:8px 0 0 8px;cursor:pointer;font-size:14px;">运营数据</button>
        <button class="page-btn" onclick="switchPage('product',this)" style="padding:8px 24px;border:1px solid var(--border);background:var(--card);color:var(--text-dim);border-radius:0 8px 8px 0;cursor:pointer;font-size:14px;border-left:none;">商品运营</button>
    </div>
    <div id="page-ops">

    <div class="filter-bar">
        <div class="filter-row">
            <span class="filter-label">📅</span>
            <div class="date-quick" id="dateQuick">
                <button class="date-btn" onclick="setDateRange('yesterday')">昨天</button>
                <button class="date-btn" onclick="setDateRange('last7')">近7天</button>
                <button class="date-btn" onclick="setDateRange('last30')">近30天</button>
                <button class="date-btn" onclick="setDateRange('thisMonth')">本月</button>
                <button class="date-btn" onclick="setDateRange('lastMonth')">上月</button>
                <button class="date-btn active" onclick="toggleCalendar('day')">自定义日</button>
                <button class="date-btn" onclick="toggleCalendar('month')">自定义月</button>
            </div>
        </div>
        <div class="calendar-wrap" id="calendarWrap">
            <div class="cal-header">
                <span class="cal-nav" onclick="calMonth(-1)">‹</span>
                <span class="cal-title" id="calTitle"></span>
                <span class="cal-nav" onclick="calMonth(1)">›</span>
            </div>
            <div class="cal-grid" id="calGrid"></div>
            <div style="display:flex;gap:6px;margin-top:6px;flex-wrap:wrap;">
                <span class="chip-action" id="calRangeLabel" style="font-size:11px;color:var(--accent);"></span>
            </div>
        </div>
        <div class="filter-row">
            <span class="filter-label">🏪</span>
            <div style="flex:1;">
                <input class="store-search" id="storeSearch" placeholder="搜索门店 / 批量粘贴ID(空格逗号换行分隔)" oninput="filterStores()">
                <div style="display:none;margin-top:2px;font-size:11px;color:var(--accent);" id="batchSelected"></div>
                <div class="chip-group" id="storeChips">''' + store_chips_html + '''</div>
            </div>
            <div class="chip-actions">
                <button class="chip-action" onclick="selectAll('stores')">全选</button>
                <button class="chip-action" onclick="deselectAll('stores')">清空</button>
            </div>
        </div>
        <div class="filter-row">
            <span class="filter-label">📡</span>
            <div class="chip-group" id="channelChips">''' + channel_chips_html + '''</div>
            <div class="chip-actions">
                <button class="chip-action" onclick="selectAll('channels')">全选</button>
                <button class="chip-action" onclick="deselectAll('channels')">清空</button>
            </div>
        </div>
    </div>

    <div class="alert-banner" id="alertBanner">
        <span class="alert-icon">⚠️</span>
        <span class="alert-msg" id="alertMsg"></span>
        <span class="alert-dismiss" onclick="document.getElementById('alertBanner').classList.remove('show')">×</span>
    </div>
    <!-- 绩效弹窗 -->
    <div class="modal-overlay" id="perfModal" onclick="if(event.target===this)this.classList.remove('show')">
        <div class="modal-box">
            <div class="modal-header">
                <h3>📊 <span id="perfMonth">6月</span>绩效进度 <span style="font-size:12px;color:#6b7280;font-weight:400" id="perfDays">(加载中...)</span></h3>
                <div style="font-size:13px;color:#374151;font-weight:600">
                    🕐 时间进度 <span id="perfTimeBar">10/30天 (33.3%)</span>
                    <span style="display:inline-block;width:120px;height:8px;background:#e5e7eb;border-radius:4px;vertical-align:middle;margin-left:6px;overflow:hidden">
                        <span id="perfTimeFill" style="display:block;height:100%;background:var(--blue);border-radius:4px;width:33%"></span>
                    </span>
                </div>
                <button class="modal-close" onclick="document.getElementById('perfModal').classList.remove('show')">×</button>
            </div>
            <table class="perf-table">
                <thead><tr><th>负责人</th><th>单量</th><th>目标</th><th>进度</th><th>去推广毛利</th><th>毛利目标</th><th>进度</th></tr></thead>
                <tbody id="perfBody"></tbody>
                <tfoot id="perfFoot"></tfoot>
            </table>
        </div>
    </div>
    <div class="summary-panel" id="summaryPanel">
        <h3 style="cursor:pointer;user-select:none" onclick="showPerfModal()" title="点击查看6月绩效进度">📋 数据快报 🔍</h3>
        <div class="summary-text" id="summaryText"></div>
    </div>

    <div class="kpi-grid" id="kpiGrid"></div>

    <div class="tab-nav-wrap">
        <div class="tab-nav">
            <button class="tab-btn active" onclick="switchTab('store',this)">门店数据</button>
            <button class="tab-btn" onclick="switchTab('channel',this)">渠道分析</button>
            <button class="tab-btn" onclick="switchTab('time',this)">时间分析</button>
            <button class="tab-btn" onclick="switchTab('neg',this)">负毛利</button>
            <button class="tab-btn" onclick="switchTab('delivery',this)">配送</button>
        </div>
    </div>

    <div id="tab-store" class="tab-content active">
        <div class="chart-section"><h3>各门店单量 vs 抽佣毛利<span class="section-date-tag"></span></h3><div id="chartStoreBar"></div></div>
        <div class="chart-section"><h3>门店明细 <span class="section-date-tag"></span><button class="export-btn" onclick="exportTable('storeTable','门店明细')">📥 导出Excel</button></h3><div class="table-scroll" id="storeTable"></div></div>
    </div>
    <div id="tab-channel" class="tab-content">
        <div class="row">
            <div class="chart-section"><h3>单量渠道分布</h3><div id="chartChannelPie"></div></div>
            <div class="chart-section"><h3>各渠道毛利率</h3><div id="chartChannelMargin"></div></div>
        </div>
        <div class="row">
            <div class="chart-section"><h3>各渠道实收</h3><div id="chartChannelRev"></div></div>
            <div class="chart-section"><h3>各渠道公司抽佣</h3><div id="chartChannelComm"></div></div>
        </div>
        <div class="chart-section"><h3>渠道明细 <span class="section-date-tag"></span><button class="export-btn" onclick="exportTable('channelTable','渠道明细')">📥 导出Excel</button></h3><div class="table-scroll" id="channelTable"></div></div>
    </div>
    <div id="tab-time" class="tab-content">
        <div class="chart-section"><h3>每日单量 & 抽佣毛利趋势<span class="section-date-tag"></span></h3><div id="chartTimeTrend"></div></div>
        <div class="chart-section"><h3>每日明细 <span class="section-date-tag"></span><button class="export-btn" onclick="exportTable('timeTable','每日明细')">📥 导出Excel</button></h3><div class="table-scroll" id="timeTable"></div></div>
    </div>
    <div id="tab-neg" class="tab-content">
        <div class="chart-section"><h3>每日负毛利占比趋势<span class="section-date-tag"></span></h3><div id="chartNegStore"></div></div>
        <div class="chart-section"><h3>门店负毛利占比<span class="section-date-tag"></span></h3><div id="chartNegByStore"></div></div>
        <div class="chart-section"><h3>负毛利明细 <span class="section-date-tag"></span><button class="export-btn" onclick="exportTable('negTable','负毛利明细')">📥 导出Excel</button></h3><div class="table-scroll" id="negTable"></div></div>
    </div>
    <div id="tab-delivery" class="tab-content">
        <div class="chart-section"><h3>每日单均配送成本<span class="section-date-tag"></span></h3><div id="chartDelCost"></div></div>
        <div class="chart-section"><h3>门店单均配送成本<span class="section-date-tag"></span></h3><div id="chartDelRatio"></div></div>
        <div class="chart-section"><h3>配送明细 <span class="section-date-tag"></span><button class="export-btn" onclick="exportTable('delTable','配送明细')">📥 导出Excel</button></h3><div class="table-scroll" id="delTable"></div></div>
    </div>
</div>

<div id="page-product" style="display:none;">
    <div class="filter-bar">
        <div class="filter-row">
            <span class="filter-label">月份</span>
            <div class="chip-group" id="prodPeriodChips"></div>
        </div>
    </div>
    <div class="kpi-grid" id="prodKpiGrid"></div>
    <div class="tab-nav-wrap">
        <div class="tab-nav">
            <button class="tab-btn active" onclick="switchProdTab('turnover',this)">动销分析</button>
            <button class="tab-btn" onclick="switchProdTab('sales',this)">品类销售</button>
        </div>
    </div>
    <div id="prod-tab-turnover" class="tab-content active">
        <div class="chart-section"><h3>大类动销率 & 毛利率对比</h3><div id="chartProdCatTurnover"></div></div>
        <div class="chart-section"><h3>门店动销率排名</h3><div id="chartProdStoreTurnover"></div></div>
        <div class="chart-section"><h3>一级分类动销明细<button class="export-btn" onclick="exportTable('prodRawcatTable','一级分类动销')">📥 导出分类表</button><button class="export-btn" style="margin-left:4px" onclick="exportStoreCat()">📥 导出各门店ABC</button></h3><div class="table-scroll" id="prodRawcatTable"></div></div>
    </div>
    <div id="prod-tab-sales" class="tab-content">
        <div class="chart-section"><h3>品类销售汇总<button class="export-btn" onclick="exportTable('prodSalesSummary','品类销售汇总')">📥 导出Excel</button></h3><div class="table-scroll" id="prodSalesSummary"></div></div>
        <div class="chart-section"><h3>一级分类销售明细<button class="export-btn" onclick="exportTable('prodSalesDetail','一级分类销售明细')">📥 导出Excel</button></h3><div class="table-scroll" id="prodSalesDetail"></div></div>
    </div>
</div>

<script>
// ============ EXPORT ============
function exportTable(tableId, filename) {
    const container = document.getElementById(tableId);
    if (!container) return;
    const table = container.querySelector('table');
    if (!table) return;

    // 检查是否有牵牛花门店ID（data-qn属性）
    const qnRows = table.querySelectorAll('tr[data-qn]');
    if (qnRows.length > 0) {
        // 有门店ID：手动构建带牵牛花门店ID、美团门店ID列的sheet
        const headers = [];
        table.querySelectorAll('tr:first-child th').forEach(th => headers.push(th.textContent.trim()));
        headers.unshift('美团门店ID');
        headers.unshift('牵牛花门店ID');

        const data = [];
        table.querySelectorAll('tr[data-qn]').forEach(tr => {
            const qn = tr.getAttribute('data-qn') || '';
            const mt = tr.getAttribute('data-meituan') || '';
            const row = [qn, mt];
            tr.querySelectorAll('td').forEach(td => row.push(td.textContent.trim()));
            data.push(row);
        });

        const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '明细');
        XLSX.writeFile(wb, filename + '.xlsx');
    } else {
        const wb = XLSX.utils.table_to_book(table, {sheet:'明细'});
        XLSX.writeFile(wb, filename + '.xlsx');
    }
}

// Export per-store ABC category turnover data (from productData, no extra embedded data)
function exportStoreCat() {
    const storeCat = getPeriodData(selectedProdPeriod).filter(r => r.store_name !== '全部门店' && ['食品','日化','百货'].includes(r.cat));
    if (!storeCat.length) { alert('暂无门店品类数据'); return; }
    const rows = storeCat.map(r => ({
        '牵牛花门店ID': storeQnMap[r.store_name] || '',
        '门店': r.store_name,
        '品类': r.cat,
        'SKU总数': Math.round(r.sku_total || 0),
        '有动销SKU': Math.round(r.sku_active || 0),
        '无动销SKU': Math.round((r.sku_total || 0) - (r.sku_active || 0)),
        '动销率': parseFloat((r.turnover_rate || 0).toFixed(1)),
        '货值': Math.round(r.goods_value || 0),
        '销量': Math.round(r.qty || 0),
        '销售额': Math.round(r.revenue || 0),
        '毛利': Math.round(r.profit || 0),
        '毛利率': parseFloat((r.margin || 0).toFixed(1)),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{wch:14},{wch:28},{wch:8},{wch:10},{wch:10},{wch:10},{wch:8},{wch:10},{wch:8},{wch:10},{wch:10},{wch:8}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '门店品类动销');
    XLSX.writeFile(wb, '各门店ABC动销.xlsx');
}


// ============ DATA ============
const DATA_URL = "https://raw.githubusercontent.com/deacyn154/smile-ke-bi/main/data.json";
let rawCompact = [];
const chNames = ['美团闪购','饿了么','京东到家','线下'];
let rawData = [];

// 从 GitHub 加载数据(支持本地回退)
async function loadData() {
    try {
        const resp = await fetch(DATA_URL, {cache: "no-cache"});
        if (!resp.ok) throw new Error("fetch failed");
        rawCompact = await resp.json();
    } catch(e) {
        console.warn("GitHub load failed");
        rawCompact = [];
    }
    // 解码
    rawData = rawCompact.map(r => ({
        '日期': String(r[0]).replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3'),
        'store_name': storeNameMap[String(r[1])] || String(r[1]),
        'qn_store_id': r[1],
        'channel': chNames[r[2]],
        'order_cnt': r[3],
        'revenue': r[4] || 0,
        'store_profit': r[5] || 0,
        'commission_profit': r[6] || 0,
        'neg_cnt': r[7] || 0,
        'delivery_fee': r[8] || 0,
        'delivery_order_cnt': r[9] || 0,
        'real_profit': r[6] || 0,
        'commission_fee': (r[4]||0) - (r[6]||0),
        'promo_fee': ((r[4]||0) - (r[5]||0)) + ((r[4]||0) - (r[6]||0)),
    }));
    rawData.sort((a,b) => (a['日期']+a['store_name']+a['channel']).localeCompare(b['日期']+b['store_name']+b['channel']));
    document.getElementById('loadingMsg').style.display = 'none';
    initDashboard();
}
// 页面加载时调用
loadData();
const promoData = ''' + promo_json + ''';
const productData = ''' + product_json + ''';
const allProductData = ''' + all_product_json + ''';
const productPeriods = ''' + product_json_all + ''';
const rawcatData = ''' + rawcat_json + ''';
const rawcatCombinedData = ''' + rawcat_combined_json + ''';
const allStoresFull = ''' + stores_json + ''';
const allChannels = ''' + channels_json + ''';
const allDates = ''' + dates_json + ''';
const shortNames = ''' + short_map_json + ''';
const searchKeys = ''' + search_map_json + ''';
const storeNameMap = ''' + store_name_map_json + ''';  // qn_store_id -> 标准名
const storeRegionMap = ''' + store_region_json + ''';  // 门店→所属地区
const storeCostMap = ''' + store_cost_json + ''';  // 门店→{cost, qnid}
const storeMeituanMap = ''' + store_meituan_json + ''';  // 门店→美团渠道门店ID
const perfData = ''' + perf_json + ''';  // 当前月绩效进度

// 门店名 → 牵牛花门店ID 映射（导出Excel时用）
const storeQnMap = {};
rawData.forEach(r => {
    if (r.store_name && r.qn_store_id !== undefined && r.qn_store_id !== null && !storeQnMap[r.store_name]) {
        storeQnMap[r.store_name] = String(r.qn_store_id);
    }
});

// 门店→省市 映射 + 按省市排序
const storeProvince = {};
(function() {
    // 从所属地区提取省份（格式：广东省深圳市龙华区 → 广东省）
    function extractProvince(region) {
        if (!region) return '其他';
        // 直辖市
        if (region.startsWith('北京')) return '北京';
        if (region.startsWith('上海')) return '上海';
        if (region.startsWith('天津')) return '天津';
        if (region.startsWith('重庆')) return '重庆';
        // 自治区
        if (region.includes('自治区')) return region.split('自治区')[0] + '自治区';
        // 普通省份
        if (region.includes('省')) return region.split('省')[0];
        return region.substring(0, 2);
    }
    for (let s of allStoresFull) {
        let region = storeRegionMap[s] || '';
        storeProvince[s] = extractProvince(region);
    }
    // 按省市排序
    const provOrder = ['广东省','湖南省','广西壮族自治区','福建省','江西省','海南省','湖北省'];
    allStoresFull.sort((a,b) => {
        let pa = provOrder.indexOf(storeProvince[a]), pb = provOrder.indexOf(storeProvince[b]);
        if (pa>=0 && pb<0) return -1;
        if (pa<0 && pb>=0) return 1;
        if (pa>=0 && pb>=0 && pa!==pb) return pa-pb;
        return a.localeCompare(b);
    });
})();

const isMobile = window.innerWidth < 768;
const chartH = isMobile ? 260 : 380;
const chartHS = isMobile ? 220 : 340;

const plotlyLayout = { paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#1f2937',family:'-apple-system,PingFang SC,Microsoft YaHei'}, xaxis:{gridcolor:'#e5e7eb',zerolinecolor:'#e5e7eb'}, yaxis:{gridcolor:'#e5e7eb',zerolinecolor:'#e5e7eb'}, margin:{l:60,r:20,t:20,b:60}, height:chartH };
const plotlyCfg = { displayModeBar:false, responsive:true };

// ============ SHORT NAME HELPERS ============
function short(full) { return shortNames[full] || full; }

// ============ DATE STATE ============
let dateFrom, dateTo;
let calDate = new Date();
let calSelect = []; // clicked dates in [start, end]
let calOpen = false;

// Convert YYYY-MM-DD to Date (local)
function dParse(s) { let p=s.split('-'); return new Date(+p[0],+p[1]-1,+p[2]); }
function dFmt(d) { return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
function today() { let d=new Date(); return dFmt(d); }
function yesterday() { let d=new Date(); d.setDate(d.getDate()-1); return dFmt(d); }

function calcWeekBounds(d) { let wd = d.getDay()||7; let mon = new Date(d); mon.setDate(d.getDate()-wd+1); let sun = new Date(mon); sun.setDate(mon.getDate()+6); return [mon,sun]; }
function calcMonthBounds(d) { return [new Date(d.getFullYear(),d.getMonth(),1), new Date(d.getFullYear(),d.getMonth()+1,0)]; }

// Find nearest available date in allDates
function nearestDate(target, dir) {
    let t = dParse(target);
    let best = null, bestDiff = Infinity;
    allDates.forEach(d => {
        let dd = dParse(d);
        let diff = Math.abs(dd - t);
        if (diff < bestDiff) { bestDiff = diff; best = d; }
    });
    return best || allDates[0];
}

// ============ DATE RANGE BUTTONS ============
function setDateRange(mode) {
    document.querySelectorAll('.date-btn').forEach(b => b.classList.remove('active'));
    let from, to;
    switch(mode) {
        case 'yesterday':
            from = yesterday(); to = from;
            break;
        case 'last7':
            to = allDates[allDates.length-1];
            from = dFmt(new Date(dParse(to).getTime()-6*86400000));
            break;
        case 'last30':
            to = allDates[allDates.length-1];
            from = dFmt(new Date(dParse(to).getTime()-29*86400000));
            break;
        case 'lastMonth':
            let d2 = new Date(); d2.setMonth(d2.getMonth()-1);
            let [lm1,lm2] = calcMonthBounds(d2);
            from = dFmt(lm1); to = dFmt(lm2);
            break;
        case 'thisMonth':
            let [tm1,tm2] = calcMonthBounds(new Date());
            from = dFmt(tm1); to = dFmt(tm2);
            break;
        default: break;
    }
    dateFrom = nearestDate(from); dateTo = nearestDate(to);
    if (dateFrom > dateTo) [dateFrom, dateTo] = [dateTo, dateFrom];
    if (mode !== 'custom') {
        let btn = document.querySelector('.date-btn[onclick*="'+mode+'"]');
        if (btn) btn.classList.add('active');
    }
    calOpen = false; document.getElementById('calendarWrap').classList.remove('show');
    // Show date range label
    let drLabel = dateFrom === dateTo ? dateFrom : dateFrom + ' ~ ' + dateTo;
    document.getElementById('dateRangeLabel').textContent = '统计时段: ' + drLabel;
    // Also update table section headers
    document.querySelectorAll('.section-date-tag').forEach(el => { el.textContent = drLabel; });
    refresh();
}

let calMode = 'day'; // 'day' or 'month'
function toggleCalendar(mode) {
    calMode = mode || 'day';
    // Update button active state
    document.querySelectorAll('.date-btn[onclick*="toggleCalendar"]').forEach(b=>b.classList.remove('active'));
    let btn = document.querySelector('.date-btn[onclick*="'+calMode+'"]');
    if (btn) btn.classList.add('active');
    
    calOpen = !calOpen;
    document.getElementById('calendarWrap').classList.toggle('show', calOpen);
    if (calOpen) {
        calDate = dParse(dateFrom || allDates[0]);
        calSelect = [];
        renderCalendar();
    }
}

// ============ CALENDAR ============
function renderCalendar() {
    const y = calDate.getFullYear(), m = calDate.getMonth();
    document.getElementById('calTitle').textContent = y+'年'+(m+1)+'月';
    
    if (calMode === 'month') {
        renderMonthSelector(y, m);
        return;
    }
    renderDayGrid(y, m);
}

function renderMonthSelector(y, m) {
    // Show 12 months grid for quick month range selection
    let html = '<div class="cal-grid" style="grid-template-columns:repeat(4,1fr);">';
    const months = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
    for (let i=0; i<12; i++) {
        let d = new Date(y, i, 1);
        let ds = dFmt(d);
        let hasData = allDates.indexOf(ds) >= 0;
        let cls = 'cal-day' + (hasData ? ' has-data' : ' no-data');
        html += '<div class="'+cls+'" style="padding:12px 8px;font-size:14px;" onclick="selectMonth('+y+','+i+')">'+months[i]+'</div>';
    }
    html += '</div>';
    document.getElementById('calGrid').innerHTML = html;
}

function selectMonth(year, month) {
    let firstDay = new Date(year, month, 1);
    let lastDay = new Date(year, month+1, 0);
    let from = dFmt(firstDay), to = dFmt(lastDay);
    
    // Find actual data range within this month
    let monthDates = allDates.filter(d => d >= from && d <= to);
    if (monthDates.length > 0) {
        from = monthDates[0]; to = monthDates[monthDates.length-1];
    }
    
    dateFrom = from; dateTo = to;
    calSelect = [from, to];
    calOpen = false;
    document.getElementById('calendarWrap').classList.remove('show');
    document.querySelectorAll('.date-btn').forEach(b=>b.classList.remove('active'));
    
    let drLabel = dateFrom === dateTo ? dateFrom : dateFrom + ' ~ ' + dateTo;
    document.getElementById('dateRangeLabel').textContent = '统计时段: ' + drLabel;
    document.querySelectorAll('.section-date-tag').forEach(el => { el.textContent = drLabel; });
    refresh();
}

function renderDayGrid(y, m) {
    const first = new Date(y,m,1).getDay() || 7;
    const daysInMonth = new Date(y,m+1,0).getDate();
    const daysInPrev = new Date(y,m,0).getDate();
    let html = '<div class="cal-weekday">一</div><div class="cal-weekday">二</div><div class="cal-weekday">三</div><div class="cal-weekday">四</div><div class="cal-weekday">五</div><div class="cal-weekday">六</div><div class="cal-weekday">日</div>';

    // prev month
    for (let i=first-2; i>=0; i--) {
        let day = daysInPrev - i;
        let d = new Date(y,m-1,day);
        let ds = dFmt(d);
        html += '<div class="cal-day other-month" data-date="'+ds+'" onclick="calClick(\\''+ds+'\\')">'+day+'</div>';
    }
    // this month
    for (let d=1; d<=daysInMonth; d++) {
        let dt = new Date(y,m,d);
        let ds = dFmt(dt);
        let hasData = allDates.indexOf(ds) >= 0;
        let cls = 'cal-day' + (hasData ? ' has-data' : ' no-data');
        if (calSelect.length===2) {
            let s=dParse(calSelect[0]), e=dParse(calSelect[1]);
            if (dt>=s && dt<=e) {
                cls += ' in-range';
                if (dFmt(dt)===calSelect[0]) cls += ' range-start';
                if (dFmt(dt)===calSelect[1]) cls += ' range-end';
            }
        } else if (calSelect.length===1 && calSelect[0]===ds) {
            cls += ' range-start range-end';
        }
        let onClick = hasData ? 'calClick(\\''+ds+'\\')' : '';
        html += '<div class="'+cls+'" data-date="'+ds+'" onclick="'+onClick+'">'+d+'</div>';
    }
    // next month
    let total = first + daysInMonth;
    let remaining = (7 - total % 7) % 7;
    for (let d=1; d<=remaining; d++) {
        let dt = new Date(y,m+1,d);
        let ds = dFmt(dt);
        html += '<div class="cal-day other-month" data-date="'+ds+'" onclick="calClick(\\''+ds+'\\')">'+d+'</div>';
    }
    document.getElementById('calGrid').innerHTML = html;
    updateCalRange();
}

function calMonth(delta) {
    calDate.setMonth(calDate.getMonth()+delta);
    renderCalendar();
}

function calClick(ds) {
    if (calSelect.length===0 || calSelect.length===2) {
        calSelect = [ds];
    } else {
        let s = dParse(calSelect[0]), e = dParse(ds);
        if (e < s) [s,e] = [e,s];
        calSelect = [dFmt(s), dFmt(e)];
        dateFrom = nearestDate(calSelect[0]);
        dateTo = nearestDate(calSelect[1]);
        document.querySelectorAll('.date-btn').forEach(b => b.classList.remove('active'));
        refresh();
    }
    renderCalendar();
}

function updateCalRange() {
    let el = document.getElementById('calRangeLabel');
    if (calSelect.length===1) el.textContent = '已选: '+calSelect[0]+' (请选结束日期)';
    else if (calSelect.length===2) el.textContent = '区间: '+calSelect[0]+' ~ '+calSelect[1];
    else el.textContent = '点击日期选开始';
}

// ============ STORE SEARCH ============
function filterStores() {
    let q = document.getElementById("storeSearch").value.trim();
    if (!q) {
        document.querySelectorAll("#storeChips .chip").forEach(c => c.classList.remove("hidden"));
        document.getElementById("batchSelected").style.display = "none";
        return;
    }
    // Detect batch paste (has spaces or commas)
    var hasDelim = q.includes(" ") || q.includes(",");
    var ids = hasDelim ? q.split(/[ ,]+/).filter(Boolean) : [q];
    var matched = [];
    document.querySelectorAll("#storeChips .chip").forEach(c => {
        var full = c.getAttribute("data-full");
        var keys = JSON.parse(searchKeys[full] || "[]");
        var m = ids.some(function(id) { return keys.some(function(k) { return k.toLowerCase().includes(id.toLowerCase()); }); });
        if (m) matched.push(full);
        else if (hasDelim) c.classList.add("hidden");
        else c.classList.toggle("hidden", !m);
    });
    if (hasDelim && matched.length > 0) {
        selectedStores.splice(0, selectedStores.length, ...matched);
        refreshChips(); refresh();
        document.getElementById("storeSearch").value = "";
        document.getElementById("batchSelected").textContent = "已选中 " + matched.length + " 家门店";
        document.getElementById("batchSelected").style.display = "block";
    } else if (!hasDelim) {
        document.getElementById("batchSelected").style.display = "none";
    }
}

// ============ FILTER STATE ============
let selectedStores = [...allStoresFull];
let selectedChannels = [...allChannels];

function createChips(containerId, items, selectedArr, onChange) {
    const container = document.getElementById(containerId);
    if (!container) return;
    // chips已预渲染，只同步激活状态
    container.querySelectorAll('.chip').forEach(chip => {
        const item = chip.getAttribute('data-full');
        chip.classList.toggle('active', selectedArr.indexOf(item) > -1);
    });
    // 同步：如果全部被selected, 确保都显示active
    if (container.querySelectorAll('.chip.active').length === 0 && selectedArr.length > 0) {
        container.querySelectorAll('.chip').forEach(chip => {
            const item = chip.getAttribute('data-full');
            chip.classList.toggle('active', selectedArr.indexOf(item) > -1);
        });
    }
}

// 全局chip点击事件委派
document.addEventListener('click', function(e) {
    const chip = e.target.closest('.chip');
    if (!chip) return;
    const container = chip.parentElement;
    if (!container) return;
    
    if (container.id === 'storeChips') {
        const item = chip.getAttribute('data-full');
        const idx = selectedStores.indexOf(item);
        if (idx > -1) { selectedStores.splice(idx,1); chip.classList.remove('active'); }
        else { selectedStores.push(item); chip.classList.add('active'); }
    } else if (container.id === 'channelChips') {
        const item = chip.getAttribute('data-full');
        const idx = selectedChannels.indexOf(item);
        if (idx > -1) { selectedChannels.splice(idx,1); chip.classList.remove('active'); }
        else { selectedChannels.push(item); chip.classList.add('active'); }
    }
    refresh();
});

function selectAll(type) {
    if (type==='stores') { selectedStores.splice(0, selectedStores.length, ...allStoresFull); }
    else { selectedChannels.splice(0, selectedChannels.length, ...allChannels); }
    refreshChips(); refresh();
}
function deselectAll(type) {
    if (type==='stores') { selectedStores.splice(0, selectedStores.length); }
    else { selectedChannels.splice(0, selectedChannels.length); }
    refreshChips(); refresh();
}
function refreshChips() {
    document.querySelectorAll('#storeChips .chip').forEach(c => {
        c.classList.toggle('active', selectedStores.includes(c.getAttribute('data-full')));
    });
    document.querySelectorAll('#channelChips .chip').forEach(c => {
        c.classList.toggle('active', selectedChannels.includes(c.textContent));
    });
}

function initFilters() {
    setDateRange('last7');
    createChips('storeChips', allStoresFull, selectedStores, refresh);
    createChips('channelChips', allChannels, selectedChannels, refresh);
}

// ============ TAB ============
function switchTab(name, btn) {
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-'+name).classList.add('active');
    // Force Plotly charts to fill container after tab switch
    setTimeout(() => {
        document.querySelectorAll('#tab-'+name+' .js-plotly-plot').forEach(el => {
            if (el._fullLayout) Plotly.Plots.resize(el);
        });
    }, 50);
}

// ============ HELPERS ============
function sum(arr, key) { return arr.reduce((s,r)=>s+(r[key]||0),0); }
function fmtY(n) { return '¥' + (n||0).toLocaleString(undefined,{maximumFractionDigits:0}); }

function groupBy(arr, keys) {
    const map={};
    arr.forEach(r=>{
        const k=keys.map(k2=>r[k2]).join('|');
        if(!map[k]){map[k]={};keys.forEach(kk=>map[k][kk]=r[kk]);map[k]._rows=[];}
        map[k]._rows.push(r);
    });
    return Object.values(map).map(g=>{
        const result={};
        keys.forEach(kk=>result[kk]=g[kk]);
        g._rows.forEach(r=>{Object.keys(r).forEach(k=>{if(!keys.includes(k))result[k]=(result[k]||0)+(r[k]||0);});});
        return result;
    });
}

function getFiltered() {
    const f=dateFrom, t=dateTo;
    return rawData.filter(r=>r['日期']>=f&&r['日期']<=t&&selectedStores.includes(String(r['qn_store_id']))&&selectedChannels.includes(r['channel']));
}
function getFilteredPromo() {
    const f=dateFrom, t=dateTo;
    return promoData.filter(r=>r['日期']>=f&&r['日期']<=t&&selectedStores.includes(String(r['qn_store_id'])));
}

// ============ 环比 ============
function getPrevData(data) {
    const from = dParse(dateFrom), to = dParse(dateTo);
    let prevStart, prevEnd;
    // 检测是否为完整月 (from=1号, to=月末最后一天)
    const lastDayOfMonth = new Date(from.getFullYear(), from.getMonth()+1, 0).getDate();
    const isFullMonth = (from.getDate() === 1 && to.getDate() === lastDayOfMonth);
    if (isFullMonth) {
        // 整月对比: 直接取上个月整月 (如3.1-3.31 vs 2.1-2.28)
        prevStart = new Date(from.getFullYear(), from.getMonth()-1, 1);
        prevEnd = new Date(from.getFullYear(), from.getMonth(), 0);
    } else {
        // 其他标签: 同比相同天数
        const days = Math.round((to - from) / (24*60*60*1000)) + 1;
        prevEnd = new Date(from); prevEnd.setDate(prevEnd.getDate() - 1);
        prevStart = new Date(prevEnd); prevStart.setDate(prevStart.getDate() - days + 1);
    }
    const pf = dFmt(prevStart), pt = dFmt(prevEnd);
    return rawData.filter(r =>
        r['日期'] >= pf && r['日期'] <= pt &&
        selectedStores.includes(String(r['qn_store_id'])) &&
        selectedChannels.includes(r['channel'])
    );
}

function deltaStr(curr, prev, isGoodUp) {
    if (prev === 0) return { txt: '--', cls: '' };
    const pct = ((curr - prev) / prev * 100).toFixed(1);
    const up = pct > 0;
    const sign = up ? '↑' : '↓';
    const cls = up ? (isGoodUp ? 'green' : 'red') : (isGoodUp ? 'red' : 'green');
    return { txt: sign + Math.abs(pct) + '%', cls: cls };
}

// ============ SMART SUMMARY ============
function renderSummary(data, prev) {
    const ord=sum(data,'order_cnt'), rev=sum(data,'revenue'), cp=sum(data,'commission_profit');
    const nc=sum(data,'neg_cnt'), pf=sum(data,'promo_fee');
    const margin=rev>0?(cp/rev*100).toFixed(1):0;
    const negPct=ord>0?(nc/ord*100).toFixed(1):0;

    const pOrd=sum(prev,'order_cnt'), pRev=sum(prev,'revenue'), pCp=sum(prev,'commission_profit');
    const ordChg=pOrd>0?((ord-pOrd)/pOrd*100).toFixed(1):0;
    const revChg=pRev>0?((rev-pRev)/pRev*100).toFixed(1):0;
    const cpChg=pCp>0?((cp-pCp)/pCp*100).toFixed(1):0;

    // Top/Bottom stores
    const stores=groupBy(data,['store_name']);
    stores.sort((a,b)=>b.commission_profit-a.commission_profit);
    const top3=stores.slice(0,3).map(s=>short(s.store_name)).join('、');
    const bot3=stores.slice(-3).reverse().filter(s=>s.commission_profit<0).map(s=>short(s.store_name));
    const botText=bot3.length>0?bot3.join('、'):null;

    // High neg-margin stores
    const negStores=stores.filter(s=>s.order_cnt>0&&(s.neg_cnt/s.order_cnt*100)>30).map(s=>short(s.store_name));
    const negText=negStores.length>0?negStores.slice(0,3).join('、'):null;

    // Channel analysis
    const chs=groupBy(data,['channel']);
    const chInfo=chs.map(c=>({n:c.channel, ord:c.order_cnt, rev:c.revenue}));
    chInfo.sort((a,b)=>b.ord-a.ord);
    const topCh=chInfo[0];

    const parts=[];
    // 1. Overall
    parts.push('本期共 <span class="highlight-blue">'+ord.toLocaleString()+'</span> 单，实收 <span class="highlight-blue">'+fmtY(rev)+
        '</span>，抽佣毛利 <span class="highlight-blue">'+fmtY(cp)+'</span>（毛利率 '+margin+'%）。');
    // 2. Change
    const chgDir=parseFloat(ordChg)>0?'增长':'下降';
    parts.push('环比'+chgDir+' <span class="'+(parseFloat(ordChg)>0?'highlight-red':'highlight-green')+'">'+Math.abs(ordChg)+'%</span>。');
    // 3. Top channel
    if(topCh) parts.push('最大渠道 <span class="highlight-blue">'+topCh.n+'</span>（'+topCh.ord.toLocaleString()+'单）。');
    // 4. Top stores
    parts.push('毛利贡献前三：<span class="highlight-red">'+top3+'</span>。');
    // 5. Warnings
    if(botText) parts.push('毛利亏损：<span class="highlight-green">'+botText+'</span>。');
    if(negText) parts.push('高负毛利门店（>30%）：<span class="highlight-green">'+negText+'</span>。');
    if(parseFloat(negPct)>25) parts.push('整体负毛利占比 <span class="highlight-green">'+negPct+'%</span>，需重点关注。');
    // 6. Promotion
    if(pf>0) parts.push('推广费合计 <span class="highlight-yellow">¥'+pf.toFixed(0)+'</span>。');

    document.getElementById('summaryText').innerHTML=parts.join('');
}

// ============ ANOMALY ALERTS ============
function renderAlerts(data, prev) {
    const alerts=[];
    // Check per-store neg margin spike
    const stores=groupBy(data,['store_name']);
    const prevStores=groupBy(prev,['store_name']);
    const prevMap={}; prevStores.forEach(s=>{prevMap[s.store_name]={ord:s.order_cnt||0,neg:s.neg_cnt||0};});

    const dangerStores=[];
    stores.forEach(s=>{
        const np=s.order_cnt>0?(s.neg_cnt/s.order_cnt*100):0;
        const p=prevMap[s.store_name]||{ord:0,neg:0};
        const pNp=p.ord>0?(p.neg/p.ord*100):0;
        if(np>35&&np-pNp>10) dangerStores.push(short(s.store_name)+' '+np.toFixed(1)+'%');
        else if(np>35) dangerStores.push(short(s.store_name)+' '+np.toFixed(1)+'%');
    });
    if(dangerStores.length>0) {
        alerts.push('负毛利过高：'+dangerStores.slice(0,5).join(' / '));
    }

    // Delivery cost anomalies
    const delStores=stores.filter(s=>(s.delivery_order_cnt||0)>3);
    const delCosts=delStores.map(s=>({
        name:short(s.store_name),
        cost:(s.delivery_order_cnt||0)>0?(s.delivery_fee/s.delivery_order_cnt):0
    }));
    delCosts.sort((a,b)=>b.cost-a.cost);
    const avgCost=delCosts.reduce((s,c)=>s+c.cost,0)/Math.max(1,delCosts.length);
    const highCost=delCosts.filter(c=>c.cost>avgCost*1.3).slice(0,3);
    if(highCost.length>0) {
        alerts.push('配送成本偏高：'+highCost.map(c=>c.name+' ¥'+c.cost.toFixed(1)).join(' / '));
    }

    if(alerts.length>0) {
        document.getElementById('alertBanner').classList.add('show');
        document.getElementById('alertMsg').innerHTML='<b>异常预警：</b>'+alerts.join('<br>');
    } else {
        document.getElementById('alertBanner').classList.remove('show');
    }
}

// ============ SPARKLINE ============
function drawSparkline(containerId, values, colorName) {
    const container=document.getElementById(containerId);
    if(!container||values.length<2) return;
    const color=getComputedStyle(document.documentElement).getPropertyValue('--'+colorName).trim()||'#6C8EF2';
    const w=container.clientWidth||80, h=28;
    const min=Math.min(...values), max=Math.max(...values);
    const range=max-min||1;

    let points=values.map((v,i)=>[(i/(values.length-1))*w, h-(v-min)/range*h]);
    const pathD=points.map((p,i)=>(i===0?'M':'L')+p[0].toFixed(1)+','+p[1].toFixed(1)).join(' ');
    const areaD=pathD+' L'+points[points.length-1][0].toFixed(1)+','+h+' L'+points[0][0].toFixed(1)+','+h+' Z';

    container.innerHTML='<svg viewBox="0 0 '+w+' '+h+'" preserveAspectRatio="none"><path class="area" d="'+areaD+'" fill="'+color+'"/><path d="'+pathD+'" stroke="'+color+'"/></svg>';
}

// ============ SORTABLE TABLE ============
const tableSortState={};
function makeSortable(tableId, data, colDefs, formatRowFn) {
    // colDefs: [{key, label, type:'num'|'str'|'pct', sortable:true}]
    let h='<table><tr>';
    colDefs.forEach((c,i)=>{
        if(c.sortable===false) h+='<th>'+c.label+'</th>';
        else h+=`<th class="sortable" onclick="sortAndRender('${tableId}',${i},'${c.type}')" title="点击排序">${c.label}</th>`;
    });
    h+='</tr>';

    const state=tableSortState[tableId]||{col:-1,asc:true};
    tableSortState[tableId]=state;

    if(state.col>=0&&state.col<colDefs.length) {
        const cdef=colDefs[state.col];
        data.sort((a,b)=>{
            let va=a[cdef.key]||0, vb=b[cdef.key]||0;
            if(cdef.type==='num'||cdef.type==='pct') { va=Number(va); vb=Number(vb); }
            else { va=String(va); vb=String(vb); }
            if(va<vb) return state.asc?-1:1;
            if(va>vb) return state.asc?1:-1;
            return 0;
        });
    }

    data.forEach((r,i)=>{ h+=formatRowFn(r,i); });
    h+='</table>';
    document.getElementById(tableId).innerHTML=h;
    // 标记排序箭头
    if(state.col>=0) {
        const ths = document.getElementById(tableId).querySelectorAll('th');
        ths.forEach((th,i) => th.classList.remove('asc','desc'));
        if(ths[state.col]) ths[state.col].classList.add(state.asc?'asc':'desc');
    }
}
function sortAndRender(tableId, colIdx, type) {
    const state=tableSortState[tableId]||{col:-1,asc:true};
    if(state.col===colIdx) state.asc=!state.asc;
    else { state.col=colIdx; state.asc=true; }
    refresh();
}

// Num to conditional className
function numCls(v, thresholds, goodLow) {
    // thresholds: [warn, bad] — values above warn get 'cell-warn', above bad get 'cell-bad'
    // goodLow=true means lower is better
    if(goodLow) {
        if(v>thresholds[1]) return 'cell-bad';
        if(v>thresholds[0]) return 'cell-warn';
        return 'cell-good';
    }
    if(v>thresholds[1]) return 'cell-good';
    if(v>thresholds[0]) return 'cell-warn';
    return 'cell-bad';
}

// ============ KPIs with Sparklines ============
function renderKPIs(data) {
    const ord=sum(data,'order_cnt'), rev=sum(data,'revenue'), gp=sum(data,'store_profit');
    const cc=sum(data,'commission_fee'), cp=sum(data,'commission_profit'), nc=sum(data,'neg_cnt'), pf=sum(data,'promo_fee');
    const df2=sum(data,'delivery_fee'), doc=sum(data,'delivery_order_cnt');
    const margin=rev>0?(cp/rev*100).toFixed(1):0;
    const negPct=ord>0?(nc/ord*100).toFixed(1):0;
    const aov=ord>0?(rev/ord).toFixed(1):0;
    const adc=doc>0?(df2/doc).toFixed(1):0;
    const ngCls=parseFloat(negPct)>25?'red':'green';

    // Get daily data for sparklines
    const byDate=groupBy(data,['日期']);
    byDate.sort((a,b)=>a.日期.localeCompare(b.日期));
    const spkOrd=byDate.map(d=>d.order_cnt||0);
    const spkRev=byDate.map(d=>d.revenue||0);
    const spkGp=byDate.map(d=>d.real_profit||0);
    const spkCp=byDate.map(d=>d.commission_profit||0);
    const spkPf=byDate.map(d=>d.promo_fee||0);
    const spkCc=byDate.map(d=>d.commission_fee||0);
    const spkNg=byDate.map(d=>d.order_cnt>0?(d.neg_cnt/d.order_cnt*100):0);
    const spkAdc=byDate.map(d=>(d.delivery_order_cnt||0)>0?(d.delivery_fee/d.delivery_order_cnt):0);
    const spkAov=byDate.map(d=>d.order_cnt>0?(d.revenue/d.order_cnt):0);
    const spkMarg=byDate.map(d=>d.revenue>0?(d.commission_profit/d.revenue*100):0);

    // 环比
    const prev = getPrevData(data);
    const pOrd=sum(prev,'order_cnt'), pRev=sum(prev,'revenue'), pGp=sum(prev,'store_profit');
    const pCc=sum(prev,'commission_fee'), pCp=sum(prev,'commission_profit'), pPf=sum(prev,'promo_fee');
    const pDf=sum(prev,'delivery_fee'), pDoc=sum(prev,'delivery_order_cnt');
    const pNeg=sum(prev,'neg_cnt');
    const pAov = pOrd > 0 ? (pRev / pOrd) : 0;
    const pAdc = pDoc > 0 ? (pDf / pDoc) : 0;
    const pNegPct = pOrd > 0 ? (pNeg / pOrd * 100) : 0;
    const pMargin = pRev > 0 ? (pCp / pRev * 100) : 0;

    const dOrd = deltaStr(ord, pOrd, false);
    const dRev = deltaStr(rev, pRev, false);
    const dGp = deltaStr(gp, pGp, false);
    const dPf = deltaStr(pf, pPf, false);
    const dCc = deltaStr(cc, pCc, false);
    const dCp = deltaStr(cp, pCp, false);
    const dMarg = deltaStr(parseFloat(margin), pMargin, false);
    const dAov = deltaStr(parseFloat(aov), pAov, false);
    const dNeg = deltaStr(parseFloat(negPct), pNegPct, false);
    const dAdc = deltaStr(parseFloat(adc), pAdc, false);

    document.getElementById('kpiGrid').innerHTML =
        '<div class="kpi-card clickable" data-kpi="order"><div class="kpi-label important">总单量 <span class="kpi-hint">📖<span class="kpi-tip">所有渠道订单总数</span></span></div><div class="kpi-value">'+ord.toLocaleString()+'</div><div class="kpi-mom" style="color:var(--'+dOrd.cls+')">环比 '+dOrd.txt+'</div><div class="kpi-sparkline" id="spkOrd"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="revenue"><div class="kpi-label">实收 <span class="kpi-hint">📖<span class="kpi-tip">客户实付金额之和</span></span></div><div class="kpi-value">'+fmtY(rev)+'</div><div class="kpi-mom" style="color:var(--'+dRev.cls+')">环比 '+dRev.txt+'</div><div class="kpi-sparkline" id="spkRev"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="aov"><div class="kpi-label">实收客单 <span class="kpi-hint">📖<span class="kpi-tip">实收÷总单量</span></span></div><div class="kpi-value">¥'+aov+'</div><div class="kpi-mom" style="color:var(--'+dAov.cls+')">环比 '+dAov.txt+'</div><div class="kpi-sparkline" id="spkAov"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="store_profit"><div class="kpi-label">门店毛利 <span class="kpi-hint">📖<span class="kpi-tip">线上毛利 - 推广费</span></span></div><div class="kpi-value">'+fmtY(gp)+'</div><div class="kpi-mom" style="color:var(--'+dGp.cls+')">环比 '+dGp.txt+'</div><div class="kpi-sparkline" id="spkGp"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="comm_profit"><div class="kpi-label important">抽佣毛利 <span class="kpi-hint">📖<span class="kpi-tip">门店毛利 - 公司抽佣</span></span></div><div class="kpi-value">'+fmtY(cp)+'</div><div class="kpi-mom" style="color:var(--'+dCp.cls+')">环比 '+dCp.txt+'</div><div class="kpi-sparkline" id="spkCp"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="promo"><div class="kpi-label">推广费 <span class="kpi-hint">📖<span class="kpi-tip">美团+饿了么推广消耗</span></span></div><div class="kpi-value">¥'+(pf.toFixed(0))+'</div><div class="kpi-mom" style="color:var(--'+dPf.cls+')">环比 '+dPf.txt+'</div><div class="kpi-sparkline" id="spkPf"></div></div>'+
        '<div class="kpi-card"><div class="kpi-label">公司抽佣 <span class="kpi-hint">📖<span class="kpi-tip">实收 × 门店抽佣点数</span></span></div><div class="kpi-value">'+fmtY(cc)+'</div><div class="kpi-mom" style="color:var(--'+dCc.cls+')">环比 '+dCc.txt+'</div><div class="kpi-sparkline" id="spkCc"></div></div>'+
        '<div class="kpi-card"><div class="kpi-label">毛利率 <span class="kpi-hint">📖<span class="kpi-tip">抽佣毛利 ÷ 实收</span></span></div><div class="kpi-value">'+margin+'%</div><div class="kpi-mom" style="color:var(--'+dMarg.cls+')">环比 '+dMarg.txt+'</div><div class="kpi-sparkline" id="spkMarg"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="neg"><div class="kpi-label">负毛利占比 <span class="kpi-hint">📖<span class="kpi-tip">线上毛利<0的单数÷总单数</span></span></div><div class="kpi-value">'+negPct+'%</div><div class="kpi-mom" style="color:var(--'+dNeg.cls+')">环比 '+dNeg.txt+'</div><div class="kpi-sparkline" id="spkNg"></div></div>'+
        '<div class="kpi-card clickable" data-kpi="delivery"><div class="kpi-label">单均配送 <span class="kpi-hint">📖<span class="kpi-tip">三方配送费÷配送单数</span></span></div><div class="kpi-value">¥'+adc+'</div><div class="kpi-mom" style="color:var(--'+dAdc.cls+')">环比 '+dAdc.txt+'</div><div class="kpi-sparkline" id="spkAdc"></div></div>';

    // Draw sparklines — all blue
    setTimeout(()=>{
        drawSparkline('spkOrd',spkOrd,'blue');
        drawSparkline('spkRev',spkRev,'blue');
        drawSparkline('spkAov',spkAov,'blue');
        drawSparkline('spkGp',spkGp,'blue');
        drawSparkline('spkCp',spkCp,'blue');
        drawSparkline('spkPf',spkPf,'blue');
        drawSparkline('spkCc',spkCc,'blue');
        drawSparkline('spkMarg',spkMarg,'blue');
        drawSparkline('spkNg',spkNg,'blue');
        drawSparkline('spkAdc',spkAdc,'blue');
    },10);
}

// ============ STORE TAB ============
function renderStore(data) {
    const stores = groupBy(data,['store_name']);
    stores.sort((a,b)=>b.revenue-a.revenue);
    const names=stores.map(s=>short(s.store_name));
    // 计算Y轴范围：0点对齐
    const maxOrd = Math.max(...stores.map(s=>s.order_cnt||0), 1);
    const pftVals = stores.map(s=>s.commission_profit||0);
    const minPft = Math.min(...pftVals, 0);
    const maxPft = Math.max(...pftVals, 1);
    const padO = maxOrd * 0.15, padP = Math.max((maxPft - minPft) * 0.15, 1);
    let leftRng, rightRng;
    if (minPft >= 0) {
        leftRng = [0, maxOrd + padO]; rightRng = [0, maxPft + padP];
    } else {
        const rMin = minPft - padP, rMax = maxPft + padP;
        rightRng = [rMin, rMax];
        const zeroRatio = Math.abs(rMin) / (rMax - rMin);
        const leftNeg = (zeroRatio * (maxOrd + padO)) / (1 - zeroRatio);
        leftRng = [-leftNeg, maxOrd + padO];
    }

    Plotly.newPlot('chartStoreBar',[
        {x:names,y:stores.map(s=>s.order_cnt||0),name:'单量',type:'bar',marker:{color:'#5B8FF9'},offsetgroup:0},
        {x:names,y:stores.map(s=>s.commission_profit||0),name:'抽佣毛利',type:'bar',marker:{color:'#FF9F43'},yaxis:'y2',offsetgroup:1}
    ],{...plotlyLayout,barmode:'group',
        xaxis:{...plotlyLayout.xaxis,tickangle:-45},
        yaxis:{title:'单量',gridcolor:'rgba(0,0,0,0)',zerolinecolor:'#e5e7eb',range:leftRng},
        yaxis2:{title:'抽佣毛利(元)',overlaying:'y',side:'right',gridcolor:'rgba(0,0,0,0)',range:rightRng},
        margin:{l:50,r:60,t:20,b:100}
    },plotlyCfg);

    // Get prev period data for MoM comparison — all metrics
    const prevStoreData = getPrevData(data);
    const prevStoreMap = {};
    prevStoreData.forEach(r => {
        const sn = r.store_name || '';
        if (!prevStoreMap[sn]) prevStoreMap[sn] = {ord:0, rev:0, rp:0, cf:0, cp:0, neg:0, doc:0, df:0, promo:0};
        prevStoreMap[sn].ord += r.order_cnt||0;
        prevStoreMap[sn].rev += r.revenue||0;
        prevStoreMap[sn].rp += r.store_profit||0;
        prevStoreMap[sn].cf += r.commission_fee||0;
        prevStoreMap[sn].cp += r.commission_profit||0;
        prevStoreMap[sn].neg += r.neg_cnt||0;
        prevStoreMap[sn].doc += r.delivery_order_cnt||0;
        prevStoreMap[sn].df += r.delivery_fee||0;
        prevStoreMap[sn].promo += r.promo_fee||0;
    });

    function momPct(cur, prev) {
        if (prev===0) return '--';
        return ((cur-prev)/Math.abs(prev)*100);
    }

    // Compute derived values for table — all metrics with MoM
    // 按省市排序
    stores.sort((a,b) => {
        let pa = storeProvince[a.store_name] || '其他', pb = storeProvince[b.store_name] || '其他';
        if (pa !== pb) {
            const po = ['广东省','湖南省','广西壮族自治区','福建省','江西省','海南省','湖北省','其他'];
            return po.indexOf(pa) - po.indexOf(pb);
        }
        return (a.store_name||'').localeCompare(b.store_name||'');
    });

    const rows=stores.map(s=>{
        const rev=s.revenue||0, sp=s.store_profit||0, rp=s.real_profit||0, cf=s.commission_fee||0, cp=s.commission_profit||0;
        const ord=s.order_cnt||0, aov=ord>0?(rev/ord):0;
        const delCost=(s.delivery_order_cnt||0)>0?(s.delivery_fee/s.delivery_order_cnt):0;
        const margin=rev>0?(cp/rev*100):0;
        const negPct=ord>0?((s.neg_cnt||0)/ord*100):0;
        const promo=s.promo_fee||0;
        const avgProfit=ord>0?(cp/ord):0;  // 单均毛利
        const p = prevStoreMap[s.store_name] || {ord:0, rev:0, rp:0, cf:0, cp:0, neg:0, doc:0, df:0, promo:0};
        const pAov = p.ord>0?(p.rev/p.ord):0;
        const pDelCost = p.doc>0?(p.df/p.doc):0;
        const pMargin = p.rev>0?(p.cp/p.rev*100):0;
        const pNegPct = p.ord>0?((p.neg||0)/p.ord*100):0;
        return {
            name:short(s.store_name),
            full: s.store_name,
            qn: storeQnMap[s.store_name] || '',
            mt: storeMeituanMap[s.store_name] || '',
            prov: storeProvince[s.store_name] || '其他',
            ord, rev, gross:sp, comm:cf, net:cp,
            promo, avgProfit,
            margin, aov, delCost, negPct,
            mom: {
                ord: momPct(ord, p.ord),
                rev: momPct(rev, p.rev),
                gross: momPct(sp, p.rp),
                comm: momPct(cf, p.cf),
                net: momPct(cp, p.cp),
                promo: momPct(promo, p.promo||0),
                avgProfit: momPct(avgProfit, p.ord>0?(p.cp/p.ord):0),
                margin: momPct(margin, pMargin),
                aov: momPct(aov, pAov),
                delCost: momPct(delCost, pDelCost),
                negPct: momPct(negPct, pNegPct)
            }
        };
    });

    // Map column index → mom key (更新为新的列顺序)
    const momKeyMap = {2:'ord',3:'rev',4:'gross',5:'promo',6:'net',7:'avgProfit',8:'comm',9:'margin',10:'aov',11:'delCost',12:'negPct'};

    function getMomDisplay(r) {
        const state = tableSortState['storeTable'] || {col:-1};
        const mk = momKeyMap[state.col] || 'net';
        const v = r.mom[mk];
        if (v==='--') return '--';
        const cls = v>0 ? 'cell-bad' : (v<0 ? 'cell-good' : '');
        return '<span class="'+cls+'">'+(v>0?'↑':v<0?'↓':'')+Math.abs(v).toFixed(1)+'%</span>';
    }

    makeSortable('storeTable', rows, [
        {key:'idx',label:'序号',type:'num',sortable:false},
        {key:'name',label:'门店名称',type:'str'},
        {key:'ord',label:'订单量',type:'num'},
        {key:'rev',label:'实收',type:'num'},
        {key:'gross',label:'门店毛利',type:'num'},
        {key:'promo',label:'推广金额',type:'num'},
        {key:'net',label:'抽佣毛利',type:'num'},
        {key:'avgProfit',label:'单均毛利',type:'num'},
        {key:'comm',label:'公司抽佣',type:'num'},
        {key:'margin',label:'抽佣毛利率',type:'pct'},
        {key:'aov',label:'实收客单价',type:'num'},
        {key:'delCost',label:'平均配送成本',type:'num'},
        {key:'negPct',label:'负毛利占比',type:'pct'},
        {key:'mom_net',label:'环比',type:'pct'}
    ], function(r,i){
        var marginCls = parseFloat(r.margin)<10?'cell-bad':'';
        var negCls = parseFloat(r.negPct)>35?'cell-bad':'';
        return '<tr data-qn="'+r.qn+'" data-meituan="'+r.mt+'">'
            +'<td>'+(i+1)+'</td>'
            +'<td>'+r.name+'</td>'
            +'<td>'+r.ord.toLocaleString()+'</td>'
            +'<td>'+fmtY(r.rev)+'</td>'
            +'<td>'+fmtY(r.gross)+'</td>'
            +'<td>'+fmtY(r.promo)+'</td>'
            +'<td>'+fmtY(r.net)+'</td>'
            +'<td>¥'+r.avgProfit.toFixed(2)+'</td>'
            +'<td>'+fmtY(r.comm)+'</td>'
            +'<td class="'+marginCls+'">'+r.margin.toFixed(1)+'%</td>'
            +'<td>¥'+r.aov.toFixed(1)+'</td>'
            +'<td>¥'+r.delCost.toFixed(1)+'</td>'
            +'<td class="'+negCls+'">'+r.negPct.toFixed(1)+'%</td>'
            +'<td>'+getMomDisplay(r)+'</td>'
            +'</tr>';
    });
}

function negBadge(np) {
    if(np>35) return '<span class="badge badge-warn">'+np.toFixed(1)+'%</span>';
    if(np>15) return '<span class="anomaly-tag info">'+np.toFixed(1)+'%</span>';
    return '<span class="badge badge-ok">'+np.toFixed(1)+'%</span>';
}

// ============ CHANNEL TAB ============
function renderChannel(data) {
    const ch = groupBy(data,['channel']);
    const to = ch.reduce((s,c)=>s+c.order_cnt,0);
    ch.forEach(c=>{c.pct=to>0?(c.order_cnt/to*100).toFixed(1):'0.0';});

    // Channel color map
    const chColors={'美团闪购':'#FFD100','饿了么':'#00AAFF','京东到家':'#E86452','线下':'#5B8FF9'};
    const chOrder=['美团闪购','饿了么','京东到家','线下'];
    const chSorted=[...ch].sort((a,b)=>chOrder.indexOf(a.channel)-chOrder.indexOf(b.channel));

    Plotly.newPlot('chartChannelPie',[{values:chSorted.map(c=>c.order_cnt),labels:chSorted.map(c=>c.channel),
        type:'pie',hole:0.45,
        marker:{colors:chSorted.map(c=>chColors[c.channel]||'#5B8FF9')},
        textinfo:'label+percent',textposition:'outside'}],
        {...plotlyLayout,height:chartHS,showlegend:false},plotlyCfg);

    const chRev=[...ch].sort((a,b)=>b.revenue-a.revenue);
    Plotly.newPlot('chartChannelRev',[{y:chRev.map(c=>c.channel),x:chRev.map(c=>c.revenue),
        type:'bar',orientation:'h',
        marker:{color:chRev.map(c=>chColors[c.channel]||'#5B8FF9')}}],
        {...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'实收(元)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    const chComm=[...ch].sort((a,b)=>b.commission_fee-a.commission_fee);
    Plotly.newPlot('chartChannelComm',[{y:chComm.map(c=>c.channel),x:chComm.map(c=>c.commission_fee),
        type:'bar',orientation:'h',
        marker:{color:chComm.map(c=>chColors[c.channel]||'#FF9F43')}}],
        {...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'公司抽佣(元)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    const chMarg=[...ch].sort((a,b)=>(b.commission_profit/(b.revenue||1))-(a.commission_profit/(a.revenue||1)));
    Plotly.newPlot('chartChannelMargin',[{y:chMarg.map(c=>c.channel),x:chMarg.map(c=>c.revenue>0?(c.commission_profit/c.revenue*100).toFixed(1):0),
        type:'bar',orientation:'h',
        marker:{color:chMarg.map(c=>chColors[c.channel]||'#F6BD16')}}],
        {...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'毛利率(%)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    // Channel badge helper
    function chBadge(name) {
        const cls={'美团闪购':'meituan','饿了么':'eleme','京东到家':'jd','线下':'pos'}[name]||'pos';
        return '<span class="channel-badge '+cls+'">'+name+'</span>';
    }

    const rows=ch.map(c=>{
        const rev=c.revenue||0, rp=c.real_profit||0, cf=c.commission_fee||0, cp=c.commission_profit||0;
        return {
            name:c.channel,
            ord:c.order_cnt||0,
            pct:parseFloat(c.pct)||0,
            rev:rev,
            gross:sp,
            comm:cf,
            net:cp,
            margin:rev>0?(cp/rev*100):0,
            negPct:c.order_cnt>0?((c.neg_cnt||0)/c.order_cnt*100):0
        };
    });

    makeSortable('channelTable', rows, [
        {key:'name',label:'渠道',type:'str',sortable:false},
        {key:'ord',label:'单量',type:'num'},
        {key:'pct',label:'占比(%)',type:'pct'},
        {key:'rev',label:'实收',type:'num'},
        {key:'gross',label:'门店毛利',type:'num'},
        {key:'comm',label:'公司抽佣',type:'num'},
        {key:'net',label:'抽佣毛利',type:'num'},
        {key:'margin',label:'毛利率(%)',type:'pct'},
        {key:'negPct',label:'负毛利占比',type:'pct'}
    ], function(r){
        return '<tr>'
            +'<td>'+chBadge(r.name)+'</td>'
            +'<td>'+r.ord.toLocaleString()+'</td>'
            +'<td>'+r.pct.toFixed(1)+'%</td>'
            +'<td>'+fmtY(r.rev)+'</td>'
            +'<td><span class="'+numCls(r.gross,[0,500],false)+'">'+fmtY(r.gross)+'</span></td>'
            +'<td class="cell-bad">'+fmtY(r.comm)+'</td>'
            +'<td><span class="'+numCls(r.net,[0,500],false)+'">'+fmtY(r.net)+'</span></td>'
            +'<td><span class="'+numCls(r.margin,[5,15],false)+'">'+r.margin.toFixed(1)+'%</span></td>'
            +'<td>'+negBadge(r.negPct)+'</td>'
            +'</tr>';
    });
}

// ============ NEG TAB ============
function renderNeg(data) {
    const sS=new Set(),cS=new Set(),mx={};
    data.forEach(r=>{sS.add(r.store_name);cS.add(r.channel);const k=r.store_name+'||'+r.channel;if(!mx[k])mx[k]={s:0,c:0};mx[k].s+=r.order_cnt>0?(r.neg_cnt/r.order_cnt*100):0;mx[k].c+=1;});
    const sL=[...sS].sort(),cL=[...cS].sort();
    const z=sL.map(s=>cL.map(c=>{const v=mx[s+'||'+c];return v?Math.round(v.s/v.c*10)/10:0;}));
    Plotly.newPlot('chartNegHeatmap',[{z:z,x:cL,y:sL.map(s=>short(s)),type:'heatmap',colorscale:[[0,'#5AD8A6'],[0.5,'#F6BD16'],[1,'#E86452']],hovertemplate:'%{y} | %{x}<br>负毛利占比: %{z}%<extra></extra>'}],{...plotlyLayout,height:Math.max(350,sL.length*26),xaxis:{...plotlyLayout.xaxis,side:'top'},margin:{l:100,r:20,t:50,b:30}},plotlyCfg);

    const sd=groupBy(data,['store_name']);
    sd.forEach(s=>{s.np=s.order_cnt>0?(s.neg_cnt/s.order_cnt*100):0;});
    sd.sort((a,b)=>b.np-a.np);
    let h='<table><tr><th>排名</th><th>门店</th><th>总单量</th><th>负毛利单数</th><th>负毛利占比</th><th>门店毛利</th></tr>';
    sd.slice(0,10).forEach((s,i)=>{
        h+='<tr><td>'+(i+1)+'</td><td>'+short(s.store_name)+'</td><td>'+s.order_cnt.toLocaleString()+'</td><td>'+(s.neg_cnt||0)+'</td><td><span class="badge badge-'+(s.np>25?'warn':'ok')+'">'+s.np.toFixed(1)+'%</span></td><td>'+fmtY(s.real_profit||0)+'</td></tr>';
    });
    h+='</table>';
    document.getElementById('negTable').innerHTML=h;
}

// ============ PROMO TAB ============
function renderPromo(pd) {
    if(!pd.length){document.getElementById('chartPromoBar').innerHTML='<p style="color:var(--text-dim);padding:30px;text-align:center;">暂无推广数据</p>';document.getElementById('promoTable').innerHTML='';return;}
    const sm={};pd.forEach(r=>{sm[r.store_name]=(sm[r.store_name]||0)+(r.promo_fee||0);});
    const so=Object.entries(sm).sort((a,b)=>b[1]-a[1]).slice(0,15);
    Plotly.newPlot('chartPromoBar',[{y:so.map(s=>short(s[0])),x:so.map(s=>s[1]),type:'bar',orientation:'h',marker:{color:so.map(s=>s[1]>100?'#E86452':s[1]>50?'#F6BD16':'#5AD8A6')}}],{...plotlyLayout,height:Math.max(340,so.length*26),xaxis:{...plotlyLayout.xaxis,title:'推广费(元)'},margin:{l:100,r:20,t:10,b:40}},plotlyCfg);
    const cm={},cs={};
    pd.forEach(r=>{cm[r.channel]=(cm[r.channel]||0)+(r.promo_fee||0);if(!cs[r.channel])cs[r.channel]=new Set();cs[r.channel].add(r.store_name);});
    let h='<table><tr><th>渠道</th><th>推广费合计</th><th>投放门店数</th></tr>';
    Object.entries(cm).sort((a,b)=>b[1]-a[1]).forEach(([ch,fee])=>{h+='<tr><td>'+ch+'</td><td>¥'+fee.toFixed(2)+'</td><td>'+cs[ch].size+'</td></tr>';});
    h+='</table>';
    document.getElementById('promoTable').innerHTML=h;
}

// ============ TIME ANALYSIS TAB ============
function renderTimeAnalysis(data) {
    const byDate = groupBy(data, ['日期']);
    byDate.sort((a,b) => a.日期.localeCompare(b.日期));
    const dates = byDate.map(d => d.日期);
    const orders = byDate.map(d => d.order_cnt || 0);
    const profits = byDate.map(d => d.commission_profit || 0);

    // Linear regression trend lines
    const n = dates.length;
    function trendLine(vals) {
        if (n < 2) return { x: dates, y: vals.map(() => null) };
        let sx=0,sy=0,sxy=0,sxx=0;
        for (let i=0;i<n;i++) { sx+=i; sy+=vals[i]; sxy+=i*vals[i]; sxx+=i*i; }
        const slope = (n*sxy - sx*sy) / (n*sxx - sx*sx);
        const intercept = (sy - slope*sx) / n;
        return { x: dates, y: vals.map((_,i) => Math.round((slope*i+intercept)*100)/100) };
    }
    const tOrders = trendLine(orders);
    const tProfits = trendLine(profits);

    // 计算范围：0点对齐（有负毛利时左轴也往下移）
    const maxO = Math.max(...orders, 1);
    const minP = Math.min(...profits, 0);
    const maxP = Math.max(...profits, 1);
    const padOT = maxO * 0.15, padPT = Math.max((maxP - minP) * 0.15, 1);
    let leftRngT, rightRngT;
    if (minP >= 0) {
        leftRngT = [0, maxO + padOT]; rightRngT = [0, maxP + padPT];
    } else {
        const rMin = minP - padPT, rMax = maxP + padPT;
        rightRngT = [rMin, rMax];
        const zeroRatio = Math.abs(rMin) / (rMax - rMin);
        const leftNeg = (zeroRatio * (maxO + padOT)) / (1 - zeroRatio);
        leftRngT = [-leftNeg, maxO + padOT];
    }
    // 双Y轴并列：蓝=单量(左轴)，橙=抽佣毛利+虚线(右轴)
    Plotly.newPlot('chartTimeTrend', [
        { x: dates, y: orders, name: '单量', type: 'bar', marker: {color:'#5B8FF9'}, offsetgroup: 0 },
        { x: dates, y: profits, name: '抽佣毛利', type: 'bar', marker: {color:'#FF9F43'}, yaxis:'y2', offsetgroup: 1 },
        { x: tProfits.x, y: tProfits.y, name: '毛利趋势', type: 'scatter', mode: 'lines', line: {color:'#FF9F43',width:2,dash:'dot'}, yaxis:'y2', showlegend: true }
    ], {
        ...plotlyLayout, height: chartH + 40, barmode: 'group',
        xaxis: { ...plotlyLayout.xaxis, title: '日期', type: 'date', tickformat: '%m-%d', dtick: 'D1' },
        yaxis: { title: '单量', gridcolor: 'rgba(0,0,0,0)', zerolinecolor: '#2a2d3a', range: leftRngT },
        yaxis2: { title: '抽佣毛利(元)', overlaying:'y', side:'right', gridcolor: 'rgba(0,0,0,0)', range: rightRngT },
        legend: { orientation: 'h', y: 1.1, font: {color:'#1f2937',size:11} },
        margin: { l: 50, r: 60, t: 20, b: 50 }
    }, plotlyCfg);

    // 每日明细表 - sortable
    const tRows = byDate.map(d=>{
        const rev=d.revenue||0, rp=d.real_profit||0, cf=d.commission_fee||0, cp=d.commission_profit||0;
        return {
            date:d.日期,
            ord:d.order_cnt||0,
            rev:rev,
            gross:sp,
            comm:cf,
            net:cp,
            margin:rev>0?(cp/rev*100):0
        };
    });
    makeSortable('timeTable', tRows, [
        {key:'date',label:'日期',type:'str'},
        {key:'ord',label:'单量',type:'num'},
        {key:'rev',label:'实收',type:'num'},
        {key:'gross',label:'门店毛利',type:'num'},
        {key:'comm',label:'抽佣',type:'num'},
        {key:'net',label:'抽佣毛利',type:'num'},
        {key:'margin',label:'毛利率',type:'pct'}
    ], function(r){
        return '<tr><td>'+r.date+'</td><td>'+r.ord.toLocaleString()+'</td><td>'+fmtY(r.rev)+'</td><td>'+fmtY(r.gross)+'</td><td>'+fmtY(r.comm)+'</td><td>'+fmtY(r.net)+'</td><td>'+r.margin.toFixed(1)+'%</td></tr>';
    });
}

// ============ NEG TAB ============
function renderNeg(data) {
    const byDate = groupBy(data, ['日期']);
    byDate.sort((a, b) => a.日期.localeCompare(b.日期));
    const dates = byDate.map(d => d.日期);
    const negPcts = byDate.map(d => d.order_cnt > 0 ? ((d.neg_cnt || 0) / d.order_cnt * 100) : 0);
    const maxP = Math.max(...negPcts, 1);

    // Trend line
    function trendLineV(vals, n) {
        if (n < 2) return { x: dates, y: vals.map(() => null) };
        let sx = 0, sy = 0, sxy = 0, sxx = 0;
        for (let i = 0; i < n; i++) { sx += i; sy += vals[i]; sxy += i * vals[i]; sxx += i * i; }
        const slope = (n * sxy - sx * sy) / (n * sxx - sx * sx);
        const intercept = (sy - slope * sx) / n;
        return { x: dates, y: vals.map((_, i) => Math.round((slope * i + intercept) * 100) / 100) };
    }
    const tNeg = trendLineV(negPcts, dates.length);

    // Bar + trend (single Y axis)
    Plotly.newPlot('chartNegStore', [
        { x: dates, y: negPcts, name: '负毛利占比', type: 'bar', marker: { color: negPcts.map(v => v > 30 ? '#E86452' : v > 15 ? '#F6BD16' : '#5AD8A6') } },
        { x: tNeg.x, y: tNeg.y, name: '趋势', type: 'scatter', mode: 'lines', line: { color: '#E86452', width: 2, dash: 'dot' }, showlegend: true }
    ], {
        ...plotlyLayout, height: chartH + 40, barmode: 'group',
        xaxis: { ...plotlyLayout.xaxis, title: '日期', type: 'date', tickformat: '%m-%d', dtick: 'D1', gridcolor: 'rgba(0,0,0,0)' },
        yaxis: { title: '负毛利占比(%)', ticksuffix: '%', gridcolor: 'rgba(0,0,0,0)', zerolinecolor: '#2a2d3a', range: [0, maxP * 1.2] },
        legend: { orientation: 'h', y: 1.1, font: { color: '#1f2937', size: 11 } },
        margin: { l: 60, r: 20, t: 20, b: 50 }
    }, plotlyCfg);

    // 门店负毛利占比 (horizontal bar)
    const storeNeg = groupBy(data, ['store_name']);
    const sNegs = storeNeg.filter(s => s.order_cnt > 0)
        .map(s => ({ name: short(s.store_name), np: (s.neg_cnt || 0) / s.order_cnt * 100 }))
        .sort((a, b) => b.np - a.np);
    Plotly.newPlot('chartNegByStore', [
        { y: sNegs.map(s => s.name), x: sNegs.map(s => s.np), type: 'bar', orientation: 'h',
          marker: { color: sNegs.map(s => s.np > 30 ? '#E86452' : s.np > 15 ? '#F6BD16' : '#5AD8A6') } }
    ], {
        ...plotlyLayout, height: Math.max(300, sNegs.length * 28),
        xaxis: { title: '负毛利占比(%)', ticksuffix: '%', gridcolor: 'rgba(0,0,0,0)', range: [0, Math.max(...sNegs.map(s=>s.np), 1) * 1.2] },
        margin: { l: 120, r: 20, t: 10, b: 40 }
    }, plotlyCfg);

    // Table - sortable
    const negRows = byDate.map(d => ({
        date: d.日期,
        ord: d.order_cnt || 0,
        neg: d.neg_cnt || 0,
        np: d.order_cnt > 0 ? ((d.neg_cnt || 0) / d.order_cnt * 100) : 0,
        gross: (d.store_profit || 0),
        net: d.commission_profit || 0
    }));
    makeSortable('negTable', negRows, [
        {key:'date',label:'日期',type:'str'},
        {key:'ord',label:'总单量',type:'num'},
        {key:'neg',label:'负毛利单',type:'num'},
        {key:'np',label:'负毛利占比',type:'pct'},
        {key:'gross',label:'门店毛利',type:'num'},
        {key:'net',label:'抽佣毛利',type:'num'}
    ], function(r){
        let cls = r.np > 30 ? 'warn' : r.np > 15 ? '' : '';
        return '<tr><td>'+r.date+'</td><td>'+r.ord.toLocaleString()+'</td><td>'+r.neg+'</td>'
            +'<td>'+negBadge(r.np)+'</td>'
            +'<td>'+fmtY(r.gross)+'</td><td>'+fmtY(r.net)+'</td></tr>';
    });
}

// ============ DELIVERY TAB ============
function renderDelivery(data) {
    const byDate = groupBy(data, ['日期']);
    byDate.sort((a, b) => a.日期.localeCompare(b.日期));
    const dates = byDate.map(d => d.日期);
    const avgCosts = byDate.map(d => d.delivery_order_cnt > 0 ? (d.delivery_fee / d.delivery_order_cnt) : 0);
    const maxC = Math.max(...avgCosts, 1);
    const n = dates.length;

    function trendLineV(vals) {
        if (n < 2) return { x: dates, y: vals.map(() => null) };
        let sx = 0, sy = 0, sxy = 0, sxx = 0;
        for (let i = 0; i < n; i++) { sx += i; sy += vals[i]; sxy += i * vals[i]; sxx += i * i; }
        const slope = (n * sxy - sx * sy) / (n * sxx - sx * sx);
        const intercept = (sy - slope * sx) / n;
        return { x: dates, y: vals.map((_, i) => Math.round((slope * i + intercept) * 100) / 100) };
    }
    const tCost = trendLineV(avgCosts);

    // 单均配送成本 + 趋势
    Plotly.newPlot('chartDelCost', [
        { x: dates, y: avgCosts, name: '单均配送', type: 'bar', marker: { color: '#FF9F43' } },
        { x: tCost.x, y: tCost.y, name: '趋势', type: 'scatter', mode: 'lines', line: { color: '#FF9F43', width: 2, dash: 'dot' }, showlegend: true }
    ], {
        ...plotlyLayout, height: chartH + 40, barmode: 'group',
        xaxis: { ...plotlyLayout.xaxis, title: '日期', type: 'date', tickformat: '%m-%d', dtick: 'D1', gridcolor: 'rgba(0,0,0,0)' },
        yaxis: { title: '单均配送(¥)', ticksuffix: '¥', gridcolor: 'rgba(0,0,0,0)', zerolinecolor: '#2a2d3a', range: [0, maxC * 1.2] },
        legend: { orientation: 'h', y: 1.1, font: { color: '#1f2937', size: 11 } },
        margin: { l: 60, r: 20, t: 20, b: 50 }
    }, plotlyCfg);

    // 门店单均配送成本 (bar)
    const storeDel = groupBy(data, ['store_name']);
    const sDels = storeDel.filter(s => s.delivery_order_cnt > 0)
        .map(s => ({ name: short(s.store_name), avg: s.delivery_fee / s.delivery_order_cnt }))
        .sort((a, b) => b.avg - a.avg);
    Plotly.newPlot('chartDelRatio', [
        { y: sDels.map(s => s.name), x: sDels.map(s => s.avg), type: 'bar', orientation: 'h',
          marker: { color: '#FF9F43' } }
    ], {
        ...plotlyLayout, height: Math.max(300, sDels.length * 28),
        xaxis: { title: '单均配送(¥)', ticksuffix: '¥', gridcolor: 'rgba(0,0,0,0)' },
        margin: { l: 120, r: 20, t: 10, b: 40 }
    }, plotlyCfg);

    // Table - sortable
    const delRows = byDate.map(d => ({
        date: d.日期,
        ord: d.order_cnt || 0,
        delCnt: d.delivery_order_cnt || 0,
        ratio: d.order_cnt > 0 ? (d.delivery_order_cnt / d.order_cnt * 100) : 0,
        fee: d.delivery_fee || 0,
        avg: d.delivery_order_cnt > 0 ? (d.delivery_fee / d.delivery_order_cnt) : 0
    }));
    makeSortable('delTable', delRows, [
        {key:'date',label:'日期',type:'str'},
        {key:'ord',label:'总单量',type:'num'},
        {key:'delCnt',label:'配送单数',type:'num'},
        {key:'ratio',label:'配送占比',type:'pct'},
        {key:'fee',label:'总配送费',type:'num'},
        {key:'avg',label:'单均配送',type:'num'}
    ], function(r){
        const avgCls = r.avg > 10 ? 'cell-bad' : r.avg > 6 ? 'cell-warn' : 'cell-good';
        return '<tr><td>'+r.date+'</td><td>'+r.ord.toLocaleString()+'</td><td>'+r.delCnt.toLocaleString()
            +'</td><td>'+r.ratio.toFixed(1)+'%</td><td>'+fmtY(r.fee)+'</td><td class="'+avgCls+'">¥'+r.avg.toFixed(2)+'</td></tr>';
    });
}

// ============ PAGE SWITCHER ============
function switchPage(name, btn) {
    document.querySelectorAll('.page-btn').forEach(b => { b.classList.remove('active'); b.style.background = 'var(--bg)'; b.style.color = 'var(--text-dim)'; });
    btn.classList.add('active'); btn.style.background = 'var(--card)'; btn.style.color = 'var(--text)';
    document.getElementById('page-ops').style.display = name === 'ops' ? 'block' : 'none';
    document.getElementById('page-product').style.display = name === 'product' ? 'block' : 'none';
    if (name === 'product') renderProduct();
}
function switchProdTab(name, btn) {
    document.querySelectorAll('#prod-tab-'+name.split('-')[0] +' .tab-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('prod-tab-turnover').classList.toggle('active', name === 'turnover');
    document.getElementById('prod-tab-sales').classList.toggle('active', name === 'sales');
}

// ============ PRODUCT RENDERING ============
let selectedProdPeriod = productPeriods.length > 0 ? productPeriods[productPeriods.length - 1] : '202605';

function initProductPeriodChips() {
    if (!productPeriods.length) return;
    const container = document.getElementById('prodPeriodChips');
    container.innerHTML = '';
    productPeriods.forEach(p => {
        const chip = document.createElement('span');
        chip.className = 'chip' + (p === selectedProdPeriod ? ' active' : '');
        const label = p.length >= 6 ? p.substring(0,4) + '-' + p.substring(4,6) : p;
        chip.textContent = label;
        chip.onclick = function() {
            selectedProdPeriod = p;
            initProductPeriodChips();
            renderProduct();
        };
        container.appendChild(chip);
    });
}

function getPeriodData(period) {
    return allProductData.filter(r => String(r.period) === String(period));
}
function getPrevPeriod(period) {
    const idx = productPeriods.indexOf(String(period));
    return idx > 0 ? productPeriods[idx - 1] : null;
}

function renderProduct() {
    const pData = getPeriodData(selectedProdPeriod);
    if (!pData.length) { document.getElementById('prodKpiGrid').innerHTML = '<p style="color:var(--text-dim);padding:20px;">暂无商品数据</p>'; return; }

    // Init period chips
    initProductPeriodChips();

    const all = pData.filter(r => r.store_name === '全部门店');
    const catOnly = all.filter(r => ['食品','日化','百货'].includes(r.cat));
    const food = all.filter(r => r.cat === '食品');
    const daily = all.filter(r => r.cat === '日化');
    const goods = all.filter(r => r.cat === '百货');

    function sum1(arr, k) { return arr.reduce((s, r) => s + (r[k] || 0), 0); }

    const tSku = Math.round(sum1(catOnly, 'sku_total'));
    const tTurn = sum1(catOnly, 'sku_total') > 0 ? (sum1(catOnly, 'sku_active') / sum1(catOnly, 'sku_total') * 100) : 0;
    const tVal = sum1(catOnly, 'goods_value');
    const tRev = sum1(catOnly, 'revenue');
    const tProf = sum1(catOnly, 'profit');
    const tMarg = tRev > 0 ? (tProf / tRev * 100) : 0;

    // Period-over-period: compare with previous period
    const prev = getPrevPeriod(selectedProdPeriod);
    const prevData = prev ? getPeriodData(prev) : [];
    const prevAll = prevData.filter(r => r.store_name === '全部门店');
    const prevCat = prevAll.filter(r => ['食品','日化','百货'].includes(r.cat));
    const pSku = Math.round(sum1(prevCat, 'sku_total'));
    const pTurn = sum1(prevCat, 'sku_total') > 0 ? (sum1(prevCat, 'sku_active') / sum1(prevCat, 'sku_total') * 100) : 0;
    const pVal = sum1(prevCat, 'goods_value');
    const pProf = sum1(prevCat, 'profit');
    const pMarg = sum1(prevCat, 'revenue') > 0 ? (sum1(prevCat, 'profit') / sum1(prevCat, 'revenue') * 100) : 0;

    function momStr(cur, pv, unit) {
        if (pv === 0) return '<div class="kpi-sub" style="color:var(--text-dim);">--</div>';
        const chg = ((cur - pv) / pv * 100).toFixed(1);
        const cls = chg > 0 ? 'green' : 'red';
        const sign = chg > 0 ? '↑' : '↓';
        return '<div class="kpi-sub" style="color:var(--'+cls+');">环比 ' + sign + Math.abs(chg) + '%</div>';
    }

    document.getElementById('prodKpiGrid').innerHTML =
        '<div class="kpi-card blue"><div class="kpi-label">SKU总数</div><div class="kpi-value">' + tSku.toLocaleString() + '</div>'+momStr(tSku, pSku)+'</div>' +
        '<div class="kpi-card accent"><div class="kpi-label">动销率</div><div class="kpi-value">' + tTurn.toFixed(1) + '%</div>'+momStr(tTurn, pTurn)+'</div>' +
        '<div class="kpi-card orange"><div class="kpi-label">货值</div><div class="kpi-value">' + fmtY(tVal) + '</div>'+momStr(tVal, pVal)+'</div>' +
        '<div class="kpi-card green"><div class="kpi-label">毛利</div><div class="kpi-value">' + fmtY(tProf) + '</div>'+momStr(tProf, pProf)+'</div>' +
        '<div class="kpi-card yellow"><div class="kpi-label">毛利率</div><div class="kpi-value">' + tMarg.toFixed(1) + '%</div>'+momStr(tMarg, pMarg)+'</div>';

    // Turnover chart
    const cats = [food, daily, goods].map((g, i) => ({
        name: ['食品', '日化', '百货'][i],
        turnover: g.length > 0 ? (sum1(g, 'sku_active') / sum1(g, 'sku_total') * 100) : 0,
        value: sum1(g, 'goods_value'),
        margin: sum1(g, 'revenue') > 0 ? (sum1(g, 'profit') / sum1(g, 'revenue') * 100) : 0
    }));
    Plotly.newPlot('chartProdCatTurnover', [
        { y: cats.map(c => c.name), x: cats.map(c => c.turnover), name: '动销率',
          type: 'bar', orientation: 'h', marker: { color: '#5B8FF9' },
          text: cats.map(c => c.turnover.toFixed(1) + '%'), textposition: 'outside', textfont: { color: '#1f2937', size: 11 } },
        { y: cats.map(c => c.name), x: cats.map(c => c.margin), name: '毛利率',
          type: 'bar', orientation: 'h', marker: { color: '#FF9F43' },
          text: cats.map(c => c.margin.toFixed(1) + '%'), textposition: 'outside', textfont: { color: '#1f2937', size: 11 } }
    ], {
        ...plotlyLayout, barmode: 'group', height: 250,
        xaxis: { ...plotlyLayout.xaxis, title: '百分比(%)', ticksuffix: '%', gridcolor: 'rgba(0,0,0,0)', range: [0, 100] },
        yaxis: { gridcolor: 'rgba(0,0,0,0)' },
        legend: { orientation: 'h', y: 1.15, x: 0.5, xanchor: 'center' },
        margin: { l: 60, r: 20, t: 10, b: 40 }
    }, plotlyCfg);

    // Store turnover ranking
    const storeMap = {};
    pData.forEach(r => {
        if (r.store_name === '全部门店') return;
        if (r.cat === '总计') return;  // skip 总计 rows to avoid double-counting
        if (!storeMap[r.store_name]) storeMap[r.store_name] = { n: 0, a: 0, v: 0 };
        storeMap[r.store_name].n += r.sku_total || 0;
        storeMap[r.store_name].a += r.sku_active || 0;
        storeMap[r.store_name].v += r.goods_value || 0;
    });
    const sList = Object.entries(storeMap).map(([n, d]) => ({
        name: n && n.includes('（') ? n.split('（')[1].replace('）', '') : (n || '').split('-').pop() || n,
        turnover: d.n > 0 ? (d.a / d.n * 100) : 0,
        value: d.v
    })).filter(s => s.turnover > 0).sort((a, b) => b.turnover - a.turnover);

    Plotly.newPlot('chartProdStoreTurnover', [{
        y: sList.map(s => s.name),
        x: sList.map(s => s.turnover),
        type: 'bar', orientation: 'h',
        marker: { color: sList.map(s => s.turnover > 60 ? '#5AD8A6' : s.turnover > 40 ? '#F6BD16' : '#E86452') }
    }], {
        ...plotlyLayout, height: Math.max(300, sList.length * 24),
        xaxis: { ...plotlyLayout.xaxis, title: '动销率(%)', ticksuffix: '%', gridcolor: 'rgba(0,0,0,0)', zerolinecolor: '#2a2d3a', range: [0, 100] },
        yaxis: { automargin: true, gridcolor: 'rgba(0,0,0,0)' },
        margin: { l: 100, r: 20, t: 10, b: 40 }
    }, plotlyCfg);

    // --- 一级分类动销明细表 ---
    if (rawcatData.length > 0) {
        const sorted = rawcatData.filter(r => r.cat && r.cat !== '总计' && r.cat !== '汇总').sort((a, b) => {
            if (a.cat !== b.cat) return a.cat.localeCompare(b.cat);
            return a.raw_cat.localeCompare(b.raw_cat);
        });
        let t = '<table><tr><th>品类</th><th>一级分类</th><th>无动销(SKU)</th><th>有动销(SKU)</th><th>SKU总计</th><th>动销率</th></tr>';
        sorted.forEach(r => {
            const inactive = Math.round(r.sku_total - r.sku_active);
            const catLabel = { '食品': '食品', '日化': '日化', '百货': '百货' }[r.cat] || r.cat;
            t += '<tr><td>' + catLabel + '</td><td>' + (r.raw_cat || '') +
                '</td><td>' + inactive.toLocaleString() + '</td><td>' + Math.round(r.sku_active || 0).toLocaleString() +
                '</td><td>' + Math.round(r.sku_total || 0).toLocaleString() +
                '</td><td>' + (r.turnover_rate || 0).toFixed(1) + '%</td></tr>';
        });
        t += '</table>';
        document.getElementById('prodRawcatTable').innerHTML = t;
    }

    // --- 品类销售汇总表 ---
    const catRows = all.filter(r => ['食品','日化','百货'].includes(r.cat));
    const totalRev = sum(catRows, 'revenue');
    const totalProf = sum(catRows, 'profit');
    const totalSku = Math.round(sum(catRows, 'sku_total'));
    const totalActive = Math.round(sum(catRows, 'sku_active'));
    const totalQty = Math.round(sum(catRows, 'qty'));
    let st = '<table><tr><th>品类</th><th>SKU总数</th><th>有动销SKU</th><th>动销率</th><th>销量</th><th>销售额</th><th>销售额占比</th><th>毛利</th><th>毛利占比</th><th>毛利率</th></tr>';
    catRows.sort((a,b) => (b.revenue||0)-(a.revenue||0)).forEach(r => formatRow(r, false));
    // 总计行
    formatRow({cat:'总计', sku_total:totalSku, sku_active:totalActive, turnover_rate: totalSku>0?(totalActive/totalSku*100):0, qty:totalQty, revenue:totalRev, profit:totalProf, margin: totalRev>0?(totalProf/totalRev*100):0}, true);
    st += '</table>';
    document.getElementById('prodSalesSummary').innerHTML = st;
    function formatRow(r, isTotal) {
        const revPct = totalRev > 0 ? (r.revenue/totalRev*100) : 0;
        const profPct = totalProf > 0 ? (r.profit/totalProf*100) : 0;
        st += '<tr' + (isTotal?' style="font-weight:bold;background:rgba(0,0,0,0.03)"':'') +
            '><td>' + r.cat + '</td><td>' + Math.round(r.sku_total||0).toLocaleString() +
            '</td><td>' + Math.round(r.sku_active||0).toLocaleString() +
            '</td><td>' + (r.turnover_rate||0).toFixed(1) + '%</td><td>' + Math.round(r.qty||0).toLocaleString() +
            '</td><td>' + fmtY(r.revenue||0) + '</td><td>' + revPct.toFixed(1) + '%</td><td>' + fmtY(r.profit||0) +
            '</td><td>' + profPct.toFixed(1) + '%</td><td>' + (r.margin||0).toFixed(1) + '%</td></tr>';
    }

    // --- 一级分类销售明细表 ---
    if (rawcatCombinedData.length > 0) {
        const detail = rawcatCombinedData.filter(r => r.revenue > 0).sort((a, b) => (b.revenue || 0) - (a.revenue || 0));
        const dtRev = detail.reduce((s,r) => s + (r.revenue||0), 0);
        const dtProf = detail.reduce((s,r) => s + (r.profit||0), 0);
        let dt = '<table><tr><th>品类</th><th>一级分类</th><th>销量</th><th>销售额</th><th>销售额占比</th><th>毛利</th><th>毛利占比</th><th>毛利率</th></tr>';
        detail.forEach(r => {
            const catLabel = { '食品': '食品', '日化': '日化', '百货': '百货' }[r.cat] || r.cat;
            const revPct = dtRev > 0 ? (r.revenue/dtRev*100) : 0;
            const profPct = dtProf > 0 ? (r.profit/dtProf*100) : 0;
            dt += '<tr><td>' + catLabel + '</td><td>' + (r.raw_cat || '') +
                '</td><td>' + Math.round(r.qty||0).toLocaleString() +
                '</td><td>' + fmtY(r.revenue||0) + '</td><td>' + revPct.toFixed(1) + '%</td><td>' + fmtY(r.profit||0) +
                '</td><td>' + profPct.toFixed(1) + '%</td><td>' + (r.margin||0).toFixed(1) + '%</td></tr>';
        });
        dt += '</table>';
        document.getElementById('prodSalesDetail').innerHTML = dt;
    }
}

// ============ REFRESH ============
function refresh() {
    const data=getFiltered(), pf=getFilteredPromo();
    let dr = dateFrom === dateTo ? dateFrom : dateFrom + ' ~ ' + dateTo;
    document.querySelectorAll('.section-date-tag').forEach(el => { el.textContent = ' (' + dr + ')'; });
    const prev=getPrevData(data);
    renderSummary(data, prev);
    renderAlerts(data, prev);
    renderKPIs(data); renderStore(data); renderChannel(data); renderTimeAnalysis(data);
    renderNeg(data); renderDelivery(data);
}

// ============ KPI DRILL-DOWN MODAL ============
function openModal(kpiType) {
    const data = getFiltered();
    const stores = groupBy(data, ['store_name']);
    const titleMap = {
        order: '各门店单量排行', revenue: '各门店实收排行', aov: '门店客单价对比',
        store_profit: '各门店毛利排行', comm_profit: '各门店抽佣毛利排行',
        promo: '各门店推广费明细', neg: '各门店负毛利情况', delivery: '各门店配送成本对比'
    };
    document.getElementById('modalTitle').textContent = titleMap[kpiType] || '门店明细';

    let chartData, chartTitle, color, vKey, vLabel;
    switch(kpiType) {
        case 'order':
            stores.sort((a,b)=>b.order_cnt-a.order_cnt);
            chartTitle='单量'; vKey='order_cnt'; color='#5B8FF9'; vLabel='单';
            break;
        case 'revenue':
            stores.sort((a,b)=>b.revenue-a.revenue);
            chartTitle='实收'; vKey='revenue'; color='#5B8FF9'; vLabel='元';
            break;
        case 'aov':
            stores.sort((a,b)=>((b.revenue/(b.order_cnt||1))-(a.revenue/(a.order_cnt||1))));
            chartTitle='客单价'; color='#6C8EF2'; vLabel='元';
            chartData = stores.map(s=>({n:short(s.store_name), v:(s.revenue/(s.order_cnt||1))}));
            break;
        case 'store_profit':
            stores.sort((a,b)=>b.real_profit-a.real_profit);
            chartTitle='门店毛利'; vKey='real_profit'; color='#5AD8A6'; vLabel='元';
            break;
        case 'comm_profit':
            stores.sort((a,b)=>b.commission_profit-a.commission_profit);
            chartTitle='抽佣毛利'; vKey='commission_profit'; color='#5AD8A6'; vLabel='元';
            break;
        case 'promo':
            stores.sort((a,b)=>b.promo_fee-a.promo_fee);
            chartTitle='推广费'; vKey='promo_fee'; color='#E86452'; vLabel='元';
            break;
        case 'neg':
            stores.sort((a,b)=>(b.neg_cnt/(b.order_cnt||1))-(a.neg_cnt/(a.order_cnt||1)));
            chartTitle='负毛利占比'; color='#E86452'; vLabel='%';
            chartData = stores.map(s=>({n:short(s.store_name), v:s.order_cnt>0?((s.neg_cnt||0)/s.order_cnt*100):0}));
            break;
        case 'delivery':
            stores.sort((a,b)=>((b.delivery_fee/(b.delivery_order_cnt||1))-(a.delivery_fee/(a.delivery_order_cnt||1))));
            chartTitle='单均配送成本'; color='#FF9F43'; vLabel='元';
            chartData = stores.map(s=>({n:short(s.store_name), v:(s.delivery_order_cnt>0?(s.delivery_fee/s.delivery_order_cnt):0)}));
            break;
    }
    if (!chartData) chartData = stores.map(s=>({n:short(s.store_name), v:s[vKey]||0}));
    const vals = chartData.map(d=>d.v);
    const names = chartData.map(d=>d.n);
    const maxV = Math.max(...vals, 1);
    const yPad = maxV * 0.25; // 顶部留25%空间给标签

    // 数值格式化 — 提前定义（表格和图表都需要）
    let valFmtFn, vTF;
    if (kpiType==='aov'||kpiType==='delivery') { valFmtFn = v=>'¥'+v.toFixed(2); vTF='.2f'; }
    else if (kpiType==='neg'||kpiType==='margin') { valFmtFn = v=>v.toFixed(2)+'%'; vTF='.2f'; }
    else { valFmtFn = v=>v.toLocaleString(undefined,{maximumFractionDigits:0}); vTF=',d'; }

    // Build content — 升序表格（从小到大）
    const tableData = chartData.map(d=>({n:d.n, v:d.v})).sort((a,b)=>b.v-a.v);
    const tabMaxV = Math.max(...tableData.map(d=>d.v), 1);
    let body = '';
    body += '<div class="table-scroll"><table><tr><th>门店</th><th>'+chartTitle+'</th>';
    if (kpiType==='order'||kpiType==='revenue'||kpiType==='store_profit'||kpiType==='comm_profit'||kpiType==='promo') {
        body += '<th>占比</th>';
    }
    body += '</tr>';
    const totalV = vals.reduce((s,v)=>s+v, 0);
    tableData.forEach((d, i) => {
        const barW = Math.round(d.v / tabMaxV * 120);
        const vDisplay = valFmtFn(d.v);
        const pctStr = totalV > 0 ? ((d.v/totalV*100).toFixed(1)+'%') : '';
        const cls = kpiType==='neg' ? (d.v>30?'cell-bad':d.v>15?'cell-warn':'cell-good') : '';
        body += '<tr><td>'+d.n+'</td>';
        body += '<td class="'+cls+'"><span class="cell-bar-wrap"><span class="cell-bar" style="width:'+barW+'px;background:'+color+'"></span>'+vDisplay+'</span></td>';
        if (pctStr) body += '<td>'+pctStr+'</td>';
        body += '</tr>';
    });
    body += '</table></div>';
    document.getElementById('modalBody').innerHTML = body;
    document.getElementById('kpiModal').classList.add('show');

    // Render chart
    setTimeout(() => {
        Plotly.newPlot('modalChart', [{
            x: names, y: vals, type: 'bar',
            cliponaxis: false,
            marker: { color: vals.map(v => {
                if (kpiType==='neg') return v>30?'#E86452':v>15?'#F6BD16':'#5AD8A6';
                if (kpiType==='promo') return '#E86452';
                return color;
            })},
            text: vals.map(valFmtFn),
            textposition: 'outside',
            textfont: { size: 10, color: '#9ca3af' }
        }], {
            ...plotlyLayout,
            height: 310,
            xaxis: { ...plotlyLayout.xaxis, tickangle: -45, categoryorder: 'total ascending' },
            yaxis: { ...plotlyLayout.yaxis, title: vLabel, gridcolor: '#e5e7eb', automargin: true, tickformat: vTF, exponentformat: 'none' },
            margin: { l: 50, r: 20, t: 50, b: 80 }
        }, plotlyCfg);
    }, 50);
}

function barPct(v, t, timePct) { if(!t)return'';let p=v/t*100;if(timePct!==undefined){let c=p>=timePct?'progress-bad':'progress-good';return'<td class="'+c+'">'+p.toFixed(1)+'%</td>';}let c=p>=45?'progress-good':p>=30?'progress-warn':'progress-bad';return'<td class="'+c+'">'+p.toFixed(1)+'%</td>'; }

function showPerfModal() {
    if (!perfData || !perfData.data || !perfData.data.length) return;
    // 标题
    document.getElementById('perfMonth').textContent = perfData.month + '月';
    // 时间进度
    let timePct = perfData.total_days ? (perfData.days / perfData.total_days * 100) : 0;
    document.getElementById('perfTimeBar').textContent = perfData.days + '/' + perfData.total_days + '天 (' + timePct.toFixed(1) + '%)';
    document.getElementById('perfTimeFill').style.width = Math.min(timePct, 100) + '%';
    document.getElementById('perfDays').textContent = '(截止' + perfData.month + '月, ' + perfData.days + '天数据)';
    // 明细行
    let rows = '';
    perfData.data.forEach(p => {
        rows += '<tr><td>'+p.name+'</td>'
            +'<td>'+p.orders.toLocaleString()+'</td>'
            +'<td>'+p.target_orders.toLocaleString()+'</td>'
            +barPct(p.orders, p.target_orders, timePct)
            +'<td>¥'+p.profit.toLocaleString(undefined,{minimumFractionDigits:2})+'</td>'
            +'<td>¥'+p.target_profit.toLocaleString(undefined,{minimumFractionDigits:2})+'</td>'
            +barPct(p.profit, p.target_profit, timePct)
            +'</tr>';
    });
    document.getElementById('perfBody').innerHTML = rows;
    // 汇总行
    if (perfData.summary) {
        let s = perfData.summary;
        document.getElementById('perfFoot').innerHTML = '<tr style="background:#f0f5ff;font-weight:700;border-top:2px solid #1a1d2e">'
            +'<td style="color:#1a1d2e">📊 合计</td>'
            +'<td>'+s.total_orders.toLocaleString()+'</td>'
            +'<td>'+s.total_target_orders.toLocaleString()+'</td>'
            +barPct(s.total_orders, s.total_target_orders, timePct)
            +'<td>¥'+s.total_profit.toLocaleString(undefined,{minimumFractionDigits:2})+'</td>'
            +'<td>¥'+s.total_target_profit.toLocaleString(undefined,{minimumFractionDigits:2})+'</td>'
            +barPct(s.total_profit, s.total_target_profit, timePct)
            +'</tr>';
    }
    document.getElementById('perfModal').classList.add('show');
}

function closeModal() {
    document.getElementById('kpiModal').classList.remove('show');
    // Clean up Plotly chart to prevent memory leaks
    const chartEl = document.getElementById('modalChart');
    if (chartEl) Plotly.purge(chartEl);
}

// Click outside modal to close
document.addEventListener('click', function(e) {
    if (e.target.id === 'kpiModal') closeModal();
});

// KPI card click delegation
document.getElementById('kpiGrid').addEventListener('click', function(e) {
    const card = e.target.closest('.kpi-card.clickable');
    if (card) {
        const kpiType = card.getAttribute('data-kpi');
        if (kpiType) openModal(kpiType);
    }
});

// ============ PROFIT ANALYSIS MODAL ============
function openProfitModal() {
    var data = rawData.filter(function(r) { return r['日期'] >= dateFrom && r['日期'] <= dateTo; });
    var stores = groupBy(data, ['store_name']);
    
    var endDate = dateTo || allDates[allDates.length-1];
    var startDate = dateFrom || allDates[0];
    var daysInRange = Math.ceil((dParse(endDate) - dParse(startDate)) / 86400000) + 1;
    var DAYS_PER_MONTH = 30;
    var costMul = daysInRange >= 28 ? 1 : daysInRange / DAYS_PER_MONTH;
    
    var totalProfit = 0, totalCost = 0, totalStoreProfit = 0;
    var profitCount = 0, lossCount = 0;
    var rows = [];
    
    stores.forEach(function(s) {
        var cp = s.commission_profit || 0;
        var shortName = short(s.store_name);
        var costInfo = storeCostMap[s.store_name] || storeCostMap[shortName];
        if (!costInfo) {
            var keys = Object.keys(storeCostMap);
            for (var i = 0; i < keys.length; i++) {
                var k = keys[i];
                if (shortName.indexOf(k) >= 0 || k.indexOf(shortName) >= 0) { costInfo = storeCostMap[k]; break; }
            }
        }
        var monthlyCost = costInfo ? costInfo.cost : 0;
        var storeCost = monthlyCost * costMul;
        var profit = cp - storeCost;
        var isProfitable = profit > 0;
        
        if (isProfitable) profitCount++; else lossCount++;
        totalProfit += profit;
        totalCost += storeCost;
        totalStoreProfit += cp;
        
        rows.push({
            name: shortName, cp: cp, cost: storeCost, profit: profit,
            isProfitable: isProfitable,
            hasCost: monthlyCost > 0,
            margin: s.revenue > 0 ? (cp / s.revenue * 100) : 0,
            rev: s.revenue || 0, ord: s.order_cnt || 0
        });
    });
    
    // Sort: profitable → loss → no-cost (bottom)
    rows.sort(function(a,b) {
        if (a.hasCost && !b.hasCost) return -1;
        if (!a.hasCost && b.hasCost) return 1;
        if (!a.hasCost && !b.hasCost) return a.name.localeCompare(b.name);
        // Both have cost: profitable first
        if (a.isProfitable && !b.isProfitable) return -1;
        if (!a.isProfitable && b.isProfitable) return 1;
        if (a.isProfitable) return b.profit - a.profit;
        return a.profit - b.profit;
    });
    
    var profitRate = rows.length > 0 ? (profitCount / rows.length * 100).toFixed(1) : 0;
    var avgProfitPerStore = rows.length > 0 ? (totalProfit / rows.length) : 0;
    
    // Build HTML
    var h = [];
    
    // KPI cards
    h.push('<div class="profit-kpi-row">');
    h.push('<div class="profit-kpi"><div class="pk-label">总盈利</div><div class="pk-value" style="color:' + (totalProfit>=0?'#dc2626':'#16a34a') + '">' + fmtY(totalProfit) + '</div></div>');
    h.push('<div class="profit-kpi"><div class="pk-label">盈利 / 亏损</div><div class="pk-value"><span style="color:#dc2626">' + profitCount + '</span> / <span style="color:#16a34a">' + lossCount + '</span></div></div>');
    h.push('<div class="profit-kpi"><div class="pk-label">盈利占比</div><div class="pk-value">' + profitRate + '%</div></div>');
    h.push('<div class="profit-kpi"><div class="pk-label">日均盈利</div><div class="pk-value">' + fmtY(avgProfitPerStore / Math.max(1,daysInRange)) + '</div></div>');
    h.push('<div class="profit-kpi"><div class="pk-label">计算周期</div><div class="pk-value" style="font-size:20px;">' + daysInRange + ' 天</div><div class="pk-sub">成本系数 ' + costMul.toFixed(2) + '</div></div>');
    h.push('</div>');
    
    // Table
    h.push('<div class="table-scroll">');
    h.push('<table>');
    h.push('<thead><tr><th>序号</th><th>门店</th><th>订单量</th><th>抽佣毛利</th><th>店铺成本</th><th>净利润</th><th>利润率</th><th>客单价</th></tr></thead>');
    h.push('<tbody>');
    
    var inGroup = null;
    for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var curGroup = r.hasCost ? (r.isProfitable ? 'profit' : 'loss') : 'nocost';
        if (curGroup !== inGroup) {
            inGroup = curGroup;
            var grpStyle, grpLabel;
            if (curGroup === 'profit') { grpStyle = 'background:#fef2f2;color:#dc2626;'; grpLabel = '盈利门店'; }
            else if (curGroup === 'loss') { grpStyle = 'background:#ecfdf5;color:#16a34a;'; grpLabel = '亏损门店'; }
            else { grpStyle = 'background:#f9fafb;color:#6b7280;'; grpLabel = '暂无成本数据'; }
            h.push('<tr class="profit-grp"><td colspan="8" style="' + grpStyle + 'font-weight:600;padding:8px 10px;">' + grpLabel + '</td></tr>');
        }
        var aov = r.ord > 0 ? (r.rev / r.ord) : 0;
        var pc = r.hasCost ? (r.isProfitable ? 'cell-bad' : 'cell-good') : '';
        h.push('<tr>'
            + '<td>' + (i+1) + '</td>'
            + '<td>' + r.name + '</td>'
            + '<td>' + r.ord.toLocaleString() + '</td>'
            + '<td>' + fmtY(r.cp) + '</td>'
            + '<td>' + (r.hasCost ? fmtY(r.cost) : '—') + '</td>'
            + '<td class="' + pc + '">' + (r.hasCost ? fmtY(r.profit) : '—') + '</td>'
            + '<td>' + r.margin.toFixed(1) + '%</td>'
            + '<td>' + fmtY(aov) + '</td>'
            + '</tr>');
    }
    
    // Summary
    var sc = totalProfit >= 0 ? 'cell-good' : 'cell-bad';
    h.push('<tr class="profit-sum" style="font-weight:600;background:#f9fafb;">'
        + '<td colspan="2">合计</td>'
        + '<td></td>'
        + '<td>' + fmtY(totalStoreProfit) + '</td>'
        + '<td>' + fmtY(totalCost) + '</td>'
        + '<td class="' + sc + '">' + fmtY(totalProfit) + '</td>'
        + '<td colspan="2"></td>'
        + '</tr>');
    
    h.push('</tbody></table></div>');
    
    document.getElementById('profitModalBody').innerHTML = h.join('');
    document.getElementById('profitModal').classList.add('show');
}

function closeProfitModal() {
    document.getElementById('profitModal').classList.remove('show');
}

document.addEventListener('click', function(e) {
    if (e.target.id === 'profitModal') closeProfitModal();
});

// ============ INIT ============
function initDashboard() {
    initFilters();
    document.getElementById('updateTime').textContent = '数据更新: ' + allDates[allDates.length - 1];
    refresh();
}
// initDashboard() 在 loadData() 完成后调用——见上方 async 函数
</script>
</body>
</html>'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)

# 同步到 index.html（GitHub Pages 入口文件）
INDEX_OUT = os.path.join(BASE, 'index.html')
with open(INDEX_OUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Generated: {OUTPUT}')
print(f'Synced  : {INDEX_OUT}')
print(f'Size: {os.path.getsize(OUTPUT) / 1024:.0f} KB')
print(f'Data: {len(df)} rows, {len(dates_all)} days, {len(stores_full)} stores')
