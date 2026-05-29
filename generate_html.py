# -*- coding: utf-8 -*-
"""
Generate self-contained HTML BI dashboard v2:
  - Calendar date picker + quick buttons
  - Store search + chips (short names)
  - New fields: 实收/门店毛利/抽佣毛利/毛利率/实收客单/单均配送成本
"""
import pandas as pd
import json, os, re

BASE = r'E:\Desktop\工作文件（月度）\claw制作BI'
WAREHOUSE = os.path.join(BASE, 'warehouse')
OUTPUT = os.path.join(BASE, 'dashboard.html')

# Load data
df = pd.read_excel(os.path.join(WAREHOUSE, 'daily_store_channel_profit.xlsx'))
df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')

# 仪表盘只取最近90天，warehouse原始文件保持全量
df['_date_sort'] = pd.to_datetime(df['日期'])
max_date = df['_date_sort'].max()
cutoff = max_date - pd.Timedelta(days=90)
df = df[df['_date_sort'] >= cutoff].drop(columns=['_date_sort'])

promo = None
promo_path = os.path.join(WAREHOUSE, 'promo_daily.xlsx')
if os.path.exists(promo_path):
    promo = pd.read_excel(promo_path)
    promo['日期'] = pd.to_datetime(promo['日期']).dt.strftime('%Y-%m-%d')
    promo['_date_sort'] = pd.to_datetime(promo['日期'])
    promo = promo[promo['_date_sort'] >= cutoff].drop(columns=['_date_sort'])

def clean(records):
    out = []
    for r in records:
        row = {}
        for k, v in r.items():
            if pd.isna(v): row[k] = 0
            elif hasattr(v, 'is_integer'): row[k] = int(v)
            elif hasattr(v, 'is_real'): row[k] = round(float(v), 2)
            else: row[k] = v
        out.append(row)
    return out

data_records = clean(df.to_dict(orient='records'))
promo_records = clean(promo.to_dict(orient='records')) if promo is not None else []

# --- Short store name extraction ---
def short_name(full):
    m = re.search(r'[（(]([^）)]+)[）)]', full)
    return m.group(1) if m else full

stores_full = sorted(df['store_name'].unique().tolist())
short_map = {}
search_map = {}
for s in stores_full:
    sn = short_name(s)
    short_map[s] = sn
    sid = s.split('-')[0] if '-' in s else ''
    search_map[s] = json.dumps([sn, sid, s], ensure_ascii=False)

channels = sorted(df['channel'].unique().tolist())
dates_all = sorted(df['日期'].unique().tolist())

data_json = json.dumps(data_records, ensure_ascii=False)
promo_json = json.dumps(promo_records, ensure_ascii=False)
stores_json = json.dumps(stores_full, ensure_ascii=False)
short_map_json = json.dumps(short_map, ensure_ascii=False)
search_map_json = json.dumps(search_map, ensure_ascii=False)
channels_json = json.dumps(channels, ensure_ascii=False)
dates_json = json.dumps(dates_all, ensure_ascii=False)

# =============================================
# HTML template
# =============================================
html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>微笑客经营看板</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
<style>
:root {
    --bg: #0f1117; --card: #1a1d29; --border: #2a2d3a;
    --text: #e8eaf0; --text-dim: #8b8fa3;
    --accent: #6C8EF2; --green: #5AD8A6; --yellow: #F6BD16; --red: #E86452; --blue: #5B8FF9;
    --orange: #FF9F43;
}
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
.cal-grid .cal-day:hover { background:rgba(108,142,242,0.15); }
.cal-grid .cal-day.in-range { background:rgba(108,142,242,0.25); color:var(--text); border-radius:0; }
.cal-grid .cal-day.range-start { background:var(--accent); color:#fff; border-radius:6px 0 0 6px; }
.cal-grid .cal-day.range-end { background:var(--accent); color:#fff; border-radius:0 6px 6px 0; }
.cal-grid .cal-day.range-start.range-end { border-radius:6px; }
.cal-grid .cal-day.other-month { color:rgba(139,143,163,0.3); }
.cal-toggle { color:var(--accent); font-size:12px; cursor:pointer; padding:3px 8px; border-radius:10px; border:1px solid var(--border); background:transparent; touch-action:manipulation; }
.cal-toggle:hover { background:rgba(108,142,242,0.1); }
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
.kpi-card { background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; }
.kpi-label { color:var(--text-dim); font-size:12px; margin-bottom:4px; }
.kpi-value { font-size:24px; font-weight:700; letter-spacing:-0.3px; }
.kpi-sub { font-size:11px; margin-top:4px; }
.accent .kpi-value { color:var(--accent); }
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
.row { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
.table-scroll { overflow-x:auto; -webkit-overflow-scrolling:touch; border-radius:8px; border:1px solid var(--border); }
.table-scroll table { min-width:700px; width:100%; border-collapse:collapse; font-size:12px; }
th { text-align:left; padding:8px 10px; background:var(--bg); color:var(--text-dim); font-weight:600; border-bottom:1px solid var(--border); position:sticky; top:0; white-space:nowrap; }
td { padding:6px 10px; border-bottom:1px solid var(--border); white-space:nowrap; }
tr:hover td { background:rgba(108,142,242,0.05); }
.badge { display:inline-block; padding:2px 6px; border-radius:3px; font-size:11px; font-weight:500; }
.badge-warn { background:rgba(232,100,82,0.15); color:var(--red); }
.badge-ok { background:rgba(90,216,166,0.15); color:var(--green); }
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
/* === Login Overlay === */
.login-overlay { position:fixed; inset:0; background:var(--bg); z-index:9999; display:flex; align-items:center; justify-content:center; }
.login-box { background:var(--card); border:1px solid var(--border); border-radius:16px; padding:32px 28px; width:340px; max-width:90vw; text-align:center; }
.login-box h2 { font-size:20px; margin-bottom:4px; }
.login-box .login-sub { color:var(--text-dim); font-size:12px; margin-bottom:20px; }
.login-box input { width:100%; padding:10px 14px; border-radius:10px; border:1px solid var(--border); background:var(--bg); color:var(--text); font-size:14px; outline:none; text-align:center; }
.login-box input:focus { border-color:var(--accent); }
.login-box .login-err { color:var(--red); font-size:12px; margin-top:10px; min-height:18px; }
.login-box button { margin-top:12px; width:100%; padding:10px; border-radius:10px; border:none; background:var(--accent); color:#fff; font-size:14px; font-weight:600; cursor:pointer; }
.login-box button:hover { opacity:0.9; }
.export-btn { padding:4px 12px; border-radius:6px; border:1px solid var(--border); background:var(--bg); color:var(--text-dim); font-size:11px; cursor:pointer; float:right; }
.export-btn:hover { border-color:var(--accent); color:var(--text); }
</style>
</head>
<body>
<div class="login-overlay" id="loginOverlay">
    <div class="login-box">
        <h2>🔐 微笑客经营看板</h2>
        <div class="login-sub">请输入访问密码</div>
        <input type="password" id="loginPwd" placeholder="请输入密码" onkeydown="if(event.key==='Enter')doLogin()">
        <button onclick="doLogin()">进入看板</button>
        <div class="login-err" id="loginErr"></div>
    </div>
</div>

<div class="dashboard" id="dashboard" style="display:none;">
    <div class="header">
        <h1>微笑客经营看板</h1>
        <div class="update-time" id="updateTime"></div>
    </div>

    <div class="filter-bar">
        <div class="filter-row">
            <span class="filter-label">📅</span>
            <div class="date-quick" id="dateQuick">
                <button class="date-btn" onclick="setDateRange('yesterday')">昨天</button>
                <button class="date-btn" onclick="setDateRange('thisWeek')">本周</button>
                <button class="date-btn" onclick="setDateRange('thisMonth')">本月</button>
                <button class="date-btn" onclick="setDateRange('lastMonth')">上月</button>
                <button class="date-btn active" onclick="toggleCalendar()">自定义</button>
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
                <input class="store-search" id="storeSearch" placeholder="搜索门店（编号/名称）" oninput="filterStores()">
                <div class="chip-group" id="storeChips"></div>
            </div>
            <div class="chip-actions">
                <button class="chip-action" onclick="selectAll('stores')">全选</button>
                <button class="chip-action" onclick="deselectAll('stores')">清空</button>
            </div>
        </div>
        <div class="filter-row">
            <span class="filter-label">📡</span>
            <div class="chip-group" id="channelChips"></div>
            <div class="chip-actions">
                <button class="chip-action" onclick="selectAll('channels')">全选</button>
                <button class="chip-action" onclick="deselectAll('channels')">清空</button>
            </div>
        </div>
    </div>

    <div class="kpi-grid" id="kpiGrid"></div>

    <div class="tab-nav-wrap">
        <div class="tab-nav">
            <button class="tab-btn active" onclick="switchTab('store',this)">门店数据</button>
            <button class="tab-btn" onclick="switchTab('channel',this)">渠道分析</button>
            <button class="tab-btn" onclick="switchTab('time',this)">时间分析</button>
        </div>
    </div>

    <div id="tab-store" class="tab-content active">
        <div class="chart-section"><h3>各门店单量 vs 抽佣毛利</h3><div id="chartStoreBar"></div></div>
        <div class="chart-section"><h3>门店明细 <button class="export-btn" onclick="exportTable('storeTable','门店明细')">📥 导出Excel</button></h3><div class="table-scroll" id="storeTable"></div></div>
    </div>
    <div id="tab-channel" class="tab-content">
        <div class="row">
            <div class="chart-section"><h3>单量渠道分布</h3><div id="chartChannelPie"></div></div>
            <div class="chart-section"><h3>各渠道毛利率</h3><div id="chartChannelMargin"></div></div>
        </div>
        <div class="row">
            <div class="chart-section"><h3>各渠道实收</h3><div id="chartChannelRev"></div></div>
            <div class="chart-section"><h3>各渠道平台抽佣</h3><div id="chartChannelComm"></div></div>
        </div>
        <div class="chart-section"><h3>渠道明细 <button class="export-btn" onclick="exportTable('channelTable','渠道明细')">📥 导出Excel</button></h3><div class="table-scroll" id="channelTable"></div></div>
    </div>
    <div id="tab-time" class="tab-content">
        <div class="chart-section"><h3>每日单量 & 抽佣毛利趋势</h3><div id="chartTimeTrend"></div></div>
        <div class="chart-section"><h3>每日明细 <button class="export-btn" onclick="exportTable('timeTable','每日明细')">📥 导出Excel</button></h3><div class="table-scroll" id="timeTable"></div></div>
    </div>
</div>

<script>
// ============ LOGIN (暂关) ============
const PWD_HASH = '33f17839cb399b359a0deaa4387a29ca211dd8be792019288aeffb029e692b46';
(async function(){
    document.getElementById('loginOverlay').style.display='none'; document.getElementById('dashboard').style.display='block'; initDashboard();
})();
async function sha256(m){const b=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(m));return Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,'0')).join('');}
async function doLogin(){
    const pwd=document.getElementById('loginPwd').value;
    if(!pwd){document.getElementById('loginErr').textContent='请输入密码';return;}
    const h=await sha256(pwd);
    if(h===PWD_HASH){sessionStorage.setItem('_auth','1');document.getElementById('loginOverlay').style.display='none';document.getElementById('dashboard').style.display='block';initDashboard();}
    else{document.getElementById('loginErr').textContent='密码错误，请重试';document.getElementById('loginPwd').value='';}
}

// ============ EXPORT ============
function exportTable(tableId, filename) {
    const table = document.getElementById(tableId).querySelector('table');
    if(!table) return;
    const wb = XLSX.utils.table_to_book(table, {sheet:'明细'});
    XLSX.writeFile(wb, filename + '.xlsx');
}


// ============ DATA ============
const rawData = ''' + data_json + ''';
const promoData = ''' + promo_json + ''';
const allStoresFull = ''' + stores_json + ''';
const allChannels = ''' + channels_json + ''';
const allDates = ''' + dates_json + ''';
const shortNames = ''' + short_map_json + ''';
const searchKeys = ''' + search_map_json + ''';

const isMobile = window.innerWidth < 768;
const chartH = isMobile ? 260 : 380;
const chartHS = isMobile ? 220 : 340;

const plotlyLayout = { paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#e8eaf0',family:'-apple-system,PingFang SC,Microsoft YaHei'}, xaxis:{gridcolor:'#2a2d3a',zerolinecolor:'#2a2d3a'}, yaxis:{gridcolor:'#2a2d3a',zerolinecolor:'#2a2d3a'}, margin:{l:60,r:20,t:20,b:60}, height:chartH };
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
        case 'thisWeek':
            let [mon,sun] = calcWeekBounds(new Date());
            from = dFmt(mon); to = dFmt(sun);
            break;
        case 'thisMonth':
            let [m1,m2] = calcMonthBounds(new Date());
            from = dFmt(m1); to = dFmt(m2);
            break;
        case 'lastMonth':
            let d2 = new Date(); d2.setMonth(d2.getMonth()-1);
            let [lm1,lm2] = calcMonthBounds(d2);
            from = dFmt(lm1); to = dFmt(lm2);
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
    refresh();
}

function toggleCalendar() {
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
        let cls = 'cal-day';
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
        html += '<div class="'+cls+'" data-date="'+ds+'" onclick="calClick(\\''+ds+'\\')">'+d+'</div>';
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
    let q = document.getElementById('storeSearch').value.trim().toLowerCase();
    document.querySelectorAll('#storeChips .chip').forEach(c => {
        let full = c.getAttribute('data-full');
        let keys = JSON.parse(searchKeys[full] || '[]');
        let match = q==='' || keys.some(k => k.toLowerCase().includes(q));
        c.classList.toggle('hidden', !match);
    });
}

// ============ FILTER STATE ============
let selectedStores = [...allStoresFull];
let selectedChannels = [...allChannels];

function createChips(containerId, items, selectedArr, onChange) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    items.forEach(item => {
        const chip = document.createElement('span');
        chip.className = 'chip active';
        chip.setAttribute('data-full', item);
        chip.textContent = short(item);
        chip.onclick = function() {
            const idx = selectedArr.indexOf(item);
            if (idx > -1) { selectedArr.splice(idx,1); this.classList.remove('active'); }
            else { selectedArr.push(item); this.classList.add('active'); }
            if (onChange) onChange();
        };
        container.appendChild(chip);
    });
}

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
    setDateRange('yesterday');
    createChips('storeChips', allStoresFull, selectedStores, refresh);
    createChips('channelChips', allChannels, selectedChannels, refresh);
}

// ============ TAB ============
function switchTab(name, btn) {
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-'+name).classList.add('active');
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
    return rawData.filter(r=>r['日期']>=f&&r['日期']<=t&&selectedStores.includes(r['store_name'])&&selectedChannels.includes(r['channel']));
}
function getFilteredPromo() {
    const f=dateFrom, t=dateTo;
    return promoData.filter(r=>r['日期']>=f&&r['日期']<=t&&selectedStores.includes(r['store_name']));
}

// ============ 环比 ============
function getPrevData(data) {
    const from = dParse(dateFrom), to = dParse(dateTo);
    const days = Math.round((to - from) / (24*60*60*1000)) + 1;
    const prevEnd = new Date(from); prevEnd.setDate(prevEnd.getDate() - 1);
    const prevStart = new Date(prevEnd); prevStart.setDate(prevStart.getDate() - days + 1);
    const pf = dFmt(prevStart), pt = dFmt(prevEnd);
    return rawData.filter(r =>
        r['日期'] >= pf && r['日期'] <= pt &&
        selectedStores.includes(r['store_name']) &&
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

// ============ KPIs ============
function renderKPIs(data) {
    const ord=sum(data,'order_cnt'), rev=sum(data,'revenue'), gp=sum(data,'real_profit');
    const cc=sum(data,'commission_fee'), cp=sum(data,'commission_profit'), nc=sum(data,'neg_cnt'), pf=sum(data,'promo_fee');
    const df2=sum(data,'delivery_fee'), doc=sum(data,'delivery_order_cnt');
    const margin=rev>0?(cp/rev*100).toFixed(1):0;
    const negPct=ord>0?(nc/ord*100).toFixed(1):0;
    const aov=ord>0?(rev/ord).toFixed(1):0;
    const adc=doc>0?(df2/doc).toFixed(1):0;
    const ngCls=negPct>25?'red':'green';

    // 环比
    const prev = getPrevData(data);
    const pOrd=sum(prev,'order_cnt'), pRev=sum(prev,'revenue'), pGp=sum(prev,'real_profit');
    const pCc=sum(prev,'commission_fee'), pCp=sum(prev,'commission_profit'), pPf=sum(prev,'promo_fee');
    const pDf=sum(prev,'delivery_fee'), pDoc=sum(prev,'delivery_order_cnt');
    const pNeg=sum(prev,'neg_cnt');
    const pAov = pOrd > 0 ? (pRev / pOrd) : 0;
    const pAdc = pDoc > 0 ? (pDf / pDoc) : 0;
    const pNegPct = pOrd > 0 ? (pNeg / pOrd * 100) : 0;
    const pMargin = pRev > 0 ? (pCp / pRev * 100) : 0;

    // 环比：上升标红↓，下降标绿↑（统一规则）
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
        // 第一行
        '<div class="kpi-card blue"><div class="kpi-label">总单量</div><div class="kpi-value">'+ord.toLocaleString()+'</div><div class="kpi-sub" style="color:var(--'+dOrd.cls+')">'+dOrd.txt+'</div></div>'+
        '<div class="kpi-card accent"><div class="kpi-label">实收</div><div class="kpi-value">'+fmtY(rev)+'</div><div class="kpi-sub" style="color:var(--'+dRev.cls+')">'+dRev.txt+'</div></div>'+
        '<div class="kpi-card accent"><div class="kpi-label">实收客单</div><div class="kpi-value">¥'+aov+'</div><div class="kpi-sub" style="color:var(--'+dAov.cls+')">'+dAov.txt+'</div></div>'+
        '<div class="kpi-card green"><div class="kpi-label">门店毛利</div><div class="kpi-value">'+fmtY(gp)+'</div><div class="kpi-sub" style="color:var(--'+dGp.cls+')">'+dGp.txt+'</div></div>'+
        '<div class="kpi-card green"><div class="kpi-label">抽佣毛利</div><div class="kpi-value">'+fmtY(cp)+'</div><div class="kpi-sub" style="color:var(--'+dCp.cls+')">'+dCp.txt+'</div></div>'+
        // 第二行
        '<div class="kpi-card red"><div class="kpi-label">推广</div><div class="kpi-value">¥'+(pf.toFixed(0))+'</div><div class="kpi-sub" style="color:var(--'+dPf.cls+')">'+dPf.txt+'</div></div>'+
        '<div class="kpi-card orange"><div class="kpi-label">抽佣</div><div class="kpi-value">'+fmtY(cc)+'</div><div class="kpi-sub" style="color:var(--'+dCc.cls+')">'+dCc.txt+'</div></div>'+
        '<div class="kpi-card yellow"><div class="kpi-label">毛利率</div><div class="kpi-value">'+margin+'%</div><div class="kpi-sub" style="color:var(--'+dMarg.cls+')">'+dMarg.txt+'</div></div>'+
        '<div class="kpi-card '+ngCls+'"><div class="kpi-label">负毛利占比</div><div class="kpi-value">'+negPct+'%</div><div class="kpi-sub" style="color:var(--'+dNeg.cls+')">'+dNeg.txt+'</div></div>'+
        '<div class="kpi-card orange"><div class="kpi-label">单均配送</div><div class="kpi-value">¥'+adc+'</div><div class="kpi-sub" style="color:var(--'+dAdc.cls+')">'+dAdc.txt+'</div></div>';
}

// ============ STORE TAB ============
function renderStore(data) {
    const stores = groupBy(data,['store_name']);
    stores.sort((a,b)=>b.revenue-a.revenue);
    const names=stores.map(s=>short(s.store_name));
    // 计算Y轴范围：0点对齐（有负毛利时左轴也往下移）
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

    // 双Y轴并列：offsetgroup让两个不同Y轴的bar分组并排
    Plotly.newPlot('chartStoreBar',[
        {x:names,y:stores.map(s=>s.order_cnt||0),name:'单量',type:'bar',marker:{color:'#5B8FF9'},offsetgroup:0},
        {x:names,y:stores.map(s=>s.commission_profit||0),name:'抽佣毛利',type:'bar',marker:{color:'#FF9F43'},yaxis:'y2',offsetgroup:1}
    ],{...plotlyLayout,barmode:'group',
        xaxis:{...plotlyLayout.xaxis,tickangle:-45},
        yaxis:{title:'单量',gridcolor:'rgba(0,0,0,0)',zerolinecolor:'#2a2d3a',range:leftRng},
        yaxis2:{title:'抽佣毛利(元)',overlaying:'y',side:'right',gridcolor:'rgba(0,0,0,0)',range:rightRng},
        margin:{l:50,r:60,t:20,b:100}
    },plotlyCfg);

    let h='<table><tr><th>门店</th><th>单量</th><th>实收</th><th>门店毛利</th><th>平台抽佣</th><th>抽佣毛利</th><th>毛利率</th><th>实收客单</th><th>配送成本</th><th>负毛利</th></tr>';
    stores.forEach(s=>{
        const rev=s.revenue||0, rp=s.real_profit||0, cf=s.commission_fee||0, cp=s.commission_profit||0;
        const mg=rev>0?(cp/rev*100).toFixed(1):'0.0';
        const aov=rev>0?(rev/(s.order_cnt||1)).toFixed(1):'0.0';
        const dc=(s.delivery_order_cnt||0)>0?(s.delivery_fee/s.delivery_order_cnt).toFixed(1):'0.0';
        const np=(s.order_cnt||0)>0?((s.neg_cnt||0)/s.order_cnt*100).toFixed(1):'0.0';
        h+='<tr><td>'+short(s.store_name)+'</td><td>'+(s.order_cnt||0).toLocaleString()+'</td><td>'+fmtY(rev)+'</td><td>'+fmtY(rp)+'</td><td>'+fmtY(cf)+'</td><td>'+fmtY(cp)+'</td><td>'+mg+'%</td><td>¥'+aov+'</td><td>¥'+dc+'</td><td><span class="badge badge-'+(np>25?'warn':'ok')+'">'+np+'%</span></td></tr>';
    });
    h+='</table>';
    document.getElementById('storeTable').innerHTML=h;
}

// ============ CHANNEL TAB ============
function renderChannel(data) {
    const ch = groupBy(data,['channel']);
    const to = ch.reduce((s,c)=>s+c.order_cnt,0);
    ch.forEach(c=>{c.pct=to>0?(c.order_cnt/to*100).toFixed(1):'0.0';});

    Plotly.newPlot('chartChannelPie',[{values:ch.map(c=>c.order_cnt),labels:ch.map(c=>c.channel),type:'pie',hole:0.45,marker:{colors:['#5B8FF9','#F6BD16','#5AD8A6','#E86452']},textinfo:'label+percent',textposition:'outside'}],{...plotlyLayout,height:chartHS,showlegend:false},plotlyCfg);

    const chRev=[...ch].sort((a,b)=>b.revenue-a.revenue);
    Plotly.newPlot('chartChannelRev',[{y:chRev.map(c=>c.channel),x:chRev.map(c=>c.revenue),type:'bar',orientation:'h',marker:{color:'#5B8FF9'}}],{...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'实收(元)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    const chComm=[...ch].sort((a,b)=>b.commission_fee-a.commission_fee);
    Plotly.newPlot('chartChannelComm',[{y:chComm.map(c=>c.channel),x:chComm.map(c=>c.commission_fee),type:'bar',orientation:'h',marker:{color:'#FF9F43'}}],{...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'平台抽佣(元)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    const chMarg=[...ch].sort((a,b)=>(b.commission_profit/(b.revenue||1))-(a.commission_profit/(a.revenue||1)));
    Plotly.newPlot('chartChannelMargin',[{y:chMarg.map(c=>c.channel),x:chMarg.map(c=>c.revenue>0?(c.commission_profit/c.revenue*100).toFixed(1):0),type:'bar',orientation:'h',marker:{color:'#F6BD16'}}],{...plotlyLayout,height:chartHS,xaxis:{...plotlyLayout.xaxis,title:'毛利率(%)'},margin:{l:80,r:20,t:10,b:40}},plotlyCfg);

    let h='<table><tr><th>渠道</th><th>单量</th><th>占比(%)</th><th>实收</th><th>门店毛利</th><th>平台抽佣</th><th>抽佣毛利</th><th>毛利率(%)</th><th>负毛利占比</th></tr>';
    ch.forEach(c=>{
        const rev=c.revenue||0, rp=c.real_profit||0, cf=c.commission_fee||0, cp=c.commission_profit||0;
        const mg=rev>0?(cp/rev*100).toFixed(1):'0.0';
        const np=c.order_cnt>0?((c.neg_cnt||0)/c.order_cnt*100).toFixed(1):'0.0';
        h+='<tr><td>'+c.channel+'</td><td>'+c.order_cnt.toLocaleString()+'</td><td>'+c.pct+'%</td><td>'+fmtY(rev)+'</td><td>'+fmtY(rp)+'</td><td>'+fmtY(cf)+'</td><td>'+fmtY(cp)+'</td><td>'+mg+'%</td><td><span class="badge badge-'+(np>25?'warn':'ok')+'">'+np+'%</span></td></tr>';
    });
    h+='</table>';
    document.getElementById('channelTable').innerHTML=h;
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
        legend: { orientation: 'h', y: 1.1, font: {color:'#e8eaf0',size:11} },
        margin: { l: 50, r: 60, t: 20, b: 50 }
    }, plotlyCfg);

    // 每日明细表
    let h='<table><tr><th>日期</th><th>单量</th><th>实收</th><th>门店毛利</th><th>抽佣</th><th>抽佣毛利</th><th>毛利率</th></tr>';
    byDate.forEach(d=>{
        const rev=d.revenue||0, rp=d.real_profit||0, cf=d.commission_fee||0, cp=d.commission_profit||0;
        const mg=rev>0?(cp/rev*100).toFixed(1):'0.0';
        h+='<tr><td>'+d.日期+'</td><td>'+(d.order_cnt||0).toLocaleString()+'</td><td>'+fmtY(rev)+'</td><td>'+fmtY(rp)+'</td><td>'+fmtY(cf)+'</td><td>'+fmtY(cp)+'</td><td>'+mg+'%</td></tr>';
    });
    h+='</table>';
    document.getElementById('timeTable').innerHTML=h;
}

// ============ REFRESH ============
function refresh() {
    const data=getFiltered(), pf=getFilteredPromo();
    renderKPIs(data); renderStore(data); renderChannel(data); renderTimeAnalysis(data);
}

// ============ INIT ============
function initDashboard() {
    initFilters();
    document.getElementById('updateTime').textContent = '数据更新: ' + allDates[allDates.length - 1];
    refresh();
}
</script>
</body>
</html>'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Generated: {OUTPUT}')
print(f'Size: {os.path.getsize(OUTPUT) / 1024:.0f} KB')
print(f'Data: {len(data_records)} rows, {len(dates_all)} days, {len(stores_full)} stores')
