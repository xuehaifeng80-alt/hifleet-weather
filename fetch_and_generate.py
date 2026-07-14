#!/usr/bin/env python3
"""
HiFleet Weather Dashboard Generator
完全匹配邮件原始排版格式：
  - 表头：日期行 + 时间行（0600/1200/1800/2400 × 7天 = 28时段）
  - 每船4行：风级 / 潮差 / 能见度 / 浪高
  - 颜色规则与邮件Excel一致
"""
import os
import json
import re
from datetime import datetime, timedelta
from imap_tools import MailBox, AND
from collections import defaultdict

# === Config ===
EMAIL_USER = os.environ.get("EMAIL_USER", "xuehaifeng666@qq.com")
EMAIL_AUTH  = os.environ.get("EMAIL_AUTH",  "ztfzkfwzzywmjdaf")
QQ_APP_PWD  = os.environ.get("QQ_APP_PWD",  "")

EXCLUDE_SHIPS = {"ORE CHINA", "ORE DONGJIAKOU", "ORE HEBEI", "ORE SHANDONG"}

WIND_THRESHOLD = 6
VIS_THRESHOLD  = 5
WAVE_RED       = 3.0
WAVE_ORANGE    = 2.5

# 邮件Excel颜色对照
WIND_COLORS = {   # 蒲福风级 → 背景色
    (1, 4):   '#C6EFCE',  # 绿色 1-4级
    (5, 5):   '#FFEB9C',  # 黄色 5级
    (6, 9):   '#FFFF00',  # 黄色 6-9级
    (10, 99): '#FF6600',  # 橙色 >=10级
}
SHIP_TYPE_COLORS = {
    'VLOC Own':  '#E8D5F5',
    'VLOC':      '#DDEEFF',
    'Capesize':  '#FFE0CC',
    'Panamax':   '#FFFFCC',
    'Ultramax':  '#CCFFE0',
    'Supramax':  '#FFFFE0',
    'Handymax':  '#E0FFFF',
    'Unknown':   '#F5F5F5',
}
PARAM_COLORS = {
    '风级':   '#E2EFDA',
    '潮差':   '#FFF2CC',
    '能见度': '#DEEBF7',
    '浪高':   '#FCE4D6',
}
WAVE_COLOR_MAP = {
    'green':  '#C6EFCE',
    'orange': '#FFEB9C',
    'red':    '#FF6600',
    'orange_wave': '#FF6600',
}

def wind_color(level_str):
    try:
        level = float(level_str)
        for (lo, hi), color in WIND_COLORS.items():
            if lo <= level <= hi:
                return color
        return '#F5F5F5'
    except:
        return '#F5F5F5'

def wave_color(val_str):
    try:
        v = float(val_str)
        if v >= WAVE_RED:    return '#FF6600'
        if v >= WAVE_ORANGE: return '#FFEB9C'
        return '#C6EFCE'
    except:
        return '#C6EFDA'

def vis_color(val_str):
    try:
        v = float(val_str)
        if v < VIS_THRESHOLD: return '#FF6600'
        if v < 10:            return '#FFEB9C'
        return '#C6EFCE'
    except:
        return '#DEEBF7'

# ── IMAP 抓取 ────────────────────────────────────────────────────────────────

def fetch_hifleet_email():
    password = QQ_APP_PWD or EMAIL_AUTH
    if not password:
        print("No IMAP password"); return None
    try:
        with MailBox("imap.qq.com").login(EMAIL_USER, password) as mb:
            msgs = list(mb.fetch(AND(from_="hifleet.com"), reverse=True, limit=1))
            if msgs: return msgs[0].html
    except Exception as e:
        print(f"IMAP error: {e}")
    return None

# ── 解析 ────────────────────────────────────────────────────────────────────

def parse_ship_name(text):
    text = text.strip()
    port = ""
    if '：' in text:
        parts = text.split('：', 1)
        text = parts[0].strip()
        port = parts[1].strip()
    types_ = sorted([
        'VLOC Own','vloc own','Capesize','capesize','Panamax','panamax',
        'Ultramax','ultramax','Supramax','supramax','Handymax','handymax',
        'VLOC','vloc'
    ], key=len, reverse=True)
    name = text; stype = 'Unknown'; is_own = False
    for st in types_:
        if name.endswith(st):
            lo = st.lower()
            if 'own' in lo: stype = 'VLOC Own'; is_own = True
            elif lo == 'vloc':     stype = 'VLOC'
            elif lo == 'capesize': stype = 'Capesize'
            elif lo == 'panamax':  stype = 'Panamax'
            elif lo == 'ultramax': stype = 'Ultramax'
            elif lo == 'supramax': stype = 'Supramax'
            elif lo == 'handymax': stype = 'Handymax'
            name = name[:-len(st)].strip()
            break
    return name, stype, port, is_own

def parse_html(html_text):
    """解析邮件HTML，提取所有船舶数据。"""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_text, 'html.parser')
    tables = soup.find_all('table')
    main_tbl = None
    for tbl in tables:
        if re.search(r'\d{4}-\d{2}-\d{2}', tbl.get_text()):
            main_tbl = tbl; break
    if not main_tbl: return [], []
    rows = main_tbl.find_all('tr')

    # 解析表头行（row 0）和时间行（row 1）
    h_cells = [c.get_text(strip=True) for c in rows[0].find_all(['td','th'])]
    t_cells = [c.get_text(strip=True) for c in rows[1].find_all(['td','th'])]
    col_dts = []
    for ci in range(2, len(h_cells)):
        d = h_cells[ci]; t = t_cells[ci] if ci < len(t_cells) else ''
        if re.match(r'\d{4}-\d{2}-\d{2}', d):
            col_dts.append((d, t))

    # 解析船舶数据行（从 row 2 开始）
    data_rows = rows[2:]
    ships = []
    known_params = {'风级','潮差','能见度','浪高'}
    i = 0
    while i < len(data_rows):
        cells0 = [c.get_text(strip=True) for c in data_rows[i].find_all(['td','th'])]
        bgs0   = [c.get('bgcolor','') or '' for c in data_rows[i].find_all(['td','th'])]
        if len(cells0) < 2 or cells0[0] in known_params:
            i += 1; continue

        name, stype, port, is_own = parse_ship_name(cells0[0])
        if not name or name.upper() in EXCLUDE_SHIPS:
            i += 1; continue

        ship = {
            'name': name, 'type': stype, 'port': port, 'is_own': is_own,
            'data': {'风级':{},'潮差':{},'能见度':{},'浪高':{}}
        }
        ships.append(ship)

        # 风级：cells[2:]
        for di, val in enumerate(cells0[2:]):
            if di >= len(col_dts): break
            d, t = col_dts[di]
            bg = bgs0[di+2] if di+2 < len(bgs0) else ''
            if val and val not in ('-',''):
                ship['data']['风级'][f"{d} {t}"] = {'value': val, 'bg': bg}

        # 潮差 / 能见度 / 浪高：cells[1:]
        for offset, param in [(1,'潮差'),(2,'能见度'),(3,'浪高')]:
            if i + offset >= len(data_rows): break
            cells = [c.get_text(strip=True) for c in data_rows[i+offset].find_all(['td','th'])]
            bgs   = [c.get('bgcolor','') or '' for c in data_rows[i+offset].find_all(['td','th'])]
            if cells and cells[0] == param:
                for di, val in enumerate(cells[1:]):
                    if di >= len(col_dts): break
                    d, t = col_dts[di]
                    bg = bgs[di+1] if di+1 < len(bgs) else ''
                    if val and val not in ('-',''):
                        ship['data'][param][f"{d} {t}"] = {'value': val, 'bg': bg}

        i += 4

    # 去重
    seen = set(); unique = []
    for s in ships:
        if s['name'] not in seen:
            seen.add(s['name']); unique.append(s)
    return unique, col_dts

# ── 从缓存加载 ────────────────────────────────────────────────────────────────

def load_cached():
    try:
        with open("hifleet_weather_76ships.json", encoding="utf-8") as f:
            d = json.load(f)
        ships = d.get("ships", [])
        # 提取col_dts：(date, time) 元组列表
        sample = ships[0]['data'] if ships else {}
        all_keys = set()
        for p in ['风级','能见度','浪高']:
            all_keys.update(sample.get(p,{}).keys())
        # 解析 "2026-07-14 0600" → ("2026-07-14", "0600")
        col_dts = []
        for key in sorted(all_keys):
            parts = key.rsplit(' ', 1)
            if len(parts) == 2:
                col_dts.append((parts[0], parts[1]))
            else:
                col_dts.append((key, ''))
        return ships, col_dts
    except Exception as e:
        print(f"Cache error: {e}")
        return [], []

# ── 生成 HTML ────────────────────────────────────────────────────────────────

def build_dashboard(ships, col_dts, output_path="output/index.html"):
    os.makedirs("output", exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")

    # 按日期分组时段
    date_groups = defaultdict(list)
    for cd in col_dts:
        date_groups[cd[0]].append(cd)

    # 统计
    alert_ships = []
    wave3_ships = {}  # ship -> [days with wave>=3]
    for ship in ships:
        max_w = 0
        max_wind = 0
        for key, vd in ship['data']['浪高'].items():
            try: max_w = max(max_w, float(vd['value']))
            except: pass
        for key, vd in ship['data']['风级'].items():
            try: max_wind = max(max_wind, float(vd['value']))
            except: pass
        if max_w >= WAVE_RED or max_wind > WIND_THRESHOLD:
            alert_ships.append({
                'name': ship['name'], 'type': ship['type'],
                'max_wave': max_w, 'max_wind': max_wind
            })
        # wave3 per day
        wave3_ships[ship['name']] = set()
        for key, vd in ship['data']['浪高'].items():
            try:
                if float(vd['value']) >= WAVE_RED:
                    wave3_ships[ship['name']].add(key.split(' ')[0])
            except: pass

    dates = sorted(date_groups.keys())

    # HTML
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>HiFleet 船舶7天气象预报</title>
<style>
* {{ box-sizing:border-box; margin:0; padding:0; }}
body {{ font-family:"Microsoft YaHei","PingFang SC",Arial,sans-serif;
        background:#1a1a2e; color:#e0e0e0; padding:16px; font-size:13px; }}
a {{ color:#7aa2f7; }}

/* ── 页眉 ── */
.header {{ background:linear-gradient(135deg,#16213e,#0f3460); border-radius:12px;
           padding:20px; margin-bottom:20px; text-align:center; }}
.header h1 {{ color:#e94560; font-size:1.6em; margin-bottom:6px; }}
.header p {{ color:#888; font-size:0.8em; }}
.stats {{ display:flex; gap:16px; justify-content:center; margin-top:14px; flex-wrap:wrap; }}
.stat {{ background:rgba(255,255,255,0.07); border-radius:10px; padding:12px 20px; text-align:center; }}
.stat .n {{ font-size:1.8em; font-weight:700; color:#e94560; }}
.stat .l {{ font-size:0.7em; color:#888; margin-top:3px; }}

/* ── 颜色图例 ── */
.legend-wrap {{ background:#16213e; border-radius:10px; padding:14px 18px; margin-bottom:20px; }}
.legend-title {{ color:#e94560; font-size:0.85em; font-weight:600; margin-bottom:10px; }}
.legend {{ display:flex; gap:18px; flex-wrap:wrap; font-size:0.75em; align-items:center; }}
.legend-item {{ display:flex; align-items:center; gap:5px; }}
.legend-dot {{ width:14px; height:14px; border-radius:2px; border:1px solid #444; }}
.legend-param {{ color:#888; }}

/* ── 今日告警 ── */
.alert-box {{ background:#16213e; border-left:4px solid #e94560;
             border-radius:8px; padding:14px 18px; margin-bottom:20px; }}
.alert-box h3 {{ color:#e94560; font-size:0.9em; margin-bottom:10px; }}
.alert-table {{ width:100%; border-collapse:collapse; font-size:0.8em; }}
.alert-table th {{ background:#0f3460; color:#fff; padding:7px 12px; text-align:center; }}
.alert-table td {{ padding:6px 12px; text-align:center; border-bottom:1px solid #2d3585; }}
.alert-table tr:hover td {{ background:rgba(255,255,255,0.03); }}
.wave-red {{ color:#ff4444; font-weight:700; }}
.wind-red {{ color:#ff6600; font-weight:700; }}

/* ── 每日浪高卡片 ── */
.wave-section {{ margin-bottom:24px; }}
.wave-section h3 {{ color:#e94560; font-size:0.9em; margin-bottom:12px; }}
.week-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(180px,1fr)); gap:10px; }}
.day-card {{ background:#16213e; border-radius:8px; padding:10px 12px; border:1px solid #2d3585; }}
.day-card.today {{ border-color:#e94560; }}
.day-label {{ font-size:0.75em; color:#888; margin-bottom:6px; }}
.today .day-label {{ color:#e94560; font-weight:600; }}
.ship-row {{ display:flex; justify-content:space-between; align-items:center;
             padding:3px 0; border-bottom:1px solid #222; font-size:0.78em; }}
.ship-row:last-child {{ border-bottom:none; }}
.ship-row .nm {{ color:#fff; flex:1; }}
.ship-row .wv {{ color:#ff6600; font-weight:600; min-width:36px; text-align:right; }}
.no-alert {{ color:#5eb92d; font-size:0.78em; }}

/* ── 主表格 ── */
.main-table-section {{ margin-bottom:24px; }}
.main-table-section h3 {{ color:#e94560; font-size:0.9em; margin-bottom:10px; }}
.table-scroll {{ overflow-x:auto; border-radius:8px; background:#16213e; }}
table.main {{ border-collapse:collapse; font-size:0.72em; white-space:nowrap; }}
table.main th {{ background:#0f3460; color:#fff; padding:6px 5px; text-align:center;
                position:sticky; top:0; z-index:2; }}
th.dt {{ background:#1a3a6e; color:#aaccff; font-size:0.68em; font-weight:400; }}
th.dt.today-col {{ background:#e94560; color:#fff; }}
th.ship-h {{ text-align:left; min-width:140px; background:#1a3a6e; }}
td {{ padding:4px 5px; text-align:center; border:1px solid #222; }}
td.ship {{ text-align:left; color:#fff; font-weight:600; font-size:0.78em; }}
td.etype {{ text-align:left; color:#888; font-size:0.7em; }}
td .param {{ font-size:0.62em; color:#888; display:block; }}
td.v {{ font-weight:600; font-size:0.78em; }}
tr:hover td {{ background:rgba(255,255,255,0.03); }}
tr.param-row td {{ background:#1e2340; }}

/* 单元格背景色 */
.wind-cell {{ background:#E2EFDA !important; }}
.tide-cell {{ background:#FFF2CC !important; }}
.vis-cell  {{ background:#DEEBF7 !important; }}
.wave-cell {{ background:#FCE4D6 !important; }}

/* 船型颜色 */
.vloc-own td.ship {{ background:#E8D5F5 !important; }}
.vloc    td.ship {{ background:#DDEEFF !important; }}
.cape    td.ship {{ background:#FFE0CC !important; }}
.pana    td.ship {{ background:#FFFFCC !important; }}
ultra    td.ship {{ background:#CCFFE0 !important; }}
supra    td.ship {{ background:#FFFFE0 !important; }}
handy    td.ship {{ background:#E0FFFF !important; }}
.unkn    td.ship {{ background:#F5F5F5 !important; color:#ccc !important; }}

/* 滚动提示 */
.hint {{ color:#555; font-size:0.7em; text-align:right; margin-bottom:4px; }}
</style>
</head>
<body>

<!-- 页眉 -->
<div class="header">
  <h1>HiFleet 船舶7天气象预报</h1>
  <p>76艘船舶 &middot; 更新: {today} &middot; 告警: 风级&gt;6 &nbsp;|&nbsp; 浪高&ge;3m &nbsp;|&nbsp; 能见度&lt;5海里</p>
  <div class="stats">
    <div class="stat"><div class="n">76</div><div class="l">监控船舶</div></div>
    <div class="stat"><div class="n">{len(dates)}</div><div class="l">预报天数</div></div>
    <div class="stat"><div class="n">{len(alert_ships)}</div><div class="l">告警船舶</div></div>
    <div class="stat"><div class="n">{sum(len(v) for v in wave3_ships.values())}</div><div class="l">高浪预警次数</div></div>
  </div>
</div>

<!-- 颜色图例 -->
<div class="legend-wrap">
  <div class="legend-title">颜色说明</div>
  <div class="legend">
    <div class="legend-item"><div class="legend-dot" style="background:#C6EFCE"></div><span>风1-4级 / 浪&lt;2.5m / 能见≥10nmi</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#FFEB9C"></div><span>风5-9级 / 浪2.5-3m / 能见5-10nmi</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#FF6600"></div><span>风≥10级 / 浪≥3m / 能见&lt;5nmi</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#E8D5F5"></div><span>VLOC Own</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#DDEEFF"></div><span>VLOC</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#FFE0CC"></div><span>Capesize</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#FFFFCC"></div><span>Panamax</span></div>
    <div class="legend-item"><div class="legend-dot" style="background:#CCFFE0"></div><span>Ultramax</span></div>
  </div>
</div>

<!-- 今日告警表 -->
<div class="alert-box">
  <h3>今日({today})告警船舶 ({len(alert_ships)}艘)</h3>
"""
    if alert_ships:
        html += """  <table class="alert-table">
    <tr><th>船名</th><th>船型</th><th>最大浪高(m)</th><th>最大风级</th><th>告警原因</th></tr>
"""
        for s in sorted(alert_ships, key=lambda x: -x['max_wave']):
            reasons = []
            if s['max_wave'] >= WAVE_RED: reasons.append(f'浪高{s["max_wave"]}m')
            if s['max_wind'] > WIND_THRESHOLD: reasons.append(f'风级{s["max_wind"]}')
            w_cls = 'wave-red' if s['max_wave'] >= WAVE_RED else ''
            wd_cls = 'wind-red' if s['max_wind'] > WIND_THRESHOLD else ''
            html += f'    <tr><td style="text-align:left;color:#fff">{s["name"]}</td><td>{s["type"]}</td>'
            html += f'<td class="{w_cls}">{s["max_wave"]}</td>'
            html += f'<td class="{wd_cls}">{s["max_wind"]}</td>'
            html += f'<td style="text-align:left;color:#ff9900">{"; ".join(reasons)}</td></tr>\n'
        html += "  </table>\n"
    else:
        html += '  <p class="no-alert">✓ 今日无告警</p>\n'
    html += "</div>\n"

    # 每日浪高卡片
    html += """<div class="wave-section">
  <h3>未来一周浪高≥3m 船舶（按日期）</h3>
  <div class="week-grid">
"""
    for d in dates:
        is_today = (d == today)
        day_ships = []
        for ship in ships:
            max_w = 0
            for key, vd in ship['data']['浪高'].items():
                if key.startswith(d):
                    try: max_w = max(max_w, float(vd['value']))
                    except: pass
            if max_w >= WAVE_RED:
                day_ships.append((ship['name'], max_w))
        day_ships.sort(key=lambda x: -x[1])

        html += f'  <div class="day-card{" today" if is_today else ""}">'
        html += f'    <div class="day-label">{"★ " if is_today else ""}{d}</div>'
        if day_ships:
            for nm, wv in day_ships[:6]:
                html += f'    <div class="ship-row"><span class="nm">{nm}</span><span class="wv">{wv}m</span></div>\n'
            if len(day_ships) > 6:
                html += f'    <div class="no-alert">...还有{len(day_ships)-6}艘</div>\n'
        else:
            html += '    <div class="no-alert">✓ 无高浪</div>\n'
        html += '  </div>\n'
    html += """</div>
</div>
"""

    # ── 主表格（与邮件1:1匹配）────────────────────────────────────────────
    html += """<div class="main-table-section">
  <h3>全览详细表格（7天 × 4时段 预报）</h3>
  <div class="hint">← 左右滑动查看全部时段 →</div>
  <div class="table-scroll">
  <table class="main">
"""
    # 表头行1：日期
    html += "<thead>\n<tr>\n"
    html += '<th class="ship-h" rowspan="2">船名 / 船型</th>\n'
    for d in dates:
        times = date_groups[d]
        is_today = (d == today)
        cls = 'today-col' if is_today else 'dt'
        if len(times) > 1:
            html += f'<th class="{cls}" colspan="{len(times)}">{d}</th>\n'
        else:
            html += f'<th class="{cls}">{d}</th>\n'
    html += "</tr>\n<tr>\n"
    # 表头行2：时间
    for d in dates:
        for (_, t) in date_groups[d]:
            is_today = (d == today)
            cls = 'today-col' if is_today else 'dt'
            html += f'<th class="{cls}">{t}</th>\n'
    html += "</tr>\n</thead>\n<tbody>\n"

    # 数据行
    type_class_map = {
        'VLOC Own': 'vloc-own', 'VLOC': 'vloc', 'Capesize': 'cape',
        'Panamax': 'pana', 'Ultramax': 'ultra', 'Supramax': 'supra',
        'Handymax': 'handy', 'Unknown': 'unkn'
    }

    for ship in ships:
        tcls = type_class_map.get(ship['type'], 'unkn')
        row_key = f'{ship["name"]}_{ship["type"]}'

        # 风级行
        html += f'<tr class="{tcls}">'
        html += f'<td class="ship">{ship["name"]}<span class="etype">{ship["type"]}{" ★" if ship.get("is_own") else ""}</span></td>\n'
        for d, t in col_dts:
            key = f"{d} {t}"
            vd  = ship['data']['风级'].get(key, {})
            val = vd.get('value', '')
            try:
                lv = float(val)
                bg = wind_color(val)
                cls = 'v wind-cell'
            except:
                bg = '#E2EFDA'; cls = 'v wind-cell'
            style = f'background:{bg}' if bg != '#E2EFDA' else ''
            html += f'<td class="{cls}" style="{style}">{val}</td>\n'
        html += "</tr>\n"

        # 潮差行
        html += '<tr class="param-row">'
        html += '<td class="ship"><span class="param">潮差</span></td>\n'
        for d, t in col_dts:
            key = f"{d} {t}"
            vd  = ship['data']['潮差'].get(key, {})
            val = vd.get('value', '')
            bg  = vd.get('bg', '') or '#FFF2CC'
            style = f'background:{bg}' if bg not in ('','None') else ''
            html += f'<td class="v tide-cell" style="{style}">{val}</td>\n'
        html += "</tr>\n"

        # 能见度行
        html += '<tr class="param-row">'
        html += '<td class="ship"><span class="param">能见度</span></td>\n'
        for d, t in col_dts:
            key = f"{d} {t}"
            vd  = ship['data']['能见度'].get(key, {})
            val = vd.get('value', '')
            try:
                bg = vis_color(val)
                style = f'background:{bg}' if bg != '#DEEBF7' else ''
            except:
                style = ''
            html += f'<td class="v vis-cell" style="{style}">{val}</td>\n'
        html += "</tr>\n"

        # 浪高行
        html += '<tr class="param-row">'
        html += '<td class="ship"><span class="param">浪高</span></td>\n'
        for d, t in col_dts:
            key = f"{d} {t}"
            vd  = ship['data']['浪高'].get(key, {})
            val = vd.get('value', '')
            try:
                bg = wave_color(val)
                style = f'background:{bg}; color:#000'
                if float(val) >= WAVE_RED: style += '; font-weight:700'
            except:
                style = ''
            html += f'<td class="v wave-cell" style="{style}">{val}</td>\n'
        html += "</tr>\n"

    html += "</tbody>\n</table>\n</div>\n</div>\n"

    # 页脚
    html += f"""<div style="text-align:center; color:#555; font-size:0.72em; padding:20px 0;">
  HiFleet 船舶气象预报自动同步 &middot; 数据更新: {today} &middot;
  排除: ORE CHINA, ORE DONGJIAKOU, ORE HEBEI, ORE SHANDONG
</div>
</body></html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Saved: {output_path}")


def main():
    print("Fetching email...")
    html = fetch_hifleet_email()

    if html:
        print("Parsing HTML...")
        try:
            from bs4 import BeautifulSoup
            ships, col_dts = parse_html(html)
            print(f"Parsed {len(ships)} ships, {len(col_dts)} time slots")
        except ImportError:
            print("BeautifulSoup not available, using cached data")
            ships, col_dts = load_cached()
    else:
        print("No email - loading cache...")
        ships, col_dts = load_cached()

    if not ships:
        print("ERROR: No ship data")
        with open("output/index.html","w",encoding="utf-8") as f:
            f.write("<html><body style='background:#1a1a2e;color:#e94560;padding:40px;font-family:Arial;text-align:center'><h1>No data</h1></body></html>")
    else:
        build_dashboard(ships, col_dts)
    print("Done!")


if __name__ == "__main__":
    main()
