#!/usr/bin/env python3
"""
HiFleet Dashboard — 100% 匹配邮件 HTML 排版格式
"""
import json, re, os
from datetime import datetime
from collections import defaultdict

# ── 颜色映射（与邮件 CSS 完全一致）
WAVE_RED = 3.0
WIND_RED = 7
VIS_THRESHOLD = 5

def wind_cls(level):
    try:
        v = int(float(level))
        v = max(1, min(10, v))
        return 'wind%d' % v
    except:
        return 'wind1'

def wave_cls(level):
    try:
        v = float(level)
        if v <= 0: return 'wave1'
        if v <= 1: return 'wave1'
        if v <= 2: return 'wave2'
        if v <= 3: return 'wave3'
        if v <= 4: return 'wave4'
        if v <= 5: return 'wave5'
        if v <= 6: return 'wave6'
        if v <= 7: return 'wave7'
        if v <= 8: return 'wave8'
        if v <= 9: return 'wave9'
        return 'wave10'
    except:
        return 'wave1'

def tide_cls(level):
    try:
        v = float(level)
        if v <= 0: return 'tide1'
        if v <= 1: return 'tide2'
        if v <= 2: return 'tide3'
        if v <= 3: return 'tide4'
        return 'tide5'
    except:
        return 'tide1'

# ── 加载数据
def load_data():
    path = r'C:\Users\HKMW\Desktop\hifleet_76.json'
    if not os.path.exists(path):
        path = r'C:\Users\HKMW\.mavis\agents\mavis\workspace\hifleet_weather_76ships.json'
    with open(path, encoding='utf-8') as f:
        d = json.load(f)
    ships = d['ships']
    col_dts = d.get('col_dts', [])
    if not col_dts or len(col_dts) <= 7:
        # 从船舶数据中直接提取所有唯一的 "日期 时间" 键
        all_keys = set()
        for ship in ships:
            for p in ['风级', '涌差', '能见度', '浪高']:
                all_keys.update(ship['data'].get(p, {}).keys())
        col_dts = []
        for key in sorted(all_keys):
            parts = key.rsplit(' ', 1)
            col_dts.append((parts[0], parts[1]) if len(parts) == 2 else (key, ''))
    return ships, col_dts

# ── HTML 片段
CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 11px; text-align: center; background: #f4f4f4; }
table { border-collapse: collapse; text-align: center; font-size: 11px; width: 100%; }
th, td { border: 1px solid black; padding: 2px; vertical-align: middle; }
th { background-color: lightgray; }
.vessel-type {
    background-color: gold; border: 1px solid #ccc; padding: 1px 4px;
    box-shadow: 2px 2px 4px rgba(0,0,0,0.4); color: blue;
    font-weight: bold; display: inline-block; font-size: 10px; margin-top: 1px;
}
.ship-name { font-weight: bold; font-size: 11px; text-align: left; vertical-align: top; }
.ship-cell { text-align: left; vertical-align: top; background: #f9f9f9; }
.param-label { font-size: 10px; color: #444; width: 42px; }
.wind1  { color:#e8f5e9; background-color: #e8f5e9; }
.wind2  { color:#c8e6c9; background-color: #c8e6c9; }
.wind3  { color:#a5d6a7; background-color: #a5d6a7; }
.wind4  { color:#81c784; background-color: #81c784; }
.wind5  { color:#0097a7; background-color: #66bb6a; }
.wind6  { color:#ffea00; background-color: #e1bee7; }
.wind7  { color:#fdd835; background-color: #ce93d8; }
.wind8  { color:#ffb300; background-color: #ab47bc; }
.wind9  { color:#ffe0b2; background-color: #8e24aa; }
.wind10 { color:#fff9c4; background-color: #6a1b9a; }
.tide1  { color:#616161; background-color: white; }
.tide2  { color:white;    background-color: #795548; }
.tide3  { color:white;    background-color: #5d4037; }
.tide4  { color:white;    background-color: #4e342e; }
.tide5  { color:white;    background-color: #3e2723; }
.vis-ok  { background-color: white; }
.vis-warn { background-color: #ff6666; color: white; font-weight: bold; }
.wave1  { color:#616161; background-color: #e1f5fe; }
.wave2  { color:#616161; background-color: #81d4fa; }
.wave3  { color:#616161; background-color: #29b6f6; }
.wave4  { color:#f44336; background-color: #fff59d; }
.wave5  { color:#f44336; background-color: #ffeb3b; }
.wave6  { color:#f44336; background-color: #ffd54f; }
.wave7  { color:#f44336; background-color: #ffc107; }
.wave8  { color:white;   background-color: #ffa000; }
.wave9  { color:white;   background-color: #ff6f00; }
.wave10 { color:white;   background-color: #bf360c; }
.today-hdr { background-color: #e94560 !important; color: white !important; font-weight: bold; }
.alert-section { background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; margin: 10px auto; max-width: 1100px; text-align: left; }
.alert-section h3 { color: #d32f2f; border-bottom: 2px solid #d32f2f; padding-bottom: 4px; margin-bottom: 8px; }
.alert-table { width: 100%; border-collapse: collapse; font-size: 11px; }
.alert-table th { background: #d32f2f; color: white; padding: 4px 8px; }
.alert-table td { padding: 4px 8px; border-bottom: 1px solid #eee; }
.wave-hi { color: #d32f2f; font-weight: bold; }
.wind-hi  { color: #e65100; font-weight: bold; }
.legend-section { background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; margin: 10px auto; max-width: 1100px; text-align: left; }
.legend-section h3 { border-bottom: 2px solid #333; padding-bottom: 4px; margin-bottom: 8px; }
.legend-row { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; margin-bottom: 6px; }
.legend-label { font-weight: bold; min-width: 70px; }
.legend-item { display: inline-flex; align-items: center; gap: 2px; font-size: 10px; }
.legend-box { width: 16px; height: 12px; border: 1px solid #999; display: inline-block; }
.page-header { background: #1a237e; color: white; padding: 12px; margin-bottom: 10px; }
.page-header h1 { font-size: 16px; margin-bottom: 4px; }
.page-header p { font-size: 11px; color: #bbdefb; }
.table-wrap { overflow-x: auto; max-height: 85vh; overflow-y: auto; }
.sticky-t { position: sticky; top: 0; z-index: 10; }
.ship-col { position: sticky; left: 0; z-index: 5; background: #f9f9f9; min-width: 130px; }
"""

def H(txt):
    """HTML 安全转义"""
    return str(txt).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def build_html(ships, col_dts, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    today = datetime.now().strftime('%Y-%m-%d')

    date_groups = defaultdict(list)
    for cd in col_dts:
        date_groups[cd[0]].append(cd)
    dates = sorted(date_groups.keys())

    # 告警统计
    alert_ships = []
    wave3_days = {}  # date -> {ship_name: max_wave}
    for ship in ships:
        max_w = 0; max_wind = 0
        for vd in ship['data']['浪高'].values():
            try: max_w = max(max_w, float(vd['value']))
            except: pass
        for vd in ship['data']['风级'].values():
            try: max_wind = max(max_wind, float(vd['value']))
            except: pass
        if max_w >= WAVE_RED or max_wind > WIND_RED:
            alert_ships.append({'name': ship['name'], 'type': ship['type'],
                               'max_wave': max_w, 'max_wind': max_wind})
        for key, vd in ship['data']['浪高'].items():
            try:
                if float(vd['value']) >= WAVE_RED:
                    d = key.split(' ')[0]
                    if d not in wave3_days:
                        wave3_days[d] = {}
                    wave3_days[d][ship['name']] = max(
                        wave3_days[d].get(ship['name'], 0), float(vd['value']))
            except: pass

    # 收集所有段落
    parts = []

    # HTML 头部
    parts.append('<!DOCTYPE html>\n<html><head>\n<meta charset="UTF-8">\n')
    parts.append('<title>HiFleet 船舶7天预报</title>\n')
    parts.append('<style>\n%s\n</style>\n</head><body>\n' % CSS)

    # 页眉
    parts.append('<div class="page-header">\n')
    parts.append('<h1>HiFleet 船舶7天气象预报</h1>\n')
    parts.append('<p>76艘船舶 &middot; 更新: %s &middot; 告警: 风级&gt;6 &nbsp;|&nbsp; 浪高&gt;=3m &nbsp;|&nbsp; 能见度&lt;5海里</p>\n' % today)
    parts.append('</div>\n')

    # 告警表
    parts.append('<div class="alert-section">\n')
    parts.append('<h3>告警船舶 (%d艘)</h3>\n' % len(alert_ships))
    if alert_ships:
        parts.append('<table class="alert-table"><tr><th>船名</th><th>船型</th><th>最大浪高(m)</th><th>最大风级</th><th>告警原因</th></tr>\n')
        for s in sorted(alert_ships, key=lambda x: -x['max_wave']):
            reasons = []
            if s['max_wave'] >= WAVE_RED: reasons.append('浪高%.1fm' % s['max_wave'])
            if s['max_wind'] > 6: reasons.append('风级%d' % int(s['max_wind']))
            wc = 'wave-hi' if s['max_wave'] >= WAVE_RED else ''
            wdc = 'wind-hi' if s['max_wind'] > 6 else ''
            parts.append('<tr><td>%s</td><td>%s</td><td class="%s">%.1f</td><td class="%s">%d</td><td>%s</td></tr>\n' % (
                H(s['name']), H(s['type']), wc, s['max_wave'], wdc, int(s['max_wind']), '; '.join(reasons)))
        parts.append('</table>\n')
    else:
        parts.append('<p style="color:green">今日无告警</p>\n')
    parts.append('</div>\n')

    # 浪高告警按日
    parts.append('<div class="alert-section">\n')
    parts.append('<h3>未来一周浪高&gt;=3m 船舶（按日期）</h3>\n')
    for d in dates:
        is_today = d == today
        day_dict = wave3_days.get(d, {})
        day_ships = sorted(day_dict.items(), key=lambda x: -x[1])
        parts.append('<p><b>%s%s:</b> ' % ('&#9733; ' if is_today else '', d))
        if day_ships:
            names = ' '.join(['%s %.1fm' % (H(n), w) for n, w in day_ships[:8]])
            parts.append(names)
            if len(day_ships) > 8:
                parts.append(' ...还有%d艘' % (len(day_ships) - 8))
        else:
            parts.append('无')
        parts.append('</p>\n')
    parts.append('</div>\n')

    # 图例
    parts.append('<div class="legend-section">\n')
    parts.append('<h3>颜色图例</h3>\n')
    parts.append('<div class="legend-row">\n')
    parts.append('<span class="legend-label">风级:</span>\n')
    wind_c = ['#e8f5e9','#c8e6c9','#a5d6a7','#81c784','#66bb6a','#e1bee7','#ce93d8','#ab47bc','#8e24aa','#6a1b9a']
    for i in range(1, 11):
        lbl = '%d级' % i if i < 10 else '&gt;=10级'
        parts.append('<span class="legend-item"><span class="legend-box" style="background:%s"></span>%s</span>\n' % (wind_c[i-1], lbl))
    parts.append('</div>\n')
    parts.append('<div class="legend-row">\n')
    parts.append('<span class="legend-label">浪高:</span>\n')
    wave_c = ['#e1f5fe','#81d4fa','#29b6f6','#fff59d','#ffeb3b','#ffd54f','#ffc107','#ffa000','#ff6f00','#bf360c']
    wave_l = ['&lt;2m','2-3m','3-4m','4-5m','5-6m','6-7m','7-8m','8-9m','9-10m','&gt;=10m']
    for i, (c, l) in enumerate(zip(wave_c, wave_l)):
        parts.append('<span class="legend-item"><span class="legend-box" style="background:%s"></span>%s</span>\n' % (c, l))
    parts.append('</div>\n')
    parts.append('<div class="legend-row">\n')
    parts.append('<span class="legend-label">涌差:</span>\n')
    tide_c = ['white','#795548','#5d4037','#4e342e','#3e2723']
    for i, c in enumerate(tide_c):
        parts.append('<span class="legend-item"><span class="legend-box" style="background:%s"></span></span>\n' % c)
    parts.append('</div>\n')
    parts.append('</div>\n')

    # 主表格
    parts.append('<div class="table-wrap">\n')
    parts.append('<table>\n')

    # 表头行1：船名 | 风浪 | 日期
    parts.append('<thead>\n')
    parts.append('<tr class="sticky-t">\n')
    parts.append('<th rowspan="2" class="sticky-t" style="min-width:130px;">船名</th>\n')
    parts.append('<th rowspan="2" class="sticky-t" style="width:42px;">风浪</th>\n')
    for d in dates:
        times = date_groups[d]
        cls = 'today-hdr' if d == today else ''
        parts.append('<th colspan="%d" class="sticky-t %s">%s</th>\n' % (len(times), cls, d))
    parts.append('</tr>\n')

    # 表头行2：时间（28列 = 7天 × 4时段）
    parts.append('<tr class="sticky-t">\n')
    time_map = {0: '0600', 1: '1200', 2: '1800', 3: '2400'}
    for d, t in col_dts:
        cls = 'today-hdr' if d == today else ''
        # 每天重复 4 个时段
        for slot_i in range(4):
            parts.append('<th class="sticky-t %s">%s</th>\n' % (cls, time_map[slot_i]))
    parts.append('</tr>\n')
    parts.append('</thead>\n')

    # 数据行
    # 关键：col_dts 只有 7 条（每天一个值），但 HTML 表头有 28 列
    # 风级行：按 col_dts 逐个填入（重复当天值填充 4 个时段格）
    # 其他行（涌差/能见度/浪高）：每个 col_dt 值填入 4 格（跨越同一日期的 4 个时段）
    parts.append('<tbody>\n')
    for ship in ships:
        # ── 风级行（rowspan=4 跨4行）────────────────────────
        parts.append('<tr>\n')
        parts.append('<td class="ship-col ship-cell" rowspan="4">\n')
        parts.append('<div class="ship-name">%s</div>\n' % H(ship['name']))
        parts.append('<div class="vessel-type">%s</div>\n' % H(ship['type']))
        parts.append('</td>\n')
        parts.append('<td class="param-label">风级</td>\n')
        # col_dts 只有 7 条，填充 28 格
        di = 0
        for d, t in col_dts:
            key = '%s %s' % (d, t)
            vd = ship['data'].get('风级', {}).get(key, {})
            val = vd.get('value', '')
            cls = wind_cls(val) if val and val not in ('-', '') else 'wind1'
            if val in ('-', ''): val = '-'
            # 每天填 4 格（同一天的值重复）
            for _ in range(4):
                parts.append('<td class="%s">%s</td>\n' % (cls, val))
        parts.append('</tr>\n')

        # ── 涌差行 ─────────────────────────────────────
        parts.append('<tr>\n')
        parts.append('<td class="param-label">涌差</td>\n')
        for d, t in col_dts:
            vd = ship['data'].get('涌差', {}).get('%s %s' % (d, t), {})
            val = vd.get('value', '')
            cls = tide_cls(val) if val and val not in ('-', '') else 'tide1'
            if val in ('-', ''): val = '-'
            for _ in range(4):
                parts.append('<td class="%s">%s</td>\n' % (cls, val))
        parts.append('</tr>\n')

        # ── 能见度行 ────────────────────────────────────
        parts.append('<tr>\n')
        parts.append('<td class="param-label">能见度</td>\n')
        for d, t in col_dts:
            vd = ship['data'].get('能见度', {}).get('%s %s' % (d, t), {})
            val = vd.get('value', '')
            if val and val not in ('-', ''):
                try:
                    cls = 'vis-warn' if float(val) < VIS_THRESHOLD else 'vis-ok'
                except:
                    cls = 'vis-ok'
            else:
                cls = 'vis-ok'
                val = '-'
            for _ in range(4):
                parts.append('<td class="%s">%s</td>\n' % (cls, val))
        parts.append('</tr>\n')

        # ── 浪高行 ────────────────────────────────────
        parts.append('<tr>\n')
        parts.append('<td class="param-label">浪高</td>\n')
        for d, t in col_dts:
            vd = ship['data'].get('浪高', {}).get('%s %s' % (d, t), {})
            val = vd.get('value', '')
            cls = wave_cls(val) if val and val not in ('-', '') else 'wave1'
            if val in ('-', ''): val = '-'
            for _ in range(4):
                parts.append('<td class="%s">%s</td>\n' % (cls, val))
        parts.append('</tr>\n')

    parts.append('</tbody>\n')
    parts.append('</table>\n</div>\n')

    # 页脚
    parts.append('<div style="text-align:center; padding:12px; color:#666; font-size:11px;">\n')
    parts.append('HiFleet 船舶气象预报自动同步 &middot; 数据更新: %s &middot; 排除: ORE CHINA, ORE DONGJIAKOU, ORE HEBEI, ORE SHANDONG\n' % today)
    parts.append('</div>\n</body></html>')

    html = ''.join(parts)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print('Saved: %s (%d bytes)' % (output_path, len(html)))


if __name__ == '__main__':
    ships, col_dts = load_data()
    print('Ships: %d, Slots: %d' % (len(ships), len(col_dts)))
    build_html(ships, col_dts, r'C:\Users\HKMW\Desktop\hifleet_email_style.html')
