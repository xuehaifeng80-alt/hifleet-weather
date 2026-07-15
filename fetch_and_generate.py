#!/usr/bin/env python3
"""
HiFleet Weather Dashboard Generator (Repo/CI Version)
完全匹配邮件原始排版格式：
  - 表头：日期行 + 时间行（0600/1200/1800/2400 × 7天 = 28时段）
  - 每船4行：风级 / 潮差 / 能见度 / 浪高
  - 颜色规则与邮件Excel一致
  - 28列直接对应28个时段，无 colspan 合并
"""
import os
import json
import re
from datetime import datetime
from collections import defaultdict

# === 找 JSON 文件（CI 和本地通用）===
def find_json():
    candidates = [
        "hifleet_weather_76ships.json",
        "../hifleet_weather_76ships.json",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("hifleet_weather_76ships.json not found in current or parent dir")

def load_data():
    path = find_json()
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    ships = d["ships"]
    col_dts = d.get("col_dts", [])
    return ships, col_dts

# === Config ===
WIND_RED   = 6
WAVE_RED   = 3.0
VIS_THRESH = 5

WIND_COLORS = {  # 蒲福风级 → 背景色
    (1, 4):   "#C6EFCE",  # 绿色 1-4级
    (5, 5):   "#FFEB9C",  # 黄色 5级
    (6, 6):   "#E2A7D7",  # 粉紫色 6级
    (7, 9):   "#FF6600",  # 橙色 7-9级
    (10, 99): "#FF0000",  # 红色 >=10级
}

WIND_CSS = {
    "wind1":  "#C6EFCE",
    "wind2":  "#C6EFCE",
    "wind3":  "#C6EFCE",
    "wind4":  "#C6EFCE",
    "wind5":  "#FFEB9C",
    "wind6":  "#E2A7D7",
    "wind7":  "#FF6600",
    "wind8":  "#FF6600",
    "wind9":  "#FF6600",
    "wind10": "#FF0000",
}

WAVE_CSS = {
    "wave1":  "#E1F5FE",
    "wave2":  "#81D4FA",
    "wave3":  "#29B6F6",
    "wave4":  "#FFF59D",
    "wave5":  "#FFEB3B",
    "wave6":  "#FFD54F",
    "wave7":  "#FFC107",
    "wave8":  "#FFA000",
    "wave9":  "#FF6F00",
    "wave10": "#BF360C",
}

TIDE_CSS = {
    "tide1":  "#FFFFFF",
    "tide2":  "#795548",
    "tide3":  "#5D4037",
    "tide4":  "#4E342E",
    "tide5":  "#3E2723",
}

SHIP_TYPE_CSS = {
    "VLOC Own":  "#E8D5F5",
    "VLOC":      "#DDEEFF",
    "Capesize":  "#FFE0CC",
    "Panamax":   "#FFFFCC",
    "Ultramax":  "#CCFFE0",
    "Supramax":  "#FFFFE0",
    "Handymax":  "#E0FFFF",
    "Unknown":   "#F5F5F5",
}

PARAM_CSS = {
    "wind":  "#E2EFDA",
    "tide":  "#FFF2CC",
    "vis":   "#DEEBF7",
    "wave":  "#FCE4D6",
}

def wind_cls(level_str):
    try:
        v = int(float(level_str))
        v = max(1, min(10, v))
        return "wind%d" % v
    except:
        return "wind1"

def wave_cls(val_str):
    try:
        v = float(val_str)
        if v <= 1: return "wave1"
        if v <= 2: return "wave2"
        if v <= 3: return "wave3"
        if v <= 4: return "wave4"
        if v <= 5: return "wave5"
        if v <= 6: return "wave6"
        if v <= 7: return "wave7"
        if v <= 8: return "wave8"
        if v <= 9: return "wave9"
        return "wave10"
    except:
        return "wave1"

def tide_cls(val_str):
    try:
        v = float(val_str)
        if v <= 0: return "tide1"
        if v <= 1: return "tide2"
        if v <= 2: return "tide3"
        if v <= 3: return "tide4"
        return "tide5"
    except:
        return "tide1"

def H(txt):
    """HTML escape"""
    return str(txt).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# === CSS ===
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
""" + "\n".join(".%s { color:%s; background-color: %s; }" % (k, v, v) for k, v in {
    **{"wind1":"#e8f5e9","wind2":"#c8e6c9","wind3":"#a5d6a7","wind4":"#81c784",
       "wind5":"#66bb6a","wind6":"#e1bee7","wind7":"#ce93d8","wind8":"#ab47bc",
       "wind9":"#8e24aa","wind10":"#6a1b9a"},
    **{"wave1":"#e1f5fe","wave2":"#81d4fa","wave3":"#29b6f6","wave4":"#fff59d",
       "wave5":"#ffeb3b","wave6":"#ffd54f","wave7":"#ffc107","wave8":"#ffa000",
       "wave9":"#ff6f00","wave10":"#bf360c"},
    **{"tide1":"#FFFFFF","tide2":"#795548","tide3":"#5D4037","tide4":"#4E342E","tide5":"#3E2723"},
    **{"vis-ok":"#FFFFFF","vis-warn":"#FF6600;color:white;font-weight:bold"},
}.items()) + """
.today-hdr { background-color: #e94560 !important; color: white !important; font-weight: bold; }
.alert-section { background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; margin: 10px auto; max-width: 1100px; text-align: left; }
.alert-section h3 { color: #d32f2f; border-bottom: 2px solid #d32f2f; padding-bottom: 4px; margin-bottom: 8px; }
.alert-table { width: 100%; border-collapse: collapse; font-size: 11px; }
.alert-table th { background: #d32f2f; color: white; padding: 4px 8px; }
.alert-table td { padding: 4px 8px; border-bottom: 1px solid #eee; }
.wave-hi { color: #d32f2f; font-weight: bold; }
.wind-hi { color: #e65000; font-weight: bold; }
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
.ship-col.th-head { background: lightgray; z-index: 15; }
.tide-cell { color: white; }
"""

def build_html(ships, col_dts, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")

    # Group col_dts by date for header row 1 (date spans 4 slots)
    date_groups = defaultdict(list)
    for item in col_dts:
        d, t = item
        date_groups[d].append(item)

    dates = sorted(date_groups.keys())
    today_str = datetime.now().strftime("%Y-%m-%d")

    # Alerts
    alert_ships = []
    wave3_days = defaultdict(dict)  # date -> {ship: max_wave}
    for ship in ships:
        max_w, max_wind = 0, 0
        for key, vd in ship["data"].get("浪高", {}).items():
            try: max_w = max(max_w, float(vd["value"]))
            except: pass
        for key, vd in ship["data"].get("风级", {}).items():
            try: max_wind = max(max_wind, float(vd["value"]))
            except: pass
        if max_w >= WAVE_RED or max_wind > WIND_RED:
            alert_ships.append({"name": ship["name"], "type": ship["type"],
                                 "max_wave": max_w, "max_wind": max_wind})
        for key, vd in ship["data"].get("浪高", {}).items():
            try:
                if float(vd["value"]) >= WAVE_RED:
                    d = key.split(" ")[0]
                    n = ship["name"]
                    wave3_days[d][n] = max(wave3_days[d].get(n, 0), float(vd["value"]))
            except: pass

    # Build HTML parts
    parts = []
    parts.append("<!DOCTYPE html>\n<html><head>\n<meta charset=\"UTF-8\">\n")
    parts.append("<title>HiFleet 气象预报 | %s</title>\n" % today)
    parts.append("<style>\n%s\n</style>\n" % CSS)
    parts.append("</head><body>\n")

    # Header
    parts.append("<div class=page-header>\n")
    parts.append("<h1>HiFleet 气象预报 | %s</h1>\n" % today)
    parts.append("<p>共 %d 艘船 &nbsp;|&nbsp; 更新: %s &nbsp;|&nbsp; 风级>%d 或 浪高>=%.1fm 预警</p>\n"
                 % (len(ships), today, WIND_RED, WAVE_RED))
    parts.append("</div>\n")

    # Alert section
    parts.append("<div class=alert-section>\n")
    parts.append("<h3>⚠ 预警船舶 (风级>%d 或 浪高>=%.1fm) (%d艘)</h3>\n" % (WIND_RED, WAVE_RED, len(alert_ships)))
    if alert_ships:
        parts.append("<table class=alert-table><tr><th>船名</th><th>类型</th><th>最大浪高</th><th>最大风级</th><th>原因</th></tr>\n")
        for s in sorted(alert_ships, key=lambda x: -x["max_wave"]):
            reasons = []
            if s["max_wave"] >= WAVE_RED: reasons.append("浪高%.1fm" % s["max_wave"])
            if s["max_wind"] > WIND_RED: reasons.append("风级%d级" % int(s["max_wind"]))
            parts.append("<tr><td>%s</td><td>%s</td><td class=wave-hi>%.1fm</td>"
                         "<td class=wind-hi>%d级</td><td>%s</td></tr>\n" %
                         (H(s["name"]), H(s["type"]), s["max_wave"],
                          int(s["max_wind"]), "; ".join(reasons)))
        parts.append("</table>\n")
    else:
        parts.append("<p style=color:green>✓ 暂无预警</p>\n")
    parts.append("</div>\n")

    # High-wave days
    if wave3_days:
        parts.append("<div class=alert-section>\n")
        parts.append("<h3>🌊 高浪日期 (浪高>=%.1fm)</h3>\n" % WAVE_RED)
        for d in sorted(wave3_days.keys()):
            is_today = (d == today_str)
            ships_d = sorted(wave3_days[d].items(), key=lambda x: -x[1])
            names = " ".join(["%s(%.1fm)" % (n, w) for n, w in ships_d[:8]])
            if len(ships_d) > 8:
                names += " ...另外%d艘" % (len(ships_d) - 8)
            parts.append("<p><b>%s%s:</b> %s</p>\n" % ("▶ " if is_today else "", d, names))
        parts.append("</div>\n")

    # Legend
    parts.append("<div class=legend-section>\n")
    parts.append("<h3>📋 图例说明</h3>\n")
    parts.append("<div class=legend-row><span class=legend-label>风级:</span>\n")
    wind_items = [("1-4级","#C6EFCE"),("5级","#FFEB9C"),("6级","#E2A7D7"),("7-9级","#FF6600"),("≥10级","#FF0000")]
    for label, color in wind_items:
        parts.append("<span class=legend-item><span class=legend-box style='background:%s'></span>%s</span>\n" % (color, label))
    parts.append("</div>\n")
    parts.append("<div class=legend-row><span class=legend-label>浪高:</span>\n")
    wave_items = [("<2m","#E1F5FE"),("2-3m","#29B6F6"),("3-4m","#FFF59D"),("4-5m","#FFEB3B"),("5-6m","#FFD54F"),("6-7m","#FFC107"),("7-8m","#FFA000"),("8-9m","#FF6F00"),("≥10m","#BF360C")]
    for label, color in wave_items:
        parts.append("<span class=legend-item><span class=legend-box style='background:%s'></span>%s</span>\n" % (color, label))
    parts.append("</div>\n")
    parts.append("<div class=legend-row><span class=legend-label>能见度:</span>\n")
    parts.append("<span class=legend-item><span class=legend-box style='background:#fff;border-color:#ccc'></span>≥5km</span>\n")
    parts.append("<span class=legend-item><span class=legend-box style='background:#FF6600;color:white'></span>&lt;5km 预警</span>\n")
    parts.append("</div>\n")
    parts.append("</div>\n")

    # Table
    parts.append("<div class=table-wrap>\n")
    parts.append("<table>\n")

    # Header row 1: 船名 | 风浪 | dates (7 dates × colspan=4)
    parts.append("<thead>\n")
    parts.append("<tr class='sticky-t'>\n")
    parts.append("<th class='sticky-t ship-col th-head' style='min-width:130px'>船名</th>\n")
    parts.append("<th class='sticky-t' style='width:42px'>风浪</th>\n")
    for d in dates:
        is_today = (d == today_str)
        cls = "today-hdr" if is_today else ""
        parts.append("<th class='sticky-t %s' colspan=4>%s</th>\n" % (cls, d))
    parts.append("</tr>\n")

    # Header row 2: 28 time slots
    parts.append("<tr class='sticky-t'>\n")
    parts.append("<th class='sticky-t ship-col th-head'></th>\n")
    parts.append("<th class='sticky-t'></th>\n")
    time_map = {0: "0600", 1: "1200", 2: "1800", 3: "2400"}
    for di, d in enumerate(dates):
        is_today = (d == today_str)
        cls = "today-hdr" if is_today else ""
        for ti in range(4):
            parts.append("<th class='sticky-t %s'>%s</th>\n" % (cls, time_map[ti]))
    parts.append("</tr>\n")
    parts.append("</thead>\n")

    # Body: 4 rows per ship
    parts.append("<tbody>\n")
    param_rows = [
        ("风级", "wind", wind_cls),
        ("潮差", "tide", tide_cls),
        ("能见度", "vis", lambda v: "vis-warn" if (try_float(v) < VIS_THRESH) else "vis-ok"),
        ("浪高", "wave", wave_cls),
    ]
    for ship in ships:
        # Build 4 rows, each with ship-name cell (rowspan=4) + param label + 28 data cells
        rows_data = [[], [], [], []]  # wind, tide, vis, wave

        for si, (label, param_key, cls_fn) in enumerate(param_rows):
            row_cells = []
            # ship name+type cell (first col, only for wind row, rowspan=4)
            if si == 0:
                row_cells.append(
                    "<td class='ship-cell' rowspan=4>"
                    "<div class=ship-name>%s</div>"
                    "<div class=vessel-type>%s</div>"
                    "</td>" % (H(ship["name"]), H(ship["type"]))
                )
            row_cells.append("<td class=param-label>%s</td>" % label)

            # 28 data cells
            for ci, item in enumerate(col_dts):
                d, t = item
                key = "%s %s" % (d, t)
                vd = ship["data"].get(param_key, {}).get(key, {})
                val = vd.get("value", "-")
                bg = vd.get("bg", "")
                cls = cls_fn(val) if val and val != "-" else param_key + "1"
                # tide cells need white text
                extra = " tide-cell" if param_key == "tide" and cls not in ("tide1",) else ""
                if val in ("-", ""):
                    val = "-"
                    cls = param_key + "1"
                row_cells.append("<td class='%s%s'>%s</td>" % (cls, extra, val))
            rows_data[si] = row_cells

            # visibility row: color override for low vis
            if param_key == "vis":
                row_cells_overrides = []
                if si == 0:
                    row_cells_overrides.append(
                        "<td class='ship-cell' rowspan=4>"
                        "<div class=ship-name>%s</div>"
                        "<div class=vessel-type>%s</div>"
                        "</td>" % (H(ship["name"]), H(ship["type"]))
                    )
                row_cells_overrides.append("<td class=param-label>%s</td>" % label)
                for ci, item in enumerate(col_dts):
                    d, t = item
                    key = "%s %s" % (d, t)
                    vd = ship["data"].get(param_key, {}).get(key, {})
                    val = vd.get("value", "-")
                    bg = vd.get("bg", "")
                    if val and val != "-" and try_float(val) < VIS_THRESH:
                        cls = "vis-warn"
                    else:
                        cls = "vis-ok"
                    if val in ("-", ""):
                        val = "-"
                        cls = "vis-ok"
                    row_cells_overrides.append("<td class='%s'>%s</td>" % (cls, val))
                rows_data[si] = row_cells_overrides

        for row_cells in rows_data:
            parts.append("<tr>%s</tr>\n" % "".join(row_cells))

    parts.append("</tbody>\n")
    parts.append("</table>\n")
    parts.append("</div>\n")

    # Footer
    parts.append("<div style='text-align:center;padding:12px;color:#666;font-size:11px;'>\n")
    parts.append("HiFleet 气象预报 &copy; %s | 更新: %s | 排除: ORE CHINA, ORE DONGJIAKOU, ORE HEBEI, ORE SHANDONG\n"
                 % (today, today))
    parts.append("</div>\n")
    parts.append("</body></html>")

    html = "".join(parts)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print("Saved: %s (%d bytes)" % (output_path, len(html)))


def try_float(v):
    try:
        return float(v)
    except:
        return 999


if __name__ == "__main__":
    ships, col_dts = load_data()
    print("Ships: %d, Slots: %d" % (len(ships), len(col_dts)))

    # Ensure output dir
    os.makedirs("output", exist_ok=True)
    out = os.path.join("output", "index.html")
    build_html(ships, col_dts, out)
    print("Done!")
