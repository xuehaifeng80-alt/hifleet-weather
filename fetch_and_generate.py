#!/usr/bin/env python3
"""Fetch HiFleet weather email and generate dashboard."""
import os
import json
import re
from datetime import datetime, timedelta
from collections import defaultdict
from imap_tools import MailBox, AND

# === Config ===
EMAIL_USER = os.environ.get("EMAIL_USER", "xuehaifeng666@qq.com")
EMAIL_AUTH = os.environ.get("EMAIL_AUTH", "ztfzkfwzzywmjdaf")
QQ_APP_PWD = os.environ.get("QQ_APP_PWD", "")

# Ships NOT on Manus platform (excluded)
EXCLUDE_SHIPS = {"ORE CHINA", "ORE DONGJIAKOU", "ORE HEBEI", "ORE SHANDONG"}

def get_imap_password():
    return QQ_APP_PWD or EMAIL_AUTH

def fetch_hifleet_email():
    """Fetch latest HiFleet weather email from QQ Mail IMAP."""
    password = get_imap_password()
    if not password:
        print("No IMAP password available")
        return None
    
    try:
        with MailBox("imap.qq.com").login(EMAIL_USER, password) as mailbox:
            criteria = AND(from_="hifleet.com")
            messages = list(mailbox.fetch(criteria, reverse=True, limit=1))
            if messages:
                print(f"Found {len(messages)} email(s)")
                return messages[0].html
            else:
                print("No hifleet emails found")
                return None
    except Exception as e:
        print(f"IMAP error: {e}")
        return None

def parse_hifleet_html(html):
    """Parse HiFleet email HTML to extract ship weather data."""
    ships_data = {}
    if not html:
        return ships_data
    
    # Match table rows: ship_name | forecast | type | port | ...
    ship_pattern = re.compile(
        r'<tr[^>]*>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*'
        r'<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*'
        r'<td[^>]*>(.*?)</td>',
        re.DOTALL
    )
    
    for row in ship_pattern.findall(html):
        ship_name = row[0].strip()
        if not ship_name or ship_name.upper() in EXCLUDE_SHIPS:
            continue
        
        forecast = row[1].strip()
        # Parse forecast days: "1d/3.5/1.5/5" = day/wind/wave/visibility
        days = re.findall(r'(\d+m?)/(\d+\.?\d*)/(\d+\.?\d*)/(\d+\.?\d*)', forecast)
        
        ships_data[ship_name] = {
            "type": row[2].strip(),
            "forecast": forecast,
            "days": days
        }
    
    return ships_data

def load_cached_data():
    """Load cached ship data from JSON file."""
    try:
        with open("hifleet_weather_76ships.json", encoding="utf-8") as f:
            cached = json.load(f)
        
        ships_data = {}
        for ship in cached.get("ships", []):
            name = ship.get("name", "")
            if not name or name.upper() in EXCLUDE_SHIPS:
                continue
            
            data = ship.get("data", {})
            # data keys are Chinese: 风级, 浪高, 能见度
            # Each contains timestamps as keys
            wind_data = data.get("风级", {})
            wave_data = data.get("浪高", {})
            vis_data = data.get("能见度", {})
            
            # Get all unique timestamps, sorted
            all_times = set(wind_data.keys()) | set(wave_data.keys()) | set(vis_data.keys())
            
            # Build days array from timestamps
            # Format each day: (day_offset, wind, wave, visibility)
            days = []
            for ts in sorted(all_times):
                # Parse timestamp like "2026-07-14 1800"
                try:
                    dt = datetime.strptime(ts, "%Y-%m-%d %H%M")
                    wind = wind_data.get(ts, {}).get("value", "")
                    wave = wave_data.get(ts, {}).get("value", "")
                    vis = vis_data.get(ts, {}).get("value", "")
                    if wind or wave or vis:
                        days.append((ts, wind, wave, vis))
                except:
                    pass
            
            # Simplify to daily: take midday values or average
            daily = simplify_to_daily(wind_data, wave_data, vis_data)
            
            ships_data[name] = {
                "type": ship.get("type", ""),
                "data": data,
                "daily": daily
            }
        
        print(f"Loaded {len(ships_data)} ships from cache")
        return ships_data
    except Exception as e:
        print(f"Failed to load cache: {e}")
        return {}

def simplify_to_daily(wind_data, wave_data, vis_data):
    """Simplify timestamp data to daily values (midday or average)."""
    # Group by day
    by_day = defaultdict(lambda: {"wind": [], "wave": [], "vis": []})
    
    for ts_str, vals in wind_data.items():
        try:
            dt = datetime.strptime(ts_str, "%Y-%m-%d %H%M")
            day = dt.date()
            v = vals.get("value", "")
            if v:
                by_day[day]["wind"].append(float(v))
        except:
            pass
    
    for ts_str, vals in wave_data.items():
        try:
            dt = datetime.strptime(ts_str, "%Y-%m-%d %H%M")
            day = dt.date()
            v = vals.get("value", "")
            if v:
                by_day[day]["wave"].append(float(v))
        except:
            pass
    
    for ts_str, vals in vis_data.items():
        try:
            dt = datetime.strptime(ts_str, "%Y-%m-%d %H%M")
            day = dt.date()
            v = vals.get("value", "")
            if v:
                by_day[day]["vis"].append(float(v))
        except:
            pass
    
    # Build 7-day forecast starting from today
    today = datetime.now().date()
    daily = []
    for i in range(7):
        day = today + timedelta(days=i)
        d = by_day.get(day, {"wind": [], "wave": [], "vis": []})
        wind = round(sum(d["wind"]) / len(d["wind"]), 1) if d["wind"] else "-"
        wave = round(sum(d["wave"]) / len(d["wave"]), 1) if d["wave"] else "-"
        vis = round(sum(d["vis"]) / len(d["vis"]), 1) if d["vis"] else "-"
        daily.append((str(i+1), str(wind), str(wave), str(vis)))
    
    return daily

def generate_dashboard(ships_data, output_path="output/index.html"):
    """Generate HTML dashboard."""
    os.makedirs("output", exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Get daily data (handle both email-parsed and cache format)
    def get_day1(data):
        if "days" in data:
            return data["days"][0] if data["days"] else None
        elif "daily" in data:
            return data["daily"][0] if data["daily"] else None
        return None
    
    def get_all_days(data):
        if "days" in data:
            return data["days"]
        elif "daily" in data:
            return data["daily"]
        return []
    
    # Build alert ships: wind > 6 OR wave >= 3m on day 1
    alert_ships = []
    for ship, data in ships_data.items():
        day1 = get_day1(data)
        if day1:
            wind = float(day1[1]) if day1[1] and day1[1] != "-" else 0
            wave = float(day1[2]) if day1[2] and day1[2] != "-" else 0
            if wind > 6 or wave >= 3:
                alert_ships.append({
                    "ship": ship,
                    "type": data.get("type", ""),
                    "wind": wind,
                    "wave": wave
                })
    
    # Wave >= 3m ships per day
    wave3_ships = {}
    for ship, data in ships_data.items():
        wave3_ships[ship] = []
        days = get_all_days(data)
        for i, day in enumerate(days):
            wave = float(day[2]) if day[2] and day[2] != "-" else 0
            if wave >= 3:
                wave3_ships[ship].append(i + 1)
    
    html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HiFleet Weather - {today}</title>
<style>
body {{ font-family: Arial; background: #0a0e27; color: #c8d3f5; margin: 0; padding: 20px; }}
.header {{ background: linear-gradient(135deg,#1a1f4e,#2d3585); padding:20px; border-radius:12px; margin-bottom:20px; }}
.header h1 {{ color:#7aa2f7; margin:0; }}
.header p {{ color:#7aa2f7; opacity:0.7; margin:5px 0 0; }}
.alert-box {{ background:#1e2340; border-left:4px solid #f7768e; padding:15px; border-radius:8px; margin-bottom:20px; }}
.alert-box h3 {{ color:#f7768e; margin:0 0 10px; }}
table {{ width:100%; border-collapse:collapse; background:#1e2340; border-radius:8px; overflow:hidden; margin-bottom:30px; }}
th {{ background:#7c3aed; color:#fff; padding:10px; text-align:center; }}
td {{ padding:8px 10px; text-align:center; border-bottom:1px solid #2d3585; }}
tr:hover {{ background:#2d3585; }}
.wave-red {{ color:#f7768e; font-weight:bold; }}
.wave-orange {{ color:#ff9e64; }}
.wind-red {{ color:#f7768e; font-weight:bold; }}
section {{ margin-bottom:30px; }}
.ship-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(300px,1fr)); gap:12px; }}
.ship-card {{ background:#1e2340; border-radius:8px; padding:12px; border:1px solid #2d3585; }}
.ship-card h4 {{ color:#7aa2f7; margin:0 0 8px; }}
.badge {{ display:inline-block; padding:2px 8px; border-radius:4px; font-size:12px; margin:2px; }}
.badge-red {{ background:#3d2048; color:#f7768e; }}
.no-data {{ color:#7aa2f7; text-align:center; padding:30px; font-size:16px; }}
</style>
</head>
<body>
<div class="header">
  <h1>HiFleet Weather Dashboard</h1>
  <p>{len(ships_data)} ships &middot; Updated: {today} &middot; Alert: wind&gt;6 OR wave&ge;3m</p>
</div>

<div class="alert-box">
  <h3>Alerts ({len(alert_ships)} ships)</h3>
"""
    if alert_ships:
        html += """  <table>
    <tr><th>Ship</th><th>Type</th><th>Wind (m/s)</th><th>Wave (m)</th></tr>
"""
        for s in sorted(alert_ships, key=lambda x: -x["wave"]):
            wind_cls = "wind-red" if s["wind"] > 6 else ""
            wave_cls = "wave-red" if s["wave"] >= 3 else "wave-orange"
            html += f'    <tr><td>{s["ship"]}</td><td>{s["type"]}</td><td class="{wind_cls}">{s["wind"]}</td><td class="{wave_cls}">{s["wave"]}</td></tr>\n'
        html += "  </table>\n"
    else:
        if ships_data:
            html += "  <p>No alerts today.</p>\n"
        else:
            html += "  <p>No data available.</p>\n"
    
    html += """</div>

<section>
  <h2 style="color:#7aa2f7;">Ships with Wave &ge; 3m (Next 7 Days)</h2>
  <div class="ship-grid">
"""
    
    wave3_count = 0
    for ship, days in sorted(wave3_ships.items()):
        if days:
            wave3_count += 1
            badges = " ".join([
                f'<span class="badge badge-red">D{d}</span>' for d in days
            ])
            html += f'    <div class="ship-card"><h4>{ship}</h4>{badges}</div>\n'
    
    if wave3_count == 0:
        html += '    <p class="no-data">No ships with wave &ge; 3m.</p>\n'
    
    html += """  </div>
</section>

<section>
  <h2 style="color:#7aa2f7;">Full 7-Day Forecast</h2>
  <div style="overflow-x:auto;">
  <table>
    <tr>
      <th>Ship</th><th>Type</th>
      <th>D1 Wind</th><th>D1 Wave</th>
      <th>D2 Wind</th><th>D2 Wave</th>
      <th>D3 Wind</th><th>D3 Wave</th>
      <th>D4 Wind</th><th>D4 Wave</th>
      <th>D5 Wind</th><th>D5 Wave</th>
      <th>D6 Wind</th><th>D6 Wave</th>
      <th>D7 Wind</th><th>D7 Wave</th>
    </tr>
"""
    
    for ship, data in sorted(ships_data.items()):
        days = get_all_days(data)
        row = f'    <tr><td>{ship}</td><td>{data.get("type","")}</td>'
        for i in range(7):
            if i < len(days):
                wind = days[i][1] if len(days[i]) > 1 else "-"
                wave = days[i][2] if len(days[i]) > 2 else "-"
                wind_cls = "wind-red" if wind != "-" and float(wind) > 6 else ""
                wave_cls = "wave-red" if wave != "-" and float(wave) >= 3 else ("wave-orange" if wave != "-" and float(wave) >= 2 else "")
                row += f'<td class="{wind_cls}">{wind}</td><td class="{wave_cls}">{wave}</td>'
            else:
                row += "<td>-</td><td>-</td>"
        row += "</tr>\n"
        html += row
    
    html += """  </table>
  </div>
</section>
</body>
</html>"""
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard saved: {output_path}")

def main():
    print("Fetching HiFleet email...")
    html = fetch_hifleet_email()
    
    if html:
        ships_data = parse_hifleet_html(html)
        print(f"Parsed {len(ships_data)} ships from email")
    else:
        print("No email - loading cached data...")
        ships_data = load_cached_data()
    
    if not ships_data:
        print("WARNING: No ship data. Creating placeholder.")
        with open("output/index.html", "w", encoding="utf-8") as f:
            f.write("<html><body style='background:#0a0e27;color:#7aa2f7;font-family:Arial;padding:40px;'>"
                    "<h1>No data available</h1><p>Please check email configuration.</p></body></html>")
    else:
        generate_dashboard(ships_data)
    
    print("Done!")

if __name__ == "__main__":
    main()
