#!/usr/bin/env python3
"""
update_dashboard.py  ·  1010 Church · Container Dashboard Updater
──────────────────────────────────────────────────────────────────
Double-click to run  (or:  python update_dashboard.py)

1. Reads  10C_Logistcs Output.xlsm  (must be in the same folder)
2. Extracts container schedule, per-unit statuses, ROJ floor dates
3. Regenerates  10C_Container_Dashboard.html  with the latest data
4. Runs  git add → commit → push  so the GitHub Pages URL auto-updates

One-time setup:  see the README or the instructions printed at the bottom.
"""

import os, sys, json, subprocess
from datetime import datetime, timedelta

# ── AUTO-INSTALL openpyxl IF MISSING ─────────────────────────
try:
    import openpyxl
    from openpyxl.utils.datetime import from_excel
except ImportError:
    print("openpyxl not found — installing now...")
    subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], check=True)
    import openpyxl
    from openpyxl.utils.datetime import from_excel

# ── CONFIG ────────────────────────────────────────────────────
HERE        = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE  = r"C:\Users\Ewan\OneDrive - Studio Snaidero Chicago\Shared\1010_Logistics Site\10C_Logistcs Output.xlsm"
OUTPUT_HTML = os.path.join(HERE, "10C_Container_Dashboard.html")
AUTO_GIT    = True   # change to False to skip the git push step

# ── SHEET NAMES (tries each variant in order, uses first match) ──
SH_CTN_VARIANTS = ["10C_CTNR Schedule (spill)", "1010C_Container Table"]
SH_BLD_VARIANTS = ["10C Building Matrix (status)", "10C_BUILDING MATRIX"]
SH_ROJ_VARIANTS = ["1010 Church ROJ Dates"]  # optional — skipped if not found

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def to_dt(v):
    """Return a Python datetime from an openpyxl cell value (datetime or Excel serial)."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)) and v > 40000:
        try:
            return from_excel(int(v))
        except Exception:
            return None
    return None

def fmt_date(v, year=False):
    """Format a cell value as 'Aug 12' or 'Aug 12, 2026'.  Returns '—' if unparseable."""
    d = to_dt(v)
    if not d:
        return "—"
    s = f"{d.strftime('%b')} {d.day}"
    return s + f", {d.year}" if year else s

def clean(v):
    """Return a string from a cell value, or '—' for None / empty."""
    if v is None:
        return "—"
    s = str(v).strip()
    return s if s else "—"

def hdr_index(row, *names):
    """Find the first column index in row whose value is in names (case-insensitive)."""
    targets = {n.lower() for n in names}
    for i, v in enumerate(row):
        if v and str(v).strip().lower() in targets:
            return i
    return None

def js_str(v):
    """Escape a Python value for embedding in a JS string literal."""
    if isinstance(v, str):
        return "'" + v.replace("'", "\\'") + "'"
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return "'—'"
    return repr(v)

# ─────────────────────────────────────────────────────────────
# LOAD WORKBOOK
# ─────────────────────────────────────────────────────────────

print(f"\n{'─'*56}")
print(f"  1010 Church · Dashboard Updater")
print(f"{'─'*56}")

if not os.path.exists(EXCEL_FILE):
    sys.exit(
        f"\nERROR: Excel file not found:\n  {EXCEL_FILE}\n"
        f"Make sure  10C_Logistcs Output.xlsm  is in the same\n"
        f"folder as this script, then double-click again.\n"
    )

print(f"\nReading  {os.path.basename(EXCEL_FILE)} …")
wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

def find_sheet(variants, required=True):
    for name in variants:
        if name in wb.sheetnames:
            return name
    if required:
        sys.exit(f"\nERROR: None of these sheets found: {variants}\n"
                 f"Available sheets: {wb.sheetnames}")
    return None

SH_CTN = find_sheet(SH_CTN_VARIANTS, required=True)
SH_BLD = find_sheet(SH_BLD_VARIANTS, required=True)
SH_ROJ = find_sheet(SH_ROJ_VARIANTS, required=False)
print(f"  Using sheets: '{SH_CTN}', '{SH_BLD}', ROJ: '{SH_ROJ or 'not found — will use defaults'}')")

# ─────────────────────────────────────────────────────────────
# 1.  CONTAINER TABLE
# ─────────────────────────────────────────────────────────────

ws = wb[SH_CTN]
all_rows = list(ws.iter_rows(values_only=True))

# Find the header row (contains "CONT. STATUS" or similar)
hdr_row, hdr = None, None
for i, row in enumerate(all_rows):
    if any(v and "STATUS" in str(v).upper() and "CONT" in str(v).upper() for v in row):
        hdr_row = i
        hdr = [str(v).strip() if v else "" for v in row]
        break

if hdr is None:
    sys.exit(f"ERROR: Could not find header row in '{SH_CTN}'.")

# Column finders
def C(*names):
    for n in names:
        try:
            return hdr.index(n)
        except ValueError:
            pass
    # Fuzzy match
    nl = n.lower()
    for i, h in enumerate(hdr):
        if nl in h.lower():
            return i
    return None

i_status  = C("CONT. STATUS", "CONT.STATUS", "STATUS")
i_num     = C("CONT. #", "CONT.#", "CTN #")
i_week    = C("WEEK")
i_floor   = C("FLOORS", "FLOOR")
i_qty     = C("UNIT Q.ty", "UNIT Q.TY", "UNIT QTY")
i_ordered = C("ORDERED")
i_vessel  = C("VESSEL")
i_load    = C("LOAD DATE")
i_ship    = C("SHIP DATE")
i_chs     = C("PORT ARRIVAL DATE/CHS", "PORT ARRIVAL DATE", "ARRIVAL DATE/CHS", "ETA CHS", "CHS")
i_rail    = C("RAIL DATE")
i_nash    = C("ARRIVAL DATE (Nashville)", "NASHVILLE DATE", "ETA NASH")
i_del     = C("DELIVERY DATE", "DELIVERED DATE", "DELIVERY")

print(f"  Container table header on row {hdr_row+1}")

CONTAINERS = []
for row in all_rows[hdr_row + 1:]:
    if not row or all(v is None for v in row):
        continue
    num_raw = row[i_num] if i_num is not None else None
    if not isinstance(num_raw, (int, float)):
        continue
    num = int(num_raw)
    if num < 1 or num > 30:
        continue

    week_raw = row[i_week] if i_week is not None else None
    if isinstance(week_raw, (int, float)):
        week = f"Wk {int(week_raw)}"
    else:
        ws_str = clean(week_raw)
        week = ("Wk " + ws_str.replace("Week ", "")) if "Week" in ws_str else ws_str

    vessel_raw = row[i_vessel] if i_vessel is not None else None
    vessel = clean(vessel_raw)
    if vessel in ("—", "--", "None"):
        vessel = "—"

    CONTAINERS.append({
        "num":      num,
        "status":   clean(row[i_status]) if i_status is not None else "PROJ",
        "week":     week,
        "floor":    int(row[i_floor]) if (i_floor is not None and isinstance(row[i_floor], (int, float))) else 0,
        "unitQty":  int(row[i_qty])   if (i_qty    is not None and isinstance(row[i_qty],   (int, float))) else 0,
        "ordered":  bool(row[i_ordered]) if i_ordered is not None else False,
        "vessel":   vessel,
        "loadDate": fmt_date(row[i_load]) if i_load is not None else "—",
        "shipDate": fmt_date(row[i_ship]) if i_ship is not None else "—",
        "etaChs":   fmt_date(row[i_chs])  if i_chs  is not None else "—",
        "railDate": fmt_date(row[i_rail]) if i_rail is not None else "—",
        "etaNash":  fmt_date(row[i_nash]) if i_nash is not None else "—",
        "delivery": fmt_date(row[i_del],  year=True) if i_del is not None else "—",
    })

CONTAINERS.sort(key=lambda c: c["num"])
print(f"  → {len(CONTAINERS)} regular containers parsed")

# Floor → main container number + status
floor_to_ctn    = {c["floor"]: c["num"]    for c in CONTAINERS}
floor_to_status = {c["floor"]: c["status"] for c in CONTAINERS}

# ─────────────────────────────────────────────────────────────
# 2.  BUILDING MATRIX  →  EXCEPTIONS
# ─────────────────────────────────────────────────────────────

ws2 = wb[SH_BLD]
bld_rows = list(ws2.iter_rows(values_only=True))

bld_hdr_row, bld_hdr = None, None
for i, row in enumerate(bld_rows):
    strs = [str(v).strip() if v else "" for v in row]
    if "UNIT #" in strs or "UNIT#" in strs:
        bld_hdr_row, bld_hdr = i, strs
        break

EXCEPTIONS = {}
if bld_hdr is not None:
    def BC(*names):
        for n in names:
            if n in bld_hdr:
                return bld_hdr.index(n)
        return None

    bi_unit  = BC("UNIT #", "UNIT#")
    bi_floor = BC("FLOOR")
    bi_tier  = BC("TIER")
    bi_kc    = BC("K CTNR #")
    bi_v1c   = BC("V1 CTNR #")
    bi_v2c   = BC("V2 CTNR #")
    bi_ks    = BC(" K STATUS", "K STATUS")
    bi_v1s   = BC("V1 STATUS")
    bi_v2s   = BC("V2 STATUS")

    for row in bld_rows[bld_hdr_row + 1:]:
        if not row or bi_unit is None:
            continue
        unit_raw = row[bi_unit]
        if not isinstance(unit_raw, (int, float)):
            continue
        unit_id = int(unit_raw)

        tier = str(row[bi_tier] or "").strip() if bi_tier is not None else "APT"
        if tier not in ("APT", "CONDO"):
            continue

        floor_raw = row[bi_floor] if bi_floor is not None else None
        floor = int(floor_raw) if isinstance(floor_raw, (int, float)) else 0
        main_ctn = floor_to_ctn.get(floor)
        if main_ctn is None:
            continue

        def _ctnr(col):
            if col is None: return None
            v = row[col]
            if v is None or str(v).strip() in ("", "--", "—"): return None
            return int(v) if isinstance(v, (int, float)) else str(v).strip()

        def _stat(col):
            if col is None: return "PROJ"
            v = row[col]
            s = str(v).strip() if v else "PROJ"
            return s if s and s not in ("--", "—") else "PROJ"

        kc  = _ctnr(bi_kc)
        v1c = _ctnr(bi_v1c)
        v2c = _ctnr(bi_v2c)
        ks  = _stat(bi_ks)
        v1s = _stat(bi_v1s)
        v2s = _stat(bi_v2s)

        exc = {}
        if kc is not None and kc != main_ctn:
            exc["kCtnr"]  = str(kc)
            exc["kStatus"] = ks
        if v1c is not None and v1c != main_ctn:
            exc["v1Ctnr"]  = str(v1c)
            exc["v1Status"] = v1s
        if v2c is not None and v2c != main_ctn:
            exc["v2Ctnr"]  = str(v2c)
            exc["v2Status"] = v2s
        if exc:
            EXCEPTIONS[unit_id] = exc
    print(f"  → {len(EXCEPTIONS)} exception units detected")
else:
    print("  ! Building matrix header not found — exceptions left empty")

# ─────────────────────────────────────────────────────────────
# 3.  ROJ DATES
# ─────────────────────────────────────────────────────────────

ROJ = {}
if SH_ROJ is None:
    print("  → ROJ sheet not present — floor-ready dates left empty")
else:
    for row in wb[SH_ROJ].iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 6:
            continue
        floor_raw = row[2]   # col C = Floor #
        if not isinstance(floor_raw, (int, float)):
            continue
        floor = int(floor_raw)
        if floor < 10 or floor > 50:
            continue

        window_raw  = row[5]   # col F = Delivery Window
        install_raw = row[7]   # col H = Install Start

        window_dt  = to_dt(window_raw)
        install_dt = to_dt(install_raw)

        window  = fmt_date(window_raw)
        # Reject install dates that precede the delivery window (formula error in source sheet)
        if install_dt and window_dt and install_dt < window_dt:
            install = "—"
        else:
            install = fmt_date(install_raw)

        if window != "—":
            ROJ[floor] = {"window": window, "install": install}
    print(f"  → ROJ dates parsed for {len(ROJ)} floors")

# ─────────────────────────────────────────────────────────────
# 4.  COMPUTE DASHBOARD METRICS
# ─────────────────────────────────────────────────────────────

n_del   = sum(1 for c in CONTAINERS if c["status"] == "D")
n_trans = sum(1 for c in CONTAINERS if c["status"] in ("PLD", "CHS", "ENR"))
n_load  = sum(1 for c in CONTAINERS if c["status"] in ("LDD", "LDG"))

del_ctns  = [c for c in CONTAINERS if c["status"] == "D"]
del_sub   = " · ".join(f"CTN {c['num']}" for c in del_ctns[-3:]) or "None yet"
trans_ctns = [c for c in CONTAINERS if c["status"] in ("PLD","CHS","ENR")]
trans_sub  = " · ".join(c["status"] for c in trans_ctns[:4]) or "None"
load_ctns  = [c for c in CONTAINERS if c["status"] in ("LDD","LDG")]
load_sub   = " · ".join(f"CTN {c['num']} ({c['status']})" for c in load_ctns[:3]) or "None"

first_ctn = CONTAINERS[0] if CONTAINERS else None
first_del = first_ctn["delivery"] if first_ctn else "—"

# Remove year from first_del if present (just "Sep 11" not "Sep 11, 2026")
if first_del and "," in first_del:
    first_del_short = first_del.split(",")[0]
else:
    first_del_short = first_del

generated_on = datetime.now().strftime("%b %-d, %Y at %-I:%M %p")

# ─────────────────────────────────────────────────────────────
# 5.  GENERATE JS DATA STRINGS
# ─────────────────────────────────────────────────────────────

def js_containers():
    lines = []
    for c in CONTAINERS:
        lines.append(
            f"  {{ num:{c['num']:2d}, status:{js_str(c['status'])}, week:{js_str(c['week'])}, "
            f"floor:{c['floor']}, unitQty:{c['unitQty']}, "
            f"loadDate:{js_str(c['loadDate'])}, shipDate:{js_str(c['shipDate'])}, "
            f"vessel:{js_str(c['vessel'])}, etaChs:{js_str(c['etaChs'])}, "
            f"railDate:{js_str(c['railDate'])}, etaNash:{js_str(c['etaNash'])}, "
            f"delivery:{js_str(c['delivery'])}, ordered:{'true' if c['ordered'] else 'false'} }}"
        )
    return "[\n" + ",\n".join(lines) + "\n]"

def js_roj():
    lines = []
    for floor in sorted(ROJ.keys()):
        r = ROJ[floor]
        lines.append(f"  {floor}: {{ window:{js_str(r['window'])}, install:{js_str(r['install'])} }}")
    return "{\n" + ",\n".join(lines) + "\n}"

def js_exceptions():
    if not EXCEPTIONS:
        return "{}"
    lines = []
    for uid in sorted(EXCEPTIONS.keys()):
        exc = EXCEPTIONS[uid]
        parts = []
        if "kCtnr"   in exc: parts.append(f"kCtnr:{js_str(exc['kCtnr'])}")
        if "kStatus" in exc: parts.append(f"kStatus:{js_str(exc['kStatus'])}")
        if "v1Ctnr"  in exc: parts.append(f"v1Ctnr:{js_str(exc['v1Ctnr'])}")
        if "v1Status" in exc: parts.append(f"v1Status:{js_str(exc['v1Status'])}")
        if "v2Ctnr"  in exc: parts.append(f"v2Ctnr:{js_str(exc['v2Ctnr'])}")
        if "v2Status" in exc: parts.append(f"v2Status:{js_str(exc['v2Status'])}")
        lines.append(f"  {uid}: {{ {', '.join(parts)} }}")
    return "{\n" + ",\n".join(lines) + "\n}"

# ─────────────────────────────────────────────────────────────
# 6.  HTML TEMPLATE
# ─────────────────────────────────────────────────────────────

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>1010 Church · Container Dashboard</title>
<style>
:root {{
  --bg: #F7F6F3;
  --surface: #FFFFFF;
  --border: rgba(0,0,0,0.10);
  --border-strong: rgba(0,0,0,0.18);
  --text: #1A1A18;
  --text-2: #5C5C58;
  --text-3: #9A9A94;
  --accent: #1A3A2A;
  --proj-bg: #FEF5E7;  --proj-text: #7A4A00;
  --ldd-bg: #E8F2EB;   --ldd-text: #1A5C2A;
  --del-bg: #E3F0EE;   --del-text: #0D5040;
  --route-bg: #E6EEF8; --route-text: #1A3A70;
  --port-bg: #EDEBFC;  --port-text: #3A2A8A;
  --rail-bg: #FAF0E8;  --rail-text: #703010;
  --load-bg: #F0EFE8;  --load-text: #5A5A56;
  --open-bg: #FEF5E7;  --open-text: #7A4A00;
  --crit-bg: #FEECEC;  --crit-text: #7A1A1A;
  --closed-bg: #E8F2EB;--closed-text: #1A5C2A;
  --scope-bg: #F4F3EF;
  --radius: 8px;
  --radius-lg: 12px;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ background: var(--bg); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; color: var(--text); font-size: 13px; line-height: 1.5; }}
.shell {{ max-width: 1240px; margin: 0 auto; padding: 24px 20px 60px; }}
.topbar {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid var(--border); }}
.topbar-left h1 {{ font-size: 18px; font-weight: 600; letter-spacing: -.3px; }}
.topbar-left .sub {{ font-size: 11px; color: var(--text-2); margin-top: 3px; }}
.date-pill {{ font-size: 11px; background: var(--accent); color: #fff; padding: 4px 10px; border-radius: 99px; }}
.metrics {{ display: grid; grid-template-columns: repeat(6,1fr); gap: 8px; margin-bottom: 20px; }}
.mc {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 10px 12px; }}
.mc .lbl {{ font-size: 10px; color: var(--text-3); text-transform: uppercase; letter-spacing: .05em; margin-bottom: 4px; }}
.mc .val {{ font-size: 22px; font-weight: 600; line-height: 1; }}
.mc .vsub {{ font-size: 10px; color: var(--text-3); margin-top: 3px; }}
.tabs {{ display: flex; gap: 2px; margin-bottom: 16px; background: var(--surface); border: 1px solid var(--border); padding: 3px; border-radius: var(--radius); width: fit-content; }}
.tab {{ padding: 6px 14px; font-size: 12px; font-weight: 500; border: none; background: none; border-radius: 6px; cursor: pointer; color: var(--text-2); font-family: inherit; transition: background .12s, color .12s; }}
.tab:hover {{ background: var(--bg); color: var(--text); }}
.tab.on {{ background: var(--accent); color: #fff; }}
.panel {{ display: none; }}
.panel.on {{ display: block; }}
.bdg {{ display: inline-flex; align-items: center; font-size: 10px; font-weight: 600; padding: 2px 7px; border-radius: 4px; white-space: nowrap; }}
.b-D    {{ background: var(--del-bg);   color: var(--del-text);   }}
.b-PLD  {{ background: var(--del-bg);   color: var(--del-text);   }}
.b-CHS  {{ background: var(--port-bg);  color: var(--port-text);  }}
.b-ENR  {{ background: var(--route-bg); color: var(--route-text); }}
.b-LDD  {{ background: var(--ldd-bg);   color: var(--ldd-text);   }}
.b-LDG  {{ background: var(--load-bg);  color: var(--load-text);  }}
.b-PROJ {{ background: var(--proj-bg);  color: var(--proj-text);  }}
.b-open  {{ background: var(--open-bg); color: var(--open-text);  }}
.b-closed{{ background: var(--closed-bg);color: var(--closed-text);}}
.b-crit  {{ background: var(--crit-bg); color: var(--crit-text);  }}
.dot {{ width: 8px; height: 8px; border-radius: 50%; display: inline-block; flex-shrink: 0; }}
.d-D    {{ background: #0D7A60; }}
.d-PLD  {{ background: #0D7A60; }}
.d-CHS  {{ background: #5A4AAA; }}
.d-ENR  {{ background: #2A5AAA; }}
.d-LDD  {{ background: #2A7A3A; }}
.d-LDG  {{ background: #888880; }}
.d-PROJ {{ background: #D4890A; }}
.d-BH   {{ background: #BBBBBB; }}
.mdot {{ width: 6px; height: 6px; border-radius: 50%; display: inline-block; flex-shrink: 0; }}
.tw {{ border: 1px solid var(--border); border-radius: var(--radius-lg); overflow: hidden; overflow-x: auto; background: var(--surface); margin-bottom: 12px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
th {{ text-align: left; padding: 8px 10px; color: var(--text-3); font-weight: 600; font-size: 10px; text-transform: uppercase; letter-spacing: .04em; border-bottom: 1px solid var(--border); background: var(--bg); white-space: nowrap; }}
td {{ padding: 7px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
tr:last-child td {{ border-bottom: none; }}
tbody tr:hover td {{ background: var(--bg); }}
.mono {{ font-family: 'SF Mono', 'Fira Code', monospace; font-size: 10px; }}
.proj-ok td {{ background: rgba(42,122,58,.03); }}
.proj-warn td {{ background: rgba(212,137,10,.04); }}
.route-pipeline {{ display: flex; align-items: center; background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius-lg); padding: 14px 16px; margin-bottom: 16px; overflow-x: auto; }}
.rp-step {{ display: flex; flex-direction: column; align-items: center; flex: 1; min-width: 64px; }}
.rp-icon {{ width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 13px; margin-bottom: 4px; }}
.rp-name {{ font-size: 9px; font-weight: 600; text-transform: uppercase; letter-spacing: .05em; color: var(--text-3); text-align: center; }}
.rp-val  {{ font-size: 10px; color: var(--text-2); text-align: center; margin-top: 1px; }}
.rp-line {{ flex: 1; height: 2px; background: var(--border); align-self: center; margin-bottom: 20px; min-width: 10px; }}
.leg {{ display: flex; gap: 14px; flex-wrap: wrap; font-size: 10px; color: var(--text-2); margin-bottom: 10px; align-items: center; }}
.leg-i {{ display: flex; align-items: center; gap: 5px; }}
.floor-block {{ margin-bottom: 10px; }}
.floor-hdr {{ font-size: 10px; font-weight: 700; color: var(--text-2); text-transform: uppercase; letter-spacing: .06em; padding: 5px 0; margin-bottom: 5px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }}
.floor-num {{ background: var(--accent); color: #fff; border-radius: 4px; padding: 1px 6px; font-size: 9px; }}
.floor-meta {{ font-size: 9px; color: var(--text-3); font-weight: 400; }}
.unit-grid {{ display: grid; grid-template-columns: repeat(12,1fr); gap: 4px; }}
.ucard {{ border: 1px solid var(--border); border-radius: 6px; padding: 4px 5px; background: var(--surface); min-width: 0; }}
.ub-D   {{ border-color: rgba(13,122,96,.35);  }}
.ub-PLD {{ border-color: rgba(13,122,96,.35);  }}
.ub-CHS {{ border-color: rgba(90,74,170,.35);  }}
.ub-ENR {{ border-color: rgba(42,90,170,.35);  }}
.ub-LDD {{ border-color: rgba(42,122,58,.5); border-width: 1.5px; }}
.ub-LDG {{ border-color: rgba(136,136,128,.5); border-width: 1.5px; }}
.ub-PROJ{{ border-color: var(--border); }}
.ub-BH  {{ background: var(--scope-bg); border-color: var(--border); opacity: .65; }}
.uid {{ font-size: 8px; color: var(--text-3); font-weight: 700; margin-bottom: 2px; }}
.ucomp {{ display: flex; align-items: center; gap: 3px; font-size: 7.5px; padding: 1px 0; }}
.ucomp-lbl {{ color: var(--text-3); font-weight: 700; min-width: 13px; }}
.ucomp-val {{ font-weight: 600; }}
.ucomp-note {{ font-size: 7px; color: var(--text-3); margin-left: 1px; }}
.form-card {{ background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius-lg); padding: 20px; margin-top: 8px; }}
.form-card h3 {{ font-size: 13px; font-weight: 600; margin-bottom: 4px; }}
.form-card .form-sub {{ font-size: 11px; color: var(--text-2); margin-bottom: 16px; }}
.form-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 12px; }}
.form-grid.single {{ grid-template-columns: 1fr; }}
.field {{ display: flex; flex-direction: column; gap: 4px; }}
.field label {{ font-size: 10px; font-weight: 600; text-transform: uppercase; letter-spacing: .05em; color: var(--text-2); }}
.field input,.field select,.field textarea {{ border: 1px solid var(--border-strong); border-radius: var(--radius); padding: 8px 10px; font-size: 12px; font-family: inherit; color: var(--text); background: var(--bg); outline: none; transition: border-color .12s; }}
.field input:focus,.field select:focus,.field textarea:focus {{ border-color: var(--accent); background: #fff; }}
.field textarea {{ resize: vertical; min-height: 70px; }}
.field .date-fixed {{ font-size: 12px; color: var(--text-2); padding: 8px 10px; background: var(--scope-bg); border: 1px solid var(--border); border-radius: var(--radius); font-family: 'SF Mono', monospace; }}
.form-actions {{ display: flex; align-items: center; gap: 10px; margin-top: 14px; }}
.btn-submit {{ background: var(--accent); color: #fff; border: none; border-radius: var(--radius); padding: 9px 18px; font-size: 12px; font-weight: 600; cursor: pointer; font-family: inherit; }}
.btn-submit:hover {{ opacity: .88; }}
.form-msg {{ font-size: 11px; color: var(--ldd-text); font-weight: 500; display: none; }}
.form-msg.err {{ color: var(--crit-text); }}
.export-bar {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }}
.export-bar .count {{ font-size: 11px; color: var(--text-2); }}
.btn-export {{ background: var(--surface); border: 1px solid var(--border-strong); border-radius: var(--radius); padding: 6px 14px; font-size: 11px; font-weight: 600; cursor: pointer; font-family: inherit; color: var(--text); display: flex; align-items: center; gap: 6px; }}
.btn-export:hover {{ background: var(--bg); }}
.btn-export svg {{ width: 14px; height: 14px; stroke: currentColor; fill: none; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }}
.dmg-group {{ border-radius: var(--radius-lg); overflow: hidden; border: 1px solid var(--border); background: var(--surface); }}
.dmg-group.open  {{ border-left: 3px solid #CC3030; }}
.dmg-group.closed{{ border-left: 3px solid #2A7A3A; }}
.dmg-group-hdr {{ padding: 6px 12px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; background: var(--bg); border-bottom: 1px solid var(--border); color: var(--text-2); display: flex; align-items: center; gap: 8px; }}
.empty-state {{ text-align: center; padding: 32px; color: var(--text-3); font-size: 12px; }}
.empty-state .big {{ font-size: 28px; margin-bottom: 8px; }}
.note-bar {{ font-size: 10px; color: var(--text-3); background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 8px 12px; margin-bottom: 10px; }}
</style>
</head>
<body>
<div class="shell">

<div class="topbar">
  <div class="topbar-left">
    <h1>1010 Church · Container Dashboard</h1>
    <div class="sub">Route: Italy (Factory) → Sea Freight → Charleston Port → Rail → Nashville → Delivery &nbsp;·&nbsp; 33 Containers · Floors 10 – 39 &nbsp;·&nbsp; <em>Updated {generated_on}</em></div>
  </div>
  <div class="topbar-right">
    <span class="date-pill" id="today-pill"></span>
  </div>
</div>

<div class="metrics">
  <div class="mc"><div class="lbl">Total containers</div><div class="val">33</div><div class="vsub">30 floors + BH + ECT + WCT</div></div>
  <div class="mc"><div class="lbl">Delivered</div><div class="val" style="color:#0D7A60" id="m-del">{n_del}</div><div class="vsub" id="m-del-sub">{del_sub}</div></div>
  <div class="mc"><div class="lbl">Pulled / In transit</div><div class="val" style="color:#0D6A50" id="m-trans">{n_trans}</div><div class="vsub" id="m-trans-sub">{trans_sub}</div></div>
  <div class="mc"><div class="lbl">Loaded / Loading</div><div class="val" style="color:#2A7A3A" id="m-load">{n_load}</div><div class="vsub" id="m-load-sub">{load_sub}</div></div>
  <div class="mc"><div class="lbl">First delivery</div><div class="val" style="font-size:15px;padding-top:4px;color:#1A3A70">{first_del_short}</div><div class="vsub">CTN 1 · Floor 10</div></div>
  <div class="mc"><div class="lbl">Damage items</div><div class="val" id="dmg-count-metric" style="color:#7A1A1A">0</div><div class="vsub" id="dmg-sub-metric">None logged yet</div></div>
</div>

<div class="tabs">
  <button class="tab on" onclick="sw('ctn',this)">Container schedule</button>
  <button class="tab" onclick="sw('proj',this)">Projections</button>
  <button class="tab" onclick="sw('bldg',this)">Building visual</button>
  <button class="tab" onclick="sw('dmg',this)">Damage / open items</button>
</div>

<!-- ===== CONTAINER SCHEDULE ===== -->
<div id="ctn" class="panel on">
  <div class="route-pipeline">
    <div class="rp-step"><div class="rp-icon" style="background:#E8F2EB">&#127981;</div><div class="rp-name">Factory IT</div><div class="rp-val">Load &amp; seal</div></div>
    <div class="rp-line"></div>
    <div class="rp-step"><div class="rp-icon" style="background:#E6EEF8">&#128674;</div><div class="rp-name">Sea Freight</div><div class="rp-val">Italy &#8594; CHS</div></div>
    <div class="rp-line"></div>
    <div class="rp-step"><div class="rp-icon" style="background:#EDEBFC">&#9875;</div><div class="rp-name">Charleston</div><div class="rp-val">Port arrival</div></div>
    <div class="rp-line"></div>
    <div class="rp-step"><div class="rp-icon" style="background:#FAF0E8">&#128642;</div><div class="rp-name">Rail</div><div class="rp-val">CHS &#8594; Nashville</div></div>
    <div class="rp-line"></div>
    <div class="rp-step"><div class="rp-icon" style="background:#FAF0E8">&#127959;</div><div class="rp-name">Nashville RR</div><div class="rp-val">Yard arrival</div></div>
    <div class="rp-line"></div>
    <div class="rp-step"><div class="rp-icon" style="background:#E8F2EB">&#128230;</div><div class="rp-name">Delivery</div><div class="rp-val">Jobsite</div></div>
  </div>
  <div class="leg">
    <span class="leg-i"><span class="dot d-D"></span> Delivered (D)</span>
    <span class="leg-i"><span class="dot d-PLD"></span> Pulled (PLD)</span>
    <span class="leg-i"><span class="dot d-CHS"></span> Charleston (CHS)</span>
    <span class="leg-i"><span class="dot d-ENR"></span> En route (ENR)</span>
    <span class="leg-i"><span class="dot d-LDD"></span> Loaded (LDD)</span>
    <span class="leg-i"><span class="dot d-LDG"></span> Loading (LDG)</span>
    <span class="leg-i"><span class="dot d-PROJ"></span> Projected (PROJ)</span>
  </div>
  <div class="tw">
    <table>
      <thead><tr>
        <th>#</th><th>Status</th><th>Load Wk</th><th>Floor</th><th>Units</th>
        <th>Load Date</th><th>Ship Date</th><th>Vessel</th>
        <th>ETA Charleston</th><th>Rail Date</th><th>ETA Nashville</th><th>Delivery</th>
      </tr></thead>
      <tbody id="ctn-tbody"></tbody>
    </table>
  </div>
  <div class="note-bar">
    <strong>Special containers:</strong> &nbsp;
    CTN <strong>BH</strong> (Buck Hoist) &#8212; all unit-07 kitchens, 30 floors, Wk 26 &nbsp;|&nbsp;
    CTN <strong>ECT</strong> (East Crane Tie-in) &#8212; V1 for units 1004/2004/2904/3604, Wk 27 &nbsp;|&nbsp;
    CTN <strong>WCT</strong> (West Crane Tie-in) &#8212; K for 1611/2511/3311 &amp; V1 for 1612/2512/3312, Wk 27
  </div>
</div>

<!-- ===== PROJECTIONS ===== -->
<div id="proj" class="panel">
  <div class="note-bar">
    Container delivery dates vs. jobsite floor-ready windows (ROJ). &nbsp;
    <strong style="color:var(--ldd-text)">Green rows</strong> = ordered/confirmed. &nbsp;
    <strong style="color:var(--proj-text)">Amber rows</strong> = projected, not yet ordered. &nbsp;
    ROJ dates sourced from 1010 Church ROJ sheet.
  </div>
  <div class="tw">
    <table>
      <thead><tr>
        <th>CTN #</th><th>Status</th><th>Floor</th><th>Load Wk</th>
        <th>Load Date</th><th>Ship Date</th><th>ETA CHS</th><th>Rail Date</th><th>ETA Nashville</th><th>CTN Delivery</th>
        <th>ROJ Window</th><th>Install Start</th><th>Ordered?</th>
      </tr></thead>
      <tbody id="proj-tbody"></tbody>
    </table>
  </div>
  <div style="font-size:10px;color:var(--text-3);margin-top:4px">
    ROJ Window = jobsite floor-ready date. Install Start = scheduled install begin. Some install dates pending GC confirmation (marked &#8212;).
  </div>
</div>

<!-- ===== BUILDING VISUAL ===== -->
<div id="bldg" class="panel">
  <div class="leg" style="margin-bottom:10px">
    <span class="leg-i"><span class="dot d-D"></span> Delivered</span>
    <span class="leg-i"><span class="dot d-PLD"></span> Pulled</span>
    <span class="leg-i"><span class="dot d-CHS"></span> Charleston</span>
    <span class="leg-i"><span class="dot d-ENR"></span> En route</span>
    <span class="leg-i"><span class="dot d-LDD"></span> Loaded</span>
    <span class="leg-i"><span class="dot d-LDG"></span> Loading</span>
    <span class="leg-i"><span class="dot d-PROJ"></span> Projected</span>
    <span class="leg-i"><span class="dot d-BH"></span> BH/ECT/WCT (special)</span>
  </div>
  <div class="note-bar">
    Each card shows per-component status: <strong>K</strong> = kitchen, <strong>V1</strong> = vanity 1, <strong>V2</strong> = vanity 2 (units 04, 06, 10, 12 only).
    Card border reflects the component furthest from delivery. Units 07 = Buck Hoist container.
  </div>
  <div id="bldg-grid"></div>
</div>

<!-- ===== DAMAGE / OPEN ITEMS ===== -->
<div id="dmg" class="panel">
  <div class="form-card">
    <h3>Log a damage or open item</h3>
    <p class="form-sub">Fill in what you found. Date is auto-stamped. Items persist in your browser and export to CSV.</p>
    <div class="form-grid">
      <div class="field">
        <label>Unit #</label>
        <select id="f-unit"><option value="">&#8212; Select unit &#8212;</option></select>
      </div>
      <div class="field">
        <label>Date reported</label>
        <div class="date-fixed" id="today-display"></div>
      </div>
      <div class="field">
        <label>Item code (if known)</label>
        <input type="text" id="f-code" placeholder="e.g. 10C-CAB-0042">
      </div>
      <div class="field">
        <label>Component / cabinet</label>
        <input type="text" id="f-component" placeholder="e.g. Island back panel">
      </div>
    </div>
    <div class="form-grid single">
      <div class="field">
        <label>Issue description</label>
        <textarea id="f-issue" placeholder="Describe what was found &#8212; damage type, location, severity&#8230;"></textarea>
      </div>
      <div class="field" style="margin-top:10px">
        <label>Photo reference / filename</label>
        <input type="text" id="f-photo" placeholder="e.g. unit_1006_damage.jpg">
      </div>
      <div class="field" style="margin-top:10px">
        <label>Comment / action taken</label>
        <input type="text" id="f-comment" placeholder="e.g. Ordered replacement, touch-up applied&#8230;">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn-submit" onclick="submitItem()">Log item</button>
      <span class="form-msg" id="form-msg"></span>
    </div>
  </div>
  <div style="margin-top:20px">
    <div class="export-bar">
      <span class="count" id="dmg-count-label">0 items logged</span>
      <button class="btn-export" onclick="exportCSV()">
        <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Export to CSV
      </button>
    </div>
    <div id="dmg-open-group" class="dmg-group open" style="display:none">
      <div class="dmg-group-hdr"><span style="color:#CC3030">&#9679;</span> Open items <span id="open-count" class="bdg b-open" style="margin-left:4px"></span></div>
      <table><thead><tr><th>Unit</th><th>Date</th><th>Code</th><th>Component</th><th>Issue</th><th>Photo</th><th>Comment</th><th>Status</th><th>Action</th></tr></thead>
      <tbody id="open-tbody"></tbody></table>
    </div>
    <div id="dmg-closed-group" class="dmg-group closed" style="display:none;margin-top:8px">
      <div class="dmg-group-hdr"><span style="color:#2A7A3A">&#9679;</span> Closed items <span id="closed-count" class="bdg b-closed" style="margin-left:4px"></span></div>
      <table><thead><tr><th>Unit</th><th>Date</th><th>Code</th><th>Component</th><th>Issue</th><th>Photo</th><th>Comment</th><th>Status</th><th>Action</th></tr></thead>
      <tbody id="closed-tbody"></tbody></table>
    </div>
    <div id="empty-state" class="empty-state">
      <div class="big">&#128203;</div>
      No items logged yet. Use the form above to log the first damage or open item.
    </div>
  </div>
</div>

</div><!-- /shell -->

<script>
// ── DATA (auto-generated {generated_on}) ─────────────────────
const CONTAINERS = {js_containers()};

const ROJ = {js_roj()};

const EXCEPTIONS = {js_exceptions()};

const V2_UNITS = new Set([4, 6, 10, 12]);

// ── STATUS HELPERS ──
const STATUS_RANK = {{ D:7, PLD:6, CHS:5, ENR:4, LDD:3, LDG:2, PROJ:1 }};
function worstStatus(arr) {{
  return arr.filter(Boolean).reduce((w, s) => (STATUS_RANK[s]||0) < (STATUS_RANK[w]||0) ? s : w);
}}
function statusDot(s)   {{ return `d-${{s}}`; }}
function statusBadge(s) {{ return `b-${{s}}`; }}
function statusColor(s) {{
  return {{ D:'#0D7A60', PLD:'#0D7A60', CHS:'#5A4AAA', ENR:'#2A5AAA', LDD:'#2A7A3A', LDG:'#888880', PROJ:'#D4890A' }}[s] || '#9A9A94';
}}

// ── TODAY ──
function today() {{ return new Date().toLocaleDateString('en-US',{{month:'short',day:'numeric',year:'numeric'}}); }}
function todayISO() {{ return new Date().toISOString().slice(0,10); }}
document.getElementById('today-pill').textContent = today();
document.getElementById('today-display').textContent = today();

function sw(id, el) {{
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.getElementById(id).classList.add('on');
  el.classList.add('on');
}}

// ── CONTAINER SCHEDULE ──
(function() {{
  let html = '';
  CONTAINERS.forEach(c => {{
    const dot = statusDot(c.status);
    const bdg = statusBadge(c.status);
    html += `<tr>
      <td><span style="display:inline-flex;align-items:center;gap:5px"><span class="dot ${{dot}}"></span>${{c.num}}</span></td>
      <td><span class="bdg ${{bdg}}">${{c.status}}</span></td>
      <td>${{c.week}}</td><td>Fl. ${{c.floor}}</td>
      <td style="color:var(--text-3)">${{c.unitQty}}</td>
      <td>${{c.loadDate}}</td><td>${{c.shipDate}}</td>
      <td style="color:var(--text-3)">${{c.vessel}}</td>
      <td>${{c.etaChs}}</td><td>${{c.railDate}}</td><td>${{c.etaNash}}</td>
      <td style="font-weight:500">${{c.delivery}}</td>
    </tr>`;
  }});
  document.getElementById('ctn-tbody').innerHTML = html;
}})();

// ── PROJECTIONS ──
(function() {{
  let html = '';
  CONTAINERS.forEach(c => {{
    const roj = ROJ[c.floor] || {{}};
    const rowCls = c.ordered ? 'proj-ok' : 'proj-warn';
    const ordBdg = c.ordered
      ? `<span class="bdg b-LDD">Ordered</span>`
      : `<span class="bdg b-PROJ">Pending</span>`;
    html += `<tr class="${{rowCls}}">
      <td>CTN ${{c.num}}</td>
      <td><span class="bdg ${{statusBadge(c.status)}}">${{c.status}}</span></td>
      <td>Floor ${{c.floor}}</td>
      <td>${{c.week}}</td>
      <td>${{c.loadDate}}</td><td>${{c.shipDate}}</td>
      <td>${{c.etaChs}}</td><td>${{c.railDate}}</td><td>${{c.etaNash}}</td>
      <td style="font-weight:500">${{c.delivery}}</td>
      <td style="font-weight:500;color:var(--accent)">${{roj.window || '&#8212;'}}</td>
      <td style="color:var(--text-2)">${{roj.install || '&#8212;'}}</td>
      <td>${{ordBdg}}</td>
    </tr>`;
  }});
  document.getElementById('proj-tbody').innerHTML = html;
}})();

// ── BUILDING GRID ──
(function() {{
  let html = '';
  for (let floor = 39; floor >= 10; floor--) {{
    const ctn = CONTAINERS[floor - 10];
    const roj = ROJ[floor] || {{}};
    html += `<div class="floor-block">
      <div class="floor-hdr">
        <span class="floor-num">${{floor}}</span>
        Floor ${{floor}}
        <span class="floor-meta">CTN ${{ctn.num}} &middot; ${{ctn.week}} &middot; ${{ctn.delivery}}</span>
        ${{roj.window ? `<span class="floor-meta">ROJ ${{roj.window}}</span>` : ''}}
        <span class="bdg ${{statusBadge(ctn.status)}}" style="margin-left:auto">${{ctn.status}}</span>
      </div>
      <div class="unit-grid">`;
    for (let u = 1; u <= 12; u++) {{
      const unitId = floor * 100 + u;
      const exc    = EXCEPTIONS[unitId] || {{}};
      const hasV2  = V2_UNITS.has(u);
      const isBH   = (u === 7);
      const kStatus  = isBH  ? 'PROJ' : (exc.kStatus  || ctn.status);
      const v1Status = exc.v1Status || ctn.status;
      const v2Status = hasV2 ? (exc.v2Status || ctn.status) : null;
      const kCtnr  = isBH  ? 'BH'  : (exc.kCtnr  || `CTN ${{ctn.num}}`);
      const v1Ctnr = exc.v1Ctnr || `CTN ${{ctn.num}}`;
      const v2Ctnr = hasV2 ? (exc.v2Ctnr || `CTN ${{ctn.num}}`) : null;
      const allStatuses = [kStatus, v1Status, v2Status].filter(Boolean);
      const worst = worstStatus(allStatuses);
      const cardCls = isBH ? 'ub-BH' : `ub-${{worst}}`;
      if (isBH) {{
        html += `<div class="ucard ${{cardCls}}">
          <div class="uid">${{unitId}}</div>
          <div class="ucomp"><span class="mdot d-${{kStatus}}"></span><span class="ucomp-lbl">K</span><span class="ucomp-val" style="color:${{statusColor(kStatus)}}">BH</span></div>
          <div class="ucomp"><span class="mdot d-${{v1Status}}"></span><span class="ucomp-lbl">V1</span><span class="ucomp-val" style="color:${{statusColor(v1Status)}}">${{v1Status}}</span></div>
        </div>`;
      }} else {{
        const kNote  = exc.kCtnr  ? `<span class="ucomp-note">${{exc.kCtnr}}</span>`  : '';
        const v1Note = exc.v1Ctnr ? `<span class="ucomp-note">${{exc.v1Ctnr}}</span>` : '';
        const v2Row  = hasV2 ? `<div class="ucomp"><span class="mdot d-${{v2Status}}"></span><span class="ucomp-lbl">V2</span><span class="ucomp-val" style="color:${{statusColor(v2Status)}}">${{v2Status}}</span></div>` : '';
        html += `<div class="ucard ${{cardCls}}">
          <div class="uid">${{unitId}}</div>
          <div class="ucomp"><span class="mdot d-${{kStatus}}"></span><span class="ucomp-lbl">K</span><span class="ucomp-val" style="color:${{statusColor(kStatus)}}">${{kStatus}}</span>${{kNote}}</div>
          <div class="ucomp"><span class="mdot d-${{v1Status}}"></span><span class="ucomp-lbl">V1</span><span class="ucomp-val" style="color:${{statusColor(v1Status)}}">${{v1Status}}</span>${{v1Note}}</div>
          ${{v2Row}}
        </div>`;
      }}
    }}
    html += `</div></div>`;
  }}
  document.getElementById('bldg-grid').innerHTML = html;
}})();

// ── UNIT DROPDOWN ──
(function() {{
  const sel = document.getElementById('f-unit');
  let html = '<option value="">&#8212; Select unit &#8212;</option>';
  for (let floor = 10; floor <= 39; floor++) {{
    const ctn = CONTAINERS[floor - 10];
    html += `<optgroup label="Floor ${{floor}} (CTN ${{ctn.num}} &middot; ${{ctn.status}})">`;
    for (let u = 1; u <= 12; u++) {{
      const unitId = floor * 100 + u;
      const exc = EXCEPTIONS[unitId] || {{}};
      const note = u === 7 ? ' [BH kitchen]' : (exc.v1Ctnr ? ` [V1: ${{exc.v1Ctnr}}]` : (exc.kCtnr ? ` [K: ${{exc.kCtnr}}]` : ''));
      html += `<option value="${{unitId}}">${{unitId}}${{note}}</option>`;
    }}
    html += '</optgroup>';
  }}
  sel.innerHTML = html;
}})();

// ── DAMAGE / OPEN ITEMS ──
const DMG_KEY = '10c_damage_v2';
function loadItems() {{ try {{ return JSON.parse(localStorage.getItem(DMG_KEY)||'[]'); }} catch(e){{ return []; }} }}
function saveItems(items) {{ localStorage.setItem(DMG_KEY, JSON.stringify(items)); }}

function renderItems() {{
  const items = loadItems();
  const open   = items.filter(i => i.status === 'Open');
  const closed = items.filter(i => i.status === 'Closed');
  document.getElementById('dmg-count-label').textContent = items.length + ' item' + (items.length!==1?'s':'') + ' logged';
  document.getElementById('dmg-count-metric').textContent = items.length;
  document.getElementById('dmg-sub-metric').textContent = items.length>0 ? open.length+' open · '+closed.length+' closed' : 'None logged yet';
  document.getElementById('open-count').textContent  = open.length;
  document.getElementById('closed-count').textContent = closed.length;
  document.getElementById('empty-state').style.display = items.length===0 ? 'block' : 'none';
  document.getElementById('dmg-open-group').style.display   = open.length>0   ? 'block' : 'none';
  document.getElementById('dmg-closed-group').style.display = closed.length>0 ? 'block' : 'none';
  function row(item) {{
    const cl = item.status==='Closed';
    return `<tr>
      <td style="font-weight:600">${{item.unit}}</td><td>${{item.date}}</td>
      <td class="mono">${{item.code||'&#8212;'}}</td><td>${{item.component||'&#8212;'}}</td>
      <td>${{item.issue}}</td><td>${{item.photo||'&#8212;'}}</td><td>${{item.comment||'&#8212;'}}</td>
      <td><span class="bdg ${{cl?'b-closed':'b-open'}}">${{item.status}}</span></td>
      <td>
        ${{!cl?`<button onclick="closeItem(${{item.id}})" style="font-size:10px;padding:2px 8px;border:1px solid var(--border-strong);border-radius:4px;background:none;cursor:pointer;font-family:inherit;color:var(--ldd-text)">Close</button>`:''}}
        <button onclick="deleteItem(${{item.id}})" style="font-size:10px;padding:2px 8px;border:1px solid var(--border-strong);border-radius:4px;background:none;cursor:pointer;font-family:inherit;color:var(--text-3);margin-left:4px">Remove</button>
      </td>
    </tr>`;
  }}
  document.getElementById('open-tbody').innerHTML   = open.map(row).join('');
  document.getElementById('closed-tbody').innerHTML = closed.map(row).join('');
}}

function submitItem() {{
  const unit  = document.getElementById('f-unit').value;
  const issue = document.getElementById('f-issue').value.trim();
  if (!unit)  {{ showMsg('Please select a unit.',      true); return; }}
  if (!issue) {{ showMsg('Please describe the issue.', true); return; }}
  const items = loadItems();
  items.unshift({{ id:Date.now(), unit, date:today(), dateISO:todayISO(),
    code:document.getElementById('f-code').value.trim(),
    component:document.getElementById('f-component').value.trim(),
    issue, photo:document.getElementById('f-photo').value.trim(),
    comment:document.getElementById('f-comment').value.trim(), status:'Open' }});
  saveItems(items); renderItems();
  ['f-unit','f-code','f-component','f-issue','f-photo','f-comment'].forEach(id => document.getElementById(id).value = '');
  showMsg('Item logged.', false);
}}
function showMsg(t,e) {{
  const m = document.getElementById('form-msg');
  m.textContent=t; m.className='form-msg'+(e?' err':''); m.style.display='inline';
  setTimeout(()=>{{ m.style.display='none'; }}, 3200);
}}
function closeItem(id) {{
  const items = loadItems(); const i = items.find(x=>x.id===id);
  if(i){{ i.status='Closed'; saveItems(items); renderItems(); }}
}}
function deleteItem(id) {{
  if(!confirm('Remove permanently?')) return;
  saveItems(loadItems().filter(x=>x.id!==id)); renderItems();
}}
function exportCSV() {{
  const items = loadItems();
  if(!items.length){{ alert('No items to export.'); return; }}
  const hdr = ['Unit','Date','Code','Component','Issue','Photo','Comment','Status'];
  const rows = items.map(i=>[i.unit,i.date,i.code||'',i.component||'',i.issue,i.photo||'',i.comment||'',i.status]
    .map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(','));
  const blob = new Blob([[hdr.join(','),...rows].join('\\r\\n')],{{type:'text/csv'}});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a'); a.href=url;
  a.download = '10C_Damage_'+todayISO()+'.csv'; a.click();
  URL.revokeObjectURL(url);
}}
renderItems();
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────
# 7.  WRITE HTML
# ─────────────────────────────────────────────────────────────

print(f"\nWriting  {os.path.basename(OUTPUT_HTML)} …")
with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(HTML)

size_kb = os.path.getsize(OUTPUT_HTML) / 1024
print(f"  → {size_kb:.1f} KB written")

# ─────────────────────────────────────────────────────────────
# 8.  GIT PUSH
# ─────────────────────────────────────────────────────────────

if AUTO_GIT:
    print("\nPushing to GitHub …")
    def run_git(*args):
        result = subprocess.run(["git"] + list(args), cwd=HERE,
                                capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip())
        return result.stdout.strip()

    try:
        run_git("add", "10C_Container_Dashboard.html")
        msg = f"Dashboard update {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        run_git("commit", "-m", msg)
        run_git("push")
        print(f"  → Pushed: \"{msg}\"")
        print("  → GitHub Pages will refresh in ~30–60 seconds.")
    except RuntimeError as e:
        err = str(e)
        if "nothing to commit" in err:
            print("  → No changes — HTML is already up to date on GitHub.")
        elif "git: command not found" in err or ("'git'" in err and "not recognized" in err.lower()):
            print("  ! Git not found. Install Git for Windows from https://git-scm.com")
            print("    Then follow the one-time setup instructions in README.md")
        else:
            print(f"  ! Git error: {err}")
            print("    Check that your GitHub remote is configured (see setup instructions).")

# ─────────────────────────────────────────────────────────────
# DONE
# ─────────────────────────────────────────────────────────────

print(f"""
{'─'*56}
  Done!  {len(CONTAINERS)} containers  ·  {len(EXCEPTIONS)} exceptions  ·  {len(ROJ)} ROJ floors
  File:  {OUTPUT_HTML}
{'─'*56}

  ONE-TIME SETUP (GitHub Pages)
  ─────────────────────────────
  1. Create a free GitHub account at  https://github.com
  2. New repository → name it  1010church-dashboard  (public)
  3. Install Git for Windows from  https://git-scm.com
  4. Open Command Prompt in this folder and run:
       git init
       git remote add origin https://github.com/YOUR_USERNAME/1010church-dashboard.git
       git add .
       git commit -m "initial"
       git branch -M main
       git push -u origin main
  5. On GitHub: Settings → Pages → Source: Deploy from branch → main
  6. Your live URL:  https://YOUR_USERNAME.github.io/1010church-dashboard/10C_Container_Dashboard.html

  After that, just double-click this script whenever you update
  the Excel file — it regenerates and pushes automatically.
{'─'*56}
""")

# Keep window open if double-clicked on Windows
if sys.platform == "win32":
    input("  Press Enter to close ...")
