# Logistics Dashboard — Setup Guide

This guide explains how the 1010 Church logistics dashboard was built so you can replicate it for your own project.

---

## What It Is

A self-contained, single-file HTML dashboard that tracks container shipments across a construction project. It reads data from an Excel file, auto-updates via a Python script, and is hosted publicly on GitHub Pages. No server, no database, no frameworks.

The dashboard has three tabs:
- **Container schedule** — a table of all containers with statuses, dates, unit counts, and a metrics/chart row at the top
- **Building matrix** — a color-coded grid showing delivery status per unit per floor
- **Damage / open items** — a form to log and track site issues (stored locally in the browser)

---

## File Structure

You need four files in one folder (your project folder):

```
10C_Container_Dashboard.html   ← the dashboard (the permanent source of truth for design)
update_dashboard.py            ← Python script that reads Excel and updates the HTML
index.html                     ← one-line redirect so your GitHub Pages URL works cleanly
10C_Logistcs Output.xlsm       ← your Excel data file (not committed to GitHub if sensitive)
```

The HTML file is **never regenerated from scratch**. The Python script only replaces a small data block inside it, leaving all layout and design untouched.

---

## Excel File Requirements

The Python script reads from three sheets in your `.xlsm` or `.xlsx` file. The sheet names are configurable at the top of the script.

### Sheet 1 — Container Schedule (spill)
A table with one row per container. Required columns (script finds them by name, case-insensitive):

| Column | Description |
|---|---|
| `CONT. STATUS` | Status code: `D`, `INYRD`, `CHS`, `ENR`, `LDD`, `LDG`, `PROJ` |
| `CONT. #` | Container number (integer 1–30) |
| `WEEK` | Load week number |
| `FLOORS` | Floor number served |
| `Kitchens` | Unit numbers for kitchen items (e.g. `1001:1006,1008:1012`) |
| `Vanity 1` | Unit numbers for vanity 1 items |
| `Vanity 2` | Unit numbers for vanity 2 items |
| `UNIT Q.ty` | Number of units in this container |
| `ORDERED` | Boolean — whether the container has been ordered |
| `LOAD DATE` | Date loaded at factory |
| `SHIP DATE` | Date shipped |
| `VESSEL` | Vessel name |
| `PORT ARRIVAL DATE` | Arrival at Charleston port |
| `RAIL DATE` | Rail departure date |
| `ARRIVAL DATE` | Arrival at Nashville yard |
| `DELIVERY DATE` | Final delivery date to site |

### Sheet 2 — Building Matrix (status)
One row per unit. Required columns:

| Column | Description |
|---|---|
| `UNIT` | Unit identifier (e.g. `1001`, `2304`) |
| `TIER` | Unit tier/type |
| `K CTNR #` | Kitchen container number |
| `V1 CTNR #` | Vanity 1 container number |
| `V2 CTNR #` | Vanity 2 container number (leave blank if unit has no V2) |
| `K STATUS` | Kitchen delivery status |
| `V1 STATUS` | Vanity 1 delivery status |
| `V2 STATUS` | Vanity 2 delivery status |

### Sheet 3 — ROJ Dates (optional)
Room-over-job schedule. If not present, the script skips it gracefully.

---

## Status Codes

The dashboard uses these status codes throughout:

| Code | Meaning | Color |
|---|---|---|
| `D` | Delivered | Dark green `#1A5C47` (white text) |
| `INYRD` | In yard — arrived at Nashville rail yard, awaiting delivery | `#22735C` |
| `CHS` | At Charleston port | `#2E8C70` |
| `ENR` | En route / sea freight | `#3EA882` |
| `LDD` | Loaded on rail | `#6DC4A8` |
| `LDG` | Loading at factory | `#A0D8C8` |
| `PROJ` | Projected — not yet loaded | No fill, dark text only |

---

## How the Python Script Works

`update_dashboard.py` does the following every time you run it:

1. Opens the Excel file with `openpyxl` (data only — no formulas)
2. Reads the container schedule, building matrix, ROJ dates, and transit performance rows
3. Opens the existing `10C_Container_Dashboard.html`
4. Finds the marker `// ── DATA` inside the HTML
5. Replaces only the data block between that marker and `const V2_UNITS` with fresh data
6. Updates the `Updated ...` timestamp in the subtitle
7. Writes `index.html` (a redirect so the GitHub Pages root URL works)
8. Runs `git add → commit → push` automatically

**The HTML design is never touched.** All layout changes you make to the HTML file are permanent and survive every data update.

### Key config at the top of the script:

```python
SITE_DIR   = r"C:\Your\Project\Folder"
EXCEL_FILE = os.path.join(SITE_DIR, "YourFile.xlsm")
OUTPUT_HTML = os.path.join(SITE_DIR, "10C_Container_Dashboard.html")
AUTO_GIT   = True  # set False to skip the git push

SH_CTN_VARIANTS = ["Your Container Sheet Name"]
SH_BLD_VARIANTS = ["Your Building Matrix Sheet Name"]
SH_ROJ_VARIANTS = ["Your ROJ Sheet Name"]  # optional
```

### Special containers (BH, ECT, WCT)
Containers that don't follow the standard floor numbering (e.g. a buck hoist, electrical, or wet column container) are hardcoded at the bottom of the script's CONTAINERS list. Edit them directly in the script.

---

## GitHub Pages Setup

1. Create a new public GitHub repository
2. In your project folder, run:
```
git init
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git add 10C_Container_Dashboard.html index.html update_dashboard.py
git commit -m "Initial commit"
git branch -M main
git push -u origin main --force
```
3. In GitHub → your repo → **Settings → Pages**, set source to **Deploy from branch → main → / (root)**
4. Your dashboard will be live at `https://YOUR-USERNAME.github.io/YOUR-REPO` within 60 seconds

For the password when pushing, use a **Personal Access Token** (not your GitHub password): GitHub → Settings → Developer settings → Personal access tokens → Tokens (classic) → Generate new token → tick **repo**.

After the first push, running `update_dashboard.py` handles all future pushes automatically.

---

## Customising the Dashboard

All visual customisation is done by editing `10C_Container_Dashboard.html` directly. The Python script will never overwrite your changes.

Key things to customise:

- **Project name** — change `1010 Church · Container Dashboard` in the `<title>` and `<h1>` tags
- **Total units** — find `of 360` and update to your unit count
- **Total containers** — find `Total containers` metric card and update the hardcoded `33`
- **Floor range** — find `Floors 10 – 39` in the subtitle
- **Accent colour** — change `--accent: #22735C` in `:root` to your brand colour; the status palette will need updating separately
- **Logos** — place image files in the same folder as the HTML and reference them as `<img src="your-logo.png">` in the topbar
- **Special containers** (BH/ECT/WCT) — edit the hardcoded entries at the bottom of `js_containers()` in the Python script

---

## Running the Update

Double-click `update_dashboard.py` from your project folder. It will:
- Print progress to a console window
- Tell you how many containers and exceptions it found
- Push to GitHub automatically
- Confirm the push with a timestamp

The site refreshes on GitHub Pages within ~30–60 seconds of a successful push.

---

## Asking Claude to Help

When starting a new Claude session on this project, paste this context to get Claude up to speed quickly:

> "I have a logistics dashboard project. It's a single HTML file updated by a Python script that reads from an Excel file. The HTML has a data block between the markers `// ── DATA` and `const V2_UNITS` that the script replaces on each run. The design lives entirely in the HTML and is never touched by the script. The site is hosted on GitHub Pages."

Then share the HTML and Python files so Claude can read the current state before making any changes.
