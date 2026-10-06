# Global IP Scan Dashboard

A live, publicly accessible dashboard for the **Global IP Scan Kanban Board** — built from the Excel tracking sheet and hosted on GitHub Pages.

**Live URL:** https://gayathri19945.github.io/ip-scan-dashboard/

---

## Repository Contents

| File | Purpose |
|---|---|
| `index.html` | Interactive dashboard (Overview + SLA Analysis tabs) |
| `scan_data.json` | Extracted scan records from the Excel file (432+ entries) |
| `github_ip_scan_sync.py` | Python script to sync Excel data → GitHub Project Board |
| `README.md` | This file |

---

## Dashboard Features

### Tab 1 — Overview
- **KPI cards**: Total / In Progress / Completed / Backlog / Cancelled
- **Filter bar**: Search, State, Priority, Location, Analyst, Category
- **Date range**: Filter completed scans by Closed Date with quick ranges (Last 7/30/90 days, MTD, YTD)
- **Daily completions chart**: Bar chart of scans closed per day
- **Sortable table**: All 430+ scan records with pagination
- **Charts**: By Location and By FOSS Analyst (In Progress)

### Tab 2 — SLA Analysis
- **KPI cards**: Within SLA / Breached SLA / No Data / Avg Days Within / Avg Overdue
- **Filter bar**: Search, SLA Status, Comment Reason, Location, Priority, Analyst, Continuous Scan
  - Default filters: **Continuous Scan = N**, **Comment Reason = Resource Constraints**
- **Date range filter**: Filter by Closed Date / End of Development / RDM / Submission Date
- **Stacked charts**: SLA status (Within vs Breached) by Comment Reason, Location, Analyst
- **Detail table**: Program, SLA Days (+/-), Status badge, Comment Reason, Location, Analyst

---

## Updating the Dashboard Data

When a new Excel file is available:

1. Copy the new `Global IP Scan Kanban Board.xlsx` into this folder
2. Run the extraction script:

```bash
python -c "
import openpyxl, json, warnings
from datetime import datetime, date
warnings.filterwarnings('ignore')
# ... (see github_ip_scan_sync.py for full extraction logic)
"
```

3. Push `scan_data.json` to GitHub — the Pages site rebuilds in ~30 seconds.

---

## GitHub Project Board Sync

The `github_ip_scan_sync.py` script syncs scan entries to a **GitHub Projects v2** board.

### Requirements
```bash
pip install requests openpyxl
```

### Usage
```bash
# Dry-run (preview only)
python github_ip_scan_sync.py --token YOUR_PAT --repo owner/repo --dry-run

# Sync only In Progress items
python github_ip_scan_sync.py --token YOUR_PAT --repo owner/repo --state-filter "In Progress"

# Full sync
python github_ip_scan_sync.py --token YOUR_PAT --repo owner/repo

# Use an existing project board
python github_ip_scan_sync.py --token YOUR_PAT --repo owner/repo --project-id PVT_xxx
```

### Required GitHub PAT scopes
- `repo` — create issues and labels
- `project` — read/write GitHub Projects v2

### Custom fields created automatically
Priority · Location · Program Category · FOSS Analyst · L1 Manager ·
End of Development · RTC · SLOC Count · Black Duck Scan · Continuous Scan ·
Crown Jewel · Severity · Owner #

---

## Data Source

Excel file: `Global IP Scan Kanban Board.xlsx` — sheet: **Active Scans**

| Field | Description |
|---|---|
| Program (Delivery) | Scan entry title |
| FOSS Scan State | Backlog / In Progress / Done / Cancelled |
| Priority | P1 / P2 / P3 |
| Location | GTLC India / Germany / China / Vietnam / GLOBAL |
| FOSS Analyst | Assigned analyst |
| Closed within SLA | Days relative to SLA deadline (positive = within, negative = breached) |
| Comments | Delay reason: Delay From Dev / Resource Constraints / etc. |
| Continuous scan | Y / N |
