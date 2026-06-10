# US CPI Dashboard

Automated pipeline that fetches US CPI-U data from the BLS API, updates a local Excel workbook, and publishes an HTML table to GitHub Pages.

## What it does

- Pulls CPI-U index values and relative importance weights for 40 categories from the BLS public API
- Updates `US_CPI_FINAL.xlsx` with the new month's data (MoM, MM3MA, MM6MA, YoY variations)
- Refreshes Excel formula blocks via COM automation
- Generates `cpi_table.html` with color-coded divergence from historical averages
- Pushes the HTML to GitHub automatically

## Quick start

```powershell
cd C:\Users\rodri\us-cpi
python update_cpi.py
```

> Make sure `US_CPI_FINAL.xlsx` is closed in Excel before running.

## Files

| File | Purpose |
|---|---|
| `build_us_cpi.py` | Fetches raw data from BLS API → `US_CPI.xlsx` |
| `update_cpi.py` | Monthly updater: extends FINAL.xlsx, fills Tabela, generates HTML, pushes to GitHub |
| `schedule_setup.py` | Creates Windows Task Scheduler tasks for each BLS release date |
| `cpi_mapping.csv` | 40 CPI categories with BLS series IDs and SA/NSA basis |
| `cpi_table.html` | Published HTML view of the Tabela (GitHub Pages) |

## Setup

### Install dependencies

```powershell
pip install requests openpyxl pywin32
```

### Schedule automatic updates

```powershell
python schedule_setup.py
```

Creates one Windows Task Scheduler task per BLS release date, running at 10:00 AM local time (30 min after BLS publishes at 8:30 AM ET / 9:30 AM BRT).

## Release schedule

BLS publishes CPI monthly. The pipeline runs automatically on each release date. To view upcoming tasks:

```powershell
schtasks /query /tn CPI_Update_* /fo LIST
```

## Output

- **`C:\Users\rodri\CPI\US_CPI_FINAL.xlsx`** — main Excel workbook (local only, not in git)
- **`cpi_table.html`** — HTML table published to this repo (GitHub Pages)

## Data source

[U.S. Bureau of Labor Statistics](https://www.bls.gov/) — CPI-U (All Urban Consumers), not seasonally adjusted and seasonally adjusted series via the [BLS Public Data API v2](https://www.bls.gov/developers/).
