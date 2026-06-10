# US CPI Dashboard

Automated pipeline that fetches US CPI-U data from the BLS API, updates a local Excel workbook, and publishes an HTML table to GitHub Pages.

## What it does

- Pulls CPI-U index values and relative importance weights for 382 series from the BLS public API
- Updates `US_CPI_FINAL.xlsx` with the new month's data (MoM, MoMA, MM3MA, MM6MA, YoY variations)
- Computes a custom metric: **Core Services Ex-Primary Rent & OER** (top-down exclusion from Services)
- Refreshes Excel formula blocks via COM automation
- Generates `cpi_table.html` with color-coded divergence from historical averages
- Pushes the HTML to GitHub automatically on each BLS release date

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
| `cpi_mapping.csv` | CPI categories with BLS series IDs and SA/NSA basis |
| `cpi_table.html` | Published HTML view of the Tabela (GitHub Pages) |

## Setup

### Install dependencies

```powershell
pip install requests openpyxl pywin32
```

`pywin32` is required for Excel COM automation (formula cache refresh).

### Schedule automatic updates

```powershell
python schedule_setup.py
```

Creates one Windows Task Scheduler task per BLS release date. The task fires at **9:32 AM BRT** — just after BLS publishes at 8:30 AM ET (9:30 AM BRT). If the BLS API is not yet updated when the script starts, it retries automatically up to 3 times with 5-minute waits between attempts.

To view scheduled tasks:

```powershell
schtasks /query /tn CPI_Update_* /fo LIST
```

### Annual update (every January)

1. Open https://www.bls.gov/schedule/news_release/cpi.htm in a browser
2. Copy the 12 release dates for the new year
3. Add a `RELEASE_DATES_20XX` list to `schedule_setup.py` and include it in `ALL_DATES`
4. Run `python schedule_setup.py`

## Output

- **`C:\Users\rodri\CPI\US_CPI_FINAL.xlsx`** — main Excel workbook (local only, not in git)
- **`cpi_table.html`** — HTML table published to this repo (GitHub Pages)

### US_CPI_FINAL.xlsx sheets

| Sheet | Description |
|---|---|
| `US_CPI` | Raw CPI-U index values + Excel formula blocks for MoM/MM3MA/MM6MA/YoY variations |
| `US_CPI_Weights` | Annual Relative Importance (BLS December RI, from December 2012) |
| `Tabela` | Summary: 4 sections × 40 categories = 160 data rows |

### Custom metric: Core Services Ex-Primary Rent & OER

Computed each update cycle and written to rows 2320–2324 of the US_CPI sheet:

```
I_core = (I_svc × w_svc − I_rent × w_rent − I_oer × w_oer) / (w_svc − w_rent − w_oer)
```

Where `I_svc` = Services less energy services, `I_rent` = Rent of primary residence, `I_oer` = Owners' equivalent rent of residences. Weights are time-varying December Relative Importance values.

## Data source

[U.S. Bureau of Labor Statistics](https://www.bls.gov/) — CPI-U (All Urban Consumers), seasonally adjusted and not seasonally adjusted series via the [BLS Public Data API v2](https://www.bls.gov/developers/).
