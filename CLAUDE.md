# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the script

All files must be in the same directory. Run from `C:\Users\rodri\Downloads\`:

```
cd C:\Users\rodri\Downloads
python build_us_cpi.py
```

Output: `US_CPI.xlsx` in the same folder. Takes ~4 minutes (32 API requests total: 16 for indexes + 16 for weights).

## Dependencies

```
pip install requests openpyxl
```

## Architecture

The script has four stages:

1. **`load_mapping`** — reads `cpi_mapping.csv` into a list of dicts with keys `indent`, `category`, `series_id`, `basis`. Rows where `series_id` is empty are header/spacer rows (written to the sheet as label-only rows with no data).

2. **`fetch_all`** — calls the BLS Public Data API v2 in batches. Hard limits enforced by BLS: max 50 series per request, max 20 years per request. With 382 series and a 2000–2026 window, this produces 16 requests (8 batches × 2 year windows). Returns `dict[series_id -> dict[(year, month) -> float]]`.

3. **`fetch_relative_importance`** — fetches the `"Relative Importance"` aspect from December data points for all NSA-equivalent series (`CUUR0000*`). SA series (`CUSR0000*`) are converted via `_nsa_sid()` (replaces `CUSR` → `CUUR`). Same 16-request batching pattern with `aspects=True`. Returns `dict[nsa_series_id -> dict[dec_year -> float]]`. Coverage: December 2012–December 2025 (the BLS API only carries this aspect from 2012 onward). 68 of 382 categories return no RI (very granular NSA sub-items not individually weighted by BLS).

4. **`build_workbook`** — constructs the xlsx with openpyxl. Produces two sheets with identical layout:
   - Row 1: title, Row 2: subtitle, Row 4: column headers, Row 5: date headers (Mon-YY), Rows 6+: one row per mapping entry
   - Column A: indent level, Column B: spacer, Column C: category name, Columns D+: values
   - Pane frozen at D6
   - **US_CPI**: monthly index values (format `0.000`)
   - **US_CPI_Weights**: annual Relative Importance — for each month of year YYYY the weight shown is the December YYYY RI (format `0.000`; pre-2013 blank)

## cpi_mapping.csv format

| Column | Description |
|---|---|
| `indent` | 0–7, hierarchy depth for the expenditure category |
| `category` | Display name (as published by BLS) |
| `series_id` | BLS series ID — `CUSR0000*` = SA (black), `CUUR0000*` = NSA (blue) |
| `basis` | `SA` or `NSA` — controls font color in the output |

SA series use prefix `CUSR0000`, NSA use `CUUR0000`. The BLS only publishes SA for aggregate categories; granular sub-items (e.g., individual food products) fall back to NSA and are colored blue in the output.

## Key constants to adjust

In `build_us_cpi.py`:
- `END_YEAR` — update each year; the API silently caps at the last published month
- `API_KEY` — BLS v2 registered key (allows 500 req/day vs 25 unregistered)
- `START_YEAR` — currently 2000; changing it shifts `min_key` and affects all column positions

## BLS access constraints

The BLS blocks all automated downloads from `download.bls.gov` and `bls.gov/cpi/tables/relative-importance/` (403 for scripts, even with browser headers). The only programmatic path that works is the API at `api.bls.gov/publicAPI/v2/timeseries/data/`. Relative importance data is only available from December 2012 onward through the API `aspects` parameter.

## Reference files (not used by the script directly)

- `cu.series` / `cu.series.txt` — full BLS series catalog, used to build `cpi_mapping.csv`
- `cu.data.1.AllItems`, `cu.data.20.USCommoditiesServicesSpecial` — historical flat files
- `cpi-u-202604.xlsx`, `news-release-table*.xlsx` — BLS press-release tables (Apr 2026)
