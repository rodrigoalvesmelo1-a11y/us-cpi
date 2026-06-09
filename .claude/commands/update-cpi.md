# Update US CPI

Runs the monthly US CPI update pipeline for the project at `C:\Users\rodri\us-cpi\`.

## Steps

1. Ask the user to confirm that **Excel is closed** (specifically `C:\Users\rodri\CPI\US_CPI_FINAL.xlsx`) before proceeding, since the script needs to save the file and open it via COM.

2. Run the update script:
```
cd C:\Users\rodri\us-cpi
python update_cpi.py
```
Use the PowerShell tool to execute this. The script takes ~4 minutes (32 BLS API requests).

3. After the script finishes, report:
   - Which month was added (look for "CPI update:" in output)
   - Whether formula blocks were extended
   - Whether GitHub push succeeded
   - Any warnings or errors

4. If the script fails with a **PermissionError** on FINAL.xlsx, ask the user to close the file in Excel and retry.

5. If the script fails with a **BLS API error** (HTTP 429 / rate limit), wait 60 seconds and retry once.

6. On success, remind the user to open `C:\Users\rodri\CPI\US_CPI_FINAL.xlsx` and verify:
   - Tabela sheet: D3 shows the new month in DD/MM/YYYY format
   - MoM/MM3MA/MM6MA/YoY sections have the new month's column populated

## Key facts

- Script: `C:\Users\rodri\us-cpi\update_cpi.py`
- Output file: `C:\Users\rodri\CPI\US_CPI_FINAL.xlsx`
- GitHub repo: `https://github.com/rodrigoalvesmelo1-a11y/us-cpi`
- BLS releases CPI monthly at 8:30 AM ET (9:30 AM BRT)
- Next release: check `schedule_setup.py` for upcoming dates
- Variation row ranges in US_CPI sheet: MoM 392–774, MM3MA 1163–1545, MM6MA 1549–1931, YoY 1935–2317
