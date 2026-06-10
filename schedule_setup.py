#!/usr/bin/env python3
"""
Creates Windows Task Scheduler tasks to run update_cpi.py on each BLS CPI
release date at the configured local time.

BLS standard release time: 8:30 AM Eastern Time (9:30 AM BRT, UTC-3).
Adjust LOCAL_HOUR / LOCAL_MINUTE so the task runs after BLS publishes.

Run once (as Administrator if needed):
    python schedule_setup.py

--- ANNUAL UPDATE (every January) ---
1. Open https://www.bls.gov/schedule/news_release/cpi.htm in a browser
2. Copy the 12 release dates for the new year
3. Add a new RELEASE_DATES_20XX list below (same format as existing ones)
4. Add it to ALL_DATES
5. Run: python schedule_setup.py
NOTE: BLS blocks automated access to this URL — manual copy-paste is required.
"""

import subprocess
import sys
from pathlib import Path

REPO_DIR = Path(__file__).parent.resolve()
PYTHON   = sys.executable
SCRIPT   = REPO_DIR / "update_cpi.py"

# Adjust these for your local timezone:
#   BLS releases at 8:30 AM ET. Eastern Daylight Time = UTC-4.
#   If you are in BRT (UTC-3): 8:30 AM ET = 9:30 AM BRT.
#   Running at 9:32 AM BRT — update_cpi.py retries up to 3x (5 min apart)
#   if the BLS API is not yet updated, so no hard buffer needed here.
LOCAL_HOUR   = 9    # local hour to run the task
LOCAL_MINUTE = 32   # local minute

# BLS CPI release dates for 2025–2026 (release date, not reference month).
# Source: https://www.bls.gov/schedule/news_release/cpi.htm
# Update this list each January with the new year's schedule.
RELEASE_DATES_2025 = [
    "2025-01-15", "2025-02-12", "2025-03-12", "2025-04-10",
    "2025-05-13", "2025-06-11", "2025-07-15", "2025-08-12",
    "2025-09-10", "2025-10-15", "2025-11-12", "2025-12-10",
]

RELEASE_DATES_2026 = [
    "2026-01-13", "2026-02-13", "2026-03-11", "2026-04-10",
    "2026-05-12", "2026-06-10", "2026-07-14", "2026-08-12",
    "2026-09-11", "2026-10-14", "2026-11-10", "2026-12-10",
]

# Add RELEASE_DATES_2027 here each January and include it in ALL_DATES.
ALL_DATES = RELEASE_DATES_2025 + RELEASE_DATES_2026


def create_task(date_str: str):
    task_name = f"CPI_Update_{date_str}"
    start_time = f"{LOCAL_HOUR:02d}:{LOCAL_MINUTE:02d}"

    # schtasks date format depends on locale; use DD/MM/YYYY for PT-BR Windows
    y, m, d = date_str.split("-")
    start_date = f"{d}/{m}/{y}"

    cmd = [
        "schtasks", "/create",
        "/tn", task_name,
        "/tr", f'"{PYTHON}" "{SCRIPT}"',
        "/sc", "ONCE",
        "/sd", start_date,
        "/st", start_time,
        "/f",   # overwrite if exists
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"  Scheduled: {task_name}  at {start_time} on {start_date}")
    else:
        print(f"  FAILED:    {task_name}: {result.stderr.strip()}")


def delete_past_tasks():
    """Remove tasks for dates already passed (optional cleanup)."""
    from datetime import date
    today = date.today().isoformat()
    for d in ALL_DATES:
        if d < today:
            task_name = f"CPI_Update_{d}"
            subprocess.run(
                ["schtasks", "/delete", "/tn", task_name, "/f"],
                capture_output=True
            )


if __name__ == "__main__":
    from datetime import date
    today = date.today().isoformat()

    print(f"Creating scheduled tasks (local {LOCAL_HOUR:02d}:{LOCAL_MINUTE:02d})...")
    created = 0
    for d in ALL_DATES:
        if d >= today:
            create_task(d)
            created += 1

    print(f"\n{created} task(s) created.")
    print(f"Script that will run: {SCRIPT}")
    print(f"Python:               {PYTHON}")
    print("\nTo view tasks: schtasks /query /tn CPI_Update_* /fo LIST")
    print("To delete all: python schedule_setup.py --delete")

    if "--delete" in sys.argv:
        print("\nDeleting past tasks...")
        delete_past_tasks()
