"""
Healthcare Reporting Automation – Jenkins Master Scheduler
-----------------------------------------------------------

This script acts as a central orchestration layer for all
healthcare monitoring reports in this repository.

What It Does:
-------------
1. Waits for the MIS_Report.xlsx file to be updated (today's date).
2. Once updated:
   - Performs pre-cleanup (deletes old Excel outputs).
   - Executes all individual report scripts sequentially.
3. Maintains execution logs.
4. Stops safely if MIS is not updated within a defined timeout window.

Key Features:
-------------
- Workspace-relative paths (GitHub/Jenkins friendly)
- Automatic pre-cleanup of old report files
- Timeout safety mechanism
- Daily log file generation
- Modular execution of independent report scripts
- Non-blocking continuation if one script fails

Designed For:
-------------
- Jenkins automation
- Scheduled enterprise reporting
- Multi-stakeholder healthcare reporting workflows
- Controlled orchestration of modular automation scripts

Author: Skanda N Raj
"""

import os
import datetime
import time
import subprocess
import sys

# ================== CONFIG ==================

# Use system Python (Jenkins environment Python)
PYTHON_EXE = "python"

# Workspace-relative MIS file
MIS_FILE_PATH = "data/MIS_Report.xlsx"

# Jenkins script paths (repo relative)
SCRIPT_PATHS = [
    "Cancelled_Appointments_Monitoring_Report/jenkins_version.py",
    "Completed_Consultations_Monitoring_Report/jenkins_version.py",
    "Dropout_Consultation_Report/jenkins_version.py",
    "Missing_Prescription_Report/jenkins_version.py",
    "Ops_Data_Sanitization/jenkins_version.py"
]

RECHECK_INTERVAL_MIN = 30  # minutes
MAX_WAIT_HOURS = 6         # safety stop

# Log directory (workspace relative)
LOG_DIR = "logs"

# Excel cleanup folders (workspace relative)
EXCEL_DELETE_FOLDERS = [
    "Dropout_Consultation_Report/output",
    "Completed_Consultations_Monitoring_Report/output",
    "Missing_Prescription_Report/output",
]

# ============================================


def get_log_file():
    os.makedirs(LOG_DIR, exist_ok=True)
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return os.path.join(LOG_DIR, f"jenkins_run_{today}.txt")


def log(message):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}"
    print(line, flush=True)
    with open(get_log_file(), "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ================= PRE-CLEANUP =================

def preclean_folders():
    log("Pre-cleanup started")

    for folder in EXCEL_DELETE_FOLDERS:
        if not os.path.exists(folder):
            log(f"Folder not found: {folder}")
            continue

        deleted = False
        for file in os.listdir(folder):
            if file.lower().endswith((".xls", ".xlsx")):
                os.remove(os.path.join(folder, file))
                log(f"Deleted Excel: {folder}\\{file}")
                deleted = True

        if not deleted:
            log(f"No Excel files in {folder}")

    log("Pre-cleanup completed")


# ================= MIS CHECK =================

def is_mis_updated_today():
    try:
        modified = datetime.datetime.fromtimestamp(
            os.path.getmtime(MIS_FILE_PATH)
        )
        log(f"MIS last modified at: {modified}")
        return modified.date() == datetime.datetime.now().date()
    except Exception as e:
        log(f"MIS check failed: {e}")
        return False


# ================= SCRIPT RUNNER =================

def run_all_scripts():
    log("Starting script execution")

    for script in SCRIPT_PATHS:
        name = os.path.basename(script)
        log(f"Running {name}")

        try:
            subprocess.run([PYTHON_EXE, script], check=True)
            log(f"{name} completed successfully")
        except subprocess.CalledProcessError as e:
            log(f"{name} FAILED: {e}")
            log("Continuing with next script")

    log("All scripts processed")


# ================= MAIN FLOW =================

def main():
    log("====================================")
    log("Jenkins Job Started")
    log("Waiting for MIS update")
    log("====================================")

    start_time = datetime.datetime.now()
    timeout = datetime.timedelta(hours=MAX_WAIT_HOURS)

    while True:

        if is_mis_updated_today():
            log("MIS updated today. Proceeding...")
            preclean_folders()
            run_all_scripts()
            log("Job completed successfully")
            sys.exit(0)

        if datetime.datetime.now() - start_time > timeout:
            log("MIS not updated within allowed window. Exiting job.")
            sys.exit(1)

        log(f"MIS not updated. Rechecking in {RECHECK_INTERVAL_MIN} minutes")
        time.sleep(RECHECK_INTERVAL_MIN * 60)


if __name__ == "__main__":
    main()
