"""
Healthcare Reporting Automation – Jenkins Master Scheduler
-----------------------------------------------------------

This script acts as a central orchestration layer for all
healthcare monitoring reports in this repository.
"""

import os
import datetime
import time
import subprocess
import sys

# ================= FIX FOR JENKINS UNICODE =================
# Prevents UnicodeEncodeError in Jenkins console
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# ================== CONFIG ==================

# Use system Python (Jenkins environment Python)
PYTHON_EXE = r"C:\Users\SKANDA NAGARAJ\AppData\Local\Programs\Python\Python311\python.exe"

# Workspace-relative MIS file
MIS_FILE_PATH = r"E:\COURSES AND PROJECTS (DATA SCIENCE)\PROJECTS (ASTER DM HEALTHCARE)\Email Automation\Dummy Dataset.xlsx"

# Jenkins script paths (repo relative)
SCRIPT_PATHS = [
    r"python file path",
    r"python file path",
    r"python file path",
    r"python file path",
    r"python file path"
]

RECHECK_INTERVAL_MIN = 30
MAX_WAIT_HOURS = 6

# Log directory
LOG_DIR = "logs"

# Excel cleanup folders
EXCEL_DELETE_FOLDER_1 = r"excel folder file path"
EXCEL_DELETE_FOLDER_2 = r"excel folder file path"
EXCEL_DELETE_FOLDER_3 = r"excel folder file path"   


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

    for folder in [EXCEL_DELETE_FOLDER_1, EXCEL_DELETE_FOLDER_2, EXCEL_DELETE_FOLDER_3]:

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
