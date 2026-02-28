"""
Healthcare Reporting Automation – Daily Scheduler (Local Execution Version)
-----------------------------------------------------------------------------

This script acts as a master scheduler for all healthcare monitoring reports
inside this repository.

What It Does:
-------------
1. Runs daily at a fixed time (CHECK_TIME).
2. Checks whether the MIS report has been updated today.
3. If updated:
      - Performs pre-cleanup (deletes old Excel outputs).
      - Executes all report scripts sequentially.
4. Logs all activities to a daily log file.
5. Sends Windows toast notifications for status updates.

Key Features:
-------------
- Workspace-relative paths (GitHub friendly)
- Automatic Excel cleanup before execution
- Daily logging system
- Windows toast notifications
- Supports both .py and .ipynb scripts
- Continuous background scheduler

Designed For:
-------------
- Local machine automation
- Scheduled healthcare reporting workflows
- Multi-stakeholder reporting execution
- Operational automation pipelines

Author: Skanda N Raj
"""

import os
import datetime
import time
import subprocess
import schedule
from win10toast import ToastNotifier


# =====================================================
#                     CONFIGURATION
# =====================================================

# Workspace-relative MIS file path
MIS_FILE_PATH = "data/MIS_Report.xlsx"

# Report scripts inside repository (executed sequentially)
SCRIPT_PATHS = [
    "Cancelled_Appointments_Monitoring_Report/main.py",
    "Completed_Consultations_Monitoring_Report/main.py",
    "Dropout_Consultation_Report/main.py",
    "Missing_Prescription_Report/main.py",
    "Ops_Data_Sanitization/main.py"
]

# Daily execution time (24-hour format)
CHECK_TIME = "10:30"

# If MIS not updated, recheck interval (in minutes)
RECHECK_INTERVAL = 30

# Log directory (workspace-relative)
LOG_DIR = "logs"

# Output folders where old Excel files should be deleted before execution
EXCEL_DELETE_FOLDER_1 = "Dropout_Consultation_Report/output"
EXCEL_DELETE_FOLDER_2 = "Completed_Consultations_Monitoring_Report/output"
EXCEL_DELETE_FOLDER_3 = "Missing_Prescription_Report/output"


# =====================================================
#                WINDOWS NOTIFICATION SETUP
# =====================================================

# Used for desktop toast notifications (Windows only)
notifier = ToastNotifier()

def notify(title, msg):
    """
    Sends a Windows toast notification.
    Fails silently if notifications are unavailable.
    """
    try:
        notifier.show_toast(title, msg, duration=10, threaded=True)
    except:
        pass


# =====================================================
#                     LOGGING SYSTEM
# =====================================================

def get_log_file():
    """
    Creates daily log file inside logs folder.
    """
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return os.path.join(LOG_DIR, f"scheduler_log_{today}.txt")


def log_message(message):
    """
    Writes timestamped log messages to console and log file.
    """
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    print(log_entry.strip())

    with open(get_log_file(), "a", encoding="utf-8") as f:
        f.write(log_entry)


# =====================================================
#                   PRE-CLEANUP LOGIC
# =====================================================

def preclean_folders():
    """
    Deletes old Excel files from specified output folders
    before running new report generation.
    """

    log_message("🧹 Running pre-cleanup...")

    excel_folders = [
        EXCEL_DELETE_FOLDER_1,
        EXCEL_DELETE_FOLDER_2,
        EXCEL_DELETE_FOLDER_3
    ]

    for folder in excel_folders:

        if folder.strip() == "":
            continue

        if os.path.exists(folder):
            deleted_excel = False

            for f in os.listdir(folder):
                if f.lower().endswith((".xls", ".xlsx")):
                    os.remove(os.path.join(folder, f))
                    log_message(f"🗑 Deleted Excel from {folder}: {f}")
                    deleted_excel = True

            if not deleted_excel:
                log_message(f"ℹ No Excel files to delete in {folder}.")
        else:
            log_message(f"⚠ Folder does not exist: {folder}")

    log_message("✅ Pre-cleanup done.")


# =====================================================
#                MIS FILE UPDATE CHECK
# =====================================================

def is_mis_updated_today():
    """
    Checks whether MIS_Report.xlsx was modified today.
    """
    try:
        modified_time = datetime.datetime.fromtimestamp(
            os.path.getmtime(MIS_FILE_PATH)
        )

        today = datetime.datetime.now().date()

        log_message(f"📄 MIS Report last modified: {modified_time}")

        return modified_time.date() == today

    except Exception as e:
        log_message(f"❌ Error checking MIS report: {e}")
        notify("MIS Check Error", "Unable to read MIS file. Check logs.")
        return False


# =====================================================
#              WAIT UNTIL MIS IS UPDATED
# =====================================================

def wait_for_update():
    """
    Keeps checking until MIS is updated today.
    Once updated:
        - Performs cleanup
        - Runs all scripts
    """

    while True:

        if is_mis_updated_today():

            log_message("✅ MIS report is updated today. Proceeding...")
            notify("MIS Ready", "MIS Report is updated. Starting automation.")

            # Perform cleanup before execution
            preclean_folders()

            # Execute all report scripts
            run_all_scripts()
            break

        else:
            msg = f"MIS report not updated. Rechecking in {RECHECK_INTERVAL} mins."
            log_message(f"⚠️ {msg}")
            notify("Waiting for MIS", msg)

            time.sleep(RECHECK_INTERVAL * 60)


# =====================================================
#                SCRIPT EXECUTION ENGINE
# =====================================================

def run_all_scripts():
    """
    Sequentially executes all scripts defined in SCRIPT_PATHS.
    Supports:
        - Python scripts (.py)
        - Jupyter notebooks (.ipynb)
    """

    for script in SCRIPT_PATHS:

        script_name = os.path.basename(script)
        log_message(f"🚀 Starting {script_name}...")
        notify("Script Started", f"Running: {script_name}")

        try:

            # If Python script
            if script.endswith(".py"):
                subprocess.run(["python", script], check=True)

            # If Jupyter notebook
            elif script.endswith(".ipynb"):
                subprocess.run([
                    "jupyter", "nbconvert", "--to", "notebook",
                    "--execute", script, "--inplace"
                ], check=True)

            else:
                log_message(f"⚠️ Unsupported file: {script}")
                notify("Unsupported File", f"Cannot run file: {script_name}")
                continue

            log_message(f"✅ {script_name} completed successfully.")
            notify("Script Completed", f"{script_name} finished successfully.")

        except subprocess.CalledProcessError as e:
            log_message(f"❌ Error running {script}: {e}")
            notify("Script Failed", f"Error running: {script_name}")


# =====================================================
#                        MAIN LOOP
# =====================================================

def main():
    """
    Initializes scheduler and runs indefinitely.
    """

    log_message(f"🕒 Scheduler started. Checking daily at {CHECK_TIME}...")
    notify("Scheduler Started", f"Daily check set at {CHECK_TIME}")

    # Schedule daily execution
    schedule.every().day.at(CHECK_TIME).do(wait_for_update)

    # Continuous background loop
    while True:
        schedule.run_pending()
        time.sleep(60)


# Entry point
if __name__ == "__main__":
    main()
