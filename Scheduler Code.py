import os
import datetime
import time
import subprocess
import schedule
from win10toast import ToastNotifier


# ============ CONFIG ==============

# Workspace-relative MIS file
MIS_FILE_PATH = "data/MIS_Report.xlsx"

# Scripts inside your repo
SCRIPT_PATHS = [
    "Cancelled_Appointments_Monitoring_Report/main.py",
    "Completed_Consultations_Monitoring_Report/main.py",
    "Dropout_Consultation_Report/main.py",
    "Missing_Prescription_Report/main.py",
    "Ops_Data_Sanitization/main.py"
]

CHECK_TIME = "10:30"
RECHECK_INTERVAL = 30  # minutes
LOG_DIR = "logs"

# Excel cleanup folders (workspace relative)
EXCEL_DELETE_FOLDER_1 = "Dropout_Consultation_Report/output"
EXCEL_DELETE_FOLDER_2 = "Completed_Consultations_Monitoring_Report/output"
EXCEL_DELETE_FOLDER_3 = "Missing_Prescription_Report/output"

# ------------------------------------

# -------- Notification Setup --------
notifier = ToastNotifier()

def notify(title, msg):
    try:
        notifier.show_toast(title, msg, duration=10, threaded=True)
    except:
        pass  # fail silently if notifications unavailable
# ====================================


def get_log_file():
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    return os.path.join(LOG_DIR, f"scheduler_log_{today}.txt")


def log_message(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    print(log_entry.strip())
    with open(get_log_file(), "a", encoding="utf-8") as f:
        f.write(log_entry)


# =====================================================
#                PRE-CLEANUP
# =====================================================

def preclean_folders():
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


def is_mis_updated_today():
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


def wait_for_update():

    while True:
        if is_mis_updated_today():

            log_message("✅ MIS report is updated today. Proceeding...")
            notify("MIS Ready", "MIS Report is updated. Starting automation.")

            # Pre-cleanup before running scripts
            preclean_folders()

            run_all_scripts()
            break

        else:
            msg = f"MIS report not updated. Rechecking in {RECHECK_INTERVAL} mins."
            log_message(f"⚠️ {msg}")
            notify("Waiting for MIS", msg)
            time.sleep(RECHECK_INTERVAL * 60)


def run_all_scripts():

    for script in SCRIPT_PATHS:

        script_name = os.path.basename(script)
        log_message(f"🚀 Starting {script_name}...")
        notify("Script Started", f"Running: {script_name}")

        try:
            if script.endswith(".py"):
                subprocess.run(["python", script], check=True)

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


def main():

    log_message(f"🕒 Scheduler started. Checking daily at {CHECK_TIME}...")
    notify("Scheduler Started", f"Daily check set at {CHECK_TIME}")

    schedule.every().day.at(CHECK_TIME).do(wait_for_update)

    while True:
        schedule.run_pending()
        time.sleep(60)


if __name__ == "__main__":
    main()
