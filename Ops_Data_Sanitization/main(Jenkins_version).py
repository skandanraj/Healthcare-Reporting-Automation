"""
Ops Data Sanitization- Jenkins Version
----------------------------------------------

Designed for CI/CD automation using Jenkins.

Features:
- UTF-8 safe console logging
- Exit codes for job monitoring
- SMTP debug logging in Jenkins console
- Environment variable based credentials
"""


#!/usr/bin/env python3

import pandas as pd
import os
import sys

# Ensure Jenkins-safe console output
sys.stdout.reconfigure(encoding="utf-8")

# ================= INPUT / OUTPUT =================
# Use workspace-relative paths (Jenkins friendly)

input_file = "data/MIS_Report.xlsx"

output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

output_file = os.path.join(output_folder, "MIS_Report_Operational_View.xlsx")


# ================= COLUMNS TO KEEP =================
columns_to_keep = [
    "UHID",
    "Patient Name",
    "Appointment Type",
    "Procedure Type",
    "Appointment Date",
    "Appointment Time",
    "Appointment End Time",
    "Hospital Name",
    "Doctor Name",
    "Doctor HIS ID",
    "Appt. Payment Status",
    "Appt. Status",
    "Booking Source",
    "Booked DateTime",
    "booked_time",
    "Doctor ID",
    "Consultation DateTime",
    "Completed DateTime",
    "Cancelled Datetime",
    "Is Re Scheduled",
    "HIS Invoice No.",
    "Invoice No",
    "Amount (₹)",
    "Registration Fee (₹)",
    "Consult Fee (₹)",
    "Payment Type",
    "Payment Reference No.",
    "Refund Amount (₹)",
    "Room ID",
    "Is Prescription Generated",
    "Prescription Generated DateTime",
    "Waiting Time Patient",
    "Event Join Time Patient",
    "Event Left Time Patient",
    "Event Join Time Doctor",
    "Event Left Time Doctor"
]


# ================= LOAD EXCEL =================
try:
    df = pd.read_excel(input_file, engine="openpyxl")
    print("[OK] Loaded rows:", len(df))
except Exception as e:
    print("[ERROR] Failed to read MIS Excel file")
    print(str(e))
    sys.exit(1)


# ================= FILTER COLUMNS =================
available_cols = [c for c in columns_to_keep if c in df.columns]
missing_cols = [c for c in columns_to_keep if c not in df.columns]

filtered_df = df[available_cols]

if missing_cols:
    print("[WARN] Missing columns (not present in MIS):")
    for col in missing_cols:
        print(" -", col)


# ================= SAVE OUTPUT =================
try:
    filtered_df.to_excel(output_file, index=False)
    print("[OK] Cleaned file created:", output_file)
except Exception as e:
    print("[ERROR] Failed to save output Excel")
    print(str(e))
    sys.exit(1)
