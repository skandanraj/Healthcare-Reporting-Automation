"""
MIS Column Standardization Script
---------------------------------

Reads raw MIS report and creates a cleaned version
containing only required business columns.

Author: SKANDA N RAJ
"""

import pandas as pd
import os


# ================= CONFIG =================

# Input MIS file
input_file = r"input folder path\Dummy Dataset.xlsx"

# Output file (same folder as input)
input_folder = os.path.dirname(input_file)
output_file = os.path.join(input_folder, "Ops_Data_Sanitization.xlsx")


# ================= REQUIRED COLUMNS =================

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
    "Event Join Time Patient",
    "Event Left Time Patient",
    "Event Join Time Doctor",
    "Event Left Time Doctor"
]


# ================= STEP 1: READ MIS FILE =================

try:
    df = pd.read_excel(input_file, engine="openpyxl")
    print(f"✅ Successfully loaded {len(df)} rows from: {os.path.basename(input_file)}")
except Exception as e:
    print(f"❌ Error reading Excel file:\n{e}")
    exit()


# ================= STEP 2: FILTER REQUIRED COLUMNS =================

available_cols = [col for col in columns_to_keep if col in df.columns]
missing_cols = [col for col in columns_to_keep if col not in df.columns]

filtered_df = df[available_cols]

print(f"✅ Columns kept: {len(available_cols)}")

if missing_cols:
    print("\n⚠️ Missing columns:")
    for col in missing_cols:
        print(" -", col)


# ================= STEP 3: SAVE CLEANED FILE =================

try:
    filtered_df.to_excel(output_file, index=False)
    print(f"\n✅ Cleaned file created successfully:\n{output_file}")
except Exception as e:
    print(f"\n❌ Error saving file:\n{e}")
