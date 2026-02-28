"""
Missing Prescription Report - Jenkins Version
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
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import sys

# Ensure Jenkins-safe console output
sys.stdout.reconfigure(encoding="utf-8")

# ================= CONFIG =================

# Use project-relative paths (Jenkins friendly)
input_file = "data/MIS_Report.xlsx"

output_file_cancelled_paid = "output/cancelled_paid_yesterday.xlsx"
output_file_cancelled = "output/cancelled_patients.xlsx"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Load credentials from Jenkins Environment Variables
FROM_EMAIL = os.getenv("EMAIL_USER")
SMTP_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Generic recipients for public repo
TO_EMAILS = ["recipient@yourdomain.com"]
CC_EMAILS = ["cc_recipient@yourdomain.com"]

SUBJECT = "Appointments Report"

BODY = """Hi Team,

Please find attached the latest reports:

1. Cancelled & Paid appointments from yesterday.
2. Cancelled appointments (yesterday and today).

Best regards,
Analytics Team
"""

# ================= LOAD MIS =================
try:
    df = pd.read_excel(input_file, engine="openpyxl")
except Exception as e:
    print("[ERROR] Failed to read MIS file:", e)
    sys.exit(1)

df.columns = df.columns.str.strip()

# Detect appointment date column
possible_date_cols = [c for c in df.columns if "date" in c.lower()]
if not possible_date_cols:
    print("[ERROR] Appointment date column not found")
    print(df.columns.tolist())
    sys.exit(1)

DATE_COL = possible_date_cols[0]
print("[INFO] Using appointment date column:", DATE_COL)

df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

# ================= FILTER DATA =================

allowed_hospitals = [
    "Hospital Name 1",
    "Hospital Name 2",
    "Hospital Name 3",
    "Hospital Name 4",
    "Hospital Name 5"
]

yesterday = datetime.today().date() - timedelta(days=1)
today = datetime.today().date()

# -------- Cancelled & Paid (Yesterday) --------
df_cp = df[df[DATE_COL].dt.date == yesterday]

if "Consider Patient" in df_cp.columns:
    df_cp = df_cp[
        df_cp["Consider Patient"]
        .astype(str).str.lower().str.strip() == "yes"
    ]

cancelled_paid = df_cp[
    (df_cp["Appt. Status"].astype(str).str.lower().str.strip() == "cancelled") &
    (df_cp["Appt. Payment Status"].astype(str).str.lower().str.strip() == "paid") &
    (df_cp["Hospital Name"].astype(str).str.strip().isin(allowed_hospitals))
].copy()

cols_cp = [
    "Patient Name", "Hospital Name", "Mobile",
    "Doctor Name", "Speciality", DATE_COL,
    "Appt. Status", "Appt. Payment Status"
]

cancelled_paid = cancelled_paid[
    [c for c in cols_cp if c in cancelled_paid.columns]
].drop_duplicates()

os.makedirs(os.path.dirname(output_file_cancelled_paid), exist_ok=True)
cancelled_paid.to_excel(output_file_cancelled_paid, index=False)

print("[OK] Cancelled & Paid report generated:", output_file_cancelled_paid)

# -------- Cancelled (Yesterday + Today) --------
df_c = df[
    (df["Appt. Status"].astype(str).str.lower().str.strip() == "cancelled") &
    (df[DATE_COL].dt.date.isin([yesterday, today])) &
    (df["Hospital Name"].astype(str).str.strip().isin(allowed_hospitals))
].copy()

if "Patient" in df_c.columns:
    df_c = df_c[
        df_c["Patient"]
        .astype(str).str.lower().str.strip() == "yes"
    ]

cols_c = [
    "Patient Name", "Hospital Name",
    "Mobile", "Doctor Name",
    "Speciality", DATE_COL
]

df_c = df_c[[c for c in cols_c if c in df_c.columns]].drop_duplicates()

os.makedirs(os.path.dirname(output_file_cancelled), exist_ok=True)
df_c.to_excel(output_file_cancelled, index=False)

print("[OK] Cancelled appointments report generated:", output_file_cancelled)

# ================= SEND EMAIL =================
try:
    msg = MIMEMultipart()
    msg["From"] = FROM_EMAIL
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(BODY, "plain"))

    for path in [output_file_cancelled_paid, output_file_cancelled]:
        with open(path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(path)}"
        )
        msg.attach(part)

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.set_debuglevel(1)  # Visible in Jenkins logs
    server.starttls()
    server.login(FROM_EMAIL, SMTP_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()

    print("[OK] Email sent successfully with both attachments")

except Exception as e:
    print("[ERROR] Email sending failed")
    print(str(e))
    sys.exit(1)
