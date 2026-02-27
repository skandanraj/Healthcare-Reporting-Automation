"""
Dropout Consultation Report - Jenkins Version
----------------------------------------------

This version is designed for Jenkins execution.

Features:
- UTF-8 safe console output
- Explicit exit codes for Jenkins job status
- SMTP debug logging enabled
- Environment-variable based credentials
"""

import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import sys

# Ensure UTF-8 safe printing (important for Jenkins logs)
sys.stdout.reconfigure(encoding="utf-8")


# ================= CONFIG =================

# Use relative paths for GitHub portability
input_file = "data/MIS_Report.xlsx"
output_file_cancelled = "output/Dropout_Consultations_Karnataka.xlsx"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Load credentials from environment variables (Jenkins credentials binding)
FROM_EMAIL = os.getenv("EMAIL_USER")
SMTP_PASSWORD = os.getenv("EMAIL_PASSWORD")

TO_EMAILS = [
    "recipient1@yourdomain.com",
    "recipient2@yourdomain.com",
    "recipient3@yourdomain.com"
]

CC_EMAILS = ["cc_recipient@yourdomain.com"]


# ================= DATE LOGIC =================
yesterday = datetime.today().date() - timedelta(days=1)

SUBJECT = f"Yesterday Dropout Consultations Report - {yesterday.strftime('%d/%m/%Y')}"

BODY = f"""Hi Team,

Attached is the dropout consultations report for {yesterday.strftime('%d/%m/%Y')}.

Regards,
Analytics Team
"""


# ================= PROCESS DATA =================
try:
    df = pd.read_excel(input_file, engine="openpyxl")
    df.columns = df.columns.str.strip()

    date_cols = [c for c in df.columns if "date" in c.lower()]
    if not date_cols:
        print("[ERROR] Appointment date column not found")
        sys.exit(1)

    DATE_COL = date_cols[0]
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

    allowed_hospitals = [
        "Hospital A",
        "Hospital B",
        "Hospital C",
    ]

    df_c = df[
        (df["Appt. Status"].astype(str).str.lower().str.strip() == "cancelled") &
        (df[DATE_COL].dt.date == yesterday) &
        (df["Hospital Name"].astype(str).str.strip().isin(allowed_hospitals))
    ].copy()

    if "Consider Patient" in df_c.columns:
        df_c = df_c[df_c["Consider Patient"].astype(str).str.lower() == "yes"]

    cols = ["Patient Name", "Hospital Name", "Mobile",
            "Doctor Name", "Speciality", DATE_COL]

    df_c = df_c[[c for c in cols if c in df_c.columns]].drop_duplicates()

    os.makedirs(os.path.dirname(output_file_cancelled), exist_ok=True)
    df_c.to_excel(output_file_cancelled, index=False)

    print("[OK] Excel report generated")

except Exception as e:
    print("[ERROR] Data processing failed")
    print(str(e))
    sys.exit(1)


# ================= SEND EMAIL =================
try:
    msg = MIMEMultipart()
    msg["From"] = FROM_EMAIL
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(BODY, "plain"))

    with open(output_file_cancelled, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(output_file_cancelled)}"
    )
    msg.attach(part)

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.set_debuglevel(1)  # Shows SMTP logs in Jenkins console
    server.starttls()
    server.login(FROM_EMAIL, SMTP_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()

    print("[OK] Email sent successfully")

except Exception as e:
    print("[ERROR] Email sending failed")
    print(str(e))
    sys.exit(1)
