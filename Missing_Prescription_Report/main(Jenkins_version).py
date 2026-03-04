"""
Missing Prescription Report - Jenkins Version
----------------------------------------------

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

# Ensure UTF-8 safe printing
sys.stdout.reconfigure(encoding="utf-8")

# --- CONFIG ---

# Jenkins / GitHub compatible paths
input_file = r"input folder path\Dummy Dataset.xlsx"
output_file = r"output folder path\prescription_no_yesterday.xlsx"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Load credentials from environment variables
FROM_EMAIL = os.getenv("EMAIL_USER")
SMTP_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Email recipients (placeholders)
TO_EMAILS = [
    "recipient@domain.com"
]

CC_EMAILS = [
    "recipient@domain.com"
]

yesterday = datetime.today().date() - timedelta(days=1)

SUBJECT = f"Missing Prescriptions - {yesterday.strftime('%d/%m/%Y')}"

BODY = f"""Hi Team,

This report contains patients who did not receive a prescription yesterday,
despite having a valid instant paid appointment.

Date: {yesterday.strftime('%d/%m/%Y')}

Best regards,
BA Team
"""

# --- LOAD DATA ---
df = pd.read_excel(input_file, engine="openpyxl")
df.columns = df.columns.str.strip()

print("[INFO] Columns found in MIS:", df.columns.tolist())

# Normalize column names
col_map = {c.lower(): c for c in df.columns}

needed_keys = [
    "is prescription generated",
    "consider patient",
    "appt. payment status",
    "procedure type",
    "appointment date",
    "hospital name",
    "mobile"
]

needed = {key: col_map.get(key) for key in needed_keys}
print("[INFO] Column mapping:", needed)

if needed["appointment date"] is None:
    print("[ERROR] Appointment Date column not found")
    sys.exit(1)

df[needed["appointment date"]] = pd.to_datetime(
    df[needed["appointment date"]], errors="coerce"
).dt.date

# --- FILTER DATA ---
filtered = df[
    (df[needed["is prescription generated"]].astype(str).str.lower().str.strip() == "no") &
    (df[needed["consider patient"]].astype(str).str.lower().str.strip() == "yes") &
    (df[needed["appt. payment status"]].astype(str).str.lower().str.strip().isin(["paid", "cash"])) &
    (df[needed["procedure type"]].astype(str).str.lower().str.strip() == "instant") &
    (df[needed["appointment date"]] == yesterday) &
    (df[needed["hospital name"]].astype(str).str.lower().str.strip() == "aster digital health")
].copy()

filtered["Missing Prescriptions (Yesterday)"] = "Yes"
filtered["Total"] = 1

required_cols = [
    needed["appointment date"],
    "Appointment Time",
    "UHID",
    "Patient Name",
    "Doctor Name",
    needed["mobile"],
    "Missing Prescriptions (Yesterday)",
    "Total"
]

final = filtered[[c for c in required_cols if c in filtered.columns]]

# Add total row
if not final.empty:
    total_row = {col: "" for col in final.columns}
    total_row["Patient Name"] = "Total Patients"
    total_row["Total"] = final["Total"].sum()
    final = pd.concat([final, pd.DataFrame([total_row])], ignore_index=True)

os.makedirs(os.path.dirname(output_file), exist_ok=True)
final.to_excel(output_file, index=False)

print("[OK] Excel report generated")

# --- SEND EMAIL ---
try:
    msg = MIMEMultipart()
    msg["From"] = FROM_EMAIL
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(BODY, "plain"))

    with open(output_file, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(output_file)}"
    )

    msg.attach(part)

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.set_debuglevel(1)   # shows SMTP conversation in Jenkins logs
    server.starttls()

    server.login(FROM_EMAIL, SMTP_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()

    print("[OK] Email sent successfully")

except Exception as e:
    print("[ERROR] Email sending failed")
    print(str(e))
    sys.exit(1)
