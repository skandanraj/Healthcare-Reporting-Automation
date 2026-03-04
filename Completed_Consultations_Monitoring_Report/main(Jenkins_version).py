"""
Completed Consultation Report - Jenkins Version
----------------------------------------------

Designed for CI/CD automation using Jenkins.

Features:
- UTF-8 safe console logging
- Exit codes for job monitoring
- SMTP debug logging in Jenkins console
- Environment variable based credentials
"""

#!/usr/bin/env python3

import os
import hashlib
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import sys

# Ensure Jenkins-safe console output
sys.stdout.reconfigure(encoding="utf-8")

# ===================== MAIL HELPER =====================
def send_mail_with_attachment(
    smtp_server, smtp_port, from_email, password,
    to_emails, cc_emails, subject, body, attachment_path
):
    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = ", ".join(to_emails)
    msg["Cc"] = ", ".join(cc_emails) if cc_emails else ""
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachment_path:
        if not os.path.exists(attachment_path):
            raise FileNotFoundError(f"Attachment not found: {attachment_path}")
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(attachment_path)}"'
        )
        msg.attach(part)

    recipients = to_emails + (cc_emails or [])

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.set_debuglevel(1)
    try:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, password)
        server.sendmail(from_email, recipients, msg.as_string())
    finally:
        server.quit()

# ===================== CONFIG =====================

INPUT_FILE = r"input folder path\Dummy Dataset.xlsx"

OUTPUT_DIR = r"output folder path"
os.makedirs(OUTPUT_DIR, exist_ok=True)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "Completed_Consultations_Last15Days_AllUnits.xlsx"
)

STATE_DIR = os.path.join(OUTPUT_DIR, "state")
os.makedirs(STATE_DIR, exist_ok=True)

STATE_FILE = os.path.join(STATE_DIR, "sent_completed_keys.csv")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Credentials from environment variables
FROM_EMAIL = os.getenv("EMAIL_USER")
SMTP_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Placeholder emails
TO_EMAILS = [
    "recepitent@domain.com"
]

CC_EMAILS = [
    "recipient@domain.com"
]

today = datetime.today().date()
end_date = today - timedelta(days=1)
start_date = end_date - timedelta(days=14)

SUBJECT = (
    f"Completed Consultations (Last 15 Days) "
    f"- {start_date:%d/%m/%Y} to {end_date:%d/%m/%Y}"
)

BODY = f"""Hi Team,

Please find attached the completed consultations (Appt. Status = done)
for the last 15 days ({start_date:%d/%m/%Y} to {end_date:%d/%m/%Y})
across all units.

Columns:
Patient Name, Contact Number, UHID,
Date of Completed Appointment, Doctor Name,
Speciality, Unit.

Best regards,
BA Team
Aster Digital Health
"""

# ===================== HELPERS =====================
def first_existing(candidates, cols):
    for c in candidates:
        if c in cols:
            return c
    return None

def to_date(series):
    s = pd.to_datetime(series, errors="coerce")
    return s.dt.date

def load_sent_keys(path):
    if not os.path.exists(path):
        return set()
    try:
        dfk = pd.read_csv(path, dtype=str)
        if "key" in dfk.columns:
            return set(dfk["key"].dropna())
    except Exception:
        pass
    return set()

def save_append_keys(path, keys):
    df = pd.DataFrame({"key": keys})
    df.to_csv(path, mode="a", index=False, header=not os.path.exists(path))

def mk_row_hash(*values):
    normed = []
    for v in values:
        s = "" if v is None else str(v)
        s = " ".join(s.strip().lower().split())
        normed.append(s)
    return hashlib.md5("|".join(normed).encode("utf-8")).hexdigest()

# ===================== LOAD MIS =====================
try:
    try:
        df = pd.read_excel(INPUT_FILE, sheet_name="Export", engine="openpyxl")
    except Exception:
        df = pd.read_excel(INPUT_FILE, engine="openpyxl")
except Exception as e:
    print("[ERROR] Could not read MIS file:", e)
    sys.exit(0)

df.columns = df.columns.map(lambda x: str(x).strip())

# Column mapping
col_patient = first_existing(["Patient Name"], df.columns)
col_mobile = first_existing(["Mobile", "Contact Number", "Phone"], df.columns)
col_uhid = first_existing(["UHID", "Uhid"], df.columns)
col_doctor = first_existing(["Doctor Name"], df.columns)
col_spec = first_existing(["Speciality", "Specialty"], df.columns)
col_unit = first_existing(["Hospital Name", "Unit"], df.columns)
col_status = first_existing(["Appt. Status", "Appointment Status"], df.columns)
col_appt_date = first_existing(["Appointment Date", "Appt Date"], df.columns)
col_completed_dt = first_existing(["Completed DateTime"], df.columns)
col_appt_id = first_existing(["Appointment ID"], df.columns)

required = [
    col_patient, col_mobile, col_uhid,
    col_doctor, col_spec, col_unit,
    col_status, col_appt_date
]

if any(c is None for c in required):
    print("[ERROR] Missing required columns")
    sys.exit(0)

df[col_appt_date] = pd.to_datetime(df[col_appt_date], errors="coerce")
df["__appt_date_only"] = df[col_appt_date].dt.date

mask = (
    (df[col_status].astype(str).str.lower().str.strip() == "done") &
    (df["__appt_date_only"] >= start_date) &
    (df["__appt_date_only"] <= end_date)
)

df_f = df.loc[mask].copy()

if "Consider Patient" in df_f.columns:
    df_f = df_f[df_f["Consider Patient"].astype(str).str.lower().str.strip() == "yes"]

done_date = (
    to_date(df_f[col_completed_dt]).fillna(to_date(df_f[col_appt_date]))
    if col_completed_dt in df_f.columns
    else to_date(df_f[col_appt_date])
)

out = pd.DataFrame({
    "Patient Name": df_f[col_patient],
    "Contact Number": df_f[col_mobile],
    "UHID": df_f[col_uhid],
    "Date of Completed Appointment": done_date,
    "Doctor Name": df_f[col_doctor],
    "Speciality": df_f[col_spec],
    "Unit": df_f[col_unit],
})

# Dedup logic
if col_appt_id and col_appt_id in df_f.columns:
    out["__key"] = df_f[col_appt_id].astype(str).apply(mk_row_hash)
else:
    out["__key"] = out.apply(
        lambda r: mk_row_hash(
            r["Patient Name"], r["UHID"],
            r["Doctor Name"], r["Unit"],
            r["Date of Completed Appointment"]
        ),
        axis=1
    )

sent_keys = load_sent_keys(STATE_FILE)
out_new = out[~out["__key"].isin(sent_keys)].drop_duplicates()

if out_new.empty:
    print("[INFO] No new completed consultations to send")
    sys.exit(0)

out_new.drop(columns="__key").to_excel(
    OUTPUT_FILE,
    index=False,
    sheet_name="Completed_Last15Days"
)

print("[OK] Excel generated:", OUTPUT_FILE)

send_mail_with_attachment(
    SMTP_SERVER, SMTP_PORT,
    FROM_EMAIL, SMTP_PASSWORD,
    TO_EMAILS, CC_EMAILS,
    SUBJECT + f" | New rows: {len(out_new)}",
    BODY,
    OUTPUT_FILE
)

print("[OK] Email sent")

save_append_keys(STATE_FILE, out_new["__key"].tolist())
print("[OK] Sent-log updated:", STATE_FILE)
