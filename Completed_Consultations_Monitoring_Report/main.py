#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Completed Consultations - Last 15 Days (All Units)

This script:
- Extracts completed appointments (Status = Done)
- Filters last 15 days window
- Keeps Consider Patient = Yes (if exists)
- Avoids re-sending already emailed records
- Maintains persistent sent-log state
- Sends only NEW rows via email

Author: Your Name
"""

import os
import hashlib
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv


# ================= ENVIRONMENT =================
load_dotenv()

# ================= CONFIG =================

INPUT_FILE = "data/MIS_Report.xlsx"

OUTPUT_DIR = "output/last_15_days"
os.makedirs(OUTPUT_DIR, exist_ok=True)

OUTPUT_FILE = os.path.join(
    OUTPUT_DIR,
    "Completed_Consultations_Last15Days_AllUnits.xlsx"
)

STATE_DIR = os.path.join(OUTPUT_DIR, "state")
os.makedirs(STATE_DIR, exist_ok=True)

STATE_FILE = os.path.join(STATE_DIR, "sent_completed_keys.csv")

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

FROM_EMAIL = os.getenv("EMAIL_USER")
SMTP_PASSWORD = os.getenv("EMAIL_PASSWORD")

TO_EMAILS = [
    "recipient1@yourdomain.com",
    "recipient2@yourdomain.com"
]

CC_EMAILS = ["cc_recipient@yourdomain.com"]

today = datetime.today().date()
end_date = today - timedelta(days=1)
start_date = end_date - timedelta(days=14)

SUBJECT = (
    f"Completed Consultations (Last 15 Days) — "
    f"{start_date:%d/%m/%Y} to {end_date:%d/%m/%Y}"
)

BODY = f"""Hi Team,

Please find attached the completed consultations (Status = Done)
for the last 15 days ({start_date:%d/%m/%Y} to {end_date:%d/%m/%Y}).

Best regards,
Analytics Team
"""


# ================= MAIL HELPER =================
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

    if attachment_path and os.path.exists(attachment_path):
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
    try:
        server.starttls()
        server.login(from_email, password)
        server.sendmail(from_email, recipients, msg.as_string())
    finally:
        server.quit()


# ================= HELPERS =================
def mk_row_hash(*values) -> str:
    normed = []
    for v in values:
        s = "" if v is None else str(v)
        s = " ".join(s.strip().lower().split())
        normed.append(s)
    joined = "|".join(normed)
    return hashlib.md5(joined.encode("utf-8")).hexdigest()


def load_sent_keys(path: str) -> set:
    if not os.path.exists(path):
        return set()
    try:
        dfk = pd.read_csv(path, dtype=str)
        return set(dfk["key"].dropna()) if "key" in dfk.columns else set()
    except Exception:
        return set()


def save_append_keys(path: str, keys: list):
    df = pd.DataFrame({"key": keys})
    df.to_csv(path, mode="a", index=False, header=not os.path.exists(path))


# ================= READ MIS =================
try:
    df = pd.read_excel(INPUT_FILE, engine="openpyxl")
except Exception as e:
    raise SystemExit(f"❌ Could not read MIS workbook: {e}")

df.columns = df.columns.map(lambda x: str(x).strip())

required_cols = [
    "Patient Name",
    "Mobile",
    "UHID",
    "Doctor Name",
    "Speciality",
    "Hospital Name",
    "Appt. Status",
    "Appointment Date",
]

missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise SystemExit(f"❌ Missing required columns: {missing}")

# ================= FILTER =================
df["Appointment Date"] = pd.to_datetime(
    df["Appointment Date"],
    errors="coerce"
)

df["__date_only"] = df["Appointment Date"].dt.date

mask = (
    (df["Appt. Status"].str.lower().str.strip() == "done") &
    (df["__date_only"] >= start_date) &
    (df["__date_only"] <= end_date)
)

df_f = df.loc[mask].copy()

if "Consider Patient" in df_f.columns:
    df_f = df_f[
        df_f["Consider Patient"].astype(str)
        .str.lower().str.strip() == "yes"
    ]

out = df_f[[
    "Patient Name",
    "Mobile",
    "UHID",
    "Appointment Date",
    "Doctor Name",
    "Speciality",
    "Hospital Name"
]].copy()

# ================= DEDUP =================
row_keys = out.apply(
    lambda r: mk_row_hash(
        r["Patient Name"],
        r["UHID"],
        r["Doctor Name"],
        r["Hospital Name"],
        r["Appointment Date"]
    ),
    axis=1
)

out["__key"] = row_keys

sent_keys = load_sent_keys(STATE_FILE)
out_new = out[~out["__key"].isin(sent_keys)].drop_duplicates()

if out_new.empty:
    print("✅ No new completed consultations to send.")
    raise SystemExit(0)

out_new.drop(columns="__key").to_excel(
    OUTPUT_FILE,
    index=False
)

print(f"✅ New rows to send: {len(out_new)}")

send_mail_with_attachment(
    SMTP_SERVER,
    SMTP_PORT,
    FROM_EMAIL,
    SMTP_PASSWORD,
    TO_EMAILS,
    CC_EMAILS,
    SUBJECT + f" | New rows: {len(out_new)}",
    BODY,
    OUTPUT_FILE
)

save_append_keys(
    STATE_FILE,
    out_new["__key"].astype(str).tolist()
)

print("📧 Email sent and state updated successfully.")
