
"""
Dropout Consultation Automated Email Report
--------------------------------------------
Generates a report of cancelled consultations from yesterday
and emails it via Outlook SMTP.

Author: Skanda N Raj
"""

import os
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# ================= CONFIG =================

INPUT_FILE = "data/sample_MIS_Report.xlsx"
OUTPUT_FILE = "output/Dropout_Consultations.xlsx"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

FROM_EMAIL = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

TO_EMAILS = ["recipient@example.com"]
CC_EMAILS = ["cc_recipient@example.com"]

# ==========================================


def generate_report():
    yesterday = datetime.today().date() - timedelta(days=1)

    df = pd.read_excel(INPUT_FILE, engine="openpyxl")
    df.columns = df.columns.str.strip()

    possible_date_cols = [col for col in df.columns if "date" in col.lower()]
    if not possible_date_cols:
        raise ValueError("Date column not found")

    date_col = possible_date_cols[0]
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

    df_filtered = df[
        (df["Appt. Status"].str.lower() == "cancelled") &
        (df[date_col].dt.date == yesterday)
    ].copy()

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    df_filtered.to_excel(OUTPUT_FILE, index=False)

    print("✅ Report generated successfully")


def send_email():
    yesterday = datetime.today().date() - timedelta(days=1)

    subject = f"Dropout Report - {yesterday.strftime('%d/%m/%Y')}"
    body = f"""Hi Team,

Please find attached the dropout consultations report 
for {yesterday.strftime('%d/%m/%Y')}.

Regards,
Analytics Team
"""

    msg = MIMEMultipart()
    msg["From"] = FROM_EMAIL
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(OUTPUT_FILE, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(OUTPUT_FILE)}"
    )
    msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(FROM_EMAIL, EMAIL_PASSWORD)
        server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
        server.quit()
        print("📧 Email sent successfully")

    except Exception as e:
        print("❌ Email error:", e)


def run():
    generate_report()
    send_email()


if __name__ == "__main__":
    run()
