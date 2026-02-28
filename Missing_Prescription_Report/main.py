"""
Missing Prescription Report Automation
---------------------------------------

Purpose:
Generates a report of patients who:
- Did NOT receive prescription
- Had valid paid/cash appointment
- Procedure type = Instant
- Appointment date = Yesterday
- Consider Patient = Yes
- Hospital = Aster Digital Health

Then emails the report automatically.

Author: SKANDA N RAJ
"""

import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
from dotenv import load_dotenv


# ================= ENVIRONMENT =================
load_dotenv()

# ================= CONFIG =================

# Use project-relative paths (GitHub friendly)
input_file = "data/MIS_Report.xlsx"
output_file = "output/prescription_no_yesterday.xlsx"

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

FROM_EMAIL = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

TO_EMAILS = [
    "recipient1@yourdomain.com",
    "recipient2@yourdomain.com",
    "recipient3@yourdomain.com"
]

CC_EMAILS = ["cc_recipient@yourdomain.com"]


# ================= DATE LOGIC =================
yesterday = (datetime.today() - timedelta(days=1)).date()
SUBJECT = f"Missing Prescriptions - {yesterday:%d/%m/%Y}"

BODY = """Hi Team,

This report contains patients who did not receive a prescription yesterday, despite having a valid instant paid appointment.

Best regards,
Analytics Team
"""


# ================= STEP 1: LOAD & FILTER DATA =================
df = pd.read_excel(input_file, engine="openpyxl")
df.columns = df.columns.str.strip()

print("Available columns:", df.columns.tolist())

# Normalize column names for mapping
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
print("Mapped columns:", needed)

# Ensure Appointment Date exists
if needed["appointment date"] is None:
    print("❌ Appointment Date column not found")
    raise SystemExit

df[needed["appointment date"]] = pd.to_datetime(
    df[needed["appointment date"]],
    errors="coerce"
).dt.date


# ================= APPLY FILTERS =================
filtered = df[
    (df[needed["is prescription generated"]].astype(str).str.strip().str.lower() == "no") &
    (df[needed["consider patient"]].astype(str).str.strip().str.lower() == "yes") &
    (df[needed["appt. payment status"]].astype(str).str.strip().str.lower().isin(["paid", "cash"])) &
    (df[needed["procedure type"]].astype(str).str.strip().str.lower() == "instant") &
    (df[needed["appointment date"]] == yesterday) &
    (df[needed["hospital name"]].astype(str).str.strip().str.lower() == "aster digital health")
].copy()


# ================= ADD BUSINESS COLUMNS =================
filtered["Missing Prescriptions (Yesterday)"] = "Yes"
filtered["Total"] = 1


# ================= SELECT REQUIRED COLUMNS =================
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

available_cols = [col for col in required_cols if col in filtered.columns]
final = filtered[available_cols]


# ================= APPEND SUMMARY ROW =================
if not final.empty:
    total_row = {col: "" for col in final.columns}
    total_row["Patient Name"] = "Total Patients"
    total_row["Total"] = final["Total"].sum()
    final = pd.concat([final, pd.DataFrame([total_row])], ignore_index=True)


# ================= EXPORT EXCEL =================
os.makedirs(os.path.dirname(output_file), exist_ok=True)
final.to_excel(output_file, index=False)

print(f"✅ Report generated: {output_file}")


# ================= SEND EMAIL =================
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

try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(FROM_EMAIL, EMAIL_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()

    print("📧 Email sent successfully!")

except Exception as e:
    print("❌ Error sending email:", e)
