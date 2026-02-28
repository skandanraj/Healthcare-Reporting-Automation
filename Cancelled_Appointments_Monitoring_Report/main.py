"""
Cancelled Appointments Monitoring Automation
--------------------------------------------

This script generates two reports:

1. Cancelled & Paid Appointments (Yesterday)
2. Cancelled Appointments (Yesterday + Today)

Both reports are emailed automatically as attachments.

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
# Load email credentials from .env file
load_dotenv()


# ================= CONFIGURATION =================

# Input MIS file
input_file = r"data/MIS_Report.xlsx"

# Output report file paths
output_file_cancelled_paid = r"output/cancelled_paid_yesterday.xlsx"
output_file_cancelled = r"output/cancelled_patients.xlsx"

# SMTP Configuration
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Load credentials securely from environment variables
FROM_EMAIL = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Email Recipients (generic for public repo)
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


# ================= STEP 1: LOAD MIS DATA =================

df = pd.read_excel(input_file, engine="openpyxl")

# Clean column names
df.columns = df.columns.str.strip()

# Detect appointment date column dynamically
possible_date_cols = [col for col in df.columns if "date" in col.lower()]

if not possible_date_cols:
    print("❌ Appointment date column not found")
    print(df.columns.tolist())
    raise SystemExit

DATE_COL = possible_date_cols[0]
print(f"✅ Using column '{DATE_COL}' as appointment date")

# Convert to datetime
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors='coerce')


# ================= BUSINESS RULE: HOSPITAL FILTER =================

# Generic hospital names for public repository
allowed_hospitals = [
    "Hospital Name 1",
    "Hospital Name 2",
    "Hospital Name 3",
    "Hospital Name 4",
    "Hospital Name 5"
]


# ================= REPORT 1: CANCELLED & PAID (YESTERDAY) =================

yesterday = datetime.today().date() - timedelta(days=1)

# Filter yesterday data
df_cp = df[df[DATE_COL].dt.date == yesterday]

# Optional filter: Consider Patient = Yes
if "Consider Patient" in df_cp.columns:
    df_cp = df_cp[
        df_cp['Consider Patient']
        .astype(str).str.lower().str.strip() == "yes"
    ]

# Apply filters
cancelled_paid = df_cp[
    (df_cp['Appt. Status'].astype(str).str.strip().str.lower() == "cancelled") &
    (df_cp['Appt. Payment Status'].astype(str).str.strip().str.lower() == "paid") &
    (df_cp['Hospital Name'].astype(str).str.strip().isin(allowed_hospitals))
].copy()

# Select required columns
cols_cp = [
    'Patient Name',
    'Hospital Name',
    'Mobile',
    'Doctor Name',
    'Speciality',
    DATE_COL,
    'Appt. Status',
    'Appt. Payment Status'
]

cols_cp_available = [col for col in cols_cp if col in cancelled_paid.columns]
cancelled_paid = cancelled_paid[cols_cp_available].drop_duplicates()

# Save report
os.makedirs(os.path.dirname(output_file_cancelled_paid), exist_ok=True)
cancelled_paid.to_excel(output_file_cancelled_paid, index=False)

print(f"✅ Cancelled & Paid report generated: {output_file_cancelled_paid}")


# ================= REPORT 2: CANCELLED (YESTERDAY + TODAY) =================

today = datetime.today().date()

df_c = df[
    (df['Appt. Status'].astype(str).str.strip().str.lower() == "cancelled") &
    (df[DATE_COL].dt.date.isin([yesterday, today])) &
    (df['Hospital Name'].astype(str).str.strip().isin(allowed_hospitals))
].copy()

# Optional column check
if "Patient" in df_c.columns:
    df_c = df_c[
        df_c['Patient']
        .astype(str).str.lower().str.strip() == "yes"
    ]

cols_c = [
    'Patient Name',
    'Hospital Name',
    'Mobile',
    'Doctor Name',
    'Speciality',
    DATE_COL
]

cols_c_available = [col for col in cols_c if col in df_c.columns]
df_c = df_c[cols_c_available].drop_duplicates()

# Save report
os.makedirs(os.path.dirname(output_file_cancelled), exist_ok=True)
df_c.to_excel(output_file_cancelled, index=False)

print(f"✅ Cancelled appointments report generated: {output_file_cancelled}")


# ================= STEP 2: SEND EMAIL =================

msg = MIMEMultipart()
msg["From"] = FROM_EMAIL
msg["To"] = ", ".join(TO_EMAILS)
msg["Cc"] = ", ".join(CC_EMAILS)
msg["Subject"] = SUBJECT
msg.attach(MIMEText(BODY, "plain"))

# Attach first report
with open(output_file_cancelled_paid, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header(
    "Content-Disposition",
    f"attachment; filename={os.path.basename(output_file_cancelled_paid)}"
)
msg.attach(part)

# Attach second report
with open(output_file_cancelled, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header(
    "Content-Disposition",
    f"attachment; filename={os.path.basename(output_file_cancelled)}"
)
msg.attach(part)

# Send email
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(FROM_EMAIL, EMAIL_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()
    print("📧 Email sent successfully with both attachments!")
except Exception as e:
    print("❌ Error sending email:", e)
