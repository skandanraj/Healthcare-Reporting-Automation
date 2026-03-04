"""
Dropout Consultation Report Automation
---------------------------------------

Purpose:
This script generates a daily dropout consultation report
for selected hospitals and sends it via Outlook SMTP.

Business Logic:
1. Reads MIS Excel report
2. Filters cancelled appointments
3. Filters only yesterday's records
4. Filters only selected hospitals
6. Sends filtered report via email

Author: SKANDA N RAJ
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

# --- CONFIG ---
input_file = r"input folder path\Dummy Dataset.xlsx"
output_file_cancelled = r"output folder path\Dropout_Consultations_Karnataka.xlsx"

# Email settings
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

FROM_EMAIL = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

TO_EMAILS = [
    "recipient@domain.com"
]

CC_EMAILS = ["recipient@domain.com"]

yesterday = datetime.today().date() - timedelta(days=1)

SUBJECT = f"Yesterday's Dropout Consultations Report - {yesterday.strftime('%d/%m/%Y')}"

BODY = f"""Hi Team,

This report contains patients who reached the payment page but did not complete the payment yesterday ({yesterday.strftime('%d/%m/%Y')}).

Best regards,
Analytics Team
"""

# --- STEP 1: Process MIS Report ---
df = pd.read_excel(input_file, engine="openpyxl")
df.columns = df.columns.str.strip()

# Detect appointment date column
possible_date_cols = [col for col in df.columns if "date" in col.lower()]
if not possible_date_cols:
    print("❌ Appointment date column not found")
    print(df.columns.tolist())
    raise SystemExit

DATE_COL = possible_date_cols[0]
df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

allowed_hospitals = [
   "Aster CMI Hospital",
    "Aster RV Hospital",
    "Aster Whitefield Hospital",
]

df_c = df[
    (df["Appt. Status"].astype(str).str.strip().str.lower() == "cancelled") &
    (df[DATE_COL].dt.date == yesterday) &
    (df["Hospital Name"].astype(str).str.strip().isin(allowed_hospitals))
].copy()

# Optional: filter "Consider Patient" = Yes if column exists
if "Consider Patient" in df_c.columns:
    df_c = df_c[df_c["Consider Patient"].astype(str).str.lower().str.strip() == "yes"]

# Select required columns that exist
cols_c = ["Patient Name", "Hospital Name", "Mobile", "Doctor Name", "Speciality", DATE_COL]
cols_c_available = [col for col in cols_c if col in df_c.columns]
df_c = df_c[cols_c_available].drop_duplicates()

# Save to Excel (make folder if needed)
os.makedirs(os.path.dirname(output_file_cancelled), exist_ok=True)
df_c.to_excel(output_file_cancelled, index=False)
print(f"✅ Cancelled appointments report generated: {output_file_cancelled}")

# --- STEP 2: Send Email ---
msg = MIMEMultipart()
msg["From"] = FROM_EMAIL
msg["To"] = ", ".join(TO_EMAILS)
msg["Cc"] = ", ".join(CC_EMAILS)
msg["Subject"] = SUBJECT
msg.attach(MIMEText(BODY, "plain"))

# Attach Excel file
with open(output_file_cancelled, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())

encoders.encode_base64(part)
part.add_header(
    "Content-Disposition",
    f"attachment; filename={os.path.basename(output_file_cancelled)}"
)
msg.attach(part)

# Connect to SMTP and send
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(FROM_EMAIL, EMAIL_PASSWORD)
    server.sendmail(FROM_EMAIL, TO_EMAILS + CC_EMAILS, msg.as_string())
    server.quit()
    print("📧 Email sent successfully with the attachment!")
except Exception as e:
    print("❌ Error sending email:", e)
