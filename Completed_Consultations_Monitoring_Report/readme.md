# 📊 Completed Consultations – Last 15 Days Automation

## 🧠 Overview

This script automates the generation and distribution of a **Completed Consultations Report (Last 15 Days)**.

The report identifies appointments where:

- Appointment Status = **Done**
- Appointment Date falls within the **last 15 days**
- (Optional) "Consider Patient" = **Yes**

The script ensures that:

- Only **new records** are sent each time
- Previously emailed rows are not re-sent
- A persistent state file tracks sent records

After filtering, the script:

1. Generates an Excel report containing only new rows  
2. Maintains a sent-log file for cross-run deduplication  
3. Sends the report automatically via Outlook SMTP  
4. Updates the sent-log after successful email  

---

## ⚙️ Business Logic

The script applies the following logic:

### 1️⃣ Status Filter
Only records where:
```
Appt. Status = "done"
```

---

### 2️⃣ Date Window Filter
Only records where:
```
Appointment Date BETWEEN (Today - 15 days) AND Yesterday
```

The script dynamically calculates:

```python
today = datetime.today().date()
end_date = today - timedelta(days=1)
start_date = end_date - timedelta(days=14)
```

---

### 3️⃣ Optional Consider Patient Filter
If column exists:
```
Consider Patient = "Yes"
```

---

### 4️⃣ Cross-Run Deduplication

The script generates a unique hash key per row using:

- Patient Name
- UHID
- Doctor Name
- Unit
- Date of Completed Appointment

These keys are stored in a persistent state file:

```
output/last_15_days/state/sent_completed_keys.csv
```

Before sending:

- Previously sent keys are loaded
- Already-sent rows are excluded
- Only new rows are emailed

---

## 🛠 Tech Stack

- Python
- Pandas
- OpenPyXL
- SMTP (Office365)
- hashlib (for deduplication)
- python-dotenv (for secure credentials)

---

## 📂 Project Structure

```
completed_consultations_15days/
│
├── main.py
├── README.md
├── requirements.txt
├── state
├── completed_consultations_15days.xlsx
```

---

## 🔐 Environment Setup

Create a `.env` file in the project directory:

```
EMAIL_USER=your_email@yourdomain.com
EMAIL_PASSWORD=your_password
```

⚠️ Do NOT commit `.env` to GitHub.

---

## 📦 Install Dependencies

Make sure Python 3.8+ is installed.

Install required libraries:

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pandas openpyxl python-dotenv
```

---

## ▶️ How to Run

1. Ensure MIS file is placed inside:
```
data/MIS_Report.xlsx
```

2. Run the script:

```bash
python main.py
```

---

## 📧 What Happens When You Run It

- MIS report is loaded
- Completed appointments from last 15 days are filtered
- Previously sent records are removed
- Excel file is generated with only new rows
- Email is sent with attachment
- Sent-log is updated
- Console shows success or status message

If no new rows exist:
- No email is sent
- Script exits gracefully

---

## 🚀 Business Impact

- Provides rolling visibility of completed consultations
- Prevents duplicate email reporting
- Maintains clean communication workflow
- Reduces operational noise
- Enables consistent monitoring across units
- Demonstrates stateful automation pipeline
