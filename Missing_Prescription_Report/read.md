# 📊 Missing Prescription Report Automation

## 🧠 Overview

This script automates the generation and distribution of a daily **Missing Prescription Report**.

The report identifies doctors and appointments where:

- A valid paid or cash appointment was completed
- Procedure type was **Instant**
- Prescription was **NOT generated**
- Appointment date was **Yesterday**
- "Consider Patient" = Yes
- Hospital matches the configured hospital

After filtering, the script:

1. Generates a structured Excel report  
2. Adds a summary row with total affected patients  
3. Sends the report automatically via Outlook SMTP  
4. Emails it to configured stakeholders  

---

## 🎯 Business Objective

This report helps maintain a healthy digital healthcare ecosystem.

It enables the organization to:

- Monitor doctors who may not be using the prescription module properly
- Identify cases where prescriptions might be shared via WhatsApp or other external platforms
- Prevent leakage outside the application
- Ensure patients receive prescriptions inside the app
- Validate whether the product is functioning correctly

If repeated cases occur, it may indicate:

- Product usability issues
- Workflow friction
- Technical bugs
- Adoption gaps among doctors

This report acts as an early monitoring and quality control mechanism.

---

## ⚙️ Business Logic

The script applies the following filters:

### 1️⃣ Prescription Generated Filter
```
Is Prescription Generated = "No"
```

### 2️⃣ Consider Patient Filter
```
Consider Patient = "Yes"
```

### 3️⃣ Payment Status Filter
```
Appt. Payment Status IN ("Paid", "Cash")
```

Ensures only valid revenue-linked appointments are considered.

### 4️⃣ Procedure Type Filter
```
Procedure Type = "Instant"
```

### 5️⃣ Date Filter
```
Appointment Date = Yesterday
```

Dynamically calculated using:

```python
(datetime.today() - timedelta(days=1)).date()
```

### 6️⃣ Hospital Filter
```
Hospital Name = Target Hospital
```

---

## ➕ Additional Enhancements in Report

- Adds column: **Missing Prescriptions (Yesterday) = Yes**
- Adds a **Total column**
- Appends a summary row showing total number of affected patients

---

## 🛠 Tech Stack

- Python
- Pandas
- OpenPyXL
- SMTP (Office365)
- python-dotenv

---

## 📂 Project Structure

```
missing_prescription_report/
│
├── main.py
├── README.md
├── requirements.txt
└── .env (not committed to GitHub)
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

1. Update input file path in the script:
```
input_file = r"path_to_your_MIS_Report.xlsx"
```

2. Update output file path:
```
output_file = r"path_where_you_want_output.xlsx"
```

3. Run the script:

```bash
python main.py
```

---

## 📧 What Happens When You Run It

- MIS report is loaded
- Required columns are dynamically mapped
- Business filters are applied
- Summary row is appended
- Excel file is generated
- Email is sent with attachment
- Console displays success or error logs

---

## 🚀 Business Impact

- Improves doctor compliance
- Strengthens in-app workflow adoption
- Protects platform ecosystem from external leakage
- Detects potential product issues early
- Supports data-driven product improvement decisions
- Ensures revenue-linked appointments are properly completed
