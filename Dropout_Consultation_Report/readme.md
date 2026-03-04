# 📊 Dropout Consultation Report Automation

## 🧠 Overview

This script automates the generation and distribution of a daily **Dropout Consultation Report**.

The report identifies patients who:

- Had an appointment status of **Cancelled**
- Belong to selected hospitals
- Had appointment date = **Yesterday**
- (Optional) Have "Consider Patient" marked as **Yes**

After filtering, the script:

1. Generates a clean Excel report
2. Sends it automatically via Outlook SMTP
3. Emails it to configured stakeholders

---

## ⚙️ Business Logic

The script applies the following filters:

### 1️⃣ Appointment Status Filter
Only records where:
```
Appt. Status = "cancelled"
```

### 2️⃣ Date Filter
Only records where:
```
Appointment Date = Yesterday
```

The script dynamically calculates yesterday’s date:
```python
datetime.today().date() - timedelta(days=1)
```

### 3️⃣ Hospital Filter
Only records where:
```
Hospital Name is in allowed_hospitals list
```

### 4️⃣ Optional Filter
If column exists:
```
Consider Patient = "Yes"
```

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
Dropout_Consultation_Report/
│
├── main.py
├── README.md
├── requirements.txt
└── Dropout_Consultations_Karnataka.xlsx
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
output_file_cancelled = r"path_where_you_want_output.xlsx"
```

3. Run the script:

```bash
python main.py
```

---

## 📧 What Happens When You Run It

- MIS report is read
- Data is filtered
- Excel report is generated
- Email is sent with attachment
- Console shows success or error message

---

## 🚀 Business Impact

- Eliminates manual Excel filtering
- Automates daily reporting workflow
- Reduces human error
- Ensures timely communication
- Improves operational efficiency


