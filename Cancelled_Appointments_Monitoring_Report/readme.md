# 📊 Cancelled Appointments Monitoring Automation

## 🧠 Overview

This script automates the generation and distribution of two operational monitoring reports focused on cancelled appointments.

The system helps identify:

- Patients who cancelled appointments  
- Paid appointments that were later cancelled  
- Potential revenue leakage  
- Possible customer drop-offs  

The script generates two reports:

1️⃣ **Cancelled & Paid Appointments (Yesterday)**  
2️⃣ **Cancelled Appointments (Yesterday + Today)**  

These reports help the business:

- Understand why patients are cancelling  
- Identify operational or product issues  
- Track potential customer loss  
- Detect revenue leakage from paid cancellations  

After filtering, the script:

1. Generates two structured Excel reports  
2. Sends both reports automatically via Outlook SMTP  
3. Emails them to configured stakeholders  

---

## ⚙️ Business Logic

The script applies the following logic:

---

### 1️⃣ Report 1 – Cancelled & Paid (Yesterday)

This report identifies:

- Appointments that were paid  
- Later cancelled  
- Belong to selected hospitals  

#### Filters Applied

```
Appointment Date = Yesterday
AND
Appt. Status = "cancelled"
AND
Appt. Payment Status = "paid"
AND
Hospital Name in allowed_hospitals
```

Optional filter (if column exists):

```
Consider Patient = "Yes"
```

### 🎯 Business Purpose

This report helps detect:

- Revenue leakage  
- Refund scenarios  
- Payment-to-cancellation patterns  
- Possible friction in booking flow  
- Post-payment dissatisfaction  

---

### 2️⃣ Report 2 – Cancelled Appointments (Yesterday + Today)

This report identifies:

- All cancelled appointments  
- Across yesterday and today  
- Across selected hospitals  

#### Filters Applied

```
Appt. Status = "cancelled"
AND
Appointment Date IN (Yesterday, Today)
AND
Hospital Name in allowed_hospitals
```

Optional filter (if column exists):

```
Patient = "Yes"
```

### 🎯 Business Purpose

This report helps:

- Identify cancellation trends  
- Track potential customer churn  
- Detect hospital-level operational issues  
- Monitor cancellation spikes  
- Investigate reasons for patient drop-off  

---

## 📅 Date Logic

The script dynamically calculates:

```python
yesterday = datetime.today().date() - timedelta(days=1)
today = datetime.today().date()
```

No manual updates required.

---

## 🛠 Tech Stack

- Python  
- Pandas  
- OpenPyXL  
- SMTP (Office365)  
- python-dotenv (for secure credentials)  

---

## 📂 Project Structure

```
cancelled_appointments_monitoring/
│
├── main.py
├── README.md
├── requirements.txt
├── data/
│   └── MIS_Report.xlsx
├── cancelled_paid_yesterday.xlsx
├──cancelled_patients.xlsx
```

---

## 🔐 Environment Setup

Create a `.env` file in the project directory:

```
EMAIL_USER=your_email@yourdomain.com
EMAIL_PASSWORD=your_password
```

⚠️ Do NOT commit `.env` to GitHub.  
Add it to `.gitignore`.

---

## 📦 Install Dependencies

Make sure Python 3.8+ is installed.

Install using:

```
pip install -r requirements.txt
```

Or manually:

```
pip install pandas openpyxl python-dotenv
```

---

## ▶️ How to Run

1️⃣ Place MIS file inside:

```
data/MIS_Report.xlsx
```

2️⃣ Run the script:

```
python main.py
```

---

## 📧 What Happens When You Run It

- MIS report is loaded  
- Appointment date column is detected automatically  
- Data is filtered based on business logic  
- Two Excel reports are generated  
- Email is sent with both attachments  
- Console displays success or error message  

---

## 🚀 Business Impact

This automation enables:

- Monitoring of cancellation behavior  
- Early detection of revenue leakage  
- Identification of potential product issues  
- Understanding of patient dissatisfaction patterns  
- Tracking of paid-but-cancelled appointments  
- Reduction of manual Excel reporting effort  
- Faster stakeholder communication  

---

## 📈 Strategic Value

By separating:

- Paid & Cancelled  
- Cancelled without payment  

The business can:

- Investigate payment failures  
- Identify UX friction  
- Analyze doctor availability issues  
- Detect potential ecosystem bypass (patients moving off-platform)  
- Improve customer retention strategy  
