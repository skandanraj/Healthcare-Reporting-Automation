# 🏥 Healthcare Reporting Automation Platform

## 🧠 Overview

This repository contains a modular healthcare reporting automation system designed to generate, clean, and distribute multiple operational monitoring reports automatically.

The platform supports **two execution modes**:

1️⃣ Local Python Scheduler Version (Windows-based automation)  
2️⃣ Jenkins Enterprise Version (CI/CD Orchestrated Automation)

Each report folder contains:

- `main.py` → Local execution version  
- `jenkins_version.py` → Jenkins-compatible execution version  

This architecture ensures flexibility across environments while maintaining identical business logic.

---

## 🏗 Architecture Overview

```
Healthcare_Reporting_Automation/
│
├── Cancelled_Appointments_Monitoring_Report/
│   ├── main.py
│   └── jenkins_version.py
│
├── Completed_Consultations_Monitoring_Report/
│   ├── main.py
│   └── jenkins_version.py
│
├── Dropout_Consultation_Report/
│   ├── main.py
│   └── jenkins_version.py
│
├── Missing_Prescription_Report/
│   ├── main.py
│   └── jenkins_version.py
│
├── Ops_Data_Sanitization/
│   ├── main.py
│   └── jenkins_version.py
│
├── scheduler.py
├── jenkins_master.py
├── requirements.txt
├── README.md
│
├── data/
│   └── MIS_Report.xlsx
│
└── logs/
```
---

# 📊 Reports Included

### 1️⃣ Cancelled Appointments Monitoring
- Cancelled & Paid (Revenue Leakage Detection)
- Cancelled (Yesterday + Today)
- Helps identify churn, dissatisfaction, and payment-to-cancellation patterns

### 2️⃣ Completed Consultations – Last 15 Days
- Tracks completed appointments
- Prevents duplicate email sending (cross-run deduplication)
- Maintains persistent sent-log state

### 3️⃣ Dropout Consultation Report
- Identifies users who reached payment stage but did not complete booking
- Detects booking funnel drop-offs

### 4️⃣ Missing Prescription Monitoring
- Detects completed appointments without prescription generation
- Monitors doctor compliance and platform usage

### 5️⃣ Operational Data Sanitization
- Removes confidential fields
- Provides controlled dataset for Ops team
- Supports data governance and access control

---

# ⚙️ Execution Modes

## 🖥 Local Python Scheduler Version

File: `scheduler.py`

### How It Works
- Runs daily at a configured time
- Checks if MIS_Report.xlsx is updated today
- Performs pre-cleanup of old Excel outputs
- Executes all report scripts sequentially
- Sends Windows toast notifications
- Maintains daily logs

### Advantages
- Easy setup
- Ideal for single-machine automation
- Real-time desktop notifications
- Lightweight execution

### Limitations
- Requires machine to remain ON
- Stops if system shuts down
- Not enterprise scalable

---

## 🏢 Jenkins Enterprise Version

File: `jenkins_master.py`

### How It Works
- Triggered by Jenkins job (cron-based or manual)
- Waits until MIS file is updated
- Enforces timeout safety window
- Performs automated cleanup
- Executes all modular report scripts
- Logs execution to workspace logs
- Exits with proper success/failure codes

### Advantages
- Runs even if personal machine is OFF
- Server-based automation
- Enterprise-grade logging
- CI/CD integration
- Failure handling with exit codes
- Timeout control
- Production-ready

Recommended for enterprise deployment.

---

# 🔁 Master Execution Flow

1. Wait for MIS update  
2. Pre-clean output folders  
3. Execute reports sequentially  
4. Log execution  
5. Exit safely  

---

# 🔐 Environment Setup

Create a `.env` file in each report folder:

EMAIL_USER=your_email@domain.com  
EMAIL_PASSWORD=your_app_password  

⚠️ Do NOT commit `.env`  
Add `.env` to `.gitignore`.

---

# 📦 Install Dependencies

pip install -r requirements.txt  

Or manually:

pip install pandas openpyxl python-dotenv schedule win10toast

---

# ▶️ How To Run

## Local Scheduler

python scheduler.py

## Jenkins Version

Configure Jenkins job and run:

python jenkins_master.py

---

# 🧹 Pre-Cleanup Logic

Before execution:
- Old Excel files are deleted from output folders
- Prevents stale file conflicts
- Ensures clean daily report generation

---

# 📄 Logging

Logs are generated inside:

logs/

Separate log file per day.

---

# 🎯 Business Impact

This automation platform enables:

- Revenue leakage detection
- Cancellation behavior monitoring
- Doctor compliance tracking
- Drop-off funnel analysis
- Operational dataset governance
- Multi-stakeholder automated reporting
- Reduced manual Excel effort
- Faster communication
- Improved ecosystem visibility

---

# 🚀 Strategic Value

By separating:

- Paid & Cancelled
- Cancelled without payment
- Dropout funnel users
- Missing prescriptions
- Sanitized operational data

The business can:

- Detect product issues early
- Identify churn patterns
- Improve doctor platform adoption
- Investigate booking friction
- Protect confidential data
- Strengthen operational intelligence

---

# 👨‍💻 Author

Skanda N Raj  
Data Analyst | Data Engineering
