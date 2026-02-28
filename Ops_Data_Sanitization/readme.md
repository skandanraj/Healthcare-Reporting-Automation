# 📊 Ops Data Sanitization

## 🧠 Overview

This script standardizes a raw MIS report by retaining only required business columns.

It helps:

- Remove unnecessary fields
- Create a clean structured dataset
- Maintain consistent reporting format
- Prepare data for downstream automation

The script:

1. Loads the MIS Excel file
2. Keeps only predefined business columns
3. Warns if any required columns are missing
4. Saves a cleaned version of the file

---

## ⚙️ Business Logic

The script:

- Reads the MIS Excel file
- Compares available columns with a predefined required column list
- Retains only matching columns
- Displays missing columns (if any)
- Generates a cleaned Excel output

No data transformation is applied.
No rows are removed.
Only column selection is performed.

---

## 🛠 Tech Stack

- Python
- Pandas
- OpenPyXL

---

## 📂 Project Structure

```
mis_column_standardization/
│
├── main.py
├── README.md
├── requirements.txt
├── data/
│   └── MIS_Report.xlsx
│
└── output/
    └── MIS_Report_Cleaned.xlsx
```

---

## 📦 Install Dependencies

```
pip install -r requirements.txt
```

Or manually:

```
pip install pandas openpyxl
```

---

## ▶️ How to Run

1. Place your MIS file inside:

```
data/MIS_Report.xlsx
```

2. Run the script:

```
python main.py
```

---

## 📧 What Happens When You Run It

- MIS file is loaded
- Only required columns are retained
- Missing columns (if any) are displayed
- Cleaned Excel file is generated inside the output folder

---

## 🚀 Business Impact

This automation helps:

- Standardize MIS structure
- Improve downstream automation reliability
- Reduce manual Excel cleaning effort
- Ensure reporting consistency across teams
