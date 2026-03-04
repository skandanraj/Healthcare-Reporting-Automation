## 📊 Dataset Overview

This dataset represents a **synthetic hospital appointment management dataset** generated for analytics, dashboarding, and data science experimentation.
It is derived from an operational Management Information System (MIS) structure used in healthcare environments such as **Aster DM Healthcare**.

The dataset simulates the **end-to-end lifecycle of patient appointments**, including booking, consultation, cancellation, and payment events.

---

## 📁 Dataset Characteristics

* **Rows:** 9,999 appointments
* **Columns:** 57 attributes
* **Type:** Structured tabular dataset
* **Format:** Excel (.xlsx)
* **Nature:** Synthetic data generated from real schema logic

---

## 🏥 Dataset Purpose

The dataset was created to support:

* Healthcare operations analytics
* Appointment workflow analysis
* Dashboard and reporting development
* Machine learning experimentation
* Data visualization projects

It preserves **realistic workflow logic while removing sensitive patient information**.

---

## 🔄 Appointment Lifecycle Logic

Each appointment follows a logical workflow depending on its **status**.

### Completed Appointment (`done`)

Booked DateTime
→ Appointment Date & Time
→ Patient Check-in
→ Consultation Start
→ Consultation Completed
→ Prescription Generated
→ Doctor/Patient Leave

---

### Cancelled Appointment (`cancelled`)

Booked DateTime
→ Cancelled Datetime
→ Appointment Date (never executed)

Consultation-related timestamps remain **null**.

---

### No Show (`no-show`)

Booked DateTime
→ Appointment Date

The patient never checks in, so consultation timestamps are **not generated**.

---

### Checked-In / Consulting

Booked DateTime
→ Appointment Date
→ Checked In Datetime
→ Consultation in progress

Completion timestamps may remain **null**.

---

## 📊 Key Data Categories

The dataset contains several groups of attributes:

### Patient Information

* Patient ID
* UHID
* Patient Name
* Mobile
* DOB
* Gender
* Location attributes

### Appointment Details

* Appointment ID
* Appointment Type
* Appointment Date & Time
* Booking Source
* Appointment Status

### Doctor & Hospital Information

* Doctor ID
* Doctor Name
* Specialty
* Hospital Name

### Consultation Timeline

* Booked DateTime
* Checked In Datetime
* Consultation DateTime
* Completed DateTime
* Prescription Generated DateTime
* Event Join/Leave timestamps

### Financial Information

* Invoice Number
* Consultation Fees
* Registration Fee
* Taxes (CGST / SGST / IGST)
* Payment Reference Number
* Refund Amount

---

## 🔐 Privacy and Data Safety

This dataset contains **synthetically generated values** for all personally identifiable information including:

* Patient names
* Mobile numbers
* Payment references
* Unique identifiers

The dataset therefore **does not contain any real patient data**.

---

## ⚙️ Data Generation Methodology

Synthetic values were generated using Python scripts that maintain realistic hospital workflow constraints, including:

* Logical ordering of timestamps
* Status-dependent event generation
* Realistic consultation durations
* Appointment booking lead times

---

## 📌 Use Cases

This dataset can be used for:

* Healthcare analytics dashboards
* Appointment scheduling analysis
* Operational efficiency studies
* Machine learning pipelines
* Data engineering practice
