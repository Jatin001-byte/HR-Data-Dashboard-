# 👥 Manpower & Attendance Tracker — Lecorb India Pvt. Ltd.

> **Macro-enabled Excel dashboard** for tracking daily attendance, absenteeism, and workforce strength across all manufacturing departments at a watch production facility (Total Strength: 2,408 employees).

---

## 📌 Problem Statement

The HR and production team at Lecorb India manually maintained attendance records across multiple departments, making it difficult to:
- Track real-time workforce availability
- Identify departments with high absenteeism
- Distinguish between Sick Leave (SL), Work Leave (WL), and unplanned absences
- Generate monthly attendance summaries for management reporting

This tool was built to **automate and centralise** all of this into a single macro-enabled Excel workbook.

---

## 📊 Dashboard Overview

| Sheet | Purpose |
|-------|---------|
| `Daily Attendance` | Date-wise attendance entry for all employees across departments |
| `Absent Reasons` | Categorises absences as SL / WL / Absent with employee-level detail |
| `Monthly Attendance` | Aggregated monthly summary with % metrics per department |
| `Holiday List` | Calendar of company holidays used for net working day calculations |

---

## 🏭 Departments Covered

The tracker covers **all production and support departments**:

`Assembly` · `Casing` · `Dial Fitting` · `Hand Fitting` · `Back Cover` · `O Ring` · `Wind Stem` · `Key Stem` · `Time Setting` · `Laser Marking` · `Strap Fitting` · `Strapping` · `Levelling` · `Modules` · `NPD / NAPS` · `Salvage` · `WRT` · `Preparation` · `Store` · `Quality` · `IGI` · `PPC` · `FG` · `Management` · `Finance`

---

## ⚙️ Key Features

- **Daily Attendance Register** — tracks Present / Absent / SL / WL per employee per date
- **Absent Reason Breakdown** — separates Sick Leave %, Work Leave %, and Absent % for each employee
- **Monthly Aggregation** — auto-calculates total strength vs available strength per month (March–February)
- **Department-wise Headcount** — tracks Regular vs Contract employees, Male vs Female, Staff vs Associate
- **Total Strength Tracker** — monitors available strength against total registered strength of 2,408
- **VBA Macros** — automate data entry, sheet updates, and summary calculations
- **Holiday Integration** — holiday list sheet feeds into net working day calculations

---

## 📈 Key Metrics Tracked

```
✅ Total Strength         → 2,408 registered employees
✅ Available Strength     → Daily headcount present
✅ Total Absent           → Daily absentee count
✅ Absent %               → (Absent / Total Strength) × 100
✅ SL %                   → Sick Leave as % of total
✅ WL %                   → Work Leave as % of total
✅ Department-wise splits → Per department attendance health
```

---

## 🛠️ Tools & Techniques Used

| Tool | Usage |
|------|-------|
| **Excel VBA (Macros)** | Automate attendance entry, sheet refresh, summary generation |
| **Power Query** | Data transformation and monthly rollup |
| **Pivot Tables** | Department-wise and month-wise aggregation |
| **Conditional Formatting** | Visual flags for high absenteeism |
| **Named Ranges** | Dynamic references across sheets |
| **Data Validation** | Dropdown lists for departments, leave types, employee categories |

---

## 📁 File Structure

```
MANPOWER_PWN_(M).xlsm
│
├── Daily Attendance       ← Main data entry sheet (date × employee matrix)
├── Absent Reasons         ← Leave categorisation (SL / WL / Absent)
├── Monthly Attendance     ← Monthly summary with % KPIs
└── Holiday List           ← Company holiday calendar
```

---

## 🚀 How to Use

1. **Enable Macros** when opening the file (required for automation to work)
2. Open the `Daily Attendance` sheet
3. Select the date and mark each employee's status (Present / Absent / SL / WL)
4. Run the macro to auto-update the `Monthly Attendance` summary
5. Use the `Absent Reasons` sheet to log and categorise leave types
6. Refer to the `Monthly Attendance` sheet for KPI reporting to management

> ⚠️ **Note:** Employee names and internal data have been used for demonstration purposes. In a production environment, sensitive data should be anonymised before sharing.

---

## 💡 Business Impact

- Eliminated manual attendance compilation — reduced HR reporting time significantly
- Enabled management to monitor **department-wise absenteeism trends** in real time
- Provided data foundation for **demand forecasting** and shift planning in production
- Standardised leave categorisation across 2,400+ employees and 25+ departments

---

## 👤 Author

**Jatin Gupta** — Data Analyst Intern, Lecorb India Pvt. Ltd.  
📧 jg6579392@gmail.com  
🔗 [linkedin.com/in/jatingupta](https://linkedin.com/in/jatingupta)  
🐙 [github.com/Jatin001-byte](https://github.com/Jatin001-byte)

---

## 📌 Related Projects

- [Watch Sales Analysis Dashboard](../watch-sales-dashboard) — Power BI + SQL sales KPI tracker
- [HR Data Dashboard](../hr-data-dashboard) — Excel HR analytics dashboard

