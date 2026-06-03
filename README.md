# 📊 Directorate Report Automation

> Automate what used to take hours. Run it in seconds.

---

## The Problem

Every reporting cycle, the process looked like this:

1. Open the master Excel file
2. Filter by Directorate 1 → copy data → paste into new file → run the report → save
3. Repeat for Directorate 2
4. Repeat for Directorate 3
5. ...repeat 13 times

That's **13 manual operations** — same steps, different directorate — every single cycle. Slow, repetitive, and one wrong filter away from an error no one catches until it's too late.

---

## The Solution

A Python script that takes in **3 source Excel files**, automatically detects all 13 directorates, and generates **13 formatted Excel reports** — each with 5 sheets — in a matter of seconds.

No repeated steps. No filter fatigue. No human error from doing the same thing 13 times in a row.

---

## What It Does

```
3 input files  →  Script runs  →  13 directorate Excel files (5 sheets each)
```

### Input Files

| File | Description |
|---|---|
| `Staff_List.xlsx` | Full staff records across all directorates |
| `FY2026LeaveReport.xlsx` | Leave utilization data for the financial year |
| `FY26TrainingReport.xlsx` | Training completion records |

### Output — Per Directorate (13 files)

Each output file is named `DirectorateName_Report.xlsx` and contains:

| Sheet | Contents |
|---|---|
| **Attrition** | Filtered staff records for that directorate |
| **FY2026LeaveReport** | Leave data — ✅ green = Taken, 🔴 red = Not Taken |
| **Leave Aggregation** | Pivot by Department: Taken / Not Taken / Total / Utilization % |
| **FY26 Training Report** | Training records — ✅ green = Completed, 🔴 red = Not Completed |
| **Training Aggregation** | Pivot by Department: Completed / Not Completed / Total / Completion % |

---

## How the Script Works

The script is structured in 4 clear stages:

```
READ 3 Excel files → SPLIT by Directorate → BUILD 5 sheets → SAVE 13 files
```

### Key Functions

**`raw_sheet()`**
Builds the data sheets (Attrition, Leave Report, Training Report). Reads a filtered
DataFrame, writes it to a formatted sheet with a title banner, styled header row,
alternating row colours, and conditional colour coding on status columns.

**`leave_agg()`**
Builds the Leave Aggregation sheet. Uses pandas `groupby` and `unstack` to pivot
leave utilization by department, calculates totals and utilization percentage,
then writes a formatted summary table.

**`train_agg()`**
Same pivot logic as `leave_agg()` but applied to the Sub_Status column in the
training data — summarising Completed vs Not Completed counts by department.

**`main()`**
The orchestrator. Loads all 3 files, collects the full directorate list across
all files, then loops through each directorate — filtering the data, building
the workbook, calling all 5 sheet functions in order, and saving the output.

---

## Column Reference

The script reads these columns from your files. Column names are configurable
in the `CONFIG` block at the top of the script.

### Staff_List.xlsx
`Employee_Number` | `Full_Name` | `Original_Date_Of_Hire` | `Directorate` | `Dept_Name` | `Team_Unit` | `Grade` | `Job` | `staff_Location`

### FY2026LeaveReport.xlsx
`Person_Number` | `Full_Name` | `Manager` | `Directorate` | `Department` | `Leave_entitlement` | `CarryOver` | `Days_Taken` | `current_Bal` | `Perce_OfOutstandingDay` | `Leave_Utilization`

> **Note:** `Leave_Utilization` should contain the values `Taken` or `Not Taken`.
> If this column is absent, the script derives it automatically from `Days_Taken > 0`.

### FY26TrainingReport.xlsx
`Learning_Item_Name` | `Person_Number` | `First_Name` | `Last_Name` | `Sub_Status` | `Department_Name` | `Directorate`

> **Note:** `Sub_Status` should contain the values `Completed` or `Not Completed`.

---

## Setup & Usage

### Requirements

- Python 3.8+
- pandas
- openpyxl

### Installation

```bash
# Clone the repository
git clone https://github.com/YOUR-USERNAME/directorate-report-automation.git
cd directorate-report-automation

# Install dependencies
pip install pandas openpyxl
```

### Running the Script

Place your 3 source Excel files in the same folder as the script, then run:

```bash
python directorate_automation_v2.py
```

The script will print progress to the terminal as it processes each directorate:

```
════════════════════════════════════════════════════════════════════
   Directorate Report Automation  v2  —  5-Sheet Build
════════════════════════════════════════════════════════════════════

  Loaded   Staff_List.xlsx                      ( 1,100 rows)
  Loaded   FY2026LeaveReport.xlsx               ( 1,100 rows)
  Loaded   FY26TrainingReport.xlsx              ( 1,100 rows)

  [01/13] Administration                      Staff: 93  Leave: 80  Training: 85  ✓
  [02/13] Corporate Strategy                  Staff: 92  Leave:103  Training: 84  ✓
  ...
  [13/13] Risk Management                     Staff: 78  Leave: 94  Training: 80  ✓

  All 13 reports saved  →  ./directorate_reports/
```

### Configuring for Your Files

If your column names differ from the defaults, update the `CONFIG` block at the
top of the script — no changes needed anywhere else:

```python
FILES = {
    "staff":    "Staff_List.xlsx",         # your staff file name
    "leave":    "FY2026LeaveReport.xlsx",  # your leave file name
    "training": "FY26TrainingReport.xlsx", # your training file name
}

DIRECTORATE_COL      = "Directorate"      # column that holds directorate names
LEAVE_UTIL_COL       = "Leave_Utilization"
LEAVE_TAKEN_VAL      = "Taken"
LEAVE_NOT_TAKEN_VAL  = "Not Taken"
TRAINING_DEPT_COL    = "Department_Name"
TRAINING_STATUS_COL  = "Sub_Status"
TRAINING_DONE_VAL    = "Completed"
TRAINING_PENDING_VAL = "Not Completed"
```

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `File not found: Staff_List.xlsx` | Excel file not in same folder as script | Move all 3 Excel files to the script folder |
| `Column 'Directorate' not found` | Column name mismatch | Check exact spelling in your Excel file, update CONFIG |
| `ModuleNotFoundError: pandas` | Libraries not installed | Run `pip install pandas openpyxl` |
| `python is not recognized` | Python not on PATH | Reinstall Python and tick "Add Python to PATH" |

---

## Project Structure

```
directorate-report-automation/
│
├── directorate_automation_v2.py   ← main script
├── Staff_List.xlsx                ← input: staff data
├── FY2026LeaveReport.xlsx         ← input: leave data
├── FY26TrainingReport.xlsx        ← input: training data
│
└── directorate_reports/           ← output folder (auto-created)
    ├── Administration_Report.xlsx
    ├── Corporate_Strategy_Report.xlsx
    ├── Finance_&_Accounts_Report.xlsx
    └── ... (13 files total)
```

---

## Skills Demonstrated

- **Python** — scripting, functions, loops, conditionals, error handling
- **pandas** — reading Excel files, filtering DataFrames, groupby, pivot (unstack), concat
- **openpyxl** — writing formatted Excel workbooks, cell styling, merged cells, freeze panes
- **Process automation** — replacing manual repetitive workflows with a single script
- **Data reporting** — aggregation, percentage calculations, conditional formatting logic

---

## Author

Built by a Risk & Data Analyst at First Bank of Nigeria to automate a manual
HR reporting process across 13 directorates.



