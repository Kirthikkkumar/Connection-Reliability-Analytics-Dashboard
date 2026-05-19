#  Interactive Uptime & Downtime Analysis Dashboard

> **An end-to-end Data Analytics project** — from raw WebSocket event logs to an interactive Power BI dashboard that monitors system availability and connection reliability.

---

##  Project Overview

This project analyzes **WebSocket connection and disconnection events** from a real operational dataset to calculate exact uptime and downtime for each session. The processed data is visualized in a Power BI dashboard that gives instant insight into system availability, disconnection patterns, and daily connection trends.

The raw dataset contained only **event codes and timestamps** — no uptime or downtime values existed. The core challenge was **deriving uptime and downtime** by pairing each WSConnection event with its next WSDisconnection event and calculating the time difference.

---

##  Objectives

- Calculate exact uptime and downtime from raw connection event logs
- Identify high downtime periods and disconnection patterns
- Track uptime percentage and operational KPIs
- Build an interactive Power BI dashboard with dynamic cross-filtering
- Deliver clean, structured data from raw event logs using Python

---

##  Tools & Technologies Used

| Tool | Purpose |
|---|---|
| **Python (Pandas)** | Data processing, uptime/downtime calculation, feature engineering |
| **datetime / timedelta** | Timestamp subtraction to calculate session durations |
| **openpyxl** | Creating formatted multi-sheet Excel output for Power BI |
| **Power BI Desktop** | Dashboard development and visualization |
| **Power Query** | Data type correction and header promotion |
| **DAX** | KPI card measures — Connected Hours, Disconnected Hours, Availability % |
| **Excel** | Raw data source and output format |

> 

---

##  Dataset Information

| Field | Details |
|---|---|
| **File** | data_2.xlsx |
| **Sheet** | Event |
| **Total Rows** | 244 events |
| **Columns** | #, EventCode, OccurredDate |
| **Event Types** | WSConnection (152 events), WSDisconnection (92 events) |
| **Date Range** | 14 February 2026 to 17 March 2026 |
| **Format** | OccurredDate stored as string `dd-mm-yyyy hh:mm:ss AM/PM` |

**Sample raw data:**

| # | EventCode | OccurredDate |
|---|---|---|
| 1 | WSConnection | 17-03-2026 06:18:19 AM |
| 2 | WSDisconnection | 17-03-2026 06:18:01 AM |
| 3 | WSConnection | 17-03-2026 04:57:33 AM |

---

##  End-to-End Workflow

### 1️⃣ Data Processing using Python

The raw dataset had no uptime or downtime column. The entire calculation logic was built in Python.

**Steps performed:**
- Read Excel file using `pd.read_excel()`
- Parse timestamp strings to Python `datetime` objects using `pd.to_datetime()`
- Sort all events chronologically
- Loop through events — pair each `WSConnection` → next `WSDisconnection` = **uptime session**
- Loop through events — pair each `WSDisconnection` → next `WSConnection` = **downtime session**
- Calculate duration using `timedelta.total_seconds()`
- Export structured output to Excel using `openpyxl`

**Core Python Code:**

```python
import pandas as pd
from datetime import timedelta

# Read and parse data
df = pd.read_excel("data_2.xlsx")
df['OccurredDate'] = pd.to_datetime(
    df['OccurredDate'], format='%d-%m-%Y %I:%M:%S %p'
)
df = df.sort_values('OccurredDate').reset_index(drop=True)

# Calculate uptime and downtime by pairing consecutive events
events = df[['EventCode', 'OccurredDate']].values.tolist()
uptime_records = []
downtime_records = []

for i in range(len(events)):
    code, ts = events[i]
    if i + 1 < len(events):
        next_code, next_ts = events[i + 1]
        duration = (next_ts - ts).total_seconds()

        # UPTIME: WSConnection → next WSDisconnection
        if code == 'WSConnection' and next_code == 'WSDisconnection':
            uptime_records.append({
                'Connection Time':    ts,
                'Disconnection Time': next_ts,
                'Uptime Seconds':     int(duration),
                'Uptime Minutes':     round(duration / 60, 2),
                'Uptime HH:MM:SS':    str(timedelta(seconds=int(duration)))
            })

        # DOWNTIME: WSDisconnection → next WSConnection
        elif code == 'WSDisconnection' and next_code == 'WSConnection':
            downtime_records.append({
                'Disconnection Time':    ts,
                'Next Connection Time':  next_ts,
                'Downtime Seconds':      int(duration),
                'Downtime Minutes':      round(duration / 60, 2),
                'Downtime HH:MM:SS':     str(timedelta(seconds=int(duration)))
            })

up_df   = pd.DataFrame(uptime_records)
down_df = pd.DataFrame(downtime_records)

print(f"Total Uptime  : {str(timedelta(seconds=int(up_df['Uptime Seconds'].sum())))}")
print(f"Total Downtime: {str(timedelta(seconds=int(down_df['Downtime Seconds'].sum())))}")
```

---

### 2️⃣ Data Export using openpyxl

After processing, the data was exported into a structured Excel file with 4 sheets ready for Power BI:

| Sheet | Rows | Content |
|---|---|---|
| **Donut Data** | 2 | Connected (354.65 hrs) and Disconnected (0.40 hrs) |
| **KPI Cards** | 3 | Total Connection, Total Disconnection, Total Sessions |
| **Daily Summary** | 26 | One row per day — uptime hours and session count |
| **Downtime Sessions** | 89 | All downtime events with exact timestamps and duration |

---

### 3️⃣ Dashboard Development in Power BI

**Steps:**
- Loaded the 4-sheet Excel file using Get Data → Excel Workbook
- Fixed column data types in Power Query (Date/Time, Decimal Number)
- Used "Use First Row as Headers" in Power Query to fix column names
- Created 3 DAX measures for accurate KPI card values
- Built visuals across 2 pages with cross-filtering enabled

**DAX Measures:**

```dax
-- Connected Hours
Connected Hours =
CALCULATE(
    MAX('KPI Cards'[Value]),
    'KPI Cards'[Metric] = "Total Connection"
)

-- Disconnected Hours
Disconnected Hours =
CALCULATE(
    MAX('KPI Cards'[Value]),
    'KPI Cards'[Metric] = "Total Disconnection"
)

-- Availability %
Availability % =
DIVIDE(
    [Connected Hours],
    ([Connected Hours] + [Disconnected Hours])
) * 100
```

**Visuals built:**

| Visual | Sheet Used | Shows |
|---|---|---|
| KPI Card — Connected Hours | DAX measure | 354.65 hrs |
| KPI Card — Disconnected Hours | DAX measure | 0.40 hrs |
| KPI Card — Availability % | DAX measure | 99.89% |
| Donut Chart | Donut Data | Connected vs Disconnected ratio |
| Bar Chart | Daily Summary | Uptime hours per day (26 days) |
| Table | Downtime Sessions | All 89 downtime events with timestamps |

---

##  Key Results

| KPI | Value |
|---|---|
| **Total Uptime** | 354.65 hours — 14 days 18 hrs 38 min 57 sec |
| **Total Downtime** | 0.40 hours — 23 minutes 43 seconds |
| **Availability** | **99.89%** |
| **Uptime Sessions** | 88 |
| **Downtime Sessions** | 89 |
| **Average Uptime per Session** | ~241.81 minutes (~4 hours) |
| **Average Downtime per Session** | ~16 seconds |
| **Total Events in Raw Dataset** | 244 |
| **Data Period** | 14 Feb 2026 — 17 Mar 2026 (32 days) |

---

##  Business Insights

- System maintained **99.89% availability** — above the typical 99.5% SLA target
- All 89 disconnection events averaged only **16 seconds** each — brief reconnection drops, not major outages
- No single downtime event exceeded 1 minute — the system is **highly stable**
- Daily bar chart shows **consistent uptime** across all 26 days with no major problem dates
- The downtime table enables operations teams to **investigate specific disconnection timestamps** for root cause analysis

---

##  Skills Demonstrated

| Skill | Tool Used |
|---|---|
| ETL Pipeline | Python, Pandas |
| Timestamp Parsing & Arithmetic | datetime, timedelta |
| Feature Engineering | Pandas |
| Excel Automation | openpyxl |
| Data Type Correction | Power Query |
| KPI Measure Calculation | DAX |
| Dashboard Design | Power BI |
| Data Visualization | Power BI |
| Business Analysis | KPI interpretation |

---

##  Project Structure

```
uptime-downtime-dashboard/
│
├── data/
│   └── data_2.xlsx                    # Raw event log dataset
│
├── output/
│   ├── Uptime_Downtime_Analysis.xlsx  # Processed output (4 sheets)
│   └── PowerBI_Clean_Data.xlsx        # Clean data for Power BI
│
├── scripts/
│   └── uptime_analysis.py             # Main Python processing script
│
├── dashboard/
│   └── project_1.pbix                 # Power BI dashboard file
│
└── README.md
```

---

##  How to Run

```bash
# 1. Clone the repository
git clone https://github.com/kirthikk/uptime-downtime-dashboard.git

# 2. Install dependencies
pip install pandas openpyxl

# 3. Run the processing script
python scripts/uptime_analysis.py

# 4. Open Power BI dashboard
# Open dashboard/project_1.pbix in Power BI Desktop
# Refresh data source to point to your local output Excel file
```

---

##  Author

**Kirthikk Kumar S**

- 📧 krithikksivakumar@gmail.com
- 💼 [LinkedIn](https://linkedin.com/in/kirthikk-kumar-s)
- 📍 Salem, Tamil Nadu, India
