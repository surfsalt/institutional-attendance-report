# Institutional Attendance Report

Pulls attendance data from **Blackboard Learn** for every course in a term and generates a formatted **Excel dashboard**.

The tool has two parts:
1. **`extract.py`** — connects to the Blackboard REST API, downloads attendance data, and saves four CSV files
2. **`build_report.py`** — reads those CSVs and builds a multi-sheet Excel workbook with charts, conditional formatting, and risk flags

---

## What You Get

| Sheet | Description |
|---|---|
| **Dashboard** | Institution-wide KPIs, charts, and at-a-glance metrics |
| **Course Summary** | One row per course — avg attendance %, student counts, status |
| **Student Detail** | One row per student per course — individual rates and risk bands |
| **Daily Trends** | Attendance by course by meeting date |
| **Compliance** | Flags courses with missing or stale attendance |
| **Risk Pivot** | Pivot table by department and risk band |
| **Instructor Pivot** | Pivot table by instructor |
| **Data Model** | Documentation of formulas, scoring, and assumptions |
| **Config** | Configurable thresholds |

---

## Requirements

- **Python 3.8+** (already installed on most machines — [download here](https://www.python.org/downloads/) if not)
- A **Blackboard Learn** REST API integration (API key + secret)
- The integration user needs these entitlements:
  - `system.useradmin.generic.VIEW`
  - `system.course.VIEW`
  - `course.attendance.VIEW`
  - `course.configure-properties.EXECUTE`
  - `system.courseuserlist.VIEW`
  - `system.multiinst.hierarchy.manager.VIEW`
  - `system.multiinst.node.course.association.VIEW`

---

## Setup (One Time)

### Step 1 — Download the project

Click the green **Code** button above, then **Download ZIP**. Unzip it to a folder on your computer.

Or if you use Git:
```
git clone https://github.com/surfsalt/institutional-attendance-report.git
cd institutional-attendance-report
```

### Step 2 — Install Python libraries

Open a terminal/command prompt **in the project folder** and run:

```
pip install -r requirements.txt
```

> **Windows tip:** If `pip` isn't recognized, try `py -m pip install -r requirements.txt`

### Step 3 — Create your config file

Copy the example config:

- **Windows:** `copy extract_config.example.ini extract_config.ini`
- **Mac/Linux:** `cp extract_config.example.ini extract_config.ini`

Open `extract_config.ini` in any text editor and fill in your Blackboard details:

```ini
[blackboard]
base_url = https://ucci.blackboard.com
api_key = your-actual-api-key
api_secret = your-actual-api-secret

[settings]
threshold_pct = 75
stale_days = 14
```

> **Important:** Never share or commit `extract_config.ini` — it contains your API credentials. The `.gitignore` file already excludes it.

---

## Running the Report

### Step 1 — Extract data from Blackboard

```
python extract.py
```

What happens:
1. Authenticates with Blackboard
2. Shows a list of terms — type the number for the term you want (e.g. `6` for Spring 2026)
3. Fetches all courses, their department (via hierarchy API), meetings, and memberships
4. Fetches attendance records for every student (cached — each student is only looked up once)
5. Saves four CSV files to the `attendance_data/` folder

> **This step takes a while** (30–60 minutes for ~340 courses). Progress is printed as it goes. The script automatically throttles API calls to avoid rate limits, but if you do hit a `429` error, wait 15–30 minutes and try again.

### Step 2 — Build the Excel report

```
python build_report.py
```

This reads the CSVs and generates `BB_Global_Attendance_Report.xlsx` in the project folder.

Open it in Excel — all charts and formatting are already in place.

---

## Output Files

After both steps, your folder looks like this:

```
institutional-attendance-report/
├── extract.py                  ← Step 1: data extraction
├── build_report.py             ← Step 2: Excel builder
├── extract_config.ini          ← your credentials (not committed)
├── requirements.txt
├── attendance_data/            ← created by extract.py
│   ├── course_summary.csv
│   ├── student_detail.csv
│   ├── daily_attendance.csv
│   └── compliance.csv
└── BB_Global_Attendance_Report.xlsx  ← created by build_report.py
```

---

## Configuration

Edit `extract_config.ini` to change:

| Setting | Default | What it does |
|---|---|---|
| `threshold_pct` | `75` | Attendance % below which a student is flagged as "below threshold" |
| `stale_days` | `14` | Days since last attendance record before a course is flagged as "Stale" |

---

## Institutional Hierarchy

Departments are resolved using Blackboard's **Institutional Hierarchy API**, not by parsing course codes. For each course, the script calls:

```
GET /learn/api/public/v1/courses/{courseId}/nodes?expand=node
```

This returns the hierarchy node(s) the course belongs to, including the node's `title` (department name) and `parentId`. The script identifies the top-level department by checking which node's `parentId` matches the "All Departments" root (`_172_1`).

If a course isn't associated with any hierarchy node, its department column will be blank.

The root node ID is set in `extract.py` as `ALL_DEPARTMENTS_NODE_ID = "_172_1"`. If your hierarchy root changes, update that value.

---

## Attendance Scoring

| Status | Weight |
|---|---|
| Present | 100% |
| Late | 50% |
| Absent | 0% |
| Excused | Excluded from calculation |
| Not Marked | Excluded from calculation |

Formula: `(Present × 100 + Late × 50) / (Present + Late + Absent)`

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `pip` not found | Try `py -m pip install -r requirements.txt` (Windows) |
| `ModuleNotFoundError: requests` | Run `pip install requests` |
| `400` error on authentication | Check `api_key` and `api_secret` in `extract_config.ini` |
| `403` on hierarchy endpoints | Your API user needs the `system.multiinst.hierarchy.manager.VIEW` entitlement |
| `429 Too Many Requests` | Blackboard rate limit — wait 15–30 minutes and try again |
| Department column is blank | The course isn't associated with a hierarchy node in Blackboard |
| Script seems stuck at Step 3 | Normal — it's fetching attendance for every student. Watch the progress counter. |

---

## License

Internal use — UCCI.
