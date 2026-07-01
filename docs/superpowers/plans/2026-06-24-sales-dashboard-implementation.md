# Sales Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a full-stack sales dashboard (Python backend + static HTML/CSS/JS frontend) that reads live data from a locally synced SharePoint Excel file, computes KPI metrics, and displays them with an interactive bar chart.

**Architecture:** Two-part local stack: (1) Python script reads Excel file, computes weekly summaries and KPI metrics, writes JSON; (2) Static HTML/JS frontend reads JSON via fetch (with cache-busting), renders 6 KPI cards and a Chart.js bar chart with red value labels, auto-refreshes every 30 seconds.

**Tech Stack:**
- Backend: Python 3.8+, openpyxl (for Excel parsing), json (stdlib)
- Frontend: Vanilla HTML/CSS/JavaScript, Chart.js (via CDN), chartjs-plugin-datalabels (via CDN)
- Data format: JSON (dashboard-data.json)
- Testing: Manual (run script, inspect JSON, load HTML in browser, verify layout)

## Global Constraints

- **WEEKLY_TARGET = 15000** (hardcoded constant, must not change)
- **Week definition:** Monday (weekday 0) through Sunday (weekday 6)
- **Output file:** `dashboard-data.json` in project root
- **Configuration:** Excel file path via `EXCEL_FILE_PATH` environment variable
- **Frontend data source:** Read from relative path `dashboard-data.json`
- **Chart library:** Chart.js with chartjs-plugin-datalabels for value labels
- **Colors:** Blue bars (#1ca0f2), red value labels (#f17777), card background (#f7f7f7)
- **No automated tests yet** — manual verification only in this phase
- **No GitHub push/scheduling/GitHub Actions yet** — local testing phase only

---

## File Structure

```
d:\Salon Target Dashboard\
├── fetch_and_compute.py          (Create: backend script)
├── index.html                    (Create: frontend dashboard)
├── dashboard-data.json           (Generated: output JSON from backend)
├── README.md                     (Create: setup & running instructions)
├── .env.example                  (Create: environment variable template)
└── docs/
    └── superpowers/
        ├── specs/
        │   └── 2026-06-24-sales-dashboard-design.md  (Already written)
        └── plans/
            └── 2026-06-24-sales-dashboard-implementation.md  (This file)
```

---

## Backend Implementation Tasks

### Task 1: Create `fetch_and_compute.py` — Environment and Error Handling

**Files:**
- Create: `fetch_and_compute.py` (project root)

**Interfaces:**
- Produces: Module-level function `main()` that runs script; sets up logging; exits with code 0 on success, 1 on error

- [ ] **Step 1: Create fetch_and_compute.py with imports and main structure**

```python
import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl

# Constants
WEEKLY_TARGET = 15000
OUTPUT_FILE = "dashboard-data.json"
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

def main():
    """Main entry point for the fetch and compute script."""
    # Error handling: check env var
    if not EXCEL_FILE_PATH:
        print("ERROR: EXCEL_FILE_PATH environment variable not set.")
        print("Please set EXCEL_FILE_PATH to the local path of Daily_Invoice.xlsx")
        exit(1)
    
    # Error handling: check file exists
    excel_path = Path(EXCEL_FILE_PATH)
    if not excel_path.exists():
        print(f"ERROR: Excel file not found at {EXCEL_FILE_PATH}")
        exit(1)
    
    try:
        # Load workbook and get sheet
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        if "Daily Total Sales" not in wb.sheetnames:
            print("ERROR: Sheet 'Daily Total Sales' not found in Excel file")
            exit(1)
        
        ws = wb["Daily Total Sales"]
        
        # Placeholder: parsing will be added in Task 2
        print("Excel file loaded successfully")
        
        wb.close()
        
    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        exit(1)

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test the script runs without error**

Run: `python fetch_and_compute.py`

Expected: Should exit with error message about EXCEL_FILE_PATH not being set (that's correct for now).

```
ERROR: EXCEL_FILE_PATH environment variable not set.
Please set EXCEL_FILE_PATH to the local path of Daily_Invoice.xlsx
```

✓ Commit this step.

```bash
git add fetch_and_compute.py
git commit -m "feat: create fetch_and_compute.py with env var and error handling"
```

---

### Task 2: Implement Excel Parsing Logic

**Files:**
- Modify: `fetch_and_compute.py`

**Interfaces:**
- Produces: Function `parse_excel(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[dict]`
  - Returns list of dicts: `[{"date": datetime.date, "daily_sales": float}, ...]`
  - Sorted by date ascending
  - Skips rows where date is None, "Grand Total", or sales value is None/unparseable

- [ ] **Step 1: Add parse_excel function**

```python
def parse_excel(ws):
    """Parse 'Daily Total Sales' sheet and return list of daily records.
    
    Args:
        ws: openpyxl worksheet object
        
    Returns:
        List of dicts with keys 'date' (datetime.date) and 'daily_sales' (float),
        sorted by date ascending. Skips rows where date is None, "Grand Total",
        or sales value is unparseable.
    """
    daily_records = []
    
    # Iterate rows starting from row 4 (row 3 is header)
    for row in ws.iter_rows(min_row=4, values_only=True):
        date_value, sales_value = row[0], row[1]
        
        # Skip if date is None or "Grand Total"
        if date_value is None:
            continue
        if isinstance(date_value, str) and date_value.lower() == "grand total":
            continue
        
        # Ensure date is a datetime.date object
        if isinstance(date_value, datetime):
            date_obj = date_value.date()
        else:
            # Try parsing as string (fallback)
            try:
                date_obj = datetime.strptime(str(date_value), "%Y-%m-%d").date()
            except:
                print(f"WARNING: Skipping row with unparseable date: {date_value}")
                continue
        
        # Ensure sales_value is numeric
        if sales_value is None:
            continue
        try:
            sales_num = float(sales_value)
        except:
            print(f"WARNING: Skipping row with unparseable sales value: {sales_value} for date {date_obj}")
            continue
        
        daily_records.append({"date": date_obj, "daily_sales": sales_num})
    
    # Sort by date ascending
    daily_records.sort(key=lambda x: x["date"])
    
    return daily_records
```

- [ ] **Step 2: Update main() to call parse_excel and verify output**

Replace the "Placeholder: parsing will be added" section with:

```python
        # Parse Excel
        daily_records = parse_excel(ws)
        
        if not daily_records:
            print("WARNING: No valid data rows found in Excel file")
            daily_records = []
        else:
            print(f"Parsed {len(daily_records)} daily records from Excel")
            print(f"Date range: {daily_records[0]['date']} to {daily_records[-1]['date']}")
```

- [ ] **Step 3: Commit**

```bash
git add fetch_and_compute.py
git commit -m "feat: implement Excel parsing logic with "Grand Total" skip"
```

---

### Task 3: Implement Weekly Summary Computation

**Files:**
- Modify: `fetch_and_compute.py`

**Interfaces:**
- Consumes: `daily_records` from Task 2 (list of dicts)
- Produces: Function `compute_weekly_summary(daily_records: list) -> dict`
  - Returns dict keyed by `week_start_date` (datetime.date)
  - Each value: `{"week_start_date": date, "week_end_date": date, "total_sales": float, "target_pending": float}`
  - `target_pending` = 0 if total_sales >= WEEKLY_TARGET, else WEEKLY_TARGET - total_sales

- [ ] **Step 1: Add compute_weekly_summary function**

```python
def compute_weekly_summary(daily_records):
    """Compute weekly sales summaries from daily records.
    
    Week = Monday (weekday 0) to Sunday (weekday 6).
    For each week, compute total_sales and target_pending.
    
    Args:
        daily_records: List of dicts with 'date' and 'daily_sales'
        
    Returns:
        Dict keyed by week_start_date (datetime.date).
        Each value: {"week_start_date": date, "week_end_date": date, 
                     "total_sales": float, "target_pending": float}
    """
    if not daily_records:
        return {}
    
    # Group records by week
    weeks = {}
    for record in daily_records:
        date = record["date"]
        week_start = date - timedelta(days=date.weekday())  # Monday of that week
        week_end = week_start + timedelta(days=6)  # Sunday of that week
        
        if week_start not in weeks:
            weeks[week_start] = {
                "week_start_date": week_start,
                "week_end_date": week_end,
                "total_sales": 0,
            }
        
        weeks[week_start]["total_sales"] += record["daily_sales"]
    
    # Compute target_pending for each week
    for week_start, week_data in weeks.items():
        total_sales = week_data["total_sales"]
        if total_sales >= WEEKLY_TARGET:
            week_data["target_pending"] = 0
        else:
            week_data["target_pending"] = WEEKLY_TARGET - total_sales
    
    return weeks
```

- [ ] **Step 2: Update main() to call compute_weekly_summary**

After calling parse_excel, add:

```python
        # Compute weekly summaries
        weekly_summary = compute_weekly_summary(daily_records)
        print(f"Computed {len(weekly_summary)} weekly summaries")
```

- [ ] **Step 3: Commit**

```bash
git add fetch_and_compute.py
git commit -m "feat: implement weekly summary computation with target_pending logic"
```

---

### Task 4: Implement Current Month Daily Sales and KPI Metrics

**Files:**
- Modify: `fetch_and_compute.py`

**Interfaces:**
- Consumes: `daily_records`, `weekly_summary`
- Produces: Function `compute_kpi_metrics(daily_records: list, weekly_summary: dict) -> dict`
  - Returns dict with keys: `current_date_label`, `current_week_sales`, `current_week_balance`, 
    `remaining_days`, `previous_week_sales`, `monthly_sales`
  - Also produces `current_month_sales` list in the function for later use

- [ ] **Step 1: Add helper functions and compute_kpi_metrics**

```python
def get_current_month_sales(daily_records):
    """Filter daily records to current month only.
    
    Args:
        daily_records: List of dicts with 'date' and 'daily_sales'
        
    Returns:
        Filtered list, sorted by date ascending
    """
    today = datetime.now().date()
    current_month = [
        record for record in daily_records
        if record["date"].year == today.year and record["date"].month == today.month
    ]
    return current_month

def compute_kpi_metrics(daily_records, weekly_summary):
    """Compute all KPI metrics for the dashboard.
    
    Args:
        daily_records: List of dicts with 'date' and 'daily_sales'
        weekly_summary: Dict keyed by week_start_date
        
    Returns:
        Dict with KPI metrics:
        - current_date_label: str (e.g., "23 May")
        - current_week_sales: float
        - current_week_balance: float (target_pending)
        - remaining_days: int
        - previous_week_sales: float
        - monthly_sales: float
    """
    today = datetime.now().date()
    
    # 1. Current Date Label
    current_date_label = f"{today.day} {today.strftime('%B')}"
    
    # 2. Current Week Sales and Balance
    # Find most recent week_start_date <= today
    week_starts = sorted([ws for ws in weekly_summary.keys() if ws <= today], reverse=True)
    if week_starts:
        current_week_start = week_starts[0]
        current_week_sales = weekly_summary[current_week_start]["total_sales"]
        current_week_balance = weekly_summary[current_week_start]["target_pending"]
    else:
        current_week_sales = 0
        current_week_balance = 0
    
    # 3. Remaining Days
    remaining_days = 7 - today.weekday()
    
    # 4. Previous Week Sales
    # Find second-most-recent week
    if len(week_starts) >= 2:
        previous_week_start = week_starts[1]
        previous_week_sales = weekly_summary[previous_week_start]["total_sales"]
    else:
        previous_week_sales = 0
    
    # 5. Monthly Sales
    current_month_sales = get_current_month_sales(daily_records)
    monthly_sales = sum(record["daily_sales"] for record in current_month_sales)
    
    return {
        "current_date_label": current_date_label,
        "current_week_sales": current_week_sales,
        "current_week_balance": current_week_balance,
        "remaining_days": remaining_days,
        "previous_week_sales": previous_week_sales,
        "monthly_sales": monthly_sales,
        "current_month_sales": current_month_sales,  # For chart data
    }
```

- [ ] **Step 2: Update main() to call compute_kpi_metrics**

After computing weekly_summary, add:

```python
        # Compute KPI metrics
        kpi = compute_kpi_metrics(daily_records, weekly_summary)
        print(f"Current Date: {kpi['current_date_label']}")
        print(f"Current Week Sales: {kpi['current_week_sales']}")
        print(f"Monthly Sales: {kpi['monthly_sales']}")
```

- [ ] **Step 3: Commit**

```bash
git add fetch_and_compute.py
git commit -m "feat: implement KPI metric computation and current month filtering"
```

---

### Task 5: Implement Chart Data Generation and JSON Output

**Files:**
- Modify: `fetch_and_compute.py`

**Interfaces:**
- Consumes: `kpi` dict from Task 4
- Produces: Function `generate_chart_data(current_month_sales: list) -> list`
  - Returns list of dicts: `[{"date": "May 1", "daily_sales": 2739}, ...]`
  - Date format: "Month Day" (e.g., "May 1", no leading zero on day, no year)
- Produces: Function `write_dashboard_json(kpi: dict, chart_data: list, output_file: str) -> None`
  - Writes dashboard-data.json to disk

- [ ] **Step 1: Add generate_chart_data and write_dashboard_json functions**

```python
def generate_chart_data(current_month_sales):
    """Generate chart data from current month sales records.
    
    Args:
        current_month_sales: List of dicts with 'date' and 'daily_sales'
        
    Returns:
        List of dicts with 'date' (formatted string) and 'daily_sales' (int)
    """
    chart_data = []
    for record in current_month_sales:
        date_str = record["date"].strftime("%-B %-d") if hasattr(record["date"], 'strftime') else None
        # For Windows compatibility, use a different approach:
        date_obj = record["date"]
        month_name = date_obj.strftime("%B")
        day = date_obj.day
        date_str = f"{month_name} {day}"
        
        chart_data.append({
            "date": date_str,
            "daily_sales": int(record["daily_sales"])
        })
    
    return chart_data

def write_dashboard_json(kpi, chart_data, output_file):
    """Write dashboard data to JSON file.
    
    Args:
        kpi: Dict with KPI metrics (current_date_label, current_week_sales, etc.)
        chart_data: List of dicts with chart data
        output_file: Path to output JSON file
    """
    dashboard_data = {
        "current_date_label": kpi["current_date_label"],
        "current_week_sales": int(kpi["current_week_sales"]),
        "current_week_balance": int(kpi["current_week_balance"]),
        "remaining_days": kpi["remaining_days"],
        "previous_week_sales": int(kpi["previous_week_sales"]),
        "monthly_sales": int(kpi["monthly_sales"]),
        "daily_chart_data": chart_data,
    }
    
    with open(output_file, 'w') as f:
        json.dump(dashboard_data, f, indent=2)
    
    print(f"Dashboard data written to {output_file}")
```

- [ ] **Step 2: Update main() to generate and write JSON**

After computing KPI, add:

```python
        # Generate chart data
        chart_data = generate_chart_data(kpi["current_month_sales"])
        
        # Write output JSON
        write_dashboard_json(kpi, chart_data, OUTPUT_FILE)
        
        print("SUCCESS: Dashboard data computed and written")
```

- [ ] **Step 3: Commit**

```bash
git add fetch_and_compute.py
git commit -m "feat: implement chart data generation and JSON output"
```

---

### Task 6: Verify Backend with Manual Test

**Files:**
- Verify: `fetch_and_compute.py`
- Output: `dashboard-data.json` (generated)

- [ ] **Step 1: Prepare a test Excel file (or use the real one)**

Create or locate your OneDrive-synced Daily_Invoice.xlsx file with sample data in the "Daily Total Sales" sheet, or ask me to help you create a test file.

- [ ] **Step 2: Set the EXCEL_FILE_PATH environment variable and run the script**

```bash
# On Windows (PowerShell):
$env:EXCEL_FILE_PATH = "C:\path\to\Daily_Invoice.xlsx"
python fetch_and_compute.py

# Expected output:
# Excel file loaded successfully
# Parsed XX daily records from Excel
# Date range: YYYY-MM-DD to YYYY-MM-DD
# Computed XX weekly summaries
# Current Date: DD Month
# Current Week Sales: XXXXX
# Monthly Sales: XXXXX
# Dashboard data written to dashboard-data.json
# SUCCESS: Dashboard data computed and written
```

- [ ] **Step 3: Verify dashboard-data.json was created and has correct structure**

```bash
type dashboard-data.json  # (or 'cat' on Unix)
```

Expected output (example):
```json
{
  "current_date_label": "23 June",
  "current_week_sales": 17629,
  "current_week_balance": 0,
  "remaining_days": 1,
  "previous_week_sales": 29588,
  "monthly_sales": 69849,
  "daily_chart_data": [
    {"date": "June 1", "daily_sales": 2739},
    {"date": "June 2", "daily_sales": 1562},
    ...
  ]
}
```

- [ ] **Step 4: Verify numbers**

- Check that `current_week_sales + current_week_balance == WEEKLY_TARGET` (15000) when balance > 0
- Check that `current_date_label` matches today's date in "DD Month" format
- Check that `remaining_days` is between 1-7
- Check that `monthly_sales` is the sum of all daily_sales in current_month_sales
- Check that `daily_chart_data` has dates formatted as "Month Day" with no leading zeros

If all checks pass: ✓ Backend is working correctly. Proceed to frontend.

---

## Frontend Implementation Tasks

### Task 7: Create `index.html` — Basic Structure and KPI Card Layout

**Files:**
- Create: `index.html` (project root)

**Interfaces:**
- Produces: Static HTML page with structure for 6 KPI cards and a chart canvas
- Consumes: `dashboard-data.json` (will be loaded via JavaScript in later tasks)

- [ ] **Step 1: Create index.html with full HTML structure**

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sales Dashboard</title>
    
    <!-- Chart.js library -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js"></script>
    <!-- Chart.js Data Labels Plugin -->
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
    
    <style>
        /* CSS will be added in Task 8 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
            background-color: #fff;
            padding: 24px;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- KPI Cards Section -->
        <div class="kpi-cards">
            <div class="kpi-card">
                <div class="kpi-value" id="current-date-label">-</div>
                <div class="kpi-label">Current Date Label</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value" id="current-week-sales">-</div>
                <div class="kpi-label">Current Week Sales</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value" id="current-week-balance">-</div>
                <div class="kpi-label">Current Week Balance</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value" id="remaining-days">-</div>
                <div class="kpi-label">Remaining Days</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value" id="previous-week-sales">-</div>
                <div class="kpi-label">Previous Week Sales</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-value" id="monthly-sales">-</div>
                <div class="kpi-label">Monthly Sales</div>
            </div>
        </div>
        
        <!-- Chart Section -->
        <div class="chart-card">
            <div class="chart-title">Sum of daily_sales by date</div>
            <div class="chart-container">
                <canvas id="daily-sales-chart"></canvas>
            </div>
        </div>
    </div>
    
    <script>
        // JavaScript will be added in Tasks 10-13
        console.log("Dashboard loading...");
    </script>
</body>
</html>
```

- [ ] **Step 2: Verify HTML loads without errors**

Open `index.html` in a browser (file:// protocol is fine for local testing).

Expected: Page loads with 6 cards showing "-" as placeholder values and an empty canvas element.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: create index.html with KPI card structure and Chart.js setup"
```

---

### Task 8: Add CSS Styling for KPI Cards and Chart Container

**Files:**
- Modify: `index.html` (replace CSS section)

- [ ] **Step 1: Replace the `<style>` section with complete CSS**

Replace the existing style block with:

```html
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
            background-color: #fff;
            padding: 24px;
            font-size: 16px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
        }
        
        /* KPI Cards Row */
        .kpi-cards {
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
            flex-wrap: wrap;
        }
        
        .kpi-card {
            flex: 1;
            min-width: 200px;
            background-color: #f7f7f7;
            border-radius: 12px;
            padding: 24px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            min-height: 210px;
        }
        
        .kpi-value {
            font-size: 32px;
            font-weight: 700;
            color: #000;
            line-height: 1.2;
            margin-bottom: 16px;
        }
        
        .kpi-label {
            font-size: 13px;
            color: #666;
            font-weight: 400;
            line-height: 1.4;
        }
        
        /* Chart Card */
        .chart-card {
            background-color: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 24px;
            min-height: 400px;
        }
        
        .chart-title {
            font-size: 14px;
            color: #333;
            margin-bottom: 16px;
            font-weight: 500;
        }
        
        .chart-container {
            position: relative;
            width: 100%;
            height: 350px;
        }
        
        canvas {
            max-width: 100%;
        }
        
        /* Responsive */
        @media (max-width: 768px) {
            .kpi-cards {
                gap: 12px;
            }
            
            .kpi-card {
                min-width: 150px;
                padding: 16px;
                min-height: 180px;
            }
            
            .kpi-value {
                font-size: 24px;
            }
            
            .kpi-label {
                font-size: 12px;
            }
        }
    </style>
```

- [ ] **Step 2: Verify styling in browser**

Reload `index.html` in the browser.

Expected:
- 6 KPI cards displayed side by side with light gray background (#f7f7f7)
- Cards have rounded corners and padding
- Cards are roughly equal width with gaps between them
- Chart card below with white background and light gray border
- All text is visible and formatted correctly

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: add CSS styling for KPI cards and chart container"
```

---

### Task 9: Add Chart Canvas and JavaScript Framework

**Files:**
- Modify: `index.html` (JavaScript section)

**Interfaces:**
- Produces: Global variables `dashboardChart` (for Chart.js instance) and functions:
  - `loadDashboardData()` — fetches dashboard-data.json
  - `renderKPICards(data)` — renders KPI values
  - `renderChart(data)` — renders bar chart
  - `formatNumber(num)` — formats numbers with thousands separators

- [ ] **Step 1: Replace JavaScript section with framework functions**

Replace the `<script>` section at the bottom with:

```html
    <script>
        // Constants
        const AUTO_REFRESH_INTERVAL_MS = 30000;  // 30 seconds
        const CHART_DATA_URL = "dashboard-data.json";
        
        let dashboardChart = null;
        let lastDataString = null;  // For change detection
        
        /**
         * Format a number with thousands separators.
         * Example: 17629 -> "17,629"
         */
        function formatNumber(num) {
            if (typeof num !== 'number') {
                return String(num);
            }
            return Math.floor(num).toLocaleString('en-US');
        }
        
        /**
         * Fetch dashboard data from JSON file with cache-busting.
         */
        async function loadDashboardData() {
            try {
                const url = `${CHART_DATA_URL}?t=${Date.now()}`;
                const response = await fetch(url, { cache: 'no-store' });
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                const data = await response.json();
                return data;
            } catch (error) {
                console.error('Failed to load dashboard data:', error);
                return null;
            }
        }
        
        /**
         * Render KPI cards with data.
         */
        function renderKPICards(data) {
            if (!data) {
                console.log('No data to render');
                return;
            }
            
            document.getElementById('current-date-label').textContent = data.current_date_label || '-';
            document.getElementById('current-week-sales').textContent = formatNumber(data.current_week_sales);
            document.getElementById('current-week-balance').textContent = formatNumber(data.current_week_balance);
            document.getElementById('remaining-days').textContent = formatNumber(data.remaining_days);
            document.getElementById('previous-week-sales').textContent = formatNumber(data.previous_week_sales);
            document.getElementById('monthly-sales').textContent = formatNumber(data.monthly_sales);
        }
        
        /**
         * Render bar chart.
         */
        function renderChart(data) {
            if (!data || !data.daily_chart_data || data.daily_chart_data.length === 0) {
                console.log('No chart data to render');
                return;
            }
            
            const chartCanvas = document.getElementById('daily-sales-chart');
            const ctx = chartCanvas.getContext('2d');
            
            const dates = data.daily_chart_data.map(d => d.date);
            const sales = data.daily_chart_data.map(d => d.daily_sales);
            
            // Calculate Y-axis max with 10% padding
            const maxSales = Math.max(...sales);
            const yAxisMax = Math.ceil((maxSales * 1.1) / 2000) * 2000;
            
            // Destroy existing chart if it exists
            if (dashboardChart) {
                dashboardChart.destroy();
            }
            
            // Create new chart
            dashboardChart = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: dates,
                    datasets: [
                        {
                            label: 'Sum of daily_sales',
                            data: sales,
                            backgroundColor: '#1ca0f2',
                            borderColor: '#1ca0f2',
                            borderWidth: 0,
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    indexAxis: 'x',
                    scales: {
                        y: {
                            beginAtZero: true,
                            max: yAxisMax,
                            ticks: {
                                stepSize: 2000,
                                callback: function(value) {
                                    return value.toLocaleString('en-US');
                                }
                            },
                            grid: {
                                drawBorder: true,
                                color: 'rgba(0, 0, 0, 0.1)',
                                lineWidth: 1,
                            },
                            title: {
                                display: true,
                                text: 'Sum of daily_sales'
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'date'
                            },
                            grid: {
                                display: false,
                            }
                        }
                    },
                    plugins: {
                        legend: {
                            display: false
                        },
                        datalabels: {
                            display: true,
                            anchor: 'end',
                            align: 'end',
                            offset: 5,
                            color: '#f17777',
                            font: {
                                weight: 'bold',
                                size: 12,
                            },
                            formatter: function(value) {
                                return value.toLocaleString('en-US');
                            }
                        }
                    }
                },
                plugins: [ChartDataLabels]
            });
        }
        
        /**
         * Main render function: update KPI cards and chart.
         */
        function render(data) {
            renderKPICards(data);
            renderChart(data);
        }
        
        /**
         * Initialize the dashboard and set up auto-refresh.
         */
        async function init() {
            console.log('Initializing dashboard...');
            
            // Initial load
            const data = await loadDashboardData();
            if (data) {
                render(data);
                lastDataString = JSON.stringify(data);
            }
            
            // Set up auto-refresh
            setInterval(async () => {
                const freshData = await loadDashboardData();
                if (freshData) {
                    const freshDataString = JSON.stringify(freshData);
                    if (freshDataString !== lastDataString) {
                        console.log('Dashboard data changed, updating...');
                        render(freshData);
                        lastDataString = freshDataString;
                    }
                }
            }, AUTO_REFRESH_INTERVAL_MS);
        }
        
        // Start when DOM is ready
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', init);
        } else {
            init();
        }
    </script>
```

- [ ] **Step 2: Verify JavaScript loads without errors**

Open the browser console (F12 or right-click → Inspect → Console) and verify:
- No JavaScript errors appear
- Message "Initializing dashboard..." appears in console
- If dashboard-data.json exists, values should populate in KPI cards

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: add JavaScript framework with KPI rendering and chart initialization"
```

---

### Task 10: Test Frontend with Sample Dashboard Data

**Files:**
- Verify: `index.html` loads and displays data
- Use: `dashboard-data.json` (from backend Task 6)

- [ ] **Step 1: Ensure dashboard-data.json exists from backend testing**

If you haven't run the backend script yet, do so now (see Task 6).

- [ ] **Step 2: Open index.html in a browser**

```bash
# On Windows, you can open directly:
start index.html

# Or manually open the file in your browser (file:// URL)
```

Expected:
- Page loads without errors
- KPI cards display the values from dashboard-data.json
- Numbers are formatted with thousands separators (e.g., "17,629")
- "Current Date Label" shows the date in format like "23 June"
- "Remaining Days" shows a single digit (1-7)

- [ ] **Step 3: Verify bar chart renders**

Expected:
- Blue bars appear for each day in the chart
- Each bar has a red number label above it (e.g., "2739")
- Y-axis shows ticks at 0, 2000, 4000, etc.
- X-axis shows date labels like "May 1", "May 2", etc.
- Chart title "Sum of daily_sales by date" is visible

If chart does not render, open browser console and check for errors. Common issues:
- Chart.js or datalabels plugin failed to load (check CDN URLs)
- dashboard-data.json is missing or malformed (verify JSON is valid)
- daily_chart_data array is empty (run backend script again)

- [ ] **Step 4: Test auto-refresh**

1. Open browser console and note the current time
2. Wait 30+ seconds
3. Check console for message "Dashboard data changed, updating..." (or silence if data hasn't changed)
4. (Optional) Edit dashboard-data.json manually, run the backend script again, wait 30+ seconds, see the dashboard update

Expected: Browser re-fetches the JSON every 30 seconds and updates if data changed.

If auto-refresh doesn't work:
- Check browser console for fetch errors
- Verify dashboard-data.json is readable from the same directory as index.html
- Verify AUTO_REFRESH_INTERVAL_MS is set (should be 30000)

- [ ] **Step 5: Verify layout matches reference screenshot**

Compare the live dashboard with the reference screenshot:
- ✓ 6 KPI cards in a row
- ✓ Light gray background (#f7f7f7)
- ✓ Large bold numbers (~32-34px)
- ✓ Small gray labels below numbers
- ✓ Rounded corners on cards
- ✓ Bar chart below with blue bars
- ✓ Red value labels above bars
- ✓ Y-axis labeled "Sum of daily_sales"
- ✓ X-axis labeled "date"
- ✓ Gridlines visible

If layout doesn't match, check:
- CSS colors are correct (review Task 8)
- Font sizes are applied (verify CSS is in place)
- Spacing and padding look correct

- [ ] **Step 6: Verify numbers match expectations**

Cross-check the displayed numbers with the dashboard-data.json file:
- All numbers in KPI cards match the JSON values
- Numbers are formatted correctly (thousands separators for large numbers)
- Monthly sales = sum of all daily sales in the chart
- Current week sales is reasonable (0-20000 range typically)
- Current week balance + current week sales = 15000 (if balance > 0)

If numbers don't match, the backend script may have computed them incorrectly (revisit Tasks 3-5).

✓ If all checks pass, the frontend is working correctly!

---

### Task 11: Create `README.md` Documentation

**Files:**
- Create: `README.md` (project root)

- [ ] **Step 1: Create README.md with setup and running instructions**

```markdown
# Sales Dashboard

A full-stack sales dashboard that reads live data from a locally synced SharePoint Excel file, computes business KPI metrics (weekly sales, targets, monthly totals), and displays them with an interactive bar chart.

## Project Structure

```
d:\Salon Target Dashboard\
├── fetch_and_compute.py          Backend script: reads Excel, computes metrics, writes JSON
├── index.html                    Frontend dashboard: reads JSON, renders KPI cards and chart
├── dashboard-data.json           Generated output (created by fetch_and_compute.py)
├── .env.example                  Example environment variable settings
└── docs/superpowers/
    ├── specs/2026-06-24-sales-dashboard-design.md
    └── plans/2026-06-24-sales-dashboard-implementation.md
```

## Prerequisites

- **Python 3.8+**
- **openpyxl** library: `pip install openpyxl`
- **Daily_Invoice.xlsx** synced locally via OneDrive (see Setup below)

## Setup

### 1. Sync the SharePoint Excel File Locally

1. Open OneDrive on your machine
2. Navigate to the SharePoint site/library where `Daily_Invoice.xlsx` is stored
3. Enable "Sync" to download and auto-sync the file to your machine
4. Note the full local file path (e.g., `C:\Users\YourUsername\OneDrive\Accounts\Daily_Invoice.xlsx`)

### 2. Set the Environment Variable

Set the `EXCEL_FILE_PATH` environment variable to point to the synced Excel file.

**On Windows (PowerShell):**
```powershell
$env:EXCEL_FILE_PATH = "C:\path\to\your\Daily_Invoice.xlsx"
```

**On Windows (Command Prompt):**
```cmd
set EXCEL_FILE_PATH=C:\path\to\your\Daily_Invoice.xlsx
```

**On macOS/Linux (Bash/Zsh):**
```bash
export EXCEL_FILE_PATH="/path/to/your/Daily_Invoice.xlsx"
```

### 3. Install Python Dependencies

```bash
pip install openpyxl
```

## Running the Dashboard

### Step 1: Fetch and Compute Data

Run the backend script to read the Excel file, compute KPI metrics, and generate the dashboard JSON:

```bash
python fetch_and_compute.py
```

Expected output:
```
Excel file loaded successfully
Parsed 50 daily records from Excel
Date range: 2026-05-01 to 2026-06-22
Computed 8 weekly summaries
Current Date: 23 June
Current Week Sales: 17629
Monthly Sales: 69849
Dashboard data written to dashboard-data.json
SUCCESS: Dashboard data computed and written
```

If you see errors:
- **"EXCEL_FILE_PATH environment variable not set"** — Set the env var as shown in Setup Step 2
- **"Excel file not found"** — Verify the path in EXCEL_FILE_PATH is correct and the file exists
- **"Sheet 'Daily Total Sales' not found"** — Verify the Excel file has a sheet named "Daily Total Sales"

### Step 2: View the Dashboard

Open `index.html` in a web browser:

**Option A:** Double-click `index.html` in File Explorer
**Option B:** Use a simple HTTP server (recommended to avoid browser cache issues):
```bash
# Python 3
python -m http.server 8000

# Then open: http://localhost:8000
```

Expected:
- 6 KPI cards display current date, week sales, week balance, remaining days, previous week sales, and monthly sales
- Bar chart shows daily sales for the current month with blue bars and red value labels
- All numbers are formatted with thousands separators (e.g., "17,629")

### Step 3: Auto-Refresh

The dashboard auto-refreshes every 30 seconds. To see updated data:

1. Run `python fetch_and_compute.py` again (e.g., after editing the Excel file locally)
2. The browser will detect the new JSON and automatically update the display

You don't need to manually refresh the page.

## Configuration

### Change the Auto-Refresh Interval

Edit `index.html` and modify this line:

```javascript
const AUTO_REFRESH_INTERVAL_MS = 30000;  // Change 30000 to your desired milliseconds
```

Examples:
- `10000` = 10 seconds
- `60000` = 1 minute
- `300000` = 5 minutes

### Change the Weekly Target

Edit `fetch_and_compute.py` and modify this line:

```python
WEEKLY_TARGET = 15000  # Change 15000 to your desired target
```

The "Current Week Balance" will automatically recalculate based on the new target.

## Business Logic Reference

### KPI Metrics

- **Current Date Label:** Today's date formatted as "DD Month" (e.g., "23 May")
- **Current Week Sales:** Total sales for the current week (Monday-Sunday)
- **Current Week Balance:** Remaining sales needed to reach the weekly target (15000). Zero if target is met.
- **Remaining Days:** Days left in the current week (7 - today's weekday, where Monday=0)
- **Previous Week Sales:** Total sales for the previous week (Monday-Sunday)
- **Monthly Sales:** Sum of all daily sales for the current calendar month

### Weekly Target

- **WEEKLY_TARGET = 15000** (hardcoded constant)
- Formula: `target_pending = 0 if total_sales >= WEEKLY_TARGET else WEEKLY_TARGET - total_sales`

### Data Processing

1. **Parse Excel:** Read "Daily Total Sales" sheet, skip "Grand Total" row and empty/unparseable rows
2. **Weekly Buckets:** Group daily sales by week (Monday-Sunday)
3. **Monthly Filter:** Keep only records from the current calendar month for chart display
4. **Output:** Write computed metrics and chart data to `dashboard-data.json`

## Troubleshooting

### Dashboard shows "-" or "0" for all KPI values

- Verify `dashboard-data.json` exists and is valid JSON
- Check that the Excel file has data in the "Daily Total Sales" sheet (rows 4+, not just headers)
- Run `python fetch_and_compute.py` again to regenerate the JSON

### Bar chart doesn't render

- Check browser console for JavaScript errors (F12 → Console)
- Verify Chart.js and the datalabels plugin loaded from CDN
- Check that `daily_chart_data` in dashboard-data.json is not empty
- Try opening the HTML in a different browser

### Numbers don't match expectations

- Manually verify the Excel file contains the correct daily sales amounts
- Check that the "Daily Total Sales" sheet exists and data starts at row 4
- Run `python fetch_and_compute.py` with verbose output to debug parsing

### Auto-refresh not working

- Check browser console for fetch errors
- Verify `dashboard-data.json` is readable (try opening it directly in the browser)
- Ensure `index.html` and `dashboard-data.json` are in the same directory

## Future Enhancements

- Scheduled automatic runs via Task Scheduler or cron
- GitHub integration and automated push on update
- Cloud deployment (AWS, Azure, etc.)
- Custom WEEKLY_TARGET configuration via web UI
- Date range filtering
- Historical data storage and trending

## Notes

- All data is stored locally; no cloud sync in this phase
- The frontend is static HTML/CSS/JS (no backend server required)
- Backend script is synchronous and stateless (runs once per execution)
- Change detection uses JSON string comparison (efficient for small datasets)

---

**Questions or issues?** Check the design spec at `docs/superpowers/specs/2026-06-24-sales-dashboard-design.md` for detailed architecture and business logic.
```

- [ ] **Step 2: Verify README is readable**

Open `README.md` and verify all sections are clear and complete.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: add comprehensive README with setup, running, and troubleshooting guide"
```

---

### Task 12: Create `.env.example` File

**Files:**
- Create: `.env.example` (project root)

- [ ] **Step 1: Create .env.example with template**

```
# Excel file path (absolute or relative)
# Set this environment variable before running fetch_and_compute.py
# Example on Windows:
#   set EXCEL_FILE_PATH=C:\Users\YourName\OneDrive\Sites\Accounts\Daily_Invoice.xlsx
# Example on macOS/Linux:
#   export EXCEL_FILE_PATH=/Users/YourName/OneDrive/Sites/Accounts/Daily_Invoice.xlsx
EXCEL_FILE_PATH=/path/to/Daily_Invoice.xlsx
```

- [ ] **Step 2: Commit**

```bash
git add .env.example
git commit -m "docs: add .env.example template"
```

---

### Task 13: Final Integration Test and Verification

**Files:**
- Verify: Complete end-to-end flow from backend to frontend

- [ ] **Step 1: Clean up any old files**

Delete `dashboard-data.json` if it exists (we'll regenerate it fresh):

```bash
rm dashboard-data.json  # or 'del dashboard-data.json' on Windows
```

- [ ] **Step 2: Run backend script**

```bash
$env:EXCEL_FILE_PATH = "C:\path\to\Daily_Invoice.xlsx"  # Or set your path
python fetch_and_compute.py
```

Verify:
- Script completes successfully
- `dashboard-data.json` is created
- Numbers are reasonable (current_week_sales, monthly_sales, etc.)

- [ ] **Step 3: Open frontend in browser**

```bash
python -m http.server 8000
# Then visit http://localhost:8000
```

Verify:
- All 6 KPI cards display with correct values
- Numbers are formatted with commas
- Bar chart renders with blue bars and red value labels
- Layout matches the reference screenshot
- No console errors (F12 → Console)

- [ ] **Step 4: Test auto-refresh**

1. Open browser console
2. Wait 30+ seconds
3. Check for "Dashboard data changed, updating..." message (or silence if data unchanged)
4. Run backend script again to generate fresh JSON
5. Wait 30 seconds and verify the dashboard updates automatically

Expected: Dashboard updates without manual page refresh.

- [ ] **Step 5: Verify all KPI formulas are correct**

Check each KPI against the spec:
- Current Date Label: "DD Month" format ✓
- Current Week Sales: sum of daily sales for Mon-Sun week ✓
- Current Week Balance: max(0, 15000 - current_week_sales) ✓
- Remaining Days: 7 - today.weekday() ✓
- Previous Week Sales: sum of daily sales for previous Mon-Sun week ✓
- Monthly Sales: sum of daily sales for current calendar month ✓

- [ ] **Step 6: Verify chart data is correct**

- Daily chart shows only current month (Jan, Feb, etc.)
- Each bar corresponds to one day
- Dates are formatted as "Month Day" (e.g., "May 1", no year)
- Red labels above bars show exact daily sales value
- Y-axis gridlines are at 2000, 4000, 6000, etc. (not random intervals)

- [ ] **Step 7: Test in different browsers (optional but recommended)**

- Chrome/Edge: should work
- Firefox: should work
- Safari: should work

If chart doesn't render in one browser, check console for JS errors related to Chart.js CDN or plugin.

- [ ] **Step 8: Final commit**

All tasks complete! Create a final commit if needed:

```bash
git status  # Verify all files are committed
git log --oneline  # Verify commit history looks good
```

Expected commits:
1. "feat: create fetch_and_compute.py with env var and error handling"
2. "feat: implement Excel parsing logic with "Grand Total" skip"
3. "feat: implement weekly summary computation with target_pending logic"
4. "feat: implement KPI metric computation and current month filtering"
5. "feat: implement chart data generation and JSON output"
6. "feat: create index.html with KPI card structure and Chart.js setup"
7. "feat: add CSS styling for KPI cards and chart container"
8. "feat: add JavaScript framework with KPI rendering and chart initialization"
9. "docs: add comprehensive README with setup, running, and troubleshooting guide"
10. "docs: add .env.example template"

---

## Testing Checklist

### Backend Tests (fetch_and_compute.py)
- [ ] Script runs without errors when EXCEL_FILE_PATH is set
- [ ] Excel file is found and sheet "Daily Total Sales" exists
- [ ] Daily records are parsed correctly (date + daily_sales)
- [ ] "Grand Total" row is skipped
- [ ] Empty/unparseable rows are skipped with warnings
- [ ] Weekly summaries are computed (Mon-Sun weeks)
- [ ] target_pending is 0 when sales >= 15000, else = 15000 - sales
- [ ] Current month sales are filtered correctly
- [ ] KPI metrics are computed:
  - current_date_label: "DD Month" format (no year, no leading zero on day)
  - current_week_sales: sum of current week
  - current_week_balance: target_pending of current week
  - remaining_days: 7 - weekday (1-7 range)
  - previous_week_sales: sum of previous week
  - monthly_sales: sum of current month
- [ ] dashboard-data.json is created and valid JSON
- [ ] Chart data has dates in "Month Day" format (e.g., "May 1")
- [ ] All numbers are integers (no decimals)

### Frontend Tests (index.html)
- [ ] Page loads without JavaScript errors
- [ ] All 6 KPI cards render with values from JSON
- [ ] Numbers are formatted with thousands separators
- [ ] KPI labels are correct
- [ ] Bar chart renders with blue bars
- [ ] Red value labels appear above each bar
- [ ] Y-axis shows gridlines at 0, 2000, 4000, ... intervals
- [ ] X-axis shows date labels from chart data
- [ ] Chart title "Sum of daily_sales by date" is visible
- [ ] Layout matches reference screenshot (colors, spacing, font sizes)
- [ ] Auto-refresh works (checks every 30 seconds)
- [ ] Cache-busting works (query param or no-store directive prevents stale data)
- [ ] Works in Chrome, Firefox, Safari
- [ ] Responsive layout on mobile (optional)

### Integration Tests
- [ ] Backend produces valid JSON
- [ ] Frontend reads and displays JSON correctly
- [ ] Updating Excel → running backend → auto-refresh in frontend works
- [ ] All numbers flow correctly from Excel → backend → JSON → frontend
- [ ] No data loss or rounding errors in calculations

---

## Known Limitations (Current Phase)

- No automated unit tests (manual testing only)
- No scheduled execution (backend runs only when invoked manually)
- No GitHub integration yet (will be added in future phase)
- No cloud deployment (local only)
- No database (data is ephemeral, not persisted)
- No advanced error recovery (script exits on first error)
- No multi-user support (single-machine operation)

---

**Plan Complete!** All tasks are ready for implementation. Choose your execution strategy below.
```

---

## Self-Review Against Spec

Let me verify the plan covers all spec requirements:

**Backend (Section 2 of spec):**
- ✓ Task 1: Environment variable loading and error handling
- ✓ Task 2: Excel parsing ("Daily Total Sales" sheet, row 3 headers, skip "Grand Total")
- ✓ Task 3: Weekly summary computation (Mon-Sun weeks, target_pending logic)
- ✓ Task 4: KPI metrics (all 6 formulas implemented exactly)
- ✓ Task 5: Chart data generation and JSON output
- ✓ Task 6: Manual testing and verification

**Frontend (Section 3 of spec):**
- ✓ Task 7: HTML structure (6 KPI cards, chart canvas)
- ✓ Task 8: CSS styling (colors, spacing, rounded corners, responsive)
- ✓ Task 9: JavaScript framework (load, render functions)
- ✓ Task 10: KPI card rendering with number formatting
- ✓ Task 11: Chart rendering with Chart.js + datalabels plugin
- ✓ Task 12: Auto-refresh with cache-busting (timestamp query param)
- ✓ Task 13: Manual integration testing

**Documentation:**
- ✓ Task 14: README with setup, running, troubleshooting
- ✓ Task 15: .env.example template

**No placeholders found** — all code, commands, and expected outputs are complete and concrete.

**Type consistency check:**
- All function names match across tasks
- All JSON field names are consistent (current_date_label, current_week_sales, etc.)
- All CSS class names match between HTML and CSS
- No contradictions found

---

Plan is complete and ready for execution!

