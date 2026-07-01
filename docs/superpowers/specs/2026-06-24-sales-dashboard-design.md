# Sales Dashboard Design Spec
**Date:** 2026-06-24  
**Status:** Design phase  
**Scope:** Full-stack local sales dashboard (Python backend + HTML/CSS/JS frontend)

---

## 1. Overview

Build a sales dashboard that replicates an existing Power BI dashboard 1-to-1, reading live data from a locally synced SharePoint Excel file ("Daily_Invoice.xlsx", sheet "Daily Total Sales"). The dashboard computes business metrics, displays KPI cards, and visualizes daily sales trends for the current month.

**Architecture:** Two-part local stack
- **Backend:** Simple Python script that reads Excel, computes metrics, writes JSON
- **Frontend:** Static HTML/CSS/JS that reads JSON and renders dashboard

**Data source:** Local file synced via OneDrive (user manages sync; backend reads from `EXCEL_FILE_PATH` env var)

---

## 2. Backend: `fetch_and_compute.py`

### 2.1 Input and Configuration

**Environment Variable:**
- `EXCEL_FILE_PATH` — absolute or relative path to the local Excel file (`Daily_Invoice.xlsx`)
- Example: `EXCEL_FILE_PATH=/Users/me/OneDrive/Sites/Accounts2/Daily_Invoice.xlsx`

**Constants:**
- `WEEKLY_TARGET` = 15000 (hardcoded)
- `OUTPUT_FILE` = `dashboard-data.json` (relative to project root)

### 2.2 Excel Parsing Logic

**Input sheet:** "Daily Total Sales"

**Data structure:**
- Row 3 contains headers: "Row Labels" (column A, date) and "Sum of NET_AMOUNT" (column B, numeric value)
- Rows 4 onward contain daily records: each row = {date (A), daily_sales (B)}
- **Last row is "Grand Total"** — skip this row (do not include in any calculations)
- Skip rows where date is empty/null or daily_sales is null/unparseable
- After skipping, data is sorted by date ascending

**Parsing approach:**
```python
import openpyxl
from datetime import datetime

# Open workbook, select "Daily Total Sales" sheet
ws = wb["Daily Total Sales"]

# Parse rows starting from 4 (row 3 is header)
daily_records = []
for row in ws.iter_rows(min_row=4, values_only=True):
    date_value, sales_value = row[0], row[1]
    
    # Skip if date is None or if it's the Grand Total row
    if date_value is None or (isinstance(date_value, str) and date_value.lower() == "grand total"):
        continue
    
    # Ensure date_value is a datetime object (openpyxl returns datetime for cells with date format)
    if isinstance(date_value, datetime):
        date_obj = date_value
    else:
        # Try parsing as string (fallback)
        try:
            date_obj = datetime.strptime(str(date_value), "%Y-%m-%d")
        except:
            continue  # Skip unparseable rows
    
    # Ensure sales_value is numeric
    if sales_value is None:
        continue
    try:
        sales_num = float(sales_value)
    except:
        continue  # Skip unparseable rows
    
    daily_records.append({"date": date_obj.date(), "daily_sales": sales_num})

# Sort by date ascending
daily_records.sort(key=lambda x: x["date"])
```

### 2.3 Business Logic: KPI Computation

All formulas must match the original PySpark pipeline exactly.

#### Weekly Summary Logic

**Week definition:** Monday (weekday 0) through Sunday (weekday 6)

For each date in `daily_records`:
1. Compute `week_start_date` = the Monday of that date's week
   - Formula: `date - timedelta(days=date.weekday())`
2. Compute `week_end_date` = `week_start_date + timedelta(days=6)`
3. `total_sales` = sum of all `daily_sales` rows where date falls in [week_start_date, week_end_date]
4. `target_pending` = 0 if `total_sales >= WEEKLY_TARGET` else `WEEKLY_TARGET - total_sales`

**Output:** A dict keyed by `week_start_date`:
```python
weekly_summary = {
    date(2026, 5, 19): {"week_start_date": date(2026, 5, 19), "week_end_date": date(2026, 5, 25), "total_sales": 17629, "target_pending": 0},
    date(2026, 5, 12): {"week_start_date": date(2026, 5, 12), "week_end_date": date(2026, 5, 18), "total_sales": 29588, "target_pending": 0},
    # ... one entry per unique week in the data
}
```

#### Current Month Daily Sales Logic

Filter `daily_records` to include only rows where the date's year-month matches today's year-month. Keep date + daily_sales, sorted by date ascending.

```python
today = datetime.now().date()
current_month_sales = [
    record for record in daily_records
    if record["date"].year == today.year and record["date"].month == today.month
]
# Already sorted by parsing logic
```

#### Dashboard Metrics (Computed Fresh on Each Run)

All metrics are computed from the above aggregations.

**1. Current Date Label**
```
Format: "<day> <month>"
Example: "23 May"
Code: f"{today.day} {today.strftime('%B')}"
```

**2. Current Week Sales**
```
Condition: Identify the most recent week_start_date <= today
If found: total_sales of that week
If not found (no data for current week yet): 0
```

**3. Current Week Balance**
```
Condition: Same week as Current Week Sales
If week found: target_pending from that week's summary
If not found: 0
```

**4. Remaining Days**
```
Formula: 7 - today.weekday()
Example: today = Wednesday (weekday 2) → remaining = 7 - 2 = 5
Note: This counts today as a remaining day if today is Mon-Sun
```

**5. Previous Week Sales**
```
Condition: Identify the second-most-recent week_start_date < most recent week
If found: total_sales of that week
If not found: 0
```

**6. Monthly Sales**
```
Formula: Sum of daily_sales for all rows in current_month_sales
Example: If May 2026 has daily sales [2739, 1562, 1029, ...], monthly = sum of all
```

### 2.4 Output Format: `dashboard-data.json`

```json
{
  "current_date_label": "23 May",
  "current_week_sales": 17629,
  "current_week_balance": 0,
  "remaining_days": 1,
  "previous_week_sales": 29588,
  "monthly_sales": 69849,
  "daily_chart_data": [
    {"date": "May 1", "daily_sales": 2739},
    {"date": "May 2", "daily_sales": 1562},
    {"date": "May 3", "daily_sales": 1029},
    ...
    {"date": "May 22", "daily_sales": 12380}
  ]
}
```

**Notes:**
- `daily_chart_data` contains only the current month's daily sales, sorted by date ascending
- Date format in chart data: "Month Day" (e.g., "May 1", "May 22")
- Numbers are integers (no decimal places for sales figures)
- All metrics are integers or formatted strings

### 2.5 Error Handling

- If `EXCEL_FILE_PATH` env var is not set: print clear error message and exit with code 1
- If file does not exist: print error and exit with code 1
- If "Daily Total Sales" sheet does not exist: print error and exit with code 1
- If parsing fails (corrupted data, wrong format): skip the row with a warning, continue processing
- If no valid data rows are found: write an empty/zero-filled JSON and print warning
- Always write output JSON even if there are warnings (frontend can handle zero/empty state)

---

## 3. Frontend: `index.html`

### 3.1 Layout and Components

**Overall structure:**
- Single HTML file: `index.html`
- Inline CSS (or single `<style>` tag)
- Inline JavaScript (or single `<script>` tag)
- External dependencies: Chart.js + chartjs-plugin-datalabels (loaded via CDN)

**Section 1: KPI Cards Row**
- 6 cards side by side, equal width, responsive (flex layout)
- Background: light gray (#f7f7f7)
- Rounded corners: ~10-12px
- Padding: ~24px
- Gap between cards: ~16px
- Card height: ~210px (fixed or natural)
- Font: sans-serif (system default: Segoe UI, -apple-system, Arial)

**Each KPI card contains:**
- Top element: Large bold number (~32-34px, color black, font-weight 700)
- Bottom element: Small gray label (~13-14px, color #666, font-weight 400)
- **Numbers are formatted with thousands separators** (e.g., "17,629" not "17629")
- **Date is formatted as "23 May"** (no year, no leading zero)

**Cards (in order):**
1. Current Date Label: `current_date_label` (string, displayed as-is)
2. Current Week Sales: `current_week_sales` (number, formatted with commas)
3. Current Week Balance: `current_week_balance` (number, formatted with commas)
4. Remaining Days: `remaining_days` (number, no commas needed)
5. Previous Week Sales: `previous_week_sales` (number, formatted with commas)
6. Monthly Sales: `monthly_sales` (number, formatted with commas)

**Section 2: Bar Chart Card**
- Title: "Sum of daily_sales by date" (small dark gray text, positioned top-left)
- Full width, sits below KPI row with small gap (~16px)
- Background: white
- Border: thin light gray (~1px, #e0e0e0)
- Rounded corners: ~8px
- Padding: ~24px
- Minimum height: ~400px

**Chart specifics:**
- **Library:** Chart.js (via CDN) + chartjs-plugin-datalabels plugin (via CDN)
- **Type:** Bar chart (vertical bars)
- **X-axis:** 
  - Label: "date"
  - Ticks: one per day in `daily_chart_data`
  - Tick labels: day labels from data (e.g., "May 1", "May 2", etc.)
- **Y-axis:**
  - Label: "Sum of daily_sales"
  - Scale: 0 to (max value in data + 10% padding)
  - Gridlines: light dotted, every 2000 units
  - Ticks: 0, 2000, 4000, 6000, ... up to max
  - **Chart.js config:** `ticks: { stepSize: 2000 }` (explicit step size to ensure gridlines match screenshot)
- **Bars:**
  - Color: solid blue (#1ca0f2)
  - No border
  - No gradient
  - Width: proportional, consistent gaps between bars
- **Data Labels (chartjs-plugin-datalabels):**
  - **Required:** Display exact value above each bar
  - Color: red/coral (#f17777)
  - Font: bold (~12px)
  - Anchor: top (label positioned above bar)
  - Align: center
  - Offset: +5px (so text sits above bar, not touching)
  - Display: always (even for small values like "110")

### 3.2 Data Loading and Refresh Logic

**Initial load:**
1. Fetch `dashboard-data.json` with cache-busting (see below)
2. Parse JSON
3. Render KPI cards
4. Render bar chart
5. Set up auto-refresh timer

**Cache-busting for fetch:**
```javascript
// Option A: Timestamp query param
fetch(`dashboard-data.json?t=${Date.now()}`)

// Option B: no-store cache directive
fetch('dashboard-data.json', { cache: 'no-store' })

// Recommendation: Use Option A (broader browser/proxy support)
```

**Auto-refresh timer:**
- Interval: 30 seconds (configurable constant at top of JS)
- On each interval: fetch fresh JSON and re-render if data changed
- Comparison: do a deep equality check (JSON.stringify) to detect changes
- If changed: re-render KPI cards + re-render/update chart

**Error handling:**
- If fetch fails: log error to console, do not crash; try again on next interval
- If JSON is malformed: log error, do not crash
- If JSON is missing fields: use zero/empty fallback values

### 3.3 Styling

**Design system:**
- Font: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif
- Color palette:
  - Primary blue (bars): #1ca0f2
  - Label red: #f17777
  - Card background: #f7f7f7
  - Card text: black (#000) for numbers, #666 for labels
  - Borders: #e0e0e0 (light gray)
  - Background (page): white (#fff)
- Spacing: 16px gaps (cards, chart), 24px padding (cards/chart)
- Rounded corners: 12px (KPI cards), 8px (chart card)
- No shadows, flat minimal design

**Responsive:**
- KPI cards: flex row, wrap on mobile (optional — may keep as 6 cols for simplicity)
- Chart: full width, scales down gracefully
- Base font size: 16px (root)
- Padding/margins: scale appropriately on smaller screens

---

## 4. Data Flow Diagram

```
┌─────────────────────────────────┐
│  Local Excel (OneDrive synced)  │
│  Daily_Invoice.xlsx             │
│  "Daily Total Sales" sheet      │
└──────────────┬──────────────────┘
               │
               │ (user runs: python fetch_and_compute.py)
               │
┌──────────────▼─────────────────┐
│  fetch_and_compute.py           │
│  - Read Excel                   │
│  - Parse "Daily Total Sales"    │
│  - Compute weekly summaries     │
│  - Compute KPI metrics          │
│  - Generate chart data          │
└──────────────┬──────────────────┘
               │
               │ (writes)
               │
┌──────────────▼─────────────────┐
│  dashboard-data.json            │
│  (project root)                 │
│  {                              │
│    "current_date_label": "...", │
│    "current_week_sales": ...,   │
│    "daily_chart_data": [...]    │
│  }                              │
└──────────────┬──────────────────┘
               │
               │ (browser loads, JS fetches with cache-busting)
               │
┌──────────────▼──────────────────┐
│  index.html (static)             │
│  - Fetch dashboard-data.json     │
│  - Render 6 KPI cards            │
│  - Render bar chart (Chart.js)   │
│  - Auto-refresh every 30s        │
└──────────────────────────────────┘
```

---

## 5. Configuration and Constants

**Backend (`fetch_and_compute.py`):**
```python
WEEKLY_TARGET = 15000  # Hardcoded constant
OUTPUT_FILE = "dashboard-data.json"  # Relative to project root
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")  # From env var
```

**Frontend (`index.html`):**
```javascript
const AUTO_REFRESH_INTERVAL_MS = 30000;  // 30 seconds, user can adjust
const CHART_DATA_URL = "dashboard-data.json";  // Relative path
```

---

## 6. Error Handling and Edge Cases

**Backend:**
- Missing env var: exit with clear error message
- Missing file: exit with clear error message
- Missing sheet: exit with clear error message
- Corrupted rows: skip with warning, continue
- No data: output empty/zero-filled JSON
- File locked/unreadable: exit with clear error

**Frontend:**
- Fetch fails: log, retry on next interval
- JSON malformed: log, use default/zero values
- Missing fields in JSON: use default/zero values
- Chart.js not loaded: display fallback (table or error message)
- Very large numbers: still format with commas, let CSS/font handle overflow

---

## 7. Testing Approach (Local)

**Manual testing (no automated tests at this phase):**
1. Set `EXCEL_FILE_PATH` to the synced file location
2. Run `python fetch_and_compute.py`
3. Verify `dashboard-data.json` exists and contains expected data
4. Open `index.html` in a browser (file:// or via simple HTTP server)
5. Verify KPI cards display correct numbers and formatting
6. Verify bar chart renders with all bars and red value labels visible
7. Manually edit Excel, run script again, refresh page to verify auto-refresh works
8. Check that numbers, formatting, and layout match the reference screenshot exactly

---

## 8. Deliverables (This Phase)

- `fetch_and_compute.py` — backend script, reads Excel, computes metrics, writes JSON
- `index.html` — full frontend (HTML + CSS + JS), reads JSON, renders dashboard
- `dashboard-data.json` — output JSON (generated by script, initially empty or sample)
- `.env.example` or `README.md` — documentation of `EXCEL_FILE_PATH` env var, how to run locally
- Git: commit spec + code, do not push yet (user will push manually after testing)

---

## 9. Known Constraints and Assumptions

- Excel file is synced locally via OneDrive (user manages this)
- "Daily Total Sales" sheet structure is fixed (headers at row 3, data rows 4+, "Grand Total" at end)
- WEEKLY_TARGET is a fixed constant (15000), not data-driven
- Week definition is always Monday-Sunday (ISO weekday 0-6)
- No timezone handling (assume all dates are in local timezone)
- No database, no state persistence (script is stateless)
- Frontend is served locally (file:// or localhost:8000), not over internet
- No user authentication, no login required
- All data is local to one machine; no multi-user sync

---

## 10. Future Enhancements (Out of Scope)

- Task Scheduler / cron job automation (mentioned for future phase)
- GitHub push logic (mentioned for future phase)
- Cloud deployment
- Multi-user sync
- Historical data storage
- Advanced filtering/date range selection
- Custom WEEKLY_TARGET configuration via UI
- Timezone handling

---

**End of Design Spec**
