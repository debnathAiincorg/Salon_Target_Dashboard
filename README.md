# Sales Dashboard

A full-stack sales dashboard that reads live data from SharePoint via Microsoft Graph API, computes business KPI metrics (weekly sales, targets, monthly totals), and displays them with an interactive bar chart.

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
- **openpyxl** and **requests** libraries: `pip install openpyxl requests`
- **Azure AD app registration** with Microsoft Graph API access (see Setup below)

## Setup

### 1. Register an Azure AD Application

The dashboard fetches `Daily_Invoice.xlsx` from SharePoint using Microsoft Graph API, which requires Azure AD authentication.

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **+ New registration**
   - **Name:** `Salon Target Dashboard` (or your preferred name)
   - **Supported account types:** "Accounts in this organizational directory only"
   - Click **Register**
4. After creation, note your **Tenant ID** and **Application (client) ID** (shown on the Overview page)

### 2. Create a Client Secret

1. In the app registration, go to **Certificates & secrets**
2. Click **+ New client secret**
   - **Description:** `Dashboard access`
   - **Expires:** Choose based on your policy (e.g., 6 months, 2 years)
   - Click **Add**
3. Copy the secret **Value** (not the ID) — this is shown only once
   - **Important:** Save this securely; you'll need it for `.env`

### 3. Grant Microsoft Graph Permissions

1. In the app registration, go to **API permissions**
2. Click **+ Add a permission**
   - Select **Microsoft Graph** → **Application permissions**
   - Search for and select **Sites.Read.All**
   - Click **Add permissions**
3. Click **Grant admin consent for [Your Org]** to approve the permission

**Permissions required:**
- `Sites.Read.All` (application permission) — Read access to SharePoint sites and files

### 4. Create `.env` File

1. In the project root (same directory as `fetch_and_compute.py`), create a file named `.env` (not `.env.example`)
2. Add your Azure credentials:
   ```
   AZURE_TENANT_ID=<your-tenant-id>
   AZURE_CLIENT_ID=<your-client-id>
   AZURE_CLIENT_SECRET=<your-client-secret>
   ```
   - Replace `<your-tenant-id>`, `<your-client-id>`, and `<your-client-secret>` with values from the Azure Portal
3. **Important:** `.env` is listed in `.gitignore` and will never be committed. Keep it secure.

### 5. Install Python Dependencies

```bash
pip install openpyxl requests python-dotenv
```

(The `python-dotenv` package loads `.env` automatically; `openpyxl` and `requests` are used by the script)

## Running the Dashboard

### Step 1: Fetch and Compute Data

Run the backend script to fetch the Excel file from SharePoint, compute KPI metrics, and generate the dashboard JSON:

```bash
python fetch_and_compute.py
```

Expected output:
```
Fetching Daily_Invoice.xlsx from SharePoint...
✓ Successfully fetched and opened Daily_Invoice.xlsx
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
- **"Azure AD credentials not configured"** — Verify `.env` file exists in the project root with `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET`
- **"Failed to authenticate with Azure AD"** — Check that credentials in `.env` are correct; verify the app was granted `Sites.Read.All` permission with admin consent
- **"Failed to resolve SharePoint share link"** — Verify the SharePoint site and file still exist; check that the Graph API can access them
- **"Failed to download file from SharePoint"** — Network or permission issue; try running the script again
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
