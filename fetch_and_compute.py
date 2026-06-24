import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl

# Constants
WEEKLY_TARGET = 15000
OUTPUT_FILE = "dashboard-data.json"
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

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

        # Parse Excel
        daily_records = parse_excel(ws)

        if not daily_records:
            print("WARNING: No valid data rows found in Excel file")
            daily_records = []
        else:
            print(f"Parsed {len(daily_records)} daily records from Excel")
            print(f"Date range: {daily_records[0]['date']} to {daily_records[-1]['date']}")

        # Compute weekly summaries
        weekly_summary = compute_weekly_summary(daily_records)
        print(f"Computed {len(weekly_summary)} weekly summaries")

        # Compute KPI metrics
        kpi = compute_kpi_metrics(daily_records, weekly_summary)
        print(f"Current Date: {kpi['current_date_label']}")
        print(f"Current Week Sales: {kpi['current_week_sales']}")
        print(f"Monthly Sales: {kpi['monthly_sales']}")

        wb.close()

    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        exit(1)

if __name__ == "__main__":
    main()
