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

        wb.close()

    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        exit(1)

if __name__ == "__main__":
    main()
