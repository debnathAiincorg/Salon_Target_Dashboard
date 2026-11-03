import json
from datetime import datetime, timedelta
from sharepoint_excel import load_daily_records, today_ist, format_ist_timestamp

# Constants
MONTHLY_TARGET = 60000
OUTPUT_FILE = "monthly-dashboard-data.json"

def compute_monthly_summary(daily_records):
    """Compute month-by-month sales totals for the current year, January through the current month.

    Only records whose year matches today's year are included, so the output
    automatically rolls over to the new year's months once the calendar year changes.

    Args:
        daily_records: List of dicts with 'date' (datetime.date) and 'daily_sales'

    Returns:
        List of dicts ordered January -> current month:
        [{"month": "January", "total_sales": 45230}, ...]
    """
    today = today_ist()

    monthly_totals = {month_num: 0 for month_num in range(1, today.month + 1)}
    for record in daily_records:
        date = record["date"]
        if date.year != today.year or date.month > today.month:
            continue
        monthly_totals[date.month] += record["daily_sales"]

    monthly_summary = []
    for month_num in range(1, today.month + 1):
        month_name = datetime(today.year, month_num, 1).strftime("%B")
        monthly_summary.append({
            "month": month_name,
            "total_sales": int(monthly_totals[month_num]),
        })

    return monthly_summary

def compute_year_stats(daily_records, monthly_summary):
    """Compute the yearly chart's stats: year total, average month and daily average.

    The data is refreshed once a day and covers sales up to yesterday, so the daily
    average uses completed days only: this year's sales up to and including
    yesterday, divided by the days from 1 January up to yesterday (0 on 1 January).

    Args:
        daily_records: List of dicts with 'date' (datetime.date) and 'daily_sales'
        monthly_summary: List of dicts with 'month' and 'total_sales' for this year

    Returns:
        Dict with year_total_sales, average_monthly_sales and year_daily_average (ints).
    """
    today = today_ist()
    yesterday = today - timedelta(days=1)
    year_start = today.replace(month=1, day=1)

    year_total = sum(entry["total_sales"] for entry in monthly_summary)
    days_done = (yesterday - year_start).days + 1
    sales_done = sum(r["daily_sales"] for r in daily_records if year_start <= r["date"] <= yesterday)

    return {
        "year_total_sales": int(year_total),
        "average_monthly_sales": int(round(year_total / len(monthly_summary))) if monthly_summary else 0,
        "year_daily_average": int(round(sales_done / days_done)) if days_done > 0 else 0,
    }

def write_monthly_json(year, monthly_summary, year_stats, output_file):
    """Write year-scoped monthly summary to JSON file.

    Args:
        year: int, the year the monthly_summary belongs to
        monthly_summary: List of dicts with 'month' and 'total_sales'
        year_stats: Dict from compute_year_stats
        output_file: Path to output JSON file
    """
    monthly_data = {
        "year": year,
        "monthly_chart_data": monthly_summary,
        "monthly_target": MONTHLY_TARGET,
        **year_stats,
        "last_updated": format_ist_timestamp(),
    }

    with open(output_file, 'w') as f:
        json.dump(monthly_data, f, indent=2)

    print(f"Monthly dashboard data written to {output_file}")

def main():
    """Main entry point for the monthly fetch and compute script."""
    try:
        # Steps 1-3: Fetch Excel from SharePoint and aggregate invoice_transactions by date
        daily_records = load_daily_records()

        # Step 4: Compute month-by-month summary for the current year
        today = today_ist()
        monthly_summary = compute_monthly_summary(daily_records)
        print(f"Computed {len(monthly_summary)} monthly summaries for {today.year}")
        for entry in monthly_summary:
            print(f"  {entry['month']}: {entry['total_sales']}")

        year_stats = compute_year_stats(daily_records, monthly_summary)

        # Step 5: Write output JSON
        write_monthly_json(today.year, monthly_summary, year_stats, OUTPUT_FILE)

        print("SUCCESS: Monthly dashboard data computed and written")

    except RuntimeError as e:
        print(f"ERROR: {e}")
        exit(1)
    except Exception as e:
        print(f"ERROR: Unexpected error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
