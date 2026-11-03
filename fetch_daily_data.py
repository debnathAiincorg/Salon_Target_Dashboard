import json
from datetime import timedelta
from sharepoint_excel import load_daily_records, today_ist, format_ist_timestamp

# Constants
WEEKLY_TARGET = 15000
WEEKS_IN_CHART = 12
OUTPUT_FILE = "daily-dashboard-data.json"

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
    today = today_ist()
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
    today = today_ist()

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

def compute_daily_averages(daily_records, weekly_summary):
    """Compute heading values and daily averages for the weekly and daily charts.

    The data is refreshed once a day and covers sales up to yesterday, so averages
    use completed days only: sales up to and including yesterday, divided by the
    number of days from the start of the period up to yesterday. Today's partial
    sales are left out. A period with no completed days yet (Monday for the week,
    the 1st for the month) has an average of 0.

    Args:
        daily_records: List of dicts with 'date' and 'daily_sales'
        weekly_summary: Dict keyed by week_start_date

    Returns:
        Dict with week_number, month_label, current/previous week daily averages
        and current/previous month daily averages (all averages rounded ints).
    """
    today = today_ist()
    yesterday = today - timedelta(days=1)

    def sales_between(start, end):
        return sum(r["daily_sales"] for r in daily_records if start <= r["date"] <= end)

    def average(total, days):
        return int(round(total / days)) if days > 0 else 0

    # Current week (Monday to Sunday containing today) and the full week before it
    week_start = today - timedelta(days=today.weekday())
    week_days_done = (yesterday - week_start).days + 1
    previous_week_start = week_start - timedelta(weeks=1)
    previous_week = weekly_summary.get(previous_week_start)

    # Current month and the full month before it
    month_start = today.replace(day=1)
    month_days_done = (yesterday - month_start).days + 1
    previous_month_end = month_start - timedelta(days=1)
    previous_month_start = previous_month_end.replace(day=1)

    return {
        "week_number": today.isocalendar()[1],
        "month_label": today.strftime("%B %Y"),
        "current_week_daily_average": average(sales_between(week_start, yesterday), week_days_done),
        "previous_week_daily_average": average(previous_week["total_sales"] if previous_week else 0, 7),
        "current_month_daily_average": average(sales_between(month_start, yesterday), month_days_done),
        "previous_month_daily_average": average(
            sales_between(previous_month_start, previous_month_end), previous_month_end.day
        ),
    }

def generate_chart_data(current_month_sales):
    """Generate chart data from current month sales records.

    Args:
        current_month_sales: List of dicts with 'date' and 'daily_sales'

    Returns:
        List of dicts with 'date' (formatted string) and 'daily_sales' (int)
    """
    chart_data = []
    for record in current_month_sales:
        date_obj = record["date"]
        # Optional: use "%b" instead of "%B" for abbreviated month names (e.g., "Jun 24" vs "June 24")
        # Abbreviated format may reduce label crowding with a full month of data
        month_name = date_obj.strftime("%B")
        day = date_obj.day
        date_str = f"{month_name} {day}"

        chart_data.append({
            "date": date_str,
            "daily_sales": int(record["daily_sales"])
        })

    return chart_data

def generate_weekly_chart_data(weekly_summary):
    """Generate weekly chart data for the last WEEKS_IN_CHART weeks.

    Covers Monday-to-Sunday weeks ending with the current week. Weeks with
    no sales are included with 0 so the chart has no gaps.

    Args:
        weekly_summary: Dict keyed by week_start_date

    Returns:
        List of dicts with 'week' (e.g., "26 Oct - 1 Nov") and 'weekly_sales' (int),
        oldest week first
    """
    today = today_ist()
    current_week_start = today - timedelta(days=today.weekday())

    chart_data = []
    for i in range(WEEKS_IN_CHART - 1, -1, -1):
        week_start = current_week_start - timedelta(weeks=i)
        week_end = week_start + timedelta(days=6)
        week = weekly_summary.get(week_start)
        label = f"{week_start.day} {week_start.strftime('%b')} - {week_end.day} {week_end.strftime('%b')}"

        chart_data.append({
            "week": label,
            "weekly_sales": int(week["total_sales"]) if week else 0
        })

    return chart_data

def write_dashboard_json(kpi, chart_data, weekly_chart_data, averages, output_file):
    """Write dashboard data to JSON file.

    Args:
        kpi: Dict with KPI metrics (current_date_label, current_week_sales, etc.)
        chart_data: List of dicts with chart data
        weekly_chart_data: List of dicts with weekly chart data
        averages: Dict from compute_daily_averages
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
        "weekly_chart_data": weekly_chart_data,
        "weekly_target": WEEKLY_TARGET,
        **averages,
        "last_updated": format_ist_timestamp(),
    }

    with open(output_file, 'w') as f:
        json.dump(dashboard_data, f, indent=2)

    print(f"Dashboard data written to {output_file}")

def main():
    """Main entry point for the fetch and compute script."""
    try:
        # Steps 1-3: Fetch Excel from SharePoint and aggregate invoice_transactions by date
        daily_records = load_daily_records()

        # Step 4: Compute weekly summaries
        weekly_summary = compute_weekly_summary(daily_records)
        print(f"Computed {len(weekly_summary)} weekly summaries")

        # Step 5: Compute KPI metrics
        kpi = compute_kpi_metrics(daily_records, weekly_summary)
        print(f"Current Date: {kpi['current_date_label']}")
        print(f"Current Week Sales: {kpi['current_week_sales']}")
        print(f"Monthly Sales: {kpi['monthly_sales']}")

        # Step 6: Generate chart data
        chart_data = generate_chart_data(kpi["current_month_sales"])
        weekly_chart_data = generate_weekly_chart_data(weekly_summary)
        averages = compute_daily_averages(daily_records, weekly_summary)

        # Step 7: Write output JSON
        write_dashboard_json(kpi, chart_data, weekly_chart_data, averages, OUTPUT_FILE)

        print("SUCCESS: Dashboard data computed and written")

    except RuntimeError as e:
        print(f"ERROR: {e}")
        exit(1)
    except Exception as e:
        print(f"ERROR: Unexpected error: {e}")
        exit(1)

if __name__ == "__main__":
    main()
