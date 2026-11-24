"""
fetch_cost.py
-------------
Reads the organization's ACTUAL Claude API cost per day from Anthropic's official
Usage and Cost Analytics API and merges it into cost_data.json (amounts in DOLLARS):

    { "2026-11-18": 0.2094, "2026-11-17": 0.1367, ... }   # one entry per UTC day

The Daily usage and Monthly usage pages (usage.html, monthly.html) read that file.
It also writes cost_meta.json with the time Anthropic last refreshed the data, which
those pages show as "Last updated".

No browser and no sign-in: it uses an API key, so it does not expire like a browser login.
It only READS cost data - it does not call Claude and does not spend credits.

Key: an Admin API credential for the organization that owns the API credits (an Admin
API key, or an organization API key of a user with the admin role), in the environment
variable ANTHROPIC_ADMIN_KEY. On GitHub it is the secret ANTHROPIC_ADMIN_KEY; for a
local run put  ANTHROPIC_ADMIN_KEY=...  in a .env file next to this script (git ignores
.env files - never commit the key).

Balance: there is no API for the prepaid credit balance, so each data refresh also adds an
ESTIMATED balance to billing_data.json (shown on the Balance page): the real balance saved in
balance_anchor.json minus the API cost since that moment. After buying credits, put the new
real balance and the time into balance_anchor.json.

Data freshness: Anthropic refreshes this data within about 4 hours (at most 24), so
today's cost fills in during the day. The files are only rewritten when the data has
changed, so running this often does not create a commit every time.

Run:  python fetch_cost.py
"""
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import anthropic

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass  # optional; GitHub Actions passes the key as an environment variable

HERE = Path(__file__).resolve().parent
DATA_FILE = HERE / "cost_data.json"
META_FILE = HERE / "cost_meta.json"
BALANCE_FILE = HERE / "billing_data.json"     # balance readings shown on the Balance page
ANCHOR_FILE = HERE / "balance_anchor.json"    # last real balance the estimate starts from
LOOKBACK_DAYS = 30   # days of daily cost refreshed on each run
HISTORY_START = "2026-01-01"  # first run reads all cost since this day, so the monthly chart has every month
MAX_SPAN_DAYS = 30   # the API allows at most 31 days per request


def load_json(path, default):
    if not path.exists():
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_if_changed(path, data, indent):
    text = json.dumps(data, indent=indent, sort_keys=True) + "\n"
    if path.exists() and path.read_text(encoding="utf-8") == text:
        return False
    path.write_text(text, encoding="utf-8")
    return True


def fetch_cost_rows(client, start, end, bucket):
    """Cost rows from `start` to `end` in `bucket` ("1d" or "1h") time buckets, in windows
    the API accepts, and the data's refresh time. Each row: (bucket start, bucket end, dollars)."""
    rows, refreshed_at = [], None
    window_start = start
    while window_start < end:
        window_end = min(window_start + timedelta(days=MAX_SPAN_DAYS), end)
        page = client.beta.organization.analytics.cost.list(
            starting_at=window_start,
            ending_at=window_end,
            time_bucket_width=bucket,
            order_by="time",
            order="asc",
            limit=1000,
        )
        refreshed_at = page.data_refreshed_at or refreshed_at
        for row in page:  # iterating fetches the following pages too
            # amount is a decimal string in cents, e.g. "41280.000000" = $412.80
            rows.append((row.starting_at, row.ending_at, float(row.cost.amount) / 100))
        window_start = window_end
    return rows, refreshed_at


def update_balance_estimate(hourly_rows, anchor, refreshed_at):
    """Append an estimated balance (last real balance minus the API cost since then) to the
    Balance page's readings, once per data refresh. Returns True if the file changed."""
    anchor_time = datetime.fromisoformat(anchor["as_of"])
    if refreshed_at <= anchor_time:
        return False  # the cost data does not reach past the real balance yet
    spent = sum(dollars for _, bucket_end, dollars in hourly_rows if bucket_end > anchor_time)
    point = {"t": refreshed_at.astimezone(timezone.utc).isoformat(timespec="seconds"),
             "balance": round(anchor["balance"] - spent, 2),
             "estimated": True}

    readings = load_json(BALANCE_FILE, [])
    if readings and datetime.fromisoformat(readings[-1]["t"]) >= refreshed_at:
        return False  # already have a reading for this data refresh (or a newer one)
    if readings and readings[-1].get("balance") == point["balance"]:
        return False  # no change; the page draws the flat line up to the last check itself
    readings.append(point)
    BALANCE_FILE.write_text(json.dumps(readings, indent=1) + "\n", encoding="utf-8")
    print(f"Estimated balance ${point['balance']:,.2f} (${anchor['balance']:,.2f} on {anchor['as_of']} "
          f"minus ${spent:,.4f} API cost since then)")
    return True


def main():
    key = os.environ.get("ANTHROPIC_ADMIN_KEY")
    if not key:
        sys.exit("ANTHROPIC_ADMIN_KEY is not set (GitHub secret, or a .env file next to this script)")

    client = anthropic.Anthropic(api_key=key.strip())
    now = datetime.now(timezone.utc)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    start = today - timedelta(days=LOOKBACK_DAYS)
    # Once, read the whole history (cost_meta.json then records it as done)
    history_done = load_json(META_FILE, {}).get("history_from") == HISTORY_START
    if not history_done:
        start = min(start, datetime.fromisoformat(HISTORY_START).replace(tzinfo=timezone.utc))
    anchor = load_json(ANCHOR_FILE, None)
    # Hourly cost is needed for today (daily buckets leave out the unfinished day) and,
    # for the balance estimate, since the last real balance
    hourly_start = today
    if anchor:
        hourly_start = min(today, datetime.fromisoformat(anchor["as_of"]).replace(minute=0, second=0, microsecond=0))

    try:
        daily_rows, refreshed_daily = fetch_cost_rows(client, start, today, "1d")
        hourly_rows, refreshed_hourly = fetch_cost_rows(client, hourly_start, now, "1h")
    except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
        sys.exit(f"The key was refused ({e.status_code}). The cost data needs an Admin API credential of the "
                 "organization that owns the API credits: an Admin API key (sk-ant-admin...) or an "
                 f"organization API key of a user with the admin role.\nDetails: {e.message}")
    except anthropic.NotFoundError as e:
        sys.exit(f"The cost endpoint was not found for this organization (404): {e.message}")

    refreshed_at = refreshed_hourly or refreshed_daily
    if refreshed_at is None:
        print("Anthropic has no cost data for this period yet - nothing changed.")
        return

    daily = {}
    for bucket_start, _, dollars in daily_rows:
        day = bucket_start.astimezone(timezone.utc).date().isoformat()
        daily[day] = daily.get(day, 0.0) + dollars
    today_cost = sum(dollars for bucket_start, _, dollars in hourly_rows if bucket_start >= today)
    if today_cost:
        daily[today.date().isoformat()] = today_cost

    # Merge: inside the fetched range the API is the source of truth for every day up to
    # the refresh time (a day missing from the reply had no cost); older days are kept.
    costs = load_json(DATA_FILE, {})
    refreshed_day = refreshed_at.astimezone(timezone.utc).date().isoformat()
    for day in list(costs):
        if start.date().isoformat() <= day < refreshed_day and day not in daily:
            del costs[day]
    for day, dollars in daily.items():
        costs[day] = round(dollars, 6)

    changed = write_if_changed(DATA_FILE, costs, indent=1)
    meta = {
        "data_refreshed_at": refreshed_at.astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
        "source": "Anthropic Usage and Cost Analytics API",
        # When this script last asked Anthropic (every run), so the pages can show that the
        # check is running even when Anthropic has no new data
        "checked_at": now.isoformat(timespec="seconds").replace("+00:00", "Z"),
        # Cost is saved for every day since this date
        "history_from": HISTORY_START,
    }
    changed = write_if_changed(META_FILE, meta, indent=2) or changed
    if anchor:
        changed = update_balance_estimate(hourly_rows, anchor, refreshed_at) or changed

    month = today.strftime("%Y-%m")
    month_total = sum(v for d, v in costs.items() if d.startswith(month))
    print(f"{len(daily)} days with cost in the last {LOOKBACK_DAYS} days; this month so far ${month_total:,.2f}; "
          f"data refreshed at {meta['data_refreshed_at']}; files {'updated' if changed else 'unchanged'}.")


if __name__ == "__main__":
    main()
