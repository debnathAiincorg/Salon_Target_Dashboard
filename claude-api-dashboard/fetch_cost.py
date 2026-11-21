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
LOOKBACK_DAYS = 30   # the API allows at most 31 days per request


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


def fetch_daily_cost(client, start):
    """Dollars per UTC day from `start` until now, and the data's refresh time."""
    page = client.beta.organization.analytics.cost.list(
        starting_at=start,
        ending_at=datetime.now(timezone.utc),
        time_bucket_width="1d",
        order_by="time",
        order="asc",
        limit=100,
    )
    refreshed_at = page.data_refreshed_at
    daily = {}
    for row in page:  # iterating fetches the following pages too
        day = row.starting_at.astimezone(timezone.utc).date().isoformat()
        # amount is a decimal string in cents, e.g. "41280.000000" = $412.80
        daily[day] = daily.get(day, 0.0) + float(row.cost.amount) / 100
    return daily, refreshed_at


def main():
    key = os.environ.get("ANTHROPIC_ADMIN_KEY")
    if not key:
        sys.exit("ANTHROPIC_ADMIN_KEY is not set (GitHub secret, or a .env file next to this script)")

    client = anthropic.Anthropic(api_key=key)
    today = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0)
    start = today - timedelta(days=LOOKBACK_DAYS)

    try:
        daily, refreshed_at = fetch_daily_cost(client, start)
    except (anthropic.AuthenticationError, anthropic.PermissionDeniedError) as e:
        sys.exit(f"The key was refused ({e.status_code}). The cost data needs an Admin API credential of the "
                 "organization that owns the API credits: an Admin API key (sk-ant-admin...) or an "
                 f"organization API key of a user with the admin role.\nDetails: {e.message}")
    except anthropic.NotFoundError as e:
        sys.exit(f"The cost endpoint was not found for this organization (404): {e.message}")

    if refreshed_at is None:
        print("Anthropic has no cost data for this period yet - nothing changed.")
        return

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
    }
    changed = write_if_changed(META_FILE, meta, indent=2) or changed

    month = today.strftime("%Y-%m")
    month_total = sum(v for d, v in costs.items() if d.startswith(month))
    print(f"{len(daily)} days with cost in the last {LOOKBACK_DAYS} days; this month so far ${month_total:,.2f}; "
          f"data refreshed at {meta['data_refreshed_at']}; files {'updated' if changed else 'unchanged'}.")


if __name__ == "__main__":
    main()
