"""
fetch_cost.py
-------------
Reads your ACTUAL API cost per day from the Claude Console "Cost" page using the
SAME login as fetch_billing.py (no API key, no account change), and merges it into
cost_data.json (amounts in DOLLARS):

    { "2026-11-18": 0.2094, "2026-11-17": 0.1367, ... }   # one entry per day

It only READS the Cost page - it does not call Claude and does not spend credits.
Each run captures roughly the last month and MERGES it in, so history builds up
over time and older days are never lost.

Laptop mode uses the signed-in Edge profile; cloud mode uses the session from the
GitHub Secret BILLING_STORAGE_STATE. Run:  python fetch_cost.py
"""
import json
import os
import sys
import time
from collections import defaultdict
from pathlib import Path

# reuse the browser/login helpers from the balance script
import fetch_billing as fb

HERE = Path(__file__).resolve().parent
DATA_FILE = HERE / "cost_data.json"
COST_URL = "https://platform.claude.com/cost"
CAPTURE_WAIT_SECONDS = 20


def _capture_daily_rows(page) -> list:
    """Load the Cost page and grab the daily (1d) cost breakdown it fetches."""
    grabbed = {"rows": None}

    def on_response(resp):
        try:
            if ("analytics/cost" in resp.url and "time_bucket_width=1d" in resp.url
                    and resp.status == 200):
                data = resp.json().get("data", [])
                # keep the richest response if several arrive
                if grabbed["rows"] is None or len(data) > len(grabbed["rows"]):
                    grabbed["rows"] = data
        except Exception:
            pass

    page.on("response", on_response)
    page.goto(COST_URL, wait_until="domcontentloaded")
    deadline = time.time() + CAPTURE_WAIT_SECONDS
    while time.time() < deadline:
        if grabbed["rows"] is not None:
            page.wait_for_timeout(1500)  # let a richer response replace it
            break
        page.wait_for_timeout(1000)
    return grabbed["rows"] or []


def fetch_cost_rows() -> list:
    from playwright.sync_api import sync_playwright

    state = os.environ.get("BILLING_STORAGE_STATE")
    with sync_playwright() as p:
        if state:  # CLOUD
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(storage_state=json.loads(state),
                                      viewport={"width": 1280, "height": 900})
            rows = _capture_daily_rows(ctx.new_page())
            browser.close()
        else:      # LAPTOP (Edge profile)
            if not fb.PROFILE_DIR.exists():
                raise fb.NotLoggedIn("Not signed in yet.")
            fb.stop_profile_edge()
            ctx = fb.open_edge(p, headless=True)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            rows = _capture_daily_rows(page)
            ctx.close()
    return rows


def load_existing() -> dict:
    if DATA_FILE.exists():
        raw = DATA_FILE.read_text(encoding="utf-8").strip()
        if raw:
            return json.loads(raw)
    return {}


def main():
    try:
        rows = fetch_cost_rows()
    except fb.NotLoggedIn as e:
        print(f"{e} Cost not updated this run.", flush=True)
        return

    per_day = defaultdict(float)
    for r in rows:
        day = r["starting_at"][:10]
        per_day[day] += float(r["cost"]["amount"]) / 100.0   # cents -> dollars

    if not per_day:
        print("No cost rows captured this run; leaving cost_data.json unchanged.", flush=True)
        return

    data = load_existing()
    for day, amt in per_day.items():
        data[day] = round(amt, 6)            # overwrite the captured days, keep older ones
    data = {d: data[d] for d in sorted(data)}

    tmp = DATA_FILE.with_name(DATA_FILE.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=1) + "\n", encoding="utf-8")
    os.replace(tmp, DATA_FILE)

    by_month = defaultdict(float)
    for d, a in data.items():
        by_month[d[:7]] += a
    print(f"Saved {len(data)} days to {DATA_FILE.name}")
    for m in sorted(by_month):
        print(f"  {m}: ${by_month[m]:.2f}")
    print(f"Total: ${sum(by_month.values()):.2f}")


if __name__ == "__main__":
    main()
