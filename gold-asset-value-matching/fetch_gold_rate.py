"""
fetch_gold_rate.py

Saves the daily gold rate in Indian rupees per gram for 24K, 22K and 18K gold to
gold-rate-data.json, which the "Gold Rate Trends" chart on daily-dashboard.html reads.

Source: GoldAPI.io (https://www.goldapi.io), an official gold-price API.
    These are international market prices converted to rupees. Jewellers such as
    Tanishq add import duty, GST and their own margin, so their store rates are higher;
    the trend over time moves the same way.

How it works:
    1. Read the saved history from gold-rate-data.json (if there is one).
    2. If the last HISTORY_DAYS days are not all there yet (e.g. the very first run),
       fetch each missing past day once from the historical endpoint.
    3. Fetch today's price and save it as today's (Indian date) point; running again
       on the same day replaces that day's point with the newer price.
    4. Keep the last KEEP_DAYS days and write the JSON file.

Requests: the first run uses about HISTORY_DAYS + 1 requests, then 1 per day: once today's
rate is saved, later runs that day make no request (set GOLD_FORCE_REFRESH=1 to fetch anyway). The free
GoldAPI plan allows 150 per month.

Needs the environment variable GOLDAPI_KEY (a GitHub secret in the workflow; for a local
run put GOLDAPI_KEY=... in a .env file next to this script, never commit it).

Run:  python fetch_gold_rate.py
"""

import json
import os
import sys
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass  # python-dotenv is optional; GitHub Actions passes the key as an environment variable

HERE = Path(__file__).resolve().parent
OUTPUT_FILE = HERE / "gold-rate-data.json"

API_BASE = "https://www.goldapi.io/api/price/XAU/INR"
HISTORY_DAYS = 30        # days shown on the chart
KEEP_DAYS = 90           # days kept in the data file
REQUEST_TIMEOUT_SECONDS = 30
SECONDS_BETWEEN_REQUESTS = 0.5   # GoldAPI allows at most 5 requests per second

GRAMS_PER_TROY_OUNCE = 31.1034768
KARATS = {"24k": 24, "22k": 22, "18k": 18}

IST = timezone(timedelta(hours=5, minutes=30))  # India has no daylight saving


def today_ist():
    return datetime.now(IST).date()


def format_ist_timestamp(moment=None):
    """Same format as the Salon data, e.g. "21 November 2026, 10:00 AM IST"."""
    moment = moment or datetime.now(IST)
    hour = moment.strftime("%I").lstrip("0") or "12"
    return f"{moment.day} {moment.strftime('%B %Y')}, {hour}:{moment.strftime('%M %p')} IST"


def fetch(url, api_key):
    response = requests.get(
        url,
        headers={"x-access-token": api_key, "Accept": "application/json"},
        params={"melt_price": "true"},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if response.status_code != 200:
        raise RuntimeError(f"GoldAPI returned HTTP {response.status_code} for {url}: {response.text[:300]}")
    return response.json()


def karat_prices(payload):
    """Rupees per gram for each karat, from one GoldAPI price response."""
    currency = payload.get("currency")
    if currency and currency != "INR":
        raise RuntimeError(f"Expected prices in INR but GoldAPI returned {currency}")

    melt = payload.get("melt_price_per_gram") or {}
    per_gram_24k = (payload.get("price_per_unit") or {}).get("gram")
    if not per_gram_24k and payload.get("price"):
        per_gram_24k = payload["price"] / GRAMS_PER_TROY_OUNCE  # "price" is per troy ounce
    if not per_gram_24k and not melt:
        raise RuntimeError(f"No price in GoldAPI response: {json.dumps(payload)[:300]}")

    prices = {}
    for key, karat in KARATS.items():
        value = melt.get(key)
        if not value:
            value = per_gram_24k * karat / 24
        prices[key] = round(float(value), 2)
    return prices


def load_history():
    if not OUTPUT_FILE.exists():
        return {}
    with open(OUTPUT_FILE, encoding="utf-8") as f:
        saved = json.load(f)
    return {point["date"]: point for point in saved.get("history", [])}


def main():
    api_key = os.environ.get("GOLDAPI_KEY")
    if not api_key:
        sys.exit("GOLDAPI_KEY is not set (GitHub secret, or a .env file next to this script)")

    history = load_history()
    today = today_ist()

    # Once a day is enough for a daily trend, and keeps within GoldAPI's free 150 requests
    # a month even though the workflow starts every 15 minutes
    have_all_days = all((today - timedelta(days=n)).isoformat() in history for n in range(HISTORY_DAYS))
    if have_all_days and os.environ.get("GOLD_FORCE_REFRESH") != "1":
        print(f"Today's gold rate ({today.isoformat()}) is already saved - no GoldAPI request made.")
        return

    # Fill in any missing past days (only many on the first run)
    for days_ago in range(HISTORY_DAYS - 1, 0, -1):
        day = today - timedelta(days=days_ago)
        if day.isoformat() in history:
            continue
        payload = fetch(f"{API_BASE}/{day.isoformat()}", api_key)
        history[day.isoformat()] = {"date": day.isoformat(), **karat_prices(payload)}
        print(f"Added {day.isoformat()}: 22K Rs. {history[day.isoformat()]['22k']:,.2f}/g")
        time.sleep(SECONDS_BETWEEN_REQUESTS)

    # Today's price
    payload = fetch(API_BASE, api_key)
    history[today.isoformat()] = {"date": today.isoformat(), **karat_prices(payload)}
    print(f"Today {today.isoformat()}: " + ", ".join(
        f"{key.upper()} Rs. {history[today.isoformat()][key]:,.2f}/g" for key in KARATS))

    cutoff = (today - timedelta(days=KEEP_DAYS - 1)).isoformat()
    points = sorted((p for d, p in history.items() if d >= cutoff), key=lambda p: p["date"])

    output = {
        "last_updated": format_ist_timestamp(),
        "source": "GoldAPI.io - international market price in INR",
        "unit": "INR per gram",
        "chart_days": HISTORY_DAYS,
        "history": points,
    }
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2)
        f.write("\n")
    print(f"Wrote {len(points)} days to {OUTPUT_FILE.name}")


if __name__ == "__main__":
    main()
