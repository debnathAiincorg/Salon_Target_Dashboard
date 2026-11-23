"""
fetch_tanishq_rate.py

Saves Tanishq's gold rate (Rs per gram for 22K, 24K and 18K) to gold-rate-data.json,
which the "Gold Rate Trends" chart on daily-dashboard.html reads.

Source: Goodreturns (www.goodreturns.in) news articles that report Tanishq's daily
gold rate. Tanishq's own website blocks automated readers, so it is not read here.
Goodreturns' robots.txt allows reading its news articles and sitemaps; this script
identifies itself honestly (USER_AGENT), reads only a few pages, and waits between them.

How it works:
    1. Find Goodreturns articles about Tanishq gold rates in its news sitemap (the
       first run also looks back LOOKBACK_DAYS days in the main sitemap).
    2. Read each article not read before, take the 22K rate per gram and its date
       (some articles also give earlier days' rates).
    3. Work out 24K and 18K: Tanishq prices them exactly as 22K x 24/22 and 22K x 18/22,
       e.g. 22K Rs 11,515 -> 24K Rs 12,562 and 18K Rs 9,421. A per-gram 24K/18K rate the
       article states is used instead when it agrees with that.
    4. Safety checks: a rate outside a plausible range, more than MAX_DAY_CHANGE away
       from the nearest saved day, or reported with two different values is skipped
       (and printed as a warning), so a misread never reaches the chart.

Goodreturns does not publish an article every day, so some days have no rate; the
chart shows only the days that were reported, never invented values.

Run:  python fetch_tanishq_rate.py
"""

import html
import json
import re
import sys
import time
import urllib.robotparser
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
OUTPUT_FILE = HERE / "gold-rate-data.json"

ROBOTS_URL = "https://www.goodreturns.in/robots.txt"
NEWS_SITEMAP = "https://www.goodreturns.in/sitemap_news.xml"   # the last ~2 days of news
MAIN_SITEMAP = "https://www.goodreturns.in/sitemap.xml"        # older articles, for the first run
USER_AGENT = "SalonTargetDashboard/1.0 (internal dashboard; +https://github.com/debnathAiincorg/Salon_Target_Dashboard)"
REQUEST_TIMEOUT_SECONDS = 30
SECONDS_BETWEEN_ARTICLES = 3
LOOKBACK_DAYS = 35        # first run: read articles from this many days back
CHART_DAYS = 30           # days shown on the chart
KEEP_DAYS = 120           # days kept in the data file
MAX_DAY_CHANGE = 0.15     # a day-to-day move larger than 15% is treated as a misread
PLAUSIBLE_22K = (3000, 60000)

IST = timezone(timedelta(hours=5, minutes=30))
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august",
          "september", "october", "november", "december"]
MONTH_RE = "(" + "|".join(m.capitalize() for m in MONTHS) + "|" + "|".join(MONTHS) + ")"
OTHER_SELLERS = re.compile(r"\b(IBJA|Malabar|Kalyan|Joyalukkas|Joyallukas|Senco|Bhima|CaratLane|PC Jeweller|MCX)\b", re.I)
GRAM_UNIT = r"(?:per\s+1\s+gram|in\s+1\s+gram|per\s+gram|/\s*g\b|a\s+gram)"
RUPEES = r"(?:Rs\.?|INR|₹)\s*([\d,]+(?:\.\d+)?)"


def format_ist_timestamp(moment=None):
    """Same format as the Salon data, e.g. "21 November 2026, 10:00 AM IST"."""
    moment = moment or datetime.now(IST)
    hour = moment.strftime("%I").lstrip("0") or "12"
    return f"{moment.day} {moment.strftime('%B %Y')}, {hour}:{moment.strftime('%M %p')} IST"


def get(session, url):
    response = session.get(url, timeout=REQUEST_TIMEOUT_SECONDS)
    response.raise_for_status()
    return response.text


def tanishq_articles(sitemap_xml, since):
    """(url, last-modified date) of Tanishq gold-rate articles in a sitemap, newest first."""
    found = []
    for block in re.findall(r"<url>(.*?)</url>", sitemap_xml, re.S):
        loc = re.search(r"<loc>\s*(.*?)\s*</loc>", block)
        if not loc:
            continue
        url = html.unescape(loc.group(1))
        slug = url.rsplit("/", 1)[-1].lower()
        if not url.startswith("https://www.goodreturns.in/") or "tanishq" not in slug or "gold" not in slug:
            continue
        when = re.search(r"<(?:lastmod|news:publication_date)>\s*(\d{4}-\d{2}-\d{2})", block)
        day = date.fromisoformat(when.group(1)) if when else None
        if day is None or day >= since:
            found.append((url, day))
    return sorted(found, key=lambda item: item[1] or date.max, reverse=True)


def article_text(page_html):
    """The article's publish date (India time) and its text, with block breaks kept as ' | '."""
    published = re.search(r'"datePublished"\s*:\s*"([^"]+)"', page_html)
    published_day = datetime.fromisoformat(published.group(1)).astimezone(IST).date() if published else None
    body = re.sub(r"<script.*?</script>|<style.*?</style>", " ", page_html, flags=re.S | re.I)
    body = re.sub(r"<(?:p|div|h\d|li|tr|td|br)[^>]*>", " | ", body, flags=re.I)
    body = html.unescape(re.sub(r"<[^>]+>", " ", body))
    return published_day, re.sub(r"\s+", " ", body)


def to_number(text):
    return float(text.replace(",", ""))


def resolve_day(month_name, day_number, published_day):
    """'November 20' in an article published on published_day -> a date (handles new year)."""
    month = MONTHS.index(month_name.lower()) + 1
    year = published_day.year if month <= published_day.month else published_day.year - 1
    return date(year, month, int(day_number))


def tanishq_sections(text):
    """Pieces of the article about Tanishq: from each 'Tanishq' to the next other seller."""
    for match in re.finditer(r"Tanishq", text, re.I):
        piece = text[match.start(): match.start() + 900]
        other = OTHER_SELLERS.search(piece, 10)
        yield piece[: other.start()] if other else piece


def karat_of(sentence):
    found = re.search(r"\b(24|22|18)\s*(?:-\s*)?(?:carat|karat|kt|k)\b", sentence, re.I)
    return found.group(1) if found else None


def extract_rates(text, published_day):
    """{date: {"22": rate, "24": rate, "18": rate}} from the Tanishq parts of an article.
    A value without its own date belongs to the article's publish date."""
    found = {}

    def add(day, karat, value):
        found.setdefault(day, {}).setdefault(karat, set()).add(value)

    for section in tanishq_sections(text):
        for sentence in re.split(r"(?<=[.;])\s+|\s\|\s", section):
            if "coin" in sentence.lower():
                continue  # coin product prices, not the per-gram rate
            karat = karat_of(sentence)
            # "... Rs 11,515 per gram on November 21, down Rs 45 from Rs 11,560 on November 20"
            dated = list(re.finditer(RUPEES + r"(?:\s*" + GRAM_UNIT + r")?\s+on\s+" + MONTH_RE + r"\s+(\d{1,2})\b", sentence))
            for m in dated:
                add(resolve_day(m.group(2), m.group(3), published_day), karat or "22", to_number(m.group(1)))
            if dated or karat is None:
                continue
            # "... stood at Rs 11,515 per 1 gram" (not "fell by Rs 470 in 1 gram")
            for m in re.finditer(r"(?<!by\s)" + RUPEES + r"\s*" + GRAM_UNIT, sentence):
                if re.search(r"\b(?:by|down|up)\s*$", sentence[: m.start()], re.I):
                    continue
                day = published_day
                date_in_sentence = re.search(MONTH_RE + r"\s+(\d{1,2})\b", sentence[: m.start()])
                if date_in_sentence:
                    day = resolve_day(date_in_sentence.group(1), date_in_sentence.group(2), published_day)
                add(day, karat, to_number(m.group(1)))
                break  # the first per-gram value of the sentence is the rate
    return found


def choose_rates(found, url, warnings):
    """One clear 22K value per day (with 24K/18K), or the day is skipped with a warning."""
    chosen = {}
    for day, karats in found.items():
        values_22 = {v for v in karats.get("22", set()) if PLAUSIBLE_22K[0] <= v <= PLAUSIBLE_22K[1]}
        if len(values_22) != 1:
            if values_22:
                warnings.append(f"{day}: different 22K values {sorted(values_22)} in {url} - skipped")
            continue
        rate_22 = values_22.pop()
        if not PLAUSIBLE_22K[0] <= rate_22 <= PLAUSIBLE_22K[1]:
            warnings.append(f"{day}: 22K Rs {rate_22:,.0f} looks implausible in {url} - skipped")
            continue
        point = {"date": day.isoformat(), "22k": rate_22, "source": url}
        for karat in ("24", "18"):
            derived = round(rate_22 * int(karat) / 22)
            stated = karats.get(karat, set())
            point[f"{karat}k"] = derived
            if len(stated) == 1:
                value = next(iter(stated))
                if abs(value - derived) <= max(2, derived * 0.002):
                    point[f"{karat}k"] = value
                else:
                    warnings.append(f"{day}: stated {karat}K Rs {value:,.0f} differs from 22K x {karat}/22 = Rs {derived:,} - using the calculated value")
        chosen[day.isoformat()] = point
    return chosen


def accept(history, point, warnings):
    """Add a point unless it jumps too far from the nearest saved day."""
    day = point["date"]
    neighbours = sorted(history, key=lambda d: abs(date.fromisoformat(d) - date.fromisoformat(day)).days)
    for other in neighbours:
        if other != day:
            previous = history[other]["22k"]
            if abs(point["22k"] - previous) / previous > MAX_DAY_CHANGE:
                warnings.append(f"{day}: 22K Rs {point['22k']:,.0f} is more than {MAX_DAY_CHANGE:.0%} away from "
                                f"Rs {previous:,.0f} on {other} - skipped as a likely misread")
                return False
            break
    history[day] = point
    return True


def main():
    saved = json.loads(OUTPUT_FILE.read_text(encoding="utf-8")) if OUTPUT_FILE.exists() else {}
    history = {p["date"]: p for p in saved.get("history", []) if "source" in p}  # Tanishq points only
    read_before = set(saved.get("articles_read", []))
    today = datetime.now(IST).date()
    if today.isoformat() in history:
        print(f"Today's Tanishq rate ({today.isoformat()}) is already saved - nothing to do.")
        return

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT
    # Read only what Goodreturns' robots.txt allows
    robots = urllib.robotparser.RobotFileParser()
    robots.parse(get(session, ROBOTS_URL).splitlines())

    def allowed(url):
        return robots.can_fetch(USER_AGENT, url)

    articles = tanishq_articles(get(session, NEWS_SITEMAP), today - timedelta(days=LOOKBACK_DAYS))
    if not history:  # first run: also look back in the main sitemap
        print("First run: looking for older articles...")
        articles += tanishq_articles(get(session, MAIN_SITEMAP), today - timedelta(days=LOOKBACK_DAYS))
    new_articles = [url for url, _ in dict.fromkeys(articles) if url not in read_before]
    new_articles = [url for url in dict.fromkeys(new_articles) if allowed(url)]

    if not new_articles:
        print("No new Tanishq gold-rate article on Goodreturns - nothing changed.")
        return

    warnings, added = [], 0
    for i, url in enumerate(new_articles):
        if i:
            time.sleep(SECONDS_BETWEEN_ARTICLES)
        try:
            published_day, text = article_text(get(session, url))
        except requests.RequestException as e:
            warnings.append(f"could not read {url}: {e}")
            continue
        read_before.add(url)
        if published_day is None:
            warnings.append(f"no publish date in {url} - skipped")
            continue
        rates = choose_rates(extract_rates(text, published_day), url, warnings)
        # the article's own day first, then the earlier days it mentions
        for day in sorted(rates, reverse=True):
            if accept(history, rates[day], warnings):
                added += 1
                print(f"{day}: 22K Rs {rates[day]['22k']:,.0f}, 24K Rs {rates[day]['24k']:,.0f}, "
                      f"18K Rs {rates[day]['18k']:,.0f}  ({url.rsplit('/', 1)[-1][:60]})")

    for warning in warnings:
        print("WARNING:", warning)

    cutoff = (today - timedelta(days=KEEP_DAYS)).isoformat()
    points = sorted((p for d, p in history.items() if d >= cutoff), key=lambda p: p["date"])
    output = {
        "last_updated": format_ist_timestamp() if added else saved.get("last_updated", format_ist_timestamp()),
        "source": "Tanishq gold rate as reported by Goodreturns (goodreturns.in)",
        "unit": "INR per gram",
        "chart_days": CHART_DAYS,
        "history": points,
        "articles_read": sorted(read_before)[-200:],
    }
    OUTPUT_FILE.write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(f"Read {len(new_articles)} new article(s), saved {added} day(s); {len(points)} days in {OUTPUT_FILE.name}.")
    if not points:
        sys.exit("No Tanishq rate could be read from Goodreturns.")


if __name__ == "__main__":
    main()
