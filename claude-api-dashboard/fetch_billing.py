"""
fetch_billing.py
----------------
Reads ONLY the "Credit balance" amount from
https://platform.claude.com/settings/billing and appends it (with a timestamp)
to billing_data.json. billing_dashboard.html reads billing_data.json and
redraws the chart when it changes.

Uses Microsoft Edge with its OWN separate profile. You log in there once with
the account that owns the API credits; your normal Edge is not touched and can
stay open.

(Copying the login from your normal Edge profile does not work: Edge encrypts
its cookies with a key tied to its own folder, so a copied login can't be read.)

Run (the only command):
    python fetch_billing.py

It fetches the balance once, saves it in billing_data.json, and ends.
If you are not signed in, it opens a sign-in window by itself.

The login is stored in ".edge_billing_profile" next to this script.
Delete that folder to log out / switch account. Don't share it.
"""

import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

BILLING_URL = "https://platform.claude.com/settings/billing"
HERE = Path(__file__).resolve().parent
HTML_FILE = HERE / "billing_dashboard.html"
DATA_FILE = HERE / "billing_data.json"
# The workflow runs from its own copy of the repo, so it sets BILLING_PROFILE_DIR
# to the login folder in "D:\Claude API Dashboard".
PROFILE_DIR = Path(os.environ.get("BILLING_PROFILE_DIR") or HERE / ".edge_billing_profile")
STORAGE_STATE_FILE = HERE / "session_state.json"  # made by --export-session; never commit it

BALANCE_LINE_RE = re.compile(r"-?\$\s*([\d,]+(?:\.\d{1,2})?)")
# Text shown on the sign-in screen (it is served at the billing URL when logged out).
LOGIN_MARKERS = ("continue with email", "continue with google")
FETCH_WAIT_SECONDS = 45


class NotLoggedIn(Exception):
    pass


def extract_balance(page_text: str):
    """Return the $ amount shown inside the 'Credit balance' box.

    Only the text between the 'Credit balance' heading and the box's
    'Buy credits' button is searched, and the amount must be on its own line.
    While the balance still says 'Loading', this returns None - it must never
    pick up other amounts further down the page (e.g. invoice history).
    """
    start = page_text.lower().find("credit balance")
    if start == -1:
        return None
    amount = None
    for line in page_text[start:].splitlines():
        line = line.strip()
        if line.lower() == "buy credits":  # the button = end of the balance box
            return amount
        match = BALANCE_LINE_RE.fullmatch(line)
        if match and amount is None:
            amount = float(match.group(1).replace(",", ""))
    return None  # box not fully rendered yet


def is_login_screen(page_text: str) -> bool:
    low = page_text.lower()
    return "credit balance" not in low and any(m in low for m in LOGIN_MARKERS)


def open_edge(p, headless: bool):
    try:
        return p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            channel="msedge",
            headless=headless,
            viewport={"width": 1280, "height": 900},
        )
    except Exception as e:
        if "exitCode=21" in str(e) or "has been closed" in str(e):
            sys.exit("The script's Edge window (sign-in window) is still open.\n"
                     "Close that window first, then run the command again.")
        raise


def find_edge() -> str:
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"),
                 os.environ.get("LOCALAPPDATA")):
        if base:
            exe = Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe"
            if exe.exists():
                return str(exe)
    sys.exit("Microsoft Edge (msedge.exe) was not found.")


def login():
    """Open a NORMAL (not automated) Edge window on the script's own profile so
    the user can sign in once. Claude's sign-in refuses automated browsers, so
    Playwright is only used afterwards, to read the page."""
    PROFILE_DIR.mkdir(exist_ok=True)
    subprocess.Popen([find_edge(), f"--user-data-dir={PROFILE_DIR}",
                      "--no-first-run", "--no-default-browser-check", BILLING_URL])
    print("A separate Edge window is open (your normal Edge is not affected).\n"
          "1. Sign in with the account that holds the API credits.\n"
          "2. If that account has several organizations, switch to the right one.\n"
          "3. When the Billing page shows the Credit balance, CLOSE that Edge window.")
    input("Then press Enter here...")
    stop_profile_edge()


def profile_edge_pids():
    """PIDs of Edge processes using the script's own profile folder (never your normal Edge)."""
    ps = ("Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | "
          f"Where-Object {{ $_.CommandLine -like '*{PROFILE_DIR.name}*' }} | "
          "ForEach-Object { $_.ProcessId }")
    out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                         capture_output=True, text=True).stdout
    return [pid for pid in out.split() if pid.isdigit()]


def stop_profile_edge():
    """Edge keeps running in the background after its window is closed. Wait a
    moment for it to exit by itself, then end it - only the script's profile."""
    for _ in range(10):
        if not profile_edge_pids():
            return
        time.sleep(1)
    for pid in profile_edge_pids():
        subprocess.run(["taskkill", "/PID", pid, "/T", "/F"], capture_output=True)
    time.sleep(2)


def _read_balance_from_page(page):
    """Open the billing page and wait for the balance. Returns (balance, logged_out)."""
    page.goto(BILLING_URL, wait_until="domcontentloaded")
    deadline = time.time() + FETCH_WAIT_SECONDS
    balance, logged_out = None, False
    while time.time() < deadline:
        try:
            if "/settings/billing" in page.url:
                text = page.inner_text("body")
                balance = extract_balance(text)
                if balance is not None:
                    break
                if is_login_screen(text):
                    logged_out = True
                    break
        except Exception:
            pass  # page still loading / redirecting
        page.wait_for_timeout(1500)
    return balance, logged_out


def fetch_balance(headless: bool = True) -> float:
    from playwright.sync_api import sync_playwright

    state = os.environ.get("BILLING_STORAGE_STATE")
    if state:
        # CLOUD mode (GitHub): use the session from the GitHub Secret, plain browser.
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(storage_state=json.loads(state),
                                      viewport={"width": 1280, "height": 900})
            balance, logged_out = _read_balance_from_page(ctx.new_page())
            browser.close()
    else:
        # LAPTOP mode: use the signed-in Edge profile on this PC.
        if not PROFILE_DIR.exists():
            raise NotLoggedIn("Not signed in yet.")
        stop_profile_edge()  # a leftover background Edge on our profile would block launch
        with sync_playwright() as p:
            ctx = open_edge(p, headless)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            balance, logged_out = _read_balance_from_page(page)
            ctx.close()

    if logged_out:
        raise NotLoggedIn("The sign-in has expired.")
    if balance is None:
        sys.exit("Could not find the Credit balance on the billing page - will try again next time.")
    return balance


def export_session():
    """Run ONCE on your laptop (after --login) to save the signed-in session to a
    file. Paste that file's contents into the GitHub Secret BILLING_STORAGE_STATE
    so the cloud workflow can read the balance without your laptop."""
    from playwright.sync_api import sync_playwright

    if not PROFILE_DIR.exists():
        sys.exit("Sign in first:  python fetch_billing.py --login")
    stop_profile_edge()
    with sync_playwright() as p:
        ctx = open_edge(p, headless=True)
        state = ctx.storage_state()
        ctx.close()
    STORAGE_STATE_FILE.write_text(json.dumps(state), encoding="utf-8")
    print(f"Saved the session to {STORAGE_STATE_FILE.name}")
    print("Next: open that file, copy ALL of its text, and paste it into the GitHub")
    print("Secret named BILLING_STORAGE_STATE (Settings -> Secrets and variables -> Actions).")
    print("Then DELETE the file - it holds the login and must not be shared or uploaded.")


def load_points() -> list:
    if not DATA_FILE.exists():
        return []
    raw = DATA_FILE.read_text(encoding="utf-8").strip()
    return json.loads(raw) if raw else []


def save_point(balance: float) -> int:
    points = load_points()
    points.append({
        "t": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "balance": round(balance, 2),
    })
    write_file(DATA_FILE, json.dumps(points, indent=1) + "\n")
    return len(points)


def write_file(path: Path, text: str):
    """Write to a temp file, then swap it in - the dashboard never reads a half-written file."""
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def record(balance: float):
    n = save_point(balance)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] Credit balance: ${balance:,.2f}  ({n} readings saved)", flush=True)


def main():
    """Fetch the balance once, save it, and end."""
    try:
        import playwright  # noqa: F401
    except ImportError:
        sys.exit("Playwright is not installed. Run:  pip install playwright")

    arg = sys.argv[1] if len(sys.argv) > 1 else ""
    if arg == "--login":
        login()
        return
    if arg == "--export-session":
        export_session()
        return

    try:
        record(fetch_balance())
    except NotLoggedIn as e:
        if os.environ.get("GITHUB_ACTIONS") == "true":  # cloud: can't sign in there
            sys.exit(f"{e}\nThe session in the GitHub Secret has expired. On your laptop run:\n"
                     "  python fetch_billing.py --export-session\n"
                     "then update the Secret BILLING_STORAGE_STATE with the new file's contents.")
        print(f"{e} Opening the sign-in window...", flush=True)
        login()
        try:
            record(fetch_balance())
        except NotLoggedIn:
            sys.exit("Still not signed in. Run the command again and sign in.")


if __name__ == "__main__":
    main()
