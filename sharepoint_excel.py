"""Shared helpers for the dashboard scripts: fetch Daily_Invoice.xlsx from SharePoint
via Microsoft Graph API and aggregate its invoice_transactions sheet into daily totals."""
import os
import base64
import io
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl
import requests

# Load .env file if available (local development)
# In CI/GitHub Actions, env vars are injected directly; dotenv is not needed
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Constants
IST = ZoneInfo("Asia/Kolkata")
SHEET_NAME = "invoice_transactions"

# Azure AD configuration
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

# SharePoint configuration
SHAREPOINT_SHARE_LINK = os.getenv("SHAREPOINT_SHARE_LINK")

def today_ist():
    """Return today's date in IST."""
    return datetime.now(IST).date()

def get_access_token():
    """Authenticate via Azure AD client credentials flow and return access token."""
    if not AZURE_TENANT_ID or not AZURE_CLIENT_ID or not AZURE_CLIENT_SECRET:
        raise ValueError(
            "Azure AD credentials not configured. "
            "Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in .env"
        )

    token_url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "grant_type": "client_credentials",
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }

    try:
        response = requests.post(token_url, data=payload, timeout=10)
        response.raise_for_status()
        return response.json()["access_token"]
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Failed to authenticate with Azure AD: {e}")
    except (KeyError, ValueError) as e:
        raise RuntimeError(f"Unexpected response from Azure AD token endpoint: {e}")

def resolve_sharepoint_link_to_driveitem(access_token, share_link):
    """Resolve SharePoint share link to driveItem using Graph API /shares endpoint."""
    # Base64-url encode the share link: prefix "u!", base64-url encode, strip padding
    share_link_bytes = share_link.encode('utf-8')
    encoded = base64.urlsafe_b64encode(share_link_bytes).decode('utf-8').rstrip('=')
    share_id = f"u!{encoded}"

    shares_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    headers = {"Authorization": f"Bearer {access_token}"}

    try:
        response = requests.get(shares_url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Failed to resolve SharePoint share link: {e}")
    except (KeyError, ValueError) as e:
        raise RuntimeError(f"Unexpected response from Graph API shares endpoint: {e}")

def download_file_from_sharepoint(access_token, drive_item):
    """Download file content from SharePoint driveItem."""
    # Use @microsoft.graph.downloadUrl to get direct download link
    if "@microsoft.graph.downloadUrl" not in drive_item:
        raise RuntimeError(
            "Share link resolved but downloadUrl not available. "
            "Check that the file exists and is accessible."
        )

    download_url = drive_item["@microsoft.graph.downloadUrl"]

    try:
        response = requests.get(download_url, timeout=30)
        response.raise_for_status()
        return response.content
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Failed to download file from SharePoint: {e}")

def fetch_excel_from_sharepoint():
    """Fetch Daily_Invoice.xlsx from SharePoint via Graph API and return openpyxl Workbook."""
    # Validate configuration
    if not SHAREPOINT_SHARE_LINK:
        raise ValueError(
            "SharePoint share link not configured. "
            "Please set SHAREPOINT_SHARE_LINK in .env"
        )

    # Step 1: Authenticate
    access_token = get_access_token()

    # Step 2: Resolve share link to driveItem
    drive_item = resolve_sharepoint_link_to_driveitem(access_token, SHAREPOINT_SHARE_LINK)

    # Step 3: Download file content
    file_content = download_file_from_sharepoint(access_token, drive_item)

    # Step 4: Load into openpyxl from bytes
    # read_only streams rows instead of building every cell in memory: much faster
    # and lighter for a sheet that is only read top to bottom once
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_content), read_only=True, data_only=True)
        return wb
    except Exception as e:
        raise RuntimeError(f"Failed to load Excel file into openpyxl: {e}")

def parse_excel(ws):
    """Parse invoice_transactions sheet and aggregate NET_AMOUNT by DATE.

    Reads raw transaction data from invoice_transactions sheet.
    Multiple rows per date are summed into one daily_sales record.

    Args:
        ws: openpyxl worksheet object (invoice_transactions sheet)

    Returns:
        List of dicts with keys 'date' (datetime.date) and 'daily_sales' (float),
        sorted by date ascending. Skips rows where DATE (column C) or
        NET_AMOUNT (column D) is None or unparseable.
    """
    # Aggregate by date: {date_obj: total_net_amount}
    daily_totals = {}

    # Iterate rows starting from row 2 (row 1 is header)
    # Column C (index 2) = DATE, Column D (index 3) = NET_AMOUNT
    for row in ws.iter_rows(min_row=2, values_only=True):
        date_value = row[2] if len(row) > 2 else None
        net_amount_value = row[3] if len(row) > 3 else None

        # Skip if date or net_amount is None
        if date_value is None or net_amount_value is None:
            continue

        # Parse date
        if isinstance(date_value, datetime):
            date_obj = date_value.date()
        else:
            try:
                date_obj = datetime.strptime(str(date_value), "%Y-%m-%d").date()
            except:
                print(f"WARNING: Skipping row with unparseable date: {date_value}")
                continue

        # Parse net amount
        try:
            net_amount = float(net_amount_value)
        except:
            print(f"WARNING: Skipping row with unparseable NET_AMOUNT: {net_amount_value} for date {date_obj}")
            continue

        # Aggregate: sum all net amounts for this date
        if date_obj not in daily_totals:
            daily_totals[date_obj] = 0
        daily_totals[date_obj] += net_amount

    # Convert to daily_records format and sort
    daily_records = [
        {"date": date_obj, "daily_sales": total}
        for date_obj, total in daily_totals.items()
    ]
    daily_records.sort(key=lambda x: x["date"])

    return daily_records

def load_daily_records():
    """Fetch the workbook, check the invoice_transactions sheet and return its daily records.

    Exits with status 1 if the sheet is missing.
    """
    print("Fetching Daily_Invoice.xlsx from SharePoint...")
    wb = fetch_excel_from_sharepoint()

    try:
        if SHEET_NAME not in wb.sheetnames:
            print(f"ERROR: Sheet '{SHEET_NAME}' not found in Excel file")
            exit(1)

        print("[OK] Successfully fetched and opened Daily_Invoice.xlsx")
        daily_records = parse_excel(wb[SHEET_NAME])
    finally:
        wb.close()

    if not daily_records:
        print("WARNING: No valid data rows found in Excel file")
    else:
        print(f"Parsed {len(daily_records)} daily records from Excel")
        print(f"Date range: {daily_records[0]['date']} to {daily_records[-1]['date']}")

    return daily_records

def format_ist_timestamp():
    """Return the current IST time formatted as 'D Month YYYY, H:MM AM/PM IST'."""
    now_ist = datetime.now(IST)
    hour_12 = now_ist.strftime("%I").lstrip("0") or "12"
    return f"{now_ist.day} {now_ist.strftime('%B %Y')}, {hour_12}:{now_ist.strftime('%M %p')} IST"
