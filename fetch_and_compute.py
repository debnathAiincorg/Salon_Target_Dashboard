import os
import json
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl

# Constants
WEEKLY_TARGET = 15000
OUTPUT_FILE = "dashboard-data.json"
EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")

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

        # Placeholder: parsing will be added in Task 2
        print("Excel file loaded successfully")

        wb.close()

    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        exit(1)

if __name__ == "__main__":
    main()
