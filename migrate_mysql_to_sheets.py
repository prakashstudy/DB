"""
MySQL -> Google Sheets Migration Script.
Reads data from local PHP MySQL database and writes to corresponding Google Sheets tabs.
"""

import os
import sys
import time
import pymysql
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# Config
CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")

# MySQL Settings
MYSQL_HOST = os.getenv("MYSQL_HOST", "localhost")
MYSQL_USER = os.getenv("MYSQL_USER", "root")
MYSQL_PASS = os.getenv("MYSQL_PASS", "")
MYSQL_DB = os.getenv("MYSQL_DB", "")

TABLES = [
    "demographics",
    "dmp",
    "rmd",
    "familypersonalhistory",
    "Pathology",
    "Followup",
    "Samples",
    "Analysis"
]

TABLE_TO_TAB = {
    "demographics": "Demographics",
    "dmp": "DMP",
    "rmd": "RMD",
    "familypersonalhistory": "FamilyPersonalHistory",
    "pathology": "Pathology",
    "followup": "Followup",
    "samples": "Samples",
    "analysis": "Analysis"
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]

def main():
    if not SPREADSHEET_ID:
        print("❌ Error: GOOGLE_SHEET_ID is not configured in your .env file.")
        sys.exit(1)
        
    if not MYSQL_DB:
        print("❌ Error: MYSQL_DB database name is not configured in your .env file.")
        sys.exit(1)

    print("\n===========================================")
    print("🚀 Starting MySQL -> Google Sheets Migration")
    print("===========================================\n")

    # Connect MySQL
    try:
        conn = pymysql.connect(
            host=MYSQL_HOST,
            user=MYSQL_USER,
            password=MYSQL_PASS,
            database=MYSQL_DB,
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor
        )
        print("✅ Connected to MySQL successfully.")
    except Exception as e:
        print(f"❌ Failed to connect to MySQL: {e}")
        sys.exit(1)

    # Connect Google Sheets
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        print("✅ Connected to Google Sheets successfully.")
    except Exception as e:
        print(f"❌ Failed to connect to Google Sheets API: {e}")
        conn.close()
        sys.exit(1)

    # Migrate each table
    for table in TABLES:
        tab_name = TABLE_TO_TAB.get(table.lower(), table)
        print(f"\n📋 Process tab: '{tab_name}' from MySQL table '{table}'...")

        try:
            with conn.cursor() as cursor:
                cursor.execute(f"SELECT * FROM `{table}`")
                rows = cursor.fetchall()
            
            if not rows:
                print(f"  ⚠️ No rows found in MySQL table '{table}'. Skipping.")
                continue

            columns = list(rows[0].keys())
            
            # Setup tab
            try:
                ws = spreadsheet.worksheet(tab_name)
                ws.clear()
            except gspread.exceptions.WorksheetNotFound:
                ws = spreadsheet.add_worksheet(title=tab_name, rows=1000, cols=len(columns))

            # Headers
            ws.update("A1", [columns])
            # Bold & styled header
            ws.format("A1:ZZ1", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": {"red": 0.0, "green": 0.09, "blue": 0.38},
                "horizontalAlignment": "CENTER"
            })

            # Rows
            all_vals = []
            for row in rows:
                val_row = []
                for col in columns:
                    val = row.get(col, "")
                    val_row.append(str(val) if val is not None else "")
                all_vals.append(val_row)

            # Batch write rows
            batch_size = 300
            for i in range(0, len(all_vals), batch_size):
                batch = all_vals[i:i+batch_size]
                start_row = i + 2
                end_col = gspread.utils.rowcol_to_a1(1, len(columns))
                end_col_letter = "".join(c for c in end_col if c.isalpha())
                range_str = f"A{start_row}:{end_col_letter}{start_row + len(batch) - 1}"
                ws.update(range_str, batch, value_input_option="USER_ENTERED")
                print(f"  👉 Wrote rows {i+1} to {i+len(batch)} to sheet.")
                time.sleep(1) # API Rate limit protection

            print(f"  🎉 Tab '{tab_name}' successfully migrated! Total rows: {len(all_vals)}")
            time.sleep(2)

        except Exception as e:
            print(f"  ❌ Error migrating '{table}': {e}")

    conn.close()
    print("\n🏁 Migration task complete!\n")

if __name__ == "__main__":
    main()
