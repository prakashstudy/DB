import os
import sys

# Add project path to sys.path
sys.path.append(r"C:\Users\hp\Desktop\cfp-database")

from dotenv import load_dotenv
load_dotenv(dotenv_path=r"C:\Users\hp\Desktop\cfp-database\.env")

from sheets import SheetsDB

script_url = os.getenv("GOOGLE_APPS_SCRIPT_URL", "")
if script_url:
    db = SheetsDB(script_url)
    try:
        cols = db.all_columns()
        print("COLUMNS:")
        for c in cols:
            if any(x in c.lower() for x in ["node", "er", "pr", "her2", "age", "status", "stage", "year", "date", "diagnos"]):
                print("  -", c)
    except Exception as e:
        print("ERROR:", e)
else:
    print("NO SCRIPT URL")
