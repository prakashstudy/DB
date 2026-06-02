import os
import sys

sys.path.append(r"C:\Users\hp\Desktop\cfp-database")
from dotenv import load_dotenv
load_dotenv(dotenv_path=r"C:\Users\hp\Desktop\cfp-database\.env")
from sheets import SheetsDB

script_url = os.getenv("GOOGLE_APPS_SCRIPT_URL", "")
if script_url:
    db = SheetsDB(script_url)
    try:
        data = db.joined_data()
        print("SAMPLE ROWS FOR SURVIVAL:")
        for i, row in enumerate(data[:10]):
            print(f"Row {i+1}:")
            for k, v in row.items():
                if any(x in k.lower() for x in ["survival", "dfs", "event", "death", "status"]):
                    print(f"  - {k}: {v}")
    except Exception as e:
        print("ERROR:", e)
else:
    print("NO SCRIPT URL")
