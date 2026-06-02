"""
Debug script to test Google Apps Script Web App connection (both GET and POST).
Reads .env, sends a test request to the Apps Script URL, and prints the raw response.
"""

import os
from dotenv import load_dotenv
import requests

load_dotenv()

SCRIPT_URL = os.getenv("GOOGLE_APPS_SCRIPT_URL", "")

print("=========================================")
print("  DEBUGGING GOOGLE APPS SCRIPT CONNECTION")
print("=========================================")
print(f"URL in .env: '{SCRIPT_URL}'\n")

if not SCRIPT_URL or "YOUR_DEPLOYED_WEB_APP_URL_HERE" in SCRIPT_URL:
    print("[-] Error: GOOGLE_APPS_SCRIPT_URL is either empty or has the placeholder value in your .env file.")
    print("Please open C:\\Users\\hp\\Desktop\\cfp-database\\.env and put your real Apps Script Web App URL there.")
    exit(1)

try:
    print("[*] Sending test GET request to get headers...")
    resp = requests.get(SCRIPT_URL, params={"action": "getHeaders"}, timeout=15)
    print(f"GET Status Code: {resp.status_code}")
    print(f"GET Content-Type: {resp.headers.get('Content-Type', '')}")
    try:
        print(f"GET Response: {resp.json()}")
    except Exception as e:
        print(f"GET JSON failed: {e}")
        print(resp.text[:400])

    print("\n[*] Sending test POST request (upsert test)...")
    test_payload = {
        "action": "upsert",
        "tabName": "Demographics",
        "headers": ["Patient_ID", "Name", "Age"],
        "rows": [{"Patient_ID": "9999", "Name": "Test Patient", "Age": "45"}]
    }
    
    resp_post = requests.post(SCRIPT_URL, json=test_payload, timeout=20)
    print(f"POST Status Code: {resp_post.status_code}")
    print(f"POST Content-Type: {resp_post.headers.get('Content-Type', '')}")
    
    try:
        data = resp_post.json()
        print("\n[+] POST Success! Deployed Apps Script accepted POST and returned correct JSON:")
        print(data)
    except Exception as je:
        print("\n[-] POST JSON Parsing Failed!")
        print(f"Error details: {je}")
        print("\n=== RAW POST RESPONSE TEXT ===")
        print(resp_post.text[:800])
        print("==============================")

except Exception as e:
    print(f"\n[-] Network Connection Failed: {e}")
