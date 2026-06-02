import os
from dotenv import load_dotenv

def verify():
    load_dotenv()
    
    print("--- Security Verification ---")
    
    # Check .env loads
    creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE")
    admin_user = os.getenv("ADMIN_USERNAME")
    
    if creds_file == "capable-mind-416116-068348e0fcd5.json":
        print("[OK] GOOGLE_CREDENTIALS_FILE correctly set in .env")
    else:
        print(f"[FAIL] GOOGLE_CREDENTIALS_FILE missing or incorrect: {creds_file}")
        
    if admin_user == "admin":
        print("[OK] ADMIN_USERNAME correctly set in .env")
    else:
        print(f"[FAIL] ADMIN_USERNAME missing or incorrect: {admin_user}")

    # Check if .gitignore exists and contains the wildcard
    gitignore_path = ".gitignore"
    if os.path.exists(gitignore_path):
        with open(gitignore_path, "r") as f:
            content = f.read()
            if "*.json" in content:
                print("[OK] .gitignore contains *.json wildcard")
            else:
                print("[FAIL] .gitignore missing *.json wildcard")
    else:
        print("[FAIL] .gitignore file not found")

if __name__ == "__main__":
    verify()
