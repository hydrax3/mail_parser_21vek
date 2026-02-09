from parser import get_log_sheet, SCOPES, CREDENTIALS_FILE, GOOGLE_SHEET_URL
import gspread
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv

load_dotenv()

def verify_gid_access():
    print("Testing GID access...")
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        target_gid = "1286665239"
        ws = get_log_sheet(client, target_gid)
        print(f"SUCCESS: Found worksheet '{ws.title}' with GID {ws.id}")
        
    except Exception as e:
        print(f"FAILURE: {e}")

if __name__ == "__main__":
    verify_gid_access()
