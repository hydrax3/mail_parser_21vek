import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

GOOGLE_SHEET_URL = os.getenv('GOOGLE_SHEET_URL')
CREDENTIALS_FILE = 'credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
TARGET_GID = 148916183

def test_sheet_access():
    print(f"Connecting to {GOOGLE_SHEET_URL}...")
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    
    try:
        sheet = client.open_by_url(GOOGLE_SHEET_URL)
        print(f"Spreadsheet Title: {sheet.title}")
        
        found = False
        print("Available Worksheets:")
        for ws in sheet.worksheets():
            print(f" - '{ws.title}' (GID: {ws.id})")
            if str(ws.id) == str(TARGET_GID):
                found = True
                print(f"   >>> TARGET FOUND: '{ws.title}'")
                
        if found:
            ws = next(w for w in sheet.worksheets() if str(w.id) == str(TARGET_GID))
            print(f"Attempting to write test row to '{ws.title}'...")
            ws.append_row(["TEST_ID", "TEST_SUBJECT", "TEST_SENDER", "TEST_TIME", "DEBUG_WRITE"])
            print("Write successful.")
        else:
            print(f"ERROR: Target GID {TARGET_GID} NOT FOUND.")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    test_sheet_access()
