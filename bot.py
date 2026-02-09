import time
import datetime
import traceback
# 'sync_emails' seems to be the main function in parser.py based on file content
from parser import sync_emails, log_sent_emails

# Loop interval in seconds (30 minutes)
INTERVAL = 30 * 60

def run_periodically():
    print(f"Bot started. Running every {INTERVAL} seconds.")
    while True:
        try:
            print(f"\n--- Starting run at {datetime.datetime.now()} ---")
            
            # 1. Parse Inbox
            print(">>> Syncing Inbox Threads...")
            sync_emails()
            
            # 2. Log Sent Emails
            print(">>> Logging Sent Emails...")
            log_sent_emails(1286665239)
            
            print("--- Run finished ---")
        except Exception as e:
            print(f"An error occurred: {e}")
            traceback.print_exc()
        
        print(f"Sleeping for {INTERVAL} seconds...")
        time.sleep(INTERVAL)

if __name__ == "__main__":
    run_periodically()
