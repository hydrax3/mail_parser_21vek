import sys
import os
import json

# Ensure we can import parser.py from current directory
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parser import sync_emails

def main():
    try:
        # Run sync and get data
        result = sync_emails()
        
        # Output JSON to stdout
        print("JSON_START")
        print(json.dumps(result, default=str)) # default=str to handle datetime objects if any
        print("JSON_END")
        
    except Exception as e:
        error_res = {"error": str(e)}
        print("JSON_START")
        print(json.dumps(error_res))
        print("JSON_END")

if __name__ == "__main__":
    main()
