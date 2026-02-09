from imap_tools import MailBox
import os
from dotenv import load_dotenv

load_dotenv()

YANDEX_EMAIL = os.getenv('YANDEX_EMAIL')
YANDEX_PASSWORD = os.getenv('YANDEX_PASSWORD')

try:
    with MailBox('mail.21vek.tech', port=993).login(YANDEX_EMAIL, YANDEX_PASSWORD) as mailbox:
        with open("folders_list.txt", "w", encoding='utf-8') as f:
            for folder in mailbox.folder.list():
                f.write(f"{folder.name} | {folder.flags}\n")
        print("Done.")
except Exception as e:
    print(f"Error: {e}")
