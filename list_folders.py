from imap_tools import MailBox
import os
from dotenv import load_dotenv

load_dotenv()

YANDEX_EMAIL = os.getenv('YANDEX_EMAIL')
YANDEX_PASSWORD = os.getenv('YANDEX_PASSWORD')

try:
    with MailBox('mail.21vek.tech', port=993).login(YANDEX_EMAIL, YANDEX_PASSWORD) as mailbox:
        print("Folders:")
        for f in mailbox.folder.list():
            print(f"- {f.name} (flags: {f.flags})")
except Exception as e:
    print(f"Error: {e}")
