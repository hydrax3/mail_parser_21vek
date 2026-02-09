import os
import datetime
import re
from datetime import timedelta
from dotenv import load_dotenv
from imap_tools import MailBox, AND
import gspread
from google.oauth2.service_account import Credentials
import email.utils

# Load environment variables
load_dotenv()

YANDEX_EMAIL = os.getenv('YANDEX_EMAIL')
YANDEX_PASSWORD = os.getenv('YANDEX_PASSWORD')
GOOGLE_SHEET_URL = os.getenv('GOOGLE_SHEET_URL')
CREDENTIALS_FILE = 'credentials.json'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SENT_FOLDER_NAMES = ['&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-', 'Sent', 'Send', 'Отправленные', 'Sent Items']

# Archive settings
ARCHIVE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1mv3pbNrGThHexCoHFxKYIbOH2U2TuAJoiVHokKqtN4g"
ARCHIVE_GID = 0  # Main archive sheet
STATS_GID = 96142908  # Statistics sheet
INACTIVE_MONTHS = 3  # Months of inactivity before archiving

# Timezone (UTC+3 for Moscow)
MSK_TZ = datetime.timezone(datetime.timedelta(hours=3))

# Operators sheet GID
OPERATORS_GID = 2115150025

# Sync state file
LAST_SYNC_FILE = '.last_sync'

def get_sheet():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f"Credentials file '{CREDENTIALS_FILE}' not found.")
    
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    try:
        sheet = client.open_by_url(GOOGLE_SHEET_URL)
        # Get worksheet by GID=0 (main emails sheet)
        for ws in sheet.worksheets():
            if ws.id == 0:
                return ws
        # Fallback to sheet1 if GID=0 not found
        return sheet.sheet1 
    except Exception as e:
        print(f"Error opening sheet: {e}")
        raise

def get_log_sheet(client, gid_str):
    """Retrieves a specific worksheet by GID."""
    try:
        sheet = client.open_by_url(GOOGLE_SHEET_URL)
        for ws in sheet.worksheets():
            if str(ws.id) == str(gid_str):
                return ws
        
        # If not found, list available GIDs for debugging
        available_gids = [f"{ws.title} (GID: {ws.id})" for ws in sheet.worksheets()]
        raise ValueError(f"Worksheet with GID {gid_str} not found. Available: {', '.join(available_gids)}")
    except Exception as e:
        print(f"Error opening log sheet: {e}")
        raise

def get_operator_emails(client):
    """
    Загружает список email операторов из листа GID=OPERATORS_GID.
    Возвращает set() email-адресов в lowercase.
    При ошибке логирует и возвращает пустой set.
    """
    try:
        sheet = client.open_by_url(GOOGLE_SHEET_URL)
        ws = None
        for worksheet in sheet.worksheets():
            if worksheet.id == OPERATORS_GID:
                ws = worksheet
                break
        
        if not ws:
            print(f"Warning: Operators sheet (GID={OPERATORS_GID}) not found.")
            return set()
        
        # Читаем столбец A
        values = ws.col_values(1)
        # Нормализуем к lowercase, пропускаем пустые
        operators = set()
        for v in values:
            if v and v.strip():
                operators.add(v.strip().lower())
        
        print(f"Loaded {len(operators)} operator emails from GID={OPERATORS_GID}")
        return operators
        
    except Exception as e:
        print(f"Error loading operator emails: {e}")
        return set()

def normalize_date(date_obj):
    # Ensure we use MSK timezone
    if date_obj.tzinfo is None:
        date_obj = date_obj.replace(tzinfo=MSK_TZ)
    return date_obj.astimezone(MSK_TZ).strftime('%Y-%m-%d %H:%M:%S')

def extract_email(sender_str):
    """
    Extracts pure email address from a 'Name <email>' string.
    Returns lowercase email or empty string.
    """
    if not sender_str:
        return ""
    # parseaddr returns (realname, email_address)
    name, addr = email.utils.parseaddr(sender_str)
    return addr.lower()

def clean_subject(subject):
    """Removes Re:, Fwd: prefixes for soft matching."""
    if not subject: return ""
    s = subject.strip()
    while True:
        # Standard Re/Fwd prefix removal
        new_s = re.sub(r'^(Re:|Fwd:|FW:|Отв:)\s*', '', s, flags=re.IGNORECASE).strip()
        
        # Date pattern removal at end of string: DD.MM.YYYY, DD.MM, etc.
        # Matches: " 16.08.2025", " 03-04.09.2025", " 22-23.08"
        # Logic: space + digits/separators at end
        # Regex: \s+\d+[-./]\d+([-./]\d+)?\s*$
        date_pattern = r'\s+\d+[-./]\d+([-./]\d+)?\s*$'
        new_s = re.sub(date_pattern, '', new_s).strip()
        
        if new_s == s: break
        s = new_s
    return s

def get_email_references(msg):
    """Extracts all related IDs (Message-ID, In-Reply-To, References) from a message."""
    refs = set()
    
    # Own ID
    msg_id = msg.headers.get('message-id', [str(msg.uid)])[0].strip('<> ')
    refs.add(msg_id)
    
    # In-Reply-To
    irt = msg.headers.get('in-reply-to', [])
    if isinstance(irt, str): refs.add(irt.strip('<> '))
    elif isinstance(irt, list): refs.update(x.strip('<> ') for x in irt)
    
    # References
    r = msg.headers.get('references', [])
    if isinstance(r, str): refs.add(r.strip('<> '))
    elif isinstance(r, list): refs.update(x.strip('<> ') for x in r)
    
    return refs, msg_id

def sync_emails():
    if not YANDEX_EMAIL or not YANDEX_PASSWORD:
        return {"error": "Yandex credentials missing in .env"}

    print(f"Connecting to IMAP for {YANDEX_EMAIL}...")
    
    try:
        # Get sheet and load operator emails
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        operator_emails = get_operator_emails(client)
        # Add bot email to operators list
        if YANDEX_EMAIL:
            operator_emails.add(YANDEX_EMAIL.lower())
        
        worksheet = get_sheet()
        all_values = worksheet.get_all_values()
        
        # Schema: id, theme_of_mail, sender, time, status_of_reply, type_of_email, last_replyer, last_activity
        header = ['id', 'theme_of_mail', 'sender', 'time', 'status_of_reply', 'type_of_email', 'last_replyer', 'last_activity']
        
        if not all_values:
            worksheet.append_row(header)
            all_values = [header] 
       
        id_map = {}  # message_id -> row_index
        subject_map = {}  # clean_subject -> row_index
        row_data = {}  # row_index -> row values (for comparison)
        
        for i, row in enumerate(all_values):
            if i == 0: continue
            if row:
                row_idx = i + 1
                if len(row) > 0: 
                    id_map[row[0]] = row_idx
                    row_data[row_idx] = row  # Store row data for later comparison
                if len(row) > 1: 
                    subj = clean_subject(row[1])
                    if subj: subject_map[subj] = row_idx


    except Exception as e:
        return {"error": f"Failed to access Google Sheets: {e}"}

    # Updates and New Rows
    updates = []
    new_rows = []
    
    # Incremental sync: read last sync date from file
    last_sync_date = None
    if os.path.exists(LAST_SYNC_FILE):
        try:
            with open(LAST_SYNC_FILE, 'r') as f:
                date_str = f.read().strip()
                last_sync_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
                print(f"Incremental sync from: {last_sync_date}")
        except:
            pass
    
    if not last_sync_date:
        print("Full sync: parsing ALL emails...")
    
    timeline = []
    
    try:
        with MailBox('mail.21vek.tech', port=993).login(YANDEX_EMAIL, YANDEX_PASSWORD) as mailbox:
            # INBOX
            mailbox.folder.set('INBOX')
            print("Scanning INBOX...")
            if last_sync_date:
                for msg in list(mailbox.fetch(AND(date_gte=last_sync_date))):
                    timeline.append({'msg': msg, 'type': 'received'})
            else:
                for msg in list(mailbox.fetch()):
                    timeline.append({'msg': msg, 'type': 'received'})

            # SENT
            sent_folder = None
            for name in SENT_FOLDER_NAMES:
                if mailbox.folder.exists(name):
                    sent_folder = name
                    break
            
            if sent_folder:
                print(f"Scanning {sent_folder}...")
                mailbox.folder.set(sent_folder)
                if last_sync_date:
                    for msg in list(mailbox.fetch(AND(date_gte=last_sync_date))):
                        timeline.append({'msg': msg, 'type': 'sent'})
                else:
                    for msg in list(mailbox.fetch()):
                        timeline.append({'msg': msg, 'type': 'sent'})
            
    except Exception as e:
        return {"error": f"IMAP Error: {e}"}
        
    # Sort by date
    def get_aware_date(x):
        d = x['msg'].date
        if d.tzinfo is None:
            d = d.replace(tzinfo=MSK_TZ)
        return d.astimezone(MSK_TZ)
        
    timeline.sort(key=get_aware_date)
    
    print(f"Processing {len(timeline)} emails from timeline...")
    processed_message_ids = set()

    for item in timeline:
        msg = item['msg']
        email_type = item['type']
        refs, msg_id = get_email_references(msg)
        
        if msg_id in processed_message_ids: continue
        processed_message_ids.add(msg_id)
        
        parent_row_idx = None
        
        # A. Strict ID Check
        for ref in refs:
            if ref in id_map:
                parent_row_idx = id_map[ref]
                break
        
        # B. Fallback Subject Check
        if not parent_row_idx:
            subj = clean_subject(msg.subject)
            if subj and subj in subject_map:
                parent_row_idx = subject_map[subj]
                print(f"Matched by Subject: '{subj}' -> Row {parent_row_idx}")

        if parent_row_idx:
            # UPDATE EXISTING ROW - but only if this is a NEW message in the thread
            existing_row = row_data.get(parent_row_idx, [])
            # Use last_activity (column H, index 7) for comparison, not time (column D)
            existing_last_activity = existing_row[7] if len(existing_row) > 7 else ""
            new_time = normalize_date(msg.date)
            
            # Skip if this is the same message (same timestamp)
            if existing_last_activity == new_time:
                continue
            
            # Check if manually closed - if so, we force update to reopen
            is_closed = str(existing_last_activity).strip().lower() == 'закрыт'
            
            # Only update if the new message is actually newer OR if thread was closed
            if not is_closed and existing_last_activity and existing_last_activity >= new_time:
                continue
                
            print(f"Updating Row {parent_row_idx} with new {email_type} email from {msg.from_}")
            # Determine status based on whether sender is an operator
            sender_email = extract_email(msg.from_).lower()
            if sender_email in operator_emails:
                new_status = 'оператор ответил'
            else:
                new_status = 'ответ не от оператора'
            
            # Note: D (time) is NOT updated - it's the original thread creation time
            updates.append({'range': f'E{parent_row_idx}', 'values': [[new_status]]})
            updates.append({'range': f'G{parent_row_idx}', 'values': [[msg.from_]]})
            # Update last_activity (column H)
            updates.append({'range': f'H{parent_row_idx}', 'values': [[new_time]]})
            
            # Update cached data to prevent duplicate updates in same run
            if parent_row_idx in row_data:
                # Extend row if needed
                while len(row_data[parent_row_idx]) <= 7:
                    row_data[parent_row_idx].append("")
                row_data[parent_row_idx][7] = new_time
        else:
            # NEW ROW
            if msg_id in id_map: continue

            print(f"New Thread: {msg.subject[:30]}")
            status = "ответа нет" if email_type == 'received' else "отправлено"
            
            row = [
                msg_id,
                msg.subject,
                msg.from_,
                normalize_date(msg.date),
                status,
                email_type,
                msg.from_,
                normalize_date(msg.date)  # last_activity
            ]
            new_rows.append(row)
            
            next_row_idx = len(all_values) + len(new_rows)
            id_map[msg_id] = next_row_idx
            subj = clean_subject(msg.subject)
            if subj: subject_map[subj] = next_row_idx


    # Execute Writes
    if new_rows:
        print(f"Adding {len(new_rows)} new threads...")
        worksheet.append_rows(new_rows)
        
    if updates:
        print(f"Updating {len(updates)} cells...")
        worksheet.batch_update(updates)
    


    print("Sync Done. Fetching final data...")
    try:
        # Save current date for next incremental sync
        with open(LAST_SYNC_FILE, 'w') as f:
            f.write(datetime.date.today().strftime('%Y-%m-%d'))
        
        final_data = worksheet.get_all_records()
        return {"status": "success", "data": final_data}
    except Exception as e:
         return {"error": f"Failed to fetch final data: {e}"}

def update_daily_stats(log_rows, stats_sheet_name="OperatorStats"):
    """
    Calculates stats from the log rows and updates the Stats sheet.
    log_rows: list of [id, sender, subject, time, ...] (raw values)
    """
    try:
        # 0. Prep Stats
        stats = {} # { date_str: { operator_email: count } }
        
        for row in log_rows:
            if len(row) < 4: continue
            # row[1] is Sender in our Loop below
            sender = extract_email(row[1]) 
            time_str = row[3] 
            
            if not sender or not time_str: continue
            
            try:
                if ' ' in time_str:
                    dt = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                else:
                    dt = datetime.datetime.strptime(time_str, '%Y-%m-%d')
                date_str = dt.strftime('%Y-%m-%d')
                
                if date_str not in stats: stats[date_str] = {}
                stats[date_str][sender] = stats[date_str].get(sender, 0) + 1
            except:
                continue

        if not stats:
            print("No stats data to update.")
            return

        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(GOOGLE_SHEET_URL)
        
        try:
            stats_ws = spreadsheet.worksheet(stats_sheet_name)
        except:
            print(f"Sheet '{stats_sheet_name}' not found. Creating...")
            stats_ws = spreadsheet.add_worksheet(title=stats_sheet_name, rows=1000, cols=10)
            stats_ws.append_row(["Date", "Operator", "Count"])

        existing_values = stats_ws.get_all_values()
        existing_map = {} # (date, operator) -> row_index
        
        for i, row in enumerate(existing_values):
            if i == 0: continue
            if len(row) >= 2:
                # Key: (Date, Operator)
                existing_map[(row[0], row[1])] = i + 1
        
        updates = []
        new_rows = []
        
        for date_str, ops in stats.items():
            for op, count in ops.items():
                if (date_str, op) in existing_map:
                    row_idx = existing_map[(date_str, op)]
                    current_val = existing_values[row_idx-1][2] if len(existing_values[row_idx-1]) > 2 else "0"
                    if str(current_val) != str(count):
                         updates.append({'range': f'C{row_idx}', 'values': [[count]]})
                else:
                    new_rows.append([date_str, op, count])
        
        if updates:
            print(f"Updating {len(updates)} stats records...")
            stats_ws.batch_update(updates)
        if new_rows:
            print(f"Adding {len(new_rows)} new stats records...")
            stats_ws.append_rows(new_rows)
            
    except Exception as e:
        print(f"Error updating stats: {e}")

def log_operator_activity(log_gid):
    """
    Scans Inbox and Sent for operator emails (from GID 2115150025).
    Logs them to sheet `log_gid`.
    Aggregates stats to 'OperatorStats'.
    """
    if not YANDEX_EMAIL or not YANDEX_PASSWORD:
        return {"error": "Yandex credentials missing"}
    
    print(f"Starting Operator Activity Log (Target GID: {log_gid})...")
    
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # 1. Operators
        operators = get_operator_emails(client)
        if not operators: return {"error": "No operators found"}
        
        # 2. Log Sheet
        ws = get_log_sheet(client, log_gid)
        
        # 3. Read Existing
        all_log_values = ws.get_all_values()
        existing_ids = set()
        if all_log_values:
            for row in all_log_values[1:]:
                if row: existing_ids.add(row[0])
        
        # 4. Scan
        new_rows = []
        date_start = datetime.datetime.now(MSK_TZ) - datetime.timedelta(hours=24)
        
        with MailBox('mail.21vek.tech', port=993).login(YANDEX_EMAIL, YANDEX_PASSWORD) as mailbox:
            folders_to_scan = ['INBOX']
            sent_folder = None
            for name in SENT_FOLDER_NAMES:
                if mailbox.folder.exists(name):
                    sent_folder = name
                    break
            if sent_folder: folders_to_scan.append(sent_folder)
            
            for folder in folders_to_scan:
                print(f"Scanning {folder} from {date_start.date()}...")
                mailbox.folder.set(folder)
                search_date = date_start.date()
                for msg in mailbox.fetch(AND(date_gte=search_date)):
                     if msg.date < date_start.astimezone(msg.date.tzinfo): continue
                     
                     sender = extract_email(msg.from_)
                     if sender not in operators: continue
                     
                     msg_id = msg.headers.get('message-id', [str(msg.uid)])[0].strip('<> ')
                     if msg_id in existing_ids: continue
                     
                     # Add [ID, Sender, Subject, Time]
                     new_rows.append([msg_id, msg.from_, msg.subject, normalize_date(msg.date)])
                     existing_ids.add(msg_id)
        
        if new_rows:
            print(f"Adding {len(new_rows)} new operator emails...")
            ws.append_rows(new_rows)
            # Add to memory for stats
            if all_log_values: all_log_values.extend(new_rows)
            else: all_log_values = [["ID", "Sender", "Subject", "Time"]] + new_rows
        else:
            print("No new operator emails found.")
            
        update_daily_stats(all_log_values[1:] if len(all_log_values)>1 else [], "OperatorStats")
        return {"status": "success", "new_count": len(new_rows)}

    except Exception as e:
        print(f"Error in log_operator_activity: {e}")
        return {"error": str(e)}

def log_overdue_emails(target_gid):
    """
    Logs overdue emails (>3 hours, status 'ответа нет') to the specified GID.
    Ignoring operator filtering (GID 2012399964).
    """
    print(f"Logging overdue emails (>3h) to sheet GID {target_gid}...")
    
    try:
        # Connect
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # 1. Open Target Sheet (Log)
        target_ws = get_log_sheet(client, target_gid)
        
        # 2. Open Main Sheet (Source)
        try:
            source_ws = client.open_by_url(GOOGLE_SHEET_URL).worksheet("Sheet1") 
        except:
             pass
        
        source_sheet = client.open_by_url(GOOGLE_SHEET_URL)
        source_ws = None
        for ws in source_sheet.worksheets():
            if ws.id == 0:
                source_ws = ws
                break
        if not source_ws: source_ws = source_sheet.sheet1
        
        # 3. Read Data
        source_values = source_ws.get_all_values() 
        
        # 4. Read Existing Data from Target (Map ID -> Row Index)
        target_values = target_ws.get_all_values()
        target_map = {} # ID -> Row Index
        if target_values:
            for i, row in enumerate(target_values):
                if i == 0: continue # Skip header
                if row:
                    target_map[row[0]] = i + 1
            print(f"DEBUG: Loaded {len(target_map)} IDs from target sheet.")

        new_rows = []
        updates = []
        now = datetime.datetime.now(MSK_TZ)
        
        # 5. Iterate Source
        for i, row in enumerate(source_values):
            if i == 0: continue 
            if len(row) < 5: continue
            
            msg_id = row[0]
            subject = row[1]
            sender = row[2]
            time_str = row[3]
            status = row[4].strip().lower()
            
            if status != 'ответа нет':
                print(f"DEBUG: Msg {msg_id} skipped. Status: '{status}'")
                continue
                
            try:
                dt = datetime.datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                # Assume stored dates are MSK if no tzinfo, but we compare with aware 'now'
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=MSK_TZ) 
            except ValueError:
                print(f"DEBUG: Msg {msg_id} date parse error: '{time_str}'")
                continue
            
            diff = now - dt
            print(f"DEBUG: Msg {msg_id} age: {diff}")
            
            # Check > 3 Hours
            if diff > datetime.timedelta(hours=3):
                duration_str = str(now - dt).split('.')[0]
                
                # Check Deduplication / Update
                if msg_id in target_map:
                    # UPDATE existing row duration (Col E)
                    row_idx = target_map[msg_id]
                    # print(f"DEBUG: Updating Row {row_idx} for ID {msg_id}")
                    updates.append({
                        'range': f'E{row_idx}',
                        'values': [[duration_str]]
                    })
                else:
                    # INSERT new row
                    new_rows.append([
                        msg_id,
                        subject,
                        sender,
                        time_str,
                        duration_str
                    ])
                    # Add to map to prevent dupes in same run
                    # (Though we don't know the row index yet, but preventing double insert is enough)
                    target_map[msg_id] = -1 
        
        # 6. Execute Writes
        if new_rows:
            print(f"Adding {len(new_rows)} new overdue emails...")
            target_ws.append_rows(new_rows)
            
        if updates:
            print(f"Updating duration for {len(updates)} existing overdue emails...")
            target_ws.batch_update(updates)
            
        if not new_rows and not updates:
            print("No changes for overdue emails.")
            
        return {"status": "success", "count": len(new_rows), "updated": len(updates)}

    except Exception as e:
        print(f"Error in log_overdue_emails: {e}")
        return {"error": str(e)}

def get_archive_sheet(gid):
    """Get worksheet from archive spreadsheet by GID."""
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(ARCHIVE_SHEET_URL)
    
    for ws in sheet.worksheets():
        if ws.id == gid:
            return ws
    
    available = [f"{ws.title} (GID: {ws.id})" for ws in sheet.worksheets()]
    raise ValueError(f"Worksheet GID {gid} not found. Available: {', '.join(available)}")

def aggregate_to_stats(rows, stats_ws):
    """
    Aggregate rows to statistics sheet.
    Format: year_month, count
    Increments existing month counts or adds new rows.
    """
    if not rows:
        return 0
    
    # Group by year-month
    month_counts = {}
    for row in rows:
        # row[3] is time column (YYYY-MM-DD HH:MM:SS)
        time_str = row[3] if len(row) > 3 else ""
        if time_str:
            year_month = time_str[:7]  # "2026-02"
        else:
            year_month = datetime.datetime.now(MSK_TZ).strftime("%Y-%m")
        month_counts[year_month] = month_counts.get(year_month, 0) + 1
    
    # Read existing stats
    existing = stats_ws.get_all_values()
    existing_map = {}  # year_month -> row_index
    for i, r in enumerate(existing):
        if i == 0 or not r:
            continue
        if r[0]:
            existing_map[r[0]] = i + 1
    
    # Update or insert
    updates = []
    new_rows = []
    
    for ym, count in month_counts.items():
        if ym in existing_map:
            row_idx = existing_map[ym]
            # Get current count and add
            current = existing[row_idx - 1][1] if len(existing[row_idx - 1]) > 1 else "0"
            try:
                new_count = int(current) + count
            except:
                new_count = count
            updates.append({'range': f'B{row_idx}', 'values': [[new_count]]})
        else:
            new_rows.append([ym, count])
    
    if updates:
        stats_ws.batch_update(updates)
    if new_rows:
        stats_ws.append_rows(new_rows)
    
    return sum(month_counts.values())

def archive_inactive_threads():
    """
    Archives email threads with no activity for > INACTIVE_MONTHS.
    1. Read main sheet
    2. Find rows where last_activity > 3 months ago
    3. Copy to archive sheet
    4. Aggregate to stats
    5. Delete from main sheet
    """
    print(f">>> Archiving threads inactive for >{INACTIVE_MONTHS} months...")
    
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # Open main sheet
        main_sheet = client.open_by_url(GOOGLE_SHEET_URL)
        main_ws = None
        for ws in main_sheet.worksheets():
            if ws.id == 0:
                main_ws = ws
                break
        if not main_ws:
            main_ws = main_sheet.sheet1
        
        # Open archive and stats sheets
        archive_ws = get_archive_sheet(ARCHIVE_GID)
        stats_ws = get_archive_sheet(STATS_GID)
        
        # Read main data
        all_values = main_ws.get_all_values()
        if len(all_values) <= 1:
            return {"status": "success", "archived": 0, "message": "No data to archive"}
        
        header = all_values[0]
        
        # Find last_activity column (H = index 7)
        # Schema: id, theme, sender, time, status, type, last_replyer, last_activity
        LAST_ACTIVITY_COL = 7
        
        now = datetime.datetime.now(MSK_TZ)
        cutoff = now - datetime.timedelta(days=INACTIVE_MONTHS * 30)
        
        rows_to_archive = []
        rows_to_delete = []  # 1-indexed row numbers
        
        for i, row in enumerate(all_values):
            if i == 0:
                continue  # Skip header
            
            # Get last_activity or fallback to time (col D)
            last_activity_str = row[LAST_ACTIVITY_COL] if len(row) > LAST_ACTIVITY_COL else ""
            if not last_activity_str:
                last_activity_str = row[3] if len(row) > 3 else ""
            
            if not last_activity_str:
                continue
            
            try:
                last_dt = datetime.datetime.strptime(last_activity_str, '%Y-%m-%d %H:%M:%S')
            except ValueError:
                continue
            
            if last_dt < cutoff:
                # Add archived_at timestamp
                archived_row = list(row) + [now.strftime('%Y-%m-%d %H:%M:%S')]
                rows_to_archive.append(archived_row)
                rows_to_delete.append(i + 1)  # 1-indexed
        
        if not rows_to_archive:
            print("No inactive threads found.")
            return {"status": "success", "archived": 0}
        
        print(f"Found {len(rows_to_archive)} inactive threads to archive...")
        
        # 1. Copy to archive
        archive_ws.append_rows(rows_to_archive)
        print(f"Copied {len(rows_to_archive)} rows to archive.")
        
        # 2. Aggregate to stats
        aggregated = aggregate_to_stats(rows_to_archive, stats_ws)
        print(f"Aggregated {aggregated} emails to stats.")
        
        # 3. Delete from main (from bottom to top to preserve indices)
        rows_to_delete.sort(reverse=True)
        for row_idx in rows_to_delete:
            main_ws.delete_rows(row_idx)
        print(f"Deleted {len(rows_to_delete)} rows from main sheet.")
        
        return {
            "status": "success",
            "archived": len(rows_to_archive),
            "aggregated": aggregated,
            "deleted": len(rows_to_delete)
        }
        
    except Exception as e:
        print(f"Error in archive_inactive_threads: {e}")
        return {"error": str(e)}

if __name__ == "__main__":
    print(">>> Running full sync (Inbox + Sent Log)...")
    
    # 1. Sync Inbox
    result = sync_emails()
    if "error" in result:
        print(f"Inbox Sync Error: {result['error']}")
    else:
        print("Inbox Sync Success.")

    # 2. Log Operator Activity (Sent + Inbox)
    # Log GID: 1286665239
    log_result = log_operator_activity(1286665239)
    if "error" in log_result:
        print(f"Operator Log Error: {log_result['error']}")
    else:
        print(f"Operator Log Success. New: {log_result.get('new_count', 0)}")
        
    # 3. Log Overdue Emails (>3h)
    # GID: 148916183
    overdue_result = log_overdue_emails(148916183)
    if "error" in overdue_result:
        print(f"Overdue Log Error: {overdue_result['error']}")
    else:
        print(f"Overdue Log Success. Count: {overdue_result.get('count', 0)}")

    # 4. Archive Inactive Threads (>3 months)
    # Run with: python parser.py --archive
    import sys
    if "--archive" in sys.argv:
        archive_result = archive_inactive_threads()
        if "error" in archive_result:
            print(f"Archive Error: {archive_result['error']}")
        else:
            print(f"Archive Success. Archived: {archive_result.get('archived', 0)}, Deleted: {archive_result.get('deleted', 0)}")
