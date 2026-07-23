import os
import time
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import traceback

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- since
# this runs unattended from the Startup folder at every login).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ==========================================
# SYSTEM CONFIGURATION
# ==========================================
PENDING_SYNC_DIR = "PENDING_CLOUD_SYNC"
QUARANTINE_DIR = os.path.join(PENDING_SYNC_DIR, "_QUARANTINE_MALFORMED")
REPORT_DIR = "JDan_Executive_Reports"
JSON_KEY_FILE = "credentials.json"
GOOGLE_SHEET_NAME = "JDAN_LIVE_OPERATIONS_REPORTS"
MAX_CONNECT_RETRIES = 3
RETRY_DELAY_SECONDS = 10

def log_error(error_msg, exc_info=None):
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with open("JDan_System_Logs.txt", "a") as f:
            f.write(f"[{current_time}] AUTO-SYNC ERROR: {error_msg}\n")
            if exc_info:
                f.write(f"{exc_info}\n")
    except:
        pass
    print(f"\n[ERROR DETECTED] => {error_msg}")

# ==========================================
# GOOGLE APIS AUTHENTICATION
# ==========================================
def setup_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    if not os.path.exists(JSON_KEY_FILE):
        log_error(f"CRITICAL: '{JSON_KEY_FILE}' is missing. Cannot connect to Google.")
        return None

    # Retry a transient failure (e.g. DNS/network not ready yet right after a
    # reboot/login) a few times before giving up -- a single failed attempt used
    # to kill the process permanently with no way back until manually relaunched
    # (root cause of the sync going dark after the 2026-07-04 07:04 AM boot-time
    # DNS blip). The Scheduled Task's own restart-on-failure setting is the
    # outer safety net if all retries here are exhausted.
    last_error = None
    for attempt in range(1, MAX_CONNECT_RETRIES + 1):
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
            client = gspread.authorize(creds)

            # 🔥 FIX: Ibabato natin ang buong Workbook, hindi lang isang specific na tab
            workbook = client.open(GOOGLE_SHEET_NAME)
            return workbook
        except Exception as e:
            last_error = e
            log_error(f"Connection attempt {attempt}/{MAX_CONNECT_RETRIES} failed: {e}", traceback.format_exc())
            if attempt < MAX_CONNECT_RETRIES:
                time.sleep(RETRY_DELAY_SECONDS)

    log_error(f"CRITICAL: Could not connect to Google after {MAX_CONNECT_RETRIES} attempts. Last error: {last_error}")
    return None

# ==========================================
# BACKGROUND COURIER PROCESS (DYNAMIC ROUTING)
# ==========================================
def run_sync(workbook):
    if not os.path.exists(PENDING_SYNC_DIR): 
        return

    pending_files = [f for f in os.listdir(PENDING_SYNC_DIR) if f.endswith('.json')]
    if not pending_files: 
        return

    print(f"\n[{time.strftime('%H:%M:%S')}] Found {len(pending_files)} pending file(s). Syncing to the cloud...")

    for filename in pending_files:
        file_path = os.path.join(PENDING_SYNC_DIR, filename)
        time.sleep(1) 
        
        try:
            with open(file_path, "r") as f:
                data = json.load(f)

            if not isinstance(data, dict):
                raise TypeError(f"Expected a JSON object (dict) but got {type(data).__name__}")

            # 🔥 NEW: Titingnan ng system kung EOD o SOD ang report
            target_tab = data.get('target_tab', 'EOD_Settlements') # Default sa EOD kung lumang file
            row_data = data.get('row_data', [])
            txt_report_name = data.get('report_filename')
            receipt_text = "No receipt text found."
            
            if txt_report_name:
                txt_path = os.path.join(REPORT_DIR, txt_report_name)
                if os.path.exists(txt_path):
                    with open(txt_path, "r", encoding="utf-8") as txt_file:
                        receipt_text = txt_file.read()
            
            # Pigilan ang Formula Parse Error (#ERROR!)
            receipt_text = "'" + receipt_text
            row_data.append(receipt_text)
            
            # Subukang ipasok sa tamang Tab
            try:
                sheet = workbook.worksheet(target_tab)
                sheet.append_row(row_data, value_input_option='USER_ENTERED')
                print(f" -> [SUCCESS] Synced to {target_tab}: {filename}")
                
                # Delete JSON after successful push
                try:
                    os.remove(file_path)
                except PermissionError:
                    print(f" -> [WARNING] Could not delete {filename} right now.")
            except gspread.exceptions.WorksheetNotFound:
                print(f" -> [FAILED] Could not find the tab '{target_tab}' in Google Sheets. Please create the tab first.")

        except json.JSONDecodeError:
            print(f" -> [WAITING] {filename} is still being written by the system.")
        except (TypeError, AttributeError, KeyError) as e:
            # Payload is structurally malformed (not a network/API hiccup) --
            # retrying forever would just spam the log every 30s without ever
            # succeeding (this is exactly what happened with
            # sync_20260621_094515.json on 2026-06-25: 'list' object has no
            # attribute 'get', retried 9 times before someone intervened).
            # Quarantine it for manual review instead of deleting it.
            os.makedirs(QUARANTINE_DIR, exist_ok=True)
            quarantine_path = os.path.join(QUARANTINE_DIR, filename)
            try:
                os.replace(file_path, quarantine_path)
                log_error(f"MALFORMED payload {filename} moved to quarantine ({QUARANTINE_DIR}). Error: {e}")
                print(f" -> [QUARANTINED] {filename}: {e}")
            except OSError as move_err:
                log_error(f"Could not quarantine malformed {filename}: {move_err}")
        except Exception as e:
            error_details = f"Could not sync {filename}. Error: {e}"
            print(f" -> [FAILED] {error_details}")
            log_error(error_details, traceback.format_exc())

if __name__ == "__main__":
    print("=======================================")
    print(" JDAN Auto-Sync Courier Started ")
    print(" (With Dynamic Tab Routing) ")
    print("=======================================")
    
    print("Connecting to Google Cloud...")
    active_workbook = setup_google_sheets()
    
    if active_workbook is None:
        print("\nFATAL ERROR: Could not connect.")
    else:
        print("Connected successfully! Watching in the background...")
        while True:
            run_sync(active_workbook)
            time.sleep(30)