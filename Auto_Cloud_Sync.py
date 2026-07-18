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
REPORT_DIR = "JDan_Executive_Reports"
JSON_KEY_FILE = "credentials.json"
GOOGLE_SHEET_NAME = "JDAN_LIVE_OPERATIONS_REPORTS"

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
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        if not os.path.exists(JSON_KEY_FILE):
            log_error(f"CRITICAL: '{JSON_KEY_FILE}' is missing. Cannot connect to Google.")
            return None
            
        creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
        client = gspread.authorize(creds)
        
        # 🔥 FIX: Ibabato natin ang buong Workbook, hindi lang isang specific na tab
        workbook = client.open(GOOGLE_SHEET_NAME)
        return workbook
    except Exception as e:
        log_error(f"Connection Error: {e}", traceback.format_exc())
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