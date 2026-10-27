"""
J-DAN PHARMACY - INVESTIGATION APP (Executive-Only)

One-stop item investigation tool: full movement history, sales/stock-in/SOD-audit
trace, auto-reconciliation, and (after investigating) a guarded manual stock
adjustment with full audit trail. Builds on the reconciliation logic proven in
Trace_Item_Movement.py, generalized to accept any item instead of a hardcoded
list.

Executive-only -- gated behind a PIN screen. Not for staff use.
"""
import os
import re
import csv
import sys
import glob
import json
import queue
import shutil
import warnings
import datetime
import traceback
import threading

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from werkzeug.security import generate_password_hash, check_password_hash

from name_cleaning import clean_text
from item_code_resolver import get_resolver, resolve_series, announce, sheet_text_code

warnings.filterwarnings('ignore', category=UserWarning)

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ==========================================
#    CRASH DETECTOR (Anti-Silent Crash)
# ==========================================
def handle_exception(exc_type, exc_value, exc_traceback):
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    try:
        messagebox.showerror("System Critical Error", f"The system encountered an error. Please screenshot this and send it to the Admin:\n\n{err_msg}")
    except Exception:
        pass

sys.excepthook = handle_exception

# ==========================================
#    SYSTEM CONFIGURATION
# ==========================================
MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
STOCKIN_DIR = 'JDan_StockIn_Reports'
EOD_DIR = os.path.join('JDan_Sales_System', 'EOD_Deduction_Reports')
ARCHIVE_DIR = os.path.join('JDan_Sales_System', 'Archived_Sales_Reports')
INBOX_DIR = 'SALES_REPORT_FROM_POS'
SOD_DIR = 'JDan_Executive_Reports'
REPORT_DIR = 'Investigation_Reports'
PENDING_SYNC_DIR = 'PENDING_CLOUD_SYNC'
ADJUSTMENT_LOG = 'Stock_Adjustments_Log.csv'
SYSTEM_LOG_FILE = 'JDan_System_Logs.txt'
ADJUSTMENT_LOG_HEADER = ['Timestamp', 'Operation_Type', 'Item_Code', 'Item_Name', 'Staff',
                         'Old_Value', 'New_Value', 'Difference', 'Reason', 'Reference', 'Source']

# This app is a fast executive override, deliberately separate from the
# staff-facing StockIn_App.py / Expiry_Scanner.py workflows -- every entry
# this app writes is tagged with this exact source string, in both the CSV
# and the cloud payload, so it's never confused with a normal staff action.
SOURCE_TAG = "Executive Override - Investigation_App"
OPERATION_TYPES = ['STOCK_IN', 'PULL_OUT', 'ADJUSTMENT']
PULL_OUT_REASONS = ['Expired', 'Damaged', 'Returned to Supplier']

ENV_FILE = '.env'
LOCKOUT_FILE = '.investigation_lockout.json'
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_COOLDOWN_MINUTES = 15

STAFF_LIST = ['Dan', 'Angelique', 'Katlyn', 'Kat']

# Color palette (matches the rest of the Pharmacy_Automation suite)
COLOR_HEADER = "#154360"
COLOR_HEADER_SUB = "#a9cce3"
COLOR_BG = "#f4f4f4"
COLOR_GREEN = "#28a745"
COLOR_RED = "#dc3545"
COLOR_ORANGE = "#f0ad4e"
COLOR_GRAY = "#6c757d"
COLOR_TEAL = "#17a2b8"
COLOR_TEXT_MUTED = "#566573"


# ==========================================
#    .env HELPERS (no external dependency --
#    plain KEY=VALUE lines, never the raw PIN,
#    only its hash)
# ==========================================
def read_env():
    values = {}
    if os.path.exists(ENV_FILE):
        with open(ENV_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, val = line.partition('=')
                values[key.strip()] = val.strip()
    return values


def write_env(values):
    with open(ENV_FILE, 'w', encoding='utf-8') as f:
        f.write("# Investigation_App secrets -- never commit this file.\n")
        for key, val in values.items():
            f.write(f"{key}={val}\n")


def get_pin_hash():
    return read_env().get('INVESTIGATION_PIN_HASH')


def set_pin_hash(pin_hash):
    values = read_env()
    values['INVESTIGATION_PIN_HASH'] = pin_hash
    write_env(values)


# ==========================================
#    LOCKOUT STATE (persisted across restarts,
#    so closing/reopening the app can't be used
#    to dodge the cooldown)
# ==========================================
def read_lockout_state():
    if not os.path.exists(LOCKOUT_FILE):
        return {'failed_attempts': 0, 'locked_until': None}
    try:
        with open(LOCKOUT_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {'failed_attempts': 0, 'locked_until': None}


def write_lockout_state(state):
    try:
        with open(LOCKOUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(state, f)
    except OSError:
        pass


def is_locked_out():
    """Returns (locked: bool, seconds_remaining: int)."""
    state = read_lockout_state()
    locked_until = state.get('locked_until')
    if not locked_until:
        return False, 0
    until_dt = datetime.datetime.fromisoformat(locked_until)
    remaining = (until_dt - datetime.datetime.now()).total_seconds()
    if remaining <= 0:
        # Cooldown expired -- clear it.
        write_lockout_state({'failed_attempts': 0, 'locked_until': None})
        return False, 0
    return True, int(remaining)


def register_failed_attempt():
    """Returns (locked_now: bool, attempts_used: int, attempts_left: int)."""
    state = read_lockout_state()
    attempts = state.get('failed_attempts', 0) + 1
    if attempts >= MAX_FAILED_ATTEMPTS:
        locked_until = (datetime.datetime.now() + datetime.timedelta(minutes=LOCKOUT_COOLDOWN_MINUTES)).isoformat()
        write_lockout_state({'failed_attempts': attempts, 'locked_until': locked_until})
        return True, attempts, 0
    write_lockout_state({'failed_attempts': attempts, 'locked_until': None})
    return False, attempts, MAX_FAILED_ATTEMPTS - attempts


def register_success():
    write_lockout_state({'failed_attempts': 0, 'locked_until': None})


def verify_pin(entered_pin):
    stored_hash = get_pin_hash()
    if not stored_hash:
        return False
    return check_password_hash(stored_hash, entered_pin)


# ==========================================
#    LOGIN / PIN SETUP UI
# ==========================================
class LoginWindow:
    """Handles both first-run PIN setup and normal PIN entry. Calls
    on_success(staff_context) once unlocked."""

    def __init__(self, on_success):
        self.on_success = on_success
        self.root = tk.Tk()
        self.root.title("J-DAN PHARMACY - INVESTIGATION APP")
        self.root.geometry("440x480")
        self.root.configure(bg=COLOR_BG)
        self.root.resizable(False, False)

        self.first_run = get_pin_hash() is None

        header = tk.Frame(self.root, bg=COLOR_HEADER, pady=18)
        header.pack(fill="x")
        tk.Label(header, text="🔒 EXECUTIVE ACCESS", fg="white", bg=COLOR_HEADER,
                 font=("Arial", 15, "bold")).pack()
        tk.Label(header, text="Investigation App -- Dan/Danny only", fg=COLOR_HEADER_SUB,
                 bg=COLOR_HEADER, font=("Arial", 9)).pack()

        self.body = tk.Frame(self.root, bg=COLOR_BG)
        self.body.pack(expand=True, fill="both", padx=30, pady=20)

        if self.first_run:
            self._build_setup_ui()
        else:
            self._build_login_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- first-run setup ----------
    def _build_setup_ui(self):
        tk.Label(self.body, text="First-time setup: create your PIN.", bg=COLOR_BG,
                 font=("Arial", 10, "bold"), fg=COLOR_HEADER, wraplength=340, justify="left").pack(pady=(0, 15))
        tk.Label(self.body, text="At least 4 characters. This PIN is stored as a secure hash --\nnever in plain text, never in the source code.",
                 bg=COLOR_BG, font=("Arial", 8), fg=COLOR_TEXT_MUTED, justify="left").pack(pady=(0, 15))

        tk.Label(self.body, text="New PIN:", bg=COLOR_BG, font=("Arial", 10)).pack(anchor="w")
        self.pin_entry = tk.Entry(self.body, show="*", font=("Arial", 14), justify="center")
        self.pin_entry.pack(fill="x", pady=(2, 12))

        tk.Label(self.body, text="Confirm PIN:", bg=COLOR_BG, font=("Arial", 10)).pack(anchor="w")
        self.confirm_entry = tk.Entry(self.body, show="*", font=("Arial", 14), justify="center")
        self.confirm_entry.pack(fill="x", pady=(2, 12))

        self.status_label = tk.Label(self.body, text="", bg=COLOR_BG, font=("Arial", 9, "bold"), fg=COLOR_RED)
        self.status_label.pack(pady=(0, 10))

        tk.Button(self.body, text="CREATE PIN & CONTINUE", font=("Arial", 11, "bold"),
                  bg=COLOR_GREEN, fg="white", command=self._submit_setup).pack(fill="x", ipady=6)
        self.pin_entry.focus_set()
        self.confirm_entry.bind('<Return>', lambda e: self._submit_setup())

    def _submit_setup(self):
        pin = self.pin_entry.get().strip()
        confirm = self.confirm_entry.get().strip()
        if len(pin) < 4:
            self.status_label.config(text="PIN must be at least 4 characters.")
            return
        if pin != confirm:
            self.status_label.config(text="PINs do not match. Try again.")
            self.confirm_entry.delete(0, tk.END)
            return
        set_pin_hash(generate_password_hash(pin))
        register_success()
        self._unlock()

    # ---------- normal login ----------
    def _build_login_ui(self):
        locked, remaining = is_locked_out()

        tk.Label(self.body, text="Enter PIN:", bg=COLOR_BG, font=("Arial", 11)).pack(anchor="w", pady=(10, 2))
        self.pin_entry = tk.Entry(self.body, show="*", font=("Arial", 16), justify="center")
        self.pin_entry.pack(fill="x", pady=(2, 15), ipady=4)
        self.pin_entry.bind('<Return>', lambda e: self._submit_login())

        self.status_label = tk.Label(self.body, text="", bg=COLOR_BG, font=("Arial", 9, "bold"), fg=COLOR_RED, wraplength=340, justify="left")
        self.status_label.pack(pady=(0, 10))

        self.unlock_btn = tk.Button(self.body, text="🔓 UNLOCK", font=("Arial", 12, "bold"),
                                     bg=COLOR_TEAL, fg="white", command=self._submit_login)
        self.unlock_btn.pack(fill="x", ipady=8)

        if locked:
            self._apply_lockout_ui(remaining)
        else:
            self.pin_entry.focus_set()

    def _apply_lockout_ui(self, remaining_seconds):
        self.pin_entry.config(state='disabled')
        self.unlock_btn.config(state='disabled')
        self._tick_cooldown(remaining_seconds)

    def _tick_cooldown(self, remaining_seconds):
        if remaining_seconds <= 0:
            self.status_label.config(fg=COLOR_GREEN, text="Cooldown finished -- you may try again.")
            self.pin_entry.config(state='normal')
            self.unlock_btn.config(state='normal')
            self.pin_entry.focus_set()
            return
        mins, secs = divmod(remaining_seconds, 60)
        self.status_label.config(
            fg=COLOR_RED,
            text=f"Too many failed attempts. Locked for {mins}m {secs:02d}s."
        )
        self.root.after(1000, lambda: self._tick_cooldown(remaining_seconds - 1))

    def _submit_login(self):
        locked, remaining = is_locked_out()
        if locked:
            self._apply_lockout_ui(remaining)
            return

        pin = self.pin_entry.get().strip()
        if verify_pin(pin):
            register_success()
            self._unlock()
            return

        locked_now, attempts_used, attempts_left = register_failed_attempt()
        self.pin_entry.delete(0, tk.END)
        if locked_now:
            self.status_label.config(text="Incorrect PIN. Maximum attempts reached.")
            self._apply_lockout_ui(LOCKOUT_COOLDOWN_MINUTES * 60)
        else:
            self.status_label.config(text=f"Incorrect PIN. {attempts_left} attempt(s) remaining before lockout.")

    def _unlock(self):
        self.root.destroy()
        self.on_success()

    def _on_close(self):
        self.root.destroy()
        sys.exit()

    def run(self):
        self.root.mainloop()



# ==========================================================================
#    PART 1 -- INVESTIGATION LOGIC (read-only)
#
#    Generalizes Trace_Item_Movement.py's proven reconciliation approach
#    (pre-correction snapshot vs. physical count, Item_Code-only matching,
#    name_cleaning for cross-checking) to accept any Item_Code instead of a
#    hardcoded list, and searches the FULL available history by default
#    instead of a manually-chosen lookback window.
# ==========================================================================

def norm(s):
    return clean_text(str(s)).strip().upper()


def parse_date_cell(value):
    try:
        return pd.to_datetime(value).date()
    except (ValueError, TypeError):
        return None


_master_cache = {'df': None, 'mtime': None}


def load_master(force=False):
    """Cached by Master's own mtime -- any external change (EOD run, another
    app's save) automatically invalidates the cache on the next call."""
    mtime = os.path.getmtime(MASTER_FILE) if os.path.exists(MASTER_FILE) else None
    if not force and _master_cache['df'] is not None and _master_cache['mtime'] == mtime:
        return _master_cache['df']
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str)
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    df['Product_Name'] = df['Product_Name'].apply(clean_text)
    _master_cache['df'] = df
    _master_cache['mtime'] = mtime
    return df


def get_item_directory():
    """One row per Item_Code for search/autocomplete: (Item_Code, canonical
    Product_Name, live Current_Stock). Mirrors Item_Settings_App's approach
    of picking the most-frequent spelling per code as the representative name."""
    df = load_master()
    name_freq = df.groupby(['Item_Code', 'Product_Name']).size().reset_index(name='freq')
    name_freq = name_freq.sort_values(['Item_Code', 'freq'], ascending=[True, False])
    representative = name_freq.drop_duplicates('Item_Code', keep='first')[['Item_Code', 'Product_Name']]
    stock = df.groupby('Item_Code')['Current_Stock'].sum().reset_index()
    merged = representative.merge(stock, on='Item_Code', how='left')
    return merged.sort_values('Product_Name').to_dict('records')


def resolve_item_code(query):
    """Resolves free-text (Item_Code or product name, exact or partial) to a
    single Item_Code. Returns (item_code_or_None, list_of_candidate_dicts).
    A single exact/unique match resolves directly; anything else returns
    candidates for the caller to show as suggestions."""
    query = str(query).strip()
    if not query:
        return None, []
    directory = get_item_directory()

    for row in directory:
        if str(row['Item_Code']) == query:
            return row['Item_Code'], [row]

    q_upper = query.upper()
    exact_name = [r for r in directory if str(r['Product_Name']).upper() == q_upper]
    if len(exact_name) == 1:
        return exact_name[0]['Item_Code'], exact_name

    partial = [r for r in directory if q_upper in str(r['Product_Name']).upper() or q_upper in str(r['Item_Code'])]
    if len(partial) == 1:
        return partial[0]['Item_Code'], partial
    return None, partial[:20]


# ---------- session-level file cache ----------
# Investigating a second/third item in the same session would otherwise
# re-read the same ~40 backups + ~35 sales exports + ~30 stock-in batches +
# ~35 EOD reports + ~27 SOD reports from disk every single time (the first
# investigation took 20-30s doing exactly that). Caching each source file's
# *fully parsed, unfiltered* contents by path+mtime makes every investigation
# after the first one in a session only pay for in-memory filtering.
_file_cache = {}


def _cached_read(path, parser):
    mtime = os.path.getmtime(path)
    key = (path, mtime)
    if key in _file_cache:
        return _file_cache[key]
    value = parser(path)
    _file_cache[key] = value
    return value


def _to_internal_codes(series, label):
    """BITPOS Item Code -> internal (old) Item_Code. Unresolved/corrupted codes stay as typed and are logged."""
    codes, report = resolve_series(series, get_resolver())
    announce(report, f'Investigation: {label}', show=None, log=True)
    return codes


# ---------- raw BITPOS POS export reader (mirrors Daily_Deduct_App.read_pos_file) ----------
def read_pos_file(filepath):
    if filepath.lower().endswith('.csv'):
        df_raw = pd.read_csv(filepath, header=None, dtype=str)
    else:
        df_raw = pd.read_excel(filepath, header=None, dtype=str)

    header_idx = None
    for i, row in df_raw.iterrows():
        row_strs = [str(x).strip().upper() for x in row.values]
        if any('ITEM CODE' in s for s in row_strs) and any('QTY' in s for s in row_strs):
            header_idx = i
            break
    if header_idx is None:
        return None

    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = [str(x).strip().title() for x in df_raw.iloc[header_idx].values]
    df = df.dropna(subset=['Item Code', 'Description'], how='all')

    if 'Item Code' in df.columns:
        df = df[~df['Item Code'].astype(str).str.contains('Total', case=False, na=False)]
        df['Item Code'] = _to_internal_codes(df['Item Code'].astype(str).str.strip(), 'BITPOS sales file')
    if 'Description' in df.columns:
        df['Description'] = df['Description'].fillna('').apply(clean_text)

    df['Qty'] = pd.to_numeric(df.get('Qty'), errors='coerce').fillna(0)
    df['_date'] = df['Date'].apply(parse_date_cell) if 'Date' in df.columns else None
    return df


def load_full_sales_history(item_code, name_norm):
    """Full available date range -- every archived BITPOS export plus
    anything still sitting in the inbox. Returns (sales_df, name_mismatches)
    where name_mismatches flags rows whose cleaned Description matches this
    item's name but NOT its Item_Code (the naming-inconsistency guard the
    user specifically asked for)."""
    files = sorted(
        glob.glob(os.path.join(ARCHIVE_DIR, '*.xls')) + glob.glob(os.path.join(ARCHIVE_DIR, '*.xlsx')) +
        glob.glob(os.path.join(ARCHIVE_DIR, '*.csv')) + glob.glob(os.path.join(INBOX_DIR, '*.xls')) +
        glob.glob(os.path.join(INBOX_DIR, '*.xlsx')) + glob.glob(os.path.join(INBOX_DIR, '*.csv'))
    )
    frames = []
    mismatches = []
    for path in files:
        try:
            df = _cached_read(path, read_pos_file)
        except Exception:
            continue
        if df is None or df.empty:
            continue
        code_match = df['Item Code'] == str(item_code)
        name_match = df['Description'].apply(norm) == name_norm
        bad = df[name_match & ~code_match]
        if not bad.empty:
            mismatches.append((os.path.basename(path), bad[['Date', 'Receipt', 'Item Code', 'Description', 'Qty']].to_dict('records')))
        m = df[code_match].copy()
        if not m.empty:
            m['source_file'] = os.path.basename(path)
            frames.append(m)
    if not frames:
        return pd.DataFrame(columns=['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', 'source_file']), mismatches
    return pd.concat(frames, ignore_index=True), mismatches


def _parse_stockin_file(path):
    df = pd.read_excel(path, dtype=str)
    if 'Item_Code' not in df.columns:
        return None
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Qty_Delivered'] = pd.to_numeric(df['Qty_Delivered'], errors='coerce').fillna(0)
    df['_date'] = df['Date_Encoded'].apply(parse_date_cell)
    return df


def load_full_stockin_history(item_code):
    frames = []
    for path in sorted(glob.glob(os.path.join(STOCKIN_DIR, '*', 'STOCKIN_BATCH_*.xlsx'))):
        try:
            df = _cached_read(path, _parse_stockin_file)
        except Exception:
            continue
        if df is None:
            continue
        m = df[df['Item_Code'] == str(item_code)].copy()
        if not m.empty:
            m['source_file'] = os.path.basename(path)
            frames.append(m)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Qty_Delivered', 'Date_Encoded', 'Encoded_By', '_date', 'source_file'])
    return pd.concat(frames, ignore_index=True)


def _parse_eod_file(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if 'Item Code' not in df.columns:
        return None
    df['Item Code'] = _to_internal_codes(df['Item Code'].astype(str).str.strip(), 'processed EOD file')
    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    df['_date'] = df['Date'].apply(parse_date_cell) if 'Date' in df.columns else None
    return df


def load_full_eod_history(item_code):
    frames = []
    for path in sorted(glob.glob(os.path.join(EOD_DIR, 'PROCESSED_EOD_*.xlsx'))):
        try:
            df = _cached_read(path, _parse_eod_file)
        except Exception:
            continue
        if df is None:
            continue
        m = df[df['Item Code'] == str(item_code)].copy()
        if not m.empty:
            m['source_file'] = os.path.basename(path)
            frames.append(m)
    if not frames:
        return pd.DataFrame(columns=['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', '_date', 'source_file'])
    return pd.concat(frames, ignore_index=True)


def _parse_backup_stock_by_item(path):
    """Returns a Series (Item_Code -> summed Current_Stock) for one backup
    file -- cached whole, so a second/third investigation in the same
    session doesn't re-open the same xlsx file just to look up a different
    Item_Code."""
    df = pd.read_excel(path, sheet_name='INVENTORY MASTER', usecols=['Item_Code', 'Current_Stock'], dtype=str)
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    return df.groupby('Item_Code')['Current_Stock'].sum()


def load_backup_timeline(item_code):
    """Every Master backup (any script's pre-write snapshot), full history --
    a genuine point-in-time record of Current_Stock for this item, ending
    with the live Master file as the final entry."""
    rows = []
    pattern = re.compile(r'Inventory_Master_BEFORE_(.+)_(\d{8}_\d{6})\.xlsx$')
    for path in glob.glob(os.path.join(BACKUP_DIR, 'Inventory_Master_*.xlsx')):
        m = pattern.search(os.path.basename(path))
        if not m:
            continue
        reason, ts_str = m.group(1), m.group(2)
        try:
            ts = datetime.datetime.strptime(ts_str, '%Y%m%d_%H%M%S')
        except ValueError:
            continue
        try:
            by_item = _cached_read(path, _parse_backup_stock_by_item)
        except Exception:
            continue
        stock = float(by_item.get(str(item_code), 0.0))
        rows.append({'timestamp': ts, 'reason': reason, 'stock': stock, 'source_file': os.path.basename(path)})
    rows.sort(key=lambda r: r['timestamp'])

    df_live = load_master()
    live_stock = df_live[df_live['Item_Code'] == str(item_code)]['Current_Stock'].sum()
    live_ts = datetime.datetime.fromtimestamp(os.path.getmtime(MASTER_FILE)) if os.path.exists(MASTER_FILE) else datetime.datetime.now()
    rows.append({'timestamp': live_ts, 'reason': 'NOW (live Master)', 'stock': live_stock, 'source_file': MASTER_FILE})
    return rows


def _parse_sod_file(path):
    """Every item block in one SOD_Summary_*.txt file, unfiltered -- cached
    whole so filtering for a different item on the next investigation is free."""
    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()
    date_match = re.search(r'Target Date:\s*([\d-]+)', text)
    auditor_match = re.search(r'Auditor:\s*(\S+)', text)
    target_date = date_match.group(1) if date_match else None
    try:
        dt = datetime.datetime.strptime(target_date, '%Y-%m-%d') if target_date else datetime.datetime.fromtimestamp(os.path.getmtime(path))
    except ValueError:
        dt = datetime.datetime.fromtimestamp(os.path.getmtime(path))

    blocks = []
    for block in text.split('---------------------------------------'):
        item_match = re.search(r'Item:\s*(.+)', block)
        qty_match = re.search(r'System Qty:\s*([\d.]+)\s*\|\s*Actual Count:\s*([\d.]+)', block)
        var_match = re.search(r'Variance:\s*(-?[\d.]+)', block)
        just_match = re.search(r'Justification:\s*(.+)', block)
        if not (item_match and qty_match):
            continue
        blocks.append({
            'item_norm': norm(item_match.group(1)),
            'date_str': target_date or '?',
            'datetime': dt,
            'system_qty': float(qty_match.group(1)),
            'actual_qty': float(qty_match.group(2)),
            'variance': float(var_match.group(1)) if var_match else None,
            'justification': just_match.group(1).strip() if just_match else '',
            'auditor': auditor_match.group(1) if auditor_match else '?',
            'source_file': os.path.basename(path),
        })
    return blocks


def load_sod_audit_history(name_norm):
    """Every SOD_Summary_*.txt occurrence of this item, chronological. Empty
    list means the item has never been part of a stratified SOD sample."""
    results = []
    for path in glob.glob(os.path.join(SOD_DIR, 'SOD_Summary_*.txt')):
        try:
            blocks = _cached_read(path, _parse_sod_file)
        except OSError:
            continue
        results.extend(b for b in blocks if b['item_norm'] == name_norm)
    results.sort(key=lambda r: r['datetime'])
    return results


def investigate_item(item_code):
    """The single entry point for the whole read-only investigation. Returns
    a plain dict (no tkinter dependency) so it can be tested headlessly."""
    df_master = load_master()
    rows = df_master[df_master['Item_Code'] == str(item_code)]
    if rows.empty:
        return None

    name_freq = rows.groupby('Product_Name').size().sort_values(ascending=False)
    canonical_name = name_freq.index[0]
    name_norm = norm(canonical_name)
    name_inconsistent = len(name_freq) > 1

    live_stock = rows['Current_Stock'].sum()

    timeline = load_backup_timeline(item_code)
    sales_df, name_mismatches = load_full_sales_history(item_code, name_norm)
    stockin_df = load_full_stockin_history(item_code)
    eod_df = load_full_eod_history(item_code)
    sod_history = load_sod_audit_history(name_norm)

    # ---- checkpoint: the most recent point Master's value is KNOWN TO BE
    # CORRECT. If the item was ever SOD-audited, that audit's Actual Count
    # is ground truth. If never audited, fall back to the earliest available
    # system record (clearly NOT a verified physical count, just the
    # earliest number we have on file).
    if sod_history:
        last_audit = sod_history[-1]
        checkpoint_stock = last_audit['actual_qty']
        checkpoint_date = last_audit['datetime'].date()
        checkpoint_label = f"Last SOD audit on {last_audit['date_str']} (physical count, verified)"
        checkpoint_verified = True
    elif timeline:
        checkpoint_stock = timeline[0]['stock']
        checkpoint_date = timeline[0]['timestamp'].date()
        checkpoint_label = f"Earliest system record found, {timeline[0]['timestamp'].strftime('%Y-%m-%d')} (not a verified physical count -- no SOD audit exists)"
        checkpoint_verified = False
    else:
        checkpoint_stock = live_stock
        checkpoint_date = None
        checkpoint_label = "No history available"
        checkpoint_verified = False

    if checkpoint_date is not None:
        # Strictly AFTER checkpoint_date, not >= : an SOD audit's Actual Count
        # verifies the closing state of that whole calendar day (it runs the
        # next morning, after that day's EOD already applied), so that day's
        # stock-in/EOD activity is already baked into checkpoint_stock. Using
        # >= here double-counts the checkpoint day's own EOD deduction --
        # exactly the bug that made LS BL CRM 7G's first pass come out wrong
        # (checkpoint 39 - that same day's -13 EOD = 26, when it should stay
        # at 39 since nothing has happened *since* the audit yet).
        stockin_since = stockin_df[stockin_df['_date'] > checkpoint_date]['Qty_Delivered'].sum() if not stockin_df.empty else 0
        eod_since = eod_df[eod_df['_date'] > checkpoint_date]['Qty'].abs().sum() if not eod_df.empty else 0
    else:
        stockin_since = stockin_df['Qty_Delivered'].sum() if not stockin_df.empty else 0
        eod_since = eod_df['Qty'].abs().sum() if not eod_df.empty else 0

    expected = checkpoint_stock + stockin_since - eod_since

    last_change = None
    for r in reversed(timeline[:-1]):  # exclude the synthetic "NOW" row
        if r['stock'] != live_stock:
            last_change = timeline[timeline.index(r) + 1]
            break
    if last_change is None and len(timeline) >= 2:
        last_change = timeline[0]

    return {
        'item_code': str(item_code),
        'canonical_name': canonical_name,
        'name_inconsistent': name_inconsistent,
        'live_stock': live_stock,
        'master_rows': rows,
        'timeline': timeline,
        'last_change': last_change,
        'sales_df': sales_df,
        'name_mismatches': name_mismatches,
        'stockin_df': stockin_df,
        'eod_df': eod_df,
        'sod_history': sod_history,
        'checkpoint_stock': checkpoint_stock,
        'checkpoint_date': checkpoint_date,
        'checkpoint_label': checkpoint_label,
        'checkpoint_verified': checkpoint_verified,
        'stockin_since': stockin_since,
        'eod_since': eod_since,
        'expected': expected,
    }


# ==========================================================================
#    PART 2 -- MANUAL STOCK ADJUSTMENT (write-capable)
#
#    The ONLY code in this entire app that writes to Inventory_Master.xlsx.
#    Every write path referenced here (MASTER_FILE, BACKUP_DIR, ADJUSTMENT_LOG,
#    PENDING_SYNC_DIR, SYSTEM_LOG_FILE) is a module-level constant so tests can
#    redirect all of them to an isolated sandbox and never touch the real files.
# ==========================================================================

def get_lock_file_path(master_file):
    directory = os.path.dirname(os.path.abspath(master_file))
    return os.path.join(directory, '~$' + os.path.basename(master_file))


def is_master_locked():
    """Proactive check: Excel creates a '~$Inventory_Master.xlsx' lock file
    while the workbook is open. This catches the common case with a clear
    message before we even attempt to write; the actual write is still
    wrapped in a PermissionError catch as the authoritative fallback (a lock
    file can be stale, and other lock mechanisms don't create one)."""
    return os.path.exists(get_lock_file_path(MASTER_FILE))


def apply_stock_adjustment_to_dataframe(df_master, item_code, new_total):
    """Mutates df_master in place: distributes the delta between the item's
    current total and new_total across its lot rows, FIFO/FEFO style --
    nearest-expiry-first is drawn down for a shortage, an overage is added to
    the nearest-expiry lot. Mirrors Daily_Dashboard_App's SOD reconciliation
    algorithm exactly, except keyed by Item_Code (the only reliable matching
    key in this codebase) instead of Product_Name. Returns (old_total, delta)."""
    mask = df_master['Item_Code'] == str(item_code)
    if not mask.any():
        raise ValueError(f"Item_Code {item_code} not found in Master.")

    old_total = df_master.loc[mask, 'Current_Stock'].sum()
    delta = new_total - old_total
    if abs(delta) < 0.0001:
        return old_total, delta

    parsed_dates = pd.to_datetime(df_master.loc[mask, 'Expiry_Date'], errors='coerce').fillna(pd.Timestamp('2099-12-31'))
    ordered_idx = parsed_dates.sort_values().index.tolist()

    if delta < 0:
        qty_to_remove = abs(delta)
        for idx in ordered_idx:
            current = df_master.at[idx, 'Current_Stock']
            if current >= qty_to_remove:
                df_master.at[idx, 'Current_Stock'] = current - qty_to_remove
                qty_to_remove = 0
                break
            qty_to_remove -= current
            df_master.at[idx, 'Current_Stock'] = 0
    else:
        df_master.at[ordered_idx[0], 'Current_Stock'] += delta

    return old_total, delta


def get_live_stock_for_item(item_code):
    """Fresh, uncached read straight from Master -- used specifically for
    stale-data checks, where the whole point is to bypass load_master()'s
    mtime cache and see the file exactly as it is right now."""
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str)
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    return df[df['Item_Code'] == str(item_code)]['Current_Stock'].sum()


def _append_adjustment_csv(ts, operation_type, item_code, name, staff, old_val, new_val, delta, reason, reference):
    is_new = not os.path.exists(ADJUSTMENT_LOG)
    with open(ADJUSTMENT_LOG, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(ADJUSTMENT_LOG_HEADER)
        writer.writerow([
            ts.strftime('%Y-%m-%d %I:%M %p'), operation_type, item_code, name, staff,
            f"{old_val:.0f}", f"{new_val:.0f}", f"{delta:+.0f}", reason, reference or '', SOURCE_TAG,
        ])


def _queue_cloud_sync(ts, operation_type, item_code, name, staff, delta, reason):
    """Same PENDING_CLOUD_SYNC JSON format Auto_Cloud_Sync.py already
    consumes, targeting the Stock_Adjustments tab. The first 7 values match
    that tab's existing live header (Date/Time, Staff, Item Code, Product
    Name, Qty Changed, Transaction Type (IN/OUT), Reason) exactly -- the
    Operation_Type and Source tags are appended as 2 extra trailing values
    (land in columns H/I). Those columns don't have header labels on the
    Sheet yet since that would mean editing live production structure --
    flag to the user separately whether to add those headers."""
    os.makedirs(PENDING_SYNC_DIR, exist_ok=True)
    txn_type = 'IN' if delta > 0 else 'OUT'
    payload = {
        'target_tab': 'Stock_Adjustments',
        'row_data': [
            ts.strftime('%Y-%m-%d %H:%M:%S'), staff.upper(), sheet_text_code(item_code), name,
            f"{abs(delta):.0f}", txn_type, reason, operation_type, SOURCE_TAG,
        ],
    }
    fname = f"sync_ADJUST_{item_code}_{ts.strftime('%Y%m%d_%H%M%S_%f')}.json"
    with open(os.path.join(PENDING_SYNC_DIR, fname), 'w', encoding='utf-8') as f:
        json.dump(payload, f)
    return fname


def commit_stock_adjustment(operation_type, item_code, canonical_name, new_total, reason, staff, reference, expected_old_value):
    """The only function in this app that writes to Master. Never raises for
    expected failure modes (stale data, locked file) -- returns a structured
    dict so the GUI can show a clear message instead of a crash."""
    if operation_type not in OPERATION_TYPES:
        return {'status': 'error', 'detail': f"Unknown operation_type: {operation_type}"}

    # 1. Stale-data check -- did Master change since the confirm screen loaded?
    live_now = get_live_stock_for_item(item_code)
    if abs(live_now - expected_old_value) > 0.0001:
        return {'status': 'stale', 'live_now': live_now, 'expected_old_value': expected_old_value}

    # 2. Proactive file-lock check
    if is_master_locked():
        return {'status': 'locked', 'detail': f"{MASTER_FILE} appears to be open in Excel (lock file present). Close it and try again."}

    # 3. Backup BEFORE any edit. Microsecond precision (not just seconds) so
    # two adjustments in quick succession can never collide on the same
    # backup filename.
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now()
    backup_path = os.path.join(BACKUP_DIR, f"Inventory_Master_BEFORE_INVESTIGATION_ADJUST_{ts.strftime('%Y%m%d_%H%M%S_%f')}.xlsx")
    try:
        shutil.copy2(MASTER_FILE, backup_path)
    except OSError as e:
        return {'status': 'error', 'detail': f"Could not create a backup before editing -- nothing was changed. Error: {e}"}

    # 4. Apply the adjustment and write (authoritative lock check -- catches
    # locks the proactive check above might miss)
    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str)
        df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
        df_master['Current_Stock'] = pd.to_numeric(df_master['Current_Stock'], errors='coerce').fillna(0)
        old_total, delta = apply_stock_adjustment_to_dataframe(df_master, item_code, new_total)
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        return {'status': 'locked', 'detail': f"{MASTER_FILE} is open in another program. A backup was made ({backup_path}) but NO changes were written."}
    except Exception as e:
        return {'status': 'error', 'detail': f"{e}\n\nA backup was made before this failure: {backup_path}"}

    # 5. System log (matches every other app's convention)
    try:
        with open(SYSTEM_LOG_FILE, "a") as f:
            f.write(f"[{ts.strftime('%Y-%m-%d %I:%M %p')}] - {operation_type} via Investigation_App ({SOURCE_TAG}) by {staff.upper()}\n")
            f.write(f"   -> {canonical_name} ({item_code}): {old_total:.0f} -> {new_total:.0f} ({delta:+.0f}) | Reason: {reason} | Ref: {reference or 'N/A'}\n")
    except OSError:
        pass

    # 6. Audit trail (CSV)
    _append_adjustment_csv(ts, operation_type, item_code, canonical_name, staff, old_total, new_total, delta, reason, reference)

    # 7. Cloud sync queue
    sync_filename = _queue_cloud_sync(ts, operation_type, item_code, canonical_name, staff, delta, reason)

    _master_cache['df'] = None  # invalidate Part 1's cache so the next investigate_item() sees the new value

    return {
        'status': 'ok', 'operation_type': operation_type, 'old_total': old_total, 'new_total': new_total, 'delta': delta,
        'backup_path': backup_path, 'sync_filename': sync_filename,
    }


# ==========================================================================
#    PART 1 -- MAIN INVESTIGATION WINDOW (GUI)
# ==========================================================================
class InvestigationApp:
    def __init__(self, staff_name):
        self.staff_name = staff_name
        self.current_result = None  # last investigate_item() result, for export/Part 2
        self.result_queue = queue.Queue()
        self.selected_item_code = None

        self.root = tk.Tk()
        self.root.title("J-DAN PHARMACY - INVESTIGATION APP (Executive)")
        self.root.geometry("1200x820")
        self.root.minsize(1000, 700)
        self.root.configure(bg=COLOR_BG)

        self._build_header()
        self._build_search_bar()
        self._build_summary_panel()
        self._build_notebook()
        self._build_footer()

        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    # ---------- layout ----------
    def _build_header(self):
        header = tk.Frame(self.root, bg=COLOR_HEADER, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="🔎 INVESTIGATION APP", fg="white", bg=COLOR_HEADER, font=("Arial", 17, "bold")).pack()
        tk.Label(header, text="Full item history, sales trace & reconciliation -- Executive Only",
                 fg=COLOR_HEADER_SUB, bg=COLOR_HEADER, font=("Arial", 9)).pack()

    def _build_search_bar(self):
        frame = tk.Frame(self.root, bg=COLOR_BG, pady=12)
        frame.pack(fill="x", padx=20)

        row = tk.Frame(frame, bg=COLOR_BG)
        row.pack(fill="x")
        tk.Label(row, text="Search Item (name or code):", font=("Arial", 11, "bold"), bg=COLOR_BG, fg=COLOR_HEADER).pack(side="left")

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(row, textvariable=self.search_var, font=("Arial", 12), width=45)
        self.search_entry.pack(side="left", padx=10, ipady=3)
        self.search_entry.bind('<KeyRelease>', self._on_search_keyrelease)
        self.search_entry.bind('<Return>', lambda e: self._investigate_from_search())

        self.investigate_btn = tk.Button(row, text="🔍 INVESTIGATE", font=("Arial", 11, "bold"),
                                          bg=COLOR_TEAL, fg="white", command=self._investigate_from_search)
        self.investigate_btn.pack(side="left", padx=5, ipady=2)

        self.suggestion_list = tk.Listbox(frame, font=("Arial", 10), height=6, activestyle='dotbox')
        self.suggestion_list.bind('<<ListboxSelect>>', self._on_suggestion_select)
        # not packed until there are suggestions to show

    def _build_summary_panel(self):
        self.summary_frame = tk.Frame(self.root, bg=COLOR_GRAY, pady=14, padx=20)
        self.summary_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.summary_status = tk.Label(self.summary_frame, text="Search for an item above to begin.",
                                        font=("Arial", 15, "bold"), bg=COLOR_GRAY, fg="white")
        self.summary_status.pack(anchor="w")

        self.summary_detail = tk.Label(self.summary_frame, text="", font=("Arial", 10), bg=COLOR_GRAY,
                                        fg="white", justify="left", wraplength=1120)
        self.summary_detail.pack(anchor="w", pady=(6, 10))

        compare_row = tk.Frame(self.summary_frame, bg=COLOR_GRAY)
        compare_row.pack(anchor="w", fill="x")
        tk.Label(compare_row, text="Physical Count (optional -- enter if you just counted):",
                 font=("Arial", 9, "bold"), bg=COLOR_GRAY, fg="white").pack(side="left")
        self.manual_count_var = tk.StringVar()
        self.manual_count_entry = tk.Entry(compare_row, textvariable=self.manual_count_var, font=("Arial", 10), width=10, state='disabled')
        self.manual_count_entry.pack(side="left", padx=8)
        self.compare_btn = tk.Button(compare_row, text="Compare Now", font=("Arial", 9, "bold"),
                                      bg=COLOR_ORANGE, fg="white", state='disabled', command=self._on_compare_now)
        self.compare_btn.pack(side="left", padx=5)

        action_row = tk.Frame(self.summary_frame, bg=COLOR_GRAY)
        action_row.pack(anchor="e", fill="x", pady=(8, 0))
        self.export_btn = tk.Button(action_row, text="📤 EXPORT REPORT", font=("Arial", 9, "bold"),
                                     bg=COLOR_HEADER, fg="white", state='disabled', command=self._export_report)
        self.export_btn.pack(side="right")
        self.adjust_btn = tk.Button(action_row, text="✏️ ADJUST STOCK", font=("Arial", 9, "bold"),
                                     bg=COLOR_RED, fg="white", state='disabled', command=self._open_adjust_step1)
        self.adjust_btn.pack(side="right", padx=(0, 8))

    def _build_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        self.tab_status = self._make_table_tab("📦 Current Status")
        self.tab_timeline = self._make_table_tab("📈 Stock Movement Timeline")
        self.tab_sales = self._make_table_tab("🧾 Sales History")
        self.tab_stockin = self._make_table_tab("📥 Stock-In Log")
        self.tab_sod = self._make_table_tab("📋 SOD Audit History")
        self.tab_recon = self._make_text_tab("🧮 Reconciliation Summary")

    def _make_table_tab(self, label):
        frame = tk.Frame(self.notebook, bg="white")
        self.notebook.add(frame, text=label)
        scroll = ttk.Scrollbar(frame, orient="vertical")
        scroll.pack(side="right", fill="y")
        tree = ttk.Treeview(frame, show="headings", yscrollcommand=scroll.set, height=20)
        scroll.config(command=tree.yview)
        tree.pack(side="left", fill="both", expand=True)
        return tree

    def _make_text_tab(self, label):
        frame = tk.Frame(self.notebook, bg="white")
        self.notebook.add(frame, text=label)
        text = tk.Text(frame, font=("Consolas", 11), wrap="word", padx=15, pady=15, state='disabled', bg="white")
        text.pack(fill="both", expand=True)
        return text

    def _build_footer(self):
        footer = tk.Frame(self.root, bg="#d5d8dc", pady=5)
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text=f"Logged in as: {self.staff_name.upper()} (Executive)", font=("Arial", 8, "bold"),
                 bg="#d5d8dc", fg=COLOR_TEXT_MUTED).pack()

    # ---------- search / autocomplete ----------
    def _on_search_keyrelease(self, event):
        if event.keysym in ('Return', 'Up', 'Down'):
            return
        query = self.search_var.get().strip().upper()
        self.suggestion_list.delete(0, tk.END)
        if not query:
            self.suggestion_list.pack_forget()
            return
        directory = get_item_directory()
        matches = [r for r in directory if query in str(r['Product_Name']).upper() or query in str(r['Item_Code'])][:15]
        if not matches:
            self.suggestion_list.pack_forget()
            return
        for r in matches:
            self.suggestion_list.insert(tk.END, f"{r['Item_Code']}  —  {r['Product_Name']}  (stock: {r['Current_Stock']:.0f})")
        self._suggestion_codes = [r['Item_Code'] for r in matches]
        self.suggestion_list.pack(fill="x", padx=0, pady=(4, 0), before=None)

    def _on_suggestion_select(self, event):
        sel = self.suggestion_list.curselection()
        if not sel:
            return
        item_code = self._suggestion_codes[sel[0]]
        self.search_var.set(self.suggestion_list.get(sel[0]))
        self.suggestion_list.pack_forget()
        self._run_investigation(item_code)

    def _investigate_from_search(self):
        query = self.search_var.get().strip()
        if not query:
            messagebox.showinfo("Search Required", "Type an item name or Item Code first.")
            return
        item_code, candidates = resolve_item_code(query)
        if item_code:
            self.suggestion_list.pack_forget()
            self._run_investigation(item_code)
            return
        if not candidates:
            messagebox.showwarning("Not Found", f"No item matches \"{query}\" in Master.")
            return
        # multiple candidates -- show them as suggestions instead of guessing
        self.suggestion_list.delete(0, tk.END)
        for r in candidates:
            self.suggestion_list.insert(tk.END, f"{r['Item_Code']}  —  {r['Product_Name']}  (stock: {r['Current_Stock']:.0f})")
        self._suggestion_codes = [r['Item_Code'] for r in candidates]
        self.suggestion_list.pack(fill="x", pady=(4, 0))
        messagebox.showinfo("Multiple Matches", f"{len(candidates)} items match \"{query}\". Please pick one from the list below the search bar.")

    # ---------- investigate ----------
    def _run_investigation(self, item_code):
        self.selected_item_code = item_code
        self.investigate_btn.config(state='disabled', text="🔄 Investigating...")
        self.summary_status.config(text="🔄 Investigating... please wait.", bg=COLOR_GRAY)
        self.summary_frame.config(bg=COLOR_GRAY)
        self.summary_detail.config(text="Reading Master backups, sales exports, stock-in records, and SOD reports. This can take a few seconds.", bg=COLOR_GRAY)
        self.root.update_idletasks()

        def worker():
            try:
                result = investigate_item(item_code)
                self.result_queue.put(('ok', result))
            except Exception as e:
                self.result_queue.put(('error', traceback.format_exc()))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(150, self._poll_investigation)

    def _poll_investigation(self):
        try:
            status, payload = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(150, self._poll_investigation)
            return

        self.investigate_btn.config(state='normal', text="🔍 INVESTIGATE")
        if status == 'error':
            messagebox.showerror("Investigation Error", f"Something went wrong:\n\n{payload}")
            self.summary_status.config(text="⚠️ Investigation failed -- see error details.", bg=COLOR_RED)
            return
        if payload is None:
            messagebox.showwarning("Not Found", "That Item_Code was not found in the current Master file.")
            return

        self.current_result = payload
        self.manual_count_entry.config(state='normal')
        self.compare_btn.config(state='normal')
        self.manual_count_var.set("")
        self.export_btn.config(state='normal')
        self.adjust_btn.config(state='normal')
        self._render_result(manual_count=None)

    def _on_compare_now(self):
        if self.current_result is None:
            return
        raw = self.manual_count_var.get().strip()
        if raw == "":
            self._render_result(manual_count=None)
            return
        try:
            manual_count = float(raw)
        except ValueError:
            messagebox.showerror("Invalid Number", "Physical Count must be a number.")
            return
        self._render_result(manual_count=manual_count)

    # ---------- rendering ----------
    def _render_result(self, manual_count):
        r = self.current_result
        expected = r['expected']

        if manual_count is not None:
            compare_value = manual_count
            compare_source = "the physical count you just entered"
            gap = expected - compare_value
            if abs(gap) < 0.01:
                self._set_summary(COLOR_GREEN, "✅ MATCHED",
                                   f"Expected: {expected:.0f} units. Found: {compare_value:.0f} units (from {compare_source}). "
                                   f"No discrepancy -- the system fully accounts for this item's stock.")
            else:
                verb = "missing" if gap > 0 else "extra"
                self._set_summary(COLOR_RED, f"⚠️ DISCREPANCY FOUND: {abs(gap):.0f} unit(s) {verb}",
                                   f"Expected: {expected:.0f} units. Found: {compare_value:.0f} units (from {compare_source}). "
                                   f"{abs(gap):.0f} unit(s) {verb}. Checkpoint used: {r['checkpoint_label']}.")
        elif r['checkpoint_verified']:
            self._set_summary(COLOR_GREEN, "✅ RECONCILED (as of last audit)",
                               f"Expected: {expected:.0f} units, based on {r['checkpoint_label']} "
                               f"plus {r['stockin_since']:.0f} stock-in and {r['eod_since']:.0f} units sold since. "
                               f"This matches Master's live stock ({r['live_stock']:.0f}). "
                               f"No new physical count entered yet -- enter one above to check for drift since the last audit.")
        else:
            self._set_summary(COLOR_ORANGE, "ℹ️ NO PRIOR SOD AUDIT EXISTS FOR THIS ITEM",
                               f"Expected (system-computed): {expected:.0f} units, based on {r['checkpoint_label']} "
                               f"plus {r['stockin_since']:.0f} stock-in and {r['eod_since']:.0f} units sold since then. "
                               f"This item has never been physically audited. Enter a Physical Count above to check for a gap.")

        self._populate_status_tab(r)
        self._populate_timeline_tab(r)
        self._populate_sales_tab(r)
        self._populate_stockin_tab(r)
        self._populate_sod_tab(r)
        self._populate_reconciliation_tab(r, manual_count)

    def _set_summary(self, color, status_text, detail_text):
        self.summary_frame.config(bg=color)
        self.summary_status.config(text=status_text, bg=color)
        self.summary_detail.config(text=detail_text, bg=color)
        for child in self.summary_frame.winfo_children():
            try:
                child.config(bg=color)
            except tk.TclError:
                pass
        for child in self.summary_frame.winfo_children():
            for grandchild in child.winfo_children():
                try:
                    if not isinstance(grandchild, (tk.Entry, tk.Button)):
                        grandchild.config(bg=color)
                except tk.TclError:
                    pass

    def _set_columns(self, tree, columns):
        tree.delete(*tree.get_children())
        tree['columns'] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=max(90, int(1000 / len(columns))), anchor='w')

    def _populate_status_tab(self, r):
        cols = ("Field", "Value")
        self._set_columns(self.tab_status, cols)
        self.tab_status.column("Field", width=260, anchor='w')
        self.tab_status.column("Value", width=700, anchor='w')
        rows = [
            ("Item Code", r['item_code']),
            ("Product Name", r['canonical_name']),
            ("Live Current Stock (Master)", f"{r['live_stock']:.0f}"),
            ("Number of lot rows in Master", f"{len(r['master_rows'])}"),
        ]
        if r['last_change']:
            rows.append(("Last recorded stock change", f"{r['last_change']['timestamp'].strftime('%Y-%m-%d %I:%M %p')} ({r['last_change']['reason']})"))
        if r['name_inconsistent']:
            rows.append(("⚠️ Name spelling", "Inconsistent across lot rows in Master -- using most common spelling."))
        for i, (k, v) in enumerate(rows):
            self.tab_status.insert('', tk.END, values=(k, v))

    def _populate_timeline_tab(self, r):
        cols = ("Date/Time", "Event", "Stock Level")
        self._set_columns(self.tab_timeline, cols)
        self.tab_timeline.column("Date/Time", width=180)
        self.tab_timeline.column("Event", width=280)
        self.tab_timeline.column("Stock Level", width=120, anchor='center')
        prev = None
        for row in r['timeline']:
            tag = ''
            if prev is not None and row['stock'] != prev:
                tag = 'change'
            self.tab_timeline.insert('', tk.END, values=(
                row['timestamp'].strftime('%Y-%m-%d %I:%M %p'), row['reason'], f"{row['stock']:.0f}"
            ), tags=(tag,) if tag else ())
            prev = row['stock']
        self.tab_timeline.tag_configure('change', background="#fff3cd")

    def _populate_sales_tab(self, r):
        cols = ("Date", "Receipt #", "Quantity Sold", "Cashier")
        self._set_columns(self.tab_sales, cols)
        df = r['sales_df']
        if df.empty:
            self.tab_sales.insert('', tk.END, values=("No sales records found", "", "", ""))
        else:
            for _, row in df.sort_values('Date').iterrows():
                self.tab_sales.insert('', tk.END, values=(
                    row.get('Date', ''), row.get('Receipt', ''), f"{row.get('Qty', 0):.0f}", row.get('Cashier Id', '')
                ))
        if r['name_mismatches']:
            total_flagged = sum(len(v) for _, v in r['name_mismatches'])
            self.tab_sales.insert('', tk.END, values=(f"⚠️ {total_flagged} row(s) matched by NAME but not Item_Code -- see export for details", "", "", ""))

    def _populate_stockin_tab(self, r):
        cols = ("Date Encoded", "Lot Number", "Expiry", "Qty Delivered", "Encoded By")
        self._set_columns(self.tab_stockin, cols)
        df = r['stockin_df']
        if df.empty:
            self.tab_stockin.insert('', tk.END, values=("No stock-in records found", "", "", "", ""))
        else:
            for _, row in df.sort_values('_date').iterrows():
                self.tab_stockin.insert('', tk.END, values=(
                    row.get('Date_Encoded', ''), row.get('Lot_Number', ''), row.get('Expiry_Date', ''),
                    f"{row.get('Qty_Delivered', 0):.0f}", row.get('Encoded_By', '')
                ))

    def _populate_sod_tab(self, r):
        cols = ("Date", "System Qty", "Actual Count", "Variance", "Justification")
        self._set_columns(self.tab_sod, cols)
        if not r['sod_history']:
            self.tab_sod.insert('', tk.END, values=("No prior SOD audit exists for this item.", "", "", "", ""))
        else:
            for s in r['sod_history']:
                tag = 'balanced' if abs(s['variance'] or 0) < 0.01 else 'variance'
                self.tab_sod.insert('', tk.END, values=(
                    s['date_str'], f"{s['system_qty']:.0f}", f"{s['actual_qty']:.0f}",
                    f"{s['variance']:+.0f}" if s['variance'] is not None else '', s['justification']
                ), tags=(tag,))
            self.tab_sod.tag_configure('variance', background="#f8d7da")
            self.tab_sod.tag_configure('balanced', background="#d4edda")

    def _populate_reconciliation_tab(self, r, manual_count):
        lines = []
        lines.append("RECONCILIATION SUMMARY")
        lines.append("=" * 60)
        lines.append(f"Item: {r['canonical_name']} ({r['item_code']})")
        lines.append("")
        lines.append(f"Checkpoint used: {r['checkpoint_label']}")
        lines.append(f"  Checkpoint stock:                {r['checkpoint_stock']:.0f}")
        lines.append(f"  + Stock-in since checkpoint:     {r['stockin_since']:.0f}")
        lines.append(f"  - Units sold since checkpoint:   {r['eod_since']:.0f}")
        lines.append(f"  = Expected stock now:            {r['expected']:.0f}")
        lines.append("")
        lines.append(f"Live Master stock right now:       {r['live_stock']:.0f}")
        match_note = "MATCHES" if abs(r['expected'] - r['live_stock']) < 0.01 else "DOES NOT MATCH -- please re-check"
        lines.append(f"  ({match_note} the expected value above)")
        lines.append("")
        if manual_count is not None:
            gap = r['expected'] - manual_count
            lines.append(f"Physical count entered:            {manual_count:.0f}")
            lines.append(f"Gap (Expected - Physical Count):   {gap:+.0f}")
        else:
            lines.append("No physical count entered this session.")
        lines.append("")
        lines.append(f"Data sources scanned: {len(r['sales_df'])} sales record(s), {len(r['stockin_df'])} stock-in record(s), "
                      f"{len(r['sod_history'])} prior SOD audit(s).")
        if r['name_mismatches']:
            lines.append(f"⚠️ {sum(len(v) for _, v in r['name_mismatches'])} sales row(s) matched by name but not Item_Code -- see export.")

        self.tab_recon.config(state='normal')
        self.tab_recon.delete('1.0', tk.END)
        self.tab_recon.insert('1.0', "\n".join(lines))
        self.tab_recon.config(state='disabled')

    # ---------- export ----------
    def _export_report(self):
        r = self.current_result
        if r is None:
            return
        os.makedirs(REPORT_DIR, exist_ok=True)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_name = re.sub(r'[^A-Za-z0-9]+', '_', r['canonical_name']).strip('_')
        out_path = os.path.join(REPORT_DIR, f'INVESTIGATION_{safe_name}_{ts}.txt')

        lines = []
        lines.append("=" * 70)
        lines.append("J-DAN PHARMACY -- ITEM INVESTIGATION REPORT")
        lines.append("=" * 70)
        lines.append(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')}")
        lines.append(f"Generated by: {self.staff_name.upper()}")
        lines.append(f"Item: {r['canonical_name']}  (Item Code: {r['item_code']})")
        lines.append("")
        lines.append(self.summary_status.cget('text'))
        lines.append(self.summary_detail.cget('text'))
        lines.append("")
        lines.append("-" * 70)
        lines.append("STOCK MOVEMENT TIMELINE")
        lines.append("-" * 70)
        for row in r['timeline']:
            lines.append(f"  {row['timestamp'].strftime('%Y-%m-%d %I:%M %p')}  [{row['reason']}]  stock={row['stock']:.0f}")
        lines.append("")
        lines.append("-" * 70)
        lines.append("SALES HISTORY")
        lines.append("-" * 70)
        if r['sales_df'].empty:
            lines.append("  No sales records found.")
        else:
            for _, row in r['sales_df'].sort_values('Date').iterrows():
                lines.append(f"  {row.get('Date','')}  Receipt #{row.get('Receipt','')}  Qty={row.get('Qty',0):.0f}  Cashier={row.get('Cashier Id','')}")
        if r['name_mismatches']:
            lines.append("")
            lines.append("  ⚠️ Rows matching by NAME but NOT Item_Code (possible naming inconsistency):")
            for fname, rows_ in r['name_mismatches']:
                for row in rows_:
                    lines.append(f"    {fname}: {row}")
        lines.append("")
        lines.append("-" * 70)
        lines.append("STOCK-IN LOG")
        lines.append("-" * 70)
        if r['stockin_df'].empty:
            lines.append("  No stock-in records found.")
        else:
            for _, row in r['stockin_df'].sort_values('_date').iterrows():
                lines.append(f"  {row.get('Date_Encoded','')}  Lot={row.get('Lot_Number','')}  Expiry={row.get('Expiry_Date','')}  "
                              f"Qty={row.get('Qty_Delivered',0):.0f}  By={row.get('Encoded_By','')}")
        lines.append("")
        lines.append("-" * 70)
        lines.append("SOD AUDIT HISTORY")
        lines.append("-" * 70)
        if not r['sod_history']:
            lines.append("  No prior SOD audit exists for this item.")
        else:
            for s in r['sod_history']:
                lines.append(f"  {s['date_str']}  System={s['system_qty']:.0f}  Actual={s['actual_qty']:.0f}  "
                              f"Variance={s['variance']:+.0f}  ({s['justification']})  Auditor={s['auditor']}")
        lines.append("")
        lines.append("-" * 70)
        lines.append("RECONCILIATION DETAIL")
        lines.append("-" * 70)
        self.tab_recon.config(state='normal')
        lines.append(self.tab_recon.get('1.0', tk.END).strip())
        self.tab_recon.config(state='disabled')

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))

        messagebox.showinfo("Exported", f"Investigation report saved to:\n{out_path}")

    # ---------- PART 2: manual stock adjustment (two-step, write-capable) ----------
    def _open_adjust_step1(self):
        r = self.current_result
        if r is None:
            return

        # Fresh snapshot right now -- this is the "old value" the whole
        # operation is based on. Re-checked again before Step 2 is shown,
        # and a THIRD time at the moment of the actual write.
        live_now = get_live_stock_for_item(r['item_code'])

        win = tk.Toplevel(self.root)
        win.title("Step 1 of 2 -- Stock Operation")
        win.geometry("560x680")
        win.configure(bg=COLOR_BG)
        win.grab_set()
        win.transient(self.root)

        tk.Label(win, text="STEP 1 OF 2 -- SELECT OPERATION & ENTER DETAILS", font=("Arial", 13, "bold"),
                 bg=COLOR_HEADER, fg="white", pady=10, wraplength=540).pack(fill="x")

        body = tk.Frame(win, bg=COLOR_BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=f"{r['canonical_name']}  (Item Code: {r['item_code']})", font=("Arial", 11, "bold"),
                 bg=COLOR_BG, fg=COLOR_HEADER, wraplength=500, justify="left").pack(anchor="w")
        tk.Label(body, text=f"Current Master stock (as of right now): {live_now:.0f}", font=("Arial", 10),
                 bg=COLOR_BG, fg=COLOR_TEXT_MUTED).pack(anchor="w", pady=(2, 15))

        # ---- Operation type selector ----
        tk.Label(body, text="Operation Type (required):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        op_type_var = tk.StringVar(value="")
        op_frame = tk.Frame(body, bg=COLOR_BG)
        op_frame.pack(anchor="w", pady=(2, 15))
        tk.Radiobutton(op_frame, text="📥 Stock In", variable=op_type_var, value="STOCK_IN",
                        bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_GREEN).pack(side="left", padx=(0, 15))
        tk.Radiobutton(op_frame, text="📤 Pull Out", variable=op_type_var, value="PULL_OUT",
                        bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_RED).pack(side="left", padx=(0, 15))
        tk.Radiobutton(op_frame, text="⚖️ Adjustment (match physical count)", variable=op_type_var, value="ADJUSTMENT",
                        bg=COLOR_BG, font=("Arial", 10, "bold"), fg=COLOR_ORANGE).pack(side="left")

        # ---- Dynamic content area: one frame per operation type ----
        dynamic_area = tk.Frame(body, bg=COLOR_BG)
        dynamic_area.pack(anchor="w", fill="x")

        # -- STOCK_IN fields --
        stock_in_frame = tk.Frame(dynamic_area, bg=COLOR_BG)
        tk.Label(stock_in_frame, text="Quantity to Add:", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        stock_in_qty_var = tk.StringVar()
        tk.Entry(stock_in_frame, textvariable=stock_in_qty_var, font=("Arial", 13), width=15).pack(anchor="w", pady=(2, 12), ipady=3)
        tk.Label(stock_in_frame, text="Reason / Justification (required):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        stock_in_reason_text = tk.Text(stock_in_frame, font=("Arial", 10), width=54, height=3, wrap="word")
        stock_in_reason_text.pack(anchor="w", pady=(2, 12))

        # -- PULL_OUT fields --
        pull_out_frame = tk.Frame(dynamic_area, bg=COLOR_BG)
        tk.Label(pull_out_frame, text="Quantity to Remove:", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        pull_out_qty_var = tk.StringVar()
        tk.Entry(pull_out_frame, textvariable=pull_out_qty_var, font=("Arial", 13), width=15).pack(anchor="w", pady=(2, 12), ipady=3)
        tk.Label(pull_out_frame, text="Reason (required -- select one):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        pull_out_reason_var = tk.StringVar()
        pull_out_combo = ttk.Combobox(pull_out_frame, textvariable=pull_out_reason_var, values=PULL_OUT_REASONS,
                                       state='readonly', font=("Arial", 10), width=25)
        pull_out_combo.pack(anchor="w", pady=(2, 12))

        # -- ADJUSTMENT fields (original flow) --
        adjustment_frame = tk.Frame(dynamic_area, bg=COLOR_BG)
        mode_var = tk.StringVar(value="absolute")
        mode_frame = tk.Frame(adjustment_frame, bg=COLOR_BG)
        mode_frame.pack(anchor="w", pady=(0, 5))
        tk.Radiobutton(mode_frame, text="Set New Actual Count", variable=mode_var, value="absolute",
                        bg=COLOR_BG, font=("Arial", 10)).pack(side="left", padx=(0, 15))
        tk.Radiobutton(mode_frame, text="Apply +/- Delta", variable=mode_var, value="delta",
                        bg=COLOR_BG, font=("Arial", 10)).pack(side="left")
        adj_value_var = tk.StringVar()
        tk.Entry(adjustment_frame, textvariable=adj_value_var, font=("Arial", 13), width=15).pack(anchor="w", pady=(2, 12), ipady=3)
        tk.Label(adjustment_frame, text="Reason / Justification (required):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        adj_reason_text = tk.Text(adjustment_frame, font=("Arial", 10), width=54, height=3, wrap="word")
        adj_reason_text.pack(anchor="w", pady=(2, 12))

        op_frames = {'STOCK_IN': stock_in_frame, 'PULL_OUT': pull_out_frame, 'ADJUSTMENT': adjustment_frame}

        def on_op_type_change(*_args):
            for f in op_frames.values():
                f.pack_forget()
            chosen = op_frames.get(op_type_var.get())
            if chosen:
                chosen.pack(anchor="w", fill="x")
            validate()

        # ---- Common fields: Staff + Reference ----
        tk.Label(body, text="Staff Making This Entry (required):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        staff_var = tk.StringVar()
        staff_combo = ttk.Combobox(body, textvariable=staff_var, values=STAFF_LIST, state='readonly', font=("Arial", 10), width=20)
        staff_combo.pack(anchor="w", pady=(2, 12))

        tk.Label(body, text="Reference (optional -- receipt #, lot #, memo, etc.):", font=("Arial", 10, "bold"), bg=COLOR_BG).pack(anchor="w")
        reference_var = tk.StringVar()
        tk.Entry(body, textvariable=reference_var, font=("Arial", 10), width=40).pack(anchor="w", pady=(2, 12))

        error_label = tk.Label(body, text="", font=("Arial", 9, "bold"), bg=COLOR_BG, fg=COLOR_RED, wraplength=500, justify="left")
        error_label.pack(anchor="w", pady=(0, 8))

        next_btn = tk.Button(body, text="Next: Review & Confirm >>", font=("Arial", 11, "bold"),
                              bg=COLOR_GRAY, fg="white", state='disabled')
        next_btn.pack(fill="x", ipady=8, pady=(5, 0), side="bottom")
        tk.Button(body, text="Cancel", font=("Arial", 10), bg="#6c757d", fg="white",
                  command=win.destroy).pack(fill="x", pady=(8, 0), side="bottom")

        def validate(*_args):
            op = op_type_var.get()
            staff_ok = bool(staff_var.get().strip())
            fields_ok = False
            if op == 'STOCK_IN':
                try:
                    fields_ok = float(stock_in_qty_var.get().strip()) > 0
                except ValueError:
                    fields_ok = False
                fields_ok = fields_ok and bool(stock_in_reason_text.get('1.0', tk.END).strip())
            elif op == 'PULL_OUT':
                try:
                    fields_ok = float(pull_out_qty_var.get().strip()) > 0
                except ValueError:
                    fields_ok = False
                fields_ok = fields_ok and bool(pull_out_reason_var.get().strip())
            elif op == 'ADJUSTMENT':
                try:
                    float(adj_value_var.get().strip())
                    fields_ok = True
                except ValueError:
                    fields_ok = False
                fields_ok = fields_ok and bool(adj_reason_text.get('1.0', tk.END).strip())

            if op and fields_ok and staff_ok:
                next_btn.config(state='normal', bg=COLOR_GREEN, command=go_next)
            else:
                next_btn.config(state='disabled', bg=COLOR_GRAY)

        def go_next():
            error_label.config(text="")
            op = op_type_var.get()
            staff = staff_var.get().strip()
            reference = reference_var.get().strip()

            if op == 'STOCK_IN':
                qty = float(stock_in_qty_var.get().strip())
                new_total = live_now + qty
                reason = stock_in_reason_text.get('1.0', tk.END).strip()
            elif op == 'PULL_OUT':
                qty = float(pull_out_qty_var.get().strip())
                new_total = live_now - qty
                if new_total < 0:
                    error_label.config(text=f"Cannot pull out {qty:.0f} -- only {live_now:.0f} currently in Master.")
                    return
                reason = pull_out_reason_var.get().strip()
            else:  # ADJUSTMENT
                raw = float(adj_value_var.get().strip())
                new_total = raw if mode_var.get() == 'absolute' else (live_now + raw)
                if new_total < 0:
                    error_label.config(text="The resulting new count cannot be negative.")
                    return
                reason = adj_reason_text.get('1.0', tk.END).strip()

            win.destroy()
            self._open_adjust_step2(r, op, live_now, new_total, reason, staff, reference)

        op_type_var.trace_add('write', on_op_type_change)
        stock_in_qty_var.trace_add('write', validate)
        pull_out_qty_var.trace_add('write', validate)
        adj_value_var.trace_add('write', validate)
        staff_combo.bind('<<ComboboxSelected>>', validate)
        pull_out_combo.bind('<<ComboboxSelected>>', validate)
        stock_in_reason_text.bind('<KeyRelease>', validate)
        adj_reason_text.bind('<KeyRelease>', validate)

    OP_TYPE_DISPLAY = {'STOCK_IN': '📥 STOCK IN', 'PULL_OUT': '📤 PULL OUT', 'ADJUSTMENT': '⚖️ ADJUSTMENT'}

    def _open_adjust_step2(self, r, op, expected_old_value, new_total, reason, staff, reference):
        # Re-check staleness BEFORE even showing the confirm screen, so we
        # never display numbers that are already known to be wrong.
        live_now = get_live_stock_for_item(r['item_code'])
        if abs(live_now - expected_old_value) > 0.0001:
            proceed = messagebox.askyesno(
                "Master Value Changed",
                f"Master's stock for {r['canonical_name']} changed from {expected_old_value:.0f} to {live_now:.0f} "
                f"since you started this {self.OP_TYPE_DISPLAY.get(op, op).split(' ', 1)[-1].title()} entry (an EOD run or another edit may have happened).\n\n"
                f"Re-investigating this item and starting over is strongly recommended.\n\n"
                f"Click 'No' to go back and re-investigate (recommended), or 'Yes' to continue anyway using "
                f"the CURRENT value ({live_now:.0f}) as the baseline instead."
            )
            if not proceed:
                return
            expected_old_value = live_now

        delta = new_total - expected_old_value

        win = tk.Toplevel(self.root)
        win.title("Step 2 of 2 -- Confirm Adjustment")
        win.geometry("520x460")
        win.configure(bg=COLOR_BG)
        win.grab_set()
        win.transient(self.root)

        tk.Label(win, text="STEP 2 OF 2 -- REVIEW & CONFIRM", font=("Arial", 13, "bold"),
                 bg=COLOR_ORANGE, fg="white", pady=10).pack(fill="x")

        body = tk.Frame(win, bg=COLOR_BG, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        tk.Label(body, text=f"{r['canonical_name']}  (Item Code: {r['item_code']})", font=("Arial", 11, "bold"),
                 bg=COLOR_BG, fg=COLOR_HEADER, wraplength=460, justify="left").pack(anchor="w", pady=(0, 5))
        tk.Label(body, text=f"Operation: {self.OP_TYPE_DISPLAY.get(op, op)}", font=("Arial", 11, "bold"),
                 bg=COLOR_BG, fg=COLOR_HEADER).pack(anchor="w", pady=(0, 15))

        table = tk.Frame(body, bg="white", relief="solid", borderwidth=1)
        table.pack(fill="x", pady=(0, 15))
        for i, (label, value) in enumerate([
            ("Old Value", f"{expected_old_value:.0f}"),
            ("New Value", f"{new_total:.0f}"),
            ("Difference", f"{delta:+.0f}"),
        ]):
            tk.Label(table, text=label, font=("Arial", 10, "bold"), bg="white", anchor="w", width=18).grid(row=i, column=0, sticky="w", padx=10, pady=6)
            tk.Label(table, text=value, font=("Arial", 12, "bold"), bg="white",
                     fg=(COLOR_GREEN if delta >= 0 else COLOR_RED)).grid(row=i, column=1, sticky="w", padx=10, pady=6)

        tk.Label(body, text=f"Reason: {reason}", font=("Arial", 10), bg=COLOR_BG, wraplength=460, justify="left").pack(anchor="w", pady=(0, 5))
        tk.Label(body, text=f"Staff: {staff}", font=("Arial", 10), bg=COLOR_BG).pack(anchor="w", pady=(0, 5))
        tk.Label(body, text=f"Reference: {reference or '(none)'}", font=("Arial", 10), bg=COLOR_BG).pack(anchor="w", pady=(0, 15))

        status_label = tk.Label(body, text="", font=("Arial", 9, "bold"), bg=COLOR_BG, fg=COLOR_RED, wraplength=460, justify="left")
        status_label.pack(anchor="w", pady=(0, 10))

        def do_confirm():
            confirm_btn.config(state='disabled', text="Saving...")
            win.update_idletasks()
            result = commit_stock_adjustment(op, r['item_code'], r['canonical_name'], new_total, reason, staff, reference, expected_old_value)
            if result['status'] == 'ok':
                messagebox.showinfo(
                    f"{self.OP_TYPE_DISPLAY.get(op, op)} Saved",
                    f"Saved.\n\n{r['canonical_name']}: {result['old_total']:.0f} -> {result['new_total']:.0f} "
                    f"({result['delta']:+.0f})\n\nBackup: {result['backup_path']}\nQueued for cloud sync: {result['sync_filename']}"
                )
                win.destroy()
                self._run_investigation(r['item_code'])  # refresh Part 1 with the new value
            elif result['status'] == 'stale':
                status_label.config(text=f"Master changed again (now {result['live_now']:.0f}) -- please close this and re-investigate.")
                confirm_btn.config(state='normal', text="✅ CONFIRM & SAVE")
            elif result['status'] == 'locked':
                status_label.config(text=result['detail'])
                confirm_btn.config(state='normal', text="✅ CONFIRM & SAVE")
            else:
                status_label.config(text=f"Error: {result.get('detail', 'unknown error')}")
                confirm_btn.config(state='normal', text="✅ CONFIRM & SAVE")

        confirm_btn = tk.Button(body, text="✅ CONFIRM & SAVE", font=("Arial", 11, "bold"),
                                 bg=COLOR_GREEN, fg="white", command=do_confirm)
        confirm_btn.pack(fill="x", ipady=8)
        tk.Button(body, text="Cancel", font=("Arial", 10), bg="#6c757d", fg="white",
                  command=win.destroy).pack(fill="x", pady=(8, 0))

    def run(self):
        self.root.mainloop()


def launch_main_app(staff_name="Dan"):
    InvestigationApp(staff_name).run()


if __name__ == "__main__":
    LoginWindow(on_success=launch_main_app).run()
