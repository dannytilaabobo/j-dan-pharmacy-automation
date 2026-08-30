from flask import Flask, render_template, request, redirect, url_for, Response, session
import sqlite3
from datetime import datetime, time
from werkzeug.security import generate_password_hash, check_password_hash
import io
import csv
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import threading
import os
import calendar

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
# This app uses bare relative paths for pharmacy_data.db and credentials.json,
# so without this it silently fails to find them when cwd isn't this folder.
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# --- LOCAL SECRETS (.env, never committed) ---
# Minimal KEY=VALUE loader so no extra dependency is needed. Ignores blank
# lines and # comments; does not override variables already set in the real
# environment. Real values live in .env (gitignored); see .env.example.
def _load_dotenv(path='.env'):
    try:
        with open(path, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, value = line.partition('=')
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))
    except FileNotFoundError:
        pass

def _require_env(key):
    value = os.environ.get(key)
    if not value:
        raise SystemExit(f"Missing required secret '{key}'. Copy .env.example to .env and fill it in.")
    return value

_load_dotenv()

app = Flask(__name__)
app.secret_key = _require_env('FLASK_SECRET_KEY')

# Effective 2026-07-15: rate is 80/hour going forward. Past records/logs are
# not recomputed -- this only affects gross pay for hours entered from now on.
HOURLY_RATE = 80

# --- EXCESS TIME-OUT CAP (payroll computation only) ---
# Effective going forward only. Existing records are NOT recomputed.
GRACE_PERIOD_MINUTES = 15  # constant, easy to adjust later

def apply_out_time_cap(time_out):
    """
    Trims the OUT time FOR PAYROLL COMPUTATION ONLY.
    This is never used when storing the raw punch -- the actual/raw time is
    always what gets saved to time_out / raw_time_out.

    Logic: each hour has 2 blocks (:00-:29 and :30-:59). Within the first
    15 minutes of a block the actual time is used (grace). Beyond that
    (minutes 16-29 of the block) the time is capped to the end of the
    grace window (:15 / :45).
    """
    block_start = 0 if time_out.minute < 30 else 30
    minutes_into_block = time_out.minute - block_start
    if minutes_into_block <= GRACE_PERIOD_MINUTES:
        return time_out
    return time_out.replace(minute=block_start + GRACE_PERIOD_MINUTES,
                            second=0, microsecond=0)

# --- ANTI-CHEAT FIREWALL (ALLOW REMOTE ADMIN & STAFF PORTAL, BLOCK REMOTE CLOCK-IN) ---
@app.before_request
def block_remote_clock_ins():
    allowed_remote_routes = [
        'portal_login', 'my_dashboard', 'portal_logout', 'change_pin', 
        'login', 'logout', 'admin', 'staff', 'view_logs', 'update_log', 
        'add_adjustment', 'insert_missing_log', 'delete_log', 'export_csv', 'payslip',
        'static'
    ]
    is_remote = request.headers.get('X-Forwarded-For') is not None
    
    if is_remote and request.endpoint not in allowed_remote_routes:
        return """
        <h1 style='color: #c62828; text-align: center; margin-top: 50px; font-family: Arial;'>
            ACCESS DENIED
        </h1>
        <p style='text-align: center; font-family: Arial;'>
            Time In / Time Out functions can only be performed on the physical computer inside Happy Med Pharmacy.
        </p>
        <div style='text-align: center; margin-top: 20px;'>
            <a href='/portal_login' style='background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-family: Arial;'>Go to Staff Portal</a>
        </div>
        """, 403

# --- SYSTEM SETTINGS ---
LATE_THRESHOLD = time(6, 51) 
LATE_CUTOFF = time(9, 0)
PENALTY = 200
ADMIN_PASSWORD = _require_env('ADMIN_PASSWORD')
PIN_EXPIRY_DAYS = 60

# --- PHILIPPINE HOLIDAYS (2026) ---
PH_HOLIDAYS = {
    '01-01': "New Year's Day", '02-17': "Chinese New Year", '02-25': "EDSA People Power Revolution Anniversary",
    '04-09': "Araw ng Kagitingan", '04-02': "Maundy Thursday", '04-03': "Good Friday", '04-04': "Black Saturday",
    '05-01': "Labor Day", '06-12': "Independence Day", '08-21': "Ninoy Aquino Day", '08-31': "National Heroes Day",
    '11-01': "All Saints' Day", '11-02': "All Souls' Day", '11-30': "Bonifacio Day", '12-08': "Feast of the Immaculate Conception",
    '12-24': "Christmas Eve", '12-25': "Christmas Day", '12-30': "Rizal Day", '12-31': "New Year's Eve"
}

# --- GOOGLE SHEETS REAL-TIME SYNC ---
def sync_to_google_sheet(emp_name, date_str, time_str, action, remarks=""):
    def task():
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
            client = gspread.authorize(creds)
            sheet = client.open("JDAN_Timekeeping_Monitor").sheet1
            sheet.append_row([date_str, emp_name, action, time_str, remarks])
        except Exception as e:
            print(f"Google Sheets Sync Error: {e}")
    threading.Thread(target=task).start()

# --- PAYROLL EDIT AUDIT LOG ---
# Every manual edit of a Time In / Time Out value is appended here. This file
# is local-only (gitignored) and is never overwritten -- append only.
PAYROLL_EDIT_AUDIT_LOG = 'PAYROLL_EDIT_AUDIT_LOG.csv'

def log_payroll_edit(log_date, staff_name, field_edited, old_value, new_value, edited_by):
    """Append one row per changed field to PAYROLL_EDIT_AUDIT_LOG.csv."""
    try:
        new_file = not os.path.exists(PAYROLL_EDIT_AUDIT_LOG)
        with open(PAYROLL_EDIT_AUDIT_LOG, 'a', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(['Date', 'Staff', 'Field_Edited', 'Old_Value',
                            'New_Value', 'Edited_By', 'Edit_Timestamp'])
            w.writerow([log_date, staff_name, field_edited,
                        old_value if old_value not in (None, '') else '(blank)',
                        new_value if new_value not in (None, '') else '(blank)',
                        edited_by, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    except Exception as e:
        print(f"Payroll Edit Audit Log Error: {e}")

# --- DATABASE SETUP ---
def init_db():
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS employees (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, pin_hash TEXT, is_temp INTEGER DEFAULT 1)''')
    c.execute('''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, emp_id INTEGER, date TEXT, time_in TEXT, time_out TEXT, hours REAL, penalty REAL DEFAULT 0)''')
    conn.commit()
    try: c.execute("ALTER TABLE logs ADD COLUMN penalty REAL DEFAULT 0"); conn.commit()
    except: pass
    try: c.execute("ALTER TABLE employees ADD COLUMN role TEXT DEFAULT 'Pharmacy Assistant'"); conn.commit()
    except: pass
    try: c.execute("ALTER TABLE logs ADD COLUMN adjustment_hours REAL DEFAULT 0"); conn.commit()
    except: pass
    try: c.execute("ALTER TABLE logs ADD COLUMN adjustment_reason TEXT"); conn.commit()
    except: pass
    try: c.execute("ALTER TABLE employees ADD COLUMN pin_last_changed TEXT"); conn.commit()
    except: pass
    # Raw punch reference columns -- preserve the original punch even after an
    # admin edits time_in / time_out. Backfilled once from current values.
    try: c.execute("ALTER TABLE logs ADD COLUMN raw_time_in TEXT"); conn.commit()
    except: pass
    try: c.execute("ALTER TABLE logs ADD COLUMN raw_time_out TEXT"); conn.commit()
    except: pass
    try:
        c.execute("UPDATE logs SET raw_time_in = time_in WHERE raw_time_in IS NULL AND time_in IS NOT NULL AND time_in != 'ADJUSTMENT'")
        c.execute("UPDATE logs SET raw_time_out = time_out WHERE raw_time_out IS NULL AND time_out IS NOT NULL AND time_out != 'ADJUSTMENT'")
        conn.commit()
    except: pass
    conn.close()

# --- ADMIN LOGIN PROTECTION ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    next_url = request.args.get('next')
    if request.method == 'POST':
        if request.form['password'] == ADMIN_PASSWORD:
            session['logged_in'] = True
            return redirect(next_url or url_for('admin'))
        else:
            return render_template('login.html', error="Invalid Master Password.")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('home'))

# --- MAIN TIMEKEEPING ---
# A single process-wide lock serializes the check-then-insert critical section
# for clock punches, so two near-simultaneous taps from the same kiosk can't
# both pass the "no open Time In" check and both insert/close a record.
clock_lock = threading.Lock()

def _fmt_dt(date_str, time_str):
    try:
        return datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S").strftime("%B %d, %Y — %I:%M:%S %p")
    except Exception:
        return f"{date_str} {time_str}"

def _get_open_time_in(c, emp_id, date):
    c.execute("""SELECT id, time_in FROM logs
                 WHERE emp_id=? AND date=? AND time_in IS NOT NULL AND time_in != 'ADJUSTMENT' AND time_out IS NULL
                 ORDER BY id DESC LIMIT 1""", (emp_id, date))
    return c.fetchone()

def _get_latest_completed_today(c, emp_id, date):
    c.execute("""SELECT id, time_in, time_out, hours FROM logs
                 WHERE emp_id=? AND date=? AND time_in IS NOT NULL AND time_in != 'ADJUSTMENT' AND time_out IS NOT NULL
                 ORDER BY id DESC LIMIT 1""", (emp_id, date))
    return c.fetchone()

def handle_time_in(c, conn, emp_id, emp_name, today, now):
    existing = _get_open_time_in(c, emp_id, today)
    if existing:
        _, time_in_str = existing
        return {
            'status': 'info', 'action': 'TIME IN', 'name': emp_name,
            'when': _fmt_dt(today, time_in_str),
            'note': "You are already clocked in for today. This is your recorded Time In. No new entry was made."
        }

    time_now_str = now.strftime("%H:%M:%S")
    current_t = now.time()
    is_late = LATE_THRESHOLD <= current_t <= LATE_CUTOFF
    penalty = PENALTY if is_late else 0
    remarks = "LATE" if is_late else ""

    c.execute("INSERT INTO logs (emp_id, date, time_in, raw_time_in, penalty) VALUES (?, ?, ?, ?, ?)", (emp_id, today, time_now_str, time_now_str, penalty))
    conn.commit()
    sync_to_google_sheet(emp_name, today, time_now_str, "Time In", remarks)

    return {
        'status': 'success', 'action': 'TIME IN', 'name': emp_name,
        'when': _fmt_dt(today, time_now_str),
        'note': "This entry has been flagged as late." if is_late else ""
    }

def handle_time_out(c, conn, emp_id, emp_name, today, now):
    open_log = _get_open_time_in(c, emp_id, today)

    if not open_log:
        completed = _get_latest_completed_today(c, emp_id, today)
        if completed:
            _, t_in_str, t_out_str, hours = completed
            return {
                'status': 'info', 'action': 'TIME OUT', 'name': emp_name,
                'when': _fmt_dt(today, t_out_str), 'hours': hours,
                'note': "You have already clocked out for today. This is your recorded Time Out. No new entry was made."
            }
        return {
            'status': 'error', 'action': 'TIME OUT', 'name': emp_name,
            'note': "No Time In record was found for today. You must clock in before you can clock out."
        }

    log_id, time_in_str = open_log
    time_now_str = now.strftime("%H:%M:%S")  # actual/raw punch -- always stored as-is
    t_in = datetime.strptime(time_in_str, "%H:%M:%S")
    # Excess time-out cap applies to the PAYROLL COMPUTATION only, not the stored punch.
    capped_now = apply_out_time_cap(now)
    raw_hours = (capped_now - t_in.replace(year=capped_now.year, month=capped_now.month, day=capped_now.day)).total_seconds() / 3600

    # --- SHORT-SHIFT ANOMALY PROTOCOL ---
    if raw_hours < 1.0:
        c.execute("UPDATE logs SET time_out=?, raw_time_out=?, hours=?, adjustment_reason=? WHERE id=?",
                  (time_now_str, time_now_str, round(raw_hours, 2), "SYSTEM FLAG: Shift under 1 hour", log_id))
        conn.commit()
        sync_to_google_sheet(emp_name, today, time_now_str, "Time Out (FLAGGED)", "Shift < 1 Hour")
        return {
            'status': 'error', 'action': 'TIME OUT', 'name': emp_name,
            'when': _fmt_dt(today, time_now_str),
            'note': "This shift did not meet the minimum duration of 1 hour and has been flagged for management review."
        }

    # --- STANDARD TIME OUT ---
    break_multiplier = int(raw_hours // 6)
    final_hours = max(0, raw_hours - (break_multiplier * 1.25))

    c.execute("UPDATE logs SET time_out=?, raw_time_out=?, hours=? WHERE id=?", (time_now_str, time_now_str, round(final_hours, 2), log_id))
    conn.commit()
    sync_to_google_sheet(emp_name, today, time_now_str, "Time Out", "")

    return {
        'status': 'success', 'action': 'TIME OUT', 'name': emp_name,
        'when': _fmt_dt(today, time_now_str), 'hours': round(final_hours, 2)
    }

@app.route('/', methods=['GET', 'POST'])
def home():
    message = ""
    banner = None
    if request.method == 'POST':
        pin = request.form['pin']
        action = request.form.get('action')
        now = datetime.now()
        conn = sqlite3.connect('pharmacy_data.db')
        c = conn.cursor()
        c.execute("SELECT id, name, pin_hash, is_temp, pin_last_changed FROM employees")
        employees = c.fetchall()

        staff = next((e for e in employees if check_password_hash(e[2], pin)), None)

        if not staff:
            conn.close()
            return render_template('index.html', message="Invalid PIN. Please try again.")

        if action not in ('time_in', 'time_out'):
            conn.close()
            return render_template('index.html', message="Invalid request. Please select TIME IN or TIME OUT.")

        is_expired = False
        days_left = PIN_EXPIRY_DAYS

        if staff[4]:
            try:
                last_changed = datetime.strptime(staff[4], "%Y-%m-%d")
                days_since = (now - last_changed).days
                days_left = PIN_EXPIRY_DAYS - days_since
                if days_left <= 0:
                    is_expired = True
            except: pass
        else:
            is_expired = True

        if staff[3] == 1 or is_expired:
            conn.close()
            return redirect(url_for('change_pin', emp_id=staff[0]))

        emp_id, emp_name = staff[0], staff[1]
        today = now.strftime("%Y-%m-%d")

        with clock_lock:
            if action == 'time_out':
                banner = handle_time_out(c, conn, emp_id, emp_name, today, now)
            else:
                banner = handle_time_in(c, conn, emp_id, emp_name, today, now)

        conn.close()

        if 1 <= days_left <= 7:
            return render_template('index.html', banner=banner, warning=True, days_left=days_left, emp_id=emp_id)
        return render_template('index.html', banner=banner)

    return render_template('index.html', message=message)

# --- PROTECTED ROUTES (NEEDS ADMIN LOGIN) ---
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if 'logged_in' not in session: return redirect(url_for('login', next=request.url))
    data = []
    start_date = end_date = ""
    late_only = False
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("SELECT id, name FROM employees")
    employees = c.fetchall()
    if request.method == 'POST':
        if 'start_date' in request.form:
            start_date = request.form['start_date']
            end_date = request.form['end_date']
            late_only = 'late_only' in request.form
            c.execute("SELECT e.id, e.name, SUM(l.hours), SUM(l.penalty), SUM(l.adjustment_hours) FROM logs l JOIN employees e ON l.emp_id = e.id WHERE l.date BETWEEN ? AND ? GROUP BY l.emp_id", (start_date, end_date))
            for row in c.fetchall():
                hrs = row[2] if row[2] is not None else 0
                pen = row[3] if row[3] is not None else 0
                adj = row[4] if row[4] is not None else 0
                if late_only and pen <= 0: continue
                total_payable_hours = hrs + adj
                gross_pay = (total_payable_hours * HOURLY_RATE)
                data.append({'id': row[0], 'name': row[1], 'hours': round(hrs, 2), 'adjustment': round(adj, 2), 'total_hours': round(total_payable_hours, 2), 'penalty': pen, 'gross': round(gross_pay, 2)})
    conn.close()
    return render_template('admin.html', data=data, start_date=start_date, end_date=end_date, employees=employees, late_only=late_only)

@app.route('/staff', methods=['GET', 'POST'])
def staff():
    if 'logged_in' not in session: return redirect(url_for('login', next=request.url))
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    if request.method == 'POST':
        if 'emp_name' in request.form:
            role = request.form.get('role', 'Pharmacy Assistant')
            c.execute("INSERT INTO employees (name, pin_hash, role, is_temp) VALUES (?, ?, ?, 1)", (request.form['emp_name'], generate_password_hash(request.form['temp_pin']), role))
        elif 'reset_id' in request.form:
            c.execute("UPDATE employees SET pin_hash=?, is_temp=1 WHERE id=?", (generate_password_hash(request.form['new_temp_pin']), request.form['reset_id']))
        conn.commit()
    c.execute("SELECT id, name, role FROM employees"); staff_list = c.fetchall(); conn.close()
    return render_template('staff.html', staff_list=staff_list)

@app.route('/view_logs/<int:emp_id>', methods=['GET'])
def view_logs(emp_id):
    if 'logged_in' not in session: return redirect(url_for('login', next=request.url))
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    late_only = request.args.get('late_only')
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("SELECT name FROM employees WHERE id=?", (emp_id,))
    emp_name = c.fetchone()[0]
    if late_only == '1':
        c.execute("SELECT id, date, time_in, time_out, hours, penalty, adjustment_hours, adjustment_reason, raw_time_in, raw_time_out FROM logs WHERE emp_id=? AND date BETWEEN ? AND ? AND penalty > 0 ORDER BY date ASC", (emp_id, start_date, end_date))
    else:
        c.execute("SELECT id, date, time_in, time_out, hours, penalty, adjustment_hours, adjustment_reason, raw_time_in, raw_time_out FROM logs WHERE emp_id=? AND date BETWEEN ? AND ? ORDER BY date ASC", (emp_id, start_date, end_date))
    logs = c.fetchall()
    conn.close()
    return render_template('view_logs.html', emp_name=emp_name, logs=logs, start_date=start_date, end_date=end_date, emp_id=emp_id, late_only=late_only)

def _parse_hms(value):
    """Accept 'HH:MM' or 'HH:MM:SS' from a time picker; return a datetime."""
    value = (value or "").strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognized time value: {value!r}")

@app.route('/update_log', methods=['POST'])
def update_log():
    if 'logged_in' not in session: return redirect(url_for('login', next=request.url))
    log_id = request.form['log_id']
    emp_id = request.form['emp_id']
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    time_in = request.form['time_in']
    time_out = request.form['time_out']
    late_only = request.form.get('late_only', '')
    pen_str = request.form.get('penalty', '0').strip()
    manual_penalty = float(pen_str) if pen_str else 0.0
    adj_str = request.form.get('adjustment_hours', '0').strip()
    adjustment_hours = float(adj_str) if adj_str else 0.0

    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()

    # Snapshot the values before the edit -- for the audit log and raw-punch backfill.
    c.execute("""SELECT l.date, l.time_in, l.time_out, l.raw_time_in, l.raw_time_out, e.name
                 FROM logs l JOIN employees e ON l.emp_id = e.id WHERE l.id=?""", (log_id,))
    prev = c.fetchone()
    old_date = prev[0] if prev else ""
    old_time_in = prev[1] if prev else None
    old_time_out = prev[2] if prev else None
    old_raw_in = prev[3] if prev else None
    old_raw_out = prev[4] if prev else None
    staff_name = prev[5] if prev else "Unknown Staff"

    try:
        t_in = _parse_hms(time_in)
        t_out = _parse_hms(time_out)
        current_t = t_in.time()
        if LATE_THRESHOLD <= current_t <= LATE_CUTOFF:
            final_penalty = manual_penalty
        else:
            final_penalty = 0.0
        # Excess time-out cap: payroll computation only. Stored time_out stays as entered.
        t_out_capped = apply_out_time_cap(t_out)
        raw_hours = (t_out_capped - t_in).total_seconds() / 3600
        break_multiplier = int(raw_hours // 6)
        final_hours = max(0, raw_hours - (break_multiplier * 1.25))
    except Exception:
        final_hours = 0
        final_penalty = manual_penalty

    c.execute("""UPDATE logs SET time_in=?, time_out=?, penalty=?, hours=?, adjustment_hours=?,
                 raw_time_in=COALESCE(raw_time_in, ?), raw_time_out=COALESCE(raw_time_out, ?)
                 WHERE id=?""",
              (time_in, time_out, final_penalty, round(final_hours, 2), adjustment_hours,
               old_raw_in if old_raw_in else old_time_in,
               old_raw_out if old_raw_out else old_time_out,
               log_id))
    conn.commit(); conn.close()

    # Audit: one row per changed field. Raw punch columns above are never overwritten.
    if (old_time_in or "") != (time_in or ""):
        log_payroll_edit(old_date, staff_name, "Time_In", old_time_in, time_in, "admin")
    if (old_time_out or "") != (time_out or ""):
        log_payroll_edit(old_date, staff_name, "Time_Out", old_time_out, time_out, "admin")

    return redirect(url_for('view_logs', emp_id=emp_id, start_date=start_date, end_date=end_date, late_only=late_only if late_only == '1' else None))

@app.route('/add_adjustment', methods=['POST'])
def add_adjustment():
    if 'logged_in' not in session: return redirect(url_for('login'))
    emp_id = request.form['emp_id']
    adj_date = request.form['adj_date']
    adj_hours = float(request.form['adj_hours'])
    reason = request.form['reason']
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("INSERT INTO logs (emp_id, date, time_in, time_out, hours, penalty, adjustment_hours, adjustment_reason) VALUES (?, ?, 'ADJUSTMENT', 'ADJUSTMENT', 0, 0, ?, ?)", (emp_id, adj_date, adj_hours, reason))
    conn.commit(); conn.close()
    return redirect(url_for('admin'))

@app.route('/insert_missing_log', methods=['POST'])
def insert_missing_log():
    if 'logged_in' not in session: return redirect(url_for('login'))
    emp_id = request.form['emp_id']
    log_date = request.form['log_date']
    time_in_str = f"{request.form['time_in']}:00"
    time_out_str = f"{request.form['time_out']}:00"
    try:
        t_in = datetime.strptime(time_in_str, "%H:%M:%S")
        t_out = datetime.strptime(time_out_str, "%H:%M:%S")
        current_t = t_in.time()
        is_late = LATE_THRESHOLD <= current_t <= LATE_CUTOFF
        penalty = PENALTY if is_late else 0.0
        remarks = "LATE" if is_late else ""
        # Excess time-out cap: payroll computation only. Stored time_out stays as entered.
        t_out_capped = apply_out_time_cap(t_out)
        raw_hours = (t_out_capped - t_in).total_seconds() / 3600
        break_multiplier = int(raw_hours // 6)
        final_hours = max(0, raw_hours - (break_multiplier * 1.25))
        conn = sqlite3.connect('pharmacy_data.db')
        c = conn.cursor()
        c.execute("SELECT name FROM employees WHERE id=?", (emp_id,))
        emp = c.fetchone()
        emp_name = emp[0] if emp else "Unknown Staff"
        c.execute("INSERT INTO logs (emp_id, date, time_in, time_out, raw_time_in, raw_time_out, hours, penalty) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                  (emp_id, log_date, time_in_str, time_out_str, time_in_str, time_out_str, round(final_hours, 2), penalty))
        conn.commit(); conn.close()
        sync_to_google_sheet(emp_name, log_date, time_in_str, "Time In", remarks)
        sync_to_google_sheet(emp_name, log_date, time_out_str, "Time Out", "")
    except Exception as e:
        print(f"Error saving missing log: {e}")
    return redirect(url_for('admin'))

@app.route('/delete_log/<int:log_id>', methods=['POST'])
def delete_log(log_id):
    if 'logged_in' not in session: return redirect(url_for('login'))
    emp_id = request.form['emp_id']
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    late_only = request.form.get('late_only', '') 
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("DELETE FROM logs WHERE id=?", (log_id,))
    conn.commit(); conn.close()
    return redirect(url_for('view_logs', emp_id=emp_id, start_date=start_date, end_date=end_date, late_only=late_only if late_only == '1' else None))

@app.route('/export_csv', methods=['POST'])
def export_csv():
    if 'logged_in' not in session: return redirect(url_for('login'))
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    late_only = 'late_only' in request.form
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("SELECT e.name, SUM(l.hours), SUM(l.penalty), SUM(l.adjustment_hours) FROM logs l JOIN employees e ON l.emp_id = e.id WHERE l.date BETWEEN ? AND ? GROUP BY l.emp_id", (start_date, end_date))
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Employee Name', 'System Hours', 'Adjustment Hours', 'Total Payable Hours', 'Rate per Hour', 'Gross Pay', 'Total Late Penalty'])
    for row in c.fetchall():
        hrs = row[1] if row[1] is not None else 0
        pen = row[2] if row[2] is not None else 0
        adj = row[3] if row[3] is not None else 0
        if late_only and pen <= 0: continue
        total_hours = hrs + adj
        writer.writerow([row[0], round(hrs, 2), round(adj, 2), round(total_hours, 2), HOURLY_RATE, round(total_hours * HOURLY_RATE, 2), pen])
    conn.close()
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment;filename=JDAN_Payroll_{start_date}to{end_date}.csv"})

@app.route('/payslip/<int:emp_id>', methods=['POST'])
def payslip(emp_id):
    if 'logged_in' not in session: return redirect(url_for('login'))
    start_date = request.form['start_date']
    end_date = request.form['end_date']
    pay_date = datetime.now().strftime("%B %d, %Y")
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("SELECT name, role FROM employees WHERE id=?", (emp_id,))
    emp_data = c.fetchone()
    c.execute("SELECT SUM(hours), SUM(penalty), SUM(adjustment_hours) FROM logs WHERE emp_id=? AND date BETWEEN ? AND ?", (emp_id, start_date, end_date))
    log_data = c.fetchone()
    hrs = log_data[0] if log_data[0] is not None else 0
    pen = log_data[1] if log_data[1] is not None else 0
    adj = log_data[2] if log_data[2] is not None else 0
    conn.close()
    total_hours = hrs + adj
    gross_pay = total_hours * HOURLY_RATE
    payroll_period = f"{datetime.strptime(start_date, '%Y-%m-%d').strftime('%B %d')} TO {datetime.strptime(end_date, '%Y-%m-%d').strftime('%B %d, %Y')}".upper()
    return render_template('payslip.html', emp_id=f"JDAN-{emp_id:04d}", name=emp_data[0].upper(), role=(emp_data[1] if len(emp_data) > 1 else 'Pharmacy Assistant').upper(), hours=round(hrs, 2), adjustment=round(adj, 2), total_hours=round(total_hours, 2), gross=round(gross_pay, 2), penalty=pen, period=payroll_period, pay_date=pay_date, rate=HOURLY_RATE)

# ==========================================
# --- STAFF PORTAL ROUTES ---
# ==========================================
@app.route('/portal_login', methods=['GET', 'POST'])
def portal_login():
    error = ""
    if request.method == 'POST':
        pin = request.form['pin']
        conn = sqlite3.connect('pharmacy_data.db')
        c = conn.cursor()
        c.execute("SELECT id, name, pin_hash, is_temp, pin_last_changed FROM employees")
        employees = c.fetchall()
        conn.close()
        
        found = False
        for emp in employees:
            if check_password_hash(emp[2], pin):
                now = datetime.now()
                is_expired = False
                days_left = PIN_EXPIRY_DAYS
                
                if emp[4]:
                    try:
                        last_changed = datetime.strptime(emp[4], "%Y-%m-%d")
                        days_since = (now - last_changed).days
                        days_left = PIN_EXPIRY_DAYS - days_since
                        if days_left <= 0:
                            is_expired = True
                    except: pass
                else:
                    is_expired = True 
                
                if emp[3] == 1 or is_expired:
                    return redirect(url_for('change_pin', emp_id=emp[0]))
                
                session['staff_id'] = emp[0]
                session['staff_name'] = emp[1]
                session['days_left'] = days_left 
                found = True
                break
                
        if found: return redirect(url_for('my_dashboard'))
        else:
            error = "Invalid PIN. Please try again."
            return render_template('staff_login.html', error=error)
    return render_template('staff_login.html', error=error)

@app.route('/my_dashboard', methods=['GET', 'POST'])
def my_dashboard():
    if 'staff_id' not in session: return redirect(url_for('portal_login'))
    emp_id = session['staff_id']
    emp_name = session['staff_name']
    days_left = session.get('days_left', 60)
    
    now = datetime.now()
    selected_month = now.strftime('%Y-%m')
    if request.method == 'POST':
        selected_month = request.form['month_filter']
        
    year, month = map(int, selected_month.split('-'))
    start_date = f"{year}-{month:02d}-01"
    last_day = calendar.monthrange(year, month)[1]
    end_date = f"{year}-{month:02d}-{last_day:02d}"
    
    month_holidays = []
    for date_key, holiday_name in PH_HOLIDAYS.items():
        if date_key.startswith(f"{month:02d}-"):
            month_holidays.append(f"{holiday_name} ({date_key})")
            
    conn = sqlite3.connect('pharmacy_data.db')
    c = conn.cursor()
    c.execute("SELECT date, time_in, time_out, hours, penalty, adjustment_hours, adjustment_reason FROM logs WHERE emp_id=? AND date BETWEEN ? AND ? ORDER BY date DESC", (emp_id, start_date, end_date))
    logs = c.fetchall()
    conn.close()
    
    total_hours = sum([l[3] for l in logs if l[3] is not None])
    total_adj = sum([l[5] for l in logs if l[5] is not None])
    total_payable = total_hours + total_adj
    total_penalty = sum([l[4] for l in logs if l[4] is not None])
    
    return render_template('staff_dashboard.html', logs=logs, selected_month=selected_month, start_date=start_date, end_date=end_date, emp_name=emp_name, emp_id=emp_id, total_payable=round(total_payable, 2), total_penalty=total_penalty, month_holidays=month_holidays, days_left=days_left)

@app.route('/portal_logout')
def portal_logout():
    session.pop('staff_id', None)
    session.pop('staff_name', None)
    session.pop('days_left', None)
    return redirect(url_for('home'))

@app.route('/change_pin/<int:emp_id>', methods=['GET', 'POST'])
def change_pin(emp_id):
    if request.method == 'POST':
        today = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect('pharmacy_data.db')
        c = conn.cursor()
        c.execute("UPDATE employees SET pin_hash=?, is_temp=0, pin_last_changed=? WHERE id=?", (generate_password_hash(request.form['new_pin']), today, emp_id))
        conn.commit(); conn.close()
        return redirect(url_for('home'))
    return render_template('change_pin.html')

if __name__ == '__main__':
    init_db()
    app.run(debug=True, host='0.0.0.0')