"""Shared Item Code resolver (barcode crosswalk).

The Master, logs, reports and the Google Sheets queue always use the OLD
(internal) Item_Code. BITPOS has ONE field for Item Code and barcode, so once an
item's code is replaced by its manufacturer barcode, BITPOS exports show the
barcode instead. Every script that reads an Item Code from BITPOS data calls this
module first to translate it back to the internal Item_Code.

Rules (never relaxed):
  * Codes are always TEXT. Nothing here converts a code to a number.
  * A code that looks like 4.80015E+12, ends in ".0", or has embedded spaces is
    CORRUPTED. It is reported, never repaired or guessed.
  * Nothing is silently dropped: every code gets a status, and anything other
    than AS_IS / TRANSLATED is counted and reported by the caller.

Statuses:
  AS_IS       code already exists in Inventory_Master (old code, or a future item
              whose barcode IS its Item_Code) -> unchanged
  TRANSLATED  code is a barcode listed in Barcode_Crosswalk.csv -> old Item_Code
  UNRESOLVED  not in Master and not in the crosswalk
  CORRUPTED   damaged text (scientific notation, ".0", embedded spaces, lost
              leading zeros that would match a known code)
  EMPTY       blank / "nan" / MANUAL-ENTRY (no code recorded)
  CONFLICT    ambiguous: barcode listed for two items, or equal to a DIFFERENT
              item's Item_Code in Master
"""
import os
import re
import csv
import sys
import shutil
import datetime

_BASE = os.path.dirname(os.path.abspath(__file__))

CROSSWALK_FILE = os.path.join(_BASE, 'Barcode_Crosswalk.csv')
CROSSWALK_COLUMNS = ['Old_Item_Code', 'Barcode', 'Product_Name', 'Status', 'Date_Added', 'Notes']
CROSSWALK_STATUSES = ('PLANNED', 'ACTIVE', 'REVERTED')
MASTER_FILE = os.path.join(_BASE, 'Inventory_Master.xlsx')
LOG_FILE = os.path.join(_BASE, 'JDan_System_Logs.txt')

AS_IS = 'AS_IS'
TRANSLATED = 'TRANSLATED'
UNRESOLVED = 'UNRESOLVED'
CORRUPTED = 'CORRUPTED'
EMPTY = 'EMPTY'
CONFLICT = 'CONFLICT'
RESOLVED_STATUSES = (AS_IS, TRANSLATED)
PROBLEM_STATUSES = (UNRESOLVED, CORRUPTED, EMPTY, CONFLICT)

EMPTY_TOKENS = {'', 'NAN', 'NONE', 'NULL', 'NAT', 'MANUAL-ENTRY'}
_SCI = re.compile(r'^[+-]?\d+(\.\d+)?[eE][+-]?\d+$')
_DECIMAL = re.compile(r'^[+-]?\d+\.\d+$')
_STATUS_ORDER = (AS_IS, TRANSLATED, UNRESOLVED, CORRUPTED, EMPTY, CONFLICT)


class Resolution:
    """Result for one code. `code` is the internal Item_Code when status is
    AS_IS or TRANSLATED, otherwise None. `raw` is exactly what was read."""
    __slots__ = ('raw', 'code', 'status', 'reason', 'barcode', 'padded')

    def __init__(self, raw, code, status, reason='', barcode=None, padded=False):
        self.raw = raw
        self.code = code
        self.status = status
        self.reason = reason
        self.barcode = barcode
        self.padded = padded

    @property
    def ok(self):
        return self.status in RESOLVED_STATUSES

    def __repr__(self):
        return f'Resolution({self.raw!r} -> {self.code!r}, {self.status}, {self.reason!r})'


# ==========================================
#    TEXT / BARCODE VALIDATION HELPERS
# ==========================================
def gtin_check_ok(code):
    """GS1 mod-10 check digit for 8/12/13/14 digit codes. None = length not checkable."""
    if not isinstance(code, str) or not code.isdigit() or len(code) not in (8, 12, 13, 14):
        return None
    total = 0
    for i, ch in enumerate(reversed(code[:-1])):
        total += int(ch) * (3 if i % 2 == 0 else 1)
    return (10 - total % 10) % 10 == int(code[-1])


def barcode_text_problems(value):
    """Problems that make a barcode unsafe to store (list of plain-English reasons)."""
    if not isinstance(value, str):
        return ['not text']
    problems = []
    if value != value.strip():
        problems.append('leading/trailing whitespace')
    v = value.strip()
    if not v:
        return problems + ['empty']
    if _SCI.match(v):
        problems.append('scientific notation')
    elif _DECIMAL.match(v):
        problems.append('decimal point (number conversion, e.g. ".0")')
    elif not v.isdigit():
        problems.append('contains non-digit characters')
    elif len(v) in (7, 11):
        problems.append(f'length {len(v)}: a leading zero may have been lost')
    return problems


def _clean_str(value):
    """Text exactly as read, or '' for None/NaN. Never converts through float."""
    if value is None:
        return ''
    if isinstance(value, float):
        # a float here means the value already went through a number conversion
        return 'nan' if value != value else repr(value)
    return str(value)


# ==========================================
#    CROSSWALK LOADING
# ==========================================
class Crosswalk:
    """In-memory crosswalk. Rows that fail validation are listed in `problems`
    and are NOT used for translation (never guessed)."""

    def __init__(self, rows=None, source='(in memory)'):
        self.source = source
        self.rows = []
        self.by_barcode = {}
        self.by_old = {}
        self.ambiguous_barcodes = set()
        self.problems = []
        for r in (rows or []):
            self._add(r)

    def _add(self, r):
        old = _clean_str(r.get('Old_Item_Code', ''))
        bc = _clean_str(r.get('Barcode', ''))
        row = {c: _clean_str(r.get(c, '')) for c in CROSSWALK_COLUMNS}
        self.rows.append(row)
        label = f'{old or "(blank old code)"} -> {bc or "(blank barcode)"}'
        bad = barcode_text_problems(bc)
        if old != old.strip() or not old.strip():
            bad.append('old Item_Code blank or has whitespace')
        if row['Status'].strip().upper() not in CROSSWALK_STATUSES:
            bad.append(f'unknown Status {row["Status"]!r}')
        if bad:
            self.problems.append(f'crosswalk row {label}: ' + '; '.join(bad) + ' - row IGNORED')
            return
        if bc == old:
            self.problems.append(f'crosswalk row {label}: barcode equals its own old code - row IGNORED')
            return
        if bc in self.by_barcode and self.by_barcode[bc]['Old_Item_Code'] != old:
            self.ambiguous_barcodes.add(bc)
            self.problems.append(f'barcode {bc} is listed for TWO different items '
                                 f'({self.by_barcode[bc]["Old_Item_Code"]} and {old}) - treated as CONFLICT')
            return
        if bc in self.ambiguous_barcodes:
            self.problems.append(f'barcode {bc} is listed for several items - treated as CONFLICT')
            return
        if old in self.by_old and self.by_old[old]['Barcode'] != bc:
            self.problems.append(f'old code {old} has more than one barcode ({self.by_old[old]["Barcode"]}, {bc}) - both kept')
        row['Status'] = row['Status'].strip().upper()
        self.by_barcode[bc] = row
        self.by_old.setdefault(old, row)

    def __len__(self):
        return len(self.by_barcode)


def load_crosswalk(path=None):
    """Read Barcode_Crosswalk.csv as TEXT (csv module, no number parsing).
    A missing file gives an empty crosswalk (today's behaviour: nothing to translate)."""
    path = path or CROSSWALK_FILE
    if not os.path.exists(path):
        cw = Crosswalk(source=path)
        cw.problems.append(f'crosswalk file not found: {path} - no barcode translation is possible')
        return cw
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        missing = [c for c in CROSSWALK_COLUMNS if c not in (reader.fieldnames or [])]
        rows = list(reader)
    cw = Crosswalk(rows, source=path)
    if missing:
        cw.problems.insert(0, f'crosswalk file is missing column(s): {missing}')
    return cw


def load_master_table(path=None):
    """Master as TEXT (Item_Code, Product_Name, ...). Returns a pandas DataFrame."""
    import pandas as pd
    return pd.read_excel(path or MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str, keep_default_na=False)


# ==========================================
#    THE RESOLVER
# ==========================================
class Resolver:
    def __init__(self, crosswalk=None, master_codes=None):
        self.crosswalk = crosswalk if crosswalk is not None else Crosswalk()
        self.master_codes = set(master_codes) if master_codes is not None else None

    @classmethod
    def from_files(cls, crosswalk_path=None, master_path=None):
        cw = load_crosswalk(crosswalk_path)
        codes = set(str(c).strip() for c in load_master_table(master_path)['Item_Code'])
        return cls(cw, codes)

    def resolve(self, raw):
        text = _clean_str(raw)
        t = text.strip()
        padded = (t != text)
        if t.upper() in EMPTY_TOKENS:
            return Resolution(raw, None, EMPTY, 'no Item Code recorded (blank or manual entry)', padded=padded)
        if _SCI.match(t):
            return Resolution(raw, None, CORRUPTED, 'scientific notation (file was opened/saved as a number)', padded=padded)
        if _DECIMAL.match(t):
            return Resolution(raw, None, CORRUPTED, 'decimal point such as ".0" (number conversion)', padded=padded)
        if re.search(r'\s', t):
            return Resolution(raw, None, CORRUPTED, 'contains spaces inside the code', padded=padded)

        cw = self.crosswalk
        in_master = self.master_codes is not None and t in self.master_codes
        if t in cw.ambiguous_barcodes:
            return Resolution(raw, None, CONFLICT, 'barcode is listed for more than one item in the crosswalk', t, padded)
        row = cw.by_barcode.get(t)
        if row is not None:
            old = row['Old_Item_Code']
            if in_master:
                return Resolution(raw, None, CONFLICT,
                                  f'is a crosswalk barcode (for {old}) AND an Item_Code already in Master', t, padded)
            if self.master_codes is not None and old not in self.master_codes:
                return Resolution(raw, None, UNRESOLVED,
                                  f'crosswalk maps it to {old}, which is not in Master', t, padded)
            return Resolution(raw, old, TRANSLATED, f'barcode of {old} (crosswalk status {row["Status"]})', t, padded)
        if in_master:
            return Resolution(raw, t, AS_IS, 'Item_Code exists in Master', None, padded)
        if self.master_codes is None:
            return Resolution(raw, None, UNRESOLVED, 'Master not loaded - cannot confirm', None, padded)
        if t.isdigit():
            for z in ('0', '00', '000'):
                if (z + t) in cw.by_barcode or (z + t) in self.master_codes:
                    return Resolution(raw, None, CORRUPTED,
                                      f'leading zero(s) probably lost: {z + t} exists but {t} does not', None, padded)
        return Resolution(raw, None, UNRESOLVED, 'not in Master and not in the crosswalk', None, padded)

    def resolve_to_internal(self, raw):
        return self.resolve(raw)

    def resolve_to_barcode(self, internal_code):
        """The crosswalk barcode for an internal Item_Code, or None. Items whose
        Item_Code is itself the barcode (future items) also return None."""
        row = self.crosswalk.by_old.get(str(internal_code).strip())
        return row['Barcode'] if row else None

    def internal_or_raw(self, raw):
        """Internal code if resolved, else the stripped raw text (so nothing is lost)."""
        res = self.resolve(raw)
        return res.code if res.ok else _clean_str(raw).strip()

    def find_collisions(self):
        """Crosswalk barcodes that equal an Item_Code in Master or another item's old code."""
        found = []
        cw = self.crosswalk
        old_codes = set(cw.by_old)
        for bc, row in cw.by_barcode.items():
            if self.master_codes is not None and bc in self.master_codes:
                found.append(f'barcode {bc} (for {row["Old_Item_Code"]}) equals an Item_Code already in Master')
            if bc in old_codes and bc != row['Old_Item_Code']:
                found.append(f'barcode {bc} (for {row["Old_Item_Code"]}) equals the old code of another crosswalk item')
        for bc in sorted(cw.ambiguous_barcodes):
            found.append(f'barcode {bc} is listed for more than one item')
        return found

    def is_safe_new_internal_code(self, candidate):
        """True if a newly generated internal code collides with nothing known."""
        c = str(candidate).strip()
        if self.master_codes is not None and c in self.master_codes:
            return False
        return c not in self.crosswalk.by_barcode and c not in self.crosswalk.by_old and c not in self.crosswalk.ambiguous_barcodes


_DEFAULT = None


def get_resolver(refresh=False):
    """Process-wide resolver built from the live Master + live crosswalk (read-only)."""
    global _DEFAULT
    if _DEFAULT is None or refresh:
        _DEFAULT = Resolver.from_files()
    return _DEFAULT


def load_crosswalk_default():
    return get_resolver().crosswalk


def resolve_to_internal(code, resolver=None):
    return (resolver or get_resolver()).resolve(code)


def resolve_to_barcode(internal_code, resolver=None):
    return (resolver or get_resolver()).resolve_to_barcode(internal_code)


def sheet_text_code(code):
    """Code as it must be queued for Google Sheets (USER_ENTERED): a leading
    apostrophe keeps leading zeros / long barcodes as text. Plain 8-digit codes
    without a leading zero are returned unchanged, so today's rows do not change."""
    c = str(code)
    if c.isdigit() and (c.startswith('0') or len(c) >= 12):
        return "'" + c
    return c


# ==========================================
#    REPORTING
# ==========================================
def log_event(message):
    """Append one line to JDan_System_Logs.txt (same timestamp style as the apps)."""
    stamp = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f'[{stamp}] - {message}\n')
    except OSError:
        pass


class ResolutionReport:
    """Counts + details for a batch of codes (rows are never dropped)."""

    def __init__(self):
        self.counts = {s: 0 for s in _STATUS_ORDER}
        self.details = {}          # (raw_stripped, status, reason) -> [rows, units]
        self.total = 0
        self.padded_rows = 0
        self.planned_translated = 0

    def add(self, res, units=0.0):
        self.total += 1
        self.counts[res.status] += 1
        if res.padded:
            self.padded_rows += 1
        if res.status == TRANSLATED and 'status PLANNED' in res.reason:
            self.planned_translated += 1
        if res.status in PROBLEM_STATUSES:
            key = (_clean_str(res.raw).strip(), res.status, res.reason)
            d = self.details.setdefault(key, [0, 0.0])
            d[0] += 1
            d[1] += float(units or 0)

    @property
    def problem_count(self):
        return sum(self.counts[s] for s in PROBLEM_STATUSES)

    @property
    def has_problems(self):
        return self.problem_count > 0

    def lines(self, max_detail=15):
        c = self.counts
        out = [f"{self.total} code(s) checked: {c[AS_IS]} resolved as-is, {c[TRANSLATED]} translated from barcode, "
               f"{c[UNRESOLVED]} unresolved, {c[CORRUPTED]} corrupted, {c[EMPTY]} empty, {c[CONFLICT]} conflicts"]
        if self.planned_translated:
            out.append(f"  note: {self.planned_translated} line(s) used a barcode whose crosswalk status is still PLANNED "
                       f"(mark it ACTIVE once the change in BITPOS is confirmed)")
        shown = 0
        for (raw, status, reason), (rows, units) in sorted(self.details.items(), key=lambda kv: (kv[0][1], kv[0][0])):
            if shown >= max_detail:
                out.append(f'  ... and {len(self.details) - max_detail} more distinct problem code(s)')
                break
            out.append(f"  {status}: {raw!r} - {reason} ({rows} line(s)" + (f", {units:g} unit(s))" if units else ')'))
            shown += 1
        return out

    def text(self, max_detail=15):
        return '\n'.join(self.lines(max_detail))


def resolve_series(series, resolver=None, weights=None):
    """Resolve a whole column of BITPOS item codes. Returns (internal_codes, report).
    internal_codes has the same index/length; unresolved codes keep their stripped
    raw text so downstream totals are unchanged (they are reported, not dropped)."""
    import pandas as pd
    r = resolver or get_resolver()
    report = ResolutionReport()
    cache = {}
    out = []
    for i, raw in enumerate(series.tolist()):
        key = raw if isinstance(raw, str) else repr(raw)
        if key not in cache:
            cache[key] = r.resolve(raw)
        res = cache[key]
        w = 0.0
        if weights is not None:
            try:
                w = float(weights.iloc[i])
            except (TypeError, ValueError):
                w = 0.0
        report.add(res, w)
        out.append(res.code if res.ok else _clean_str(raw).strip())
    return pd.Series(out, index=series.index, dtype=object), report


def announce(report, context, show=print, log=True):
    """Print (and log) a report only when something needs attention. Returns the
    text, or '' when everything resolved cleanly."""
    if not report.has_problems:
        return ''
    text = f'ITEM CODE CHECK [{context}]\n' + report.text()
    if show:
        show(text)
    if log:
        log_event(f'ITEM CODE CHECK [{context}] ' + ' | '.join(report.lines(8)))
    return text


def resolve_sales_history(all_sales, context, resolver=None, code_col='Item Code', qty_col='Qty',
                          date_col='Date', recent_days=60, show=print):
    """For the velocity/ranking tools: translate the Item Code column of the
    PROCESSED_EOD history to internal codes. Returns (internal_codes, popup_text).
    Problems are always counted and logged; `popup_text` is non-empty only when a
    problem line is from the last `recent_days` days (old June lines for items that
    have since left the Master are logged, not shown every time)."""
    import pandas as pd
    codes, report = resolve_series(all_sales[code_col].astype(str).str.strip(), resolver, weights=all_sales[qty_col])
    if not report.has_problems:
        return codes, ''
    recent_text = ''
    recent_report = None
    if date_col in all_sales.columns:
        dates = pd.to_datetime(all_sales[date_col], errors='coerce')
        cutoff = pd.Timestamp.now().normalize() - pd.Timedelta(days=recent_days)
        recent = (dates >= cutoff) | dates.isna()
        _, recent_report = resolve_series(all_sales.loc[recent, code_col].astype(str).str.strip(), resolver, weights=all_sales.loc[recent, qty_col])
    else:
        recent_report = report
    text = announce(report, context + ' (all sales history)', show=show, log=True)
    if recent_report.has_problems:
        recent_text = f'ITEM CODE CHECK [{context}] - sales from the last {recent_days} days\n' + recent_report.text()
    return codes, recent_text


def ask_yes_dialog(parent, title, text, prompt='Type YES to continue, or press Cancel to stop. Nothing is changed until you type YES.'):
    """Modal window: scrollable text + an entry box. Returns True only if the
    operator typed YES (any case). Used by the EOD and SOD apps."""
    import tkinter as tk
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry('780x520')
    win.grab_set()
    win.transient(parent)
    result = {'ok': False}
    tk.Label(win, text=title, font=('Arial', 12, 'bold'), fg='#cc0000').pack(pady=(10, 4))
    frame = tk.Frame(win)
    frame.pack(fill='both', expand=True, padx=12)
    box = tk.Text(frame, wrap='word', font=('Consolas', 9), height=20)
    scroll = tk.Scrollbar(frame, command=box.yview)
    box.configure(yscrollcommand=scroll.set)
    scroll.pack(side='right', fill='y')
    box.pack(side='left', fill='both', expand=True)
    box.insert('1.0', text)
    box.configure(state='disabled')
    tk.Label(win, text=prompt, font=('Arial', 10, 'bold')).pack(pady=(8, 2))
    entry = tk.Entry(win, font=('Arial', 12), justify='center', width=12)
    entry.pack()
    entry.focus_set()

    def go(event=None):
        if entry.get().strip().upper() == 'YES':
            result['ok'] = True
            win.destroy()
        else:
            entry.configure(bg='#ffcccc')

    btns = tk.Frame(win)
    btns.pack(pady=10)
    tk.Button(btns, text='CONTINUE', width=14, bg='#003366', fg='white', command=go).pack(side='left', padx=8)
    tk.Button(btns, text='CANCEL', width=14, command=win.destroy).pack(side='left', padx=8)
    win.bind('<Return>', go)
    parent.wait_window(win)
    return result['ok']


# ==========================================
#    SALES LINE RECONCILIATION (display only)
# ==========================================
BOTH, NAME_ONLY, HELD_CONFLICT, UNMATCHED = 'BOTH', 'NAME_ONLY', 'CONFLICT', 'UNMATCHED'


def _name_key(value):
    return str(value).strip().upper()


class Reconciliation:
    """Per-line comparison of 'item found by code' vs 'item found by name'.
    Names stay the deduction key; the code is only a cross-check."""

    def __init__(self, lines, units_file):
        self.lines = lines                                   # list of dicts
        self.units_file = units_file
        self.count = {k: 0 for k in (BOTH, NAME_ONLY, HELD_CONFLICT, UNMATCHED)}
        self.units = {k: 0.0 for k in (BOTH, NAME_ONLY, HELD_CONFLICT, UNMATCHED)}
        self.corrupted_lines = 0
        self.code_problem_name_only = 0
        for ln in lines:
            self.count[ln['category']] += 1
            self.units[ln['category']] += ln['qty']
            if ln['code_status'] in (CORRUPTED, CONFLICT):
                self.corrupted_lines += 1
            if ln['category'] == NAME_ONLY and ln['code_status'] in PROBLEM_STATUSES:
                self.code_problem_name_only += 1

    @property
    def units_to_deduct(self):
        """Units sent to the name-based deduction (BOTH + NAME_ONLY)."""
        return self.units[BOTH] + self.units[NAME_ONLY]

    @property
    def units_not_deducted(self):
        return self.units[UNMATCHED] + self.units[HELD_CONFLICT]

    @property
    def reconciles(self):
        return abs(self.units_file - (self.units_to_deduct + self.units_not_deducted)) < 1e-9

    @property
    def held_indices(self):
        return [ln['index'] for ln in self.lines if ln['category'] == HELD_CONFLICT]

    @property
    def needs_confirmation(self):
        """Operator must type YES: any unmatched line, any conflict, any damaged or ambiguous code."""
        return bool(self.count[UNMATCHED] or self.count[HELD_CONFLICT] or self.corrupted_lines)

    def of(self, category):
        return [ln for ln in self.lines if ln['category'] == category]

    def summary_lines(self, max_detail=25):
        n = len(self.lines)
        out = [
            f"Sales lines checked: {n}   Units in the BITPOS file: {self.units_file:g}",
            f"  Matched by code AND name: {self.count[BOTH]} line(s), {self.units[BOTH]:g} unit(s)",
            f"  Matched by name only:     {self.count[NAME_ONLY]} line(s), {self.units[NAME_ONLY]:g} unit(s)"
            + (f"  ({self.code_problem_name_only} of them have an unreadable or unknown code)" if self.code_problem_name_only else ''),
            f"  Code and name disagree (NOT deducted): {self.count[HELD_CONFLICT]} line(s), {self.units[HELD_CONFLICT]:g} unit(s)",
            f"  No matching item (NOT deducted):       {self.count[UNMATCHED]} line(s), {self.units[UNMATCHED]:g} unit(s)",
        ]
        if self.corrupted_lines:
            out.append(f"  {self.corrupted_lines} line(s) have a DAMAGED or AMBIGUOUS Item Code (listed below)")
        out.append(f"Units to be deducted from stock: {self.units_to_deduct:g}   Units not deducted: {self.units_not_deducted:g}   "
                   f"Check: {self.units_to_deduct:g} + {self.units_not_deducted:g} = {self.units_to_deduct + self.units_not_deducted:g} "
                   f"({'RECONCILES' if self.reconciles else 'DOES NOT RECONCILE'})")
        shown = 0
        for cat, title in ((HELD_CONFLICT, 'CODE/NAME CONFLICT (held back, not deducted)'),
                           (UNMATCHED, 'NO MATCHING ITEM (not deducted)'),
                           (NAME_ONLY, 'MATCHED BY NAME ONLY - code not usable')):
            rows = [ln for ln in self.of(cat) if cat != NAME_ONLY or ln['code_status'] in PROBLEM_STATUSES]
            if not rows:
                continue
            out.append(f'{title}:')
            for ln in rows:
                if shown >= max_detail:
                    out.append('  ... more lines are listed in the system log')
                    return out
                out.append(f"  {ln['name']}  |  qty {ln['qty']:g}  |  code {ln['code_raw'].strip()!r}: {ln['note']}")
                shown += 1
        return out

    def log(self, context):
        log_event(f'EOD RECONCILIATION [{context}] lines={len(self.lines)} units_file={self.units_file:g} '
                  f'both={self.count[BOTH]} name_only={self.count[NAME_ONLY]} conflicts={self.count[HELD_CONFLICT]} '
                  f'unmatched={self.count[UNMATCHED]} units_to_deduct={self.units_to_deduct:g} '
                  f'units_not_deducted={self.units_not_deducted:g} reconciles={self.reconciles}')
        for cat in (HELD_CONFLICT, UNMATCHED):
            for ln in self.of(cat):
                log_event(f"   -> {cat}: {ln['name']} | qty {ln['qty']:g} | code {ln['code_raw'].strip()!r} | {ln['note']}")
        for ln in self.of(NAME_ONLY):
            if ln['code_status'] in PROBLEM_STATUSES:
                log_event(f"   -> NAME_ONLY (code {ln['code_status']}): {ln['name']} | qty {ln['qty']:g} | "
                          f"code {ln['code_raw'].strip()!r} | {ln['note']}")


def reconcile_sales(df, master_df, resolver, code_col='Item Code', name_col='Description', qty_col='Qty'):
    """Classify every sales line (Qty > 0 only; the same lines the deduction uses).
    `master_df` needs Item_Code and Product_Name as text. Pure function: no files, no GUI."""
    name_to_codes = {}
    code_to_name = {}
    for code, nm in zip(master_df['Item_Code'].astype(str).str.strip(), master_df['Product_Name']):
        name_to_codes.setdefault(_name_key(nm), set()).add(code)
        code_to_name.setdefault(code, str(nm))
    lines = []
    units_file = 0.0
    has_code_col = code_col in df.columns
    cache = {}
    for idx, row in df.iterrows():
        try:
            qty = float(row[qty_col])
        except (TypeError, ValueError):
            continue
        if not qty > 0:
            continue
        units_file += qty
        raw = row[code_col] if has_code_col else None
        key = raw if isinstance(raw, str) else repr(raw)
        if key not in cache:
            cache[key] = resolver.resolve(raw) if has_code_col else Resolution(raw, None, EMPTY, 'no Item Code column in this file')
        res = cache[key]
        name = _name_key(row[name_col])
        codes_by_name = name_to_codes.get(name)
        if not codes_by_name:
            cat = UNMATCHED
            note = 'no Master item has this name'
            if res.ok:
                note += f"; its code points to {res.code} ({code_to_name.get(res.code, '?')})"
            else:
                note += f'; code is {res.status}: {res.reason}'
        elif res.ok:
            if res.code in codes_by_name:
                cat, note = BOTH, 'code and name agree'
            else:
                cat = HELD_CONFLICT
                note = (f"code points to {res.code} ({code_to_name.get(res.code, '?')}) but the name belongs to "
                        f"{', '.join(sorted(codes_by_name))}")
        else:
            cat, note = NAME_ONLY, f'code is {res.status}: {res.reason}'
        lines.append({'index': idx, 'name': str(row[name_col]).strip(), 'qty': qty,
                      'code_raw': _clean_str(raw), 'code_status': res.status, 'code_internal': res.code,
                      'name_codes': sorted(codes_by_name) if codes_by_name else [], 'category': cat, 'note': note})
    return Reconciliation(lines, units_file)


# ==========================================
#    SUMMARY OF A WHOLE BITPOS FILE
# ==========================================
def read_bitpos_text(path):
    """Read a BITPOS export as TEXT. Returns (header_list, rows(list of dict), numeric_code_cells).
    Handles the item/price CSV and the sales export (.xls/.xlsx/.csv with a junk header block).
    `numeric_code_cells` counts Item Code cells stored as NUMBERS in a spreadsheet file
    (a sign that the file was opened and saved in Excel)."""
    lower = path.lower()
    numeric_cells = 0
    if lower.endswith('.csv'):
        with open(path, 'r', encoding='utf-8-sig', newline='') as f:
            grid = list(csv.reader(f))
    else:
        import pandas as pd
        raw = pd.read_excel(path, header=None, dtype=str)
        grid = [['' if (v is None or (isinstance(v, float) and v != v)) else str(v) for v in row] for row in raw.values.tolist()]
    header_idx = None
    for i, row in enumerate(grid):
        up = [c.strip().upper() for c in row]
        if any('ITEM CODE' in c for c in up):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"No 'ITEM CODE' column found in {os.path.basename(path)}")
    header = [c.strip() for c in grid[header_idx]]
    rows = []
    for row in grid[header_idx + 1:]:
        if not any(c.strip() for c in row):
            continue
        rows.append({h: (row[j] if j < len(row) else '') for j, h in enumerate(header)})
    if not lower.endswith('.csv'):
        numeric_cells = _count_numeric_code_cells(path, header_idx, header)
    return header, rows, numeric_cells


def _count_numeric_code_cells(path, header_idx, header):
    try:
        col = next(j for j, h in enumerate(header) if 'ITEM CODE' in h.upper())
        if path.lower().endswith('.xls'):
            import xlrd
            sheet = xlrd.open_workbook(path).sheet_by_index(0)
            return sum(1 for r in range(header_idx + 1, sheet.nrows) if col < sheet.ncols and sheet.cell_type(r, col) == xlrd.XL_CELL_NUMBER)
        import openpyxl
        ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
        return sum(1 for r in ws.iter_rows(min_row=header_idx + 2, min_col=col + 1, max_col=col + 1) if r[0].data_type == 'n')
    except Exception:   # noqa: BLE001 - the numeric-cell check is a bonus signal only
        return 0


def _clean_name(value):
    try:
        from name_cleaning import clean_text
        return str(clean_text(value))
    except ImportError:
        return str(value)


class FileSummary:
    def __init__(self, path):
        self.path = path
        self.report = ResolutionReport()
        self.rows = 0
        self.blank_records = 0
        self.inactive_skipped = 0
        self.numeric_code_cells = 0
        self.recon = None
        self.crosswalk_problems = []

    @property
    def stop(self):
        return bool(self.report.problem_count or self.numeric_code_cells or
                    (self.recon is not None and (self.recon.count[HELD_CONFLICT])))

    def lines(self):
        out = [f'FILE: {self.path}', f'Records checked: {self.rows}'
               + (f' (skipped {self.inactive_skipped} inactive BITPOS record(s))' if self.inactive_skipped else '')
               + (f' (skipped {self.blank_records} malformed/blank record(s) with no Item Code)' if self.blank_records else '')]
        out += self.report.lines(30)
        if self.report.padded_rows:
            out.append(f'  ({self.report.padded_rows} code(s) carry a trailing space - normal for the BITPOS sales export, removed before matching)')
        if self.numeric_code_cells:
            out.append(f'  CORRUPTED: {self.numeric_code_cells} Item Code cell(s) are stored as NUMBERS - the file was probably opened and saved in Excel')
        if self.recon is not None:
            out += ['Name cross-check against Master:'] + ['  ' + s for s in self.recon.summary_lines(10)]
        for p in self.crosswalk_problems:
            out.append('CROSSWALK PROBLEM: ' + p)
        out.append('RESULT: ' + ('STOP - do not continue; see the lines above.' if self.stop else 'OK - every code resolved.'))
        return out


def summarize_file(path, resolver=None):
    """Resolver summary of a real BITPOS file (item export, price list or sales export)."""
    import pandas as pd
    r = resolver or get_resolver()
    header, rows, numeric_cells = read_bitpos_text(path)
    fs = FileSummary(path)
    fs.numeric_code_cells = numeric_cells
    fs.crosswalk_problems = list(r.crosswalk.problems)
    code_h = next(h for h in header if 'ITEM CODE' in h.upper())
    name_h = next((h for h in header if h.strip().upper() in ('DESCRIPTION',)), None)
    qty_h = next((h for h in header if h.strip().upper() == 'QTY'), None)
    inactive_h = next((h for h in header if h.strip().upper() == 'INACTIVE'), None)
    kept = []
    for row in rows:
        raw = row.get(code_h, '')
        if 'TOTAL' in raw.upper():
            continue
        # same rules as the live EOD reader: a row with no code AND no description is not an item line
        if name_h and not raw.strip() and (inactive_h or not (row.get(name_h) or '').strip()):
            fs.blank_records += 1      # (item exports: the known malformed multi-line records have no code)
            continue
        # item exports keep retired items as INACTIVE=1; those are not checked against Master
        if inactive_h and (row.get(inactive_h) or '').strip() == '1':
            fs.inactive_skipped += 1
            continue
        if name_h:
            row[name_h] = _clean_name(row.get(name_h, ''))
        kept.append(row)
        fs.rows += 1
        units = 0.0
        if qty_h:
            try:
                units = float(row.get(qty_h) or 0)
            except ValueError:
                units = 0.0
        fs.report.add(r.resolve(raw), units)
    rows = kept
    if qty_h and name_h and r.master_codes is not None:
        df = pd.DataFrame(rows)
        df['Qty'] = pd.to_numeric(df[qty_h], errors='coerce').fillna(0)
        fs.recon = reconcile_sales(df, load_master_table(), r, code_col=code_h, name_col=name_h, qty_col='Qty')
    return fs


# ==========================================
#    CROSSWALK STATUS EDITS (always backed up first)
# ==========================================
def backup_crosswalk(path=None):
    path = path or CROSSWALK_FILE
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    dest = os.path.join(os.path.dirname(path), f'Barcode_Crosswalk_BACKUP_{stamp}.csv')
    shutil.copy2(path, dest)
    return dest


def set_crosswalk_status(old_codes, new_status, path=None):
    """Change Status for the listed old Item_Codes. A dated backup is made first.
    Returns (changed_count, backup_path). Unknown codes raise, nothing is half-written."""
    path = path or CROSSWALK_FILE
    new_status = new_status.strip().upper()
    if new_status not in CROSSWALK_STATUSES:
        raise ValueError(f'Status must be one of {CROSSWALK_STATUSES}')
    wanted = [str(c).strip() for c in old_codes]
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    known = {r['Old_Item_Code'] for r in rows}
    unknown = [c for c in wanted if c not in known]
    if unknown:
        raise ValueError(f'Not in the crosswalk (nothing changed): {unknown}')
    backup = backup_crosswalk(path)
    changed = 0
    for r in rows:
        if r['Old_Item_Code'] in wanted and r['Status'] != new_status:
            r['Status'] = new_status
            changed += 1
    with open(path, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=CROSSWALK_COLUMNS, lineterminator='\n')
        w.writeheader()
        w.writerows(rows)
    log_event(f'CROSSWALK STATUS set to {new_status} for {changed} item(s) (backup {os.path.basename(backup)})')
    return changed, backup


# ==========================================
#    SELF TEST (synthetic data only - never touches live files)
# ==========================================
def run_self_test(verbose=True):
    cw = Crosswalk([
        {'Old_Item_Code': '11111111', 'Barcode': '4800153151255', 'Product_Name': 'A', 'Status': 'PLANNED'},
        {'Old_Item_Code': '22222222', 'Barcode': '020800306348', 'Product_Name': 'B', 'Status': 'ACTIVE'},
        {'Old_Item_Code': '33333333', 'Barcode': '48033077', 'Product_Name': 'C', 'Status': 'ACTIVE'},
        {'Old_Item_Code': '44444444', 'Barcode': '4800000000017', 'Product_Name': 'D', 'Status': 'ACTIVE'},
        {'Old_Item_Code': '55555555', 'Barcode': '4800000000017', 'Product_Name': 'E', 'Status': 'ACTIVE'},   # duplicate barcode
        {'Old_Item_Code': '66666666', 'Barcode': '4806666666666', 'Product_Name': 'F', 'Status': 'ACTIVE'},   # collides with Master code below
        {'Old_Item_Code': '77777777', 'Barcode': '4.80015E+12', 'Product_Name': 'G', 'Status': 'ACTIVE'},      # corrupted row
    ])
    master = {'11111111', '22222222', '33333333', '44444444', '55555555', '66666666', '77777777',
              '4806666666666', '4809999999991'}   # last two: a Master item that collides, and a future barcode-as-code item
    r = Resolver(cw, master)
    checks = []

    def expect(label, got, want):
        ok = got == want
        checks.append((label, ok, got, want))

    expect('old code unchanged', (r.resolve('11111111 ').status, r.resolve('11111111 ').code), (AS_IS, '11111111'))
    expect('barcode translated', (r.resolve('4800153151255').status, r.resolve('4800153151255').code), (TRANSLATED, '11111111'))
    expect('leading-zero barcode translated (text)', r.resolve('020800306348').code, '22222222')
    expect('8-digit barcode translated', r.resolve('48033077').code, '33333333')
    expect('future barcode-as-Item_Code accepted as-is', (r.resolve('4809999999991').status, r.resolve('4809999999991').code), (AS_IS, '4809999999991'))
    expect('unknown barcode unresolved', r.resolve('4801111111111').status, UNRESOLVED)
    expect('scientific notation corrupted', r.resolve('4.80015E+12').status, CORRUPTED)
    expect('trailing .0 corrupted', r.resolve('4800153151255.0').status, CORRUPTED)
    expect('lost leading zero reported, not repaired', (r.resolve('20800306348').status, r.resolve('20800306348').code), (CORRUPTED, None))
    expect('embedded space corrupted', r.resolve('4800 153151255').status, CORRUPTED)
    expect('empty string', r.resolve('').status, EMPTY)
    expect('None', r.resolve(None).status, EMPTY)
    expect('nan', r.resolve('nan').status, EMPTY)
    expect('MANUAL-ENTRY', r.resolve('MANUAL-ENTRY').status, EMPTY)
    expect('float input is not trusted', r.resolve(4800153151255.0).status, CORRUPTED)
    expect('duplicate barcode is a conflict', r.resolve('4800000000017').status, CONFLICT)
    expect('barcode equal to a different Master code is a conflict', r.resolve('4806666666666').status, CONFLICT)
    expect('bad crosswalk row ignored, reported', any('scientific notation' in p for p in cw.problems), True)
    expect('resolve_to_barcode', r.resolve_to_barcode('11111111'), '4800153151255')
    expect('resolve_to_barcode unknown', r.resolve_to_barcode('99999999'), None)
    expect('collision list finds the Master collision', any('4806666666666' in c for c in r.find_collisions()), True)
    expect('new generated code colliding with a barcode is unsafe', r.is_safe_new_internal_code('48033077'), False)
    expect('new generated code with nothing known is safe', r.is_safe_new_internal_code('12345678'), True)
    expect('sheet text code keeps leading zero', sheet_text_code('020800306348'), "'020800306348")
    expect('sheet text code leaves old codes alone', sheet_text_code('90566769'), '90566769')
    expect('gtin check ok', gtin_check_ok('4800153151255'), True)
    expect('gtin check bad', gtin_check_ok('4800153151256'), False)
    expect('barcode_text_problems sci', 'scientific notation' in barcode_text_problems('4.80015E+12'), True)

    import pandas as pd
    mdf = pd.DataFrame({'Item_Code': ['11111111', '22222222', '33333333'], 'Product_Name': ['ITEM A', 'ITEM B', 'ITEM C']})
    sales = pd.DataFrame({'Item Code': ['11111111 ', '020800306348', '33333333', '4801111111111', '4.80015E+12', 'MANUAL-ENTRY', '22222222', '33333333'],
                          'Description': ['ITEM A', 'ITEM B', 'ITEM A', 'ITEM C', 'ITEM C', 'ITEM C', 'NOT IN MASTER', 'ITEM C'],
                          'Qty': [1, 2, 3, 4, 5, 6, 7, 8]})
    rc = reconcile_sales(sales, mdf, r)
    cats = [ln['category'] for ln in rc.lines]
    expect('reconcile categories', cats, [BOTH, BOTH, HELD_CONFLICT, NAME_ONLY, NAME_ONLY, NAME_ONLY, UNMATCHED, BOTH])
    expect('reconcile units add up', rc.reconciles and rc.units_file == 36, True)
    expect('reconcile asks for confirmation', rc.needs_confirmation, True)
    expect('conflict is held back', rc.held_indices, [2])
    clean = pd.DataFrame({'Item Code': ['11111111 ', '4800153151255'], 'Description': ['ITEM A', 'ITEM A'], 'Qty': [1, 2]})
    expect('clean file needs no confirmation', reconcile_sales(clean, mdf, r).needs_confirmation, False)

    failed = [c for c in checks if not c[1]]
    if verbose:
        for label, ok, got, want in checks:
            print(('PASS  ' if ok else 'FAIL  ') + label + ('' if ok else f'   got {got!r}, wanted {want!r}'))
        print(f'{len(checks) - len(failed)} of {len(checks)} self-test checks passed.')
    return not failed


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    args = sys.argv[1:]
    if not args or args[0] in ('-h', '--help'):
        print('Usage:\n'
              '  python item_code_resolver.py --self-test\n'
              '  python item_code_resolver.py --summary <BITPOS file>     (item export, price list or sales export)\n'
              '  python item_code_resolver.py --check-collisions\n'
              '  python item_code_resolver.py --set-status <PLANNED|ACTIVE|REVERTED> <code1,code2,...>')
    elif args[0] == '--self-test':
        sys.exit(0 if run_self_test() else 1)
    elif args[0] == '--summary' and len(args) > 1:
        fs = summarize_file(args[1])
        print('\n'.join(fs.lines()))
        log_event(f'RESOLVER SUMMARY of {os.path.basename(args[1])}: ' + ' | '.join(fs.report.lines(5)) + f' | stop={fs.stop}')
        sys.exit(1 if fs.stop else 0)
    elif args[0] == '--check-collisions':
        rv = get_resolver()
        found = rv.find_collisions() + rv.crosswalk.problems
        print(f'Crosswalk rows in use: {len(rv.crosswalk)}; Master item codes: {len(rv.master_codes)}')
        print('\n'.join(found) if found else 'No collisions and no crosswalk problems.')
        sys.exit(1 if found else 0)
    elif args[0] == '--set-status' and len(args) > 2:
        n, b = set_crosswalk_status([c for c in args[2].split(',') if c.strip()], args[1])
        print(f'{n} item(s) set to {args[1].upper()}. Backup: {b}')
    else:
        print('Unrecognised arguments. Run with --help.')
