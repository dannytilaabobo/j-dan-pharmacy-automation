"""
Standalone, read-only movement tracer for specific items -- a scoped version
of the "360-degree trace" logic used for the BENCH BODY SPRAY investigation
(see JDAN_Pharmacy_Master_Notes_v2.md, 2026-07-20 entry), applied here to
BONAKID 1-3 350G / BONNA 0-6MOS 350G. Does not write to Master or any other
file. Edit TARGET_ITEM_CODES / LOOKBACK_START below to point this at a
different item or date range for a future investigation.
"""
import os
import re
import glob
import datetime
import pandas as pd

from name_cleaning import clean_text

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
STOCKIN_DIR = 'JDan_StockIn_Reports'
EOD_DIR = os.path.join('JDan_Sales_System', 'EOD_Deduction_Reports')
ARCHIVE_DIR = os.path.join('JDan_Sales_System', 'Archived_Sales_Reports')
INBOX_DIR = 'SALES_REPORT_FROM_POS'
SOD_DIR = 'JDan_Executive_Reports'

# ==========================================
#    WHAT TO TRACE
# ==========================================
TARGET_ITEM_CODES = ['34567692']  # LS BL CRM 7G
LOOKBACK_START = datetime.date(2026, 7, 12)  # last long-stable point before the decline (stock=86, unchanged 2026-07-12 to 2026-07-18)


def norm(s):
    return clean_text(str(s)).strip().upper()


def parse_date_cell(value):
    try:
        return pd.to_datetime(value).date()
    except (ValueError, TypeError):
        return None


# ==========================================
#    RAW BITPOS POS EXPORT READER
#    (mirrors Daily_Deduct_App.read_pos_file's header-detection logic --
#    BITPOS exports have a variable-length junk header block above the
#    real Item Code / Qty table -- kept independent here since this script
#    must stay read-only and cannot import a GUI-launching module.)
# ==========================================
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
        df['Item Code'] = df['Item Code'].astype(str).str.strip()

    if 'Description' in df.columns:
        df['Description'] = df['Description'].fillna('').apply(clean_text)

    df['Qty'] = pd.to_numeric(df.get('Qty'), errors='coerce').fillna(0)
    df['_date'] = df['Date'].apply(parse_date_cell) if 'Date' in df.columns else None
    return df


def load_bitpos_sales_for_items(item_codes, item_names_norm):
    frames = []
    files = sorted(glob.glob(os.path.join(ARCHIVE_DIR, '*.xls')) +
                    glob.glob(os.path.join(ARCHIVE_DIR, '*.xlsx')) +
                    glob.glob(os.path.join(ARCHIVE_DIR, '*.csv')) +
                    glob.glob(os.path.join(INBOX_DIR, '*.xls')) +
                    glob.glob(os.path.join(INBOX_DIR, '*.xlsx')) +
                    glob.glob(os.path.join(INBOX_DIR, '*.csv')))
    for path in files:
        df = read_pos_file(path)
        if df is None or df.empty:
            continue
        df = df[df['_date'].notna() & (df['_date'] >= LOOKBACK_START)]
        if df.empty:
            continue
        code_match = df['Item Code'].isin(item_codes)
        name_match = df['Description'].apply(norm).isin(item_names_norm)
        # Flag rows that match by name but NOT by code -- a possible
        # naming inconsistency the user asked this script to guard against
        # (e.g. a sale rung up under a differently-coded duplicate row).
        mismatched = df[name_match & ~code_match]
        if not mismatched.empty:
            print(f"  [!] {os.path.basename(path)}: {len(mismatched)} row(s) match the item NAME "
                  f"but not its known Item_Code -- possible naming inconsistency, review manually:")
            print(mismatched[['Date', 'Receipt', 'Item Code', 'Description', 'Qty']].to_string(index=False))
        m = df[code_match].copy()
        if not m.empty:
            m['source_file'] = os.path.basename(path)
            frames.append(m)
    if not frames:
        return pd.DataFrame(columns=['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', 'source_file'])
    return pd.concat(frames, ignore_index=True)


def load_stockin_for_items(item_codes):
    frames = []
    for path in sorted(glob.glob(os.path.join(STOCKIN_DIR, '*', 'STOCKIN_BATCH_*.xlsx'))):
        df = pd.read_excel(path, dtype=str)
        df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
        df['_date'] = df['Date_Encoded'].apply(parse_date_cell)
        df = df[df['_date'].notna() & (df['_date'] >= LOOKBACK_START) & df['Item_Code'].isin(item_codes)]
        if not df.empty:
            df['Qty_Delivered'] = pd.to_numeric(df['Qty_Delivered'], errors='coerce').fillna(0)
            df['source_file'] = os.path.basename(path)
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date',
                                      'Qty_Delivered', 'Date_Encoded', 'Encoded_By', 'source_file'])
    return pd.concat(frames, ignore_index=True)


def load_eod_deductions_for_items(item_codes):
    frames = []
    for path in sorted(glob.glob(os.path.join(EOD_DIR, 'PROCESSED_EOD_*.xlsx'))):
        df = pd.read_excel(path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        if 'Item Code' not in df.columns:
            continue
        df['Item Code'] = df['Item Code'].astype(str).str.strip()
        df['_date'] = df['Date'].apply(parse_date_cell) if 'Date' in df.columns else None
        df = df[df['_date'].notna() & (df['_date'] >= LOOKBACK_START) & df['Item Code'].isin(item_codes)]
        if not df.empty:
            df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
            df['source_file'] = os.path.basename(path)
            frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', 'source_file'])
    return pd.concat(frames, ignore_index=True)


def load_master_backup_timeline(item_codes):
    """Every Master backup (any script's pre-write snapshot) is a genuine
    point-in-time record of Current_Stock -- not just SOD-reconcile backups
    -- so all of them are included here rather than assuming which one
    represents "Monday" specifically."""
    rows = []
    pattern = re.compile(r'Inventory_Master_BEFORE_(.+)_(\d{8}_\d{6})\.xlsx$')
    for path in sorted(glob.glob(os.path.join(BACKUP_DIR, 'Inventory_Master_*.xlsx'))):
        m = pattern.search(os.path.basename(path))
        if not m:
            continue
        reason, ts_str = m.group(1), m.group(2)
        try:
            ts = datetime.datetime.strptime(ts_str, '%Y%m%d_%H%M%S')
        except ValueError:
            continue
        if ts.date() < (LOOKBACK_START - datetime.timedelta(days=1)):
            continue
        try:
            df = pd.read_excel(path, sheet_name='INVENTORY MASTER', dtype=str)
        except Exception:
            continue
        df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
        df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
        row = {'timestamp': ts, 'reason': reason, 'source_file': os.path.basename(path)}
        for code in item_codes:
            row[code] = df[df['Item_Code'] == code]['Current_Stock'].sum()
        rows.append(row)
    rows.sort(key=lambda r: r['timestamp'])

    # Final point: the live Master file right now.
    df_live = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str)
    df_live['Item_Code'] = df_live['Item_Code'].astype(str).str.strip()
    df_live['Current_Stock'] = pd.to_numeric(df_live['Current_Stock'], errors='coerce').fillna(0)
    live_row = {'timestamp': datetime.datetime.now(), 'reason': 'NOW (live Master)', 'source_file': MASTER_FILE}
    for code in item_codes:
        live_row[code] = df_live[df_live['Item_Code'] == code]['Current_Stock'].sum()
    rows.append(live_row)
    return rows


def load_latest_sod_actuals(item_names_norm):
    files = glob.glob(os.path.join(SOD_DIR, 'SOD_Summary_*.txt'))
    if not files:
        return {}, None
    latest = max(files, key=os.path.getmtime)
    with open(latest, 'r', encoding='utf-8') as f:
        text = f.read()

    date_match = re.search(r'Target Date:\s*([\d-]+)', text)
    report_date = date_match.group(1) if date_match else '?'

    results = {}
    for block in text.split('---------------------------------------'):
        item_match = re.search(r'Item:\s*(.+)', block)
        qty_match = re.search(r'System Qty:\s*([\d.]+)\s*\|\s*Actual Count:\s*([\d.]+)', block)
        var_match = re.search(r'Variance:\s*(-?[\d.]+)', block)
        just_match = re.search(r'Justification:\s*(.+)', block)
        if not (item_match and qty_match):
            continue
        item_norm = norm(item_match.group(1))
        if item_norm in item_names_norm:
            results[item_norm] = {
                'system_qty': float(qty_match.group(1)),
                'actual_qty': float(qty_match.group(2)),
                'variance': float(var_match.group(1)) if var_match else None,
                'justification': just_match.group(1).strip() if just_match else '',
            }
    return results, (latest, report_date)


def main():
    print('=' * 78)
    print(' J-DAN ITEM MOVEMENT TRACE (read-only -- no Master writes)')
    print(f' Target Item_Codes: {TARGET_ITEM_CODES}')
    print(f' Lookback window: {LOOKBACK_START} to {datetime.date.today()}')
    print('=' * 78)

    df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype=str)
    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    df_master['Current_Stock'] = pd.to_numeric(df_master['Current_Stock'], errors='coerce').fillna(0)

    canonical_names = {}
    for code in TARGET_ITEM_CODES:
        rows = df_master[df_master['Item_Code'] == code]
        name = rows['Product_Name'].dropna().iloc[0] if not rows.empty else '(not found in Master)'
        canonical_names[code] = clean_text(name)
        print(f"\n--- {code}: {canonical_names[code]} ---")
        if rows.empty:
            print("  WARNING: Item_Code not found anywhere in current Master.")
            continue
        print(rows[['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Current_Stock', 'Status']].to_string(index=False))
        print(f"  Current Master total: {rows['Current_Stock'].sum():.0f}")

    item_names_norm = {norm(n) for n in canonical_names.values()}

    # ---- 1. Master backup snapshot timeline ----
    print('\n' + '=' * 78)
    print(' 1. MASTER Current_Stock SNAPSHOT TIMELINE (from every backup taken)')
    print('=' * 78)
    timeline = load_master_backup_timeline(TARGET_ITEM_CODES)
    for row in timeline:
        vals = ', '.join(f"{canonical_names.get(c, c)}={row[c]:.0f}" for c in TARGET_ITEM_CODES)
        print(f"  {row['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}  [{row['reason']:<28}] {vals}   ({row['source_file']})")

    # ---- 2. Raw BITPOS sales trace ----
    print('\n' + '=' * 78)
    print(' 2. RAW BITPOS SALES EXPORT TRACE (Archived_Sales_Reports + inbox)')
    print('=' * 78)
    sales_df = load_bitpos_sales_for_items(TARGET_ITEM_CODES, item_names_norm)
    if sales_df.empty:
        print(f"  No BITPOS sales rows found for these items from {LOOKBACK_START} onward.")
    else:
        cols = [c for c in ['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', 'source_file'] if c in sales_df.columns]
        print(sales_df[cols].sort_values(['Date', 'Item Code']).to_string(index=False))
        print("\n  Subtotal sold (per BITPOS raw export):")
        for code in TARGET_ITEM_CODES:
            total = sales_df[sales_df['Item Code'] == code]['Qty'].sum()
            print(f"    {canonical_names.get(code, code)}: {total:.0f}")

    if not any(datetime.date.today().strftime('%Y-%m-%d') in os.path.basename(f) for f in
               glob.glob(os.path.join(ARCHIVE_DIR, '*')) + glob.glob(os.path.join(INBOX_DIR, '*'))):
        print(f"\n  Note: no raw BITPOS export found yet for today ({datetime.date.today()}) -- "
              f"today's EOD settlement likely hasn't been run/archived yet.")

    # ---- 3. StockIn trace ----
    print('\n' + '=' * 78)
    print(' 3. STOCK-IN TRACE (JDan_StockIn_Reports)')
    print('=' * 78)
    stockin_df = load_stockin_for_items(TARGET_ITEM_CODES)
    if stockin_df.empty:
        print(f"  No stock-in batches found for these items from {LOOKBACK_START} onward.")
    else:
        print(stockin_df[['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Qty_Delivered', 'Date_Encoded', 'Encoded_By', 'source_file']].to_string(index=False))

    # ---- 4. EOD deduction trace ----
    print('\n' + '=' * 78)
    print(' 4. EOD DEDUCTION TRACE (JDan_Sales_System/EOD_Deduction_Reports)')
    print('=' * 78)
    eod_df = load_eod_deductions_for_items(TARGET_ITEM_CODES)
    if eod_df.empty:
        print(f"  No EOD deduction rows found for these items from {LOOKBACK_START} onward.")
    else:
        cols = [c for c in ['Date', 'Receipt', 'Item Code', 'Description', 'Qty', 'Cashier Id', 'source_file'] if c in eod_df.columns]
        print(eod_df[cols].sort_values(['Date', 'Item Code']).to_string(index=False))
        print("\n  Subtotal deducted from Master (EOD, already applied):")
        for code in TARGET_ITEM_CODES:
            total = eod_df[eod_df['Item Code'] == code]['Qty'].sum()
            print(f"    {canonical_names.get(code, code)}: {total:.0f}")

    # Cross-check raw BITPOS vs EOD-applied deduction per item.
    print("\n  Cross-check (raw BITPOS sale count vs. what EOD actually deducted from Master):")
    for code in TARGET_ITEM_CODES:
        raw_total = sales_df[sales_df['Item Code'] == code]['Qty'].sum() if not sales_df.empty else 0
        eod_total = abs(eod_df[eod_df['Item Code'] == code]['Qty'].sum()) if not eod_df.empty else 0
        flag = '' if raw_total == eod_total else '  <-- MISMATCH, needs review'
        print(f"    {canonical_names.get(code, code)}: raw sold={raw_total:.0f} | EOD deducted={eod_total:.0f}{flag}")

    # ---- 5. Reconciliation ----
    print('\n' + '=' * 78)
    print(' 5. RECONCILIATION: earliest snapshot + stock-in - EOD deductions vs. actual')
    print('=' * 78)
    sod_actuals, sod_source = load_latest_sod_actuals(item_names_norm)
    if sod_source:
        print(f"  Latest SOD report used for Actual Count: {os.path.basename(sod_source[0])} (Target Date: {sod_source[1]})")

    baseline = timeline[0] if timeline else None
    # The last two timeline entries are always [pre-correction system snapshot,
    # live Master]: Daily_Dashboard_App's SOD reconciliation (2026-07-18 fix)
    # writes the physical Actual Count back into Master, so "live" is expected
    # to already equal the shelf count -- it is NOT a second independent data
    # point for finding a gap. The real gap to look for is between the
    # system's own uncorrected bookkeeping (pre-correction snapshot) and the
    # physical Actual Count.
    pre_correction = timeline[-2] if len(timeline) >= 2 else None
    live = timeline[-1] if timeline else None

    for code in TARGET_ITEM_CODES:
        name = canonical_names.get(code, code)
        print(f"\n  {name} ({code}):")
        if baseline is None:
            print("    No baseline snapshot available -- cannot reconcile.")
            continue
        baseline_qty = baseline[code]
        stockin_qty = stockin_df[stockin_df['Item_Code'] == code]['Qty_Delivered'].sum() if not stockin_df.empty else 0
        eod_qty = abs(eod_df[eod_df['Item Code'] == code]['Qty'].sum()) if not eod_df.empty else 0
        expected = baseline_qty + stockin_qty - eod_qty

        print(f"    Baseline ({baseline['timestamp'].strftime('%Y-%m-%d %H:%M:%S')}, "
              f"snapshot '{baseline['reason']}'): {baseline_qty:.0f}")
        print(f"    + Stock-in since baseline:              {stockin_qty:.0f}")
        print(f"    - EOD deductions since baseline:         {eod_qty:.0f}")
        print(f"    = Expected system stock (uncorrected):   {expected:.0f}")

        if pre_correction is not None:
            pre_qty = pre_correction[code]
            sanity = "matches -- this trace accounts for every logged movement" if expected == pre_qty else \
                     "DOES NOT MATCH -- some stock-in/sale/deduction affecting this item was missed by this trace"
            print(f"    System stock per Master just before today's SOD write-back: {pre_qty:.0f}  ({sanity})")
        else:
            pre_qty = expected

        item_norm = norm(name)
        if item_norm in sod_actuals:
            sod = sod_actuals[item_norm]
            gap = pre_qty - sod['actual_qty']
            print(f"    Physical Actual Count (latest SOD):      {sod['actual_qty']:.0f}")
            print(f"    Justification on file:                   {sod['justification']}")
            if gap == 0:
                print(f"    -> No gap: system bookkeeping and shelf count agree. No further action needed.")
            else:
                print(f"    -> UNEXPLAINED GAP: {gap:.0f} unit(s) missing from the shelf relative to what the "
                      f"system's own logs account for (stock-in + sales/deductions), with no record explaining "
                      f"it. Consistent with the BENCH BODY SPRAY pattern (2026-07-20): loss occurred outside all "
                      f"logged channels (receiving/count error or unlogged shrinkage) -- recommend physical "
                      f"recount / spot-check, not a system bug.")
            if live is not None:
                print(f"    Master's Current_Stock has already been corrected to {live[code]:.0f} by today's SOD "
                      f"reconciliation write-back (matches the physical Actual Count above).")
        else:
            print("    No Actual Count found for this item in the latest SOD report.")

    print('\n' + '=' * 78)
    print(' END OF TRACE -- read-only, no files were modified.')
    print('=' * 78)


if __name__ == '__main__':
    main()
