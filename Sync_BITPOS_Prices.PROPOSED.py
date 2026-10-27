# ============================================================================
# PROPOSED VERSION -- NOT ACTIVE. The live Sync_BITPOS_Prices.py is unchanged.
# Differences from the live file (see Sync_BITPOS_Prices.PROPOSED.diff and
# Sync_BITPOS_Prices.PROPOSED_EXPLANATION.md in Foundation_Build_*):
#   1. Every BITPOS ITEM CODE is translated to the internal (old) Item_Code through
#      item_code_resolver BEFORE matching, so a barcode never looks like a new item
#      and its old code never looks "gone".
#   2. An item whose code resolves to an existing internal item is never marked
#      DISCONTINUED.
#   3. HARD STOP (nothing is asked, nothing is applied) if more than 10 items would
#      be marked DISCONTINUED, or more than 10 new items would be created, or any
#      BITPOS code is damaged / empty / ambiguous.
#   4. --dry-run prints the preview and writes NOTHING (no preview file, no log).
#   5. Master is read with Item_Code as TEXT (leading zeros of barcode codes survive).
# The preview, the per-name questions and the append-only audit log are unchanged.
# ============================================================================
import os
import sys
import glob
import shutil
import datetime
import pandas as pd

from name_cleaning import clean_text
from item_code_resolver import (Resolver, ResolutionReport, load_crosswalk, announce,
                                CORRUPTED, EMPTY, CONFLICT, UNRESOLVED)

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BITPOS_GLOB = 'bitpos_sales_pricing_*.csv'
BACKUP_DIR = 'JDan_Master_Backups'
REPORT_DIR = 'JDan_Executive_Reports'
AUDIT_LOG_PATH = os.path.join(REPORT_DIR, 'BITPOS_SYNC_AUDIT_LOG.csv')
PENDING_PATH = os.path.join(REPORT_DIR, 'BITPOS_PENDING_NAME_VERIFICATION.csv')
PENDING_COLUMNS = ['Item_Code', 'Old_Name', 'New_Name', 'Flagged_Date', 'Status', 'Note']

# BITPOS is the authority for Retail_Price, Product_Name, and whether an item
# is still carried. Item_Code is the matching key (never Product_Name, since
# a display name can change without the code changing). This script NEVER
# touches Current_Stock -- that field is owned exclusively by
# StockIn_App/Daily_Deduct_App/Expiry_Scanner/Daily_Dashboard_App.

# HARD STOP: if more than this many items would be newly marked DISCONTINUED in
# one run, the script stops and applies nothing -- it almost certainly means the
# BITPOS export is incomplete or its codes changed (e.g. barcodes), not that many
# items were really discontinued at once.
MASS_DISCONTINUE_HARD_STOP = 10

# Same for new items: an unusually large batch is more likely an incomplete or
# mismatched BITPOS export than a real bulk product launch.
MASS_NEW_ITEM_HARD_STOP = 10

# Filled by detect_changes(): what the code translation found (read by main()/tests).
CODE_CHECK = {'blocked': [], 'ambiguous': [], 'report': None}

# Columns written to Inventory_Master.xlsx for a brand-new item added via
# BITPOS sync. Only Item_Code/Product_Name/Retail_Price come from BITPOS --
# Current_Stock starts at 0 (no stock-in has happened yet), and
# Category/Supplier/Reorder_Point/Is_Generic are left blank for the user to
# fill in later via Item_Settings_App. Date_Encoded/Encoded_By/Lot_Number/
# Expiry_Date are owned by StockIn_App and stay blank until an actual
# stock-in happens.
NEW_ITEM_DEFAULTS = {
    'Current_Stock': 0,
    'Lot_Number': '',
    'Expiry_Date': '',
    'Date_Encoded': '',
    'Encoded_By': '',
    'Pulled_Out_By': '',
    'LOCATION': '',
    'Status': 'ACTIVE',
    'Discontinued_Date': '',
    'Discontinued_Reason': '',
    'Reorder_Point': '',
    'Supplier': '',
    'Category': '',
    'Is_Generic': '',
}


def latest_bitpos_file():
    matches = glob.glob(BITPOS_GLOB)
    if not matches:
        raise FileNotFoundError(f'No file matching {BITPOS_GLOB} found in this folder.')
    return max(matches, key=os.path.getmtime)


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_BITPOS_SYNC_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def load_bitpos(source_file):
    df = pd.read_csv(source_file, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df['ITEM CODE'] = df['ITEM CODE'].astype(str).str.strip()
    df['Retail'] = pd.to_numeric(df['Retail'], errors='coerce')
    df['DESCRIPTION'] = df['DESCRIPTION'].apply(clean_text)
    return df


def build_resolver(df_master):
    return Resolver(load_crosswalk(), set(df_master['Item_Code'].astype(str).str.strip()))


def resolve_bitpos_codes(df_bitpos, resolver):
    """Adds INTERNAL_CODE / CODE_STATUS / CODE_REASON / CODE_BARCODE columns. Rows are never dropped here."""
    results = [resolver.resolve(c) for c in df_bitpos['ITEM CODE']]
    df = df_bitpos.copy()
    df['CODE_STATUS'] = [r.status for r in results]
    df['CODE_REASON'] = [r.reason for r in results]
    df['CODE_BARCODE'] = [r.barcode for r in results]
    df['INTERNAL_CODE'] = [r.code if r.ok else None for r in results]
    report = ResolutionReport()
    for r in results:
        report.add(r)
    return df, report


def load_pending():
    if os.path.exists(PENDING_PATH):
        df = pd.read_csv(PENDING_PATH, dtype=str)
        return df
    return pd.DataFrame(columns=PENDING_COLUMNS)


def save_pending(df):
    os.makedirs(REPORT_DIR, exist_ok=True)
    df.to_csv(PENDING_PATH, index=False)


def detect_changes(df_master, df_bitpos, resolver=None):
    # BITPOS code -> internal Item_Code first. Codes that cannot be trusted (damaged, empty,
    # ambiguous, or a crosswalk barcode whose old code is missing from Master) are NOT matched
    # and NOT turned into new items; they are collected in CODE_CHECK['blocked'].
    if resolver is None:
        resolver = build_resolver(df_master)
    df_bitpos, code_report = resolve_bitpos_codes(df_bitpos, resolver)
    blocked_mask = df_bitpos['CODE_STATUS'].isin([CORRUPTED, EMPTY, CONFLICT]) | \
        ((df_bitpos['CODE_STATUS'] == UNRESOLVED) & df_bitpos['CODE_BARCODE'].notna())
    blocked = df_bitpos[blocked_mask][['ITEM CODE', 'DESCRIPTION', 'CODE_STATUS', 'CODE_REASON']].to_dict('records')
    usable = df_bitpos[~blocked_mask].copy()
    # key used for matching: the internal code when known, otherwise the code exactly as typed (a real new item)
    usable['MATCH_CODE'] = usable['INTERNAL_CODE'].where(usable['INTERNAL_CODE'].notna(), usable['ITEM CODE'])
    dup_keys = usable[usable['MATCH_CODE'].duplicated(keep=False)]
    ambiguous = dup_keys[['ITEM CODE', 'DESCRIPTION', 'MATCH_CODE']].to_dict('records')
    usable = usable.drop_duplicates('MATCH_CODE', keep='first')
    CODE_CHECK.update({'blocked': blocked, 'ambiguous': ambiguous, 'report': code_report})

    bitpos_codes = set(usable['MATCH_CODE'])
    bitpos_price = dict(zip(usable['MATCH_CODE'], usable['Retail']))
    bitpos_name = dict(zip(usable['MATCH_CODE'], usable['DESCRIPTION']))

    price_changes = []
    name_changes = []
    discontinued = []
    new_items = []

    active_mask = df_master['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED' if 'Status' in df_master.columns else pd.Series(True, index=df_master.index)
    master_codes = df_master.loc[active_mask, 'Item_Code'].unique()

    # Checked against ALL Master codes (active AND discontinued) -- a code
    # that already exists in Master, even discontinued, is a reactivation,
    # not a "new" item, and is out of scope here.
    all_master_codes = set(df_master['Item_Code'].unique())
    seen_new_codes = set()
    for code in usable['MATCH_CODE']:
        if code in all_master_codes or code in seen_new_codes:
            continue
        seen_new_codes.add(code)
        new_items.append({
            'Item_Code': code,
            'Description': bitpos_name.get(code, ''),
            'Change_Type': 'NEW ITEM ADDED via BITPOS sync',
            'Old_Value': '',
            'New_Value': bitpos_price.get(code),
        })

    for code in master_codes:
        mask = (df_master['Item_Code'] == code) & active_mask
        rows = df_master.loc[mask]
        current_name = clean_text(rows['Product_Name'].iloc[0])

        if code not in bitpos_codes:
            discontinued.append({
                'Item_Code': code,
                'Description': current_name,
                'Change_Type': 'DISCONTINUED',
                'Old_Value': 'ACTIVE',
                'New_Value': 'DISCONTINUED',
            })
            continue

        new_price = bitpos_price.get(code)
        if pd.notna(new_price):
            old_prices = pd.to_numeric(rows['Retail_Price'], errors='coerce').unique()
            old_price = old_prices[0] if len(old_prices) else None
            if old_price is None or pd.isna(old_price) or abs(old_price - new_price) > 0.001:
                price_changes.append({
                    'Item_Code': code,
                    'Description': current_name,
                    'Change_Type': 'PRICE',
                    'Old_Value': old_price,
                    'New_Value': new_price,
                })

        new_name = bitpos_name.get(code, '')
        if new_name and new_name != current_name:
            name_changes.append({
                'Item_Code': code,
                'Description': current_name,
                'Change_Type': 'NAME',
                'Old_Value': current_name,
                'New_Value': new_name,
            })

    return price_changes, name_changes, discontinued, new_items


def print_preview(price_changes, name_changes, discontinued, new_items):
    print('=' * 70)
    print(' PREVIEW -- NO CHANGES HAVE BEEN APPLIED YET')
    print('=' * 70)

    print(f'\n--- (1) PRICE CHANGES: {len(price_changes)} item(s) ---')
    if price_changes:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(price_changes)[['Item_Code', 'Description', 'Old_Value', 'New_Value']].to_string(index=False))
    else:
        print('  None.')

    print(f'\n--- (2) NAME CHANGES: {len(name_changes)} item(s) (each will be asked individually below) ---')
    if name_changes:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(name_changes)[['Item_Code', 'Old_Value', 'New_Value']].rename(
                columns={'Old_Value': 'Old_Name', 'New_Value': 'New_Name'}).to_string(index=False))
    else:
        print('  None.')

    print(f'\n--- (3) NEWLY DISCONTINUED (no longer in the BITPOS export): {len(discontinued)} item(s) ---')
    if discontinued:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(discontinued)[['Item_Code', 'Description']].to_string(index=False))
        if len(discontinued) > MASS_DISCONTINUE_HARD_STOP:
            print(f'\n  WARNING: {len(discontinued)} items would lose their match in the BITPOS export -- '
                  f'more than the limit of {MASS_DISCONTINUE_HARD_STOP}. This script will STOP and apply nothing.')
    else:
        print('  None.')

    print(f'\n--- (4) NEW ITEMS FOUND (in BITPOS, no match in Master): {len(new_items)} item(s) ---')
    if new_items:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(new_items)[['Item_Code', 'Description', 'New_Value']].rename(
                columns={'Description': 'Product_Name', 'New_Value': 'Retail_Price'}).to_string(index=False))
        if len(new_items) > MASS_NEW_ITEM_HARD_STOP:
            print(f'\n  WARNING: {len(new_items)} items have no match in Master -- '
                  f'more than the limit of {MASS_NEW_ITEM_HARD_STOP}. This script will STOP and apply nothing.')
        print('  New items will be added with Current_Stock = 0 (no stock-in yet) and '
              'Category/Supplier/Reorder_Point left blank for you to fill in later via Item Settings.')
    else:
        print('  None.')

    print('\nCurrent_Stock is NEVER touched by this script (except new items, which start at 0).')


def resolve_name_changes(name_changes, pending_df):
    """Interactively decide each detected name change: apply now, or defer to
    the Pending Verification list. Items already PENDING from a previous run
    are skipped automatically -- not re-asked every sync."""
    pending_lookup = {}
    if not pending_df.empty:
        for _, r in pending_df[pending_df['Status'] == 'PENDING'].iterrows():
            pending_lookup[r['Item_Code']] = r

    confirmed = []
    newly_deferred = []
    already_pending = []

    for c in name_changes:
        code = c['Item_Code']
        if code in pending_lookup:
            already_pending.append(c)
            continue

        print(f"\nItem_Code {code}: '{c['Old_Value']}' -> '{c['New_Value']}'")
        answer = input("  Apply this rename now? (yes = apply / no = defer to Pending Verification): ").strip().lower()
        if answer == 'yes':
            confirmed.append(c)
        else:
            newly_deferred.append(c)

    if already_pending:
        print(f"\n{len(already_pending)} name change(s) AUTO-SKIPPED (already pending verification from a previous run):")
        for c in already_pending:
            print(f"  {c['Item_Code']}: '{c['Old_Value']}' -> '{c['New_Value']}'")

    return confirmed, newly_deferred, already_pending


def apply_changes(df_master, price_changes, name_changes, discontinued, new_items):
    if 'Status' not in df_master.columns:
        df_master['Status'] = 'ACTIVE'
    if 'Discontinued_Date' not in df_master.columns:
        df_master['Discontinued_Date'] = ''
    if 'Discontinued_Reason' not in df_master.columns:
        df_master['Discontinued_Reason'] = ''
    df_master['Status'] = df_master['Status'].fillna('ACTIVE').replace('', 'ACTIVE')

    today = datetime.datetime.now().strftime('%Y-%m-%d')

    for c in price_changes:
        df_master.loc[df_master['Item_Code'] == c['Item_Code'], 'Retail_Price'] = c['New_Value']

    for c in name_changes:
        df_master.loc[df_master['Item_Code'] == c['Item_Code'], 'Product_Name'] = c['New_Value']

    for c in discontinued:
        mask = df_master['Item_Code'] == c['Item_Code']
        df_master.loc[mask, 'Status'] = 'DISCONTINUED'
        df_master.loc[mask, 'Discontinued_Date'] = today
        df_master.loc[mask, 'Discontinued_Reason'] = 'No longer in the BITPOS export -- automatically detected by Sync_BITPOS_Prices.py'

    if new_items:
        new_rows = []
        for c in new_items:
            row = {col: NEW_ITEM_DEFAULTS.get(col, '') for col in df_master.columns}
            row['Item_Code'] = c['Item_Code']
            row['Product_Name'] = c['Description']
            row['Retail_Price'] = c['New_Value']
            new_rows.append(row)
        df_master = pd.concat([df_master, pd.DataFrame(new_rows)], ignore_index=True)

    return df_master


def update_pending_file(pending_df, newly_deferred, resolved_codes):
    today = datetime.datetime.now().strftime('%Y-%m-%d')

    if not pending_df.empty and resolved_codes:
        resolve_mask = pending_df['Item_Code'].isin(resolved_codes) & (pending_df['Status'] == 'PENDING')
        pending_df.loc[resolve_mask, 'Status'] = 'RESOLVED'

    if newly_deferred:
        new_rows = pd.DataFrame([{
            'Item_Code': c['Item_Code'],
            'Old_Name': c['Old_Value'],
            'New_Name': c['New_Value'],
            'Flagged_Date': today,
            'Status': 'PENDING',
            'Note': 'No stock/verification opportunity yet -- verify on the next stock-in.',
        } for c in newly_deferred])
        pending_df = pd.concat([pending_df, new_rows], ignore_index=True)

    save_pending(pending_df)
    return pending_df


def append_audit_log(all_changes, source_file):
    os.makedirs(REPORT_DIR, exist_ok=True)
    log_df = pd.DataFrame(all_changes)
    log_df.insert(0, 'Synced_At', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    log_df.insert(1, 'Source_File', source_file)
    header_needed = not os.path.exists(AUDIT_LOG_PATH)
    log_df.to_csv(AUDIT_LOG_PATH, mode='a', header=header_needed, index=False)


def save_preview_report(price_changes, name_changes, discontinued, new_items):
    os.makedirs(REPORT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_DIR, f'BITPOS_SYNC_PREVIEW_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        pd.DataFrame(price_changes).to_excel(writer, sheet_name='Price_Changes', index=False)
        pd.DataFrame(name_changes).to_excel(writer, sheet_name='Name_Changes', index=False)
        pd.DataFrame(discontinued).to_excel(writer, sheet_name='Newly_Discontinued', index=False)
        pd.DataFrame(new_items).to_excel(writer, sheet_name='New_Items_Found', index=False)
    return out_path


def hard_stop_reasons(discontinued, new_items):
    reasons = []
    if len(discontinued) > MASS_DISCONTINUE_HARD_STOP:
        reasons.append(f'{len(discontinued)} items would be marked DISCONTINUED (limit {MASS_DISCONTINUE_HARD_STOP}).')
    if len(new_items) > MASS_NEW_ITEM_HARD_STOP:
        reasons.append(f'{len(new_items)} new items would be created (limit {MASS_NEW_ITEM_HARD_STOP}).')
    if CODE_CHECK['blocked']:
        reasons.append(f"{len(CODE_CHECK['blocked'])} BITPOS code(s) are damaged, empty or ambiguous: "
                       + '; '.join(f"{b['ITEM CODE']!r} ({b['CODE_REASON']})" for b in CODE_CHECK['blocked'][:5]))
    if CODE_CHECK['ambiguous']:
        reasons.append(f"{len(CODE_CHECK['ambiguous'])} BITPOS record(s) resolve to the same internal item (old code AND barcode both present?): "
                       + '; '.join(f"{b['ITEM CODE']!r}->{b['MATCH_CODE']}" for b in CODE_CHECK['ambiguous'][:6]))
    return reasons


def compute_plan(df_master, df_bitpos, resolver=None):
    """Everything main() would show, as data (used by the dry-run tests). Writes nothing."""
    price_changes, name_changes, discontinued, new_items = detect_changes(df_master, df_bitpos, resolver)
    return {'price_changes': price_changes, 'name_changes': name_changes, 'discontinued': discontinued,
            'new_items': new_items, 'stop_reasons': hard_stop_reasons(discontinued, new_items),
            'code_check': dict(CODE_CHECK)}


def main(dry_run=False, bitpos_path=None, master_path=None):
    source_file = bitpos_path or latest_bitpos_file()
    print(f'Using BITPOS source: {source_file}')
    if dry_run:
        print('DRY RUN -- nothing will be asked, written or changed.')

    df_bitpos = load_bitpos(source_file)

    try:
        df_master = pd.read_excel(master_path or MASTER_FILE, sheet_name='INVENTORY MASTER', dtype={'Item_Code': str})
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    df_master['Retail_Price'] = pd.to_numeric(df_master['Retail_Price'], errors='coerce')

    price_changes, name_changes, discontinued, new_items = detect_changes(df_master, df_bitpos)

    if (not price_changes and not name_changes and not discontinued and not new_items
            and not CODE_CHECK['blocked'] and not CODE_CHECK['ambiguous']):
        print('\nNo differences -- Master already matches the BITPOS export.')
        return

    announce(CODE_CHECK['report'], 'Sync_BITPOS_Prices', log=not dry_run)
    print_preview(price_changes, name_changes, discontinued, new_items)
    stop_reasons = hard_stop_reasons(discontinued, new_items)
    if dry_run:
        print(f'\nDRY RUN RESULT: {"WOULD STOP - " + " | ".join(stop_reasons) if stop_reasons else "no stop condition"}. Nothing was written.')
        return
    preview_path = save_preview_report(price_changes, name_changes, discontinued, new_items)
    print(f'\nFull preview also saved to: {preview_path}')
    if stop_reasons:
        print('\n' + '!' * 70)
        print(' STOPPED -- NOTHING WAS APPLIED AND NOTHING WAS ASKED.')
        for reason in stop_reasons:
            print(f'  - {reason}')
        print(' Check that the BITPOS export is complete and that every barcode is listed in Barcode_Crosswalk.csv,')
        print(' then run this again. (The preview above was saved so you can review it.)')
        print('!' * 70)
        return

    pending_df = load_pending()
    confirmed_names, newly_deferred, already_pending = resolve_name_changes(name_changes, pending_df)

    print('\n--- FINAL SUMMARY BEFORE APPLYING ---')
    print(f'Price changes: {len(price_changes)}')
    print(f'Name changes to apply now: {len(confirmed_names)}')
    print(f'Name changes to defer to Pending Verification: {len(newly_deferred)}')
    print(f'Name changes auto-skipped (already pending): {len(already_pending)}')
    print(f'Newly discontinued: {len(discontinued)}')
    print(f'New items to add: {len(new_items)}')

    answer = input('\nProceed and apply the CONFIRMED changes to Master (price, approved name changes, discontinued, new items)? (yes/no): ').strip().lower()
    if answer != 'yes':
        print('\nCANCELLED. No changes were applied to Master.')
        if newly_deferred:
            print('Note: the deferred items were also not saved to the Pending Verification list -- re-run the script if you still want to record them.')
        return

    backup_path = backup_master()
    print(f'Backup created: {backup_path}')

    df_master = apply_changes(df_master, price_changes, confirmed_names, discontinued, new_items)

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    resolved_codes = [c['Item_Code'] for c in confirmed_names]
    update_pending_file(pending_df, newly_deferred, resolved_codes)

    all_changes = price_changes + confirmed_names + discontinued + new_items
    if newly_deferred:
        all_changes += [{**c, 'Change_Type': 'NAME_DEFERRED'} for c in newly_deferred]
    append_audit_log(all_changes, source_file)

    print(f'\nAPPLIED: {len(price_changes)} price change(s), {len(confirmed_names)} name change(s), '
          f'{len(discontinued)} newly discontinued item(s), {len(new_items)} new item(s) added.')
    if newly_deferred:
        print(f'{len(newly_deferred)} name change(s) saved to the Pending Verification list: {PENDING_PATH}')
    print(f'Audit trail appended to: {AUDIT_LOG_PATH}')
    if new_items:
        print('NOTE: new items were appended as new rows -- this rewrite may have cleared the Category/Supplier/'
              'Is_Generic Excel dropdowns (a known side effect of any full-sheet rewrite); re-run '
              '"Add_Reorder_And_Supplier_Columns.py" if you need them reattached.')
    print('Current_Stock was not touched for existing items (new items start at 0).')


if __name__ == '__main__':
    argv = sys.argv[1:]
    def _opt(name):
        return argv[argv.index(name) + 1] if name in argv and argv.index(name) + 1 < len(argv) else None
    main(dry_run='--dry-run' in argv, bitpos_path=_opt('--bitpos'), master_path=_opt('--master'))
