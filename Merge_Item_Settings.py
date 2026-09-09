import os
import shutil
import datetime
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'

TARGET_FIELDS = ['Category', 'Supplier', 'Reorder_Point', 'Is_Generic']


def pick_worksheet_path():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title='Select the edited Item Settings worksheet',
        initialdir=os.getcwd(),
        filetypes=[('Excel files', '*.xlsx')],
    )
    root.destroy()
    return path


def is_blank(value):
    if pd.isna(value):
        return True
    return str(value).strip() == ''


def main():
    worksheet_path = pick_worksheet_path()
    if not worksheet_path:
        print('No file selected. Cancelled -- nothing was changed.')
        return

    staff_name = input('Enter your name (for the update log): ').strip()
    if not staff_name:
        print('A name is required to log this update. Cancelled.')
        return

    try:
        # keep_default_na=False stops pandas from silently reading the literal
        # text "N/A" (a valid, intentional Branded_Generic value) as a blank
        # cell -- it would otherwise be indistinguishable from an actually
        # empty cell and get wrongly skipped. Genuinely empty cells still read
        # as NaN regardless of this setting (nothing there for it to affect).
        # sheet_name=0 (first sheet) instead of a hardcoded name so this reads
        # either export's data sheet -- both put their DropdownLists helper
        # sheet after the data sheet, so index 0 is always the right one.
        ws_df = pd.read_excel(worksheet_path, sheet_name=0, keep_default_na=False, na_values=[''])
    except Exception as e:
        print(f'Could not read that worksheet: {e}')
        return

    columns = set(ws_df.columns)
    if 'Branded_Generic' in columns:
        merge_full_worksheet(ws_df, worksheet_path, staff_name)
    elif {'Product_Name', 'Current_Stock'} <= columns:
        merge_missing_category_supplier(ws_df, worksheet_path, staff_name)
    else:
        print('This file does not look like either the full Item Settings Worksheet')
        print('or an Items_Missing_Category_Supplier.xlsx export.')
        print(f'Columns found: {sorted(columns)}')


def merge_full_worksheet(ws_df, worksheet_path, staff_name):
    required_cols = {'Item_Code', 'Category', 'Supplier', 'Reorder_Point', 'Branded_Generic'}
    missing_cols = required_cols - set(ws_df.columns)
    if missing_cols:
        print(f'This file is missing expected column(s): {sorted(missing_cols)}')
        print('Make sure you selected the worksheet produced by Export_Item_Settings.py.')
        return

    ws_df = ws_df[ws_df['Item_Code'].notna()]
    ws_df['Item_Code'] = ws_df['Item_Code'].astype(str).str.strip()
    ws_df = ws_df[ws_df['Item_Code'] != '']
    ws_df = ws_df.drop_duplicates('Item_Code', keep='last')

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'Cannot open {MASTER_FILE} -- it may be open in Excel. Close it and try again.')
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    for col, default in [('Reorder_Point', 10), ('Supplier', ''), ('Category', ''), ('Is_Generic', '')]:
        if col not in df_master.columns:
            df_master[col] = default
    # Text columns can be all-blank/NaN and get read back as float64 -- force
    # to object dtype before writing strings into them (same fix as
    # Item_Settings_App.py's save_all_to_master()).
    for col in ['Supplier', 'Category', 'Is_Generic']:
        df_master[col] = df_master[col].astype('object')

    master_codes = set(df_master['Item_Code'])
    # Only compare against active (non-discontinued) items, since
    # Export_Item_Settings.py excludes DISCONTINUED items from the worksheet
    # -- those wouldn't be "missing," they were never meant to be there.
    if 'Status' in df_master.columns:
        active_codes = set(df_master.loc[df_master['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED', 'Item_Code'])
    else:
        active_codes = master_codes

    updated_codes = []
    field_change_log = {}
    skipped_no_match = []

    for _, row in ws_df.iterrows():
        code = row['Item_Code']
        if code not in master_codes:
            skipped_no_match.append(code)
            continue

        mask = df_master['Item_Code'] == code
        changes = {}

        if not is_blank(row['Category']):
            changes['Category'] = str(row['Category']).strip()
        if not is_blank(row['Supplier']):
            changes['Supplier'] = str(row['Supplier']).strip()
        if not is_blank(row['Reorder_Point']):
            rp = pd.to_numeric(row['Reorder_Point'], errors='coerce')
            if not pd.isna(rp):
                changes['Reorder_Point'] = rp
        if not is_blank(row['Branded_Generic']):
            bg = str(row['Branded_Generic']).strip().upper()
            changes['Is_Generic'] = '' if bg == 'N/A' else bg

        if not changes:
            continue

        for col, val in changes.items():
            df_master.loc[mask, col] = val

        updated_codes.append(code)
        field_change_log[code] = changes

    missing_from_worksheet = sorted(active_codes - set(ws_df['Item_Code']))

    if not updated_codes:
        print('No matching updates found in the worksheet -- Master was not modified.')
    else:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        shutil.copy2(MASTER_FILE, os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_MERGE_ITEM_SETTINGS_{ts}.xlsx'))

        try:
            with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
                df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
        except PermissionError:
            print(f'Cannot save {MASTER_FILE} -- it may be open in Excel. Close it and try again. No changes were saved.')
            return

        log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
        with open('JDan_System_Logs.txt', 'a') as logbook:
            logbook.write(f'[{log_time}] - ITEM SETTINGS MERGE by {staff_name.upper()} '
                          f'from {os.path.basename(worksheet_path)} ({len(updated_codes)} item/s)\n')
            for code in updated_codes:
                changes = field_change_log[code]
                parts = ', '.join(f'{k}={v}' for k, v in changes.items())
                logbook.write(f'   -> {code}: {parts}\n')

    print('=' * 70)
    print(' ITEM SETTINGS MERGE SUMMARY')
    print('=' * 70)
    print(f'Worksheet: {worksheet_path}')
    print(f'Updated: {len(updated_codes)} item(s)')
    print()

    if skipped_no_match:
        print(f'SKIPPED -- {len(skipped_no_match)} Item_Code(s) in worksheet had no match in Master (check for typos):')
        for code in skipped_no_match:
            print(f'   {code}')
    else:
        print('SKIPPED (no match): none.')
    print()

    if missing_from_worksheet:
        print(f'IN MASTER BUT NOT IN WORKSHEET -- {len(missing_from_worksheet)} active Item_Code(s) '
              f'(may not have been downloaded/edited, or were removed from the worksheet):')
        for code in missing_from_worksheet:
            print(f'   {code}')
    else:
        print('IN MASTER BUT NOT IN WORKSHEET: none -- every active item was covered.')


def merge_missing_category_supplier(ws_df, worksheet_path, staff_name):
    # Unlike merge_full_worksheet, this path touches ONLY Category/Supplier --
    # Reorder_Point/Current_Stock in this file are context columns for the
    # person filling it in, not fields to write back to Master.
    ws_df = ws_df[ws_df['Item_Code'].notna()]
    ws_df['Item_Code'] = ws_df['Item_Code'].astype(str).str.strip()
    ws_df = ws_df[ws_df['Item_Code'] != '']
    ws_df = ws_df.drop_duplicates('Item_Code', keep='last')

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'Cannot open {MASTER_FILE} -- it may be open in Excel. Close it and try again.')
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    for col, default in [('Category', ''), ('Supplier', '')]:
        if col not in df_master.columns:
            df_master[col] = default
    for col in ['Category', 'Supplier']:
        df_master[col] = df_master[col].astype('object')

    master_codes = set(df_master['Item_Code'])

    planned = []  # each: {'code', 'field', 'before', 'after'}
    skipped_no_match = []

    for _, row in ws_df.iterrows():
        code = row['Item_Code']
        if code not in master_codes:
            skipped_no_match.append(code)
            continue

        mask = df_master['Item_Code'] == code
        current_category = df_master.loc[mask, 'Category'].iloc[0]
        current_supplier = df_master.loc[mask, 'Supplier'].iloc[0]
        current_category = '' if pd.isna(current_category) else str(current_category).strip()
        current_supplier = '' if pd.isna(current_supplier) else str(current_supplier).strip()

        if not is_blank(row['Category']):
            new_category = str(row['Category']).strip()
            if new_category != current_category:
                planned.append({'code': code, 'field': 'Category', 'before': current_category, 'after': new_category})
        if not is_blank(row['Supplier']):
            new_supplier = str(row['Supplier']).strip()
            if new_supplier != current_supplier:
                planned.append({'code': code, 'field': 'Supplier', 'before': current_supplier, 'after': new_supplier})

    affected_codes = sorted({p['code'] for p in planned})

    print('=' * 70)
    print(' DRY-RUN PREVIEW -- MISSING CATEGORY/SUPPLIER MERGE')
    print('=' * 70)
    print(f'Worksheet: {worksheet_path}')
    print(f'Items that will change: {len(affected_codes)}')
    print()
    if planned:
        for code in affected_codes:
            item_changes = [p for p in planned if p['code'] == code]
            parts = ', '.join(f"{p['field']}: '{p['before']}' -> '{p['after']}'" for p in item_changes)
            print(f'  {code}: {parts}')
        print()
    else:
        print('Nothing to update -- no Category/Supplier values differ from Master')
        print('(or the worksheet rows were left blank).')

    if skipped_no_match:
        print()
        print(f'SKIPPED -- {len(skipped_no_match)} Item_Code(s) in worksheet had no match in Master (check for typos):')
        for code in skipped_no_match:
            print(f'   {code}')

    if not planned:
        print()
        print('Master was not modified.')
        return

    print()
    confirm = input(f'Apply these {len(affected_codes)} item update(s) to Master? Type YES to apply: ').strip().upper()
    if confirm != 'YES':
        print('Cancelled -- no changes were applied to Master.')
        return

    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_MISSING_CATSUP_MERGE_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)

    for p in planned:
        mask = df_master['Item_Code'] == p['code']
        df_master.loc[mask, p['field']] = p['after']

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'Cannot save {MASTER_FILE} -- it may be open in Excel. Close it and try again. No changes were saved.')
        return

    log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    with open('JDan_System_Logs.txt', 'a') as logbook:
        logbook.write(f'[{log_time}] - MISSING CATEGORY/SUPPLIER MERGE by {staff_name.upper()} '
                      f'from {os.path.basename(worksheet_path)} ({len(affected_codes)} item/s)\n')
        for code in affected_codes:
            item_changes = [p for p in planned if p['code'] == code]
            parts = ', '.join(f"{p['field']}={p['after']}" for p in item_changes)
            logbook.write(f'   -> {code}: {parts}\n')

    print()
    print('=' * 70)
    print(' MISSING CATEGORY/SUPPLIER MERGE -- SUMMARY')
    print('=' * 70)
    print(f'Items updated: {len(affected_codes)}')
    print(f'Timestamp: {log_time}')
    print(f'Backup saved to: {backup_path}')


if __name__ == '__main__':
    main()
