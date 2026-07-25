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
        ws_df = pd.read_excel(worksheet_path, sheet_name='Item Settings', keep_default_na=False, na_values=[''])
    except Exception as e:
        print(f'Could not read that worksheet: {e}')
        return

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


if __name__ == '__main__':
    main()
