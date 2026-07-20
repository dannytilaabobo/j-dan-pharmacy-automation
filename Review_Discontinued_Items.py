import os
import shutil
import datetime
import pandas as pd

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
REPORT_DIR = 'JDan_Executive_Reports'
REMOVAL_AUDIT_LOG = os.path.join(REPORT_DIR, 'DISCONTINUED_REMOVAL_AUDIT_LOG.csv')


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_DISCONTINUED_REMOVAL_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def append_removal_log(removed_rows):
    os.makedirs(REPORT_DIR, exist_ok=True)
    log_df = removed_rows[['Item_Code', 'Product_Name', 'Current_Stock', 'Discontinued_Date', 'Discontinued_Reason']].copy()
    log_df.insert(0, 'Removed_At', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    log_df['Action'] = 'PERMANENTLY REMOVED FROM MASTER'
    header_needed = not os.path.exists(REMOVAL_AUDIT_LOG)
    log_df.to_csv(REMOVAL_AUDIT_LOG, mode='a', header=header_needed, index=False)


def print_section(title, rows_df, show_stock):
    print(f'\n=== {title} ({len(rows_df)}) ===')
    if rows_df.empty:
        print('  (none)')
        return
    for _, row in rows_df.iterrows():
        stock_part = f" | Current_Stock: {int(row['Current_Stock'])}" if show_stock else ''
        print(f"  {row['Item_Code']} | {row['Product_Name']}{stock_part}")
        reason = row.get('Discontinued_Reason', '')
        if reason and str(reason).strip():
            print(f'      Reason: {reason}')


def main():
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    if 'Status' not in df.columns:
        print('No Status column found in Master -- no items have been marked DISCONTINUED yet.')
        return

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)

    discontinued = df[df['Status'].astype(str).str.strip().str.upper() == 'DISCONTINUED'].copy()

    if discontinued.empty:
        print('No DISCONTINUED items found in Master. Nothing to review.')
        return

    safe_to_remove = discontinued[discontinued['Current_Stock'] == 0].copy()
    has_stock = discontinued[discontinued['Current_Stock'] > 0].copy()

    print(f'Total DISCONTINUED items in Master: {len(discontinued)}')
    print_section('SAFE TO REMOVE -- Current_Stock = 0', safe_to_remove, show_stock=False)
    print_section('HAS REMAINING STOCK -- sell through or return to supplier first', has_stock, show_stock=True)

    if safe_to_remove.empty:
        print('\nNo items are currently safe to remove (all DISCONTINUED items still have stock).')
        return

    print('\n--- Remove from Master? ---')
    print('  [1] Review and remove one by one (confirm each)')
    print('  [2] Remove ALL safe-to-remove items at once (single confirmation)')
    print('  [3] Exit without removing anything')
    choice = input('Choose 1, 2, or 3: ').strip()

    to_remove_codes = []

    if choice == '1':
        for _, row in safe_to_remove.iterrows():
            answer = input(
                f"\nRemove {row['Item_Code']} | {row['Product_Name']} permanently from Master? (yes/no): "
            ).strip().lower()
            if answer == 'yes':
                to_remove_codes.append(row['Item_Code'])
            else:
                print('  Skipped.')
    elif choice == '2':
        print(f'\nAbout to PERMANENTLY remove {len(safe_to_remove)} item(s) from Master:')
        for _, row in safe_to_remove.iterrows():
            print(f"  {row['Item_Code']} | {row['Product_Name']}")
        answer = input('\nType "yes" to confirm this batch removal (cannot be undone except via backup): ').strip().lower()
        if answer == 'yes':
            to_remove_codes = safe_to_remove['Item_Code'].tolist()
    else:
        print('Exiting -- no changes made.')
        return

    if not to_remove_codes:
        print('\nNo items confirmed for removal. Nothing changed.')
        return

    removed_rows = df[df['Item_Code'].isin(to_remove_codes)].copy()

    backup_path = backup_master()
    print(f'\nBackup created: {backup_path}')

    df = df[~df['Item_Code'].isin(to_remove_codes)].copy()

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    append_removal_log(removed_rows)

    print(f'\n{len(to_remove_codes)} item(s) permanently removed from Master:')
    for _, row in removed_rows.iterrows():
        print(f"  {row['Item_Code']} | {row['Product_Name']}")
    print(f'Removal logged to: {REMOVAL_AUDIT_LOG}')
    print('NOTE: this rewrite may have cleared the Category/Supplier/Is_Generic Excel dropdowns '
          '(a known side effect of any full-sheet rewrite) -- re-run "Add_Reorder_And_Supplier_Columns.py" '
          'if you need them reattached.')


if __name__ == '__main__':
    main()
