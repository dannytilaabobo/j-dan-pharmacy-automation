import os
import sys
import shutil
import datetime
import pandas as pd

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'

EXFLEM_ITEM_CODE = '65845431'
EXFLEM_CORRECT_NAME = 'EXFLEM (ACETYLCYSTEINE) 600MG 10S TUBE'


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_EXFLEM_FIX_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def main():
    apply = '--apply' in sys.argv

    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    mask = df['Item_Code'] == EXFLEM_ITEM_CODE

    if not mask.any():
        print(f'WARNING: Item_Code {EXFLEM_ITEM_CODE} not found -- nothing to do.')
        return

    old_names = df.loc[mask, 'Product_Name'].tolist()
    rows_changed = int((df.loc[mask, 'Product_Name'] != EXFLEM_CORRECT_NAME).sum())

    print('=' * 70)
    print(' EXFLEM (Item_Code 65845431) SPELLING CORRECTION' + (' -- APPLY MODE' if apply else ' -- DRY RUN (no changes made)'))
    print('=' * 70)
    print(f'Rows matching Item_Code {EXFLEM_ITEM_CODE}: {mask.sum()}')
    print(f'Rows that will change: {rows_changed}')
    for i, (idx, name) in enumerate(zip(df.loc[mask].index, old_names)):
        arrow = '->' if name != EXFLEM_CORRECT_NAME else '== (already correct)'
        print(f'  Excel row {idx + 2}: "{name}" {arrow} "{EXFLEM_CORRECT_NAME}"')

    if not apply:
        print('\nDRY RUN ONLY -- no backup was made, no file was written.')
        print('Re-run with --apply after confirming this preview to save the change.')
        return

    if rows_changed == 0:
        print('\nNothing to change -- all rows already have the correct spelling.')
        return

    backup_path = backup_master()
    print(f'\nBackup created: {backup_path}')

    df.loc[mask, 'Product_Name'] = EXFLEM_CORRECT_NAME

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{log_time}] - MANUAL NAME CORRECTION (EXFLEM spelling inconsistency)\n")
        logbook.write(f"   -> Item_Code {EXFLEM_ITEM_CODE}: {old_names} -> \"{EXFLEM_CORRECT_NAME}\" (standardized to BITPOS source-of-truth name, {rows_changed} row(s) changed, backup: {backup_path})\n")

    print(f'\nDone. {rows_changed} row(s) corrected and logged to JDan_System_Logs.txt.')


if __name__ == '__main__':
    main()
