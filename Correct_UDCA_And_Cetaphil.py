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
PENDING_PATH = os.path.join('JDan_Executive_Reports', 'BITPOS_PENDING_NAME_VERIFICATION.csv')

UDCA_ITEM_CODE = '33053837'
UDCA_CORRECT_NAME = 'URSODEOXYCHOLIC ACID (UDCA) 300MG'

CETAPHIL_ITEM_CODE = '51881290'
CETAPHIL_NEW_NAME = 'CETAPHIL BABY GENTLE WASH SHAMPOO 230ML'


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_UDCA_CETAPHIL_FIX_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def main():
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()

    # --- Correction 1: standardize UDCA/UCDA typo ---
    udca_mask = df['Item_Code'] == UDCA_ITEM_CODE
    if not udca_mask.any():
        print(f'WARNING: Item_Code {UDCA_ITEM_CODE} not found -- skipping UDCA correction.')
    else:
        old_names_udca = df.loc[udca_mask, 'Product_Name'].tolist()
        df.loc[udca_mask, 'Product_Name'] = UDCA_CORRECT_NAME
        print(f'UDCA fix: {old_names_udca} -> all set to "{UDCA_CORRECT_NAME}" ({udca_mask.sum()} row(s))')

    # --- Correction 2: apply pending CETAPHIL rename ---
    cetaphil_mask = df['Item_Code'] == CETAPHIL_ITEM_CODE
    if not cetaphil_mask.any():
        print(f'WARNING: Item_Code {CETAPHIL_ITEM_CODE} not found -- skipping CETAPHIL rename.')
        old_name_cetaphil = None
    else:
        old_name_cetaphil = df.loc[cetaphil_mask, 'Product_Name'].iloc[0]
        df.loc[cetaphil_mask, 'Product_Name'] = CETAPHIL_NEW_NAME
        print(f'CETAPHIL fix: "{old_name_cetaphil}" -> "{CETAPHIL_NEW_NAME}" ({cetaphil_mask.sum()} row(s))')

    backup_path = backup_master()
    print(f'\nBackup created: {backup_path}')

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{log_time}] - MANUAL NAME CORRECTIONS (UDCA typo + CETAPHIL rename)\n")
        if udca_mask.any():
            logbook.write(f"   -> Item_Code {UDCA_ITEM_CODE}: {old_names_udca} -> \"{UDCA_CORRECT_NAME}\" (typo standardization, {udca_mask.sum()} row(s))\n")
        if cetaphil_mask.any():
            logbook.write(f"   -> Item_Code {CETAPHIL_ITEM_CODE}: \"{old_name_cetaphil}\" -> \"{CETAPHIL_NEW_NAME}\" (pending BITPOS rename resolved, confirmed by distinct pricing vs Item_Code 24681902)\n")

    # --- Remove resolved entry from Pending Verification list ---
    if os.path.exists(PENDING_PATH):
        pending_df = pd.read_csv(PENDING_PATH, dtype=str)
        before_count = len(pending_df)
        pending_df = pending_df[pending_df['Item_Code'] != CETAPHIL_ITEM_CODE]
        removed = before_count - len(pending_df)
        pending_df.to_csv(PENDING_PATH, index=False)
        print(f'\nRemoved {removed} entry/entries for Item_Code {CETAPHIL_ITEM_CODE} from Pending Verification list.')
    else:
        print('\nNo Pending Verification file found -- nothing to remove.')

    print('\nDone. Both corrections applied and logged.')


if __name__ == '__main__':
    main()
