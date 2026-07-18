import os
import shutil
import datetime
import pandas as pd

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
ITEM_CODE = '99795658'  # BAND AID PLASTER (PC)
ACTUAL_TOTAL = 187      # confirmed by physical count, 2026-07-18
REASON = "Stock-in encoding error, corrected after physical count verification"


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_BANDAID_CORRECTION_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def main():
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)

    mask = df['Item_Code'] == ITEM_CODE
    if not mask.any():
        print(f'CRITICAL: Item_Code {ITEM_CODE} not found in Master.')
        return

    old_rows = df.loc[mask, ['Lot_Number', 'Expiry_Date', 'Current_Stock']].copy()
    old_total = old_rows['Current_Stock'].sum()
    delta = ACTUAL_TOTAL - old_total

    # Same FIFO/FEFO convention as the SOD reconciliation fix: earliest-expiry
    # lot absorbs the shortage first.
    parsed_dates = pd.to_datetime(df.loc[mask, 'Expiry_Date'], errors='coerce').fillna(pd.Timestamp('2099-12-31'))
    ordered_idx = parsed_dates.sort_values().index.tolist()

    qty_to_remove = abs(delta)
    for idx in ordered_idx:
        if qty_to_remove <= 0:
            break
        current = df.at[idx, 'Current_Stock']
        if current >= qty_to_remove:
            df.at[idx, 'Current_Stock'] = current - qty_to_remove
            qty_to_remove = 0
        else:
            qty_to_remove -= current
            df.at[idx, 'Current_Stock'] = 0

    new_rows = df.loc[mask, ['Lot_Number', 'Expiry_Date', 'Current_Stock']].copy()
    new_total = df.loc[mask, 'Current_Stock'].sum()

    print('BEFORE:')
    print(old_rows.to_string(index=False))
    print(f'Old total: {old_total}')
    print()
    print('AFTER:')
    print(new_rows.to_string(index=False))
    print(f'New total: {new_total}')

    if abs(new_total - ACTUAL_TOTAL) > 0.01:
        print(f'\nCRITICAL: New total ({new_total}) does not match target ({ACTUAL_TOTAL}) -- aborting, no changes saved.')
        return

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
        logbook.write(f"[{log_time}] - MANUAL STOCK CORRECTION: BAND AID PLASTER (PC) [{ITEM_CODE}]\n")
        logbook.write(f"   -> Old Total: {old_total} | New Total: {new_total}\n")
        for (_, old_r), (_, new_r) in zip(old_rows.iterrows(), new_rows.iterrows()):
            logbook.write(f"   -> Lot {old_r['Lot_Number']} ({old_r['Expiry_Date']}): {old_r['Current_Stock']} -> {new_r['Current_Stock']}\n")
        logbook.write(f"   -> Reason: {REASON}\n")

    print(f'\nSaved and logged to JDan_System_Logs.txt.')


if __name__ == '__main__':
    main()
