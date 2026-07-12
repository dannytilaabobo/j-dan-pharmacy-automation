import os
import shutil
import datetime
import pandas as pd

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'

# Item_Code -> reason. Keeps the record (and any remaining stock) but flags it
# as discontinued so dashboards/audits can filter it out of "active catalog" views.
DISCONTINUED_ITEMS = {
    '37776367': 'Repacked by supplier as SILVER SULFADIAZINE 10G (code 26246000). '
                'Sell through or return remaining stock.',
    '77925626': 'Repacked by supplier as FUSIDIC ACID TOPICAL CRM 10G (code 10320655). '
                'Sell through or return remaining stock.',
}


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_DISCONTINUE_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def main():
    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    if 'Status' not in df.columns:
        df['Status'] = 'ACTIVE'
    if 'Discontinued_Date' not in df.columns:
        df['Discontinued_Date'] = ''
    if 'Discontinued_Reason' not in df.columns:
        df['Discontinued_Reason'] = ''
    df['Status'] = df['Status'].fillna('ACTIVE').replace('', 'ACTIVE')

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    today = datetime.datetime.now().strftime('%Y-%m-%d')

    updated = []
    for code, reason in DISCONTINUED_ITEMS.items():
        mask = df['Item_Code'] == code
        if not mask.any():
            print(f'WARNING: Item_Code {code} not found in Master -- skipped.')
            continue
        df.loc[mask, 'Status'] = 'DISCONTINUED'
        df.loc[mask, 'Discontinued_Date'] = today
        df.loc[mask, 'Discontinued_Reason'] = reason
        updated.append((code, df.loc[mask, 'Product_Name'].iloc[0], df.loc[mask, 'Current_Stock'].iloc[0]))

    if not updated:
        print('No matching items found. Nothing changed.')
        return

    backup_path = backup_master()
    print(f'Backup created: {backup_path}')

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    print(f'\n{len(updated)} item(s) marked DISCONTINUED:')
    for code, name, stock in updated:
        print(f'  {code} | {name} | remaining stock: {stock}')


if __name__ == '__main__':
    main()
