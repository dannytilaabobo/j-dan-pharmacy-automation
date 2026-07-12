import os
import glob
import shutil
import datetime
import pandas as pd

MASTER_FILE = 'Inventory_Master.xlsx'
BITPOS_GLOB = 'bitpos_sales_pricing_*.csv'
BACKUP_DIR = 'JDan_Master_Backups'
AUDIT_LOG_PATH = os.path.join('JDan_Executive_Reports', 'PRICE_SYNC_AUDIT_LOG.csv')

# BITPOS is the price authority; this script only ever touches Retail_Price.
# Stock/quantity fields in Master are owned by StockIn/Deduct/PullOut, never by this script.


def latest_bitpos_file():
    matches = glob.glob(BITPOS_GLOB)
    if not matches:
        raise FileNotFoundError(f'No file matching {BITPOS_GLOB} found in this folder.')
    return max(matches, key=os.path.getmtime)


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_PRICE_SYNC_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def append_audit_log(rows, source_file):
    os.makedirs(os.path.dirname(AUDIT_LOG_PATH), exist_ok=True)
    log_df = pd.DataFrame(rows)
    log_df.insert(0, 'Synced_At', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    log_df.insert(1, 'Source_File', source_file)
    header_needed = not os.path.exists(AUDIT_LOG_PATH)
    log_df.to_csv(AUDIT_LOG_PATH, mode='a', header=header_needed, index=False)


def main():
    source_file = latest_bitpos_file()
    print(f'Using price source: {source_file}')

    df_bitpos = pd.read_csv(source_file, dtype=str)
    df_bitpos.columns = [c.strip() for c in df_bitpos.columns]
    df_bitpos['ITEM CODE'] = df_bitpos['ITEM CODE'].astype(str).str.strip()
    df_bitpos['Retail'] = pd.to_numeric(df_bitpos['Retail'], errors='coerce')
    bitpos_prices = dict(zip(df_bitpos['ITEM CODE'], df_bitpos['Retail']))
    bitpos_desc = dict(zip(df_bitpos['ITEM CODE'], df_bitpos['DESCRIPTION']))

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    df_master['Retail_Price'] = pd.to_numeric(df_master['Retail_Price'], errors='coerce')

    changes = []
    for code, new_price in bitpos_prices.items():
        if pd.isna(new_price):
            continue
        mask = df_master['Item_Code'] == code
        if not mask.any():
            continue  # new item, not this script's job -- see Compare_BITPOS_Prices.py
        old_prices = df_master.loc[mask, 'Retail_Price'].unique()
        old_price = old_prices[0] if len(old_prices) else None
        if old_price is None or abs(old_price - new_price) > 0.001 or pd.isna(old_price):
            changes.append({
                'Item_Code': code,
                'Description': bitpos_desc.get(code, ''),
                'Old_Price': old_price,
                'New_Price': new_price,
            })
            df_master.loc[mask, 'Retail_Price'] = new_price

    if not changes:
        print('No price differences found -- Master is already in sync with BITPOS.')
        return

    backup_path = backup_master()
    print(f'Backup created: {backup_path}')

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    append_audit_log(changes, source_file)

    print(f'\n{len(changes)} price(s) synced from {source_file}:')
    for c in changes:
        print(f"  {c['Item_Code']} | {c['Description']}: P{c['Old_Price']} -> P{c['New_Price']}")
    print(f'\nAudit trail appended to: {AUDIT_LOG_PATH}')


if __name__ == '__main__':
    main()
