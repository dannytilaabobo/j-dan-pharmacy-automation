import os
import glob
import shutil
import datetime
import pandas as pd

from name_cleaning import clean_text

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

# If more than this many items would be newly marked DISCONTINUED in one run,
# show an extra warning -- it likely means the BITPOS export is incomplete
# rather than that many items actually got discontinued at once.
MASS_DISCONTINUE_WARNING_THRESHOLD = 20


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


def load_pending():
    if os.path.exists(PENDING_PATH):
        df = pd.read_csv(PENDING_PATH, dtype=str)
        return df
    return pd.DataFrame(columns=PENDING_COLUMNS)


def save_pending(df):
    os.makedirs(REPORT_DIR, exist_ok=True)
    df.to_csv(PENDING_PATH, index=False)


def detect_changes(df_master, df_bitpos):
    bitpos_codes = set(df_bitpos['ITEM CODE'])
    bitpos_price = dict(zip(df_bitpos['ITEM CODE'], df_bitpos['Retail']))
    bitpos_name = dict(zip(df_bitpos['ITEM CODE'], df_bitpos['DESCRIPTION']))

    price_changes = []
    name_changes = []
    discontinued = []

    active_mask = df_master['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED' if 'Status' in df_master.columns else pd.Series(True, index=df_master.index)
    master_codes = df_master.loc[active_mask, 'Item_Code'].unique()

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

    return price_changes, name_changes, discontinued


def print_preview(price_changes, name_changes, discontinued):
    print('=' * 70)
    print(' PREVIEW -- WALA PANG NA-A-APPLY NA PAGBABAGO')
    print('=' * 70)

    print(f'\n--- (1) PRICE CHANGES: {len(price_changes)} item(s) ---')
    if price_changes:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(price_changes)[['Item_Code', 'Description', 'Old_Value', 'New_Value']].to_string(index=False))
    else:
        print('  Wala.')

    print(f'\n--- (2) PANGALAN NA NAGBAGO: {len(name_changes)} item(s) (isa-isang tatanungin sa ibaba) ---')
    if name_changes:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(name_changes)[['Item_Code', 'Old_Value', 'New_Value']].rename(
                columns={'Old_Value': 'Old_Name', 'New_Value': 'New_Name'}).to_string(index=False))
    else:
        print('  Wala.')

    print(f'\n--- (3) BAGONG DISCONTINUED (wala na sa BITPOS export): {len(discontinued)} item(s) ---')
    if discontinued:
        with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
            print(pd.DataFrame(discontinued)[['Item_Code', 'Description']].to_string(index=False))
        if len(discontinued) > MASS_DISCONTINUE_WARNING_THRESHOLD:
            print(f'\n  ⚠️  WARNING: {len(discontinued)} items ang mawawalan ng match sa BITPOS export -- '
                  f'hindi pangkaraniwan ito. I-verify muna kung kumpleto ang CSV file bago mag-confirm, '
                  f'baka partial/truncated na export lang ito, hindi totoong mass discontinuation.')
    else:
        print('  Wala.')

    print('\nCurrent_Stock ay HINDI kailanman ginagalaw ng script na ito.')


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
        answer = input("  I-apply na ba ang rename na ito ngayon? (yes = apply / no = i-defer sa Pending Verification): ").strip().lower()
        if answer == 'yes':
            confirmed.append(c)
        else:
            newly_deferred.append(c)

    if already_pending:
        print(f"\n{len(already_pending)} name change(s) ay AUTO-SKIPPED (naka-pending na verification mula sa nakaraang run):")
        for c in already_pending:
            print(f"  {c['Item_Code']}: '{c['Old_Value']}' -> '{c['New_Value']}'")

    return confirmed, newly_deferred, already_pending


def apply_changes(df_master, price_changes, name_changes, discontinued):
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
        df_master.loc[mask, 'Discontinued_Reason'] = 'Wala na sa BITPOS export -- awtomatikong na-detect ng Sync_BITPOS_Prices.py'

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
            'Note': 'Wala pang stock/verification opportunity -- i-verify sa susunod na stock-in.',
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


def save_preview_report(price_changes, name_changes, discontinued):
    os.makedirs(REPORT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_DIR, f'BITPOS_SYNC_PREVIEW_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        pd.DataFrame(price_changes).to_excel(writer, sheet_name='Price_Changes', index=False)
        pd.DataFrame(name_changes).to_excel(writer, sheet_name='Name_Changes', index=False)
        pd.DataFrame(discontinued).to_excel(writer, sheet_name='Newly_Discontinued', index=False)
    return out_path


def main():
    source_file = latest_bitpos_file()
    print(f'Using BITPOS source: {source_file}')

    df_bitpos = load_bitpos(source_file)

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    df_master['Retail_Price'] = pd.to_numeric(df_master['Retail_Price'], errors='coerce')

    price_changes, name_changes, discontinued = detect_changes(df_master, df_bitpos)

    if not price_changes and not name_changes and not discontinued:
        print('\nWalang pagkakaiba -- Master ay tugma na sa BITPOS export.')
        return

    print_preview(price_changes, name_changes, discontinued)
    preview_path = save_preview_report(price_changes, name_changes, discontinued)
    print(f'\nBuong preview na-save din sa: {preview_path}')

    pending_df = load_pending()
    confirmed_names, newly_deferred, already_pending = resolve_name_changes(name_changes, pending_df)

    print('\n--- FINAL SUMMARY BAGO I-APPLY ---')
    print(f'Price changes: {len(price_changes)}')
    print(f'Name changes na i-a-apply ngayon: {len(confirmed_names)}')
    print(f'Name changes na i-de-defer sa Pending Verification: {len(newly_deferred)}')
    print(f'Name changes na auto-skipped (dati nang pending): {len(already_pending)}')
    print(f'Bagong discontinued: {len(discontinued)}')

    answer = input('\nItuloy ba at i-apply ang mga NAKUMPIRMANG pagbabago sa Master (price, approved name changes, discontinued)? (yes/no): ').strip().lower()
    if answer != 'yes':
        print('\nKINANSELA. Walang na-apply na pagbabago sa Master.')
        if newly_deferred:
            print('Note: hindi rin na-save sa Pending Verification list ang mga i-de-defer -- i-rerun ang script kung gusto mo pa ring i-record ang mga ito.')
        return

    backup_path = backup_master()
    print(f'Backup created: {backup_path}')

    df_master = apply_changes(df_master, price_changes, confirmed_names, discontinued)

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    resolved_codes = [c['Item_Code'] for c in confirmed_names]
    update_pending_file(pending_df, newly_deferred, resolved_codes)

    all_changes = price_changes + confirmed_names + discontinued
    if newly_deferred:
        all_changes += [{**c, 'Change_Type': 'NAME_DEFERRED'} for c in newly_deferred]
    append_audit_log(all_changes, source_file)

    print(f'\nNA-APPLY: {len(price_changes)} price change/s, {len(confirmed_names)} name change/s, '
          f'{len(discontinued)} bagong discontinued item/s.')
    if newly_deferred:
        print(f'{len(newly_deferred)} name change/s na-save sa Pending Verification list: {PENDING_PATH}')
    print(f'Audit trail appended to: {AUDIT_LOG_PATH}')
    print('Current_Stock ay hindi ginalaw.')


if __name__ == '__main__':
    main()
