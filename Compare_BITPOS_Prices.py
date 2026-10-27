import os
import glob
import datetime
import pandas as pd

from item_code_resolver import resolve_series, announce

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BITPOS_GLOB = 'bitpos_sales_pricing_*.csv'


def load_bitpos():
    matches = sorted(glob.glob(BITPOS_GLOB))
    if not matches:
        raise FileNotFoundError(f'No file matching {BITPOS_GLOB} found in this folder.')
    path = matches[-1]
    df = pd.read_csv(path, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    df['ITEM CODE'] = df['ITEM CODE'].astype(str).str.strip()
    # barcode -> internal (old) Item_Code; anything unresolved stays as typed and is reported
    df['ITEM CODE'], _code_report = resolve_series(df['ITEM CODE'])
    announce(_code_report, 'Compare BITPOS Prices')
    df['Retail'] = pd.to_numeric(df['Retail'], errors='coerce')
    return path, df


def load_master():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype={'Item_Code': str})
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Retail_Price'] = pd.to_numeric(df['Retail_Price'], errors='coerce')
    return df


def main():
    bitpos_path, df_bitpos = load_bitpos()
    df_master = load_master()

    print('=' * 70)
    print(' BITPOS PRICE LIST vs INVENTORY_MASTER COMPARISON')
    print('=' * 70)
    print(f'BITPOS file: {bitpos_path} ({len(df_bitpos)} items)')
    print(f'Master file: {MASTER_FILE} ({df_master["Item_Code"].nunique()} unique item codes, {len(df_master)} rows)')

    # Master can carry multiple rows per Item_Code (one per lot). If those rows
    # disagree on Retail_Price, that is itself a data-integrity issue worth
    # surfacing before comparing against BITPOS.
    price_groups = df_master.groupby('Item_Code')['Retail_Price'].agg(['nunique', 'first', 'max', 'min'])
    internal_conflicts = price_groups[price_groups['nunique'] > 1]

    master_prices = df_master.groupby('Item_Code').agg(
        Product_Name=('Product_Name', 'first'),
        Master_Retail=('Retail_Price', 'first'),
    ).reset_index()

    merged = df_bitpos.merge(master_prices, left_on='ITEM CODE', right_on='Item_Code', how='outer', indicator=True)

    # (a) Price mismatches
    both = merged[merged['_merge'] == 'both'].copy()
    both['price_diff'] = both['Retail'] - both['Master_Retail']
    mismatches = both[both['price_diff'].abs() > 0.001].copy()
    mismatches = mismatches.sort_values('price_diff', key=lambda s: s.abs(), ascending=False)

    # (b) New items in BITPOS, not in Master
    new_items = merged[merged['_merge'] == 'left_only'][['ITEM CODE', 'DESCRIPTION', 'Retail']].copy()

    # (c) Items in Master, not in current BITPOS export
    delisted = merged[merged['_merge'] == 'right_only'][['Item_Code', 'Product_Name', 'Master_Retail']].copy()

    print('\n--- (0) MASTER INTERNAL PRICE CONFLICTS (same Item_Code, different Retail_Price across lots) ---')
    if internal_conflicts.empty:
        print('  None found.')
    else:
        print(internal_conflicts.to_string())

    print(f'\n--- (a) PRICE MISMATCHES: {len(mismatches)} items ---')
    if mismatches.empty:
        print('  None found -- all matched prices agree.')
    else:
        show = mismatches[['ITEM CODE', 'DESCRIPTION', 'Master_Retail', 'Retail', 'price_diff']].copy()
        show.columns = ['Item_Code', 'Description', 'Master Price', 'BITPOS Price', 'Difference']
        with pd.option_context('display.max_colwidth', 50, 'display.width', 200):
            print(show.to_string(index=False))

    print(f'\n--- (b) NEW ITEMS IN BITPOS, NOT YET IN MASTER: {len(new_items)} items ---')
    if new_items.empty:
        print('  None found.')
    else:
        with pd.option_context('display.max_colwidth', 50, 'display.width', 200):
            print(new_items.to_string(index=False))

    print(f'\n--- (c) ITEMS IN MASTER, NOT IN CURRENT BITPOS LIST: {len(delisted)} items ---')
    if delisted.empty:
        print('  None found.')
    else:
        with pd.option_context('display.max_colwidth', 50, 'display.width', 200):
            print(delisted.to_string(index=False))

    out_path = f'JDan_Executive_Reports/BITPOS_PRICE_COMPARISON_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        mismatches[['ITEM CODE', 'DESCRIPTION', 'Master_Retail', 'Retail', 'price_diff']].to_excel(
            writer, sheet_name='Price_Mismatches', index=False)
        new_items.to_excel(writer, sheet_name='New_Items_Not_In_Master', index=False)
        delisted.to_excel(writer, sheet_name='In_Master_Not_In_BITPOS', index=False)
        if not internal_conflicts.empty:
            internal_conflicts.to_excel(writer, sheet_name='Master_Internal_Conflicts')
    print(f'\nFull report saved to: {out_path}')


if __name__ == '__main__':
    main()
