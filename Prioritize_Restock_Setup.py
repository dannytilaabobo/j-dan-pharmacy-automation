import glob
import datetime
import pandas as pd

MASTER_FILE = 'Inventory_Master.xlsx'
EOD_DIR = 'JDan_Sales_System/EOD_Deduction_Reports'
REPORT_DIR = 'JDan_Executive_Reports'

TIER1_CUTOFF = 50   # top N by total qty sold -- customize Reorder_Point/Supplier first
TIER2_CUTOFF = 150  # next N -- do next


def load_sales_velocity():
    frames = []
    for path in glob.glob(f'{EOD_DIR}/PROCESSED_EOD_*.xlsx'):
        df = pd.read_excel(path)
        df['source_file'] = path
        frames.append(df)
    if not frames:
        raise FileNotFoundError('No PROCESSED_EOD_*.xlsx files found -- no sales history to rank by.')
    all_sales = pd.concat(frames, ignore_index=True)
    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    all_sales['Item_Code'] = all_sales['Item Code'].astype(str).str.strip()

    velocity = all_sales.groupby('Item_Code').agg(
        Total_Qty_Sold=('Qty', 'sum'),
        Days_Appeared=('source_file', 'nunique'),
    ).reset_index()
    return velocity


def main():
    velocity = load_sales_velocity()

    df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    df_master['Item_Code'] = df_master['Item_Code'].astype(str).str.strip()
    if 'Status' in df_master.columns:
        df_master = df_master[df_master['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED']
    df_master['Current_Stock'] = pd.to_numeric(df_master['Current_Stock'], errors='coerce').fillna(0)
    df_master['Reorder_Point'] = pd.to_numeric(df_master.get('Reorder_Point', 10), errors='coerce').fillna(10)
    df_master['Supplier'] = df_master['Supplier'].fillna('').astype(str).str.strip() if 'Supplier' in df_master.columns else ''

    # Item_Code is the only stable identifier -- Product_Name can drift
    # (typos/spelling variants) across a code's lot rows, so grouping by name
    # would silently split one product's stock/sales across "different" rows.
    df_master['_name_clean'] = df_master['Product_Name'].astype(str).str.strip()
    name_freq = df_master.groupby(['Item_Code', '_name_clean']).size().reset_index(name='freq')
    name_freq = name_freq.sort_values(['Item_Code', 'freq'], ascending=[True, False])
    representative_name = name_freq.drop_duplicates('Item_Code', keep='first').set_index('Item_Code')['_name_clean']

    agg_master = df_master.groupby('Item_Code').agg(
        Current_Stock=('Current_Stock', 'sum'),
        Reorder_Point=('Reorder_Point', 'first'),
        Supplier=('Supplier', 'first'),
    ).reset_index()
    agg_master['Product_Name'] = agg_master['Item_Code'].map(representative_name)

    merged = agg_master.merge(velocity, on='Item_Code', how='left')
    merged['Total_Qty_Sold'] = merged['Total_Qty_Sold'].fillna(0)
    merged['Days_Appeared'] = merged['Days_Appeared'].fillna(0)
    merged['Supplier_Assigned'] = merged['Supplier'] != ''

    merged = merged.sort_values('Total_Qty_Sold', ascending=False).reset_index(drop=True)
    merged['Rank'] = merged.index + 1

    def tier(rank):
        if rank <= TIER1_CUTOFF:
            return 'TIER 1 - DO FIRST (fast-moving)'
        elif rank <= TIER2_CUTOFF:
            return 'TIER 2 - DO NEXT (moderate movers)'
        return 'TIER 3 - CAN WAIT (slow/no sales)'

    merged['Priority_Tier'] = merged['Rank'].apply(tier)

    print('=' * 70)
    print(' RESTOCK SETUP PRIORITY GUIDE (based on sales velocity)')
    print('=' * 70)
    print(f'Total active products: {len(merged)}')
    print(f'Products with zero recorded sales in available history: {(merged["Total_Qty_Sold"] == 0).sum()}')
    print()

    tier1 = merged[merged['Priority_Tier'].str.startswith('TIER 1')]
    print(f'--- TIER 1: TOP {len(tier1)} FAST-MOVING ITEMS (customize these first) ---')
    show_cols = ['Rank', 'Item_Code', 'Product_Name', 'Total_Qty_Sold', 'Days_Appeared',
                 'Current_Stock', 'Reorder_Point', 'Supplier_Assigned']
    with pd.option_context('display.max_colwidth', 45, 'display.width', 200):
        print(tier1[show_cols].to_string(index=False))

    tier1_no_supplier = tier1[~tier1['Supplier_Assigned']]
    print(f'\n{len(tier1_no_supplier)} of the Tier 1 items still have NO supplier assigned.')

    out_path = f'{REPORT_DIR}/RESTOCK_PRIORITY_GUIDE_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    export_cols = ['Rank', 'Priority_Tier', 'Item_Code', 'Product_Name', 'Total_Qty_Sold',
                   'Days_Appeared', 'Current_Stock', 'Reorder_Point', 'Supplier', 'Supplier_Assigned']
    merged[export_cols].to_excel(out_path, index=False)
    print(f'\nFull ranked list (all {len(merged)} products, all 3 tiers) saved to: {out_path}')


if __name__ == '__main__':
    main()
