import os
import glob
import datetime
import pandas as pd

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
EOD_DIR = 'JDan_Sales_System/EOD_Deduction_Reports'

# Same tiering convention as Prioritize_Restock_Setup.py's RESTOCK_PRIORITY_GUIDE.
TIER1_CUTOFF = 50   # top N by total qty sold
TIER2_CUTOFF = 150  # next N


def load_sales_velocity():
    frames = []
    for path in glob.glob(f'{EOD_DIR}/PROCESSED_EOD_*.xlsx'):
        try:
            df = pd.read_excel(path)
        except Exception:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Total_Qty_Sold'])
    all_sales = pd.concat(frames, ignore_index=True)
    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    all_sales['Item_Code'] = all_sales['Item Code'].astype(str).str.strip()
    return all_sales.groupby('Item_Code').agg(Total_Qty_Sold=('Qty', 'sum')).reset_index()


def main():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()

    if 'Status' in df.columns:
        before = df['Item_Code'].nunique()
        df = df[df['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED']
        after = df['Item_Code'].nunique()
        discontinued_excluded = before - after
    else:
        discontinued_excluded = 0

    for col, default in [('Reorder_Point', 10), ('Supplier', ''), ('Category', ''), ('Is_Generic', '')]:
        if col not in df.columns:
            df[col] = default
    df['Reorder_Point'] = pd.to_numeric(df['Reorder_Point'], errors='coerce').fillna(10)
    df['Supplier'] = df['Supplier'].fillna('').astype(str).str.strip()
    df['Category'] = df['Category'].fillna('').astype(str).str.strip()
    df['Is_Generic'] = df['Is_Generic'].fillna('').astype(str).str.strip().str.upper()

    # Item_Code is the only stable identifier -- Product_Name can drift across
    # a code's lot rows, so grouping by name would silently split one product
    # into several. Pick the most frequent spelling per code as the display name.
    df['_name_clean'] = df['Product_Name'].astype(str).str.strip()
    name_freq = df.groupby(['Item_Code', '_name_clean']).size().reset_index(name='freq')
    name_freq = name_freq.sort_values(['Item_Code', 'freq'], ascending=[True, False])
    representative_name = name_freq.drop_duplicates('Item_Code', keep='first').set_index('Item_Code')['_name_clean']

    agg = df.groupby('Item_Code').agg(
        Reorder_Point=('Reorder_Point', 'first'),
        Supplier=('Supplier', 'first'),
        Category=('Category', 'first'),
        Is_Generic=('Is_Generic', 'first'),
    ).reset_index()
    agg['Item_Name'] = agg['Item_Code'].map(representative_name)

    velocity = load_sales_velocity()
    agg = agg.merge(velocity, on='Item_Code', how='left')
    agg['Total_Qty_Sold'] = agg['Total_Qty_Sold'].fillna(0)
    agg = agg.sort_values('Total_Qty_Sold', ascending=False).reset_index(drop=True)
    agg['Rank'] = agg.index + 1

    def tier(rank):
        if rank <= TIER1_CUTOFF:
            return 1
        elif rank <= TIER2_CUTOFF:
            return 2
        return 3

    agg['_tier'] = agg['Rank'].apply(tier)

    # Is_Generic blank means "not applicable" (non-medicine item) in Master --
    # spell that out as N/A in the worksheet so it's unambiguous while editing.
    agg['Branded_Generic'] = agg['Is_Generic'].replace('', 'N/A')

    out_df = agg[['Item_Code', 'Item_Name', 'Category', 'Supplier', 'Reorder_Point', 'Branded_Generic']].copy()

    out_path = f"Item_Settings_Worksheet_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        out_df.to_excel(writer, sheet_name='Item Settings', index=False)
        ws = writer.sheets['Item Settings']
        # Cells are editable by default (no protect_sheet call) -- just widen
        # columns for usability while editing at home.
        widths = {'A': 14, 'B': 45, 'C': 40, 'D': 22, 'E': 16, 'F': 16}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    tier1_count = (agg['_tier'] == 1).sum()
    tier2_count = (agg['_tier'] == 2).sum()
    tier3_count = (agg['_tier'] == 3).sum()

    print('=' * 70)
    print(' ITEM SETTINGS WORKSHEET EXPORTED')
    print('=' * 70)
    print(f'Total active items exported: {len(out_df)}')
    if discontinued_excluded:
        print(f'(Excluded {discontinued_excluded} DISCONTINUED item(s) -- not part of active settings work.)')
    print(f'  Tier 1 (fast-movers, top {TIER1_CUTOFF}): {tier1_count} -- listed first')
    print(f'  Tier 2 (moderate movers, next {TIER2_CUTOFF - TIER1_CUTOFF}): {tier2_count}')
    print(f'  Tier 3 (slow/no recorded sales): {tier3_count}')
    print()
    print(f'Saved to: {os.path.abspath(out_path)}')
    print()
    print('All cells are unlocked/editable. Bring this file home, fill in')
    print('Category / Supplier / Reorder_Point / Branded_Generic, then run')
    print('Merge_Item_Settings.py when you get back to merge it into Master.')


if __name__ == '__main__':
    main()
