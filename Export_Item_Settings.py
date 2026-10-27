import os
import glob
import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

from item_code_resolver import resolve_sales_history

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
EOD_DIR = 'JDan_Sales_System/EOD_Deduction_Reports'
LISTS_SHEET_NAME = 'DropdownLists'
MISSING_OUTPUT_FILE = 'Items_Missing_Category_Supplier.xlsx'

# Same tiering convention as Prioritize_Restock_Setup.py's RESTOCK_PRIORITY_GUIDE.
TIER1_CUTOFF = 50   # top N by total qty sold
TIER2_CUTOFF = 150  # next N

# Same canonical 17-category list as Item_Settings_App.py -- kept as one
# vocabulary across the whole system so the offline worksheet and the
# in-store app never offer conflicting category names for the same item.
# Alphabetized here per Danny's request (values are identical, display
# order is just different from the GUI app's grouped-by-type order).
CATEGORY_LIST = sorted([
    'Antibiotics & Anti-infectives', 'Antidiabetic Medicines', 'Beverages',
    'Cardiovascular Medicines', 'Central Nervous System (CNS) Medicines',
    'Cough, Cold, Allergy & Asthma', 'Eye, Ear & Nasal Preparations',
    "Family Planning & Women's Health", 'Gastrointestinal (GI) Medicines',
    'General Merchandise', 'Hormones & Steroids', 'Medical Supplies & Devices',
    'Others / Uncategorized', 'Pain, Fever & Inflammation',
    'Personal Care, Baby Care & Feminine Care',
    'Topical Preparations (Creams, Ointments, Gels & Sprays)',
    'Vitamins, Minerals & Supplements',
], key=str.upper)

# Same fixed list as Item_Settings_App.py's SUPPLIER_LIST.
SUPPLIER_LIST = [
    'Dyna Drug', 'Comjem', 'Copharm', 'Latriel', 'RGSOI', 'Ice Climber',
    'Metrogen', 'Otsuka', 'Pocari', 'MedSupplies/Shopee', 'Shopee',
    'MULTIPLE/FLEXIBLE', 'Others',
]

BRANDED_GENERIC_OPTIONS = ['GENERIC', 'BRANDED', 'N/A']


def next_available_path(base_path):
    if not os.path.exists(base_path):
        return base_path
    stem, ext = os.path.splitext(base_path)
    n = 2
    while True:
        candidate = f'{stem}_v{n}{ext}'
        if not os.path.exists(candidate):
            return candidate
        n += 1


def load_sales_velocity():
    frames = []
    for path in glob.glob(f'{EOD_DIR}/PROCESSED_EOD_*.xlsx'):
        try:
            df = pd.read_excel(path, dtype={'Item Code': str})
        except Exception:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Total_Qty_Sold'])
    all_sales = pd.concat(frames, ignore_index=True)
    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    # BITPOS codes are translated to the internal Item_Code first (barcode -> old code).
    # Unresolved / corrupted / empty codes are kept as typed, counted and reported -- never dropped.
    all_sales['Item_Code'], _ = resolve_sales_history(all_sales, 'Export Item Settings')
    return all_sales.groupby('Item_Code').agg(Total_Qty_Sold=('Qty', 'sum')).reset_index()


def attach_dropdowns(path, sheet_name, last_row, dropdown_cols):
    # Shared by both export modes. Category's 17-item list joined is ~500
    # chars, over Excel's 255-char inline-list limit, so -- same fix as the
    # "Removed Feature: Data validation" corruption found 2026-07-19/20 in
    # Add_Reorder_And_Supplier_Columns.py -- all dropdowns source from a
    # hidden helper sheet via DefinedName ranges instead of inline formula1
    # strings, for consistency and to avoid that failure mode entirely.
    wb = load_workbook(path)
    ws = wb[sheet_name]

    if LISTS_SHEET_NAME in wb.sheetnames:
        del wb[LISTS_SHEET_NAME]
    lists_ws = wb.create_sheet(LISTS_SHEET_NAME)
    lists_ws.sheet_state = 'hidden'

    for list_col_idx, (col_name, col_letter, options) in enumerate(dropdown_cols, start=1):
        list_col_letter = get_column_letter(list_col_idx)
        for row_idx, value in enumerate(options, start=1):
            lists_ws.cell(row=row_idx, column=list_col_idx, value=value)

        range_name = f'{col_name}DropdownRange'
        range_ref = f"'{LISTS_SHEET_NAME}'!${list_col_letter}$1:${list_col_letter}${len(options)}"
        wb.defined_names[range_name] = DefinedName(range_name, attr_text=range_ref)

        dv = DataValidation(
            type='list',
            formula1=range_name,
            allow_blank=True,
            showErrorMessage=True,
            errorTitle='Invalid entry',
            error=f'Please pick a value from the {col_name} dropdown list (or leave blank).',
        )
        ws.add_data_validation(dv)
        dv.add(f'{col_letter}2:{col_letter}{last_row}')

    wb.save(path)


def export_missing_category_supplier(agg, discontinued_excluded):
    missing = agg[(agg['Category'] == '') | (agg['Supplier'] == '')].copy()
    missing_category_only = int(((missing['Category'] == '') & (missing['Supplier'] != '')).sum())
    missing_supplier_only = int(((missing['Supplier'] == '') & (missing['Category'] != '')).sum())
    missing_both = int(((missing['Category'] == '') & (missing['Supplier'] == '')).sum())

    out_df = missing[['Item_Code', 'Item_Name', 'Category', 'Supplier', 'Reorder_Point', 'Current_Stock']].copy()
    out_df = out_df.rename(columns={'Item_Name': 'Product_Name'})
    out_df = out_df.sort_values('Item_Code').reset_index(drop=True)

    out_path = MISSING_OUTPUT_FILE
    sheet_name = 'Missing Category-Supplier'
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        out_df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        widths = {'A': 14, 'B': 45, 'C': 40, 'D': 22, 'E': 16, 'F': 16}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    last_row = len(out_df) + 1
    if len(out_df) > 0:
        dropdown_cols = [('Category', 'C', CATEGORY_LIST), ('Supplier', 'D', SUPPLIER_LIST)]
        attach_dropdowns(out_path, sheet_name, last_row, dropdown_cols)

    print('=' * 70)
    print(' ITEMS MISSING CATEGORY / SUPPLIER -- EXPORTED')
    print('=' * 70)
    if discontinued_excluded:
        print(f'(Excluded {discontinued_excluded} DISCONTINUED item(s) from the scan.)')
    print(f'Total items flagged (missing Category and/or Supplier): {len(out_df)}')
    print(f'  Missing Category only: {missing_category_only}')
    print(f'  Missing Supplier only: {missing_supplier_only}')
    print(f'  Missing BOTH Category and Supplier: {missing_both}')
    print()
    if len(out_df) == 0:
        print('Nothing flagged -- every active item already has Category and Supplier filled in.')
    else:
        print(f'Saved to: {os.path.abspath(out_path)}')
        print()
        print('Fill in Category / Supplier for the flagged items, then run')
        print('Merge_Item_Settings.py and select this file -- it will update')
        print('ONLY Category and Supplier in Master, leaving every other column untouched.')


def main():
    print('=' * 70)
    print(' ITEM SETTINGS EXPORT')
    print('=' * 70)
    print('Select export mode:')
    print('  1. Full Item Settings Worksheet (all active items)')
    print('  2. Items Missing Category and/or Supplier only')
    mode_choice = input('Enter 1 or 2 (default 1): ').strip()
    mode = '2' if mode_choice == '2' else '1'
    print()

    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype={'Item_Code': str})
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()

    if 'Status' in df.columns:
        before = df['Item_Code'].nunique()
        df = df[df['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED']
        after = df['Item_Code'].nunique()
        discontinued_excluded = before - after
    else:
        discontinued_excluded = 0

    for col, default in [('Reorder_Point', 10), ('Supplier', ''), ('Category', ''), ('Is_Generic', ''), ('Current_Stock', 0)]:
        if col not in df.columns:
            df[col] = default
    df['Reorder_Point'] = pd.to_numeric(df['Reorder_Point'], errors='coerce').fillna(10)
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
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
        Current_Stock=('Current_Stock', 'sum'),
    ).reset_index()
    agg['Item_Name'] = agg['Item_Code'].map(representative_name)

    if mode == '2':
        export_missing_category_supplier(agg, discontinued_excluded)
        return

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

    base_path = f"Item_Settings_Worksheet_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    out_path = next_available_path(base_path)
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        out_df.to_excel(writer, sheet_name='Item Settings', index=False)
        ws = writer.sheets['Item Settings']
        # Cells are editable by default (no protect_sheet call) -- just widen
        # columns for usability while editing at home.
        widths = {'A': 14, 'B': 45, 'C': 40, 'D': 22, 'E': 16, 'F': 16}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    # Reorder_Point stays manual/numeric (no dropdown, per Danny's request).
    last_row = len(out_df) + 1
    dropdown_cols = [('Category', 'C', CATEGORY_LIST), ('Supplier', 'D', SUPPLIER_LIST),
                      ('Branded_Generic', 'F', BRANDED_GENERIC_OPTIONS)]
    attach_dropdowns(out_path, 'Item Settings', last_row, dropdown_cols)

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
    existing_categories = sorted(set(agg['Category']) - {''})
    existing_suppliers = sorted(set(agg['Supplier']) - {''})
    print(f'Existing unique Category values found in Master: {len(existing_categories)} '
          f'(all already covered by the {len(CATEGORY_LIST)}-item canonical dropdown list)')
    print(f'Existing unique Supplier values found in Master: {len(existing_suppliers)} '
          f'(all already covered by the {len(SUPPLIER_LIST)}-item fixed dropdown list)')
    print()
    print(f'Dropdowns attached: Category ({len(CATEGORY_LIST)} options, alphabetized), '
          f'Supplier ({len(SUPPLIER_LIST)} options), Branded_Generic (GENERIC/BRANDED/N-A).')
    print('Reorder_Point stays manual/numeric -- no dropdown.')
    print('All cells are unlocked/editable. Bring this file home, fill in')
    print('Category / Supplier / Reorder_Point / Branded_Generic, then run')
    print('Merge_Item_Settings.py when you get back to merge it into Master.')


if __name__ == '__main__':
    main()
