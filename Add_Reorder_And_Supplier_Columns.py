import os
import shutil
import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'

DEFAULT_REORDER_POINT = 10

SUPPLIER_LIST = [
    'Dyna Drug', 'Comjem', 'Copharm', 'Latriel', 'RGSOI', 'Ice Climber',
    'Metrogen', 'Otsuka', 'Pocari', 'MedSupplies/Shopee', 'Shopee',
    'MULTIPLE/FLEXIBLE', 'Others',
]

# 17-category list, replaced 2026-07-19 (0/1,610 items had Category filled
# at the time, so this is a clean replacement -- no migration of old values).
CATEGORY_LIST = [
    # Medicines
    'Vitamins, Minerals & Supplements', 'Pain, Fever & Inflammation',
    'Cough, Cold, Allergy & Asthma', 'Gastrointestinal (GI) Medicines',
    'Antibiotics & Anti-infectives', 'Cardiovascular Medicines',
    'Antidiabetic Medicines', 'Hormones & Steroids',
    'Central Nervous System (CNS) Medicines',
    'Topical Preparations (Creams, Ointments, Gels & Sprays)',
    'Eye, Ear & Nasal Preparations', "Family Planning & Women's Health",
    # Health & Wellness
    'Medical Supplies & Devices', 'Personal Care, Baby Care & Feminine Care',
    'Beverages',
    # Retail
    'General Merchandise', 'Others / Uncategorized',
]

# Medicine-only attribute: GENERIC, BRANDED, or blank/N-A for non-medicine
# items (General Merchandise, Personal Care, etc.) -- never forced.
IS_GENERIC_OPTIONS = ['', 'GENERIC', 'BRANDED']

# How far past the current row count to extend the Supplier dropdown, so
# rows added later by StockIn_App still get the picker without re-running this.
DROPDOWN_ROW_BUFFER = 3000

# Hidden helper sheet + named ranges for dropdown sources. Excel's inline
# comma-list validation (formula1='"A,B,C"') is capped at 255 characters --
# CATEGORY_LIST alone is ~500 with the newer, more descriptive names -- so
# every list is sourced from here instead, regardless of length.
LISTS_SHEET_NAME = 'DropdownLists'


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_REORDER_SUPPLIER_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def main():
    backup_path = backup_master()
    print(f'Backup created: {backup_path}')

    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        print(f'CRITICAL: Cannot read {MASTER_FILE} -- close it in Excel first, then re-run.')
        return

    # Reorder_Level was a pre-existing but dead column (flat default of 20,
    # no script ever read or wrote it) -- superseded by Reorder_Point below.
    if 'Reorder_Level' in df.columns:
        df = df.drop(columns=['Reorder_Level'])
        print('Dropped unused legacy column: Reorder_Level (superseded by Reorder_Point).')

    added_cols = []
    if 'Reorder_Point' not in df.columns:
        df['Reorder_Point'] = DEFAULT_REORDER_POINT
        added_cols.append('Reorder_Point')
    else:
        blank_mask = df['Reorder_Point'].isna()
        if blank_mask.any():
            df.loc[blank_mask, 'Reorder_Point'] = DEFAULT_REORDER_POINT
            print(f"Reorder_Point already existed -- filled {blank_mask.sum()} blank cell(s) with the default.")
        else:
            print('Reorder_Point already existed and fully populated -- left as-is.')

    if 'Supplier' not in df.columns:
        df['Supplier'] = ''
        added_cols.append('Supplier')
    else:
        print('Supplier column already existed -- left existing values as-is.')

    if 'Category' not in df.columns:
        df['Category'] = ''
        added_cols.append('Category')
    else:
        print('Category column already existed -- left existing values as-is.')

    if 'Is_Generic' not in df.columns:
        df['Is_Generic'] = ''
        added_cols.append('Is_Generic')
    else:
        print('Is_Generic column already existed -- left existing values as-is.')

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    if added_cols:
        print(f"Added column(s): {', '.join(added_cols)}")

    # Second pass with openpyxl to attach dropdowns (data validation with
    # showErrorMessage=False so free-text entries are still allowed -- these
    # are convenience pickers, not enforced rules). NOTE: this dropdown gets
    # wiped the next time any other script rewrites Master via
    # pd.ExcelWriter().to_excel() -- Item_Settings_App.py is the durable way
    # to edit these fields, this is just a fallback for direct Excel use.
    wb = load_workbook(MASTER_FILE)
    ws = wb['INVENTORY MASTER']
    headers = [str(c.value).strip() if c.value else '' for c in ws[1]]
    last_row = max(ws.max_row, 2) + DROPDOWN_ROW_BUFFER

    # Clear any dropdowns from previous runs of this script first -- re-running
    # without this accumulates duplicate/overlapping validation definitions
    # for the same ranges, which Excel considers corrupt and auto-repairs by
    # stripping the feature (harmless to data, but loses the dropdown).
    existing_count = len(ws.data_validations.dataValidation)
    ws.data_validations.dataValidation = []
    if existing_count:
        print(f"Cleared {existing_count} pre-existing data validation(s) before re-adding (prevents duplicate/corrupt dropdowns).")

    # Rebuild the hidden helper sheet fresh each run too -- idempotent, and
    # avoids stale leftover rows if a list shrinks between runs.
    if LISTS_SHEET_NAME in wb.sheetnames:
        del wb[LISTS_SHEET_NAME]
    lists_ws = wb.create_sheet(LISTS_SHEET_NAME)
    lists_ws.sheet_state = 'hidden'

    for list_col_idx, (col_name, options) in enumerate([('Supplier', SUPPLIER_LIST), ('Category', CATEGORY_LIST), ('Is_Generic', IS_GENERIC_OPTIONS)], start=1):
        col_idx = headers.index(col_name) + 1
        col_letter = get_column_letter(col_idx)
        list_options = [o for o in options if o]  # allow_blank already covers the blank/N-A case

        list_col_letter = get_column_letter(list_col_idx)
        for row_idx, value in enumerate(list_options, start=1):
            lists_ws.cell(row=row_idx, column=list_col_idx, value=value)

        range_name = f'{col_name}DropdownRange'
        range_ref = f"'{LISTS_SHEET_NAME}'!${list_col_letter}$1:${list_col_letter}${len(list_options)}"
        wb.defined_names[range_name] = DefinedName(range_name, attr_text=range_ref)

        dv = DataValidation(
            type='list',
            formula1=range_name,
            allow_blank=True,
            showErrorMessage=False,
        )
        ws.add_data_validation(dv)
        dv.add(f'{col_letter}2:{col_letter}{last_row}')
        print(f"{col_name} dropdown attached to column {col_letter} (rows 2-{last_row}), sourced from hidden '{LISTS_SHEET_NAME}' sheet ({len(list_options)} option(s)), free-text entry still allowed.")

    wb.save(MASTER_FILE)

    print(f"Supplier options: {', '.join(SUPPLIER_LIST)}")
    print(f"Category options: {', '.join(CATEGORY_LIST)}")
    print(f"Is_Generic options: GENERIC, BRANDED, or blank/N-A (medicine-only, optional)")
    print(f"\nFinal columns: {df.columns.tolist()}")
    print(f"Total rows: {len(df)}")


if __name__ == '__main__':
    main()
