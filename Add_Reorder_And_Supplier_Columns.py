import os
import shutil
import datetime
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'

DEFAULT_REORDER_POINT = 10

SUPPLIER_LIST = [
    'Dyna Drug', 'Comjem', 'Copharm', 'Latriel', 'RGSOI', 'Ice Climber',
    'Metrogen', 'Otsuka', 'Pocari', 'MedSupplies/Shopee', 'Shopee', 'Others',
]

# How far past the current row count to extend the Supplier dropdown, so
# rows added later by StockIn_App still get the picker without re-running this.
DROPDOWN_ROW_BUFFER = 3000


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

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f'CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.')
        return

    if added_cols:
        print(f"Added column(s): {', '.join(added_cols)}")

    # Second pass with openpyxl to attach the Supplier dropdown (data
    # validation with showErrorMessage=False so free-text entries are still
    # allowed -- the list is just a convenience picker, not an enforced rule).
    wb = load_workbook(MASTER_FILE)
    ws = wb['INVENTORY MASTER']
    headers = [str(c.value).strip() if c.value else '' for c in ws[1]]
    supplier_col_idx = headers.index('Supplier') + 1
    supplier_col_letter = get_column_letter(supplier_col_idx)

    dv = DataValidation(
        type='list',
        formula1='"' + ','.join(SUPPLIER_LIST) + '"',
        allow_blank=True,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv)
    last_row = max(ws.max_row, 2) + DROPDOWN_ROW_BUFFER
    dv.add(f'{supplier_col_letter}2:{supplier_col_letter}{last_row}')
    wb.save(MASTER_FILE)

    print(f"Supplier dropdown attached to column {supplier_col_letter} (rows 2-{last_row}), free-text entry still allowed.")
    print(f"Supplier options: {', '.join(SUPPLIER_LIST)}")
    print(f"\nFinal columns: {df.columns.tolist()}")
    print(f"Total rows: {len(df)}")


if __name__ == '__main__':
    main()
