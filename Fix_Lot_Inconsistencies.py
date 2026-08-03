"""
One-time cleanup: normalize Reorder_Point/Supplier across lot rows of the
same Item_Code in Inventory_Master.xlsx.

Root cause (see StockIn_App.py investigation): new lots added for an
EXISTING Item_Code were getting a hardcoded default (Reorder_Point=10,
Supplier=blank) instead of inheriting the value already set on that
product's other lot rows.

Usage:
    python Fix_Lot_Inconsistencies.py            -> dry run only (default,
                                                      safe, no writes)
    python Fix_Lot_Inconsistencies.py --apply    -> backs up Master, then
                                                      applies the AUTO-FIX
                                                      rows only
"""
import argparse
import csv
import datetime
import glob
import json
import os
import shutil
import sys

import pandas as pd

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
REPORT_DIR = 'JDan_Executive_Reports'
AUDIT_LOG = os.path.join(REPORT_DIR, 'LOT_CLEANUP_AUDIT_LOG.csv')
AUDIT_LOG_HEADER = ['Timestamp', 'Item_Code', 'Product_Name', 'Old_Reorder_Point',
                     'New_Reorder_Point', 'Old_Supplier', 'New_Supplier', 'Lots_Updated', 'Reason', 'Applied_By']
REASON = 'Lot inconsistency cleanup - normalized to established value'

DEFAULT_REORDER_POINT = 10.0


def load_active_lots():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED'].copy()
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Reorder_Point'] = pd.to_numeric(df['Reorder_Point'], errors='coerce').fillna(DEFAULT_REORDER_POINT)
    df['Supplier'] = df['Supplier'].fillna('').astype(str).str.strip() if 'Supplier' in df.columns else ''
    df['_name_clean'] = df['Product_Name'].astype(str).str.strip()
    return df


def find_inconsistent_codes(df):
    inconsistency = df.groupby('Item_Code').agg(
        RP_Nunique=('Reorder_Point', 'nunique'),
        Supplier_Nunique=('Supplier', lambda s: s[s != ''].nunique()),
        Name_Nunique=('_name_clean', 'nunique'),
    ).reset_index()
    return inconsistency[
        (inconsistency['RP_Nunique'] > 1) | (inconsistency['Supplier_Nunique'] > 1) | (inconsistency['Name_Nunique'] > 1)
    ]['Item_Code'].tolist()


def is_default_pair(rp, supplier):
    return float(rp) == DEFAULT_REORDER_POINT and str(supplier).strip() == ''


def build_plan():
    """Returns a list of dict rows, one per flagged Item_Code, with a
    Category of AUTO-FIX or NEEDS MANUAL REVIEW and (for AUTO-FIX) the
    single established Reorder_Point/Supplier value to apply to every lot."""
    df = load_active_lots()
    flagged_codes = find_inconsistent_codes(df)

    plan = []
    for code in flagged_codes:
        group = df[df['Item_Code'] == code]
        product_name = group['_name_clean'].mode().iloc[0]
        name_nunique = group['_name_clean'].nunique()

        rp_list = group['Reorder_Point'].tolist()
        supplier_list = group['Supplier'].tolist()
        current_rp_display = ', '.join(f'{v:.1f}' for v in rp_list)
        current_supplier_display = ', '.join((s if s else '(blank)') for s in supplier_list)

        pairs = group[['Reorder_Point', 'Supplier']].drop_duplicates()

        if name_nunique > 1:
            plan.append({
                'Item_Code': code, 'Product_Name': product_name, 'Category': 'NEEDS MANUAL REVIEW',
                'Current_Reorder_Point_Per_Lot': current_rp_display,
                'Current_Supplier_Per_Lot': current_supplier_display,
                'Proposed_Final_Reorder_Point': '', 'Proposed_Final_Supplier': '',
                'Notes': 'Product_Name also differs across lots -- name conflict is out of scope for this '
                         'Reorder_Point/Supplier cleanup, needs separate manual correction.',
            })
            continue

        non_default = pairs[~pairs.apply(lambda r: is_default_pair(r['Reorder_Point'], r['Supplier']), axis=1)]
        distinct_established = non_default.drop_duplicates()

        if len(distinct_established) == 1:
            est_rp = float(distinct_established.iloc[0]['Reorder_Point'])
            est_supplier = distinct_established.iloc[0]['Supplier']
            plan.append({
                'Item_Code': code, 'Product_Name': product_name, 'Category': 'AUTO-FIX',
                'Current_Reorder_Point_Per_Lot': current_rp_display,
                'Current_Supplier_Per_Lot': current_supplier_display,
                'Proposed_Final_Reorder_Point': f'{est_rp:.1f}',
                'Proposed_Final_Supplier': est_supplier if est_supplier else '(blank)',
                'Notes': '',
            })
        elif len(distinct_established) == 0:
            plan.append({
                'Item_Code': code, 'Product_Name': product_name, 'Category': 'NEEDS MANUAL REVIEW',
                'Current_Reorder_Point_Per_Lot': current_rp_display,
                'Current_Supplier_Per_Lot': current_supplier_display,
                'Proposed_Final_Reorder_Point': '', 'Proposed_Final_Supplier': '',
                'Notes': 'No non-default (non 10.0/blank) value found on any lot -- unexpected, verify manually.',
            })
        else:
            conflict_desc = '; '.join(
                f"{r['Reorder_Point']:.1f}/{r['Supplier'] if r['Supplier'] else '(blank)'}"
                for _, r in distinct_established.iterrows()
            )
            plan.append({
                'Item_Code': code, 'Product_Name': product_name, 'Category': 'NEEDS MANUAL REVIEW',
                'Current_Reorder_Point_Per_Lot': current_rp_display,
                'Current_Supplier_Per_Lot': current_supplier_display,
                'Proposed_Final_Reorder_Point': '', 'Proposed_Final_Supplier': '',
                'Notes': f'Multiple differing non-default values found: {conflict_desc} -- true conflict, pick manually.',
            })

    return plan


def dry_run():
    plan = build_plan()
    auto_fix = [p for p in plan if p['Category'] == 'AUTO-FIX']
    manual_review = [p for p in plan if p['Category'] == 'NEEDS MANUAL REVIEW']

    os.makedirs(REPORT_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    out_path = os.path.join(REPORT_DIR, f'LOT_CLEANUP_DRYRUN_{ts}.csv')

    columns = ['Item_Code', 'Product_Name', 'Category', 'Current_Reorder_Point_Per_Lot',
               'Current_Supplier_Per_Lot', 'Proposed_Final_Reorder_Point', 'Proposed_Final_Supplier', 'Notes']
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for row in plan:
            writer.writerow(row)

    print(f"Dry-run report saved to: {out_path}")
    print()
    print(f"Total flagged Item_Codes: {len(plan)}")
    print(f"  AUTO-FIX (single clear established value): {len(auto_fix)}")
    print(f"  NEEDS MANUAL REVIEW: {len(manual_review)}")
    if manual_review:
        print()
        print("Manual review items:")
        for row in manual_review:
            print(f"  {row['Item_Code']} ({row['Product_Name']}): {row['Notes']}")

    return out_path, plan


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_LOT_CLEANUP_{ts}.xlsx')
    shutil.copy2(MASTER_FILE, backup_path)
    return backup_path


def apply_fix(staff_name):
    plan = build_plan()
    auto_fix = [p for p in plan if p['Category'] == 'AUTO-FIX']

    if not auto_fix:
        print("No AUTO-FIX items to apply. Nothing to do.")
        return

    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()

    backup_path = backup_master()
    print(f"Backup created: {backup_path}")

    ts = datetime.datetime.now()
    log_time = ts.strftime('%Y-%m-%d %I:%M %p')
    audit_rows = []

    for row in auto_fix:
        code = row['Item_Code']
        mask = df['Item_Code'] == code
        old_rp_values = df.loc[mask, 'Reorder_Point'].tolist()
        old_supplier_values = df.loc[mask, 'Supplier'].tolist()
        new_rp = float(row['Proposed_Final_Reorder_Point'])
        new_supplier = '' if row['Proposed_Final_Supplier'] == '(blank)' else row['Proposed_Final_Supplier']

        df.loc[mask, 'Reorder_Point'] = new_rp
        df.loc[mask, 'Supplier'] = new_supplier

        audit_rows.append([
            log_time, code, row['Product_Name'],
            ', '.join(f'{v:.1f}' if pd.notna(v) else '(blank)' for v in old_rp_values), f'{new_rp:.1f}',
            ', '.join((str(s).strip() if pd.notna(s) and str(s).strip() else '(blank)') for s in old_supplier_values),
            (new_supplier if new_supplier else '(blank)'),
            mask.sum(), REASON, staff_name.upper(),
        ])

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        print(f"CRITICAL: Cannot save {MASTER_FILE} -- close it in Excel and re-run. No changes were written.")
        return

    is_new = not os.path.exists(AUDIT_LOG)
    os.makedirs(REPORT_DIR, exist_ok=True)
    with open(AUDIT_LOG, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(AUDIT_LOG_HEADER)
        writer.writerows(audit_rows)

    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{log_time}] - USER: {staff_name.upper()} - LOT INCONSISTENCY CLEANUP: "
                       f"{len(auto_fix)} Item_Code(s) normalized\n")
        for r in audit_rows:
            logbook.write(f"   -> {r[1]} ({r[2]}): Reorder_Point [{r[3]}] -> {r[4]}, "
                           f"Supplier [{r[5]}] -> {r[6]} ({r[7]} lot rows)\n")
        logbook.write(f"   -> Reason: {REASON}\n")

    print(f"Applied {len(auto_fix)} AUTO-FIX Item_Code(s). Audit log: {AUDIT_LOG}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--apply', action='store_true', help='Actually write the fix to Master (default is dry-run only)')
    parser.add_argument('--staff', default=None, help='Staff name for the audit log (required with --apply)')
    args = parser.parse_args()

    if args.apply:
        if not args.staff:
            print("ERROR: --staff NAME is required when using --apply (for the audit log).")
            sys.exit(1)
        apply_fix(args.staff)
    else:
        dry_run()
