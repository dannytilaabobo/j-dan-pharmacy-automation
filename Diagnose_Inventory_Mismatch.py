import os
import re
import glob
import html
import difflib
import datetime
import pandas as pd

MASTER_FILE = 'Inventory_Master.xlsx'
SOD_DIR = 'JDan_Executive_Reports'
STOCKIN_DIR = 'JDan_StockIn_Reports'
EOD_DIR = os.path.join('JDan_Sales_System', 'EOD_Deduction_Reports')
PULLOUT_DIR = os.path.join('JDan_PullOut_System', '1_Excel_Reports')

LOOKBACK_DAYS = 14
MIN_MISMATCHES = 2
TOP_N = 8


def norm(s):
    return str(s).strip().upper()


def load_sod_reports(days=LOOKBACK_DAYS):
    cutoff = datetime.datetime.now() - datetime.timedelta(days=days)
    rows = []
    for path in glob.glob(os.path.join(SOD_DIR, 'SOD_Summary_*.txt')):
        with open(path, 'r', encoding='utf-8') as f:
            text = f.read()

        date_match = re.search(r'Target Date:\s*([\d-]+)', text)
        if not date_match:
            continue
        report_date = datetime.datetime.strptime(date_match.group(1), '%Y-%m-%d')
        if report_date < cutoff:
            continue

        for block in text.split('---------------------------------------'):
            item_match = re.search(r'Item:\s*(.+)', block)
            qty_match = re.search(r'System Qty:\s*([\d.]+)\s*\|\s*Actual Count:\s*([\d.]+)', block)
            var_match = re.search(r'Variance:\s*(-?[\d.]+)', block)
            just_match = re.search(r'Justification:\s*(.+)', block)
            if item_match and qty_match and var_match:
                rows.append({
                    'date': report_date,
                    'item_raw': item_match.group(1).strip(),
                    'item': norm(item_match.group(1)),
                    'system_qty': float(qty_match.group(1)),
                    'actual_qty': float(qty_match.group(2)),
                    'variance': float(var_match.group(1)),
                    'justification': just_match.group(1).strip() if just_match else '',
                    'source_file': os.path.basename(path),
                })
    return pd.DataFrame(rows)


def find_repeat_offenders(df_sod):
    if df_sod.empty:
        return pd.DataFrame()
    mismatches = df_sod[df_sod['variance'] != 0]
    agg = mismatches.groupby('item').agg(
        item_raw=('item_raw', 'first'),
        times_mismatched=('variance', 'count'),
        total_abs_variance=('variance', lambda s: s.abs().sum()),
        avg_variance=('variance', 'mean'),
        min_variance=('variance', 'min'),
        max_variance=('variance', 'max'),
        dates=('date', lambda s: ', '.join(d.strftime('%m/%d') for d in sorted(s))),
    ).reset_index()
    agg = agg[agg['times_mismatched'] >= MIN_MISMATCHES]
    agg = agg.sort_values(['times_mismatched', 'total_abs_variance'], ascending=False)
    return agg.head(TOP_N)


def load_master():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    df['_norm_name'] = df['Product_Name'].apply(norm)
    return df


def load_all_stockin():
    frames = []
    for path in glob.glob(os.path.join(STOCKIN_DIR, '*', 'STOCKIN_BATCH_*.xlsx')):
        df = pd.read_excel(path)
        df['source_file'] = os.path.basename(path)
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Product_Name', 'Qty_Delivered', 'source_file'])
    out = pd.concat(frames, ignore_index=True)
    out['_norm_name'] = out['Product_Name'].apply(norm)
    out['Qty_Delivered'] = pd.to_numeric(out['Qty_Delivered'], errors='coerce').fillna(0)
    return out


def load_all_sales():
    frames = []
    for path in glob.glob(os.path.join(EOD_DIR, 'PROCESSED_EOD_*.xlsx')):
        df = pd.read_excel(path)
        df['source_file'] = os.path.basename(path)
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Description', 'Qty', 'source_file'])
    out = pd.concat(frames, ignore_index=True)
    out['_norm_name'] = out['Description'].apply(norm)
    out['Qty'] = pd.to_numeric(out['Qty'], errors='coerce').fillna(0)
    return out


def load_all_pullout():
    frames = []
    for path in glob.glob(os.path.join(PULLOUT_DIR, 'COMPLETED_PULLOUT_*.xlsx')):
        df = pd.read_excel(path)
        df['source_file'] = os.path.basename(path)
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Product_Name', 'Actual_Pulled_Qty', 'source_file'])
    out = pd.concat(frames, ignore_index=True)
    out['_norm_name'] = out['Product_Name'].apply(norm)
    out['Actual_Pulled_Qty'] = pd.to_numeric(out['Actual_Pulled_Qty'], errors='coerce').fillna(0)
    return out


def diagnose_item(item_norm, master_df, stockin_df, sales_df, pullout_df):
    result = {'item': item_norm}

    master_rows = master_df[master_df['_norm_name'] == item_norm]
    result['master_rows'] = len(master_rows)
    result['master_item_codes'] = sorted(master_rows['Item_Code'].astype(str).unique().tolist())
    result['master_locations'] = sorted(master_rows['LOCATION'].dropna().astype(str).unique().tolist())
    result['master_current_stock'] = master_rows['Current_Stock'].sum()

    stockin_rows = stockin_df[stockin_df['_norm_name'] == item_norm]
    result['total_stockin'] = stockin_rows['Qty_Delivered'].sum()

    sales_rows = sales_df[sales_df['_norm_name'] == item_norm]
    result['total_sales'] = sales_rows['Qty'].abs().sum()

    pullout_rows = pullout_df[pullout_df['_norm_name'] == item_norm]
    result['total_pullout'] = pullout_rows['Actual_Pulled_Qty'].sum()

    # Replicates Daily_Deduct_App's matching logic (exact strip+upper equality against
    # Master Product_Name). If a sale's description normalizes to this item but no
    # Master row shares that exact normalized name, that sale's deduction silently
    # never happened -- Master would overstate stock versus the physical shelf count.
    all_master_names = set(master_df['_norm_name'])
    unmatched_sales_descriptions = sorted({
        raw for raw in sales_rows['Description'].astype(str).unique()
        if norm(raw) not in all_master_names
    })
    result['unmatched_sales_descriptions'] = unmatched_sales_descriptions

    # Near-duplicate product names elsewhere in Master (packaging/spelling variants
    # that staff or encoders could confuse with this item).
    other_names = [n for n in master_df['_norm_name'].unique() if n != item_norm]
    close = difflib.get_close_matches(item_norm, other_names, n=5, cutoff=0.82)
    result['similar_master_names'] = close

    result['html_entity_artifact'] = '&#' in item_norm or any('&#' in n for n in close)

    causes = []
    if len(result['master_item_codes']) > 1:
        causes.append('MASTER DUPLICATE/SPLIT ITEM_CODE for same product name')
    if len(result['master_locations']) > 1:
        causes.append('ITEM SPLIT ACROSS MULTIPLE LOCATIONS/GONDOLAS')
    if unmatched_sales_descriptions:
        causes.append('SALES DEDUCTION MISMATCH (POS description text does not match Master name -> stock not deducted)')
    if close:
        causes.append('POSSIBLE NAME CONFUSION WITH SIMILAR ITEM(S) IN MASTER')
    if result['html_entity_artifact']:
        causes.append('CORRUPTED/HTML-ENCODED CHARACTERS IN NAME (e.g. &#039; instead of an apostrophe)')
    if not causes:
        causes.append('UNRESOLVED -- no encoding/matching issue found; recommend manual shelf audit (possible miscount, damage, or theft)')
    result['root_cause_guess'] = '; '.join(causes)

    return result


def main():
    print('=' * 70)
    print(' J-DAN INVENTORY MISMATCH DIAGNOSTIC')
    print(f' Scanning SOD reports from the last {LOOKBACK_DAYS} days')
    print('=' * 70)

    df_sod = load_sod_reports()
    if df_sod.empty:
        print('No SOD reports found in the lookback window. Nothing to diagnose.')
        return
    print(f'\nLoaded {len(df_sod)} item-entries from {df_sod["source_file"].nunique()} SOD reports '
          f'({df_sod["date"].min().date()} to {df_sod["date"].max().date()}).')

    offenders = find_repeat_offenders(df_sod)
    if offenders.empty:
        print(f'\nNo items mismatched {MIN_MISMATCHES}+ times in the window. No repeat offenders found.')
        return

    print(f'\n--- REPEAT OFFENDERS (mismatched {MIN_MISMATCHES}+ times) ---')
    print(offenders[['item_raw', 'times_mismatched', 'total_abs_variance', 'avg_variance',
                      'min_variance', 'max_variance', 'dates']].to_string(index=False))

    master_df = load_master()
    stockin_df = load_all_stockin()
    sales_df = load_all_sales()
    pullout_df = load_all_pullout()

    diagnostics = []
    for _, row in offenders.iterrows():
        diagnostics.append(diagnose_item(row['item'], master_df, stockin_df, sales_df, pullout_df))

    diag_df = pd.DataFrame(diagnostics)

    print('\n' + '=' * 70)
    print(' ROOT-CAUSE DIAGNOSTIC TABLE')
    print('=' * 70)
    display_df = diag_df[['item', 'master_rows', 'master_item_codes', 'master_locations',
                           'master_current_stock', 'total_stockin', 'total_sales',
                           'total_pullout', 'root_cause_guess']].copy()
    display_df.columns = ['Item', 'Master Rows', 'Item_Codes', 'Locations', 'Current Stock (Master)',
                           'Total Stock-In (all-time)', 'Total Sales Deducted (all-time)',
                           'Total Pulled-Out (all-time)', 'Root Cause Guess']
    with pd.option_context('display.max_colwidth', 60, 'display.width', 220):
        print(display_df.to_string(index=False))

    print('\n--- DETAILS: Unmatched sales descriptions per item (if any) ---')
    any_unmatched = False
    for d in diagnostics:
        if d['unmatched_sales_descriptions']:
            any_unmatched = True
            print(f"  {d['item']}: {d['unmatched_sales_descriptions']}")
    if not any_unmatched:
        print('  None found -- all sales descriptions for the flagged items match a Master row exactly.')

    print('\n--- DETAILS: Similar/near-duplicate Master names per item (if any) ---')
    any_similar = False
    for d in diagnostics:
        if d['similar_master_names']:
            any_similar = True
            print(f"  {d['item']}: {d['similar_master_names']}")
    if not any_similar:
        print('  None found.')

    out_dir = 'JDan_Executive_Reports'
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'DIAGNOSTIC_MISMATCH_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        offenders.to_excel(writer, sheet_name='Repeat_Offenders', index=False)
        display_df.to_excel(writer, sheet_name='Root_Cause_Diagnostic', index=False)
        df_sod.to_excel(writer, sheet_name='Raw_SOD_Data', index=False)
    print(f'\nFull report saved to: {out_path}')


if __name__ == '__main__':
    main()
