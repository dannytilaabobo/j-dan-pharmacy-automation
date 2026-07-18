import html
import difflib
import datetime
import pandas as pd

from name_cleaning import DASH_CHARS, QUOTE_MAP

MASTER_FILE = 'Inventory_Master.xlsx'
REPORT_DIR = 'JDan_Executive_Reports'


def normalize_dash(text):
    for d in DASH_CHARS:
        text = text.replace(d, '-')
    return text


def normalize_quotes(text):
    for smart, straight in QUOTE_MAP.items():
        text = text.replace(smart, straight)
    return text


def collapse_ws(text):
    return ' '.join(text.split())


def classify_diff(a, b):
    """Return the list of normalization steps needed to make a and b equal,
    or a final verdict if they still differ after all known cleanups."""
    steps = []

    a1, b1 = html.unescape(a), html.unescape(b)
    if a1 != a or b1 != b:
        steps.append('HTML_ENTITY')

    a2, b2 = normalize_dash(a1), normalize_dash(b1)
    if a2 != a1 or b2 != b1:
        steps.append('DASH')

    a3, b3 = normalize_quotes(a2), normalize_quotes(b2)
    if a3 != a2 or b3 != b2:
        steps.append('SMART_QUOTE')

    a4, b4 = collapse_ws(a3), collapse_ws(b3)
    if a4 != a3 or b4 != b3:
        steps.append('WHITESPACE')

    if a4 == b4:
        return steps if steps else ['UNKNOWN_BUT_RESOLVED']

    if a4.upper() == b4.upper():
        steps.append('CASE_DIFFERENCE')
        return steps

    ratio = difflib.SequenceMatcher(None, a4.upper(), b4.upper()).ratio()
    if ratio > 0.85:
        steps.append('LIKELY_TYPO_OR_TRANSPOSITION')
    else:
        steps.append('SUBSTANTIALLY_DIFFERENT -- VERIFY MANUALLY (possible wrong code)')
    return steps


def scan_split_by_code(df):
    """Same Item_Code, different Product_Name spellings across lot rows."""
    df = df.copy()
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['_name'] = df['Product_Name'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)

    variant_counts = df.groupby('Item_Code')['_name'].nunique()
    flagged_codes = variant_counts[variant_counts > 1].index

    rows = []
    for code in flagged_codes:
        sub = df[df['Item_Code'] == code]
        name_freq = sub.groupby('_name').agg(count=('_name', 'size'), stock=('Current_Stock', 'sum')).reset_index()
        name_freq = name_freq.sort_values(['count', 'stock'], ascending=False)
        baseline = name_freq.iloc[0]['_name']
        variants = name_freq['_name'].tolist()

        issue_types = set()
        for v in variants:
            if v == baseline:
                continue
            issue_types.update(classify_diff(baseline, v))

        statuses = sub['Status'].dropna().unique().tolist() if 'Status' in sub.columns else []

        rows.append({
            'Item_Code': code,
            'Total_Stock_Affected': sub['Current_Stock'].sum(),
            'Num_Variants': len(variants),
            'Baseline_Name': baseline,
            'All_Variant_Names': ' || '.join(variants),
            'Issue_Type(s)': ', '.join(sorted(issue_types)) if issue_types else 'UNKNOWN',
            'Status(es)': ', '.join(str(s) for s in statuses),
        })

    result = pd.DataFrame(rows)
    if not result.empty:
        result = result.sort_values('Total_Stock_Affected', ascending=False).reset_index(drop=True)
    return result


def scan_same_name_diff_code(df):
    """Exact same Product_Name text shared by two or more different Item_Codes."""
    df = df.copy()
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['_name'] = df['Product_Name'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)

    code_counts = df.groupby('_name')['Item_Code'].nunique()
    flagged_names = code_counts[code_counts > 1].index

    rows = []
    for name in flagged_names:
        sub = df[df['_name'] == name]
        codes = sub.groupby('Item_Code').agg(stock=('Current_Stock', 'sum')).reset_index()
        rows.append({
            'Product_Name': name,
            'Num_Item_Codes': sub['Item_Code'].nunique(),
            'Item_Codes_And_Stock': ' || '.join(f"{r['Item_Code']} (stock={r['stock']:.0f})" for _, r in codes.iterrows()),
            'Total_Stock_Combined': sub['Current_Stock'].sum(),
        })

    result = pd.DataFrame(rows)
    if not result.empty:
        result = result.sort_values('Total_Stock_Combined', ascending=False).reset_index(drop=True)
    return result


def main():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')

    print('=' * 70)
    print(' NAME INCONSISTENCY SCAN -- Inventory_Master.xlsx')
    print(' (SCAN ONLY -- walang na-apply na correction)')
    print('=' * 70)
    print(f'Total rows scanned: {len(df)}')

    split_by_code = scan_split_by_code(df)
    print(f'\n--- (1) SAME Item_Code, MAGKAIBANG Product_Name: {len(split_by_code)} code(s) ---')
    if not split_by_code.empty:
        with pd.option_context('display.max_colwidth', 55, 'display.width', 250):
            print(split_by_code[['Item_Code', 'Total_Stock_Affected', 'Num_Variants', 'Baseline_Name', 'Issue_Type(s)']].to_string(index=False))
    else:
        print('  Wala.')

    same_name = scan_same_name_diff_code(df)
    print(f'\n--- (2) SAME Product_Name, MAGKAIBANG Item_Code: {len(same_name)} name(s) ---')
    if not same_name.empty:
        with pd.option_context('display.max_colwidth', 55, 'display.width', 250):
            print(same_name[['Product_Name', 'Num_Item_Codes', 'Total_Stock_Combined']].to_string(index=False))
    else:
        print('  Wala.')

    out_path = f'{REPORT_DIR}/NAME_INCONSISTENCY_SCAN_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        split_by_code.to_excel(writer, sheet_name='Same_Code_Diff_Name', index=False)
        same_name.to_excel(writer, sheet_name='Same_Name_Diff_Code', index=False)
    print(f'\nFull report saved to: {out_path}')
    print('\nWala pang na-apply na correction -- scan lang ito.')


if __name__ == '__main__':
    main()
