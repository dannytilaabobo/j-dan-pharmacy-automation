"""Check_Pending_Sync_Codes.py - READ-ONLY check of item codes waiting for / written to the cloud logs.

Scans PENDING_CLOUD_SYNC\\*.json (Stock_Adjustments rows) and Stock_Adjustments_Log.csv and reports any Item Code that
  - shows scientific notation or a decimal point,
  - lost its leading zero (a zero-padded version of it is a known barcode or Master code),
  - is a crosswalk BARCODE where the internal (old) code is expected.
Codes that start with an apostrophe (the protection used for Google Sheets) are checked without it.
Results go only into a new folder Barcode_Verification_<date_time>.

Optional arguments (used by the tests): --pending-dir D --adjust-log F --master F --crosswalk F --out-root D
"""
import os
import sys
import csv
import glob
import json
import shutil
import tempfile
import argparse
import datetime
import warnings

warnings.filterwarnings('ignore')
os.chdir(os.path.dirname(os.path.abspath(__file__)))
PROJECT = os.getcwd()
sys.path.insert(0, PROJECT)
import item_code_resolver as ic   # noqa: E402


def classify(raw, resolver):
    """(problem text or '', protected_by_apostrophe)"""
    code = raw.strip()
    protected = code.startswith("'")
    if protected:
        code = code[1:]
    if not code:
        return 'empty Item Code', protected
    res = resolver.resolve(code)
    if res.status == ic.CORRUPTED:
        return 'DAMAGED: ' + res.reason, protected
    if code in resolver.crosswalk.by_barcode and code not in (resolver.master_codes or set()):
        old = resolver.crosswalk.by_barcode[code]['Old_Item_Code']
        return f'BARCODE where the internal (old) code is expected (should be {old})', protected
    return '', protected


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--pending-dir', default=os.path.join(PROJECT, 'PENDING_CLOUD_SYNC'))
    ap.add_argument('--adjust-log', default=os.path.join(PROJECT, 'Stock_Adjustments_Log.csv'))
    ap.add_argument('--master', default=os.path.join(PROJECT, 'Inventory_Master.xlsx'))
    ap.add_argument('--crosswalk', default=ic.CROSSWALK_FILE)
    ap.add_argument('--out-root', default=PROJECT)
    args = ap.parse_args(argv)

    import pandas as pd
    tmp = tempfile.mkdtemp(prefix='pending_check_master_')
    mc = os.path.join(tmp, 'Master_COPY.xlsx')
    shutil.copy2(args.master, mc)
    master_codes = set(pd.read_excel(mc, sheet_name='INVENTORY MASTER', dtype=str, keep_default_na=False)['Item_Code'].str.strip())
    shutil.rmtree(tmp, ignore_errors=True)
    resolver = ic.Resolver(ic.load_crosswalk(args.crosswalk), master_codes)

    findings, checked, protected_n, unknown_n = [], 0, 0, 0
    files = sorted(glob.glob(os.path.join(args.pending_dir, '*.json')) + glob.glob(os.path.join(args.pending_dir, '_QUARANTINE_MALFORMED', '*.json')))
    for p in files:
        try:
            with open(p, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except (OSError, ValueError):
            findings.append({'Where': os.path.basename(p), 'Item_Code': '', 'Problem': 'file could not be read (still being written, or malformed)'})
            continue
        if not isinstance(data, dict) or data.get('target_tab') != 'Stock_Adjustments':
            continue
        row = data.get('row_data') or []
        if len(row) < 3:
            findings.append({'Where': os.path.basename(p), 'Item_Code': '', 'Problem': 'Stock_Adjustments row has no Item Code column'})
            continue
        checked += 1
        raw = str(row[2])
        prob, prot = classify(raw, resolver)
        protected_n += prot
        if prob:
            findings.append({'Where': 'PENDING_CLOUD_SYNC\\' + os.path.basename(p), 'Item_Code': raw, 'Problem': prob})
        elif raw.strip().lstrip("'") not in master_codes:
            unknown_n += 1
    log_rows = 0
    if os.path.exists(args.adjust_log):
        with open(args.adjust_log, 'r', encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f):
                log_rows += 1
                raw = r.get('Item_Code', '')
                checked += 1
                prob, prot = classify(raw, resolver)
                if prob:
                    findings.append({'Where': f'Stock_Adjustments_Log.csv line {log_rows + 1}', 'Item_Code': raw, 'Problem': prob})
                elif raw.strip().lstrip("'") not in master_codes:
                    unknown_n += 1

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out_dir = os.path.join(args.out_root, f'Barcode_Verification_{stamp}')
    n = 2
    while os.path.exists(out_dir):
        out_dir = os.path.join(args.out_root, f'Barcode_Verification_{stamp}_{n}')
        n += 1
    os.makedirs(out_dir)
    lines = [f'PENDING SYNC / ADJUSTMENT LOG CODE CHECK  {datetime.datetime.now():%Y-%m-%d %H:%M}', '',
             f'Files waiting in PENDING_CLOUD_SYNC (incl. quarantine): {len(files)}', f'Stock_Adjustments_Log.csv rows: {log_rows}',
             f'Item codes checked: {checked}   protected by an apostrophe: {protected_n}   not found in the Master: {unknown_n}', '']
    if findings:
        lines.append(f'PROBLEMS FOUND: {len(findings)}')
        for fnd in findings:
            lines.append(f"  {fnd['Where']}: {fnd['Item_Code']!r} - {fnd['Problem']}")
        lines.append('\nRESULT: STOP - do not let these rows reach Google Sheets until they are corrected. Nothing was changed by this tool.')
    else:
        lines.append('RESULT: OK - no damaged, zero-less or barcode-instead-of-old-code Item Code found.')
    text = '\n'.join(lines)
    print(text)
    with open(os.path.join(out_dir, 'PENDING_CODE_CHECK.txt'), 'w', encoding='utf-8') as f:
        f.write(text)
    with open(os.path.join(out_dir, 'PENDING_CODE_CHECK.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=['Where', 'Item_Code', 'Problem'], quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(findings)
    print(f'\nSaved in: {out_dir}')
    return 1 if findings else 0


if __name__ == '__main__':
    sys.exit(main())
