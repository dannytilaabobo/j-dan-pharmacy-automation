"""Snapshot_BITPOS_Before.py - the "before the batch" record. READ-ONLY on live data.

Copies the newest BITPOS item export and the newest sales export into a new folder
Barcode_Verification_<date_time>\\snapshots (read-only copies, fingerprints recorded).
The exports are never opened in Excel and never modified.

Optional arguments (used by the tests): --items <file> --sales <file> --out-root <folder> --yes
"""
import os
import sys
import csv
import json
import glob
import stat
import shutil
import hashlib
import argparse
import datetime

os.chdir(os.path.dirname(os.path.abspath(__file__)))
PROJECT = os.getcwd()
DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')
DESKTOP = os.path.join(os.path.expanduser('~'), 'Desktop')
INBOX = os.path.join(PROJECT, 'SALES_REPORT_FROM_POS')
ARCHIVE = os.path.join(PROJECT, 'JDan_Sales_System', 'Archived_Sales_Reports')
CROSSWALK = os.path.join(PROJECT, 'Barcode_Crosswalk.csv')


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def looks_like_item_export(path):
    try:
        with open(path, 'r', encoding='utf-8-sig', errors='replace', newline='') as f:
            header = next(csv.reader(f))
    except (OSError, StopIteration, csv.Error):
        return False
    up = [h.strip().upper() for h in header]
    return 'ITEM CODE' in up and 'DESCRIPTION' in up and 'QTY' not in up


def find_item_exports(extra_dirs=()):
    found = []
    for d in [DOWNLOADS, DESKTOP, PROJECT] + list(extra_dirs):
        for p in glob.glob(os.path.join(d, '*.csv')):
            if 'Barcode_Verification_' in p or os.path.basename(p).startswith('Barcode_Crosswalk'):
                continue
            if looks_like_item_export(p):
                found.append(p)
    return sorted(set(found), key=os.path.getmtime, reverse=True)


def find_sales_exports():
    inbox = [p for p in glob.glob(os.path.join(INBOX, '*')) if p.lower().endswith(('.xls', '.xlsx', '.csv'))]
    arch = [p for p in glob.glob(os.path.join(ARCHIVE, '*')) if p.lower().endswith(('.xls', '.xlsx', '.csv'))]
    return sorted(inbox, key=os.path.getmtime, reverse=True), sorted(arch, key=os.path.getmtime, reverse=True)


def describe(p):
    t = datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%Y-%m-%d %H:%M')
    return f'{p}   (saved {t}, {os.path.getsize(p):,} bytes)'


def choose(candidates, label, input_fn):
    if not candidates:
        return None
    print(f'\nNewest {label}:\n  {describe(candidates[0])}')
    if len(candidates) > 1:
        print('Other recent files:')
        for i, p in enumerate(candidates[1:6], 2):
            print(f'  {i}. {describe(p)}')
    ans = input_fn('Press Enter to use the newest file, or type the number of another one (or N to skip): ').strip().lower()
    if ans in ('n', 'no', 'skip'):
        return None
    if ans.isdigit() and 1 <= int(ans) <= min(len(candidates), 6):
        return candidates[int(ans) - 1]
    return candidates[0]


def count_item_rows(path):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        rows = list(csv.reader(f))
    header = [h.strip().upper() for h in rows[0]]
    ci = header.index('ITEM CODE')
    codes = [r[ci].strip() for r in rows[1:] if len(r) > ci]
    return len(rows) - 1, codes


def main(argv=None, input_fn=input):
    ap = argparse.ArgumentParser()
    ap.add_argument('--items')
    ap.add_argument('--sales')
    ap.add_argument('--out-root', default=PROJECT)
    ap.add_argument('--yes', action='store_true')
    args = ap.parse_args(argv)
    ask = (lambda prompt: '') if args.yes else input_fn

    print('=' * 70)
    print(' SNAPSHOT OF BITPOS - BEFORE THE BARCODE BATCH (read-only)')
    print('=' * 70)
    items = args.items or choose(find_item_exports(), 'BITPOS item export', ask)
    if not items:
        print('\nNo BITPOS item export was found or chosen. Export the item list from BITPOS first (do not open it in Excel).')
        return 2
    if args.sales:
        sales = args.sales
    else:
        inbox, arch = find_sales_exports()
        sales = choose(inbox or arch, 'BITPOS sales export' + ('' if inbox else ' (inbox is empty - newest ARCHIVED day)'), ask)

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out_dir = os.path.join(args.out_root, f'Barcode_Verification_{stamp}')
    snap_dir = os.path.join(out_dir, 'snapshots')
    os.makedirs(snap_dir, exist_ok=True)
    manifest = {'taken_at': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'purpose': 'BEFORE the barcode batch', 'files': []}
    lines = []
    for role, src in (('item export', items), ('sales export', sales)):
        if not src:
            continue
        dest = os.path.join(snap_dir, f'BEFORE_{stamp}__{os.path.basename(src)}')
        shutil.copy2(src, dest)
        os.chmod(dest, stat.S_IREAD)
        entry = {'role': role, 'source': os.path.abspath(src), 'source_modified': datetime.datetime.fromtimestamp(os.path.getmtime(src)).strftime('%Y-%m-%d %H:%M:%S'),
                 'snapshot': os.path.abspath(dest), 'bytes': os.path.getsize(dest), 'sha256': sha256(dest)}
        if sha256(src) != entry['sha256']:
            print(f'STOP: the copy of {src} does not match the original. Nothing was recorded.')
            return 1
        if role == 'item export':
            n, codes = count_item_rows(dest)
            entry['rows'] = n
            odd = [c for c in codes if c and not (c.isdigit() and len(c) == 8)]
            entry['codes_not_8_digits'] = len(odd)
            barcodes = set()
            if os.path.exists(CROSSWALK):
                with open(CROSSWALK, encoding='utf-8-sig', newline='') as f:
                    barcodes = {r['Barcode'] for r in csv.DictReader(f)}
            entry['codes_equal_to_a_crosswalk_barcode'] = sum(1 for c in codes if c in barcodes)
            lines.append(f'Item export: {n} records, {entry["codes_equal_to_a_crosswalk_barcode"]} already carry a crosswalk barcode, '
                         f'{len(odd)} code(s) are not 8 digits.')
            if entry['codes_equal_to_a_crosswalk_barcode']:
                lines.append('  NOTE: some barcodes are already in BITPOS - this is not a clean "before" picture of the batch.')
        manifest['files'].append(entry)
        lines.append(f'{role}: {os.path.basename(src)} -> saved as {os.path.basename(dest)}\n   fingerprint (SHA-256): {entry["sha256"]}')
    with open(os.path.join(out_dir, 'SNAPSHOT_MANIFEST.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    with open(os.path.join(out_dir, 'SNAPSHOT_MANIFEST.txt'), 'w', encoding='utf-8') as f:
        f.write('BEFORE-BATCH SNAPSHOT  ' + manifest['taken_at'] + '\n\n' + '\n'.join(lines) + '\n')
    print('\n' + '\n'.join(lines))
    print(f'\nSaved in: {out_dir}')
    print('These copies are read-only. Keep this folder: it is your "before" record for Verify_Barcode_Batch.')
    return 0


if __name__ == '__main__':
    code = main()
    sys.exit(code)
