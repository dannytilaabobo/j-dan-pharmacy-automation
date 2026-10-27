"""Verify_Barcode_Batch.py - PASS or STOP verdict for a barcode batch. READ-ONLY on live data.

Inputs (all read as TEXT, never through Excel):
  1. the BEFORE item snapshot (from Snapshot_BITPOS_Before.py)
  2. the AFTER item export (exported from BITPOS after the batch was entered)
  3. the pilot batch CSV (Old_Item_Code, Barcode, ...)
  4. optionally a sales export from after the batch

Results go only into a new folder Barcode_Verification_<date_time>: VERIFY_RESULT.txt,
VERIFY_CHECKS.csv, VERIFY_DIFFERENCES.csv and VERIFY_RESULT.json.

Optional arguments (used by the tests):
  --before F --after F --batch F [--sales F] [--master F] [--crosswalk F] [--out-root DIR] [--declared-new N]
"""
import os
import sys
import csv
import json
import shutil
import hashlib
import tempfile
import argparse
import datetime
import warnings
from collections import Counter, defaultdict

warnings.filterwarnings('ignore')
os.chdir(os.path.dirname(os.path.abspath(__file__)))
PROJECT = os.getcwd()
sys.path.insert(0, PROJECT)
import item_code_resolver as ic   # noqa: E402

DOWNLOADS = os.path.join(os.path.expanduser('~'), 'Downloads')


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def read_export(path):
    """Header (stripped) and rows (tuples of raw text) of a BITPOS item export / price list."""
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        grid = list(csv.reader(f))
    header = [h.strip() for h in grid[0]]
    return header, [tuple(r) for r in grid[1:] if any(c.strip() for c in r)]


def read_batch(path):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        rows = list(csv.DictReader(f))
    out = []
    for r in rows:
        low = {k.strip().lower(): v for k, v in r.items() if k}
        out.append({'old': low.get('old_item_code', ''), 'barcode': low.get('barcode', ''), 'name': low.get('product_name', '')})
    return out


class Report:
    def __init__(self):
        self.checks = []
        self.diffs = []

    def add(self, key, title, result, detail=''):
        self.checks.append({'Check': key, 'Result': result, 'What_was_checked': title, 'Plain_English': detail})

    def diff(self, kind, code, name, detail):
        self.diffs.append({'Kind': kind, 'Item_Code': code, 'Product': name, 'Detail': detail})

    @property
    def stops(self):
        return [c for c in self.checks if c['Result'] == 'STOP']


def changed_columns(header, a, b, skip=()):
    out = []
    for i, h in enumerate(header):
        if h in skip:
            continue
        va = a[i] if i < len(a) else ''
        vb = b[i] if i < len(b) else ''
        if va != vb:
            out.append(f'{h}: {va!r} -> {vb!r}')
    if len(a) != len(b):
        out.append(f'row length {len(a)} -> {len(b)}')
    return out


def verify_items(rep, before_path, after_path, batch, xw, resolver, declared_new):
    hb, rb = read_export(before_path)
    ha, ra = read_export(after_path)
    if hb != ha:
        rep.add('1 same layout', 'BEFORE and AFTER files have the same columns', 'STOP',
                f'The two files have different columns (before {hb[:6]}..., after {ha[:6]}...). They are not the same kind of export.')
        return
    rep.add('1 same layout', 'BEFORE and AFTER files have the same columns', 'PASS', f'{len(hb)} columns.')
    up = [h.upper() for h in hb]
    ci = up.index('ITEM CODE')
    di = up.index('DESCRIPTION') if 'DESCRIPTION' in up else None
    name_of = lambda row: row[di] if di is not None and len(row) > di else ''
    code_of = lambda row: row[ci] if len(row) > ci else ''

    batch_old = {b['old']: b['barcode'] for b in batch}
    batch_bc = {b['barcode']: b['old'] for b in batch}
    before_by = defaultdict(list)
    after_by = defaultdict(list)
    for r in rb:
        before_by[code_of(r)].append(r)
    for r in ra:
        after_by[code_of(r)].append(r)

    # 2. row count
    delta = len(ra) - len(rb)
    rep.add('2 row count', f'number of records is unchanged (before {len(rb)}, after {len(ra)}, declared new items {declared_new})',
            'PASS' if delta == declared_new else 'STOP',
            'Same number of records.' if delta == declared_new else
            f'BITPOS now has {delta:+d} record(s) compared with before (you declared {declared_new}). An item may have been created, duplicated or removed.')

    # 3. batch items: old code gone, barcode present once, nothing else changed
    problems = []
    for b in batch:
        old, bc = b['old'], b['barcode']
        nb = len(before_by.get(old, []))
        if nb != 1:
            problems.append(f'{old} ({b["name"]}): found {nb} time(s) in the BEFORE snapshot, expected exactly once')
            rep.diff('batch item not in before', old, b['name'], f'{nb} record(s) with the old code before')
            continue
        old_after = len(after_by.get(old, []))
        if old_after:
            problems.append(f'{old} ({b["name"]}): the OLD code is still in BITPOS ({old_after} record) - BITPOS may have created a new record instead of changing the old one')
            rep.diff('old code still present', old, b['name'], f'{old_after} record(s) still have the old code')
        n_bc = len(after_by.get(bc, []))
        if n_bc == 0:
            problems.append(f'{old} ({b["name"]}): the barcode {bc} is NOT in BITPOS')
            rep.diff('barcode missing', old, b['name'], f'expected barcode {bc}')
            continue
        if n_bc > 1:
            problems.append(f'{old} ({b["name"]}): the barcode {bc} appears {n_bc} times')
            rep.diff('barcode duplicated', bc, b['name'], f'{n_bc} records')
            continue
        diffs = changed_columns(hb, before_by[old][0], after_by[bc][0], skip=('ITEM CODE',))
        if diffs:
            problems.append(f'{old} ({b["name"]}): something besides the code changed - ' + '; '.join(diffs))
            rep.diff('field changed on batch item', old, b['name'], '; '.join(diffs))
    rep.add('3 batch items', f'all {len(batch)} batch items now use their barcode, the old code is gone, and no other field changed',
            'PASS' if not problems else 'STOP',
            'Every batch item shows its barcode and is otherwise identical to before.' if not problems else ' | '.join(problems))

    # 4. everything outside the batch identical
    rest_before = Counter(r for r in rb if code_of(r) not in batch_old)
    rest_after = Counter(r for r in ra if code_of(r) not in batch_bc and code_of(r) not in batch_old)
    only_b, only_a = rest_before - rest_after, rest_after - rest_before
    codes_b = defaultdict(list)
    codes_a = defaultdict(list)
    for r in only_b.elements():
        codes_b[code_of(r)].append(r)
    for r in only_a.elements():
        codes_a[code_of(r)].append(r)
    msgs = []
    outside_barcode = []
    for code in sorted(set(codes_b) | set(codes_a)):
        if code in codes_b and code in codes_a:
            d = changed_columns(hb, codes_b[code][0], codes_a[code][0])
            msgs.append(f'{code} ({name_of(codes_b[code][0])}) changed: ' + '; '.join(d))
            rep.diff('item outside batch changed', code, name_of(codes_b[code][0]), '; '.join(d))
        elif code in codes_b:
            msgs.append(f'{code} ({name_of(codes_b[code][0])}) is MISSING from BITPOS now')
            rep.diff('item missing', code, name_of(codes_b[code][0]), 'present before, gone after')
        else:
            row = codes_a[code][0]
            if code in xw.by_barcode:
                msgs.append(f'barcode {code} ({name_of(row)}) was put on an item that is NOT in the batch (it belongs to {xw.by_barcode[code]["Old_Item_Code"]})')
                rep.diff('barcode on item outside batch', code, name_of(row), 'not in the pilot list')
                outside_barcode.append(code)
            else:
                msgs.append(f'{code} ({name_of(row)}) is a NEW record that was not there before')
                rep.diff('new record', code, name_of(row), 'not in before snapshot')
    rep.add('4 nothing else changed', 'every record outside the batch is identical to before (code, name, price and every other field)',
            'PASS' if not msgs else 'STOP',
            f'{sum(rest_before.values())} other records are identical to before.' if not msgs else ' | '.join(msgs[:12]) + (f' | ... and {len(msgs) - 12} more (see VERIFY_DIFFERENCES.csv)' if len(msgs) > 12 else ''))

    # 5. duplicates
    dup_msgs = []
    for code, rows in after_by.items():
        if len(rows) > 1 and (code in xw.by_barcode or len(before_by.get(code, [])) < len(rows)):
            dup_msgs.append(f'{code!r} appears {len(rows)} times ({name_of(rows[0])})')
            rep.diff('duplicate code', code, name_of(rows[0]), f'{len(rows)} records')
    rep.add('5 no duplicates', 'no barcode (and no code) appears twice',
            'PASS' if not dup_msgs else 'STOP', 'No duplicated barcode or code.' if not dup_msgs else ' | '.join(dup_msgs[:8]))

    # 6. damaged codes
    bad = []
    for code in after_by:
        if not code:
            continue
        problems_ = [p for p in ic.barcode_text_problems(code) if p.startswith(('scientific', 'decimal', 'leading/trailing', 'contains non-digit'))]
        st = resolver.resolve(code).status
        if st == ic.CORRUPTED or (problems_ and code not in before_by):
            bad.append(f'{code!r}: ' + (resolver.resolve(code).reason if st == ic.CORRUPTED else '; '.join(problems_)))
            rep.diff('damaged code', code, name_of(after_by[code][0]), bad[-1])
    rep.add('6 no damaged codes', 'no code looks damaged (4.80015E+12, ending in .0, spaces, lost leading zero)',
            'PASS' if not bad else 'STOP', 'No damaged code.' if not bad else ' | '.join(bad[:8]))

    # 7. resolver summary on the AFTER file (active records)
    inact = up.index('INACTIVE') if 'INACTIVE' in up else None
    counts = Counter()
    pre_existing, new_problems = [], []
    before_rows_set = set(rb)
    for r in ra:
        code = code_of(r)
        if inact is not None and len(r) > inact and r[inact].strip() == '1':
            continue
        if not code.strip() and inact is not None:
            continue                                  # the known malformed multi-line records
        res = resolver.resolve(code)
        counts[res.status] += 1
        if res.status in ic.PROBLEM_STATUSES:
            (pre_existing if r in before_rows_set else new_problems).append(f'{code!r} ({name_of(r)}): {res.status} - {res.reason}')
    translated = counts[ic.TRANSLATED]
    ok7 = translated == len(batch) and not new_problems
    detail = (f'{counts[ic.AS_IS]} resolved as-is, {translated} translated from barcode (expected exactly {len(batch)}), '
              f'{counts[ic.UNRESOLVED]} unresolved, {counts[ic.CORRUPTED]} corrupted, {counts[ic.CONFLICT]} conflicts, {counts[ic.EMPTY]} empty.')
    if translated != len(batch):
        detail += f' PROBLEM: {translated} codes were translated, not {len(batch)}.'
    if new_problems:
        detail += ' NEW problems caused by this export: ' + ' | '.join(new_problems[:6])
    rep.add('7 resolver summary', 'resolver summary of the AFTER file: 0 new unresolved / corrupted / conflicts and exactly the batch size translated',
            'PASS' if ok7 else 'STOP', detail)
    if pre_existing:
        rep.add('7b known old problems', 'codes that were already unresolved BEFORE the batch (not caused by it)', 'NOTE',
                ' | '.join(pre_existing[:6]) + ' - for example a new product that is in BITPOS but not yet in the Master. Not a barcode problem.')


def verify_sales(rep, sales_path, batch, master_df, resolver):
    import pandas as pd
    import Daily_Deduct_App as dd          # reads files only; the app window is not started
    from collections import Counter as C
    df = dd.read_pos_file(sales_path)
    valid_master = master_df.copy()
    valid_master['Current_Stock'] = pd.to_numeric(valid_master['Current_Stock'], errors='coerce').fillna(0)
    rc = ic.reconcile_sales(df[df['Qty'] > 0], valid_master, resolver)
    batch_old = {b['old']: b for b in batch}
    batch_bc = {b['barcode']: b for b in batch}
    problems, touched, units = [], 0, 0.0
    for ln in rc.lines:
        code = ln['code_raw'].strip()
        b = batch_bc.get(code) or batch_old.get(code)
        if not b:
            continue
        touched += 1
        units += ln['qty']
        if ln['category'] != ic.BOTH or ln['code_internal'] != b['old']:
            problems.append(f"{ln['name']} (code {code}, qty {ln['qty']:g}): {ln['category']} - {ln['note']}")
            rep.diff('sales line of batch item not matched', code, ln['name'], f"{ln['category']}: {ln['note']}")
    rep.add('8 batch sales lines', f'every sales line of a batch item resolved to the correct item and matched by code AND name ({touched} line(s), {units:g} unit(s))',
            'PASS' if not problems else 'STOP',
            ('All matched.' if touched else 'No batch item was sold in this file, so this check proves little. Sell one of them and export again.') if not problems else ' | '.join(problems[:8]))
    if not touched and not problems:
        rep.checks[-1]['Result'] = 'NOTE'
    bad = []
    if rc.count[ic.HELD_CONFLICT]:
        bad.append(f'{rc.count[ic.HELD_CONFLICT]} line(s) where the code and the name point to different items')
    if rc.corrupted_lines:
        bad.append(f'{rc.corrupted_lines} line(s) with a damaged or ambiguous code')
    if not rc.reconciles:
        bad.append('units do not balance')
    for ln in rc.of(ic.HELD_CONFLICT):
        rep.diff('code/name conflict in sales', ln['code_raw'].strip(), ln['name'], ln['note'])
    for ln in rc.lines:
        if ln['code_status'] == ic.CORRUPTED or ln['code_status'] == ic.CONFLICT:
            rep.diff('damaged code in sales', ln['code_raw'].strip(), ln['name'], ln['note'])
    rep.add('9 sales file clean', f'whole sales file: no conflicts, no damaged codes, units balance ({rc.units_file:g} units in file = {rc.units_to_deduct:g} to deduct + {rc.units_not_deducted:g} not deducted)',
            'PASS' if not bad else 'STOP', 'Clean.' if not bad else '; '.join(bad))
    if rc.count[ic.UNMATCHED]:
        lst = [f"{ln['name']} x{ln['qty']:g}" for ln in rc.of(ic.UNMATCHED)]
        rep.add('9b unmatched lines', 'sales lines with no matching Master item (these would need a YES at the real EOD)', 'NOTE', ' | '.join(lst[:8]))
        for ln in rc.of(ic.UNMATCHED):
            rep.diff('unmatched sales line', ln['code_raw'].strip(), ln['name'], ln['note'])

    # deductions identical to what the old codes would have produced
    def run(frame):
        r = ic.reconcile_sales(frame[frame['Qty'] > 0], valid_master, resolver)
        v = frame[frame['Qty'] > 0].copy()
        if r.held_indices:
            v = v.drop(index=r.held_indices)
        ded = v.groupby('Description')['Qty'].sum().to_dict()
        return dd.apply_fifo_deduction(valid_master.copy(), ded)

    df_old = df.copy()
    df_old['Item Code'] = [resolver.internal_or_raw(c) for c in df['Item Code']]
    a, b = run(df), run(df_old)
    same = list(a['Current_Stock'].round(6)) == list(b['Current_Stock'].round(6))
    rep.add('10 same deductions', 'stock deducted with the codes as BITPOS shows them is identical (every lot) to stock deducted with every code replaced by its old code',
            'PASS' if same else 'STOP', 'Identical.' if same else 'The two runs removed different quantities - do not process this EOD.')


def pick_files(args):
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    snaps = sorted([p for p in os.listdir(PROJECT) if p.startswith('Barcode_Verification_')], reverse=True)
    init_before = os.path.join(PROJECT, snaps[0], 'snapshots') if snaps else DOWNLOADS
    if not args.before:
        messagebox.showinfo('Verify barcode batch', 'Step 1 of 4: choose the BEFORE item snapshot (made by Snapshot_BITPOS_Before).')
        args.before = filedialog.askopenfilename(title='BEFORE item snapshot', initialdir=init_before, filetypes=[('CSV', '*.csv')])
    if not args.after and args.before:
        messagebox.showinfo('Verify barcode batch', 'Step 2 of 4: choose the AFTER item export (exported from BITPOS after the batch).')
        args.after = filedialog.askopenfilename(title='AFTER item export', initialdir=DOWNLOADS, filetypes=[('CSV', '*.csv')])
    if not args.batch and args.after:
        pilots = sorted([p for p in os.listdir(PROJECT) if p.startswith('Pilot_Batch_Prep_')], reverse=True)
        messagebox.showinfo('Verify barcode batch', 'Step 3 of 4: choose the pilot batch CSV (PILOT_BATCH_15_FINAL_DRAFT.csv).')
        args.batch = filedialog.askopenfilename(title='Pilot batch CSV', initialdir=os.path.join(PROJECT, pilots[0]) if pilots else PROJECT, filetypes=[('CSV', '*.csv')])
    if args.batch and args.sales is None:
        if messagebox.askyesno('Verify barcode batch', 'Step 4 of 4: do you have a SALES export from after the batch to check as well?'):
            args.sales = filedialog.askopenfilename(title='Sales export (do not open it in Excel)', initialdir=PROJECT, filetypes=[('Sales exports', '*.xls *.xlsx *.csv')])
        else:
            args.sales = ''
    root.destroy()


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--before')
    ap.add_argument('--after')
    ap.add_argument('--batch')
    ap.add_argument('--sales')
    ap.add_argument('--master', default=os.path.join(PROJECT, 'Inventory_Master.xlsx'))
    ap.add_argument('--crosswalk', default=ic.CROSSWALK_FILE)
    ap.add_argument('--out-root', default=PROJECT)
    ap.add_argument('--declared-new', type=int, default=0)
    args = ap.parse_args(argv)
    if not (args.before and args.after and args.batch):
        pick_files(args)
    if not (args.before and args.after and args.batch):
        print('A file was not chosen. Nothing was checked.')
        return 2

    print('=' * 70)
    print(' VERIFY BARCODE BATCH (read-only)')
    print('=' * 70)
    import pandas as pd
    tmp = tempfile.mkdtemp(prefix='verify_master_copy_')
    master_copy = os.path.join(tmp, 'Inventory_Master_COPY.xlsx')
    shutil.copy2(args.master, master_copy)               # the live Master is only copied, never opened for writing
    master_df = pd.read_excel(master_copy, sheet_name='INVENTORY MASTER', dtype=str, keep_default_na=False)
    xw = ic.load_crosswalk(args.crosswalk)
    resolver = ic.Resolver(xw, set(master_df['Item_Code'].str.strip()))
    batch = read_batch(args.batch)
    rep = Report()

    # batch file sanity
    probs = []
    for b in batch:
        row = xw.by_old.get(b['old'])
        tp = ic.barcode_text_problems(b['barcode'])
        if tp:
            probs.append(f"{b['old']}: barcode {b['barcode']!r} is damaged ({'; '.join(tp)})")
        elif not row or row['Barcode'] != b['barcode']:
            probs.append(f"{b['old']}: the pilot file says {b['barcode']} but the crosswalk says {row['Barcode'] if row else 'nothing'}")
    if len({b['barcode'] for b in batch}) != len(batch) or len({b['old'] for b in batch}) != len(batch):
        probs.append('a barcode or an old code is listed twice in the pilot file')
    rep.add('0 pilot list', f'the pilot file ({len(batch)} items) agrees with Barcode_Crosswalk.csv and its barcodes are clean text',
            'PASS' if not probs and batch else 'STOP', 'Agrees.' if not probs and batch else (' | '.join(probs) if probs else 'The pilot file is empty.'))
    if batch and not probs:
        verify_items(rep, args.before, args.after, batch, xw, resolver, args.declared_new)
        if args.sales:
            try:
                verify_sales(rep, args.sales, batch, master_df, resolver)
            except Exception as e:   # noqa: BLE001
                rep.add('8 sales check', 'the sales export could be read and checked', 'STOP', f'The sales file could not be checked: {e}')
    shutil.rmtree(tmp, ignore_errors=True)

    verdict = 'PASS' if not rep.stops else 'STOP'
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    out_dir = os.path.join(args.out_root, f'Barcode_Verification_{stamp}')
    n = 2
    while os.path.exists(out_dir):
        out_dir = os.path.join(args.out_root, f'Barcode_Verification_{stamp}_{n}')
        n += 1
    os.makedirs(out_dir)
    lines = [f'VERIFY BARCODE BATCH  {datetime.datetime.now():%Y-%m-%d %H:%M}', '',
             f'BEFORE snapshot : {args.before}', f'AFTER export    : {args.after}', f'Pilot list      : {args.batch}',
             f'Sales export    : {args.sales or "(none supplied)"}', '']
    for c in rep.checks:
        lines.append(f"[{c['Result']:4}] {c['Check']} - {c['What_was_checked']}")
        lines.append(f"        {c['Plain_English']}")
    lines += ['', '=' * 70, f'VERDICT: {verdict}']
    if verdict == 'STOP':
        lines.append('Do NOT continue. Reasons:')
        for c in rep.stops:
            lines.append(f"  - {c['What_was_checked']}: {c['Plain_English']}")
        lines.append('See STOP_WHAT_TO_DO in PILOT_PREP_REPORT.md. Nothing was changed by this tool.')
    else:
        lines.append('All checks passed. The batch is exactly as planned.' + ('' if args.sales else ' (No sales export was checked yet.)'))
    lines.append('=' * 70)
    text = '\n'.join(lines)
    print(text)
    with open(os.path.join(out_dir, 'VERIFY_RESULT.txt'), 'w', encoding='utf-8') as f:
        f.write(text)
    with open(os.path.join(out_dir, 'VERIFY_CHECKS.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=['Check', 'Result', 'What_was_checked', 'Plain_English'], quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(rep.checks)
    with open(os.path.join(out_dir, 'VERIFY_DIFFERENCES.csv'), 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=['Kind', 'Item_Code', 'Product', 'Detail'], quoting=csv.QUOTE_ALL)
        w.writeheader()
        w.writerows(rep.diffs)
    result = {'verdict': verdict, 'created': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
              'batch_items': [{'old': b['old'], 'barcode': b['barcode']} for b in batch],
              'sales_checked': bool(args.sales), 'stop_reasons': [c['What_was_checked'] for c in rep.stops],
              'crosswalk_sha256': sha256(args.crosswalk) if os.path.exists(args.crosswalk) else '',
              'before_sha256': sha256(args.before), 'after_sha256': sha256(args.after), 'batch_sha256': sha256(args.batch),
              'sales_sha256': sha256(args.sales) if args.sales else ''}
    with open(os.path.join(out_dir, 'VERIFY_RESULT.json'), 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2)
    print(f'\nResults saved in: {out_dir}')
    return 0 if verdict == 'PASS' else 1


if __name__ == '__main__':
    sys.exit(main())
