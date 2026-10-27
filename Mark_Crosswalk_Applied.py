"""Mark_Crosswalk_Applied.py - records in Barcode_Crosswalk.csv that a verified batch is live in BITPOS.

It is a DRY RUN until you type YES. It refuses to run unless the latest Verify_Barcode_Batch result is PASS
for every item it would change. Before any change it makes a dated backup of the crosswalk and afterwards
it writes a line to JDan_System_Logs.txt.

IMPORTANT about the status name: the resolver only understands the statuses PLANNED, ACTIVE and REVERTED.
A row with any other status text (for example APPLIED_IN_BITPOS) would be IGNORED by the resolver and its
barcode would stop being translated. So "applied in BITPOS" is recorded as ACTIVE. The status is otherwise
a record only: PLANNED and ACTIVE rows are translated in exactly the same way.

Optional arguments (used by the tests): --crosswalk F --verify-root DIR --verify-json F --log F --yes-text YES
"""
import os
import sys
import csv
import glob
import json
import hashlib
import argparse

os.chdir(os.path.dirname(os.path.abspath(__file__)))
PROJECT = os.getcwd()
sys.path.insert(0, PROJECT)
import item_code_resolver as ic   # noqa: E402

NEW_STATUS = 'ACTIVE'


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def latest_verify_json(root):
    found = glob.glob(os.path.join(root, 'Barcode_Verification_*', 'VERIFY_RESULT.json'))
    return max(found, key=os.path.getmtime) if found else None


def main(argv=None, input_fn=input):
    ap = argparse.ArgumentParser()
    ap.add_argument('--crosswalk', default=ic.CROSSWALK_FILE)
    ap.add_argument('--verify-root', default=PROJECT)
    ap.add_argument('--verify-json')
    ap.add_argument('--log')
    args = ap.parse_args(argv)
    if args.log:
        ic.LOG_FILE = args.log

    print('=' * 70)
    print(' MARK CROSSWALK APPLIED  (nothing changes until you type YES)')
    print('=' * 70)
    vj = args.verify_json or latest_verify_json(args.verify_root)
    if not vj:
        print('\nREFUSED: no Verify_Barcode_Batch result was found. Run VERIFY BARCODE BATCH first.')
        return 2
    with open(vj, encoding='utf-8') as f:
        v = json.load(f)
    print(f'\nLatest verification: {vj}\n  verdict {v["verdict"]}, made {v["created"]}, {len(v["batch_items"])} items, '
          f'sales export checked: {"yes" if v["sales_checked"] else "NO"}')
    if v['verdict'] != 'PASS':
        print('\nREFUSED: the latest verification is STOP, not PASS. Fix the problems listed in VERIFY_RESULT.txt and verify again.')
        return 2
    if not os.path.exists(args.crosswalk):
        print('\nREFUSED: Barcode_Crosswalk.csv was not found.')
        return 2
    with open(args.crosswalk, encoding='utf-8-sig', newline='') as f:
        rows = {r['Old_Item_Code']: r for r in csv.DictReader(f)}
    todo, already, missing = [], [], []
    for it in v['batch_items']:
        r = rows.get(it['old'])
        if r is None or r['Barcode'] != it['barcode']:
            missing.append(it['old'])
        elif r['Status'] == NEW_STATUS:
            already.append(it['old'])
        else:
            todo.append(it['old'])
    if missing:
        print(f'\nREFUSED: these items are not in the crosswalk with the verified barcode: {missing}')
        return 2
    if not todo:
        print(f'\nNothing to change: every verified item is already {NEW_STATUS}.')
        return 0
    if v.get('crosswalk_sha256') and v['crosswalk_sha256'] != sha256(args.crosswalk):
        print('\nREFUSED: Barcode_Crosswalk.csv has changed since the verification was made. Verify the batch again.')
        return 2
    print(f'\nRows that would change from PLANNED/REVERTED to {NEW_STATUS}: {len(todo)}')
    for c in todo:
        print(f'   {c}  {rows[c]["Barcode"]}  {rows[c]["Product_Name"]}')
    if already:
        print(f'Already {NEW_STATUS} (no change): {len(already)}')
    if not v['sales_checked']:
        print('\nWARNING: the verification did not include a sales export. The plan is to mark items applied only after the first clean EOD.')
    if not todo:
        print('\nNothing to change.')
        return 0
    ans = input_fn(f'\nThis was a DRY RUN. To really change these {len(todo)} rows, type YES (anything else leaves everything unchanged): ').strip()
    if ans != 'YES':
        print('\nNothing was changed.')
        return 0
    changed, backup = ic.set_crosswalk_status(todo, NEW_STATUS, args.crosswalk)
    cw = ic.load_crosswalk(args.crosswalk)
    bad = [c for c in todo if cw.by_old.get(c) is None or cw.by_old[c]['Status'] != NEW_STATUS or cw.by_old[c]['Barcode'] not in cw.by_barcode]
    if bad or cw.problems:
        import shutil
        shutil.copy2(backup, args.crosswalk)
        print(f'\nPROBLEM after the change ({bad or cw.problems[:3]}). The backup was restored; the crosswalk is as it was before.')
        ic.log_event(f'MARK CROSSWALK APPLIED FAILED and was rolled back (backup {os.path.basename(backup)})')
        return 1
    ic.log_event(f'MARK CROSSWALK APPLIED: {changed} item(s) set to {NEW_STATUS} after a PASS verification ({os.path.basename(os.path.dirname(vj))}); '
                 f'backup {os.path.basename(backup)}; items {",".join(todo)}')
    print(f'\nDONE: {changed} row(s) set to {NEW_STATUS}. Backup of the crosswalk: {backup}')
    print('The resolver was re-read and every marked barcode still translates.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
