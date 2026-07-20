# J-DAN / Happy Med Pharmacy Automation — Master Notes v2

Living reference for the whole `Pharmacy_Automation` system. Kept inside the
repo (not just in chat) so any future session — Claude or human — can pick
up work without re-deriving history. Update this file whenever a system
changes status, a bug gets fixed, or a priority shifts.

Separate, unrelated repo: `Pharmacy_System` (staff timekeeping/payroll,
Flask + SQLite) has its own history — not covered here except where the two
intersect (shared identity/security items are noted below).

---

## 1. Systems Currently Built and Working

| Script / App | Purpose | Status |
|---|---|---|
| `StockIn_App.py` | Encode incoming deliveries (Item_Code, Lot, Expiry, Qty) into Master | Working. Has a confirmation prompt for any single quantity over 5,000 (typo guard) and preventive product-name cleaning. |
| `Daily_Deduct_App.py` | Applies POS EOD sales as FIFO/FEFO stock deductions against Master | Working. Matches sales to Master strictly by `Item_Code`, never by name. |
| `Daily_Dashboard_App.py` | SOD (start-of-day) physical count stratifier/reconciliation | Working. Since 2026-07-18 it also **writes the corrected count back to Master** (previously it only logged variance without correcting the system-of-record — see bug log below). |
| `Expiry_Scanner.py` | Flags near-expiry / expired lots | Working. |
| `Restock_Check_App.py` | Flags products at/below `Reorder_Point`, groups by `Supplier`, suggests order qty | Working. First run flags ~73% of products "low stock" because `Reorder_Point` is still a flat default (10) for most items — expected, not a bug, until Supplier/Reorder_Point get customized per item. |
| `Item_Settings_App.py` | Primary, durable editor for `Category` / `Supplier` / `Reorder_Point` / `Is_Generic` | Working. Built because the Excel dropdown itself is structurally fragile (see below) — this app is unaffected by that fragility. Has a progress readout, priority-by-sales-velocity ordering, search box, and batches edits into one `SAVE ALL CHANGES` write. |
| `Add_Reorder_And_Supplier_Columns.py` | Schema/dropdown maintenance for Master (Reorder_Point, Supplier, Category, Is_Generic columns + Excel dropdowns) | Working. Re-run any time the dropdown needs reattaching (it gets wiped by any other script's full-sheet rewrite — known, accepted limitation, see below). |
| `Sync_BITPOS_Prices.py` | Syncs Master against the latest BITPOS export | Working. Detects **4 categories**: price changes, name changes, discontinued items (in Master, gone from BITPOS), and new items (in BITPOS, not in Master). Always shows a full preview before applying; price/discontinued/new-items get one blanket confirmation, each name change gets its own individual confirmation (renames are riskier — sales matching depends on Item_Code, not name, but a wrong rename still misleads staff). |
| `Review_Discontinued_Items.py` (+ `.bat`) | Reviews DISCONTINUED items, splits into "safe to remove" (0 stock) vs "still has stock," removes only with explicit confirmation | Working, new 2026-07-20. |
| `Mark_Discontinued_Items.py` | Manually mark specific Item_Codes as DISCONTINUED (hardcoded dict + reason) | Working — used for supplier repacks/replacements. |
| `Scan_Name_Inconsistencies.py` | Read-only sweep for Item_Code/Product_Name mismatches across all of Master | Working. Safe to re-run any time (e.g., after bulk stock-ins or BITPOS syncs). |
| `name_cleaning.py` | Shared text-cleaning helpers (`clean_text`, `is_date_corrupted`) | Working, used by StockIn/Deduct/BITPOS sync. |
| `Auto_Cloud_Sync.py` | Syncs to Google Sheets | Working (uses a service-account key that still needs rotation — see Priorities). |
| `Diagnose_Inventory_Mismatch.py` | Diagnostic over SOD reports to find repeat-offender products | Working, read-only. |
| `Cost_Tracker_App.py` | Cost/margin tracker, isolated from staff-facing apps | **Design approved, NOT YET BUILT.** This is Step 2 of a two-tool plan — do not start until explicitly told to; Item_Settings_App (Step 1) needed real per-item completion first. |
| `Admin_Audit_App.py` | "Stealth investigator / zero-movement theft trap" referenced in the user's own notes | **Not found on this PC.** May live on a different machine — unconfirmed. Don't assume it runs from here. |
| `Desktop\dist\app_ui.exe`, `Desktop\dist\sync_worker.exe` | Compiled binaries referenced elsewhere | **No source available in this environment** — cannot be patched or audited without it. |

### Master schema (`Inventory_Master.xlsx`, sheet `INVENTORY MASTER`)
Core columns: `Item_Code` (only reliable matching key — never Product_Name), `Product_Name`, `Retail_Price`, `Current_Stock`, `Lot_Number`, `Expiry_Date`, `Date_Encoded`, `Encoded_By`, `Pulled_Out_By`, `LOCATION`, `Status` (ACTIVE/DISCONTINUED), `Discontinued_Date`, `Discontinued_Reason`, `Reorder_Point`, `Supplier`, `Category`, `Is_Generic`.

**Category list — 17 items (replaced the old 14-item list 2026-07-19/20, clean replacement, no migration needed):**
Vitamins/Minerals & Supplements, Pain/Fever & Inflammation, Cough/Cold/Allergy & Asthma, Gastrointestinal (GI) Medicines, Antibiotics & Anti-infectives, Cardiovascular Medicines, Antidiabetic Medicines, Hormones & Steroids, Central Nervous System (CNS) Medicines, Topical Preparations (Creams/Ointments/Gels & Sprays), Eye/Ear & Nasal Preparations, Family Planning & Women's Health, Medical Supplies & Devices, Personal Care/Baby Care & Feminine Care, Beverages, General Merchandise, Others/Uncategorized.

**Every script that writes to Master backs it up first** to `JDan_Master_Backups/` (timestamped). This is a hard convention — follow it in any new script.

---

## 2. Bugs Found and Fixed (chronological)

- **2026-07-08 — Security audit baseline.** Found plaintext `credentials.json` files and a hardcoded admin password. Git set up with a strict code-only `.gitignore` (no secrets, no data, no reports) as the first fix. Rotation/removal of the actual secrets is still open (see Priorities).
- **2026-07-12 — Corrupted product name.** One HTML-entity-corrupted name found and fixed; preventive cleaning (`name_cleaning.py`) wired into StockIn/Deduct apps so it can't silently recur.
- **2026-07-18 — SOD reconciliation never wrote back to Master.** `Daily_Dashboard_App.py` logged SOD variance but never corrected `Current_Stock` in Master — a likely root cause of repeat mismatches. Fixed: reconciliation now re-reads Master and applies the corrected count via FIFO/FEFO lot distribution, backs up first, logs every adjustment.
- **2026-07-18 — Excel dropdown "Removed Feature" corruption (no data lost).** Caused by re-running `Add_Reorder_And_Supplier_Columns.py` multiple times without clearing previous data-validation definitions — they piled up until Excel flagged the sheet as corrupt and stripped the feature (cell data untouched). Fixed by clearing validations before re-adding each run; confirmed idempotent.
- **2026-07-18 — `FileNotFoundError` / cwd landing in `C:\Windows\System32`.** Root cause: launch-time working directory wasn't reliable across shortcuts/elevation/schedulers. Fixed by adding `os.chdir(os.path.dirname(os.path.abspath(__file__)))` to all path-dependent scripts (18 scripts in this repo, plus `Pharmacy_System/app.py` separately).
- **2026-07-18 — Instant `.bat` crash (flash and close).** Root cause: bare `python`/`py` resolved to the Microsoft Store execution-alias stub, not the real interpreter, and printed its install-nudge message then exited before any `pause` could show it. Fixed in all `.bat` launchers (both repos) by hardcoding the real interpreter path and adding `pause`.
- **2026-07-18 — Per-item "SAVE" button confused with "SAVE ALL CHANGES."** A real edit appeared to vanish because the per-item button only staged it in memory. Fixed: renamed to "STAGE THIS EDIT (not saved to file yet)," recolored teal to visually distinguish it.
- **2026-07-18 — `TypeError: Invalid value ... for dtype 'float64'`.** All-blank Supplier/Category/Is_Generic columns get read back as `float64`; assigning a string then crashed. Fixed with `.astype('object')` before assignment.
- **2026-07-18 — Grouping bug: must key by Item_Code, never Product_Name.** A typo'd duplicate name ("UDCA" vs "UCDA") crashed a Treeview insert. Fixed in all 3 aggregating scripts to group strictly by `Item_Code`.
- **2026-07-18 — BAND AID PLASTER 240,522-unit outlier.** Traced to a single stock-in typo (240,326 instead of ~200). Corrected via FIFO/FEFO lot adjustment; added a 5,000-unit confirmation prompt to `StockIn_App.py` to catch this class of typo earlier.
- **2026-07-19 — Timekeeping `ERR_CONNECTION_REFUSED` despite healthy ngrok tunnel.** Same Microsoft Store python-stub root cause, but `Pharmacy_System`'s `.bat` had been missed in the first pass. Fixed and verified live (HTTP 200).
- **2026-07-19 — Payslip Gross Pay not reflecting the new ₱80/hour rate.** Found only by testing the live page, not code review: `templates/payslip.html` computes Gross Pay via client-side JavaScript reading its own hardcoded rate input, completely independent of the server-side value. Fixed by passing the real rate into the template.
- **2026-07-19/20 — Excel 255-character dropdown limit.** The new 17-item Category list joins to ~498 characters, over Excel's inline-list validation limit — would have caused the same corruption as the July 18 bug, but from length this time. Fixed by sourcing all 3 dropdowns (Supplier/Category/Is_Generic) from a hidden `DropdownLists` worksheet via named ranges instead of inline strings. Verified idempotent (2 runs, no growth in validations/names).
- **2026-07-20 — Physical-count discrepancy investigated (BENCH BODY SPRAY EIGHT 100ML, Item_Code `61800403`): 11 in Master, 10 counted on shelf.** Full trace of every stock-in, sale, and pull-out record for this item found **zero gap in the system's own logs** — 1 stock-in (+10, 2026-07-17, by KAT) and 5 sales (-5 total, cashiers angelique/DANNY/katlyn) reconcile exactly to Master's current 11. No pull-out record, no manual override, no dropped sales record. **Conclusion: the missing unit happened outside all logged channels** — most likely either a receiving/count error at the 2026-07-17 stock-in itself, or unlogged shrinkage afterward. **Not yet resolved** — recommended next step: physically recount the 10 other items that shared the same delivery lot (`4321224`, same 2026-07-17 batch) to tell whether this is an isolated shrinkage case or a whole-delivery undercount.

---

## 3. Remaining Priorities / TODO

**Security (open since 2026-07-08 audit — highest priority, not yet done):**
1. Move `Pharmacy_System/app.py`'s hardcoded `ADMIN_PASSWORD` out of source (env var or hashed config file).
2. Rotate service-account key `jdan-sheets-updater@jdan-pharmacy-automation` (used by `Auto_Cloud_Sync.py`).
3. Rotate service-account key `timekeeping-bot@synthetic-song-499811-e9` (used by `Pharmacy_System/app.py`).
4. Verify/deactivate staff PINs in the timekeeping system for departed staff.

**Data completeness:**
- `Category` and `Supplier` are still blank for the large majority of the 1,610+ active products — needs real per-item completion via `Item_Settings_App.py` before Restock_Check's supplier grouping or any Category-based reporting is trustworthy.
- CETIRIZINE 10MG TAB (Item_Code `65789667`) — user was going to fill this in themselves via the now-fixed app; verify it actually got saved next time it comes up.

**Open investigation:**
- BENCH BODY SPRAY EIGHT 100ML (`61800403`) 1-unit shortfall — see bug log above. Recount the other 10 items from Lot `4321224` to narrow down cause.
- BITPOS pending rename: Item_Code `51881290` — zero stock, can't verify until next stock-in of that code (tracked in `JDan_Executive_Reports/BITPOS_PENDING_NAME_VERIFICATION.csv`).
- Two known DISCONTINUED items (SILVER SULFADIAZINE 20G / `37776367`, FUSIDIC ACID TOPICAL CRM 5G / `77925626`) still carry 5 units each — not yet eligible for removal via `Review_Discontinued_Items.py` until sold through or returned.

**Planned but not started:**
- `Cost_Tracker_App.py` (Step 2) — approved design, isolated read-only access to Master, hashed password, cost/margin tracking. Build only when explicitly requested.
- Excel dropdown fragility is an accepted limitation, not fixed at the root — would require switching every Master-writing script from full-sheet rewrites to in-place cell edits (bigger refactor, not scheduled).
- Supplier Price Variance Engine, WFH Executive Hub, Google Drive Bridge — all still just planned, no work started.

**Resolved questions:**
- Physical location: **Malate** (confirmed 2026-07-20; "Caloocan" seen in some exports was incorrect/stale).
- Collaboration model: **local-folder** — Claude Code on desktop directly accessing this folder (confirmed 2026-07-20).

---

*Last updated: 2026-07-20. Update this file (don't just rely on chat) whenever a system's status changes, a bug gets fixed, or a priority is resolved/added.*
