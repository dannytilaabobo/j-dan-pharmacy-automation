import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import datetime
import os
import glob
import shutil
import sys
import traceback

from item_code_resolver import resolve_sales_history

# Anchor to this script's own folder, not the launch-time working directory
# (which can end up somewhere unexpected -- e.g. C:\Windows\System32 -- if
# launched via a shortcut, elevation, or scheduler that doesn't set cwd).
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ==========================================
#    CRASH DETECTOR (Anti-Silent Crash)
# ==========================================
def handle_exception(exc_type, exc_value, exc_traceback):
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    messagebox.showerror("System Critical Error", f"The system encountered an error. Please screenshot this and send it to the Admin:\n\n{err_msg}")

sys.excepthook = handle_exception

# ==========================================
#    SECURITY LOGIN & LOGBOOK
# ==========================================
root = tk.Tk()
root.withdraw()

staff_name = simpledialog.askstring("J-DAN PHARMACY - SECURITY", "Enter Authorized User Name:", parent=root)
if not staff_name or staff_name.strip() == "":
    messagebox.showerror("Access Denied", "Authorized user name is required to access Item Settings.")
    sys.exit()

current_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
try:
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{current_time}] - USER: {staff_name.upper()} - Accessed ITEM SETTINGS (Category/Supplier/Reorder_Point editor)\n")
except Exception:
    pass

root.deiconify()

# ==========================================
#    SYSTEM CONFIGURATION
# ==========================================
MASTER_FILE = 'Inventory_Master.xlsx'
BACKUP_DIR = 'JDan_Master_Backups'
EOD_DIR = 'JDan_Sales_System/EOD_Deduction_Reports'

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

# Medicine-only attribute: blank/N-A is valid and expected for non-medicine
# items (General Merchandise, Personal Care, etc.) -- never forced.
IS_GENERIC_OPTIONS = ['(N/A)', 'GENERIC', 'BRANDED']


def norm(s):
    return str(s).strip().upper()


code_check = {'text': ''}  # Item Code check result of the last load_sales_velocity() call


def load_sales_velocity():
    frames = []
    for path in glob.glob(f'{EOD_DIR}/PROCESSED_EOD_*.xlsx'):
        try:
            df = pd.read_excel(path, dtype={'Item Code': str})
        except Exception:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Total_Qty_Sold'])
    all_sales = pd.concat(frames, ignore_index=True)
    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    # BITPOS codes are translated to the internal Item_Code first (barcode -> old code).
    # Unresolved / corrupted / empty codes are kept as typed, counted and reported -- never dropped.
    all_sales['Item_Code'], code_check['text'] = resolve_sales_history(all_sales, 'Item Settings', show=None)
    return all_sales.groupby('Item_Code').agg(Total_Qty_Sold=('Qty', 'sum')).reset_index()


def load_working_data():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype={'Item_Code': str})
    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED']
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    if 'Reorder_Point' not in df.columns:
        df['Reorder_Point'] = 10
    df['Reorder_Point'] = pd.to_numeric(df['Reorder_Point'], errors='coerce').fillna(10)
    if 'Supplier' not in df.columns:
        df['Supplier'] = ''
    df['Supplier'] = df['Supplier'].fillna('').astype(str).str.strip()
    if 'Category' not in df.columns:
        df['Category'] = ''
    df['Category'] = df['Category'].fillna('').astype(str).str.strip()
    if 'Is_Generic' not in df.columns:
        df['Is_Generic'] = ''
    df['Is_Generic'] = df['Is_Generic'].fillna('').astype(str).str.strip().str.upper()

    # Item_Code is the only stable identifier -- Product_Name can drift
    # (typos, spelling variants) across a code's lot rows, so grouping by
    # name would silently split one product into several (and can crash the
    # Treeview, which requires a unique iid per row). Pick the most frequent
    # spelling per code as the representative name, and flag codes where the
    # spelling actually disagrees so it can be corrected in Master.
    df['_name_clean'] = df['Product_Name'].astype(str).str.strip()
    name_variant_counts = df.groupby('Item_Code')['_name_clean'].nunique()
    inconsistent_codes = set(name_variant_counts[name_variant_counts > 1].index)

    name_freq = df.groupby(['Item_Code', '_name_clean']).size().reset_index(name='freq')
    name_freq = name_freq.sort_values(['Item_Code', 'freq'], ascending=[True, False])
    representative_name = name_freq.drop_duplicates('Item_Code', keep='first').set_index('Item_Code')['_name_clean']

    agg = df.groupby('Item_Code').agg(
        Current_Stock=('Current_Stock', 'sum'),
        Reorder_Point=('Reorder_Point', 'first'),
        Supplier=('Supplier', 'first'),
        Category=('Category', 'first'),
        Is_Generic=('Is_Generic', 'first'),
    ).reset_index()
    agg['Product_Name'] = agg['Item_Code'].map(representative_name)
    agg['Name_Inconsistent'] = agg['Item_Code'].isin(inconsistent_codes)

    velocity = load_sales_velocity()
    agg = agg.merge(velocity, on='Item_Code', how='left')
    agg['Total_Qty_Sold'] = agg['Total_Qty_Sold'].fillna(0)
    agg = agg.sort_values('Total_Qty_Sold', ascending=False).reset_index(drop=True)
    agg['Rank'] = agg.index + 1
    return agg


try:
    working_df = load_working_data()
except FileNotFoundError:
    messagebox.showerror(
        "File Not Found",
        f"Cannot find '{MASTER_FILE}' in this folder:\n{os.getcwd()}\n\n"
        f"Please make sure Item_Settings_App.py (or its .bat launcher) is run from inside "
        f"the Pharmacy_Automation folder, and that {MASTER_FILE} has not been moved, renamed, or deleted."
    )
    sys.exit()
except PermissionError:
    messagebox.showerror(
        "File In Use",
        f"Cannot open '{MASTER_FILE}' -- it may be open in Excel or locked by another program.\n\n"
        f"Please close it and try again."
    )
    sys.exit()

if code_check['text']:
    messagebox.showwarning("Item Code Check", code_check['text'])

pending_edits = {}  # Item_Code -> {'Reorder_Point': x, 'Supplier': y, 'Category': z, 'Is_Generic': w}


def refresh_progress_label():
    total = len(working_df)
    supplier_filled = (working_df['Supplier'].astype(str).str.strip() != '').sum()
    category_filled = (working_df['Category'].astype(str).str.strip() != '').sum()
    fully_set = ((working_df['Supplier'].astype(str).str.strip() != '') & (working_df['Category'].astype(str).str.strip() != '')).sum()
    pending_count = len(pending_edits)
    progress_label.config(
        text=f"Supplier: {supplier_filled}/{total}   |   Category: {category_filled}/{total}   |   "
             f"Complete (both Category+Supplier): {fully_set}/{total}"
             + (f"   |   {pending_count} unsaved edit(s)" if pending_count else "")
    )
    save_all_btn.config(state='normal' if pending_edits else 'disabled')


def matches_search(row, query):
    if not query:
        return True
    query = query.strip().upper()
    return query in str(row['Product_Name']).upper() or query in str(row['Item_Code'])


def populate_tree():
    tree.delete(*tree.get_children())
    query = search_var.get()
    name_conflicts = []
    for _, row in working_df.iterrows():
        if not matches_search(row, query):
            continue
        supplier_display = row['Supplier'] if row['Supplier'] else '(blank)'
        category_display = row['Category'] if row['Category'] else '(blank)'
        is_generic_display = row['Is_Generic'] if row['Is_Generic'] else '(N/A)'
        product_display = row['Product_Name']
        if row['Name_Inconsistent']:
            product_display = f"⚠️ {product_display} (inconsistent spelling in Master)"
            name_conflicts.append((row['Item_Code'], row['Product_Name']))
        is_unassigned = not row['Supplier'] or not row['Category']
        tag = 'inconsistent' if row['Name_Inconsistent'] else ('unassigned' if is_unassigned else '')
        tree.insert('', tk.END, iid=str(row['Item_Code']), values=(
            int(row['Rank']), row['Item_Code'], product_display,
            f"{row['Total_Qty_Sold']:.0f}", f"{row['Current_Stock']:.0f}",
            f"{row['Reorder_Point']:.0f}", category_display, supplier_display, is_generic_display,
        ), tags=(tag,) if tag else ())
    refresh_progress_label()
    return name_conflicts


def open_editor(item_code=None):
    if item_code is None:
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an item from the list first.")
            return
        item_code = selected[0]

    row_match = working_df[working_df['Item_Code'].astype(str) == str(item_code)]
    if row_match.empty:
        return
    row = row_match.iloc[0]

    win = tk.Toplevel(root)
    win.title(f"Edit: {row['Product_Name']}")
    win.geometry("500x400")
    win.configure(bg="#eaeded")
    win.grab_set()

    tk.Label(win, text=row['Product_Name'], font=("Arial", 12, "bold"), bg="#eaeded", wraplength=460, justify="left").pack(pady=(15, 5), padx=15)
    tk.Label(win, text=f"Item Code: {row['Item_Code']}  |  Current Stock: {row['Current_Stock']:.0f}  |  Total Sold: {row['Total_Qty_Sold']:.0f}",
             font=("Arial", 9), bg="#eaeded", fg="#566573").pack(pady=(0, 15))

    form = tk.Frame(win, bg="#eaeded")
    form.pack(pady=5)

    tk.Label(form, text="Category:", font=("Arial", 10, "bold"), bg="#eaeded").grid(row=0, column=0, sticky="w", pady=8, padx=5)
    category_combo = ttk.Combobox(form, font=("Arial", 10), width=30, values=CATEGORY_LIST)
    category_combo.set(row['Category'])
    category_combo.grid(row=0, column=1, pady=8, padx=5)

    tk.Label(form, text="Supplier:", font=("Arial", 10, "bold"), bg="#eaeded").grid(row=1, column=0, sticky="w", pady=8, padx=5)
    supplier_combo = ttk.Combobox(form, font=("Arial", 10), width=30, values=SUPPLIER_LIST)
    supplier_combo.set(row['Supplier'])
    supplier_combo.grid(row=1, column=1, pady=8, padx=5)

    tk.Label(form, text="Reorder Point:", font=("Arial", 10, "bold"), bg="#eaeded").grid(row=2, column=0, sticky="w", pady=8, padx=5)
    reorder_entry = tk.Entry(form, font=("Arial", 10), width=20)
    reorder_entry.insert(0, f"{row['Reorder_Point']:.0f}")
    reorder_entry.grid(row=2, column=1, sticky="w", pady=8, padx=5)

    tk.Label(form, text="Generic/Branded\n(medicines only, optional):", font=("Arial", 10, "bold"), bg="#eaeded", justify="left").grid(row=3, column=0, sticky="w", pady=8, padx=5)
    is_generic_var = tk.StringVar(value=row['Is_Generic'] if row['Is_Generic'] else '(N/A)')
    is_generic_frame = tk.Frame(form, bg="#eaeded")
    is_generic_frame.grid(row=3, column=1, sticky="w", pady=8, padx=5)
    for option in IS_GENERIC_OPTIONS:
        tk.Radiobutton(is_generic_frame, text=option, variable=is_generic_var, value=option,
                        bg="#eaeded", font=("Arial", 9)).pack(anchor="w")

    def stage_and_close():
        try:
            new_reorder = float(reorder_entry.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Input", "Reorder Point must be a number.")
            return
        new_supplier = supplier_combo.get().strip()
        new_category = category_combo.get().strip()
        new_is_generic = is_generic_var.get()
        if new_is_generic == '(N/A)':
            new_is_generic = ''

        idx = working_df.index[working_df['Item_Code'] == row['Item_Code']][0]
        working_df.at[idx, 'Reorder_Point'] = new_reorder
        working_df.at[idx, 'Supplier'] = new_supplier
        working_df.at[idx, 'Category'] = new_category
        working_df.at[idx, 'Is_Generic'] = new_is_generic
        pending_edits[row['Item_Code']] = {
            'Reorder_Point': new_reorder, 'Supplier': new_supplier,
            'Category': new_category, 'Is_Generic': new_is_generic,
        }

        populate_tree()
        win.destroy()
        select_next_unassigned(after_item_code=row['Item_Code'])

    btn_frame = tk.Frame(win, bg="#eaeded")
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="📋 STAGE THIS EDIT (not saved to file yet)", font=("Arial", 10, "bold"),
              bg="#17a2b8", fg="white", command=stage_and_close).grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="Cancel", font=("Arial", 10), bg="#6c757d", fg="white", command=win.destroy).grid(row=0, column=1, padx=5)


def select_next_unassigned(after_item_code=None):
    all_ids = list(tree.get_children())
    if not all_ids:
        return
    start = 0
    if after_item_code is not None and str(after_item_code) in all_ids:
        start = all_ids.index(str(after_item_code)) + 1

    for iid in all_ids[start:] + all_ids[:start]:
        row_match = working_df[working_df['Item_Code'].astype(str) == iid]
        if not row_match.empty:
            r = row_match.iloc[0]
            if not r['Supplier'] or not r['Category']:
                tree.selection_set(iid)
                tree.see(iid)
                return
    messagebox.showinfo("All Done!", "Every item in the current list already has a Category and Supplier.")


def on_search_change(*args):
    populate_tree()


def save_all_to_master():
    if not pending_edits:
        messagebox.showinfo("No Edits Yet", "There are no changes to save yet.")
        return
    if not messagebox.askyesno("Confirm", f"Save {len(pending_edits)} changed item(s) to {MASTER_FILE}?"):
        return

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER', dtype={'Item_Code': str})
    except PermissionError:
        messagebox.showerror("File In Use", f"Please close {MASTER_FILE} in Excel first before saving.")
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str)
    for col, default in [('Reorder_Point', 10), ('Supplier', ''), ('Category', ''), ('Is_Generic', '')]:
        if col not in df_master.columns:
            df_master[col] = default

    # Text columns that were all-blank/NaN get read back by pandas as
    # float64 (nothing but NaN to infer a type from), so assigning a string
    # into them raises "Invalid value ... for dtype 'float64'". Force them
    # to object dtype before writing any text into them.
    for col in ['Supplier', 'Category', 'Is_Generic']:
        df_master[col] = df_master[col].astype('object')

    for code, edit in pending_edits.items():
        mask = df_master['Item_Code'] == str(code)
        df_master.loc[mask, 'Reorder_Point'] = edit['Reorder_Point']
        df_master.loc[mask, 'Supplier'] = edit['Supplier']
        df_master.loc[mask, 'Category'] = edit['Category']
        df_master.loc[mask, 'Is_Generic'] = edit['Is_Generic']

    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy2(MASTER_FILE, os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_ITEM_SETTINGS_{ts}.xlsx'))

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        messagebox.showerror("File In Use", f"Please close {MASTER_FILE} in Excel. The changes were not saved.")
        return

    log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{log_time}] - ITEM SETTINGS UPDATE by {staff_name.upper()} ({len(pending_edits)} item/s)\n")
        for code, edit in pending_edits.items():
            logbook.write(f"   -> {code}: Category={edit['Category']}, Supplier={edit['Supplier']}, "
                          f"Reorder_Point={edit['Reorder_Point']:.0f}, Is_Generic={edit['Is_Generic'] or '(N/A)'}\n")

    messagebox.showinfo("Saved!", f"{len(pending_edits)} item(s) saved to {MASTER_FILE}.")
    pending_edits.clear()
    refresh_progress_label()


# ==========================================
#    MAIN GUI
# ==========================================
root.title("J-DAN PHARMACY - ITEM SETTINGS (Category / Supplier / Reorder Point)")
root.geometry("1150x720")
root.configure(bg="#f4f4f4")

header_frame = tk.Frame(root, bg="#154360", pady=15)
header_frame.pack(fill="x")
tk.Label(header_frame, text="ITEM SETTINGS", fg="white", bg="#154360", font=("Arial", 16, "bold")).pack()
tk.Label(header_frame, text="Category, Supplier & Reorder Point -- ranked by sales velocity (fast-movers first)", fg="#a9cce3", bg="#154360", font=("Arial", 10)).pack()

progress_label = tk.Label(root, text="", bg="#f4f4f4", font=("Arial", 10, "bold"), fg="#154360")
progress_label.pack(pady=8)

search_frame = tk.Frame(root, bg="#f4f4f4")
search_frame.pack(pady=(0, 5))
tk.Label(search_frame, text="🔍 Search (name or Item Code):", font=("Arial", 9), bg="#f4f4f4").pack(side="left", padx=5)
search_var = tk.StringVar()
search_var.trace_add('write', on_search_change)
search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 10), width=40)
search_entry.pack(side="left", padx=5)

btn_frame = tk.Frame(root, bg="#f4f4f4")
btn_frame.pack(pady=5)
tk.Button(btn_frame, text="✏️ EDIT SELECTED", font=("Helvetica", 10, "bold"), bg="#17a2b8", fg="white", width=20, command=lambda: open_editor()).grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="⏭ NEXT UNASSIGNED", font=("Helvetica", 10, "bold"), bg="#f0ad4e", fg="white", width=20, command=lambda: select_next_unassigned(tree.selection()[0] if tree.selection() else None)).grid(row=0, column=1, padx=5)
save_all_btn = tk.Button(btn_frame, text="💾 SAVE ALL CHANGES", font=("Helvetica", 11, "bold"), bg="#28a745", fg="white", width=22, state='disabled', command=save_all_to_master)
save_all_btn.grid(row=0, column=2, padx=5)

tree_frame = tk.Frame(root, bg="#f4f4f4")
tree_frame.pack(pady=10, padx=20, fill="both", expand=True)

tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scroll.pack(side="right", fill="y")

columns = ("Rank", "Code", "Product", "Sold", "Stock", "Reorder", "Category", "Supplier", "IsGeneric")
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=25, yscrollcommand=tree_scroll.set)
tree_scroll.config(command=tree.yview)

tree.heading("Rank", text="Rank")
tree.heading("Code", text="Item Code")
tree.heading("Product", text="Product Name")
tree.heading("Sold", text="Total Sold")
tree.heading("Stock", text="Current Stock")
tree.heading("Reorder", text="Reorder Point")
tree.heading("Category", text="Category")
tree.heading("Supplier", text="Supplier")
tree.heading("IsGeneric", text="Generic/Branded")

tree.column("Rank", width=50, anchor='center')
tree.column("Code", width=90, anchor='center')
tree.column("Product", width=260)
tree.column("Sold", width=70, anchor='center')
tree.column("Stock", width=80, anchor='center')
tree.column("Reorder", width=80, anchor='center')
tree.column("Category", width=150)
tree.column("Supplier", width=130)
tree.column("IsGeneric", width=100, anchor='center')

tree.tag_configure('unassigned', background="#fff3cd")
tree.tag_configure('inconsistent', background="#f8d7da")

tree.bind('<Double-1>', lambda e: open_editor())
tree.pack(side="left", fill="both", expand=True)


def on_close():
    if pending_edits:
        if messagebox.askyesno(
            "Unsaved Changes",
            f"You have {len(pending_edits)} edited item(s) that have NOT been saved to {MASTER_FILE} yet "
            f"(they are only staged in this window).\n\n"
            f"Click SAVE ALL CHANGES first if you want to keep them.\n\n"
            f"Close anyway and lose these changes?"
        ):
            root.destroy()
    else:
        root.destroy()


root.protocol("WM_DELETE_WINDOW", on_close)

footer_frame = tk.Frame(root, bg="#d5d8dc", pady=5)
footer_frame.pack(fill="x", side="bottom")
tk.Label(footer_frame, text=f"Authorized System User: {staff_name.upper()}", font=("Arial", 8, "bold"), bg="#d5d8dc", fg="#566573").pack()

startup_conflicts = populate_tree()
if startup_conflicts:
    conflict_lines = "\n".join(f"  {code}: {name}" for code, name in startup_conflicts[:15])
    more = f"\n  ... and {len(startup_conflicts) - 15} more" if len(startup_conflicts) > 15 else ""
    messagebox.showwarning(
        "Inconsistent Spelling Under the Same Item_Code",
        f"{len(startup_conflicts)} item(s) have inconsistent Product_Name spelling across their lot rows in Master "
        f"(merged here using the most common spelling for now):\n\n{conflict_lines}{more}\n\n"
        f"Recommendation: fix the typo in Master (standardize the spelling) so this doesn't recur."
    )

root.mainloop()
