import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import datetime
import os
import glob
import shutil
import sys
import traceback

# ==========================================
#    CRASH DETECTOR (Anti-Silent Crash)
# ==========================================
def handle_exception(exc_type, exc_value, exc_traceback):
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    messagebox.showerror("System Critical Error", f"Nakatagpo ng error ang system. Paki-screenshot ito at ipadala sa Admin:\n\n{err_msg}")

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

CATEGORY_LIST = [
    'MAINTENANCE MEDS', 'PAIN & FEVER RELIEF', 'COUGH, COLD & ALLERGY',
    'GI / DIGESTIVE CARE', 'VITAMINS & SUPPLEMENTS', 'WOUND CARE & TOPICALS',
    'BABY CARE', 'FEMININE CARE', 'PERSONAL CARE & COSMETICS',
    'MEDICAL SUPPLIES & DEVICES', 'HOUSEHOLD & GROCERY', 'OTHERS / UNCATEGORIZED',
]


def norm(s):
    return str(s).strip().upper()


def load_sales_velocity():
    frames = []
    for path in glob.glob(f'{EOD_DIR}/PROCESSED_EOD_*.xlsx'):
        try:
            df = pd.read_excel(path)
        except Exception:
            continue
        frames.append(df)
    if not frames:
        return pd.DataFrame(columns=['Item_Code', 'Total_Qty_Sold'])
    all_sales = pd.concat(frames, ignore_index=True)
    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    all_sales['Item_Code'] = all_sales['Item Code'].astype(str).str.strip()
    return all_sales.groupby('Item_Code').agg(Total_Qty_Sold=('Qty', 'sum')).reset_index()


def load_working_data():
    df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
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
    ).reset_index()
    agg['Product_Name'] = agg['Item_Code'].map(representative_name)
    agg['Name_Inconsistent'] = agg['Item_Code'].isin(inconsistent_codes)

    velocity = load_sales_velocity()
    agg = agg.merge(velocity, on='Item_Code', how='left')
    agg['Total_Qty_Sold'] = agg['Total_Qty_Sold'].fillna(0)
    agg = agg.sort_values('Total_Qty_Sold', ascending=False).reset_index(drop=True)
    agg['Rank'] = agg.index + 1
    return agg


working_df = load_working_data()
pending_edits = {}  # Item_Code -> {'Reorder_Point': x, 'Supplier': y, 'Category': z}


def refresh_progress_label():
    total = len(working_df)
    supplier_filled = (working_df['Supplier'].astype(str).str.strip() != '').sum()
    category_filled = (working_df['Category'].astype(str).str.strip() != '').sum()
    fully_set = ((working_df['Supplier'].astype(str).str.strip() != '') & (working_df['Category'].astype(str).str.strip() != '')).sum()
    pending_count = len(pending_edits)
    progress_label.config(
        text=f"Supplier: {supplier_filled}/{total}   |   Category: {category_filled}/{total}   |   "
             f"Kumpleto (parehong Category+Supplier): {fully_set}/{total}"
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
        supplier_display = row['Supplier'] if row['Supplier'] else '(blangko)'
        category_display = row['Category'] if row['Category'] else '(blangko)'
        product_display = row['Product_Name']
        if row['Name_Inconsistent']:
            product_display = f"⚠️ {product_display} (magkaibang spelling sa Master)"
            name_conflicts.append((row['Item_Code'], row['Product_Name']))
        is_unassigned = not row['Supplier'] or not row['Category']
        tag = 'inconsistent' if row['Name_Inconsistent'] else ('unassigned' if is_unassigned else '')
        tree.insert('', tk.END, iid=str(row['Item_Code']), values=(
            int(row['Rank']), row['Item_Code'], product_display,
            f"{row['Total_Qty_Sold']:.0f}", f"{row['Current_Stock']:.0f}",
            f"{row['Reorder_Point']:.0f}", category_display, supplier_display,
        ), tags=(tag,) if tag else ())
    refresh_progress_label()
    return name_conflicts


def open_editor(item_code=None):
    if item_code is None:
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Walang Napili", "Pumili muna ng item sa listahan.")
            return
        item_code = selected[0]

    row_match = working_df[working_df['Item_Code'].astype(str) == str(item_code)]
    if row_match.empty:
        return
    row = row_match.iloc[0]

    win = tk.Toplevel(root)
    win.title(f"I-edit: {row['Product_Name']}")
    win.geometry("500x320")
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

    def save_and_close():
        try:
            new_reorder = float(reorder_entry.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Input", "Ang Reorder Point ay dapat numero.")
            return
        new_supplier = supplier_combo.get().strip()
        new_category = category_combo.get().strip()

        idx = working_df.index[working_df['Item_Code'] == row['Item_Code']][0]
        working_df.at[idx, 'Reorder_Point'] = new_reorder
        working_df.at[idx, 'Supplier'] = new_supplier
        working_df.at[idx, 'Category'] = new_category
        pending_edits[row['Item_Code']] = {'Reorder_Point': new_reorder, 'Supplier': new_supplier, 'Category': new_category}

        populate_tree()
        win.destroy()
        select_next_unassigned(after_item_code=row['Item_Code'])

    btn_frame = tk.Frame(win, bg="#eaeded")
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="💾 SAVE (sa listahan lang, hindi pa sa file)", font=("Arial", 10, "bold"),
              bg="#28a745", fg="white", command=save_and_close).grid(row=0, column=0, padx=5)
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
    messagebox.showinfo("Tapos na!", "Lahat ng items sa kasalukuyang listahan ay may Category at Supplier na.")


def on_search_change(*args):
    populate_tree()


def save_all_to_master():
    if not pending_edits:
        messagebox.showinfo("Wala pang Edit", "Wala pang binago na i-se-save.")
        return
    if not messagebox.askyesno("Kumpirmahin", f"I-save ang {len(pending_edits)} item(s) na binago papunta sa {MASTER_FILE}?"):
        return

    try:
        df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        messagebox.showerror("File In Use", f"Isara muna ang {MASTER_FILE} sa Excel bago mag-save.")
        return

    df_master['Item_Code'] = df_master['Item_Code'].astype(str)
    for col, default in [('Reorder_Point', 10), ('Supplier', ''), ('Category', '')]:
        if col not in df_master.columns:
            df_master[col] = default

    for code, edit in pending_edits.items():
        mask = df_master['Item_Code'] == str(code)
        df_master.loc[mask, 'Reorder_Point'] = edit['Reorder_Point']
        df_master.loc[mask, 'Supplier'] = edit['Supplier']
        df_master.loc[mask, 'Category'] = edit['Category']

    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    shutil.copy2(MASTER_FILE, os.path.join(BACKUP_DIR, f'Inventory_Master_BEFORE_ITEM_SETTINGS_{ts}.xlsx'))

    try:
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
            df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
    except PermissionError:
        messagebox.showerror("File In Use", f"Isara muna ang {MASTER_FILE} sa Excel. Hindi na-save ang mga pagbabago.")
        return

    log_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{log_time}] - ITEM SETTINGS UPDATE by {staff_name.upper()} ({len(pending_edits)} item/s)\n")
        for code, edit in pending_edits.items():
            logbook.write(f"   -> {code}: Category={edit['Category']}, Supplier={edit['Supplier']}, Reorder_Point={edit['Reorder_Point']:.0f}\n")

    messagebox.showinfo("Na-save!", f"{len(pending_edits)} item(s) na-save sa {MASTER_FILE}.")
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
tk.Label(search_frame, text="🔍 Hanapin (pangalan o Item Code):", font=("Arial", 9), bg="#f4f4f4").pack(side="left", padx=5)
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

columns = ("Rank", "Code", "Product", "Sold", "Stock", "Reorder", "Category", "Supplier")
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

tree.column("Rank", width=50, anchor='center')
tree.column("Code", width=90, anchor='center')
tree.column("Product", width=280)
tree.column("Sold", width=75, anchor='center')
tree.column("Stock", width=85, anchor='center')
tree.column("Reorder", width=85, anchor='center')
tree.column("Category", width=160)
tree.column("Supplier", width=140)

tree.tag_configure('unassigned', background="#fff3cd")
tree.tag_configure('inconsistent', background="#f8d7da")

tree.bind('<Double-1>', lambda e: open_editor())
tree.pack(side="left", fill="both", expand=True)

footer_frame = tk.Frame(root, bg="#d5d8dc", pady=5)
footer_frame.pack(fill="x", side="bottom")
tk.Label(footer_frame, text=f"Authorized System User: {staff_name.upper()}", font=("Arial", 8, "bold"), bg="#d5d8dc", fg="#566573").pack()

startup_conflicts = populate_tree()
if startup_conflicts:
    conflict_lines = "\n".join(f"  {code}: {name}" for code, name in startup_conflicts[:15])
    more = f"\n  ... at {len(startup_conflicts) - 15} pa" if len(startup_conflicts) > 15 else ""
    messagebox.showwarning(
        "Magkaibang Spelling sa Ilalim ng Parehong Item_Code",
        f"{len(startup_conflicts)} item(s) ang may magkaibang Product_Name spelling sa iba't ibang lot row nito sa Master "
        f"(pinag-isa muna sila dito gamit ang pinaka-madalas na spelling):\n\n{conflict_lines}{more}\n\n"
        f"Rekomendasyon: ayusin ang typo sa Master (i-standardize ang spelling) para hindi ito maulit."
    )

root.mainloop()
