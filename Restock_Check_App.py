import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import datetime
import os
import glob
import sys
import traceback

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
#    SECURITY LOGIN & LOGBOOK SYSTEM
# ==========================================
root = tk.Tk()
root.withdraw()

staff_name = simpledialog.askstring("J-DAN PHARMACY - SECURITY", "Enter Authorized User Name:", parent=root)

if not staff_name or staff_name.strip() == "":
    messagebox.showerror("Access Denied", "Authorized user name is required to access the Restock Check module.")
    sys.exit()

current_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
try:
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{current_time}] - USER: {staff_name.upper()} - Accessed RESTOCK CHECK module\n")
except Exception:
    pass

root.deiconify()

# ==========================================
#    SYSTEM CONFIGURATION
# ==========================================
MASTER_FILE = 'Inventory_Master.xlsx'
REPORT_DIR = 'JDan_Executive_Reports'
EOD_DIR = 'JDan_Sales_System/EOD_Deduction_Reports'
UNASSIGNED_LABEL = '⚠️ UNASSIGNED / NEEDS SUPPLIER'

# Sales-velocity thresholds for Priority Mode -- adjust here if Dan wants a
# higher bar for "fast-moving" (e.g. require 5+ units instead of just 1+).
PRIORITY_WINDOW_DAYS = 60
MIN_UNITS_SOLD_FOR_PRIORITY = 1

last_scan_results = None  # populated by run_scan(), read by export_list()
last_inconsistent_display = []
last_inconsistent_item_codes = []  # raw Item_Codes, read by export_lot_inconsistency_report()
last_raw_lots_df = None  # per-lot rows (pre-aggregation), read by export_lot_inconsistency_report()


def load_sales_velocity(window_days):
    """Total units sold per Item_Code in the last `window_days` days, from
    JDan_Sales_System/EOD_Deduction_Reports/PROCESSED_EOD_*.xlsx (read-only).
    Item Code is already resolved in these files, so no name-matching needed."""
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
    all_sales['Date'] = pd.to_datetime(all_sales['Date'], errors='coerce')
    cutoff = pd.Timestamp.now().normalize() - pd.Timedelta(days=window_days)
    all_sales = all_sales[all_sales['Date'] >= cutoff]

    all_sales['Qty'] = pd.to_numeric(all_sales['Qty'], errors='coerce').fillna(0).abs()
    all_sales['Item_Code'] = all_sales['Item Code'].astype(str).str.strip()
    return all_sales.groupby('Item_Code').agg(Total_Qty_Sold=('Qty', 'sum')).reset_index()


def run_scan():
    global last_scan_results, last_inconsistent_display, last_inconsistent_item_codes, last_raw_lots_df

    try:
        df = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
    except PermissionError:
        messagebox.showerror("File In Use", f"Please close {MASTER_FILE} in Excel first before scanning.")
        return
    except Exception as e:
        messagebox.showerror("Error", f"Could not open {MASTER_FILE}.\n\nDetails: {e}")
        return

    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.strip().str.upper() != 'DISCONTINUED']

    df['Item_Code'] = df['Item_Code'].astype(str).str.strip()
    df['Current_Stock'] = pd.to_numeric(df['Current_Stock'], errors='coerce').fillna(0)
    if 'Reorder_Point' not in df.columns:
        df['Reorder_Point'] = 10
    df['Reorder_Point'] = pd.to_numeric(df['Reorder_Point'], errors='coerce').fillna(10)
    if 'Supplier' not in df.columns:
        df['Supplier'] = ''
    df['Supplier'] = df['Supplier'].fillna('').astype(str).str.strip()
    if 'Lot_Number' not in df.columns:
        df['Lot_Number'] = ''
    if 'Expiry_Date' not in df.columns:
        df['Expiry_Date'] = ''

    # Item_Code is the only stable identifier -- Product_Name can drift
    # (typos/spelling variants) across a code's lot rows, so grouping by name
    # would silently split one product into several. Group by Item_Code, and
    # flag codes where Reorder_Point/Supplier/name actually disagree across lots.
    df['_name_clean'] = df['Product_Name'].astype(str).str.strip()
    inconsistency = df.groupby('Item_Code').agg(
        RP_Nunique=('Reorder_Point', 'nunique'),
        Supplier_Nunique=('Supplier', lambda s: s[s != ''].nunique()),
        Name_Nunique=('_name_clean', 'nunique'),
    ).reset_index()
    inconsistent_items = inconsistency[
        (inconsistency['RP_Nunique'] > 1) | (inconsistency['Supplier_Nunique'] > 1) | (inconsistency['Name_Nunique'] > 1)
    ]['Item_Code'].tolist()

    name_freq = df.groupby(['Item_Code', '_name_clean']).size().reset_index(name='freq')
    name_freq = name_freq.sort_values(['Item_Code', 'freq'], ascending=[True, False])
    representative_name = name_freq.drop_duplicates('Item_Code', keep='first').set_index('Item_Code')['_name_clean']

    agg = df.groupby('Item_Code').agg(
        Current_Stock=('Current_Stock', 'sum'),
        Reorder_Point=('Reorder_Point', 'first'),
        Supplier=('Supplier', 'first'),
    ).reset_index()
    agg['Product_Name'] = agg['Item_Code'].map(representative_name)

    velocity = load_sales_velocity(PRIORITY_WINDOW_DAYS)
    agg = agg.merge(velocity, on='Item_Code', how='left')
    agg['Total_Qty_Sold'] = agg['Total_Qty_Sold'].fillna(0)
    agg['Is_Active'] = agg['Total_Qty_Sold'] >= MIN_UNITS_SOLD_FOR_PRIORITY

    low_stock = agg[agg['Current_Stock'] <= agg['Reorder_Point']].copy()
    low_stock['Suggested_Order_Qty'] = (low_stock['Reorder_Point'] - low_stock['Current_Stock']).clip(lower=0)
    low_stock['Supplier_Group'] = low_stock['Supplier'].apply(lambda s: s if s else UNASSIGNED_LABEL)
    low_stock = low_stock.sort_values(['Supplier_Group', 'Current_Stock'])

    inconsistent_display = [f"{code}: {representative_name.get(code, '')}" for code in inconsistent_items]

    last_scan_results = low_stock
    last_inconsistent_display = inconsistent_display
    last_inconsistent_item_codes = inconsistent_items
    last_raw_lots_df = df[['Item_Code', 'Product_Name', 'Reorder_Point', 'Supplier', 'Current_Stock', 'Lot_Number', 'Expiry_Date']].copy()
    export_inconsistency_btn.config(state='normal' if inconsistent_items else 'disabled')
    render_results()

    if inconsistent_display:
        messagebox.showwarning(
            "Data Inconsistency Detected",
            "Some items have mismatched Reorder_Point, Supplier, or Product_Name (typo) across their lot rows in Master "
            "(should be one value per product):\n\n" + "\n".join(inconsistent_display[:15]) +
            ("\n... and more" if len(inconsistent_display) > 15 else "")
        )


def get_display_results():
    """Applies the current Priority Mode checkbox state to the last scan."""
    if last_scan_results is None:
        return None
    if priority_mode_var.get():
        return last_scan_results[last_scan_results['Is_Active']]
    return last_scan_results


def render_results():
    display_df = get_display_results()
    if display_df is None:
        return
    populate_tree(display_df, last_inconsistent_display)


def populate_tree(low_stock, inconsistent_display):
    tree.delete(*tree.get_children())

    if low_stock.empty:
        message = 'No fast-moving items are low on stock right now.' if priority_mode_var.get() else 'No items are low on stock right now.'
        tree.insert('', tk.END, values=(message, '', '', '', ''), tags=('info',))
    else:
        for supplier, group in low_stock.groupby('Supplier_Group', sort=False):
            tree.insert('', tk.END, values=(f'--- {supplier} ({len(group)} item(s)) ---', '', '', '', ''), tags=('group_header',))
            for _, row in group.iterrows():
                product_display = row['Product_Name'] + (' [ACTIVE]' if row['Is_Active'] else '')
                tree.insert('', tk.END, values=(
                    row['Item_Code'], product_display, f"{row['Current_Stock']:.0f}",
                    f"{row['Reorder_Point']:.0f}", f"{row['Suggested_Order_Qty']:.0f}",
                ))

    summary = f"{len(low_stock)} item(s) shown"
    if priority_mode_var.get():
        summary += f" (Priority Mode ON: {PRIORITY_WINDOW_DAYS}-day fast-movers only)"
    if not low_stock.empty:
        unassigned_count = (low_stock['Supplier_Group'] == UNASSIGNED_LABEL).sum()
        summary += f" | {unassigned_count} with no supplier assigned"
    if inconsistent_display:
        summary += f" | ⚠️ {len(inconsistent_display)} item(s) with mismatched Reorder_Point/Supplier/name across lots"
    summary += f" | Last scan: {datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')}"
    summary_label.config(text=summary)

    export_btn.config(state='normal' if not low_stock.empty else 'disabled')
    export_txt_btn.config(state='normal' if not low_stock.empty else 'disabled')


def export_list():
    export_source = get_display_results()
    if export_source is None or export_source.empty:
        messagebox.showinfo("Nothing to Export", "There is no scan result to export yet.")
        return
    if not messagebox.askyesno("Confirm", "Export the restock list as an Excel file?"):
        return

    os.makedirs(REPORT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_DIR, f'RESTOCK_LIST_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    export_df = export_source[['Supplier_Group', 'Item_Code', 'Product_Name', 'Current_Stock', 'Reorder_Point', 'Suggested_Order_Qty', 'Is_Active']].copy()
    export_df.columns = ['Supplier', 'Item_Code', 'Product_Name', 'Current_Stock', 'Reorder_Point', 'Suggested_Order_Qty', 'Fast_Moving']
    export_df['Fast_Moving'] = export_df['Fast_Moving'].map({True: 'ACTIVE', False: ''})
    export_df.to_excel(out_path, index=False)
    messagebox.showinfo("Exported", f"Restock list saved to:\n{out_path}")


def export_list_txt():
    export_source = get_display_results()
    if export_source is None or export_source.empty:
        messagebox.showinfo("Nothing to Export", "There is no scan result to export yet.")
        return
    if not messagebox.askyesno("Confirm", "Export the restock list as a TXT file (for phone editing)?"):
        return

    os.makedirs(REPORT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_DIR, f'RESTOCK_LIST_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.txt')

    lines = []
    for supplier, group in export_source.groupby('Supplier_Group', sort=False):
        lines.append(f"=== {supplier} ===")
        for _, row in group.iterrows():
            tag = ' [ACTIVE]' if row['Is_Active'] else ''
            lines.append(f"{row['Product_Name']}{tag} - {row['Current_Stock']:.0f} pcs left")
        lines.append("")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines).strip() + "\n")

    messagebox.showinfo("Exported", f"Restock list (TXT) saved to:\n{out_path}")


def export_lot_inconsistency_report():
    if not last_inconsistent_item_codes or last_raw_lots_df is None:
        messagebox.showinfo("Nothing to Export", "No lot inconsistencies were found in the last scan.")
        return
    if not messagebox.askyesno("Confirm", "Export the full lot inconsistency report as a CSV file?"):
        return

    report_df = last_raw_lots_df[last_raw_lots_df['Item_Code'].isin(last_inconsistent_item_codes)].copy()
    report_df['Lot_Row_Number'] = report_df.groupby('Item_Code').cumcount() + 1
    report_df['Expiry_Date'] = pd.to_datetime(report_df['Expiry_Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')

    report_df = report_df[[
        'Item_Code', 'Lot_Row_Number', 'Product_Name', 'Reorder_Point', 'Supplier',
        'Current_Stock', 'Lot_Number', 'Expiry_Date',
    ]].sort_values(['Item_Code', 'Lot_Row_Number'])

    os.makedirs(REPORT_DIR, exist_ok=True)
    out_path = os.path.join(REPORT_DIR, f'LOT_INCONSISTENCY_REPORT_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    report_df.to_csv(out_path, index=False)
    messagebox.showinfo(
        "Exported",
        f"Lot inconsistency report saved to:\n{out_path}\n\n"
        f"{report_df['Item_Code'].nunique()} product(s), {len(report_df)} lot row(s) total."
    )


# ==========================================
#    MAIN GUI
# ==========================================
root.title("J-DAN PHARMACY - RESTOCK / ORDERING CHECK")
root.geometry("1000x700")
root.configure(bg="#f4f4f4")

header_frame = tk.Frame(root, bg="#154360", pady=15)
header_frame.pack(fill="x")
tk.Label(header_frame, text="RESTOCK CHECK", fg="white", bg="#154360", font=("Arial", 16, "bold")).pack()
tk.Label(header_frame, text="Current_Stock vs Reorder_Point, grouped by Supplier", fg="#a9cce3", bg="#154360", font=("Arial", 10)).pack()

btn_frame = tk.Frame(root, bg="#f4f4f4")
btn_frame.pack(pady=10)
tk.Button(btn_frame, text="🔍 SCAN NOW", font=("Helvetica", 11, "bold"), bg="#28a745", fg="white", width=20, height=2, command=run_scan).grid(row=0, column=0, padx=10)
export_btn = tk.Button(btn_frame, text="📤 EXPORT LIST (Optional)", font=("Helvetica", 11, "bold"), bg="#f0ad4e", fg="white", width=25, height=2, state='disabled', command=export_list)
export_btn.grid(row=0, column=1, padx=10)
export_txt_btn = tk.Button(btn_frame, text="📱 EXPORT TXT (Phone-Friendly)", font=("Helvetica", 11, "bold"), bg="#5dade2", fg="white", width=27, height=2, state='disabled', command=export_list_txt)
export_txt_btn.grid(row=0, column=2, padx=10)
export_inconsistency_btn = tk.Button(btn_frame, text="⚠️ EXPORT LOT INCONSISTENCY REPORT", font=("Helvetica", 11, "bold"), bg="#c0392b", fg="white", width=32, height=2, state='disabled', command=export_lot_inconsistency_report)
export_inconsistency_btn.grid(row=0, column=3, padx=10)

priority_mode_var = tk.BooleanVar(value=False)
priority_check = tk.Checkbutton(
    root, text=f"Show Priority Items Only (Fast-Moving, sold within last {PRIORITY_WINDOW_DAYS} days)",
    variable=priority_mode_var, bg="#f4f4f4", font=("Arial", 10, "bold"), fg="#154360",
    command=render_results,
)
priority_check.pack(pady=(0, 5))

summary_label = tk.Label(root, text="Click SCAN NOW to see items that are low on stock.", bg="#f4f4f4", font=("Arial", 9, "italic"), fg="#566573")
summary_label.pack(pady=(0, 5))

tree_frame = tk.Frame(root, bg="#f4f4f4")
tree_frame.pack(pady=5, padx=20, fill="both", expand=True)

tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scroll.pack(side="right", fill="y")

columns = ("Code", "Product", "Current", "Reorder", "Suggested")
tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=22, yscrollcommand=tree_scroll.set)
tree_scroll.config(command=tree.yview)

tree.heading("Code", text="Item Code")
tree.heading("Product", text="Product Name / Supplier Group")
tree.heading("Current", text="Current Stock")
tree.heading("Reorder", text="Reorder Point")
tree.heading("Suggested", text="Suggested Order Qty")

tree.column("Code", width=100, anchor='center')
tree.column("Product", width=380)
tree.column("Current", width=110, anchor='center')
tree.column("Reorder", width=110, anchor='center')
tree.column("Suggested", width=140, anchor='center')

tree.tag_configure('group_header', background="#d5d8dc", foreground="#154360")
tree.tag_configure('info', foreground="#566573")

tree.pack(side="left", fill="both", expand=True)

footer_frame = tk.Frame(root, bg="#d5d8dc", pady=5)
footer_frame.pack(fill="x", side="bottom")
tk.Label(footer_frame, text=f"Authorized System User: {staff_name.upper()}", font=("Arial", 8, "bold"), bg="#d5d8dc", fg="#566573").pack()

root.mainloop()
