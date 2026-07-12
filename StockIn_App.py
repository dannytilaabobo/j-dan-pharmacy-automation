import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import datetime
import os
import sys

from name_cleaning import clean_text

# ==========================================
#    GUI INITIALIZATION & AUDIT TRAIL
# ==========================================
root = tk.Tk()
root.withdraw() # Hide main window during authentication

# Standardized authentication prompt
staff_name = simpledialog.askstring("J-DAN PHARMACY - SECURITY", "Enter Authorized Encoder Name:")

if not staff_name or staff_name.strip() == "":
    messagebox.showerror("Access Denied", "Encoder name is required to proceed with Stock-In operations.")
    sys.exit()

current_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
try:
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{current_time}] - USER: {staff_name.upper()} - Accessed INCOMING SUPPLIES module.\n")
except Exception:
    pass # Failsafe to prevent crash if log text file is locked

root.deiconify() # Reveal main window
root.title("J-DAN PHARMACY - CONTROLLED INVENTORY STOCK-IN")
root.geometry("880x650")
root.configure(bg="#f4f4f4")

# ==========================================
#    SYSTEM VARIABLES & FOLDERS
# ==========================================
master_file = 'Inventory_Master.xlsx'
base_report_folder = 'JDan_StockIn_Reports'

if not os.path.exists(base_report_folder):
    os.makedirs(base_report_folder)

# --- SYSTEM LOGIC ---
def load_items():
    try:
        df = pd.read_excel(master_file, sheet_name='INVENTORY MASTER', dtype=str)
        if 'Product_Name' not in df.columns:
            return {}
        items = df[['Item_Code', 'Product_Name']].drop_duplicates().dropna()
        items = items.sort_values(by='Product_Name')
        return dict(zip(items['Product_Name'], items['Item_Code']))
    except Exception as e:
        messagebox.showerror("FILE IN USE", f"Please close the Inventory Masterlist Excel file before operating the system.\n\nError: {e}")
        return {}

def filter_dropdown():
    typed = combo_product.get().strip()
    if typed == '':
        combo_product['values'] = product_names
    else:
        filtered = [item for item in product_names if typed.lower() in item.lower()]
        combo_product['values'] = filtered

def on_keyrelease(event):
    if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'BackSpace'):
        return
    filter_dropdown()

def on_select(event):
    selected = combo_product.get()
    entry_item_code.config(state='normal')
    entry_item_code.delete(0, tk.END)
    if selected in item_dict:
        entry_item_code.insert(0, item_dict[selected])
    entry_item_code.config(state='readonly')
    combo_product['values'] = product_names
    entry_lot.focus()

def add_to_staging_list():
    prod = clean_text(combo_product.get().strip())
    code = entry_item_code.get().strip()
    lot = clean_text(entry_lot.get().strip())
    exp = entry_exp.get().strip().upper()
    qty = entry_qty.get().strip()

    if prod not in item_dict:
        messagebox.showerror("Invalid Product", "Product name not found. Please select a valid item from the dropdown list.")
        return
    if not lot or not exp or not qty:
        messagebox.showerror("Incomplete Data", "All fields are required. Please fill out the form completely.")
        return
    try:
        float(qty)
    except:
        messagebox.showerror("Invalid Quantity", "Quantity must be a numerical value.")
        return

    # Expiry Validation (Blueprint Compliance: MM/YYYY or NO-EXP)
    if exp != 'NO-EXP':
        try:
            exp_dt = pd.to_datetime(exp, format='%m/%Y', errors='coerce')
            if pd.isna(exp_dt):
                raise ValueError
            
            today = pd.Timestamp.now().normalize()
            days_left = (exp_dt - today).days
            
            if days_left < 365:
                override = messagebox.askyesno(
                    "⚠️ RETURN TO SUPPLIER WARNING ⚠️", 
                    f"POTENTIAL FINANCIAL LOSS DETECTED.\n\n"
                    f"This product will expire in {days_left} days (Less than 1 Year).\n"
                    f"Standard operating procedure dictates returning this batch to the supplier.\n\n"
                    f"Do you wish to OVERRIDE and accept this batch into inventory?"
                )
                if not override:
                    return 
        except:
            messagebox.showerror("Invalid Date Format", "Expiry Date must follow the MM/YYYY format (e.g., 12/2027) or be declared as 'NO-EXP'.")
            return

    # ==========================================
    # ✨ DUAL-LAYER DOUBLE POSTING CHECKER ✨
    # ==========================================
    
    # CHECK 1: Staging List Validation (Screen)
    for item in tree_stage.get_children():
        vals = tree_stage.item(item, 'values')
        if str(vals[0]).strip() == code and str(vals[2]).strip() == lot:
            ans_screen = messagebox.askyesno(
                "⚠️ DUPLICATE ENTRY DETECTED", 
                f"The product '{prod}' with Lot Number '{lot}' is already in your staging list below.\n\n"
                f"Are you sure you want to add this duplicate entry to the current batch?"
            )
            if not ans_screen:
                return # Cancel addition

    # CHECK 2: Masterlist Validation (Excel)
    try:
        df_check = pd.read_excel(master_file, sheet_name='INVENTORY MASTER', dtype=str)
        df_check['Item_Code'] = df_check['Item_Code'].astype(str).str.strip()
        df_check['Lot_Number'] = df_check['Lot_Number'].astype(str).str.strip()
        
        mask = (df_check['Item_Code'] == code) & (df_check['Lot_Number'] == lot)
        if mask.any():
            ans_excel = messagebox.askyesno(
                "⚠️ EXISTING MASTERLIST ENTRY", 
                f"WARNING: The product '{prod}' with Lot Number '{lot}' ALREADY EXISTS in the Master Inventory.\n\n"
                f"Are you receiving an additional delivery of an existing lot?\n"
                f"Click YES to restock the existing lot.\n"
                f"Click NO to cancel if this is an accidental double-encode."
            )
            if not ans_excel:
                return # Cancel addition
    except Exception as e:
        pass # Bypass gracefully if master file is inaccessible temporarily

    # ==========================================

    # Append to treeview
    tree_stage.insert("", tk.END, values=(code, prod, lot, exp, qty))
    
    # Reset inputs
    combo_product.set('')
    entry_item_code.config(state='normal')
    entry_item_code.delete(0, tk.END)
    entry_item_code.config(state='readonly')
    entry_lot.delete(0, tk.END)
    entry_exp.delete(0, tk.END)
    entry_qty.delete(0, tk.END)
    combo_product.focus()

def delete_selected_staged():
    selected = tree_stage.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select an item from the staging list to remove.")
        return
    for item in selected:
        tree_stage.delete(item)

def process_batch_to_master():
    all_staged = tree_stage.get_children()
    if not all_staged:
        messagebox.showwarning("Empty Batch", "The staging list is empty. Please add items before processing.")
        return

    if not messagebox.askyesno("CONFIRM BATCH PROCESSING", f"Are you sure you want to process and finalize these {len(all_staged)} items into the Live Inventory Masterlist?"):
        return

    try:
        df = pd.read_excel(master_file, sheet_name='INVENTORY MASTER', dtype=str)
        
        # Ensure requisite columns exist
        for col in ['Lot_Number', 'Date_Encoded', 'Current_Stock', 'Encoded_By']:
            if col not in df.columns:
                df[col] = ""

        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        report_data = []

        for item in all_staged:
            vals = tree_stage.item(item, 'values')
            code, prod, lot, exp, qty = vals[0], vals[1], vals[2], vals[3], vals[4]
            
            # Append audit trail data for reporting
            report_data.append([code, prod, lot, exp, qty, today_str, staff_name.upper()])

            duplicate_mask = (df['Item_Code'] == code) & (df['Lot_Number'] == lot)
            if duplicate_mask.any():
                idx = df[duplicate_mask].index[0]
                old_qty = float(df.at[idx, 'Current_Stock']) if pd.notna(df.at[idx, 'Current_Stock']) else 0
                df.at[idx, 'Current_Stock'] = str(old_qty + float(qty))
                df.at[idx, 'Date_Encoded'] = today_str
                df.at[idx, 'Encoded_By'] = staff_name.upper()
            else:
                new_row = {
                    'Item_Code': code,
                    'Product_Name': prod,
                    'Lot_Number': lot,
                    'Expiry_Date': exp,
                    'Current_Stock': qty,
                    'Date_Encoded': today_str,
                    'Encoded_By': staff_name.upper()
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        with pd.ExcelWriter(master_file, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)

        date_folder = os.path.join(base_report_folder, today_str)
        if not os.path.exists(date_folder):
            os.makedirs(date_folder)
            
        timestamp = datetime.datetime.now().strftime("%H%M%S")
        report_file = os.path.join(date_folder, f'STOCKIN_BATCH_{timestamp}.xlsx')
        
        df_rep = pd.DataFrame(report_data, columns=['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Qty_Delivered', 'Date_Encoded', 'Encoded_By'])
        df_rep.to_excel(report_file, index=False)

        for item in all_staged:
            tree_stage.delete(item)

        messagebox.showinfo("OPERATION SUCCESSFUL", f"Batch successfully processed into the Inventory Masterlist.\n\nAudit Report Generated:\n{report_file}")

    except Exception as e:
        messagebox.showerror("CRITICAL ERROR", f"Failed to save data. Ensure the 'Inventory_Master.xlsx' file is completely closed.\n\nSystem Error Trace: {e}")

# --- GUI DESIGN ---
item_dict = load_items()
product_names = list(item_dict.keys()) if item_dict else []

tk.Label(root, text="J-DAN PHARMACY BATCH STOCK-IN", font=("Helvetica", 16, "bold"), bg="#f4f4f4", fg="#003366").pack(pady=10)

frame_input = tk.Frame(root, bg="#f4f4f4")
frame_input.pack(pady=5)

tk.Label(frame_input, text="Search Product:", font=("Helvetica", 10, "bold"), bg="#f4f4f4").grid(row=0, column=0, sticky="w", pady=5)

combo_product = ttk.Combobox(frame_input, width=35, font=("Helvetica", 10), postcommand=filter_dropdown)
combo_product.grid(row=0, column=1, pady=5, padx=10)
combo_product['values'] = product_names
combo_product.bind("<KeyRelease>", on_keyrelease)
combo_product.bind("<<ComboboxSelected>>", on_select)

tk.Label(frame_input, text="Item Code (Auto):", font=("Helvetica", 10), bg="#f4f4f4").grid(row=0, column=2, sticky="w", pady=5)
entry_item_code = tk.Entry(frame_input, width=20, font=("Helvetica", 10), state='readonly')
entry_item_code.grid(row=0, column=3, pady=5, padx=10)

tk.Label(frame_input, text="Lot Number:", font=("Helvetica", 10), bg="#f4f4f4").grid(row=1, column=0, sticky="w", pady=5)
entry_lot = tk.Entry(frame_input, width=37, font=("Helvetica", 10))
entry_lot.grid(row=1, column=1, pady=5, padx=10)

tk.Label(frame_input, text="Expiry (MM/YYYY or NO-EXP):", font=("Helvetica", 10), bg="#f4f4f4").grid(row=1, column=2, sticky="w", pady=5)
entry_exp = tk.Entry(frame_input, width=20, font=("Helvetica", 10))
entry_exp.grid(row=1, column=3, pady=5, padx=10)

tk.Label(frame_input, text="Quantity Delivered:", font=("Helvetica", 10, "bold"), bg="#f4f4f4").grid(row=2, column=0, sticky="w", pady=5)
entry_qty = tk.Entry(frame_input, width=37, font=("Helvetica", 10))
entry_qty.grid(row=2, column=1, pady=5, padx=10)

frame_action_btns = tk.Frame(root, bg="#f4f4f4")
frame_action_btns.pack(pady=10)

btn_add_list = tk.Button(frame_action_btns, text="⬇️ ADD TO STAGING BATCH", font=("Helvetica", 10, "bold"), bg="#17a2b8", fg="white", width=25, command=add_to_staging_list)
btn_add_list.grid(row=0, column=0, padx=10)

btn_del_list = tk.Button(frame_action_btns, text="❌ REMOVE SELECTED ITEM", font=("Helvetica", 10, "bold"), bg="#6c757d", fg="white", width=25, command=delete_selected_staged)
btn_del_list.grid(row=0, column=1, padx=10)

tk.Label(root, text="📋 STAGED ITEMS FOR BATCH PROCESSING (REVIEW BEFORE SAVING):", font=("Helvetica", 10, "bold"), bg="#f4f4f4", fg="#333").pack(anchor="w", padx=25, pady=5)

columns = ("Code", "Product", "Lot", "Expiry", "Qty")
tree_stage = ttk.Treeview(root, columns=columns, show="headings", height=12)
tree_stage.heading("Code", text="Item Code")
tree_stage.heading("Product", text="Product Name")
tree_stage.heading("Lot", text="Lot Number")
tree_stage.heading("Expiry", text="Expiry Date")
tree_stage.heading("Qty", text="Quantity")

tree_stage.column("Code", width=100, anchor='center')
tree_stage.column("Product", width=280)
tree_stage.column("Lot", width=130, anchor='center')
tree_stage.column("Expiry", width=110, anchor='center')
tree_stage.column("Qty", width=80, anchor='center')
tree_stage.pack(pady=5, padx=20, fill="both", expand=True)

btn_final_process = tk.Button(root, text="🚀 PROCESS AND COMMIT BATCH TO MASTERLIST", font=("Helvetica", 12, "bold"), bg="#28a745", fg="white", width=50, height=2, command=process_batch_to_master)
btn_final_process.pack(pady=15)

root.mainloop()