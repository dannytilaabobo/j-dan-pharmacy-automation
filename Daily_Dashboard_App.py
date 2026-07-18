import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import datetime
import os
import random
import warnings
import sys
import traceback
import json
import shutil

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
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
#    SECURITY LOGIN & LOGBOOK SYSTEM
# ==========================================
root = tk.Tk()
root.withdraw()

staff_name = simpledialog.askstring("Security Authentication", "J-DAN EXECUTIVE DASHBOARD\n\nPlease enter Authorized User Name:", parent=root)

if not staff_name or staff_name.strip() == "":
    messagebox.showerror("Access Denied", "System locked. Authorized user name is required to access the dashboard.")
    sys.exit()

current_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
try:
    with open("JDan_System_Logs.txt", "a") as logbook:
        logbook.write(f"[{current_time}] - USER: {staff_name.upper()} - Accessed EXECUTIVE DASHBOARD\n")
except PermissionError:
    pass 

root.deiconify() 

# ==========================================
#    SYSTEM DIRECTORIES
# ==========================================
stockin_base = 'JDan_StockIn_Reports'
sales_folder = os.path.join('JDan_Sales_System', 'EOD_Deduction_Reports')
base_sales_folder = 'JDan_Sales_System' 
pullout_folder = os.path.join('JDan_PullOut_System', '1_Excel_Reports')
exec_folder = 'JDan_Executive_Reports'
master_file = 'Inventory_Master.xlsx' 
sync_folder = 'PENDING_CLOUD_SYNC'

for folder in [stockin_base, sales_folder, pullout_folder, exec_folder, base_sales_folder, sync_folder]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# ==========================================
#    CORE DATA PROCESSING
# ==========================================
def generate_dashboard():
    target_date = entry_date.get().strip()
    
    try:
        target_date_obj = datetime.datetime.strptime(target_date, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Invalid Date Format", "Please enter the date in YYYY-MM-DD format (e.g., 2026-06-20).")
        return

    # --- SMART DATE MATCHER FOR EOD ---
    target_yyyymmdd = target_date_obj.strftime("%Y%m%d")          
    target_mmdd = target_date_obj.strftime("%m%d")                
    target_mon_dd = target_date_obj.strftime("%b%d").upper()      
    target_month_dd = target_date_obj.strftime("%B%d").upper()    
    
    valid_eod_markers = [target_yyyymmdd, target_mmdd, target_mon_dd, target_month_dd]
    
    # ------------------------------------------
    # 1. STOCK-IN DATA 
    # ------------------------------------------
    stockin_today_folder = os.path.join(stockin_base, target_date)
    stockin_data = []
    if os.path.exists(stockin_today_folder):
        for file in os.listdir(stockin_today_folder):
            if file.endswith(('.csv', '.xlsx', '.xls')) and not file.startswith('~'):
                filepath = os.path.join(stockin_today_folder, file)
                if file.endswith('.csv'):
                    df = pd.read_csv(filepath, dtype=str)
                else:
                    df = pd.read_excel(filepath, dtype=str)
                stockin_data.append(df)
    
    df_stockin = pd.concat(stockin_data, ignore_index=True) if stockin_data else pd.DataFrame(columns=['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Qty_Delivered', 'Date_Encoded', 'Encoded_By'])

    # ------------------------------------------
    # 2. SOLD (EOD) DATA 
    # ------------------------------------------
    sales_data = []
    folders_to_check = [sales_folder, base_sales_folder] 
    
    for sf in folders_to_check:
        if os.path.exists(sf):
            for file in os.listdir(sf):
                if file.endswith(('.csv', '.xlsx', '.xls')) and not file.startswith('~'):
                    file_upper = file.upper()
                    if any(marker in file_upper for marker in valid_eod_markers):
                        filepath = os.path.join(sf, file)
                        if os.path.isfile(filepath):
                            if file.endswith('.csv'):
                                df = pd.read_csv(filepath, dtype=str)
                            else:
                                df = pd.read_excel(filepath, dtype=str)
                            sales_data.append(df)
                
    df_sales = pd.concat(sales_data, ignore_index=True) if sales_data else pd.DataFrame()

    # ------------------------------------------
    # 3. PULL-OUTS DATA 
    # ------------------------------------------
    pullout_data = []
    if os.path.exists(pullout_folder):
        for file in os.listdir(pullout_folder):
            if file.endswith(('.csv', '.xlsx', '.xls')) and target_yyyymmdd in file and not file.startswith('~'):
                filepath = os.path.join(pullout_folder, file)
                if file.endswith('.csv'):
                    df = pd.read_csv(filepath, dtype=str)
                else:
                    df = pd.read_excel(filepath, dtype=str)
                pullout_data.append(df)
                
    df_pullout = pd.concat(pullout_data, ignore_index=True) if pullout_data else pd.DataFrame(columns=['Item_Code', 'Product_Name', 'Lot_Number', 'Expiry_Date', 'Current_Stock', 'Status', 'Pulled_Out_By'])

    # ------------------------------------------
    # 4. MASTER INVENTORY READ 
    # ------------------------------------------
    stock_dict = {}
    if os.path.exists(master_file):
        df_master = pd.read_excel(master_file, sheet_name='INVENTORY MASTER')
        if 'Expiry_Date' in df_master.columns:
            df_master['Expiry_Date'] = df_master['Expiry_Date'].fillna('NO-EXP')
        else:
            df_master['Expiry_Date'] = 'NO-EXP'
            
        group_cols = [c for c in ['Item_Code', 'Product_Name', 'Expiry_Date'] if c in df_master.columns]
        
        if 'Current_Stock' in df_master.columns:
            df_master['Current_Stock'] = pd.to_numeric(df_master['Current_Stock'], errors='coerce').fillna(0)
            df_final_inventory = df_master.groupby(group_cols, dropna=False).agg({'Current_Stock': 'sum'}).reset_index()
            
            if 'Product_Name' in df_final_inventory.columns:
                stock_dict = df_final_inventory.groupby(df_final_inventory['Product_Name'].astype(str).str.strip().str.upper())['Current_Stock'].sum().to_dict()
        else:
            df_final_inventory = df_master
    else:
        df_final_inventory = pd.DataFrame({'SYSTEM_ERROR': ['CRITICAL: Inventory_Master.xlsx NOT FOUND.']})
        df_master = pd.DataFrame()

    # ------------------------------------------
    # 5. DYNAMIC 20 SOD STRATIFIER GENERATOR 
    # ------------------------------------------
    sod_descriptions = []
    df_top_20_sod = pd.DataFrame()

    if not df_sales.empty:
        desc_col = 'Description' if 'Description' in df_sales.columns else ('Product_Name' if 'Product_Name' in df_sales.columns else None)
        qty_col = 'Qty' if 'Qty' in df_sales.columns else ('Qty_Deducted' if 'Qty_Deducted' in df_sales.columns else None)
        sales_col = 'Sales' if 'Sales' in df_sales.columns else None
        
        if desc_col and qty_col and sales_col:
            df_calc = df_sales.copy()
            df_calc[qty_col] = pd.to_numeric(df_calc[qty_col], errors='coerce').fillna(0).abs()
            df_calc[sales_col] = pd.to_numeric(df_calc[sales_col], errors='coerce').fillna(0).abs()
            
            df_calc['Unit_Price'] = (df_calc[sales_col] / df_calc[qty_col]).fillna(0)
            
            item_stats = df_calc.groupby(desc_col).agg({qty_col: 'sum', 'Unit_Price': 'max'}).reset_index()
            
            fast_moving = item_stats.sort_values(by=qty_col, ascending=False).head(10)
            remaining_for_exp = item_stats[~item_stats[desc_col].isin(fast_moving[desc_col])]
            expensive = remaining_for_exp.sort_values(by='Unit_Price', ascending=False).head(5)
            
            picked_so_far = pd.concat([fast_moving, expensive])[desc_col].tolist()
            if 'Product_Name' in df_master.columns:
                master_items = df_master['Product_Name'].dropna().unique().tolist()
                remaining_master = [str(item) for item in master_items if str(item) not in picked_so_far]
                
                random_count = min(5, len(remaining_master))
                random_picks = random.sample(remaining_master, random_count) if random_count > 0 else []
                
                random_df = pd.DataFrame({
                    desc_col: random_picks,
                    qty_col: 0,
                    'Unit_Price': 0
                })
            else:
                random_df = pd.DataFrame()

            df_top_20_sod = pd.concat([fast_moving, expensive, random_df], ignore_index=True)
            df_top_20_sod.rename(columns={desc_col: 'Item_Name', qty_col: 'Sold_Today'}, inplace=True)
            
            df_top_20_sod['System_Expected_Qty'] = df_top_20_sod['Item_Name'].apply(lambda x: stock_dict.get(str(x).strip().upper(), 0.0))
            df_top_20_sod = df_top_20_sod[['Item_Name', 'Sold_Today', 'System_Expected_Qty']]
            sod_descriptions = df_top_20_sod['Item_Name'].astype(str).str.strip().str.upper().tolist()

    if df_sales.empty:
        df_sales = pd.DataFrame(columns=['Item_Code', 'Description', 'Qty', 'Sales', 'Receipt'])
    if df_top_20_sod.empty:
        messagebox.showinfo("No Data", f"No transaction data found for SOD Stratification for {target_date}.\n\nGenerating Inventory Snapshot instead.")
        df_top_20_sod = pd.DataFrame(columns=['Item_Name', 'Sold_Today', 'System_Expected_Qty', 'Actual_SOD_Count', 'Variance', 'Justification'])

    if 'Product_Name' in df_final_inventory.columns and sod_descriptions:
        df_final_inventory['Priority'] = df_final_inventory['Product_Name'].apply(
            lambda x: 1 if str(x).strip().upper() in sod_descriptions else 2
        )
        df_final_inventory['SOD_Stratifier_Tag'] = df_final_inventory['Priority'].apply(
            lambda x: '🎯 SELECTED FOR SOD AUDIT' if x == 1 else ''
        )
        df_final_inventory = df_final_inventory.sort_values(by=['Priority', 'Product_Name', 'Expiry_Date']).drop(columns=['Priority'])
        cols = ['SOD_Stratifier_Tag'] + [c for c in df_final_inventory.columns if c != 'SOD_Stratifier_Tag']
        df_final_inventory = df_final_inventory[cols]

    if not df_top_20_sod.empty:
        launch_stratifier_ui(target_date, df_stockin, df_sales, df_pullout, df_final_inventory, df_top_20_sod)
    else:
        export_final_report(target_date, df_stockin, df_sales, df_pullout, df_final_inventory, df_top_20_sod)

# ==========================================
#    PHASE 2: INTERACTIVE STRATIFIER UI
# ==========================================
def launch_stratifier_ui(target_date, df_stockin, df_sales, df_pullout, df_final_inventory, df_top_20_sod):
    strat_win = tk.Toplevel(root)
    strat_win.title(f"Phase 2: SOD Stratification Audit - {target_date}")
    strat_win.geometry("900x700") 
    strat_win.configure(bg="#eaeded")
    strat_win.grab_set() 
    
    tk.Label(strat_win, text=f"START OF DAY (SOD) STRATIFIER CHECK", fg="white", bg="#154360", font=("Arial", 14, "bold"), pady=10).pack(fill="x")
    tk.Label(strat_win, text="Fill out the Actual Count. Justification is MANDATORY if Variance is detected.", bg="#eaeded", font=("Arial", 10, "italic"), fg="#cc0000").pack(pady=5)

    grid_frame = tk.Frame(strat_win, bg="#eaeded")
    grid_frame.pack(fill="both", expand=True, padx=20, pady=10)

    headers = ["Item Name", "System Qty", "Actual Count", "Justification (If Variance)"]
    widths = [45, 12, 12, 40]
    
    for col, head in enumerate(headers):
        tk.Label(grid_frame, text=head, font=("Arial", 10, "bold"), bg="#d5d8dc", borderwidth=1, relief="solid", width=widths[col]).grid(row=0, column=col, sticky="ew")

    entry_actuals = []
    entry_justs = []
    
    for idx, (index, row) in enumerate(df_top_20_sod.iterrows()):
        tk.Label(grid_frame, text=row['Item_Name'], font=("Arial", 9), bg="white", borderwidth=1, relief="solid", anchor="w", padx=5).grid(row=idx+1, column=0, sticky="ew")
        tk.Label(grid_frame, text=f"{row['System_Expected_Qty']:.0f}", font=("Arial", 10, "bold"), bg="#fcf3cf", borderwidth=1, relief="solid").grid(row=idx+1, column=1, sticky="ew")
        
        act_entry = tk.Entry(grid_frame, font=("Arial", 10), justify="center")
        act_entry.grid(row=idx+1, column=2, sticky="ew", padx=2, pady=2)
        entry_actuals.append(act_entry)
        
        just_entry = tk.Entry(grid_frame, font=("Arial", 9))
        just_entry.grid(row=idx+1, column=3, sticky="ew", padx=2, pady=2)
        entry_justs.append(just_entry)

    def submit_and_export():
        final_sod_data = []
        
        for idx, (index, row) in enumerate(df_top_20_sod.iterrows()):
            act_val = entry_actuals[idx].get().strip().replace(',', '')
            just_val = entry_justs[idx].get().strip()
            
            entry_justs[idx].config(bg="white")
            
            if not act_val:
                messagebox.showerror("Validation Error", f"Missing Actual Count for:\n\n{row['Item_Name']}")
                return
                
            try:
                actual_qty = float(act_val)
            except ValueError:
                messagebox.showerror("Validation Error", f"Invalid number format (e.g. letters found) for:\n\n{row['Item_Name']}")
                return
                
            expected_qty = float(row['System_Expected_Qty'])
            variance = actual_qty - expected_qty
            
            if abs(variance) > 0.01 and just_val == "":
                entry_justs[idx].config(bg="#ffcccc")
                messagebox.showwarning(
                    "MANDATORY JUSTIFICATION REQUIRED", 
                    f"⚠️ INVALID! VARIANCE DETECTED!\n\n"
                    f"Item: {row['Item_Name']}\n"
                    f"System Qty: {expected_qty:.0f}\n"
                    f"Actual Count: {actual_qty:.0f}\n"
                    f"Variance: {variance:.0f}\n\n"
                    f"You CANNOT proceed. Please type a justification for this specific item before finalizing."
                )
                return
                
            final_sod_data.append({
                'Item_Name': row['Item_Name'],
                'Sold_Today': row['Sold_Today'],
                'System_Expected_Qty': expected_qty,
                'Actual_SOD_Count': actual_qty,
                'Variance': variance,
                'Justification': just_val if abs(variance) > 0.01 else "Balanced"
            })

        df_export_sod = pd.DataFrame(final_sod_data)

        # ==========================================
        # 🔥 RECONCILE MASTER: ang Actual Count na binilang ang dapat
        # maging bagong Current_Stock -- dati, dito lang na-lo-log ang
        # variance/justification pero hindi na kino-correct ang Master,
        # kaya paulit-ulit na nagpapatuloy ang parehong mismatch araw-araw.
        # ==========================================
        try:
            df_master_live = pd.read_excel(master_file, sheet_name='INVENTORY MASTER')
        except PermissionError:
            messagebox.showerror("Inventory Locked", f"Cannot reconcile {master_file} because it is open in Excel. Please close it and FINALIZE again.")
            return
        df_master_live['Current_Stock'] = pd.to_numeric(df_master_live['Current_Stock'], errors='coerce').fillna(0)

        reconcile_log = []
        for r in final_sod_data:
            item_name_clean = str(r['Item_Name']).strip().upper()
            actual_qty = r['Actual_SOD_Count']
            mask = df_master_live['Product_Name'].astype(str).str.strip().str.upper() == item_name_clean
            if not mask.any():
                continue
            old_total = df_master_live.loc[mask, 'Current_Stock'].sum()
            delta = actual_qty - old_total
            if abs(delta) < 0.01:
                continue

            # Ipinapamahagi ang adjustment FIFO/FEFO-style (kaparehong convention
            # ng Daily_Deduct_App/Expiry_Scanner): ang pinaka-malapit-nang-mag-expire
            # na lot ang unang babawasan kung shortage; ang overage naman ay
            # ilalagay din doon dahil walang paraan para malaman kung aling
            # aktwal na lot napunta ang sobrang units.
            parsed_dates = pd.to_datetime(df_master_live.loc[mask, 'Expiry_Date'], errors='coerce').fillna(pd.Timestamp('2099-12-31'))
            ordered_idx = parsed_dates.sort_values().index.tolist()

            if delta < 0:
                qty_to_remove = abs(delta)
                for idx in ordered_idx:
                    current = df_master_live.at[idx, 'Current_Stock']
                    if current >= qty_to_remove:
                        df_master_live.at[idx, 'Current_Stock'] = current - qty_to_remove
                        qty_to_remove = 0
                        break
                    qty_to_remove -= current
                    df_master_live.at[idx, 'Current_Stock'] = 0
            else:
                df_master_live.at[ordered_idx[0], 'Current_Stock'] += delta

            reconcile_log.append({
                'Item_Name': r['Item_Name'],
                'Old_System_Total': old_total,
                'New_Actual_Total': actual_qty,
                'Adjustment': delta,
            })

        if reconcile_log:
            os.makedirs('JDan_Master_Backups', exist_ok=True)
            backup_ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            shutil.copy2(master_file, os.path.join('JDan_Master_Backups', f'Inventory_Master_BEFORE_SOD_RECONCILE_{backup_ts}.xlsx'))
            try:
                with pd.ExcelWriter(master_file, engine='openpyxl', mode='w') as writer:
                    df_master_live.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
            except PermissionError:
                messagebox.showerror("Inventory Locked", f"The reconciliation was not saved because {master_file} is open. Please close it and FINALIZE again.")
                return

            reconcile_time = datetime.datetime.now().strftime('%Y-%m-%d %I:%M %p')
            with open("JDan_System_Logs.txt", "a") as logbook:
                logbook.write(f"[{reconcile_time}] - SOD RECONCILIATION by {staff_name.upper()} ({target_date})\n")
                for entry in reconcile_log:
                    logbook.write(f"   -> {entry['Item_Name']} | System: {entry['Old_System_Total']:.0f} -> Actual: {entry['New_Actual_Total']:.0f} | Adjustment: {entry['Adjustment']:+.0f}\n")
        # ==========================================

        # 🔥 NEW: GUMAWA NG TEXT REPORT AT I-QUEUE SA JSON PARA SA CLOUD SYNC
        timestamp = datetime.datetime.now().strftime("%H%M%S")
        txt_report_name = f"SOD_Summary_{target_date}_{timestamp}.txt"
        txt_path = os.path.join(exec_folder, txt_report_name)
        
        total_variances = 0
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(f"=======================================\n")
            f.write(f"START OF DAY (SOD) STRATIFICATION AUDIT\n")
            f.write(f"Target Date: {target_date} | Auditor: {staff_name.upper()}\n")
            f.write(f"=======================================\n")
            for _, r in df_export_sod.iterrows():
                f.write(f"Item: {r['Item_Name']}\n")
                f.write(f"System Qty: {r['System_Expected_Qty']} | Actual Count: {r['Actual_SOD_Count']}\n")
                f.write(f"Variance: {r['Variance']}\n")
                f.write(f"Justification: {r['Justification']}\n")
                f.write(f"---------------------------------------\n")
                if abs(r['Variance']) > 0:
                    total_variances += 1
            f.write(f"Total Items with Variance: {total_variances}\n")
            f.write(f"=======================================\n")

        # JSON Payload para sa Courier
        queue_payload = {
            "target_tab": "SOD_Audits", # Ito ang babasahin ng Courier
            "row_data": [datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), staff_name.upper(), target_date, f"{total_variances} mismatches"],
            "report_filename": txt_report_name
        }
        
        pending_file = os.path.join(sync_folder, f"sync_SOD_{target_date}_{timestamp}.json")
        with open(pending_file, "w") as f:
            json.dump(queue_payload, f)

        # Ituloy ang normal na Excel Export
        export_success = export_final_report(target_date, df_stockin, df_sales, df_pullout, df_final_inventory, df_export_sod)
        
        if export_success:
            strat_win.destroy()
            entry_date.delete(0, tk.END)
            entry_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))

    btn_frame = tk.Frame(strat_win, bg="#eaeded")
    btn_frame.pack(pady=15)
    tk.Button(btn_frame, text="✅ FINALIZE EOD & GENERATE REPORT", font=("Arial", 11, "bold"), bg="#28b463", fg="white", height=2, width=35, command=submit_and_export).pack()

def export_final_report(target_date, df_stockin, df_sales, df_pullout, df_final_inventory, df_top_20_sod):
    timestamp = datetime.datetime.now().strftime("%H%M%S")
    report_name = f"EXECUTIVE_REPORT_{target_date}_{timestamp}.xlsx"
    report_path = os.path.join(exec_folder, report_name)

    try:
        with pd.ExcelWriter(report_path, engine='openpyxl', mode='w') as writer:
            df_stockin.to_excel(writer, sheet_name='1_STOCK_IN', index=False)
            df_sales.to_excel(writer, sheet_name='2_SOLD', index=False)
            df_pullout.to_excel(writer, sheet_name='3_PULLED_OUT', index=False)
            df_final_inventory.to_excel(writer, sheet_name='4_FINAL_INVENTORY', index=False)
            df_top_20_sod.to_excel(writer, sheet_name='5_SOD_AUDIT_REPORT', index=False)
            
        messagebox.showinfo("Export Success", f"EOD Report and Stratification complete!\n\n1. File Saved: {report_name}\n2. SOD Report is queued in the background to sync to Google Sheets.")
        return True
            
    except PermissionError:
        messagebox.showerror("Permission Error", "The report could not be saved because a file with the same name is still open in Excel. Please close Excel first.")
        return False
    except Exception as e:
        messagebox.showerror("Export Error", f"Failed to save the report.\n\nError Details: {e}")
        return False

# ==========================================
#    MAIN DASHBOARD GUI
# ==========================================
root.title("J-DAN PHARMACY | Executive Operations Dashboard")
root.geometry("550x380")
root.configure(bg="#eaeded")
root.resizable(False, False)

header_frame = tk.Frame(root, bg="#154360", pady=20)
header_frame.pack(fill="x")
tk.Label(header_frame, text="EXECUTIVE DASHBOARD", fg="white", bg="#154360", font=("Arial", 16, "bold")).pack()
tk.Label(header_frame, text="Centralized Operations & Phase 2 Stratification", fg="#a9cce3", bg="#154360", font=("Arial", 10)).pack()

main_frame = tk.Frame(root, bg="#eaeded", pady=30)
main_frame.pack(fill="both", expand=True)

tk.Label(main_frame, text="Select Target Date for Consolidation:", font=("Arial", 11, "bold"), bg="#eaeded", fg="#333333").pack(pady=(0, 10))

frame_date = tk.Frame(main_frame, bg="#eaeded")
frame_date.pack(pady=5)
tk.Label(frame_date, text="Reporting Date (YYYY-MM-DD):", font=("Arial", 10), bg="#eaeded").pack(side="left", padx=5)

entry_date = ttk.Entry(frame_date, font=("Arial", 11), width=18)
entry_date.insert(0, datetime.datetime.now().strftime("%Y-%m-%d")) 
entry_date.pack(side="left", padx=5)

btn_generate = tk.Button(main_frame, text="ENTER STRATIFICATION (PHASE 2)", font=("Arial", 11, "bold"), bg="#cb4335", fg="white", height=2, width=35, relief="flat", cursor="hand2", command=generate_dashboard)
btn_generate.pack(pady=25)

footer_frame = tk.Frame(root, bg="#d5d8dc", pady=5)
footer_frame.pack(fill="x", side="bottom")
tk.Label(footer_frame, text=f"Authorized System User: {staff_name.upper()}", font=("Arial", 8, "bold"), bg="#d5d8dc", fg="#566573").pack()

root.mainloop()