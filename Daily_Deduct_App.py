import os
import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import warnings
import json
import shutil
import traceback
import sys

from name_cleaning import clean_text

# ==========================================
#    CRASH DETECTOR (Anti-Silent Crash)
# ==========================================
def handle_exception(exc_type, exc_value, exc_traceback):
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    messagebox.showerror("System Critical Error", f"Nakatagpo ng error ang system. Paki-screenshot ito at ipadala sa Admin:\n\n{err_msg}")
    
sys.excepthook = handle_exception
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==========================================
# SYSTEM CONFIGURATION DIRECTORIES
# ==========================================
INBOX_DIR = "SALES_REPORT_FROM_POS"
BASE_SYS_DIR = "JDan_Sales_System"
DEDUCTION_DIR = os.path.join(BASE_SYS_DIR, "EOD_Deduction_Reports")
ARCHIVE_DIR = os.path.join(BASE_SYS_DIR, "Archived_Sales_Reports")
REPORT_DIR = "JDan_Executive_Reports"
PENDING_SYNC_DIR = "PENDING_CLOUD_SYNC"
MASTER_FILE = "Inventory_Master.xlsx" 

for directory in [INBOX_DIR, DEDUCTION_DIR, ARCHIVE_DIR, REPORT_DIR, PENDING_SYNC_DIR]:
    if not os.path.exists(directory):
        os.makedirs(directory)

# ==========================================
# ROBUST POS FILE READER (FINAL FIXED LOGIC)
# ==========================================
def read_pos_file(filepath):
    if filepath.endswith('.csv'):
        df_raw = pd.read_csv(filepath, header=None, dtype=str)
    else:
        df_raw = pd.read_excel(filepath, header=None, dtype=str)

    header_idx = None
    for i, row in df_raw.iterrows():
        row_strs = [str(x).strip().upper() for x in row.values]
        if any('ITEM CODE' in s for s in row_strs) and any('QTY' in s for s in row_strs):
            header_idx = i
            break
            
    if header_idx is None:
        raise ValueError("Hindi mahanap ang 'ITEM CODE' o 'QTY' columns sa file.\nSiguraduhing EOD SALES REPORT ito.")

    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = [str(x).strip().title() for x in df_raw.iloc[header_idx].values]
    
    # TANGGALIN ANG BLANGKONG LINYA PARA HINDI MAGING VOID
    df = df.dropna(subset=['Item Code', 'Description'], how='all')

    if 'Item Code' in df.columns:
        df['Item Code'] = df['Item Code'].fillna('MANUAL-ENTRY')
        df.loc[df['Item Code'].astype(str).str.strip().isin(['Nan', 'nan', '', 'None']), 'Item Code'] = 'MANUAL-ENTRY'
    
    if 'Description' in df.columns:
        df['Description'] = df['Description'].fillna('Unknown / Manual Entry').apply(clean_text)

    if 'Item Code' in df.columns:
        df = df[~df['Item Code'].astype(str).str.contains('Total', case=False, na=False)]
    if 'Date' in df.columns:
        df = df[~df['Date'].astype(str).str.contains('Total', case=False, na=False)]

    df['Qty'] = pd.to_numeric(df['Qty'], errors='coerce').fillna(0)
    df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce').fillna(0.0)
    
    return df

# ==========================================
# MAIN APPLICATION GUI
# ==========================================
class EODSettlementApp:
    def __init__(self, root, df, file_path, processor_name):
        self.root = root
        self.root.title("J-DAN Pharmacy - EOD Settlement System (Offline-Ready)")
        self.root.geometry("550x740")
        self.root.resizable(False, False)
        
        self.df = df
        self.file_path = file_path
        self.processor_name = processor_name
        
        if 'Cashier' in df.columns and not pd.isna(df['Cashier'].iloc[0]):
            self.cashier_name = str(df['Cashier'].iloc[0])
        elif 'Cashier Id' in df.columns and not pd.isna(df['Cashier Id'].iloc[0]):
            self.cashier_name = str(df['Cashier Id'].iloc[0])
        else:
            self.cashier_name = "Unknown"
            
        # 🔥 FIX: IBINALIK NATIN SA DATING COMPUTATION, WALA NANG DISCOUNT NA IDADAGDAG
        self.system_total_sales = float(df['Sales'].sum()) if 'Sales' in df.columns else 0.0
        
        self.voids_df = df[df['Qty'] == 0] if 'Qty' in df.columns else pd.DataFrame()
        self.void_count = len(self.voids_df)
        
        self.void_item_names = ""
        self.default_void_text = ""
        
        if self.void_count > 0:
            void_display_list = []
            void_text_box_list = []
            for _, row in self.voids_df.iterrows():
                receipt_no = str(row.get('Receipt', 'UNKNOWN')).strip()
                desc = str(row.get('Description', 'Unknown Item')).strip()
                void_display_list.append(f"Receipt #{receipt_no} ({desc})")
                void_text_box_list.append(f"Receipt #{receipt_no} - [Type Reason Here]")
                
            self.void_item_names = "\n".join(void_display_list)
            self.default_void_text = "\n".join(void_text_box_list)
        else:
            self.void_item_names = "None"
            self.default_void_text = "N/A - No voids detected"
            
        self.z_reading_justification = "Matched"
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        header_frame = tk.Frame(self.root, bg="#003366", pady=15)
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="END OF DAY SETTLEMENT", fg="white", bg="#003366", font=("Arial", 14, "bold")).pack()
        tk.Label(header_frame, text=f"Processed By: {self.processor_name.upper()} | Cashier: {self.cashier_name.upper()}", fg="white", bg="#003366", font=("Arial", 10)).pack()
        tk.Label(header_frame, text=f"Date: {datetime.datetime.now().strftime('%Y-%m-%d')}", fg="white", bg="#003366", font=("Arial", 9)).pack()

        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="System Extracted Data (From POS CSV)", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10), columnspan=2)
        
        ttk.Label(main_frame, text="System Total Sales:").grid(row=1, column=0, sticky="w", pady=2)
        tk.Label(main_frame, text=f"₱ {self.system_total_sales:,.2f}", font=("Arial", 11, "bold"), fg="#003366").grid(row=1, column=1, sticky="w")

        ttk.Label(main_frame, text="Void Transactions Detected:").grid(row=2, column=0, sticky="nw", pady=5)
        void_color = "red" if self.void_count > 0 else "green"
        void_display_text = f"{self.void_count} Item/s:\n{self.void_item_names}" if self.void_count > 0 else "0"
        tk.Label(main_frame, text=void_display_text, font=("Arial", 9, "bold"), fg=void_color, justify="left", wraplength=280).grid(row=2, column=1, sticky="w")

        ttk.Label(main_frame, text="Void Justification:").grid(row=3, column=0, sticky="nw", pady=5)
        self.void_reason_entry = tk.Text(main_frame, width=35, height=4, font=("Arial", 10))
        self.void_reason_entry.grid(row=3, column=1, sticky="w", pady=5)
        self.void_reason_entry.insert("1.0", self.default_void_text)
        if self.void_count == 0:
            self.void_reason_entry.config(state="disabled")

        ttk.Separator(main_frame, orient='horizontal').grid(row=4, column=0, columnspan=2, sticky='ew', pady=15)

        ttk.Label(main_frame, text="Cash Reconciliation (MOP)", font=("Arial", 11, "bold")).grid(row=5, column=0, sticky="w", pady=(0, 10), columnspan=2)

        ttk.Label(main_frame, text="Z-Reading Total Sales (₱):").grid(row=6, column=0, sticky="w", pady=5)
        self.z_reading_entry = ttk.Entry(main_frame, width=20)
        self.z_reading_entry.grid(row=6, column=1, sticky="w", pady=5)

        ttk.Label(main_frame, text="Initial Cash / Pondo (₱):").grid(row=7, column=0, sticky="w", pady=5)
        self.initial_cash_entry = ttk.Entry(main_frame, width=20)
        self.initial_cash_entry.grid(row=7, column=1, sticky="w", pady=5)

        ttk.Label(main_frame, text="Actual Cash Counted (₱):").grid(row=8, column=0, sticky="w", pady=5)
        self.actual_cash_entry = ttk.Entry(main_frame, width=20)
        self.actual_cash_entry.grid(row=8, column=1, sticky="w", pady=5)
        
        ttk.Label(main_frame, text="GCash Total (₱):").grid(row=9, column=0, sticky="w", pady=5)
        self.gcash_entry = ttk.Entry(main_frame, width=20)
        self.gcash_entry.grid(row=9, column=1, sticky="w", pady=5)
        self.gcash_entry.insert(0, "0.00") 
        
        ttk.Label(main_frame, text="Charge Total (₱):").grid(row=10, column=0, sticky="w", pady=5)
        self.charge_entry = ttk.Entry(main_frame, width=20)
        self.charge_entry.grid(row=10, column=1, sticky="w", pady=5)
        self.charge_entry.insert(0, "0.00") 

        btn_frame = tk.Frame(self.root, pady=20)
        btn_frame.pack(fill="x", side="bottom")
        
        self.submit_btn = tk.Button(btn_frame, text="PROCESS SETTLEMENT", bg="#003366", fg="white", font=("Arial", 11, "bold"), width=28, height=2, command=self.process_settlement)
        self.submit_btn.pack()

    def process_settlement(self):
        void_just = self.void_reason_entry.get("1.0", "end-1c").strip()
        if self.void_count > 0:
            if "[Type Reason Here]" in void_just or not void_just:
                messagebox.showwarning("Validation Error", "Mandatory: Please replace '[Type Reason Here]' with a specific justification for EACH receipt.")
                return

        try:
            z_reading = float(self.z_reading_entry.get().replace(',', '') or 0)
            initial_cash = float(self.initial_cash_entry.get().replace(',', '') or 0)
            actual_cash = float(self.actual_cash_entry.get().replace(',', '') or 0)
            gcash_total = float(self.gcash_entry.get().replace(',', '') or 0)
            charge_total = float(self.charge_entry.get().replace(',', '') or 0)
        except ValueError:
            messagebox.showerror("Format Error", "Please enter valid numeric values for all cash fields. Do not use letters.")
            return

        # Babala kapag magkaiba ang Tape at Export (Halimbawa: dahil sa Non-Inventory fees na tinype ng cashier sa POS)
        if abs(z_reading - self.system_total_sales) > 0.1:
            z_diff = z_reading - self.system_total_sales
            ans = simpledialog.askstring(
                "Z-Reading Mismatch Detected", 
                f"Magkaiba ang Z-Reading Tape (₱{z_reading:,.2f}) at System CSV Export (₱{self.system_total_sales:,.2f}).\n\nMay variance na ₱{z_diff:,.2f}. Marahil ito ay GCash Fee o Special Charge.\n\nPara makatuloy, i-type ang justification (Halimbawa: 'GCash Fee'):"
            )
            if not ans or ans.strip() == "":
                messagebox.showerror("Process Aborted", "Kailangan ng justification para sa Z-Reading Mismatch.")
                return
            self.z_reading_justification = f"Mismatch (Diff: {z_diff:,.2f}) - Reason: {ans}"

        total_accounted = actual_cash + gcash_total + charge_total
        expected_total = z_reading + initial_cash
        variance = total_accounted - expected_total
        cash_just = "Balanced"

        if abs(variance) > 0.1:
            status = "OVER" if variance > 0 else "SHORT"
            cash_just = simpledialog.askstring(
                "Cash Variance Detected", 
                f"Settlement is {status} by ₱ {abs(variance):,.2f}.\nExpected: ₱ {expected_total:,.2f}\nAccounted (Cash+GCash+Charge): ₱ {total_accounted:,.2f}\n\nEnter Override Justification:"
            )
            if not cash_just or cash_just.strip() == "":
                messagebox.showerror("Process Aborted", "Settlement halted. Variance requires a mandatory justification.")
                return

        self.execute_ledger_deduction(initial_cash, actual_cash, gcash_total, charge_total, variance, void_just, cash_just, z_reading)

    def execute_ledger_deduction(self, initial_cash, actual_cash, gcash_total, charge_total, variance, void_just, cash_just, z_reading):
        expected_total = z_reading + initial_cash
        total_accounted = actual_cash + gcash_total + charge_total
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        timestamp_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        
        original_filename = os.path.basename(self.file_path)
        base_name = os.path.splitext(original_filename)[0].upper() 
        
        valid_sales = self.df[self.df['Qty'] > 0].copy()
        if not valid_sales.empty:
            valid_sales['Qty'] = valid_sales['Qty'] * -1
            
        output_filename = f"PROCESSED_EOD_{base_name}.xlsx" 
        output_filepath = os.path.join(DEDUCTION_DIR, output_filename)
        valid_sales.to_excel(output_filepath, index=False)

        # =======================================================
        # ─── AUTOMATED LOCAL INVENTORY DEDUCTION (FIFO RULE) ───
        # =======================================================
        sync_success = False
        if os.path.exists(MASTER_FILE):
            try:
                df_master = pd.read_excel(MASTER_FILE, sheet_name='INVENTORY MASTER')
                if 'Current_Stock' in df_master.columns and 'Product_Name' in df_master.columns:
                    df_master['Current_Stock'] = pd.to_numeric(df_master['Current_Stock'], errors='coerce').fillna(0)
                    
                    df_sales_calc = valid_sales.copy()
                    df_sales_calc['Qty'] = pd.to_numeric(df_sales_calc['Qty'], errors='coerce').fillna(0).abs()
                    desc_col = 'Description' if 'Description' in df_sales_calc.columns else ('Product_Name' if 'Product_Name' in df_sales_calc.columns else 'Item Code')
                    deductions = df_sales_calc.groupby(desc_col)['Qty'].sum().to_dict()
                    
                    for item_name, deduct_qty in deductions.items():
                        if deduct_qty <= 0: continue
                        item_name_clean = str(item_name).strip().upper()
                        
                        matching_mask = df_master['Product_Name'].astype(str).str.strip().str.upper() == item_name_clean
                        matching_rows = df_master[matching_mask]
                        
                        if not matching_rows.empty:
                            df_master['Temp_Is_Blank'] = df_master['Expiry_Date'].isna() | (df_master['Expiry_Date'].astype(str).str.strip().str.upper() == 'NO-EXP')
                            df_master['Temp_Parsed_Date'] = pd.to_datetime(df_master['Expiry_Date'], errors='coerce').fillna(pd.Timestamp('2099-12-31'))
                            
                            indices_to_process = matching_rows.index.tolist()
                            indices_to_process.sort(key=lambda idx: (df_master.loc[idx, 'Temp_Is_Blank'], df_master.loc[idx, 'Temp_Parsed_Date']))
                            
                            qty_to_deduct = deduct_qty
                            for idx in indices_to_process:
                                current_stock = df_master.at[idx, 'Current_Stock']
                                if current_stock >= qty_to_deduct:
                                    df_master.at[idx, 'Current_Stock'] -= qty_to_deduct
                                    qty_to_deduct = 0
                                    break
                                else:
                                    df_master.at[idx, 'Current_Stock'] = 0
                                    qty_to_deduct -= current_stock
                                    
                    if 'Temp_Is_Blank' in df_master.columns: df_master = df_master.drop(columns=['Temp_Is_Blank'])
                    if 'Temp_Parsed_Date' in df_master.columns: df_master = df_master.drop(columns=['Temp_Parsed_Date'])
                    
                    with pd.ExcelWriter(MASTER_FILE, engine='openpyxl', mode='w') as writer:
                        df_master.to_excel(writer, sheet_name='INVENTORY MASTER', index=False)
                    sync_success = True
            except PermissionError:
                messagebox.showerror("Inventory Locked", f"CRITICAL: Hindi nabawasan ang {MASTER_FILE} dahil kasalukuyan itong NAKABUKAS sa PC mo. Pakisara ang Excel file at i-run ulit ang EOD.")
                return
            except Exception as e:
                print(f"Auto Sync Error: {e}")

        # Write Local Text Summary
        txt_report_name = f"EOD_Summary_{base_name}.txt"
        report_path = os.path.join(REPORT_DIR, txt_report_name)
        with open(report_path, "a", encoding="utf-8") as f:
            f.write(f"=======================================\n")
            f.write(f"EOD SETTLEMENT REPORT - {current_time}\n")
            f.write(f"Processed By: {self.processor_name} | Cashier: {self.cashier_name}\n")
            f.write(f"Source File Date Tag: {base_name}\n")
            f.write(f"=======================================\n")
            f.write(f"[ INVENTORY STATUS ]\n")
            f.write(f"System Export Sales: ₱ {self.system_total_sales:,.2f}\n")
            f.write(f"Z-Reading Match Status: {self.z_reading_justification}\n")
            f.write(f"Local Master Stock Deducted: {'YES' if sync_success else 'FAILED'}\n")
            f.write(f"Void Transactions: {self.void_count}\n")
            f.write(f"Void Justifications:\n{void_just}\n\n")
            f.write(f"[ FUND RECONCILIATION ]\n")
            f.write(f"Initial Cash (Float): ₱ {initial_cash:,.2f}\n")
            f.write(f"POS Z-Reading: ₱ {z_reading:,.2f}\n")
            f.write(f"Expected Total Funds: ₱ {expected_total:,.2f}\n")
            f.write(f"---------------------------------------\n")
            f.write(f"Actual Cash Counted: ₱ {actual_cash:,.2f}\n")
            f.write(f"GCash Declared: ₱ {gcash_total:,.2f}\n")
            f.write(f"Charge Declared: ₱ {charge_total:,.2f}\n")
            f.write(f"Total Funds Accounted: ₱ {total_accounted:,.2f}\n")
            f.write(f"---------------------------------------\n")
            f.write(f"Variance (Over/Short): ₱ {variance:,.2f}\n")
            f.write(f"Cash Justification: {cash_just}\n")
            f.write(f"=======================================\n")

        # Queue payload for Cloud Sync Courier
        row_data = [
            current_time, self.processor_name, self.cashier_name, self.system_total_sales, 
            initial_cash, actual_cash, gcash_total, charge_total, variance, cash_just, 
            self.void_count, void_just
        ]
        
        queue_payload = {
            "row_data": row_data,
            "report_filename": txt_report_name
        }
        
        pending_file = os.path.join(PENDING_SYNC_DIR, f"sync_{base_name}_{timestamp_str}.json")
        with open(pending_file, "w") as f:
            json.dump(queue_payload, f)

        # ARCHIVE ORIGINAL POS FILE
        try:
            archive_path = os.path.join(ARCHIVE_DIR, f"{base_name}_{timestamp_str}_{original_filename}")
            shutil.move(self.file_path, archive_path)
            archive_status = "Awtomatikong inilipat ang POS file sa Archive."
        except PermissionError:
            archive_status = f"BABALA: Hindi nailipat ang '{original_filename}' dahil NAKABUKAS pa ito sa Excel. Pakisara at ilipat manual sa Archive folder."
        except Exception as e:
            archive_status = f"BABALA: Hindi nailipat ang POS file: {e}"
        
        messagebox.showinfo("EOD SUCCESS", f"Settlement Complete!\n\n1. Nag-generate ng report para sa: {base_name}\n2. Awtomatikong NABAWASAN na ang Live Master Inventory sa PC.\n3. Naka-queue na sa background para mag-sync sa Google Sheets.\n\n{archive_status}")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() 
    
    processor = simpledialog.askstring("Auditor Authentication", "EOD PROCESSOR LOGIN\n\nPlease enter your name to proceed:", parent=root)
    
    if not processor or processor.strip() == "":
        messagebox.showerror("Access Denied", "EOD processing cannot begin without an authorized processor name.")
        root.destroy()
    else:
        files = [f for f in os.listdir(INBOX_DIR) if f.endswith(('.csv', '.xlsx', '.xls'))]
        
        if not files:
            messagebox.showwarning("No Data Found", "No POS export detected in SALES_REPORT_FROM_POS folder.\n\nPlease export data from POS first.")
            root.destroy()
        else:
            file_path = os.path.join(INBOX_DIR, files[0])
            try:
                df = read_pos_file(file_path)
                root.deiconify() 
                app = EODSettlementApp(root, df, file_path, processor.strip())
                root.mainloop()
            except Exception as e:
                messagebox.showerror("Data Read Error", f"Failed to parse POS data.\n\nError details: {e}")
                root.destroy()