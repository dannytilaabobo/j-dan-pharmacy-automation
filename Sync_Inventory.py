import pandas as pd
import os

# CONFIG (Siguraduhin na tugma ang folder path)
master_file = 'Inventory_Master_20.xlsx'
sales_report = 'JDan_Sales_System/EOD_Deduction_Reports/SALES_JUNE_06_2026.xlsx' 

def sync_inventory():
    df_master = pd.read_excel(master_file, sheet_name='Stocks')
    df_sales = pd.read_excel(sales_report)
    
    # I-sync ang stock
    for index, row in df_sales.iterrows():
        code = str(row['Item_Code']).strip()
        sold = float(row['Qty_Deducted'])
        
        # Bawas sa Master
        mask = df_master['Item_Code'].astype(str).str.strip() == code
        df_master.loc[mask, 'Current_Stock'] -= sold
        
    # Save as new "Expected" Inventory
    df_master.to_excel('Inventory_Master_20_Expected_June07.xlsx', index=False)
    print("Sync Done! Check 'Inventory_Master_20_Expected_June07.xlsx'")

sync_inventory()