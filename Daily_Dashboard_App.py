import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# Configurations Matching the Master Blueprint
DB_FILE = 'Inventory_Master.xlsx'
LOG_FILE = 'JDan_System_Logs.txt'

st.set_page_config(page_title="J-DAN Pharmacy Analytics", layout="wide", page_icon="🏥")

def load_data():
    if not os.path.exists(DB_FILE):
        st.error(f"🚨 Master Database '{DB_FILE}' not found. Please run your backup or sync engine.")
        return None, None
    
    try:
        xls = pd.ExcelFile(DB_FILE)
        ledger = pd.read_excel(xls, sheet_name='Transactions')
        summary = pd.read_excel(xls, sheet_name='Inventory_Summary')
        return ledger, summary
    except Exception as e:
        st.error(f"❌ Error loading Excel data: {e}")
        return None, None

# --- HEADER SECTION ---
st.title("🏥 J-DAN Pharmacy Operations & Analytics Dashboard")
st.markdown("Real-time executive monitoring system for transaction ledger, automated summary, and risk control anomalies.")
st.hr()

ledger_df, summary_df = load_data()

if ledger_df is not None and summary_df is not None:
    # --- KEY PERFORMANCE METRICS (KPIs) ---
    total_products = len(summary_df)
    total_stock_units = int(summary_df['Current_Stock'].sum())
    
    # Calculate daily sales transactions from ledger
    today_str = datetime.now().strftime('%m/%d/%Y')
    sales_today = ledger_df[(ledger_df['Transaction_Type'] == 'SALES')]
    total_sales_count = len(sales_today)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="Total Tracked SKUs", value=f"{total_products} Items")
    with col2:
        st.metric(label="Total Physical Stock Volume", value=f"{total_stock_units} Units")
    with col3:
        st.metric(label="Historical Sales Records Count", value=f"{total_sales_count} Batches")

    st.write(" ")

    # --- MAIN CONTENT TABS ---
    tab1, tab2, tab3 = st.tabs(["📊 Inventory Analysis", "📑 Live Transaction Ledger", "🚨 System Logs & Audit Trail"])

    with tab1:
        st.subheader("Current Master Inventory Stock Levels")
        
        # Highlighting items requiring stock replenishment
        low_stock_threshold = 10
        low_stock_df = summary_df[summary_df['Current_Stock'] <= low_stock_threshold]
        
        if not low_stock_df.empty:
            st.warning(f"⚠️ **Alert:** Found {len(low_stock_df)} items that are low in stock (below threshold of {low_stock_threshold} units).")
            st.dataframe(low_stock_df, use_container_width=True)
        else:
            st.success("✅ All stock levels are optimal.")

        # Interactive Data Visualizations
        st.markdown("### Product Stock Breakdown")
        fig = px.bar(summary_df, x='Product_Name', y='Current_Stock', 
                     title="Current Stock Counts Per SKU",
                     labels={'Current_Stock': 'Stock Volume', 'Product_Name': 'Item Description'},
                     color='Current_Stock', color_continuous_scale=px.colors.sequential.Viridis)
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("Immutable Transaction Ledger Log")
        st.markdown("This view reflects the exact row-by-row immutable inputs from the `Transactions` worksheet data pipeline.")
        
        # Filter operations by movement type
        txn_filter = st.multiselect("Filter by Transaction Event:", 
                                    options=ledger_df['Transaction_Type'].unique(), 
                                    default=ledger_df['Transaction_Type'].unique())
        
        filtered_ledger = ledger_df[ledger_df['Transaction_Type'].isin(txn_filter)]
        st.dataframe(filtered_ledger.sort_values(by='Date', ascending=False), use_container_width=True)

    with tab3:
        st.subheader("System Audit Logs (`JDan_System_Logs.txt`)")
        if os.path.exists(LOG_FILE):
            with open(LOG_FILE, "r") as f:
                log_lines = f.readlines()
            
            # Show the latest logs first
            log_text = "".join(log_lines[-20:])  # Read last 20 log actions
            st.code(log_text, language="text")
            st.caption("Displaying the 20 most recent system execution signals.")
        else:
            st.info("ℹ️ No active system logs file found in the current root directory.")
