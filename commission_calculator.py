"""
Commission Calculator Module for Calyx Containers
Updated: Google Sheets Integration ('NS Invoices' Logic)
Focus: Brad Sherman (Oct 2025) - Paid In Full / Date Closed Logic
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# ==========================================
# ⚙️ COLUMN MAPPING CONFIGURATION
# Updated based on user feedback:
# Col O = Sales Rep (for Shopify exclusion)
# Col R = Amount Tax (for Math)
# Col U = Rep Master (for Brad filtering)
# ==========================================
COLS = {
    "STATUS": "Status",                     # Col B
    "TRANS_TOTAL": "Amount Transaction Total", # Col K
    "DATE_CLOSED": "Date Closed",           # Col N
    "SALES_REP": "Sales Rep",               # Col O (Filter OUT Shopify)
    "TAX": "Amount Tax",                    # Col R (Math)
    "REP_MASTER": "Rep Master",             # Col U (Filter FOR Rep)
    "SHIPPING": "Amount Shipping"           # Needed for Net Revenue calc
}

# ==========================================
# COMMISSION CONFIGURATION
# ==========================================

TARGET_REP = "Brad Sherman"
TARGET_RATE = 0.07 
TARGET_MONTH = 10 # October
TARGET_YEAR = 2025

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def load_google_sheet(sheet_url):
    """
    Loads data from a Google Sheet CSV export link.
    """
    try:
        # Auto-convert edit links to export links if user mistakes them
        if '/edit' in sheet_url:
            sheet_url = sheet_url.replace('/edit#gid=', '/export?format=csv&gid=')
            sheet_url = sheet_url.replace('/edit?gid=', '/export?format=csv&gid=')
        
        df = pd.read_csv(sheet_url)
        return df
    except Exception as e:
        st.error(f"Error loading Google Sheet: {e}")
        return None

def process_brad_logic(df):
    """
    Logic:
    1. Filter Rep Master (Col U) == Brad Sherman
    2. Filter Status (Col B) == Paid In Full
    3. Filter Date Closed (Col N) == October 2025
    4. Exclude Sales Rep (Col O) == Shopify Ecommerce
    5. Math: Total (Col K) - Shipping - Tax (Col R)
    """
    if df.empty:
        return pd.DataFrame()

    # Clean Column Names
    df.columns = df.columns.str.strip()

    # Check for required columns
    missing_cols = [c for c in COLS.values() if c not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing columns: {missing_cols}")
        st.write("Columns found in file:", df.columns.tolist())
        return pd.DataFrame()

    # ==========================================
    # 🔍 FILTERING
    # ==========================================

    # 1. Filter by Rep Master (Col U) -> Brad Sherman
    mask_rep = df[COLS['REP_MASTER']].astype(str).str.strip().str.upper() == TARGET_REP.upper()
    df = df[mask_rep].copy()
    
    # 2. Filter by Status (Col B) -> Paid In Full
    mask_status = df[COLS['STATUS']].astype(str).str.strip().str.upper() == "PAID IN FULL"
    df = df[mask_status].copy()

    # 3. Filter by Date Closed (Col N) -> Oct 2025
    df['Date_Closed_DT'] = pd.to_datetime(df[COLS['DATE_CLOSED']], errors='coerce')
    mask_date = (df['Date_Closed_DT'].dt.month == TARGET_MONTH) & \
                (df['Date_Closed_DT'].dt.year == TARGET_YEAR)
    df = df[mask_date].copy()

    # 4. Filter OUT Shopify (Col O)
    # We look at 'Sales Rep' column (O) and exclude if it contains 'Shopify'
    df[COLS['SALES_REP']] = df[COLS['SALES_REP']].fillna('')
    mask_shopify = df[COLS['SALES_REP']].astype(str).str.upper().str.contains("SHOPIFY", na=False)
    df = df[~mask_shopify].copy()

    if df.empty:
        return df

    # ==========================================
    # 🧮 CALCULATION
    # ==========================================
    
    def clean_currency(series):
        # Remove '$' and ',' and convert to float
        return pd.to_numeric(series.astype(str).str.replace('$','').str.replace(',',''), errors='coerce').fillna(0)

    # Use Col K (Total), Col R (Tax), and Shipping
    df['Calc_Total'] = clean_currency(df[COLS['TRANS_TOTAL']])
    df['Calc_Shipping'] = clean_currency(df[COLS['SHIPPING']])
    df['Calc_Tax'] = clean_currency(df[COLS['TAX']])

    # Net Revenue = Total - Shipping - Tax
    df['Commissionable_Revenue'] = df['Calc_Total'] - df['Calc_Shipping'] - df['Calc_Tax']
    
    # Commission = Net Revenue * 7%
    df['Commission_Amount'] = df['Commissionable_Revenue'] * TARGET_RATE

    return df

# ==========================================
# UI
# ==========================================

def main():
    st.set_page_config(page_title="Brad Commission Calc", layout="wide")
    
    st.title(f"💰 Commission: {TARGET_REP} (Oct 2025)")
    st.markdown("""
    **Rules:**
    1. **Rep Master (Col U):** Must be 'Brad Sherman'
    2. **Status (Col B):** Must be 'Paid In Full'
    3. **Date Closed (Col N):** Must be October
    4. **Sales Rep (Col O):** Cannot be 'Shopify Ecommerce'
    5. **Math:** Trans Total (K) - Shipping - Tax (R)
    """)

    sheet_url = st.text_input("Paste Google Sheet CSV URL (from 'NS Invoices' tab):")

    if sheet_url:
        df = load_google_sheet(sheet_url)

        if df is not None:
            results = process_brad_logic(df)

            if not results.empty:
                # Top Level Metrics
                col1, col2, col3 = st.columns(3)
                total_rev = results['Commissionable_Revenue'].sum()
                total_comm = results['Commission_Amount'].sum()
                
                col1.metric("Commissionable Revenue", f"${total_rev:,.2f}")
                col2.metric("Total Commission (7%)", f"${total_comm:,.2f}")
                col3.metric("Deal Count", len(results))

                st.divider()

                # Detailed Table
                st.subheader("Transaction Details")
                
                # Select columns to display
                display_cols = [
                    COLS['DATE_CLOSED'], 
                    COLS['SALES_REP'], # Show Col O to prove Shopify is gone
                    COLS['TRANS_TOTAL'], 
                    COLS['SHIPPING'], 
                    COLS['TAX'], 
                    'Commissionable_Revenue', 
                    'Commission_Amount'
                ]
                
                # Format for readability
                display_df = results[display_cols].copy()
                money_cols = ['Commissionable_Revenue', 'Commission_Amount', COLS['TRANS_TOTAL'], COLS['TAX']]
                
                for c in money_cols:
                    if c in display_df.columns:
                        display_df[c] = display_df[c].apply(lambda x: f"${x:,.2f}" if isinstance(x, (int, float)) else x)

                st.dataframe(display_df, use_container_width=True)
            else:
                st.warning("No transactions matched criteria (Brad Sherman / Paid in Full / Oct 2025 / Not Shopify).")

if __name__ == "__main__":
    main()
