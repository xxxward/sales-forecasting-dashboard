"""
Commission Calculator Module for Calyx Containers
Integrated with Google Sheets (NS Invoices & NS Sales Orders)
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# GOOGLE SHEETS CONFIGURATION
# ==========================================
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# ==========================================
# PASSWORD CONFIGURATION (Xander Only)
# ==========================================
ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

# ==========================================
# COMMISSION CONFIGURATION
# ==========================================
COMMISSION_REPS = ["Dave Borkowski", "Jake Lynch", "Brad Sherman", "Lance Mitton"]

# Commission rates by rep
REP_COMMISSION_RATES = {
    "Dave Borkowski": 0.05,      # 5% flat rate
    "Jake Lynch": 0.07,          # 7% flat rate
    "Brad Sherman": 0.07,        # 7% flat rate
    "Lance Mitton": 0.07,        # 7% flat rate
}

BRAD_OVERRIDE_RATE = 0.01
COMMISSION_MONTHS = ['2025-09', '2025-10'] 

# ==========================================
# EXCLUSION RULES
# ==========================================
EXCLUDED_ITEMS = [
    "Convenience Fee 3.5%", "SHIPPING", "UPS", "FEDEX", "AVATAX", "TAX"
]

# ==========================================
# DATA LOADING FUNCTIONS (Google Sheets)
# ==========================================

@st.cache_data(ttl=3600)
def fetch_google_sheet_data(sheet_name, range_name):
    """
    Fetch data from Google Sheets using Streamlit Secrets
    """
    try:
        # Check for secrets
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing 'gcp_service_account' in Streamlit secrets.")
            return pd.DataFrame()

        # Create credentials
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        
        # Build service
        service = build('sheets', 'v4', credentials=creds)
        
        # Fetch data
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame()
        
        # Handle headers and data
        headers = values[0]
        data = values[1:]
        
        # Create DataFrame
        df = pd.DataFrame(data, columns=headers)
        return df

    except Exception as e:
        st.error(f"Error loading {sheet_name}: {str(e)}")
        return pd.DataFrame()

def process_ns_invoices(df):
    """
    Clean NS Invoices data and map to Calculator standard columns
    Fixed to handle duplicate 'Date' column collision
    """
    if df.empty:
        return df

    # 1. Clean Column Names
    df.columns = df.columns.str.strip()
    
    # 2. Handle Column Name Collisions BEFORE Renaming
    # We want 'Date Closed' to become 'Date'.
    # But 'Date' likely already exists (Col C). We must drop the old 'Date' first.
    if 'Date' in df.columns and 'Date Closed' in df.columns:
        df = df.drop(columns=['Date'])
        
    # We want 'Rep Master' to become 'Sales Rep'.
    # But 'Sales Rep' likely already exists (Col O). 
    # We want to keep Col O as 'Original Sales Rep' for Shopify filtering.
    if 'Sales Rep' in df.columns:
        df = df.rename(columns={'Sales Rep': 'Original Sales Rep'})
    
    # 3. Map Columns (Google Sheet Header -> Calculator Internal Name)
    rename_map = {
        'Rep Master': 'Sales Rep',
        'Amount Transaction Total': 'Amount',
        'Date Closed': 'Date',
        'Item': 'Item',
        'Document Number': 'Document Number',
        'Status': 'Status',
        'Amount Tax': 'Tax Amount',
        'Amount Shipping': 'Shipping Amount'
    }
    
    # Only rename columns that exist
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # 4. Clean Numeric Columns
    numeric_cols = ['Amount', 'Tax Amount', 'Shipping Amount']
    for col in numeric_cols:
        if col in df.columns:
            # Remove $, commas, whitespace
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 5. Calculate Net Commissionable Amount (Total - Tax - Shipping)
    if 'Amount' in df.columns:
        df['Subtotal'] = df['Amount']
        if 'Tax Amount' in df.columns:
            df['Subtotal'] = df['Subtotal'] - df['Tax Amount']
        if 'Shipping Amount' in df.columns:
            df['Subtotal'] = df['Subtotal'] - df['Shipping Amount']

    return df

def process_ns_sales_orders(df):
    """
    Clean NS Sales Orders data
    """
    if df.empty:
        return df
        
    df.columns = df.columns.str.strip()
    
    # Map Rep Master to Sales Rep
    if 'Rep Master' in df.columns:
        df['Sales Rep'] = df['Rep Master']
        
    # Clean Amount
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].astype(str).str.replace(r'[$,]', '', regex=True)
        df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
        
    return df

# ==========================================
# AUTH & HELPERS
# ==========================================

def verify_admin(email, password):
    if email != ADMIN_EMAIL:
        return False
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    return password_hash == ADMIN_PASSWORD_HASH

def parse_invoice_month(date_value):
    if pd.isna(date_value) or str(date_value).strip() == '': 
        return None
    try:
        date_obj = pd.to_datetime(date_value, errors='coerce')
        if pd.isna(date_obj): return None
        return date_obj.strftime('%Y-%m')
    except: return None

def calculate_payout_date(invoice_month_str):
    try:
        if not invoice_month_str: return None
        year, month = map(int, invoice_month_str.split('-'))
        if month == 12:
            month_end = datetime(year, 12, 31)
        else:
            month_end = datetime(year, month + 1, 1) - timedelta(days=1)
        return month_end + timedelta(weeks=4)
    except: return None

# ==========================================
# MAIN CALCULATION ENGINE
# ==========================================

def process_commission_data_fast(df):
    """
    Main logic to filter rows and calculate commission
    """
    if df.empty: return pd.DataFrame()

    # 1. Filter to 4 Reps Only
    if 'Sales Rep' not in df.columns:
        st.error("❌ 'Sales Rep' column missing after processing. Check 'Rep Master' column in Sheet.")
        return pd.DataFrame()
        
    df = df[df['Sales Rep'].isin(COMMISSION_REPS)].copy()
    
    # 2. Filter: Status = Paid In Full
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.upper().str.strip() == "PAID IN FULL"]

    # 3. Filter: Exclude Shopify (Column O in sheet, mapped to Original Sales Rep)
    if 'Original Sales Rep' in df.columns:
        df = df[~df['Original Sales Rep'].astype(str).str.upper().str.contains("SHOPIFY", na=False)]

    # 4. Date Parsing
    if 'Date' not in df.columns:
        st.error("❌ 'Date' column missing. Check 'Date Closed' column in Sheet.")
        return pd.DataFrame()

    st.caption("📅 Using 'Date Closed' to determine commission month")
    df['Invoice Month'] = df['Date'].apply(parse_invoice_month)
    df = df[df['Invoice Month'].isin(COMMISSION_MONTHS)].copy()

    # 5. Exclude non-commissionable items (Shipping, Tax labels, etc)
    if 'Item' in df.columns:
        df['Item Upper'] = df['Item'].astype(str).str.upper()
        def is_excluded(item_upper):
            if 'TOOLING' in item_upper: return False
            return any(excl in item_upper for excl in EXCLUDED_ITEMS)
        
        df = df[~df['Item Upper'].apply(is_excluded)]

    # 6. Apply Commission Rates
    # Assign Base Rate
    df['Commission Rate'] = df['Sales Rep'].map(REP_COMMISSION_RATES).fillna(0.0)
    
    # Calculate Commission (Subtotal * Rate)
    if 'Subtotal' not in df.columns:
        df['Subtotal'] = df['Amount'] # Fallback
        
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']

    # 7. Brad's Override (1% on Lance)
    df['Brad Override'] = 0.0
    lance_mask = df['Sales Rep'] == 'Lance Mitton'
    df.loc[lance_mask, 'Brad Override'] = df.loc[lance_mask, 'Subtotal'] * BRAD_OVERRIDE_RATE

    # 8. Payout Date
    df['Payout Date'] = df['Invoice Month'].apply(calculate_payout_date)

    return df

# ==========================================
# UI COMPONENTS
# ==========================================

def display_password_gate():
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 padding: 20px; border-radius: 10px; color: white; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0;'>🔒 Commission Calculator</h2>
        <p style='margin: 5px 0 0 0;'>Admin access required</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        email = st.text_input("Email:", key="commission_email")
        password = st.text_input("Password:", type="password", key="commission_password")
        if st.button("🔓 Login", use_container_width=True):
            if verify_admin(email, password):
                st.session_state.commission_authenticated = True
                st.rerun()
            else:
                st.error("❌ Invalid credentials")

def display_commission_dashboard(invoice_df):
    """Display the commission dashboard"""
    # Header
    col1, col2 = st.columns([4, 1])
    with col1:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                      padding: 15px; border-radius: 10px; color: white;'>
            <h2 style='color: white; margin: 0;'>💰 Commission Dashboard</h2>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.commission_authenticated = False
            st.rerun()

    # Process Data
    with st.spinner("Calculating commissions from live Sheet data..."):
        commission_df = process_commission_data_fast(invoice_df)

    if commission_df.empty:
        st.warning("⚠️ No commissionable transactions found for the selected months (Sep/Oct 2025).")
        return

    # Overall Metrics
    st.markdown("### 📊 Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sales", f"${commission_df['Subtotal'].sum():,.2f}")
    c2.metric("Total Commission", f"${commission_df['Commission Amount'].sum():,.2f}")
    c3.metric("Brad Override", f"${commission_df['Brad Override'].sum():,.2f}")
    c4.metric("Line Items", len(commission_df))

    st.markdown("---")

    # Group by Rep
    st.markdown("### 👤 By Sales Rep")
    rep_summary = commission_df.groupby('Sales Rep').agg({
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    # Add Total Override Row for Brad
    total_override = commission_df['Brad Override'].sum()
    if total_override > 0:
        brad_row = pd.DataFrame([{
            'Sales Rep': 'Brad Sherman (Override)', 
            'Subtotal': 0, 
            'Commission Amount': 0, 
            'Brad Override': total_override
        }])
        rep_summary = pd.concat([rep_summary, brad_row], ignore_index=True)

    # Format for display
    display_rep = rep_summary.copy()
    display_rep['Subtotal'] = display_rep['Subtotal'].apply(lambda x: f"${x:,.2f}")
    display_rep['Commission Amount'] = display_rep['Commission Amount'].apply(lambda x: f"${x:,.2f}")
    display_rep['Brad Override'] = display_rep['Brad Override'].apply(lambda x: f"${x:,.2f}")

    st.dataframe(display_rep, use_container_width=True, hide_index=True)

    st.markdown("---")
    
    # Detailed Data
    with st.expander("📋 View Detailed Transactions"):
        st.dataframe(commission_df)

# ==========================================
# MAIN ENTRY POINT
# ==========================================

def display_commission_section(invoices_df=None, sales_orders_df=None):
    """
    Main function called by the dashboard. 
    Ignores passed dfs if they aren't processed correctly, loads fresh from Sheets.
    """
    
    # Check Auth
    if not st.session_state.get('commission_authenticated', False):
        display_password_gate()
        return

    # Load Data Directly from Google Sheets (Fresh Pull)
    with st.spinner("🔄 Fetching live data from 'NS Invoices' and 'NS Sales Orders'..."):
        raw_invoices = fetch_google_sheet_data("NS Invoices", "A:U")
        raw_orders = fetch_google_sheet_data("NS Sales Orders", "A:AF")
        
        if raw_invoices.empty:
            st.error("Could not load 'NS Invoices' from Google Sheet.")
            return

        # Clean and Prepare
        clean_invoices = process_ns_invoices(raw_invoices)
        clean_orders = process_ns_sales_orders(raw_orders)
        
        st.success(f"✅ Loaded {len(clean_invoices)} Invoices and {len(clean_orders)} Sales Orders")

    # Display Dashboard
    display_commission_dashboard(clean_invoices)

if __name__ == "__main__":
    st.set_page_config(page_title="Commission Calc", layout="wide")
    display_commission_section()
