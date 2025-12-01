import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. CONFIGURATION
# ==========================================
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

COMMISSION_REPS = ["Dave Borkowski", "Jake Lynch", "Brad Sherman", "Lance Mitton"]

REP_COMMISSION_RATES = {
    "Dave Borkowski": 0.05,      # 5% flat rate
    "Jake Lynch": 0.07,          # 7% flat rate
    "Brad Sherman": 0.07,        # 7% flat rate
    "Lance Mitton": 0.07,        # 7% flat rate
}

BRAD_OVERRIDE_RATE = 0.01

# ==========================================
# 2. DATA LOADING & PROCESSING
# ==========================================

@st.cache_data(ttl=3600)
def fetch_google_sheet_data(sheet_name, range_name):
    """Fetch data from Google Sheets using Streamlit Secrets"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing 'gcp_service_account' in Streamlit secrets.")
            return pd.DataFrame()

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        
        service = build('sheets', 'v4', credentials=creds)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        if not values: return pd.DataFrame()
        
        headers = values[0]
        data = values[1:]
        return pd.DataFrame(data, columns=headers)

    except Exception as e:
        st.error(f"Error loading {sheet_name}: {str(e)}")
        return pd.DataFrame()

def process_ns_invoices(df):
    """Clean and standardize invoice data"""
    if df is None or df.empty: return pd.DataFrame()
    
    # Avoid modifying the original dataframe from cache/parent
    df = df.copy()
    
    # Basic Cleanup
    df.columns = df.columns.str.strip()
    
    # Handle Date Columns
    if 'Date' in df.columns and 'Date Closed' in df.columns:
        df = df.drop(columns=['Date'])
    
    # Mappings
    rename_map = {
        'Amount (Transaction Total)': 'Amount',
        'Date Closed': 'Date',
        'Document Number': 'Document Number',
        'Status': 'Status',
        'Amount (Transaction Tax Total)': 'Tax Amount',
        'Amount (Shipping)': 'Shipping Amount',
        'Corrected Customer Name': 'Customer',
        'HubSpot Pipeline': 'Pipeline',
        'Rep Master': 'Sales Rep',
        'Sales Rep': 'Original Sales Rep' # Preserve original if needed
    }
    
    # Only rename columns that actually exist
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    
    # If Sales Rep is still missing, check if it was mapped correctly
    if 'Sales Rep' not in df.columns and 'Rep Master' in df.columns:
         df['Sales Rep'] = df['Rep Master']

    # Handle Customer Name Prefix
    if 'Customer' in df.columns:
        df['Customer'] = df['Customer'].astype(str).str.replace('^Customer ', '', regex=True)

    # Numeric Cleanup
    numeric_cols = ['Amount', 'Tax Amount', 'Shipping Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Date Cleanup
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df['Month_Str'] = df['Date'].dt.strftime('%Y-%m')

    # Calculate Net
    if 'Amount' in df.columns:
        df['Subtotal'] = df['Amount']
        if 'Tax Amount' in df.columns: df['Subtotal'] -= df['Tax Amount']
        if 'Shipping Amount' in df.columns: df['Subtotal'] -= df['Shipping Amount']
    
    return df

def calculate_commissions(df, status_filter, month_filter):
    """Core logic to calculate commissions"""
    # 1. Filter Reps
    if 'Sales Rep' not in df.columns:
        st.error("Missing 'Sales Rep' column in data.")
        return pd.DataFrame()
        
    df = df[df['Sales Rep'].isin(COMMISSION_REPS)].copy()
    
    # 2. Filter Status (Dynamic)
    if status_filter:
        # Normalize status for comparison
        df['Status_Clean'] = df['Status'].astype(str).str.upper().str.strip()
        status_clean = [s.upper().strip() for s in status_filter]
        df = df[df['Status_Clean'].isin(status_clean)]
    
    # 3. Filter Month
    if month_filter:
        df = df[df['Month_Str'].isin(month_filter)]

    # 4. Calculate Rates
    df['Commission Rate'] = df['Sales Rep'].map(REP_COMMISSION_RATES).fillna(0.0)
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']
    
    # 5. Brad's Override Calculation
    # Brad gets override on Lance's deals
    df['Brad Override'] = 0.0
    lance_mask = df['Sales Rep'] == 'Lance Mitton'
    df.loc[lance_mask, 'Brad Override'] = df.loc[lance_mask, 'Subtotal'] * BRAD_OVERRIDE_RATE
    
    return df

# ==========================================
# 3. UI HELPERS
# ==========================================

def display_password_gate():
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("""
        <div style='background-color: #262730; padding: 20px; border-radius: 10px; border: 1px solid #464b5c; margin-bottom: 20px;'>
            <h3 style='margin-top:0;'>🔒 Commission Access</h3>
        </div>
        """, unsafe_allow_html=True)
        
        email = st.text_input("Email", key="comm_login_email")
        password = st.text_input("Password", type="password", key="comm_login_pass")
        
        if st.button("Unlock Calculator", use_container_width=True, type="primary"):
            if email == ADMIN_EMAIL and hashlib.sha256(password.encode()).hexdigest() == ADMIN_PASSWORD_HASH:
                st.session_state.commission_authenticated = True
                st.rerun()
            else:
                st.error("Invalid Credentials")

# ==========================================
# 4. MAIN MODULE ENTRY POINT
# ==========================================

def display_commission_section(invoices_df=None, sales_orders_df=None):
    """
    Main function called by sales_dashboard.py
    """
    
    # --- Custom CSS ---
    st.markdown("""
    <style>
        /* Gradient Headers */
        .main-header {
            background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%);
            padding: 20px;
            border-radius: 10px;
            color: white;
            text-align: center;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .main-header h2 { color: white; margin:0; }
        
        /* Metrics Styling */
        div[data-testid="stMetric"] {
            background-color: #f0f2f6;
            border: 1px solid #e0e0e0;
            padding: 15px;
            border-radius: 10px;
        }
    </style>
    """, unsafe_allow_html=True)

    # --- Authentication Gate ---
    if 'commission_authenticated' not in st.session_state:
        st.session_state.commission_authenticated = False

    if not st.session_state.commission_authenticated:
        display_password_gate()
        return

    # --- Data Handling ---
    # If invoices_df is passed from parent, use it, otherwise fetch
    if invoices_df is None or invoices_df.empty:
        with st.spinner("Fetching invoice data..."):
            raw_data = fetch_google_sheet_data("NS Invoices", "A:Z")
    else:
        raw_data = invoices_df

    # Process Data
    clean_data = process_ns_invoices(raw_data)
    
    if clean_data.empty:
        st.error("No data available to process.")
        return

    # --- Header with Logout ---
    c1, c2 = st.columns([6, 1])
    with c1:
        st.markdown("""
        <div class="main-header">
            <h2>💰 Commission Calculator</h2>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        if st.button("Logout", key="comm_logout"):
            st.session_state.commission_authenticated = False
            st.rerun()

    # --- Filters (Sidebar or Top) ---
    # Use columns for filters to avoid clashing with main app sidebar
    st.subheader("⚙️ Filters")
    f1, f2 = st.columns(2)
    
    with f1:
        available_months = sorted(clean_data['Month_Str'].dropna().unique(), reverse=True)
        selected_months = st.multiselect("📅 Close Month", available_months, default=available_months[:1])
    
    with f2:
        all_statuses = clean_data['Status'].unique()
        # Default to Paid In Full if available
        default_status = ["Paid In Full"] if "Paid In Full" in all_statuses else [all_statuses[0]]
        selected_status = st.multiselect("🏷️ Status", all_statuses, default=default_status)

    st.markdown("---")

    # --- Calculation ---
    processed_df = calculate_commissions(clean_data, selected_status, selected_months)

    if processed_df.empty:
        st.warning("⚠️ No transactions found matching these filters.")
        return

    # --- Global KPI Row ---
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    
    total_comm = processed_df['Commission Amount'].sum()
    total_override = processed_df['Brad Override'].sum()
    total_payout = total_comm + total_override
    
    kpi1.metric("Total Payout", f"${total_payout:,.2f}", delta="Includes Overrides")
    kpi2.metric("Total Sales Volume", f"${processed_df['Subtotal'].sum():,.2f}")
    kpi3.metric("Deals Closed", len(processed_df))
    kpi4.metric("Avg Deal Size", f"${processed_df['Subtotal'].mean():,.2f}")

    # --- Visualization ---
    with st.expander("📊 Performance Visuals", expanded=True):
        chart_data = processed_df.groupby('Sales Rep')[['Commission Amount', 'Brad Override']].sum().reset_index()
        
        fig = px.bar(
            chart_data, 
            x='Sales Rep', 
            y=['Commission Amount', 'Brad Override'], 
            title="Commission Breakdown by Rep",
            labels={'value': 'USD ($)', 'variable': 'Type'},
            color_discrete_sequence=['#4b6cb7', '#182848'],
            template="plotly_white"
        )
        st.plotly_chart(fig, use_container_width=True)

    # --- Rep Sections (Tabs) ---
    tabs = st.tabs(["📋 Overview"] + COMMISSION_REPS)

    # 1. OVERVIEW TAB
    with tabs[0]:
        st.subheader("Master Ledger")
        
        column_cfg = {
            "Subtotal": st.column_config.NumberColumn("Net Sales", format="$%d"),
            "Commission Amount": st.column_config.NumberColumn("Comm.", format="$%.2f"),
            "Brad Override": st.column_config.NumberColumn("Override", format="$%.2f"),
            "Date": st.column_config.DateColumn("Close Date", format="YYYY-MM-DD"),
            "Status": st.column_config.TextColumn("Status"),
            "Sales Rep": st.column_config.TextColumn("Rep"),
        }
        
        display_cols = ['Date', 'Document Number', 'Customer', 'Sales Rep', 'Status', 'Subtotal', 'Commission Amount', 'Brad Override']
        st.dataframe(
            processed_df[display_cols].sort_values(by='Date', ascending=False),
            use_container_width=True,
            column_config=column_cfg,
            hide_index=True
        )

    # 2. INDIVIDUAL REP TABS
    for i, rep in enumerate(COMMISSION_REPS):
        with tabs[i+1]:
            rep_df = processed_df[processed_df['Sales Rep'] == rep].copy()
            
            if rep_df.empty:
                st.info(f"No transactions for {rep} in this period.")
                continue

            # Rep Metrics
            rep_sales = rep_df['Subtotal'].sum()
            rep_comm = rep_df['Commission Amount'].sum()
            
            # Special Logic for Brad (Show his override earnings separately)
            is_brad = (rep == "Brad Sherman")
            rep_override = 0
            if is_brad:
                lance_df = processed_df[processed_df['Sales Rep'] == "Lance Mitton"]
                rep_override = lance_df['Brad Override'].sum()

            c1, c2, c3 = st.columns(3)
            c1.metric(f"{rep} Sales", f"${rep_sales:,.2f}")
            c2.metric(f"Direct Commission", f"${rep_comm:,.2f}")
            if is_brad:
                c3.metric("Override Earnings", f"${rep_override:,.2f}", help="1% of Lance's Sales")
            else:
                c3.metric("Deal Count", len(rep_df))

            st.divider()
            
            # Rep Data Table
            st.subheader(f"📄 {rep}'s Deal Sheet")
            
            rep_display_cols = ['Date', 'Document Number', 'Customer', 'Pipeline', 'Status', 'Subtotal', 'Commission Amount']
            
            st.dataframe(
                rep_df[rep_display_cols].sort_values('Date', ascending=False),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Subtotal": st.column_config.ProgressColumn(
                        "Deal Value",
                        format="$%f",
                        min_value=0,
                        max_value=max(processed_df['Subtotal'].max(), 1000)
                    ),
                    "Commission Amount": st.column_config.NumberColumn("Commission", format="$%.2f"),
                    "Date": st.column_config.DateColumn("Date", format="MM/DD/YYYY"),
                }
            )
            
            # If Brad, show the Override Source Table
            if is_brad and rep_override > 0:
                st.divider()
                st.subheader("🕵️ Override Source (Lance's Deals)")
                lance_source = processed_df[processed_df['Sales Rep'] == "Lance Mitton"][['Date', 'Customer', 'Subtotal', 'Brad Override']]
                st.dataframe(
                    lance_source, 
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Subtotal": st.column_config.NumberColumn("Lance Sales", format="$%d"),
                        "Brad Override": st.column_config.NumberColumn("Brad Cut (1%)", format="$%.2f")
                    }
                )

# Standalone execution for testing
if __name__ == "__main__":
    st.set_page_config(layout="wide")
    display_commission_section()
