import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# 1. APP CONFIG & STYLING
# ==========================================
st.set_page_config(
    page_title="Calyx Commissions",
    page_icon="💸",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for that "Sexy" Look
st.markdown("""
<style>
    /* Gradient Headers */
    .stApp header {background-color: transparent;}
    .main-header {
        background: linear-gradient(90deg, #4b6cb7 0%, #182848 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 { color: white; margin:0; font-family: 'Helvetica Neue', sans-serif; font-weight: 700; }
    .main-header p { color: #e0e0e0; margin:0; font-size: 1.1rem; }
    
    /* Metrics Styling */
    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"]:hover {
        border-color: #4b6cb7;
        transform: translateY(-2px);
        transition: all 0.3s ease;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CONFIGURATION
# ==========================================
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

COMMISSION_REPS = ["Dave Borkowski", "Jake Lynch", "Brad Sherman", "Lance Mitton"]

REP_COMMISSION_RATES = {
    "Dave Borkowski": 0.05,
    "Jake Lynch": 0.07,
    "Brad Sherman": 0.07,
    "Lance Mitton": 0.07,
}

BRAD_OVERRIDE_RATE = 0.01

# ==========================================
# 3. DATA LOADING
# ==========================================

@st.cache_data(ttl=3600)
def fetch_google_sheet_data(sheet_name, range_name):
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
    if df.empty: return df
    
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
        'Rep Master': 'Sales Rep'
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    
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

# ==========================================
# 4. COMMISSION ENGINE
# ==========================================

def calculate_commissions(df, status_filter, month_filter):
    # 1. Filter Reps
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
# 5. UI COMPONENTS
# ==========================================

def login_screen():
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("""
        <div style='background-color: #1e1e1e; padding: 30px; border-radius: 15px; text-align: center; border: 1px solid #333;'>
            <h2>🔒 Restricted Access</h2>
            <p>Please log in to view commissions.</p>
        </div>
        """, unsafe_allow_html=True)
        
        email = st.text_input("Email", key="login_email")
        password = st.text_input("Password", type="password", key="login_pass")
        
        if st.button("Unlock Dashboard", use_container_width=True, type="primary"):
            if email == ADMIN_EMAIL and hashlib.sha256(password.encode()).hexdigest() == ADMIN_PASSWORD_HASH:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Credentials")

def main_dashboard(raw_df):
    # --- Sidebar Filters ---
    with st.sidebar:
        st.title("⚙️ Filters")
        
        # Date Filter
        available_months = sorted(raw_df['Month_Str'].dropna().unique(), reverse=True)
        selected_months = st.multiselect("📅 Close Month", available_months, default=available_months[:1])
        
        # Status Filter
        all_statuses = raw_df['Status'].unique()
        # Default to Paid In Full if available, otherwise first item
        default_status = ["Paid In Full"] if "Paid In Full" in all_statuses else [all_statuses[0]]
        selected_status = st.multiselect("🏷️ Status", all_statuses, default=default_status)
        
        st.divider()
        st.caption(f"Admin: {ADMIN_EMAIL}")
        if st.button("Logout"):
            st.session_state.authenticated = False
            st.rerun()

    # --- Processing ---
    processed_df = calculate_commissions(raw_df, selected_status, selected_months)

    # --- Main Header ---
    st.markdown(f"""
    <div class="main-header">
        <h1>💰 Commission Calculator</h1>
        <p>Period: {", ".join(selected_months) if selected_months else "All Time"}</p>
    </div>
    """, unsafe_allow_html=True)

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
        # Prepare data for plotting
        chart_data = processed_df.groupby('Sales Rep')[['Commission Amount', 'Brad Override']].sum().reset_index()
        chart_data['Total Earnings'] = chart_data['Commission Amount'] + chart_data['Brad Override']
        
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
        
        # Configure the big table
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
                # Brad gets override from ALL Lance rows in the main dataframe, not just Brad's rows
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

# ==========================================
# 6. APP EXECUTION
# ==========================================

if __name__ == "__main__":
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        login_screen()
    else:
        # Load Data once authenticated
        with st.spinner("🔄 Fetching live data from Netsuite Invoices..."):
            raw_data = fetch_google_sheet_data("NS Invoices", "A:Z")
            
            if not raw_data.empty:
                clean_data = process_ns_invoices(raw_data)
                main_dashboard(clean_data)
            else:
                st.error("Unable to load data. Please check Google Sheets connection.")
