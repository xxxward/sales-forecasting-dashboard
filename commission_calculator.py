"""
Commission Calculator Module for Calyx Containers
Enhanced version with interactive SO selection
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

# Available months for commission
AVAILABLE_MONTHS = ['2025-09', '2025-10', '2025-11', '2025-12']

# ==========================================
# DATA LOADING FUNCTIONS (Google Sheets)
# ==========================================

@st.cache_data(ttl=3600)
def fetch_google_sheet_data(sheet_name, range_name):
    """
    Fetch data from Google Sheets using Streamlit Secrets
    """
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
        
        if not values:
            return pd.DataFrame()
        
        headers = values[0]
        data = values[1:]
        
        df = pd.DataFrame(data, columns=headers)
        return df

    except Exception as e:
        st.error(f"Error loading {sheet_name}: {str(e)}")
        return pd.DataFrame()

def process_ns_invoices(df):
    """
    Clean NS Invoices data and map to Calculator standard columns
    """
    if df.empty:
        return df
    
    # Clean Column Names
    df.columns = df.columns.str.strip()
    
    # Handle Column Name Collisions
    # There's a 'Date' column (Col C) and 'Date Closed' column (Col N)
    if 'Date' in df.columns and 'Date Closed' in df.columns:
        df = df.drop(columns=['Date'])
    
    # There's a 'Sales Rep' column (Col O) - preserve it as 'Original Sales Rep' for Shopify filter
    if 'Sales Rep' in df.columns:
        df = df.rename(columns={'Sales Rep': 'Original Sales Rep'})
    
    # Map Columns - preserve all display columns
    rename_map = {
        'Amount (Transaction Total)': 'Amount',
        'Date Closed': 'Date',
        'Document Number': 'Document Number',
        'Status': 'Status',
        'Amount (Transaction Tax Total)': 'Tax Amount',
        'Amount (Shipping)': 'Shipping Amount',
        'Corrected Customer Name': 'Customer',
        'Created From': 'Created From',
        'CSM': 'CSM',
        'HubSpot Pipeline': 'HubSpot Pipeline',
        'Rep Master': 'Rep Master'
    }
    
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    
    # Add Sales Rep as a separate field for filtering (from Rep Master)
    if 'Rep Master' in df.columns:
        df['Sales Rep'] = df['Rep Master']
    
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Clean Numeric Columns
    numeric_cols = ['Amount', 'Tax Amount', 'Shipping Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Calculate Net Commissionable Amount
    if 'Amount' in df.columns:
        df['Subtotal'] = df['Amount']
        if 'Tax Amount' in df.columns:
            df['Subtotal'] = df['Subtotal'] - df['Tax Amount']
        if 'Shipping Amount' in df.columns:
            df['Subtotal'] = df['Subtotal'] - df['Shipping Amount']
    else:
        st.error("❌ Critical Error: 'Amount' column not found.")
        return pd.DataFrame()

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

def process_commission_data(df, selected_months):
    """
    Main logic to filter rows and calculate commission
    """
    if df.empty: 
        return pd.DataFrame()

    # Filter to 4 Reps Only
    if 'Sales Rep' not in df.columns:
        st.error("❌ 'Sales Rep' column missing.")
        return pd.DataFrame()
        
    df = df[df['Sales Rep'].isin(COMMISSION_REPS)].copy()
    
    # Filter: Status = Paid In Full
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.upper().str.strip() == "PAID IN FULL"]

    # Filter: Exclude Shopify
    if 'Original Sales Rep' in df.columns:
        df = df[~df['Original Sales Rep'].astype(str).str.upper().str.contains("SHOPIFY", na=False)]

    # Date Parsing
    if 'Date' not in df.columns:
        st.error("❌ 'Date' column missing.")
        return pd.DataFrame()

    df['Invoice Month'] = df['Date'].apply(parse_invoice_month)
    df = df[df['Invoice Month'].isin(selected_months)].copy()

    # Apply Commission Rates
    df['Commission Rate'] = df['Sales Rep'].map(REP_COMMISSION_RATES).fillna(0.0)
    
    if 'Subtotal' not in df.columns:
        st.error("❌ Subtotal missing!")
        return pd.DataFrame()
        
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']

    # Brad's Override (1% on Lance)
    df['Brad Override'] = 0.0
    lance_mask = df['Sales Rep'] == 'Lance Mitton'
    df.loc[lance_mask, 'Brad Override'] = df.loc[lance_mask, 'Subtotal'] * BRAD_OVERRIDE_RATE

    # Payout Date
    df['Payout Date'] = df['Invoice Month'].apply(calculate_payout_date)
    
    # Add unique ID for selection
    df['Row_ID'] = range(len(df))

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
    """Display the enhanced commission dashboard"""
    
    # Header with Logout
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

    st.markdown("---")
    
    # Month Selector
    st.markdown("### 📅 Select Commission Period")
    selected_months = st.multiselect(
        "Choose months to include:",
        options=AVAILABLE_MONTHS,
        default=['2025-10'],
        help="Select one or more months to calculate commissions"
    )
    
    if not selected_months:
        st.warning("⚠️ Please select at least one month.")
        return

    # Process Data
    with st.spinner("Loading commission data..."):
        commission_df = process_commission_data(invoice_df, selected_months)

    if commission_df.empty:
        st.warning(f"⚠️ No commissionable transactions found for selected months.")
        return

    # Initialize session state for selections
    if 'selected_rows' not in st.session_state:
        st.session_state.selected_rows = set(commission_df['Row_ID'].tolist())

    st.markdown("---")
    
    # Rep Selector
    st.markdown("### 👤 Sales Rep Filter")
    selected_reps = st.multiselect(
        "Choose reps to view:",
        options=COMMISSION_REPS,
        default=COMMISSION_REPS,
        help="Filter by sales representative"
    )
    
    if not selected_reps:
        st.warning("⚠️ Please select at least one rep.")
        return
    
    # Filter by selected reps
    filtered_df = commission_df[commission_df['Sales Rep'].isin(selected_reps)].copy()
    
    # Apply row selections
    included_df = filtered_df[filtered_df['Row_ID'].isin(st.session_state.selected_rows)].copy()
    
    st.markdown("---")

    # Overall Summary
    st.markdown("### 📊 Commission Summary (Selected Transactions)")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Sales", f"${included_df['Subtotal'].sum():,.2f}")
    c2.metric("Total Commission", f"${included_df['Commission Amount'].sum():,.2f}")
    c3.metric("Brad Override", f"${included_df['Brad Override'].sum():,.2f}")
    c4.metric("Transactions", f"{len(included_df)} / {len(filtered_df)}")

    st.markdown("---")

    # By Rep Summary
    st.markdown("### 👥 By Sales Rep")
    
    rep_summary = included_df.groupby('Sales Rep').agg({
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum',
        'Document Number': 'count'
    }).reset_index()
    
    rep_summary.columns = ['Sales Rep', 'Total Sales', 'Commission', 'Brad Override', 'Transaction Count']
    
    # Add Brad's override as a separate row
    total_override = included_df['Brad Override'].sum()
    if total_override > 0:
        brad_row = pd.DataFrame([{
            'Sales Rep': 'Brad Sherman (Override)', 
            'Total Sales': 0, 
            'Commission': 0, 
            'Brad Override': total_override,
            'Transaction Count': 0
        }])
        rep_summary = pd.concat([rep_summary, brad_row], ignore_index=True)

    # Format for display
    display_rep = rep_summary.copy()
    display_rep['Total Sales'] = display_rep['Total Sales'].apply(lambda x: f"${x:,.2f}")
    display_rep['Commission'] = display_rep['Commission'].apply(lambda x: f"${x:,.2f}")
    display_rep['Brad Override'] = display_rep['Brad Override'].apply(lambda x: f"${x:,.2f}")

    st.dataframe(display_rep, use_container_width=True, hide_index=True)

    st.markdown("---")
    
    # Detailed Transactions by Rep
    st.markdown("### 📋 Transaction Details - Select to Include/Exclude")
    
    for rep in selected_reps:
        rep_data = filtered_df[filtered_df['Sales Rep'] == rep].copy()
        
        if rep_data.empty:
            continue
        
        # Calculate rep totals
        rep_total_sales = rep_data['Subtotal'].sum()
        rep_included = included_df[included_df['Sales Rep'] == rep]
        rep_included_total = rep_included['Subtotal'].sum() if not rep_included.empty else 0
        rep_included_commission = rep_included['Commission Amount'].sum() if not rep_included.empty else 0
        
        # Header checkbox and expander
        col_check, col_expand = st.columns([0.5, 9.5])
        
        with col_check:
            all_rep_selected = all(row_id in st.session_state.selected_rows for row_id in rep_data['Row_ID'].tolist())
            select_all = st.checkbox(
                f"select_all_{rep}",
                value=all_rep_selected,
                key=f"select_all_check_{rep}",
                label_visibility="collapsed"
            )
            
            if select_all and not all_rep_selected:
                st.session_state.selected_rows.update(rep_data['Row_ID'].tolist())
                st.rerun()
            elif not select_all and all_rep_selected:
                st.session_state.selected_rows -= set(rep_data['Row_ID'].tolist())
                st.rerun()
        
        with col_expand:
            with st.expander(
                f"**{rep}** - Commission: ${rep_included_commission:,.2f} ({len(rep_included)}/{len(rep_data)} transactions)",
                expanded=False
            ):
                # Table header row
                header_cols = st.columns([0.4, 1.3, 0.6, 0.8, 2.2, 0.7, 1.3, 0.8, 0.7, 1.0])
                with header_cols[0]:
                    st.markdown("**☑️**")
                with header_cols[1]:
                    st.markdown("**Doc #**")
                with header_cols[2]:
                    st.markdown("**Status**")
                with header_cols[3]:
                    st.markdown("**Created From**")
                with header_cols[4]:
                    st.markdown("**Customer**")
                with header_cols[5]:
                    st.markdown("**CSM**")
                with header_cols[6]:
                    st.markdown("**Pipeline**")
                with header_cols[7]:
                    st.markdown("**Amount**")
                with header_cols[8]:
                    st.markdown("**Rep**")
                with header_cols[9]:
                    st.markdown("**Commission**")
                
                st.markdown("---")
                
                # Data rows
                for idx, row in rep_data.iterrows():
                    row_id = row['Row_ID']
                    is_selected = row_id in st.session_state.selected_rows
                    
                    cols = st.columns([0.4, 1.3, 0.6, 0.8, 2.2, 0.7, 1.3, 0.8, 0.7, 1.0])
                    
                    with cols[0]:
                        selected = st.checkbox(
                            "sel",
                            value=is_selected,
                            key=f"check_{row_id}",
                            label_visibility="collapsed"
                        )
                        
                        if selected != is_selected:
                            if selected:
                                st.session_state.selected_rows.add(row_id)
                            else:
                                st.session_state.selected_rows.discard(row_id)
                            st.rerun()
                    
                    with cols[1]:
                        st.text(str(row.get('Document Number', 'N/A')))
                    
                    with cols[2]:
                        status = str(row.get('Status', 'N/A'))
                        st.text(status[:10])
                    
                    with cols[3]:
                        created_from = str(row.get('Created From', 'N/A'))
                        st.text(created_from[:10])
                    
                    with cols[4]:
                        customer_name = str(row.get('Customer', 'N/A'))
                        if customer_name == 'N/A' or pd.isna(customer_name) or customer_name.strip() == '':
                            customer_name = 'No Customer Name'
                        st.text(customer_name)  # Show full customer name
                    
                    with cols[5]:
                        csm = str(row.get('CSM', 'N/A'))
                        st.text(csm[:8])
                    
                    with cols[6]:
                        pipeline = str(row.get('HubSpot Pipeline', 'N/A'))
                        st.text(pipeline[:15])
                    
                    with cols[7]:
                        st.text(f"${row.get('Amount', 0):,.0f}")
                    
                    with cols[8]:
                        rep_master = str(row.get('Rep Master', 'N/A'))
                        st.text(rep_master[:8])
                    
                    with cols[9]:
                        st.text(f"${row.get('Commission Amount', 0):,.2f}")

    st.markdown("---")
    
    # Export Options
    with st.expander("📥 Export Selected Data"):
        if not included_df.empty:
            # Select columns for export
            export_cols = [
                'Document Number', 'Status', 'Created From', 'Customer', 'CSM', 
                'Rep Master', 'HubSpot Pipeline', 'Date', 'Amount', 'Subtotal',
                'Commission Rate', 'Commission Amount', 'Brad Override'
            ]
            # Only include columns that exist
            export_cols = [col for col in export_cols if col in included_df.columns]
            
            export_df = included_df[export_cols].copy()
            
            # Format for export
            if 'Commission Rate' in export_df.columns:
                export_df['Commission Rate'] = export_df['Commission Rate'].apply(lambda x: f"{x:.1%}")
            
            csv = export_df.to_csv(index=False)
            st.download_button(
                label="Download Selected Transactions as CSV",
                data=csv,
                file_name=f"commission_report_{'-'.join(selected_months)}.csv",
                mime="text/csv"
            )

# ==========================================
# MAIN ENTRY POINT
# ==========================================

def display_commission_section(invoices_df=None, sales_orders_df=None):
    """
    Main function called by the dashboard
    """
    
    # Check Auth
    if not st.session_state.get('commission_authenticated', False):
        display_password_gate()
        return

    # Load Data
    with st.spinner("🔄 Loading invoice data..."):
        raw_invoices = fetch_google_sheet_data("NS Invoices", "A:U")
        
        if raw_invoices.empty:
            st.error("Could not load 'NS Invoices' from Google Sheet.")
            return

        clean_invoices = process_ns_invoices(raw_invoices)
        
        if clean_invoices.empty:
            st.error("❌ No data after processing invoices.")
            return

    # Display Dashboard
    display_commission_dashboard(clean_invoices)

if __name__ == "__main__":
    st.set_page_config(page_title="Commission Calculator", layout="wide")
    display_commission_section()
