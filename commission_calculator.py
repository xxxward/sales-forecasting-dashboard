"""
Commission Calculator Module for Calyx Containers
Simplified version - processes stored invoice data for commission calculations
Focuses on Sep/Oct 2025 for Dave, Jake, Brad, and Lance only
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib

# ==========================================
# PASSWORD CONFIGURATION (Xander Only)
# ==========================================
ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

# ==========================================
# COMMISSION CONFIGURATION - 4 REPS ONLY
# ==========================================

# Only these 4 reps get commissions
COMMISSION_REPS = ["Dave Borkowski", "Jake Lynch", "Brad Sherman", "Lance Mitton"]

# Commission rates by rep (simplified - no pipeline needed)
REP_COMMISSION_RATES = {
    "Dave Borkowski": 0.05,      # 5% flat rate
    "Jake Lynch": 0.07,          # 7% flat rate
    "Brad Sherman": 0.07,        # 7% flat rate
    "Lance Mitton": 0.07,        # 7% flat rate
}

# Brad's 1% override on Lance's deals
BRAD_OVERRIDE_RATE = 0.01

# Focus months for commission calculation
COMMISSION_MONTHS = ['2025-09', '2025-10']  # September and October 2025

# ==========================================
# EXCLUSION RULES
# ==========================================

EXCLUDED_ITEMS = [
    "Convenience Fee 3.5%",
    "SHIPPING",
    "UPS",
    "FEDEX",
    "AVATAX",
    "TAX",
]

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def verify_admin(email, password):
    """Verify admin credentials"""
    if email != ADMIN_EMAIL:
        return False
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    return password_hash == ADMIN_PASSWORD_HASH

def load_invoice_data():
    """
    Load invoice line data from the stored CSV file
    This should be a file you upload once and store
    """
    try:
        # Try to load from session state first (if already uploaded this session)
        if 'invoice_data' in st.session_state:
            return st.session_state.invoice_data
        
        # Otherwise, need to upload
        return None
    except:
        return None

def should_include_line(row):
    """
    Quick check if this line should be included in commission calculation
    Returns True/False
    """
    # Check rep
    sales_rep = str(row.get('Sales Rep', '')).strip()
    if sales_rep not in COMMISSION_REPS:
        return False
    
    # Check if paid (Amount Remaining = 0)
    amount_remaining = pd.to_numeric(row.get('Amount Remaining', 0), errors='coerce')
    if pd.isna(amount_remaining):
        amount_remaining = 0
    if amount_remaining != 0:
        return False
    
    # Check item for exclusions
    item_name = str(row.get('Item', '')).strip().upper()
    for excluded in EXCLUDED_ITEMS:
        if excluded.upper() in item_name:
            return False
    
    return True

def parse_invoice_month(date_value):
    """Parse date and return YYYY-MM format"""
    if pd.isna(date_value):
        return None
    try:
        date_obj = pd.to_datetime(date_value, errors='coerce')
        if pd.isna(date_obj):
            return None
        return date_obj.strftime('%Y-%m')
    except:
        return None

def calculate_payout_date(invoice_month_str):
    """Calculate payout date: 4 weeks after month close"""
    try:
        year, month = map(int, invoice_month_str.split('-'))
        if month == 12:
            month_end = datetime(year, 12, 31)
        else:
            month_end = datetime(year, month + 1, 1) - timedelta(days=1)
        payout_date = month_end + timedelta(weeks=4)
        return payout_date
    except:
        return None

def process_commission_data_fast(df):
    """
    Fast processing - filter and calculate in one pass
    Only processes Sep/Oct 2025 for 4 reps
    """
    if df.empty:
        return pd.DataFrame()
    
    # Clean BOM
    df.columns = df.columns.str.replace('﻿', '')
    
    st.info(f"📊 Starting with {len(df):,} total invoice lines")
    
    # Step 1: Quick filter to 4 reps only
    df = df[df['Sales Rep'].isin(COMMISSION_REPS)].copy()
    st.caption(f"✓ Filtered to 4 commission reps: {len(df):,} lines")
    
    # Step 2: Only paid invoices (Amount Remaining = 0)
    df['Amount Remaining Numeric'] = pd.to_numeric(df.get('Amount Remaining', 0), errors='coerce').fillna(0)
    df = df[df['Amount Remaining Numeric'] == 0].copy()
    st.caption(f"✓ Filtered to paid invoices: {len(df):,} lines")
    
    # Step 3: Parse dates and filter to Sep/Oct 2025
    df['Invoice Month'] = df['Date'].apply(parse_invoice_month)
    df = df[df['Invoice Month'].isin(COMMISSION_MONTHS)].copy()
    st.caption(f"✓ Filtered to Sep/Oct 2025: {len(df):,} lines")
    
    # Step 4: Exclude shipping, tax, fees
    df['Item Upper'] = df['Item'].str.upper()
    
    def is_excluded(item_upper):
        if pd.isna(item_upper):
            return False
        return any(excl in item_upper for excl in EXCLUDED_ITEMS)
    
    df['Is Excluded'] = df['Item Upper'].apply(is_excluded)
    df = df[~df['Is Excluded']].copy()
    st.caption(f"✓ Excluded shipping/tax/fees: {len(df):,} commissionable lines")
    
    if df.empty:
        return df
    
    # Step 5: Calculate commissions
    df['Amount Numeric'] = pd.to_numeric(df['Amount'], errors='coerce').fillna(0)
    df['Subtotal'] = df['Amount Numeric']
    
    # Get commission rate per rep
    df['Commission Rate'] = df['Sales Rep'].map(REP_COMMISSION_RATES)
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']
    
    # Brad's override on Lance's deals
    df['Brad Override'] = 0.0
    lance_mask = df['Sales Rep'] == 'Lance Mitton'
    df.loc[lance_mask, 'Brad Override'] = df.loc[lance_mask, 'Subtotal'] * BRAD_OVERRIDE_RATE
    
    # Add payout dates
    df['Payout Date'] = df['Invoice Month'].apply(calculate_payout_date)
    
    # Clean up temp columns
    df = df.drop(columns=['Amount Remaining Numeric', 'Item Upper', 'Is Excluded', 'Amount Numeric'], errors='ignore')
    
    st.success(f"✅ Commission calculated on {len(df):,} line items!")
    
    return df

# ==========================================
# SUMMARY FUNCTIONS
# ==========================================

def calculate_summary_by_month(commission_df):
    """Summary by month"""
    if commission_df.empty:
        return pd.DataFrame()
    
    summary = commission_df.groupby('Invoice Month').agg({
        'Document Number': 'nunique',
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    summary.columns = ['Invoice Month', 'Unique Invoices', 'Total Sales', 'Total Commission', 'Brad Override']
    summary['Payout Date'] = summary['Invoice Month'].apply(calculate_payout_date)
    
    return summary

def calculate_summary_by_rep(commission_df):
    """Summary by rep"""
    if commission_df.empty:
        return pd.DataFrame()
    
    summary = commission_df.groupby('Sales Rep').agg({
        'Document Number': 'nunique',
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    summary.columns = ['Sales Rep', 'Unique Invoices', 'Total Sales', 'Total Commission', 'Brad Override']
    
    # Add Brad's total override earnings
    brad_total_override = commission_df['Brad Override'].sum()
    if brad_total_override > 0:
        # Add a row for Brad's override
        brad_override_row = pd.DataFrame({
            'Sales Rep': ['Brad Sherman (Override)'],
            'Unique Invoices': [0],
            'Total Sales': [0],
            'Total Commission': [0],
            'Brad Override': [brad_total_override]
        })
        summary = pd.concat([summary, brad_override_row], ignore_index=True)
    
    return summary

def calculate_rep_month_summary(commission_df):
    """Summary by rep and month"""
    if commission_df.empty:
        return pd.DataFrame()
    
    summary = commission_df.groupby(['Sales Rep', 'Invoice Month']).agg({
        'Document Number': 'nunique',
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    summary.columns = ['Sales Rep', 'Invoice Month', 'Unique Invoices', 'Total Sales', 'Total Commission', 'Brad Override']
    
    return summary

# ==========================================
# STREAMLIT DISPLAY FUNCTIONS
# ==========================================

def display_password_gate():
    """Display password entry form"""
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 padding: 20px; border-radius: 10px; color: white; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0;'>🔒 Commission Calculator</h2>
        <p style='margin: 5px 0 0 0;'>Admin access required</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        email = st.text_input("Email:", value="", key="commission_email")
        password = st.text_input("Password:", type="password", key="commission_password")
        
        col_a, col_b, col_c = st.columns([1, 1, 1])
        with col_b:
            if st.button("🔓 Login", use_container_width=True):
                if verify_admin(email, password):
                    st.session_state.commission_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ Invalid credentials")

def display_data_manager():
    """Manage the stored invoice data"""
    st.markdown("### 📁 Invoice Data Management")
    
    # Check if data is already loaded
    if 'invoice_data' in st.session_state and st.session_state.invoice_data is not None:
        df = st.session_state.invoice_data
        st.success(f"✅ Invoice data loaded: {len(df):,} total rows")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Upload New Data"):
                del st.session_state.invoice_data
                st.rerun()
        
        return df
    
    # No data loaded - show uploader
    st.caption("Upload your Invoice Line Level CSV export from NetSuite")
    
    uploaded_file = st.file_uploader(
        "Choose CSV file",
        type=['csv'],
        key="invoice_data_upload"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file, low_memory=False)
            df.columns = df.columns.str.replace('﻿', '')  # Clean BOM
            
            st.success(f"✅ Loaded {len(df):,} invoice lines")
            
            # Store in session state
            st.session_state.invoice_data = df
            
            st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            return None
    
    return None

def display_commission_dashboard(invoice_df):
    """Display the commission dashboard"""
    
    # Header
    col1, col2 = st.columns([4, 1])
    
    with col1:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                     padding: 15px; border-radius: 10px; color: white;'>
            <h2 style='color: white; margin: 0;'>💰 Commission Dashboard - Sep/Oct 2025</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.commission_authenticated = False
            if 'invoice_data' in st.session_state:
                del st.session_state.invoice_data
            st.rerun()
    
    st.markdown("---")
    
    # Process commissions
    with st.spinner("Calculating commissions..."):
        commission_df = process_commission_data_fast(invoice_df)
    
    if commission_df.empty:
        st.warning("⚠️ No commissionable transactions found")
        return
    
    # Overall Summary
    st.markdown("### 📊 Overall Summary (Sep + Oct 2025)")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Line Items", f"{len(commission_df):,}")
    
    with col2:
        total_sales = commission_df['Subtotal'].sum()
        st.metric("Total Sales", f"${total_sales:,.2f}")
    
    with col3:
        total_commission = commission_df['Commission Amount'].sum()
        st.metric("Total Commission", f"${total_commission:,.2f}")
    
    with col4:
        total_override = commission_df['Brad Override'].sum()
        st.metric("Brad's Override", f"${total_override:,.2f}")
    
    st.markdown("---")
    
    # By Month
    st.markdown("### 📅 Commission by Month")
    month_summary = calculate_summary_by_month(commission_df)
    
    if not month_summary.empty:
        display_month = month_summary.copy()
        display_month['Total Sales'] = display_month['Total Sales'].apply(lambda x: f"${x:,.2f}")
        display_month['Total Commission'] = display_month['Total Commission'].apply(lambda x: f"${x:,.2f}")
        display_month['Brad Override'] = display_month['Brad Override'].apply(lambda x: f"${x:,.2f}")
        display_month['Payout Date'] = display_month['Payout Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isna(x) else '')
        
        st.dataframe(display_month, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # By Rep
    st.markdown("### 👤 Commission by Sales Rep")
    rep_summary = calculate_summary_by_rep(commission_df)
    
    if not rep_summary.empty:
        display_rep = rep_summary.copy()
        display_rep['Total Sales'] = display_rep['Total Sales'].apply(lambda x: f"${x:,.2f}")
        display_rep['Total Commission'] = display_rep['Total Commission'].apply(lambda x: f"${x:,.2f}")
        display_rep['Brad Override'] = display_rep['Brad Override'].apply(lambda x: f"${x:,.2f}")
        
        st.dataframe(display_rep, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # Rep x Month breakdown
    st.markdown("### 📊 Commission by Rep & Month")
    rep_month_summary = calculate_rep_month_summary(commission_df)
    
    if not rep_month_summary.empty:
        display_rm = rep_month_summary.copy()
        display_rm['Total Sales'] = display_rm['Total Sales'].apply(lambda x: f"${x:,.2f}")
        display_rm['Total Commission'] = display_rm['Total Commission'].apply(lambda x: f"${x:,.2f}")
        display_rm['Brad Override'] = display_rm['Brad Override'].apply(lambda x: f"${x:,.2f}")
        
        st.dataframe(display_rm, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # Detailed view with filters
    st.markdown("### 📋 Detailed Transaction View")
    
    col1, col2 = st.columns(2)
    
    with col1:
        filter_reps = ['All'] + COMMISSION_REPS
        selected_rep = st.selectbox("Filter by Rep:", filter_reps)
    
    with col2:
        filter_months = ['All'] + COMMISSION_MONTHS
        selected_month = st.selectbox("Filter by Month:", filter_months)
    
    # Apply filters
    filtered_df = commission_df.copy()
    
    if selected_rep != 'All':
        filtered_df = filtered_df[filtered_df['Sales Rep'] == selected_rep]
    
    if selected_month != 'All':
        filtered_df = filtered_df[filtered_df['Invoice Month'] == selected_month]
    
    st.caption(f"Showing {len(filtered_df):,} line items")
    
    # Display columns
    display_columns = [
        'Document Number', 'Date', 'Sales Rep', 'Customer', 'Item',
        'Subtotal', 'Commission Rate', 'Commission Amount', 'Brad Override',
        'Invoice Month', 'Payout Date'
    ]
    
    display_df = filtered_df[display_columns].copy()
    
    # Format
    display_df['Subtotal'] = display_df['Subtotal'].apply(lambda x: f"${x:,.2f}")
    display_df['Commission Rate'] = (display_df['Commission Rate'] * 100).apply(lambda x: f"{x:.1f}%")
    display_df['Commission Amount'] = display_df['Commission Amount'].apply(lambda x: f"${x:,.2f}")
    display_df['Brad Override'] = display_df['Brad Override'].apply(lambda x: f"${x:,.2f}" if x > 0 else "")
    display_df['Payout Date'] = display_df['Payout Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isna(x) else '')
    
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # Download
    csv = filtered_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Filtered Data (CSV)",
        data=csv,
        file_name=f"commission_data_{selected_rep}_{selected_month}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

def display_commission_section(invoices_df=None, sales_orders_df=None):
    """Main entry point"""
    
    # Check authentication
    if not st.session_state.get('commission_authenticated', False):
        display_password_gate()
        return
    
    # Manage data
    invoice_df = display_data_manager()
    
    if invoice_df is not None:
        display_commission_dashboard(invoice_df)
    else:
        st.info("👆 Please upload your invoice data CSV to begin")

if __name__ == "__main__":
    st.set_page_config(page_title="Commission Calculator", page_icon="💰", layout="wide")
    display_commission_section()
