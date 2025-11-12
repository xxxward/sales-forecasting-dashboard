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
import re

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
# EXPECTED COMMISSION VALUES (For Reconciliation)
# ==========================================

EXPECTED_COMMISSIONS = {
    "Brad Sherman": {
        "2025-09": {
            "acquisition_total": 124022.00,
            "acquisition_commission_7pct": 8681.53,
            "lance_override_base": 19093.00,
            "lance_override_1pct": 190.93,
            "total_commission": 8872.45,
            "sales_orders": [
                "SO13290", "SO13087", "SO13194", "SO13097", "SO13105",
                "SO13259", "SO13241", "SO13137", "SO13094", "SO12943",
                "SO13109", "SO13057", "SO13006", "SO12445", "SO13041",
                "SO13231", "SO13147"
            ]
        }
    },
    # Add more reps and months as you get the expected values
}

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

def extract_so_number(created_from):
    """Extract SO number from 'Sales Order #SO12345' format"""
    if pd.isna(created_from):
        return None
    
    created_from = str(created_from)
    
    # Look for SO followed by numbers
    match = re.search(r'SO(\d+)', created_from)
    if match:
        return f"SO{match.group(1)}"
    
    return None

def process_commission_data_fast(df):
    """
    Fast processing - filter and calculate in one pass
    Only processes Sep/Oct 2025 for 4 reps
    Uses Date Closed (payment date) to determine commission month
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
    
    # Step 3: Parse CLOSE dates (when payment received) and filter to Sep/Oct 2025
    st.caption("⏰ Using 'Date Closed' to determine commission month (when payment was received)")
    df['Invoice Month'] = df['Date Closed'].apply(parse_invoice_month)
    
    # Show how many have no close date
    no_close_date = df['Invoice Month'].isna().sum()
    if no_close_date > 0:
        st.warning(f"⚠️ {no_close_date:,} lines have no Date Closed value and will be excluded")
    
    df = df[df['Invoice Month'].isin(COMMISSION_MONTHS)].copy()
    st.caption(f"✓ Filtered to Sep/Oct 2025 (by Date Closed): {len(df):,} lines")
    
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
    
    # Add tabs for Dashboard vs Reconciliation
    tab1, tab2 = st.tabs(["📊 Commission Dashboard", "🔍 Reconciliation Tool"])
    
    with tab1:
        display_commission_calculations(invoice_df)
    
    with tab2:
        display_reconciliation_tool(invoice_df)

def display_commission_calculations(invoice_df):
    """Display the main commission calculations"""
    
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

def display_reconciliation_tool(invoice_df):
    """Reconciliation tool to compare expected vs calculated commissions"""
    st.markdown("### 🔍 Commission Reconciliation Tool")
    st.caption("Compare your boss's expected values against calculated commissions")
    
    # Select rep and month
    col1, col2 = st.columns(2)
    
    with col1:
        available_reps = list(EXPECTED_COMMISSIONS.keys())
        if not available_reps:
            st.warning("No expected commission data loaded yet. Add values to EXPECTED_COMMISSIONS in the code.")
            return
        selected_rep = st.selectbox("Select Rep:", available_reps, key="recon_rep")
    
    with col2:
        if selected_rep in EXPECTED_COMMISSIONS:
            available_months = list(EXPECTED_COMMISSIONS[selected_rep].keys())
            selected_month = st.selectbox("Select Month:", available_months, key="recon_month")
        else:
            return
    
    if not selected_rep or not selected_month:
        return
    
    expected_data = EXPECTED_COMMISSIONS[selected_rep][selected_month]
    
    st.markdown("---")
    st.markdown(f"### Reconciling: {selected_rep} - {selected_month}")
    
    # Prepare data
    df = invoice_df.copy()
    df.columns = df.columns.str.replace('﻿', '')
    
    # Parse dates and extract SO numbers
    df['Date Parsed'] = pd.to_datetime(df['Date Closed'], errors='coerce')  # Use Date Closed!
    df['Invoice Month'] = df['Date Parsed'].apply(
        lambda x: x.strftime('%Y-%m') if not pd.isna(x) else None
    )
    df['SO Number'] = df['Created From'].apply(extract_so_number)
    
    # Filter to this rep and month
    rep_data = df[(df['Sales Rep'] == selected_rep) & (df['Invoice Month'] == selected_month)].copy()
    
    st.info(f"Found {len(rep_data):,} total invoice lines")
    
    # Filter to paid only
    rep_data['Amount Remaining Numeric'] = pd.to_numeric(rep_data.get('Amount Remaining', 0), errors='coerce').fillna(0)
    paid_data = rep_data[rep_data['Amount Remaining Numeric'] == 0].copy()
    st.caption(f"Paid invoices: {len(paid_data):,} lines")
    
    # Exclude shipping/tax
    paid_data['Item Upper'] = paid_data['Item'].str.upper()
    excluded_keywords = ['SHIPPING', 'UPS', 'FEDEX', 'AVATAX', 'TAX', 'CONVENIENCE FEE']
    paid_data['Is Excluded'] = paid_data['Item Upper'].apply(
        lambda x: any(kw in str(x) for kw in excluded_keywords) if not pd.isna(x) else False
    )
    commissionable = paid_data[~paid_data['Is Excluded']].copy()
    st.caption(f"Commissionable lines: {len(commissionable):,}")
    
    # Calculate amounts
    commissionable['Amount Numeric'] = pd.to_numeric(commissionable['Amount'], errors='coerce').fillna(0)
    
    # Compare SOs
    st.markdown("#### 📦 Sales Order Comparison")
    
    our_sos = set(commissionable['SO Number'].dropna().unique())
    expected_sos = set(expected_data.get('sales_orders', []))
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Expected SOs", len(expected_sos))
        with st.expander("View List"):
            for so in sorted(expected_sos):
                st.text(so)
    
    with col2:
        st.metric("Found SOs", len(our_sos))
        with st.expander("View List"):
            for so in sorted(our_sos):
                st.text(so)
    
    with col3:
        missing_sos = expected_sos - our_sos
        extra_sos = our_sos - expected_sos
        matched_sos = expected_sos & our_sos
        
        st.metric("Matched", len(matched_sos))
        
        if missing_sos:
            with st.expander(f"❌ Missing ({len(missing_sos)})"):
                for so in sorted(missing_sos):
                    st.text(so)
        
        if extra_sos:
            with st.expander(f"➕ Extra ({len(extra_sos)})"):
                for so in sorted(extra_sos):
                    st.text(so)
    
    # Compare amounts
    st.markdown("---")
    st.markdown("#### 💰 Amount Comparison")
    
    calculated_total = commissionable['Amount Numeric'].sum()
    expected_total = expected_data.get('acquisition_total', 0)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Expected Total", f"${expected_total:,.2f}")
    
    with col2:
        st.metric("Calculated Total", f"${calculated_total:,.2f}")
    
    with col3:
        diff = calculated_total - expected_total
        diff_pct = (diff / expected_total * 100) if expected_total > 0 else 0
        st.metric("Difference", f"${diff:,.2f}", delta=f"{diff_pct:+.2f}%")
    
    # Commission comparison
    st.markdown("---")
    st.markdown("#### 🧮 Commission Comparison")
    
    calc_comm = calculated_total * 0.07
    exp_comm = expected_data.get('acquisition_commission_7pct', 0)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Expected Commission", f"${exp_comm:,.2f}")
    
    with col2:
        st.metric("Calculated Commission", f"${calc_comm:,.2f}")
    
    with col3:
        comm_diff = calc_comm - exp_comm
        st.metric("Difference", f"${comm_diff:,.2f}", delta=f"${comm_diff:,.2f}")
    
    # Investigate missing SOs
    if missing_sos:
        st.markdown("---")
        st.markdown("#### 🔎 Investigating Missing Sales Orders")
        
        for so in sorted(missing_sos):
            st.markdown(f"**{so}**")
            
            # Check if SO exists anywhere
            so_data = df[df['SO Number'] == so]
            
            if so_data.empty:
                st.error(f"❌ Not found in invoice data at all")
            else:
                st.warning(f"⚠️ Found but filtered out")
                
                # Show why
                so_rep = so_data['Sales Rep'].iloc[0]
                so_month = so_data['Invoice Month'].iloc[0]
                so_amt_remain = pd.to_numeric(so_data['Amount Remaining'].iloc[0], errors='coerce')
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    if so_rep != selected_rep:
                        st.error(f"Rep: {so_rep}")
                    else:
                        st.success(f"Rep: {so_rep} ✓")
                
                with col2:
                    if so_month != selected_month:
                        st.error(f"Month: {so_month}")
                    else:
                        st.success(f"Month: {so_month} ✓")
                
                with col3:
                    if pd.isna(so_amt_remain):
                        so_amt_remain = 0
                    if so_amt_remain > 0:
                        st.error(f"Unpaid: ${so_amt_remain:,.2f}")
                    else:
                        st.success("Paid ✓")
                
                with st.expander(f"View {so} details"):
                    display_cols = ['Document Number', 'Date', 'Item', 'Amount', 'Amount Remaining', 'Status']
                    display_cols = [c for c in display_cols if c in so_data.columns]
                    st.dataframe(so_data[display_cols], use_container_width=True)
    
    # Download reconciliation report
    st.markdown("---")
    
    export_df = commissionable[[
        'Document Number', 'Date', 'SO Number', 'Sales Rep', 'Customer',
        'Item', 'Amount Numeric', 'Amount Remaining', 'Status'
    ]].copy()
    
    export_df['In Expected List'] = export_df['SO Number'].isin(expected_sos)
    export_df.columns = ['Invoice', 'Date', 'SO#', 'Rep', 'Customer', 'Item', 'Amount', 'Amt Remaining', 'Status', 'Expected?']
    
    csv = export_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Reconciliation Report (CSV)",
        data=csv,
        file_name=f"reconciliation_{selected_rep.replace(' ', '_')}_{selected_month}.csv",
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
