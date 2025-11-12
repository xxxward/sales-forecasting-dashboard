"""
Commission Calculator Module for Calyx Containers
Handles commission calculations based on uploaded line-level invoice data from NetSuite
Processes invoices by month with 4-week payout delay
Password-protected access for Xander only
Supports large XLS/XLSX files (100MB+)
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib

# Note: openpyxl is used by pandas for reading Excel files
# Make sure it's in your requirements.txt: openpyxl>=3.0.0

# ==========================================
# PASSWORD CONFIGURATION (Xander Only)
# ==========================================
ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

# ==========================================
# COMMISSION RATE CONFIGURATION
# ==========================================

COMMISSION_RATES = {
    ("Dave Borkowski", "Growth Pipeline (Upsell/Cross-sell)"): 0.05,
    ("Jake Lynch", "Growth Pipeline (Upsell/Cross-sell)"): 0.07,
    ("Dave Borkowski", "Retention (Existing Product)"): 0.005,
    ("Jake Lynch", "Retention (Existing Product)"): 0.005,
    ("Brad Sherman", "Acquisition (New Customer)"): 0.07,
    ("Lance Mitton", "Acquisition (New Customer)"): 0.07,
    ("Alex Gonzalez", "Acquisition (New Customer)"): 0.07,
}

# Brad's override rate on Lance's deals
BRAD_OVERRIDE_RATE = 0.01
BRAD_OVERRIDE_REPS = ["Lance Mitton"]

# ==========================================
# EXCLUSION RULES
# ==========================================

EXCLUDED_ITEMS = [
    "Convenience Fee 3.5%",
    "Shipping",
    "Tax",
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

def parse_invoice_month(date_value):
    """Parse date and return the invoice month"""
    if pd.isna(date_value):
        return None
    
    try:
        if isinstance(date_value, str):
            date_obj = pd.to_datetime(date_value, errors='coerce')
        else:
            date_obj = pd.to_datetime(date_value)
        
        if pd.isna(date_obj):
            return None
        
        return date_obj.strftime('%Y-%m')  # Return as YYYY-MM
    except:
        return None

def calculate_payout_date(invoice_month_str):
    """
    Calculate payout date: 4 weeks after month close
    invoice_month_str format: 'YYYY-MM'
    """
    try:
        year, month = map(int, invoice_month_str.split('-'))
        # Last day of invoice month
        if month == 12:
            month_end = datetime(year, 12, 31)
        else:
            month_end = datetime(year, month + 1, 1) - timedelta(days=1)
        
        # Add 4 weeks
        payout_date = month_end + timedelta(weeks=4)
        return payout_date
    except:
        return None

def calculate_subtotal(row):
    """
    Calculate clean subtotal for commission calculation
    Excludes shipping, tax, and convenience fees
    Formula: netamountnotax - shippingamount
    """
    # Check for exclusions by item name
    item_name = str(row.get('Item Name', '')).strip() if not pd.isna(row.get('Item Name', '')) else ''
    
    # Skip excluded items
    if any(excluded in item_name for excluded in EXCLUDED_ITEMS):
        return 0
    
    # Skip tax and shipping lines (if these flags exist)
    if row.get('taxline', False) or row.get('shippingline', False):
        return 0
    
    # Get net amount and shipping amount
    net_amount = pd.to_numeric(row.get('netamountnotax', 0), errors='coerce')
    shipping_amount = pd.to_numeric(row.get('shippingamount', 0), errors='coerce')
    
    if pd.isna(net_amount):
        net_amount = 0
    if pd.isna(shipping_amount):
        shipping_amount = 0
    
    subtotal = net_amount - shipping_amount
    
    # Subtotal should not be negative
    return max(0, subtotal)

def get_payment_status(row):
    """Determine payment status of invoice"""
    amount_paid = pd.to_numeric(row.get('Amount Paid', 0), errors='coerce')
    amount_due = pd.to_numeric(row.get('Amount Due', 0), errors='coerce')
    
    if pd.isna(amount_paid):
        amount_paid = 0
    if pd.isna(amount_due):
        amount_due = 0
    
    if amount_paid == 0:
        return "No Payment"
    elif amount_due == 0:
        return "Fully Paid"
    else:
        return "Partially Paid"

def resolve_sales_rep(row):
    """Get the sales rep for commission calculation"""
    sales_rep = row.get('Sales Rep', '')
    
    # Handle any null/empty values
    if pd.isna(sales_rep) or sales_rep == '':
        return 'Unknown'
    
    return str(sales_rep).strip()

def get_commission_rate(sales_rep, pipeline):
    """Get commission rate for a given rep and pipeline combination"""
    # Normalize pipeline name
    if pd.isna(pipeline):
        pipeline = ''
    
    # Look up in commission rates table
    rate = COMMISSION_RATES.get((sales_rep, pipeline), 0)
    return rate

def calculate_brad_override(row, subtotal):
    """Calculate Brad's 1% override on Lance Mitton's deals"""
    sales_rep = row.get('Final Sales Rep', '')
    
    if sales_rep in BRAD_OVERRIDE_REPS:
        return subtotal * BRAD_OVERRIDE_RATE
    
    return 0

def process_commission_data(invoices_df):
    """
    Main function to process commission data from uploaded file
    Optimized for large datasets (100k+ rows)
    Returns a DataFrame with calculated commissions
    """
    if invoices_df.empty:
        return pd.DataFrame()
    
    # Make a copy to avoid modifying original
    df = invoices_df.copy()
    
    total_rows = len(df)
    st.info(f"Processing {total_rows:,} invoice line items...")
    
    # Create a progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Step 1: Resolve sales rep
    status_text.text("Step 1/8: Resolving sales reps...")
    df['Final Sales Rep'] = df.apply(resolve_sales_rep, axis=1)
    progress_bar.progress(12)
    
    # Step 2: Calculate subtotal (excluding fees, shipping, tax)
    status_text.text("Step 2/8: Calculating subtotals...")
    df['Subtotal'] = df.apply(calculate_subtotal, axis=1)
    progress_bar.progress(25)
    
    # Step 3: Get payment status
    status_text.text("Step 3/8: Checking payment status...")
    df['Payment Status'] = df.apply(get_payment_status, axis=1)
    progress_bar.progress(37)
    
    # Step 4: Filter to only paid/partially paid invoices
    status_text.text("Step 4/8: Filtering to paid invoices...")
    initial_count = len(df)
    df = df[df['Payment Status'].isin(['Fully Paid', 'Partially Paid'])].copy()
    filtered_count = len(df)
    st.caption(f"Filtered from {initial_count:,} to {filtered_count:,} paid/partially paid line items")
    progress_bar.progress(50)
    
    # Step 5: Parse invoice month and calculate payout date
    status_text.text("Step 5/8: Calculating invoice months and payout dates...")
    df['Invoice Month'] = df.apply(lambda row: parse_invoice_month(row.get('Transaction Date', '')), axis=1)
    df['Payout Date'] = df['Invoice Month'].apply(lambda x: calculate_payout_date(x) if x else None)
    progress_bar.progress(62)
    
    # Step 6: Get commission rate
    status_text.text("Step 6/8: Looking up commission rates...")
    df['Pipeline'] = df.get('custbody_calyx_hs_pipeline', '')
    df['Commission Rate'] = df.apply(
        lambda row: get_commission_rate(row['Final Sales Rep'], row['Pipeline']), 
        axis=1
    )
    progress_bar.progress(75)
    
    # Step 7: Calculate commission amount
    status_text.text("Step 7/8: Calculating commission amounts...")
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']
    progress_bar.progress(87)
    
    # Step 8: Calculate Brad's override
    status_text.text("Step 8/8: Calculating Brad's override...")
    df['Brad Override'] = df.apply(lambda row: calculate_brad_override(row, row['Subtotal']), axis=1)
    progress_bar.progress(100)
    
    # Clean up progress indicators
    status_text.text("✅ Processing complete!")
    progress_bar.empty()
    status_text.empty()
    
    return df

def calculate_summary_by_month(commission_df):
    """Calculate summary statistics by invoice month"""
    if commission_df.empty or 'Invoice Month' not in commission_df.columns:
        return pd.DataFrame()
    
    summary = commission_df.groupby('Invoice Month').agg({
        'Document Number': 'count',
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    summary.columns = ['Invoice Month', 'Transaction Count', 'Total Sales', 'Total Commission', 'Brad Override']
    
    # Add payout date
    summary['Payout Date'] = summary['Invoice Month'].apply(calculate_payout_date)
    
    return summary

def calculate_summary_by_rep(commission_df):
    """Calculate summary statistics by rep"""
    if commission_df.empty:
        return pd.DataFrame()
    
    summary = commission_df.groupby('Final Sales Rep').agg({
        'Document Number': 'count',
        'Subtotal': 'sum',
        'Commission Amount': 'sum',
        'Brad Override': 'sum'
    }).reset_index()
    
    summary.columns = ['Sales Rep', 'Transaction Count', 'Total Sales', 'Total Commission', 'Brad Override']
    
    return summary

# ==========================================
# STREAMLIT DISPLAY FUNCTIONS
# ==========================================

def display_password_gate():
    """Display password entry form for Xander"""
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
    
    return False

def display_file_uploader():
    """Display file upload interface for large XLS/XLSX/CSV files"""
    st.markdown("### 📁 Upload Invoice Data")
    st.caption("Upload your line-level invoice export from NetSuite (CSV, XLS, or XLSX format)")
    st.caption("⚠️ Large files (100MB+) may take a few minutes to process")
    
    uploaded_file = st.file_uploader(
        "Choose a file",
        type=['csv', 'xls', 'xlsx'],
        key="commission_file_upload"
    )
    
    if uploaded_file is not None:
        try:
            # Show file info
            file_size_mb = uploaded_file.size / (1024 * 1024)
            st.info(f"📦 File size: {file_size_mb:.2f} MB")
            
            # Read the file based on type
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            with st.spinner(f"Reading {file_extension.upper()} file... This may take a moment for large files"):
                if file_extension == 'csv':
                    df = pd.read_csv(uploaded_file, low_memory=False)
                elif file_extension in ['xls', 'xlsx']:
                    # Use openpyxl engine for better performance with large files
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                else:
                    st.error("Unsupported file format")
                    return None
            
            st.success(f"✅ File loaded successfully: {len(df):,} rows")
            
            # Show memory usage
            memory_usage_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)
            st.caption(f"Memory usage: {memory_usage_mb:.2f} MB")
            
            # Show column mapping info
            with st.expander("📋 View Column Names & Data Preview"):
                st.write("**Detected columns:**", df.columns.tolist())
                st.caption("Make sure your file includes: Sales Rep, Item Name, netamountnotax, shippingamount, Transaction Date, custbody_calyx_hs_pipeline, Amount Paid, Amount Due")
                
                st.markdown("**First 5 rows:**")
                st.dataframe(df.head(), use_container_width=True)
            
            return df
        
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.caption("If the file is very large, try exporting as CSV instead of XLS for better performance")
            return None
    
    return None

def display_commission_dashboard(commission_df):
    """Display the commission dashboard with all calculations"""
    
    # Header with logout
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
    
    # Overall Summary
    st.markdown("### 📊 Overall Summary")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total_transactions = len(commission_df)
    total_sales = commission_df['Subtotal'].sum()
    total_commission = commission_df['Commission Amount'].sum()
    total_override = commission_df['Brad Override'].sum()
    
    with col1:
        st.metric("Total Transactions", f"{total_transactions:,}")
    
    with col2:
        st.metric("Total Sales Volume", f"${total_sales:,.2f}")
    
    with col3:
        st.metric("Total Commission", f"${total_commission:,.2f}")
    
    with col4:
        st.metric("Brad's Override", f"${total_override:,.2f}")
    
    st.markdown("---")
    
    # Summary by Month
    st.markdown("### 📅 Commission by Invoice Month")
    st.caption("Shows when invoices were created and when commissions will be paid (4 weeks after month close)")
    
    month_summary = calculate_summary_by_month(commission_df)
    
    if not month_summary.empty:
        # Format for display
        display_month = month_summary.copy()
        display_month['Total Sales'] = display_month['Total Sales'].apply(lambda x: f"${x:,.2f}")
        display_month['Total Commission'] = display_month['Total Commission'].apply(lambda x: f"${x:,.2f}")
        display_month['Brad Override'] = display_month['Brad Override'].apply(lambda x: f"${x:,.2f}")
        display_month['Payout Date'] = display_month['Payout Date'].apply(lambda x: x.strftime('%Y-%m-%d') if not pd.isna(x) else '')
        
        st.dataframe(display_month, use_container_width=True, hide_index=True)
    else:
        st.info("No monthly data available")
    
    st.markdown("---")
    
    # Summary by Rep
    st.markdown("### 👤 Commission by Sales Rep")
    
    rep_summary = calculate_summary_by_rep(commission_df)
    
    if not rep_summary.empty:
        # Format for display
        display_rep = rep_summary.copy()
        display_rep['Total Sales'] = display_rep['Total Sales'].apply(lambda x: f"${x:,.2f}")
        display_rep['Total Commission'] = display_rep['Total Commission'].apply(lambda x: f"${x:,.2f}")
        display_rep['Brad Override'] = display_rep['Brad Override'].apply(lambda x: f"${x:,.2f}")
        
        st.dataframe(display_rep, use_container_width=True, hide_index=True)
    else:
        st.info("No rep data available")
    
    st.markdown("---")
    
    # Detailed Transaction View
    st.markdown("### 📋 Detailed Transaction View")
    st.caption(f"Showing all {len(commission_df)} commissionable transactions")
    
    # Filter by rep
    all_reps = ['All Reps'] + sorted(commission_df['Final Sales Rep'].unique().tolist())
    selected_rep = st.selectbox("Filter by Rep:", all_reps)
    
    # Filter by month
    if 'Invoice Month' in commission_df.columns:
        all_months = ['All Months'] + sorted(commission_df['Invoice Month'].dropna().unique().tolist(), reverse=True)
        selected_month = st.selectbox("Filter by Month:", all_months)
    else:
        selected_month = 'All Months'
    
    # Apply filters
    filtered_df = commission_df.copy()
    
    if selected_rep != 'All Reps':
        filtered_df = filtered_df[filtered_df['Final Sales Rep'] == selected_rep]
    
    if selected_month != 'All Months' and 'Invoice Month' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Invoice Month'] == selected_month]
    
    st.caption(f"Filtered to {len(filtered_df)} transactions")
    
    # Select display columns
    display_columns = [
        'Document Number', 'Transaction Date', 'Final Sales Rep', 'Pipeline',
        'Item Name', 'Subtotal', 'Commission Rate', 'Commission Amount',
        'Brad Override', 'Payment Status', 'Invoice Month', 'Payout Date'
    ]
    
    # Only show columns that exist
    available_columns = [col for col in display_columns if col in filtered_df.columns]
    display_df = filtered_df[available_columns].copy()
    
    # Format currency columns
    currency_columns = ['Subtotal', 'Commission Amount', 'Brad Override']
    for col in currency_columns:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(lambda x: f"${x:,.2f}" if not pd.isna(x) else "$0.00")
    
    # Format percentage
    if 'Commission Rate' in display_df.columns:
        display_df['Commission Rate'] = (display_df['Commission Rate'] * 100).apply(
            lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%"
        )
    
    # Format dates
    if 'Payout Date' in display_df.columns:
        display_df['Payout Date'] = display_df['Payout Date'].apply(
            lambda x: x.strftime('%Y-%m-%d') if not pd.isna(x) else ''
        )
    
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # Download button
    csv = filtered_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Filtered Data (CSV)",
        data=csv,
        file_name=f"commission_data_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )

def display_commission_section(invoices_df=None, sales_orders_df=None):
    """
    Main entry point for commission section
    Handles authentication, file upload, and display
    """
    # Check if user is authenticated
    if not st.session_state.get('commission_authenticated', False):
        display_password_gate()
        return
    
    # Display file uploader
    uploaded_df = display_file_uploader()
    
    if uploaded_df is not None:
        # Process the uploaded data
        with st.spinner("Calculating commissions..."):
            commission_df = process_commission_data(uploaded_df)
        
        if commission_df.empty:
            st.warning("⚠️ No commissionable transactions found in uploaded file.")
            st.info("Make sure your file includes paid/partially paid invoices with the required fields.")
            return
        
        # Display the dashboard
        display_commission_dashboard(commission_df)
    else:
        st.info("👆 Please upload your invoice data file to begin")

if __name__ == "__main__":
    st.set_page_config(page_title="Commission Calculator", page_icon="💰", layout="wide")
    display_commission_section()
