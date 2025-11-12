"""
Commission Calculator Module for Calyx Containers
Handles commission calculations based on NetSuite invoice and sales order data
Password-protected access per rep
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import hashlib

# ==========================================
# PASSWORD CONFIGURATION
# ==========================================
# Store password hashes (not plain text)
# To generate: hashlib.sha256("password".encode()).hexdigest()

REP_PASSWORDS = {
    "Dave Borkowski": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password"
    "Jake Lynch": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password"
    "Brad Sherman": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password"
    "Lance Mitton": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password"
    "Alex Gonzalez": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password"
    "Kyle Bissell": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password" (VP view all)
    "Xander": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # "password" (admin view all)
}

# Reps who can view all commission data (typically managers/admins)
ADMIN_REPS = ["Kyle Bissell", "Xander"]

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
    ("Alex Gonzalez", "Acquisition (New Customer)"): 0.07,  # Assuming same rate
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

def hash_password(password):
    """Hash a password for storing or comparison"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(rep_name, password):
    """Verify password for a given rep"""
    if rep_name not in REP_PASSWORDS:
        return False
    return hash_password(password) == REP_PASSWORDS[rep_name]

def is_admin(rep_name):
    """Check if rep has admin access"""
    return rep_name in ADMIN_REPS

def calculate_subtotal(row):
    """
    Calculate clean subtotal for commission calculation
    Excludes shipping, tax, and convenience fees
    """
    # Check for exclusions
    if pd.isna(row.get('Item Name', '')):
        item_name = ''
    else:
        item_name = str(row.get('Item Name', ''))
    
    # Skip excluded items
    if any(excluded in item_name for excluded in EXCLUDED_ITEMS):
        return 0
    
    # Skip tax and shipping lines
    if row.get('taxline', False) or row.get('shippingline', False):
        return 0
    
    # Calculate subtotal
    net_amount = pd.to_numeric(row.get('netamountnotax', 0), errors='coerce')
    shipping_amount = pd.to_numeric(row.get('shippingamount', 0), errors='coerce')
    
    if pd.isna(net_amount):
        net_amount = 0
    if pd.isna(shipping_amount):
        shipping_amount = 0
    
    return net_amount - shipping_amount

def get_payment_status(row):
    """
    Determine payment status of invoice
    """
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

def resolve_sales_rep(row, sales_orders_df=None):
    """
    Handle Shopify ECommerce → Customer Owner mapping
    Returns the final sales rep to use for commission
    """
    sales_rep = row.get('Sales Rep', '')
    
    # If rep is Shopify ECommerce, look up Customer Owner from linked Sales Order
    if sales_rep == "Shopify ECommerce" and sales_orders_df is not None:
        created_from = row.get('Created From', '')
        if created_from and not pd.isna(created_from):
            # Find matching sales order
            matching_so = sales_orders_df[sales_orders_df['Document Number'] == created_from]
            if not matching_so.empty:
                customer_owner = matching_so.iloc[0].get('Customer Owner', '')
                if customer_owner and not pd.isna(customer_owner):
                    return customer_owner
    
    return sales_rep

def get_commission_rate(sales_rep, pipeline):
    """
    Get commission rate for a given rep and pipeline combination
    """
    # Normalize pipeline name
    if pd.isna(pipeline):
        pipeline = ''
    
    # Look up in commission rates table
    rate = COMMISSION_RATES.get((sales_rep, pipeline), 0)
    return rate

def calculate_brad_override(row, subtotal):
    """
    Calculate Brad's 1% override on Lance Mitton's deals
    """
    sales_rep = row.get('Final Sales Rep', '')
    
    if sales_rep in BRAD_OVERRIDE_REPS:
        return subtotal * BRAD_OVERRIDE_RATE
    
    return 0

def process_commission_data(invoices_df, sales_orders_df=None):
    """
    Main function to process commission data
    Returns a DataFrame with calculated commissions
    """
    if invoices_df.empty:
        return pd.DataFrame()
    
    # Make a copy to avoid modifying original
    df = invoices_df.copy()
    
    # Step 1: Resolve sales rep (handle Shopify mapping)
    df['Final Sales Rep'] = df.apply(lambda row: resolve_sales_rep(row, sales_orders_df), axis=1)
    
    # Step 2: Calculate subtotal (excluding fees, shipping, tax)
    df['Subtotal'] = df.apply(calculate_subtotal, axis=1)
    
    # Step 3: Get payment status
    df['Payment Status'] = df.apply(get_payment_status, axis=1)
    
    # Step 4: Filter to only paid/partially paid invoices
    df = df[df['Payment Status'].isin(['Fully Paid', 'Partially Paid'])].copy()
    
    # Step 5: Get commission rate
    df['Pipeline'] = df.get('custbody_calyx_hs_pipeline', '')
    df['Commission Rate'] = df.apply(
        lambda row: get_commission_rate(row['Final Sales Rep'], row['Pipeline']), 
        axis=1
    )
    
    # Step 6: Calculate commission amount
    df['Commission Amount'] = df['Subtotal'] * df['Commission Rate']
    
    # Step 7: Calculate Brad's override
    df['Brad Override'] = df.apply(lambda row: calculate_brad_override(row, row['Subtotal']), axis=1)
    
    # Step 8: Select and format output columns
    output_columns = [
        'Document Number',
        'Transaction Date',
        'Final Sales Rep',
        'Pipeline',
        'Subtotal',
        'Commission Rate',
        'Commission Amount',
        'Brad Override',
        'Payment Status',
        'Amount Paid',
        'Amount Due'
    ]
    
    # Only include columns that exist
    available_columns = [col for col in output_columns if col in df.columns]
    result_df = df[available_columns].copy()
    
    # Format commission rate as percentage
    if 'Commission Rate' in result_df.columns:
        result_df['Commission Rate %'] = (result_df['Commission Rate'] * 100).round(2)
    
    return result_df

def calculate_rep_summary(commission_df, rep_name=None):
    """
    Calculate summary statistics for a rep
    """
    if commission_df.empty:
        return {
            'total_transactions': 0,
            'total_subtotal': 0,
            'total_commission': 0,
            'total_override': 0,
            'avg_commission': 0
        }
    
    # Filter to specific rep if provided
    if rep_name and not is_admin(rep_name):
        df = commission_df[commission_df['Final Sales Rep'] == rep_name].copy()
    else:
        df = commission_df.copy()
    
    summary = {
        'total_transactions': len(df),
        'total_subtotal': df['Subtotal'].sum(),
        'total_commission': df['Commission Amount'].sum(),
        'total_override': df['Brad Override'].sum() if 'Brad Override' in df.columns else 0,
        'avg_commission': df['Commission Amount'].mean() if len(df) > 0 else 0
    }
    
    return summary

# ==========================================
# STREAMLIT DISPLAY FUNCTIONS
# ==========================================

def display_password_gate():
    """
    Display password entry form
    Returns (is_authenticated, rep_name)
    """
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 padding: 20px; border-radius: 10px; color: white; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0;'>🔒 Commission Portal</h2>
        <p style='margin: 5px 0 0 0;'>Enter your credentials to view commission data</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        rep_name = st.selectbox(
            "Select Your Name:",
            options=[""] + list(REP_PASSWORDS.keys()),
            key="commission_rep_select"
        )
        
        if rep_name:
            password = st.text_input(
                "Enter Password:",
                type="password",
                key="commission_password"
            )
            
            col_a, col_b, col_c = st.columns([1, 1, 1])
            with col_b:
                if st.button("🔓 Access Commission Data", use_container_width=True):
                    if verify_password(rep_name, password):
                        st.session_state.commission_authenticated = True
                        st.session_state.commission_rep = rep_name
                        st.rerun()
                    else:
                        st.error("❌ Invalid password. Please try again.")
    
    return False, None

def display_commission_dashboard(invoices_df, sales_orders_df, rep_name):
    """
    Display the main commission dashboard for authenticated user
    """
    # Header with logout
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                     padding: 15px; border-radius: 10px; color: white;'>
            <h2 style='color: white; margin: 0;'>💰 Commission Dashboard - {rep_name}</h2>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.commission_authenticated = False
            st.session_state.commission_rep = None
            st.rerun()
    
    st.markdown("---")
    
    # Process commission data
    with st.spinner("Calculating commissions..."):
        commission_df = process_commission_data(invoices_df, sales_orders_df)
    
    if commission_df.empty:
        st.warning("⚠️ No commission data available for the selected period.")
        return
    
    # Calculate summary for this rep
    summary = calculate_rep_summary(commission_df, rep_name)
    
    # Display summary metrics
    st.markdown("### 📊 Commission Summary")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "Total Transactions",
            f"{summary['total_transactions']:,}"
        )
    
    with col2:
        st.metric(
            "Total Sales Volume",
            f"${summary['total_subtotal']:,.2f}"
        )
    
    with col3:
        st.metric(
            "Total Commission",
            f"${summary['total_commission']:,.2f}"
        )
    
    with col4:
        if summary['total_override'] > 0:
            st.metric(
                "Override Commission",
                f"${summary['total_override']:,.2f}"
            )
        else:
            st.metric(
                "Avg Commission/Deal",
                f"${summary['avg_commission']:,.2f}"
            )
    
    st.markdown("---")
    
    # Filter commission data for display
    if is_admin(rep_name):
        display_df = commission_df.copy()
        st.info("👑 Admin View: Showing commission data for all reps")
    else:
        display_df = commission_df[commission_df['Final Sales Rep'] == rep_name].copy()
    
    # Date range filter
    st.markdown("### 📅 Filter by Date Range")
    col1, col2 = st.columns(2)
    
    with col1:
        if 'Transaction Date' in display_df.columns:
            display_df['Transaction Date'] = pd.to_datetime(display_df['Transaction Date'], errors='coerce')
            min_date = display_df['Transaction Date'].min()
            if pd.isna(min_date):
                min_date = datetime.now() - pd.Timedelta(days=365)
            start_date = st.date_input("Start Date", value=min_date)
        else:
            start_date = st.date_input("Start Date", value=datetime.now() - pd.Timedelta(days=365))
    
    with col2:
        if 'Transaction Date' in display_df.columns:
            max_date = display_df['Transaction Date'].max()
            if pd.isna(max_date):
                max_date = datetime.now()
            end_date = st.date_input("End Date", value=max_date)
        else:
            end_date = st.date_input("End Date", value=datetime.now())
    
    # Apply date filter
    if 'Transaction Date' in display_df.columns:
        display_df = display_df[
            (display_df['Transaction Date'] >= pd.Timestamp(start_date)) &
            (display_df['Transaction Date'] <= pd.Timestamp(end_date))
        ]
    
    st.markdown("---")
    
    # Detailed table
    st.markdown("### 📋 Commission Detail")
    st.caption(f"Showing {len(display_df)} transactions")
    
    # Format currency columns
    currency_columns = ['Subtotal', 'Commission Amount', 'Brad Override', 'Amount Paid', 'Amount Due']
    for col in currency_columns:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(lambda x: f"${x:,.2f}" if not pd.isna(x) else "$0.00")
    
    # Format percentage
    if 'Commission Rate %' in display_df.columns:
        display_df['Commission Rate %'] = display_df['Commission Rate %'].apply(
            lambda x: f"{x:.2f}%" if not pd.isna(x) else "0.00%"
        )
    
    # Display dataframe
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )
    
    # Download button
    csv = display_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Commission Data (CSV)",
        data=csv,
        file_name=f"commission_data_{rep_name}_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
    
    # Show commission breakdown by pipeline
    if is_admin(rep_name):
        st.markdown("---")
        st.markdown("### 📊 Commission Breakdown by Rep & Pipeline")
        
        # Convert back to numeric for grouping
        summary_df = commission_df.copy()
        summary_df['Commission Amount Numeric'] = pd.to_numeric(
            summary_df['Commission Amount'], errors='coerce'
        )
        summary_df['Subtotal Numeric'] = pd.to_numeric(
            summary_df['Subtotal'], errors='coerce'
        )
        
        pivot_df = summary_df.groupby(['Final Sales Rep', 'Pipeline']).agg({
            'Document Number': 'count',
            'Subtotal Numeric': 'sum',
            'Commission Amount Numeric': 'sum'
        }).reset_index()
        
        pivot_df.columns = ['Sales Rep', 'Pipeline', 'Transaction Count', 'Total Sales', 'Total Commission']
        
        # Format currency
        pivot_df['Total Sales'] = pivot_df['Total Sales'].apply(lambda x: f"${x:,.2f}")
        pivot_df['Total Commission'] = pivot_df['Total Commission'].apply(lambda x: f"${x:,.2f}")
        
        st.dataframe(
            pivot_df,
            use_container_width=True,
            hide_index=True
        )

def display_commission_section(invoices_df, sales_orders_df):
    """
    Main entry point for commission section
    Handles authentication and display
    """
    # Check if user is authenticated
    if not st.session_state.get('commission_authenticated', False):
        is_auth, rep_name = display_password_gate()
    else:
        rep_name = st.session_state.get('commission_rep')
        display_commission_dashboard(invoices_df, sales_orders_df, rep_name)

# ==========================================
# UTILITY FUNCTIONS FOR TESTING
# ==========================================

def generate_password_hash(password):
    """
    Utility function to generate password hash
    Use this to create new password hashes for REP_PASSWORDS
    """
    return hashlib.sha256(password.encode()).hexdigest()

if __name__ == "__main__":
    # Example: Generate password hash
    print("Password Hash Generator")
    print("=======================")
    test_password = "password"
    print(f"Hash for '{test_password}':")
    print(generate_password_hash(test_password))
