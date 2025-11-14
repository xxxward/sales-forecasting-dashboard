"""
Q4 2025 Shipping Planning Tool - CLEAN VERSION
Simple checkbox selection interface like the original
"""

import streamlit as st
import pandas as pd
import numpy as np
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "v3_clean"

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing Google Cloud credentials")
            return pd.DataFrame()
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame()
        
        # Handle mismatched column counts
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return pd.DataFrame()

def load_all_data():
    """Load and process all data - EXACT PATTERN FROM SALES DASHBOARD"""
    
    # Load deals
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
    # Load dashboard info
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoices
    invoices_df = load_google_sheets_data("NS Invoices", "A:U", version=CACHE_VERSION)
    
    # Load sales orders
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AD", version=CACHE_VERSION)
    
    # Process deals
    if not deals_df.empty and len(deals_df.columns) >= 6:
        col_names = deals_df.columns.tolist()
        rename_dict = {}
        
        for col in col_names:
            if col == 'Record ID':
                rename_dict[col] = 'Record ID'
            elif col == 'Deal Name':
                rename_dict[col] = 'Deal Name'
            elif col == 'Close Date':
                rename_dict[col] = 'Close Date'
            elif col == 'Amount':
                rename_dict[col] = 'Amount'
            elif col == 'Close Status':
                rename_dict[col] = 'Status'
            elif col == 'Q1 2026 Spillover':
                rename_dict[col] = 'Q1 2026 Spillover'
            elif col == 'Account Name':
                rename_dict[col] = 'Account Name'
        
        deals_df = deals_df.rename(columns=rename_dict)
        
        # Clean Amount
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        if 'Amount' in deals_df.columns:
            deals_df['Amount'] = deals_df['Amount'].apply(clean_numeric)
        
        # Convert Close Date
        if 'Close Date' in deals_df.columns:
            deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2026-01-01')
        
        if 'Close Date' in deals_df.columns:
            deals_df = deals_df[(deals_df['Close Date'] >= q4_start) & (deals_df['Close Date'] < q4_end)]
        
        # Add Counts_In_Q4 flag
        if 'Q1 2026 Spillover' in deals_df.columns:
            deals_df['Counts_In_Q4'] = deals_df['Q1 2026 Spillover'] != 'Q1 2026'
        else:
            deals_df['Counts_In_Q4'] = True
    
    # Process dashboard info
    if not dashboard_df.empty and len(dashboard_df.columns) >= 3:
        dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
        dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
        
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        dashboard_df['Quota'] = dashboard_df['Quota'].apply(clean_numeric)
        dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(clean_numeric)
    
    # Process invoices - EXACT PATTERN FROM SALES DASHBOARD
    if not invoices_df.empty and len(invoices_df.columns) >= 15:
        rename_dict = {
            invoices_df.columns[0]: 'Invoice Number',
            invoices_df.columns[1]: 'Status',
            invoices_df.columns[2]: 'Date',
            invoices_df.columns[6]: 'Customer',
            invoices_df.columns[10]: 'Amount',
            invoices_df.columns[14]: 'Sales Rep'
        }
        
        if len(invoices_df.columns) > 19:
            rename_dict[invoices_df.columns[19]] = 'Corrected Customer Name'
        if len(invoices_df.columns) > 20:
            rename_dict[invoices_df.columns[20]] = 'Rep Master'
        
        invoices_df = invoices_df.rename(columns=rename_dict)
        
        # Apply Rep Master override - EXACT PATTERN FROM SALES DASHBOARD
        if 'Rep Master' in invoices_df.columns:
            invoices_df['Rep Master'] = invoices_df['Rep Master'].astype(str).str.strip()
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            mask = invoices_df['Rep Master'].isin(invalid_values)
            invoices_df.loc[~mask, 'Sales Rep'] = invoices_df.loc[~mask, 'Rep Master']
            invoices_df = invoices_df.drop(columns=['Rep Master'])
        
        # Apply customer name correction - EXACT PATTERN FROM SALES DASHBOARD
        if 'Corrected Customer Name' in invoices_df.columns:
            invoices_df['Corrected Customer Name'] = invoices_df['Corrected Customer Name'].astype(str).str.strip()
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            mask = invoices_df['Corrected Customer Name'].isin(invalid_values)
            invoices_df.loc[~mask, 'Customer'] = invoices_df.loc[~mask, 'Corrected Customer Name']
            invoices_df = invoices_df.drop(columns=['Corrected Customer Name'])
        
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        invoices_df['Amount'] = invoices_df['Amount'].apply(clean_numeric)
        invoices_df['Date'] = pd.to_datetime(invoices_df['Date'], errors='coerce')
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2025-12-31')
        
        invoices_df = invoices_df[
            (invoices_df['Date'] >= q4_start) & 
            (invoices_df['Date'] <= q4_end)
        ]
        
        # Remove invalid reps
        invoices_df['Sales Rep'] = invoices_df['Sales Rep'].astype(str).str.strip()
        
        invoices_df = invoices_df[
            (invoices_df['Sales Rep'].notna()) & 
            (invoices_df['Sales Rep'] != '') &
            (invoices_df['Sales Rep'].str.lower() != 'nan') &
            (invoices_df['Sales Rep'].str.lower() != 'house')
        ]
        
        # Deduplicate
        if 'Invoice Number' in invoices_df.columns:
            invoices_df = invoices_df.drop_duplicates(subset=['Invoice Number'], keep='first')
    
    # Process sales orders - EXACT PATTERN FROM SALES DASHBOARD
    if not sales_orders_df.empty and len(sales_orders_df.columns) >= 10:
        # Map columns by index
        rename_dict = {
            sales_orders_df.columns[0]: 'Internal ID',
            sales_orders_df.columns[1]: 'Document Number',
            sales_orders_df.columns[2]: 'Status',
            sales_orders_df.columns[3]: 'Order Start Date',
            sales_orders_df.columns[7]: 'Customer',
            sales_orders_df.columns[10]: 'Amount',
            sales_orders_df.columns[19]: 'Sales Rep'
        }
        
        # Add date columns
        if len(sales_orders_df.columns) > 23:
            rename_dict[sales_orders_df.columns[23]] = 'Customer Promise Date'
        if len(sales_orders_df.columns) > 24:
            rename_dict[sales_orders_df.columns[24]] = 'Projected Date'
        if len(sales_orders_df.columns) > 25:
            rename_dict[sales_orders_df.columns[25]] = 'Pending Approval Date'
        if len(sales_orders_df.columns) > 29:
            rename_dict[sales_orders_df.columns[29]] = 'Rep Master'
        
        sales_orders_df = sales_orders_df.rename(columns=rename_dict)
        
        # Apply Rep Master override - EXACT PATTERN FROM SALES DASHBOARD (DIRECT REPLACEMENT)
        if 'Rep Master' in sales_orders_df.columns:
            sales_orders_df['Sales Rep'] = sales_orders_df['Rep Master']
            sales_orders_df = sales_orders_df.drop(columns=['Rep Master'])
        
        # CRITICAL: Remove any duplicate columns
        if sales_orders_df.columns.duplicated().any():
            sales_orders_df = sales_orders_df.loc[:, ~sales_orders_df.columns.duplicated()]
        
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric)
        
        # Convert dates
        for date_col in ['Order Start Date', 'Customer Promise Date', 'Projected Date', 'Pending Approval Date']:
            if date_col in sales_orders_df.columns:
                sales_orders_df[date_col] = pd.to_datetime(sales_orders_df[date_col], errors='coerce')
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2025-12-31')
        
        if 'Order Start Date' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Order Start Date'] >= q4_start) & 
                (sales_orders_df['Order Start Date'] <= q4_end)
            ]
        
        # Calculate business days for aging
        if 'Pending Approval Date' in sales_orders_df.columns:
            today = pd.Timestamp.now().normalize()
            
            def calculate_business_days(date_val):
                if pd.isna(date_val):
                    return 0
                date_val = pd.to_datetime(date_val).normalize()
                return max(0, len(pd.bdate_range(date_val, today)))
            
            sales_orders_df['Age_Business_Days'] = sales_orders_df['Pending Approval Date'].apply(calculate_business_days)
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def calculate_team_metrics(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Calculate team-level metrics"""
    
    metrics = {
        'orders': 0,
        'pending_fulfillment': 0,
        'pending_approval': 0,
        'expect_commit': 0,
        'pending_fulfillment_no_date': 0,
        'pending_approval_no_date': 0,
        'pending_approval_old': 0,
        'q1_spillover_expect_commit': 0,
        'q1_spillover_best_opp': 0
    }
    
    # Invoiced & Shipped
    if invoices_df is not None and not invoices_df.empty:
        metrics['orders'] = invoices_df['Amount'].sum()
    
    # Sales Orders
    if sales_orders_df is not None and not sales_orders_df.empty:
        # Pending Fulfillment
        pf_orders = sales_orders_df[sales_orders_df['Status'] == 'Pending Fulfillment'].copy()
        if not pf_orders.empty:
            pf_orders['Has_Date'] = (
                pf_orders['Customer Promise Date'].notna() | 
                pf_orders['Projected Date'].notna()
            )
            
            metrics['pending_fulfillment'] = pf_orders[pf_orders['Has_Date'] == True]['Amount'].sum()
            metrics['pending_fulfillment_no_date'] = pf_orders[pf_orders['Has_Date'] == False]['Amount'].sum()
        
        # Pending Approval
        pa_orders = sales_orders_df[sales_orders_df['Status'] == 'Pending Approval'].copy()
        if not pa_orders.empty:
            pa_with_date = pa_orders[pa_orders['Pending Approval Date'].notna()]
            pa_no_date = pa_orders[pa_orders['Pending Approval Date'].isna()]
            
            # Old PA
            if 'Age_Business_Days' in pa_orders.columns:
                pa_old = pa_orders[pa_orders['Age_Business_Days'] >= 10]
                metrics['pending_approval_old'] = pa_old['Amount'].sum()
                
                # Young PA with date
                pa_young_with_date = pa_with_date[pa_with_date['Age_Business_Days'] < 10]
                metrics['pending_approval'] = pa_young_with_date['Amount'].sum()
            else:
                metrics['pending_approval'] = pa_with_date['Amount'].sum()
            
            metrics['pending_approval_no_date'] = pa_no_date['Amount'].sum()
    
    # HubSpot Deals
    if deals_df is not None and not deals_df.empty and 'Status' in deals_df.columns:
        deals_df['Amount_Numeric'] = pd.to_numeric(deals_df['Amount'], errors='coerce')
        
        # Q4 deals
        q4_deals = deals_df[deals_df.get('Counts_In_Q4', True) == True]
        
        expect_deals = q4_deals[q4_deals['Status'] == 'Expect']
        commit_deals = q4_deals[q4_deals['Status'] == 'Commit']
        
        metrics['expect_commit'] = expect_deals['Amount_Numeric'].sum() + commit_deals['Amount_Numeric'].sum()
        
        # Q1 Spillover
        if 'Q1 2026 Spillover' in deals_df.columns:
            q1_deals = deals_df[deals_df['Q1 2026 Spillover'] == 'Q1 2026']
            
            q1_expect_commit = q1_deals[q1_deals['Status'].isin(['Expect', 'Commit'])]
            q1_best_opp = q1_deals[q1_deals['Status'].isin(['Best Case', 'Opportunity'])]
            
            metrics['q1_spillover_expect_commit'] = q1_expect_commit['Amount_Numeric'].sum()
            metrics['q1_spillover_best_opp'] = q1_best_opp['Amount_Numeric'].sum()
    
    return metrics

def build_shipping_plan_section(metrics, quota, deals_df=None, invoices_df=None, sales_orders_df=None):
    """Simple checkbox selection - like the original"""
    
    st.markdown("### 📦 Build Your Shipping Plan")
    st.caption("Select the components you want to include in your custom forecast calculation")
    
    # Create columns for checkboxes
    col1, col2, col3 = st.columns(3)
    
    # Available data sources
    sources = {
        'Invoiced & Shipped': metrics.get('orders', 0),
        'Pending Fulfillment (with date)': metrics.get('pending_fulfillment', 0),
        'Pending Approval (with date)': metrics.get('pending_approval', 0),
        'HubSpot Expect/Commit': metrics.get('expect_commit', 0),
        'Pending Fulfillment (without date)': metrics.get('pending_fulfillment_no_date', 0),
        'Pending Approval (without date)': metrics.get('pending_approval_no_date', 0),
        'Pending Approval (>2 weeks old)': metrics.get('pending_approval_old', 0),
        'Q1 Spillover - Expect/Commit': metrics.get('q1_spillover_expect_commit', 0),
        'Q1 Spillover - Best Case': metrics.get('q1_spillover_best_opp', 0)
    }
    
    # Create checkboxes
    selected_sources = {}
    source_list = list(sources.keys())
    
    with col1:
        for source in source_list[0:3]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    with col2:
        for source in source_list[3:6]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    with col3:
        for source in source_list[6:]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    # Calculate totals
    custom_forecast = sum(sources[source] for source, selected in selected_sources.items() if selected)
    
    # Only count invoiced_shipped if it was selected
    if selected_sources.get('Invoiced & Shipped', False):
        invoiced_shipped = sources.get('Invoiced & Shipped', 0)
    else:
        invoiced_shipped = 0
    
    to_ship = custom_forecast - invoiced_shipped
    gap_to_quota = quota - custom_forecast
    
    # Calculate working days remaining
    today = datetime.now()
    q4_end = datetime(2025, 12, 31)
    remaining_calendar_days = (q4_end - today).days
    working_days_remaining = max(1, int(remaining_calendar_days * (5/7)))
    per_day_needed = to_ship / working_days_remaining if working_days_remaining > 0 else 0
    
    # Display results
    st.markdown("---")
    st.markdown("#### 📦 Shipping Plan Summary")
    
    result_col1, result_col2, result_col3, result_col4 = st.columns(4)
    
    with result_col1:
        st.metric("Q4 Quota", f"${quota:,.0f}")
    
    with result_col2:
        shipped_pct = (invoiced_shipped / quota * 100) if quota > 0 else 0
        st.metric("✅ Shipped", f"${invoiced_shipped:,.0f}", 
                 delta=f"{shipped_pct:.1f}%",
                 delta_color="normal")
    
    with result_col3:
        st.metric("📦 To Ship", f"${to_ship:,.0f}",
                 delta=f"Per day: ${per_day_needed:,.0f}",
                 delta_color="off",
                 help=f"Based on ~{working_days_remaining} working days remaining in Q4")
    
    with result_col4:
        gap_pct = (gap_to_quota / quota * 100) if quota > 0 else 0
        st.metric("Gap to Quota", f"${gap_to_quota:,.0f}",
                 delta=f"{gap_pct:.1f}% short" if gap_to_quota > 0 else f"{abs(gap_pct):.1f}% over",
                 delta_color="inverse")

def main():
    st.markdown("""
    <div style='text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 color: white; border-radius: 10px; margin-bottom: 20px;'>
        <h3>📦 Q4 2025 Shipping Planning</h3>
        <p style='font-size: 14px; margin: 0;'>Simple Checkbox Selection</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Load data
    with st.spinner("🔄 Loading data..."):
        deals_df, dashboard_df, invoices_df, sales_orders_df = load_all_data()
    
    if deals_df.empty and dashboard_df.empty:
        st.error("❌ Unable to load data")
        return
    
    # Calculate metrics
    metrics = calculate_team_metrics(deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    # Get total quota
    quota = dashboard_df['Quota'].sum() if not dashboard_df.empty else 5_021_440
    
    # Call the shipping plan function
    build_shipping_plan_section(
        metrics=metrics,
        quota=quota,
        deals_df=deals_df,
        invoices_df=invoices_df,
        sales_orders_df=sales_orders_df
    )

if __name__ == "__main__":
    main()
