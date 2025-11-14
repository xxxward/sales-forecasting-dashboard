"""
Q4 2025 Shipping Planning Tool - SIMPLIFIED VERSION
Just loads data and uses the exact Build Your Own Forecast interface
"""

import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "v1_simple"

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets - EXACT COPY from main dashboard"""
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
    """Load data - EXACT COPY from main dashboard load_all_data function"""
    
    # Load deals
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
    # Load dashboard info
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoices
    invoices_df = load_google_sheets_data("NS Invoices", "A:U", version=CACHE_VERSION)
    
    # Load sales orders
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AD", version=CACHE_VERSION)
    
    # Process deals - SIMPLIFIED VERSION (just the essentials)
    if not deals_df.empty and len(deals_df.columns) >= 6:
        col_names = deals_df.columns.tolist()
        rename_dict = {}
        
        for col in col_names:
            if col == 'Record ID':
                rename_dict[col] = 'Record ID'
            elif col == 'Deal Name':
                rename_dict[col] = 'Deal Name'
            elif col == 'Deal Stage':
                rename_dict[col] = 'Deal Stage'
            elif col == 'Close Date':
                rename_dict[col] = 'Close Date'
            elif 'Deal Owner First Name' in col and 'Deal Owner Last Name' in col:
                rename_dict[col] = 'Deal Owner'
            elif col == 'Amount':
                rename_dict[col] = 'Amount'
            elif col == 'Close Status':
                rename_dict[col] = 'Status'
            elif col == 'Pipeline':
                rename_dict[col] = 'Pipeline'
            elif col == 'Deal Type':
                rename_dict[col] = 'Product Type'
            elif col == 'Q1 2026 Spillover':
                rename_dict[col] = 'Q1 2026 Spillover'
        
        deals_df = deals_df.rename(columns=rename_dict)
        
        # Create Deal Owner if needed
        if 'Deal Owner' not in deals_df.columns:
            if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
        
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
        
        # Add Counts_In_Q4 flag (simplified - just check if Q1 2026 Spillover column says "Q1 2026")
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
    
    # Process invoices
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
        
        # Apply Rep Master override
        if 'Rep Master' in invoices_df.columns:
            invoices_df['Rep Master'] = invoices_df['Rep Master'].astype(str).str.strip()
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            mask = invoices_df['Rep Master'].isin(invalid_values)
            invoices_df.loc[~mask, 'Sales Rep'] = invoices_df.loc[~mask, 'Rep Master']
            invoices_df = invoices_df.drop(columns=['Rep Master'])
        
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
        invoices_df = invoices_df[(invoices_df['Date'] >= q4_start) & (invoices_df['Date'] <= q4_end)]
        
        # Clean Sales Rep
        invoices_df['Sales Rep'] = invoices_df['Sales Rep'].astype(str).str.strip()
        invoices_df = invoices_df[
            (invoices_df['Sales Rep'].notna()) & 
            (invoices_df['Sales Rep'] != '') &
            (invoices_df['Sales Rep'].str.lower() != 'nan') &
            (invoices_df['Sales Rep'].str.lower() != 'house')
        ]
        
        # Remove duplicates
        if 'Invoice Number' in invoices_df.columns:
            invoices_df = invoices_df.drop_duplicates(subset=['Invoice Number'], keep='first')
        
        # Calculate invoice totals by rep
        invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
        invoice_totals.columns = ['Rep Name', 'Invoice Total']
        
        dashboard_df['Rep Name'] = dashboard_df['Rep Name'].str.strip()
        dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
        dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
        dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
        dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
    
    # Process sales orders  
    if not sales_orders_df.empty:
        col_names = sales_orders_df.columns.tolist()
        rename_dict = {}
        
        # Find standard columns
        for idx, col in enumerate(col_names):
            col_lower = str(col).lower()
            if 'status' in col_lower and 'Status' not in rename_dict.values():
                rename_dict[col] = 'Status'
            elif ('amount' in col_lower or 'total' in col_lower) and 'Amount' not in rename_dict.values():
                rename_dict[col] = 'Amount'
            elif ('sales rep' in col_lower or 'salesrep' in col_lower) and 'Sales Rep' not in rename_dict.values():
                rename_dict[col] = 'Sales Rep'
            elif 'customer' in col_lower and 'customer promise' not in col_lower and 'Customer' not in rename_dict.values():
                rename_dict[col] = 'Customer'
            elif ('doc' in col_lower or 'document' in col_lower) and 'Document Number' not in rename_dict.values():
                rename_dict[col] = 'Document Number'
        
        # Map specific columns by position
        if len(col_names) > 8:
            rename_dict[col_names[8]] = 'Order Start Date'
        if len(col_names) > 11:
            rename_dict[col_names[11]] = 'Customer Promise Date'
        if len(col_names) > 12:
            rename_dict[col_names[12]] = 'Projected Date'
        
        # Map Rep Master
        if len(col_names) > 29:
            rename_dict[col_names[29]] = 'Rep Master'
        
        sales_orders_df = sales_orders_df.rename(columns=rename_dict)
        
        # Apply Rep Master override
        if 'Rep Master' in sales_orders_df.columns:
            sales_orders_df['Sales Rep'] = sales_orders_df['Rep Master']
            sales_orders_df = sales_orders_df.drop(columns=['Rep Master'])
        
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        if 'Amount' in sales_orders_df.columns:
            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric)
        
        if 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].astype(str).str.strip()
        
        if 'Status' in sales_orders_df.columns:
            sales_orders_df['Status'] = sales_orders_df['Status'].astype(str).str.strip()
        
        # Convert date columns
        date_columns = ['Order Start Date', 'Customer Promise Date', 'Projected Date']
        for col in date_columns:
            if col in sales_orders_df.columns:
                sales_orders_df[col] = pd.to_datetime(sales_orders_df[col], errors='coerce')
        
        # Filter to Pending statuses
        if 'Status' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
            ]
        
        # Calculate age
        if 'Order Start Date' in sales_orders_df.columns:
            today = pd.Timestamp.now()
            
            def business_days_between(start_date, end_date):
                if pd.isna(start_date):
                    return 0
                days = pd.bdate_range(start=start_date, end=end_date).size - 1
                return max(0, days)
            
            sales_orders_df['Age_Business_Days'] = sales_orders_df['Order Start Date'].apply(
                lambda x: business_days_between(x, today)
            )
        
        # Remove invalid rows
        if 'Amount' in sales_orders_df.columns and 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & 
                (sales_orders_df['Sales Rep'].notna()) & 
                (sales_orders_df['Sales Rep'] != '') &
                (sales_orders_df['Sales Rep'] != 'nan') &
                (~sales_orders_df['Sales Rep'].str.lower().isin(['house']))
            ]
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def calculate_team_metrics(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Calculate metrics - SIMPLIFIED VERSION"""
    
    metrics = {
        'orders': 0,
        'pending_fulfillment': 0,
        'pending_fulfillment_no_date': 0,
        'pending_approval': 0,
        'pending_approval_no_date': 0,
        'pending_approval_old': 0,
        'expect_commit': 0,
        'q1_spillover_expect_commit': 0,
        'q1_spillover_best_opp': 0
    }
    
    # Get total invoiced from dashboard
    if not dashboard_df.empty and 'NetSuite Orders' in dashboard_df.columns:
        metrics['orders'] = dashboard_df['NetSuite Orders'].sum()
    
    # Calculate SO metrics
    if not sales_orders_df.empty:
        # Create Estimated Ship Date column
        if 'Customer Promise Date' in sales_orders_df.columns and 'Projected Date' in sales_orders_df.columns:
            sales_orders_df['Estimated Ship Date'] = sales_orders_df['Customer Promise Date'].fillna(sales_orders_df['Projected Date'])
        elif 'Customer Promise Date' in sales_orders_df.columns:
            sales_orders_df['Estimated Ship Date'] = sales_orders_df['Customer Promise Date']
        elif 'Projected Date' in sales_orders_df.columns:
            sales_orders_df['Estimated Ship Date'] = sales_orders_df['Projected Date']
        
        # Pending Fulfillment
        pf_df = sales_orders_df[sales_orders_df.get('Status', '') == 'Pending Fulfillment']
        metrics['pending_fulfillment'] = pf_df[pf_df['Estimated Ship Date'].notna()]['Amount'].sum()
        metrics['pending_fulfillment_no_date'] = pf_df[pf_df['Estimated Ship Date'].isna()]['Amount'].sum()
        
        # Pending Approval
        pa_df = sales_orders_df[sales_orders_df.get('Status', '') == 'Pending Approval']
        metrics['pending_approval'] = pa_df[pa_df['Estimated Ship Date'].notna()]['Amount'].sum()
        metrics['pending_approval_no_date'] = pa_df[pa_df['Estimated Ship Date'].isna()]['Amount'].sum()
        
        # Pending Approval > 2 weeks old
        if 'Age_Business_Days' in pa_df.columns:
            old_pa = pa_df[pa_df['Age_Business_Days'] >= 10]
            metrics['pending_approval_old'] = old_pa['Amount'].sum()
    
    # Calculate HubSpot metrics
    if not deals_df.empty and 'Status' in deals_df.columns:
        deals_df['Amount_Numeric'] = pd.to_numeric(deals_df['Amount'], errors='coerce')
        
        # Q4 deals
        q4_deals = deals_df[deals_df.get('Counts_In_Q4', True) == True]
        expect_commit = q4_deals[q4_deals['Status'].isin(['Expect', 'Commit'])]['Amount_Numeric'].sum()
        metrics['expect_commit'] = expect_commit
        
        # Q1 Spillover
        if 'Q1 2026 Spillover' in deals_df.columns:
            q1_deals = deals_df[deals_df['Q1 2026 Spillover'] == 'Q1 2026']
            metrics['q1_spillover_expect_commit'] = q1_deals[q1_deals['Status'].isin(['Expect', 'Commit'])]['Amount_Numeric'].sum()
            metrics['q1_spillover_best_opp'] = q1_deals[q1_deals['Status'].isin(['Best Case', 'Opportunity'])]['Amount_Numeric'].sum()
    
    return metrics

# Import the exact Build Your Own Forecast function from main dashboard
def build_your_own_forecast_section(metrics, quota, rep_name=None, deals_df=None, invoices_df=None, sales_orders_df=None):
    """
    Interactive section where users can select which data sources to include in their forecast
    """
    st.markdown("### 🎯 Build Your Own Forecast")
    st.caption("Select the components you want to include in your custom forecast calculation")
    
    # Initialize session state for individual selections if not exists
    if 'selected_individual_items' not in st.session_state:
        st.session_state.selected_individual_items = {}
    
    # Create columns for checkboxes
    col1, col2, col3 = st.columns(3)
    
    # Available data sources with their values
    sources = {
        'Invoiced & Shipped': metrics.get('orders', 0),
        'Pending Fulfillment (with date)': metrics.get('pending_fulfillment', 0),
        'Pending Approval (with date)': metrics.get('pending_approval', 0),
        'HubSpot Expect': metrics.get('expect_commit', 0) if 'expect_commit' in metrics else 0,
        'HubSpot Commit': 0,  # Will calculate separately
        'HubSpot Best Case': 0,  # Will calculate separately
        'HubSpot Opportunity': 0,  # Will calculate separately
        'Pending Fulfillment (without date)': metrics.get('pending_fulfillment_no_date', 0),
        'Pending Approval (without date)': metrics.get('pending_approval_no_date', 0),
        'Pending Approval (>2 weeks old)': metrics.get('pending_approval_old', 0),
        'Q1 Spillover - Expect/Commit': metrics.get('q1_spillover_expect_commit', 0),
        'Q1 Spillover - Best Case': metrics.get('q1_spillover_best_opp', 0)
    }
    
    # Track which categories allow individual selection
    individual_select_categories = [
        'Invoiced & Shipped',  # NEW - allow drilling into invoices
        'Pending Fulfillment (with date)',  # NEW - allow drilling into PF with dates
        'Pending Approval (with date)',  # NEW - allow drilling into PA with dates
        'HubSpot Expect', 'HubSpot Commit', 'HubSpot Best Case', 'HubSpot Opportunity',
        'Pending Fulfillment (without date)', 'Pending Approval (without date)', 
        'Pending Approval (>2 weeks old)', 'Q1 Spillover - Expect/Commit', 'Q1 Spillover - Best Case'
    ]
    
    # Calculate individual HubSpot categories
    if deals_df is not None and not deals_df.empty:
        if rep_name:
            rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
        else:
            rep_deals = deals_df.copy()
        
        if not rep_deals.empty and 'Status' in rep_deals.columns:
            rep_deals['Amount_Numeric'] = pd.to_numeric(rep_deals['Amount'], errors='coerce')
            
            # Filter for Q4 only
            q4_deals = rep_deals[rep_deals.get('Counts_In_Q4', True) == True]
            
            sources['HubSpot Expect'] = q4_deals[q4_deals['Status'] == 'Expect']['Amount_Numeric'].sum()
            sources['HubSpot Commit'] = q4_deals[q4_deals['Status'] == 'Commit']['Amount_Numeric'].sum()
            sources['HubSpot Best Case'] = q4_deals[q4_deals['Status'] == 'Best Case']['Amount_Numeric'].sum()
            sources['HubSpot Opportunity'] = q4_deals[q4_deals['Status'] == 'Opportunity']['Amount_Numeric'].sum()
    
    # Create checkboxes in columns with individual selection option
    selected_sources = {}
    individual_selection_mode = {}
    source_list = list(sources.keys())
    
    with col1:
        for source in source_list[0:4]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            # Add "Select Individual" option for applicable categories
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    with col2:
        for source in source_list[4:8]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    with col3:
        for source in source_list[8:]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    # Show individual selection interfaces for each category
    individual_selections = {}
    
    for category, is_individual in individual_selection_mode.items():
        if is_individual:
            st.markdown(f"#### 🛒 Select Individual Items: {category}")
            
            # Get the relevant data for this category
            items_to_select = []
            
            # NEW: Invoiced & Shipped
            if 'Invoiced & Shipped' in category and invoices_df is not None:
                if rep_name and 'Sales Rep' in invoices_df.columns:
                    inv_data = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
                else:
                    inv_data = invoices_df.copy()
                
                items_to_select = inv_data.copy()
            
            # NEW: Pending Fulfillment WITH date
            elif 'Pending Fulfillment (with date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Fulfillment') &
                    (so_data['Customer Promise Date'].notna() | so_data['Projected Date'].notna())
                ].copy()
            
            # NEW: Pending Approval WITH date
            elif 'Pending Approval (with date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Approval') &
                    (so_data['Customer Promise Date'].notna() | so_data['Projected Date'].notna())
                ].copy()
            
            # Sales Orders categories
            elif 'Pending Fulfillment (without date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Fulfillment') &
                    (so_data['Customer Promise Date'].isna()) &
                    (so_data['Projected Date'].isna())
                ].copy()
                
            elif 'Pending Approval (without date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Approval') &
                    (so_data['Customer Promise Date'].isna()) &
                    (so_data['Projected Date'].isna())
                ].copy()
                
            elif 'Pending Approval (>2 weeks old)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                if 'Age_Business_Days' in so_data.columns:
                    items_to_select = so_data[
                        (so_data['Status'] == 'Pending Approval') &
                        (so_data['Age_Business_Days'] >= 10)
                    ].copy()
                    
            # HubSpot deals categories
            elif 'HubSpot' in category and deals_df is not None:
                if rep_name:
                    hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
                else:
                    hs_deals = deals_df.copy()
                
                if not hs_deals.empty and 'Status' in hs_deals.columns:
                    hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                    q4_deals = hs_deals[hs_deals.get('Counts_In_Q4', True) == True]
                    
                    if 'Expect' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Expect'].copy()
                    elif 'Commit' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Commit'].copy()
                    elif 'Best Case' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Best Case'].copy()
                    elif 'Opportunity' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Opportunity'].copy()
            
            # Q1 Spillover deals
            elif 'Q1 Spillover' in category and deals_df is not None:
                if rep_name:
                    hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
                else:
                    hs_deals = deals_df.copy()
                
                if not hs_deals.empty and 'Status' in hs_deals.columns:
                    hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                    
                    # Determine which status to filter by
                    if 'Expect/Commit' in category:
                        status_filter = ['Expect', 'Commit']
                    elif 'Best Case' in category:
                        status_filter = ['Best Case', 'Opportunity']
                    else:
                        status_filter = ['Expect', 'Commit']  # Default
                    
                    # Get Q1 spillover deals using the Q1 2026 Spillover column
                    # Must match EXACTLY the logic in calculate_rep_metrics
                    if 'Q1 2026 Spillover' in hs_deals.columns:
                        items_to_select = hs_deals[
                            (hs_deals['Q1 2026 Spillover'] == 'Q1 2026') &
                            (hs_deals['Status'].isin(status_filter))
                        ].copy()
                        
                        # Debug info
                        total_spillover = hs_deals[hs_deals['Q1 2026 Spillover'] == 'Q1 2026']
                        st.caption(f"🔍 Debug: Total Q1 spillover deals = {len(total_spillover)}, {'/'.join(status_filter)} only = {len(items_to_select)}")
                        st.caption(f"Total amount in Q1 spillover {'/'.join(status_filter)} = ${items_to_select['Amount_Numeric'].sum():,.0f}")
                    else:
                        # Fallback to old logic if column doesn't exist
                        items_to_select = hs_deals[
                            (hs_deals.get('Counts_In_Q4', True) == False) &
                            (hs_deals['Status'].isin(status_filter))
                        ].copy()
                        st.caption("⚠️ Using fallback logic - Q1 2026 Spillover column not found")
            
            # Display selection interface
            if not items_to_select.empty:
                st.caption(f"Found {len(items_to_select)} items - select the ones you want to include")
                
                selected_items = []
                
                # Create a more compact selection interface
                for idx, row in items_to_select.iterrows():
                    # Determine display info based on type
                    if 'Deal Name' in row:
                        item_id = row.get('Record ID', idx)
                        item_name = row.get('Deal Name', 'Unknown')
                        item_customer = row.get('Account Name', '')
                        item_amount = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                    else:
                        item_id = row.get('Document Number', idx)
                        item_name = f"SO #{item_id}"
                        item_customer = row.get('Customer', '')
                        item_amount = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                    
                    # Checkbox for each item
                    is_selected = st.checkbox(
                        f"{item_name} - {item_customer} - ${item_amount:,.0f}",
                        value=False,
                        key=f"{'team' if rep_name is None else rep_name}_{category}_{item_id}"
                    )
                    
                    if is_selected:
                        selected_items.append({
                            'id': item_id,
                            'amount': item_amount,
                            'row': row
                        })
                
                individual_selections[category] = selected_items
                st.caption(f"✓ Selected {len(selected_items)} of {len(items_to_select)} items")
            else:
                st.info(f"No items found in this category")
    
    # Calculate custom forecast
    custom_forecast = 0
    
    for source, selected in selected_sources.items():
        if selected:
            if individual_selection_mode.get(source, False):
                # Use individual selections
                if source in individual_selections:
                    custom_forecast += sum(item['amount'] for item in individual_selections[source])
            else:
                # Use full category amount
                custom_forecast += sources[source]
    
    gap_to_quota = quota - custom_forecast
    attainment_pct = (custom_forecast / quota * 100) if quota > 0 else 0
    
    # Calculate shipping metrics
    # Only count invoiced_shipped if it was actually selected
    if selected_sources.get('Invoiced & Shipped', False):
        if individual_selection_mode.get('Invoiced & Shipped', False):
            # Use individual selections if in individual mode
            invoiced_shipped = sum(item['amount'] for item in individual_selections.get('Invoiced & Shipped', []))
        else:
            # Use full category amount
            invoiced_shipped = sources.get('Invoiced & Shipped', 0)
    else:
        invoiced_shipped = 0
    
    to_ship = custom_forecast - invoiced_shipped
    
    # Calculate working days remaining in Q4 (approximate)
    today = datetime.now()
    q4_end = datetime(2025, 12, 31)
    remaining_calendar_days = (q4_end - today).days
    # Rough estimate: 5/7 of days are working days
    working_days_remaining = max(1, int(remaining_calendar_days * (5/7)))
    per_day_needed = to_ship / working_days_remaining if working_days_remaining > 0 else 0
    
    # Display results - OPS FOCUSED
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
    
    # Export functionality
    if any(selected_sources.values()):
        st.markdown("---")
        
        # Collect data for export with summary
        export_summary = []
        export_data = []
        
        # Build summary section
        export_summary.append({
            'Category': '=== FORECAST SUMMARY ===',
            'Amount': ''
        })
        export_summary.append({
            'Category': 'Quota',
            'Amount': f"${quota:,.0f}"
        })
        export_summary.append({
            'Category': 'Custom Forecast',
            'Amount': f"${custom_forecast:,.0f}"
        })
        export_summary.append({
            'Category': 'Gap to Quota',
            'Amount': f"${gap_to_quota:,.0f}"
        })
        export_summary.append({
            'Category': 'Attainment %',
            'Amount': f"{attainment_pct:.1f}%"
        })
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        export_summary.append({
            'Category': '=== SELECTED COMPONENTS ===',
            'Amount': ''
        })
        
        # Add each selected component total
        for source, selected in selected_sources.items():
            if selected:
                if individual_selection_mode.get(source, False) and source in individual_selections:
                    # Show individual selection count
                    item_count = len(individual_selections[source])
                    item_total = sum(item['amount'] for item in individual_selections[source])
                    export_summary.append({
                        'Category': f"{source} ({item_count} items selected)",
                        'Amount': f"${item_total:,.0f}"
                    })
                else:
                    export_summary.append({
                        'Category': source,
                        'Amount': f"${sources[source]:,.0f}"
                    })
        
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        export_summary.append({
            'Category': '=== DETAILED LINE ITEMS ===',
            'Amount': ''
        })
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        
        # Get invoices data (always bulk)
        if selected_sources.get('Invoiced & Shipped', False) and invoices_df is not None:
            if rep_name and 'Sales Rep' in invoices_df.columns:
                inv_data = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
            else:
                inv_data = invoices_df.copy()
            
            if not inv_data.empty:
                for _, row in inv_data.iterrows():
                    export_data.append({
                        'Type': 'Invoice',
                        'ID': row.get('Document Number', row.get('Invoice Number', '')),
                        'Name': '',
                        'Customer': row.get('Account Name', row.get('Customer', '')),
                        'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                        'Date': row.get('Date', row.get('Transaction Date', '')),
                        'Sales Rep': row.get('Sales Rep', '')
                    })
        
        # Get sales orders data - check individual vs bulk
        if sales_orders_df is not None:
            if rep_name and 'Sales Rep' in sales_orders_df.columns:
                so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
            else:
                so_data = sales_orders_df.copy()
            
            if not so_data.empty:
                # Pending Fulfillment with date (always bulk)
                if selected_sources.get('Pending Fulfillment (with date)', False):
                    pf_data = so_data[so_data['Status'] == 'Pending Fulfillment'].copy()
                    for _, row in pf_data.iterrows():
                        if pd.notna(row.get('Customer Promise Date')) or pd.notna(row.get('Projected Date')):
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Approval with date (always bulk)
                if selected_sources.get('Pending Approval (with date)', False):
                    pa_data = so_data[so_data['Status'] == 'Pending Approval'].copy()
                    for _, row in pa_data.iterrows():
                        if pd.notna(row.get('Customer Promise Date')) or pd.notna(row.get('Projected Date')):
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Fulfillment without date - check individual mode
                if selected_sources.get('Pending Fulfillment (without date)', False):
                    category = 'Pending Fulfillment (without date)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        pf_no_date = so_data[
                            (so_data['Status'] == 'Pending Fulfillment') &
                            (so_data['Customer Promise Date'].isna()) &
                            (so_data['Projected Date'].isna())
                        ].copy()
                        for _, row in pf_no_date.iterrows():
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Approval without date - check individual mode
                if selected_sources.get('Pending Approval (without date)', False):
                    category = 'Pending Approval (without date)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        pa_no_date = so_data[
                            (so_data['Status'] == 'Pending Approval') &
                            (so_data['Customer Promise Date'].isna()) &
                            (so_data['Projected Date'].isna())
                        ].copy()
                        for _, row in pa_no_date.iterrows():
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Old Pending Approval - check individual mode
                if selected_sources.get('Pending Approval (>2 weeks old)', False):
                    category = 'Pending Approval (>2 weeks old)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Old Pending Approval',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        if 'Age_Business_Days' in so_data.columns:
                            old_pa = so_data[
                                (so_data['Status'] == 'Pending Approval') &
                                (so_data['Age_Business_Days'] >= 10)
                            ].copy()
                            for _, row in old_pa.iterrows():
                                export_data.append({
                                    'Type': 'Sales Order - Old Pending Approval',
                                    'ID': row.get('Document Number', ''),
                                    'Name': '',
                                    'Customer': row.get('Customer', ''),
                                    'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                    'Date': row.get('Order Start Date', ''),
                                    'Sales Rep': row.get('Sales Rep', '')
                                })
        
        # Get HubSpot deals data - check individual vs bulk
        if deals_df is not None and not deals_df.empty:
            if rep_name:
                hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
            else:
                hs_deals = deals_df.copy()
            
            if not hs_deals.empty and 'Status' in hs_deals.columns:
                hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                
                # Filter for selected categories with individual selection support
                for status_name, checkbox_name in [
                    ('Expect', 'HubSpot Expect'),
                    ('Commit', 'HubSpot Commit'),
                    ('Best Case', 'HubSpot Best Case'),
                    ('Opportunity', 'HubSpot Opportunity')
                ]:
                    if selected_sources.get(checkbox_name, False):
                        if individual_selection_mode.get(checkbox_name, False) and checkbox_name in individual_selections:
                            # Use individual selections
                            for item in individual_selections[checkbox_name]:
                                row = item['row']
                                export_data.append({
                                    'Type': f'HubSpot Deal - {status_name}',
                                    'ID': row.get('Record ID', ''),
                                    'Name': row.get('Deal Name', ''),
                                    'Customer': row.get('Account Name', ''),
                                    'Amount': item['amount'],
                                    'Date': row.get('Close Date', ''),
                                    'Sales Rep': row.get('Deal Owner', '')
                                })
                        else:
                            # Bulk export
                            status_deals = hs_deals[hs_deals['Status'] == status_name].copy()
                            for _, row in status_deals.iterrows():
                                export_data.append({
                                    'Type': f'HubSpot Deal - {status_name}',
                                    'ID': row.get('Record ID', ''),
                                    'Name': row.get('Deal Name', ''),
                                    'Customer': row.get('Account Name', ''),
                                    'Amount': row.get('Amount_Numeric', 0),
                                    'Date': row.get('Close Date', ''),
                                    'Sales Rep': row.get('Deal Owner', '')
                                })
                
                # Q1 Spillover - check individual mode
                if selected_sources.get('Q1 Spillover - Expect/Commit', False):
                    category = 'Q1 Spillover - Expect/Commit'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
                    else:
                        # Bulk export
                        q1_deals = hs_deals[
                            (hs_deals.get('Counts_In_Q4', True) == False) &
                            (hs_deals['Status'].isin(['Expect', 'Commit']))
                        ].copy()
                        for _, row in q1_deals.iterrows():
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': row.get('Amount_Numeric', 0),
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
        
        # Create export dataframes
        if export_data:
            summary_df = pd.DataFrame(export_summary)
            export_df = pd.DataFrame(export_data)
            
            # Format amounts for export
            export_df['Amount'] = export_df['Amount'].apply(lambda x: f"${x:,.0f}" if pd.notna(x) else "$0")
            
            # Combine summary and detail
            final_export = summary_df.to_csv(index=False) + '\n' + export_df.to_csv(index=False)
            
            st.download_button(
                label="📥 Download Your Winning Pipeline",
                data=final_export,
                file_name=f"winning_pipeline_{'team' if rep_name is None else rep_name}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                help="Download your selected forecast components with summary and details",
                key=f"download_pipeline_{'team' if rep_name is None else rep_name}_v1"
            )
            
            st.caption(f"Export includes summary + {len(export_df)} line items from your selected categories")


def main():
    st.markdown("""
    <div style='text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 color: white; border-radius: 10px; margin-bottom: 20px;'>
        <h3>📦 Q4 2025 Shipping Planning</h3>
        <p style='font-size: 14px; margin: 0;'>Build Your Shipping Plan - Same as Forecast Tool</p>
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
    
    # Call the EXACT Build Your Own Forecast function
    build_your_own_forecast_section(
        metrics=metrics,
        quota=quota,
        rep_name=None,  # Team view
        deals_df=deals_df,
        invoices_df=invoices_df,
        sales_orders_df=sales_orders_df
    )

if __name__ == "__main__":
    main()
