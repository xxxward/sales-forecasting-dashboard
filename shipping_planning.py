"""
Q4 2025 Shipping Planning Tool - ENHANCED VERSION
Features:
- Proper dropdown displays with company names, amounts, links to NetSuite/HubSpot
- Dynamic ship date inputs for orders without dates
- Ship date visualization chart
- Export functionality with ship dates
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "v2_enhanced"

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
    """Load and process all data"""
    
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
            elif col == 'Account Name':
                rename_dict[col] = 'Account Name'
        
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
        
        # Apply Rep Master override - EXACT PATTERN FROM SALES_DASHBOARD.PY
        if 'Rep Master' in invoices_df.columns:
            # Rep Master takes priority - replace Sales Rep with Rep Master values
            # But first, clean up Rep Master to handle any formula errors or blanks
            invoices_df['Rep Master'] = invoices_df['Rep Master'].astype(str).str.strip()
            # Replace empty strings, 'nan', 'None', '#N/A', '#REF!' with original Sales Rep
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            mask = invoices_df['Rep Master'].isin(invalid_values)
            
            # Only replace Sales Rep where Rep Master has a valid value
            invoices_df.loc[~mask, 'Sales Rep'] = invoices_df.loc[~mask, 'Rep Master']
            # Drop the Rep Master column since we've copied it to Sales Rep
            invoices_df = invoices_df.drop(columns=['Rep Master'])
        
        # Apply customer name correction - EXACT PATTERN FROM SALES_DASHBOARD.PY
        if 'Corrected Customer Name' in invoices_df.columns:
            # Corrected Customer Name takes priority - replace Customer with corrected values
            invoices_df['Corrected Customer Name'] = invoices_df['Corrected Customer Name'].astype(str).str.strip()
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            mask = invoices_df['Corrected Customer Name'].isin(invalid_values)
            invoices_df.loc[~mask, 'Customer'] = invoices_df.loc[~mask, 'Corrected Customer Name']
            # Drop the Corrected Customer Name column since we've copied it to Customer
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
    
    # Process sales orders
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
        
        # Apply Rep Master override - EXACT PATTERN FROM SALES_DASHBOARD.PY (DIRECT REPLACEMENT)
        if 'Rep Master' in sales_orders_df.columns:
            # Rep Master takes priority - replace Sales Rep with Rep Master values
            sales_orders_df['Sales Rep'] = sales_orders_df['Rep Master']
            # Drop the Rep Master column since we've copied it to Sales Rep
            sales_orders_df = sales_orders_df.drop(columns=['Rep Master'])
        
        # CRITICAL: Remove any duplicate columns that may have been created
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
    """Calculate team-level metrics with detail dataframes"""
    
    metrics = {
        'orders': 0,
        'pending_fulfillment': 0,
        'pending_approval': 0,
        'expect_commit': 0,
        'pending_fulfillment_no_date': 0,
        'pending_approval_no_date': 0,
        'pending_approval_old': 0,
        'q1_spillover_expect_commit': 0,
        'q1_spillover_best_opp': 0,
        # Detail dataframes
        'invoices_details': pd.DataFrame(),
        'pending_fulfillment_details': pd.DataFrame(),
        'pending_approval_details': pd.DataFrame(),
        'pending_fulfillment_no_date_details': pd.DataFrame(),
        'pending_approval_no_date_details': pd.DataFrame(),
        'pending_approval_old_details': pd.DataFrame(),
        'expect_commit_deals': pd.DataFrame(),
        'commit_deals': pd.DataFrame(),
        'best_case_deals': pd.DataFrame(),
        'opportunity_deals': pd.DataFrame(),
        'q1_spillover_expect_commit_deals': pd.DataFrame(),
        'q1_spillover_best_opp_deals': pd.DataFrame()
    }
    
    # Invoiced & Shipped
    if invoices_df is not None and not invoices_df.empty:
        metrics['invoices_details'] = invoices_df.copy()
        metrics['orders'] = invoices_df['Amount'].sum()
    
    # Sales Orders
    if sales_orders_df is not None and not sales_orders_df.empty:
        # Pending Fulfillment - with date
        pf_orders = sales_orders_df[sales_orders_df['Status'] == 'Pending Fulfillment'].copy()
        if not pf_orders.empty:
            # Check for Customer Promise Date or Projected Date
            pf_orders['Has_Date'] = (
                pf_orders['Customer Promise Date'].notna() | 
                pf_orders['Projected Date'].notna()
            )
            
            pf_with_date = pf_orders[pf_orders['Has_Date'] == True].copy()
            pf_no_date = pf_orders[pf_orders['Has_Date'] == False].copy()
            
            metrics['pending_fulfillment_details'] = pf_with_date
            metrics['pending_fulfillment'] = pf_with_date['Amount'].sum()
            
            metrics['pending_fulfillment_no_date_details'] = pf_no_date
            metrics['pending_fulfillment_no_date'] = pf_no_date['Amount'].sum()
        
        # Pending Approval
        pa_orders = sales_orders_df[sales_orders_df['Status'] == 'Pending Approval'].copy()
        if not pa_orders.empty:
            # With date
            pa_with_date = pa_orders[pa_orders['Pending Approval Date'].notna()].copy()
            pa_no_date = pa_orders[pa_orders['Pending Approval Date'].isna()].copy()
            
            # Old PA (>= 10 business days)
            if 'Age_Business_Days' in pa_orders.columns:
                pa_old = pa_orders[pa_orders['Age_Business_Days'] >= 10].copy()
                metrics['pending_approval_old_details'] = pa_old
                metrics['pending_approval_old'] = pa_old['Amount'].sum()
                
                # Young PA with date
                pa_young_with_date = pa_with_date[pa_with_date['Age_Business_Days'] < 10].copy()
                metrics['pending_approval_details'] = pa_young_with_date
                metrics['pending_approval'] = pa_young_with_date['Amount'].sum()
            else:
                metrics['pending_approval_details'] = pa_with_date
                metrics['pending_approval'] = pa_with_date['Amount'].sum()
            
            metrics['pending_approval_no_date_details'] = pa_no_date
            metrics['pending_approval_no_date'] = pa_no_date['Amount'].sum()
    
    # HubSpot Deals
    if deals_df is not None and not deals_df.empty and 'Status' in deals_df.columns:
        deals_df['Amount_Numeric'] = pd.to_numeric(deals_df['Amount'], errors='coerce')
        
        # Q4 deals
        q4_deals = deals_df[deals_df.get('Counts_In_Q4', True) == True]
        
        expect_deals = q4_deals[q4_deals['Status'] == 'Expect'].copy()
        commit_deals = q4_deals[q4_deals['Status'] == 'Commit'].copy()
        best_case_deals = q4_deals[q4_deals['Status'] == 'Best Case'].copy()
        opportunity_deals = q4_deals[q4_deals['Status'] == 'Opportunity'].copy()
        
        metrics['expect_commit_deals'] = expect_deals
        metrics['commit_deals'] = commit_deals
        metrics['best_case_deals'] = best_case_deals
        metrics['opportunity_deals'] = opportunity_deals
        
        metrics['expect_commit'] = expect_deals['Amount_Numeric'].sum() + commit_deals['Amount_Numeric'].sum()
        
        # Q1 Spillover
        if 'Q1 2026 Spillover' in deals_df.columns:
            q1_deals = deals_df[deals_df['Q1 2026 Spillover'] == 'Q1 2026']
            
            q1_expect_commit = q1_deals[q1_deals['Status'].isin(['Expect', 'Commit'])].copy()
            q1_best_opp = q1_deals[q1_deals['Status'].isin(['Best Case', 'Opportunity'])].copy()
            
            metrics['q1_spillover_expect_commit_deals'] = q1_expect_commit
            metrics['q1_spillover_best_opp_deals'] = q1_best_opp
            
            metrics['q1_spillover_expect_commit'] = q1_expect_commit['Amount_Numeric'].sum()
            metrics['q1_spillover_best_opp'] = q1_best_opp['Amount_Numeric'].sum()
    
    return metrics

def display_drill_down_with_ship_dates(title, amount, details_df, category_key, ship_dates_dict, selected_items_dict):
    """Display collapsible section with checkboxes for individual selection and ship date inputs"""
    
    item_count = len(details_df)
    if item_count == 0:
        return
    
    with st.expander(f"{title}: ${amount:,.0f} (👀 Click to see {item_count} {'item' if item_count == 1 else 'items'})"):
        # Determine data type
        is_hubspot = 'Deal Name' in details_df.columns
        is_invoice = 'Invoice Number' in details_df.columns
        is_netsuite = 'Document Number' in details_df.columns or 'Internal ID' in details_df.columns
        
        # Determine if this category needs ship dates
        needs_ship_dates = category_key in [
            'pf_no_date', 'pa_no_date', 'pa_old',
            'hs_expect', 'hs_commit', 'hs_best_case', 'hs_opportunity',
            'q1_expect_commit', 'q1_best_opp'
        ]
        
        st.markdown("**Select items to include:**")
        
        # Track selected items for this category
        if category_key not in selected_items_dict:
            selected_items_dict[category_key] = []
        
        # Display items with checkboxes
        for idx, row in details_df.iterrows():
            col1, col2 = st.columns([3, 1])
            
            with col1:
                # Build display label based on type
                if is_hubspot:
                    item_id = row.get('Record ID', idx)
                    deal_name = row.get('Deal Name', 'Unknown Deal')
                    company = row.get('Account Name', '')
                    amount_val = row.get('Amount', 0)
                    close_date = row.get('Close Date', '')
                    if pd.notna(close_date) and hasattr(close_date, 'strftime'):
                        close_date_str = close_date.strftime('%Y-%m-%d')
                    else:
                        close_date_str = str(close_date) if pd.notna(close_date) else ''
                    
                    label = f"**{deal_name}** | {company} | ${amount_val:,.0f}"
                    if close_date_str:
                        label += f" | Close: {close_date_str}"
                    
                    link = f"https://app.hubspot.com/contacts/6712259/record/0-3/{item_id}/"
                
                elif is_invoice:
                    item_id = row.get('Invoice Number', idx)
                    company = row.get('Customer', '')
                    amount_val = row.get('Amount', 0)
                    invoice_date = row.get('Date', '')
                    if pd.notna(invoice_date) and hasattr(invoice_date, 'strftime'):
                        date_str = invoice_date.strftime('%Y-%m-%d')
                    else:
                        date_str = str(invoice_date) if pd.notna(invoice_date) else ''
                    
                    label = f"**INV #{item_id}** | {company} | ${amount_val:,.0f}"
                    if date_str:
                        label += f" | {date_str}"
                    link = None
                
                else:  # NetSuite
                    item_id = row.get('Document Number', idx)
                    internal_id = row.get('Internal ID', '')
                    company = row.get('Customer', '')
                    amount_val = row.get('Amount', 0)
                    status = row.get('Status', '')
                    
                    # Show the most relevant date
                    date_to_show = None
                    date_label = ""
                    if pd.notna(row.get('Customer Promise Date')):
                        date_to_show = row.get('Customer Promise Date')
                        date_label = "Cust Promise"
                    elif pd.notna(row.get('Projected Date')):
                        date_to_show = row.get('Projected Date')
                        date_label = "Projected"
                    elif pd.notna(row.get('Pending Approval Date')):
                        date_to_show = row.get('Pending Approval Date')
                        date_label = "PA Date"
                    
                    if date_to_show and hasattr(date_to_show, 'strftime'):
                        date_str = f"{date_label}: {date_to_show.strftime('%Y-%m-%d')}"
                    else:
                        date_str = ""
                    
                    label = f"**SO# {item_id}** | {company} | ${amount_val:,.0f} | {status}"
                    if date_str:
                        label += f" | {date_str}"
                    
                    link = f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={internal_id}&whence=" if pd.notna(internal_id) else None
                
                # Checkbox for selection
                is_selected = st.checkbox(
                    label,
                    value=False,
                    key=f"{category_key}_{item_id}_{idx}"
                )
            
            with col2:
                # Show link if available
                if link:
                    st.markdown(f"[View →]({link})")
            
            # If selected, add to tracking and show ship date input if needed
            if is_selected:
                # Track this selection
                selected_items_dict[category_key].append({
                    'id': item_id,
                    'amount': amount_val,
                    'row': row
                })
                
                # Show ship date input if needed
                if needs_ship_dates:
                    default_date = datetime.now() + timedelta(days=14)
                    ship_date = st.date_input(
                        f"📅 Ship Date",
                        value=default_date,
                        key=f"ship_date_{category_key}_{item_id}_{idx}"
                    )
                    
                    # Store in ship dates dictionary
                    ship_dates_dict[f"{category_key}_{item_id}"] = {
                        'ship_date': ship_date,
                        'amount': amount_val,
                        'row': row
                    }
        
        # Summary
        selected_count = len(selected_items_dict.get(category_key, []))
        if selected_count > 0:
            selected_total = sum(item['amount'] for item in selected_items_dict[category_key])
            st.success(f"✓ Selected {selected_count} of {item_count} items | Total: ${selected_total:,.0f}")
        else:
            st.info(f"No items selected (0 of {item_count})")

def create_ship_date_chart(ship_dates_dict, custom_forecast):
    """Create a timeline chart showing when things will ship"""
    
    if not ship_dates_dict:
        return None
    
    # Prepare data for chart
    ship_data = []
    for item_id, data in ship_dates_dict.items():
        ship_data.append({
            'date': data['ship_date'],
            'amount': data['amount']
        })
    
    if not ship_data:
        return None
    
    # Create dataframe and aggregate by date
    ship_df = pd.DataFrame(ship_data)
    ship_df['date'] = pd.to_datetime(ship_df['date'])
    daily_ships = ship_df.groupby('date')['amount'].sum().reset_index()
    daily_ships = daily_ships.sort_values('date')
    
    # Calculate cumulative
    daily_ships['cumulative'] = daily_ships['amount'].cumsum()
    
    # Create chart
    fig = go.Figure()
    
    # Daily shipping bars
    fig.add_trace(go.Bar(
        x=daily_ships['date'],
        y=daily_ships['amount'],
        name='Daily Ship Amount',
        marker_color='lightblue',
        yaxis='y'
    ))
    
    # Cumulative line
    fig.add_trace(go.Scatter(
        x=daily_ships['date'],
        y=daily_ships['cumulative'],
        name='Cumulative Shipped',
        line=dict(color='darkblue', width=3),
        yaxis='y2'
    ))
    
    # Add forecast line
    fig.add_hline(
        y=custom_forecast,
        line_dash="dash",
        line_color="green",
        annotation_text=f"Total Forecast: ${custom_forecast:,.0f}",
        annotation_position="right"
    )
    
    fig.update_layout(
        title="📦 Shipping Timeline",
        xaxis_title="Ship Date",
        yaxis_title="Daily Ship Amount ($)",
        yaxis2=dict(
            title="Cumulative Shipped ($)",
            overlaying='y',
            side='right'
        ),
        hovermode='x unified',
        height=400,
        showlegend=True
    )
    
    return fig

def build_shipping_plan_section(metrics, quota, deals_df=None, invoices_df=None, sales_orders_df=None):
    """
    Interactive shipping planning section with proper dropdowns and ship date inputs
    """
    st.markdown("### 📦 Build Your Shipping Plan")
    st.caption("Select components and set ship dates for items without dates")
    
    # Initialize session state for ship dates
    if 'ship_dates' not in st.session_state:
        st.session_state.ship_dates = {}
    
    ship_dates_dict = {}
    selected_items_dict = {}
    
    # Create columns for checkboxes
    col1, col2, col3 = st.columns(3)
    
    # Available data sources
    sources = {
        'Invoiced & Shipped': metrics.get('orders', 0),
        'Pending Fulfillment (with date)': metrics.get('pending_fulfillment', 0),
        'Pending Approval (with date)': metrics.get('pending_approval', 0),
        'HubSpot Expect': 0,
        'HubSpot Commit': 0,
        'HubSpot Best Case': 0,
        'HubSpot Opportunity': 0,
        'Pending Fulfillment (without date)': metrics.get('pending_fulfillment_no_date', 0),
        'Pending Approval (without date)': metrics.get('pending_approval_no_date', 0),
        'Pending Approval (>2 weeks old)': metrics.get('pending_approval_old', 0),
        'Q1 Spillover - Expect/Commit': metrics.get('q1_spillover_expect_commit', 0),
        'Q1 Spillover - Best Case': metrics.get('q1_spillover_best_opp', 0)
    }
    
    # Calculate individual HubSpot categories
    if 'expect_commit_deals' in metrics and not metrics['expect_commit_deals'].empty:
        sources['HubSpot Expect'] = metrics['expect_commit_deals']['Amount_Numeric'].sum()
    if 'commit_deals' in metrics and not metrics['commit_deals'].empty:
        sources['HubSpot Commit'] = metrics['commit_deals']['Amount_Numeric'].sum()
    if 'best_case_deals' in metrics and not metrics['best_case_deals'].empty:
        sources['HubSpot Best Case'] = metrics['best_case_deals']['Amount_Numeric'].sum()
    if 'opportunity_deals' in metrics and not metrics['opportunity_deals'].empty:
        sources['HubSpot Opportunity'] = metrics['opportunity_deals']['Amount_Numeric'].sum()
    
    # Create checkboxes
    selected_sources = {}
    source_list = list(sources.keys())
    
    with col1:
        for source in source_list[0:4]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    with col2:
        for source in source_list[4:8]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    with col3:
        for source in source_list[8:]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"team_{source}"
            )
    
    # Display drill-downs for selected categories
    st.markdown("---")
    st.markdown("#### 🔍 Selected Components Details")
    
    if selected_sources.get('Invoiced & Shipped', False):
        display_drill_down_with_ship_dates(
            "✅ Invoiced & Shipped",
            sources['Invoiced & Shipped'],
            metrics.get('invoices_details', pd.DataFrame()),
            'invoices',
            ship_dates_dict,
            selected_items_dict
        )
    
    if selected_sources.get('Pending Fulfillment (with date)', False):
        display_drill_down_with_ship_dates(
            "📦 Pending Fulfillment (with date)",
            sources['Pending Fulfillment (with date)'],
            metrics.get('pending_fulfillment_details', pd.DataFrame()),
            'pf_date',
            ship_dates_dict,
            selected_items_dict
        )
    
    if selected_sources.get('Pending Approval (with date)', False):
        display_drill_down_with_ship_dates(
            "⏳ Pending Approval (with date)",
            sources['Pending Approval (with date)'],
            metrics.get('pending_approval_details', pd.DataFrame()),
            'pa_date',
            ship_dates_dict,
            selected_items_dict
        )
    
    if selected_sources.get('HubSpot Expect', False):
        display_drill_down_with_ship_dates(
            "🎯 HubSpot Expect",
            sources['HubSpot Expect'],
            metrics.get('expect_commit_deals', pd.DataFrame()),
            'hs_expect',
            ship_dates_dict,
            selected_items_dict
        )
    
    if selected_sources.get('HubSpot Commit', False):
        display_drill_down_with_ship_dates(
            "🎯 HubSpot Commit",
            sources['HubSpot Commit'],
            metrics.get('commit_deals', pd.DataFrame()),
            'hs_commit',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('HubSpot Best Case', False):
        display_drill_down_with_ship_dates(
            "🎲 HubSpot Best Case",
            sources['HubSpot Best Case'],
            metrics.get('best_case_deals', pd.DataFrame()),
            'hs_best_case',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('HubSpot Opportunity', False):
        display_drill_down_with_ship_dates(
            "🌟 HubSpot Opportunity",
            sources['HubSpot Opportunity'],
            metrics.get('opportunity_deals', pd.DataFrame()),
            'hs_opportunity',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('Pending Fulfillment (without date)', False):
        display_drill_down_with_ship_dates(
            "📦 Pending Fulfillment (without date)",
            sources['Pending Fulfillment (without date)'],
            metrics.get('pending_fulfillment_no_date_details', pd.DataFrame()),
            'pf_no_date',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('Pending Approval (without date)', False):
        display_drill_down_with_ship_dates(
            "⏳ Pending Approval (without date)",
            sources['Pending Approval (without date)'],
            metrics.get('pending_approval_no_date_details', pd.DataFrame()),
            'pa_no_date',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('Pending Approval (>2 weeks old)', False):
        display_drill_down_with_ship_dates(
            "⏰ Pending Approval (>2 weeks old)",
            sources['Pending Approval (>2 weeks old)'],
            metrics.get('pending_approval_old_details', pd.DataFrame()),
            'pa_old',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('Q1 Spillover - Expect/Commit', False):
        display_drill_down_with_ship_dates(
            "🔄 Q1 Spillover - Expect/Commit",
            sources['Q1 Spillover - Expect/Commit'],
            metrics.get('q1_spillover_expect_commit_deals', pd.DataFrame()),
            'q1_expect_commit',
            ship_dates_dict, selected_items_dict
        )
    
    if selected_sources.get('Q1 Spillover - Best Case', False):
        display_drill_down_with_ship_dates(
            "🔄 Q1 Spillover - Best Case",
            sources['Q1 Spillover - Best Case'],
            metrics.get('q1_spillover_best_opp_deals', pd.DataFrame()),
            'q1_best_opp',
            ship_dates_dict, selected_items_dict
        )
    
    # Calculate totals based on individually selected items
    custom_forecast = 0
    for category, items in selected_items_dict.items():
        custom_forecast += sum(item['amount'] for item in items)
    
    # Only count invoiced_shipped if items were selected
    if 'invoices' in selected_items_dict and selected_items_dict['invoices']:
        invoiced_shipped = sum(item['amount'] for item in selected_items_dict['invoices'])
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
    
    # Display shipping timeline chart
    if ship_dates_dict:
        st.markdown("---")
        st.markdown("#### 📅 Shipping Timeline")
        
        fig = create_ship_date_chart(ship_dates_dict, custom_forecast)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    
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
    
    # Export functionality
    if any(selected_sources.values()):
        st.markdown("---")
        
        # Build export data
        export_data = []
        
        # Add all selected items with ship dates
        for item_id, data in ship_dates_dict.items():
            row = data['row']
            
            # Determine type and get appropriate fields
            if 'Deal Name' in row:
                export_data.append({
                    'Type': 'HubSpot Deal',
                    'ID': row.get('Record ID', ''),
                    'Name': row.get('Deal Name', ''),
                    'Company': row.get('Account Name', ''),
                    'Amount': f"${data['amount']:,.0f}",
                    'Status': row.get('Status', ''),
                    'Ship Date': data['ship_date'].strftime('%Y-%m-%d'),
                    'Link': f"https://app.hubspot.com/contacts/6712259/record/0-3/{row.get('Record ID', '')}/"
                })
            elif 'Invoice Number' in row:
                export_data.append({
                    'Type': 'Invoice',
                    'ID': row.get('Invoice Number', ''),
                    'Name': '',
                    'Company': row.get('Customer', ''),
                    'Amount': f"${data['amount']:,.0f}",
                    'Status': 'Invoiced',
                    'Ship Date': 'Already Shipped',
                    'Link': ''
                })
            else:
                # Sales order
                export_data.append({
                    'Type': 'Sales Order',
                    'ID': row.get('Document Number', ''),
                    'Name': '',
                    'Company': row.get('Customer', ''),
                    'Amount': f"${data['amount']:,.0f}",
                    'Status': row.get('Status', ''),
                    'Ship Date': data['ship_date'].strftime('%Y-%m-%d'),
                    'Link': f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={row.get('Internal ID', '')}&whence="
                })
        
        if export_data:
            export_df = pd.DataFrame(export_data)
            csv = export_df.to_csv(index=False)
            
            st.download_button(
                label="📥 Download Shipping Plan",
                data=csv,
                file_name=f"shipping_plan_team_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                help="Download your shipping plan with ship dates"
            )
            
            st.caption(f"Export includes {len(export_df)} line items with ship dates")

def main():
    st.markdown("""
    <div style='text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 color: white; border-radius: 10px; margin-bottom: 20px;'>
        <h3>📦 Q4 2025 Shipping Planning - Enhanced</h3>
        <p style='font-size: 14px; margin: 0;'>Build Your Shipping Plan with Ship Date Tracking</p>
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
