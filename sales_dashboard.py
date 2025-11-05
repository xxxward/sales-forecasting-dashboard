"""
Sales Forecasting Dashboard - Enhanced Version with Drill-Down Capability
Reads from Google Sheets and displays gap-to-goal analysis with interactive visualizations
Includes lead time logic for Q4/Q1 fulfillment determination and detailed order drill-downs
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
from datetime import datetime, timedelta
import time
import base64
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Sales Forecasting Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
    <style>
    .big-font {
        font-size: 28px !important;
        font-weight: bold;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 1px 1px 3px rgba(0,0,0,0.1);
    }
    .progress-breakdown {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 25px;
        border-radius: 15px;
        color: white;
        margin: 20px 0;
        box-shadow: 0 10px 20px rgba(0,0,0,0.2);
    }
    .progress-breakdown h3 {
        color: white;
        margin-bottom: 15px;
        font-size: 24px;
    }
    .progress-item {
        display: flex;
        justify-content: space-between;
        padding: 10px 0;
        border-bottom: 1px solid rgba(255,255,255,0.2);
    }
    .progress-item:last-child {
        border-bottom: none;
        font-weight: bold;
        font-size: 18px;
        padding-top: 15px;
        border-top: 2px solid rgba(255,255,255,0.4);
    }
    .progress-label {
        font-size: 16px;
    }
    .progress-value {
        font-size: 16px;
        font-weight: 600;
    }
    .reconciliation-table {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .section-header {
        background: #f0f2f6;
        padding: 10px 15px;
        border-radius: 8px;
        margin: 15px 0;
        font-weight: bold;
    }
    .drill-down-section {
        background: #f8f9fa;
        padding: 10px;
        border-radius: 8px;
        margin: 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v25_pa_q4_filter"

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """
    Load data from Google Sheets with caching and enhanced error handling
    """
    try:
        # Check if secrets exist
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing Google Cloud credentials in Streamlit secrets")
            return pd.DataFrame()
        
        # Load credentials from Streamlit secrets
        creds_dict = dict(st.secrets["gcp_service_account"])
        
        # Create credentials
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=SCOPES
        )
        
        # Build service
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        # Fetch data
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning(f"⚠️ No data found in {sheet_name}!{range_name}")
            return pd.DataFrame()
        
        # Handle mismatched column counts - pad shorter rows with empty strings
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        # Convert to DataFrame
        df = pd.DataFrame(values[1:], columns=values[0])
        
        # Success message (only in sidebar to avoid clutter)
        st.sidebar.success(f"✅ Loaded {len(df)} rows from {sheet_name}")
        
        return df
        
    except Exception as e:
        error_msg = str(e)
        st.error(f"❌ Error loading data from {sheet_name}: {error_msg}")
        
        # Provide specific troubleshooting based on error type
        if "403" in error_msg or "permission" in error_msg.lower():
            st.warning("""
            **Permission Error:**
            - Make sure you've shared the Google Sheet with your service account email
            - The service account email looks like: `your-service-account@project.iam.gserviceaccount.com`
            - Share the sheet with 'Viewer' access
            """)
        elif "404" in error_msg or "not found" in error_msg.lower():
            st.warning("""
            **Sheet Not Found:**
            - Check that the spreadsheet ID is correct
            - Check that the sheet name matches exactly (case-sensitive)
            - Current spreadsheet ID: `12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk`
            """)
        elif "401" in error_msg or "authentication" in error_msg.lower():
            st.warning("""
            **Authentication Error:**
            - Your service account credentials may be invalid
            - Try regenerating the service account key in Google Cloud Console
            """)
        
        return pd.DataFrame()

def apply_q4_fulfillment_logic(deals_df):
    """
    Apply lead time logic to filter out deals that close late in Q4 
    but won't ship until Q1 based on product type
    """
    # Lead time mapping based on your image
    lead_time_map = {
        'Labeled - Labels In Stock': 10,
        'Outer Boxes': 20,
        'Non-Labeled - 1 Week Lead Time': 5,
        'Non-Labeled - 2 Week Lead Time': 10,
        'Labeled - Print & Apply': 20,
        'Non-Labeled - Custom Lead Time': 30,
        'Labeled with FEP - Print & Apply': 35,
        'Labeled - Custom Lead Time': 40,
        'Flexpack': 25,
        'Labels Only - Direct to Customer': 15,
        'Labels Only - For Inventory': 15,
        'Labeled with FEP - Labels In Stock': 25,
        'Labels Only (deprecated)': 15
    }
    
    # Calculate cutoff date for each product type
    q4_end = pd.Timestamp('2025-12-31')
    
    def get_business_days_before(end_date, business_days):
        """Calculate date that is N business days before end_date"""
        current = end_date
        days_counted = 0
        
        while days_counted < business_days:
            current -= timedelta(days=1)
            # Skip weekends (Monday=0, Sunday=6)
            if current.weekday() < 5:
                days_counted += 1
        
        return current
    
    # Add a column to track if deal counts for Q4
    deals_df['Counts_In_Q4'] = True
    deals_df['Q1_Spillover_Amount'] = 0
    
    # Check if we have a Product Type column
    if 'Product Type' in deals_df.columns:
        for product_type, lead_days in lead_time_map.items():
            cutoff_date = get_business_days_before(q4_end, lead_days)
            
            # Mark deals closing after cutoff as Q1
            mask = (
                (deals_df['Product Type'] == product_type) & 
                (deals_df['Close Date'] > cutoff_date) &
                (deals_df['Close Date'].notna())
            )
            deals_df.loc[mask, 'Counts_In_Q4'] = False
            deals_df.loc[mask, 'Q1_Spillover_Amount'] = deals_df.loc[mask, 'Amount']
            
        # Log how many deals were excluded
        excluded_count = (~deals_df['Counts_In_Q4']).sum()
        excluded_value = deals_df[~deals_df['Counts_In_Q4']]['Amount'].sum()
        
        if excluded_count > 0:
            st.sidebar.info(f"📊 {excluded_count} deals (${excluded_value:,.0f}) deferred to Q1 2026 due to lead times")
    else:
        st.sidebar.warning("⚠️ No 'Product Type' column found - lead time logic not applied")
    
    return deals_df

def load_all_data():
    """Load all necessary data from Google Sheets"""
    
    st.sidebar.info("🔄 Loading data from Google Sheets...")
    
    # Load deals data - extend range to include all columns (A through Q for Pending Approval Date)
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:Q", version=CACHE_VERSION)
    
    # DEBUG: Show what we got from HubSpot
    if not deals_df.empty:
        st.sidebar.success(f"📊 HubSpot raw data: {len(deals_df)} rows, {len(deals_df.columns)} columns")
        with st.sidebar.expander("🔍 HubSpot Column Names"):
            for i, col in enumerate(deals_df.columns):
                st.write(f"{i}: {col}")
    else:
        st.sidebar.error("❌ No HubSpot data loaded!")
    
    # Load dashboard info (rep quotas and orders)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoice data from NetSuite
    invoices_df = load_google_sheets_data("NS Invoices", "A:Z", version=CACHE_VERSION)
    
    # Load sales orders data from NetSuite - EXTEND to include Column AB
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AB", version=CACHE_VERSION)
    
    # Clean and process deals data - FIXED VERSION to match actual sheet
    if not deals_df.empty and len(deals_df.columns) >= 6:
        # Get column names from first row
        if len(deals_df) > 0:
            # Get actual column names
            col_names = deals_df.columns.tolist()
            
            st.sidebar.info(f"Processing {len(col_names)} HubSpot columns")
            st.sidebar.info(f"First 10 columns: {col_names[:10]}")
            
            # Map based on ACTUAL column names from your sheet
            # Note: Column 4 appears to be "Deal Owner First Name Deal Owner Last Name" combined
            
            rename_dict = {}
            
            # Map columns by actual names (case-sensitive)
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
                    # This column has both names already combined
                    rename_dict[col] = 'Deal Owner'
                elif col == 'Deal Owner First Name':
                    rename_dict[col] = 'Deal Owner First Name'
                elif col == 'Deal Owner Last Name':
                    rename_dict[col] = 'Deal Owner Last Name'
                elif col == 'Amount':
                    rename_dict[col] = 'Amount'
                elif col == 'Close Status':
                    rename_dict[col] = 'Status'  # Map Close Status to Status
                elif col == 'Pipeline':
                    rename_dict[col] = 'Pipeline'
                elif col == 'Deal Type':
                    rename_dict[col] = 'Product Type'  # Map Deal Type to Product Type for lead time logic
                elif col == 'Average Leadtime':
                    rename_dict[col] = 'Average Leadtime'
            
            deals_df = deals_df.rename(columns=rename_dict)
            
            # Check if Deal Owner already exists (from combined column)
            if 'Deal Owner' not in deals_df.columns:
                # Create a combined "Deal Owner" field from First Name + Last Name if they're separate
                if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                    deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                    deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
                    st.sidebar.success("✅ Created Deal Owner from First + Last Name")
                else:
                    st.sidebar.error("❌ Missing Deal Owner column!")
            else:
                st.sidebar.success("✅ Deal Owner column already exists")
                # Clean up the Deal Owner field
                deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
            
            # Show what we have after renaming
            st.sidebar.success(f"✅ Columns after rename: {', '.join([c for c in deals_df.columns.tolist()[:10] if c])}")
            
            # Check if we have required columns
            required_cols = ['Deal Name', 'Status', 'Close Date', 'Deal Owner', 'Amount', 'Pipeline']
            missing_cols = [col for col in required_cols if col not in deals_df.columns]
            if missing_cols:
                st.sidebar.error(f"❌ Missing required columns: {missing_cols}")
            
            # Clean and convert amount to numeric
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
            else:
                st.sidebar.error("❌ No Amount column found!")
            
            # Convert close date to datetime
            if 'Close Date' in deals_df.columns:
                deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
                
                # Debug: Show date range in the data
                valid_dates = deals_df['Close Date'].dropna()
                if len(valid_dates) > 0:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    st.sidebar.info(f"📅 Date range in data: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")
                    
                    # Count deals in each quarter
                    q4_2024_count = len(deals_df[(deals_df['Close Date'] >= '2024-10-01') & (deals_df['Close Date'] <= '2024-12-31')])
                    q1_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-01-01') & (deals_df['Close Date'] <= '2025-03-31')])
                    q2_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-04-01') & (deals_df['Close Date'] <= '2025-06-30')])
                    q3_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-07-01') & (deals_df['Close Date'] <= '2025-09-30')])
                    q4_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-10-01') & (deals_df['Close Date'] <= '2025-12-31')])
                    
                    st.sidebar.info(f"Q4 2024: {q4_2024_count} | Q1 2025: {q1_2025_count} | Q2 2025: {q2_2025_count} | Q3 2025: {q3_2025_count} | Q4 2025: {q4_2025_count}")
                else:
                    st.sidebar.error("❌ No valid dates found in Close Date column!")
            else:
                st.sidebar.error("❌ No Close Date column found!")
            
            # Show data before filtering
            total_deals_before = len(deals_df)
            total_amount_before = deals_df['Amount'].sum() if 'Amount' in deals_df.columns else 0
            st.sidebar.info(f"📊 Before filtering: {total_deals_before} deals, ${total_amount_before:,.0f}")
            
            # Show unique values in Status column
            if 'Status' in deals_df.columns:
                unique_statuses = deals_df['Status'].unique()
                st.sidebar.info(f"🏷️ Unique Status values: {', '.join([str(s) for s in unique_statuses[:10]])}")
            else:
                st.sidebar.error("❌ No Status column found! Check 'Close Status' mapping")
            
            # FILTER: Only Q4 2025 deals (Oct 1 - Dec 31, 2025)
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            if 'Close Date' in deals_df.columns:
                before_count = len(deals_df)
                deals_df = deals_df[
                    (deals_df['Close Date'] >= q4_start) & 
                    (deals_df['Close Date'] <= q4_end)
                ]
                after_count = len(deals_df)
                
                st.sidebar.info(f"📅 Q4 2025 Filter: {before_count} deals → {after_count} deals")
                
                if after_count == 0:
                    st.sidebar.error("❌ No deals found in Q4 2025 (Oct-Dec 2025)")
                    st.sidebar.info("💡 Your data range is 2019-2021. You may need to refresh your Google Sheet with current HubSpot data.")
                else:
                    st.sidebar.success(f"✅ Found {after_count} Q4 2025 deals worth ${deals_df['Amount'].sum():,.0f}")
            else:
                st.sidebar.error("❌ Cannot apply date filter - no Close Date column")
            
            # FILTER OUT unwanted deal stages
            excluded_stages = [
                '', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 
                'closed lost', 'closed won', 'sales order created in NS', 
                'NCR', 'Shipped'
            ]
            
            # Convert Deal Stage to string and handle NaN
            if 'Deal Stage' in deals_df.columns:
                deals_df['Deal Stage'] = deals_df['Deal Stage'].fillna('')
                deals_df['Deal Stage'] = deals_df['Deal Stage'].astype(str).str.strip()
                
                # Show unique stages before filtering
                unique_stages = deals_df['Deal Stage'].unique()
                st.sidebar.info(f"🎯 Unique Deal Stages: {', '.join([str(s) for s in unique_stages[:10]])}")
                
                # Filter out excluded stages
                deals_df = deals_df[~deals_df['Deal Stage'].str.lower().isin([s.lower() if s else '' for s in excluded_stages])]
                
                st.sidebar.success(f"✅ After stage filter: {len(deals_df)} deals, ${deals_df['Amount'].sum():,.0f}")
            else:
                st.sidebar.warning("⚠️ No Deal Stage column found")
            
            # Apply Q4 fulfillment logic
            deals_df = apply_q4_fulfillment_logic(deals_df)
    else:
        st.sidebar.error(f"❌ HubSpot data has insufficient columns: {len(deals_df.columns) if not deals_df.empty else 0}")
    
    if not dashboard_df.empty:
        # Ensure we have the right column names
        if len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            
            # Remove any empty rows
            dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
            
            # Clean and convert numeric columns
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
    
    # Process invoice data
    if not invoices_df.empty:
        if len(invoices_df.columns) >= 15:
            invoices_df = invoices_df.rename(columns={
                invoices_df.columns[0]: 'Invoice Number',
                invoices_df.columns[1]: 'Status',
                invoices_df.columns[2]: 'Date',
                invoices_df.columns[6]: 'Customer',
                invoices_df.columns[10]: 'Amount',
                invoices_df.columns[14]: 'Sales Rep'
            })
            
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
            
            # Filter to Q4 2025 only
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            invoices_df = invoices_df[
                (invoices_df['Date'] >= q4_start) & 
                (invoices_df['Date'] <= q4_end)
            ]
            
            invoices_df['Sales Rep'] = invoices_df['Sales Rep'].str.strip()
            
            invoices_df = invoices_df[
                (invoices_df['Amount'] > 0) & 
                (invoices_df['Sales Rep'].notna()) & 
                (invoices_df['Sales Rep'] != '')
            ]
            
            # Calculate total invoices by rep
            invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
            invoice_totals.columns = ['Rep Name', 'Invoice Total']
            
            dashboard_df['Rep Name'] = dashboard_df['Rep Name'].str.strip()
            
            dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
            dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
            
            dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
            dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
    
    # Process sales orders data with NEW LOGIC
    if not sales_orders_df.empty:
        # Map column positions
        col_names = sales_orders_df.columns.tolist()
        
        rename_dict = {}
        
        # Find standard columns - only map FIRST occurrence
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
        
        # Map specific columns by position (0-indexed) - be more careful
        if len(col_names) > 8 and 'Order Start Date' not in rename_dict.values():
            rename_dict[col_names[8]] = 'Order Start Date'  # Column I
        if len(col_names) > 11 and 'Customer Promise Date' not in rename_dict.values():
            rename_dict[col_names[11]] = 'Customer Promise Date'  # Column L
        if len(col_names) > 12 and 'Projected Date' not in rename_dict.values():
            rename_dict[col_names[12]] = 'Projected Date'  # Column M
        if len(col_names) > 27 and 'Pending Approval Date' not in rename_dict.values():
            rename_dict[col_names[27]] = 'Pending Approval Date'  # Column AB
        
        sales_orders_df = sales_orders_df.rename(columns=rename_dict)
        
        # CRITICAL: Remove any duplicate columns that may have been created
        if sales_orders_df.columns.duplicated().any():
            st.sidebar.warning(f"⚠️ Removed duplicate columns in Sales Orders: {sales_orders_df.columns[sales_orders_df.columns.duplicated()].tolist()}")
            sales_orders_df = sales_orders_df.loc[:, ~sales_orders_df.columns.duplicated()]
        
        # Clean numeric values
        def clean_numeric_so(value):
            value_str = str(value).strip()
            if value_str == '' or value_str == 'nan' or value_str == 'None':
                return 0
            cleaned = value_str.replace(',', '').replace('$', '').replace(' ', '')
            try:
                return float(cleaned)
            except:
                return 0
        
        if 'Amount' in sales_orders_df.columns:
            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric_so)
        
        if 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].astype(str).str.strip()
        
        if 'Status' in sales_orders_df.columns:
            sales_orders_df['Status'] = sales_orders_df['Status'].astype(str).str.strip()
        
        # Convert date columns
        date_columns = ['Order Start Date', 'Customer Promise Date', 'Projected Date', 'Pending Approval Date']
        for col in date_columns:
            if col in sales_orders_df.columns:
                sales_orders_df[col] = pd.to_datetime(sales_orders_df[col], errors='coerce')
        
        # Filter to include Pending Approval, Pending Fulfillment, AND Pending Billing/Partially Fulfilled
        if 'Status' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
            ]
        
        # Calculate age for Old Pending Approval
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
        else:
            sales_orders_df['Age_Business_Days'] = 0
        
        # Remove rows without amount or sales rep
        if 'Amount' in sales_orders_df.columns and 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & 
                (sales_orders_df['Sales Rep'].notna()) & 
                (sales_orders_df['Sales Rep'] != '') &
                (sales_orders_df['Sales Rep'] != 'nan')
            ]
    else:
        st.warning("Could not find required columns in NS Sales Orders")
        sales_orders_df = pd.DataFrame()
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def calculate_team_metrics(deals_df, dashboard_df):
    """Calculate overall team metrics"""
    
    total_quota = dashboard_df['Quota'].sum()
    total_orders = dashboard_df['NetSuite Orders'].sum()
    
    # Filter for Q4 fulfillment only
    deals_q4 = deals_df[deals_df.get('Counts_In_Q4', True) == True]
    
    # Calculate Expect/Commit forecast (Q4 only)
    expect_commit = deals_q4[deals_q4['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    
    # Calculate Best Case/Opportunity (Q4 only)
    best_opp = deals_q4[deals_q4['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
    # Calculate Q1 spillover
    q1_spillover = deals_df[deals_df.get('Counts_In_Q4', True) == False]['Amount'].sum()
    
    # Calculate gap
    gap = total_quota - expect_commit - total_orders
    
    # Calculate attainment percentage
    current_forecast = expect_commit + total_orders
    attainment_pct = (current_forecast / total_quota * 100) if total_quota > 0 else 0
    
    # Potential attainment (if all deals close)
    potential_attainment = ((expect_commit + best_opp + total_orders) / total_quota * 100) if total_quota > 0 else 0
    
    return {
        'total_quota': total_quota,
        'total_orders': total_orders,
        'expect_commit': expect_commit,
        'best_opp': best_opp,
        'gap': gap,
        'attainment_pct': attainment_pct,
        'potential_attainment': potential_attainment,
        'current_forecast': current_forecast,
        'q1_spillover': q1_spillover
    }

def calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df=None):
    """Calculate metrics for a specific rep with detailed order lists for drill-down"""
    
    # Get rep's quota and orders
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    
    if rep_info.empty:
        return None
    
    quota = rep_info['Quota'].iloc[0]
    orders = rep_info['NetSuite Orders'].iloc[0]
    
    # Filter deals for this rep - ALL Q4 2025 deals (regardless of spillover)
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    
    # NEW: Check if we have the Q1 2026 Spillover column
    has_spillover_column = 'Q1 2026 Spillover' in rep_deals.columns
    
    if has_spillover_column:
        # Separate deals by shipping timeline
        rep_deals['Ships_In_Q4'] = rep_deals['Q1 2026 Spillover'] != 'Q1 2026'
        rep_deals['Ships_In_Q1'] = rep_deals['Q1 2026 Spillover'] == 'Q1 2026'
        
        # Deals that ship in Q4 2025
        rep_deals_ship_q4 = rep_deals[rep_deals['Ships_In_Q4'] == True].copy()
        
        # Deals that ship in Q1 2026 (spillover)
        rep_deals_ship_q1 = rep_deals[rep_deals['Ships_In_Q1'] == True].copy()
    else:
        # Fallback if column doesn't exist - treat all as Q4
        rep_deals_ship_q4 = rep_deals.copy()
        rep_deals_ship_q1 = pd.DataFrame()
    
    # Calculate metrics for DEALS SHIPPING IN Q4 (this counts toward quota)
    expect_commit_q4_deals = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Expect', 'Commit'])].copy()
    if expect_commit_q4_deals.columns.duplicated().any():
        expect_commit_q4_deals = expect_commit_q4_deals.loc[:, ~expect_commit_q4_deals.columns.duplicated()]
    expect_commit_q4 = expect_commit_q4_deals['Amount'].sum() if not expect_commit_q4_deals.empty else 0
    
    best_opp_q4_deals = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Best Case', 'Opportunity'])].copy()
    if best_opp_q4_deals.columns.duplicated().any():
        best_opp_q4_deals = best_opp_q4_deals.loc[:, ~best_opp_q4_deals.columns.duplicated()]
    best_opp_q4 = best_opp_q4_deals['Amount'].sum() if not best_opp_q4_deals.empty else 0
    
    # Calculate metrics for Q1 SPILLOVER DEALS (closing in Q4 but shipping in Q1)
    expect_commit_q1_deals = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Expect', 'Commit'])].copy()
    if expect_commit_q1_deals.columns.duplicated().any():
        expect_commit_q1_deals = expect_commit_q1_deals.loc[:, ~expect_commit_q1_deals.columns.duplicated()]
    expect_commit_q1_spillover = expect_commit_q1_deals['Amount'].sum() if not expect_commit_q1_deals.empty else 0
    
    best_opp_q1_deals = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Best Case', 'Opportunity'])].copy()
    if best_opp_q1_deals.columns.duplicated().any():
        best_opp_q1_deals = best_opp_q1_deals.loc[:, ~best_opp_q1_deals.columns.duplicated()]
    best_opp_q1_spillover = best_opp_q1_deals['Amount'].sum() if not best_opp_q1_deals.empty else 0
    
    # Total Q1 spillover
    q1_spillover_total = expect_commit_q1_spillover + best_opp_q1_spillover
    
    # Initialize detail dataframes for sales orders
    pending_approval_details = pd.DataFrame()
    pending_approval_no_date_details = pd.DataFrame()
    pending_approval_old_details = pd.DataFrame()
    pending_fulfillment_details = pd.DataFrame()
    pending_fulfillment_no_date_details = pd.DataFrame()
    
    # Calculate sales order metrics
    pending_approval = 0
    pending_approval_no_date = 0
    pending_approval_old = 0
    pending_fulfillment = 0
    pending_fulfillment_no_date = 0
    
    if sales_orders_df is not None and not sales_orders_df.empty:
        # Make a clean copy and ensure no duplicate columns
        rep_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
        
        # Remove duplicate columns if any exist
        if rep_orders.columns.duplicated().any():
            rep_orders = rep_orders.loc[:, ~rep_orders.columns.duplicated()]
        
        # PENDING APPROVAL LOGIC
        pending_approval_orders = rep_orders[rep_orders['Status'] == 'Pending Approval'].copy()
        
        if not pending_approval_orders.empty:
            # Define Q4 2025 date range
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Pending Approval WITH dates (Column AB has a date AND it's in Q4 2025)
            if 'Pending Approval Date' in pending_approval_orders.columns:
                pa_with_date_mask = (
                    (pending_approval_orders['Pending Approval Date'].notna()) &
                    (pending_approval_orders['Pending Approval Date'] != 'No Date') &
                    (pending_approval_orders['Pending Approval Date'] >= q4_start) &
                    (pending_approval_orders['Pending Approval Date'] <= q4_end)
                )
                pending_approval_details = pending_approval_orders[pa_with_date_mask].copy()
                # Remove duplicate columns
                if pending_approval_details.columns.duplicated().any():
                    pending_approval_details = pending_approval_details.loc[:, ~pending_approval_details.columns.duplicated()]
                pending_approval = pending_approval_details['Amount'].sum() if not pending_approval_details.empty else 0
                
                # Pending Approval WITHOUT dates OR outside Q4
                pa_no_date_mask = ~pa_with_date_mask
                pending_approval_no_date_details = pending_approval_orders[pa_no_date_mask].copy()
                # Remove duplicate columns
                if pending_approval_no_date_details.columns.duplicated().any():
                    pending_approval_no_date_details = pending_approval_no_date_details.loc[:, ~pending_approval_no_date_details.columns.duplicated()]
                pending_approval_no_date = pending_approval_no_date_details['Amount'].sum() if not pending_approval_no_date_details.empty else 0
            
            # Old Pending Approval (Order Start Date > 14 business days)
            if 'Age_Business_Days' in pending_approval_orders.columns:
                old_mask = pending_approval_orders['Age_Business_Days'] > 14
                pending_approval_old_details = pending_approval_orders[old_mask].copy()
                # Remove duplicate columns
                if pending_approval_old_details.columns.duplicated().any():
                    pending_approval_old_details = pending_approval_old_details.loc[:, ~pending_approval_old_details.columns.duplicated()]
                pending_approval_old = pending_approval_old_details['Amount'].sum() if not pending_approval_old_details.empty else 0
        
        # PENDING FULFILLMENT LOGIC
        fulfillment_orders = rep_orders[
            rep_orders['Status'].isin(['Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
        ].copy()
        
        if not fulfillment_orders.empty:
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Check if Customer Promise Date OR Projected Date is in Q4
            def has_q4_date(row):
                if pd.notna(row.get('Customer Promise Date')):
                    if q4_start <= row['Customer Promise Date'] <= q4_end:
                        return True
                if pd.notna(row.get('Projected Date')):
                    if q4_start <= row['Projected Date'] <= q4_end:
                        return True
                return False
            
            fulfillment_orders['Has_Q4_Date'] = fulfillment_orders.apply(has_q4_date, axis=1)
            
            # Pending Fulfillment WITH Q4 dates
            pending_fulfillment_details = fulfillment_orders[fulfillment_orders['Has_Q4_Date'] == True].copy()
            # Remove duplicate columns
            if pending_fulfillment_details.columns.duplicated().any():
                pending_fulfillment_details = pending_fulfillment_details.loc[:, ~pending_fulfillment_details.columns.duplicated()]
            pending_fulfillment = pending_fulfillment_details['Amount'].sum() if not pending_fulfillment_details.empty else 0
            
            # Pending Fulfillment WITHOUT dates
            no_date_mask = (
                (fulfillment_orders['Customer Promise Date'].isna()) &
                (fulfillment_orders['Projected Date'].isna())
            )
            pending_fulfillment_no_date_details = fulfillment_orders[no_date_mask].copy()
            # Remove duplicate columns
            if pending_fulfillment_no_date_details.columns.duplicated().any():
                pending_fulfillment_no_date_details = pending_fulfillment_no_date_details.loc[:, ~pending_fulfillment_no_date_details.columns.duplicated()]
            pending_fulfillment_no_date = pending_fulfillment_no_date_details['Amount'].sum() if not pending_fulfillment_no_date_details.empty else 0
    
    # Total calculations - ONLY Q4 SHIPPING DEALS COUNT TOWARD QUOTA
    total_pending_fulfillment = pending_fulfillment + pending_fulfillment_no_date
    total_progress = orders + expect_commit_q4 + pending_approval + pending_fulfillment
    gap = quota - total_progress
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
    potential_attainment = ((total_progress + best_opp_q4) / quota * 100) if quota > 0 else 0
    
    return {
        'quota': quota,
        'orders': orders,
        'expect_commit': expect_commit_q4,  # Only Q4 shipping deals
        'best_opp': best_opp_q4,  # Only Q4 shipping deals
        'gap': gap,
        'attainment_pct': attainment_pct,
        'potential_attainment': potential_attainment,
        'total_progress': total_progress,
        'pending_approval': pending_approval,
        'pending_approval_no_date': pending_approval_no_date,
        'pending_approval_old': pending_approval_old,
        'pending_fulfillment': pending_fulfillment,
        'pending_fulfillment_no_date': pending_fulfillment_no_date,
        'total_pending_fulfillment': total_pending_fulfillment,
        
        # NEW: Q1 Spillover metrics
        'q1_spillover_expect_commit': expect_commit_q1_spillover,
        'q1_spillover_best_opp': best_opp_q1_spillover,
        'q1_spillover_total': q1_spillover_total,
        
        # ALL Q4 2025 closing deals (for reference)
        'total_q4_closing_deals': len(rep_deals),
        'total_q4_closing_amount': rep_deals['Amount'].sum() if not rep_deals.empty else 0,
        
        'deals': rep_deals_ship_q4,  # Deals shipping in Q4
        
        # Add detail dataframes for drill-down
        'pending_approval_details': pending_approval_details,
        'pending_approval_no_date_details': pending_approval_no_date_details,
        'pending_approval_old_details': pending_approval_old_details,
        'pending_fulfillment_details': pending_fulfillment_details,
        'pending_fulfillment_no_date_details': pending_fulfillment_no_date_details,
        'expect_commit_deals': expect_commit_q4_deals,
        'best_opp_deals': best_opp_q4_deals,
        
        # NEW: Q1 Spillover deal details
        'expect_commit_q1_spillover_deals': expect_commit_q1_deals,
        'best_opp_q1_spillover_deals': best_opp_q1_deals,
        'all_q1_spillover_deals': rep_deals_ship_q1
    }

def create_gap_chart(metrics, title):
    """Create a waterfall/combo chart showing progress to goal"""
    
    fig = go.Figure()
    
    # Create stacked bar
    fig.add_trace(go.Bar(
        name='NetSuite Orders',
        x=['Progress'],
        y=[metrics['total_orders'] if 'total_orders' in metrics else metrics['orders']],
        marker_color='#1E88E5',
        text=[f"${metrics['total_orders'] if 'total_orders' in metrics else metrics['orders']:,.0f}"],
        textposition='inside'
    ))
    
    fig.add_trace(go.Bar(
        name='Expect/Commit',
        x=['Progress'],
        y=[metrics['expect_commit']],
        marker_color='#43A047',
        text=[f"${metrics['expect_commit']:,.0f}"],
        textposition='inside'
    ))
    
    # Add quota line
    fig.add_trace(go.Scatter(
        name='Quota Goal',
        x=['Progress'],
        y=[metrics['total_quota'] if 'total_quota' in metrics else metrics['quota']],
        mode='markers',
        marker=dict(size=12, color='#DC3912', symbol='diamond'),
        text=[f"Goal: ${metrics['total_quota'] if 'total_quota' in metrics else metrics['quota']:,.0f}"],
        textposition='top center'
    ))
    
    # Add potential attainment line
    potential = metrics['expect_commit'] + metrics['best_opp'] + (metrics['total_orders'] if 'total_orders' in metrics else metrics['orders'])
    fig.add_trace(go.Scatter(
        name='Potential (if all deals close)',
        x=['Progress'],
        y=[potential],
        mode='markers',
        marker=dict(size=12, color='#FB8C00', symbol='diamond'),
        text=[f"Potential: ${potential:,.0f}"],
        textposition='bottom center'
    ))
    
    fig.update_layout(
        title=title,
        barmode='stack',
        height=400,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis_title="Amount ($)",
        xaxis_title="",
        hovermode='x unified'
    )
    
    return fig

def create_status_breakdown_chart(deals_df, rep_name=None):
    """Create a pie chart showing deal distribution by status"""
    
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    
    # Only show Q4 deals
    deals_df = deals_df[deals_df.get('Counts_In_Q4', True) == True]
    
    if deals_df.empty:
        return None
    
    status_summary = deals_df.groupby('Status')['Amount'].sum().reset_index()
    
    color_map = {
        'Expect': '#1E88E5',
        'Commit': '#43A047',
        'Best Case': '#FB8C00',
        'Opportunity': '#8E24AA'
    }
    
    fig = px.pie(
        status_summary,
        values='Amount',
        names='Status',
        title='Deal Amount by Forecast Category (Q4 Only)',
        color='Status',
        color_discrete_map=color_map,
        hole=0.4
    )
    
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=400)
    
    return fig

def create_pipeline_breakdown_chart(deals_df, rep_name=None):
    """Create a stacked bar chart showing pipeline breakdown"""
    
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    
    # Only show Q4 deals
    deals_df = deals_df[deals_df.get('Counts_In_Q4', True) == True]
    
    if deals_df.empty:
        return None
    
    # Group by pipeline and status
    pipeline_summary = deals_df.groupby(['Pipeline', 'Status'])['Amount'].sum().reset_index()
    
    color_map = {
        'Expect': '#1E88E5',
        'Commit': '#43A047',
        'Best Case': '#FB8C00',
        'Opportunity': '#8E24AA'
    }
    
    fig = px.bar(
        pipeline_summary,
        x='Pipeline',
        y='Amount',
        color='Status',
        title='Pipeline Breakdown by Forecast Category (Q4 Only)',
        color_discrete_map=color_map,
        text_auto='.2s',
        barmode='stack'
    )
    
    fig.update_layout(
        height=400,
        yaxis_title="Amount ($)",
        xaxis_title="Pipeline"
    )
    
    return fig

def create_deals_timeline(deals_df, rep_name=None):
    """Create a timeline showing when deals are expected to close"""
    
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    
    # Filter out deals without close dates
    timeline_df = deals_df[deals_df['Close Date'].notna()].copy()
    
    if timeline_df.empty:
        return None
    
    # Sort by close date
    timeline_df = timeline_df.sort_values('Close Date')
    
    # Add Q4/Q1 indicator to color map
    timeline_df['Quarter'] = timeline_df.apply(
        lambda x: 'Q4 2025' if x.get('Counts_In_Q4', True) else 'Q1 2026', 
        axis=1
    )
    
    color_map = {
        'Expect': '#1E88E5',
        'Commit': '#43A047',
        'Best Case': '#FB8C00',
        'Opportunity': '#8E24AA'
    }
    
    fig = px.scatter(
        timeline_df,
        x='Close Date',
        y='Amount',
        color='Status',
        size='Amount',
        hover_data=['Deal Name', 'Amount', 'Pipeline', 'Quarter'],
        title='Deal Close Date Timeline',
        color_discrete_map=color_map
    )
    
    # Fixed: Use datetime object for the vertical line
    from datetime import datetime
    q4_boundary = datetime(2025, 12, 31)
    
    try:
        fig.add_vline(
            x=q4_boundary, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Q4/Q1 Boundary"
        )
    except:
        pass
    
    fig.update_layout(
        height=400,
        yaxis_title="Deal Amount ($)",
        xaxis_title="Expected Close Date"
    )
    
    return fig

def create_invoice_status_chart(invoices_df, rep_name=None):
    """Create a chart showing invoice breakdown by status"""
    
    if invoices_df.empty:
        return None
    
    if rep_name:
        invoices_df = invoices_df[invoices_df['Sales Rep'] == rep_name]
    
    if invoices_df.empty:
        return None
    
    status_summary = invoices_df.groupby('Status')['Amount'].sum().reset_index()
    
    fig = px.pie(
        status_summary,
        values='Amount',
        names='Status',
        title='Invoice Amount by Status',
        hole=0.4
    )
    
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=400)
    
    return fig

def display_drill_down_section(title, amount, details_df, key_suffix):
    """Display a collapsible section with order details - FIXED for duplicate columns"""
    
    with st.expander(f"{title}: ${amount:,.2f} (Click to see {len(details_df)} items)"):
        if not details_df.empty:
            # DEBUG: Check for duplicate columns
            if details_df.columns.duplicated().any():
                st.warning(f"⚠️ Duplicate columns detected: {details_df.columns[details_df.columns.duplicated()].tolist()}")
                # Remove duplicates
                details_df = details_df.loc[:, ~details_df.columns.duplicated()]
            
            # Select relevant columns to display
            display_cols = []
            
            # Prioritize important columns based on dataframe type
            if 'Deal Name' in details_df.columns:  # HubSpot deals
                priority_cols = ['Deal Name', 'Amount', 'Status', 'Pipeline', 'Close Date', 'Product Type']
            else:  # Sales Orders
                priority_cols = ['Document Number', 'Customer', 'Amount', 'Status', 
                               'Order Start Date', 'Pending Approval Date', 
                               'Customer Promise Date', 'Projected Date']
            
            for col in priority_cols:
                if col in details_df.columns and col not in display_cols:
                    display_cols.append(col)
            
            # Limit to available columns (no duplicates)
            display_cols_unique = []
            for col in display_cols:
                if col not in display_cols_unique:
                    display_cols_unique.append(col)
            
            if display_cols_unique:
                try:
                    display_df = details_df[display_cols_unique].copy()
                    
                    # Double-check for duplicates in display_df
                    if display_df.columns.duplicated().any():
                        st.error(f"Still have duplicates: {display_df.columns.tolist()}")
                        display_df = display_df.loc[:, ~display_df.columns.duplicated()]
                    
                    # Format amount column
                    if 'Amount' in display_df.columns:
                        display_df['Amount'] = display_df['Amount'].apply(lambda x: f"${x:,.2f}")
                    
                    # Format date columns
                    for col in display_df.columns:
                        if 'Date' in col and pd.api.types.is_datetime64_any_dtype(display_df[col]):
                            display_df[col] = display_df[col].dt.strftime('%Y-%m-%d')
                    
                    # Final check before display
                    if display_df.columns.duplicated().any():
                        st.error("Cannot display: duplicate columns persist")
                        st.write(f"Columns: {display_df.columns.tolist()}")
                    else:
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
                        
                        # Summary statistics
                        st.caption(f"Total: ${details_df['Amount'].sum():,.2f} | Count: {len(details_df)} items")
                except Exception as e:
                    st.error(f"Error displaying data: {str(e)}")
                    st.write(f"Available columns: {details_df.columns.tolist()}")
                    st.write(f"Attempted to display: {display_cols_unique}")
            else:
                st.info("No columns available to display")
        else:
            st.info("No items in this category")

def display_progress_breakdown(metrics):
    """Display a beautiful progress breakdown card"""
    
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>💰 Section 1: Q4 Gap to Goal</h3>
        <div class="progress-item">
            <span class="progress-label">📦 Invoiced (Orders Shipped)</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📤 Pending Fulfillment (with dates)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with dates)</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">✅ HubSpot Expect/Commit (Q4)</span>
            <span class="progress-value">${metrics['expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 SECTION 1 TOTAL</span>
            <span class="progress-value">${metrics['total_progress']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Add attainment info below
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Current Attainment", f"{metrics['attainment_pct']:.1f}%", 
                 delta=f"${metrics['total_progress']:,.0f} of ${metrics['quota']:,.0f}")
    with col2:
        st.metric("Potential with Upside", f"{metrics['potential_attainment']:.1f}%",
                 delta=f"+${metrics['best_opp']:,.0f} Best Case/Opp")

def display_reconciliation_view(deals_df, dashboard_df, sales_orders_df):
    """Show a reconciliation view to compare with boss's numbers"""
    
    st.title("🔍 Forecast Reconciliation with Boss's Numbers")
    
    # Boss's Q4 numbers from the LATEST screenshot
    boss_rep_numbers = {
        'Jake Lynch': {
            'invoiced': 577540,
            'pending_fulfillment': 263183,
            'pending_approval': 45220,
            'hubspot': 340756,
            'total': 1226699,
            'pending_fulfillment_so_no_date': 87891,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 33741,
            'total_q4': 1350638
        },
        'Dave Borkowski': {
            'invoiced': 237849,
            'pending_fulfillment': 160537,
            'pending_approval': 13390,
            'hubspot': 414768,
            'total': 826545,
            'pending_fulfillment_so_no_date': 45471,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 12244,
            'total_q4': 884260
        },
        'Alex Gonzalez': {
            'invoiced': 314523,
            'pending_fulfillment': 190865,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 505387,
            'pending_fulfillment_so_no_date': 41710,
            'pending_approval_so_no_date': 79361,
            'old_pending_approval': 4900,
            'total_q4': 631358
        },
        'Brad Sherman': {
            'invoiced': 118211,
            'pending_fulfillment': 38330,
            'pending_approval': 28984,
            'hubspot': 183103,
            'total': 368629,
            'pending_fulfillment_so_no_date': 29970,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 884,
            'total_q4': 399482
        },
        'Lance Mitton': {
            'invoiced': 23948,
            'pending_fulfillment': 2027,
            'pending_approval': 3331,
            'hubspot': 11000,
            'total': 38356,
            'pending_fulfillment_so_no_date': 1613,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 60527,
            'total_q4': 100496
        },
        'House': {
            'invoiced': 0,
            'pending_fulfillment': 899,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 0,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 0,
            'total_q4': 0
        },
        'Shopify ECommerce': {
            'invoiced': 21348,
            'pending_fulfillment': 0,
            'pending_approval': 1174,
            'hubspot': 0,
            'total': 23421,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1544,
            'total_q4': 24965
        }
    }
    
    # Tab selection for Rep vs Pipeline view
    tab1, tab2 = st.tabs(["By Rep", "By Pipeline"])
    
    with tab1:
        st.markdown('<div class="section-header">Section 1: Q4 Gap to Goal</div>', unsafe_allow_html=True)
        
        # Create the comparison table
        comparison_data = []
        totals = {
            'invoiced_you': 0, 'invoiced_boss': 0,
            'pf_you': 0, 'pf_boss': 0,
            'pa_you': 0, 'pa_boss': 0,
            'hs_you': 0, 'hs_boss': 0,
            'total_you': 0, 'total_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            metrics = None
            if rep_name in dashboard_df['Rep Name'].values:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics or rep_name in ['House', 'Shopify ECommerce']:
                boss = boss_rep_numbers[rep_name]
                
                # Get your values
                your_invoiced = metrics['orders'] if metrics else 0
                your_pf = metrics['pending_fulfillment'] if metrics else 0
                your_pa = metrics['pending_approval'] if metrics else 0
                your_hs = metrics['expect_commit'] if metrics else 0
                your_total = your_invoiced + your_pf + your_pa + your_hs
                
                # Update totals
                totals['invoiced_you'] += your_invoiced
                totals['invoiced_boss'] += boss['invoiced']
                totals['pf_you'] += your_pf
                totals['pf_boss'] += boss['pending_fulfillment']
                totals['pa_you'] += your_pa
                totals['pa_boss'] += boss['pending_approval']
                totals['hs_you'] += your_hs
                totals['hs_boss'] += boss['hubspot']
                totals['total_you'] += your_total
                totals['total_boss'] += boss['total']
                
                comparison_data.append({
                    'Rep': rep_name,
                    'Invoiced': f"${your_invoiced:,.0f}",
                    'Invoiced (Boss)': f"${boss['invoiced']:,.0f}",
                    'Pending Fulfillment': f"${your_pf:,.0f}",
                    'PF (Boss)': f"${boss['pending_fulfillment']:,.0f}",
                    'Pending Approval': f"${your_pa:,.0f}",
                    'PA (Boss)': f"${boss['pending_approval']:,.0f}",
                    'HubSpot Expect/Commit': f"${your_hs:,.0f}",
                    'HS (Boss)': f"${boss['hubspot']:,.0f}",
                    'Total': f"${your_total:,.0f}",
                    'Total (Boss)': f"${boss['total']:,.0f}"
                })
        
        # Add totals row
        comparison_data.append({
            'Rep': 'TOTAL',
            'Invoiced': f"${totals['invoiced_you']:,.0f}",
            'Invoiced (Boss)': f"${totals['invoiced_boss']:,.0f}",
            'Pending Fulfillment': f"${totals['pf_you']:,.0f}",
            'PF (Boss)': f"${totals['pf_boss']:,.0f}",
            'Pending Approval': f"${totals['pa_you']:,.0f}",
            'PA (Boss)': f"${totals['pa_boss']:,.0f}",
            'HubSpot Expect/Commit': f"${totals['hs_you']:,.0f}",
            'HS (Boss)': f"${totals['hs_boss']:,.0f}",
            'Total': f"${totals['total_you']:,.0f}",
            'Total (Boss)': f"${totals['total_boss']:,.0f}"
        })
        
        if comparison_data:
            comparison_df = pd.DataFrame(comparison_data)
            st.dataframe(comparison_df, use_container_width=True, hide_index=True)
        
        # Section 2: Additional Orders
        st.markdown('<div class="section-header">Section 2: Additional Orders (Can be included)</div>', unsafe_allow_html=True)
        
        additional_data = []
        additional_totals = {
            'pf_no_date_you': 0, 'pf_no_date_boss': 0,
            'pa_no_date_you': 0, 'pa_no_date_boss': 0,
            'old_pa_you': 0, 'old_pa_boss': 0,
            'final_you': 0, 'final_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            metrics = None
            if rep_name in dashboard_df['Rep Name'].values:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics or rep_name in ['House', 'Shopify ECommerce']:
                boss = boss_rep_numbers[rep_name]
                
                # Calculate additional metrics
                your_pf_no_date = metrics['pending_fulfillment_no_date'] if metrics else 0
                your_pa_no_date = metrics['pending_approval_no_date'] if metrics else 0
                your_old_pa = metrics['pending_approval_old'] if metrics else 0
                
                # Calculate final total
                section1_total = (metrics['orders'] + metrics['pending_fulfillment'] + 
                                 metrics['pending_approval'] + metrics['expect_commit']) if metrics else 0
                your_final = section1_total + your_pf_no_date + your_pa_no_date + your_old_pa
                
                # Update totals
                additional_totals['pf_no_date_you'] += your_pf_no_date
                additional_totals['pf_no_date_boss'] += boss['pending_fulfillment_so_no_date']
                additional_totals['pa_no_date_you'] += your_pa_no_date
                additional_totals['pa_no_date_boss'] += boss['pending_approval_so_no_date']
                additional_totals['old_pa_you'] += your_old_pa
                additional_totals['old_pa_boss'] += boss['old_pending_approval']
                additional_totals['final_you'] += your_final
                additional_totals['final_boss'] += boss['total_q4']
                
                additional_data.append({
                    'Rep': rep_name,
                    'PF SO\'s No Date': f"${your_pf_no_date:,.0f}",
                    'PF No Date (Boss)': f"${boss['pending_fulfillment_so_no_date']:,.0f}",
                    'PA SO\'s No Date': f"${your_pa_no_date:,.0f}",
                    'PA No Date (Boss)': f"${boss['pending_approval_so_no_date']:,.0f}",
                    'Old PA (>2 weeks)': f"${your_old_pa:,.0f}",
                    'Old PA (Boss)': f"${boss['old_pending_approval']:,.0f}",
                    'Total Q4': f"${your_final:,.0f}",
                    'Total Q4 (Boss)': f"${boss['total_q4']:,.0f}"
                })
        
        # Add totals row
        additional_data.append({
            'Rep': 'TOTAL',
            'PF SO\'s No Date': f"${additional_totals['pf_no_date_you']:,.0f}",
            'PF No Date (Boss)': f"${additional_totals['pf_no_date_boss']:,.0f}",
            'PA SO\'s No Date': f"${additional_totals['pa_no_date_you']:,.0f}",
            'PA No Date (Boss)': f"${additional_totals['pa_no_date_boss']:,.0f}",
            'Old PA (>2 weeks)': f"${additional_totals['old_pa_you']:,.0f}",
            'Old PA (Boss)': f"${additional_totals['old_pa_boss']:,.0f}",
            'Total Q4': f"${additional_totals['final_you']:,.0f}",
            'Total Q4 (Boss)': f"${additional_totals['final_boss']:,.0f}"
        })
        
        if additional_data:
            additional_df = pd.DataFrame(additional_data)
            st.dataframe(additional_df, use_container_width=True, hide_index=True)
    
    with tab2:
        st.markdown("### Pipeline-Level Comparison")
        st.info("Pipeline breakdown in development - need to map invoices and sales orders to pipelines")
    
    # Summary
    st.markdown("### 📊 Key Insights")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        diff = totals['total_boss'] - totals['total_you']
        st.metric("Section 1 Variance", f"${abs(diff):,.0f}", 
                 delta=f"{'Under' if diff > 0 else 'Over'} by ${abs(diff):,.0f}")
    
    with col2:
        final_diff = additional_totals['final_boss'] - additional_totals['final_you']
        st.metric("Total Q4 Variance", f"${abs(final_diff):,.0f}",
                 delta=f"{'Under' if final_diff > 0 else 'Over'} by ${abs(final_diff):,.0f}")
    
    with col3:
        accuracy = (1 - abs(final_diff) / additional_totals['final_boss']) * 100 if additional_totals['final_boss'] > 0 else 0
        st.metric("Accuracy", f"{accuracy:.1f}%")

def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display the team-level dashboard"""
    
    st.title("🎯 Team Sales Dashboard - Q4 2025")
    
    # Calculate metrics
    metrics = calculate_team_metrics(deals_df, dashboard_df)
    
    # Display Q1 spillover warning if applicable
    if metrics.get('q1_spillover', 0) > 0:
        st.warning(
            f"📅 **Q1 2026 Spillover**: ${metrics['q1_spillover']:,.0f} in deals will close in late December "
            f"but ship in Q1 2026 based on product lead times"
        )
    
    # Display key metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Total Quota",
            value=f"${metrics['total_quota']:,.0f}",
            delta=None
        )
    
    with col2:
        st.metric(
            label="Current Forecast",
            value=f"${metrics['current_forecast']:,.0f}",
            delta=f"{metrics['attainment_pct']:.1f}% of quota"
        )
    
    with col3:
        st.metric(
            label="Gap to Goal",
            value=f"${metrics['gap']:,.0f}",
            delta=f"${-metrics['gap']:,.0f}" if metrics['gap'] < 0 else None,
            delta_color="inverse"
        )
    
    with col4:
        st.metric(
            label="Potential Attainment",
            value=f"{metrics['potential_attainment']:.1f}%",
            delta=f"+{metrics['potential_attainment'] - metrics['attainment_pct']:.1f}% upside"
        )
    
    # Progress bar
    st.markdown("### 📈 Progress to Quota")
    progress = min(metrics['attainment_pct'] / 100, 1.0)
    st.progress(progress)
    st.caption(f"Current: {metrics['attainment_pct']:.1f}% | Potential: {metrics['potential_attainment']:.1f}%")
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        gap_chart = create_gap_chart(metrics, "Team Progress to Goal")
        st.plotly_chart(gap_chart, use_container_width=True)
    
    with col2:
        status_chart = create_status_breakdown_chart(deals_df)
        if status_chart:
            st.plotly_chart(status_chart, use_container_width=True)
        else:
            st.info("No deal data available for status breakdown")
    
    # Pipeline breakdown
    st.markdown("### 🔄 Pipeline Analysis")
    pipeline_chart = create_pipeline_breakdown_chart(deals_df)
    if pipeline_chart:
        st.plotly_chart(pipeline_chart, use_container_width=True)
    else:
        st.info("No deal data available for pipeline analysis")
    
    # Timeline
    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    else:
        st.info("No deal data available for timeline")
    
    # Invoice Status Breakdown
    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Status Breakdown")
        invoice_chart = create_invoice_status_chart(invoices_df)
        if invoice_chart:
            st.plotly_chart(invoice_chart, use_container_width=True)
    
    # Rep summary table
    st.markdown("### 👥 Rep Summary")
    
    rep_summary = []
    for rep_name in dashboard_df['Rep Name']:
        rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
        if rep_metrics:
            rep_summary.append({
                'Rep': rep_name,
                'Quota': f"${rep_metrics['quota']:,.0f}",
                'Orders': f"${rep_metrics['orders']:,.0f}",
                'Expect/Commit': f"${rep_metrics['expect_commit']:,.0f}",
                'Pending Approval': f"${rep_metrics['pending_approval']:,.0f}",
                'Pending Fulfillment': f"${rep_metrics['pending_fulfillment']:,.0f}",
                'Total Progress': f"${rep_metrics['total_progress']:,.0f}",
                'Gap': f"${rep_metrics['gap']:,.0f}",
                'Attainment': f"{rep_metrics['attainment_pct']:.1f}%",
                'Q1 Spillover': f"${rep_metrics['q1_spillover']:,.0f}"
            })
    
    if rep_summary:
        rep_summary_df = pd.DataFrame(rep_summary)
        st.dataframe(rep_summary_df, use_container_width=True, hide_index=True)
    else:
        st.warning("No rep data available")

def display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display individual rep dashboard with drill-down capability"""
    
    st.title(f"👤 {rep_name}'s Q4 2025 Forecast")
    
    # Calculate metrics with details
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
    
    if not metrics:
        st.error(f"No data found for {rep_name}")
        return
    
    # Show summary of ALL Q4 2025 deals (closing in Q4)
    if metrics.get('total_q4_closing_deals', 0) > 0:
        st.info(
            f"📋 **Total Q4 2025 Pipeline**: {metrics['total_q4_closing_deals']} deals worth "
            f"${metrics['total_q4_closing_amount']:,.0f} closing in Q4 2025. "
            f"Of these, ${metrics.get('q1_spillover_total', 0):,.0f} will ship in Q1 2026 based on lead times."
        )
    
    # Display key metrics - Section 1
    st.markdown("### 💰 Section 1: Q4 Gap to Goal Components")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            label="Quota",
            value=f"${metrics['quota']/1000:.0f}K" if metrics['quota'] < 1000000 else f"${metrics['quota']/1000000:.1f}M"
        )
    
    with col2:
        st.metric(
            label="Invoiced",
            value=f"${metrics['orders']/1000:.0f}K" if metrics['orders'] < 1000000 else f"${metrics['orders']/1000000:.1f}M",
            help="Orders shipped and invoiced - click sections below for details"
        )
    
    with col3:
        st.metric(
            label="Pending Fulfillment",
            value=f"${metrics['pending_fulfillment']/1000:.0f}K",
            help="Click sections below to see order details"
        )
    
    with col4:
        st.metric(
            label="Pending Approval",
            value=f"${metrics['pending_approval']/1000:.0f}K",
            help="Click sections below to see order details"
        )
    
    with col5:
        st.metric(
            label="HubSpot (Q4)",
            value=f"${metrics['expect_commit']/1000:.0f}K",
            help="Click sections below to see deal details"
        )
    
    with col6:
        st.metric(
            label="Gap to Goal",
            value=f"${metrics['gap']/1000:.0f}K",
            delta=f"${-metrics['gap']/1000:.0f}K" if metrics['gap'] < 0 else None,
            delta_color="inverse"
        )
    
    # Beautiful progress breakdown
    display_progress_breakdown(metrics)
    
    # Drill-down sections for Section 1
    st.markdown("#### 📊 Section 1 Detail Drill-Downs")
    
    col1, col2 = st.columns(2)
    
    with col1:
        display_drill_down_section(
            "📤 Pending Fulfillment (with Q4 dates)",
            metrics['pending_fulfillment'],
            metrics.get('pending_fulfillment_details', pd.DataFrame()),
            f"{rep_name}_pf"
        )
        
        display_drill_down_section(
            "⏳ Pending Approval (with dates)",
            metrics['pending_approval'],
            metrics.get('pending_approval_details', pd.DataFrame()),
            f"{rep_name}_pa"
        )
    
    with col2:
        display_drill_down_section(
            "✅ HubSpot Expect/Commit Deals",
            metrics['expect_commit'],
            metrics.get('expect_commit_deals', pd.DataFrame()),
            f"{rep_name}_hs"
        )
        
        display_drill_down_section(
            "🎯 Best Case/Opportunity Deals",
            metrics['best_opp'],
            metrics.get('best_opp_deals', pd.DataFrame()),
            f"{rep_name}_bo"
        )
    
    # Section 2: Additional Orders
    st.markdown("### 📊 Section 2: Additional Orders (Can be included)")
    
    warning_col1, warning_col2, warning_col3 = st.columns(3)
    
    with warning_col1:
        st.metric(
            label="PF SO's No Date",
            value=f"${metrics['pending_fulfillment_no_date']:,.0f}",
            help="Click below for details"
        )
        display_drill_down_section(
            "PF Orders Without Dates",
            metrics['pending_fulfillment_no_date'],
            metrics.get('pending_fulfillment_no_date_details', pd.DataFrame()),
            f"{rep_name}_pf_no_date"
        )
    
    with warning_col2:
        st.metric(
            label="PA SO's No Date",
            value=f"${metrics['pending_approval_no_date']:,.0f}",
            help="Click below for details"
        )
        display_drill_down_section(
            "PA Orders Without Dates",
            metrics['pending_approval_no_date'],
            metrics.get('pending_approval_no_date_details', pd.DataFrame()),
            f"{rep_name}_pa_no_date"
        )
    
    with warning_col3:
        st.metric(
            label="Old PA (>2 weeks)",
            value=f"${metrics['pending_approval_old']:,.0f}",
            help="Click below for details",
            delta="Needs attention" if metrics['pending_approval_old'] > 0 else None,
            delta_color="off" if metrics['pending_approval_old'] > 0 else "normal"
        )
        display_drill_down_section(
            "Old Pending Approval Orders",
            metrics['pending_approval_old'],
            metrics.get('pending_approval_old_details', pd.DataFrame()),
            f"{rep_name}_pa_old"
        )
    
    # NEW: Section 3 - Q1 2026 Spillover
    st.markdown("### 🚢 Section 3: Q1 2026 Spillover (Closing in Q4, Shipping in Q1)")
    
    st.info(f"💡 These deals close in Q4 2025 but will ship in Q1 2026 based on product lead times")
    
    spillover_col1, spillover_col2, spillover_col3 = st.columns(3)
    
    with spillover_col1:
        st.metric(
            label="Expect/Commit (Q1 Ship)",
            value=f"${metrics.get('q1_spillover_expect_commit', 0):,.0f}",
            help="Deals closing in Q4 but shipping in Q1"
        )
        display_drill_down_section(
            "Expect/Commit Deals Shipping Q1 2026",
            metrics.get('q1_spillover_expect_commit', 0),
            metrics.get('expect_commit_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_ec_q1"
        )
    
    with spillover_col2:
        st.metric(
            label="Best Case/Opp (Q1 Ship)",
            value=f"${metrics.get('q1_spillover_best_opp', 0):,.0f}",
            help="Potential deals closing in Q4 but shipping in Q1"
        )
        display_drill_down_section(
            "Best Case/Opp Deals Shipping Q1 2026",
            metrics.get('q1_spillover_best_opp', 0),
            metrics.get('best_opp_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_bo_q1"
        )
    
    with spillover_col3:
        st.metric(
            label="Total Q1 Spillover",
            value=f"${metrics.get('q1_spillover_total', 0):,.0f}",
            help="All deals shipping in Q1 2026"
        )
        display_drill_down_section(
            "All Q1 2026 Spillover Deals",
            metrics.get('q1_spillover_total', 0),
            metrics.get('all_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_all_q1"
        )
    
    # Final Total
    final_total = (metrics['total_progress'] + metrics['pending_fulfillment_no_date'] + 
                   metrics['pending_approval_no_date'] + metrics['pending_approval_old'])
    st.metric(
        label="📊 FINAL TOTAL Q4",
        value=f"${final_total:,.0f}",
        delta=f"Section 1: ${metrics['total_progress']:,.0f} + Section 2: ${final_total - metrics['total_progress']:,.0f}"
    )
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        gap_chart = create_gap_chart(metrics, f"{rep_name}'s Progress to Goal")
        st.plotly_chart(gap_chart, use_container_width=True)
    
    with col2:
        status_chart = create_status_breakdown_chart(deals_df, rep_name)
        if status_chart:
            st.plotly_chart(status_chart, use_container_width=True)
        else:
            st.info("No deal data available for this rep")
    
    # Pipeline breakdown
    st.markdown("### 🔄 Pipeline Analysis")
    pipeline_chart = create_pipeline_breakdown_chart(deals_df, rep_name)
    if pipeline_chart:
        st.plotly_chart(pipeline_chart, use_container_width=True)
    else:
        st.info("No deal data available for this rep")
    
    # Timeline
    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df, rep_name)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    else:
        st.info("No deal data available for this rep")

# Main app
def main():
    
    # Sidebar
    with st.sidebar:
        # Display company name
        st.markdown("""
        <div style="text-align: center; padding: 20px;">
            <h2>CALYX</h2>
            <p style="font-size: 12px; letter-spacing: 3px;">CONTAINERS</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        
        st.markdown("### 🎯 Dashboard Navigation")
        view_mode = st.radio(
            "Select View:",
            ["Team Overview", "Individual Rep", "Reconciliation"],
            label_visibility="collapsed"
        )
        
        st.markdown("---")
        
        # Connection diagnostics
        with st.expander("🔧 Connection Diagnostics"):
            st.write("**Spreadsheet ID:**")
            st.code(SPREADSHEET_ID)
            
            if "gcp_service_account" in st.secrets:
                st.success("✅ GCP credentials found")
                try:
                    creds_dict = dict(st.secrets["gcp_service_account"])
                    if 'client_email' in creds_dict:
                        st.info(f"Service account: {creds_dict['client_email']}")
                        st.caption("Make sure this email has 'Viewer' access to your Google Sheet")
                except:
                    st.error("Error reading credentials")
            else:
                st.error("❌ GCP credentials missing")
        
        # Last updated
        current_time = datetime.now()
        st.caption(f"Last updated: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption("Dashboard refreshes every hour")
        
        if st.button("🔄 Refresh Data Now"):
            st.cache_data.clear()
            st.rerun()
    
    # Load data
    with st.spinner("Loading data from Google Sheets..."):
        deals_df, dashboard_df, invoices_df, sales_orders_df = load_all_data()
    
    # Check if data loaded successfully
    if deals_df.empty and dashboard_df.empty:
        st.error("❌ Unable to load data. Please check your Google Sheets connection.")
        
        with st.expander("📋 Setup Checklist"):
            st.markdown("""
            ### Quick Setup Guide:
            
            1. **Google Cloud Setup:**
               - Create a service account in Google Cloud Console
               - Download the JSON key file
               - Note the service account email (ends with @iam.gserviceaccount.com)
            
            2. **Share Your Google Sheet:**
               - Open your Google Sheet
               - Click 'Share' button
               - Add the service account email
               - Give 'Viewer' permission
            
            3. **Add Credentials to Streamlit:**
               - Go to your Streamlit Cloud dashboard
               - Click on your app
               - Go to Settings → Secrets
               - Paste your service account JSON in the format shown in diagnostics above
            
            4. **Verify Sheet Structure:**
               - Ensure sheet names match: 'All Reps All Pipelines', 'Dashboard Info', 'NS Invoices', 'NS Sales Orders'
               - Verify columns are in the expected positions
            """)
        
        return
    elif deals_df.empty:
        st.warning("⚠️ Deals data is empty. Check 'All Reps All Pipelines' sheet.")
    elif dashboard_df.empty:
        st.warning("⚠️ Dashboard info is empty. Check 'Dashboard Info' sheet.")
    
    # Display appropriate dashboard
    if view_mode == "Team Overview":
        display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df)
    elif view_mode == "Individual Rep":
        if not dashboard_df.empty:
            rep_name = st.selectbox(
                "Select Rep:",
                options=dashboard_df['Rep Name'].tolist()
            )
            if rep_name:
                display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df)
        else:
            st.error("No rep data available")
    else:  # Reconciliation view
        display_reconciliation_view(deals_df, dashboard_df, sales_orders_df)

if __name__ == "__main__":
    main()
