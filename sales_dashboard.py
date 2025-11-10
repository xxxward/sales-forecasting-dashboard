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

# Custom CSS for styling - Dark Mode Compatible
st.markdown("""
    <style>
    /* Force light mode compatibility for embedded iframes */
    .stApp {
        color-scheme: light dark;
    }
    
    /* Ensure text is always visible in both modes */
    .stMarkdown, .stText, p, span, div {
        color: inherit !important;
    }
    
    .big-font {
        font-size: 28px !important;
        font-weight: bold;
    }
    
    /* Metric cards - adapt to theme */
    .metric-card {
        background-color: rgba(240, 242, 246, 0.5);
        padding: 20px;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    
    .stMetric {
        background-color: rgba(255, 255, 255, 0.05);
        padding: 15px;
        border-radius: 8px;
        box-shadow: 1px 1px 3px rgba(0,0,0,0.1);
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    
    /* Progress breakdown - high contrast gradient */
    .progress-breakdown {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 25px;
        border-radius: 15px;
        color: white !important;
        margin: 20px 0;
        box-shadow: 0 10px 20px rgba(0,0,0,0.3);
    }
    
    .progress-breakdown h3 {
        color: white !important;
        margin-bottom: 15px;
        font-size: 24px;
    }
    
    .progress-item {
        display: flex;
        justify-content: space-between;
        padding: 10px 0;
        border-bottom: 1px solid rgba(255,255,255,0.3);
        color: white !important;
    }
    
    .progress-item:last-child {
        border-bottom: none;
        font-weight: bold;
        font-size: 18px;
        padding-top: 15px;
        border-top: 2px solid rgba(255,255,255,0.5);
    }
    
    .progress-label {
        font-size: 16px;
        color: white !important;
    }
    
    .progress-value {
        font-size: 16px;
        font-weight: 600;
        color: white !important;
    }
    
    /* Tables and sections - theme adaptive */
    .reconciliation-table {
        background: rgba(248, 249, 250, 0.5);
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    
    .section-header {
        background: rgba(240, 242, 246, 0.5);
        padding: 10px 15px;
        border-radius: 8px;
        margin: 15px 0;
        font-weight: bold;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    
    .drill-down-section {
        background: rgba(248, 249, 250, 0.5);
        padding: 10px;
        border-radius: 8px;
        margin: 10px 0;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    
    /* Ensure dataframes are readable in dark mode */
    .stDataFrame, [data-testid="stDataFrame"] {
        border: 1px solid rgba(128, 128, 128, 0.3);
    }
    
    /* Force metric labels to be visible */
    [data-testid="stMetricLabel"] {
        color: inherit !important;
    }
    
    [data-testid="stMetricValue"] {
        color: inherit !important;
    }
    
    /* Captions should be visible in both modes */
    .caption, [data-testid="caption"] {
        opacity: 0.7;
        color: inherit !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v34_calyx_logo_added"

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
    
    # Load deals data - extend range to include Q1 2026 Spillover column
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
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
                elif col == 'Q1 2026 Spillover':
                    rename_dict[col] = 'Q1 2026 Spillover'
            
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
            # Map additional columns for Shopify identification
            rename_dict = {
                invoices_df.columns[0]: 'Invoice Number',
                invoices_df.columns[1]: 'Status',
                invoices_df.columns[2]: 'Date',
                invoices_df.columns[6]: 'Customer',
                invoices_df.columns[10]: 'Amount',
                invoices_df.columns[14]: 'Sales Rep'
            }
            
            # Try to find HubSpot Pipeline and CSM columns
            for idx, col in enumerate(invoices_df.columns):
                col_str = str(col).lower()
                if 'hubspot' in col_str and 'pipeline' in col_str:
                    rename_dict[col] = 'HubSpot_Pipeline'
                elif col_str == 'csm' or 'csm' in col_str:
                    rename_dict[col] = 'CSM'
            
            invoices_df = invoices_df.rename(columns=rename_dict)
            
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
                (invoices_df['Sales Rep'] != '') &
                (~invoices_df['Sales Rep'].str.lower().isin(['house']))
            ]
            
            # NEW: Create Shopify ECommerce virtual rep for invoices
            # Priority: If actual sales rep is mentioned, attribute to them. Otherwise, Shopify bucket.
            actual_sales_reps = ['Brad Sherman', 'Lance Mitton', 'Dave Borkowski', 'Jake Lynch', 'Alex Gonzalez']
            
            if 'HubSpot_Pipeline' in invoices_df.columns:
                invoices_df['HubSpot_Pipeline'] = invoices_df['HubSpot_Pipeline'].astype(str).str.strip()
            if 'CSM' in invoices_df.columns:
                invoices_df['CSM'] = invoices_df['CSM'].astype(str).str.strip()
            
            # Identify potential Shopify/Ecommerce orders
            shopify_invoice_mask = pd.Series([False] * len(invoices_df), index=invoices_df.index)
            
            if 'HubSpot_Pipeline' in invoices_df.columns:
                shopify_invoice_mask |= (invoices_df['HubSpot_Pipeline'] == 'Ecommerce Pipeline')
            if 'CSM' in invoices_df.columns:
                shopify_invoice_mask |= (invoices_df['CSM'] == 'Shopify ECommerce')
            
            # For Shopify/Ecommerce orders, check if an actual sales rep is mentioned
            if shopify_invoice_mask.any():
                shopify_candidates = invoices_df[shopify_invoice_mask].copy()
                
                for idx, row in shopify_candidates.iterrows():
                    # Check if any actual sales rep is mentioned in Sales Rep or CSM fields
                    sales_rep = str(row.get('Sales Rep', ''))
                    csm = str(row.get('CSM', '')) if 'CSM' in shopify_candidates.columns else ''
                    
                    # Check if one of the 5 actual reps is mentioned
                    actual_rep_found = None
                    for rep in actual_sales_reps:
                        if rep in sales_rep or rep in csm:
                            actual_rep_found = rep
                            break
                    
                    if actual_rep_found:
                        # Attribute to the actual sales rep
                        invoices_df.at[idx, 'Sales Rep'] = actual_rep_found
                    else:
                        # Attribute to Shopify ECommerce
                        invoices_df.at[idx, 'Sales Rep'] = 'Shopify ECommerce'
            
            # Calculate total invoices by rep
            invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
            invoice_totals.columns = ['Rep Name', 'Invoice Total']
            
            dashboard_df['Rep Name'] = dashboard_df['Rep Name'].str.strip()
            
            dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
            dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
            
            dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
            dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
            
            # NEW: Add Shopify ECommerce to dashboard if it has invoices but isn't already in dashboard
            if 'Shopify ECommerce' in invoice_totals['Rep Name'].values:
                if 'Shopify ECommerce' not in dashboard_df['Rep Name'].values:
                    shopify_invoice_total = invoice_totals[invoice_totals['Rep Name'] == 'Shopify ECommerce']['Invoice Total'].iloc[0]
                    new_shopify_row = pd.DataFrame([{
                        'Rep Name': 'Shopify ECommerce',
                        'Quota': 0,  # No quota for Shopify
                        'NetSuite Orders': shopify_invoice_total
                    }])
                    dashboard_df = pd.concat([dashboard_df, new_shopify_row], ignore_index=True)
    
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
        
        # NEW: Map PI || CSM column (Column G based on screenshot)
        for idx, col in enumerate(col_names):
            col_str = str(col).lower()
            if ('pi' in col_str and 'csm' in col_str) or col_str == 'pi || csm':
                rename_dict[col] = 'PI_CSM'
                break
        
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
                (sales_orders_df['Sales Rep'] != 'nan') &
                (~sales_orders_df['Sales Rep'].str.lower().isin(['house']))
            ]
        
        # NEW: Create Shopify ECommerce virtual rep for sales orders
        # Priority: If actual sales rep is mentioned, attribute to them. Otherwise, Shopify bucket.
        if 'PI_CSM' in sales_orders_df.columns:
            sales_orders_df['PI_CSM'] = sales_orders_df['PI_CSM'].astype(str).str.strip()
            
            actual_sales_reps = ['Brad Sherman', 'Lance Mitton', 'Dave Borkowski', 'Jake Lynch', 'Alex Gonzalez']
            
            # Identify potential Shopify orders
            shopify_mask = (
                (sales_orders_df['Sales Rep'] == 'Shopify ECommerce') | 
                (sales_orders_df['PI_CSM'] == 'Shopify ECommerce')
            )
            
            # For Shopify orders, check if an actual sales rep is mentioned
            if shopify_mask.any():
                shopify_candidates = sales_orders_df[shopify_mask].copy()
                
                for idx, row in shopify_candidates.iterrows():
                    # Check if any actual sales rep is mentioned in Sales Rep or PI_CSM fields
                    sales_rep = str(row.get('Sales Rep', ''))
                    pi_csm = str(row.get('PI_CSM', ''))
                    
                    # Check if one of the 5 actual reps is mentioned
                    actual_rep_found = None
                    for rep in actual_sales_reps:
                        if rep in sales_rep or rep in pi_csm:
                            actual_rep_found = rep
                            break
                    
                    if actual_rep_found:
                        # Attribute to the actual sales rep
                        sales_orders_df.at[idx, 'Sales Rep'] = actual_rep_found
                    else:
                        # Attribute to Shopify ECommerce
                        sales_orders_df.at[idx, 'Sales Rep'] = 'Shopify ECommerce'
                
                # Ensure Shopify ECommerce exists in dashboard_df only if it has attributed orders
                shopify_order_count = (sales_orders_df['Sales Rep'] == 'Shopify ECommerce').sum()
                if shopify_order_count > 0 and 'Shopify ECommerce' not in dashboard_df['Rep Name'].values:
                    new_shopify_row = pd.DataFrame([{
                        'Rep Name': 'Shopify ECommerce',
                        'Quota': 0,
                        'NetSuite Orders': 0  # Will be calculated from sales orders in calculate_rep_metrics
                    }])
                    dashboard_df = pd.concat([dashboard_df, new_shopify_row], ignore_index=True)
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
        
        # PENDING APPROVAL LOGIC - FIXED TO PREVENT DOUBLE COUNTING
        # Priority: Age > 14 days takes precedence over date field status
        pending_approval_orders = rep_orders[rep_orders['Status'] == 'Pending Approval'].copy()
        
        if not pending_approval_orders.empty:
            # Define Q4 2025 date range
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Check if we have the Age_Business_Days column
            if 'Age_Business_Days' not in pending_approval_orders.columns:
                st.warning("⚠️ Age_Business_Days column missing from Sales Orders. Cannot calculate Old PA correctly.")
            
            # CATEGORY 3 (PRIORITY): Old Pending Approval (Age > 14 business days)
            # This takes priority and removes orders from other categories
            if 'Age_Business_Days' in pending_approval_orders.columns:
                old_mask = pending_approval_orders['Age_Business_Days'] > 14
                pending_approval_old_details = pending_approval_orders[old_mask].copy()
                # Remove duplicate columns
                if pending_approval_old_details.columns.duplicated().any():
                    pending_approval_old_details = pending_approval_old_details.loc[:, ~pending_approval_old_details.columns.duplicated()]
                pending_approval_old = pending_approval_old_details['Amount'].sum() if not pending_approval_old_details.empty else 0
                
                # Create a mask for orders that are NOT old (Age <= 14 days)
                # These are the only ones eligible for Categories 1 and 2
                young_orders_mask = pending_approval_orders['Age_Business_Days'] <= 14
                young_orders = pending_approval_orders[young_orders_mask].copy()
            else:
                pending_approval_old = 0
                pending_approval_old_details = pd.DataFrame()
                young_orders = pending_approval_orders.copy()
            
            # Now process only the "young" orders (Age <= 14 days) for Categories 1 and 2
            if not young_orders.empty and 'Pending Approval Date' in young_orders.columns:
                
                # CATEGORY 1: Pending Approval WITH valid Q4 dates (and Age <= 14 days)
                pa_with_date_mask = (
                    (young_orders['Pending Approval Date'].notna()) &
                    (young_orders['Pending Approval Date'] != 'No Date') &
                    (young_orders['Pending Approval Date'] >= q4_start) &
                    (young_orders['Pending Approval Date'] <= q4_end)
                )
                pending_approval_details = young_orders[pa_with_date_mask].copy()
                # Remove duplicate columns
                if pending_approval_details.columns.duplicated().any():
                    pending_approval_details = pending_approval_details.loc[:, ~pending_approval_details.columns.duplicated()]
                pending_approval = pending_approval_details['Amount'].sum() if not pending_approval_details.empty else 0
                
                # CATEGORY 2: Pending Approval with "No Date" string (and Age <= 14 days)
                # Robust matching for various "No Date" formats from Google Sheets
                pa_no_date_mask = (
                    (young_orders['Pending Approval Date'].astype(str).str.strip() == 'No Date') |
                    (young_orders['Pending Approval Date'].astype(str).str.strip() == '') |
                    (young_orders['Pending Approval Date'].isna())
                )
                pending_approval_no_date_details = young_orders[pa_no_date_mask].copy()
                # Remove duplicate columns
                if pending_approval_no_date_details.columns.duplicated().any():
                    pending_approval_no_date_details = pending_approval_no_date_details.loc[:, ~pending_approval_no_date_details.columns.duplicated()]
                pending_approval_no_date = pending_approval_no_date_details['Amount'].sum() if not pending_approval_no_date_details.empty else 0
            else:
                # No young orders or missing date column
                pending_approval = 0
                pending_approval_details = pd.DataFrame()
                pending_approval_no_date = 0
                pending_approval_no_date_details = pd.DataFrame()
        
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
        textposition='auto',
        textfont=dict(size=14)
    ))

    fig.add_trace(go.Bar(
        name='Expect/Commit',
        x=['Progress'],
        y=[metrics['expect_commit']],
        marker_color='#43A047',
        text=[f"${metrics['expect_commit']:,.0f}"],
        textposition='auto',
        textfont=dict(size=14)
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

def create_enhanced_waterfall_chart(metrics, title, mode):
    """
    Creates a waterfall chart for forecast progress to address visibility issues with small segments.
    Each component gets its own visible bar height proportional to its value, making small segments readable.
    
    Args:
        metrics: dict with keys like 'orders', 'pending_fulfillment', etc., and 'total_quota', 'total_progress'
        title: Chart title
        mode: 'base' or 'full' to determine which components to include
    """
    # Define the steps based on mode
    if mode == "base":
        steps = [
            {'label': 'Invoiced', 'value': metrics['orders'], 'color': '#1E88E5'},
            {'label': 'Pending Fulfillment', 'value': metrics['pending_fulfillment'], 'color': '#FFC107'},
            {'label': 'Pending Approval', 'value': metrics['pending_approval'], 'color': '#FB8C00'},
            {'label': 'HubSpot Expect/Commit', 'value': metrics['expect_commit'], 'color': '#43A047'},
        ]
    elif mode == "full":
        steps = [
            {'label': 'Invoiced', 'value': metrics['orders'], 'color': '#1E88E5'},
            {'label': 'Pending Fulfillment', 'value': metrics['pending_fulfillment'], 'color': '#FFC107'},
            {'label': 'PF No Date', 'value': metrics.get('pending_fulfillment_no_date', 0), 'color': '#FFE082'},
            {'label': 'Pending Approval', 'value': metrics['pending_approval'], 'color': '#FB8C00'},
            {'label': 'PA No Date', 'value': metrics.get('pending_approval_no_date', 0), 'color': '#FFCC80'},
            {'label': 'Old PA (>2 weeks)', 'value': metrics.get('pending_approval_old', 0), 'color': '#FF9800'},
            {'label': 'HubSpot Expect/Commit', 'value': metrics['expect_commit'], 'color': '#43A047'},
        ]
    else:
        return None

    # Filter out zero-value steps to avoid clutter
    steps = [step for step in steps if step['value'] > 0]
    
    if not steps:
        return None

    # Calculate totals
    current_total = sum(step['value'] for step in steps)
    quota = metrics.get('total_quota', metrics.get('quota', 0))
    gap = quota - current_total
    
    # Create figure
    fig = go.Figure()
    
    # Add each component as a separate bar trace for full color control
    cumulative = 0
    for step in steps:
        fig.add_trace(go.Bar(
            name=step['label'],
            x=[step['label']],
            y=[step['value']],
            marker_color=step['color'],
            text=[f"${step['value']:,.0f}"],
            textposition='outside',
            textfont=dict(size=12, color='black'),
            hovertemplate=f"<b>{step['label']}</b><br>${step['value']:,.0f}<br>Cumulative: ${cumulative + step['value']:,.0f}<extra></extra>",
            showlegend=True
        ))
        cumulative += step['value']
    
    # Add total bar showing cumulative sum
    fig.add_trace(go.Bar(
        name='TOTAL FORECAST',
        x=['TOTAL'],
        y=[current_total],
        marker_color='#7B1FA2',
        marker_line=dict(width=2, color='#4A148C'),
        text=[f"${current_total:,.0f}"],
        textposition='outside',
        textfont=dict(size=14, color='black', family='Arial Black'),
        hovertemplate=f"<b>Total Forecast</b><br>${current_total:,.0f}<extra></extra>",
        showlegend=True
    ))
    
    # Add gap bar if exists
    if gap != 0:
        gap_color = '#DC3912' if gap > 0 else '#43A047'
        gap_label = 'Gap to Goal' if gap > 0 else 'Over Goal'
        fig.add_trace(go.Bar(
            name=gap_label,
            x=[gap_label],
            y=[abs(gap)],
            marker_color=gap_color,
            text=[f"${gap:,.0f}"],
            textposition='outside',
            textfont=dict(size=12, color='black'),
            hovertemplate=f"<b>{gap_label}</b><br>${gap:,.0f}<extra></extra>",
            showlegend=True
        ))
    
    # Add quota reference line
    fig.add_hline(
        y=quota,
        line_dash="dash",
        line_color="#DC3912",
        line_width=2,
        annotation_text=f"Quota Goal: ${quota:,.0f}",
        annotation_position="right"
    )
    
    # Add best case potential line if in base mode
    best_opp = metrics.get('best_opp', 0)
    if best_opp > 0:
        potential = current_total + best_opp
        fig.add_hline(
            y=potential,
            line_dash="dot",
            line_color="#FB8C00",
            line_width=2,
            annotation_text=f"Potential: ${potential:,.0f}",
            annotation_position="right"
        )
    
    # Customize layout
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=18, color='#333333')
        ),
        xaxis_title="Forecast Components",
        yaxis_title="Amount ($)",
        barmode='group',
        height=600,
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02,
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor="#333333",
            borderwidth=1
        ),
        plot_bgcolor='white',
        yaxis=dict(
            gridcolor='#E5E5E5',
            zeroline=True,
            zerolinecolor='#999999',
            zerolinewidth=1
        ),
        xaxis=dict(
            tickangle=-45,
            automargin=True
        ),
        margin=dict(l=70, r=200, t=100, b=120),
        annotations=[
            dict(
                x=1.02,
                y=1.05,
                xref='paper',
                yref='paper',
                text=f"<b>Current Total:</b> ${current_total:,.0f}<br><b>Quota:</b> ${quota:,.0f}<br><b>Gap:</b> ${gap:,.0f}",
                showarrow=False,
                font=dict(size=13, color="black"),
                align="left",
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#333333",
                borderwidth=1,
                borderpad=8
            )
        ]
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

    fig.update_traces(textfont_size=14, textposition='auto')

    fig.update_layout(
        height=450,
        yaxis_title="Amount ($)",
        xaxis_title="Pipeline",
        xaxis=dict(
            automargin=True,
            tickangle=-45
        ),
        yaxis=dict(automargin=True),
        margin=dict(l=50, r=50, t=80, b=100),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
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
    
    item_count = len(details_df)
    with st.expander(f"{title}: ${amount:,.2f} (👀 Click to see {item_count} {'item' if item_count == 1 else 'items'})"):
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
            st.info("📭 Nothing to see here... yet!")

def display_progress_breakdown(metrics):
    """Display a beautiful progress breakdown card"""
    
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>💰 Section 1: The Money We Can Count On</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Already Celebrating (Invoiced & Shipped)</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 In the Warehouse (Just Add Shipping Label)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Waiting for the Magic Signature</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Deals We're Banking On (HubSpot Expect/Commit)</span>
            <span class="progress-value">${metrics['expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 THE SAFE BET TOTAL</span>
            <span class="progress-value">${metrics['total_progress']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Add attainment info below
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Current Attainment", f"{metrics['attainment_pct']:.1f}%", 
                 delta=f"${metrics['total_progress']:,.0f} of ${metrics['quota']:,.0f}",
                 help="This is real money! 💵")
    with col2:
        st.metric("If Everything Goes Right", f"{metrics['potential_attainment']:.1f}%",
                 delta=f"+${metrics['best_opp']:,.0f} Best Case/Opp",
                 help="The optimist's view (we believe! 🌟)")

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
    
    # Calculate differences first
    diff = totals['total_boss'] - totals['total_you']
    final_diff = additional_totals['final_boss'] - additional_totals['final_you']
    
    # Debug: Show the actual totals being compared
    st.caption(f"Debug: Your Total = ${additional_totals['final_you']:,.0f} | Boss Total = ${additional_totals['final_boss']:,.0f} | Diff = ${abs(final_diff):,.0f}")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Section 1 Variance", f"${abs(diff):,.0f}", 
                 delta=f"{'Under' if diff > 0 else 'Over'} by ${abs(diff):,.0f}")
    
    with col2:
        st.metric("Total Q4 Variance", f"${abs(final_diff):,.0f}",
                 delta=f"{'Under' if final_diff > 0 else 'Over'} by ${abs(final_diff):,.0f}")
    
    with col3:
        if additional_totals['final_boss'] > 0:
            # Calculate accuracy as: 100% - (percentage difference)
            accuracy = (1 - abs(final_diff) / additional_totals['final_boss']) * 100
            
            # Show color coding
            if accuracy >= 95:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Excellent match")
            elif accuracy >= 90:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Good match", delta_color="normal")
            elif accuracy >= 80:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Needs review", delta_color="off")
            else:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Large variance", delta_color="inverse")
        else:
            st.metric("Accuracy", "N/A", delta="No boss data")

def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display the team-level dashboard"""
   
    st.title("🎯 Team Sales Dashboard - Q4 2025")
   
    # Calculate basic metrics
    basic_metrics = calculate_team_metrics(deals_df, dashboard_df)
   
    # Aggregate full team metrics from per-rep calculations
    team_quota = basic_metrics['total_quota']
    team_best_opp = basic_metrics['best_opp']
    team_q1_spillover = 0  # Calculate fresh from rep metrics
   
    team_invoiced = 0
    team_pf = 0
    team_pa = 0
    team_hs = 0
    team_pf_no_date = 0
    team_pa_no_date = 0
    team_old_pa = 0
   
    section1_data = []
    section2_data = []
   
    # Filter out unwanted reps
    excluded_reps = ['House', 'house', 'HOUSE']
    
    for rep_name in dashboard_df['Rep Name']:
        # Skip excluded reps
        if rep_name in excluded_reps:
            continue
            
        rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
        if rep_metrics:
            section1_total = (rep_metrics['orders'] + rep_metrics['pending_fulfillment'] +
                              rep_metrics['pending_approval'] + rep_metrics['expect_commit'])
            final_total = (section1_total + rep_metrics['pending_fulfillment_no_date'] +
                           rep_metrics['pending_approval_no_date'] + rep_metrics['pending_approval_old'])
           
            section1_data.append({
                'Rep': rep_name,
                'Invoiced': f"${rep_metrics['orders']:,.0f}",
                'Pending Fulfillment': f"${rep_metrics['pending_fulfillment']:,.0f}",
                'Pending Approval': f"${rep_metrics['pending_approval']:,.0f}",
                'HubSpot Expect/Commit': f"${rep_metrics['expect_commit']:,.0f}",
                'Total': f"${section1_total:,.0f}"
            })
           
            section2_data.append({
                'Rep': rep_name,
                'PF SO\'s No Date': f"${rep_metrics['pending_fulfillment_no_date']:,.0f}",
                'PA SO\'s No Date': f"${rep_metrics['pending_approval_no_date']:,.0f}",
                'Old PA (>2 weeks)': f"${rep_metrics['pending_approval_old']:,.0f}",
                'Total Q4': f"${final_total:,.0f}"
            })
           
            # Aggregate sums
            team_invoiced += rep_metrics['orders']
            team_pf += rep_metrics['pending_fulfillment']
            team_pa += rep_metrics['pending_approval']
            team_hs += rep_metrics['expect_commit']
            team_pf_no_date += rep_metrics['pending_fulfillment_no_date']
            team_pa_no_date += rep_metrics['pending_approval_no_date']
            team_old_pa += rep_metrics['pending_approval_old']
            team_q1_spillover += rep_metrics.get('q1_spillover_total', 0)
   
    # Calculate team totals
    base_forecast = team_invoiced + team_pf + team_pa + team_hs
    full_forecast = base_forecast + team_pf_no_date + team_pa_no_date + team_old_pa
    base_gap = team_quota - base_forecast
    full_gap = team_quota - full_forecast
    base_attainment_pct = (base_forecast / team_quota * 100) if team_quota > 0 else 0
    full_attainment_pct = (full_forecast / team_quota * 100) if team_quota > 0 else 0
    potential_attainment = ((base_forecast + team_best_opp) / team_quota * 100) if team_quota > 0 else 0
   
    # Add total rows to data
    section1_data.append({
        'Rep': 'TOTAL',
        'Invoiced': f"${team_invoiced:,.0f}",
        'Pending Fulfillment': f"${team_pf:,.0f}",
        'Pending Approval': f"${team_pa:,.0f}",
        'HubSpot Expect/Commit': f"${team_hs:,.0f}",
        'Total': f"${base_forecast:,.0f}"
    })
   
    section2_data.append({
        'Rep': 'TOTAL',
        'PF SO\'s No Date': f"${team_pf_no_date:,.0f}",
        'PA SO\'s No Date': f"${team_pa_no_date:,.0f}",
        'Old PA (>2 weeks)': f"${team_old_pa:,.0f}",
        'Total Q4': f"${full_forecast:,.0f}"
    })
   
    # Display Q1 spillover info if applicable
    if team_q1_spillover > 0:
        st.info(
            f"ℹ️ **Q1 2026 Spillover**: ${team_q1_spillover:,.0f} in deals closing late Q4 2025 "
            f"will ship in Q1 2026 due to product lead times. These are excluded from Q4 revenue recognition."
        )
   
    # Display key metrics with two breakdowns
    st.markdown("### 📊 Team Scorecard")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
   
    with col1:
        st.metric(
            label="🎯 Total Quota",
            value=f"${team_quota:,.0f}",
            delta=None,
            help="Q4 2025 Sales Target"
        )
   
    with col2:
        st.metric(
            label="💪 High Confidence Forecast",
            value=f"${base_forecast:,.0f}",
            delta=f"{base_attainment_pct:.1f}% of quota",
            help="Invoiced Orders + Pending Fulfillment (with date) + Pending Approval (with date) + HubSpot Expect/Commit"
        )
   
    with col3:
        st.metric(
            label="📊 Full Forecast (All Sources)",
            value=f"${full_forecast:,.0f}",
            delta=f"{full_attainment_pct:.1f}% of quota",
            help="High Confidence Forecast + Pending Fulfillment (without date) + Pending Approval (without date) + Old Pending Approval (>2 weeks)"
        )
    
    with col4:
        adjusted_forecast = full_forecast - team_q1_spillover
        adjusted_attainment = (adjusted_forecast / team_quota * 100) if team_quota > 0 else 0
        st.metric(
            label="🎯 Q4 Adjusted Forecast",
            value=f"${adjusted_forecast:,.0f}",
            delta=f"{adjusted_attainment:.1f}% of quota",
            help="Full Forecast (All Sources) minus Q1 2026 Spillover deals"
        )
   
    with col5:
        st.metric(
            label="📈 Gap to Quota",
            value=f"${base_gap:,.0f}",
            delta=f"${-base_gap:,.0f}" if base_gap < 0 else None,
            delta_color="inverse",
            help="Remaining amount needed to reach quota (based on High Confidence Forecast)"
        )

    with col6:
        st.metric(
            label="🌟 Potential Attainment",
            value=f"{potential_attainment:.1f}%",
            delta=f"+{potential_attainment - base_attainment_pct:.1f}% upside",
            help="If all Best Case + Opportunity deals close"
        )
   
    # Progress bars for both breakdowns
    st.markdown("### 📈 Progress to Quota")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💪 High Confidence Forecast Progress**")
        st.caption("Confirmed orders and forecast with dates")
        base_progress = min(base_attainment_pct / 100, 1.0)
        st.progress(base_progress)
        st.caption(f"Current: {base_attainment_pct:.1f}% | Potential: {potential_attainment:.1f}%")
   
    with col2:
        st.markdown("**📊 Full Forecast Progress**")
        st.caption("All sources including orders without dates")
        full_progress = min(full_attainment_pct / 100, 1.0)
        st.progress(full_progress)
        st.caption(f"Current: {full_attainment_pct:.1f}%")
   
    # Base Forecast Chart with Enhanced Annotations
    st.markdown("### 💪 High Confidence Forecast Breakdown")
    st.caption("Orders and deals with confirmed dates and high confidence")
    
    # Create metrics dict for base chart
    base_metrics = {
        'orders': team_invoiced,
        'pending_fulfillment': team_pf,
        'pending_approval': team_pa,
        'expect_commit': team_hs,
        'best_opp': team_best_opp,
        'total_progress': base_forecast,
        'total_quota': team_quota
    }
    
    base_chart = create_enhanced_waterfall_chart(base_metrics, "💪 High Confidence Forecast - Path to Quota", "base")
    st.plotly_chart(base_chart, use_container_width=True)

    # Full Forecast Chart with Enhanced Annotations
    st.markdown("### 📊 Full Forecast Breakdown")
    st.caption("Complete view including all orders and pending items")
    
    full_metrics = {
        'orders': team_invoiced,
        'pending_fulfillment': team_pf,
        'pending_fulfillment_no_date': team_pf_no_date,
        'pending_approval': team_pa,
        'pending_approval_no_date': team_pa_no_date,
        'pending_approval_old': team_old_pa,
        'expect_commit': team_hs,
        'best_opp': team_best_opp,
        'total_progress': base_forecast,
        'total_quota': team_quota
    }
    
    full_chart = create_enhanced_waterfall_chart(full_metrics, "📊 Full Forecast - All Sources Included", "full")
    st.plotly_chart(full_chart, use_container_width=True)

    # Quick Translation Guide
    st.markdown("""
    ---
    ### 📋 Metric Definitions

    **High Confidence Forecast includes:**
    - ✅ **Invoiced Orders** - Billed and recognized revenue
    - 📦 **Pending Fulfillment (with date)** - Orders with confirmed Q4 ship dates
    - ⏳ **Pending Approval (with date)** - Orders with Q4 approval dates scheduled
    - 🎯 **HubSpot Expect/Commit** - High probability forecasted deals

    **Full Forecast adds:**
    - 📦 **Pending Fulfillment (without date)** - Orders without confirmed ship dates
    - ⏳ **Pending Approval (without date)** - Orders without approval dates (age ≤2 weeks)
    - ⏱️ **Old Pending Approval (>2 weeks)** - Orders over 14 business days old

    **Other:**
    - 🦘 **Q1 2026 Spillover** - Deals closing in Q4 2025 but shipping in Q1 2026 due to lead times
    - 🎲 **Best Case/Opportunity** - Lower confidence pipeline deals
    ---
    """)

    # Other charts remain the same
    col1, col2 = st.columns(2)
   
    with col1:
        st.markdown("#### 🎯 Deal Confidence Levels")
        status_chart = create_status_breakdown_chart(deals_df)
        if status_chart:
            st.plotly_chart(status_chart, use_container_width=True)
        else:
            st.info("📭 Nothing to see here... yet!")
   
    with col2:
        st.markdown("#### 🔮 The Crystal Ball: Where Our Deals Stand")
        pipeline_chart = create_pipeline_breakdown_chart(deals_df)
        if pipeline_chart:
            st.plotly_chart(pipeline_chart, use_container_width=True)
        else:
            st.info("📭 Nothing to see here... yet!")
   
    st.markdown("### 📅 When the Magic Happens (Expected Close Dates)")
    timeline_chart = create_deals_timeline(deals_df)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    else:
        st.info("📭 Nothing to see here... yet!")
   
    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Status (Show Me the Money!)")
        invoice_chart = create_invoice_status_chart(invoices_df)
        if invoice_chart:
            st.plotly_chart(invoice_chart, use_container_width=True)
   
    # Display the two sections
    st.markdown("### 👥 High Confidence Forecast by Rep")
    st.caption("Invoiced + Pending Fulfillment (with date) + Pending Approval (with date) + HubSpot Expect/Commit")
    if section1_data:
        section1_df = pd.DataFrame(section1_data)
        st.dataframe(section1_df, use_container_width=True, hide_index=True)
    else:
        st.warning("📭 No data for High Confidence Forecast")
   
    st.markdown("### 👥 Additional Forecast Items by Rep")
    st.caption("Section 1 (above) + items below = Total Q4. Items below: Pending Fulfillment (without date) + Pending Approval (without date) + Old Pending Approval (>2 weeks)")
    if section2_data:
        section2_df = pd.DataFrame(section2_data)
        st.dataframe(section2_df, use_container_width=True, hide_index=True)
    else:
        st.warning("📭 No additional forecast items")
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
        total_deals = metrics['total_q4_closing_deals']
        total_amount = metrics['total_q4_closing_amount']
        spillover_amount = metrics.get('q1_spillover_total', 0)
        
        st.info(
            f"📋 **Total Q4 2025 Pipeline Summary**\n\n"
            f"• **{total_deals} deals** worth **${total_amount:,.0f}** closing in Q4 2025\n\n"
            f"• **${spillover_amount:,.0f}** will ship in Q1 2026 based on lead times"
        )
    
    # Display key metrics - Section 1
    st.markdown("### 💪 High Confidence Forecast (With Confirmed Dates)")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            label="Quota",
            value=f"${metrics['quota']/1000:.0f}K" if metrics['quota'] < 1000000 else f"${metrics['quota']/1000000:.1f}M",
            help="Q4 2025 Sales Target"
        )
    
    with col2:
        st.metric(
            label="✅ Invoiced",
            value=f"${metrics['orders']/1000:.0f}K" if metrics['orders'] < 1000000 else f"${metrics['orders']/1000000:.1f}M",
            help="Invoiced Orders - Billed and recognized revenue"
        )
    
    with col3:
        st.metric(
            label="📦 Pending Fulfillment",
            value=f"${metrics['pending_fulfillment']/1000:.0f}K",
            help="Pending Fulfillment (with date) - Orders with confirmed Q4 ship dates"
        )
    
    with col4:
        st.metric(
            label="⏳ Pending Approval",
            value=f"${metrics['pending_approval']/1000:.0f}K",
            help="Pending Approval (with date) - Orders with Q4 approval dates scheduled"
        )
    
    with col5:
        st.metric(
            label="🎯 HubSpot (Q4)",
            value=f"${metrics['expect_commit']/1000:.0f}K",
            help="HubSpot Expect/Commit - High probability forecasted deals"
        )
    
    with col6:
        st.metric(
            label="Gap to Quota",
            value=f"${metrics['gap']/1000:.0f}K",
            delta=f"${-metrics['gap']/1000:.0f}K" if metrics['gap'] < 0 else None,
            delta_color="inverse",
            help="Remaining amount needed to reach quota"
        )
    
    # Beautiful progress breakdown
    display_progress_breakdown(metrics)
    
    # Drill-down sections for Section 1
    st.markdown("#### 📊 High Confidence Forecast Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        display_drill_down_section(
            "📦 Pending Fulfillment (with date)",
            metrics['pending_fulfillment'],
            metrics.get('pending_fulfillment_details', pd.DataFrame()),
            f"{rep_name}_pf"
        )
        
        display_drill_down_section(
            "⏳ Pending Approval (with date)",
            metrics['pending_approval'],
            metrics.get('pending_approval_details', pd.DataFrame()),
            f"{rep_name}_pa"
        )
    
    with col2:
        display_drill_down_section(
            "🎯 HubSpot Expect/Commit",
            metrics['expect_commit'],
            metrics.get('expect_commit_deals', pd.DataFrame()),
            f"{rep_name}_hs"
        )
        
        display_drill_down_section(
            "🎲 Best Case/Opportunity",
            metrics['best_opp'],
            metrics.get('best_opp_deals', pd.DataFrame()),
            f"{rep_name}_bo"
        )
    
    # Section 2: Additional Orders
    st.markdown("### 📊 Additional Forecast Items (Without Confirmed Dates)")
    
    warning_col1, warning_col2, warning_col3 = st.columns(3)
    
    with warning_col1:
        st.metric(
            label="📦 PF (without date)",
            value=f"${metrics['pending_fulfillment_no_date']:,.0f}",
            help="Pending Fulfillment (without date) - Orders without confirmed ship dates"
        )
        display_drill_down_section(
            "📦 Pending Fulfillment (without date)",
            metrics['pending_fulfillment_no_date'],
            metrics.get('pending_fulfillment_no_date_details', pd.DataFrame()),
            f"{rep_name}_pf_no_date"
        )
    
    with warning_col2:
        st.metric(
            label="⏳ PA (without date)",
            value=f"${metrics['pending_approval_no_date']:,.0f}",
            help="Pending Approval (without date) - Orders without approval dates (age ≤2 weeks)"
        )
        display_drill_down_section(
            "⏳ Pending Approval (without date)",
            metrics['pending_approval_no_date'],
            metrics.get('pending_approval_no_date_details', pd.DataFrame()),
            f"{rep_name}_pa_no_date"
        )
    
    with warning_col3:
        st.metric(
            label="⏱️ Old PA (>2 weeks)",
            value=f"${metrics['pending_approval_old']:,.0f}",
            help="Old Pending Approval (>2 weeks) - Orders over 14 business days old",
            delta="Review needed" if metrics['pending_approval_old'] > 0 else None,
            delta_color="off" if metrics['pending_approval_old'] > 0 else "normal"
        )
        display_drill_down_section(
            "⏱️ Old Pending Approval (>2 weeks)",
            metrics['pending_approval_old'],
            metrics.get('pending_approval_old_details', pd.DataFrame()),
            f"{rep_name}_pa_old"
        )
    
    # NEW: Section 3 - Q1 2026 Spillover
    st.markdown("### 🦘 Q1 2026 Spillover Deals")
    st.caption("Deals closing in Q4 2025 but shipping in Q1 2026 due to lead times")
    
    spillover_col1, spillover_col2, spillover_col3 = st.columns(3)
    
    with spillover_col1:
        st.metric(
            label="🎯 Expect/Commit (Q1 Spillover)",
            value=f"${metrics.get('q1_spillover_expect_commit', 0):,.0f}",
            help="High confidence deals closing Q4 2025, shipping Q1 2026"
        )
        display_drill_down_section(
            "🎯 Expect/Commit (Q1 Spillover)",
            metrics.get('q1_spillover_expect_commit', 0),
            metrics.get('expect_commit_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_ec_q1"
        )
    
    with spillover_col2:
        st.metric(
            label="🎲 Best Case/Opp (Q1 Spillover)",
            value=f"${metrics.get('q1_spillover_best_opp', 0):,.0f}",
            help="Lower probability deals closing Q4 2025, shipping Q1 2026"
        )
        display_drill_down_section(
            "🎲 Best Case/Opp (Q1 Spillover)",
            metrics.get('q1_spillover_best_opp', 0),
            metrics.get('best_opp_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_bo_q1"
        )
    
    with spillover_col3:
        st.metric(
            label="📦 Total Q1 Spillover",
            value=f"${metrics.get('q1_spillover_total', 0):,.0f}",
            help="All deals closing Q4 2025 but shipping Q1 2026"
        )
        display_drill_down_section(
            "📦 Total Q1 2026 Spillover",
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
        gap_chart = create_gap_chart(metrics, f"{rep_name} - Q4 2025 Forecast Progress")
        st.plotly_chart(gap_chart, use_container_width=True)
    
    with col2:
        status_chart = create_status_breakdown_chart(deals_df, rep_name)
        if status_chart:
            st.plotly_chart(status_chart, use_container_width=True)
        else:
            st.info("No deal data available for this rep")
    
    # Pipeline breakdown
    st.markdown("### 📊 Pipeline Breakdown by Status")
    pipeline_chart = create_pipeline_breakdown_chart(deals_df, rep_name)
    if pipeline_chart:
        st.plotly_chart(pipeline_chart, use_container_width=True)
    else:
        st.info("📭 Nothing to see here... yet!")
    
    # Timeline
    st.markdown("### 📅 Deal Timeline by Expected Close Date")
    timeline_chart = create_deals_timeline(deals_df, rep_name)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    else:
        st.info("📭 Nothing to see here... yet!")

# Main app
def main():
    
    # Dashboard tagline
    st.markdown("""
    <div style='text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 color: white; border-radius: 10px; margin-bottom: 20px;'>
        <h3>📊 Sales Forecast Dashboard</h3>
        <p style='font-size: 14px; margin: 0;'>Where numbers meet reality (and sometimes they argue)</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        # Add Calyx logo
        logo_base64 = "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"
        st.logo(f"data:image/png;base64,{logo_base64}")
        
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
