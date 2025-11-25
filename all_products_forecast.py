"""
All Products Forecasting Module - Multi-Source
===============================================
Creates 2026 forecasts combining:
1. Historical Invoice Line Item data (baseline trend)
2. Pending Sales Orders (committed revenue not yet invoiced)

Uses weighted historical analysis (2024 weighted higher than 2025).
Allows filtering and analysis by product type, item type, customer, and sales rep.

Navigation: "📦 All Products Forecast" in sidebar menu

Data Sources:
-------------
Invoice Line Item tab (A:X):
A: Document Number, B: Status, C: Date, D: Due Date, E: Created From
F: Created By, G: Customer, H: Item, I: Quantity, J: Account
K: Period, L: Department, M: Amount, N: Amount (Transaction Total)
O: Amount Remaining, P: CSM, Q: Date Closed, R: Sales Rep
S: External ID, T: Amount (Shipping), U: Amount (Transaction Tax Total)
V: Terms, W: Calyx | Item Type, X: PI || Product Type

Sales Order Line Item tab (A:W):
A: Internal ID, B: Document Number, C: Item, D: Amount, E: Item Rate
F: Quantity Ordered, G: Quantity Fulfilled, H: Purchase Order, I: Invoice
J: PI || Product Type, K: Transaction Discount, L: Income Account, M: Class
N: Quote, O: Pending Fulfillment Date, P: Actual Ship Date, Q: Date Created
R: Date Billed, S: Date Closed, T: Customer Companyname, U: HubSpot Pipeline
V: Calyx | Item Type, W: Status
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Google Sheets Configuration (same as main dashboard)
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "all_products_v3"  # Updated version for three-source

# Confidence factor for active sales orders (95% typically convert to invoices)
ACTIVE_ORDER_CONFIDENCE = 0.95

# HubSpot Close Status probability mapping (Column P)
HUBSPOT_CLOSE_STATUS_PROBABILITY = {
    'Commit': 0.90,        # 90% - Highly likely to close
    'Expect': 0.75,        # 75% - Expected to close
    'Best Case': 0.50,     # 50% - Possible to close
    'Opportunity': 0.25,   # 25% - Early stage opportunity
}

# Sales Rep Configuration
SALES_REPS = [
    'Alex Gonzalez',
    'Lance Mitton', 
    'Dave Borkowski',
    'Jake Lynch',
    'Brad Sherman'
]

# Outlier and data quality settings
MIN_REVENUE_FOR_RECOMMENDATION = 50000  # Minimum $50K revenue to be recommended
MIN_ORDERS_FOR_ANALYSIS = 5  # Minimum 5 orders to be considered
OUTLIER_THRESHOLD_STD = 2.0  # Standard deviations for planning recommendations (2σ = ~95% of data)

# =============================================================================
# DATA LOADING
# =============================================================================

@st.cache_data(ttl=CACHE_TTL)
def load_invoice_line_items(version=CACHE_VERSION):
    """
    Load data from Invoice Line Item tab in Google Sheets
    """
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
        
        # Load from Invoice Line Item tab - columns A:X
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Invoice Line Item!A:X"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning("⚠️ No data found in 'Invoice Line Item' tab")
            return pd.DataFrame()
        
        # Pad rows to match column count
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"❌ Error loading Invoice Line Item data: {str(e)}")
        return pd.DataFrame()


@st.cache_data(ttl=CACHE_TTL)
def load_sales_order_line_items(version=CACHE_VERSION):
    """
    Load data from Sales Order Line Item tab in Google Sheets
    
    Columns (A:W):
    A: Internal ID, B: Document Number, C: Item, D: Amount, E: Item Rate
    F: Quantity Ordered, G: Quantity Fulfilled, H: Purchase Order, I: Invoice
    J: PI || Product Type, K: Transaction Discount, L: Income Account, M: Class
    N: Quote, O: Pending Fulfillment Date, P: Actual Ship Date, Q: Date Created
    R: Date Billed, S: Date Closed, T: Customer Companyname, U: HubSpot Pipeline
    V: Calyx | Item Type, W: Status
    """
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
        
        # Load from Sales Order Line Item tab - columns A:W
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Sales Order Line Item!A:W"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning("⚠️ No data found in 'Sales Order Line Item' tab")
            return pd.DataFrame()
        
        # Pad rows to match column count
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"❌ Error loading Sales Order Line Item data: {str(e)}")
        return pd.DataFrame()


def load_hubspot_data(version=CACHE_VERSION):
    """
    Load data from Hubspot Data tab in Google Sheets
    
    Contains pipeline opportunities with Product Name, SKU, Deal Stage, Close Date, etc.
    Only includes forecastable deals (Closed Won and Closed Lost already removed)
    """
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
        
        # Load from Hubspot Data tab
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Hubspot Data!A:Z"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.info("ℹ️ No data found in 'Hubspot Data' tab")
            return pd.DataFrame()
        
        # Pad rows to match column count
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"❌ Error loading HubSpot data: {str(e)}")
        return pd.DataFrame()


def clean_numeric(value):
    """Clean and convert value to numeric"""
    if pd.isna(value) or str(value).strip() == '':
        return 0
    cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except:
        return 0


def format_currency(x):
    """Format currency with M for millions, K for thousands"""
    if x >= 1000000:
        return f"${x/1000000:.1f}M"
    elif x >= 1000:
        return f"${x/1000:.0f}K"
    else:
        return f"${int(x)}"


def format_quantity(x):
    """Format quantity with M for millions, K for thousands"""
    if x >= 1000000:
        return f"{x/1000000:.1f}M"
    elif x >= 1000:
        return f"{x/1000:.0f}K"
    else:
        return str(int(x))


def process_invoice_data(df):
    """
    Process the invoice line item data with known column mappings.
    """
    if df.empty:
        return df
    
    # Standardize column names (handle variations)
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'document number':
            col_mapping[col] = 'Document Number'
        elif col_lower == 'status':
            col_mapping[col] = 'Status'
        elif col_lower == 'date':
            col_mapping[col] = 'Date'
        elif col_lower == 'due date':
            col_mapping[col] = 'Due Date'
        elif col_lower == 'created from':
            col_mapping[col] = 'Created From'
        elif col_lower == 'created by':
            col_mapping[col] = 'Created By'
        elif col_lower == 'customer':
            col_mapping[col] = 'Customer'
        elif col_lower == 'item':
            col_mapping[col] = 'Item'
        elif col_lower == 'quantity':
            col_mapping[col] = 'Quantity'
        elif col_lower == 'account':
            col_mapping[col] = 'Account'
        elif col_lower == 'period':
            col_mapping[col] = 'Period'
        elif col_lower == 'department':
            col_mapping[col] = 'Department'
        elif col_lower == 'amount' and 'transaction' not in col_lower and 'shipping' not in col_lower and 'tax' not in col_lower and 'remaining' not in col_lower:
            col_mapping[col] = 'Amount'
        elif 'amount (transaction total)' in col_lower or col_lower == 'amount (transaction total)':
            col_mapping[col] = 'Amount_Transaction_Total'
        elif 'amount remaining' in col_lower:
            col_mapping[col] = 'Amount_Remaining'
        elif col_lower == 'csm':
            col_mapping[col] = 'CSM'
        elif 'date closed' in col_lower:
            col_mapping[col] = 'Date Closed'
        elif col_lower == 'sales rep':
            col_mapping[col] = 'Sales Rep'
        elif col_lower == 'external id':
            col_mapping[col] = 'External ID'
        elif 'amount (shipping)' in col_lower:
            col_mapping[col] = 'Amount_Shipping'
        elif 'amount (transaction tax' in col_lower:
            col_mapping[col] = 'Amount_Tax'
        elif col_lower == 'terms':
            col_mapping[col] = 'Terms'
        elif 'calyx' in col_lower and 'item type' in col_lower:
            col_mapping[col] = 'Item Type'
        elif 'pi' in col_lower and 'product type' in col_lower:
            col_mapping[col] = 'Product Type'
    
    df = df.rename(columns=col_mapping)
    
    # Parse Date column
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df[df['Date'].notna()].copy()
    else:
        st.warning("⚠️ 'Date' column not found in data")
        return pd.DataFrame()
    
    # Add time-based columns
    df['Year'] = df['Date'].dt.year
    df['Month'] = df['Date'].dt.month
    df['YearMonth'] = df['Date'].dt.to_period('M')
    df['Quarter'] = df['Date'].dt.quarter
    df['MonthName'] = df['Date'].dt.strftime('%b')
    df['MonthLabel'] = df['Date'].dt.strftime('%b %Y')
    
    # Clean numeric columns
    if 'Quantity' in df.columns:
        df['Quantity'] = df['Quantity'].apply(clean_numeric)
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
    
    # Clean up Item Type and Product Type
    if 'Item Type' in df.columns:
        df['Item Type'] = df['Item Type'].fillna('Unknown').replace('', 'Unknown')
    if 'Product Type' in df.columns:
        df['Product Type'] = df['Product Type'].fillna('Unknown').replace('', 'Unknown')
    
    return df


def process_sales_order_data(df):
    """
    Process the sales order line item data.
    Uses Date Created (column Q) and includes only ACTIVE orders.
    Filters out old/closed/inactive orders even if technically uninvoiced.
    """
    if df.empty:
        return df
    
    # Standardize column names
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'internal id' in col_lower:
            col_mapping[col] = 'Internal_ID'
        elif 'document number' in col_lower:
            col_mapping[col] = 'Document Number'
        elif col_lower == 'item':
            col_mapping[col] = 'Item'
        elif col_lower == 'amount' and 'transaction' not in col_lower:
            col_mapping[col] = 'Amount'
        elif 'item rate' in col_lower:
            col_mapping[col] = 'Item_Rate'
        elif 'quantity ordered' in col_lower:
            col_mapping[col] = 'Quantity'
        elif 'quantity fulfilled' in col_lower:
            col_mapping[col] = 'Quantity_Fulfilled'
        elif 'purchase order' in col_lower:
            col_mapping[col] = 'Purchase_Order'
        elif col_lower == 'invoice':
            col_mapping[col] = 'Invoice'
        elif 'pi' in col_lower and 'product type' in col_lower:
            col_mapping[col] = 'Product Type'
        elif 'transaction discount' in col_lower:
            col_mapping[col] = 'Transaction_Discount'
        elif 'income account' in col_lower:
            col_mapping[col] = 'Income_Account'
        elif col_lower == 'class':
            col_mapping[col] = 'Class'
        elif col_lower == 'quote':
            col_mapping[col] = 'Quote'
        elif 'pending fulfillment date' in col_lower:
            col_mapping[col] = 'Pending_Fulfillment_Date'
        elif 'actual ship date' in col_lower:
            col_mapping[col] = 'Actual_Ship_Date'
        elif 'date created' in col_lower:
            col_mapping[col] = 'Date_Created'
        elif 'date billed' in col_lower:
            col_mapping[col] = 'Date_Billed'
        elif 'date closed' in col_lower:
            col_mapping[col] = 'Date_Closed'
        elif 'customer companyname' in col_lower or 'customer' in col_lower:
            col_mapping[col] = 'Customer'
        elif 'hubspot pipeline' in col_lower:
            col_mapping[col] = 'HubSpot_Pipeline'
        elif 'calyx' in col_lower and 'item type' in col_lower:
            col_mapping[col] = 'Item Type'
        elif col_lower == 'status':
            col_mapping[col] = 'Status'
    
    df = df.rename(columns=col_mapping)
    
    # Filter out closed/billed orders (even if technically uninvoiced)
    original_count = len(df)
    
    # Exclude orders with Date Closed populated
    if 'Date_Closed' in df.columns:
        closed_mask = df['Date_Closed'].notna() & (df['Date_Closed'] != '')
        df = df[~closed_mask].copy()
        st.info(f"🔍 Excluded {closed_mask.sum():,} closed orders")
    
    # Exclude orders with Date Billed populated  
    if 'Date_Billed' in df.columns:
        billed_mask = df['Date_Billed'].notna() & (df['Date_Billed'] != '')
        df = df[~billed_mask].copy()
        st.info(f"🔍 Excluded {billed_mask.sum():,} billed orders")
    
    # Use Date Created to filter for recent orders only
    if 'Date_Created' not in df.columns:
        st.warning("⚠️ 'Date Created' column not found in Sales Order data")
        return pd.DataFrame()
    
    df['Date_Created'] = pd.to_datetime(df['Date_Created'], errors='coerce')
    
    # Remove rows with no valid date
    df = df[df['Date_Created'].notna()].copy()
    
    if df.empty:
        st.info("ℹ️ No sales orders with valid Date Created")
        return df
    
    # Filter for orders created in the last 12 months (adjust as needed)
    # This captures truly "active" orders, not ancient uninvoiced orders
    cutoff_date = datetime.now() - timedelta(days=365)
    before_date_filter = len(df)
    df = df[df['Date_Created'] >= cutoff_date].copy()
    after_date_filter = len(df)
    
    if before_date_filter > after_date_filter:
        st.info(f"🔍 Filtered to orders from last 12 months: {after_date_filter:,} orders (excluded {before_date_filter - after_date_filter:,} older orders)")
    
    if df.empty:
        st.info("ℹ️ No active sales orders from the last 12 months")
        return df
    
    # For 2026 forecast: Use the creation month to assign to corresponding month in 2026
    df['Month'] = df['Date_Created'].dt.month
    df['Year'] = 2026  # All orders contribute to 2026 forecast
    
    # Clean numeric columns
    if 'Quantity' in df.columns:
        df['Quantity'] = df['Quantity'].apply(clean_numeric)
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
    
    # Clean up Item Type and Product Type
    if 'Item Type' in df.columns:
        df['Item Type'] = df['Item Type'].fillna('Unknown').replace('', 'Unknown')
    if 'Product Type' in df.columns:
        df['Product Type'] = df['Product Type'].fillna('Unknown').replace('', 'Unknown')
    
    # Mark this as active order data
    df['Data_Source'] = 'Active Order'
    
    st.success(f"✅ Final result: {len(df):,} active orders (${df['Amount'].sum():,.0f} total)")
    
    return df


def process_hubspot_data(df):
    """
    Process the HubSpot pipeline data.
    Applies probability weighting based on Close Status (Column P).
    Uses Close Date to assign to months in 2026.
    """
    if df.empty:
        return df
    
    original_count = len(df)
    
    # Standardize column names
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'product name' in col_lower:
            col_mapping[col] = 'Product_Name'
        elif col_lower == 'sku':
            col_mapping[col] = 'SKU'
        elif 'close status' in col_lower:
            col_mapping[col] = 'Close_Status'
        elif 'deal stage' in col_lower:
            col_mapping[col] = 'Deal_Stage'
        elif 'amount in company currency' in col_lower or (col_lower == 'amount' and 'company' in col.lower()):
            col_mapping[col] = 'Amount'
        elif col_lower == 'quantity':
            col_mapping[col] = 'Quantity'
        elif 'close date' in col_lower:
            col_mapping[col] = 'Close_Date'
        elif 'company name' in col_lower:
            col_mapping[col] = 'Company_Name'
        elif 'deal name' in col_lower:
            col_mapping[col] = 'Deal_Name'
        elif 'deal id' in col_lower:
            col_mapping[col] = 'Deal_ID'
        elif 'pipeline' in col_lower and 'hubspot' not in col_lower:
            col_mapping[col] = 'Pipeline'
    
    df = df.rename(columns=col_mapping)
    
    # Filter out rows with no amount or zero amount
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
        df = df[df['Amount'] > 0].copy()
    else:
        st.warning("⚠️ 'Amount' column not found in HubSpot data")
        return pd.DataFrame()
    
    if df.empty:
        st.info("ℹ️ No HubSpot deals with valid amounts")
        return df
    
    # Parse Close Date
    if 'Close_Date' not in df.columns:
        st.warning("⚠️ 'Close Date' column not found in HubSpot data")
        return pd.DataFrame()
    
    df['Close_Date'] = pd.to_datetime(df['Close_Date'], errors='coerce')
    df = df[df['Close_Date'].notna()].copy()
    
    if df.empty:
        st.info("ℹ️ No HubSpot deals with valid Close Date")
        return df
    
    # Filter for deals closing in 2026
    df = df[df['Close_Date'].dt.year == 2026].copy()
    
    if df.empty:
        st.info("ℹ️ No HubSpot deals with 2026 close dates")
        return df
    
    # Extract Year and Month from close date
    df['Year'] = df['Close_Date'].dt.year
    df['Month'] = df['Close_Date'].dt.month
    
    # Apply probability weighting based on Close Status (Column P)
    if 'Close_Status' in df.columns:
        # Map Close Status to probability
        df['Status_Probability'] = df['Close_Status'].map(HUBSPOT_CLOSE_STATUS_PROBABILITY).fillna(0.25)
        df['Weighted_Amount'] = df['Amount'] * df['Status_Probability']
        
        # Show status distribution
        status_summary = df.groupby('Close_Status').agg({
            'Amount': 'sum',
            'Weighted_Amount': 'sum',
            'Deal_ID': 'nunique' if 'Deal_ID' in df.columns else 'count'
        }).reset_index()
        
        st.info(f"🔍 HubSpot: {len(df)} deals across {df['Close_Status'].nunique()} statuses (Commit/Expect/Best Case/Opportunity)")
        
        # Show breakdown by status
        for _, row in status_summary.iterrows():
            status = row['Close_Status']
            count = row['Deal_ID']
            total = row['Amount']
            weighted = row['Weighted_Amount']
            prob = HUBSPOT_CLOSE_STATUS_PROBABILITY.get(status, 0.25)
            st.info(f"   • {status}: {count} deals, ${total:,.0f} total → ${weighted:,.0f} weighted ({prob:.0%})")
        
    else:
        st.warning("⚠️ 'Close Status' column not found, using 25% default probability")
        df['Status_Probability'] = 0.25
        df['Weighted_Amount'] = df['Amount'] * 0.25
    
    # Clean numeric columns
    if 'Quantity' in df.columns:
        df['Quantity'] = df['Quantity'].apply(clean_numeric)
        df['Weighted_Quantity'] = df['Quantity'] * df['Status_Probability']
    else:
        df['Quantity'] = 0
        df['Weighted_Quantity'] = 0
    
    # Try to match Product Name to existing Product Type for filtering
    # Map HubSpot product names to standardized Product Types
    if 'Product_Name' in df.columns:
        # Start with Product_Name as base
        df['Product Type'] = df['Product_Name']
        
        # Apply common mappings (customize based on your naming conventions)
        # This helps match HubSpot names to NetSuite product types
        df['Product Type'] = df['Product Type'].str.strip()
        
        # You can add custom mappings here:
        # df['Product Type'] = df['Product Type'].replace({
        #     'Concentrate Jar - 5ml': 'Concentrate Jars',
        #     'Flower Container': 'Flower Jars',
        #     # Add more mappings as needed
        # })
        
        # Log unique product types for visibility
        unique_hs_products = df['Product Type'].nunique()
        st.info(f"🔍 HubSpot contains {unique_hs_products} unique product types")
    else:
        df['Product Type'] = 'Unknown'
        st.warning("⚠️ 'Product Name' column not found in HubSpot data")
    
    df['Item Type'] = 'Unknown'  # HubSpot data doesn't have this granularity yet
    
    # Mark this as pipeline data
    df['Data_Source'] = 'Pipeline'
    
    st.success(f"✅ Processed {len(df):,} HubSpot pipeline deals (${df['Weighted_Amount'].sum():,.0f} weighted)")
    
    return df


def get_invoiced_order_numbers(invoice_df):
    """
    Extract set of sales order numbers that have already been invoiced
    to prevent double-counting
    """
    if invoice_df.empty or 'Created From' not in invoice_df.columns:
        return set()
    
    # Extract order numbers from "Created From" field
    invoiced_orders = invoice_df['Created From'].dropna().unique()
    
    order_numbers = set()
    for order_ref in invoiced_orders:
        order_str = str(order_ref).strip()
        if order_str and order_str != '' and order_str != 'nan':
            order_numbers.add(order_str)
            # Also extract just number portion if it has a delimiter
            if '#' in order_str:
                num_part = order_str.split('#')[-1].strip()
                order_numbers.add(num_part)
    
    return order_numbers


def filter_uninvoiced_sales_orders(sales_order_df, invoice_df):
    """
    Remove sales orders that have already been invoiced
    """
    if sales_order_df.empty:
        return sales_order_df
    
    # First, exclude any orders where Invoice column is populated
    if 'Invoice' in sales_order_df.columns:
        mask = sales_order_df['Invoice'].isna() | (sales_order_df['Invoice'] == '')
        sales_order_df = sales_order_df[mask].copy()
    
    # Get list of invoiced order numbers
    invoiced_orders = get_invoiced_order_numbers(invoice_df)
    
    if not invoiced_orders or 'Document Number' not in sales_order_df.columns:
        return sales_order_df
    
    original_count = len(sales_order_df)
    original_amount = sales_order_df['Amount'].sum() if 'Amount' in sales_order_df.columns else 0
    
    # Check both full document number and extracted number
    def check_if_invoiced(doc_num):
        doc_str = str(doc_num).strip()
        if doc_str in invoiced_orders:
            return True
        if '#' in doc_str:
            num_part = doc_str.split('#')[-1].strip()
            if num_part in invoiced_orders:
                return True
        return False
    
    mask = ~sales_order_df['Document Number'].apply(check_if_invoiced)
    deduplicated = sales_order_df[mask].copy()
    
    # Log the filtering
    removed_count = original_count - len(deduplicated)
    if removed_count > 0:
        removed_amount = original_amount - (deduplicated['Amount'].sum() if 'Amount' in deduplicated.columns else 0)
        st.info(
            f"🔍 Deduplication: Excluded {removed_count:,} sales order lines "
            f"(${removed_amount:,.0f}) already invoiced"
        )
    
    return deduplicated


# =============================================================================
# OUTLIER HANDLING
# =============================================================================

def soften_outliers(series, limits=(0.05, 0.95)):
    """
    Soften the impact of outliers using winsorization.
    Caps extreme values at the 5th and 95th percentile.
    """
    if len(series) == 0 or series.sum() == 0:
        return series
    
    lower = series.quantile(limits[0])
    upper = series.quantile(limits[1])
    return series.clip(lower=lower, upper=upper)


# =============================================================================
# FORECASTING ENGINE
# =============================================================================

def calculate_weighted_monthly_averages(df, weight_2024=0.6, weight_2025=0.4, group_by=None):
    """
    Calculate weighted monthly averages using historical data.
    
    2024 is weighted higher (default 60%) because stock was healthy.
    2025 is weighted lower (default 40%) due to stock constraints.
    
    If group_by is specified, returns grouped forecasts.
    """
    if df.empty or 'Year' not in df.columns:
        return pd.DataFrame()
    
    # Aggregate by Year, Month, and optionally group_by
    if group_by and group_by in df.columns:
        monthly_data = df.groupby(['Year', 'Month', group_by]).agg({
            'Quantity': 'sum',
            'Amount': 'sum'
        }).reset_index()
    else:
        monthly_data = df.groupby(['Year', 'Month']).agg({
            'Quantity': 'sum',
            'Amount': 'sum'
        }).reset_index()
    
    # Separate years
    data_2023 = monthly_data[monthly_data['Year'] == 2023]
    data_2024 = monthly_data[monthly_data['Year'] == 2024]
    data_2025 = monthly_data[monthly_data['Year'] == 2025]
    
    weighted_avgs = []
    
    if group_by and group_by in df.columns:
        groups = df[group_by].unique()
        
        for group in groups:
            d2023 = data_2023[data_2023[group_by] == group].set_index('Month')
            d2024 = data_2024[data_2024[group_by] == group].set_index('Month')
            d2025 = data_2025[data_2025[group_by] == group].set_index('Month')
            
            for month in range(1, 13):
                qty_2023 = d2023.loc[month, 'Quantity'] if month in d2023.index else 0
                qty_2024 = d2024.loc[month, 'Quantity'] if month in d2024.index else 0
                qty_2025 = d2025.loc[month, 'Quantity'] if month in d2025.index else 0
                amt_2023 = d2023.loc[month, 'Amount'] if month in d2023.index else 0
                amt_2024 = d2024.loc[month, 'Amount'] if month in d2024.index else 0
                amt_2025 = d2025.loc[month, 'Amount'] if month in d2025.index else 0
                
                # Primary weighting: 2024 (60%) + 2025 (40%)
                if qty_2024 > 0 and qty_2025 > 0:
                    weighted_qty = (qty_2024 * weight_2024) + (qty_2025 * weight_2025)
                    weighted_amt = (amt_2024 * weight_2024) + (amt_2025 * weight_2025)
                elif qty_2024 > 0:
                    weighted_qty = qty_2024
                    weighted_amt = amt_2024
                elif qty_2025 > 0:
                    weighted_qty = qty_2025
                    weighted_amt = amt_2025
                elif qty_2023 > 0:
                    weighted_qty = qty_2023
                    weighted_amt = amt_2023
                else:
                    weighted_qty = 0
                    weighted_amt = 0
                
                weighted_avgs.append({
                    group_by: group,
                    'Month': month,
                    'MonthName': datetime(2024, month, 1).strftime('%b'),
                    'Weighted_Quantity': weighted_qty,
                    'Weighted_Amount': weighted_amt,
                    'Qty_2023': qty_2023,
                    'Qty_2024': qty_2024,
                    'Qty_2025': qty_2025,
                    'Amt_2023': amt_2023,
                    'Amt_2024': amt_2024,
                    'Amt_2025': amt_2025
                })
    else:
        d2023 = data_2023.set_index('Month')
        d2024 = data_2024.set_index('Month')
        d2025 = data_2025.set_index('Month')
        
        for month in range(1, 13):
            qty_2023 = d2023.loc[month, 'Quantity'] if month in d2023.index else 0
            qty_2024 = d2024.loc[month, 'Quantity'] if month in d2024.index else 0
            qty_2025 = d2025.loc[month, 'Quantity'] if month in d2025.index else 0
            amt_2023 = d2023.loc[month, 'Amount'] if month in d2023.index else 0
            amt_2024 = d2024.loc[month, 'Amount'] if month in d2024.index else 0
            amt_2025 = d2025.loc[month, 'Amount'] if month in d2025.index else 0
            
            # Primary weighting: 2024 (60%) + 2025 (40%)
            if qty_2024 > 0 and qty_2025 > 0:
                weighted_qty = (qty_2024 * weight_2024) + (qty_2025 * weight_2025)
                weighted_amt = (amt_2024 * weight_2024) + (amt_2025 * weight_2025)
            elif qty_2024 > 0:
                weighted_qty = qty_2024
                weighted_amt = amt_2024
            elif qty_2025 > 0:
                weighted_qty = qty_2025
                weighted_amt = amt_2025
            elif qty_2023 > 0:
                weighted_qty = qty_2023
                weighted_amt = amt_2023
            else:
                weighted_qty = 0
                weighted_amt = 0
            
            weighted_avgs.append({
                'Month': month,
                'MonthName': datetime(2024, month, 1).strftime('%b'),
                'Weighted_Quantity': weighted_qty,
                'Weighted_Amount': weighted_amt,
                'Qty_2023': qty_2023,
                'Qty_2024': qty_2024,
                'Qty_2025': qty_2025,
                'Amt_2023': amt_2023,
                'Amt_2024': amt_2024,
                'Amt_2025': amt_2025
            })
    
    return pd.DataFrame(weighted_avgs)


def generate_2026_forecast(df, weight_2024=0.6, weight_2025=0.4, group_by=None):
    """
    Generate month-by-month and quarter-by-quarter 2026 forecast.
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    # Use historical data (before current date for accuracy)
    cutoff_date = datetime.now()
    historical_df = df[df['Date'] < cutoff_date].copy()
    
    # Soften outliers in historical data
    if not historical_df.empty and historical_df['Quantity'].sum() > 0:
        historical_df['Quantity'] = soften_outliers(historical_df['Quantity'])
        historical_df['Amount'] = soften_outliers(historical_df['Amount'])
    
    # Calculate weighted monthly baselines
    monthly_baselines = calculate_weighted_monthly_averages(historical_df, weight_2024, weight_2025, group_by)
    
    if monthly_baselines.empty:
        st.warning("⚠️ Not enough historical data to generate forecast")
        return pd.DataFrame(), pd.DataFrame(), monthly_baselines
    
    # Generate 2026 monthly forecast
    forecast_2026 = []
    
    if group_by and group_by in monthly_baselines.columns:
        for group in monthly_baselines[group_by].unique():
            group_baselines = monthly_baselines[monthly_baselines[group_by] == group]
            
            for month in range(1, 13):
                baseline_row = group_baselines[group_baselines['Month'] == month]
                if baseline_row.empty:
                    continue
                    
                baseline = baseline_row.iloc[0]
                
                forecasted_qty = baseline['Weighted_Quantity']
                forecasted_amt = baseline['Weighted_Amount']
                
                # Confidence range based on forecast horizon
                if month <= 3:
                    confidence = 0.20
                elif month <= 6:
                    confidence = 0.25
                else:
                    confidence = 0.30
                
                quarter_num = (month - 1) // 3 + 1
                
                forecast_2026.append({
                    group_by: group,
                    'Month': month,
                    'MonthName': datetime(2026, month, 1).strftime('%B'),
                    'MonthShort': datetime(2026, month, 1).strftime('%b'),
                    'QuarterNum': quarter_num,
                    'Quarter': f"Q{quarter_num} 2026",
                    'Forecasted_Quantity': int(forecasted_qty),
                    'Forecasted_Amount': round(forecasted_amt, 2),
                    'Qty_Low': int(forecasted_qty * (1 - confidence)),
                    'Qty_High': int(forecasted_qty * (1 + confidence)),
                    'Amt_Low': round(forecasted_amt * (1 - confidence), 2),
                    'Amt_High': round(forecasted_amt * (1 + confidence), 2),
                    'Historical_Qty_2023': int(baseline['Qty_2023']),
                    'Historical_Qty_2024': int(baseline['Qty_2024']),
                    'Historical_Qty_2025': int(baseline['Qty_2025']),
                    'Historical_Amt_2024': round(baseline['Amt_2024'], 2),
                    'Historical_Amt_2025': round(baseline['Amt_2025'], 2),
                    'Confidence': f"±{int(confidence*100)}%"
                })
    else:
        for month in range(1, 13):
            baseline_row = monthly_baselines[monthly_baselines['Month'] == month]
            if baseline_row.empty:
                continue
                
            baseline = baseline_row.iloc[0]
            
            forecasted_qty = baseline['Weighted_Quantity']
            forecasted_amt = baseline['Weighted_Amount']
            
            # Confidence range based on forecast horizon
            if month <= 3:
                confidence = 0.20
            elif month <= 6:
                confidence = 0.25
            else:
                confidence = 0.30
            
            quarter_num = (month - 1) // 3 + 1
            
            forecast_2026.append({
                'Month': month,
                'MonthName': datetime(2026, month, 1).strftime('%B'),
                'MonthShort': datetime(2026, month, 1).strftime('%b'),
                'QuarterNum': quarter_num,
                'Quarter': f"Q{quarter_num} 2026",
                'Forecasted_Quantity': int(forecasted_qty),
                'Forecasted_Amount': round(forecasted_amt, 2),
                'Qty_Low': int(forecasted_qty * (1 - confidence)),
                'Qty_High': int(forecasted_qty * (1 + confidence)),
                'Amt_Low': round(forecasted_amt * (1 - confidence), 2),
                'Amt_High': round(forecasted_amt * (1 + confidence), 2),
                'Historical_Qty_2023': int(baseline['Qty_2023']),
                'Historical_Qty_2024': int(baseline['Qty_2024']),
                'Historical_Qty_2025': int(baseline['Qty_2025']),
                'Historical_Amt_2024': round(baseline['Amt_2024'], 2),
                'Historical_Amt_2025': round(baseline['Amt_2025'], 2),
                'Confidence': f"±{int(confidence*100)}%"
            })
    
    monthly_forecast = pd.DataFrame(forecast_2026)
    
    # Generate quarterly summary
    quarterly_forecast = pd.DataFrame()
    if not monthly_forecast.empty:
        group_cols = ['QuarterNum', 'Quarter']
        if group_by and group_by in monthly_forecast.columns:
            group_cols = [group_by] + group_cols
        
        quarterly_forecast = monthly_forecast.groupby(group_cols).agg({
            'Forecasted_Quantity': 'sum',
            'Forecasted_Amount': 'sum',
            'Qty_Low': 'sum',
            'Qty_High': 'sum',
            'Amt_Low': 'sum',
            'Amt_High': 'sum'
        }).reset_index()
        
        quarterly_forecast = quarterly_forecast.sort_values('QuarterNum')
    
    return monthly_forecast, quarterly_forecast, monthly_baselines


def calculate_pending_orders_forecast(sales_order_df):
    """
    Calculate monthly forecast contribution from active sales orders.
    Groups by expected month and applies confidence factor.
    """
    if sales_order_df.empty:
        return pd.DataFrame()
    
    # Group by Year and Month
    monthly_pending = sales_order_df.groupby(['Year', 'Month']).agg({
        'Amount': 'sum',
        'Quantity': 'sum',
        'Document Number': 'nunique'  # Count distinct orders
    }).reset_index()
    
    monthly_pending.columns = ['Year', 'Month', 'Pending_Amount', 'Pending_Quantity', 'Order_Count']
    
    # Apply confidence factor (95% of active orders typically convert)
    monthly_pending['Weighted_Amount'] = monthly_pending['Pending_Amount'] * ACTIVE_ORDER_CONFIDENCE
    monthly_pending['Weighted_Quantity'] = monthly_pending['Pending_Quantity'] * ACTIVE_ORDER_CONFIDENCE
    
    # Filter for 2026 only
    monthly_pending = monthly_pending[monthly_pending['Year'] == 2026].copy()
    
    # Add month names
    monthly_pending['MonthName'] = monthly_pending['Month'].apply(
        lambda m: datetime(2026, m, 1).strftime('%B')
    )
    monthly_pending['MonthShort'] = monthly_pending['Month'].apply(
        lambda m: datetime(2026, m, 1).strftime('%b')
    )
    monthly_pending['QuarterNum'] = ((monthly_pending['Month'] - 1) // 3) + 1
    monthly_pending['Quarter'] = monthly_pending['QuarterNum'].apply(lambda q: f"Q{q} 2026")
    
    return monthly_pending


def calculate_hubspot_forecast(hubspot_df):
    """
    Calculate monthly forecast contribution from HubSpot pipeline.
    Groups by close date month and uses weighted amounts (already probability-adjusted).
    """
    if hubspot_df.empty:
        return pd.DataFrame()
    
    # Group by Year and Month
    monthly_hs = hubspot_df.groupby(['Year', 'Month']).agg({
        'Weighted_Amount': 'sum',
        'Weighted_Quantity': 'sum',
        'Deal_ID': 'nunique' if 'Deal_ID' in hubspot_df.columns else 'count'
    }).reset_index()
    
    monthly_hs.columns = ['Year', 'Month', 'Weighted_Amount', 'Weighted_Quantity', 'Deal_Count']
    
    # Filter for 2026 only
    monthly_hs = monthly_hs[monthly_hs['Year'] == 2026].copy()
    
    # Add month names
    monthly_hs['MonthName'] = monthly_hs['Month'].apply(
        lambda m: datetime(2026, m, 1).strftime('%B')
    )
    monthly_hs['MonthShort'] = monthly_hs['Month'].apply(
        lambda m: datetime(2026, m, 1).strftime('%b')
    )
    monthly_hs['QuarterNum'] = ((monthly_hs['Month'] - 1) // 3) + 1
    monthly_hs['Quarter'] = monthly_hs['QuarterNum'].apply(lambda q: f"Q{q} 2026")
    
    return monthly_hs


def combine_forecast_sources(historical_forecast, pending_forecast, hubspot_forecast=None):
    """
    Combine historical baseline + active orders + HubSpot pipeline to create unified forecast.
    """
    if historical_forecast.empty:
        st.warning("⚠️ No historical forecast data")
        return pd.DataFrame(), pd.DataFrame()
    
    combined = historical_forecast.copy()
    
    # Add active order columns (initialize to 0)
    combined['Pending_Amount'] = 0
    combined['Pending_Quantity'] = 0
    combined['Order_Count'] = 0
    
    # Add HubSpot pipeline columns (initialize to 0)
    combined['Pipeline_Amount'] = 0
    combined['Pipeline_Quantity'] = 0
    combined['Deal_Count'] = 0
    
    # Merge in active orders by month
    if not pending_forecast.empty:
        for _, pending_row in pending_forecast.iterrows():
            month = pending_row['Month']
            mask = combined['Month'] == month
            
            if mask.any():
                combined.loc[mask, 'Pending_Amount'] = pending_row['Weighted_Amount']
                combined.loc[mask, 'Pending_Quantity'] = pending_row['Weighted_Quantity']
                combined.loc[mask, 'Order_Count'] = pending_row['Order_Count']
    
    # Merge in HubSpot pipeline by month
    if hubspot_forecast is not None and not hubspot_forecast.empty:
        for _, hs_row in hubspot_forecast.iterrows():
            month = hs_row['Month']
            mask = combined['Month'] == month
            
            if mask.any():
                combined.loc[mask, 'Pipeline_Amount'] = hs_row['Weighted_Amount']
                combined.loc[mask, 'Pipeline_Quantity'] = hs_row['Weighted_Quantity']
                combined.loc[mask, 'Deal_Count'] = hs_row['Deal_Count']
    
    # Rename baseline columns for clarity
    combined['Historical_Baseline_Amount'] = combined['Forecasted_Amount']
    combined['Historical_Baseline_Quantity'] = combined['Forecasted_Quantity']
    
    # Calculate combined totals (all three sources)
    combined['Forecasted_Amount'] = (
        combined['Historical_Baseline_Amount'] + 
        combined['Pending_Amount'] + 
        combined['Pipeline_Amount']
    )
    combined['Forecasted_Quantity'] = (
        combined['Historical_Baseline_Quantity'] + 
        combined['Pending_Quantity'] + 
        combined['Pipeline_Quantity']
    )
    
    # Calculate source mix percentages
    combined['Pct_Historical'] = np.where(
        combined['Forecasted_Amount'] > 0,
        combined['Historical_Baseline_Amount'] / combined['Forecasted_Amount'],
        1.0
    )
    combined['Pct_Pending'] = np.where(
        combined['Forecasted_Amount'] > 0,
        combined['Pending_Amount'] / combined['Forecasted_Amount'],
        0.0
    )
    combined['Pct_Pipeline'] = np.where(
        combined['Forecasted_Amount'] > 0,
        combined['Pipeline_Amount'] / combined['Forecasted_Amount'],
        0.0
    )
    
    # Adjust confidence ranges based on source mix
    # More certain sources = tighter confidence range
    base_confidence = []
    for _, row in combined.iterrows():
        # Base confidence by quarter
        if row['Month'] <= 3:
            base_conf = 0.20
        elif row['Month'] <= 6:
            base_conf = 0.25
        else:
            base_conf = 0.30
        
        # Adjust based on data source mix
        # Historical: ±20-30% confidence
        # Active orders: ±10% confidence (95% certain)
        # Pipeline: ±35% confidence (varies by stage)
        adjusted_conf = (
            row['Pct_Historical'] * base_conf +
            row['Pct_Pending'] * 0.10 +  # Active orders are more certain
            row['Pct_Pipeline'] * 0.35    # Pipeline is less certain
        )
        base_confidence.append(adjusted_conf)
    
    combined['Confidence_Factor'] = base_confidence
    combined['Confidence'] = combined['Confidence_Factor'].apply(lambda x: f"±{int(x*100)}%")
    
    # Recalculate confidence ranges
    combined['Amt_Low'] = combined['Forecasted_Amount'] * (1 - combined['Confidence_Factor'])
    combined['Amt_High'] = combined['Forecasted_Amount'] * (1 + combined['Confidence_Factor'])
    combined['Qty_Low'] = combined['Forecasted_Quantity'] * (1 - combined['Confidence_Factor'])
    combined['Qty_High'] = combined['Forecasted_Quantity'] * (1 + combined['Confidence_Factor'])
    
    # Generate quarterly summary
    quarterly_combined = combined.groupby(['QuarterNum', 'Quarter']).agg({
        'Historical_Baseline_Amount': 'sum',
        'Pending_Amount': 'sum',
        'Pipeline_Amount': 'sum',
        'Forecasted_Amount': 'sum',
        'Forecasted_Quantity': 'sum',
        'Amt_Low': 'sum',
        'Amt_High': 'sum',
        'Qty_Low': 'sum',
        'Qty_High': 'sum',
        'Order_Count': 'sum',
        'Deal_Count': 'sum'
    }).reset_index()
    
    return combined, quarterly_combined


# =============================================================================
# VISUALIZATION
# =============================================================================

def create_historical_trend_chart(df, title_suffix=""):
    """
    Create a stacked bar chart showing historical trends by month.
    """
    if df.empty:
        return None
    
    # Aggregate by YearMonth
    monthly = df.groupby(['Year', 'Month']).agg({
        'Quantity': 'sum',
        'Amount': 'sum'
    }).reset_index()
    
    monthly['MonthLabel'] = monthly.apply(
        lambda x: f"{datetime(int(x['Year']), int(x['Month']), 1).strftime('%b %Y')}", 
        axis=1
    )
    monthly['SortKey'] = monthly['Year'] * 100 + monthly['Month']
    monthly = monthly.sort_values('SortKey')
    
    # Color by year
    color_map = {2023: '#6366f1', 2024: '#8b5cf6', 2025: '#a855f7'}
    monthly['Color'] = monthly['Year'].map(color_map).fillna('#6366f1')
    
    fig = go.Figure()
    
    # Add bars for each year
    for year in sorted(monthly['Year'].unique()):
        year_data = monthly[monthly['Year'] == year]
        fig.add_trace(go.Bar(
            x=year_data['MonthLabel'],
            y=year_data['Amount'],
            name=str(int(year)),
            marker=dict(
                color=color_map.get(year, '#6366f1'),
                line=dict(color='rgba(255,255,255,0.3)', width=1)
            ),
            text=year_data['Amount'].apply(format_currency),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'📊 Historical Revenue by Month{title_suffix}',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
        ),
        barmode='group',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=80)
    )
    
    return fig


def create_forecast_chart(monthly_forecast, metric='Amount'):
    """
    Create a forecast chart with confidence bands for 2026.
    """
    if monthly_forecast.empty:
        return None
    
    fig = go.Figure()
    
    if metric == 'Amount':
        y_col = 'Forecasted_Amount'
        low_col = 'Amt_Low'
        high_col = 'Amt_High'
        hist_2024 = 'Historical_Amt_2024'
        hist_2025 = 'Historical_Amt_2025'
        y_title = 'Revenue ($)'
        format_func = format_currency
        hover_format = 'Revenue: $%{y:,.0f}'
    else:
        y_col = 'Forecasted_Quantity'
        low_col = 'Qty_Low'
        high_col = 'Qty_High'
        hist_2024 = 'Historical_Qty_2024'
        hist_2025 = 'Historical_Qty_2025'
        y_title = 'Quantity (Units)'
        format_func = format_quantity
        hover_format = 'Quantity: %{y:,.0f}'
    
    # Add confidence band (shaded area)
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'].tolist() + monthly_forecast['MonthShort'].tolist()[::-1],
        y=monthly_forecast[high_col].tolist() + monthly_forecast[low_col].tolist()[::-1],
        fill='toself',
        fillcolor='rgba(99, 102, 241, 0.2)',
        line=dict(color='rgba(255,255,255,0)'),
        hoverinfo='skip',
        showlegend=True,
        name='Confidence Range'
    ))
    
    # Add forecast line
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast[y_col],
        mode='lines+markers',
        name='2026 Forecast',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10, color='#10b981', line=dict(color='white', width=2)),
        text=monthly_forecast[y_col].apply(format_func),
        textposition='top center',
        hovertemplate=f'<b>%{{x}} 2026</b><br>{hover_format}<extra></extra>'
    ))
    
    # Add 2024 historical reference line
    if hist_2024 in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast[hist_2024],
            mode='lines+markers',
            name='2024 Actual',
            line=dict(color='#8b5cf6', width=2, dash='dot'),
            marker=dict(size=6, color='#8b5cf6'),
            hovertemplate=f'<b>%{{x}} 2024</b><br>{hover_format}<extra></extra>'
        ))
    
    # Add 2025 historical reference line
    if hist_2025 in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast[hist_2025],
            mode='lines+markers',
            name='2025 Actual',
            line=dict(color='#f59e0b', width=2, dash='dash'),
            marker=dict(size=6, color='#f59e0b'),
            hovertemplate=f'<b>%{{x}} 2025</b><br>{hover_format}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'🔮 2026 {"Revenue" if metric == "Amount" else "Quantity"} Forecast vs Historical',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title=y_title,
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f' if metric == 'Amount' else ',.0f'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


def create_multi_source_stacked_chart(monthly_forecast):
    """
    Create a stacked area chart showing historical baseline + active orders + HubSpot pipeline.
    """
    if monthly_forecast.empty:
        return None
    
    # Check if we have multi-source data
    has_pending = 'Pending_Amount' in monthly_forecast.columns and monthly_forecast['Pending_Amount'].sum() > 0
    has_pipeline = 'Pipeline_Amount' in monthly_forecast.columns and monthly_forecast['Pipeline_Amount'].sum() > 0
    
    fig = go.Figure()
    
    # Historical baseline (bottom layer)
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast['Historical_Baseline_Amount'] if 'Historical_Baseline_Amount' in monthly_forecast.columns else monthly_forecast['Forecasted_Amount'],
        name='Historical Baseline',
        mode='lines',
        line=dict(width=0),
        fillcolor='rgba(34, 197, 94, 0.5)',  # Green
        fill='tozeroy',
        hovertemplate='<b>%{x}</b><br>Historical: $%{y:,.0f}<extra></extra>',
        stackgroup='one'
    ))
    
    # Active orders (middle layer) - only if we have active order data
    if has_pending:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast['Pending_Amount'],
            name='Active Orders',
            mode='lines',
            line=dict(width=0),
            fillcolor='rgba(59, 130, 246, 0.5)',  # Blue
            fill='tonexty',
            hovertemplate='<b>%{x}</b><br>Active Orders: $%{y:,.0f}<extra></extra>',
            stackgroup='one'
        ))
    
    # HubSpot Pipeline (top layer) - only if we have pipeline data
    if has_pipeline:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast['Pipeline_Amount'],
            name='HubSpot Pipeline',
            mode='lines',
            line=dict(width=0),
            fillcolor='rgba(249, 115, 22, 0.5)',  # Orange
            fill='tonexty',
            hovertemplate='<b>%{x}</b><br>Pipeline: $%{y:,.0f}<extra></extra>',
            stackgroup='one'
        ))
    
    # Total line (on top, bold)
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast['Forecasted_Amount'],
        name='Total Forecast',
        mode='lines+markers',
        line=dict(color='rgb(99, 102, 241)', width=3),
        marker=dict(size=8, color='rgb(99, 102, 241)', line=dict(width=2, color='white')),
        hovertemplate='<b>%{x}</b><br>Total: $%{y:,.0f}<extra></extra>'
    ))
    
    # Build title dynamically based on what sources are included
    title_parts = ['Historical']
    if has_pending:
        title_parts.append('Active Orders')
    if has_pipeline:
        title_parts.append('Pipeline')
    title_text = f"📊 2026 Forecast by Source ({' + '.join(title_parts)})"
    
    fig.update_layout(
        title=dict(
            text=title_text,
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


def create_year_comparison_chart(filtered_df, monthly_forecast):
    """
    Create a comparison chart showing 2024 actual, 2025 actual, and 2026 forecast
    """
    if filtered_df.empty or monthly_forecast.empty:
        return None
    
    # Get 2024 and 2025 actual data by month
    monthly_actuals = filtered_df.groupby(['Year', 'Month']).agg({
        'Amount': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    
    # Prepare data for all 12 months
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    month_nums = list(range(1, 13))
    
    # Extract 2024 data
    data_2024 = monthly_actuals[monthly_actuals['Year'] == 2024].set_index('Month')
    revenue_2024 = [data_2024.loc[m, 'Amount'] if m in data_2024.index else 0 for m in month_nums]
    
    # Extract 2025 data
    data_2025 = monthly_actuals[monthly_actuals['Year'] == 2025].set_index('Month')
    revenue_2025 = [data_2025.loc[m, 'Amount'] if m in data_2025.index else 0 for m in month_nums]
    
    # Extract 2026 forecast
    forecast_2026 = monthly_forecast.set_index('Month')
    revenue_2026 = [forecast_2026.loc[m, 'Forecasted_Amount'] if m in forecast_2026.index else 0 for m in month_nums]
    
    fig = go.Figure()
    
    # 2024 bars
    fig.add_trace(go.Bar(
        x=months,
        y=revenue_2024,
        name='2024 Actual',
        marker_color='rgba(139, 92, 246, 0.7)',
        hovertemplate='<b>%{x} 2024</b><br>$%{y:,.0f}<extra></extra>'
    ))
    
    # 2025 bars
    fig.add_trace(go.Bar(
        x=months,
        y=revenue_2025,
        name='2025 Actual',
        marker_color='rgba(249, 115, 22, 0.7)',
        hovertemplate='<b>%{x} 2025</b><br>$%{y:,.0f}<extra></extra>'
    ))
    
    # 2026 bars
    fig.add_trace(go.Bar(
        x=months,
        y=revenue_2026,
        name='2026 Forecast',
        marker_color='rgba(34, 197, 94, 0.7)',
        hovertemplate='<b>%{x} 2026</b><br>$%{y:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='📊 Year-Over-Year Comparison: 2024 vs 2025 vs 2026 Forecast',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
        ),
        barmode='group',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


def create_year_summary_table(filtered_df, monthly_forecast):
    """
    Create a summary table comparing 2024, 2025, and 2026
    """
    if filtered_df.empty or monthly_forecast.empty:
        return None
    
    # Calculate totals
    total_2024_rev = filtered_df[filtered_df['Year'] == 2024]['Amount'].sum()
    total_2024_qty = filtered_df[filtered_df['Year'] == 2024]['Quantity'].sum()
    
    total_2025_rev = filtered_df[filtered_df['Year'] == 2025]['Amount'].sum()
    total_2025_qty = filtered_df[filtered_df['Year'] == 2025]['Quantity'].sum()
    
    total_2026_rev = monthly_forecast['Forecasted_Amount'].sum()
    total_2026_qty = monthly_forecast['Forecasted_Quantity'].sum()
    
    # Calculate growth rates
    growth_25_vs_24 = ((total_2025_rev - total_2024_rev) / total_2024_rev * 100) if total_2024_rev > 0 else 0
    growth_26_vs_25 = ((total_2026_rev - total_2025_rev) / total_2025_rev * 100) if total_2025_rev > 0 else 0
    
    summary_data = {
        'Year': ['2024 Actual', '2025 Actual', '2026 Forecast'],
        'Total Revenue': [
            f'${total_2024_rev:,.0f}',
            f'${total_2025_rev:,.0f}',
            f'${total_2026_rev:,.0f}'
        ],
        'Total Quantity': [
            f'{total_2024_qty:,.0f}',
            f'{total_2025_qty:,.0f}',
            f'{total_2026_qty:,.0f}'
        ],
        'Growth vs Prior Year': [
            '-',
            f'{growth_25_vs_24:+.1f}%',
            f'{growth_26_vs_25:+.1f}%'
        ]
    }
    
    return pd.DataFrame(summary_data)


def create_quarterly_chart(quarterly_forecast, metric='Amount'):
    """
    Create a quarterly summary bar chart.
    """
    if quarterly_forecast.empty:
        return None
    
    if metric == 'Amount':
        y_col = 'Forecasted_Amount'
        y_title = 'Revenue ($)'
        format_func = format_currency
        hover_format = 'Revenue: $%{y:,.0f}'
    else:
        y_col = 'Forecasted_Quantity'
        y_title = 'Quantity (Units)'
        format_func = format_quantity
        hover_format = 'Quantity: %{y:,.0f}'
    
    fig = go.Figure()
    
    # Add bars
    fig.add_trace(go.Bar(
        x=quarterly_forecast['Quarter'],
        y=quarterly_forecast[y_col],
        name=f'Forecasted {metric}',
        marker=dict(
            color=['#6366f1', '#8b5cf6', '#a855f7', '#c084fc'],
            line=dict(color='rgba(255,255,255,0.3)', width=2)
        ),
        text=quarterly_forecast[y_col].apply(format_func),
        textposition='outside',
        hovertemplate=f'<b>%{{x}}</b><br>{hover_format}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text=f'📈 2026 Quarterly {"Revenue" if metric == "Amount" else "Quantity"} Forecast',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Quarter',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title=y_title,
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f' if metric == 'Amount' else ',.0f'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=400,
        margin=dict(t=80, b=60),
        showlegend=False
    )
    
    return fig


def create_product_breakdown_chart(df, group_col='Product Type', metric='Amount'):
    """
    Create a pie/donut chart showing breakdown by product type or item type.
    """
    if df.empty or group_col not in df.columns:
        return None
    
    # Aggregate by group
    grouped = df.groupby(group_col).agg({
        'Amount': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    
    grouped = grouped.sort_values(metric, ascending=False)
    
    # Limit to top 10 + "Other"
    if len(grouped) > 10:
        top_10 = grouped.head(10)
        other_val = grouped.iloc[10:][metric].sum()
        other_row = pd.DataFrame({group_col: ['Other'], metric: [other_val]})
        grouped = pd.concat([top_10, other_row], ignore_index=True)
    
    fig = go.Figure(data=[go.Pie(
        labels=grouped[group_col],
        values=grouped[metric],
        hole=0.4,
        marker=dict(
            colors=px.colors.qualitative.Set3[:len(grouped)]
        ),
        textinfo='label+percent',
        textposition='outside',
        hovertemplate=f'<b>%{{label}}</b><br>{metric}: %{{value:,.0f}}<br>Share: %{{percent}}<extra></extra>'
    )])
    
    fig.update_layout(
        title=dict(
            text=f'📊 {"Revenue" if metric == "Amount" else "Quantity"} by {group_col}',
            font=dict(size=18),
            x=0.5
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=450,
        margin=dict(t=60, b=40),
        showlegend=True,
        legend=dict(
            orientation='v',
            yanchor='middle',
            y=0.5,
            xanchor='left',
            x=1.05
        )
    )
    
    return fig


def create_top_customers_chart(df, top_n=15):
    """
    Create a horizontal bar chart of top customers.
    """
    if df.empty or 'Customer' not in df.columns:
        return None
    
    # Aggregate by customer
    customer_totals = df.groupby('Customer').agg({
        'Amount': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    
    customer_totals = customer_totals.sort_values('Amount', ascending=True).tail(top_n)
    
    # Truncate long customer names
    customer_totals['Display_Name'] = customer_totals['Customer'].apply(
        lambda x: x[:30] + '...' if len(str(x)) > 30 else x
    )
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        y=customer_totals['Display_Name'],
        x=customer_totals['Amount'],
        orientation='h',
        marker=dict(
            color=customer_totals['Amount'],
            colorscale='Viridis',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=customer_totals['Amount'].apply(format_currency),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Revenue: $%{x:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text=f'🏆 Top {top_n} Customers by Revenue',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(
            title='Total Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
        ),
        yaxis=dict(
            title='',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=500,
        margin=dict(l=180, r=80, t=60, b=40)
    )
    
    return fig


def create_product_forecast_comparison(monthly_forecast, group_col):
    """
    Create a stacked area chart comparing forecasts across products/item types.
    """
    if monthly_forecast.empty or group_col not in monthly_forecast.columns:
        return None
    
    # Pivot data for stacked chart
    pivot_df = monthly_forecast.pivot(
        index='MonthShort', 
        columns=group_col, 
        values='Forecasted_Amount'
    ).fillna(0)
    
    fig = go.Figure()
    
    colors = px.colors.qualitative.Set3
    
    for i, col in enumerate(pivot_df.columns):
        fig.add_trace(go.Scatter(
            x=pivot_df.index,
            y=pivot_df[col],
            name=str(col)[:25],
            mode='lines',
            stackgroup='one',
            line=dict(width=0.5, color=colors[i % len(colors)]),
            fillcolor=colors[i % len(colors)],
            hovertemplate=f'<b>{col}</b><br>%{{x}}: $%{{y:,.0f}}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'📊 2026 Revenue Forecast by {group_col}',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


# =============================================================================
# FORECAST ADJUSTMENTS
# =============================================================================

def apply_forecast_adjustments(monthly_forecast, quarterly_forecast, overall_multiplier=1.0, growth_trend=0.0, quarterly_adjustments=None):
    """
    Apply dynamic adjustments to the forecast.
    """
    if monthly_forecast.empty:
        return monthly_forecast, quarterly_forecast
    
    adjusted_monthly = monthly_forecast.copy()
    
    # Apply overall multiplier
    qty_cols = ['Forecasted_Quantity', 'Qty_Low', 'Qty_High']
    amt_cols = ['Forecasted_Amount', 'Amt_Low', 'Amt_High']
    
    for col in qty_cols + amt_cols:
        if col in adjusted_monthly.columns:
            adjusted_monthly[col] = adjusted_monthly[col] * overall_multiplier
    
    # Apply monthly growth trend (compound)
    if growth_trend != 0:
        for i in range(len(adjusted_monthly)):
            growth_factor = (1 + growth_trend/100) ** i
            for col in qty_cols + amt_cols:
                if col in adjusted_monthly.columns:
                    adjusted_monthly.loc[adjusted_monthly.index[i], col] *= growth_factor
    
    # Apply quarterly adjustments
    if quarterly_adjustments:
        for quarter, adj in quarterly_adjustments.items():
            if adj != 0:
                mask = adjusted_monthly['QuarterNum'] == quarter
                for col in qty_cols + amt_cols:
                    if col in adjusted_monthly.columns:
                        adjusted_monthly.loc[mask, col] *= (1 + adj)
    
    # Round quantity columns to integers
    for col in qty_cols:
        if col in adjusted_monthly.columns:
            adjusted_monthly[col] = adjusted_monthly[col].astype(int)
    
    # Round amount columns to 2 decimals
    for col in amt_cols:
        if col in adjusted_monthly.columns:
            adjusted_monthly[col] = adjusted_monthly[col].round(2)
    
    # Regenerate quarterly summary from adjusted monthly
    if not adjusted_monthly.empty:
        adjusted_quarterly = adjusted_monthly.groupby(['QuarterNum', 'Quarter']).agg({
            'Forecasted_Quantity': 'sum',
            'Forecasted_Amount': 'sum',
            'Qty_Low': 'sum',
            'Qty_High': 'sum',
            'Amt_Low': 'sum',
            'Amt_High': 'sum'
        }).reset_index()
        adjusted_quarterly = adjusted_quarterly.sort_values('QuarterNum')
    else:
        adjusted_quarterly = quarterly_forecast.copy()
    
    return adjusted_monthly, adjusted_quarterly


def create_customer_acquisition_analysis(df):
    """
    Analyze customer acquisition patterns - new vs returning customers by month
    """
    if df.empty or 'Customer' not in df.columns or 'Date' not in df.columns:
        return None, None, None
    
    # Get first order date for each customer
    customer_first_order = df.groupby('Customer')['Date'].min().reset_index()
    customer_first_order.columns = ['Customer', 'First_Order_Date']
    customer_first_order['First_Order_Month'] = customer_first_order['First_Order_Date'].dt.to_period('M')
    
    # Merge back to main df
    df_with_cohort = df.merge(customer_first_order[['Customer', 'First_Order_Month']], on='Customer')
    df_with_cohort['Order_Month'] = df_with_cohort['Date'].dt.to_period('M')
    
    # Calculate monthly active customers and revenue
    monthly_stats = df_with_cohort.groupby('Order_Month').agg({
        'Customer': 'nunique',
        'Amount': 'sum'
    }).reset_index()
    monthly_stats.columns = ['Month', 'Active_Customers', 'Revenue']
    monthly_stats['Month_Str'] = monthly_stats['Month'].astype(str)
    
    # Calculate new customer acquisitions by month
    new_customers = customer_first_order.groupby('First_Order_Month').size().reset_index()
    new_customers.columns = ['Month', 'New_Customers']
    new_customers['Month_Str'] = new_customers['Month'].astype(str)
    
    return monthly_stats, new_customers, customer_first_order


def create_monthly_active_customers_chart(monthly_stats):
    """
    Create chart showing monthly active customers and revenue
    """
    if monthly_stats is None or monthly_stats.empty:
        return None
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    # Active customers bars
    fig.add_trace(
        go.Bar(
            x=monthly_stats['Month_Str'],
            y=monthly_stats['Active_Customers'],
            name='Active Customers',
            marker_color='rgba(99, 102, 241, 0.7)',
            yaxis='y',
            hovertemplate='<b>%{x}</b><br>Customers: %{y}<extra></extra>'
        ),
        secondary_y=False
    )
    
    # Revenue line
    fig.add_trace(
        go.Scatter(
            x=monthly_stats['Month_Str'],
            y=monthly_stats['Revenue'],
            name='Revenue',
            mode='lines+markers',
            line=dict(color='rgba(34, 197, 94, 1)', width=3),
            marker=dict(size=8),
            yaxis='y2',
            hovertemplate='<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>'
        ),
        secondary_y=True
    )
    
    fig.update_xaxes(
        title_text="Month",
        gridcolor='rgba(128,128,128,0.1)',
        tickangle=-45
    )
    
    fig.update_yaxes(
        title_text="Active Customers",
        secondary_y=False,
        gridcolor='rgba(128,128,128,0.2)'
    )
    
    fig.update_yaxes(
        title_text="Revenue ($)",
        secondary_y=True,
        tickformat='$,.0f'
    )
    
    fig.update_layout(
        title=dict(
            text='📊 Monthly Active Customers & Revenue',
            font=dict(size=20),
            x=0.5
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1
        ),
        height=500,
        margin=dict(t=80, b=100),
        hovermode='x unified'
    )
    
    return fig


def create_new_customer_acquisition_chart(new_customers):
    """
    Create chart showing new customer acquisition by cohort month
    """
    if new_customers is None or new_customers.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=new_customers['Month_Str'],
        y=new_customers['New_Customers'],
        marker_color='rgba(139, 92, 246, 0.7)',
        text=new_customers['New_Customers'],
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>New Customers: %{y}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='📈 New Customer Acquisition by Month',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Cohort (First Order Month)',
            gridcolor='rgba(128,128,128,0.1)',
            tickangle=-45
        ),
        yaxis=dict(
            title='New Customers',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=500,
        margin=dict(t=80, b=100),
        showlegend=False
    )
    
    return fig


def create_product_acquisition_analysis(df):
    """
    Analyze product adoption - which products attract new customers
    """
    if df.empty or 'Customer' not in df.columns or 'Product Type' not in df.columns:
        return None
    
    # Get first order for each customer
    customer_first_order = df.sort_values('Date').groupby('Customer').first().reset_index()
    
    # Count which product types brought in new customers
    product_acquisition = customer_first_order.groupby('Product Type').agg({
        'Customer': 'count',
        'Amount': 'sum'
    }).reset_index()
    product_acquisition.columns = ['Product Type', 'New_Customers', 'First_Order_Revenue']
    product_acquisition = product_acquisition.sort_values('New_Customers', ascending=False)
    
    return product_acquisition


def create_product_acquisition_chart(product_acquisition):
    """
    Create chart showing which products acquire new customers
    """
    if product_acquisition is None or product_acquisition.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=product_acquisition['Product Type'].head(15),
        y=product_acquisition['New_Customers'].head(15),
        marker_color='rgba(249, 115, 22, 0.7)',
        text=product_acquisition['New_Customers'].head(15),
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>New Customers: %{y}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='🎯 New Customer Acquisition by Product Type',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Product Type',
            gridcolor='rgba(128,128,128,0.1)',
            tickangle=-45
        ),
        yaxis=dict(
            title='New Customers Acquired',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=500,
        margin=dict(t=80, b=120),
        showlegend=False
    )
    
    return fig


def remove_outliers(df, column='Amount', threshold_std=3.0):
    """
    Remove statistical outliers from dataset
    Uses z-score method: removes values beyond threshold standard deviations
    """
    if df.empty or column not in df.columns:
        return df
    
    # Calculate z-scores
    mean = df[column].mean()
    std = df[column].std()
    
    if std == 0:
        return df
    
    z_scores = np.abs((df[column] - mean) / std)
    
    # Filter out outliers
    filtered_df = df[z_scores < threshold_std].copy()
    
    removed_count = len(df) - len(filtered_df)
    if removed_count > 0:
        removed_value = df[z_scores >= threshold_std][column].sum()
        st.sidebar.caption(f"🧹 Removed {removed_count} outliers (${removed_value:,.0f})")
    
    return filtered_df


def smooth_product_data(df, min_revenue=50000, min_orders=5):
    """
    Filter products to only include meaningful ones for analysis
    - Removes low-revenue products
    - Removes products with too few orders (noise)
    """
    if df.empty or 'Product Type' not in df.columns:
        return df
    
    # Calculate product statistics
    product_stats = df.groupby('Product Type').agg({
        'Amount': ['sum', 'count'],
        'Customer': 'nunique'
    }).reset_index()
    
    product_stats.columns = ['Product Type', 'Total_Revenue', 'Order_Count', 'Customer_Count']
    
    # Filter criteria
    valid_products = product_stats[
        (product_stats['Total_Revenue'] >= min_revenue) &
        (product_stats['Order_Count'] >= min_orders)
    ]['Product Type'].tolist()
    
    filtered_df = df[df['Product Type'].isin(valid_products)].copy()
    
    removed_products = len(product_stats) - len(valid_products)
    if removed_products > 0:
        st.sidebar.caption(f"📊 Filtered to {len(valid_products)} meaningful products ({removed_products} low-volume removed)")
    
    return filtered_df


def calculate_product_growth_potential(df):
    """
    Calculate growth potential for each product type based on historical trends
    Only includes products meeting minimum revenue thresholds
    """
    if df.empty or 'Product Type' not in df.columns:
        return None
    
    # Calculate year-over-year growth by product
    product_by_year = df.groupby(['Product Type', 'Year']).agg({
        'Amount': 'sum',
        'Customer': 'nunique'
    }).reset_index()
    
    growth_rates = []
    for product in product_by_year['Product Type'].unique():
        product_data = product_by_year[product_by_year['Product Type'] == product].sort_values('Year')
        
        if len(product_data) >= 2:
            # Get 2024 and 2025 values
            rev_2024 = product_data[product_data['Year'] == 2024]['Amount'].sum()
            rev_2025 = product_data[product_data['Year'] == 2025]['Amount'].sum()
            
            # Only include if 2025 revenue meets minimum threshold
            if rev_2024 > 0 and rev_2025 >= MIN_REVENUE_FOR_RECOMMENDATION:
                growth_rate = ((rev_2025 - rev_2024) / rev_2024) * 100
                
                growth_rates.append({
                    'Product Type': product,
                    'Revenue_2024': rev_2024,
                    'Revenue_2025': rev_2025,
                    'Growth_Rate': growth_rate,
                    'Revenue_Increase': rev_2025 - rev_2024
                })
    
    if not growth_rates:
        return None
    
    growth_df = pd.DataFrame(growth_rates)
    growth_df = growth_df.sort_values('Growth_Rate', ascending=False)
    
    return growth_df


def calculate_gap_to_goal(current_forecast, goal_amount):
    """
    Calculate the gap between current forecast and goal
    """
    gap = goal_amount - current_forecast
    gap_pct = (gap / current_forecast * 100) if current_forecast > 0 else 0
    
    return {
        'current_forecast': current_forecast,
        'goal': goal_amount,
        'gap': gap,
        'gap_pct': gap_pct,
        'needs_growth': gap > 0
    }


def recommend_product_focus(growth_potential, gap_amount, top_n=5):
    """
    Recommend which products to focus on to close the revenue gap
    """
    if growth_potential is None or growth_potential.empty:
        return None
    
    # Sort by growth rate and revenue size
    # Prioritize products that are both growing fast AND have meaningful revenue
    growth_potential['Score'] = (
        growth_potential['Growth_Rate'] * 0.6 +  # 60% weight on growth rate
        (growth_potential['Revenue_2025'] / growth_potential['Revenue_2025'].max() * 100) * 0.4  # 40% weight on size
    )
    
    recommendations = growth_potential.sort_values('Score', ascending=False).head(top_n)
    
    # Calculate how much each product could contribute if growth continues
    recommendations['Potential_2026'] = recommendations['Revenue_2025'] * (1 + recommendations['Growth_Rate']/100)
    recommendations['Additional_Revenue'] = recommendations['Potential_2026'] - recommendations['Revenue_2025']
    
    # Calculate what % of gap each could close
    recommendations['Gap_Coverage'] = (recommendations['Additional_Revenue'] / gap_amount * 100) if gap_amount > 0 else 0
    
    return recommendations


def apply_product_scenario(monthly_forecast, product_adjustments):
    """
    Apply growth rate adjustments to specific product types
    product_adjustments: dict like {'Concentrate Jars': 1.25, 'Flower Jars': 1.10}
    """
    # This would need product-level forecast data
    # For now, return a scenario summary
    adjusted_forecast = monthly_forecast.copy()
    
    # Apply overall multiplier based on weighted average of adjustments
    if product_adjustments:
        avg_multiplier = np.mean(list(product_adjustments.values()))
        adjusted_forecast['Forecasted_Amount'] = adjusted_forecast['Forecasted_Amount'] * avg_multiplier
        adjusted_forecast['Forecasted_Quantity'] = adjusted_forecast['Forecasted_Quantity'] * avg_multiplier
    
    return adjusted_forecast


def create_gap_analysis_visual(gap_info):
    """
    Create a visual showing current forecast vs goal
    """
    fig = go.Figure()
    
    # Current forecast bar
    fig.add_trace(go.Bar(
        x=['Current Forecast'],
        y=[gap_info['current_forecast']],
        name='Current Forecast',
        marker_color='rgba(99, 102, 241, 0.7)',
        text=[f"${gap_info['current_forecast']:,.0f}"],
        textposition='inside',
        textfont=dict(size=16, color='white')
    ))
    
    # Goal bar
    fig.add_trace(go.Bar(
        x=['2026 Goal'],
        y=[gap_info['goal']],
        name='Goal',
        marker_color='rgba(34, 197, 94, 0.7)',
        text=[f"${gap_info['goal']:,.0f}"],
        textposition='inside',
        textfont=dict(size=16, color='white')
    ))
    
    # Gap indicator
    if gap_info['needs_growth']:
        fig.add_annotation(
            x=1,
            y=gap_info['current_forecast'] + (gap_info['gap'] / 2),
            text=f"GAP: ${gap_info['gap']:,.0f}<br>({gap_info['gap_pct']:.1f}% growth needed)",
            showarrow=True,
            arrowhead=2,
            arrowsize=1,
            arrowwidth=2,
            arrowcolor='rgba(249, 115, 22, 1)',
            font=dict(size=14, color='rgba(249, 115, 22, 1)'),
            bgcolor='rgba(0,0,0,0.7)',
            bordercolor='rgba(249, 115, 22, 1)',
            borderwidth=2
        )
    
    fig.update_layout(
        title=dict(
            text='🎯 Revenue Goal vs Current Forecast',
            font=dict(size=20),
            x=0.5
        ),
        yaxis=dict(
            title='Revenue ($)',
            tickformat='$,.0f',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=400,
        showlegend=False
    )
    
    return fig


def create_product_recommendation_chart(recommendations):
    """
    Create chart showing recommended product focus areas
    """
    if recommendations is None or recommendations.empty:
        return None
    
    fig = go.Figure()
    
    # Growth rate bars
    fig.add_trace(go.Bar(
        x=recommendations['Product Type'],
        y=recommendations['Growth_Rate'],
        marker_color='rgba(34, 197, 94, 0.7)',
        text=recommendations['Growth_Rate'].apply(lambda x: f"{x:+.1f}%"),
        textposition='outside',
        name='Historical Growth Rate',
        hovertemplate='<b>%{x}</b><br>Growth: %{y:.1f}%<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='🎯 Recommended Product Focus (By Growth Potential)',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Product Type',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='YoY Growth Rate (%)',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=500,
        margin=dict(t=80, b=120),
        showlegend=False
    )
    
    return fig


def get_rep_customers(df, sales_rep):
    """
    Get list of customers for a specific sales rep
    """
    if 'Sales Rep' not in df.columns:
        return []
    
    rep_customers = df[df['Sales Rep'] == sales_rep]['Customer'].unique().tolist()
    return sorted(rep_customers)


def calculate_rep_forecast(monthly_forecast, df, sales_rep, excluded_customers=None):
    """
    Calculate forecast for a specific sales rep's territory
    """
    if 'Sales Rep' not in df.columns:
        return None, None
    
    # Filter to rep's historical data
    rep_df = df[df['Sales Rep'] == sales_rep].copy()
    
    # Exclude churned customers if specified
    if excluded_customers:
        rep_df = rep_df[~rep_df['Customer'].isin(excluded_customers)]
    
    # Calculate rep's % of total revenue
    total_revenue = df['Amount'].sum()
    rep_revenue = rep_df['Amount'].sum()
    rep_pct = (rep_revenue / total_revenue) if total_revenue > 0 else 0
    
    # Apply rep % to forecast
    rep_forecast = monthly_forecast.copy()
    rep_forecast['Forecasted_Amount'] = rep_forecast['Forecasted_Amount'] * rep_pct
    rep_forecast['Forecasted_Quantity'] = rep_forecast['Forecasted_Quantity'] * rep_pct
    
    rep_stats = {
        'total_revenue': rep_revenue,
        'pct_of_total': rep_pct * 100,
        'customer_count': rep_df['Customer'].nunique(),
        'avg_order_size': rep_revenue / len(rep_df) if len(rep_df) > 0 else 0
    }
    
    return rep_forecast, rep_stats


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """
    Main function for All Products Forecasting dashboard.
    Combines historical invoice data with pending sales orders.
    """
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.2) 100%);
        border: 1px solid rgba(99, 102, 241, 0.3);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 24px;
    ">
        <h1 style="margin: 0; font-size: 28px; display: flex; align-items: center; gap: 12px;">
            📦 All Products Forecast - 2026 (Multi-Source)
        </h1>
        <p style="margin: 8px 0 0 0; opacity: 0.8;">
            Revenue and quantity forecasting combining historical invoices + active sales orders + HubSpot pipeline
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # =========================
    # LOAD DATA FROM ALL THREE SOURCES
    # =========================
    
    # Load Invoice Line Items (historical data)
    with st.spinner("Loading Invoice Line Item data..."):
        raw_invoice_df = load_invoice_line_items()
    
    if raw_invoice_df.empty:
        st.error("❌ No invoice data loaded. Check your Google Sheets connection and ensure the 'Invoice Line Item' tab exists.")
        return
    
    # Process invoice data
    invoice_df = process_invoice_data(raw_invoice_df)
    
    if invoice_df.empty:
        st.error("❌ Invoice data processing failed. Check that required columns exist.")
        return
    
    # Load Sales Order Line Items (active orders)
    with st.spinner("Loading Sales Order Line Item data..."):
        raw_so_df = load_sales_order_line_items()
    
    # Process and filter sales orders
    sales_order_df = pd.DataFrame()
    pending_order_count = 0
    pending_order_amount = 0
    
    if not raw_so_df.empty:
        sales_order_df = process_sales_order_data(raw_so_df)
        
        # Deduplicate - remove already invoiced orders
        if not sales_order_df.empty:
            sales_order_df = filter_uninvoiced_sales_orders(sales_order_df, invoice_df)
            
            if not sales_order_df.empty:
                pending_order_count = len(sales_order_df)
                pending_order_amount = sales_order_df['Amount'].sum()
    
    # Load HubSpot Pipeline Data (future opportunities)
    with st.spinner("Loading HubSpot Pipeline data..."):
        raw_hubspot_df = load_hubspot_data()
    
    # Process HubSpot data
    hubspot_df = pd.DataFrame()
    hubspot_deal_count = 0
    hubspot_weighted_amount = 0
    
    if not raw_hubspot_df.empty:
        hubspot_df = process_hubspot_data(raw_hubspot_df)
        
        if not hubspot_df.empty:
            hubspot_deal_count = len(hubspot_df)
            hubspot_weighted_amount = hubspot_df['Weighted_Amount'].sum()
    
    # Show data summary
    col1, col2, col3 = st.columns(3)
    with col1:
        st.success(f"✅ Loaded {len(invoice_df):,} invoice line items (historical)")
    with col2:
        if pending_order_count > 0:
            st.success(f"✅ Loaded {pending_order_count:,} active sales order lines (${pending_order_amount:,.0f})")
        else:
            st.info("ℹ️ No active sales orders found")
    with col3:
        if hubspot_deal_count > 0:
            st.success(f"✅ Loaded {hubspot_deal_count:,} HubSpot deals (${hubspot_weighted_amount:,.0f} weighted)")
        else:
            st.info("ℹ️ No HubSpot pipeline data found")
    
    # Use invoice_df as the base for filtering (historical data)
    df = invoice_df
    
    # =========================
    # DEBUG INFO - Show actual totals
    # =========================
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Data Loaded")
    
    if not df.empty:
        total_2024 = df[df['Year'] == 2024]['Amount'].sum()
        total_2025 = df[df['Year'] == 2025]['Amount'].sum()
        count_2024 = len(df[df['Year'] == 2024])
        count_2025 = len(df[df['Year'] == 2025])
        
        st.sidebar.metric("2024 Revenue", f"${total_2024:,.0f}", delta=f"{count_2024:,} invoices")
        st.sidebar.metric("2025 Revenue", f"${total_2025:,.0f}", delta=f"{count_2025:,} invoices")
        
        if total_2024 > 0:
            yoy_change = ((total_2025 - total_2024) / total_2024) * 100
            st.sidebar.caption(f"YoY Change: {yoy_change:+.1f}%")
        
        # Product Type diagnostic
        if 'Product Type' in df.columns:
            unique_products = df['Product Type'].nunique()
            st.sidebar.caption(f"📦 {unique_products} Product Types in invoices")
            
            with st.sidebar.expander("🔍 View All Product Types"):
                product_list = sorted(df['Product Type'].unique().tolist())
                for product in product_list:
                    product_total = df[df['Product Type'] == product]['Amount'].sum()
                    st.sidebar.caption(f"• {product}: ${product_total:,.0f}")
    
    # =========================
    # SIDEBAR CONTROLS
    # =========================
    
    st.sidebar.markdown("## 📦 Forecast Controls")
    
    # Data source toggle
    st.sidebar.markdown("### 📊 Data Sources")
    include_pending_orders = st.sidebar.checkbox(
        "Include Active Sales Orders",
        value=True if not sales_order_df.empty else False,
        disabled=sales_order_df.empty,
        help="Add active sales orders to the forecast (all uninvoiced orders)"
    )
    
    if include_pending_orders and not sales_order_df.empty:
        st.sidebar.success(f"✅ {pending_order_count:,} active orders included")
    
    include_hubspot_pipeline = st.sidebar.checkbox(
        "Include HubSpot Pipeline",
        value=True if not hubspot_df.empty else False,
        disabled=hubspot_df.empty,
        help="Add HubSpot pipeline deals to the forecast (weighted by stage probability)"
    )
    
    if include_hubspot_pipeline and not hubspot_df.empty:
        st.sidebar.success(f"✅ {hubspot_deal_count:,} pipeline deals included")
    
    # Historical weighting
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚖️ Historical Weights")
    weight_2024 = st.sidebar.slider(
        "2024 Weight", 0.0, 1.0, 0.3, 0.05,  # Changed from 0.6 to 0.3
        help="Weight given to 2024 data (lower = less influence, more weight on recent 2025 trends)"
    )
    weight_2025 = 1.0 - weight_2024
    st.sidebar.caption(f"2025 Weight: {weight_2025:.0%}")
    
    # Filter options
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filters")
    
    # Churn Management
    st.sidebar.markdown("#### ⚠️ Churn Risk Management")
    
    if 'Customer' in df.columns:
        all_customers = sorted(df['Customer'].dropna().unique().tolist())
        excluded_customers = st.sidebar.multiselect(
            "Exclude Customers (Churn Risk)",
            options=all_customers,
            default=[],
            help="Select customers to exclude from forecast (at-risk of churning)"
        )
        
        if excluded_customers:
            excluded_revenue = df[df['Customer'].isin(excluded_customers)]['Amount'].sum()
            st.sidebar.error(f"Excluding ${excluded_revenue:,.0f} in historical revenue from {len(excluded_customers)} customer(s)")
    else:
        excluded_customers = []
    
    st.sidebar.markdown("---")
    
    # Product Type filter
    if 'Product Type' in df.columns:
        product_types = ['All'] + sorted(df['Product Type'].unique().tolist())
        selected_product_type = st.sidebar.selectbox(
            "Product Type",
            product_types,
            help="Filter by PI || Product Type"
        )
    else:
        selected_product_type = 'All'
    
    # Item Type filter
    if 'Item Type' in df.columns:
        item_types = ['All'] + sorted(df['Item Type'].unique().tolist())
        selected_item_type = st.sidebar.selectbox(
            "Item Type",
            item_types,
            help="Filter by Calyx | Item Type"
        )
    else:
        selected_item_type = 'All'
    
    # Customer filter
    if 'Customer' in df.columns:
        customers = ['All'] + sorted(df['Customer'].dropna().unique().tolist())
        selected_customer = st.sidebar.selectbox(
            "Customer",
            customers[:100],  # Limit dropdown size
            help="Filter by customer (top 100 shown)"
        )
    else:
        selected_customer = 'All'
    
    # Sales Rep filter
    if 'Sales Rep' in df.columns:
        sales_reps = ['All'] + sorted(df['Sales Rep'].dropna().unique().tolist())
        selected_rep = st.sidebar.selectbox(
            "Sales Rep",
            sales_reps,
            help="Filter by sales representative"
        )
    else:
        selected_rep = 'All'
    
    # Apply filters
    filtered_df = df.copy()
    filter_desc = []
    
    # Apply churn exclusions first
    if excluded_customers:
        filtered_df = filtered_df[~filtered_df['Customer'].isin(excluded_customers)]
        filter_desc.append(f"Excluded {len(excluded_customers)} at-risk customer(s)")
    
    if selected_product_type != 'All':
        filtered_df = filtered_df[filtered_df['Product Type'] == selected_product_type]
        filter_desc.append(f"Product Type: {selected_product_type}")
    
    if selected_item_type != 'All':
        filtered_df = filtered_df[filtered_df['Item Type'] == selected_item_type]
        filter_desc.append(f"Item Type: {selected_item_type}")
    
    if selected_customer != 'All':
        filtered_df = filtered_df[filtered_df['Customer'] == selected_customer]
        filter_desc.append(f"Customer: {selected_customer}")
    
    if selected_rep != 'All':
        filtered_df = filtered_df[filtered_df['Sales Rep'] == selected_rep]
        filter_desc.append(f"Sales Rep: {selected_rep}")
    
    # Data quality settings for PLANNING/RECOMMENDATIONS only
    # These do NOT affect historical actuals or baseline forecast
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🧹 Planning Data Quality")
    st.sidebar.caption("⚠️ Only affects Goal Planning recommendations, NOT historical actuals")
    
    apply_outlier_removal_planning = st.sidebar.checkbox(
        "Remove Outliers (Planning Only)",
        value=False,
        help=f"For recommendations: Remove transactions beyond {OUTLIER_THRESHOLD_STD} standard deviations. Does NOT affect historical baseline."
    )
    
    apply_smoothing_planning = st.sidebar.checkbox(
        "Filter Low-Volume Products (Planning)",
        value=False,
        help=f"For recommendations: Only include products with >${MIN_REVENUE_FOR_RECOMMENDATION/1000:.0f}K revenue and >{MIN_ORDERS_FOR_ANALYSIS} orders. Does NOT affect historical baseline."
    )
    
    # Create separate planning dataframe (for Goal Planning tab recommendations)
    # This is ONLY used for product recommendations, NOT for historical baseline
    planning_df = filtered_df.copy()
    
    if apply_outlier_removal_planning:
        planning_df = remove_outliers(planning_df, column='Amount', threshold_std=OUTLIER_THRESHOLD_STD)
    
    if apply_smoothing_planning:
        planning_df = smooth_product_data(planning_df, min_revenue=MIN_REVENUE_FOR_RECOMMENDATION, min_orders=MIN_ORDERS_FOR_ANALYSIS)
    
    # Dynamic forecast adjustments
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎛️ Forecast Adjustments")
    
    overall_multiplier = st.sidebar.slider(
        "Overall Forecast Multiplier", 
        0.5, 2.0, 1.0, 0.05,
        help="Adjust entire forecast up or down (1.0 = no change)"
    )
    
    growth_trend = st.sidebar.slider(
        "Monthly Growth Trend %",
        -5.0, 5.0, 0.0, 0.5,
        help="Apply compound monthly growth/decline"
    )
    
    # Quarterly adjustments
    with st.sidebar.expander("📅 Quarterly Adjustments", expanded=False):
        q1_adj = st.slider("Q1 Adjustment %", -50, 50, 0, 5, key="q1_adj_all")
        q2_adj = st.slider("Q2 Adjustment %", -50, 50, 0, 5, key="q2_adj_all")
        q3_adj = st.slider("Q3 Adjustment %", -50, 50, 0, 5, key="q3_adj_all")
        q4_adj = st.slider("Q4 Adjustment %", -50, 50, 0, 5, key="q4_adj_all")
    
    quarterly_adjustments = {1: q1_adj/100, 2: q2_adj/100, 3: q3_adj/100, 4: q4_adj/100}
    
    # Show active filters
    if filter_desc:
        st.info(f"🔍 Active Filters: {' | '.join(filter_desc)} ({len(filtered_df):,} records)")
    
    # =========================
    # GENERATE FORECASTS
    # =========================
    
    if filtered_df.empty:
        st.warning("⚠️ No data matches the selected filters.")
        return
    
    # Generate historical baseline forecast
    monthly_historical, quarterly_historical, monthly_baselines = generate_2026_forecast(
        filtered_df, weight_2024=weight_2024, weight_2025=weight_2025
    )
    
    # Calculate pending orders forecast (if enabled and available)
    monthly_pending = pd.DataFrame()
    if include_pending_orders and not sales_order_df.empty:
        # Apply same filters to sales orders
        filtered_so_df = sales_order_df.copy()
        
        if selected_product_type != 'All' and 'Product Type' in filtered_so_df.columns:
            filtered_so_df = filtered_so_df[filtered_so_df['Product Type'] == selected_product_type]
        
        if selected_item_type != 'All' and 'Item Type' in filtered_so_df.columns:
            filtered_so_df = filtered_so_df[filtered_so_df['Item Type'] == selected_item_type]
        
        if selected_customer != 'All' and 'Customer' in filtered_so_df.columns:
            filtered_so_df = filtered_so_df[filtered_so_df['Customer'] == selected_customer]
        
        if not filtered_so_df.empty:
            monthly_pending = calculate_pending_orders_forecast(filtered_so_df)
    
    # Calculate HubSpot pipeline forecast (if enabled and available)
    monthly_hubspot = pd.DataFrame()
    if include_hubspot_pipeline and not hubspot_df.empty:
        # Apply same filters to HubSpot data (where applicable)
        filtered_hs_df = hubspot_df.copy()
        
        # Note: HubSpot might not have exact same Product Type/Item Type structure
        # Filter by Product_Name if it matches selected_product_type
        if selected_product_type != 'All' and 'Product Type' in filtered_hs_df.columns:
            filtered_hs_df = filtered_hs_df[filtered_hs_df['Product Type'] == selected_product_type]
        
        if selected_customer != 'All' and 'Company_Name' in filtered_hs_df.columns:
            filtered_hs_df = filtered_hs_df[filtered_hs_df['Company_Name'] == selected_customer]
        
        if not filtered_hs_df.empty:
            monthly_hubspot = calculate_hubspot_forecast(filtered_hs_df)
    
    # Combine all three sources
    has_pending = not monthly_pending.empty
    has_hubspot = not monthly_hubspot.empty
    
    if has_pending or has_hubspot:
        monthly_forecast, quarterly_forecast = combine_forecast_sources(
            monthly_historical, 
            monthly_pending if has_pending else pd.DataFrame(),
            monthly_hubspot if has_hubspot else pd.DataFrame()
        )
    else:
        monthly_forecast = monthly_historical
        quarterly_forecast = quarterly_historical
        # Add source columns for consistency
        if not monthly_forecast.empty:
            monthly_forecast['Pending_Amount'] = 0
            monthly_forecast['Pending_Quantity'] = 0
            monthly_forecast['Order_Count'] = 0
            monthly_forecast['Pipeline_Amount'] = 0
            monthly_forecast['Pipeline_Quantity'] = 0
            monthly_forecast['Deal_Count'] = 0
            monthly_forecast['Historical_Baseline_Amount'] = monthly_forecast['Forecasted_Amount']
            monthly_forecast['Historical_Baseline_Quantity'] = monthly_forecast['Forecasted_Quantity']
    
    # Store base forecast for comparison
    base_monthly_forecast = monthly_forecast.copy() if not monthly_forecast.empty else pd.DataFrame()
    base_total_amt = monthly_forecast['Forecasted_Amount'].sum() if not monthly_forecast.empty else 0
    
    # Apply adjustments
    if not monthly_forecast.empty:
        monthly_forecast, quarterly_forecast = apply_forecast_adjustments(
            monthly_forecast, 
            quarterly_forecast,
            overall_multiplier=overall_multiplier,
            growth_trend=growth_trend,
            quarterly_adjustments=quarterly_adjustments
        )
    
    # =========================
    # TOP METRICS ROW
    # =========================
    
    if not monthly_forecast.empty:
        total_qty_2026 = monthly_forecast['Forecasted_Quantity'].sum()
        total_amt_2026 = monthly_forecast['Forecasted_Amount'].sum()
        total_historical = monthly_forecast['Historical_Baseline_Amount'].sum()
        total_pending = monthly_forecast['Pending_Amount'].sum()
        total_pipeline = monthly_forecast['Pipeline_Amount'].sum() if 'Pipeline_Amount' in monthly_forecast.columns else 0
        q1_qty = monthly_forecast[monthly_forecast['QuarterNum'] == 1]['Forecasted_Quantity'].sum()
        q1_amt = monthly_forecast[monthly_forecast['QuarterNum'] == 1]['Forecasted_Amount'].sum()
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            delta_components = []
            if total_pending > 0:
                delta_components.append(f"${total_pending:,.0f} active")
            if total_pipeline > 0:
                delta_components.append(f"${total_pipeline:,.0f} pipeline")
            delta_text = " + ".join(delta_components) if delta_components else None
            
            st.metric(
                "2026 Total Revenue",
                f"${total_amt_2026:,.0f}",
                delta=delta_text
            )
        
        with col2:
            st.metric(
                "Historical Baseline",
                f"${total_historical:,.0f}",
                delta=f"{(total_historical/total_amt_2026*100):.0f}% of total"
            )
        
        with col3:
            if total_pending > 0:
                pending_orders_count = monthly_forecast['Order_Count'].sum()
                st.metric(
                    "Active Orders",
                    f"${total_pending:,.0f}",
                    delta=f"{int(pending_orders_count)} orders"
                )
            else:
                st.metric(
                    "Active Orders",
                    "$0",
                    delta="None loaded"
                )
        
        with col4:
            if total_pipeline > 0:
                pipeline_deals_count = monthly_forecast['Deal_Count'].sum()
                st.metric(
                    "HubSpot Pipeline",
                    f"${total_pipeline:,.0f}",
                    delta=f"{int(pipeline_deals_count)} deals"
                )
            else:
                st.metric(
                    "HubSpot Pipeline",
                    "$0",
                    delta="None loaded"
                )
        
        with col5:
            st.metric(
                "Q1 2026 Revenue",
                f"${q1_amt:,.0f}",
                delta="Highest confidence"
            )
    
    st.markdown("---")
    
    # =========================
    # TABS FOR DIFFERENT VIEWS
    # =========================
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📈 Forecast Overview",
        "🎯 Goal Planning",
        "👥 Sales Rep Planning",
        "📊 Product Breakdown",
        "🏆 Customer Analysis",
        "📋 Detailed Data",
        "📖 Methodology"
    ])
    
    with tab1:
        # Year-over-year comparison
        st.markdown("### 📊 Year-Over-Year Comparison")
        
        year_comparison_chart = create_year_comparison_chart(filtered_df, monthly_forecast)
        if year_comparison_chart:
            st.plotly_chart(year_comparison_chart, use_container_width=True)
        
        # Summary table
        summary_table = create_year_summary_table(filtered_df, monthly_forecast)
        if summary_table is not None:
            st.markdown("#### 📋 Annual Summary")
            st.dataframe(summary_table, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # Historical trend
        st.markdown("### 📈 Historical Revenue Trend")
        hist_chart = create_historical_trend_chart(filtered_df)
        if hist_chart:
            st.plotly_chart(hist_chart, use_container_width=True)
        
        st.markdown("---")
        
        # 2026 Forecast
        st.markdown("### 🔮 2026 Forecast")
        
        metric_choice = st.radio(
            "View metric:",
            ["Revenue", "Quantity"],
            horizontal=True,
            key="forecast_metric"
        )
        
        metric_type = 'Amount' if metric_choice == 'Revenue' else 'Quantity'
        
        if not monthly_forecast.empty:
            # Multi-source breakdown chart (if we have pending orders)
            has_pending = 'Pending_Amount' in monthly_forecast.columns and monthly_forecast['Pending_Amount'].sum() > 0
            if has_pending:
                st.markdown("#### 📊 Forecast by Source")
                multi_source_chart = create_multi_source_stacked_chart(monthly_forecast)
                if multi_source_chart:
                    st.plotly_chart(multi_source_chart, use_container_width=True)
                
                st.markdown("---")
            
            # Standard forecast chart
            forecast_chart = create_forecast_chart(monthly_forecast, metric=metric_type)
            if forecast_chart:
                st.plotly_chart(forecast_chart, use_container_width=True)
        
        # Quarterly summary
        st.markdown("### 📊 Quarterly Summary")
        
        if not quarterly_forecast.empty:
            col1, col2 = st.columns(2)
            
            with col1:
                quarterly_chart = create_quarterly_chart(quarterly_forecast, metric=metric_type)
                if quarterly_chart:
                    st.plotly_chart(quarterly_chart, use_container_width=True)
            
            with col2:
                for _, row in quarterly_forecast.iterrows():
                    st.markdown(f"""
                    <div style="
                        background: linear-gradient(135deg, rgba(99, 102, 241, 0.15) 0%, rgba(139, 92, 246, 0.15) 100%);
                        border: 1px solid rgba(99, 102, 241, 0.3);
                        border-radius: 12px;
                        padding: 16px;
                        margin-bottom: 12px;
                    ">
                        <div style="font-size: 18px; font-weight: 700; margin-bottom: 8px;">{row['Quarter']}</div>
                        <div style="display: flex; justify-content: space-between;">
                            <div>
                                <div style="font-size: 12px; opacity: 0.7;">Revenue</div>
                                <div style="font-size: 20px; font-weight: 600; color: #10b981;">${row['Forecasted_Amount']:,.0f}</div>
                                <div style="font-size: 11px; opacity: 0.6;">${row['Amt_Low']:,.0f} - ${row['Amt_High']:,.0f}</div>
                            </div>
                            <div style="text-align: right;">
                                <div style="font-size: 12px; opacity: 0.7;">Quantity</div>
                                <div style="font-size: 20px; font-weight: 600;">{row['Forecasted_Quantity']:,.0f}</div>
                                <div style="font-size: 11px; opacity: 0.6;">{row['Qty_Low']:,.0f} - {row['Qty_High']:,.0f}</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### 🎯 Strategic Goal Planning")
        
        # Explanation box
        st.info("""
        **📌 Note on Data Quality Settings:**
        
        The "Planning Data Quality" settings in the sidebar (Remove Outliers, Filter Low-Volume Products) **only affect 
        product recommendations** on this tab. They do NOT change your historical actuals or baseline forecast.
        
        - ✅ Historical 2024/2025 revenue = Always shows actual data (never filtered)
        - ✅ Baseline forecast = Based on actual historical data (never filtered)
        - 🎯 Product recommendations below = Optionally filtered to remove noise from planning
        """)
        
        # Goal setting
        col1, col2 = st.columns([1, 1])
        
        with col1:
            revenue_goal_2026 = st.number_input(
                "💰 Set 2026 Revenue Goal ($)",
                min_value=0,
                value=20000000,
                step=100000,
                format="%d",
                help="Enter your target revenue for 2026"
            )
        
        with col2:
            if not monthly_forecast.empty:
                current_forecast_total = monthly_forecast['Forecasted_Amount'].sum()
                gap_info = calculate_gap_to_goal(current_forecast_total, revenue_goal_2026)
                
                if gap_info['needs_growth']:
                    st.error(f"⚠️ **Gap to Goal:** ${gap_info['gap']:,.0f} ({gap_info['gap_pct']:.1f}% growth needed)")
                else:
                    st.success(f"✅ **Exceeding Goal by:** ${abs(gap_info['gap']):,.0f}")
        
        st.markdown("---")
        
        # Visual goal comparison
        if not monthly_forecast.empty:
            gap_visual = create_gap_analysis_visual(gap_info)
            if gap_visual:
                st.plotly_chart(gap_visual, use_container_width=True)
        
        st.markdown("---")
        
        # Product focus recommendations
        if gap_info['needs_growth']:
            st.markdown("### 🎯 Recommended Product Focus Areas")
            st.markdown(f"*To close the ${gap_info['gap']:,.0f} gap, consider focusing on these high-growth products:*")
            
            # Use planning_df (with optional outlier removal) for recommendations
            # This does NOT affect the historical baseline forecast
            growth_potential = calculate_product_growth_potential(planning_df)
            
            if growth_potential is not None:
                recommendations = recommend_product_focus(growth_potential, gap_info['gap'], top_n=5)
                
                if recommendations is not None:
                    # Chart
                    rec_chart = create_product_recommendation_chart(recommendations)
                    if rec_chart:
                        st.plotly_chart(rec_chart, use_container_width=True)
                    
                    # Detailed recommendations table
                    st.markdown("#### 📋 Detailed Growth Analysis")
                    
                    rec_display = recommendations[['Product Type', 'Revenue_2025', 'Growth_Rate', 
                                                   'Additional_Revenue', 'Gap_Coverage']].copy()
                    rec_display.columns = ['Product Type', '2025 Revenue', 'YoY Growth %', 
                                          'Potential 2026 Increase', '% of Gap Covered']
                    
                    rec_display['2025 Revenue'] = rec_display['2025 Revenue'].apply(lambda x: f"${x:,.0f}")
                    rec_display['YoY Growth %'] = rec_display['YoY Growth %'].apply(lambda x: f"{x:+.1f}%")
                    rec_display['Potential 2026 Increase'] = rec_display['Potential 2026 Increase'].apply(lambda x: f"${x:,.0f}")
                    rec_display['% of Gap Covered'] = rec_display['% of Gap Covered'].apply(lambda x: f"{x:.1f}%")
                    
                    st.dataframe(rec_display, use_container_width=True, hide_index=True)
                    
                    # Strategic insights
                    st.markdown("#### 💡 Strategic Insights")
                    top_product = recommendations.iloc[0]
                    st.info(f"""
                    **Top Opportunity:** {top_product['Product Type']}
                    - Currently growing at **{top_product['Growth_Rate']:.1f}%** YoY
                    - Could contribute **${top_product['Additional_Revenue']:,.0f}** if growth continues
                    - Would cover **{top_product['Gap_Coverage']:.1f}%** of your revenue gap
                    """)
        
        st.markdown("---")
        
        # Scenario Planning
        st.markdown("### 🎨 Scenario Planning")
        st.markdown("*Adjust product category growth rates to model different scenarios*")
        
        if 'Product Type' in filtered_df.columns:
            # Get ALL product types (not just top 8)
            all_products = filtered_df.groupby('Product Type')['Amount'].sum().sort_values(ascending=False)
            
            # Show product type selection
            st.markdown("#### Product Categories Available")
            
            # Show summary of all product types
            with st.expander("📋 View All Product Types & Revenue", expanded=False):
                product_summary = pd.DataFrame({
                    'Product Type': all_products.index,
                    'Total Revenue': all_products.values
                })
                product_summary['Revenue'] = product_summary['Total Revenue'].apply(lambda x: f"${x:,.0f}")
                product_summary['% of Total'] = (product_summary['Total Revenue'] / product_summary['Total Revenue'].sum() * 100).apply(lambda x: f"{x:.1f}%")
                st.dataframe(product_summary[['Product Type', 'Revenue', '% of Total']], use_container_width=True, hide_index=True)
            
            st.markdown("#### Adjust Growth Rates by Product Category")
            st.caption(f"📊 Showing all {len(all_products)} product types (sorted by revenue)")
            
            product_scenarios = {}
            
            # Create sliders for ALL products (organized in rows of 4)
            product_list = list(all_products.items())
            
            for i in range(0, len(product_list), 4):
                cols = st.columns(4)
                batch = product_list[i:i+4]
                
                for idx, (product, revenue) in enumerate(batch):
                    with cols[idx]:
                        multiplier = st.slider(
                            f"{product[:25]}..." if len(product) > 25 else product,
                            min_value=0.5,
                            max_value=2.0,
                            value=1.0,
                            step=0.1,
                            format="%.1fx",
                            key=f"scenario_{product}",
                            help=f"Historical Revenue: ${revenue:,.0f}"
                        )
                        product_scenarios[product] = multiplier
            
            # Show scenario impact
            if any(v != 1.0 for v in product_scenarios.values()):
                st.markdown("#### 📊 Scenario Impact")
                
                # Calculate weighted impact
                total_revenue = filtered_df['Amount'].sum()
                weighted_impact = 0
                
                for product, multiplier in product_scenarios.items():
                    product_revenue = filtered_df[filtered_df['Product Type'] == product]['Amount'].sum()
                    product_pct = product_revenue / total_revenue if total_revenue > 0 else 0
                    weighted_impact += (multiplier - 1.0) * product_pct
                
                scenario_forecast = current_forecast_total * (1 + weighted_impact)
                scenario_gap = revenue_goal_2026 - scenario_forecast
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(
                        "Scenario Forecast",
                        f"${scenario_forecast:,.0f}",
                        delta=f"${scenario_forecast - current_forecast_total:,.0f}"
                    )
                with col2:
                    st.metric(
                        "vs Goal",
                        f"${scenario_gap:,.0f}" if scenario_gap > 0 else "✅ Goal Met",
                        delta="Gap" if scenario_gap > 0 else "Surplus"
                    )
                with col3:
                    growth_needed = ((scenario_forecast / current_forecast_total) - 1) * 100
                    st.metric(
                        "Total Growth",
                        f"{growth_needed:+.1f}%"
                    )
    
    with tab3:
        st.markdown("### 👥 Sales Rep Territory Planning")
        
        # Check if Sales Rep column exists
        if 'Sales Rep' not in filtered_df.columns:
            st.error("❌ Sales Rep column not found in data")
            st.info("""
            **To use Sales Rep Planning:**
            1. Ensure your invoice data has a 'Sales Rep' column
            2. Column should contain one of: Alex Gonzalez, Lance Mitton, Dave Borkowski, Jake Lynch, Brad Sherman
            3. Re-load the data
            """)
        else:
            try:
                # Rep selector
                selected_rep = st.selectbox(
                    "Select Sales Rep",
                    SALES_REPS,
                    help="Choose a sales rep to view their territory forecast"
                )
                
                if selected_rep:
                    # Get rep's customers
                    rep_customers = get_rep_customers(filtered_df, selected_rep)
                    
                    st.markdown(f"#### 📊 Territory Overview: {selected_rep}")
                    
                    # Churn management
                    st.markdown("##### ⚠️ Churn Risk Management")
                    st.caption("Select customers you expect might churn in 2026:")
                    
                    churned_customers = st.multiselect(
                        "At-Risk Customers",
                        options=rep_customers,
                        default=[],
                        key=f"churn_{selected_rep}",
                        help="These customers will be excluded from the forecast"
                    )
                    
                    # Calculate rep forecast
                    rep_forecast, rep_stats = calculate_rep_forecast(
                        monthly_forecast, 
                        filtered_df, 
                        selected_rep,
                        excluded_customers=churned_customers
                    )
                    
                    if rep_stats:
                        st.markdown("---")
                        
                        # Rep metrics
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric(
                                "Historical Revenue",
                                f"${rep_stats['total_revenue']:,.0f}"
                            )
                        
                        with col2:
                            st.metric(
                                "% of Total",
                                f"{rep_stats['pct_of_total']:.1f}%"
                            )
                        
                        with col3:
                            st.metric(
                                "Active Customers",
                                f"{rep_stats['customer_count']}"
                            )
                        
                        with col4:
                            st.metric(
                                "Avg Order Size",
                                f"${rep_stats['avg_order_size']:,.0f}"
                            )
                        
                        # Show impact of churn
                        if churned_customers:
                            churned_revenue = filtered_df[
                                (filtered_df['Sales Rep'] == selected_rep) & 
                                (filtered_df['Customer'].isin(churned_customers))
                            ]['Amount'].sum()
                            
                            st.warning(f"⚠️ Excluding {len(churned_customers)} at-risk customers (${churned_revenue:,.0f} historical revenue)")
                        
                        st.markdown("---")
                        
                        # Rep forecast chart
                        if rep_forecast is not None and not rep_forecast.empty:
                            st.markdown("##### 📈 Monthly Forecast by Territory")
                            
                            fig = go.Figure()
                            
                            fig.add_trace(go.Bar(
                                x=rep_forecast['MonthShort'],
                                y=rep_forecast['Forecasted_Amount'],
                                marker_color='rgba(99, 102, 241, 0.7)',
                                text=rep_forecast['Forecasted_Amount'].apply(lambda x: f"${x/1000:.0f}K"),
                                textposition='outside'
                            ))
                            
                            fig.update_layout(
                                title=f"2026 Forecast - {selected_rep}",
                                xaxis_title="Month",
                                yaxis_title="Revenue ($)",
                                yaxis_tickformat='$,.0f',
                                plot_bgcolor='rgba(0,0,0,0)',
                                paper_bgcolor='rgba(0,0,0,0)',
                                font=dict(color='white'),
                                height=400
                            )
                            
                            st.plotly_chart(fig, use_container_width=True)
                            
                            # Territory goal setting
                            st.markdown("---")
                            st.markdown("##### 🎯 Territory Goal Planning")
                            
                            rep_forecast_total = rep_forecast['Forecasted_Amount'].sum()
                            
                            rep_goal = st.number_input(
                                f"Set 2026 Goal for {selected_rep} ($)",
                                min_value=0,
                                value=int(rep_forecast_total * 1.15),  # Default to 15% growth
                                step=10000,
                                key=f"goal_{selected_rep}"
                            )
                            
                            rep_gap = rep_goal - rep_forecast_total
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("Current Forecast", f"${rep_forecast_total:,.0f}")
                            with col2:
                                if rep_gap > 0:
                                    st.metric("Gap to Goal", f"${rep_gap:,.0f}", delta=f"{(rep_gap/rep_forecast_total*100):.1f}% growth needed")
                                else:
                                    st.success(f"✅ Exceeding goal by ${abs(rep_gap):,.0f}")
                            
                            # Customer breakdown
                            st.markdown("---")
                            st.markdown("##### 👥 Customer Breakdown")
                            
                            rep_customer_summary = filtered_df[filtered_df['Sales Rep'] == selected_rep].groupby('Customer').agg({
                                'Amount': 'sum',
                                'Date': ['count', 'max']
                            }).reset_index()
                            
                            rep_customer_summary.columns = ['Customer', 'Revenue', 'Orders', 'Last Order']
                            rep_customer_summary = rep_customer_summary.sort_values('Revenue', ascending=False)
                            
                            # Mark churned customers
                            rep_customer_summary['Status'] = rep_customer_summary['Customer'].apply(
                                lambda x: '⚠️ At Risk' if x in churned_customers else '✅ Active'
                            )
                            
                            rep_customer_summary['Revenue'] = rep_customer_summary['Revenue'].apply(lambda x: f"${x:,.0f}")
                            rep_customer_summary['Last Order'] = pd.to_datetime(rep_customer_summary['Last Order']).dt.strftime('%b %Y')
                            
                            st.dataframe(
                                rep_customer_summary[['Customer', 'Status', 'Revenue', 'Orders', 'Last Order']],
                                use_container_width=True,
                                hide_index=True
                            )
            except Exception as e:
                st.error(f"Error loading rep planning: {str(e)}")
                st.info("Please ensure Sales Rep column is properly formatted")
    
    with tab4:
        st.markdown("### 📊 Revenue Breakdown by Product")
        
        breakdown_col = st.radio(
            "Group by:",
            ["Product Type", "Item Type"],
            horizontal=True,
            key="breakdown_col"
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            pie_chart = create_product_breakdown_chart(filtered_df, breakdown_col, 'Amount')
            if pie_chart:
                st.plotly_chart(pie_chart, use_container_width=True)
        
        with col2:
            qty_pie = create_product_breakdown_chart(filtered_df, breakdown_col, 'Quantity')
            if qty_pie:
                st.plotly_chart(qty_pie, use_container_width=True)
        
        # Product-level forecast table
        st.markdown("### 📋 Forecast by Product")
        
        if breakdown_col in filtered_df.columns:
            product_forecasts = []
            for product in filtered_df[breakdown_col].unique():
                product_df = filtered_df[filtered_df[breakdown_col] == product]
                prod_monthly, prod_quarterly, _ = generate_2026_forecast(product_df, weight_2024, weight_2025)
                
                if not prod_monthly.empty:
                    product_forecasts.append({
                        breakdown_col: product,
                        '2026 Revenue': prod_monthly['Forecasted_Amount'].sum(),
                        '2026 Quantity': prod_monthly['Forecasted_Quantity'].sum(),
                        'Q1 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 1]['Forecasted_Amount'].sum(),
                        'Q2 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 2]['Forecasted_Amount'].sum(),
                        'Q3 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 3]['Forecasted_Amount'].sum(),
                        'Q4 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 4]['Forecasted_Amount'].sum()
                    })
            
            if product_forecasts:
                forecast_table = pd.DataFrame(product_forecasts)
                forecast_table = forecast_table.sort_values('2026 Revenue', ascending=False)
                
                # Format columns
                for col in ['2026 Revenue', 'Q1 Revenue', 'Q2 Revenue', 'Q3 Revenue', 'Q4 Revenue']:
                    forecast_table[col] = forecast_table[col].apply(format_currency)
                forecast_table['2026 Quantity'] = forecast_table['2026 Quantity'].apply(format_quantity)
                
                st.dataframe(forecast_table, use_container_width=True, hide_index=True)
    
    with tab5:
        st.markdown("### 🏆 Customer Analysis")
        
        try:
            # Customer Acquisition Analysis
            st.markdown("#### 📊 Customer Acquisition Trends")
            
            monthly_stats, new_customers, customer_cohorts = create_customer_acquisition_analysis(filtered_df)
            
            if monthly_stats is not None:
                # Monthly active customers & revenue chart
                active_customers_chart = create_monthly_active_customers_chart(monthly_stats)
                if active_customers_chart:
                    st.plotly_chart(active_customers_chart, use_container_width=True)
                
                # New customer acquisition chart
                new_customer_chart = create_new_customer_acquisition_chart(new_customers)
                if new_customer_chart:
                    st.plotly_chart(new_customer_chart, use_container_width=True)
                
                # Product acquisition analysis
                product_acquisition = create_product_acquisition_analysis(filtered_df)
                if product_acquisition is not None:
                    product_acq_chart = create_product_acquisition_chart(product_acquisition)
                    if product_acq_chart:
                        st.plotly_chart(product_acq_chart, use_container_width=True)
                
                st.markdown("---")
        
        # Top customers chart (existing)
            st.markdown("#### 💰 Top Customers by Revenue")
            top_customers_chart = create_top_customers_chart(filtered_df, top_n=15)
            if top_customers_chart:
                st.plotly_chart(top_customers_chart, use_container_width=True)
            
            # Customer summary table
            st.markdown("### 📋 Customer Summary")
            
            if 'Customer' in filtered_df.columns:
                customer_summary = filtered_df.groupby('Customer').agg({
                    'Amount': 'sum',
                    'Quantity': 'sum',
                    'Date': ['count', 'min', 'max']
                }).reset_index()
                
                customer_summary.columns = ['Customer', 'Total Revenue', 'Total Quantity', 'Orders', 'First Order', 'Last Order']
                customer_summary = customer_summary.sort_values('Total Revenue', ascending=False)
                
                # Format columns
                customer_summary['Total Revenue'] = customer_summary['Total Revenue'].apply(format_currency)
                customer_summary['Total Quantity'] = customer_summary['Total Quantity'].apply(format_quantity)
                customer_summary['First Order'] = pd.to_datetime(customer_summary['First Order']).dt.strftime('%b %Y')
                customer_summary['Last Order'] = pd.to_datetime(customer_summary['Last Order']).dt.strftime('%b %Y')
                
                st.dataframe(customer_summary.head(50), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Error loading customer analysis: {str(e)}")
            st.info("Please ensure Customer and Date columns exist in data")
    
    with tab6:
        st.markdown("### 📋 Monthly Forecast Details")
        
        if not monthly_forecast.empty:
            display_df = monthly_forecast[[
                'MonthName', 'Forecasted_Amount', 'Amt_Low', 'Amt_High',
                'Forecasted_Quantity', 'Qty_Low', 'Qty_High',
                'Historical_Amt_2024', 'Historical_Amt_2025', 'Confidence'
            ]].copy()
            
            display_df.columns = [
                'Month', 'Forecast Revenue', 'Rev Low', 'Rev High',
                'Forecast Qty', 'Qty Low', 'Qty High',
                '2024 Revenue', '2025 Revenue', 'Confidence'
            ]
            
            # Format currency columns
            for col in ['Forecast Revenue', 'Rev Low', 'Rev High', '2024 Revenue', '2025 Revenue']:
                display_df[col] = display_df[col].apply(format_currency)
            
            # Format quantity columns
            for col in ['Forecast Qty', 'Qty Low', 'Qty High']:
                display_df[col] = display_df[col].apply(format_quantity)
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🔍 Raw Data Sample")
        
        st.dataframe(filtered_df.head(100), use_container_width=True)
        st.caption(f"Showing first 100 of {len(filtered_df):,} records")
    
    with tab7:
        st.markdown(f"""
        ### How This Forecast is Calculated
        
        **Data Source:** Invoice Line Item tab containing historical invoice data with product categorization.
        
        **Weighting:**
        - **2024:** {weight_2024:.0%} weight (healthy stock levels, more representative demand)
        - **2025:** {weight_2025:.0%} weight (stock constraints may have suppressed true demand)
        
        **Outlier Treatment:** Extreme values are softened using winsorization (capped at 5th and 95th percentiles) 
        to prevent single large orders from skewing the forecast.
        
        **Confidence Ranges:**
        - Q1 2026: ±20% (highest confidence - closest to current date)
        - Q2 2026: ±25% (moderate confidence)
        - Q3-Q4 2026: ±30% (lower confidence - further from current date)
        
        **Filters:** Use the sidebar filters to drill down into specific product types, item types, 
        customers, or sales reps for more targeted forecasts.
        
        **Adjustments:** Use the forecast adjustment sliders to apply scenarios like optimistic (+20%) 
        or conservative (-20%) projections, or to model expected growth trends.
        """)
        
        st.markdown("---")
        st.markdown("### 🔍 Product Type Diagnostic")
        st.markdown("*Check consistency across data sources*")
        
        # Product type comparison across data sources
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**📄 Invoice Data**")
            if 'Product Type' in df.columns:
                invoice_products = sorted(df['Product Type'].unique().tolist())
                st.caption(f"{len(invoice_products)} product types")
                with st.expander("View List"):
                    for product in invoice_products:
                        st.caption(f"• {product}")
            else:
                st.warning("No Product Type column")
        
        with col2:
            st.markdown("**📦 Sales Orders**")
            if not sales_order_df.empty and 'Product Type' in sales_order_df.columns:
                so_products = sorted(sales_order_df['Product Type'].unique().tolist())
                st.caption(f"{len(so_products)} product types")
                with st.expander("View List"):
                    for product in so_products:
                        st.caption(f"• {product}")
            else:
                st.caption("No active sales orders")
        
        with col3:
            st.markdown("**🔮 HubSpot Pipeline**")
            if not hubspot_df.empty and 'Product Type' in hubspot_df.columns:
                hs_products = sorted(hubspot_df['Product Type'].unique().tolist())
                st.caption(f"{len(hs_products)} product types")
                with st.expander("View List"):
                    for product in hs_products:
                        st.caption(f"• {product}")
            else:
                st.caption("No HubSpot data loaded")
        
        # Show mismatches
        if 'Product Type' in df.columns:
            invoice_set = set(df['Product Type'].unique())
            
            if not sales_order_df.empty and 'Product Type' in sales_order_df.columns:
                so_set = set(sales_order_df['Product Type'].unique())
                only_in_so = so_set - invoice_set
                if only_in_so:
                    st.warning(f"⚠️ Sales Orders contain {len(only_in_so)} product types not in invoices: {', '.join(sorted(only_in_so))}")
            
            if not hubspot_df.empty and 'Product Type' in hubspot_df.columns:
                hs_set = set(hubspot_df['Product Type'].unique())
                only_in_hs = hs_set - invoice_set
                if only_in_hs:
                    st.warning(f"⚠️ HubSpot contains {len(only_in_hs)} product types not in invoices: {', '.join(sorted(only_in_hs))}")
                    st.info("💡 **Tip:** Product names in HubSpot may need mapping to match Invoice Product Types. Edit lines 590-604 in the code to add custom mappings.")


# Entry point when called from main dashboard
if __name__ == "__main__":
    main()
