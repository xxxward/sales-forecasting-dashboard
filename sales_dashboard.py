"""
Sales Forecasting Dashboard - Enhanced Version
Reads from Google Sheets and displays gap-to-goal analysis with interactive visualizations
Includes lead time logic for Q4/Q1 fulfillment determination
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
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v15"

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """
    Load data from Google Sheets with caching
    """
    try:
        # Load credentials from Streamlit secrets
        creds_dict = st.secrets["gcp_service_account"]
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
            st.warning(f"No data found in {sheet_name}!{range_name}")
            return pd.DataFrame()
        
        # Handle mismatched column counts - pad shorter rows with empty strings
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        # Convert to DataFrame
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"Error loading data from {sheet_name}: {str(e)}")
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
    
    # Load deals data
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:I", version=CACHE_VERSION)
    
    # Load dashboard info (rep quotas and orders)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoice data from NetSuite
    invoices_df = load_google_sheets_data("NS Invoices", "A:Z", version=CACHE_VERSION)
    
    # Load sales orders data from NetSuite
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:Z", version=CACHE_VERSION)
    
    # Clean and process deals data
    if not deals_df.empty and len(deals_df.columns) >= 8:
        # Get column names from first row
        if len(deals_df) > 0:
            # Standardize column names based on position
            col_names = deals_df.columns.tolist()
            
            # Map columns by position
            rename_dict = {
                col_names[1]: 'Deal Name',
                col_names[3]: 'Close Date',
                col_names[4]: 'Deal Owner',
                col_names[5]: 'Amount',
                col_names[6]: 'Status',
                col_names[7]: 'Pipeline'
            }
            
            # Check if we have Product Type column (might be column 8)
            if len(col_names) > 8:
                rename_dict[col_names[8]] = 'Product Type'
            
            deals_df = deals_df.rename(columns=rename_dict)
            
            # Clean and convert amount to numeric
            def clean_numeric(value):
                if pd.isna(value) or value == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            deals_df['Amount'] = deals_df['Amount'].apply(clean_numeric)
            
            # Convert close date to datetime
            deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
            
            # Apply Q4 fulfillment logic
            deals_df = apply_q4_fulfillment_logic(deals_df)
    
    if not dashboard_df.empty:
        # Ensure we have the right column names
        if len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            
            # Remove any empty rows
            dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
            
            # Clean and convert numeric columns
            def clean_numeric(value):
                if pd.isna(value) or value == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            dashboard_df['Quota'] = dashboard_df['Quota'].apply(clean_numeric)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(clean_numeric)
        else:
            st.error(f"Dashboard Info sheet has wrong number of columns: {len(dashboard_df.columns)}")
    
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
                if pd.isna(value) or value == '':
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
    
    # Process sales orders data
    if not sales_orders_df.empty:
        status_col = None
        amount_col = None
        sales_rep_col = None
        date_col = None
        pending_fulfillment_date_col = None
        projected_date_col = None
        customer_promise_col = None
        
        for col in sales_orders_df.columns:
            col_lower = str(col).lower()
            if 'status' in col_lower and not status_col:
                status_col = col
            if ('amount' in col_lower or 'total' in col_lower) and not amount_col:
                amount_col = col
            if ('sales rep' in col_lower or 'salesrep' in col_lower) and not sales_rep_col:
                sales_rep_col = col
            if col_lower == 'date' and not date_col:
                date_col = col
            if 'pending fulfillment date' in col_lower:
                pending_fulfillment_date_col = col
            if 'projected date' in col_lower:
                projected_date_col = col
            if 'customer promise last date to ship' in col_lower:
                customer_promise_col = col
        
        if status_col and amount_col and sales_rep_col:
            rename_dict = {
                status_col: 'Status',
                amount_col: 'Amount',
                sales_rep_col: 'Sales Rep'
            }
            
            if date_col:
                rename_dict[date_col] = 'Date'
            if pending_fulfillment_date_col:
                rename_dict[pending_fulfillment_date_col] = 'Pending Fulfillment Date'
            if projected_date_col:
                rename_dict[projected_date_col] = 'Projected Date'
            if customer_promise_col:
                rename_dict[customer_promise_col] = 'Customer Promise Date'
            
            sales_orders_df = sales_orders_df.rename(columns=rename_dict)
            
            def clean_numeric_so(value):
                if pd.isna(value) or value == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric_so)
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].str.strip()
            sales_orders_df['Status'] = sales_orders_df['Status'].str.strip()
            
            if 'Date' in sales_orders_df.columns:
                sales_orders_df['Date'] = pd.to_datetime(sales_orders_df['Date'], errors='coerce')
            
            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment'])
            ]
            
            if 'Pending Fulfillment Date' in sales_orders_df.columns:
                def get_fulfillment_date(row):
                    if row['Status'] == 'Pending Approval':
                        # For Pending Approval, check Projected Date and Customer Promise Date
                        if 'Projected Date' in row and pd.notna(row['Projected Date']) and row['Projected Date'] != '':
                            return row['Projected Date']
                        elif 'Customer Promise Date' in row and pd.notna(row['Customer Promise Date']) and row['Customer Promise Date'] != '':
                            return row['Customer Promise Date']
                        return None
                    
                    # For Pending Fulfillment, use waterfall J→M→L
                    date_val = None
                    
                    if 'Pending Fulfillment Date' in row and pd.notna(row['Pending Fulfillment Date']) and row['Pending Fulfillment Date'] != '':
                        date_val = row['Pending Fulfillment Date']
                    elif 'Projected Date' in row and pd.notna(row['Projected Date']) and row['Projected Date'] != '':
                        date_val = row['Projected Date']
                    elif 'Customer Promise Date' in row and pd.notna(row['Customer Promise Date']) and row['Customer Promise Date'] != '':
                        date_val = row['Customer Promise Date']
                    
                    return date_val
                
                sales_orders_df['Effective Date'] = sales_orders_df.apply(get_fulfillment_date, axis=1)
                sales_orders_df['Effective Date'] = pd.to_datetime(sales_orders_df['Effective Date'], errors='coerce')
                
                q4_start = pd.Timestamp('2025-10-01')
                q4_end = pd.Timestamp('2025-12-31')
                
                # Keep all pending approval (will filter by date later)
                pending_approval = sales_orders_df[sales_orders_df['Status'] == 'Pending Approval']
                
                pending_fulfillment_q4 = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Fulfillment') &
                    (sales_orders_df['Effective Date'] >= q4_start) &
                    (sales_orders_df['Effective Date'] <= q4_end)
                ]
                
                pending_fulfillment_no_date = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Fulfillment') &
                    (sales_orders_df['Effective Date'].isna())
                ]
                
                sales_orders_df = pd.concat([pending_approval, pending_fulfillment_q4, pending_fulfillment_no_date])
            
            if 'Date' in sales_orders_df.columns:
                today = pd.Timestamp.now()
                sales_orders_df['Age_Days'] = (today - sales_orders_df['Date']).dt.days
            else:
                sales_orders_df['Age_Days'] = 0
            
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & 
                (sales_orders_df['Sales Rep'].notna()) & 
                (sales_orders_df['Sales Rep'] != '')
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
    """Calculate metrics for a specific rep"""
    
    # Get rep's quota and orders
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    
    if rep_info.empty:
        return None
    
    quota = rep_info['Quota'].iloc[0]
    orders = rep_info['NetSuite Orders'].iloc[0]
    
    # Filter deals for this rep
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name]
    
    # Separate Q4 and Q1 deals
    rep_deals_q4 = rep_deals[rep_deals.get('Counts_In_Q4', True) == True]
    rep_deals_q1 = rep_deals[rep_deals.get('Counts_In_Q4', True) == False]
    
    # Calculate Expect/Commit (Q4 only)
    expect_commit = rep_deals_q4[rep_deals_q4['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    
    # Calculate Best Case/Opportunity (Q4 only)
    best_opp = rep_deals_q4[rep_deals_q4['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
    # Track Q1 spillover
    q1_spillover = rep_deals_q1['Amount'].sum()
    
    # Calculate sales order metrics
    pending_approval = 0
    pending_approval_no_date = 0
    pending_approval_old = 0
    pending_fulfillment = 0
    pending_fulfillment_no_date = 0
    
    if sales_orders_df is not None and not sales_orders_df.empty:
        rep_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name]
        
        # Pending Approval - only with dates in Projected Date or Customer Promise Date
        pending_approval_orders = rep_orders[rep_orders['Status'] == 'Pending Approval']
        
        # Count pending approval WITH dates
        if 'Effective Date' in pending_approval_orders.columns:
            pending_approval = pending_approval_orders[
                pending_approval_orders['Effective Date'].notna()
            ]['Amount'].sum()
            
            # Count pending approval WITHOUT dates
            pending_approval_no_date = pending_approval_orders[
                pending_approval_orders['Effective Date'].isna()
            ]['Amount'].sum()
        
        # Old Pending Approval (over 2 weeks)
        if 'Age_Days' in pending_approval_orders.columns:
            pending_approval_old = pending_approval_orders[
                pending_approval_orders['Age_Days'] > 14
            ]['Amount'].sum()
        
        # Pending Fulfillment
        if 'Effective Date' in rep_orders.columns:
            pending_fulfillment = rep_orders[
                (rep_orders['Status'] == 'Pending Fulfillment') &
                (rep_orders['Effective Date'].notna())
            ]['Amount'].sum()
            
            # Pending Fulfillment without dates
            pending_fulfillment_no_date = rep_orders[
                (rep_orders['Status'] == 'Pending Fulfillment') &
                (rep_orders['Effective Date'].isna())
            ]['Amount'].sum()
        else:
            pending_fulfillment = rep_orders[
                rep_orders['Status'] == 'Pending Fulfillment'
            ]['Amount'].sum()
    
    # Total Pending Fulfillment
    total_pending_fulfillment = pending_fulfillment + pending_fulfillment_no_date
    
    # Total Progress calculation
    total_progress = orders + expect_commit + pending_approval + pending_fulfillment
    
    # Calculate gap
    gap = quota - total_progress
    
    # Calculate attainment
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
    
    # Potential attainment
    potential_attainment = ((total_progress + best_opp) / quota * 100) if quota > 0 else 0
    
    return {
        'quota': quota,
        'orders': orders,
        'expect_commit': expect_commit,
        'best_opp': best_opp,
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
        'q1_spillover': q1_spillover,
        'deals': rep_deals_q4
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
        # Try with datetime object
        fig.add_vline(
            x=q4_boundary, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Q4/Q1 Boundary"
        )
    except:
        # If that fails, try without annotation
        try:
            fig.add_shape(
                type="line",
                x0=q4_boundary, x1=q4_boundary,
                y0=0, y1=1,
                yref="paper",
                line=dict(color="red", dash="dash")
            )
            fig.add_annotation(
                x=q4_boundary,
                y=1,
                yref="paper",
                text="Q4/Q1 Boundary",
                showarrow=False,
                yshift=10
            )
        except:
            # If all else fails, just skip the boundary line
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

def create_customer_invoice_table(invoices_df, rep_name):
    """Create a detailed customer invoice breakdown table"""
    
    if invoices_df.empty:
        return pd.DataFrame()
    
    rep_invoices = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
    
    if rep_invoices.empty:
        return pd.DataFrame()
    
    # Group by customer and status
    customer_summary = rep_invoices.groupby(['Customer', 'Status'])['Amount'].sum().reset_index()
    
    # Pivot to show statuses as columns
    pivot_table = customer_summary.pivot_table(
        index='Customer',
        columns='Status',
        values='Amount',
        fill_value=0,
        aggfunc='sum'
    ).reset_index()
    
    # Add total column
    status_cols = [col for col in pivot_table.columns if col != 'Customer']
    pivot_table['Total'] = pivot_table[status_cols].sum(axis=1)
    
    # Sort by total descending
    pivot_table = pivot_table.sort_values('Total', ascending=False)
    
    return pivot_table

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
    
    # Boss's Q4 numbers from the UPDATED screenshot
    boss_rep_numbers = {
        'Jake Lynch': {
            'invoiced': 518981,
            'pending_fulfillment': 291888,
            'pending_approval': 42002,
            'hubspot': 350386,
            'total': 1203256,
            'pending_fulfillment_so_no_date': 108306,
            'pending_approval_so_no_date': 2107,
            'old_pending_approval': 33741,
            'total_q4': 1347410
        },
        'Dave Borkowski': {
            'invoiced': 223593,
            'pending_fulfillment': 146068,
            'pending_approval': 15702,
            'hubspot': 396043,
            'total': 781406,
            'pending_fulfillment_so_no_date': 48150,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 81737,
            'total_q4': 911294
        },
        'Alex Gonzalez': {
            'invoiced': 311101,
            'pending_fulfillment': 190589,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 501691,
            'pending_fulfillment_so_no_date': 3183,
            'pending_approval_so_no_date': 34846,
            'old_pending_approval': 19300,
            'total_q4': 559019
        },
        'Brad Sherman': {
            'invoiced': 107166,
            'pending_fulfillment': 39759,
            'pending_approval': 16878,
            'hubspot': 211062,
            'total': 374865,
            'pending_fulfillment_so_no_date': 35390,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1006,
            'total_q4': 411262
        },
        'Lance Mitton': {
            'invoiced': 21998,
            'pending_fulfillment': 0,
            'pending_approval': 2758,
            'hubspot': 11000,
            'total': 35756,
            'pending_fulfillment_so_no_date': 3735,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 60527,
            'total_q4': 100019
        },
        'House': {
            'invoiced': 0,
            'pending_fulfillment': 0,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 0,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 0,
            'total_q4': 0
        },
        'Shopify ECommerce': {
            'invoiced': 20404,
            'pending_fulfillment': 1406,
            'pending_approval': 1174,
            'hubspot': 0,
            'total': 22984,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1544,
            'total_q4': 24528
        }
    }
    
    # Boss's pipeline numbers - UPDATED
    boss_pipeline_numbers = {
        'Retention (Existing Product)': {
            'invoiced': 425514,
            'pending_fulfillment': 387936,
            'pending_approval': 46383,
            'hubspot': 570164,
            'total': 1429997,
            'pending_fulfillment_so_no_date': 105199,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 114779,
            'total_q4': 1649975
        },
        'Growth Pipeline (Upsell/Cross-sell)': {
            'invoiced': 219590,
            'pending_fulfillment': 31655,
            'pending_approval': 11321,
            'hubspot': 176265,
            'total': 438830,
            'pending_fulfillment_so_no_date': 46582,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1583,
            'total_q4': 486994
        },
        'Acquisition (New Customer)': {
            'invoiced': 108041,
            'pending_fulfillment': 10066,
            'pending_approval': 25295,
            'hubspot': 222062,
            'total': 365463,
            'pending_fulfillment_so_no_date': 7447,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 60527,
            'total_q4': 433437
        },
        'SO Manually Built': {
            'invoiced': 429222,
            'pending_fulfillment': 301858,
            'pending_approval': 1174,
            'hubspot': 0,
            'total': 732254,
            'pending_fulfillment_so_no_date': 39537,
            'pending_approval_so_no_date': 36953,
            'old_pending_approval': 20967,
            'total_q4': 829711
        },
        'Ecommerce Pipeline': {
            'invoiced': 25388,
            'pending_fulfillment': 1406,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 26793,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 0,
            'total_q4': 26793
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
        st.plotly_chart(status_chart, use_container_width=True)
    
    # Pipeline breakdown
    st.markdown("### 🔄 Pipeline Analysis")
    pipeline_chart = create_pipeline_breakdown_chart(deals_df)
    st.plotly_chart(pipeline_chart, use_container_width=True)
    
    # Timeline
    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    
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
    
    rep_summary_df = pd.DataFrame(rep_summary)
    st.dataframe(rep_summary_df, use_container_width=True, hide_index=True)

def display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display individual rep dashboard"""
    
    st.title(f"👤 {rep_name}'s Q4 2025 Forecast")
    
    # Calculate metrics
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
    
    if not metrics:
        st.error(f"No data found for {rep_name}")
        return
    
    # Display Q1 spillover info if applicable
    if metrics.get('q1_spillover', 0) > 0:
        st.info(
            f"💡 **Q1 2026 Spillover**: ${metrics['q1_spillover']:,.0f} in deals will close in late December "
            f"but ship in Q1 2026 based on product lead times"
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
            help="Orders shipped and invoiced"
        )
    
    with col3:
        st.metric(
            label="Pending Fulfillment",
            value=f"${metrics['pending_fulfillment']/1000:.0f}K" if metrics['pending_fulfillment'] < 1000000 else f"${metrics['pending_fulfillment']/1000000:.1f}M",
            help="Orders with Q4 ship dates"
        )
    
    with col4:
        st.metric(
            label="Pending Approval",
            value=f"${metrics['pending_approval']/1000:.0f}K" if metrics['pending_approval'] < 1000000 else f"${metrics['pending_approval']/1000000:.1f}M",
            help="Orders with dates in Projected/Customer Promise fields"
        )
    
    with col5:
        st.metric(
            label="HubSpot (Q4)",
            value=f"${metrics['expect_commit']/1000:.0f}K" if metrics['expect_commit'] < 1000000 else f"${metrics['expect_commit']/1000000:.1f}M",
            help="Expect/Commit deals shipping in Q4"
        )
    
    with col6:
        st.metric(
            label="Gap to Goal",
            value=f"${metrics['gap']/1000:.0f}K" if abs(metrics['gap']) < 1000000 else f"${metrics['gap']/1000000:.1f}M",
            delta=f"${-metrics['gap']/1000:.0f}K" if metrics['gap'] < 0 else None,
            delta_color="inverse"
        )
    
    # Beautiful progress breakdown
    display_progress_breakdown(metrics)
    
    # Section 2: Additional Orders
    st.markdown("### 📊 Section 2: Additional Orders (Can be included)")
    warning_col1, warning_col2, warning_col3 = st.columns(3)
    
    with warning_col1:
        st.metric(
            label="PF SO's No Date",
            value=f"${metrics['pending_fulfillment_no_date']:,.0f}",
            help="Pending Fulfillment without ship dates"
        )
    
    with warning_col2:
        st.metric(
            label="PA SO's No Date",
            value=f"${metrics['pending_approval_no_date']:,.0f}",
            help="Pending Approval without dates"
        )
    
    with warning_col3:
        st.metric(
            label="Old PA (>2 weeks)",
            value=f"${metrics['pending_approval_old']:,.0f}",
            help="Pending Approval older than 14 days",
            delta="Needs attention" if metrics['pending_approval_old'] > 0 else None,
            delta_color="off" if metrics['pending_approval_old'] > 0 else "normal"
        )
    
    # Final Total
    final_total = (metrics['total_progress'] + metrics['pending_fulfillment_no_date'] + 
                   metrics['pending_approval_no_date'] + metrics['pending_approval_old'])
    st.metric(
        label="📊 FINAL TOTAL Q4",
        value=f"${final_total:,.0f}",
        delta=f"Including all additional orders"
    )
    
    # Charts
    col1, col2 = st.columns(2)
    
    with col1:
        gap_chart = create_gap_chart(metrics, f"{rep_name}'s Progress to Goal")
        st.plotly_chart(gap_chart, use_container_width=True)
    
    with col2:
        status_chart = create_status_breakdown_chart(deals_df, rep_name)
        st.plotly_chart(status_chart, use_container_width=True)
    
    # Pipeline breakdown
    st.markdown("### 🔄 Pipeline Analysis")
    pipeline_chart = create_pipeline_breakdown_chart(deals_df, rep_name)
    st.plotly_chart(pipeline_chart, use_container_width=True)
    
    # Timeline
    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df, rep_name)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    
    # Detailed deals table
    st.markdown("### 📋 Deal Details (Q4 Only)")
    
    # Add filters
    col1, col2 = st.columns(2)
    with col1:
        status_filter = st.multiselect(
            "Filter by Status",
            options=['Expect', 'Commit', 'Best Case', 'Opportunity'],
            default=['Expect', 'Commit', 'Best Case', 'Opportunity']
        )
    
    with col2:
        if 'Pipeline' in metrics['deals'].columns:
            pipeline_filter = st.multiselect(
                "Filter by Pipeline",
                options=metrics['deals']['Pipeline'].unique(),
                default=metrics['deals']['Pipeline'].unique()
            )
        else:
            pipeline_filter = None
    
    # Filter deals
    filtered_deals = metrics['deals'][metrics['deals']['Status'].isin(status_filter)]
    if pipeline_filter:
        filtered_deals = filtered_deals[filtered_deals['Pipeline'].isin(pipeline_filter)]
    
    # Display deals table
    if not filtered_deals.empty:
        display_deals = filtered_deals[['Deal Name', 'Close Date', 'Amount', 'Status', 'Pipeline']].copy()
        display_deals['Amount'] = display_deals['Amount'].apply(lambda x: f"${x:,.0f}")
        display_deals['Close Date'] = display_deals['Close Date'].dt.strftime('%Y-%m-%d')
        st.dataframe(display_deals, use_container_width=True, hide_index=True)
    else:
        st.info("No deals match the selected filters.")

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
    
    if deals_df.empty or dashboard_df.empty:
        st.error("Unable to load data. Please check your Google Sheets connection.")
        st.info("""
        **Setup Instructions:**
        1. Add your Google Service Account credentials to Streamlit secrets
        2. Share your Google Sheet with the service account email
        3. Verify the spreadsheet ID in the code
        4. Ensure 'Product Type' column exists in HubSpot data for lead time calculations
        """)
        return
    
    # Display appropriate dashboard
    if view_mode == "Team Overview":
        display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df)
    elif view_mode == "Individual Rep":
        rep_name = st.selectbox(
            "Select Rep:",
            options=dashboard_df['Rep Name'].tolist()
        )
        display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df)
    else:  # Reconciliation view
        display_reconciliation_view(deals_df, dashboard_df, sales_orders_df)

if __name__ == "__main__":
    main()
