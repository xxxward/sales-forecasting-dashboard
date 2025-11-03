"""
Sales Forecasting Dashboard
Reads from Google Sheets and displays gap-to-goal analysis with interactive visualizations
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
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v10"

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
        
        # Debug output
        st.sidebar.write(f"**DEBUG - {sheet_name}:**")
        st.sidebar.write(f"Total rows loaded: {len(values)}")
        st.sidebar.write(f"First 3 rows: {values[:3]}")
        
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
        st.sidebar.write(f"**ERROR in {sheet_name}:** {str(e)}")
        return pd.DataFrame()

def load_all_data():
    """Load all necessary data from Google Sheets"""
    
    # Load deals data
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:H", version=CACHE_VERSION)
    
    # Load dashboard info (rep quotas and orders)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoice data from NetSuite
    invoices_df = load_google_sheets_data("NS Invoices", "A:Z", version=CACHE_VERSION)
    
    # Load sales orders data from NetSuite
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:Z", version=CACHE_VERSION)
    
    # Clean and process data
    if not deals_df.empty and len(deals_df.columns) >= 8:
        # Get column names from first row
        if len(deals_df) > 0:
            # Standardize column names based on position
            # Columns: 0=Record ID, 1=Deal Name, 2=Deal Stage, 3=Close Date, 4=Deal Owner, 5=Amount, 6=Status, 7=Pipeline
            col_names = deals_df.columns.tolist()
            
            # Map columns by position
            deals_df = deals_df.rename(columns={
                col_names[1]: 'Deal Name',
                col_names[3]: 'Close Date',
                col_names[4]: 'Deal Owner',
                col_names[5]: 'Amount',
                col_names[6]: 'Status',
                col_names[7]: 'Pipeline'
            })
            
            # Clean and convert amount to numeric (remove commas, dollar signs)
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
    
    if not dashboard_df.empty:
        # Debug: Show raw data
        st.sidebar.write("**DEBUG - Dashboard Info Raw:**")
        st.sidebar.dataframe(dashboard_df.head())
        
        # The first row should be headers, data starts from row 2
        # Columns should be: Rep Name, Quota, Orders
        
        # Ensure we have the right column names
        if len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            
            # Remove any empty rows
            dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
            
            # Clean and convert numeric columns (remove commas, dollar signs, etc.)
            def clean_numeric(value):
                if pd.isna(value) or value == '':
                    return 0
                # Remove commas, dollar signs, and spaces
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            dashboard_df['Quota'] = dashboard_df['Quota'].apply(clean_numeric)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(clean_numeric)
            
            # Debug: Show after conversion
            st.sidebar.write("**DEBUG - After Conversion:**")
            st.sidebar.dataframe(dashboard_df)
        else:
            st.error(f"Dashboard Info sheet has wrong number of columns: {len(dashboard_df.columns)}")
    
    # Process invoice data
    if not invoices_df.empty:
        st.sidebar.write("**DEBUG - NS Invoices loaded:**", len(invoices_df), "rows")
        
        # Clean invoice data
        if len(invoices_df.columns) >= 15:
            # Map important columns by position (0-indexed)
            # Columns: 0=Doc#, 1=Status, 2=Date, 6=Customer, 10=Amount, 14=Sales Rep
            invoices_df = invoices_df.rename(columns={
                invoices_df.columns[0]: 'Invoice Number',
                invoices_df.columns[1]: 'Status',
                invoices_df.columns[2]: 'Date',
                invoices_df.columns[6]: 'Customer',
                invoices_df.columns[10]: 'Amount',
                invoices_df.columns[14]: 'Sales Rep'
            })
            
            # Clean numeric amount
            def clean_numeric(value):
                if pd.isna(value) or value == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            invoices_df['Amount'] = invoices_df['Amount'].apply(clean_numeric)
            
            # Convert Date column to datetime
            invoices_df['Date'] = pd.to_datetime(invoices_df['Date'], errors='coerce')
            
            # Filter to Q4 2025 only (October 1 - December 31, 2025)
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            invoices_df = invoices_df[
                (invoices_df['Date'] >= q4_start) & 
                (invoices_df['Date'] <= q4_end)
            ]
            
            st.sidebar.write(f"**DEBUG - After Q4 2025 filter:** {len(invoices_df)} rows")
            
            # Normalize rep names (trim spaces, consistent capitalization)
            invoices_df['Sales Rep'] = invoices_df['Sales Rep'].str.strip()
            
            # Remove rows without amount or sales rep
            invoices_df = invoices_df[
                (invoices_df['Amount'] > 0) & 
                (invoices_df['Sales Rep'].notna()) & 
                (invoices_df['Sales Rep'] != '')
            ]
            
            st.sidebar.write(f"**DEBUG - After cleaning:** {len(invoices_df)} rows")
            st.sidebar.write("**DEBUG - Unique sales reps in invoices:**", invoices_df['Sales Rep'].unique().tolist())
            
            # Calculate total invoices by rep
            invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
            invoice_totals.columns = ['Rep Name', 'Invoice Total']
            
            st.sidebar.write("**DEBUG - Invoice totals by rep:**")
            st.sidebar.dataframe(invoice_totals)
            
            # Normalize dashboard rep names too
            dashboard_df['Rep Name'] = dashboard_df['Rep Name'].str.strip()
            
            st.sidebar.write("**DEBUG - Rep names in Dashboard Info:**", dashboard_df['Rep Name'].tolist())
            
            # Merge with dashboard_df to update NetSuite Orders
            dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
            dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
            
            # Override NetSuite Orders with calculated invoice totals
            dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
            dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
            
            st.sidebar.write("**DEBUG - Final dashboard_df with invoice totals:**")
            st.sidebar.dataframe(dashboard_df)
        else:
            st.warning(f"NS Invoices sheet doesn't have enough columns (has {len(invoices_df.columns)})")
            invoices_df = pd.DataFrame()
    
    # Process sales orders data
    if not sales_orders_df.empty:
        st.sidebar.write("**DEBUG - NS Sales Orders loaded:**", len(sales_orders_df), "rows")
        
        # Find required columns
        status_col = None
        amount_col = None
        sales_rep_col = None
        pending_fulfillment_date_col = None  # Column J
        projected_date_col = None  # Column M
        customer_promise_col = None  # Column L
        
        for col in sales_orders_df.columns:
            col_lower = str(col).lower()
            if 'status' in col_lower and not status_col:
                status_col = col
            if ('amount' in col_lower or 'total' in col_lower) and not amount_col:
                amount_col = col
            if ('sales rep' in col_lower or 'salesrep' in col_lower) and not sales_rep_col:
                sales_rep_col = col
            if 'pending fulfillment date' in col_lower:
                pending_fulfillment_date_col = col
            if 'projected date' in col_lower:
                projected_date_col = col
            if 'customer promise last date to ship' in col_lower:
                customer_promise_col = col
        
        st.sidebar.write(f"**DEBUG - Found columns:**")
        st.sidebar.write(f"Status={status_col}, Amount={amount_col}, Rep={sales_rep_col}")
        st.sidebar.write(f"Dates: J={pending_fulfillment_date_col}, M={projected_date_col}, L={customer_promise_col}")
        
        if status_col and amount_col and sales_rep_col:
            # Standardize column names
            rename_dict = {
                status_col: 'Status',
                amount_col: 'Amount',
                sales_rep_col: 'Sales Rep'
            }
            
            if pending_fulfillment_date_col:
                rename_dict[pending_fulfillment_date_col] = 'Pending Fulfillment Date'
            if projected_date_col:
                rename_dict[projected_date_col] = 'Projected Date'
            if customer_promise_col:
                rename_dict[customer_promise_col] = 'Customer Promise Date'
            
            sales_orders_df = sales_orders_df.rename(columns=rename_dict)
            
            # Clean numeric amount
            def clean_numeric_so(value):
                if pd.isna(value) or value == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric_so)
            
            # Normalize rep names and status
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].str.strip()
            sales_orders_df['Status'] = sales_orders_df['Status'].str.strip()
            
            # Filter to ONLY Pending Approval and Pending Fulfillment
            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment'])
            ]
            
            st.sidebar.write(f"**DEBUG - After status filter:** {len(sales_orders_df)} rows")
            
            # For Pending Fulfillment orders, determine the date using waterfall logic
            if 'Pending Fulfillment Date' in sales_orders_df.columns:
                # Create a unified date column using waterfall logic
                def get_fulfillment_date(row):
                    if row['Status'] == 'Pending Approval':
                        return None  # No date for Pending Approval
                    
                    # Waterfall: J → M → L → None
                    date_val = None
                    
                    if 'Pending Fulfillment Date' in row and pd.notna(row['Pending Fulfillment Date']) and row['Pending Fulfillment Date'] != '':
                        date_val = row['Pending Fulfillment Date']
                    elif 'Projected Date' in row and pd.notna(row['Projected Date']) and row['Projected Date'] != '':
                        date_val = row['Projected Date']
                    elif 'Customer Promise Date' in row and pd.notna(row['Customer Promise Date']) and row['Customer Promise Date'] != '':
                        date_val = row['Customer Promise Date']
                    
                    return date_val
                
                sales_orders_df['Effective Date'] = sales_orders_df.apply(get_fulfillment_date, axis=1)
                
                # Convert to datetime
                sales_orders_df['Effective Date'] = pd.to_datetime(sales_orders_df['Effective Date'], errors='coerce')
                
                # Filter Pending Fulfillment to Q4 2025 only
                q4_start = pd.Timestamp('2025-10-01')
                q4_end = pd.Timestamp('2025-12-31')
                
                # Separate into categories
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
                
                # Combine back
                sales_orders_df = pd.concat([pending_approval, pending_fulfillment_q4, pending_fulfillment_no_date])
                
                st.sidebar.write(f"**DEBUG - After Q4 filter:**")
                st.sidebar.write(f"Pending Approval: {len(pending_approval)} rows")
                st.sidebar.write(f"Pending Fulfillment Q4: {len(pending_fulfillment_q4)} rows")
                st.sidebar.write(f"Pending Fulfillment No Date: {len(pending_fulfillment_no_date)} rows")
            
            # Remove rows without amount or sales rep
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & 
                (sales_orders_df['Sales Rep'].notna()) & 
                (sales_orders_df['Sales Rep'] != '')
            ]
            
            st.sidebar.write(f"**DEBUG - Final count:** {len(sales_orders_df)} rows")
        else:
            st.warning("Could not find required columns in NS Sales Orders")
            sales_orders_df = pd.DataFrame()
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def calculate_team_metrics(deals_df, dashboard_df):
    """Calculate overall team metrics"""
    
    total_quota = dashboard_df['Quota'].sum()
    total_orders = dashboard_df['NetSuite Orders'].sum()
    
    # Calculate Expect/Commit forecast
    expect_commit = deals_df[deals_df['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    
    # Calculate Best Case/Opportunity
    best_opp = deals_df[deals_df['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
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
        'current_forecast': current_forecast
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
    
    # Calculate Expect/Commit
    expect_commit = rep_deals[rep_deals['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    
    # Calculate Best Case/Opportunity
    best_opp = rep_deals[rep_deals['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
    # Calculate sales order metrics
    pending_approval = 0
    pending_fulfillment = 0
    pending_fulfillment_no_date = 0
    
    if sales_orders_df is not None and not sales_orders_df.empty:
        rep_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name]
        
        # Pending Approval (all amounts, no date filter)
        pending_approval = rep_orders[
            rep_orders['Status'] == 'Pending Approval'
        ]['Amount'].sum()
        
        # Pending Fulfillment with Q4 dates
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
            # If no date column, all pending fulfillment counts
            pending_fulfillment = rep_orders[
                rep_orders['Status'] == 'Pending Fulfillment'
            ]['Amount'].sum()
    
    # NEW CALCULATION: Total Progress includes Orders + Expect/Commit + Pending Approval + Pending Fulfillment
    total_progress = orders + expect_commit + pending_approval + pending_fulfillment
    
    # Calculate gap based on new formula
    gap = quota - total_progress
    
    # Calculate attainment based on total progress
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
    
    # Potential attainment (add Best Case/Opportunity upside)
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
        'pending_fulfillment': pending_fulfillment,
        'pending_fulfillment_no_date': pending_fulfillment_no_date,
        'deals': rep_deals
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
        title='Deal Amount by Forecast Category',
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
        title='Pipeline Breakdown by Forecast Category',
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
        hover_data=['Deal Name', 'Amount', 'Pipeline'],
        title='Deal Close Date Timeline',
        color_discrete_map=color_map
    )
    
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

def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display the team-level dashboard"""
    
    st.title("🎯 Team Sales Dashboard - Q4 2025")
    
    # Calculate metrics
    metrics = calculate_team_metrics(deals_df, dashboard_df)
    
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
            label="Orders Shipped NetSuite",
            value=f"${metrics['current_forecast']:,.0f}",
            delta=f"{metrics['attainment_pct']:.1f}% of quota"
        )
    
    with col3:
        st.metric(
            label="Gap to Goal",
            value=f"${metrics['gap']:,.0f}",
            delta=f"{-metrics['gap']:,.0f}" if metrics['gap'] < 0 else None,
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
                'Orders Shipped': f"${rep_metrics['current_forecast']:,.0f}",
                'Pending Approval': f"${rep_metrics['pending_approval']:,.0f}",
                'Pending Fulfillment': f"${rep_metrics['pending_fulfillment']:,.0f}",
                'Gap': f"${rep_metrics['gap']:,.0f}",
                'Attainment': f"{rep_metrics['attainment_pct']:.1f}%"
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
    
    # Display key metrics - NEW STRUCTURE
    st.markdown("### 💰 Revenue Progress")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            label="Quota",
            value=f"${metrics['quota']:,.0f}"
        )
    
    with col2:
        st.metric(
            label="Orders Shipped",
            value=f"${metrics['orders']:,.0f}",
            help="Invoiced and shipped orders from NetSuite"
        )
    
    with col3:
        st.metric(
            label="Expect/Commit",
            value=f"${metrics['expect_commit']:,.0f}",
            help="HubSpot deals likely to close this quarter"
        )
    
    with col4:
        st.metric(
            label="Pending Approval",
            value=f"${metrics['pending_approval']:,.0f}",
            help="Sales orders awaiting approval"
        )
    
    with col5:
        st.metric(
            label="Pending Fulfillment",
            value=f"${metrics['pending_fulfillment']:,.0f}",
            help="Sales orders awaiting shipment (Q4 only)"
        )
    
    with col6:
        st.metric(
            label="Gap to Goal",
            value=f"${metrics['gap']:,.0f}",
            delta=f"{-metrics['gap']:,.0f}" if metrics['gap'] < 0 else None,
            delta_color="inverse"
        )
    
    # Progress bar
    st.markdown("### 📈 Progress to Quota")
    progress = min(metrics['attainment_pct'] / 100, 1.0)
    st.progress(progress)
    st.caption(
        f"Total Progress: ${metrics['total_progress']:,.0f} = "
        f"Orders (${metrics['orders']:,.0f}) + "
        f"Expect/Commit (${metrics['expect_commit']:,.0f}) + "
        f"Pending Approval (${metrics['pending_approval']:,.0f}) + "
        f"Pending Fulfillment (${metrics['pending_fulfillment']:,.0f})"
    )
    st.caption(f"Attainment: {metrics['attainment_pct']:.1f}% | Potential with Best Case/Opp: {metrics['potential_attainment']:.1f}%")
    
    # Pending Fulfillment No Date warning
    if metrics['pending_fulfillment_no_date'] > 0:
        st.warning(
            f"⚠️ **Pending Fulfillment - No Ship Date:** ${metrics['pending_fulfillment_no_date']:,.0f} "
            f"(Not included in totals - needs ship date to count toward Q4)"
        )
    
    
    # Sales Order Pipeline Metrics
    if metrics['pending_approval'] > 0 or metrics['pending_fulfillment'] > 0:
        st.markdown("### 📦 Sales Order Pipeline")
        so_col1, so_col2 = st.columns(2)
        
        with so_col1:
            st.metric(
                label="⏳ Pending Approval",
                value=f"${metrics['pending_approval']:,.0f}",
                help="Sales orders awaiting approval"
            )
        
        with so_col2:
            st.metric(
                label="📤 Pending Fulfillment",
                value=f"${metrics['pending_fulfillment']:,.0f}",
                help="Sales orders pending fulfillment or billing"
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
    
    # Invoice Status Breakdown
    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Breakdown")
        
        col1, col2 = st.columns(2)
        
        with col1:
            invoice_chart = create_invoice_status_chart(invoices_df, rep_name)
            if invoice_chart:
                st.plotly_chart(invoice_chart, use_container_width=True)
        
        with col2:
            # Customer invoice breakdown table
            customer_table = create_customer_invoice_table(invoices_df, rep_name)
            if not customer_table.empty:
                st.markdown("**Invoice Amounts by Customer**")
                
                # Format currency columns
                for col in customer_table.columns:
                    if col != 'Customer':
                        customer_table[col] = customer_table[col].apply(lambda x: f"${x:,.0f}")
                
                st.dataframe(customer_table, use_container_width=True, hide_index=True)
            else:
                st.info("No invoice data available for this rep")
    
    # Detailed deals table
    st.markdown("### 📋 Deal Details")
    
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
        st.image("https://via.placeholder.com/200x80/1E88E5/FFFFFF?text=Your+Logo", use_container_width=True)
        st.markdown("---")
        
        st.markdown("### 🎯 Dashboard Navigation")
        view_mode = st.radio(
            "Select View:",
            ["Team Overview", "Individual Rep"],
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
        """)
        return
    
    # Display appropriate dashboard
    if view_mode == "Team Overview":
        display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df)
    else:
        rep_name = st.selectbox(
            "Select Rep:",
            options=dashboard_df['Rep Name'].tolist()
        )
        display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df)

if __name__ == "__main__":
    main()
