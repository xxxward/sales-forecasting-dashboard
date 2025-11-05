"""
Sales Forecasting Dashboard - Refined Version
Streamlit app for Q4 2025 sales forecasting with gap-to-goal analysis
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import numpy as np

# ==================== CONFIGURATION ====================
st.set_page_config(
    page_title="Sales Forecasting Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constants
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "v30"

# Color scheme
COLORS = {
    'primary': '#1E88E5',
    'success': '#43A047',
    'warning': '#FB8C00',
    'danger': '#DC3912',
    'expect': '#1E88E5',
    'commit': '#43A047',
    'best_case': '#FB8C00',
    'opportunity': '#8E24AA'
}

# Lead time mapping
LEAD_TIME_MAP = {
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

# ==================== STYLES ====================
def load_custom_css():
    """Load custom CSS styles"""
    st.markdown("""
        <style>
        .metric-card {
            background-color: #f0f2f6;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
        }
        .progress-breakdown {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px;
            border-radius: 15px;
            color: white;
            margin: 20px 0;
            box-shadow: 0 10px 20px rgba(0,0,0,0.2);
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
        .section-header {
            background: #f0f2f6;
            padding: 10px 15px;
            border-radius: 8px;
            margin: 15px 0;
            font-weight: bold;
        }
        .info-card {
            background: #e3f2fd;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #1E88E5;
            margin: 10px 0;
        }
        </style>
        """, unsafe_allow_html=True)

# ==================== DATA LOADING ====================
@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets with caching"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing Google Cloud credentials")
            return pd.DataFrame()
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame()
        
        # Pad shorter rows
        max_cols = max(len(row) for row in values) if len(values) > 1 else len(values[0])
        for row in values:
            while len(row) < max_cols:
                row.append('')
        
        return pd.DataFrame(values[1:], columns=values[0])
        
    except Exception as e:
        st.error(f"Error loading {sheet_name}: {str(e)}")
        return pd.DataFrame()

def clean_numeric(value):
    """Clean and convert string to numeric value"""
    if pd.isna(value) or str(value).strip() == '':
        return 0
    cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except:
        return 0

def get_business_days_before(end_date, business_days):
    """Calculate date that is N business days before end_date"""
    current = end_date
    days_counted = 0
    
    while days_counted < business_days:
        current -= timedelta(days=1)
        if current.weekday() < 5:  # Monday=0, Sunday=6
            days_counted += 1
    
    return current

def business_days_between(start_date, end_date):
    """Calculate business days between two dates"""
    if pd.isna(start_date):
        return 0
    days = pd.bdate_range(start=start_date, end=end_date).size - 1
    return max(0, days)

# ==================== DATA PROCESSING ====================
def process_deals_data(deals_df):
    """Process and clean deals data from HubSpot"""
    if deals_df.empty:
        return deals_df
    
    # Rename columns based on position/content
    col_renames = {
        'Record ID': 'Record ID',
        'Deal Name': 'Deal Name',
        'Deal Stage': 'Deal Stage',
        'Close Date': 'Close Date',
        'Amount': 'Amount',
        'Close Status': 'Status',
        'Pipeline': 'Pipeline',
        'Deal Type': 'Product Type',
        'Q1 2026 Spillover': 'Q1 2026 Spillover'
    }
    
    # Handle combined Deal Owner column
    for col in deals_df.columns:
        if 'Deal Owner First Name' in col and 'Deal Owner Last Name' in col:
            col_renames[col] = 'Deal Owner'
        elif col in col_renames:
            continue
        else:
            for key in col_renames:
                if key.lower() in col.lower():
                    col_renames[col] = col_renames[key]
                    break
    
    deals_df = deals_df.rename(columns=col_renames)
    
    # Clean Deal Owner
    if 'Deal Owner' in deals_df.columns:
        deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
    
    # Convert amount to numeric
    if 'Amount' in deals_df.columns:
        deals_df['Amount'] = deals_df['Amount'].apply(clean_numeric)
    
    # Convert close date
    if 'Close Date' in deals_df.columns:
        deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
    
    # Filter to Q4 2025
    q4_start = pd.Timestamp('2025-10-01')
    q4_end = pd.Timestamp('2025-12-31')
    deals_df = deals_df[
        (deals_df['Close Date'] >= q4_start) & 
        (deals_df['Close Date'] <= q4_end)
    ]
    
    # Filter out unwanted stages
    excluded_stages = ['', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 
                      'closed lost', 'closed won', 'sales order created in NS', 
                      'NCR', 'Shipped']
    
    if 'Deal Stage' in deals_df.columns:
        deals_df['Deal Stage'] = deals_df['Deal Stage'].fillna('').astype(str).str.strip()
        deals_df = deals_df[~deals_df['Deal Stage'].str.lower().isin(
            [s.lower() if s else '' for s in excluded_stages]
        )]
    
    # Apply Q1 spillover logic
    if 'Q1 2026 Spillover' in deals_df.columns:
        deals_df['Ships_In_Q4'] = deals_df['Q1 2026 Spillover'] != 'Q1 2026'
        deals_df['Ships_In_Q1'] = deals_df['Q1 2026 Spillover'] == 'Q1 2026'
    else:
        deals_df['Ships_In_Q4'] = True
        deals_df['Ships_In_Q1'] = False
    
    return deals_df

def process_sales_orders(sales_orders_df):
    """Process sales orders data from NetSuite"""
    if sales_orders_df.empty:
        return sales_orders_df
    
    # Map columns
    col_mapping = {
        8: 'Order Start Date',      # Column I
        11: 'Customer Promise Date', # Column L
        12: 'Projected Date',        # Column M
        27: 'Pending Approval Date'  # Column AB
    }
    
    for idx, name in col_mapping.items():
        if len(sales_orders_df.columns) > idx:
            sales_orders_df.rename(columns={sales_orders_df.columns[idx]: name}, inplace=True)
    
    # Standard column mapping
    for col in sales_orders_df.columns:
        col_lower = str(col).lower()
        if 'status' in col_lower and 'Status' not in sales_orders_df.columns:
            sales_orders_df.rename(columns={col: 'Status'}, inplace=True)
        elif 'amount' in col_lower and 'Amount' not in sales_orders_df.columns:
            sales_orders_df.rename(columns={col: 'Amount'}, inplace=True)
        elif 'sales rep' in col_lower and 'Sales Rep' not in sales_orders_df.columns:
            sales_orders_df.rename(columns={col: 'Sales Rep'}, inplace=True)
        elif 'customer' in col_lower and 'Customer' not in sales_orders_df.columns:
            sales_orders_df.rename(columns={col: 'Customer'}, inplace=True)
        elif 'document' in col_lower and 'Document Number' not in sales_orders_df.columns:
            sales_orders_df.rename(columns={col: 'Document Number'}, inplace=True)
    
    # Remove duplicate columns
    if sales_orders_df.columns.duplicated().any():
        sales_orders_df = sales_orders_df.loc[:, ~sales_orders_df.columns.duplicated()]
    
    # Clean data
    if 'Amount' in sales_orders_df.columns:
        sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric)
    
    if 'Sales Rep' in sales_orders_df.columns:
        sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].astype(str).str.strip()
    
    if 'Status' in sales_orders_df.columns:
        sales_orders_df['Status'] = sales_orders_df['Status'].astype(str).str.strip()
        sales_orders_df = sales_orders_df[
            sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
        ]
    
    # Convert date columns
    date_columns = ['Order Start Date', 'Customer Promise Date', 'Projected Date', 'Pending Approval Date']
    for col in date_columns:
        if col in sales_orders_df.columns:
            sales_orders_df[col] = pd.to_datetime(sales_orders_df[col], errors='coerce')
    
    # Calculate age for old pending approval
    if 'Order Start Date' in sales_orders_df.columns:
        today = pd.Timestamp.now()
        sales_orders_df['Age_Business_Days'] = sales_orders_df['Order Start Date'].apply(
            lambda x: business_days_between(x, today)
        )
    
    # Filter valid records
    if 'Amount' in sales_orders_df.columns and 'Sales Rep' in sales_orders_df.columns:
        sales_orders_df = sales_orders_df[
            (sales_orders_df['Amount'] > 0) & 
            (sales_orders_df['Sales Rep'].notna()) & 
            (sales_orders_df['Sales Rep'] != '') &
            (sales_orders_df['Sales Rep'] != 'nan')
        ]
    
    return sales_orders_df

# ==================== METRICS CALCULATION ====================
def calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df=None):
    """Calculate comprehensive metrics for a specific rep"""
    
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    if rep_info.empty:
        return None
    
    quota = rep_info['Quota'].iloc[0]
    orders = rep_info['NetSuite Orders'].iloc[0]
    
    # Filter deals for this rep
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    
    # Separate by shipping timeline
    rep_deals_ship_q4 = rep_deals[rep_deals.get('Ships_In_Q4', True) == True].copy()
    rep_deals_ship_q1 = rep_deals[rep_deals.get('Ships_In_Q1', False) == True].copy()
    
    # Calculate HubSpot metrics
    expect_commit_q4 = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    best_opp_q4 = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
    expect_commit_q1 = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    best_opp_q1 = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
    # Initialize sales order metrics
    pending_approval = 0
    pending_approval_no_date = 0
    pending_approval_old = 0
    pending_fulfillment = 0
    pending_fulfillment_no_date = 0
    
    pending_approval_details = pd.DataFrame()
    pending_fulfillment_details = pd.DataFrame()
    
    # Process sales orders
    if sales_orders_df is not None and not sales_orders_df.empty:
        rep_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
        
        if not rep_orders.empty:
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Pending Approval
            pa_orders = rep_orders[rep_orders['Status'] == 'Pending Approval'].copy()
            if not pa_orders.empty and 'Pending Approval Date' in pa_orders.columns:
                # With dates in Q4
                pa_with_date = pa_orders[
                    (pa_orders['Pending Approval Date'].notna()) &
                    (pa_orders['Pending Approval Date'] >= q4_start) &
                    (pa_orders['Pending Approval Date'] <= q4_end)
                ]
                pending_approval_details = pa_with_date.copy()
                pending_approval = pa_with_date['Amount'].sum()
                
                # Without dates
                pa_no_date = pa_orders[~pa_orders.index.isin(pa_with_date.index)]
                pending_approval_no_date = pa_no_date['Amount'].sum()
                
                # Old orders
                if 'Age_Business_Days' in pa_orders.columns:
                    old_orders = pa_orders[pa_orders['Age_Business_Days'] > 14]
                    pending_approval_old = old_orders['Amount'].sum()
            
            # Pending Fulfillment
            pf_orders = rep_orders[
                rep_orders['Status'].isin(['Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
            ].copy()
            
            if not pf_orders.empty:
                # Check for Q4 dates
                def has_q4_date(row):
                    if pd.notna(row.get('Customer Promise Date')) and q4_start <= row['Customer Promise Date'] <= q4_end:
                        return True
                    if pd.notna(row.get('Projected Date')) and q4_start <= row['Projected Date'] <= q4_end:
                        return True
                    return False
                
                pf_orders['Has_Q4_Date'] = pf_orders.apply(has_q4_date, axis=1)
                
                # With Q4 dates
                pf_with_date = pf_orders[pf_orders['Has_Q4_Date'] == True]
                pending_fulfillment_details = pf_with_date.copy()
                pending_fulfillment = pf_with_date['Amount'].sum()
                
                # Without dates
                pf_no_date = pf_orders[
                    (pf_orders['Customer Promise Date'].isna()) &
                    (pf_orders['Projected Date'].isna())
                ]
                pending_fulfillment_no_date = pf_no_date['Amount'].sum()
    
    # Calculate totals
    total_progress = orders + expect_commit_q4 + pending_approval + pending_fulfillment
    gap = quota - total_progress
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
    potential_attainment = ((total_progress + best_opp_q4) / quota * 100) if quota > 0 else 0
    
    return {
        'quota': quota,
        'orders': orders,
        'expect_commit': expect_commit_q4,
        'best_opp': best_opp_q4,
        'gap': gap,
        'attainment_pct': attainment_pct,
        'potential_attainment': potential_attainment,
        'total_progress': total_progress,
        'pending_approval': pending_approval,
        'pending_approval_no_date': pending_approval_no_date,
        'pending_approval_old': pending_approval_old,
        'pending_fulfillment': pending_fulfillment,
        'pending_fulfillment_no_date': pending_fulfillment_no_date,
        'q1_spillover_expect_commit': expect_commit_q1,
        'q1_spillover_best_opp': best_opp_q1,
        'q1_spillover_total': expect_commit_q1 + best_opp_q1,
        'pending_approval_details': pending_approval_details,
        'pending_fulfillment_details': pending_fulfillment_details,
        'deals': rep_deals_ship_q4
    }

# ==================== VISUALIZATION ====================
def create_gap_chart(metrics, title):
    """Create waterfall chart showing progress to goal"""
    fig = go.Figure()
    
    # Stacked bar
    fig.add_trace(go.Bar(
        name='NetSuite Orders',
        x=['Progress'],
        y=[metrics.get('orders', 0)],
        marker_color=COLORS['primary'],
        text=[f"${metrics.get('orders', 0):,.0f}"],
        textposition='inside'
    ))
    
    fig.add_trace(go.Bar(
        name='Expect/Commit',
        x=['Progress'],
        y=[metrics.get('expect_commit', 0)],
        marker_color=COLORS['success'],
        text=[f"${metrics.get('expect_commit', 0):,.0f}"],
        textposition='inside'
    ))
    
    # Quota line
    fig.add_trace(go.Scatter(
        name='Quota Goal',
        x=['Progress'],
        y=[metrics.get('quota', 0)],
        mode='markers',
        marker=dict(size=12, color=COLORS['danger'], symbol='diamond')
    ))
    
    fig.update_layout(
        title=title,
        barmode='stack',
        height=400,
        showlegend=True,
        yaxis_title="Amount ($)",
        xaxis_title=""
    )
    
    return fig

def display_progress_breakdown(metrics):
    """Display progress breakdown in a clean format"""
    st.markdown(f"""
    <div class="info-card">
        <h4>📊 Q4 Gap to Goal Components</h4>
        <table style="width:100%; font-size: 14px;">
            <tr><td>📦 Invoiced (Orders Shipped)</td><td style="text-align:right"><strong>${metrics['orders']:,.0f}</strong></td></tr>
            <tr><td>📤 Pending Fulfillment (with dates)</td><td style="text-align:right"><strong>${metrics['pending_fulfillment']:,.0f}</strong></td></tr>
            <tr><td>⏳ Pending Approval (with dates)</td><td style="text-align:right"><strong>${metrics['pending_approval']:,.0f}</strong></td></tr>
            <tr><td>✅ HubSpot Expect/Commit (Q4)</td><td style="text-align:right"><strong>${metrics['expect_commit']:,.0f}</strong></td></tr>
            <tr style="border-top: 2px solid #1E88E5; font-weight:bold;">
                <td>🎯 TOTAL PROGRESS</td><td style="text-align:right">${metrics['total_progress']:,.0f}</td>
            </tr>
        </table>
        <br/>
        <div style="display:flex; justify-content:space-around; text-align:center;">
            <div>
                <small>Current Attainment</small><br/>
                <strong style="font-size:20px; color:#1E88E5">{metrics['attainment_pct']:.1f}%</strong>
            </div>
            <div>
                <small>Gap to Goal</small><br/>
                <strong style="font-size:20px; color:{'#43A047' if metrics['gap'] <= 0 else '#DC3912'}">${abs(metrics['gap']):,.0f}</strong>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ==================== MAIN APP ====================
def main():
    load_custom_css()
    
    # Sidebar
    with st.sidebar:
        st.markdown("## 🎯 CALYX CONTAINERS")
        st.markdown("### Sales Dashboard")
        st.markdown("---")
        
        view_mode = st.radio(
            "Select View:",
            ["Individual Rep", "Team Overview", "Reconciliation"]
        )
        
        st.markdown("---")
        
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
        
        st.caption(f"Last updated: {datetime.now().strftime('%H:%M:%S')}")
    
    # Load data
    with st.spinner("Loading data..."):
        # Load all sheets
        deals_df = load_google_sheets_data("All Reps All Pipelines", "A:Q")
        dashboard_df = load_google_sheets_data("Dashboard Info", "A:C")
        sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AB")
        invoices_df = load_google_sheets_data("NS Invoices", "A:Z")
        
        # Process data
        deals_df = process_deals_data(deals_df)
        sales_orders_df = process_sales_orders(sales_orders_df)
        
        # Process dashboard data
        if not dashboard_df.empty and len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
            dashboard_df['Quota'] = dashboard_df['Quota'].apply(clean_numeric)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(clean_numeric)
    
    # Check data
    if deals_df.empty or dashboard_df.empty:
        st.error("❌ Unable to load required data. Please check your Google Sheets connection.")
        return
    
    # Display selected view
    if view_mode == "Individual Rep":
        st.title("📊 Sales Forecasting Dashboard - Q4 2025")
        
        rep_name = st.selectbox("Select Sales Rep:", dashboard_df['Rep Name'].tolist())
        
        if rep_name:
            metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics:
                # Display pipeline summary
                total_deals = len(deals_df[deals_df['Deal Owner'] == rep_name])
                total_amount = deals_df[deals_df['Deal Owner'] == rep_name]['Amount'].sum()
                
                if total_deals > 0:
                    st.markdown(f"""
                    <div class="info-card">
                        <strong>📋 Total Q4 2025 Pipeline:</strong> {total_deals} deals worth ${total_amount:,.0f}<br/>
                        <small>Note: ${metrics.get('q1_spillover_total', 0):,.0f} will ship in Q1 2026 based on lead times</small>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Section 1: Main metrics
                st.markdown("### 💰 Section 1: Q4 Gap to Goal Components")
                display_progress_breakdown(metrics)
                
                # Progress chart
                col1, col2 = st.columns(2)
                with col1:
                    fig = create_gap_chart(metrics, f"{rep_name}'s Progress to Goal")
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.metric("Quota", f"${metrics['quota']:,.0f}")
                    st.metric("Gap to Goal", f"${abs(metrics['gap']):,.0f}", 
                             delta="Over quota!" if metrics['gap'] <= 0 else "Under quota",
                             delta_color="normal" if metrics['gap'] <= 0 else "inverse")
                
                # Section 2: Additional orders
                st.markdown("### 📊 Section 2: Additional Orders")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("PF SO's No Date", f"${metrics['pending_fulfillment_no_date']:,.0f}")
                
                with col2:
                    st.metric("PA SO's No Date", f"${metrics['pending_approval_no_date']:,.0f}")
                
                with col3:
                    st.metric("Old PA (>14 days)", f"${metrics['pending_approval_old']:,.0f}",
                             delta="Needs attention" if metrics['pending_approval_old'] > 0 else None)
                
                # Section 3: Q1 Spillover
                if metrics.get('q1_spillover_total', 0) > 0:
                    st.markdown("### 🚢 Section 3: Q1 2026 Spillover")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("Expect/Commit (Q1)", f"${metrics.get('q1_spillover_expect_commit', 0):,.0f}")
                    
                    with col2:
                        st.metric("Best Case/Opp (Q1)", f"${metrics.get('q1_spillover_best_opp', 0):,.0f}")
                    
                    with col3:
                        st.metric("Total Q1 Spillover", f"${metrics.get('q1_spillover_total', 0):,.0f}")
                
                # Final total
                final_total = (metrics['total_progress'] + metrics['pending_fulfillment_no_date'] + 
                              metrics['pending_approval_no_date'] + metrics['pending_approval_old'])
                st.markdown("---")
                st.metric("📊 FINAL TOTAL Q4", f"${final_total:,.0f}",
                         delta=f"Section 1: ${metrics['total_progress']:,.0f} + Section 2: ${final_total - metrics['total_progress']:,.0f}")
    
    elif view_mode == "Team Overview":
        st.title("🎯 Team Sales Dashboard - Q4 2025")
        
        # Calculate team metrics
        total_quota = dashboard_df['Quota'].sum()
        total_orders = dashboard_df['NetSuite Orders'].sum()
        
        deals_q4 = deals_df[deals_df.get('Ships_In_Q4', True) == True]
        expect_commit = deals_q4[deals_q4['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
        best_opp = deals_q4[deals_q4['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
        
        current_forecast = expect_commit + total_orders
        gap = total_quota - current_forecast
        attainment_pct = (current_forecast / total_quota * 100) if total_quota > 0 else 0
        
        # Display metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Quota", f"${total_quota:,.0f}")
        
        with col2:
            st.metric("Current Forecast", f"${current_forecast:,.0f}",
                     delta=f"{attainment_pct:.1f}% of quota")
        
        with col3:
            st.metric("Gap to Goal", f"${abs(gap):,.0f}",
                     delta="Over quota!" if gap <= 0 else "Under quota",
                     delta_color="normal" if gap <= 0 else "inverse")
        
        with col4:
            potential = ((current_forecast + best_opp) / total_quota * 100) if total_quota > 0 else 0
            st.metric("Potential", f"{potential:.1f}%",
                     delta=f"+{potential - attainment_pct:.1f}% upside")
        
        # Rep summary
        st.markdown("### 👥 Rep Performance Summary")
        
        rep_summary = []
        for rep_name in dashboard_df['Rep Name']:
            rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            if rep_metrics:
                rep_summary.append({
                    'Rep': rep_name,
                    'Quota': f"${rep_metrics['quota']:,.0f}",
                    'Progress': f"${rep_metrics['total_progress']:,.0f}",
                    'Gap': f"${rep_metrics['gap']:,.0f}",
                    'Attainment': f"{rep_metrics['attainment_pct']:.1f}%"
                })
        
        if rep_summary:
            st.dataframe(pd.DataFrame(rep_summary), use_container_width=True, hide_index=True)
    
    else:  # Reconciliation view
        st.title("🔍 Forecast Reconciliation")
        st.info("Compare your numbers with the boss's forecast")
        
        # Simplified reconciliation table
        st.markdown("### Section 1: Q4 Gap to Goal Comparison")
        
        comparison_data = []
        for rep_name in dashboard_df['Rep Name']:
            metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            if metrics:
                comparison_data.append({
                    'Rep': rep_name,
                    'Invoiced': f"${metrics['orders']:,.0f}",
                    'Pending Fulfillment': f"${metrics['pending_fulfillment']:,.0f}",
                    'Pending Approval': f"${metrics['pending_approval']:,.0f}",
                    'HubSpot': f"${metrics['expect_commit']:,.0f}",
                    'Total': f"${metrics['total_progress']:,.0f}"
                })
        
        if comparison_data:
            st.dataframe(pd.DataFrame(comparison_data), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
