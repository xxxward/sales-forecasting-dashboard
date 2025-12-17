"""
Q1 2026 Sales Forecasting Module
Based on Sales Dashboard 43 architecture
Helps sales reps plan and forecast their Q1 2026 quarter
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import numpy as np

# ========== DATE CONSTANTS ==========
Q1_2026_START = pd.Timestamp('2026-01-01')
Q1_2026_END = pd.Timestamp('2026-03-31')
Q4_2025_START = pd.Timestamp('2025-10-01')
Q4_2025_END = pd.Timestamp('2025-12-31')
Q2_2026_START = pd.Timestamp('2026-04-01')

def get_mst_time():
    """Get current time in Mountain Standard Time"""
    return datetime.now(ZoneInfo("America/Denver"))

def calculate_business_days_until_q1():
    """Calculate business days from today until Q1 2026 starts"""
    from datetime import date
    
    today = date.today()
    q1_start = date(2026, 1, 1)
    
    if today >= q1_start:
        return 0
    
    holidays = [
        date(2025, 11, 27),  # Thanksgiving
        date(2025, 11, 28),  # Day after Thanksgiving
        date(2025, 12, 25),  # Christmas
        date(2025, 12, 26),  # Day after Christmas
    ]
    
    business_days = 0
    current_date = today
    
    while current_date < q1_start:
        if current_date.weekday() < 5 and current_date not in holidays:
            business_days += 1
        current_date += timedelta(days=1)
    
    return business_days

# ========== CUSTOM CSS FOR Q1 2026 MODULE ==========
def inject_custom_css():
    st.markdown("""
    <style>
    /* Q1 2026 Theme - Slightly different accent to distinguish from Q4 */
    .q1-header {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%);
        padding: 20px;
        border-radius: 16px;
        margin-bottom: 20px;
        text-align: center;
    }
    
    .q1-header h1 {
        color: white;
        margin: 0;
        font-size: 28px;
    }
    
    .q1-header p {
        color: rgba(255,255,255,0.8);
        margin: 5px 0 0 0;
    }
    
    /* Goal input styling */
    .goal-input-container {
        background: linear-gradient(135deg, rgba(5, 150, 105, 0.2) 0%, rgba(16, 185, 129, 0.2) 100%);
        border: 1px solid rgba(16, 185, 129, 0.3);
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 20px;
    }
    
    /* Sticky forecast bar for Q1 */
    .sticky-forecast-bar-q1 {
        position: fixed;
        bottom: 0;
        left: 22rem;
        right: 0;
        z-index: 9999;
        background: linear-gradient(135deg, rgba(5, 46, 22, 0.98) 0%, rgba(20, 83, 45, 0.98) 100%);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-top: 2px solid rgba(16, 185, 129, 0.5);
        padding: 16px 32px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 24px;
        box-shadow: 0 -8px 32px rgba(0, 0, 0, 0.4);
    }
    
    .sticky-forecast-item {
        text-align: center;
        flex: 1;
    }
    
    .sticky-forecast-label {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        opacity: 0.8;
        margin-bottom: 4px;
        font-weight: 600;
    }
    
    .sticky-forecast-value {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    
    .sticky-forecast-value.scheduled {
        color: #4ade80;
        text-shadow: 0 0 20px rgba(74, 222, 128, 0.5);
    }
    
    .sticky-forecast-value.pipeline {
        color: #60a5fa;
        text-shadow: 0 0 20px rgba(96, 165, 250, 0.5);
    }
    
    .sticky-forecast-value.total {
        font-size: 30px;
        background: linear-gradient(135deg, #10b981 0%, #3b82f6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        filter: drop-shadow(0 0 20px rgba(16, 185, 129, 0.5));
    }
    
    .sticky-forecast-value.gap-behind {
        color: #f87171;
        text-shadow: 0 0 20px rgba(248, 113, 113, 0.5);
    }
    
    .sticky-forecast-value.gap-ahead {
        color: #4ade80;
        text-shadow: 0 0 20px rgba(74, 222, 128, 0.5);
    }
    
    .sticky-forecast-divider {
        width: 1px;
        height: 50px;
        background: linear-gradient(180deg, transparent, rgba(255, 255, 255, 0.2), transparent);
    }
    
    @media (max-width: 768px) {
        .sticky-forecast-bar-q1 {
            left: 0;
        }
    }
    
    /* Add padding at bottom for sticky bar */
    .main .block-container {
        padding-bottom: 100px !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ========== DATA LOADING ==========
def load_q1_data(sales_orders_df, deals_df, dashboard_df):
    """
    Process data for Q1 2026 forecasting
    Returns filtered DataFrames for Q1 2026 planning
    """
    
    q1_data = {
        'sales_orders': pd.DataFrame(),
        'hubspot_q1_close': pd.DataFrame(),
        'hubspot_q4_spillover': pd.DataFrame(),
        'reps': []
    }
    
    # Get rep list from dashboard
    if dashboard_df is not None and not dashboard_df.empty:
        if 'Rep Name' in dashboard_df.columns:
            q1_data['reps'] = dashboard_df['Rep Name'].tolist()
    
    # Process Sales Orders for Q1 2026
    if sales_orders_df is not None and not sales_orders_df.empty:
        q1_data['sales_orders'] = sales_orders_df.copy()
    
    # Process HubSpot deals
    if deals_df is not None and not deals_df.empty:
        df = deals_df.copy()
        
        # Ensure Close Date is datetime
        if 'Close Date' in df.columns:
            df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
        
        # Parse Pending Approval Date (Column P in HubSpot data)
        if 'Pending Approval Date' in df.columns:
            df['PA_Date_Parsed'] = pd.to_datetime(df['Pending Approval Date'], errors='coerce')
        else:
            df['PA_Date_Parsed'] = pd.NaT
        
        # Q1 2026 Close Date deals
        q1_close_mask = (
            (df['Close Date'] >= Q1_2026_START) & 
            (df['Close Date'] <= Q1_2026_END)
        )
        q1_data['hubspot_q1_close'] = df[q1_close_mask].copy()
        
        # Q4 2025 Spillover: Close Date in Q4 2025 BUT PA Date in 2026
        q4_spillover_mask = (
            (df['Close Date'] >= Q4_2025_START) & 
            (df['Close Date'] <= Q4_2025_END) &
            (df['PA_Date_Parsed'] >= Q1_2026_START)
        )
        q1_data['hubspot_q4_spillover'] = df[q4_spillover_mask].copy()
    
    return q1_data


# ========== SALES ORDER CATEGORIZATION FOR Q1 2026 ==========
def categorize_sales_orders_q1(sales_orders_df, rep_name=None):
    """
    Categorize sales orders for Q1 2026 planning
    Returns orders with Promise/Projected Date OR PA Date in Q1 2026
    """
    
    empty_result = {
        'pf_date_ext': pd.DataFrame(), 'pf_date_ext_amount': 0,
        'pf_date_int': pd.DataFrame(), 'pf_date_int_amount': 0,
        'pa_date': pd.DataFrame(), 'pa_date_amount': 0,
        'pf_q2_spillover': pd.DataFrame(), 'pf_q2_spillover_amount': 0,
    }
    
    if sales_orders_df is None or sales_orders_df.empty:
        return empty_result
    
    # Filter by rep if specified
    if rep_name and 'Sales Rep' in sales_orders_df.columns:
        orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
    else:
        orders = sales_orders_df.copy()
    
    if orders.empty:
        return empty_result
    
    # Remove duplicate columns
    if orders.columns.duplicated().any():
        orders = orders.loc[:, ~orders.columns.duplicated()]
    
    # Helper to get column by index
    def get_col_by_index(df, index):
        if df is not None and len(df.columns) > index:
            return df.iloc[:, index]
        return pd.Series()
    
    # Add display columns
    orders['Display_SO_Num'] = get_col_by_index(orders, 1)  # Col B: SO#
    orders['Display_Type'] = get_col_by_index(orders, 17).fillna('Standard')  # Col R: Order Type
    orders['Display_Promise_Date'] = pd.to_datetime(get_col_by_index(orders, 11), errors='coerce')  # Col L
    orders['Display_Projected_Date'] = pd.to_datetime(get_col_by_index(orders, 12), errors='coerce')  # Col M
    
    # PA Date handling
    if 'Pending Approval Date' in orders.columns:
        orders['Display_PA_Date'] = pd.to_datetime(orders['Pending Approval Date'], errors='coerce')
    else:
        orders['Display_PA_Date'] = pd.to_datetime(get_col_by_index(orders, 29), errors='coerce')  # Col AD
    
    # Parse Customer Promise Date and Projected Date
    if 'Customer Promise Date' not in orders.columns:
        orders['Customer Promise Date'] = orders['Display_Promise_Date']
    else:
        orders['Customer Promise Date'] = pd.to_datetime(orders['Customer Promise Date'], errors='coerce')
    
    if 'Projected Date' not in orders.columns:
        orders['Projected Date'] = orders['Display_Projected_Date']
    else:
        orders['Projected Date'] = pd.to_datetime(orders['Projected Date'], errors='coerce')
    
    # === PENDING FULFILLMENT CATEGORIZATION FOR Q1 2026 ===
    pf_orders = orders[orders['Status'].isin(['Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])].copy()
    
    if not pf_orders.empty:
        # Check if dates are in Q1 2026 range
        def has_q1_2026_date(row):
            if pd.notna(row.get('Customer Promise Date')):
                if Q1_2026_START <= row['Customer Promise Date'] <= Q1_2026_END:
                    return True
            if pd.notna(row.get('Projected Date')):
                if Q1_2026_START <= row['Projected Date'] <= Q1_2026_END:
                    return True
            return False
        
        # Check if dates are in Q2 2026+ (spillover beyond Q1)
        def has_q2_2026_date(row):
            if pd.notna(row.get('Customer Promise Date')):
                if row['Customer Promise Date'] >= Q2_2026_START:
                    return True
            if pd.notna(row.get('Projected Date')):
                if row['Projected Date'] >= Q2_2026_START:
                    return True
            return False
        
        pf_orders['Has_Q1_2026_Date'] = pf_orders.apply(has_q1_2026_date, axis=1)
        pf_orders['Has_Q2_2026_Date'] = pf_orders.apply(has_q2_2026_date, axis=1)
        
        # Check External/Internal flag
        is_ext = pd.Series(False, index=pf_orders.index)
        if 'Calyx External Order' in pf_orders.columns:
            is_ext = pf_orders['Calyx External Order'].astype(str).str.strip().str.upper() == 'YES'
        
        # Q2 2026 Spillover (beyond Q1)
        pf_q2_spillover = pf_orders[pf_orders['Has_Q2_2026_Date'] == True].copy()
        spillover_ids = pf_q2_spillover.index
        
        # Q1 2026 PF orders (exclude Q2 spillover)
        q1_candidates = pf_orders[(pf_orders['Has_Q1_2026_Date'] == True) & (~pf_orders.index.isin(spillover_ids))]
        pf_date_ext = q1_candidates[is_ext.loc[q1_candidates.index]].copy()
        pf_date_int = q1_candidates[~is_ext.loc[q1_candidates.index]].copy()
    else:
        pf_date_ext = pf_date_int = pf_q2_spillover = pd.DataFrame()
    
    # === PENDING APPROVAL CATEGORIZATION FOR Q1 2026 ===
    pa_orders = orders[orders['Status'] == 'Pending Approval'].copy()
    
    if not pa_orders.empty:
        # Parse PA Date
        if 'Pending Approval Date' in pa_orders.columns:
            pa_orders['PA_Date_Parsed'] = pd.to_datetime(pa_orders['Pending Approval Date'], errors='coerce')
            
            # Fix 2-digit year issue (26 -> 1926 instead of 2026)
            if pa_orders['PA_Date_Parsed'].notna().any():
                mask_1900s = (pa_orders['PA_Date_Parsed'].dt.year < 2000) & (pa_orders['PA_Date_Parsed'].notna())
                if mask_1900s.any():
                    pa_orders.loc[mask_1900s, 'PA_Date_Parsed'] = pa_orders.loc[mask_1900s, 'PA_Date_Parsed'] + pd.DateOffset(years=100)
        else:
            pa_orders['PA_Date_Parsed'] = pd.NaT
        
        # PA with Q1 2026 Date
        has_q1_pa_date = (
            (pa_orders['PA_Date_Parsed'].notna()) &
            (pa_orders['PA_Date_Parsed'] >= Q1_2026_START) &
            (pa_orders['PA_Date_Parsed'] <= Q1_2026_END)
        )
        
        pa_date = pa_orders[has_q1_pa_date].copy()
    else:
        pa_date = pd.DataFrame()
    
    # Calculate amounts
    def get_amount(df):
        return df['Amount'].sum() if not df.empty and 'Amount' in df.columns else 0
    
    return {
        'pf_date_ext': pf_date_ext,
        'pf_date_ext_amount': get_amount(pf_date_ext),
        'pf_date_int': pf_date_int,
        'pf_date_int_amount': get_amount(pf_date_int),
        'pa_date': pa_date,
        'pa_date_amount': get_amount(pa_date),
        'pf_q2_spillover': pf_q2_spillover,
        'pf_q2_spillover_amount': get_amount(pf_q2_spillover),
    }


# ========== GAUGE CHART ==========
def create_q1_gauge(value, goal, title="Q1 2026 Progress"):
    """Create a gauge chart for Q1 2026 progress"""
    
    if goal <= 0:
        goal = 1  # Prevent division by zero
    
    percentage = (value / goal) * 100
    
    # Color based on progress
    if percentage >= 100:
        bar_color = "#4ade80"  # Green
    elif percentage >= 75:
        bar_color = "#fbbf24"  # Yellow
    elif percentage >= 50:
        bar_color = "#fb923c"  # Orange
    else:
        bar_color = "#f87171"  # Red
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        number={'prefix': "$", 'valueformat': ",.0f"},
        delta={'reference': goal, 'valueformat': ",.0f", 'prefix': "$"},
        title={'text': title, 'font': {'size': 16}},
        gauge={
            'axis': {'range': [0, goal * 1.2], 'tickformat': "$,.0f"},
            'bar': {'color': bar_color},
            'bgcolor': "rgba(255,255,255,0.1)",
            'borderwidth': 0,
            'steps': [
                {'range': [0, goal * 0.5], 'color': "rgba(248, 113, 113, 0.2)"},
                {'range': [goal * 0.5, goal * 0.75], 'color': "rgba(251, 191, 36, 0.2)"},
                {'range': [goal * 0.75, goal], 'color': "rgba(74, 222, 128, 0.2)"},
                {'range': [goal, goal * 1.2], 'color': "rgba(96, 165, 250, 0.2)"},
            ],
            'threshold': {
                'line': {'color': "white", 'width': 3},
                'thickness': 0.8,
                'value': goal
            }
        }
    ))
    
    fig.update_layout(
        height=300,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white'}
    )
    
    return fig


# ========== MAIN FUNCTION ==========
def main():
    """Main function for Q1 2026 Forecasting module"""
    
    inject_custom_css()
    
    # Header
    st.markdown("""
    <div class="q1-header">
        <h1>📦 Q1 2026 Sales Forecast</h1>
        <p>Plan and forecast your Q1 2026 quarter</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Calculate days until Q1
    days_until_q1 = calculate_business_days_until_q1()
    
    # Info bar
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📅 Q1 2026", "Jan 1 - Mar 31, 2026")
    with col2:
        st.metric("⏱️ Business Days Until Q1", days_until_q1)
    with col3:
        st.metric("🕐 Last Updated", get_mst_time().strftime('%I:%M %p MST'))
    
    st.markdown("---")
    
    # Load data from Google Sheets (standalone loading)
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        
        SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        
        @st.cache_data
        def load_sheet_data(sheet_name, range_name):
            """Load data from Google Sheets"""
            try:
                if "gcp_service_account" not in st.secrets:
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
                if len(values) > 1:
                    max_cols = max(len(row) for row in values)
                    for row in values:
                        while len(row) < max_cols:
                            row.append('')
                
                return pd.DataFrame(values[1:], columns=values[0])
            except Exception as e:
                st.error(f"Error loading {sheet_name}: {e}")
                return pd.DataFrame()
        
        # Load necessary data
        deals_df = load_sheet_data("All Reps All Pipelines", "A:R")
        dashboard_df = load_sheet_data("Dashboard Info", "A:C")
        sales_orders_df = load_sheet_data("NS Sales Orders", "A:AF")
        
    except ImportError as e:
        st.error(f"❌ Unable to import Google Sheets libraries: {e}")
        st.info("Make sure google-auth and google-api-python-client are installed")
        return
    
    # Process deals data
    if not deals_df.empty and len(deals_df.columns) >= 6:
        # Column mapping (same as main dashboard)
        col_names = deals_df.columns.tolist()
        rename_dict = {}
        
        for col in col_names:
            if col == 'Record ID':
                rename_dict[col] = 'Record ID'
            elif col == 'Deal Name':
                rename_dict[col] = 'Deal Name'
            elif col == 'Close Date':
                rename_dict[col] = 'Close Date'
            elif 'Deal Owner First Name' in col and 'Deal Owner Last Name' in col:
                rename_dict[col] = 'Deal Owner'
            elif col == 'Deal Owner First Name':
                rename_dict[col] = 'Deal Owner First Name'
            elif col == 'Deal Owner Last Name':
                rename_dict[col] = 'Deal Owner Last Name'
            elif col == 'Amount':
                rename_dict[col] = 'Amount'
            elif col == 'Close Status':
                rename_dict[col] = 'Status'
            elif col == 'Pipeline':
                rename_dict[col] = 'Pipeline'
            elif col == 'Deal Type':
                rename_dict[col] = 'Product Type'
            elif col == 'Pending Approval Date':
                rename_dict[col] = 'Pending Approval Date'
        
        deals_df = deals_df.rename(columns=rename_dict)
        
        # Create Deal Owner if not exists
        if 'Deal Owner' not in deals_df.columns:
            if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
        
        # Clean amount
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
        
        # Convert dates
        if 'Close Date' in deals_df.columns:
            deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
        
        if 'Pending Approval Date' in deals_df.columns:
            deals_df['Pending Approval Date'] = pd.to_datetime(deals_df['Pending Approval Date'], errors='coerce')
    
    # Process dashboard data
    if not dashboard_df.empty and len(dashboard_df.columns) >= 3:
        dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
        dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
    
    # Process sales orders data
    if not sales_orders_df.empty:
        # Map columns similar to main dashboard
        if len(sales_orders_df.columns) > 29:
            # Rename key columns
            col_mapping = {
                sales_orders_df.columns[1]: 'SO Number',
                sales_orders_df.columns[3]: 'Status',
                sales_orders_df.columns[6]: 'Customer',
                sales_orders_df.columns[10]: 'Amount',
                sales_orders_df.columns[11]: 'Customer Promise Date',
                sales_orders_df.columns[12]: 'Projected Date',
                sales_orders_df.columns[17]: 'Order Type',
                sales_orders_df.columns[29]: 'Pending Approval Date',
            }
            
            if len(sales_orders_df.columns) > 30:
                col_mapping[sales_orders_df.columns[30]] = 'Calyx External Order'
            if len(sales_orders_df.columns) > 31:
                col_mapping[sales_orders_df.columns[31]] = 'Sales Rep'
            
            sales_orders_df = sales_orders_df.rename(columns=col_mapping)
            
            # Clean Amount
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
            
            # Convert dates
            for date_col in ['Customer Promise Date', 'Projected Date', 'Pending Approval Date']:
                if date_col in sales_orders_df.columns:
                    sales_orders_df[date_col] = pd.to_datetime(sales_orders_df[date_col], errors='coerce')
    
    # Get rep list
    reps = dashboard_df['Rep Name'].tolist() if not dashboard_df.empty else []
    
    if not reps:
        st.warning("No reps found in Dashboard Info")
        return
    
    # Rep selector
    rep_name = st.selectbox("Select Sales Rep:", options=reps, key="q1_rep_selector")
    
    # === USER-DEFINED GOAL INPUT ===
    st.markdown("### 🎯 Set Your Q1 2026 Goal")
    
    goal_key = f"q1_goal_{rep_name}"
    if goal_key not in st.session_state:
        st.session_state[goal_key] = 1000000  # Default $1M
    
    col1, col2 = st.columns([2, 1])
    with col1:
        q1_goal = st.number_input(
            "Enter your Q1 2026 quota/goal ($):",
            min_value=0,
            max_value=50000000,
            value=st.session_state[goal_key],
            step=50000,
            format="%d",
            key=f"q1_goal_input_{rep_name}"
        )
        st.session_state[goal_key] = q1_goal
    
    with col2:
        st.metric("Your Q1 Goal", f"${q1_goal:,.0f}")
    
    st.markdown("---")
    
    # === BUILD YOUR OWN FORECAST SECTION ===
    st.markdown("### 🎯 Build Your Q1 2026 Forecast")
    st.caption("Select components to include in your Q1 2026 forecast.")
    
    # Load Q1 2026 data
    q1_data = load_q1_data(sales_orders_df, deals_df, dashboard_df)
    
    # Categorize sales orders for this rep
    so_categories = categorize_sales_orders_q1(sales_orders_df, rep_name)
    
    # === CATEGORY DEFINITIONS ===
    ns_categories = {
        'PF_Date_Ext': {'label': 'PF (Q1 Date) - External'},
        'PF_Date_Int': {'label': 'PF (Q1 Date) - Internal'},
        'PA_Date': {'label': 'PA (Q1 Date)'},
        'PF_Q2_Spillover': {'label': '⚠️ PF Spillover (Q2 2026+)'},
    }
    
    hs_categories = {
        'Q1_Expect': {'label': 'Q1 Close - Expect'},
        'Q1_Commit': {'label': 'Q1 Close - Commit'},
        'Q1_BestCase': {'label': 'Q1 Close - Best Case'},
        'Q1_Opp': {'label': 'Q1 Close - Opportunity'},
        'Q4_Spillover_Expect': {'label': 'Q4 Spillover - Expect'},
        'Q4_Spillover_Commit': {'label': 'Q4 Spillover - Commit'},
        'Q4_Spillover_BestCase': {'label': 'Q4 Spillover - Best Case'},
        'Q4_Spillover_Opp': {'label': 'Q4 Spillover - Opportunity'},
    }
    
    # Helper function to format NS view
    def format_ns_view(df, date_col_name):
        if df.empty:
            return df
        d = df.copy()
        
        if 'Sales Rep' not in d.columns and 'Rep Master' in d.columns:
            d['Sales Rep'] = d['Rep Master']
        
        # Add display columns
        def get_col_by_index(df, index):
            if df is not None and len(df.columns) > index:
                return df.iloc[:, index]
            return pd.Series()
        
        if 'Internal ID' in d.columns:
            d['Link'] = d['Internal ID'].apply(lambda x: f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}" if pd.notna(x) else "")
        
        if 'Display_SO_Num' in d.columns:
            d['SO #'] = d['Display_SO_Num']
        elif 'SO Number' in d.columns:
            d['SO #'] = d['SO Number']
        
        if 'Display_Type' in d.columns:
            d['Type'] = d['Display_Type']
        elif 'Order Type' in d.columns:
            d['Type'] = d['Order Type']
        
        # Ship Date
        if date_col_name == 'Promise':
            d['Ship Date'] = ''
            if 'Display_Promise_Date' in d.columns:
                promise_dates = pd.to_datetime(d['Display_Promise_Date'], errors='coerce')
                d.loc[promise_dates.notna(), 'Ship Date'] = promise_dates.dt.strftime('%Y-%m-%d')
            if 'Display_Projected_Date' in d.columns:
                projected_dates = pd.to_datetime(d['Display_Projected_Date'], errors='coerce')
                mask = (d['Ship Date'] == '') & projected_dates.notna()
                if mask.any():
                    d.loc[mask, 'Ship Date'] = projected_dates.loc[mask].dt.strftime('%Y-%m-%d')
        elif date_col_name == 'PA_Date':
            if 'Display_PA_Date' in d.columns:
                pa_dates = pd.to_datetime(d['Display_PA_Date'], errors='coerce')
                d['Ship Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('')
            else:
                d['Ship Date'] = ''
        else:
            d['Ship Date'] = ''
        
        return d.sort_values('Amount', ascending=False) if 'Amount' in d.columns else d
    
    # Map NS categories to dataframes
    ns_dfs = {
        'PF_Date_Ext': format_ns_view(so_categories['pf_date_ext'], 'Promise'),
        'PF_Date_Int': format_ns_view(so_categories['pf_date_int'], 'Promise'),
        'PA_Date': format_ns_view(so_categories['pa_date'], 'PA_Date'),
        'PF_Q2_Spillover': format_ns_view(so_categories['pf_q2_spillover'], 'Promise'),
    }
    
    # Process HubSpot data for this rep
    hs_dfs = {}
    
    # Q1 2026 Close Date deals
    q1_close = q1_data['hubspot_q1_close']
    if not q1_close.empty and 'Deal Owner' in q1_close.columns:
        rep_q1 = q1_close[q1_close['Deal Owner'] == rep_name]
        
        def format_hs_view(df):
            if df.empty:
                return df
            d = df.copy()
            if 'Record ID' in d.columns:
                d['Deal ID'] = d['Record ID']
                d['Link'] = d['Record ID'].apply(lambda x: f"https://app.hubspot.com/contacts/6712259/record/0-3/{x}/" if pd.notna(x) else "")
            if 'Close Date' in d.columns:
                d['Close'] = pd.to_datetime(d['Close Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
            if 'Amount' in d.columns:
                d['Amount_Numeric'] = pd.to_numeric(d['Amount'], errors='coerce').fillna(0)
            return d.sort_values('Amount_Numeric', ascending=False) if 'Amount_Numeric' in d.columns else d
        
        hs_dfs['Q1_Expect'] = format_hs_view(rep_q1[rep_q1['Status'] == 'Expect']) if 'Status' in rep_q1.columns else pd.DataFrame()
        hs_dfs['Q1_Commit'] = format_hs_view(rep_q1[rep_q1['Status'] == 'Commit']) if 'Status' in rep_q1.columns else pd.DataFrame()
        hs_dfs['Q1_BestCase'] = format_hs_view(rep_q1[rep_q1['Status'] == 'Best Case']) if 'Status' in rep_q1.columns else pd.DataFrame()
        hs_dfs['Q1_Opp'] = format_hs_view(rep_q1[rep_q1['Status'] == 'Opportunity']) if 'Status' in rep_q1.columns else pd.DataFrame()
    else:
        hs_dfs['Q1_Expect'] = pd.DataFrame()
        hs_dfs['Q1_Commit'] = pd.DataFrame()
        hs_dfs['Q1_BestCase'] = pd.DataFrame()
        hs_dfs['Q1_Opp'] = pd.DataFrame()
    
    # Q4 2025 Spillover deals
    q4_spillover = q1_data['hubspot_q4_spillover']
    if not q4_spillover.empty and 'Deal Owner' in q4_spillover.columns:
        rep_spillover = q4_spillover[q4_spillover['Deal Owner'] == rep_name]
        
        def format_hs_view(df):
            if df.empty:
                return df
            d = df.copy()
            if 'Record ID' in d.columns:
                d['Deal ID'] = d['Record ID']
                d['Link'] = d['Record ID'].apply(lambda x: f"https://app.hubspot.com/contacts/6712259/record/0-3/{x}/" if pd.notna(x) else "")
            if 'Close Date' in d.columns:
                d['Close'] = pd.to_datetime(d['Close Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
            if 'PA_Date_Parsed' in d.columns:
                d['PA Date'] = pd.to_datetime(d['PA_Date_Parsed'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
            if 'Amount' in d.columns:
                d['Amount_Numeric'] = pd.to_numeric(d['Amount'], errors='coerce').fillna(0)
            return d.sort_values('Amount_Numeric', ascending=False) if 'Amount_Numeric' in d.columns else d
        
        hs_dfs['Q4_Spillover_Expect'] = format_hs_view(rep_spillover[rep_spillover['Status'] == 'Expect']) if 'Status' in rep_spillover.columns else pd.DataFrame()
        hs_dfs['Q4_Spillover_Commit'] = format_hs_view(rep_spillover[rep_spillover['Status'] == 'Commit']) if 'Status' in rep_spillover.columns else pd.DataFrame()
        hs_dfs['Q4_Spillover_BestCase'] = format_hs_view(rep_spillover[rep_spillover['Status'] == 'Best Case']) if 'Status' in rep_spillover.columns else pd.DataFrame()
        hs_dfs['Q4_Spillover_Opp'] = format_hs_view(rep_spillover[rep_spillover['Status'] == 'Opportunity']) if 'Status' in rep_spillover.columns else pd.DataFrame()
    else:
        hs_dfs['Q4_Spillover_Expect'] = pd.DataFrame()
        hs_dfs['Q4_Spillover_Commit'] = pd.DataFrame()
        hs_dfs['Q4_Spillover_BestCase'] = pd.DataFrame()
        hs_dfs['Q4_Spillover_Opp'] = pd.DataFrame()
    
    # === EXPORT BUCKETS ===
    export_buckets = {}
    
    # === SELECT ALL / UNSELECT ALL BUTTONS ===
    sel_col1, sel_col2, sel_col3 = st.columns([1, 1, 2])
    with sel_col1:
        if st.button("☑️ Select All", key=f"q1_select_all_{rep_name}", use_container_width=True):
            for key in ns_categories.keys():
                df = ns_dfs.get(key, pd.DataFrame())
                val = df['Amount'].sum() if not df.empty and 'Amount' in df.columns else 0
                if val > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
            for key in hs_categories.keys():
                df = hs_dfs.get(key, pd.DataFrame())
                val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
                if val > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
            st.rerun()
    
    with sel_col2:
        if st.button("☐ Unselect All", key=f"q1_unselect_all_{rep_name}", use_container_width=True):
            for key in ns_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
            for key in hs_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
            st.rerun()
    
    # === RENDER UI ===
    with st.container():
        col_ns, col_hs = st.columns(2)
        
        # === NETSUITE COLUMN ===
        with col_ns:
            st.markdown("#### 📦 NetSuite Orders (Scheduled for Q1)")
            
            for key, data in ns_categories.items():
                df = ns_dfs.get(key, pd.DataFrame())
                val = df['Amount'].sum() if not df.empty and 'Amount' in df.columns else 0
                
                checkbox_key = f"q1_chk_{key}_{rep_name}"
                
                if val > 0:
                    is_checked = st.checkbox(
                        f"{data['label']}: ${val:,.0f}",
                        key=checkbox_key
                    )
                    
                    if is_checked:
                        with st.expander(f"🔎 View Orders ({data['label']})"):
                            if not df.empty:
                                enable_edit = st.toggle("Customize", key=f"q1_tgl_{key}_{rep_name}")
                                
                                display_cols = []
                                if 'Link' in df.columns: display_cols.append('Link')
                                if 'SO #' in df.columns: display_cols.append('SO #')
                                if 'Type' in df.columns: display_cols.append('Type')
                                if 'Customer' in df.columns: display_cols.append('Customer')
                                if 'Ship Date' in df.columns: display_cols.append('Ship Date')
                                if 'Amount' in df.columns: display_cols.append('Amount')
                                
                                if enable_edit and display_cols:
                                    df_edit = df.copy()
                                    
                                    unselected_key = f"q1_unselected_{key}_{rep_name}"
                                    if unselected_key not in st.session_state:
                                        st.session_state[unselected_key] = set()
                                    
                                    id_col = 'SO #' if 'SO #' in df_edit.columns else None
                                    
                                    # Row-level select/unselect buttons
                                    row_col1, row_col2, row_col3 = st.columns([1, 1, 2])
                                    with row_col1:
                                        if st.button("☑️ All", key=f"q1_row_sel_{key}_{rep_name}"):
                                            st.session_state[unselected_key] = set()
                                            st.rerun()
                                    with row_col2:
                                        if st.button("☐ None", key=f"q1_row_unsel_{key}_{rep_name}"):
                                            if id_col and id_col in df_edit.columns:
                                                st.session_state[unselected_key] = set(df_edit[id_col].astype(str).tolist())
                                            st.rerun()
                                    
                                    if id_col and id_col in df_edit.columns:
                                        df_edit.insert(0, "Select", df_edit[id_col].apply(
                                            lambda x: str(x) not in st.session_state[unselected_key]
                                        ))
                                    else:
                                        df_edit.insert(0, "Select", True)
                                    
                                    display_with_select = ['Select'] + display_cols
                                    
                                    edited = st.data_editor(
                                        df_edit[display_with_select],
                                        column_config={
                                            "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                            "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                            "SO #": st.column_config.TextColumn("SO #", width="small"),
                                            "Type": st.column_config.TextColumn("Type", width="small"),
                                            "Ship Date": st.column_config.TextColumn("Ship Date", width="small"),
                                            "Amount": st.column_config.NumberColumn("Amount", format="$%d")
                                        },
                                        disabled=[c for c in display_with_select if c != 'Select'],
                                        hide_index=True,
                                        key=f"q1_edit_{key}_{rep_name}",
                                        num_rows="fixed"
                                    )
                                    
                                    # Update unselected set
                                    if id_col and id_col in edited.columns:
                                        current_unselected = set()
                                        for idx, row in edited.iterrows():
                                            if not row['Select']:
                                                current_unselected.add(str(row[id_col]))
                                        st.session_state[unselected_key] = current_unselected
                                    
                                    selected_indices = edited[edited['Select']].index
                                    selected_rows = df.loc[selected_indices].copy()
                                    export_buckets[key] = selected_rows
                                    
                                    current_total = selected_rows['Amount'].sum() if 'Amount' in selected_rows.columns else 0
                                    st.caption(f"Selected: ${current_total:,.0f}")
                                else:
                                    if display_cols:
                                        st.dataframe(
                                            df[display_cols],
                                            column_config={
                                                "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                                "SO #": st.column_config.TextColumn("SO #", width="small"),
                                                "Type": st.column_config.TextColumn("Type", width="small"),
                                                "Ship Date": st.column_config.TextColumn("Ship Date", width="small"),
                                                "Amount": st.column_config.NumberColumn("Amount", format="$%d")
                                            },
                                            hide_index=True,
                                            use_container_width=True
                                        )
                                    export_buckets[key] = df
        
        # === HUBSPOT COLUMN ===
        with col_hs:
            st.markdown("#### 🎯 HubSpot Pipeline (Q1 2026)")
            
            for key, data in hs_categories.items():
                df = hs_dfs.get(key, pd.DataFrame())
                val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
                
                checkbox_key = f"q1_chk_{key}_{rep_name}"
                
                if val > 0:
                    is_checked = st.checkbox(
                        f"{data['label']}: ${val:,.0f}",
                        key=checkbox_key
                    )
                    
                    if is_checked:
                        with st.expander(f"🔎 View Deals ({data['label']})"):
                            if not df.empty:
                                enable_edit = st.toggle("Customize", key=f"q1_tgl_{key}_{rep_name}")
                                
                                display_cols = ['Link', 'Deal ID', 'Deal Name', 'Close', 'Amount_Numeric']
                                if 'PA Date' in df.columns:
                                    display_cols.insert(4, 'PA Date')
                                
                                if enable_edit:
                                    df_edit = df.copy()
                                    
                                    unselected_key = f"q1_unselected_{key}_{rep_name}"
                                    if unselected_key not in st.session_state:
                                        st.session_state[unselected_key] = set()
                                    
                                    id_col = 'Deal ID' if 'Deal ID' in df_edit.columns else None
                                    
                                    # Row-level select/unselect buttons
                                    row_col1, row_col2, row_col3 = st.columns([1, 1, 2])
                                    with row_col1:
                                        if st.button("☑️ All", key=f"q1_row_sel_{key}_{rep_name}"):
                                            st.session_state[unselected_key] = set()
                                            st.rerun()
                                    with row_col2:
                                        if st.button("☐ None", key=f"q1_row_unsel_{key}_{rep_name}"):
                                            if id_col and id_col in df_edit.columns:
                                                st.session_state[unselected_key] = set(df_edit[id_col].astype(str).tolist())
                                            st.rerun()
                                    
                                    if id_col and id_col in df_edit.columns:
                                        df_edit.insert(0, "Select", df_edit[id_col].apply(
                                            lambda x: str(x) not in st.session_state[unselected_key]
                                        ))
                                    else:
                                        df_edit.insert(0, "Select", True)
                                    
                                    display_with_select = ['Select'] + [c for c in display_cols if c in df_edit.columns]
                                    
                                    edited = st.data_editor(
                                        df_edit[display_with_select],
                                        column_config={
                                            "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                            "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                            "Deal ID": st.column_config.TextColumn("Deal ID", width="small"),
                                            "Deal Name": st.column_config.TextColumn("Deal Name", width="medium"),
                                            "Close": st.column_config.TextColumn("Close Date", width="small"),
                                            "PA Date": st.column_config.TextColumn("PA Date", width="small"),
                                            "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")
                                        },
                                        disabled=[c for c in display_with_select if c != 'Select'],
                                        hide_index=True,
                                        key=f"q1_edit_{key}_{rep_name}",
                                        num_rows="fixed"
                                    )
                                    
                                    # Update unselected set
                                    if id_col and id_col in edited.columns:
                                        current_unselected = set()
                                        for idx, row in edited.iterrows():
                                            if not row['Select']:
                                                current_unselected.add(str(row[id_col]))
                                        st.session_state[unselected_key] = current_unselected
                                    
                                    selected_indices = edited[edited['Select']].index
                                    selected_rows = df.loc[selected_indices].copy()
                                    export_buckets[key] = selected_rows
                                    
                                    current_total = selected_rows['Amount_Numeric'].sum() if 'Amount_Numeric' in selected_rows.columns else 0
                                    st.caption(f"Selected: ${current_total:,.0f}")
                                else:
                                    avail_cols = [c for c in display_cols if c in df.columns]
                                    if avail_cols:
                                        st.dataframe(
                                            df[avail_cols],
                                            column_config={
                                                "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                                "Deal ID": st.column_config.TextColumn("Deal ID", width="small"),
                                                "Deal Name": st.column_config.TextColumn("Deal Name", width="medium"),
                                                "Close": st.column_config.TextColumn("Close Date", width="small"),
                                                "PA Date": st.column_config.TextColumn("PA Date", width="small"),
                                                "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")
                                            },
                                            hide_index=True,
                                            use_container_width=True
                                        )
                                    export_buckets[key] = df
    
    # === CALCULATE RESULTS ===
    def safe_sum(df):
        if df.empty:
            return 0
        if 'Amount_Numeric' in df.columns:
            return df['Amount_Numeric'].sum()
        elif 'Amount' in df.columns:
            return df['Amount'].sum()
        return 0
    
    selected_scheduled = sum(safe_sum(df) for k, df in export_buckets.items() if k in ns_categories)
    selected_pipeline = sum(safe_sum(df) for k, df in export_buckets.items() if k in hs_categories)
    
    total_forecast = selected_scheduled + selected_pipeline
    gap_to_goal = q1_goal - total_forecast
    
    # === STICKY FORECAST SUMMARY BAR ===
    gap_class = "gap-behind" if gap_to_goal > 0 else "gap-ahead"
    gap_label = "GAP" if gap_to_goal > 0 else "AHEAD"
    gap_display = f"${abs(gap_to_goal):,.0f}"
    
    st.markdown(f"""
    <div class="sticky-forecast-bar-q1">
        <div class="sticky-forecast-item">
            <div class="sticky-forecast-label">Scheduled</div>
            <div class="sticky-forecast-value scheduled">${selected_scheduled:,.0f}</div>
        </div>
        <div class="sticky-forecast-divider"></div>
        <div class="sticky-forecast-item">
            <div class="sticky-forecast-label">+ Pipeline</div>
            <div class="sticky-forecast-value pipeline">${selected_pipeline:,.0f}</div>
        </div>
        <div class="sticky-forecast-divider"></div>
        <div class="sticky-forecast-item">
            <div class="sticky-forecast-label">= Forecast</div>
            <div class="sticky-forecast-value total">${total_forecast:,.0f}</div>
        </div>
        <div class="sticky-forecast-divider"></div>
        <div class="sticky-forecast-item">
            <div class="sticky-forecast-label">{gap_label}</div>
            <div class="sticky-forecast-value {gap_class}">{gap_display}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # === FORECAST RESULTS SECTION ===
    st.markdown("---")
    st.markdown("### 🔮 Q1 2026 Forecast Results")
    
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("1. Scheduled (NS)", f"${selected_scheduled:,.0f}")
    with m2:
        st.metric("2. Pipeline (HS)", f"${selected_pipeline:,.0f}")
    with m3:
        st.metric("🏁 Total Forecast", f"${total_forecast:,.0f}", delta="Sum of 1+2")
    with m4:
        if gap_to_goal > 0:
            st.metric("Gap to Goal", f"${gap_to_goal:,.0f}", delta="Behind", delta_color="inverse")
        else:
            st.metric("Gap to Goal", f"${abs(gap_to_goal):,.0f}", delta="Ahead!", delta_color="normal")
    
    # Gauge chart
    col1, col2 = st.columns([2, 1])
    with col1:
        fig = create_q1_gauge(total_forecast, q1_goal, "Q1 2026 Progress to Goal")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("#### 📊 Breakdown")
        st.markdown(f"""
        **Scheduled Orders (NetSuite):** ${selected_scheduled:,.0f}
        - Orders already in NetSuite with Q1 2026 dates
        
        **Pipeline (HubSpot):** ${selected_pipeline:,.0f}
        - Q1 2026 close date deals
        - Q4 2025 spillover (PA date in 2026)
        
        **Your Q1 Goal:** ${q1_goal:,.0f}
        """)
    
    # === EXPORT SECTION ===
    st.markdown("---")
    st.markdown("### 📤 Export Forecast")
    
    if export_buckets:
        # Combine all selected data for export
        all_ns_data = []
        all_hs_data = []
        
        for key, df in export_buckets.items():
            if df.empty:
                continue
            
            export_df = df.copy()
            export_df['Category'] = ns_categories.get(key, hs_categories.get(key, {})).get('label', key)
            
            if key in ns_categories:
                all_ns_data.append(export_df)
            else:
                all_hs_data.append(export_df)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if all_ns_data:
                ns_export = pd.concat(all_ns_data, ignore_index=True)
                csv_ns = ns_export.to_csv(index=False)
                st.download_button(
                    label="📥 Download NetSuite Orders (CSV)",
                    data=csv_ns,
                    file_name=f"q1_2026_netsuite_{rep_name.replace(' ', '_')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                st.caption(f"{len(ns_export)} orders, ${ns_export['Amount'].sum() if 'Amount' in ns_export.columns else 0:,.0f}")
        
        with col2:
            if all_hs_data:
                hs_export = pd.concat(all_hs_data, ignore_index=True)
                csv_hs = hs_export.to_csv(index=False)
                st.download_button(
                    label="📥 Download HubSpot Deals (CSV)",
                    data=csv_hs,
                    file_name=f"q1_2026_hubspot_{rep_name.replace(' ', '_')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                st.caption(f"{len(hs_export)} deals, ${hs_export['Amount_Numeric'].sum() if 'Amount_Numeric' in hs_export.columns else 0:,.0f}")
    else:
        st.info("Select some items above to enable export")


# Run if called directly
if __name__ == "__main__":
    main()
