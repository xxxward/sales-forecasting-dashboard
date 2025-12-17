"""
Q1 2026 Sales Forecasting Module
Based on Sales Dashboard architecture

KEY INSIGHT: The main dashboard's "spillover" buckets ARE the Q1 2026 scheduled orders!
- pf_spillover = PF orders with Q1 2026 dates
- pa_spillover = PA orders with PA Date in Q1 2026

This module imports directly from the main dashboard to reuse all data loading and categorization logic.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# ========== DATE CONSTANTS ==========
Q1_2026_START = pd.Timestamp('2026-01-01')
Q1_2026_END = pd.Timestamp('2026-03-31')
Q4_2025_START = pd.Timestamp('2025-10-01')
Q4_2025_END = pd.Timestamp('2025-12-31')


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


# ========== CUSTOM CSS ==========
def inject_custom_css():
    st.markdown("""
    <style>
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
    
    .main .block-container {
        padding-bottom: 100px !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ========== GAUGE CHART ==========
def create_q1_gauge(value, goal, title="Q1 2026 Progress"):
    """Create a clean gauge chart for Q1 2026 progress"""
    
    if goal <= 0:
        goal = 1
    
    percentage = (value / goal) * 100
    
    # Determine color based on progress
    if percentage >= 100:
        bar_color = "#4ade80"  # Green - at or above goal
    elif percentage >= 75:
        bar_color = "#fbbf24"  # Yellow
    elif percentage >= 50:
        bar_color = "#fb923c"  # Orange
    else:
        bar_color = "#f87171"  # Red
    
    # Set gauge range - adapt to actual value if it exceeds goal
    max_range = max(goal * 1.1, value * 1.05)
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={
            'prefix': "$", 
            'valueformat': ",.0f",
            'font': {'size': 48, 'color': 'white'}
        },
        title={
            'text': f"{title}<br><span style='font-size:14px;color:#888'>Goal: ${goal:,.0f}</span>",
            'font': {'size': 18, 'color': 'white'}
        },
        gauge={
            'axis': {
                'range': [0, max_range], 
                'tickmode': 'array',
                'tickvals': [0, goal * 0.5, goal, max_range],
                'ticktext': ['$0', f'${goal*0.5/1000:.0f}K', f'${goal/1000:.0f}K', ''],
                'tickfont': {'size': 10, 'color': '#888'},
                'showticklabels': True
            },
            'bar': {'color': bar_color, 'thickness': 0.75},
            'bgcolor': "rgba(255,255,255,0.05)",
            'borderwidth': 0,
            'steps': [
                {'range': [0, goal * 0.5], 'color': "rgba(248, 113, 113, 0.15)"},
                {'range': [goal * 0.5, goal * 0.75], 'color': "rgba(251, 191, 36, 0.15)"},
                {'range': [goal * 0.75, goal], 'color': "rgba(74, 222, 128, 0.15)"},
                {'range': [goal, max_range], 'color': "rgba(96, 165, 250, 0.15)"},
            ],
            'threshold': {
                'line': {'color': "#10b981", 'width': 4},
                'thickness': 0.85,
                'value': goal
            }
        }
    ))
    
    # Add percentage annotation
    fig.add_annotation(
        x=0.5, y=0.25,
        text=f"{percentage:.0f}% of Goal",
        showarrow=False,
        font=dict(size=16, color='#888' if percentage < 100 else '#4ade80'),
        xref="paper", yref="paper"
    )
    
    fig.update_layout(
        height=280,
        margin=dict(l=30, r=30, t=60, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white'}
    )
    
    return fig


# ========== FORMAT FUNCTIONS ==========
def get_col_by_index(df, index):
    """Safely get column by index"""
    if df is not None and len(df.columns) > index:
        return df.iloc[:, index]
    return pd.Series()


def format_ns_view(df, date_col_name):
    """Format NetSuite orders for display"""
    if df.empty:
        return df
    d = df.copy()
    
    # Remove duplicate columns
    if d.columns.duplicated().any():
        d = d.loc[:, ~d.columns.duplicated()]
    
    # Add Link column
    if 'Internal ID' in d.columns:
        d['Link'] = d['Internal ID'].apply(lambda x: f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}" if pd.notna(x) else "")
    
    # SO Number
    if 'Display_SO_Num' in d.columns:
        d['SO #'] = d['Display_SO_Num']
    elif 'Document Number' in d.columns:
        d['SO #'] = d['Document Number']
    
    # Type
    if 'Display_Type' in d.columns:
        d['Type'] = d['Display_Type']
    
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
        elif 'PA_Date_Parsed' in d.columns:
            pa_dates = pd.to_datetime(d['PA_Date_Parsed'], errors='coerce')
            d['Ship Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('')
        else:
            d['Ship Date'] = ''
    else:
        d['Ship Date'] = ''
    
    return d.sort_values('Amount', ascending=False) if 'Amount' in d.columns else d


def format_hs_view(df):
    """Format HubSpot deals for display"""
    if df.empty:
        return df
    d = df.copy()
    
    if 'Record ID' in d.columns:
        d['Deal ID'] = d['Record ID']
        d['Link'] = d['Record ID'].apply(lambda x: f"https://app.hubspot.com/contacts/6712259/record/0-3/{x}/" if pd.notna(x) else "")
    if 'Close Date' in d.columns:
        d['Close'] = pd.to_datetime(d['Close Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
    if 'Pending Approval Date' in d.columns:
        d['PA Date'] = pd.to_datetime(d['Pending Approval Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
    if 'Amount' in d.columns:
        d['Amount_Numeric'] = pd.to_numeric(d['Amount'], errors='coerce').fillna(0)
    return d.sort_values('Amount_Numeric', ascending=False) if 'Amount_Numeric' in d.columns else d


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
    
    # Days until Q1
    days_until_q1 = calculate_business_days_until_q1()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📅 Q1 2026", "Jan 1 - Mar 31, 2026")
    with col2:
        st.metric("⏱️ Business Days Until Q1", days_until_q1)
    with col3:
        st.metric("🕐 Last Updated", get_mst_time().strftime('%I:%M %p MST'))
    
    st.markdown("---")
    
    # Show data source info in sidebar
    st.sidebar.markdown("### 📊 Q1 2026 Data")
    st.sidebar.caption("HubSpot: Copy of All Reps All Pipelines")
    st.sidebar.caption("NetSuite: NS Sales Orders (spillover)")
    
    # === IMPORT FROM MAIN DASHBOARD ===
    # The main dashboard already has all the data loading and categorization logic
    # We import it directly to ensure consistency
    try:
        # Import the main dashboard module (it's named sales_dashboard.py in the repo)
        import sales_dashboard as main_dash
        
        # Load sales orders and dashboard data using the EXACT SAME function as the main dashboard
        deals_df_q4, dashboard_df, invoices_df, sales_orders_df, q4_push_df = main_dash.load_all_data()
        
        # Get the categorization function
        categorize_sales_orders = main_dash.categorize_sales_orders
        
        # NOW: Load Q1 2026 deals from "Copy of All Reps All Pipelines" 
        # This sheet includes BOTH Q4 2025 and Q1 2026 close dates
        deals_df = main_dash.load_google_sheets_data("Copy of All Reps All Pipelines", "A:R", version=main_dash.CACHE_VERSION)
        
        # Process the deals data (same logic as main dashboard)
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
                elif col == 'Q1 2026 Spillover':
                    rename_dict[col] = 'Q1 2026 Spillover'
            
            deals_df = deals_df.rename(columns=rename_dict)
            
            # Create Deal Owner if not exists
            if 'Deal Owner' not in deals_df.columns:
                if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                    deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                    deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
            else:
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
            
            # Filter out excluded deal stages
            excluded_stages = [
                '', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 
                'closed lost', 'closed won', 'sales order created in NS', 
                'NCR', 'Shipped'
            ]
            
            if 'Deal Stage' in deals_df.columns:
                deals_df['Deal Stage'] = deals_df['Deal Stage'].fillna('')
                deals_df['Deal Stage'] = deals_df['Deal Stage'].astype(str).str.strip()
                deals_df = deals_df[~deals_df['Deal Stage'].str.lower().isin([s.lower() if s else '' for s in excluded_stages])]
        
    except ImportError as e:
        st.error(f"❌ Unable to import main dashboard: {e}")
        st.info("Make sure sales_dashboard.py is in the same directory")
        return
    except Exception as e:
        st.error(f"❌ Error loading data: {e}")
        st.exception(e)
        return
    
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
        st.session_state[goal_key] = 1000000
    
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
    
    # === GET Q1 2026 DATA ===
    # The main dashboard's "spillover" buckets ARE the Q1 2026 scheduled orders!
    # - pf_spillover = PF orders with Q1 2026 Promise/Projected dates
    # - pa_spillover = PA orders with PA Date in Q1 2026
    
    so_categories = categorize_sales_orders(sales_orders_df, rep_name)
    
    # Map spillover to Q1 categories
    ns_categories = {
        'PF_Spillover': {'label': '📦 PF (Q1 2026 Date)', 'df': so_categories['pf_spillover'], 'amount': so_categories['pf_spillover_amount']},
        'PA_Spillover': {'label': '⏳ PA (Q1 2026 PA Date)', 'df': so_categories['pa_spillover'], 'amount': so_categories['pa_spillover_amount']},
    }
    
    # Format for display
    ns_dfs = {
        'PF_Spillover': format_ns_view(so_categories['pf_spillover'], 'Promise'),
        'PA_Spillover': format_ns_view(so_categories['pa_spillover'], 'PA_Date'),
    }
    
    # === HUBSPOT Q1 2026 PIPELINE ===
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
    
    hs_dfs = {}
    
    if not deals_df.empty and 'Deal Owner' in deals_df.columns:
        rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
        
        if 'Close Date' in rep_deals.columns:
            # Q1 2026 Close Date deals
            q1_close_mask = (rep_deals['Close Date'] >= Q1_2026_START) & (rep_deals['Close Date'] <= Q1_2026_END)
            q1_deals = rep_deals[q1_close_mask]
            
            # Q4 2025 Spillover - use the Q1 2026 Spillover column from spreadsheet
            if 'Q1 2026 Spillover' in rep_deals.columns:
                q4_spillover = rep_deals[rep_deals['Q1 2026 Spillover'] == 'Q1 2026']
            else:
                q4_spillover = pd.DataFrame()
            
            # Q1 Close deals by status
            if 'Status' in q1_deals.columns:
                hs_dfs['Q1_Expect'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Expect'])
                hs_dfs['Q1_Commit'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Commit'])
                hs_dfs['Q1_BestCase'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Best Case'])
                hs_dfs['Q1_Opp'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Opportunity'])
            
            # Q4 Spillover deals by status
            if not q4_spillover.empty and 'Status' in q4_spillover.columns:
                hs_dfs['Q4_Spillover_Expect'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Expect'])
                hs_dfs['Q4_Spillover_Commit'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Commit'])
                hs_dfs['Q4_Spillover_BestCase'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Best Case'])
                hs_dfs['Q4_Spillover_Opp'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Opportunity'])
    
    # Fill missing
    for key in hs_categories.keys():
        if key not in hs_dfs:
            hs_dfs[key] = pd.DataFrame()
    
    # === BUILD YOUR OWN FORECAST ===
    st.markdown("### 🎯 Build Your Q1 2026 Forecast")
    st.caption("Select components to include in your Q1 2026 forecast. NetSuite spillover orders are already scheduled for Q1.")
    
    export_buckets = {}
    
    # === CLEAR ALL SELECTIONS BUTTON (top right) ===
    clear_col1, clear_col2 = st.columns([3, 1])
    with clear_col2:
        if st.button("🗑️ Clear All Selections", key=f"q1_clear_all_{rep_name}"):
            # Clear all category checkboxes
            for key in ns_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
                # Also clear row-level selections
                st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            for key in hs_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
                st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            st.rerun()
    
    # === SELECT ALL / UNSELECT ALL ===
    sel_col1, sel_col2, sel_col3 = st.columns([1, 1, 2])
    with sel_col1:
        if st.button("☑️ Select All", key=f"q1_select_all_{rep_name}", use_container_width=True):
            for key in ns_categories.keys():
                if ns_categories[key]['amount'] > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
                    # Also clear row-level unselections so all rows are selected
                    st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            for key in hs_categories.keys():
                df = hs_dfs.get(key, pd.DataFrame())
                val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
                if val > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
                    st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
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
            st.markdown("#### 📦 NetSuite Orders (Q1 2026 Scheduled)")
            st.caption("These are spillover orders from Q4 with Q1 2026 ship/PA dates")
            
            for key, data in ns_categories.items():
                df = ns_dfs.get(key, pd.DataFrame())
                val = data['amount']
                
                checkbox_key = f"q1_chk_{key}_{rep_name}"
                
                if val > 0:
                    is_checked = st.checkbox(
                        f"{data['label']}: ${val:,.0f}",
                        key=checkbox_key
                    )
                    
                    if is_checked:
                        with st.expander(f"🔎 View Orders ({data['label']})"):
                            if not df.empty:
                                # Customize toggle
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
                                    
                                    # Session state for unselected rows
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
                                    
                                    # Add Select column
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
                                    
                                    # Get selected rows for export
                                    selected_indices = edited[edited['Select']].index
                                    selected_rows = df.loc[selected_indices].copy()
                                    export_buckets[key] = selected_rows
                                    
                                    current_total = selected_rows['Amount'].sum() if 'Amount' in selected_rows.columns else 0
                                    st.caption(f"Selected: ${current_total:,.0f}")
                                else:
                                    # Read-only view
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
                else:
                    st.caption(f"{data['label']}: $0")
        
        # === HUBSPOT COLUMN ===
        with col_hs:
            st.markdown("#### 🎯 HubSpot Pipeline (Q1 2026)")
            st.caption("Q1 close dates + Q4 spillover deals")
            
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
                                # Customize toggle
                                enable_edit = st.toggle("Customize", key=f"q1_tgl_{key}_{rep_name}")
                                
                                display_cols = ['Link', 'Deal ID', 'Deal Name', 'Close', 'Amount_Numeric']
                                if 'PA Date' in df.columns:
                                    display_cols.insert(4, 'PA Date')
                                
                                if enable_edit:
                                    df_edit = df.copy()
                                    
                                    # Session state for unselected rows
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
                                    
                                    # Add Select column
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
                                    
                                    # Get selected rows for export
                                    selected_indices = edited[edited['Select']].index
                                    selected_rows = df.loc[selected_indices].copy()
                                    export_buckets[key] = selected_rows
                                    
                                    current_total = selected_rows['Amount_Numeric'].sum() if 'Amount_Numeric' in selected_rows.columns else 0
                                    st.caption(f"Selected: ${current_total:,.0f}")
                                else:
                                    # Read-only view
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
    
    # === STICKY SUMMARY BAR ===
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
    
    # === RESULTS SECTION ===
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
    
    # Gauge
    col1, col2 = st.columns([2, 1])
    with col1:
        fig = create_q1_gauge(total_forecast, q1_goal, "Q1 2026 Progress to Goal")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("#### 📊 Breakdown")
        st.markdown(f"""
        **Scheduled Orders (NetSuite):** ${selected_scheduled:,.0f}
        - PF Spillover: Orders with Q1 2026 ship dates
        - PA Spillover: Orders with Q1 2026 PA dates
        
        **Pipeline (HubSpot):** ${selected_pipeline:,.0f}
        - Q1 2026 close date deals
        - Q4 2025 spillover deals
        
        **Your Q1 Goal:** ${q1_goal:,.0f}
        """)
    
    # === EXPORT SECTION ===
    st.markdown("---")
    st.markdown("### 📤 Export Q1 2026 Forecast")
    
    if export_buckets:
        # Combine all selected data for export
        all_ns_data = []
        all_hs_data = []
        
        for key, df in export_buckets.items():
            if df.empty:
                continue
            
            export_df = df.copy()
            
            # Add category label
            if key in ns_categories:
                export_df['Category'] = ns_categories[key]['label']
                all_ns_data.append(export_df)
            elif key in hs_categories:
                export_df['Category'] = hs_categories[key]['label']
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
                ns_total = ns_export['Amount'].sum() if 'Amount' in ns_export.columns else 0
                st.caption(f"{len(ns_export)} orders, ${ns_total:,.0f}")
            else:
                st.info("No NetSuite orders selected")
        
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
                hs_total = hs_export['Amount_Numeric'].sum() if 'Amount_Numeric' in hs_export.columns else 0
                st.caption(f"{len(hs_export)} deals, ${hs_total:,.0f}")
            else:
                st.info("No HubSpot deals selected")
    else:
        st.info("Select items above to enable export")
    
    # === DEBUG INFO ===
    with st.expander("🔧 Debug: Data Summary"):
        st.write("**Data Source:** Copy of All Reps All Pipelines (Q4 2025 + Q1 2026 deals)")
        st.write(f"**Total Deals Loaded:** {len(deals_df)}")
        st.write(f"**PF Spillover:** {len(so_categories['pf_spillover'])} orders, ${so_categories['pf_spillover_amount']:,.0f}")
        st.write(f"**PA Spillover:** {len(so_categories['pa_spillover'])} orders, ${so_categories['pa_spillover_amount']:,.0f}")
        
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
            st.write(f"**{key}:** {len(df)} deals, ${val:,.0f}")


# Run if called directly
if __name__ == "__main__":
    main()
