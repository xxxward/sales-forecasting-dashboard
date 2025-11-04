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
from io import BytesIO
from PIL import Image

# =========================
# Page / Branding Settings
# =========================

# Put your logo file path here (local in the app repo or mounted path)
LOGO_PATH = "/mnt/data/25a8bd3a-7e41-4692-b9d0-c06e29aa57f7.png"  # change if you store it elsewhere

def _load_logo(path_or_fallback="📊"):
    try:
        return Image.open(path_or_fallback)
    except Exception:
        return "📊"

PAGE_ICON = _load_logo(LOGO_PATH)

st.set_page_config(
    page_title="Sales Forecasting Dashboard",
    page_icon=PAGE_ICON,  # real image => shows as favicon/tab icon
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============
# Custom Styles
# =============
st.markdown("""
<style>
.big-font { font-size: 28px !important; font-weight: bold; }
.metric-card { background-color: #f0f2f6; padding: 20px; border-radius: 10px; box-shadow: 2px 2px 5px rgba(0,0,0,0.1); }
.stMetric { background-color: #ffffff; padding: 15px; border-radius: 8px; box-shadow: 1px 1px 3px rgba(0,0,0,0.1); }
</style>
""", unsafe_allow_html=True)

# =================
# Google Sheets IO
# =================
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600
# Version number to force cache refresh when code changes
CACHE_VERSION = "v12"

# Global Q4 window reused across functions
Q4_START = pd.Timestamp('2025-10-01')
Q4_END   = pd.Timestamp('2025-12-31')

# ========
# Debugging
# ========
# Toggle in sidebar (dev only). Default False in prod.
with st.sidebar:
    st.image(LOGO_PATH, use_container_width=True)
    st.markdown("---")
    st.markdown("### 🎯 Dashboard Navigation (use main area)")
    st.markdown("---")
    DEBUG = st.checkbox("Show debug logs", value=False)

def dbg(*args, **kwargs):
    if DEBUG:
        st.sidebar.write(*args, **kwargs)

def dbg_df(df, **kwargs):
    if DEBUG:
        st.sidebar.dataframe(df, **kwargs)

# ============
# Data Loading
# ============
@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets with caching"""
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

        dbg(f"**DEBUG - {sheet_name}:**")
        dbg(f"Total rows loaded: {len(values)}")
        dbg(f"First 3 rows: {values[:3]}")

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
        dbg(f"**ERROR in {sheet_name}:** {str(e)}")
        return pd.DataFrame()

def _clean_numeric(value):
    if pd.isna(value) or value == '':
        return 0
    cleaned = (str(value)
               .replace(',', '')
               .replace('$', '')
               .replace('′', '')
               .replace('’', '')
               .replace(' ', '')
               .strip())
    try:
        return float(cleaned)
    except:
        return 0

def load_all_data():
    """Load and clean all necessary data from Google Sheets"""
    # Load sheets
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:H", version=CACHE_VERSION)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    invoices_df = load_google_sheets_data("NS Invoices", "A:Z", version=CACHE_VERSION)
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:Z", version=CACHE_VERSION)

    # Deals cleanup
    if not deals_df.empty and len(deals_df.columns) >= 8:
        col_names = deals_df.columns.tolist()
        deals_df = deals_df.rename(columns={
            col_names[1]: 'Deal Name',
            col_names[3]: 'Close Date',
            col_names[4]: 'Deal Owner',
            col_names[5]: 'Amount',
            col_names[6]: 'Status',
            col_names[7]: 'Pipeline'
        })
        deals_df['Amount'] = deals_df['Amount'].apply(_clean_numeric)
        deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')

    # Dashboard Info cleanup
    if not dashboard_df.empty:
        dbg("**DEBUG - Dashboard Info Raw:**")
        dbg_df(dashboard_df.head())
        if len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            dashboard_df = dashboard_df[
                dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')
            ]
            dashboard_df['Quota'] = dashboard_df['Quota'].apply(_clean_numeric)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(_clean_numeric)
            dbg("**DEBUG - After Conversion:**")
            dbg_df(dashboard_df)
        else:
            st.error(f"Dashboard Info sheet has wrong number of columns: {len(dashboard_df.columns)}")

    # Invoices cleanup (Q4 only)
    if not invoices_df.empty:
        dbg("**DEBUG - NS Invoices loaded:**", len(invoices_df), "rows")
        if len(invoices_df.columns) >= 15:
            invoices_df = invoices_df.rename(columns={
                invoices_df.columns[0]: 'Invoice Number',
                invoices_df.columns[1]: 'Status',
                invoices_df.columns[2]: 'Date',
                invoices_df.columns[6]: 'Customer',
                invoices_df.columns[10]: 'Amount',
                invoices_df.columns[14]: 'Sales Rep'
            })
            invoices_df['Amount'] = invoices_df['Amount'].apply(_clean_numeric)
            invoices_df['Date'] = pd.to_datetime(invoices_df['Date'], errors='coerce')

            invoices_df = invoices_df[
                (invoices_df['Date'] >= Q4_START) & (invoices_df['Date'] <= Q4_END)
            ]
            dbg(f"**DEBUG - After Q4 2025 filter:** {len(invoices_df)} rows")

            invoices_df['Sales Rep'] = invoices_df['Sales Rep'].astype(str).str.strip()
            invoices_df = invoices_df[
                (invoices_df['Amount'] > 0) & (invoices_df['Sales Rep'].notna()) & (invoices_df['Sales Rep'] != '')
            ]
            dbg(f"**DEBUG - After cleaning:** {len(invoices_df)} rows")
            dbg("**DEBUG - Unique sales reps in invoices:**", invoices_df['Sales Rep'].unique().tolist())

            invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
            invoice_totals.columns = ['Rep Name', 'Invoice Total']
            dbg("**DEBUG - Invoice totals by rep:**")
            dbg_df(invoice_totals)

            if not dashboard_df.empty:
                dashboard_df['Rep Name'] = dashboard_df['Rep Name'].astype(str).str.strip()
                dbg("**DEBUG - Rep names in Dashboard Info:**", dashboard_df['Rep Name'].tolist())
                dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
                dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
                dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
                dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
                dbg("**DEBUG - Final dashboard_df with invoice totals:**")
                dbg_df(dashboard_df)
        else:
            st.warning(f"NS Invoices sheet doesn't have enough columns (has {len(invoices_df.columns)})")
            invoices_df = pd.DataFrame()

    # Sales Orders cleanup
    if not sales_orders_df.empty:
        dbg("**DEBUG - NS Sales Orders loaded:**", len(sales_orders_df), "rows")
        status_col = amount_col = sales_rep_col = None
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
            if 'pending fulfillment date' in col_lower:
                pending_fulfillment_date_col = col
            if 'projected date' in col_lower:
                projected_date_col = col
            if 'customer promise last date to ship' in col_lower or 'customer promise date' in col_lower:
                customer_promise_col = col

        dbg("**DEBUG - Found columns:**")
        dbg(f"Status={status_col}, Amount={amount_col}, Rep={sales_rep_col}")
        dbg(f"Dates: J={pending_fulfillment_date_col}, M={projected_date_col}, L={customer_promise_col}")

        if status_col and amount_col and sales_rep_col:
            rename_dict = {status_col: 'Status', amount_col: 'Amount', sales_rep_col: 'Sales Rep'}
            if pending_fulfillment_date_col: rename_dict[pending_fulfillment_date_col] = 'Pending Fulfillment Date'
            if projected_date_col: rename_dict[projected_date_col] = 'Projected Date'
            if customer_promise_col: rename_dict[customer_promise_col] = 'Customer Promise Date'
            sales_orders_df = sales_orders_df.rename(columns=rename_dict)

            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(_clean_numeric)
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].astype(str).str.strip()
            sales_orders_df['Status'] = sales_orders_df['Status'].astype(str).str.strip()

            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment'])
            ]
            dbg(f"**DEBUG - After status filter:** {len(sales_orders_df)} rows")

            if 'Pending Fulfillment Date' in sales_orders_df.columns:
                def get_fulfillment_date(row):
                    if row['Status'] == 'Pending Approval':
                        return None
                    if pd.notna(row.get('Pending Fulfillment Date')) and row.get('Pending Fulfillment Date') != '':
                        return row.get('Pending Fulfillment Date')
                    if pd.notna(row.get('Projected Date')) and row.get('Projected Date') != '':
                        return row.get('Projected Date')
                    if pd.notna(row.get('Customer Promise Date')) and row.get('Customer Promise Date') != '':
                        return row.get('Customer Promise Date')
                    return None

                sales_orders_df['Effective Date'] = sales_orders_df.apply(get_fulfillment_date, axis=1)
                sales_orders_df['Effective Date'] = pd.to_datetime(sales_orders_df['Effective Date'], errors='coerce')

                pending_approval = sales_orders_df[sales_orders_df['Status'] == 'Pending Approval']
                pending_fulfillment_q4 = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Fulfillment') &
                    (sales_orders_df['Effective Date'] >= Q4_START) &
                    (sales_orders_df['Effective Date'] <= Q4_END)
                ]
                pending_fulfillment_no_date = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Fulfillment') &
                    (sales_orders_df['Effective Date'].isna())
                ]
                sales_orders_df = pd.concat([pending_approval, pending_fulfillment_q4, pending_fulfillment_no_date])

                dbg("**DEBUG - After Q4 filter:**")
                dbg(f"Pending Approval: {len(pending_approval)} rows")
                dbg(f"Pending Fulfillment Q4: {len(pending_fulfillment_q4)} rows")
                dbg(f"Pending Fulfillment No Date: {len(pending_fulfillment_no_date)} rows")

            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & (sales_orders_df['Sales Rep'].notna()) & (sales_orders_df['Sales Rep'] != '')
            ]
            dbg(f"**DEBUG - Final count:** {len(sales_orders_df)} rows")
        else:
            st.warning("Could not find required columns in NS Sales Orders")
            sales_orders_df = pd.DataFrame()

    return deals_df, dashboard_df, invoices_df, sales_orders_df

# ==================
# Metrics Calculators
# ==================
def calculate_team_metrics(deals_df, dashboard_df):
    """Calculate overall team metrics"""
    total_quota = dashboard_df['Quota'].sum()
    total_orders = dashboard_df['NetSuite Orders'].sum()

    expect_commit = deals_df[deals_df['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    best_opp = deals_df[deals_df['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()

    gap = total_quota - expect_commit - total_orders
    current_forecast = expect_commit + total_orders
    attainment_pct = (current_forecast / total_quota * 100) if total_quota > 0 else 0
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
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    if rep_info.empty:
        return None

    quota = rep_info['Quota'].iloc[0]
    orders = rep_info['NetSuite Orders'].iloc[0]

    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name]
    expect_commit = rep_deals[rep_deals['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    best_opp = rep_deals[rep_deals['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()

    pending_approval = 0
    pending_fulfillment = 0
    pending_fulfillment_no_date = 0

    if sales_orders_df is not None and not sales_orders_df.empty:
        rep_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name]
        pending_approval = rep_orders[rep_orders['Status'] == 'Pending Approval']['Amount'].sum()

        if 'Effective Date' in rep_orders.columns:
            pending_fulfillment = rep_orders[
                (rep_orders['Status'] == 'Pending Fulfillment') & (rep_orders['Effective Date'].notna())
            ]['Amount'].sum()
            pending_fulfillment_no_date = rep_orders[
                (rep_orders['Status'] == 'Pending Fulfillment') & (rep_orders['Effective Date'].isna())
            ]['Amount'].sum()
        else:
            pending_fulfillment = rep_orders[rep_orders['Status'] == 'Pending Fulfillment']['Amount'].sum()

    total_progress = orders + expect_commit + pending_approval + pending_fulfillment
    gap = quota - total_progress
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
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

# ==========
# Visuals
# ==========
def create_gap_chart(metrics, title):
    """Create a stacked bar + markers showing progress to goal"""
    fig = go.Figure()

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

    # Quota goal
    fig.add_trace(go.Scatter(
        name='Quota Goal',
        x=['Progress'],
        y=[metrics['total_quota'] if 'total_quota' in metrics else metrics['quota']],
        mode='markers',
        marker=dict(size=12, color='#DC3912', symbol='diamond'),
        text=[f"Goal: ${metrics['total_quota'] if 'total_quota' in metrics else metrics['quota']:,.0f}"],
        textposition='top center'
    ))

    # Potential
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
        status_summary, values='Amount', names='Status',
        title='Deal Amount by Forecast Category',
        color='Status', color_discrete_map=color_map, hole=0.4
    )
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=400)
    return fig

def create_pipeline_breakdown_chart(deals_df, rep_name=None):
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    pipeline_summary = deals_df.groupby(['Pipeline', 'Status'])['Amount'].sum().reset_index()
    color_map = {
        'Expect': '#1E88E5',
        'Commit': '#43A047',
        'Best Case': '#FB8C00',
        'Opportunity': '#8E24AA'
    }
    fig = px.bar(
        pipeline_summary, x='Pipeline', y='Amount', color='Status',
        title='Pipeline Breakdown by Forecast Category',
        color_discrete_map=color_map, text_auto='.2s', barmode='stack'
    )
    fig.update_layout(height=400, yaxis_title="Amount ($)", xaxis_title="Pipeline")
    return fig

def create_deals_timeline(deals_df, rep_name=None):
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    timeline_df = deals_df[deals_df['Close Date'].notna()].copy()
    if timeline_df.empty:
        return None
    timeline_df = timeline_df.sort_values('Close Date')
    color_map = {
        'Expect': '#1E88E5',
        'Commit': '#43A047',
        'Best Case': '#FB8C00',
        'Opportunity': '#8E24AA'
    }
    fig = px.scatter(
        timeline_df, x='Close Date', y='Amount', color='Status', size='Amount',
        hover_data=['Deal Name', 'Amount', 'Pipeline'],
        title='Deal Close Date Timeline',
        color_discrete_map=color_map
    )
    fig.update_layout(height=400, yaxis_title="Deal Amount ($)", xaxis_title="Expected Close Date")
    return fig

def create_invoice_status_chart(invoices_df, rep_name=None):
    if invoices_df.empty:
        return None
    if rep_name:
        invoices_df = invoices_df[invoices_df['Sales Rep'] == rep_name]
        if invoices_df.empty:
            return None
    status_summary = invoices_df.groupby('Status')['Amount'].sum().reset_index()
    fig = px.pie(
        status_summary, values='Amount', names='Status',
        title='Invoice Amount by Status', hole=0.4
    )
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=400)
    return fig

def create_customer_invoice_table(invoices_df, rep_name):
    if invoices_df.empty:
        return pd.DataFrame()
    rep_invoices = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
    if rep_invoices.empty:
        return pd.DataFrame()
    customer_summary = rep_invoices.groupby(['Customer', 'Status'])['Amount'].sum().reset_index()
    pivot_table = customer_summary.pivot_table(
        index='Customer', columns='Status', values='Amount',
        fill_value=0, aggfunc='sum'
    ).reset_index()
    status_cols = [c for c in pivot_table.columns if c != 'Customer']
    pivot_table['Total'] = pivot_table[status_cols].sum(axis=1)
    pivot_table = pivot_table.sort_values('Total', ascending=False)
    return pivot_table

def render_total_progress(metrics):
    """Render chip-style breakdown under progress bar"""
    html = f"""
    <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;font-size:14px;">
      <div><strong>Total Progress:</strong> ${metrics['total_progress']:,.0f}</div>
      <span style="background:#eef3ff;padding:4px 8px;border-radius:999px;">Orders: ${metrics['orders']:,.0f}</span>
      <span style="background:#e8f5e9;padding:4px 8px;border-radius:999px;">Expect/Commit: ${metrics['expect_commit']:,.0f}</span>
      <span style="background:#fff8e1;padding:4px 8px;border-radius:999px;">Pending Approval: ${metrics['pending_approval']:,.0f}</span>
      <span style="background:#fff3e0;padding:4px 8px;border-radius:999px;">Pending Fulfillment: ${metrics['pending_fulfillment']:,.0f}</span>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

# ===================
# Dashboard Rendering
# ===================
def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    st.title("🎯 Team Sales Dashboard - Q4 2025")

    metrics = calculate_team_metrics(deals_df, dashboard_df)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Quota", value=f"${metrics['total_quota']:,.0f}")
    with col2:
        # FIX: show actual orders shipped, not current_forecast
        st.metric("Orders Shipped NetSuite", value=f"${metrics['total_orders']:,.0f}",
                  delta=f"{metrics['attainment_pct']:.1f}% of quota")
    with col3:
        st.metric("Gap to Goal", value=f"${metrics['gap']:,.0f}",
                  delta=f"{-metrics['gap']:,.0f}" if metrics['gap'] < 0 else None,
                  delta_color="inverse")
    with col4:
        st.metric("Potential Attainment", value=f"{metrics['potential_attainment']:.1f}%",
                  delta=f"+{metrics['potential_attainment'] - metrics['attainment_pct']:.1f}% upside")

    st.markdown("### 📈 Progress to Quota")
    progress = min(metrics['attainment_pct'] / 100, 1.0)
    st.progress(progress)
    st.caption(f"Current: {metrics['attainment_pct']:.1f}% | Potential: {metrics['potential_attainment']:.1f}%")

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(create_gap_chart(metrics, "Team Progress to Goal"), use_container_width=True)
    with col2:
        st.plotly_chart(create_status_breakdown_chart(deals_df), use_container_width=True)

    st.markdown("### 🔄 Pipeline Analysis")
    st.plotly_chart(create_pipeline_breakdown_chart(deals_df), use_container_width=True)

    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)

    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Status Breakdown")
        invoice_chart = create_invoice_status_chart(invoices_df)
        if invoice_chart:
            st.plotly_chart(invoice_chart, use_container_width=True)

    # Rep summary table (FIX: use 'orders' not 'current_forecast')
    st.markdown("### 👥 Rep Summary")
    rep_summary = []
    for rep_name in dashboard_df['Rep Name']:
        rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
        if rep_metrics:
            rep_summary.append({
                'Rep': rep_name,
                'Quota': f"${rep_metrics['quota']:,.0f}",
                'Orders Shipped': f"${rep_metrics['orders']:,.0f}",
                'Pending Approval': f"${rep_metrics['pending_approval']:,.0f}",
                'Pending Fulfillment': f"${rep_metrics['pending_fulfillment']:,.0f}",
                'Gap': f"${rep_metrics['gap']:,.0f}",
                'Attainment': f"{rep_metrics['attainment_pct']:.1f}%"
            })
    rep_summary_df = pd.DataFrame(rep_summary)
    st.dataframe(rep_summary_df, use_container_width=True, hide_index=True)

def display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    st.title(f"👤 {rep_name}'s Q4 2025 Forecast")
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
    if not metrics:
        st.error(f"No data found for {rep_name}")
        return

    st.markdown("### 💰 Revenue Progress")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric("Quota", value=f"${metrics['quota']/1000:.0f}K" if metrics['quota'] < 1_000_000 else f"${metrics['quota']/1_000_000:.1f}M")
    with col2:
        st.metric("Orders Shipped", value=f"${metrics['orders']/1000:.0f}K" if metrics['orders'] < 1_000_000 else f"${metrics['orders']/1_000_000:.1f}M",
                  help="Invoiced and shipped orders from NetSuite")
    with col3:
        st.metric("Expect/Commit", value=f"${metrics['expect_commit']/1000:.0f}K" if metrics['expect_commit'] < 1_000_000 else f"${metrics['expect_commit']/1_000_000:.1f}M",
                  help="HubSpot deals likely to close this quarter")
    with col4:
        st.metric("Pending Approval", value=f"${metrics['pending_approval']/1000:.0f}K" if metrics['pending_approval'] < 1_000_000 else f"${metrics['pending_approval']/1_000_000:.1f}M",
                  help="Sales orders awaiting approval")
    with col5:
        st.metric("Pending Fulfillment", value=f"${metrics['pending_fulfillment']/1000:.0f}K" if metrics['pending_fulfillment'] < 1_000_000 else f"${metrics['pending_fulfillment']/1_000_000:.1f}M",
                  help="Sales orders awaiting shipment (Q4 only)")
    with col6:
        st.metric("Gap to Goal", value=f"${metrics['gap']/1000:.0f}K" if abs(metrics['gap']) < 1_000_000 else f"${metrics['gap']/1_000_000:.1f}M",
                  delta=f"${-metrics['gap']/1000:.0f}K" if metrics['gap'] < 0 else None, delta_color="inverse")

    st.markdown("### 📈 Progress to Quota")
    progress = min(metrics['attainment_pct'] / 100, 1.0)
    st.progress(progress)
    # Better formatted chips
    render_total_progress(metrics)
    st.caption(f"Attainment: {metrics['attainment_pct']:.1f}% | Potential with Best Case/Opp: {metrics['potential_attainment']:.1f}%")

    if metrics['pending_fulfillment_no_date'] > 0:
        st.warning(
            f"⚠️ **Pending Fulfillment - No Ship Date:** ${metrics['pending_fulfillment_no_date']:,.0f} "
            f"(Not included in totals - needs ship date to count toward Q4)"
        )

    if metrics['pending_approval'] > 0 or metrics['pending_fulfillment'] > 0:
        st.markdown("### 📦 Sales Order Pipeline")
        so_col1, so_col2 = st.columns(2)
        with so_col1:
            st.metric("⏳ Pending Approval", value=f"${metrics['pending_approval']:,.0f}", help="Sales orders awaiting approval")
        with so_col2:
            st.metric("📤 Pending Fulfillment", value=f"${metrics['pending_fulfillment']:,.0f}", help="Sales orders pending fulfillment or billing")

    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(create_gap_chart(metrics, f"{rep_name}'s Progress to Goal"), use_container_width=True)
    with col2:
        st.plotly_chart(create_status_breakdown_chart(deals_df, rep_name), use_container_width=True)

    st.markdown("### 🔄 Pipeline Analysis")
    st.plotly_chart(create_pipeline_breakdown_chart(deals_df, rep_name), use_container_width=True)

    st.markdown("### 📅 Deal Close Timeline")
    timeline_chart = create_deals_timeline(deals_df, rep_name)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)

    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Breakdown")
        col1, col2 = st.columns(2)
        with col1:
            invoice_chart = create_invoice_status_chart(invoices_df, rep_name)
            if invoice_chart:
                st.plotly_chart(invoice_chart, use_container_width=True)
        with col2:
            customer_table = create_customer_invoice_table(invoices_df, rep_name)
            if not customer_table.empty:
                st.markdown("**Invoice Amounts by Customer**")
                for c in customer_table.columns:
                    if c != 'Customer':
                        customer_table[c] = customer_table[c].apply(lambda x: f"${x:,.0f}")
                st.dataframe(customer_table, use_container_width=True, hide_index=True)
            else:
                st.info("No invoice data available for this rep")

    st.markdown("### 📋 Deal Details")
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

    filtered_deals = metrics['deals'][metrics['deals']['Status'].isin(status_filter)]
    if pipeline_filter is not None:
        filtered_deals = filtered_deals[filtered_deals['Pipeline'].isin(pipeline_filter)]

    if not filtered_deals.empty:
        display_deals = filtered_deals[['Deal Name', 'Close Date', 'Amount', 'Status', 'Pipeline']].copy()
        display_deals['Amount'] = display_deals['Amount'].apply(lambda x: f"${x:,.0f}")
        display_deals['Close Date'] = pd.to_datetime(display_deals['Close Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        st.dataframe(display_deals, use_container_width=True, hide_index=True)
    else:
        st.info("No deals match the selected filters.")

# ========
# Main App
# ========
def main():
    # Left sidebar above already shows logo + debug toggle

    st.markdown("### 🎯 Dashboard Navigation")
    view_mode = st.radio("Select View:", ["Team Overview", "Individual Rep"], label_visibility="collapsed")
    st.markdown("---")

    current_time = datetime.now()
    st.caption(f"Last updated: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
    st.caption("Dashboard refreshes every hour")

    if st.button("🔄 Refresh Data Now"):
        st.cache_data.clear()
        st.rerun()

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

    if view_mode == "Team Overview":
        display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df)
    else:
        rep_name = st.selectbox("Select Rep:", options=dashboard_df['Rep Name'].tolist())
        display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df)

if __name__ == "__main__":
    main()
