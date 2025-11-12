"""
Sales Forecasting Dashboard - Enhanced Version with Change Tracking
Reads from Google Sheets and displays gap-to-goal analysis with interactive visualizations
Includes lead time logic for Q4/Q1 fulfillment determination and detailed order drill-downs
NEW: Invoices section, change detection, and day-over-day audit snapshot
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import plotly.io as pio
from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
from datetime import datetime, timedelta
import time
import base64
import numpy as np
import claude_insights

# Configure Plotly for dark mode compatibility
pio.templates.default = "plotly"  # Use default template that adapts to theme

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
    
    /* Responsive metrics - prevent truncation on small screens */
    [data-testid="stMetricValue"] {
        font-size: 1.5rem !important;
        overflow: visible !important;
        text-overflow: clip !important;
        white-space: nowrap !important;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 0.875rem !important;
        overflow: visible !important;
        text-overflow: clip !important;
        white-space: normal !important;
        line-height: 1.2 !important;
    }
    
    [data-testid="stMetricDelta"] {
        font-size: 0.875rem !important;
        white-space: normal !important;
    }
    
    /* Make metric containers not overflow */
    [data-testid="stMetric"] {
        overflow: visible !important;
        min-width: 140px !important;
    }
    
    /* Ensure columns wrap on small screens */
    [data-testid="column"] > div {
        overflow: visible !important;
    }
    
    /* For very small screens, stack metrics vertically */
    @media (max-width: 768px) {
        [data-testid="stMetricValue"] {
            font-size: 1.2rem !important;
        }
        [data-testid="stMetric"] {
            margin-bottom: 1rem !important;
        }
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
    
    /* Change tracking styles */
    .change-positive {
        color: #28a745;
        font-weight: bold;
    }
    
    .change-negative {
        color: #dc3545;
        font-weight: bold;
    }
    
    .change-neutral {
        color: #6c757d;
        font-weight: bold;
    }
    
    .audit-section {
        background: rgba(240, 242, 246, 0.5);
        padding: 20px;
        border-radius: 10px;
        margin: 15px 0;
        border: 1px solid rgba(128, 128, 128, 0.2);
    }
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v43_invoices_change_tracking"

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
            range=f"'{sheet_name}'!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame()
        
        # Convert to DataFrame
        df = pd.DataFrame(values[1:], columns=values[0])
        
        return df
        
    except Exception as e:
        st.error(f"Error loading data from '{sheet_name}': {str(e)}")
        return pd.DataFrame()

def load_all_data():
    """
    Load all required data from different sheets
    Returns: deals_df, dashboard_df, invoices_df, sales_orders_df
    """
    # Load deals data (All Reps All Pipelines)
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:Z")
    
    # Load dashboard info
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:Z")
    
    # Load NetSuite Invoices
    invoices_df = load_google_sheets_data("NS Invoices", "A:Z")
    
    # Load NetSuite Sales Orders
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:Z")
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def store_snapshot(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """
    Store a snapshot of current data for change tracking
    """
    snapshot = {
        'timestamp': datetime.now(),
        'deals': deals_df.copy() if not deals_df.empty else pd.DataFrame(),
        'dashboard': dashboard_df.copy() if not dashboard_df.empty else pd.DataFrame(),
        'invoices': invoices_df.copy() if not invoices_df.empty else pd.DataFrame(),
        'sales_orders': sales_orders_df.copy() if not sales_orders_df.empty else pd.DataFrame()
    }
    
    # Store in session state
    if 'previous_snapshot' not in st.session_state:
        st.session_state.previous_snapshot = snapshot
    else:
        # Move current to previous
        st.session_state.previous_snapshot = st.session_state.current_snapshot
    
    st.session_state.current_snapshot = snapshot

def detect_changes(current, previous):
    """
    Detect changes between current and previous snapshots
    Returns a dictionary of changes
    """
    changes = {
        'new_invoices': [],
        'new_sales_orders': [],
        'updated_deals': [],
        'rep_changes': {}
    }
    
    if previous is None:
        return changes
    
    try:
        # Detect new invoices
        if not current['invoices'].empty and not previous['invoices'].empty:
            if 'Document Number' in current['invoices'].columns:
                current_invoices = set(current['invoices']['Document Number'].dropna())
                previous_invoices = set(previous['invoices']['Document Number'].dropna())
                new_invoices = current_invoices - previous_invoices
                changes['new_invoices'] = list(new_invoices)
        
        # Detect new sales orders
        if not current['sales_orders'].empty and not previous['sales_orders'].empty:
            if 'Document Number' in current['sales_orders'].columns:
                current_orders = set(current['sales_orders']['Document Number'].dropna())
                previous_orders = set(previous['sales_orders']['Document Number'].dropna())
                new_orders = current_orders - previous_orders
                changes['new_sales_orders'] = list(new_orders)
        
        # Detect rep-level changes in forecasts
        if not current['dashboard'].empty and not previous['dashboard'].empty:
            if 'Rep Name' in current['dashboard'].columns:
                for rep in current['dashboard']['Rep Name'].unique():
                    current_rep = current['dashboard'][current['dashboard']['Rep Name'] == rep]
                    previous_rep = previous['dashboard'][previous['dashboard']['Rep Name'] == rep]
                    
                    if not previous_rep.empty:
                        rep_change = {}
                        
                        # Check for changes in key metrics
                        if 'Q4 Goal' in current_rep.columns:
                            current_val = pd.to_numeric(current_rep['Q4 Goal'].iloc[0], errors='coerce')
                            previous_val = pd.to_numeric(previous_rep['Q4 Goal'].iloc[0], errors='coerce')
                            if not pd.isna(current_val) and not pd.isna(previous_val):
                                if current_val != previous_val:
                                    rep_change['goal_change'] = current_val - previous_val
                        
                        if 'NetSuite Actual' in current_rep.columns:
                            current_val = pd.to_numeric(current_rep['NetSuite Actual'].iloc[0], errors='coerce')
                            previous_val = pd.to_numeric(previous_rep['NetSuite Actual'].iloc[0], errors='coerce')
                            if not pd.isna(current_val) and not pd.isna(previous_val):
                                if current_val != previous_val:
                                    rep_change['actual_change'] = current_val - previous_val
                        
                        if rep_change:
                            changes['rep_changes'][rep] = rep_change
    
    except Exception as e:
        st.error(f"Error detecting changes: {str(e)}")
    
    return changes

def show_change_dialog(changes):
    """
    Display a dialog showing what changed since last refresh
    """
    if not any([changes['new_invoices'], changes['new_sales_orders'], changes['rep_changes']]):
        st.info("ℹ️ No changes detected since last refresh")
        return
    
    st.markdown("""
    <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 padding: 20px; border-radius: 10px; color: white; margin: 15px 0;'>
        <h3 style='color: white; margin: 0 0 10px 0;'>🔄 Changes Detected!</h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if changes['new_invoices']:
            st.metric("New Invoices", len(changes['new_invoices']))
            with st.expander("View New Invoices"):
                for inv in changes['new_invoices'][:10]:  # Show first 10
                    st.write(f"• {inv}")
                if len(changes['new_invoices']) > 10:
                    st.caption(f"...and {len(changes['new_invoices']) - 10} more")
    
    with col2:
        if changes['new_sales_orders']:
            st.metric("New Sales Orders", len(changes['new_sales_orders']))
            with st.expander("View New Sales Orders"):
                for so in changes['new_sales_orders'][:10]:
                    st.write(f"• {so}")
                if len(changes['new_sales_orders']) > 10:
                    st.caption(f"...and {len(changes['new_sales_orders']) - 10} more")
    
    with col3:
        if changes['rep_changes']:
            st.metric("Reps with Changes", len(changes['rep_changes']))
            with st.expander("View Rep Changes"):
                for rep, change in changes['rep_changes'].items():
                    st.write(f"**{rep}:**")
                    if 'actual_change' in change:
                        delta = change['actual_change']
                        color = "green" if delta > 0 else "red"
                        st.markdown(f"- Actual: <span style='color:{color}'>${delta:,.0f}</span>", unsafe_allow_html=True)
                    if 'goal_change' in change:
                        st.markdown(f"- Goal: ${change['goal_change']:,.0f}")

def create_dod_audit_section(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """
    Create a day-over-day audit section showing changes
    """
    st.markdown("### 📊 Day-Over-Day Audit Snapshot")
    st.caption("Track changes in key metrics to audit data quality")
    
    # Get previous snapshot if it exists
    if 'previous_snapshot' in st.session_state and st.session_state.previous_snapshot:
        previous = st.session_state.previous_snapshot
        
        # Calculate time difference
        time_diff = datetime.now() - previous['timestamp']
        hours_ago = time_diff.total_seconds() / 3600
        
        st.markdown(f"""
        <div class='audit-section'>
            <p><strong>Previous Snapshot:</strong> {previous['timestamp'].strftime('%Y-%m-%d %H:%M:%S')} 
            ({hours_ago:.1f} hours ago)</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Team-level changes
        st.markdown("#### 👥 Team Overview")
        team_col1, team_col2, team_col3, team_col4 = st.columns(4)
        
        with team_col1:
            current_invoices = len(invoices_df) if not invoices_df.empty else 0
            previous_invoices = len(previous['invoices']) if not previous['invoices'].empty else 0
            delta_invoices = current_invoices - previous_invoices
            st.metric("Total Invoices", current_invoices, delta=delta_invoices)
        
        with team_col2:
            current_orders = len(sales_orders_df) if not sales_orders_df.empty else 0
            previous_orders = len(previous['sales_orders']) if not previous['sales_orders'].empty else 0
            delta_orders = current_orders - previous_orders
            st.metric("Total Sales Orders", current_orders, delta=delta_orders)
        
        with team_col3:
            current_deals = len(deals_df) if not deals_df.empty else 0
            previous_deals = len(previous['deals']) if not previous['deals'].empty else 0
            delta_deals = current_deals - previous_deals
            st.metric("Total Deals", current_deals, delta=delta_deals)
        
        with team_col4:
            # Calculate total invoice amount change
            if not invoices_df.empty and 'Amount' in invoices_df.columns:
                current_inv_total = pd.to_numeric(invoices_df['Amount'], errors='coerce').sum()
            else:
                current_inv_total = 0
            
            if not previous['invoices'].empty and 'Amount' in previous['invoices'].columns:
                previous_inv_total = pd.to_numeric(previous['invoices']['Amount'], errors='coerce').sum()
            else:
                previous_inv_total = 0
            
            delta_inv_amount = current_inv_total - previous_inv_total
            st.metric("Invoice Amount", f"${current_inv_total:,.0f}", delta=f"${delta_inv_amount:,.0f}")
        
        # Rep-level changes
        st.markdown("#### 👤 Rep-Level Changes")
        
        if not dashboard_df.empty and not previous['dashboard'].empty:
            rep_comparison = []
            
            for rep in dashboard_df['Rep Name'].unique():
                current_rep = dashboard_df[dashboard_df['Rep Name'] == rep]
                previous_rep = previous['dashboard'][previous['dashboard']['Rep Name'] == rep]
                
                if not previous_rep.empty:
                    rep_data = {'Rep': rep}
                    
                    # NetSuite Actual change
                    if 'NetSuite Actual' in current_rep.columns:
                        current_val = pd.to_numeric(current_rep['NetSuite Actual'].iloc[0], errors='coerce')
                        previous_val = pd.to_numeric(previous_rep['NetSuite Actual'].iloc[0], errors='coerce')
                        if not pd.isna(current_val) and not pd.isna(previous_val):
                            rep_data['Current Actual'] = current_val
                            rep_data['Previous Actual'] = previous_val
                            rep_data['Δ Actual'] = current_val - previous_val
                    
                    # Pending Fulfillment change
                    if 'Pending Fulfillment' in current_rep.columns:
                        current_val = pd.to_numeric(current_rep['Pending Fulfillment'].iloc[0], errors='coerce')
                        previous_val = pd.to_numeric(previous_rep['Pending Fulfillment'].iloc[0], errors='coerce')
                        if not pd.isna(current_val) and not pd.isna(previous_val):
                            rep_data['Current PF'] = current_val
                            rep_data['Previous PF'] = previous_val
                            rep_data['Δ PF'] = current_val - previous_val
                    
                    if len(rep_data) > 1:  # If we have any changes
                        rep_comparison.append(rep_data)
            
            if rep_comparison:
                comparison_df = pd.DataFrame(rep_comparison)
                
                # Format for display
                if 'Δ Actual' in comparison_df.columns:
                    comparison_df = comparison_df[comparison_df['Δ Actual'] != 0]
                
                if not comparison_df.empty:
                    st.dataframe(
                        comparison_df.style.format({
                            'Current Actual': '${:,.0f}',
                            'Previous Actual': '${:,.0f}',
                            'Δ Actual': '${:,.0f}',
                            'Current PF': '${:,.0f}',
                            'Previous PF': '${:,.0f}',
                            'Δ PF': '${:,.0f}'
                        }),
                        use_container_width=True
                    )
                else:
                    st.info("No significant changes in rep metrics")
            else:
                st.info("No rep-level data available for comparison")
        
    else:
        st.info("📸 No previous snapshot available. Changes will be tracked after the next refresh.")

def parse_date(date_str):
    """
    Parse date from string, handling various formats
    Returns datetime or None
    """
    if pd.isna(date_str) or str(date_str).strip() == '':
        return None
    
    try:
        # Try ISO format first
        return pd.to_datetime(date_str)
    except:
        try:
            # Try common date formats
            for fmt in ['%m/%d/%Y', '%Y-%m-%d', '%d-%b-%y', '%m-%d-%Y']:
                try:
                    return datetime.strptime(str(date_str), fmt)
                except:
                    continue
        except:
            pass
    
    return None

def calculate_q4_end():
    """Calculate Q4 2025 end date"""
    return datetime(2025, 12, 31)

def get_fulfillment_quarter(close_date_str):
    """
    Determine which quarter an order will fulfill in based on close date and lead time
    Returns 'Q4 2025' or 'Q1 2026'
    """
    if pd.isna(close_date_str) or str(close_date_str).strip() == '':
        return 'Unknown'
    
    close_date = parse_date(close_date_str)
    if close_date is None:
        return 'Unknown'
    
    # Q4 ends December 31, 2025
    q4_end = calculate_q4_end()
    
    # 21-day lead time threshold
    lead_time_days = 21
    latest_close_for_q4 = q4_end - timedelta(days=lead_time_days)
    
    if close_date <= latest_close_for_q4:
        return 'Q4 2025'
    else:
        return 'Q1 2026'

def display_drill_down_section(title, total_value, details_df, section_key):
    """
    Display a drill-down section with expandable details
    """
    st.markdown(f"**{title}**")
    st.metric("", f"${total_value:,.0f}")
    
    if not details_df.empty:
        with st.expander(f"View Details ({len(details_df)} items)"):
            # Select relevant columns for display
            display_columns = []
            possible_columns = [
                'Document Number', 'Sales Order #', 'Deal Name', 'Account Name',
                'Amount', 'Close Date', 'Expected Fulfillment Date', 
                'Status', 'Sales Rep', 'Deal Owner', 'Fulfillment Quarter'
            ]
            
            for col in possible_columns:
                if col in details_df.columns:
                    display_columns.append(col)
            
            if display_columns:
                display_df = details_df[display_columns].copy()
                
                # Format currency columns
                if 'Amount' in display_df.columns:
                    display_df['Amount'] = pd.to_numeric(display_df['Amount'], errors='coerce')
                    display_df['Amount'] = display_df['Amount'].apply(lambda x: f"${x:,.0f}" if not pd.isna(x) else "")
                
                st.dataframe(display_df, use_container_width=True, hide_index=True)
            else:
                st.write(details_df)
    else:
        st.caption("No items in this category")

def display_invoices_drill_down(invoices_df, rep_name=None):
    """
    Display invoices with drill-down capability, similar to sales orders
    """
    st.markdown("### 💰 Invoices Detail")
    st.caption("Completed and billed orders from NetSuite")
    
    if invoices_df.empty:
        st.info("No invoice data available")
        return
    
    # Filter by rep if specified
    if rep_name and 'Sales Rep' in invoices_df.columns:
        filtered_invoices = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
    else:
        filtered_invoices = invoices_df.copy()
    
    if filtered_invoices.empty:
        st.info(f"No invoices found{' for ' + rep_name if rep_name else ''}")
        return
    
    # Calculate totals
    if 'Amount' in filtered_invoices.columns:
        filtered_invoices['Amount_Numeric'] = pd.to_numeric(filtered_invoices['Amount'], errors='coerce')
        total_invoiced = filtered_invoices['Amount_Numeric'].sum()
    else:
        total_invoiced = 0
    
    # Display summary metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Invoices", len(filtered_invoices))
    
    with col2:
        st.metric("Total Amount", f"${total_invoiced:,.0f}")
    
    with col3:
        if len(filtered_invoices) > 0:
            avg_invoice = total_invoiced / len(filtered_invoices)
            st.metric("Avg Invoice", f"${avg_invoice:,.0f}")
    
    # Display invoices table
    with st.expander("📋 View All Invoices", expanded=False):
        display_columns = []
        possible_columns = [
            'Document Number', 'Transaction Date', 'Account Name', 'Customer',
            'Amount', 'Status', 'Sales Rep', 'Sales Order #', 'Terms'
        ]
        
        for col in possible_columns:
            if col in filtered_invoices.columns:
                display_columns.append(col)
        
        if display_columns:
            display_df = filtered_invoices[display_columns].copy()
            
            # Format currency
            if 'Amount' in display_df.columns:
                display_df['Amount'] = display_df['Amount_Numeric'].apply(
                    lambda x: f"${x:,.0f}" if not pd.isna(x) else ""
                )
            
            # Remove the numeric helper column if it exists
            if 'Amount_Numeric' in display_df.columns:
                display_df = display_df.drop('Amount_Numeric', axis=1)
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        else:
            st.dataframe(filtered_invoices, use_container_width=True, hide_index=True)

def calculate_rep_metrics(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    """
    Calculate all metrics for a specific rep with enhanced lead time logic
    """
    metrics = {}
    
    # Get rep's goal from dashboard
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    if not rep_info.empty:
        metrics['goal'] = pd.to_numeric(rep_info['Q4 Goal'].iloc[0], errors='coerce')
        metrics['orders'] = pd.to_numeric(rep_info['NetSuite Actual'].iloc[0], errors='coerce')
        
        # Get Pending Fulfillment values
        if 'Pending Fulfillment' in rep_info.columns:
            metrics['pending_fulfillment'] = pd.to_numeric(rep_info['Pending Fulfillment'].iloc[0], errors='coerce')
        else:
            metrics['pending_fulfillment'] = 0
        
        if 'Pending Fulfillment No Date' in rep_info.columns:
            metrics['pending_fulfillment_no_date'] = pd.to_numeric(rep_info['Pending Fulfillment No Date'].iloc[0], errors='coerce')
        else:
            metrics['pending_fulfillment_no_date'] = 0
        
        # Get Pending Approval values
        if 'Pending Approval' in rep_info.columns:
            metrics['pending_approval'] = pd.to_numeric(rep_info['Pending Approval'].iloc[0], errors='coerce')
        else:
            metrics['pending_approval'] = 0
        
        if 'Pending Approval No Date' in rep_info.columns:
            metrics['pending_approval_no_date'] = pd.to_numeric(rep_info['Pending Approval No Date'].iloc[0], errors='coerce')
        else:
            metrics['pending_approval_no_date'] = 0
        
        if 'Pending Approval >2wk' in rep_info.columns:
            metrics['pending_approval_old'] = pd.to_numeric(rep_info['Pending Approval >2wk'].iloc[0], errors='coerce')
        else:
            metrics['pending_approval_old'] = 0
    else:
        # Defaults if rep not found
        metrics['goal'] = 0
        metrics['orders'] = 0
        metrics['pending_fulfillment'] = 0
        metrics['pending_fulfillment_no_date'] = 0
        metrics['pending_approval'] = 0
        metrics['pending_approval_no_date'] = 0
        metrics['pending_approval_old'] = 0
    
    # Get rep's deals from HubSpot
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    
    # Initialize HubSpot metrics
    metrics['expect_commit'] = 0
    metrics['best_case_opp'] = 0
    metrics['q1_spillover_expect_commit'] = 0
    metrics['q1_spillover_best_opp'] = 0
    metrics['q1_spillover_total'] = 0
    
    # DataFrames for drill-down
    metrics['expect_commit_deals'] = pd.DataFrame()
    metrics['best_opp_deals'] = pd.DataFrame()
    metrics['expect_commit_q1_spillover_deals'] = pd.DataFrame()
    metrics['best_opp_q1_spillover_deals'] = pd.DataFrame()
    metrics['all_q1_spillover_deals'] = pd.DataFrame()
    
    if not rep_deals.empty and 'Amount' in rep_deals.columns and 'Close Date' in rep_deals.columns:
        rep_deals['Amount_Numeric'] = pd.to_numeric(rep_deals['Amount'], errors='coerce')
        rep_deals['Fulfillment Quarter'] = rep_deals['Close Date'].apply(get_fulfillment_quarter)
        
        # Q4 2025 deals (will fulfill in Q4)
        q4_deals = rep_deals[rep_deals['Fulfillment Quarter'] == 'Q4 2025']
        
        # Q1 2026 spillover deals (close in Q4 but fulfill in Q1)
        q1_spillover_deals = rep_deals[rep_deals['Fulfillment Quarter'] == 'Q1 2026']
        
        # Calculate Q4 metrics (Expect/Commit and Best Case/Opp)
        if 'Forecast Category' in q4_deals.columns:
            expect_commit_q4 = q4_deals[q4_deals['Forecast Category'].isin(['Commit', 'Best Case'])]
            metrics['expect_commit'] = expect_commit_q4['Amount_Numeric'].sum()
            metrics['expect_commit_deals'] = expect_commit_q4.copy()
            
            best_opp_q4 = q4_deals[q4_deals['Forecast Category'].isin(['Pipeline', 'Best Case'])]
            metrics['best_case_opp'] = best_opp_q4['Amount_Numeric'].sum()
            metrics['best_opp_deals'] = best_opp_q4.copy()
        
        # Calculate Q1 spillover metrics
        if 'Forecast Category' in q1_spillover_deals.columns:
            expect_commit_q1 = q1_spillover_deals[q1_spillover_deals['Forecast Category'].isin(['Commit', 'Best Case'])]
            metrics['q1_spillover_expect_commit'] = expect_commit_q1['Amount_Numeric'].sum()
            metrics['expect_commit_q1_spillover_deals'] = expect_commit_q1.copy()
            
            best_opp_q1 = q1_spillover_deals[q1_spillover_deals['Forecast Category'].isin(['Pipeline', 'Best Case'])]
            metrics['q1_spillover_best_opp'] = best_opp_q1['Amount_Numeric'].sum()
            metrics['best_opp_q1_spillover_deals'] = best_opp_q1.copy()
            
            metrics['q1_spillover_total'] = q1_spillover_deals['Amount_Numeric'].sum()
            metrics['all_q1_spillover_deals'] = q1_spillover_deals.copy()
    
    # Get detailed sales order data for this rep
    if not sales_orders_df.empty and 'Sales Rep' in sales_orders_df.columns:
        rep_sales_orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
        
        if not rep_sales_orders.empty and 'Status' in rep_sales_orders.columns:
            # Pending Fulfillment with date
            pf_with_date = rep_sales_orders[
                (rep_sales_orders['Status'] == 'Pending Fulfillment') &
                (rep_sales_orders['Expected Fulfillment Date'].notna()) &
                (rep_sales_orders['Expected Fulfillment Date'] != '')
            ]
            metrics['pending_fulfillment_details'] = pf_with_date
            
            # Pending Fulfillment without date
            pf_no_date = rep_sales_orders[
                (rep_sales_orders['Status'] == 'Pending Fulfillment') &
                ((rep_sales_orders['Expected Fulfillment Date'].isna()) |
                 (rep_sales_orders['Expected Fulfillment Date'] == ''))
            ]
            metrics['pending_fulfillment_no_date_details'] = pf_no_date
            
            # Pending Approval with date
            pa_with_date = rep_sales_orders[
                (rep_sales_orders['Status'] == 'Pending Approval') &
                (rep_sales_orders['Expected Fulfillment Date'].notna()) &
                (rep_sales_orders['Expected Fulfillment Date'] != '')
            ]
            metrics['pending_approval_details'] = pa_with_date
            
            # Pending Approval without date
            pa_no_date = rep_sales_orders[
                (rep_sales_orders['Status'] == 'Pending Approval') &
                ((rep_sales_orders['Expected Fulfillment Date'].isna()) |
                 (rep_sales_orders['Expected Fulfillment Date'] == ''))
            ]
            metrics['pending_approval_no_date_details'] = pa_no_date
            
            # Old Pending Approval (>2 weeks)
            if 'Transaction Date' in rep_sales_orders.columns:
                two_weeks_ago = datetime.now() - timedelta(days=14)
                rep_sales_orders['Transaction Date Parsed'] = rep_sales_orders['Transaction Date'].apply(parse_date)
                
                pa_old = rep_sales_orders[
                    (rep_sales_orders['Status'] == 'Pending Approval') &
                    (rep_sales_orders['Transaction Date Parsed'].notna()) &
                    (rep_sales_orders['Transaction Date Parsed'] < two_weeks_ago)
                ]
                metrics['pending_approval_old_details'] = pa_old
    
    # Calculate gaps
    metrics['gap_to_goal'] = metrics['goal'] - metrics['orders']
    metrics['conservative_gap'] = max(0, metrics['gap_to_goal'] - metrics['pending_fulfillment'] - 
                                      metrics['pending_approval'] - metrics['expect_commit'])
    metrics['optimistic_gap'] = max(0, metrics['gap_to_goal'] - metrics['pending_fulfillment'] - 
                                    metrics['pending_approval'] - metrics['expect_commit'] - 
                                    metrics['best_case_opp'] - metrics['pending_fulfillment_no_date'] - 
                                    metrics['pending_approval_no_date'] - metrics['pending_approval_old'])
    
    return metrics

def calculate_team_metrics(deals_df, dashboard_df):
    """
    Calculate aggregate team metrics with enhanced lead time logic
    """
    metrics = {}
    
    # Sum up team goals and actuals
    if not dashboard_df.empty:
        metrics['total_goal'] = pd.to_numeric(dashboard_df['Q4 Goal'], errors='coerce').sum()
        metrics['total_orders'] = pd.to_numeric(dashboard_df['NetSuite Actual'], errors='coerce').sum()
        
        if 'Pending Fulfillment' in dashboard_df.columns:
            metrics['total_pending_fulfillment'] = pd.to_numeric(dashboard_df['Pending Fulfillment'], errors='coerce').sum()
        else:
            metrics['total_pending_fulfillment'] = 0
        
        if 'Pending Fulfillment No Date' in dashboard_df.columns:
            metrics['total_pending_fulfillment_no_date'] = pd.to_numeric(dashboard_df['Pending Fulfillment No Date'], errors='coerce').sum()
        else:
            metrics['total_pending_fulfillment_no_date'] = 0
        
        if 'Pending Approval' in dashboard_df.columns:
            metrics['total_pending_approval'] = pd.to_numeric(dashboard_df['Pending Approval'], errors='coerce').sum()
        else:
            metrics['total_pending_approval'] = 0
        
        if 'Pending Approval No Date' in dashboard_df.columns:
            metrics['total_pending_approval_no_date'] = pd.to_numeric(dashboard_df['Pending Approval No Date'], errors='coerce').sum()
        else:
            metrics['total_pending_approval_no_date'] = 0
        
        if 'Pending Approval >2wk' in dashboard_df.columns:
            metrics['total_pending_approval_old'] = pd.to_numeric(dashboard_df['Pending Approval >2wk'], errors='coerce').sum()
        else:
            metrics['total_pending_approval_old'] = 0
    else:
        metrics['total_goal'] = 0
        metrics['total_orders'] = 0
        metrics['total_pending_fulfillment'] = 0
        metrics['total_pending_fulfillment_no_date'] = 0
        metrics['total_pending_approval'] = 0
        metrics['total_pending_approval_no_date'] = 0
        metrics['total_pending_approval_old'] = 0
    
    # Calculate HubSpot pipeline by fulfillment quarter
    metrics['total_expect_commit'] = 0
    metrics['total_best_case_opp'] = 0
    metrics['total_q1_spillover_expect_commit'] = 0
    metrics['total_q1_spillover_best_opp'] = 0
    metrics['total_q1_spillover'] = 0
    
    if not deals_df.empty and 'Amount' in deals_df.columns and 'Close Date' in deals_df.columns:
        deals_df_copy = deals_df.copy()
        deals_df_copy['Amount_Numeric'] = pd.to_numeric(deals_df_copy['Amount'], errors='coerce')
        deals_df_copy['Fulfillment Quarter'] = deals_df_copy['Close Date'].apply(get_fulfillment_quarter)
        
        # Q4 2025 deals
        q4_deals = deals_df_copy[deals_df_copy['Fulfillment Quarter'] == 'Q4 2025']
        
        # Q1 2026 spillover deals
        q1_spillover_deals = deals_df_copy[deals_df_copy['Fulfillment Quarter'] == 'Q1 2026']
        
        if 'Forecast Category' in q4_deals.columns:
            expect_commit_q4 = q4_deals[q4_deals['Forecast Category'].isin(['Commit', 'Best Case'])]
            metrics['total_expect_commit'] = expect_commit_q4['Amount_Numeric'].sum()
            
            best_opp_q4 = q4_deals[q4_deals['Forecast Category'].isin(['Pipeline', 'Best Case'])]
            metrics['total_best_case_opp'] = best_opp_q4['Amount_Numeric'].sum()
        
        if 'Forecast Category' in q1_spillover_deals.columns:
            expect_commit_q1 = q1_spillover_deals[q1_spillover_deals['Forecast Category'].isin(['Commit', 'Best Case'])]
            metrics['total_q1_spillover_expect_commit'] = expect_commit_q1['Amount_Numeric'].sum()
            
            best_opp_q1 = q1_spillover_deals[q1_spillover_deals['Forecast Category'].isin(['Pipeline', 'Best Case'])]
            metrics['total_q1_spillover_best_opp'] = best_opp_q1['Amount_Numeric'].sum()
            
            metrics['total_q1_spillover'] = q1_spillover_deals['Amount_Numeric'].sum()
    
    # Calculate team gaps
    metrics['gap_to_goal'] = metrics['total_goal'] - metrics['total_orders']
    metrics['conservative_gap'] = max(0, metrics['gap_to_goal'] - metrics['total_pending_fulfillment'] - 
                                      metrics['total_pending_approval'] - metrics['total_expect_commit'])
    metrics['optimistic_gap'] = max(0, metrics['gap_to_goal'] - metrics['total_pending_fulfillment'] - 
                                    metrics['total_pending_approval'] - metrics['total_expect_commit'] - 
                                    metrics['total_best_case_opp'] - metrics['total_pending_fulfillment_no_date'] - 
                                    metrics['total_pending_approval_no_date'] - metrics['total_pending_approval_old'])
    
    return metrics

def create_gap_chart(metrics, title):
    """
    Create a waterfall chart showing gap to goal with enhanced lead time logic
    """
    # Calculate Q4 Adjusted Total (same as in progress breakdown)
    q4_adjusted = (
        metrics['orders'] +
        metrics['pending_fulfillment'] +
        metrics['pending_approval'] +
        metrics['expect_commit'] +
        metrics['pending_fulfillment_no_date'] +
        metrics['pending_approval_no_date'] +
        metrics['pending_approval_old'] +
        metrics.get('q1_spillover_expect_commit', 0)
    )
    
    # Components for waterfall
    categories = [
        'Q4 Goal',
        'Invoiced',
        'Pend. Fulfill.',
        'Pend. Approval',
        'Expect/Commit',
        'Pend. Fulf. (No Date)',
        'Pend. App. (No Date)',
        'PA >2wk',
        'Q1 Spillover',
        'Gap to Goal'
    ]
    
    values = [
        metrics['goal'],
        metrics['orders'],
        metrics['pending_fulfillment'],
        metrics['pending_approval'],
        metrics['expect_commit'],
        metrics['pending_fulfillment_no_date'],
        metrics['pending_approval_no_date'],
        metrics['pending_approval_old'],
        metrics.get('q1_spillover_expect_commit', 0),
        max(0, metrics['goal'] - q4_adjusted)
    ]
    
    # Create waterfall chart
    fig = go.Figure(go.Waterfall(
        name="Q4 Forecast",
        orientation="v",
        measure=["relative", "relative", "relative", "relative", "relative", "relative", "relative", "relative", "relative", "total"],
        x=categories,
        y=values,
        text=[f"${v:,.0f}" for v in values],
        textposition="outside",
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        decreasing={"marker": {"color": "#ef5350"}},
        increasing={"marker": {"color": "#66bb6a"}},
        totals={"marker": {"color": "#42a5f5"}}
    ))
    
    fig.update_layout(
        title=title,
        showlegend=False,
        height=500,
        template="plotly"
    )
    
    return fig

def create_status_breakdown_chart(deals_df, rep_name):
    """
    Create a pie chart showing breakdown by forecast category
    """
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name]
    
    if rep_deals.empty or 'Forecast Category' not in rep_deals.columns:
        return None
    
    # Get value by forecast category
    rep_deals_copy = rep_deals.copy()
    rep_deals_copy['Amount_Numeric'] = pd.to_numeric(rep_deals_copy['Amount'], errors='coerce')
    
    category_totals = rep_deals_copy.groupby('Forecast Category')['Amount_Numeric'].sum().reset_index()
    
    if category_totals.empty:
        return None
    
    fig = px.pie(
        category_totals,
        values='Amount_Numeric',
        names='Forecast Category',
        title=f"{rep_name} - Pipeline by Forecast Category",
        hole=0.4
    )
    
    fig.update_traces(textposition='inside', textinfo='label+percent')
    fig.update_layout(height=500)
    
    return fig

def create_pipeline_breakdown_chart(deals_df, rep_name):
    """
    Create a horizontal bar chart showing pipeline breakdown by deal stage
    """
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name]
    
    if rep_deals.empty or 'Deal Stage' not in rep_deals.columns:
        return None
    
    rep_deals_copy = rep_deals.copy()
    rep_deals_copy['Amount_Numeric'] = pd.to_numeric(rep_deals_copy['Amount'], errors='coerce')
    
    stage_totals = rep_deals_copy.groupby('Deal Stage')['Amount_Numeric'].sum().reset_index()
    stage_totals = stage_totals.sort_values('Amount_Numeric', ascending=True)
    
    if stage_totals.empty:
        return None
    
    fig = px.bar(
        stage_totals,
        x='Amount_Numeric',
        y='Deal Stage',
        orientation='h',
        title=f"{rep_name} - Pipeline by Deal Stage",
        text='Amount_Numeric'
    )
    
    fig.update_traces(texttemplate='$%{text:,.0f}', textposition='outside')
    fig.update_layout(height=400, xaxis_title="Amount ($)", yaxis_title="")
    
    return fig

def create_deals_timeline(deals_df, rep_name):
    """
    Create a timeline showing deals by expected close date
    """
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name]
    
    if rep_deals.empty or 'Close Date' not in rep_deals.columns:
        return None
    
    rep_deals_copy = rep_deals.copy()
    rep_deals_copy['Amount_Numeric'] = pd.to_numeric(rep_deals_copy['Amount'], errors='coerce')
    rep_deals_copy['Close Date Parsed'] = rep_deals_copy['Close Date'].apply(parse_date)
    
    # Filter out deals without valid dates
    rep_deals_copy = rep_deals_copy[rep_deals_copy['Close Date Parsed'].notna()]
    
    if rep_deals_copy.empty:
        return None
    
    # Group by date
    daily_totals = rep_deals_copy.groupby('Close Date Parsed').agg({
        'Amount_Numeric': 'sum',
        'Deal Name': 'count'
    }).reset_index()
    
    daily_totals = daily_totals.rename(columns={'Deal Name': 'Deal Count'})
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=daily_totals['Close Date Parsed'],
        y=daily_totals['Amount_Numeric'],
        mode='markers+lines',
        name='Deal Value',
        marker=dict(size=daily_totals['Deal Count']*5, sizemode='diameter'),
        text=[f"${v:,.0f}<br>{c} deal(s)" for v, c in zip(daily_totals['Amount_Numeric'], daily_totals['Deal Count'])],
        hovertemplate='%{text}<extra></extra>'
    ))
    
    fig.update_layout(
        title=f"{rep_name} - Deals Timeline",
        xaxis_title="Expected Close Date",
        yaxis_title="Total Deal Value ($)",
        height=400,
        template="plotly"
    )
    
    return fig

def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """
    Display team-level dashboard with enhanced lead time logic
    """
    st.markdown("## 👥 Team Overview")
    
    # Calculate team metrics
    metrics = calculate_team_metrics(deals_df, dashboard_df)
    
    # Q4 Adjusted Total (includes Q1 spillover)
    q4_adjusted = (
        metrics['total_orders'] +
        metrics['total_pending_fulfillment'] +
        metrics['total_pending_approval'] +
        metrics['total_expect_commit'] +
        metrics['total_pending_fulfillment_no_date'] +
        metrics['total_pending_approval_no_date'] +
        metrics['total_pending_approval_old'] +
        metrics.get('total_q1_spillover_expect_commit', 0)
    )
    
    # Top-level metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Q4 Goal", f"${metrics['total_goal']:,.0f}")
    
    with col2:
        st.metric("Invoiced & Shipped", f"${metrics['total_orders']:,.0f}")
    
    with col3:
        st.metric("Gap to Goal", f"${metrics['gap_to_goal']:,.0f}")
    
    with col4:
        progress_pct = (metrics['total_orders'] / metrics['total_goal'] * 100) if metrics['total_goal'] > 0 else 0
        st.metric("Progress", f"{progress_pct:.1f}%")
    
    with col5:
        st.metric("Q4 Adjusted Total", f"${q4_adjusted:,.0f}")
    
    st.markdown("---")
    
    # Change detection and audit section
    if st.checkbox("📊 Show Day-Over-Day Audit", value=False):
        create_dod_audit_section(deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    st.markdown("---")
    
    # Invoices section
    display_invoices_drill_down(invoices_df)
    
    st.markdown("---")
    
    # SECTION 1: Current State
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>📍 Section 1: Current State (NetSuite Actuals)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['total_orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['total_pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['total_pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 CONSERVATIVE FORECAST</span>
            <span class="progress-value">${metrics['total_orders'] + metrics['total_pending_fulfillment'] + metrics['total_pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Conservative Gap to Goal</span>
            <span class="progress-value">${metrics['conservative_gap']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # SECTION 2: Optimistic View
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>🌟 Section 2: Optimistic View (Full Pipeline)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['total_orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['total_pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['total_pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 HubSpot Expect/Commit</span>
            <span class="progress-value">${metrics['total_expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎲 HubSpot Best Case/Opp</span>
            <span class="progress-value">${metrics['total_best_case_opp']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (without date)</span>
            <span class="progress-value">${metrics['total_pending_fulfillment_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (without date)</span>
            <span class="progress-value">${metrics['total_pending_approval_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏱️ Pending Approval (>2 weeks old)</span>
            <span class="progress-value">${metrics['total_pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 OPTIMISTIC FORECAST</span>
            <span class="progress-value">${metrics['total_orders'] + metrics['total_pending_fulfillment'] + metrics['total_pending_approval'] + metrics['total_expect_commit'] + metrics['total_best_case_opp'] + metrics['total_pending_fulfillment_no_date'] + metrics['total_pending_approval_no_date'] + metrics['total_pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Optimistic Gap to Goal</span>
            <span class="progress-value">${metrics['optimistic_gap']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # SECTION 3: Q4 Adjusted Forecast
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>📈 Section 3: Q4 Adjusted Forecast (Includes Q1 Spillover - Review Needed)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['total_orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['total_pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['total_pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 HubSpot Expect/Commit</span>
            <span class="progress-value">${metrics['total_expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (without date)</span>
            <span class="progress-value">${metrics['total_pending_fulfillment_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (without date)</span>
            <span class="progress-value">${metrics['total_pending_approval_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏱️ Pending Approval (>2 weeks old)</span>
            <span class="progress-value">${metrics['total_pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🦘 Q1 Spillover - Expect/Commit (⚠️ Review)</span>
            <span class="progress-value">${metrics.get('total_q1_spillover_expect_commit', 0):,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 Q4 ADJUSTED TOTAL</span>
            <span class="progress-value">${q4_adjusted:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Rep-by-Rep Breakdown
    st.markdown("### 👤 Individual Rep Performance")
    
    if not dashboard_df.empty:
        for _, rep_row in dashboard_df.iterrows():
            with st.expander(f"📊 {rep_row['Rep Name']}"):
                rep_metrics = calculate_rep_metrics(
                    rep_row['Rep Name'], 
                    deals_df, 
                    dashboard_df, 
                    invoices_df, 
                    sales_orders_df
                )
                
                rep_col1, rep_col2, rep_col3, rep_col4 = st.columns(4)
                
                with rep_col1:
                    st.metric("Goal", f"${rep_metrics['goal']:,.0f}")
                
                with rep_col2:
                    st.metric("Actual", f"${rep_metrics['orders']:,.0f}")
                
                with rep_col3:
                    st.metric("Gap", f"${rep_metrics['gap_to_goal']:,.0f}")
                
                with rep_col4:
                    progress = (rep_metrics['orders'] / rep_metrics['goal'] * 100) if rep_metrics['goal'] > 0 else 0
                    st.metric("Progress", f"{progress:.1f}%")

def display_reconciliation_view(deals_df, dashboard_df, sales_orders_df):
    """
    Display reconciliation view comparing HubSpot deals with NetSuite sales orders
    """
    st.markdown("## 🔍 Reconciliation View")
    st.caption("Compare HubSpot deals with NetSuite sales orders to identify discrepancies")
    
    if deals_df.empty or sales_orders_df.empty:
        st.warning("⚠️ Need both HubSpot deals and NetSuite sales orders data for reconciliation")
        return
    
    # Create mapping tables
    st.markdown("### 📋 Sales Orders by Status")
    
    if 'Status' in sales_orders_df.columns:
        status_summary = sales_orders_df.groupby('Status').agg({
            'Document Number': 'count',
            'Amount': lambda x: pd.to_numeric(x, errors='coerce').sum()
        }).reset_index()
        
        status_summary.columns = ['Status', 'Count', 'Total Amount']
        status_summary['Total Amount'] = status_summary['Total Amount'].apply(lambda x: f"${x:,.0f}")
        
        st.dataframe(status_summary, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # Deal stage breakdown
    st.markdown("### 📊 HubSpot Deal Stages")
    
    if 'Deal Stage' in deals_df.columns:
        deals_copy = deals_df.copy()
        deals_copy['Amount_Numeric'] = pd.to_numeric(deals_copy['Amount'], errors='coerce')
        
        stage_summary = deals_copy.groupby('Deal Stage').agg({
            'Deal Name': 'count',
            'Amount_Numeric': 'sum'
        }).reset_index()
        
        stage_summary.columns = ['Deal Stage', 'Count', 'Total Amount']
        stage_summary['Total Amount'] = stage_summary['Total Amount'].apply(lambda x: f"${x:,.0f}")
        
        st.dataframe(stage_summary, use_container_width=True, hide_index=True)

def display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    """
    Display individual rep dashboard with enhanced lead time logic
    """
    st.markdown(f"## 👤 {rep_name}")
    
    # Calculate rep metrics
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    # Q4 Adjusted Total
    q4_adjusted = (
        metrics['orders'] +
        metrics['pending_fulfillment'] +
        metrics['pending_approval'] +
        metrics['expect_commit'] +
        metrics['pending_fulfillment_no_date'] +
        metrics['pending_approval_no_date'] +
        metrics['pending_approval_old'] +
        metrics.get('q1_spillover_expect_commit', 0)
    )
    
    # Top-level metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Q4 Goal", f"${metrics['goal']:,.0f}")
    
    with col2:
        st.metric("Invoiced & Shipped", f"${metrics['orders']:,.0f}")
    
    with col3:
        st.metric("Gap to Goal", f"${metrics['gap_to_goal']:,.0f}")
    
    with col4:
        progress_pct = (metrics['orders'] / metrics['goal'] * 100) if metrics['goal'] > 0 else 0
        st.metric("Progress", f"{progress_pct:.1f}%")
    
    with col5:
        st.metric("Q4 Adjusted Total", f"${q4_adjusted:,.0f}")
    
    st.markdown("---")
    
    # Invoices section for this rep
    display_invoices_drill_down(invoices_df, rep_name)
    
    st.markdown("---")
    
    # SECTION 1: Current State
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>📍 Section 1: Current State (NetSuite Actuals)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 CONSERVATIVE FORECAST</span>
            <span class="progress-value">${metrics['orders'] + metrics['pending_fulfillment'] + metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Conservative Gap to Goal</span>
            <span class="progress-value">${metrics['conservative_gap']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Drill-down for Section 1
    st.markdown("#### 🔍 Current State Details")
    
    section1_col1, section1_col2, section1_col3 = st.columns(3)
    
    with section1_col1:
        display_drill_down_section(
            "📦 Pending Fulfillment",
            metrics['pending_fulfillment'],
            metrics.get('pending_fulfillment_details', pd.DataFrame()),
            f"{rep_name}_pf"
        )
    
    with section1_col2:
        display_drill_down_section(
            "⏳ Pending Approval",
            metrics['pending_approval'],
            metrics.get('pending_approval_details', pd.DataFrame()),
            f"{rep_name}_pa"
        )
    
    st.markdown("---")
    
    # SECTION 2: Optimistic View
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>🌟 Section 2: Optimistic View (Full Pipeline)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 HubSpot Expect/Commit</span>
            <span class="progress-value">${metrics['expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎲 HubSpot Best Case/Opp</span>
            <span class="progress-value">${metrics['best_case_opp']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (without date)</span>
            <span class="progress-value">${metrics['pending_fulfillment_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (without date)</span>
            <span class="progress-value">${metrics['pending_approval_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏱️ Pending Approval (>2 weeks old)</span>
            <span class="progress-value">${metrics['pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 OPTIMISTIC FORECAST</span>
            <span class="progress-value">${metrics['orders'] + metrics['pending_fulfillment'] + metrics['pending_approval'] + metrics['expect_commit'] + metrics['best_case_opp'] + metrics['pending_fulfillment_no_date'] + metrics['pending_approval_no_date'] + metrics['pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Optimistic Gap to Goal</span>
            <span class="progress-value">${metrics['optimistic_gap']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Drill-down for Section 2
    st.markdown("#### 🔍 Pipeline Details")
    
    section2_col1, section2_col2 = st.columns(2)
    
    with section2_col1:
        display_drill_down_section(
            "🎯 Expect/Commit Deals",
            metrics['expect_commit'],
            metrics.get('expect_commit_deals', pd.DataFrame()),
            f"{rep_name}_ec"
        )
    
    with section2_col2:
        display_drill_down_section(
            "🎲 Best Case/Opp Deals",
            metrics['best_case_opp'],
            metrics.get('best_opp_deals', pd.DataFrame()),
            f"{rep_name}_bo"
        )
    
    st.markdown("#### 📋 Additional Categories")
    
    section2_col3, section2_col4, section2_col5 = st.columns(3)
    
    with section2_col3:
        display_drill_down_section(
            "📦 Pend. Fulf. (No Date)",
            metrics['pending_fulfillment_no_date'],
            metrics.get('pending_fulfillment_no_date_details', pd.DataFrame()),
            f"{rep_name}_pf_no_date"
        )
    
    with section2_col4:
        display_drill_down_section(
            "⏳ Pend. App. (No Date)",
            metrics['pending_approval_no_date'],
            metrics.get('pending_approval_no_date_details', pd.DataFrame()),
            f"{rep_name}_pa_no_date"
        )
    
    with section2_col5:
        display_drill_down_section(
            "⏱️ Old Pending Approval (>2 weeks)",
            metrics['pending_approval_old'],
            metrics.get('pending_approval_old_details', pd.DataFrame()),
            f"{rep_name}_pa_old"
        )
    
    st.markdown("---")
    
    # SECTION 3: Q4 Adjusted Forecast
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>📈 Section 3: Q4 Adjusted Forecast (Includes Q1 Spillover - Review Needed)</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Invoiced & Shipped</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (with date)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (with date)</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 HubSpot Expect/Commit</span>
            <span class="progress-value">${metrics['expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 Pending Fulfillment (without date)</span>
            <span class="progress-value">${metrics['pending_fulfillment_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Pending Approval (without date)</span>
            <span class="progress-value">${metrics['pending_approval_no_date']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏱️ Pending Approval (>2 weeks old)</span>
            <span class="progress-value">${metrics['pending_approval_old']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🦘 Q1 Spillover - Expect/Commit (⚠️ Review)</span>
            <span class="progress-value">${metrics.get('q1_spillover_expect_commit', 0):,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📈 Q4 ADJUSTED TOTAL</span>
            <span class="progress-value">${q4_adjusted:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Drill-down for Section 3 (Q1 Spillover)
    st.markdown("#### 🦘 Q1 2026 Spillover Details")
    st.caption("⚠️ These deals close in Q4 2025 but will ship in Q1 2026 due to lead times")
    
    spillover_col1, spillover_col2, spillover_col3 = st.columns(3)
    
    with spillover_col1:
        display_drill_down_section(
            "🎯 Expect/Commit (Q1 Spillover)",
            metrics.get('q1_spillover_expect_commit', 0),
            metrics.get('expect_commit_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_ec_q1"
        )
    
    with spillover_col2:
        display_drill_down_section(
            "🎲 Best Case/Opp (Q1 Spillover)",
            metrics.get('q1_spillover_best_opp', 0),
            metrics.get('best_opp_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_bo_q1"
        )
    
    with spillover_col3:
        display_drill_down_section(
            "📦 All Q1 2026 Spillover",
            metrics.get('q1_spillover_total', 0),
            metrics.get('all_q1_spillover_deals', pd.DataFrame()),
            f"{rep_name}_all_q1"
        )
    
    st.markdown("---")
    
    # Charts
    st.markdown("### 📊 Visual Analysis")
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
        st.image("calyx_logo.png", width=200)
        
        st.markdown("---")
        
        st.markdown("### 🎯 Dashboard Navigation")
        view_mode = st.radio(
            "Select View:",
            ["Team Overview", "Individual Rep", "Reconciliation", "AI Insights"],
            label_visibility="collapsed"
        )
        
        st.markdown("---")
        
        # Last updated and refresh button (always visible)
        current_time = datetime.now()
        st.caption(f"Last updated: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        st.caption("Dashboard refreshes every hour")
        
        if st.button("🔄 Refresh Data Now"):
            # Store snapshot before clearing cache
            if 'current_snapshot' in st.session_state:
                st.session_state.previous_snapshot = st.session_state.current_snapshot
            
            st.cache_data.clear()
            st.rerun()
        
        st.markdown("---")
        
        # Sync Status - collapsed by default, for Xander
        with st.expander("🔧 Sync Status (for Xander)"):
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
    
    # Load data
    with st.spinner("Loading data from Google Sheets..."):
        deals_df, dashboard_df, invoices_df, sales_orders_df = load_all_data()
    
    # Store snapshot for change tracking
    store_snapshot(deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    # Show change detection dialog if there's a previous snapshot
    if 'previous_snapshot' in st.session_state and st.session_state.previous_snapshot:
        with st.expander("🔄 View Changes Since Last Refresh", expanded=False):
            changes = detect_changes(st.session_state.current_snapshot, st.session_state.previous_snapshot)
            show_change_dialog(changes)
    
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
    elif view_mode == "AI Insights":
        # Calculate team metrics for Claude to use
        team_metrics = calculate_team_metrics(deals_df, dashboard_df)
        claude_insights.display_insights_dashboard(deals_df, dashboard_df, team_metrics)
    else:  # Reconciliation view
        display_reconciliation_view(deals_df, dashboard_df, sales_orders_df)

if __name__ == "__main__":
    main()
