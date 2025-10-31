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
CACHE_VERSION = "v2"

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
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:H")
    
    # Load dashboard info (rep quotas and orders)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C")
    
    # Clean and process data
    if not deals_df.empty and len(deals_df.columns) >= 8:
        # The sheet has columns A-H, so let's use proper column names
        # Based on the Google Apps Script: A=?, B=Deal Name, C=?, D=Close Date, E=Deal Owner, F=Amount, G=Status, H=Pipeline
        
        # Get column names from first row (if they exist) or use indices
        if len(deals_df) > 0:
            # Standardize column names based on position
            col_names = deals_df.columns.tolist()
            
            # Map columns by position if column names aren't what we expect
            deals_df = deals_df.rename(columns={
                col_names[1]: 'Deal Name',
                col_names[3]: 'Close Date',
                col_names[4]: 'Deal Owner',
                col_names[5]: 'Amount',
                col_names[6]: 'Status',
                col_names[7]: 'Pipeline'
            })
            
            # Convert amount to numeric
            deals_df['Amount'] = pd.to_numeric(deals_df['Amount'], errors='coerce').fillna(0)
            
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
            
            # Convert to numeric, keeping original values if conversion fails
            dashboard_df['Quota'] = pd.to_numeric(dashboard_df['Quota'], errors='coerce')
            dashboard_df['NetSuite Orders'] = pd.to_numeric(dashboard_df['NetSuite Orders'], errors='coerce')
            
            # Fill NaN with 0
            dashboard_df['Quota'] = dashboard_df['Quota'].fillna(0)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].fillna(0)
            
            # Debug: Show after conversion
            st.sidebar.write("**DEBUG - After Conversion:**")
            st.sidebar.dataframe(dashboard_df)
        else:
            st.error(f"Dashboard Info sheet has wrong number of columns: {len(dashboard_df.columns)}")
    
    return deals_df, dashboard_df

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

def calculate_rep_metrics(rep_name, deals_df, dashboard_df):
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
    
    # Calculate gap
    gap = quota - expect_commit - orders
    
    # Calculate attainment
    current_forecast = expect_commit + orders
    attainment_pct = (current_forecast / quota * 100) if quota > 0 else 0
    
    # Potential attainment
    potential_attainment = ((expect_commit + best_opp + orders) / quota * 100) if quota > 0 else 0
    
    return {
        'quota': quota,
        'orders': orders,
        'expect_commit': expect_commit,
        'best_opp': best_opp,
        'gap': gap,
        'attainment_pct': attainment_pct,
        'potential_attainment': potential_attainment,
        'current_forecast': current_forecast,
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

def display_team_dashboard(deals_df, dashboard_df):
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
            label="Current Forecast",
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
    
    # Rep summary table
    st.markdown("### 👥 Rep Summary")
    
    rep_summary = []
    for rep_name in dashboard_df['Rep Name']:
        rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df)
        if rep_metrics:
            rep_summary.append({
                'Rep': rep_name,
                'Quota': f"${rep_metrics['quota']:,.0f}",
                'Forecast': f"${rep_metrics['current_forecast']:,.0f}",
                'Gap': f"${rep_metrics['gap']:,.0f}",
                'Attainment': f"{rep_metrics['attainment_pct']:.1f}%",
                'Potential': f"{rep_metrics['potential_attainment']:.1f}%"
            })
    
    rep_summary_df = pd.DataFrame(rep_summary)
    st.dataframe(rep_summary_df, use_container_width=True, hide_index=True)

def display_rep_dashboard(rep_name, deals_df, dashboard_df):
    """Display individual rep dashboard"""
    
    st.title(f"👤 {rep_name}'s Q4 2025 Forecast")
    
    # Calculate metrics
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df)
    
    if not metrics:
        st.error(f"No data found for {rep_name}")
        return
    
    # Display key metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            label="Quota",
            value=f"${metrics['quota']:,.0f}"
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
        deals_df, dashboard_df = load_all_data()
    
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
        display_team_dashboard(deals_df, dashboard_df)
    else:
        rep_name = st.selectbox(
            "Select Rep:",
            options=dashboard_df['Rep Name'].tolist()
        )
        display_rep_dashboard(rep_name, deals_df, dashboard_df)

if __name__ == "__main__":
    main()
