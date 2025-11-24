"""
Concentrate Jar Forecasting Module (replaces shipping_planning.py)
===================================================================
Creates 2026 forecasts for 4ml concentrate jars based on ALL historical glass SKU data.
Uses weighted historical analysis (2024 weighted higher than 2025 due to healthy stock levels).
Incorporates pipeline data for Nov-Mar and softens outlier impact.

Navigation: "📦 Q4 Shipping Plan" -> Will show as Concentrate Jar Forecast

Columns from Google Sheet (Concentrate Jar Forecasting tab):
A: Close Date
B: Quantity
C: Product
D: Product Name
E: Amount
F: Close Status
G: Pipeline
H: Deal Stage
I: Deal ID
J: Ticket ID
K: Line item ID
L: Company ID
M: Contact ID
N: Company Name
O: Company Owner
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Google Sheets Configuration (same as main dashboard)
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "concentrate_v3"

# =============================================================================
# DATA LOADING
# =============================================================================

@st.cache_data(ttl=CACHE_TTL)
def load_concentrate_data(version=CACHE_VERSION):
    """
    Load data from Concentrate Jar Forecasting tab in Google Sheets
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
        
        # Load from Concentrate Jar Forecasting tab - columns A:O
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Concentrate Jar Forecasting!A:O"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning("⚠️ No data found in 'Concentrate Jar Forecasting' tab")
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
        st.error(f"❌ Error loading Concentrate Jar Forecasting data: {str(e)}")
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


def process_concentrate_data(df):
    """
    Process the concentrate jar data with known column mappings.
    """
    if df.empty:
        return df
    
    # Standardize column names (handle variations)
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if 'close date' in col_lower or col_lower == 'close date':
            col_mapping[col] = 'Close Date'
        elif col_lower == 'quantity':
            col_mapping[col] = 'Quantity'
        elif col_lower == 'product name':
            col_mapping[col] = 'Product Name'
        elif col_lower == 'product' and 'product name' not in col_lower:
            col_mapping[col] = 'Product'
        elif col_lower == 'amount':
            col_mapping[col] = 'Amount'
        elif 'close status' in col_lower:
            col_mapping[col] = 'Close Status'
        elif col_lower == 'pipeline':
            col_mapping[col] = 'Pipeline'
        elif 'deal stage' in col_lower:
            col_mapping[col] = 'Deal Stage'
        elif 'company name' in col_lower:
            col_mapping[col] = 'Company Name'
        elif 'company owner' in col_lower:
            col_mapping[col] = 'Company Owner'
    
    df = df.rename(columns=col_mapping)
    
    # Parse Close Date
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    df = df[df['Close Date'].notna()].copy()
    
    # Add time-based columns
    df['Year'] = df['Close Date'].dt.year
    df['Month'] = df['Close Date'].dt.month
    df['YearMonth'] = df['Close Date'].dt.to_period('M')
    df['Quarter'] = df['Close Date'].dt.quarter
    df['MonthName'] = df['Close Date'].dt.strftime('%b')
    df['MonthLabel'] = df['Close Date'].dt.strftime('%b %Y')
    
    # Clean numeric columns
    df['Quantity'] = df['Quantity'].apply(clean_numeric)
    df['Amount'] = df['Amount'].apply(clean_numeric)
    
    return df


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

def calculate_weighted_monthly_averages(df, weight_2024=0.6, weight_2025=0.4):
    """
    Calculate weighted monthly averages using historical data.
    
    2024 is weighted higher (default 60%) because stock was healthy.
    2025 is weighted lower (default 40%) due to stock constraints.
    """
    if df.empty or 'Year' not in df.columns:
        return pd.DataFrame()
    
    # Aggregate by Year and Month
    monthly_data = df.groupby(['Year', 'Month']).agg({
        'Quantity': 'sum',
        'Amount': 'sum'
    }).reset_index()
    
    # Separate years
    data_2023 = monthly_data[monthly_data['Year'] == 2023].set_index('Month')
    data_2024 = monthly_data[monthly_data['Year'] == 2024].set_index('Month')
    data_2025 = monthly_data[monthly_data['Year'] == 2025].set_index('Month')
    
    weighted_avgs = []
    
    for month in range(1, 13):
        qty_2023 = data_2023.loc[month, 'Quantity'] if month in data_2023.index else 0
        qty_2024 = data_2024.loc[month, 'Quantity'] if month in data_2024.index else 0
        qty_2025 = data_2025.loc[month, 'Quantity'] if month in data_2025.index else 0
        amt_2023 = data_2023.loc[month, 'Amount'] if month in data_2023.index else 0
        amt_2024 = data_2024.loc[month, 'Amount'] if month in data_2024.index else 0
        amt_2025 = data_2025.loc[month, 'Amount'] if month in data_2025.index else 0
        
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


def identify_pipeline_data(df):
    """
    Identify pipeline/pending deals for Nov 2025 - Mar 2026.
    """
    if df.empty:
        return pd.DataFrame()
    
    # Define pipeline date range
    pipeline_start = datetime(2025, 11, 1)
    pipeline_end = datetime(2026, 3, 31)
    
    # Filter for pipeline period
    pipeline_df = df[
        (df['Close Date'] >= pipeline_start) & 
        (df['Close Date'] <= pipeline_end)
    ].copy()
    
    return pipeline_df


def generate_2026_forecast(df, weight_2024=0.6, weight_2025=0.4):
    """
    Generate month-by-month and quarter-by-quarter 2026 forecast.
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    # Separate historical (before Nov 2025) and pipeline data
    cutoff_date = datetime(2025, 11, 1)
    historical_df = df[df['Close Date'] < cutoff_date].copy()
    pipeline_df = identify_pipeline_data(df)
    
    # Soften outliers in historical data
    if not historical_df.empty and historical_df['Quantity'].sum() > 0:
        historical_df['Quantity'] = soften_outliers(historical_df['Quantity'])
        historical_df['Amount'] = soften_outliers(historical_df['Amount'])
    
    # Calculate weighted monthly baselines
    monthly_baselines = calculate_weighted_monthly_averages(historical_df, weight_2024, weight_2025)
    
    if monthly_baselines.empty:
        st.warning("⚠️ Not enough historical data to generate forecast")
        return pd.DataFrame(), pd.DataFrame(), monthly_baselines
    
    # Aggregate pipeline by month
    pipeline_monthly = pd.DataFrame()
    if not pipeline_df.empty:
        pipeline_df['ForecastMonth'] = pipeline_df['Month']
        # Adjust for year (Nov/Dec 2025 -> month 11/12, Jan-Mar 2026 -> month 1-3)
        pipeline_monthly = pipeline_df.groupby('Month').agg({
            'Quantity': 'sum',
            'Amount': 'sum'
        }).reset_index()
    
    # Generate 2026 monthly forecast
    forecast_2026 = []
    
    for month in range(1, 13):
        baseline_row = monthly_baselines[monthly_baselines['Month'] == month]
        if baseline_row.empty:
            continue
            
        baseline = baseline_row.iloc[0]
        
        forecasted_qty = baseline['Weighted_Quantity']
        forecasted_amt = baseline['Weighted_Amount']
        
        # For Q1, blend with pipeline data if available
        pipeline_qty = 0
        pipeline_amt = 0
        if not pipeline_monthly.empty and month in [1, 2, 3]:
            if month in pipeline_monthly['Month'].values:
                pipeline_row = pipeline_monthly[pipeline_monthly['Month'] == month].iloc[0]
                pipeline_qty = pipeline_row['Quantity']
                pipeline_amt = pipeline_row['Amount']
                
                # Blend: 60% historical baseline + 40% pipeline
                forecasted_qty = (forecasted_qty * 0.6) + (pipeline_qty * 0.4)
                forecasted_amt = (forecasted_amt * 0.6) + (pipeline_amt * 0.4)
        
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
            'Pipeline_Qty': int(pipeline_qty),
            'Pipeline_Amt': round(pipeline_amt, 2),
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
        quarterly_forecast = monthly_forecast.groupby(['QuarterNum', 'Quarter']).agg({
            'Forecasted_Quantity': 'sum',
            'Forecasted_Amount': 'sum',
            'Qty_Low': 'sum',
            'Qty_High': 'sum',
            'Amt_Low': 'sum',
            'Amt_High': 'sum',
            'Pipeline_Qty': 'sum',
            'Pipeline_Amt': 'sum'
        }).reset_index()
        
        quarterly_forecast = quarterly_forecast.sort_values('QuarterNum')
    
    return monthly_forecast, quarterly_forecast, monthly_baselines


# =============================================================================
# VISUALIZATION
# =============================================================================

def create_historical_trend_chart(df):
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
            y=year_data['Quantity'],
            name=str(int(year)),
            marker=dict(
                color=color_map.get(year, '#6366f1'),
                line=dict(color='rgba(255,255,255,0.3)', width=1)
            ),
            text=year_data['Quantity'].apply(lambda x: f"{x/1000:.0f}K" if x >= 1000 else f"{int(x)}"),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Quantity: %{y:,.0f}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text='📊 Historical Concentrate Jar Demand by Month',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Quantity (Units)',
            gridcolor='rgba(128,128,128,0.2)'
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


def create_forecast_chart(monthly_forecast):
    """
    Create a forecast chart with confidence bands for 2026.
    """
    if monthly_forecast.empty:
        return None
    
    fig = go.Figure()
    
    # Add confidence band (shaded area)
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'].tolist() + monthly_forecast['MonthShort'].tolist()[::-1],
        y=monthly_forecast['Qty_High'].tolist() + monthly_forecast['Qty_Low'].tolist()[::-1],
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
        y=monthly_forecast['Forecasted_Quantity'],
        mode='lines+markers',
        name='2026 Forecast',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10, color='#10b981', line=dict(color='white', width=2)),
        text=monthly_forecast['Forecasted_Quantity'].apply(lambda x: f"{x/1000:.0f}K" if x >= 1000 else str(int(x))),
        textposition='top center',
        hovertemplate='<b>%{x} 2026</b><br>Forecast: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add 2024 historical reference line
    if 'Historical_Qty_2024' in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast['Historical_Qty_2024'],
            mode='lines+markers',
            name='2024 Actual',
            line=dict(color='#8b5cf6', width=2, dash='dot'),
            marker=dict(size=6, color='#8b5cf6'),
            hovertemplate='<b>%{x} 2024</b><br>Actual: %{y:,.0f} units<extra></extra>'
        ))
    
    # Add 2025 historical reference line
    if 'Historical_Qty_2025' in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast['Historical_Qty_2025'],
            mode='lines+markers',
            name='2025 Actual',
            line=dict(color='#f59e0b', width=2, dash='dash'),
            marker=dict(size=6, color='#f59e0b'),
            hovertemplate='<b>%{x} 2025</b><br>Actual: %{y:,.0f} units<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text='🔮 2026 Concentrate Jar Forecast vs Historical',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Quantity (Units)',
            gridcolor='rgba(128,128,128,0.2)'
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


def create_quarterly_chart(quarterly_forecast):
    """
    Create a quarterly summary bar chart.
    """
    if quarterly_forecast.empty:
        return None
    
    fig = go.Figure()
    
    # Add quantity bars
    fig.add_trace(go.Bar(
        x=quarterly_forecast['Quarter'],
        y=quarterly_forecast['Forecasted_Quantity'],
        name='Forecasted Quantity',
        marker=dict(
            color=['#6366f1', '#8b5cf6', '#a855f7', '#c084fc'],
            line=dict(color='rgba(255,255,255,0.3)', width=2)
        ),
        text=quarterly_forecast['Forecasted_Quantity'].apply(lambda x: f"{x/1000:.0f}K"),
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>Quantity: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add error bars for confidence range
    fig.add_trace(go.Scatter(
        x=quarterly_forecast['Quarter'],
        y=quarterly_forecast['Forecasted_Quantity'],
        error_y=dict(
            type='data',
            symmetric=False,
            array=quarterly_forecast['Qty_High'] - quarterly_forecast['Forecasted_Quantity'],
            arrayminus=quarterly_forecast['Forecasted_Quantity'] - quarterly_forecast['Qty_Low'],
            color='rgba(255,255,255,0.5)',
            thickness=2,
            width=10
        ),
        mode='markers',
        marker=dict(size=0),
        showlegend=False,
        hoverinfo='skip'
    ))
    
    fig.update_layout(
        title=dict(
            text='📈 2026 Quarterly Forecast Summary',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Quarter',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Quantity (Units)',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=400,
        margin=dict(t=80, b=60),
        showlegend=False
    )
    
    return fig


def create_revenue_forecast_chart(monthly_forecast):
    """
    Create a revenue forecast chart for 2026.
    """
    if monthly_forecast.empty:
        return None
    
    fig = go.Figure()
    
    # Add confidence band
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'].tolist() + monthly_forecast['MonthShort'].tolist()[::-1],
        y=monthly_forecast['Amt_High'].tolist() + monthly_forecast['Amt_Low'].tolist()[::-1],
        fill='toself',
        fillcolor='rgba(16, 185, 129, 0.2)',
        line=dict(color='rgba(255,255,255,0)'),
        hoverinfo='skip',
        showlegend=True,
        name='Confidence Range'
    ))
    
    # Add forecast line
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast['Forecasted_Amount'],
        mode='lines+markers',
        name='2026 Revenue Forecast',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10, color='#10b981', line=dict(color='white', width=2)),
        text=monthly_forecast['Forecasted_Amount'].apply(lambda x: f"${x/1000:.0f}K"),
        textposition='top center',
        hovertemplate='<b>%{x} 2026</b><br>Revenue: $%{y:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='💰 2026 Revenue Forecast',
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
        height=450,
        margin=dict(t=80, b=60)
    )
    
    return fig


# =============================================================================
# MAIN DISPLAY FUNCTION
# =============================================================================

def main():
    """
    Main function to display the Concentrate Jar Forecasting dashboard.
    Called by main dashboard when this nav item is selected.
    """
    
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.2) 100%);
        border: 1px solid rgba(99, 102, 241, 0.3);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 24px;
    ">
        <h1 style="margin: 0; font-size: 28px;">🧪 Concentrate Jar Forecasting</h1>
        <p style="margin: 8px 0 0 0; opacity: 0.8;">
            2026 4ml Forecast based on all historical glass SKU data • 2024 weighted 60% (healthy stock) • 2025 weighted 40%
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Load data
    with st.spinner("Loading Concentrate Jar Forecasting data..."):
        raw_df = load_concentrate_data()
    
    if raw_df.empty:
        st.error("❌ No data found. Please ensure the 'Concentrate Jar Forecasting' tab exists in your Google Sheet.")
        st.info("Expected columns: Close Date, Quantity, Product, Product Name, Amount, Close Status, Pipeline, Deal Stage, etc.")
        return
    
    # Process data
    df = process_concentrate_data(raw_df)
    
    if df.empty:
        st.error("❌ Could not process data. Check date format in 'Close Date' column.")
        return
    
    # Display data stats in sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Data Summary")
    st.sidebar.metric("Total Records", f"{len(df):,}")
    st.sidebar.metric("Date Range", f"{df['Close Date'].min().strftime('%b %Y')} - {df['Close Date'].max().strftime('%b %Y')}")
    st.sidebar.metric("Total Quantity", f"{df['Quantity'].sum():,.0f}")
    st.sidebar.metric("Total Revenue", f"${df['Amount'].sum():,.0f}")
    
    # Weighting controls
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚖️ Forecast Weights")
    weight_2024 = st.sidebar.slider("2024 Weight (healthy stock)", 0.0, 1.0, 0.6, 0.05)
    weight_2025 = 1.0 - weight_2024
    st.sidebar.caption(f"2025 Weight: {weight_2025:.0%}")
    
    # Generate forecasts
    monthly_forecast, quarterly_forecast, monthly_baselines = generate_2026_forecast(
        df, weight_2024=weight_2024, weight_2025=weight_2025
    )
    
    # =========================
    # TOP METRICS ROW
    # =========================
    
    if not monthly_forecast.empty:
        total_qty_2026 = monthly_forecast['Forecasted_Quantity'].sum()
        total_amt_2026 = monthly_forecast['Forecasted_Amount'].sum()
        q1_qty = monthly_forecast[monthly_forecast['QuarterNum'] == 1]['Forecasted_Quantity'].sum()
        q1_amt = monthly_forecast[monthly_forecast['QuarterNum'] == 1]['Forecasted_Amount'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "2026 Total Quantity",
                f"{total_qty_2026:,.0f}",
                delta=f"{total_qty_2026/1000:.0f}K units"
            )
        
        with col2:
            st.metric(
                "2026 Total Revenue",
                f"${total_amt_2026:,.0f}",
                delta=f"${total_amt_2026/1000:.0f}K"
            )
        
        with col3:
            st.metric(
                "Q1 2026 Quantity",
                f"{q1_qty:,.0f}",
                delta="Highest confidence"
            )
        
        with col4:
            st.metric(
                "Q1 2026 Revenue",
                f"${q1_amt:,.0f}",
                delta="±20% confidence"
            )
    
    st.markdown("---")
    
    # =========================
    # HISTORICAL TREND
    # =========================
    
    st.markdown("### 📈 Historical Demand Trend")
    
    hist_chart = create_historical_trend_chart(df)
    if hist_chart:
        st.plotly_chart(hist_chart, use_container_width=True)
    
    # =========================
    # 2026 FORECAST
    # =========================
    
    st.markdown("---")
    st.markdown("### 🔮 2026 Monthly Forecast")
    
    if not monthly_forecast.empty:
        # Quantity forecast chart
        forecast_chart = create_forecast_chart(monthly_forecast)
        if forecast_chart:
            st.plotly_chart(forecast_chart, use_container_width=True)
        
        # Monthly forecast table
        with st.expander("📋 View Monthly Forecast Details", expanded=False):
            display_df = monthly_forecast[[
                'MonthName', 'Forecasted_Quantity', 'Qty_Low', 'Qty_High',
                'Forecasted_Amount', 'Amt_Low', 'Amt_High',
                'Historical_Qty_2024', 'Historical_Qty_2025', 'Confidence'
            ]].copy()
            
            display_df.columns = [
                'Month', 'Forecast Qty', 'Qty Low', 'Qty High',
                'Forecast $', '$ Low', '$ High',
                '2024 Actual', '2025 Actual', 'Confidence'
            ]
            
            # Format currency columns
            for col in ['Forecast $', '$ Low', '$ High']:
                display_df[col] = display_df[col].apply(lambda x: f"${x:,.0f}")
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # =========================
    # QUARTERLY SUMMARY
    # =========================
    
    st.markdown("---")
    st.markdown("### 📊 Quarterly Summary")
    
    if not quarterly_forecast.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            quarterly_chart = create_quarterly_chart(quarterly_forecast)
            if quarterly_chart:
                st.plotly_chart(quarterly_chart, use_container_width=True)
        
        with col2:
            # Quarterly metrics cards
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
                            <div style="font-size: 12px; opacity: 0.7;">Quantity</div>
                            <div style="font-size: 20px; font-weight: 600;">{row['Forecasted_Quantity']:,.0f}</div>
                            <div style="font-size: 11px; opacity: 0.6;">{row['Qty_Low']:,.0f} - {row['Qty_High']:,.0f}</div>
                        </div>
                        <div style="text-align: right;">
                            <div style="font-size: 12px; opacity: 0.7;">Revenue</div>
                            <div style="font-size: 20px; font-weight: 600; color: #10b981;">${row['Forecasted_Amount']:,.0f}</div>
                            <div style="font-size: 11px; opacity: 0.6;">${row['Amt_Low']:,.0f} - ${row['Amt_High']:,.0f}</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
    
    # =========================
    # REVENUE FORECAST
    # =========================
    
    st.markdown("---")
    st.markdown("### 💰 Revenue Forecast")
    
    if not monthly_forecast.empty:
        revenue_chart = create_revenue_forecast_chart(monthly_forecast)
        if revenue_chart:
            st.plotly_chart(revenue_chart, use_container_width=True)
    
    # =========================
    # METHODOLOGY NOTES
    # =========================
    
    st.markdown("---")
    with st.expander("📖 Forecast Methodology", expanded=False):
        st.markdown(f"""
        ### How This Forecast is Calculated
        
        **Data Source:** All historical glass SKU data from the Concentrate Jar Forecasting tab, 
        used to forecast 4ml demand for 2026 (since many customers switched from other glass SKUs to 4ml).
        
        **Weighting:**
        - **2024:** {weight_2024:.0%} weight (healthy stock levels, more representative demand)
        - **2025:** {weight_2025:.0%} weight (stock constraints may have suppressed true demand)
        
        **Outlier Treatment:** Extreme values are softened using winsorization (capped at 5th and 95th percentiles) 
        to prevent single large orders from skewing the forecast.
        
        **Pipeline Integration:** For Q1 2026, the forecast blends historical patterns (60%) with 
        current pipeline data (40%) for deals expected to close Nov 2025 - Mar 2026.
        
        **Confidence Ranges:**
        - Q1 2026: ±20% (highest confidence - closest to current date + pipeline data)
        - Q2 2026: ±25% (moderate confidence)
        - Q3-Q4 2026: ±30% (lower confidence - further from current date)
        
        **Note:** This forecast treats ALL historical glass SKU volume as potential 4ml demand, 
        based on the product transition strategy.
        """)
    
    # =========================
    # RAW DATA VIEW
    # =========================
    
    with st.expander("🔍 View Raw Data", expanded=False):
        st.dataframe(df.head(100), use_container_width=True)
        st.caption(f"Showing first 100 of {len(df):,} records")


# Entry point when called from main dashboard
if __name__ == "__main__":
    main()
