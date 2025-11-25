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


def format_number(value, include_dollar=False):
    """
    Format numbers with K or M suffix.
    - >= 1,000,000: show as millions (e.g., "1.23M")
    - >= 1,000: show as thousands (e.g., "500K")
    - < 1,000: show as whole number (e.g., "500")
    """
    if value >= 1_000_000:
        formatted = f"{value/1_000_000:.2f}M"
    elif value >= 1_000:
        formatted = f"{value/1_000:.0f}K"
    else:
        formatted = f"{int(value)}"
    
    if include_dollar:
        return f"${formatted}"
    return formatted


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
            text=year_data['Quantity'].apply(format_number),
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
        text=monthly_forecast['Forecasted_Quantity'].apply(format_number),
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
        text=quarterly_forecast['Forecasted_Quantity'].apply(format_number),
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
        text=monthly_forecast['Forecasted_Amount'].apply(lambda x: format_number(x, include_dollar=True)),
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


def apply_forecast_adjustments(monthly_forecast, quarterly_forecast, overall_multiplier=1.0, growth_trend=0.0, quarterly_adjustments=None):
    """
    Apply dynamic adjustments to the forecast.
    
    Args:
        monthly_forecast: Base monthly forecast DataFrame
        quarterly_forecast: Base quarterly forecast DataFrame
        overall_multiplier: Multiply all values by this factor (1.0 = no change)
        growth_trend: Monthly compound growth rate (0.0 = no trend)
        quarterly_adjustments: Dict of {quarter: adjustment_pct} for quarterly tweaks
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
            'Amt_High': 'sum',
            'Pipeline_Qty': 'sum',
            'Pipeline_Amt': 'sum'
        }).reset_index()
        adjusted_quarterly = adjusted_quarterly.sort_values('QuarterNum')
    else:
        adjusted_quarterly = quarterly_forecast.copy()
    
    return adjusted_monthly, adjusted_quarterly


def create_base_vs_adjusted_chart(base_forecast, adjusted_forecast):
    """
    Create a comparison chart showing base forecast vs adjusted forecast.
    """
    if base_forecast.empty or adjusted_forecast.empty:
        return None
    
    fig = go.Figure()
    
    # Base forecast (dotted line)
    fig.add_trace(go.Scatter(
        x=base_forecast['MonthShort'],
        y=base_forecast['Forecasted_Quantity'],
        name='Base Forecast',
        mode='lines+markers',
        line=dict(color='rgba(156, 163, 175, 0.7)', width=2, dash='dot'),
        marker=dict(size=6, color='rgba(156, 163, 175, 0.7)'),
        hovertemplate='<b>%{x} 2026</b><br>Base: %{y:,.0f} units<extra></extra>'
    ))
    
    # Adjusted forecast (solid line)
    fig.add_trace(go.Scatter(
        x=adjusted_forecast['MonthShort'],
        y=adjusted_forecast['Forecasted_Quantity'],
        name='Adjusted Forecast',
        mode='lines+markers',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10, color='#10b981', line=dict(color='white', width=2)),
        fill='tonexty',
        fillcolor='rgba(16, 185, 129, 0.1)',
        hovertemplate='<b>%{x} 2026</b><br>Adjusted: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add difference annotations for significant changes
    for i in range(len(base_forecast)):
        base_val = base_forecast.iloc[i]['Forecasted_Quantity']
        adj_val = adjusted_forecast.iloc[i]['Forecasted_Quantity']
        diff_pct = ((adj_val - base_val) / base_val * 100) if base_val > 0 else 0
        
        if abs(diff_pct) >= 10:  # Only annotate significant differences
            color = '#10b981' if diff_pct > 0 else '#ef4444'
            fig.add_annotation(
                x=adjusted_forecast.iloc[i]['MonthShort'],
                y=adj_val,
                text=f"{diff_pct:+.0f}%",
                showarrow=False,
                font=dict(size=10, color=color),
                yshift=15
            )
    
    fig.update_layout(
        title=dict(
            text='📊 Base vs Adjusted Forecast Comparison',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(title='Month', gridcolor='rgba(128,128,128,0.1)'),
        yaxis=dict(title='Quantity (Units)', gridcolor='rgba(128,128,128,0.2)'),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        height=400,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


# =============================================================================
# PURCHASING STRATEGY CHARTS
# =============================================================================

def create_demand_vs_order_chart(monthly_forecast, order_quantity):
    """
    Create a chart comparing forecasted demand vs order quantity.
    """
    if monthly_forecast.empty:
        return None
    
    # Calculate cumulative demand
    monthly_forecast = monthly_forecast.copy()
    monthly_forecast['Cumulative_Demand'] = monthly_forecast['Forecasted_Quantity'].cumsum()
    
    fig = go.Figure()
    
    # Add monthly demand bars
    fig.add_trace(go.Bar(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast['Forecasted_Quantity'],
        name='Monthly Demand',
        marker=dict(
            color='rgba(99, 102, 241, 0.7)',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=monthly_forecast['Forecasted_Quantity'].apply(format_number),
        textposition='outside',
        hovertemplate='<b>%{x} 2026</b><br>Demand: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add cumulative demand line
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=monthly_forecast['Cumulative_Demand'],
        name='Cumulative Demand',
        mode='lines+markers',
        line=dict(color='#f59e0b', width=3),
        marker=dict(size=8),
        yaxis='y2',
        hovertemplate='<b>%{x} 2026</b><br>Cumulative: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add order quantity reference line
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'],
        y=[order_quantity] * len(monthly_forecast),
        name=f'Order Qty ({order_quantity/1000000:.1f}M)',
        mode='lines',
        line=dict(color='#10b981', width=3, dash='dash'),
        yaxis='y2',
        hovertemplate=f'Order: {order_quantity:,} units<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='📊 2026 Demand Forecast vs Order Quantity',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(title='Month', gridcolor='rgba(128,128,128,0.1)'),
        yaxis=dict(title='Monthly Demand', gridcolor='rgba(128,128,128,0.2)', side='left'),
        yaxis2=dict(title='Cumulative / Order Qty', overlaying='y', side='right', gridcolor='rgba(128,128,128,0.1)'),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        height=450,
        margin=dict(t=80, b=60)
    )
    
    return fig


def calculate_inventory_depletion(monthly_forecast, starting_inventory):
    """
    Calculate month-by-month inventory depletion.
    """
    if monthly_forecast.empty:
        return pd.DataFrame()
    
    depletion = []
    inventory = starting_inventory
    
    # Extend to 2027 if inventory lasts that long
    months_to_extend = max(24, int(starting_inventory / (monthly_forecast['Forecasted_Quantity'].mean() or 1)) + 6)
    
    for i in range(min(months_to_extend, 36)):  # Cap at 3 years
        year = 2026 + (i // 12)
        month = (i % 12) + 1
        
        # Use forecast data for 2026, extrapolate for 2027+
        if i < 12 and i < len(monthly_forecast):
            demand = monthly_forecast.iloc[i]['Forecasted_Quantity']
        else:
            # Use average monthly demand for projection
            demand = monthly_forecast['Forecasted_Quantity'].mean()
        
        ending_inv = max(0, inventory - demand)
        
        depletion.append({
            'Month_Num': i + 1,
            'Month_Label': f"{datetime(year, month, 1).strftime('%b %Y')}",
            'Year': year,
            'Month': month,
            'Starting_Inventory': inventory,
            'Demand': demand,
            'Ending_Inventory': ending_inv
        })
        
        inventory = ending_inv
        
        if inventory <= 0:
            break
    
    return pd.DataFrame(depletion)


def create_inventory_depletion_chart(depletion_data, order_quantity):
    """
    Create inventory depletion timeline chart.
    """
    if depletion_data.empty:
        return None
    
    fig = go.Figure()
    
    # Add inventory level area
    fig.add_trace(go.Scatter(
        x=depletion_data['Month_Label'],
        y=depletion_data['Ending_Inventory'],
        fill='tozeroy',
        name='Inventory Level',
        mode='lines',
        line=dict(color='#6366f1', width=2),
        fillcolor='rgba(99, 102, 241, 0.3)',
        hovertemplate='<b>%{x}</b><br>Inventory: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add demand bars
    fig.add_trace(go.Bar(
        x=depletion_data['Month_Label'],
        y=depletion_data['Demand'],
        name='Monthly Demand',
        marker=dict(color='rgba(239, 68, 68, 0.5)'),
        hovertemplate='<b>%{x}</b><br>Demand: %{y:,.0f} units<extra></extra>'
    ))
    
    # Add safety stock line (e.g., 3 months of demand)
    avg_demand = depletion_data['Demand'].mean()
    safety_stock = avg_demand * 3
    
    fig.add_hline(
        y=safety_stock, 
        line_dash="dot", 
        line_color="#f59e0b",
        annotation_text=f"Safety Stock ({safety_stock/1000:.0f}K)",
        annotation_position="right"
    )
    
    fig.update_layout(
        title=dict(
            text='📦 Inventory Depletion Timeline',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(title='Month', tickangle=-45, gridcolor='rgba(128,128,128,0.1)'),
        yaxis=dict(title='Units', gridcolor='rgba(128,128,128,0.2)'),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        height=400,
        margin=dict(t=60, b=80),
        barmode='overlay'
    )
    
    return fig


def create_cashflow_comparison_chart(tooling_cost, down_payment_a, remaining_a, down_payment_b, remaining_b, monthly_forecast, order_quantity, unit_cost):
    """
    Create cash flow comparison chart for both options.
    """
    if monthly_forecast.empty:
        return None
    
    # Timeline: Order placed (Month 0), Shipment (Month 2), then monthly revenue
    months = ['Order\n(Jan)', 'Production\n(Feb)', 'Shipment\n(Mar)'] + [f"{m}\n(Revenue)" for m in monthly_forecast['MonthShort'].tolist()[:9]]
    
    # Option A: Tooling Upfront
    cashflow_a = [
        -(tooling_cost + down_payment_a),  # Order: tooling + down payment
        0,  # Production
        -remaining_a,  # Shipment: remaining balance
    ]
    
    # Option B: Tooling Baked In
    cashflow_b = [
        -down_payment_b,  # Order: just down payment (no separate tooling)
        0,  # Production
        -remaining_b,  # Shipment: remaining balance
    ]
    
    # Add revenue months (same for both options)
    for i in range(min(9, len(monthly_forecast))):
        revenue = monthly_forecast.iloc[i]['Forecasted_Amount']
        cashflow_a.append(revenue)
        cashflow_b.append(revenue)
    
    # Calculate cumulative
    cumulative_a = np.cumsum(cashflow_a)
    cumulative_b = np.cumsum(cashflow_b)
    
    fig = go.Figure()
    
    # Option A cumulative
    fig.add_trace(go.Scatter(
        x=months,
        y=cumulative_a,
        name='Option A: Tooling Upfront',
        mode='lines+markers',
        line=dict(color='#10b981', width=3),
        marker=dict(size=8),
        hovertemplate='<b>%{x}</b><br>Cumulative: $%{y:,.0f}<extra></extra>'
    ))
    
    # Option B cumulative
    fig.add_trace(go.Scatter(
        x=months,
        y=cumulative_b,
        name='Option B: Tooling Baked In',
        mode='lines+markers',
        line=dict(color='#f59e0b', width=3),
        marker=dict(size=8),
        hovertemplate='<b>%{x}</b><br>Cumulative: $%{y:,.0f}<extra></extra>'
    ))
    
    # Add break-even line
    fig.add_hline(y=0, line_dash="dash", line_color="white", opacity=0.5)
    
    fig.update_layout(
        title=dict(
            text='💵 Cumulative Cash Flow Comparison',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(title='Timeline', gridcolor='rgba(128,128,128,0.1)', tickangle=-30),
        yaxis=dict(title='Cumulative Cash Flow ($)', gridcolor='rgba(128,128,128,0.2)', tickformat='$,.0f'),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        height=450,
        margin=dict(t=80, b=80)
    )
    
    return fig


# =============================================================================
# CUSTOMER ANALYSIS
# =============================================================================

def analyze_customers(df):
    """
    Analyze customer ordering patterns and behavior.
    """
    if df.empty or 'Company Name' not in df.columns:
        return pd.DataFrame()
    
    # Filter out empty company names
    df_customers = df[df['Company Name'].notna() & (df['Company Name'] != '')].copy()
    
    if df_customers.empty:
        return pd.DataFrame()
    
    # Aggregate by customer
    customer_summary = df_customers.groupby('Company Name').agg({
        'Quantity': 'sum',
        'Amount': 'sum',
        'Close Date': ['count', 'min', 'max'],
        'Year': lambda x: list(x.unique())
    }).reset_index()
    
    # Flatten column names
    customer_summary.columns = [
        'Company Name', 'Total_Quantity', 'Total_Amount', 
        'Order_Count', 'First_Order', 'Last_Order', 'Years_Active'
    ]
    
    # Calculate additional metrics
    customer_summary['Months_Active'] = (
        (customer_summary['Last_Order'] - customer_summary['First_Order']).dt.days / 30
    ).round(0).astype(int)
    
    customer_summary['Avg_Order_Value'] = (
        customer_summary['Total_Amount'] / customer_summary['Order_Count']
    ).round(2)
    
    customer_summary['Avg_Order_Qty'] = (
        customer_summary['Total_Quantity'] / customer_summary['Order_Count']
    ).round(0)
    
    # Check if customer was active in specific years
    customer_summary['Is_2024_Customer'] = customer_summary['Years_Active'].apply(
        lambda x: 2024 in x if isinstance(x, list) else False
    )
    customer_summary['Is_2025_Customer'] = customer_summary['Years_Active'].apply(
        lambda x: 2025 in x if isinstance(x, list) else False
    )
    
    # Format dates for display
    customer_summary['First_Order'] = customer_summary['First_Order'].dt.strftime('%b %Y')
    customer_summary['Last_Order'] = customer_summary['Last_Order'].dt.strftime('%b %Y')
    
    # Sort by total revenue
    customer_summary = customer_summary.sort_values('Total_Amount', ascending=False)
    
    return customer_summary


def identify_sticky_customers(df, forecast_total_revenue=None):
    """
    Identify customers most likely to continue ordering (sticky customers).
    
    Criteria:
    - High: 3+ orders AND ordered in 2025
    - Medium: 2 orders OR (1 order in last 6 months)
    - Low: 1 order more than 6 months ago
    
    Projected 2026 revenue is scaled to match the forecast total.
    """
    customer_analysis = analyze_customers(df)
    
    if customer_analysis.empty:
        return pd.DataFrame()
    
    # Get last order as datetime for comparison
    df_customers = df[df['Company Name'].notna() & (df['Company Name'] != '')].copy()
    last_orders = df_customers.groupby('Company Name')['Close Date'].max().reset_index()
    last_orders.columns = ['Company Name', 'Last_Order_Date']
    
    customer_analysis = customer_analysis.merge(last_orders, on='Company Name', how='left')
    
    # Calculate months since last order
    today = datetime.now()
    customer_analysis['Months_Since_Last'] = (
        (today - customer_analysis['Last_Order_Date']).dt.days / 30
    ).round(1)
    
    # Get annual revenue by customer for 2024 and 2025
    annual_revenue = df_customers.groupby(['Company Name', 'Year'])['Amount'].sum().reset_index()
    annual_2024 = annual_revenue[annual_revenue['Year'] == 2024].set_index('Company Name')['Amount']
    annual_2025 = annual_revenue[annual_revenue['Year'] == 2025].set_index('Company Name')['Amount']
    
    # Calculate average annual spend
    def calc_avg_annual(row):
        name = row['Company Name']
        rev_2024 = annual_2024.get(name, 0)
        rev_2025 = annual_2025.get(name, 0)
        
        if rev_2024 > 0 and rev_2025 > 0:
            return (rev_2024 + rev_2025) / 2
        elif rev_2024 > 0:
            return rev_2024
        elif rev_2025 > 0:
            return rev_2025
        else:
            # Fallback: use total amount / years active
            years = len(row['Years_Active']) if isinstance(row['Years_Active'], list) else 1
            return row['Total_Amount'] / max(years, 1)
    
    customer_analysis['Avg_Annual_Revenue'] = customer_analysis.apply(calc_avg_annual, axis=1)
    
    # Assign stickiness
    def calc_stickiness(row):
        if row['Order_Count'] >= 3 and row['Is_2025_Customer']:
            return 'High'
        elif row['Order_Count'] >= 2:
            return 'Medium'
        elif row['Months_Since_Last'] <= 6:
            return 'Medium'
        else:
            return 'Low'
    
    customer_analysis['Stickiness'] = customer_analysis.apply(calc_stickiness, axis=1)
    
    # Calculate raw projected 2026 revenue based on stickiness weights
    def calc_raw_projection(row):
        avg_annual = row['Avg_Annual_Revenue']
        if row['Stickiness'] == 'High':
            return avg_annual * 1.0  # 100% weight
        elif row['Stickiness'] == 'Medium':
            return avg_annual * 0.5  # 50% weight
        else:  # Low
            return avg_annual * 0.25  # 25% weight
    
    customer_analysis['Raw_Projection'] = customer_analysis.apply(calc_raw_projection, axis=1)
    
    # Scale projections to match forecast total
    raw_total = customer_analysis['Raw_Projection'].sum()
    
    if forecast_total_revenue and raw_total > 0:
        scale_factor = forecast_total_revenue / raw_total
        customer_analysis['Projected_2026_Revenue'] = customer_analysis['Raw_Projection'] * scale_factor
    else:
        customer_analysis['Projected_2026_Revenue'] = customer_analysis['Raw_Projection']
    
    return customer_analysis


def analyze_customer_cohorts(df):
    """
    Analyze customer cohorts by first order date.
    """
    if df.empty or 'Company Name' not in df.columns:
        return pd.DataFrame()
    
    df_customers = df[df['Company Name'].notna() & (df['Company Name'] != '')].copy()
    
    # Get first order month for each customer
    first_orders = df_customers.groupby('Company Name')['Close Date'].min().reset_index()
    first_orders.columns = ['Company Name', 'Cohort_Date']
    first_orders['Cohort'] = first_orders['Cohort_Date'].dt.to_period('M')
    
    # Count customers per cohort
    cohort_counts = first_orders.groupby('Cohort').size().reset_index(name='New_Customers')
    cohort_counts['Cohort'] = cohort_counts['Cohort'].astype(str)
    
    return cohort_counts


def create_top_customers_chart(customer_df):
    """
    Create a horizontal bar chart of top customers.
    """
    if customer_df.empty:
        return None
    
    fig = go.Figure()
    
    # Truncate long customer names
    customer_df = customer_df.copy()
    customer_df['Display_Name'] = customer_df['Company Name'].apply(
        lambda x: x[:25] + '...' if len(str(x)) > 25 else x
    )
    
    # Reverse for horizontal bar (top at top)
    customer_df = customer_df.iloc[::-1]
    
    fig.add_trace(go.Bar(
        y=customer_df['Display_Name'],
        x=customer_df['Total_Amount'],
        orientation='h',
        marker=dict(
            color=customer_df['Total_Amount'],
            colorscale='Viridis',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=customer_df['Total_Amount'].apply(lambda x: format_number(x, include_dollar=True)),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Revenue: $%{x:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='🏆 Top Customers by Revenue',
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
        margin=dict(l=150, r=80, t=60, b=40)
    )
    
    return fig


def create_customer_trends_chart(df):
    """
    Create a chart showing customer ordering trends over time.
    """
    if df.empty or 'Company Name' not in df.columns:
        return None
    
    df_customers = df[df['Company Name'].notna() & (df['Company Name'] != '')].copy()
    
    # Count unique customers per month
    monthly_customers = df_customers.groupby(['Year', 'Month']).agg({
        'Company Name': 'nunique',
        'Quantity': 'sum',
        'Amount': 'sum'
    }).reset_index()
    
    monthly_customers.columns = ['Year', 'Month', 'Unique_Customers', 'Total_Qty', 'Total_Revenue']
    
    monthly_customers['MonthLabel'] = monthly_customers.apply(
        lambda x: f"{datetime(int(x['Year']), int(x['Month']), 1).strftime('%b %Y')}", 
        axis=1
    )
    monthly_customers['SortKey'] = monthly_customers['Year'] * 100 + monthly_customers['Month']
    monthly_customers = monthly_customers.sort_values('SortKey')
    
    fig = go.Figure()
    
    # Add customer count bars
    fig.add_trace(go.Bar(
        x=monthly_customers['MonthLabel'],
        y=monthly_customers['Unique_Customers'],
        name='Active Customers',
        marker=dict(
            color='rgba(99, 102, 241, 0.8)',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=monthly_customers['Unique_Customers'],
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>Active Customers: %{y}<extra></extra>'
    ))
    
    # Add revenue line on secondary axis
    fig.add_trace(go.Scatter(
        x=monthly_customers['MonthLabel'],
        y=monthly_customers['Total_Revenue'],
        name='Revenue',
        mode='lines+markers',
        line=dict(color='#10b981', width=3),
        marker=dict(size=8),
        yaxis='y2',
        hovertemplate='<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='📈 Monthly Active Customers & Revenue',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Active Customers',
            gridcolor='rgba(128,128,128,0.2)',
            side='left'
        ),
        yaxis2=dict(
            title='Revenue ($)',
            overlaying='y',
            side='right',
            tickformat='$,.0f',
            gridcolor='rgba(128,128,128,0.1)'
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
        margin=dict(t=80, b=80)
    )
    
    return fig


def create_cohort_chart(cohort_data):
    """
    Create a chart showing new customer acquisition by cohort.
    """
    if cohort_data.empty:
        return None
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=cohort_data['Cohort'],
        y=cohort_data['New_Customers'],
        marker=dict(
            color='rgba(139, 92, 246, 0.8)',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=cohort_data['New_Customers'],
        textposition='outside',
        hovertemplate='<b>%{x}</b><br>New Customers: %{y}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text='🆕 New Customer Acquisition by Month',
            font=dict(size=18),
            x=0.5
        ),
        xaxis=dict(
            title='Cohort (First Order Month)',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='New Customers',
            gridcolor='rgba(128,128,128,0.2)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=400,
        margin=dict(t=60, b=80)
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
    
    # Dynamic forecast adjustments
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎛️ Forecast Adjustments")
    
    # Overall multiplier
    overall_multiplier = st.sidebar.slider(
        "Overall Forecast Multiplier", 
        0.5, 2.0, 1.0, 0.05,
        help="Adjust entire forecast up or down (1.0 = no change)"
    )
    
    # Growth trend
    growth_trend = st.sidebar.slider(
        "Monthly Growth Trend %",
        -5.0, 5.0, 0.0, 0.5,
        help="Apply compound monthly growth/decline"
    )
    
    # Initialize quarterly adjustments with defaults
    q1_adj = 0
    q2_adj = 0
    q3_adj = 0
    q4_adj = 0
    
    # Quarterly adjustments
    with st.sidebar.expander("📅 Quarterly Adjustments", expanded=False):
        q1_adj = st.slider("Q1 Adjustment %", -50, 50, 0, 5, key="q1_adj")
        q2_adj = st.slider("Q2 Adjustment %", -50, 50, 0, 5, key="q2_adj")
        q3_adj = st.slider("Q3 Adjustment %", -50, 50, 0, 5, key="q3_adj")
        q4_adj = st.slider("Q4 Adjustment %", -50, 50, 0, 5, key="q4_adj")
    
    quarterly_adjustments = {1: q1_adj/100, 2: q2_adj/100, 3: q3_adj/100, 4: q4_adj/100}
    
    # Churning customer exclusions
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🚫 Exclude Churning Customers")
    
    # Get unique customers sorted by total amount (biggest first)
    if 'Company Name' in df.columns:
        customer_totals = df.groupby('Company Name').agg({
            'Amount': 'sum',
            'Quantity': 'sum'
        }).reset_index().sort_values('Amount', ascending=False)
        
        # Filter out empty names
        customer_totals = customer_totals[
            customer_totals['Company Name'].notna() & 
            (customer_totals['Company Name'] != '')
        ]
        
        # Create display labels with revenue
        customer_options = []
        customer_revenue_map = {}
        customer_qty_map = {}
        for _, row in customer_totals.iterrows():
            name = row['Company Name']
            amt = row['Amount']
            qty = row['Quantity']
            label = f"{name} (${amt:,.0f})"
            customer_options.append(name)
            customer_revenue_map[name] = amt
            customer_qty_map[name] = qty
        
        # Initialize session state for excluded customers if not exists
        if 'excluded_customers' not in st.session_state:
            st.session_state['excluded_customers'] = []
        
        # Multiselect for excluding customers
        excluded_customers = st.sidebar.multiselect(
            "Select customers to exclude:",
            options=customer_options,
            default=st.session_state.get('excluded_customers', []),
            help="These customers will be removed from the forecast",
            key="customer_exclusion_select"
        )
        
        # Store in session state
        st.session_state['excluded_customers'] = excluded_customers
        
        # Show impact of exclusions
        if excluded_customers:
            excluded_revenue = sum(customer_revenue_map.get(c, 0) for c in excluded_customers)
            excluded_qty = sum(customer_qty_map.get(c, 0) for c in excluded_customers)
            total_revenue = df['Amount'].sum()
            total_qty = df['Quantity'].sum()
            pct_excluded = (excluded_revenue / total_revenue * 100) if total_revenue > 0 else 0
            
            st.sidebar.markdown(f"""
            <div style="
                background: rgba(239, 68, 68, 0.15);
                border: 1px solid rgba(239, 68, 68, 0.3);
                border-radius: 8px;
                padding: 10px;
                margin-top: 8px;
            ">
                <div style="font-size: 11px; opacity: 0.7;">Excluding from forecast:</div>
                <div style="font-size: 14px; font-weight: 600; color: #ef4444;">{len(excluded_customers)} customers</div>
                <div style="font-size: 12px; margin-top: 4px;">
                    -{excluded_qty:,.0f} units ({excluded_qty/total_qty*100:.1f}%)<br>
                    -${excluded_revenue:,.0f} ({pct_excluded:.1f}%)
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Quick actions
            if st.sidebar.button("🔄 Clear Exclusions", use_container_width=True):
                st.session_state['excluded_customers'] = []
                st.rerun()
    else:
        excluded_customers = []
        customer_revenue_map = {}
        customer_qty_map = {}
    
    # Scenario presets
    st.sidebar.markdown("**Quick Scenarios:**")
    scenario_col1, scenario_col2 = st.sidebar.columns(2)
    with scenario_col1:
        if st.button("📈 Optimistic", use_container_width=True, help="+20% overall"):
            st.session_state['scenario'] = 'optimistic'
    with scenario_col2:
        if st.button("📉 Conservative", use_container_width=True, help="-20% overall"):
            st.session_state['scenario'] = 'conservative'
    
    # Check for scenario overrides
    if 'scenario' in st.session_state:
        if st.session_state['scenario'] == 'optimistic':
            overall_multiplier = 1.2
            st.sidebar.success("📈 Optimistic scenario: +20%")
        elif st.session_state['scenario'] == 'conservative':
            overall_multiplier = 0.8
            st.sidebar.warning("📉 Conservative scenario: -20%")
        
        if st.sidebar.button("🔄 Reset to Base", use_container_width=True):
            del st.session_state['scenario']
            st.rerun()
    
    # Generate forecasts
    # First, filter out excluded customers
    df_for_forecast = df.copy()
    if excluded_customers:
        df_for_forecast = df_for_forecast[~df_for_forecast['Company Name'].isin(excluded_customers)]
    
    monthly_forecast, quarterly_forecast, monthly_baselines = generate_2026_forecast(
        df_for_forecast, weight_2024=weight_2024, weight_2025=weight_2025
    )
    
    # Store base forecast for comparison before adjustments
    base_monthly_forecast = monthly_forecast.copy() if not monthly_forecast.empty else pd.DataFrame()
    base_total_qty = monthly_forecast['Forecasted_Quantity'].sum() if not monthly_forecast.empty else 0
    base_total_amt = monthly_forecast['Forecasted_Amount'].sum() if not monthly_forecast.empty else 0
    
    # Apply dynamic adjustments
    if not monthly_forecast.empty:
        monthly_forecast, quarterly_forecast = apply_forecast_adjustments(
            monthly_forecast, 
            quarterly_forecast,
            overall_multiplier=overall_multiplier,
            growth_trend=growth_trend,
            quarterly_adjustments=quarterly_adjustments
        )
    
    # Calculate adjustment impact
    adjusted_total_qty = monthly_forecast['Forecasted_Quantity'].sum() if not monthly_forecast.empty else 0
    adjusted_total_amt = monthly_forecast['Forecasted_Amount'].sum() if not monthly_forecast.empty else 0
    qty_change_pct = ((adjusted_total_qty - base_total_qty) / base_total_qty * 100) if base_total_qty > 0 else 0
    amt_change_pct = ((adjusted_total_amt - base_total_amt) / base_total_amt * 100) if base_total_amt > 0 else 0
    
    # Show adjustment impact if any adjustments are active
    adjustments_active = (overall_multiplier != 1.0 or growth_trend != 0.0 or 
                          any(v != 0 for v in quarterly_adjustments.values()))
    
    if adjustments_active:
        change_color = "#10b981" if qty_change_pct >= 0 else "#ef4444"
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, rgba(251, 191, 36, 0.15) 0%, rgba(245, 158, 11, 0.15) 100%);
            border: 1px solid rgba(251, 191, 36, 0.4);
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 16px;
        ">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <span style="font-size: 24px;">🎛️</span>
                    <div>
                        <div style="font-weight: 600; color: #fbbf24;">Forecast Adjustments Active</div>
                        <div style="font-size: 12px; opacity: 0.8;">Modified from base calculation</div>
                    </div>
                </div>
                <div style="display: flex; gap: 24px;">
                    <div style="text-align: center;">
                        <div style="font-size: 11px; opacity: 0.7;">Base Qty</div>
                        <div style="font-size: 14px; text-decoration: line-through; opacity: 0.6;">{base_total_qty:,.0f}</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 11px; opacity: 0.7;">Adjusted Qty</div>
                        <div style="font-size: 16px; font-weight: 700; color: {change_color};">{adjusted_total_qty:,.0f}</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 11px; opacity: 0.7;">Change</div>
                        <div style="font-size: 16px; font-weight: 700; color: {change_color};">{qty_change_pct:+.1f}%</div>
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # Show customer exclusion banner if any customers are excluded
    if excluded_customers:
        excluded_revenue = sum(customer_revenue_map.get(c, 0) for c in excluded_customers)
        excluded_qty = sum(customer_qty_map.get(c, 0) for c in excluded_customers)
        
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, rgba(239, 68, 68, 0.15) 0%, rgba(220, 38, 38, 0.15) 100%);
            border: 1px solid rgba(239, 68, 68, 0.4);
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 16px;
        ">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <span style="font-size: 24px;">🚫</span>
                    <div>
                        <div style="font-weight: 600; color: #ef4444;">Customers Excluded from Forecast</div>
                        <div style="font-size: 12px; opacity: 0.8;">{len(excluded_customers)} churning customers removed</div>
                    </div>
                </div>
                <div style="display: flex; gap: 24px;">
                    <div style="text-align: center;">
                        <div style="font-size: 11px; opacity: 0.7;">Excluded Qty</div>
                        <div style="font-size: 16px; font-weight: 700; color: #ef4444;">-{excluded_qty:,.0f}</div>
                    </div>
                    <div style="text-align: center;">
                        <div style="font-size: 11px; opacity: 0.7;">Excluded Revenue</div>
                        <div style="font-size: 16px; font-weight: 700; color: #ef4444;">-${excluded_revenue:,.0f}</div>
                    </div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Show excluded customers in expandable section
        with st.expander("👀 View Excluded Customers", expanded=False):
            excl_data = []
            for cust in excluded_customers:
                excl_data.append({
                    'Customer': cust,
                    'Historical Qty': f"{customer_qty_map.get(cust, 0):,.0f}",
                    'Historical Revenue': f"${customer_revenue_map.get(cust, 0):,.0f}"
                })
            st.dataframe(pd.DataFrame(excl_data), use_container_width=True, hide_index=True)
    
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
        
        # Show base vs adjusted comparison if adjustments active
        if adjustments_active and not base_monthly_forecast.empty:
            with st.expander("📊 Base vs Adjusted Comparison", expanded=True):
                comparison_chart = create_base_vs_adjusted_chart(base_monthly_forecast, monthly_forecast)
                if comparison_chart:
                    st.plotly_chart(comparison_chart, use_container_width=True)
                
                # Side by side quarterly comparison
                st.markdown("**Quarterly Impact:**")
                comp_cols = st.columns(4)
                for i, q in enumerate([1, 2, 3, 4]):
                    base_q = base_monthly_forecast[base_monthly_forecast['QuarterNum'] == q]['Forecasted_Quantity'].sum()
                    adj_q = monthly_forecast[monthly_forecast['QuarterNum'] == q]['Forecasted_Quantity'].sum()
                    change = ((adj_q - base_q) / base_q * 100) if base_q > 0 else 0
                    
                    with comp_cols[i]:
                        delta_color = "normal" if change >= 0 else "inverse"
                        st.metric(
                            f"Q{q} 2026",
                            f"{adj_q:,.0f}",
                            delta=f"{change:+.1f}%",
                            delta_color=delta_color
                        )
        
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
    # CUSTOMER BREAKDOWN
    # =========================
    
    st.markdown("---")
    st.markdown("### 👥 Customer Breakdown")
    
    # Note about exclusions
    if excluded_customers:
        st.caption(f"⚠️ {len(excluded_customers)} churning customers excluded from this analysis")
    
    # Generate customer analysis (using filtered df)
    customer_analysis = analyze_customers(df_for_forecast)
    
    if not customer_analysis.empty:
        # Top metrics row
        total_customers = len(customer_analysis)
        repeat_customers = len(customer_analysis[customer_analysis['Order_Count'] > 1])
        repeat_rate = (repeat_customers / total_customers * 100) if total_customers > 0 else 0
        top_customer_revenue = customer_analysis['Total_Amount'].max()
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Customers", f"{total_customers:,}")
        with col2:
            st.metric("Repeat Customers", f"{repeat_customers:,}", delta=f"{repeat_rate:.0f}% repeat rate")
        with col3:
            st.metric("Avg Orders/Customer", f"{customer_analysis['Order_Count'].mean():.1f}")
        with col4:
            st.metric("Top Customer Revenue", f"${top_customer_revenue:,.0f}")
        
        st.markdown("")
        
        # Create tabs for different views
        cust_tab1, cust_tab2, cust_tab3 = st.tabs(["🏆 Top Customers", "📈 Customer Trends", "🔍 Full Customer List"])
        
        with cust_tab1:
            # Top customers chart
            top_customers_chart = create_top_customers_chart(customer_analysis.head(15))
            if top_customers_chart:
                st.plotly_chart(top_customers_chart, use_container_width=True)
            
            # Top customers table
            st.markdown("#### Top 15 Customers by Revenue")
            top_display = customer_analysis.head(15)[[
                'Company Name', 'Total_Quantity', 'Total_Amount', 'Order_Count', 
                'First_Order', 'Last_Order', 'Months_Active', 'Avg_Order_Value'
            ]].copy()
            
            top_display.columns = [
                'Customer', 'Total Qty', 'Total Revenue', 'Orders', 
                'First Order', 'Last Order', 'Months Active', 'Avg Order $'
            ]
            
            top_display['Total Revenue'] = top_display['Total Revenue'].apply(lambda x: f"${x:,.0f}")
            top_display['Avg Order $'] = top_display['Avg Order $'].apply(lambda x: f"${x:,.0f}")
            top_display['Total Qty'] = top_display['Total Qty'].apply(lambda x: f"{x:,.0f}")
            
            st.dataframe(top_display, use_container_width=True, hide_index=True)
        
        with cust_tab2:
            # Customer ordering patterns over time
            customer_trends_chart = create_customer_trends_chart(df_for_forecast)
            if customer_trends_chart:
                st.plotly_chart(customer_trends_chart, use_container_width=True)
            
            # Cohort analysis - customers by first order date
            st.markdown("#### Customer Cohort Analysis")
            cohort_data = analyze_customer_cohorts(df_for_forecast)
            if not cohort_data.empty:
                cohort_chart = create_cohort_chart(cohort_data)
                if cohort_chart:
                    st.plotly_chart(cohort_chart, use_container_width=True)
        
        with cust_tab3:
            # Search/filter
            search_term = st.text_input("🔍 Search customers", placeholder="Type customer name...")
            
            filtered_customers = customer_analysis.copy()
            if search_term:
                filtered_customers = filtered_customers[
                    filtered_customers['Company Name'].str.lower().str.contains(search_term.lower(), na=False)
                ]
            
            # Full customer table
            full_display = filtered_customers[[
                'Company Name', 'Total_Quantity', 'Total_Amount', 'Order_Count',
                'First_Order', 'Last_Order', 'Months_Active', 'Avg_Order_Value', 
                'Is_2024_Customer', 'Is_2025_Customer'
            ]].copy()
            
            full_display.columns = [
                'Customer', 'Total Qty', 'Total Revenue', 'Orders',
                'First Order', 'Last Order', 'Months Active', 'Avg Order $',
                '2024 Customer', '2025 Customer'
            ]
            
            full_display['Total Revenue'] = full_display['Total Revenue'].apply(lambda x: f"${x:,.0f}")
            full_display['Avg Order $'] = full_display['Avg Order $'].apply(lambda x: f"${x:,.0f}")
            full_display['Total Qty'] = full_display['Total Qty'].apply(lambda x: f"{x:,.0f}")
            full_display['2024 Customer'] = full_display['2024 Customer'].apply(lambda x: "✅" if x else "")
            full_display['2025 Customer'] = full_display['2025 Customer'].apply(lambda x: "✅" if x else "")
            
            st.dataframe(full_display, use_container_width=True, hide_index=True, height=400)
            st.caption(f"Showing {len(filtered_customers):,} customers")
            
            # Download button
            csv = customer_analysis.to_csv(index=False)
            st.download_button(
                label="📥 Download Full Customer List (CSV)",
                data=csv,
                file_name="concentrate_jar_customers.csv",
                mime="text/csv"
            )
    
    # =========================
    # CUSTOMER STICKINESS ANALYSIS
    # =========================
    
    st.markdown("---")
    st.markdown("### 🎯 Customer Stickiness Analysis")
    st.caption("Identifying customers most likely to continue ordering in 2026")
    
    # Get the 2026 forecast total to scale customer projections
    forecast_2026_total = monthly_forecast['Forecasted_Amount'].sum() if not monthly_forecast.empty else None
    
    # Use filtered df (excludes churning customers)
    sticky_customers = identify_sticky_customers(df_for_forecast, forecast_total_revenue=forecast_2026_total)
    
    if not sticky_customers.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            # High likelihood customers
            st.markdown("""
            <div style="
                background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.15) 100%);
                border: 1px solid rgba(16, 185, 129, 0.3);
                border-radius: 12px;
                padding: 16px;
                margin-bottom: 12px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #10b981; margin-bottom: 12px;">
                    🟢 High Likelihood (3+ orders, active in 2025)
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            high_likelihood = sticky_customers[sticky_customers['Stickiness'] == 'High'].sort_values('Projected_2026_Revenue', ascending=False).head(10)
            if not high_likelihood.empty:
                for _, row in high_likelihood.iterrows():
                    st.markdown(f"**{row['Company Name']}** - {row['Order_Count']} orders, ${row['Projected_2026_Revenue']:,.0f} projected")
            else:
                st.info("No high-likelihood customers found")
        
        with col2:
            # Medium likelihood customers
            st.markdown("""
            <div style="
                background: linear-gradient(135deg, rgba(251, 191, 36, 0.15) 0%, rgba(245, 158, 11, 0.15) 100%);
                border: 1px solid rgba(251, 191, 36, 0.3);
                border-radius: 12px;
                padding: 16px;
                margin-bottom: 12px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #fbbf24; margin-bottom: 12px;">
                    🟡 Medium Likelihood (2 orders OR recent single order)
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            medium_likelihood = sticky_customers[sticky_customers['Stickiness'] == 'Medium'].sort_values('Projected_2026_Revenue', ascending=False).head(10)
            if not medium_likelihood.empty:
                for _, row in medium_likelihood.iterrows():
                    st.markdown(f"**{row['Company Name']}** - {row['Order_Count']} orders, ${row['Projected_2026_Revenue']:,.0f} projected")
            else:
                st.info("No medium-likelihood customers found")
        
        # Stickiness summary with PROJECTED 2026 revenue
        stickiness_summary = sticky_customers.groupby('Stickiness').agg({
            'Company Name': 'count',
            'Projected_2026_Revenue': 'sum',
            'Total_Quantity': 'sum'
        }).reset_index()
        stickiness_summary.columns = ['Likelihood', 'Customer Count', 'Projected 2026 Revenue', 'Total Quantity']
        
        st.markdown("#### Stickiness Summary - Projected 2026 Revenue")
        summary_col1, summary_col2, summary_col3 = st.columns(3)
        
        for i, likelihood in enumerate(['High', 'Medium', 'Low']):
            row = stickiness_summary[stickiness_summary['Likelihood'] == likelihood]
            if not row.empty:
                with [summary_col1, summary_col2, summary_col3][i]:
                    color = {'High': '#10b981', 'Medium': '#fbbf24', 'Low': '#ef4444'}[likelihood]
                    projected_rev = row['Projected 2026 Revenue'].values[0]
                    st.markdown(f"""
                    <div style="text-align: center; padding: 12px; background: rgba(0,0,0,0.2); border-radius: 8px;">
                        <div style="font-size: 12px; opacity: 0.7;">{likelihood} Likelihood</div>
                        <div style="font-size: 24px; font-weight: 700; color: {color};">{int(row['Customer Count'].values[0])}</div>
                        <div style="font-size: 13px; opacity: 0.9; color: {color};">${projected_rev:,.0f}</div>
                        <div style="font-size: 10px; opacity: 0.5;">projected 2026</div>
                    </div>
                    """, unsafe_allow_html=True)
        
        # Show total projected revenue
        total_projected = stickiness_summary['Projected 2026 Revenue'].sum()
        st.markdown(f"""
        <div style="
            text-align: center; 
            padding: 16px; 
            margin-top: 16px;
            background: linear-gradient(135deg, rgba(99, 102, 241, 0.15) 0%, rgba(139, 92, 246, 0.15) 100%);
            border: 1px solid rgba(99, 102, 241, 0.3);
            border-radius: 12px;
        ">
            <div style="font-size: 14px; opacity: 0.7;">2026 Forecast Revenue - Customer Breakdown</div>
            <div style="font-size: 28px; font-weight: 700; color: #10b981;">${total_projected:,.0f}</div>
            <div style="font-size: 11px; opacity: 0.5;">Distributed by customer likelihood (matches 2026 forecast)</div>
        </div>
        """, unsafe_allow_html=True)
    
    # =========================
    # PURCHASING STRATEGY - HEINZ 2M UNIT ANALYSIS
    # =========================
    
    st.markdown("---")
    st.markdown("### 💰 Purchasing Strategy: Heinz 2M Unit Analysis")
    st.caption("Cash flow modeling for 2 million unit order decision")
    
    # Get forecast quantities
    if not monthly_forecast.empty:
        total_qty_2026 = monthly_forecast['Forecasted_Quantity'].sum()
        total_amt_2026 = monthly_forecast['Forecasted_Amount'].sum()
    else:
        total_qty_2026 = 0
        total_amt_2026 = 0
    
    # Configuration inputs in sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🏭 Heinz Order Config")
    
    order_quantity = st.sidebar.number_input("Order Quantity", value=2000000, step=100000, format="%d")
    unit_cost = st.sidebar.number_input("Unit Cost ($)", value=0.35, step=0.01, format="%.2f")
    tooling_cost = st.sidebar.number_input("Tooling Cost ($)", value=75000, step=1000, format="%d")
    down_payment_pct = st.sidebar.slider("Down Payment %", 10, 50, 20, 5)
    
    # Calculate key metrics
    total_unit_cost = order_quantity * unit_cost
    down_payment = total_unit_cost * (down_payment_pct / 100)
    remaining_balance = total_unit_cost - down_payment
    
    # Calculate months of inventory
    monthly_avg_demand = total_qty_2026 / 12 if total_qty_2026 > 0 else 50000
    months_of_inventory = order_quantity / monthly_avg_demand if monthly_avg_demand > 0 else 0
    
    # Top metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("2026 Forecasted Demand", f"{total_qty_2026:,.0f}", delta=f"{total_qty_2026/1000:.0f}K units")
    with col2:
        st.metric("Order vs Demand", f"{(order_quantity/total_qty_2026*100):.0f}%" if total_qty_2026 > 0 else "N/A", 
                  delta=f"{order_quantity - total_qty_2026:+,.0f} units")
    with col3:
        st.metric("Inventory Runway", f"{months_of_inventory:.1f} months", delta=f"~{months_of_inventory/12:.1f} years")
    with col4:
        st.metric("Total Investment", f"${total_unit_cost + tooling_cost:,.0f}")
    
    st.markdown("")
    
    # Create tabs for different analyses
    purch_tab1, purch_tab2, purch_tab3 = st.tabs(["📊 Demand vs Order", "💵 Cash Flow Scenarios", "🎯 Recommendation"])
    
    with purch_tab1:
        # Demand vs Order visualization
        demand_vs_order_chart = create_demand_vs_order_chart(monthly_forecast, order_quantity)
        if demand_vs_order_chart:
            st.plotly_chart(demand_vs_order_chart, use_container_width=True)
        
        # Inventory depletion timeline
        st.markdown("#### 📦 Inventory Depletion Timeline")
        
        depletion_data = calculate_inventory_depletion(monthly_forecast, order_quantity)
        if not depletion_data.empty:
            depletion_chart = create_inventory_depletion_chart(depletion_data, order_quantity)
            if depletion_chart:
                st.plotly_chart(depletion_chart, use_container_width=True)
            
            # Find when inventory runs out
            runout_row = depletion_data[depletion_data['Ending_Inventory'] <= 0].head(1)
            if not runout_row.empty:
                runout_month = runout_row['Month_Label'].values[0]
                st.info(f"📅 **Projected Inventory Runout:** {runout_month}")
            else:
                st.success(f"✅ Inventory lasts beyond forecast period ({months_of_inventory:.1f} months)")
    
    with purch_tab2:
        st.markdown("#### 💵 Cash Flow Scenario Comparison")
        
        # Scenario 1: Tooling Paid Upfront (lower unit cost negotiations later)
        # Scenario 2: Tooling Baked Into Unit Cost (higher per-unit but spread out)
        # Scenario 3: Delayed Order (Wait until 2027)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            <div style="
                background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.15) 100%);
                border: 1px solid rgba(16, 185, 129, 0.3);
                border-radius: 12px;
                padding: 16px;
                margin-bottom: 12px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #10b981; margin-bottom: 12px;">
                    Option A: Pay Tooling Upfront
                </div>
                <div style="font-size: 13px; opacity: 0.9;">
                    ✓ Own the tooling outright<br>
                    ✓ Leverage for lower unit costs later<br>
                    ✓ Clear cost separation<br>
                    ✗ Higher upfront cash outlay
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Cash flow timeline for Option A
            st.markdown("**Cash Flow Timeline:**")
            st.markdown(f"""
            - **Upfront (Order):** ${tooling_cost:,.0f} (tooling) + ${down_payment:,.0f} ({down_payment_pct}% down) = **${tooling_cost + down_payment:,.0f}**
            - **On Shipment:** ${remaining_balance:,.0f}
            - **Total:** ${total_unit_cost + tooling_cost:,.0f}
            - **Per Unit (excl. tooling):** ${unit_cost:.3f}
            """)
        
        with col2:
            # Calculate baked-in unit cost
            tooling_per_unit = tooling_cost / order_quantity
            baked_unit_cost = unit_cost + tooling_per_unit
            baked_total = order_quantity * baked_unit_cost
            baked_down = baked_total * (down_payment_pct / 100)
            baked_remaining = baked_total - baked_down
            
            st.markdown("""
            <div style="
                background: linear-gradient(135deg, rgba(251, 191, 36, 0.15) 0%, rgba(245, 158, 11, 0.15) 100%);
                border: 1px solid rgba(251, 191, 36, 0.3);
                border-radius: 12px;
                padding: 16px;
                margin-bottom: 12px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #fbbf24; margin-bottom: 12px;">
                    Option B: Bake Tooling Into Unit Cost
                </div>
                <div style="font-size: 13px; opacity: 0.9;">
                    ✓ Lower upfront cash outlay<br>
                    ✓ Simpler single line item<br>
                    ✗ Higher per-unit cost forever<br>
                    ✗ Harder to negotiate down later
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("**Cash Flow Timeline:**")
            st.markdown(f"""
            - **Upfront (Order):** ${baked_down:,.0f} ({down_payment_pct}% down)
            - **On Shipment:** ${baked_remaining:,.0f}
            - **Total:** ${baked_total:,.0f}
            - **Per Unit (incl. tooling):** ${baked_unit_cost:.4f} (+${tooling_per_unit:.4f})
            """)
        
        # Cash Flow Comparison Chart
        st.markdown("---")
        st.markdown("#### 📈 Cash Flow Comparison Over Time")
        
        cashflow_chart = create_cashflow_comparison_chart(
            tooling_cost, down_payment, remaining_balance,
            baked_down, baked_remaining,
            monthly_forecast, order_quantity, unit_cost
        )
        if cashflow_chart:
            st.plotly_chart(cashflow_chart, use_container_width=True)
        
        # Break-even analysis
        st.markdown("---")
        st.markdown("#### ⚖️ Break-Even Analysis")
        
        # Calculate when revenue covers costs
        avg_selling_price = total_amt_2026 / total_qty_2026 if total_qty_2026 > 0 else 1.50
        margin_per_unit_a = avg_selling_price - unit_cost
        margin_per_unit_b = avg_selling_price - baked_unit_cost
        
        breakeven_units_a = (tooling_cost + total_unit_cost) / margin_per_unit_a if margin_per_unit_a > 0 else 0
        breakeven_units_b = baked_total / margin_per_unit_b if margin_per_unit_b > 0 else 0
        
        breakeven_months_a = breakeven_units_a / monthly_avg_demand if monthly_avg_demand > 0 else 0
        breakeven_months_b = breakeven_units_b / monthly_avg_demand if monthly_avg_demand > 0 else 0
        
        be_col1, be_col2, be_col3 = st.columns(3)
        
        with be_col1:
            st.metric("Avg Selling Price", f"${avg_selling_price:.2f}/unit")
        with be_col2:
            st.metric("Option A Break-Even", f"{breakeven_months_a:.1f} months", delta=f"{breakeven_units_a:,.0f} units")
        with be_col3:
            st.metric("Option B Break-Even", f"{breakeven_months_b:.1f} months", delta=f"{breakeven_units_b:,.0f} units")
    
    with purch_tab3:
        st.markdown("#### 🎯 Strategic Recommendation")
        
        # Decision matrix
        order_now_score = 0
        wait_score = 0
        
        # Factor 1: Demand coverage
        demand_coverage = order_quantity / total_qty_2026 if total_qty_2026 > 0 else 0
        if 1.0 <= demand_coverage <= 2.0:
            order_now_score += 2
        elif demand_coverage > 2.0:
            wait_score += 1
        
        # Factor 2: Inventory runway
        if 12 <= months_of_inventory <= 24:
            order_now_score += 2
        elif months_of_inventory > 24:
            wait_score += 1
        
        # Factor 3: Cash flow (20% down is favorable)
        if down_payment_pct <= 25:
            order_now_score += 1
        
        # Generate recommendation
        if order_now_score > wait_score:
            recommendation = "ORDER_NOW"
            rec_color = "#10b981"
        else:
            recommendation = "WAIT"
            rec_color = "#f59e0b"
        
        # Display recommendation
        st.markdown(f"""
        <div style="
            background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.2) 100%);
            border: 2px solid rgba(99, 102, 241, 0.4);
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 20px;
            text-align: center;
        ">
            <div style="font-size: 14px; opacity: 0.7; margin-bottom: 8px;">ANALYSIS SUGGESTS</div>
            <div style="font-size: 32px; font-weight: 700; color: {rec_color};">
                {"✅ ORDER 2M UNITS IN 2026" if recommendation == "ORDER_NOW" else "⏳ WAIT UNTIL 2027"}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Detailed reasoning
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### ✅ Reasons to Order Now")
            st.markdown(f"""
            - **Demand Coverage:** 2M units = {demand_coverage:.1f}x your 2026 forecast ({total_qty_2026:,.0f} units)
            - **Inventory Runway:** {months_of_inventory:.1f} months (~{months_of_inventory/12:.1f} years of stock)
            - **Favorable Terms:** {down_payment_pct}% down (vs typical 33%) reduces upfront cash
            - **Price Lock:** Locks in current pricing before potential increases
            - **Supply Security:** Ensures no stockouts during growth
            """)
        
        with col2:
            st.markdown("##### ⚠️ Risks / Considerations")
            st.markdown(f"""
            - **Capital Tied Up:** ${total_unit_cost + tooling_cost:,.0f} total investment
            - **Demand Uncertainty:** Forecast confidence decreases for Q3-Q4
            - **Storage Costs:** {months_of_inventory:.0f} months of inventory to store
            - **Cash Flow Timing:** When does revenue catch up to costs?
            - **Market Changes:** Product/customer preferences could shift
            """)
        
        # Tooling recommendation
        st.markdown("---")
        st.markdown("##### 🔧 Tooling Payment Recommendation")
        
        tooling_upfront_advantage = margin_per_unit_a - margin_per_unit_b
        future_order_savings = tooling_upfront_advantage * 2000000  # Savings on next 2M order
        
        if tooling_upfront_advantage > 0.01:  # More than 1 cent per unit advantage
            st.markdown(f"""
            <div style="
                background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(5, 150, 105, 0.15) 100%);
                border: 1px solid rgba(16, 185, 129, 0.3);
                border-radius: 12px;
                padding: 16px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #10b981;">
                    💡 Recommend: Pay Tooling Upfront (Option A)
                </div>
                <div style="font-size: 13px; margin-top: 8px;">
                    <b>Why:</b> Paying ${tooling_cost:,.0f} upfront gives you negotiating leverage. 
                    On your NEXT 2M unit order, you could save ${future_order_savings:,.0f} by not paying the 
                    ${tooling_per_unit:.4f}/unit tooling amortization again.<br><br>
                    <b>Negotiation angle:</b> "We've already paid for tooling - this order should be at ${unit_cost:.3f}/unit, not ${baked_unit_cost:.4f}/unit."
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="
                background: linear-gradient(135deg, rgba(251, 191, 36, 0.15) 0%, rgba(245, 158, 11, 0.15) 100%);
                border: 1px solid rgba(251, 191, 36, 0.3);
                border-radius: 12px;
                padding: 16px;
            ">
                <div style="font-size: 16px; font-weight: 700; color: #fbbf24;">
                    💡 Consider: Bake Tooling Into Unit Cost (Option B)
                </div>
                <div style="font-size: 13px; margin-top: 8px;">
                    Lower upfront cash outlay (${tooling_cost:,.0f} less at order time). 
                    The per-unit difference is minimal at scale.
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        # Suggested order timing
        st.markdown("---")
        st.markdown("##### 📅 Suggested Order Timing (Cash Flow Optimization)")
        
        # Calculate optimal order month based on revenue accumulation
        if not monthly_forecast.empty:
            cumulative_revenue = monthly_forecast['Forecasted_Amount'].cumsum()
            
            # Find month where cumulative revenue covers down payment + tooling
            initial_outlay = tooling_cost + down_payment
            coverage_month = monthly_forecast[cumulative_revenue >= initial_outlay].head(1)
            
            if not coverage_month.empty:
                optimal_month = coverage_month['MonthName'].values[0]
                st.info(f"""
                📊 **Optimal Order Timing:** By **{optimal_month} 2026**, your cumulative revenue (${initial_outlay:,.0f}+) 
                would cover the initial outlay (tooling + down payment).
                
                **Suggested approach:**
                1. Place order in **Q1 2026** to ensure inventory arrives for peak demand
                2. Use Q1 revenue to cover the {down_payment_pct}% down payment
                3. Remaining balance due on shipment (typically 60-90 days after order)
                """)
            else:
                st.info("📊 Based on forecast, consider ordering at start of Q1 2026 to align with demand cycle.")
    
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
