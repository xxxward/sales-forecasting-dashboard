"""
All Products Forecasting Module
================================
Creates 2026 forecasts for ALL products based on Invoice Line Item data.
Uses weighted historical analysis (2024 weighted higher than 2025).
Allows filtering and analysis by product type, item type, customer, and sales rep.

Navigation: "📦 All Products Forecast" in sidebar menu

Columns from Google Sheet (Invoice Line Item tab):
A: Document Number
B: Status
C: Date
D: Due Date
E: Created From
F: Created By
G: Customer
H: Item
I: Quantity
J: Account
K: Period
L: Department
M: Amount
N: Amount (Transaction Total)
O: Amount Remaining
P: CSM
Q: Date Closed
R: Sales Rep
S: External ID
T: Amount (Shipping)
U: Amount (Transaction Tax Total)
V: Terms
W: Calyx | Item Type
X: PI || Product Type
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
CACHE_VERSION = "all_products_v1"

# =============================================================================
# DATA LOADING
# =============================================================================

@st.cache_data(ttl=CACHE_TTL)
def load_invoice_line_items(version=CACHE_VERSION):
    """
    Load data from Invoice Line Item tab in Google Sheets
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
        
        # Load from Invoice Line Item tab - columns A:X
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="Invoice Line Item!A:X"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning("⚠️ No data found in 'Invoice Line Item' tab")
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
        st.error(f"❌ Error loading Invoice Line Item data: {str(e)}")
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


def format_currency(x):
    """Format currency with M for millions, K for thousands"""
    if x >= 1000000:
        return f"${x/1000000:.1f}M"
    elif x >= 1000:
        return f"${x/1000:.0f}K"
    else:
        return f"${int(x)}"


def format_quantity(x):
    """Format quantity with M for millions, K for thousands"""
    if x >= 1000000:
        return f"{x/1000000:.1f}M"
    elif x >= 1000:
        return f"{x/1000:.0f}K"
    else:
        return str(int(x))


def process_invoice_data(df):
    """
    Process the invoice line item data with known column mappings.
    """
    if df.empty:
        return df
    
    # Standardize column names (handle variations)
    col_mapping = {}
    for col in df.columns:
        col_lower = col.lower().strip()
        if col_lower == 'document number':
            col_mapping[col] = 'Document Number'
        elif col_lower == 'status':
            col_mapping[col] = 'Status'
        elif col_lower == 'date':
            col_mapping[col] = 'Date'
        elif col_lower == 'due date':
            col_mapping[col] = 'Due Date'
        elif col_lower == 'created from':
            col_mapping[col] = 'Created From'
        elif col_lower == 'created by':
            col_mapping[col] = 'Created By'
        elif col_lower == 'customer':
            col_mapping[col] = 'Customer'
        elif col_lower == 'item':
            col_mapping[col] = 'Item'
        elif col_lower == 'quantity':
            col_mapping[col] = 'Quantity'
        elif col_lower == 'account':
            col_mapping[col] = 'Account'
        elif col_lower == 'period':
            col_mapping[col] = 'Period'
        elif col_lower == 'department':
            col_mapping[col] = 'Department'
        elif col_lower == 'amount' and 'transaction' not in col_lower and 'shipping' not in col_lower and 'tax' not in col_lower and 'remaining' not in col_lower:
            col_mapping[col] = 'Amount'
        elif 'amount (transaction total)' in col_lower or col_lower == 'amount (transaction total)':
            col_mapping[col] = 'Amount_Transaction_Total'
        elif 'amount remaining' in col_lower:
            col_mapping[col] = 'Amount_Remaining'
        elif col_lower == 'csm':
            col_mapping[col] = 'CSM'
        elif 'date closed' in col_lower:
            col_mapping[col] = 'Date Closed'
        elif col_lower == 'sales rep':
            col_mapping[col] = 'Sales Rep'
        elif col_lower == 'external id':
            col_mapping[col] = 'External ID'
        elif 'amount (shipping)' in col_lower:
            col_mapping[col] = 'Amount_Shipping'
        elif 'amount (transaction tax' in col_lower:
            col_mapping[col] = 'Amount_Tax'
        elif col_lower == 'terms':
            col_mapping[col] = 'Terms'
        elif 'calyx' in col_lower and 'item type' in col_lower:
            col_mapping[col] = 'Item Type'
        elif 'pi' in col_lower and 'product type' in col_lower:
            col_mapping[col] = 'Product Type'
    
    df = df.rename(columns=col_mapping)
    
    # Parse Date column
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df[df['Date'].notna()].copy()
    else:
        st.warning("⚠️ 'Date' column not found in data")
        return pd.DataFrame()
    
    # Add time-based columns
    df['Year'] = df['Date'].dt.year
    df['Month'] = df['Date'].dt.month
    df['YearMonth'] = df['Date'].dt.to_period('M')
    df['Quarter'] = df['Date'].dt.quarter
    df['MonthName'] = df['Date'].dt.strftime('%b')
    df['MonthLabel'] = df['Date'].dt.strftime('%b %Y')
    
    # Clean numeric columns
    if 'Quantity' in df.columns:
        df['Quantity'] = df['Quantity'].apply(clean_numeric)
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
    
    # Clean up Item Type and Product Type
    if 'Item Type' in df.columns:
        df['Item Type'] = df['Item Type'].fillna('Unknown').replace('', 'Unknown')
    if 'Product Type' in df.columns:
        df['Product Type'] = df['Product Type'].fillna('Unknown').replace('', 'Unknown')
    
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

def calculate_weighted_monthly_averages(df, weight_2024=0.6, weight_2025=0.4, group_by=None):
    """
    Calculate weighted monthly averages using historical data.
    
    2024 is weighted higher (default 60%) because stock was healthy.
    2025 is weighted lower (default 40%) due to stock constraints.
    
    If group_by is specified, returns grouped forecasts.
    """
    if df.empty or 'Year' not in df.columns:
        return pd.DataFrame()
    
    # Aggregate by Year, Month, and optionally group_by
    if group_by and group_by in df.columns:
        monthly_data = df.groupby(['Year', 'Month', group_by]).agg({
            'Quantity': 'sum',
            'Amount': 'sum'
        }).reset_index()
    else:
        monthly_data = df.groupby(['Year', 'Month']).agg({
            'Quantity': 'sum',
            'Amount': 'sum'
        }).reset_index()
    
    # Separate years
    data_2023 = monthly_data[monthly_data['Year'] == 2023]
    data_2024 = monthly_data[monthly_data['Year'] == 2024]
    data_2025 = monthly_data[monthly_data['Year'] == 2025]
    
    weighted_avgs = []
    
    if group_by and group_by in df.columns:
        groups = df[group_by].unique()
        
        for group in groups:
            d2023 = data_2023[data_2023[group_by] == group].set_index('Month')
            d2024 = data_2024[data_2024[group_by] == group].set_index('Month')
            d2025 = data_2025[data_2025[group_by] == group].set_index('Month')
            
            for month in range(1, 13):
                qty_2023 = d2023.loc[month, 'Quantity'] if month in d2023.index else 0
                qty_2024 = d2024.loc[month, 'Quantity'] if month in d2024.index else 0
                qty_2025 = d2025.loc[month, 'Quantity'] if month in d2025.index else 0
                amt_2023 = d2023.loc[month, 'Amount'] if month in d2023.index else 0
                amt_2024 = d2024.loc[month, 'Amount'] if month in d2024.index else 0
                amt_2025 = d2025.loc[month, 'Amount'] if month in d2025.index else 0
                
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
                    group_by: group,
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
    else:
        d2023 = data_2023.set_index('Month')
        d2024 = data_2024.set_index('Month')
        d2025 = data_2025.set_index('Month')
        
        for month in range(1, 13):
            qty_2023 = d2023.loc[month, 'Quantity'] if month in d2023.index else 0
            qty_2024 = d2024.loc[month, 'Quantity'] if month in d2024.index else 0
            qty_2025 = d2025.loc[month, 'Quantity'] if month in d2025.index else 0
            amt_2023 = d2023.loc[month, 'Amount'] if month in d2023.index else 0
            amt_2024 = d2024.loc[month, 'Amount'] if month in d2024.index else 0
            amt_2025 = d2025.loc[month, 'Amount'] if month in d2025.index else 0
            
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


def generate_2026_forecast(df, weight_2024=0.6, weight_2025=0.4, group_by=None):
    """
    Generate month-by-month and quarter-by-quarter 2026 forecast.
    """
    if df.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    
    # Use historical data (before current date for accuracy)
    cutoff_date = datetime.now()
    historical_df = df[df['Date'] < cutoff_date].copy()
    
    # Soften outliers in historical data
    if not historical_df.empty and historical_df['Quantity'].sum() > 0:
        historical_df['Quantity'] = soften_outliers(historical_df['Quantity'])
        historical_df['Amount'] = soften_outliers(historical_df['Amount'])
    
    # Calculate weighted monthly baselines
    monthly_baselines = calculate_weighted_monthly_averages(historical_df, weight_2024, weight_2025, group_by)
    
    if monthly_baselines.empty:
        st.warning("⚠️ Not enough historical data to generate forecast")
        return pd.DataFrame(), pd.DataFrame(), monthly_baselines
    
    # Generate 2026 monthly forecast
    forecast_2026 = []
    
    if group_by and group_by in monthly_baselines.columns:
        for group in monthly_baselines[group_by].unique():
            group_baselines = monthly_baselines[monthly_baselines[group_by] == group]
            
            for month in range(1, 13):
                baseline_row = group_baselines[group_baselines['Month'] == month]
                if baseline_row.empty:
                    continue
                    
                baseline = baseline_row.iloc[0]
                
                forecasted_qty = baseline['Weighted_Quantity']
                forecasted_amt = baseline['Weighted_Amount']
                
                # Confidence range based on forecast horizon
                if month <= 3:
                    confidence = 0.20
                elif month <= 6:
                    confidence = 0.25
                else:
                    confidence = 0.30
                
                quarter_num = (month - 1) // 3 + 1
                
                forecast_2026.append({
                    group_by: group,
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
                    'Historical_Qty_2023': int(baseline['Qty_2023']),
                    'Historical_Qty_2024': int(baseline['Qty_2024']),
                    'Historical_Qty_2025': int(baseline['Qty_2025']),
                    'Historical_Amt_2024': round(baseline['Amt_2024'], 2),
                    'Historical_Amt_2025': round(baseline['Amt_2025'], 2),
                    'Confidence': f"±{int(confidence*100)}%"
                })
    else:
        for month in range(1, 13):
            baseline_row = monthly_baselines[monthly_baselines['Month'] == month]
            if baseline_row.empty:
                continue
                
            baseline = baseline_row.iloc[0]
            
            forecasted_qty = baseline['Weighted_Quantity']
            forecasted_amt = baseline['Weighted_Amount']
            
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
        group_cols = ['QuarterNum', 'Quarter']
        if group_by and group_by in monthly_forecast.columns:
            group_cols = [group_by] + group_cols
        
        quarterly_forecast = monthly_forecast.groupby(group_cols).agg({
            'Forecasted_Quantity': 'sum',
            'Forecasted_Amount': 'sum',
            'Qty_Low': 'sum',
            'Qty_High': 'sum',
            'Amt_Low': 'sum',
            'Amt_High': 'sum'
        }).reset_index()
        
        quarterly_forecast = quarterly_forecast.sort_values('QuarterNum')
    
    return monthly_forecast, quarterly_forecast, monthly_baselines


# =============================================================================
# VISUALIZATION
# =============================================================================

def create_historical_trend_chart(df, title_suffix=""):
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
            y=year_data['Amount'],
            name=str(int(year)),
            marker=dict(
                color=color_map.get(year, '#6366f1'),
                line=dict(color='rgba(255,255,255,0.3)', width=1)
            ),
            text=year_data['Amount'].apply(format_currency),
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Revenue: $%{y:,.0f}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'📊 Historical Revenue by Month{title_suffix}',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            tickangle=-45,
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title='Revenue ($)',
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f'
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


def create_forecast_chart(monthly_forecast, metric='Amount'):
    """
    Create a forecast chart with confidence bands for 2026.
    """
    if monthly_forecast.empty:
        return None
    
    fig = go.Figure()
    
    if metric == 'Amount':
        y_col = 'Forecasted_Amount'
        low_col = 'Amt_Low'
        high_col = 'Amt_High'
        hist_2024 = 'Historical_Amt_2024'
        hist_2025 = 'Historical_Amt_2025'
        y_title = 'Revenue ($)'
        format_func = format_currency
        hover_format = 'Revenue: $%{y:,.0f}'
    else:
        y_col = 'Forecasted_Quantity'
        low_col = 'Qty_Low'
        high_col = 'Qty_High'
        hist_2024 = 'Historical_Qty_2024'
        hist_2025 = 'Historical_Qty_2025'
        y_title = 'Quantity (Units)'
        format_func = format_quantity
        hover_format = 'Quantity: %{y:,.0f}'
    
    # Add confidence band (shaded area)
    fig.add_trace(go.Scatter(
        x=monthly_forecast['MonthShort'].tolist() + monthly_forecast['MonthShort'].tolist()[::-1],
        y=monthly_forecast[high_col].tolist() + monthly_forecast[low_col].tolist()[::-1],
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
        y=monthly_forecast[y_col],
        mode='lines+markers',
        name='2026 Forecast',
        line=dict(color='#10b981', width=3),
        marker=dict(size=10, color='#10b981', line=dict(color='white', width=2)),
        text=monthly_forecast[y_col].apply(format_func),
        textposition='top center',
        hovertemplate=f'<b>%{{x}} 2026</b><br>{hover_format}<extra></extra>'
    ))
    
    # Add 2024 historical reference line
    if hist_2024 in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast[hist_2024],
            mode='lines+markers',
            name='2024 Actual',
            line=dict(color='#8b5cf6', width=2, dash='dot'),
            marker=dict(size=6, color='#8b5cf6'),
            hovertemplate=f'<b>%{{x}} 2024</b><br>{hover_format}<extra></extra>'
        ))
    
    # Add 2025 historical reference line
    if hist_2025 in monthly_forecast.columns:
        fig.add_trace(go.Scatter(
            x=monthly_forecast['MonthShort'],
            y=monthly_forecast[hist_2025],
            mode='lines+markers',
            name='2025 Actual',
            line=dict(color='#f59e0b', width=2, dash='dash'),
            marker=dict(size=6, color='#f59e0b'),
            hovertemplate=f'<b>%{{x}} 2025</b><br>{hover_format}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'🔮 2026 {"Revenue" if metric == "Amount" else "Quantity"} Forecast vs Historical',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Month',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title=y_title,
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f' if metric == 'Amount' else ',.0f'
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


def create_quarterly_chart(quarterly_forecast, metric='Amount'):
    """
    Create a quarterly summary bar chart.
    """
    if quarterly_forecast.empty:
        return None
    
    if metric == 'Amount':
        y_col = 'Forecasted_Amount'
        y_title = 'Revenue ($)'
        format_func = format_currency
        hover_format = 'Revenue: $%{y:,.0f}'
    else:
        y_col = 'Forecasted_Quantity'
        y_title = 'Quantity (Units)'
        format_func = format_quantity
        hover_format = 'Quantity: %{y:,.0f}'
    
    fig = go.Figure()
    
    # Add bars
    fig.add_trace(go.Bar(
        x=quarterly_forecast['Quarter'],
        y=quarterly_forecast[y_col],
        name=f'Forecasted {metric}',
        marker=dict(
            color=['#6366f1', '#8b5cf6', '#a855f7', '#c084fc'],
            line=dict(color='rgba(255,255,255,0.3)', width=2)
        ),
        text=quarterly_forecast[y_col].apply(format_func),
        textposition='outside',
        hovertemplate=f'<b>%{{x}}</b><br>{hover_format}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text=f'📈 2026 Quarterly {"Revenue" if metric == "Amount" else "Quantity"} Forecast',
            font=dict(size=20),
            x=0.5
        ),
        xaxis=dict(
            title='Quarter',
            gridcolor='rgba(128,128,128,0.1)'
        ),
        yaxis=dict(
            title=y_title,
            gridcolor='rgba(128,128,128,0.2)',
            tickformat='$,.0f' if metric == 'Amount' else ',.0f'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=400,
        margin=dict(t=80, b=60),
        showlegend=False
    )
    
    return fig


def create_product_breakdown_chart(df, group_col='Product Type', metric='Amount'):
    """
    Create a pie/donut chart showing breakdown by product type or item type.
    """
    if df.empty or group_col not in df.columns:
        return None
    
    # Aggregate by group
    grouped = df.groupby(group_col).agg({
        'Amount': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    
    grouped = grouped.sort_values(metric, ascending=False)
    
    # Limit to top 10 + "Other"
    if len(grouped) > 10:
        top_10 = grouped.head(10)
        other_val = grouped.iloc[10:][metric].sum()
        other_row = pd.DataFrame({group_col: ['Other'], metric: [other_val]})
        grouped = pd.concat([top_10, other_row], ignore_index=True)
    
    fig = go.Figure(data=[go.Pie(
        labels=grouped[group_col],
        values=grouped[metric],
        hole=0.4,
        marker=dict(
            colors=px.colors.qualitative.Set3[:len(grouped)]
        ),
        textinfo='label+percent',
        textposition='outside',
        hovertemplate=f'<b>%{{label}}</b><br>{metric}: %{{value:,.0f}}<br>Share: %{{percent}}<extra></extra>'
    )])
    
    fig.update_layout(
        title=dict(
            text=f'📊 {"Revenue" if metric == "Amount" else "Quantity"} by {group_col}',
            font=dict(size=18),
            x=0.5
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=450,
        margin=dict(t=60, b=40),
        showlegend=True,
        legend=dict(
            orientation='v',
            yanchor='middle',
            y=0.5,
            xanchor='left',
            x=1.05
        )
    )
    
    return fig


def create_top_customers_chart(df, top_n=15):
    """
    Create a horizontal bar chart of top customers.
    """
    if df.empty or 'Customer' not in df.columns:
        return None
    
    # Aggregate by customer
    customer_totals = df.groupby('Customer').agg({
        'Amount': 'sum',
        'Quantity': 'sum'
    }).reset_index()
    
    customer_totals = customer_totals.sort_values('Amount', ascending=True).tail(top_n)
    
    # Truncate long customer names
    customer_totals['Display_Name'] = customer_totals['Customer'].apply(
        lambda x: x[:30] + '...' if len(str(x)) > 30 else x
    )
    
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        y=customer_totals['Display_Name'],
        x=customer_totals['Amount'],
        orientation='h',
        marker=dict(
            color=customer_totals['Amount'],
            colorscale='Viridis',
            line=dict(color='rgba(255,255,255,0.3)', width=1)
        ),
        text=customer_totals['Amount'].apply(format_currency),
        textposition='outside',
        hovertemplate='<b>%{y}</b><br>Revenue: $%{x:,.0f}<extra></extra>'
    ))
    
    fig.update_layout(
        title=dict(
            text=f'🏆 Top {top_n} Customers by Revenue',
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
        margin=dict(l=180, r=80, t=60, b=40)
    )
    
    return fig


def create_product_forecast_comparison(monthly_forecast, group_col):
    """
    Create a stacked area chart comparing forecasts across products/item types.
    """
    if monthly_forecast.empty or group_col not in monthly_forecast.columns:
        return None
    
    # Pivot data for stacked chart
    pivot_df = monthly_forecast.pivot(
        index='MonthShort', 
        columns=group_col, 
        values='Forecasted_Amount'
    ).fillna(0)
    
    fig = go.Figure()
    
    colors = px.colors.qualitative.Set3
    
    for i, col in enumerate(pivot_df.columns):
        fig.add_trace(go.Scatter(
            x=pivot_df.index,
            y=pivot_df[col],
            name=str(col)[:25],
            mode='lines',
            stackgroup='one',
            line=dict(width=0.5, color=colors[i % len(colors)]),
            fillcolor=colors[i % len(colors)],
            hovertemplate=f'<b>{col}</b><br>%{{x}}: $%{{y:,.0f}}<extra></extra>'
        ))
    
    fig.update_layout(
        title=dict(
            text=f'📊 2026 Revenue Forecast by {group_col}',
            font=dict(size=18),
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
        height=500,
        margin=dict(t=80, b=60),
        hovermode='x unified'
    )
    
    return fig


# =============================================================================
# FORECAST ADJUSTMENTS
# =============================================================================

def apply_forecast_adjustments(monthly_forecast, quarterly_forecast, overall_multiplier=1.0, growth_trend=0.0, quarterly_adjustments=None):
    """
    Apply dynamic adjustments to the forecast.
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
            'Amt_High': 'sum'
        }).reset_index()
        adjusted_quarterly = adjusted_quarterly.sort_values('QuarterNum')
    else:
        adjusted_quarterly = quarterly_forecast.copy()
    
    return adjusted_monthly, adjusted_quarterly


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def main():
    """
    Main function for All Products Forecasting dashboard.
    """
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, rgba(99, 102, 241, 0.2) 0%, rgba(139, 92, 246, 0.2) 100%);
        border: 1px solid rgba(99, 102, 241, 0.3);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 24px;
    ">
        <h1 style="margin: 0; font-size: 28px; display: flex; align-items: center; gap: 12px;">
            📦 All Products Forecast - 2026
        </h1>
        <p style="margin: 8px 0 0 0; opacity: 0.8;">
            Revenue and quantity forecasting for all product lines based on Invoice Line Item historical data
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Load data
    with st.spinner("Loading Invoice Line Item data..."):
        raw_df = load_invoice_line_items()
    
    if raw_df.empty:
        st.error("❌ No data loaded. Check your Google Sheets connection and ensure the 'Invoice Line Item' tab exists.")
        return
    
    # Process data
    df = process_invoice_data(raw_df)
    
    if df.empty:
        st.error("❌ Data processing failed. Check that required columns exist.")
        return
    
    # Show data summary
    st.success(f"✅ Loaded {len(df):,} invoice line items")
    
    # =========================
    # SIDEBAR CONTROLS
    # =========================
    
    st.sidebar.markdown("## 📦 Forecast Controls")
    
    # Historical weighting
    st.sidebar.markdown("### ⚖️ Historical Weights")
    weight_2024 = st.sidebar.slider(
        "2024 Weight", 0.0, 1.0, 0.6, 0.05,
        help="Weight given to 2024 data (healthier stock levels)"
    )
    weight_2025 = 1.0 - weight_2024
    st.sidebar.caption(f"2025 Weight: {weight_2025:.0%}")
    
    # Filter options
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filters")
    
    # Product Type filter
    if 'Product Type' in df.columns:
        product_types = ['All'] + sorted(df['Product Type'].unique().tolist())
        selected_product_type = st.sidebar.selectbox(
            "Product Type",
            product_types,
            help="Filter by PI || Product Type"
        )
    else:
        selected_product_type = 'All'
    
    # Item Type filter
    if 'Item Type' in df.columns:
        item_types = ['All'] + sorted(df['Item Type'].unique().tolist())
        selected_item_type = st.sidebar.selectbox(
            "Item Type",
            item_types,
            help="Filter by Calyx | Item Type"
        )
    else:
        selected_item_type = 'All'
    
    # Customer filter
    if 'Customer' in df.columns:
        customers = ['All'] + sorted(df['Customer'].dropna().unique().tolist())
        selected_customer = st.sidebar.selectbox(
            "Customer",
            customers[:100],  # Limit dropdown size
            help="Filter by customer (top 100 shown)"
        )
    else:
        selected_customer = 'All'
    
    # Sales Rep filter
    if 'Sales Rep' in df.columns:
        sales_reps = ['All'] + sorted(df['Sales Rep'].dropna().unique().tolist())
        selected_rep = st.sidebar.selectbox(
            "Sales Rep",
            sales_reps,
            help="Filter by sales representative"
        )
    else:
        selected_rep = 'All'
    
    # Apply filters
    filtered_df = df.copy()
    filter_desc = []
    
    if selected_product_type != 'All':
        filtered_df = filtered_df[filtered_df['Product Type'] == selected_product_type]
        filter_desc.append(f"Product Type: {selected_product_type}")
    
    if selected_item_type != 'All':
        filtered_df = filtered_df[filtered_df['Item Type'] == selected_item_type]
        filter_desc.append(f"Item Type: {selected_item_type}")
    
    if selected_customer != 'All':
        filtered_df = filtered_df[filtered_df['Customer'] == selected_customer]
        filter_desc.append(f"Customer: {selected_customer}")
    
    if selected_rep != 'All':
        filtered_df = filtered_df[filtered_df['Sales Rep'] == selected_rep]
        filter_desc.append(f"Sales Rep: {selected_rep}")
    
    # Dynamic forecast adjustments
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🎛️ Forecast Adjustments")
    
    overall_multiplier = st.sidebar.slider(
        "Overall Forecast Multiplier", 
        0.5, 2.0, 1.0, 0.05,
        help="Adjust entire forecast up or down (1.0 = no change)"
    )
    
    growth_trend = st.sidebar.slider(
        "Monthly Growth Trend %",
        -5.0, 5.0, 0.0, 0.5,
        help="Apply compound monthly growth/decline"
    )
    
    # Quarterly adjustments
    with st.sidebar.expander("📅 Quarterly Adjustments", expanded=False):
        q1_adj = st.slider("Q1 Adjustment %", -50, 50, 0, 5, key="q1_adj_all")
        q2_adj = st.slider("Q2 Adjustment %", -50, 50, 0, 5, key="q2_adj_all")
        q3_adj = st.slider("Q3 Adjustment %", -50, 50, 0, 5, key="q3_adj_all")
        q4_adj = st.slider("Q4 Adjustment %", -50, 50, 0, 5, key="q4_adj_all")
    
    quarterly_adjustments = {1: q1_adj/100, 2: q2_adj/100, 3: q3_adj/100, 4: q4_adj/100}
    
    # Show active filters
    if filter_desc:
        st.info(f"🔍 Active Filters: {' | '.join(filter_desc)} ({len(filtered_df):,} records)")
    
    # =========================
    # GENERATE FORECASTS
    # =========================
    
    if filtered_df.empty:
        st.warning("⚠️ No data matches the selected filters.")
        return
    
    # Generate aggregate forecast
    monthly_forecast, quarterly_forecast, monthly_baselines = generate_2026_forecast(
        filtered_df, weight_2024=weight_2024, weight_2025=weight_2025
    )
    
    # Store base forecast for comparison
    base_monthly_forecast = monthly_forecast.copy() if not monthly_forecast.empty else pd.DataFrame()
    base_total_amt = monthly_forecast['Forecasted_Amount'].sum() if not monthly_forecast.empty else 0
    
    # Apply adjustments
    if not monthly_forecast.empty:
        monthly_forecast, quarterly_forecast = apply_forecast_adjustments(
            monthly_forecast, 
            quarterly_forecast,
            overall_multiplier=overall_multiplier,
            growth_trend=growth_trend,
            quarterly_adjustments=quarterly_adjustments
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
                "2026 Total Revenue",
                f"${total_amt_2026:,.0f}",
                delta=format_currency(total_amt_2026)
            )
        
        with col2:
            st.metric(
                "2026 Total Quantity",
                f"{total_qty_2026:,.0f}",
                delta=f"{format_quantity(total_qty_2026)} units"
            )
        
        with col3:
            st.metric(
                "Q1 2026 Revenue",
                f"${q1_amt:,.0f}",
                delta="Highest confidence"
            )
        
        with col4:
            st.metric(
                "Q1 2026 Quantity",
                f"{q1_qty:,.0f}",
                delta="±20% confidence"
            )
    
    st.markdown("---")
    
    # =========================
    # TABS FOR DIFFERENT VIEWS
    # =========================
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 Forecast Overview",
        "📊 Product Breakdown",
        "🏆 Customer Analysis",
        "📋 Detailed Data",
        "📖 Methodology"
    ])
    
    with tab1:
        # Historical trend
        st.markdown("### 📈 Historical Revenue Trend")
        hist_chart = create_historical_trend_chart(filtered_df)
        if hist_chart:
            st.plotly_chart(hist_chart, use_container_width=True)
        
        st.markdown("---")
        
        # 2026 Forecast
        st.markdown("### 🔮 2026 Forecast")
        
        metric_choice = st.radio(
            "View metric:",
            ["Revenue", "Quantity"],
            horizontal=True,
            key="forecast_metric"
        )
        
        metric_type = 'Amount' if metric_choice == 'Revenue' else 'Quantity'
        
        if not monthly_forecast.empty:
            forecast_chart = create_forecast_chart(monthly_forecast, metric=metric_type)
            if forecast_chart:
                st.plotly_chart(forecast_chart, use_container_width=True)
        
        # Quarterly summary
        st.markdown("### 📊 Quarterly Summary")
        
        if not quarterly_forecast.empty:
            col1, col2 = st.columns(2)
            
            with col1:
                quarterly_chart = create_quarterly_chart(quarterly_forecast, metric=metric_type)
                if quarterly_chart:
                    st.plotly_chart(quarterly_chart, use_container_width=True)
            
            with col2:
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
                                <div style="font-size: 12px; opacity: 0.7;">Revenue</div>
                                <div style="font-size: 20px; font-weight: 600; color: #10b981;">${row['Forecasted_Amount']:,.0f}</div>
                                <div style="font-size: 11px; opacity: 0.6;">${row['Amt_Low']:,.0f} - ${row['Amt_High']:,.0f}</div>
                            </div>
                            <div style="text-align: right;">
                                <div style="font-size: 12px; opacity: 0.7;">Quantity</div>
                                <div style="font-size: 20px; font-weight: 600;">{row['Forecasted_Quantity']:,.0f}</div>
                                <div style="font-size: 11px; opacity: 0.6;">{row['Qty_Low']:,.0f} - {row['Qty_High']:,.0f}</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
    
    with tab2:
        st.markdown("### 📊 Revenue Breakdown by Product")
        
        breakdown_col = st.radio(
            "Group by:",
            ["Product Type", "Item Type"],
            horizontal=True,
            key="breakdown_col"
        )
        
        col1, col2 = st.columns(2)
        
        with col1:
            pie_chart = create_product_breakdown_chart(filtered_df, breakdown_col, 'Amount')
            if pie_chart:
                st.plotly_chart(pie_chart, use_container_width=True)
        
        with col2:
            qty_pie = create_product_breakdown_chart(filtered_df, breakdown_col, 'Quantity')
            if qty_pie:
                st.plotly_chart(qty_pie, use_container_width=True)
        
        # Product-level forecast table
        st.markdown("### 📋 Forecast by Product")
        
        if breakdown_col in filtered_df.columns:
            product_forecasts = []
            for product in filtered_df[breakdown_col].unique():
                product_df = filtered_df[filtered_df[breakdown_col] == product]
                prod_monthly, prod_quarterly, _ = generate_2026_forecast(product_df, weight_2024, weight_2025)
                
                if not prod_monthly.empty:
                    product_forecasts.append({
                        breakdown_col: product,
                        '2026 Revenue': prod_monthly['Forecasted_Amount'].sum(),
                        '2026 Quantity': prod_monthly['Forecasted_Quantity'].sum(),
                        'Q1 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 1]['Forecasted_Amount'].sum(),
                        'Q2 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 2]['Forecasted_Amount'].sum(),
                        'Q3 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 3]['Forecasted_Amount'].sum(),
                        'Q4 Revenue': prod_monthly[prod_monthly['QuarterNum'] == 4]['Forecasted_Amount'].sum()
                    })
            
            if product_forecasts:
                forecast_table = pd.DataFrame(product_forecasts)
                forecast_table = forecast_table.sort_values('2026 Revenue', ascending=False)
                
                # Format columns
                for col in ['2026 Revenue', 'Q1 Revenue', 'Q2 Revenue', 'Q3 Revenue', 'Q4 Revenue']:
                    forecast_table[col] = forecast_table[col].apply(format_currency)
                forecast_table['2026 Quantity'] = forecast_table['2026 Quantity'].apply(format_quantity)
                
                st.dataframe(forecast_table, use_container_width=True, hide_index=True)
    
    with tab3:
        st.markdown("### 🏆 Customer Analysis")
        
        top_customers_chart = create_top_customers_chart(filtered_df, top_n=15)
        if top_customers_chart:
            st.plotly_chart(top_customers_chart, use_container_width=True)
        
        # Customer summary table
        st.markdown("### 📋 Customer Summary")
        
        if 'Customer' in filtered_df.columns:
            customer_summary = filtered_df.groupby('Customer').agg({
                'Amount': 'sum',
                'Quantity': 'sum',
                'Date': ['count', 'min', 'max']
            }).reset_index()
            
            customer_summary.columns = ['Customer', 'Total Revenue', 'Total Quantity', 'Orders', 'First Order', 'Last Order']
            customer_summary = customer_summary.sort_values('Total Revenue', ascending=False)
            
            # Format columns
            customer_summary['Total Revenue'] = customer_summary['Total Revenue'].apply(format_currency)
            customer_summary['Total Quantity'] = customer_summary['Total Quantity'].apply(format_quantity)
            customer_summary['First Order'] = pd.to_datetime(customer_summary['First Order']).dt.strftime('%b %Y')
            customer_summary['Last Order'] = pd.to_datetime(customer_summary['Last Order']).dt.strftime('%b %Y')
            
            st.dataframe(customer_summary.head(50), use_container_width=True, hide_index=True)
    
    with tab4:
        st.markdown("### 📋 Monthly Forecast Details")
        
        if not monthly_forecast.empty:
            display_df = monthly_forecast[[
                'MonthName', 'Forecasted_Amount', 'Amt_Low', 'Amt_High',
                'Forecasted_Quantity', 'Qty_Low', 'Qty_High',
                'Historical_Amt_2024', 'Historical_Amt_2025', 'Confidence'
            ]].copy()
            
            display_df.columns = [
                'Month', 'Forecast Revenue', 'Rev Low', 'Rev High',
                'Forecast Qty', 'Qty Low', 'Qty High',
                '2024 Revenue', '2025 Revenue', 'Confidence'
            ]
            
            # Format currency columns
            for col in ['Forecast Revenue', 'Rev Low', 'Rev High', '2024 Revenue', '2025 Revenue']:
                display_df[col] = display_df[col].apply(format_currency)
            
            # Format quantity columns
            for col in ['Forecast Qty', 'Qty Low', 'Qty High']:
                display_df[col] = display_df[col].apply(format_quantity)
            
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🔍 Raw Data Sample")
        
        st.dataframe(filtered_df.head(100), use_container_width=True)
        st.caption(f"Showing first 100 of {len(filtered_df):,} records")
    
    with tab5:
        st.markdown(f"""
        ### How This Forecast is Calculated
        
        **Data Source:** Invoice Line Item tab containing historical invoice data with product categorization.
        
        **Weighting:**
        - **2024:** {weight_2024:.0%} weight (healthy stock levels, more representative demand)
        - **2025:** {weight_2025:.0%} weight (stock constraints may have suppressed true demand)
        
        **Outlier Treatment:** Extreme values are softened using winsorization (capped at 5th and 95th percentiles) 
        to prevent single large orders from skewing the forecast.
        
        **Confidence Ranges:**
        - Q1 2026: ±20% (highest confidence - closest to current date)
        - Q2 2026: ±25% (moderate confidence)
        - Q3-Q4 2026: ±30% (lower confidence - further from current date)
        
        **Filters:** Use the sidebar filters to drill down into specific product types, item types, 
        customers, or sales reps for more targeted forecasts.
        
        **Adjustments:** Use the forecast adjustment sliders to apply scenarios like optimistic (+20%) 
        or conservative (-20%) projections, or to model expected growth trends.
        """)


# Entry point when called from main dashboard
if __name__ == "__main__":
    main()
