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
import re
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


# ========== CUSTOM CSS (ENHANCED) ==========
def inject_custom_css():
    st.markdown("""
    <style>
    /* MAIN APP BACKGROUND */
    .stApp {
        background: radial-gradient(circle at top left, #1e293b, #0f172a 60%, #020617);
        color: #e2e8f0;
    }

    /* REMOVE DEFAULT STREAMLIT PADDING */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 8rem;
        max-width: 95% !important;
    }

    /* GLASSMORPHISM CARDS */
    .glass-card {
        background: rgba(30, 41, 59, 0.4);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
        margin-bottom: 20px;
    }
    
    .glass-card:hover {
        border-color: rgba(255, 255, 255, 0.15);
    }

    /* SECTION HEADERS */
    .section-header {
        font-size: 1.5rem;
        font-weight: 700;
        margin-bottom: 1rem;
        margin-top: 1.5rem;
        background: linear-gradient(90deg, #fff, #94a3b8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* HERO METRICS */
    .hero-metric {
        background: linear-gradient(145deg, rgba(30, 41, 59, 0.7), rgba(15, 23, 42, 0.8));
        border-radius: 12px;
        padding: 15px 25px;
        border-left: 4px solid #3b82f6;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        text-align: center;
    }
    .hero-label {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #94a3b8;
        margin-bottom: 5px;
    }
    .hero-value {
        font-size: 1.6rem;
        font-weight: 800;
        color: #fff;
    }

    /* STICKY FOOTER (HUD STYLE) */
    .sticky-forecast-bar-q1 {
        position: fixed;
        bottom: 20px;
        left: 50%;
        transform: translateX(-50%);
        width: 90%;
        max-width: 1400px;
        z-index: 99999;
        background: rgba(15, 23, 42, 0.95);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(59, 130, 246, 0.3);
        border-radius: 24px;
        padding: 12px 30px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 0 30px rgba(59, 130, 246, 0.15), 0 10px 15px -3px rgba(0, 0, 0, 0.5);
    }
    
    .sticky-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        flex: 1;
    }
    .sticky-label {
        font-size: 0.7rem;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 2px;
    }
    .sticky-val {
        font-size: 1.4rem;
        font-weight: 700;
        font-variant-numeric: tabular-nums;
    }
    
    .val-sched { color: #10b981; text-shadow: 0 0 15px rgba(16, 185, 129, 0.4); }
    .val-pipe { color: #3b82f6; text-shadow: 0 0 15px rgba(59, 130, 246, 0.4); }
    .val-reorder { color: #f59e0b; text-shadow: 0 0 15px rgba(245, 158, 11, 0.4); }
    .val-total { 
        font-size: 1.6rem;
        background: linear-gradient(135deg, #fff 0%, #cbd5e1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .val-gap-behind { color: #ef4444; text-shadow: 0 0 15px rgba(239, 68, 68, 0.4); }
    .val-gap-ahead { color: #10b981; text-shadow: 0 0 15px rgba(16, 185, 129, 0.4); }
    
    .sticky-sep {
        width: 1px;
        height: 40px;
        background: linear-gradient(to bottom, transparent, #334155, transparent);
    }

    /* TIER BADGES */
    .tier-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-right: 10px;
    }
    .tier-likely { background: rgba(16, 185, 129, 0.2); color: #34d399; border: 1px solid rgba(16, 185, 129, 0.4); }
    .tier-possible { background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.4); }
    .tier-longshot { background: rgba(148, 163, 184, 0.2); color: #cbd5e1; border: 1px solid rgba(148, 163, 184, 0.4); }

    /* STREAMLIT WIDGET OVERRIDES */
    div[data-testid="stCheckbox"] label {
        color: #e2e8f0 !important;
        font-weight: 500;
    }
    div[data-testid="stExpander"] details {
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 8px;
        background: rgba(0,0,0,0.2);
    }
    
    /* CUSTOM BUTTONS */
    div.stButton > button {
        background: linear-gradient(to right, #1e293b, #0f172a);
        color: white;
        border: 1px solid rgba(255,255,255,0.1);
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        border-color: #3b82f6;
        box-shadow: 0 0 10px rgba(59, 130, 246, 0.3);
    }

    @media (max-width: 768px) {
        .sticky-forecast-bar-q1 {
            width: 95%;
            padding: 10px 15px;
        }
        .sticky-val { font-size: 1.1rem; }
    }
    
    .main .block-container {
        padding-bottom: 120px !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ========== GAUGE CHART (ENHANCED) ==========
def create_q1_gauge(value, goal, title="Q1 2026 Progress"):
    """Create a modern gauge chart for Q1 2026 progress"""
    
    if goal <= 0:
        goal = 1
    
    percentage = (value / goal) * 100
    
    # Modern color palette
    if percentage >= 100:
        bar_color = "#10b981"  # Emerald - at or above goal
    elif percentage >= 75:
        bar_color = "#3b82f6"  # Blue - close
    elif percentage >= 50:
        bar_color = "#f59e0b"  # Amber - mid
    else:
        bar_color = "#ef4444"  # Red - behind
    
    # Set gauge range - adapt to actual value if it exceeds goal
    max_range = max(goal * 1.1, value * 1.05)
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={
            'prefix': "$", 
            'valueformat': ",.0f",
            'font': {'size': 50, 'color': 'white', 'family': 'Inter, sans-serif'}
        },
        title={
            'text': f"<span style='font-size:14px;color:#94a3b8;letter-spacing:1px'>{title.upper()}</span>",
            'font': {'size': 14}
        },
        gauge={
            'axis': {
                'range': [0, max_range], 
                'tickmode': 'array',
                'tickvals': [0, goal, max_range],
                'ticktext': ['0', 'GOAL', ''],
                'tickfont': {'size': 12, 'color': '#64748b'},
                'showticklabels': True
            },
            'bar': {'color': bar_color, 'thickness': 0.8},
            'bgcolor': "rgba(255,255,255,0.05)",
            'borderwidth': 0,
            'steps': [
                {'range': [0, goal], 'color': "rgba(255,255,255,0.03)"}
            ],
            'threshold': {
                'line': {'color': "#fff", 'width': 3},
                'thickness': 0.9,
                'value': goal
            }
        }
    ))
    
    # Add percentage annotation with glow effect
    fig.add_annotation(
        x=0.5, y=0.15,
        text=f"{percentage:.0f}%",
        showarrow=False,
        font=dict(size=24, color=bar_color, family="Inter, sans-serif"),
        xref="paper", yref="paper"
    )
    
    fig.update_layout(
        height=300,
        margin=dict(l=30, r=30, t=50, b=20),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white', 'family': 'Inter, sans-serif'}
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


# ========== HISTORICAL ANALYSIS FUNCTIONS ==========

def load_historical_orders(main_dash, rep_name):
    """
    Load 2025 completed orders for historical analysis
    
    Filters:
    - Date Range: 2025-01-01 to 2025-12-31
    - Status: "Billed" or "Closed" only
    - Rep Master: Match selected rep
    - Amount > 0
    """
    
    # Load raw sales orders data
    historical_df = main_dash.load_google_sheets_data("NS Sales Orders", "A:AF", version=main_dash.CACHE_VERSION)
    
    if historical_df.empty:
        return pd.DataFrame()
    
    col_names = historical_df.columns.tolist()
    
    # Map columns by position (same as main dashboard)
    rename_dict = {}
    
    # Column A: Internal ID
    if len(col_names) > 0:
        rename_dict[col_names[0]] = 'Internal ID'
    
    # Column B: Document Number (SO#) - IMPORTANT for line item matching
    if len(col_names) > 1:
        rename_dict[col_names[1]] = 'SO_Number'
    
    # Column C: Status
    if len(col_names) > 2:
        rename_dict[col_names[2]] = 'Status'
    
    # Column H: Amount (Transaction Total)
    if len(col_names) > 7:
        rename_dict[col_names[7]] = 'Amount'
    
    # Column I: Order Start Date
    if len(col_names) > 8:
        rename_dict[col_names[8]] = 'Order Start Date'
    
    # Column R: Order Type (Product Type)
    if len(col_names) > 17:
        rename_dict[col_names[17]] = 'Order Type'
    
    # Column AE: Corrected Customer Name
    if len(col_names) > 30:
        rename_dict[col_names[30]] = 'Customer'
    
    # Column AF: Rep Master
    if len(col_names) > 31:
        rename_dict[col_names[31]] = 'Rep Master'
    
    historical_df = historical_df.rename(columns=rename_dict)
    
    # Remove duplicate columns
    if historical_df.columns.duplicated().any():
        historical_df = historical_df.loc[:, ~historical_df.columns.duplicated()]
    
    # Clean SO_Number immediately after rename
    if 'SO_Number' in historical_df.columns:
        historical_df['SO_Number'] = historical_df['SO_Number'].astype(str).str.strip().str.upper()
    
    # Clean Status column
    if 'Status' in historical_df.columns:
        historical_df['Status'] = historical_df['Status'].astype(str).str.strip()
        # Filter to Billed and Closed only
        historical_df = historical_df[historical_df['Status'].isin(['Billed', 'Closed'])]
    else:
        return pd.DataFrame()
    
    # Clean Rep Master and filter to selected rep
    if 'Rep Master' in historical_df.columns:
        historical_df['Rep Master'] = historical_df['Rep Master'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        historical_df = historical_df[~historical_df['Rep Master'].isin(invalid_values)]
        historical_df = historical_df[historical_df['Rep Master'] == rep_name]
    else:
        return pd.DataFrame()
    
    # Clean Customer column
    if 'Customer' in historical_df.columns:
        historical_df['Customer'] = historical_df['Customer'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        historical_df = historical_df[~historical_df['Customer'].isin(invalid_values)]
    
    # Clean Amount
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '':
            return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try:
            return float(cleaned)
        except:
            return 0
    
    if 'Amount' in historical_df.columns:
        historical_df['Amount'] = historical_df['Amount'].apply(clean_numeric)
        historical_df = historical_df[historical_df['Amount'] > 0]
    
    # Parse Order Start Date and filter to 2025
    if 'Order Start Date' in historical_df.columns:
        historical_df['Order Start Date'] = pd.to_datetime(historical_df['Order Start Date'], errors='coerce')
        
        # Fix 2-digit year issue
        if historical_df['Order Start Date'].notna().any():
            mask = (historical_df['Order Start Date'].dt.year < 2000) & (historical_df['Order Start Date'].notna())
            if mask.any():
                historical_df.loc[mask, 'Order Start Date'] = historical_df.loc[mask, 'Order Start Date'] + pd.DateOffset(years=100)
        
        # Filter to 2025 only
        year_2025_start = pd.Timestamp('2025-01-01')
        year_2025_end = pd.Timestamp('2025-12-31')
        historical_df = historical_df[
            (historical_df['Order Start Date'] >= year_2025_start) & 
            (historical_df['Order Start Date'] <= year_2025_end)
        ]
    
    # Clean Order Type
    if 'Order Type' in historical_df.columns:
        historical_df['Order Type'] = historical_df['Order Type'].astype(str).str.strip()
        historical_df.loc[historical_df['Order Type'].isin(['', 'nan', 'None']), 'Order Type'] = 'Standard'
    else:
        historical_df['Order Type'] = 'Standard'
    
    return historical_df


def load_invoices(main_dash, rep_name):
    """
    Load 2025 invoices for actual revenue figures
    
    NS Invoice tab columns:
    - Column C: Date (Invoice Date)
    - Column E: Created From (SO# to match with Sales Orders)
    - Column K: Amount (Transaction Total)
    - Column T: Corrected Customer Name
    - Column U: Rep Master
    """
    
    invoice_df = main_dash.load_google_sheets_data("NS Invoices", "A:U", version=main_dash.CACHE_VERSION)
    
    if invoice_df.empty:
        return pd.DataFrame()
    
    col_names = invoice_df.columns.tolist()
    
    rename_dict = {}
    
    # Column C: Date
    if len(col_names) > 2:
        rename_dict[col_names[2]] = 'Invoice_Date'
    
    # Column E: Created From (SO#)
    if len(col_names) > 4:
        rename_dict[col_names[4]] = 'SO_Number'
    
    # Column K: Amount (Transaction Total)
    if len(col_names) > 10:
        rename_dict[col_names[10]] = 'Invoice_Amount'
    
    # Column T: Corrected Customer Name
    if len(col_names) > 19:
        rename_dict[col_names[19]] = 'Customer'
    
    # Column U: Rep Master
    if len(col_names) > 20:
        rename_dict[col_names[20]] = 'Rep Master'
    
    invoice_df = invoice_df.rename(columns=rename_dict)
    
    # Remove duplicate columns
    if invoice_df.columns.duplicated().any():
        invoice_df = invoice_df.loc[:, ~invoice_df.columns.duplicated()]
    
    # Clean Rep Master and filter to selected rep
    if 'Rep Master' in invoice_df.columns:
        invoice_df['Rep Master'] = invoice_df['Rep Master'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        invoice_df = invoice_df[~invoice_df['Rep Master'].isin(invalid_values)]
        invoice_df = invoice_df[invoice_df['Rep Master'] == rep_name]
    else:
        return pd.DataFrame()
    
    # Clean Customer column
    if 'Customer' in invoice_df.columns:
        invoice_df['Customer'] = invoice_df['Customer'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        invoice_df = invoice_df[~invoice_df['Customer'].isin(invalid_values)]
    
    # Clean Amount
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '':
            return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try:
            return float(cleaned)
        except:
            return 0
    
    if 'Invoice_Amount' in invoice_df.columns:
        invoice_df['Invoice_Amount'] = invoice_df['Invoice_Amount'].apply(clean_numeric)
        invoice_df = invoice_df[invoice_df['Invoice_Amount'] > 0]
    
    # Parse Invoice Date and filter to 2025
    if 'Invoice_Date' in invoice_df.columns:
        invoice_df['Invoice_Date'] = pd.to_datetime(invoice_df['Invoice_Date'], errors='coerce')
        
        # Fix 2-digit year issue
        if invoice_df['Invoice_Date'].notna().any():
            mask = (invoice_df['Invoice_Date'].dt.year < 2000) & (invoice_df['Invoice_Date'].notna())
            if mask.any():
                invoice_df.loc[mask, 'Invoice_Date'] = invoice_df.loc[mask, 'Invoice_Date'] + pd.DateOffset(years=100)
        
        # Filter to 2025 only
        year_2025_start = pd.Timestamp('2025-01-01')
        year_2025_end = pd.Timestamp('2025-12-31')
        invoice_df = invoice_df[
            (invoice_df['Invoice_Date'] >= year_2025_start) & 
            (invoice_df['Invoice_Date'] <= year_2025_end)
        ]
    
    # Clean SO_Number for matching - keep full format
    if 'SO_Number' in invoice_df.columns:
        invoice_df['SO_Number'] = invoice_df['SO_Number'].astype(str).str.strip().str.upper()
    
    return invoice_df


def load_line_items(main_dash):
    """
    Load Sales Order Line Items for item-level detail
    
    Sales Order Line Item tab columns:
    - Column B: Document Number (SO#)
    - Column C: Item
    - Column E: Item Rate (price per unit)
    - Column F: Quantity Ordered
    """
    
    line_items_df = main_dash.load_google_sheets_data("Sales Order Line Item", "A:F", version=main_dash.CACHE_VERSION)
    
    if line_items_df.empty:
        return pd.DataFrame()
    
    col_names = line_items_df.columns.tolist()
    
    rename_dict = {}
    
    # Column B: Document Number (SO#)
    if len(col_names) > 1:
        rename_dict[col_names[1]] = 'SO_Number'
    
    # Column C: Item
    if len(col_names) > 2:
        rename_dict[col_names[2]] = 'Item'
    
    # Column E: Item Rate
    if len(col_names) > 4:
        rename_dict[col_names[4]] = 'Item_Rate'
    
    # Column F: Quantity Ordered
    if len(col_names) > 5:
        rename_dict[col_names[5]] = 'Quantity'
    
    line_items_df = line_items_df.rename(columns=rename_dict)
    
    # Remove duplicate columns
    if line_items_df.columns.duplicated().any():
        line_items_df = line_items_df.loc[:, ~line_items_df.columns.duplicated()]
    
    # Clean SO_Number - keep full format (e.g., "SO13778")
    if 'SO_Number' in line_items_df.columns:
        line_items_df['SO_Number'] = line_items_df['SO_Number'].astype(str).str.strip().str.upper()
        line_items_df = line_items_df[line_items_df['SO_Number'] != '']
        line_items_df = line_items_df[line_items_df['SO_Number'].str.lower() != 'nan']
    
    # Clean Item
    if 'Item' in line_items_df.columns:
        line_items_df['Item'] = line_items_df['Item'].astype(str).str.strip()
        line_items_df = line_items_df[line_items_df['Item'] != '']
        line_items_df = line_items_df[line_items_df['Item'].str.lower() != 'nan']
        
        # === COMPREHENSIVE NON-PRODUCT EXCLUSION ===
        
        # Pattern-based exclusions (case-insensitive contains)
        exclude_patterns = [
            # Tax & Fees
            'avatax', 'tax', 'fee', 'convenience', 'surcharge', 'handling',
            # Shipping
            'shipping', 'freight', 'fedex', 'ups ', 'usps', 'ltl', 'truckload',
            'customer pickup', 'client arranged', 'generic ship', 'send to inventory',
            'default shipping', 'best way', 'ground', 'next day', '2nd day', '3rd day',
            'overnight', 'standard', 'saver', 'express', 'priority',
            # Carriers
            'estes', 't-force', 'ward trucking', 'old dominion', 'roadrunner', 
            'xpo logistics', 'abf', 'a. duie pyle', 'frontline freight', 'saia',
            'dependable highway', 'cross country', 'oak harbor',
            # Discounts & Credits
            'discount', 'credit', 'adjustment', 'replacement order', 'partner discount',
            # Creative/Design Services
            'creative', 'pre-press', 'retrofit', 'press proof', 'design', 'die cut sample',
            'label appl', 'application', 'changeover',
            # Misc
            'expedite', 'rush', 'sample', 'testimonial', 'cm-for sos',
            'wip', 'work in progress', 'end of group', 'other', '-not taxable-',
            'fep-liner insert', 'cc payment', 'waive', 'modular plus',
            'canadian business', 'canadian goods'
        ]
        
        # Exact match exclusions (case-insensitive)
        exclude_exact = [
            # Discount codes
            'brad10', 'blake10', '420ten', 'oil10', 'welcome10', 'take10', 'jack', 'jake',
            'james20off', 'lpp15', 'brad', 'davis', 'mjbiz2023', 'blackfriday10',
            'danksggivingtubes', 'legends20', 'mjbizlastcall', '$100off',
            # Kits (not actual products)
            'sb-45d-kit', 'sb-25d-kit', 'sb-145d-kit', 'sb-15d-kit',
            # Special items
            'flexpack', 'bb-dml-000-00', '145d-blk-blk', 'bisonbotanics45d',
            'samples2023', 'samples2023-inactive', 'jake-inactive', 'replacement order-inactive',
            'every-other-label-free', 'free-application', 'single item discount', 
            'single line item discount', 'general discount', 'rist/howards',
            # Tier labels
            'diamond creative tier', 'silver creative tier', 'platinum creative tier'
        ]
        
        # Regex patterns for location/warehouse codes (STATE_COUNTY_CITY format)
        state_pattern = re.compile(r'^[A-Z]{2}_')  # Starts with 2-letter state code + underscore
        
        # Create exclusion mask
        item_lower = line_items_df['Item'].str.lower()
        item_upper = line_items_df['Item'].str.upper()
        
        # Pattern-based exclusion
        pattern_mask = item_lower.apply(
            lambda x: any(pattern in x for pattern in exclude_patterns)
        )
        
        # Exact match exclusion
        exact_mask = item_lower.isin([e.lower() for e in exclude_exact])
        
        # State/location code exclusion (e.g., "CA_LOS ANGELES_ZFYC")
        location_mask = item_upper.apply(lambda x: bool(state_pattern.match(x)))
        
        # Combine all exclusions
        exclude_mask = pattern_mask | exact_mask | location_mask
        
        # Keep only actual product line items
        excluded_count = exclude_mask.sum()
        line_items_df = line_items_df[~exclude_mask]
    
    # Clean numeric columns
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '':
            return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try:
            return float(cleaned)
        except:
            return 0
    
    if 'Item_Rate' in line_items_df.columns:
        line_items_df['Item_Rate'] = line_items_df['Item_Rate'].apply(clean_numeric)
    
    if 'Quantity' in line_items_df.columns:
        line_items_df['Quantity'] = line_items_df['Quantity'].apply(clean_numeric)
    
    # Calculate line total
    line_items_df['Line_Total'] = line_items_df['Quantity'] * line_items_df['Item_Rate']
    
    return line_items_df


def merge_orders_with_invoices(orders_df, invoices_df):
    """
    Merge sales orders with invoice data to get actual revenue
    
    Returns orders_df with Invoice_Amount added (actual invoiced revenue)
    Cadence still based on Order Start Date
    """
    
    if orders_df.empty:
        return orders_df
    
    if invoices_df.empty:
        # No invoices - fall back to order amounts
        orders_df['Invoice_Amount'] = orders_df['Amount']
        return orders_df
    
    # Clean SO numbers for matching - keep full format, uppercase for consistency
    orders_df['SO_Number_Clean'] = orders_df['SO_Number'].astype(str).str.strip().str.upper()
    
    # Aggregate invoice amounts by SO#
    invoices_df['SO_Number_Clean'] = invoices_df['SO_Number'].astype(str).str.strip().str.upper()
    invoice_totals = invoices_df.groupby('SO_Number_Clean')['Invoice_Amount'].sum().reset_index()
    
    # Merge
    merged = orders_df.merge(invoice_totals, on='SO_Number_Clean', how='left')
    
    # Fill missing invoice amounts with order amounts (for orders not yet invoiced)
    merged['Invoice_Amount'] = merged['Invoice_Amount'].fillna(merged['Amount'])
    
    return merged


def calculate_customer_metrics(historical_df):
    """
    Calculate metrics for each customer based on historical orders
    
    Returns DataFrame with:
    - Customer name
    - Total orders in 2025
    - Total revenue (from invoices)
    - Weighted avg order value (H2 weighted 1.25x)
    - Avg days between orders (cadence - based on order dates)
    - Last order date
    - Days since last order
    - Product types purchased
    - Confidence tier
    """
    
    if historical_df.empty:
        return pd.DataFrame()
    
    today = pd.Timestamp.now()
    
    # Determine which amount column to use (Invoice_Amount if available, else Amount)
    amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
    
    # Group by customer
    customer_metrics = []
    
    for customer in historical_df['Customer'].unique():
        cust_orders = historical_df[historical_df['Customer'] == customer].copy()
        cust_orders = cust_orders.sort_values('Order Start Date')
        
        # Basic metrics - use invoice amounts for revenue
        order_count = len(cust_orders)
        total_revenue = cust_orders[amount_col].sum()
        
        # Order dates for cadence calculation (still based on order dates, not invoice dates)
        order_dates = cust_orders['Order Start Date'].dropna().tolist()
        
        # Weighted average order value (H2 = 1.25x weight) - use invoice amounts
        weighted_sum = 0
        weight_total = 0
        for _, row in cust_orders.iterrows():
            order_date = row['Order Start Date']
            amount = row[amount_col]
            if pd.notna(order_date) and order_date.month >= 7:  # H2
                weight = 1.25
            else:  # H1
                weight = 1.0
            weighted_sum += amount * weight
            weight_total += weight
        
        weighted_avg = weighted_sum / weight_total if weight_total > 0 else 0
        
        # Cadence calculation (avg days between orders)
        cadence_days = None
        if len(order_dates) >= 2:
            gaps = []
            for i in range(len(order_dates) - 1):
                gap = (order_dates[i + 1] - order_dates[i]).days
                if gap > 0:  # Ignore same-day orders
                    gaps.append(gap)
            if gaps:
                cadence_days = sum(gaps) / len(gaps)
        
        # Last order info
        last_order_date = cust_orders['Order Start Date'].max()
        days_since_last = (today - last_order_date).days if pd.notna(last_order_date) else 999
        
        # Product types
        product_types = cust_orders['Order Type'].value_counts().to_dict()
        product_types_str = ', '.join([f"{k} ({v})" for k, v in product_types.items()])
        
        # Confidence tier
        if order_count >= 3:
            confidence_tier = 'Likely'
            confidence_pct = 0.75
        elif order_count >= 2:
            confidence_tier = 'Possible'
            confidence_pct = 0.50
        else:
            confidence_tier = 'Long Shot'
            confidence_pct = 0.25
        
        # Calculate expected orders in Q1 based on cadence
        # Q1 2026 = 90 days (Jan 1 - Mar 31)
        q1_days = 90
        if cadence_days and cadence_days > 0:
            expected_orders_q1 = q1_days / cadence_days
            # Cap at reasonable max (6 orders = roughly every 2 weeks)
            expected_orders_q1 = min(expected_orders_q1, 6.0)
            # Floor at 1 order minimum
            expected_orders_q1 = max(expected_orders_q1, 1.0)
        else:
            # No cadence data (only 1 order) - assume 1 order in Q1
            expected_orders_q1 = 1.0
        
        # Projected value = Avg Order × Expected Orders × Confidence %
        projected_value = weighted_avg * expected_orders_q1 * confidence_pct
        
        # Get rep name if available (for team view)
        rep_for_customer = cust_orders['Rep'].iloc[0] if 'Rep' in cust_orders.columns else ''
        
        # Get list of SO numbers for line item lookup
        so_numbers = []
        if 'SO_Number' in cust_orders.columns:
            so_numbers = cust_orders['SO_Number'].dropna().unique().tolist()
        
        customer_metrics.append({
            'Customer': customer,
            'Rep': rep_for_customer,
            'Order_Count': order_count,
            'Total_Revenue': total_revenue,
            'Weighted_Avg_Order': weighted_avg,
            'Cadence_Days': cadence_days,
            'Expected_Orders_Q1': expected_orders_q1,
            'Last_Order_Date': last_order_date,
            'Days_Since_Last': days_since_last,
            'Product_Types': product_types_str,
            'Product_Types_Dict': product_types,
            'Confidence_Tier': confidence_tier,
            'Confidence_Pct': confidence_pct,
            'Projected_Value': projected_value,
            'SO_Numbers': so_numbers
        })
    
    return pd.DataFrame(customer_metrics)


def calculate_customer_product_metrics(historical_df, line_items_df):
    """
    Calculate metrics by Customer + Product Type combination.
    This gives accurate cadence per product line, not per customer overall.
    
    Returns DataFrame with:
    - Customer, Product Type, Order count, Revenue, Cadence, Expected Q1 orders
    - Aggregated line item totals (qty, avg rate, total value)
    """
    
    if historical_df.empty:
        return pd.DataFrame()
    
    today = pd.Timestamp.now()
    amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
    
    metrics = []
    
    # Group by Customer + Product Type
    for (customer, product_type), group in historical_df.groupby(['Customer', 'Order Type']):
        group = group.sort_values('Order Start Date')
        
        # Basic metrics
        order_count = len(group)
        total_revenue = group[amount_col].sum()
        avg_order_value = total_revenue / order_count if order_count > 0 else 0
        
        # Cadence for THIS product type
        order_dates = group['Order Start Date'].dropna().tolist()
        cadence_days = None
        if len(order_dates) >= 2:
            gaps = []
            for i in range(len(order_dates) - 1):
                gap = (order_dates[i + 1] - order_dates[i]).days
                if gap > 0:
                    gaps.append(gap)
            if gaps:
                cadence_days = sum(gaps) / len(gaps)
        
        # Last order for this product type
        last_order_date = group['Order Start Date'].max()
        days_since_last = (today - last_order_date).days if pd.notna(last_order_date) else 999
        
        # Expected Q1 orders for this product type
        q1_days = 90
        if cadence_days and cadence_days > 0:
            expected_orders_q1 = q1_days / cadence_days
            expected_orders_q1 = min(expected_orders_q1, 6.0)
            expected_orders_q1 = max(expected_orders_q1, 1.0)
        else:
            expected_orders_q1 = 1.0
        
        # Confidence based on order count for THIS product type
        if order_count >= 3:
            confidence_tier = 'Likely'
            confidence_pct = 0.75
        elif order_count >= 2:
            confidence_tier = 'Possible'
            confidence_pct = 0.50
        else:
            confidence_tier = 'Long Shot'
            confidence_pct = 0.25
        
        # Get SO numbers for this customer + product type
        so_numbers = group['SO_Number'].dropna().unique().tolist() if 'SO_Number' in group.columns else []
        
        # Get line items for these SOs and aggregate
        total_qty = 0
        total_line_value = 0
        avg_rate = 0
        sku_count = 0
        
        if so_numbers and not line_items_df.empty:
            product_line_items = line_items_df[line_items_df['SO_Number'].isin(so_numbers)]
            if not product_line_items.empty:
                total_qty = int(product_line_items['Quantity'].sum())
                total_line_value = product_line_items['Line_Total'].sum()
                avg_rate = total_line_value / total_qty if total_qty > 0 else 0
                sku_count = product_line_items['Item'].nunique()
        
        # Calculate Q1 projection
        # Use line item data if available, otherwise use order amounts
        if total_qty > 0:
            avg_qty_per_order = total_qty / order_count
            q1_qty = int(round(avg_qty_per_order * expected_orders_q1))
            q1_value = q1_qty * avg_rate
        else:
            q1_qty = 0
            q1_value = avg_order_value * expected_orders_q1
        
        # Apply confidence
        q1_forecast = q1_value * confidence_pct
        
        # Rep
        rep = group['Rep'].iloc[0] if 'Rep' in group.columns else ''
        
        metrics.append({
            'Customer': customer,
            'Product_Type': product_type,
            'Rep': rep,
            'Order_Count': order_count,
            'Total_Revenue': total_revenue,
            'Avg_Order_Value': avg_order_value,
            'Cadence_Days': cadence_days,
            'Last_Order_Date': last_order_date,
            'Days_Since_Last': days_since_last,
            'Expected_Orders_Q1': expected_orders_q1,
            'Confidence_Tier': confidence_tier,
            'Confidence_Pct': confidence_pct,
            'SO_Numbers': so_numbers,
            'Total_Qty_2025': total_qty,
            'Avg_Rate': avg_rate,
            'SKU_Count': sku_count,
            'Q1_Qty': q1_qty,
            'Q1_Value': q1_value,
            'Q1_Forecast': q1_forecast
        })
    
    return pd.DataFrame(metrics)


def identify_reorder_opportunities(customer_metrics_df, pending_customers, pipeline_customers):
    """
    Filter out customers who already have pending orders or pipeline deals
    
    Args:
        customer_metrics_df: DataFrame from calculate_customer_metrics()
        pending_customers: Set of customer names with pending NS orders
        pipeline_customers: Set of customer names in Q1 HubSpot pipeline
    
    Returns:
        DataFrame with only customers who are reorder opportunities
    """
    
    if customer_metrics_df.empty:
        return customer_metrics_df
    
    # Normalize customer names for matching
    def normalize(name):
        return str(name).lower().strip()
    
    pending_normalized = {normalize(c) for c in pending_customers}
    pipeline_normalized = {normalize(c) for c in pipeline_customers}
    active_customers = pending_normalized | pipeline_normalized
    
    # Filter out active customers
    opportunities_df = customer_metrics_df[
        ~customer_metrics_df['Customer'].apply(normalize).isin(active_customers)
    ].copy()
    
    return opportunities_df


def get_customer_line_items(so_numbers, line_items_df):
    """
    Get aggregated line items for a customer based on their SO numbers
    
    Groups by Item and sums quantities, calculates weighted average rate
    
    Returns DataFrame with columns: Item, Total_Qty, Avg_Rate, Total_Value
    """
    
    if not so_numbers or line_items_df.empty:
        return pd.DataFrame()
    
    # Clean SO numbers for matching - keep full format (e.g., "SO13778")
    so_numbers_clean = [str(so).strip().upper() for so in so_numbers if str(so).strip()]
    
    if not so_numbers_clean:
        return pd.DataFrame()
    
    # Filter line items to customer's SO numbers
    customer_items = line_items_df[line_items_df['SO_Number'].isin(so_numbers_clean)].copy()
    
    if customer_items.empty:
        return pd.DataFrame()
    
    # Aggregate by Item - sum quantities, weighted average rate
    aggregated = customer_items.groupby('Item').agg({
        'Quantity': 'sum',
        'Item_Rate': 'mean',  # Average rate across orders
        'Line_Total': 'sum'
    }).reset_index()
    
    aggregated.columns = ['Item', 'Total_Qty', 'Avg_Rate', 'Total_Value']
    
    # Sort by total value descending
    aggregated = aggregated.sort_values('Total_Value', ascending=False)
    
    return aggregated


def get_product_type_summary(historical_df, opportunities_df):
    """
    Summarize reorder opportunities by product type
    
    Returns dict with:
    {
        'FlexPack': {
            'customers': ['AYR Wellness', 'Curaleaf'],
            'historical_total': 150000,
            'projected_total': 75000,
            'order_count': 15
        },
        ...
    }
    """
    
    if historical_df.empty or opportunities_df.empty:
        return {}
    
    # Get list of opportunity customers
    opp_customers = set(opportunities_df['Customer'].tolist())
    
    # Filter historical to only opportunity customers
    opp_historical = historical_df[historical_df['Customer'].isin(opp_customers)]
    
    if opp_historical.empty:
        return {}
    
    # Group by product type
    product_summary = {}
    
    for product_type in opp_historical['Order Type'].unique():
        prod_orders = opp_historical[opp_historical['Order Type'] == product_type]
        
        # Get unique customers for this product type
        prod_customers = prod_orders['Customer'].unique().tolist()
        
        # Calculate totals
        historical_total = prod_orders['Amount'].sum()
        order_count = len(prod_orders)
        
        # Calculate projected based on customer confidence levels
        projected_total = 0
        for customer in prod_customers:
            cust_metrics = opportunities_df[opportunities_df['Customer'] == customer]
            if not cust_metrics.empty:
                conf_pct = cust_metrics['Confidence_Pct'].iloc[0]
                cust_prod_avg = prod_orders[prod_orders['Customer'] == customer]['Amount'].mean()
                projected_total += cust_prod_avg * conf_pct
        
        product_summary[product_type] = {
            'customers': prod_customers,
            'historical_total': historical_total,
            'projected_total': projected_total,
            'order_count': order_count
        }
    
    # Sort by projected total descending
    product_summary = dict(sorted(product_summary.items(), key=lambda x: x[1]['projected_total'], reverse=True))
    
    return product_summary


# ========== MAIN FUNCTION ==========
def main():
    """Main function for Q1 2026 Forecasting module"""
    
    inject_custom_css()
    
    # === HEADER / HERO SECTION ===
    days_until_q1 = calculate_business_days_until_q1()
    
    st.markdown("""
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="font-size: 2.8rem; font-weight: 800; background: linear-gradient(to right, #10b981, #3b82f6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin: 0;">
                Q1 2026 FORECAST
            </h1>
            <p style="color: #94a3b8; font-size: 1.1rem; margin-top: 10px;">Strategic Planning & Revenue Projection</p>
        </div>
    """, unsafe_allow_html=True)
    
    # Hero Metrics Grid
    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1:
        st.markdown(f"""
        <div class="hero-metric">
            <div class="hero-label">Timeline</div>
            <div class="hero-value">Jan 1 - Mar 31</div>
            <div style="color: #64748b; font-size: 0.8rem; margin-top: 5px;">2026 Fiscal Quarter</div>
        </div>
        """, unsafe_allow_html=True)
    with col_h2:
        st.markdown(f"""
        <div class="hero-metric" style="border-left-color: #10b981;">
            <div class="hero-label">Countdown</div>
            <div class="hero-value">{days_until_q1} Days</div>
            <div style="color: #64748b; font-size: 0.8rem; margin-top: 5px;">Business days remaining</div>
        </div>
        """, unsafe_allow_html=True)
    with col_h3:
        st.markdown(f"""
        <div class="hero-metric" style="border-left-color: #f59e0b;">
            <div class="hero-label">Last Sync</div>
            <div class="hero-value">{get_mst_time().strftime('%I:%M %p')}</div>
            <div style="color: #64748b; font-size: 0.8rem; margin-top: 5px;">Mountain Standard Time</div>
        </div>
        """, unsafe_allow_html=True)
    
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
    
    # Define the team reps for "All Reps" aggregate view
    TEAM_REPS = ['Alex Gonzalez', 'Jake Lynch', 'Dave Borkowski', 'Lance Mitton', 'Shopify E-commerce', 'Brad Sherman']
    
    # Add "All Reps" option at the beginning
    rep_options = ["👥 All Reps (Team View)"] + reps
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 1: WHO ARE YOU?
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("### 👋 Step 1: Let's Get Started")
    
    # Rep selector
    selected_option = st.selectbox("Who are you?", options=rep_options, key="q1_rep_selector")
    
    # Determine if we're in team view mode
    is_team_view = selected_option == "👥 All Reps (Team View)"
    
    if is_team_view:
        rep_name = "All Reps"
        first_name = "Team"
        active_team_reps = [r for r in TEAM_REPS if r in reps]
        st.markdown(f"""
        <div style="background: rgba(59, 130, 246, 0.1); border-left: 4px solid #3b82f6; padding: 15px; border-radius: 8px; margin: 10px 0;">
            <div style="font-size: 1.1rem;">📊 <strong>Team View Active</strong></div>
            <div style="color: #94a3b8; margin-top: 5px;">Showing combined data for: {', '.join(active_team_reps)}</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        rep_name = selected_option
        first_name = rep_name.split()[0]  # Get first name
        active_team_reps = [rep_name]
        
        # Personalized greeting
        st.markdown(f"""
        <div style="background: rgba(16, 185, 129, 0.1); border-left: 4px solid #10b981; padding: 15px; border-radius: 8px; margin: 10px 0;">
            <div style="font-size: 1.2rem;">👋 <strong>Hey {first_name}!</strong> Let's build out your Q1 2026 forecast.</div>
            <div style="color: #94a3b8; margin-top: 5px;">I'll walk you through this step by step. First, let's set your quota.</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 2: SET YOUR GOAL
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown(f"### 🎯 Step 2: {'Set Team Goal' if is_team_view else f'{first_name}, Set Your Q1 Quota'}")
    
    goal_key = f"q1_goal_{rep_name}"
    if goal_key not in st.session_state:
        st.session_state[goal_key] = 5000000 if is_team_view else 1000000
    
    team_prompt = "What's the team target" if is_team_view else "What are you committing to"
    st.markdown(f"*{team_prompt} for Q1 2026?*")
    
    col1, col2 = st.columns([2, 1])
    with col1:
        q1_goal = st.number_input(
            "Q1 2026 Quota ($)",
            min_value=0,
            max_value=50000000,
            value=st.session_state[goal_key],
            step=50000,
            format="%d",
            key=f"q1_goal_input_{rep_name}",
            label_visibility="collapsed"
        )
        st.session_state[goal_key] = q1_goal
    
    with col2:
        st.metric("🎯 Q1 Goal", f"${q1_goal:,.0f}")
    
    # Confirmation message
    if q1_goal > 0:
        st.markdown(f"""
        <div style="color: #10b981; font-size: 0.95rem; margin-top: 5px;">
            ✅ {'Team is' if is_team_view else f"{first_name}, you're"} targeting <strong>${q1_goal:,.0f}</strong> for Q1 2026. Let's build the plan to get there!
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # === GET Q1 2026 DATA ===
    # The main dashboard's "spillover" buckets ARE the Q1 2026 scheduled orders!
    # - pf_spillover = PF orders with Q1 2026 Promise/Projected dates
    # - pa_spillover = PA orders with PA Date in Q1 2026
    
    # Aggregate data from all active reps
    all_pf_spillover = []
    all_pa_spillover = []
    total_pf_amount = 0
    total_pa_amount = 0
    
    for r in active_team_reps:
        so_cats = categorize_sales_orders(sales_orders_df, r)
        if not so_cats['pf_spillover'].empty:
            all_pf_spillover.append(so_cats['pf_spillover'])
            total_pf_amount += so_cats['pf_spillover_amount']
        if not so_cats['pa_spillover'].empty:
            all_pa_spillover.append(so_cats['pa_spillover'])
            total_pa_amount += so_cats['pa_spillover_amount']
    
    # Combine into single dataframes
    combined_pf = pd.concat(all_pf_spillover, ignore_index=True) if all_pf_spillover else pd.DataFrame()
    combined_pa = pd.concat(all_pa_spillover, ignore_index=True) if all_pa_spillover else pd.DataFrame()
    
    # ═══════════════════════════════════════════════════════════════════════════
    # PRODUCT TYPE FILTER (GLOBAL)
    # ═══════════════════════════════════════════════════════════════════════════
    
    # Collect all product types from NetSuite data
    all_product_types = set()
    if not combined_pf.empty and 'Order Type' in combined_pf.columns:
        all_product_types.update(combined_pf['Order Type'].dropna().unique())
    if not combined_pa.empty and 'Order Type' in combined_pa.columns:
        all_product_types.update(combined_pa['Order Type'].dropna().unique())
    
    # Also check deals_df for HubSpot product types
    if not deals_df.empty and 'Product Type' in deals_df.columns:
        all_product_types.update(deals_df['Product Type'].dropna().unique())
    
    # Clean and sort product types
    all_product_types = sorted([pt for pt in all_product_types if pt and str(pt).strip() not in ['', 'nan', 'None']])
    
    # Product type filter in sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🏷️ Filter by Product Type")
    
    filter_key = f"product_type_filter_{rep_name}"
    if filter_key not in st.session_state:
        st.session_state[filter_key] = all_product_types  # Default: all selected
    
    # Select all / none buttons
    fcol1, fcol2 = st.sidebar.columns(2)
    with fcol1:
        if st.button("Select All", key=f"pt_all_{rep_name}", use_container_width=True):
            st.session_state[filter_key] = all_product_types
            st.rerun()
    with fcol2:
        if st.button("Clear All", key=f"pt_clear_{rep_name}", use_container_width=True):
            st.session_state[filter_key] = []
            st.rerun()
    
    selected_product_types = st.sidebar.multiselect(
        "Show product types:",
        options=all_product_types,
        default=st.session_state[filter_key],
        key=f"pt_multiselect_{rep_name}"
    )
    st.session_state[filter_key] = selected_product_types
    
    if len(selected_product_types) < len(all_product_types):
        st.sidebar.caption(f"Showing {len(selected_product_types)} of {len(all_product_types)} product types")
    
    # Apply product type filter to NetSuite data
    if selected_product_types:
        if not combined_pf.empty and 'Order Type' in combined_pf.columns:
            combined_pf = combined_pf[combined_pf['Order Type'].isin(selected_product_types)]
        if not combined_pa.empty and 'Order Type' in combined_pa.columns:
            combined_pa = combined_pa[combined_pa['Order Type'].isin(selected_product_types)]
    
    # Recalculate amounts after filtering
    total_pf_amount = combined_pf['Amount'].sum() if not combined_pf.empty and 'Amount' in combined_pf.columns else 0
    total_pa_amount = combined_pa['Amount'].sum() if not combined_pa.empty and 'Amount' in combined_pa.columns else 0
    
    # Map spillover to Q1 categories
    ns_categories = {
        'PF_Spillover': {'label': '📦 PF (Q1 2026 Date)', 'df': combined_pf, 'amount': total_pf_amount},
        'PA_Spillover': {'label': '⏳ PA (Q1 2026 PA Date)', 'df': combined_pa, 'amount': total_pa_amount},
    }
    
    # Format for display
    ns_dfs = {
        'PF_Spillover': format_ns_view(combined_pf, 'Promise'),
        'PA_Spillover': format_ns_view(combined_pa, 'PA_Date'),
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
        # Filter to active team reps (supports both single rep and team view)
        rep_deals = deals_df[deals_df['Deal Owner'].isin(active_team_reps)].copy()
        
        # Apply product type filter to HubSpot deals
        if selected_product_types and 'Product Type' in rep_deals.columns:
            rep_deals = rep_deals[rep_deals['Product Type'].isin(selected_product_types)]
        
        if 'Close Date' in rep_deals.columns:
            # Q1 2026 Close Date deals (Close Date in Q1 2026)
            q1_close_mask = (rep_deals['Close Date'] >= Q1_2026_START) & (rep_deals['Close Date'] <= Q1_2026_END)
            q1_deals = rep_deals[q1_close_mask]
            
            # Q4 2025 Spillover - deals with Q4 close date BUT marked as Q1 2026 Spillover
            # IMPORTANT: Only include deals with Q4 close dates to avoid double counting with Q1 deals
            q4_close_mask = (rep_deals['Close Date'] >= Q4_2025_START) & (rep_deals['Close Date'] <= Q4_2025_END)
            
            if 'Q1 2026 Spillover' in rep_deals.columns:
                # Q4 Spillover = Q4 close date AND spillover flag is set
                q4_spillover = rep_deals[q4_close_mask & (rep_deals['Q1 2026 Spillover'] == 'Q1 2026')]
            else:
                q4_spillover = pd.DataFrame()
            
            # Debug info
            with st.expander("🔧 Debug: HubSpot Deal Counts"):
                if is_team_view:
                    st.write(f"**Team View - Reps included:** {', '.join(active_team_reps)}")
                st.write(f"**Total deals loaded:** {len(rep_deals)}")
                st.write(f"**Q1 Close Date deals:** {len(q1_deals)} (Close Date in Jan-Mar 2026)")
                st.write(f"**Q4 Spillover deals:** {len(q4_spillover)} (Q4 Close Date + Spillover flag)")
                if 'Amount' in rep_deals.columns:
                    q1_total = q1_deals['Amount'].sum() if not q1_deals.empty else 0
                    q4_spill_total = q4_spillover['Amount'].sum() if not q4_spillover.empty else 0
                    st.write(f"**Q1 deals total:** ${q1_total:,.0f}")
                    st.write(f"**Q4 spillover total:** ${q4_spill_total:,.0f}")
                    st.write(f"**Combined total:** ${q1_total + q4_spill_total:,.0f}")
            
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
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 3: BUILD YOUR FORECAST - CURRENT PIPELINE
    # ═══════════════════════════════════════════════════════════════════════════
    
    step3_title = "Review Pipeline" if is_team_view else f"{first_name}, Let's Review Your Pipeline"
    st.markdown(f"### 📊 Step 3: {step3_title}")
    
    st.markdown(f"""
    <div style="background: rgba(59, 130, 246, 0.1); border-left: 4px solid #3b82f6; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-size: 1rem;">
            {"Here's what the team has" if is_team_view else "Here's what you've got"} in the pipeline for Q1. Check the boxes to include them in your forecast.
        </div>
        <div style="color: #94a3b8; margin-top: 5px; font-size: 0.9rem;">
            💡 <strong>Tip:</strong> NetSuite orders are already confirmed. HubSpot deals are your opportunities to close.
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    export_buckets = {}
    
    # === CLEAR ALL SELECTIONS BUTTON (top right) ===
    clear_col1, clear_col2 = st.columns([3, 1])
    with clear_col2:
        if st.button("🗑️ Reset", key=f"q1_clear_all_{rep_name}"):
            for key in ns_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
                st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            for key in hs_categories.keys():
                st.session_state[f"q1_chk_{key}_{rep_name}"] = False
                st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            st.rerun()
    
    # === SELECT ALL / UNSELECT ALL ===
    sel_col1, sel_col2, sel_col3 = st.columns([1, 1, 2])
    with sel_col1:
        if st.button("☑️ Select All Pipeline", key=f"q1_select_all_{rep_name}", use_container_width=True):
            for key in ns_categories.keys():
                if ns_categories[key]['amount'] > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
                    st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            for key in hs_categories.keys():
                df = hs_dfs.get(key, pd.DataFrame())
                val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
                if val > 0:
                    st.session_state[f"q1_chk_{key}_{rep_name}"] = True
                    st.session_state[f"q1_unselected_{key}_{rep_name}"] = set()
            st.rerun()
    
    with sel_col2:
        if st.button("☐ Clear Pipeline", key=f"q1_unselect_all_{rep_name}", use_container_width=True):
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
            st.markdown("#### 📦 Confirmed Orders (NetSuite)")
            st.caption("These are locked in - spillover orders shipping in Q1")
            
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
            st.markdown("#### 🎯 Open Deals (HubSpot)")
            st.caption("Your opportunities - close these to hit your number!")
            
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
    
    # ═══════════════════════════════════════════════════════════════════════════
    # PIPELINE PRODUCT TYPE BREAKDOWN
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    st.markdown("#### 📊 Pipeline by Product Type")
    
    # Combine all NS data for breakdown
    all_ns_data = pd.concat([combined_pf, combined_pa], ignore_index=True) if not combined_pf.empty or not combined_pa.empty else pd.DataFrame()
    
    # Combine all HS data for breakdown
    all_hs_dfs = [df for df in hs_dfs.values() if not df.empty]
    all_hs_data = pd.concat(all_hs_dfs, ignore_index=True) if all_hs_dfs else pd.DataFrame()
    
    pt_col1, pt_col2 = st.columns(2)
    
    with pt_col1:
        st.markdown("**📦 NetSuite by Order Type**")
        if not all_ns_data.empty and 'Order Type' in all_ns_data.columns:
            ns_by_type = all_ns_data.groupby('Order Type').agg({
                'Amount': ['sum', 'count']
            }).reset_index()
            ns_by_type.columns = ['Order Type', 'Amount', 'Orders']
            ns_by_type = ns_by_type.sort_values('Amount', ascending=False)
            
            # Display as nice table
            for _, row in ns_by_type.iterrows():
                pct = (row['Amount'] / ns_by_type['Amount'].sum() * 100) if ns_by_type['Amount'].sum() > 0 else 0
                st.markdown(f"""
                <div style="display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <span>{row['Order Type']}</span>
                    <span style="color: #10b981;"><strong>${row['Amount']:,.0f}</strong> ({int(row['Orders'])} orders)</span>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style="display: flex; justify-content: space-between; padding: 8px 0; margin-top: 5px; border-top: 2px solid rgba(255,255,255,0.2);">
                <span><strong>Total</strong></span>
                <span style="color: #10b981;"><strong>${ns_by_type['Amount'].sum():,.0f}</strong></span>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.caption("No NetSuite orders in selected product types")
    
    with pt_col2:
        st.markdown("**🎯 HubSpot by Product Type**")
        if not all_hs_data.empty and 'Product Type' in all_hs_data.columns:
            # Use Amount_Numeric for HubSpot
            amount_col = 'Amount_Numeric' if 'Amount_Numeric' in all_hs_data.columns else 'Amount'
            hs_by_type = all_hs_data.groupby('Product Type').agg({
                amount_col: ['sum', 'count']
            }).reset_index()
            hs_by_type.columns = ['Product Type', 'Amount', 'Deals']
            hs_by_type = hs_by_type.sort_values('Amount', ascending=False)
            
            for _, row in hs_by_type.iterrows():
                st.markdown(f"""
                <div style="display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <span>{row['Product Type']}</span>
                    <span style="color: #3b82f6;"><strong>${row['Amount']:,.0f}</strong> ({int(row['Deals'])} deals)</span>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown(f"""
            <div style="display: flex; justify-content: space-between; padding: 8px 0; margin-top: 5px; border-top: 2px solid rgba(255,255,255,0.2);">
                <span><strong>Total</strong></span>
                <span style="color: #3b82f6;"><strong>${hs_by_type['Amount'].sum():,.0f}</strong></span>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.caption("No HubSpot deals in selected product types")
    
    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 3: REORDER FORECAST (Historical Analysis)
    # ═══════════════════════════════════════════════════════════════════════════
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 4: REORDER OPPORTUNITIES
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    st.markdown(f"### 🔄 Step 4: {'Team Reorder Opportunities' if is_team_view else f'{first_name}, Find Your Reorder Opportunities'}")
    
    st.markdown(f"""
    <div style="background: rgba(245, 158, 11, 0.1); border-left: 4px solid #f59e0b; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-size: 1rem;">
            {"These are customers the team served" if is_team_view else "These are your customers from"} 2025 who <strong>don't have pending orders or active deals</strong>. 
            They're likely to reorder — let's figure out how much!
        </div>
        <div style="color: #94a3b8; margin-top: 8px; font-size: 0.9rem;">
            <strong>How it works:</strong><br>
            • Grouped by <strong>Product Type</strong> so cadence is accurate (not mixing Jars with Flex Pkg)<br>
            • <strong>🟢 Likely</strong> = 3+ orders (75% confidence) | <strong>🟡 Possible</strong> = 2 orders (50%) | <strong>⚪ Long Shot</strong> = 1 order (25%)<br>
            • Edit the Q1 Value column if you know better — you're the expert on your accounts!
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize reorder buckets
    reorder_buckets = {}
    
    # Load all data
    with st.spinner("Loading historical data and line items..."):
        # Load historical orders
        if is_team_view:
            all_historical = []
            all_invoices = []
            for r in active_team_reps:
                rep_hist = load_historical_orders(main_dash, r)
                rep_inv = load_invoices(main_dash, r)
                if not rep_hist.empty:
                    rep_hist['Rep'] = r
                    all_historical.append(rep_hist)
                if not rep_inv.empty:
                    all_invoices.append(rep_inv)
            historical_df = pd.concat(all_historical, ignore_index=True) if all_historical else pd.DataFrame()
            invoices_df = pd.concat(all_invoices, ignore_index=True) if all_invoices else pd.DataFrame()
        else:
            historical_df = load_historical_orders(main_dash, rep_name)
            invoices_df = load_invoices(main_dash, rep_name)
            if not historical_df.empty:
                historical_df['Rep'] = rep_name
        
        # Merge with invoices for accurate revenue
        if not historical_df.empty:
            historical_df = merge_orders_with_invoices(historical_df, invoices_df)
        
        # Apply product type filter to historical data
        if selected_product_types and not historical_df.empty and 'Order Type' in historical_df.columns:
            historical_df = historical_df[historical_df['Order Type'].isin(selected_product_types)]
        
        # Load line items - THIS IS THE KEY DATA
        line_items_df = load_line_items(main_dash)
    
    # Debug section - EXPANDED
    with st.expander("🔧 Debug: Data Loading Status", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Historical Orders (NS Sales Orders):**")
            if historical_df.empty:
                st.error("❌ No historical orders loaded")
            else:
                st.success(f"✅ {len(historical_df)} orders loaded")
                st.write(f"Columns: {historical_df.columns.tolist()}")
                if 'SO_Number' in historical_df.columns:
                    sample_sos = historical_df['SO_Number'].dropna().head(10).tolist()
                    st.write(f"**Sample SO Numbers:** {sample_sos}")
                    st.write(f"**Unique SOs:** {historical_df['SO_Number'].nunique()}")
                else:
                    st.error("❌ SO_Number column MISSING from historical_df!")
        
        with col2:
            st.write("**Line Items (Sales Order Line Item):**")
            if line_items_df.empty:
                st.error("❌ No line items loaded - check tab name 'Sales Order Line Item'")
            else:
                st.success(f"✅ {len(line_items_df)} line items loaded")
                st.write(f"Columns: {line_items_df.columns.tolist()}")
                if 'SO_Number' in line_items_df.columns:
                    sample_sos = line_items_df['SO_Number'].dropna().head(10).tolist()
                    st.write(f"**Sample SO Numbers:** {sample_sos}")
                    st.write(f"**Unique SOs:** {line_items_df['SO_Number'].nunique()}")
                else:
                    st.error("❌ SO_Number column MISSING!")
                
                if 'Item' in line_items_df.columns:
                    st.write(f"**Sample Items:** {line_items_df['Item'].head(5).tolist()}")
                if 'Quantity' in line_items_df.columns:
                    st.write(f"**Sample Qty:** {line_items_df['Quantity'].head(5).tolist()}")
                if 'Item_Rate' in line_items_df.columns:
                    st.write(f"**Sample Rates:** {line_items_df['Item_Rate'].head(5).tolist()}")
        
        # Test matching
        if not historical_df.empty and not line_items_df.empty:
            if 'SO_Number' in historical_df.columns and 'SO_Number' in line_items_df.columns:
                hist_sos = set(historical_df['SO_Number'].dropna().unique())
                line_sos = set(line_items_df['SO_Number'].dropna().unique())
                matching = hist_sos.intersection(line_sos)
                st.write(f"**SO Number Matching Test:**")
                st.write(f"- Historical unique SOs: {len(hist_sos)}")
                st.write(f"- Line Item unique SOs: {len(line_sos)}")
                st.write(f"- **Matching SOs: {len(matching)}**")
                if len(matching) == 0:
                    st.error("❌ NO MATCHING SO NUMBERS! Check format - Historical: " + 
                             str(list(hist_sos)[:3]) + " vs Line Items: " + str(list(line_sos)[:3]))
                else:
                    st.success(f"✅ {len(matching)} SOs match between datasets")
                    st.write(f"Sample matches: {list(matching)[:5]}")
    
    if historical_df.empty:
        st.info("No 2025 historical orders found for this rep")
    elif line_items_df.empty:
        st.warning("⚠️ Line item data not available. Please check the 'Sales Order Line Item' tab in your spreadsheet.")
    else:
        # Calculate customer metrics (old method - for exclusion logic)
        customer_metrics_df = calculate_customer_metrics(historical_df)
        
        # Exclude customers with pending orders or pipeline deals
        pending_customers = set()
        for key in ns_categories.keys():
            df = ns_dfs.get(key, pd.DataFrame())
            if not df.empty and 'Customer' in df.columns:
                pending_customers.update(df['Customer'].dropna().tolist())
        
        pipeline_customers = set()
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            if not df.empty and 'Deal Name' in df.columns:
                pipeline_customers.update(df['Deal Name'].dropna().tolist())
        
        # Get list of customers to exclude
        def normalize(name): 
            return str(name).lower().strip()
        excluded_customers = {normalize(c) for c in pending_customers | pipeline_customers}
        
        # Calculate NEW product-level metrics
        product_metrics_df = calculate_customer_product_metrics(historical_df, line_items_df)
        
        if product_metrics_df.empty:
            st.warning("No product metrics calculated")
        else:
            # Filter out customers with pending orders/deals
            product_metrics_df['Customer_Normalized'] = product_metrics_df['Customer'].apply(normalize)
            opportunities_df = product_metrics_df[
                ~product_metrics_df['Customer_Normalized'].isin(excluded_customers)
            ].copy()
            
            if opportunities_df.empty:
                st.success("✅ All 2025 customers already have pending orders or pipeline deals!")
            else:
                # Show 2025 summary
                st.markdown("#### 📊 2025 Performance Summary")
                amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Invoiced Revenue", f"${historical_df[amount_col].sum():,.0f}")
                with c2:
                    st.metric("Customers", f"{opportunities_df['Customer'].nunique()}")
                with c3:
                    st.metric("Product Lines", f"{len(opportunities_df)}")
                with c4:
                    st.metric("Projected Q1", f"${opportunities_df['Q1_Forecast'].sum():,.0f}")
                
                st.markdown("---")
                
                # === PRODUCT TYPE BASED FORECASTING ===
                st.markdown(f"#### {'Pick Your Reorder Targets' if is_team_view else f'{first_name}, Pick Your Reorder Targets'}")
                st.markdown("""
                <div style="color: #94a3b8; margin-bottom: 15px; font-size: 0.95rem;">
                    Start with the <strong style="color: #10b981;">🟢 Likely</strong> ones — these are your best bets. 
                    Then add <strong style="color: #f59e0b;">🟡 Possible</strong> customers you're confident about.
                    <strong style="color: #94a3b8;">⚪ Long Shots</strong> are bonus if you can close them!
                </div>
                """, unsafe_allow_html=True)
                
                # Group by confidence tier
                tiers = [
                    ('Likely', '🟢', 0.75, 'Your best customers — 3+ orders in 2025. Include these!'),
                    ('Possible', '🟡', 0.50, 'Solid customers — 2 orders in 2025. Good upside here.'),
                    ('Long Shot', '⚪', 0.25, 'One-time buyers — worth a call to see if they\'ll reorder.')
                ]
                
                for tier_name, emoji, conf_pct, tier_desc in tiers:
                    tier_data = opportunities_df[opportunities_df['Confidence_Tier'] == tier_name].copy()
                    
                    if tier_data.empty:
                        continue
                    
                    # Calculate tier totals
                    tier_historical = tier_data['Total_Revenue'].sum()
                    tier_projected = tier_data['Q1_Forecast'].sum()
                    tier_rows = len(tier_data)
                    tier_customers = tier_data['Customer'].nunique()
                    
                    # Tier header checkbox with friendlier text
                    checkbox_key = f"q1_reorder_{tier_name}_{rep_name}"
                    
                    # Different prompts per tier
                    if tier_name == 'Likely':
                        checkbox_label = f"{emoji} **{tier_name}** — {tier_customers} customers who WILL reorder (${tier_projected:,.0f} projected)"
                    elif tier_name == 'Possible':
                        checkbox_label = f"{emoji} **{tier_name}** — {tier_customers} customers who might reorder (${tier_projected:,.0f} projected)"
                    else:
                        checkbox_label = f"{emoji} **{tier_name}** — {tier_customers} worth a follow-up (${tier_projected:,.0f} projected)"
                    
                    is_checked = st.checkbox(
                        checkbox_label,
                        key=checkbox_key,
                        help=tier_desc
                    )
                    
                    if is_checked:
                        with st.expander(f"📋 {tier_name} - Review & Edit", expanded=True):
                            st.caption(f"💡 {tier_desc}")
                            
                            # Session state for this tier
                            edited_key = f"q1_products_{tier_name}_{rep_name}"
                            if edited_key not in st.session_state:
                                st.session_state[edited_key] = {}
                            
                            # Build display dataframe
                            display_data = []
                            for _, row in tier_data.iterrows():
                                key = f"{row['Customer']}|{row['Product_Type']}"
                                
                                # Use edited value if available
                                if key in st.session_state[edited_key]:
                                    q1_value = st.session_state[edited_key][key]
                                else:
                                    q1_value = int(row['Q1_Value'])
                                
                                # Format cadence
                                cadence = row['Cadence_Days']
                                if pd.notna(cadence) and cadence > 0:
                                    cadence_str = f"Every {int(cadence)}d"
                                    days_since = row['Days_Since_Last']
                                    if days_since > cadence * 1.5:
                                        status = f"🔴 {int(days_since - cadence)}d overdue"
                                    elif days_since > cadence:
                                        status = "🟡 Due now"
                                    elif days_since > cadence * 0.75:
                                        status = "🟢 Due soon"
                                    else:
                                        status = "⚪ On track"
                                else:
                                    cadence_str = "1 order"
                                    status = "⚪ New"
                                
                                display_data.append({
                                    'Select': True,
                                    'Customer': row['Customer'],
                                    'Product Type': row['Product_Type'],
                                    '2025 #': int(row['Order_Count']),
                                    '2025 $': int(row['Total_Revenue']),
                                    'Cadence': cadence_str,
                                    'Q1 Est': round(row['Expected_Orders_Q1'], 1),
                                    'Status': status,
                                    'Q1 Value': q1_value,
                                    'SKUs': int(row['SKU_Count']) if row['SKU_Count'] > 0 else 0
                                })
                            
                            display_df = pd.DataFrame(display_data)
                            
                            # Editable data editor
                            edited_df = st.data_editor(
                                display_df,
                                column_config={
                                    "Select": st.column_config.CheckboxColumn("✓", width="small", help="Include in forecast"),
                                    "Customer": st.column_config.TextColumn("Customer", width="large"),
                                    "Product Type": st.column_config.TextColumn("Product", width="medium"),
                                    "2025 #": st.column_config.NumberColumn("2025 #", format="%d", width="small", help="Orders in 2025"),
                                    "2025 $": st.column_config.NumberColumn("2025 $", format="$%d", width="small", help="Revenue in 2025"),
                                    "Cadence": st.column_config.TextColumn("Cadence", width="small"),
                                    "Q1 Est": st.column_config.NumberColumn("Q1 Est", format="%.1f", width="small", help="Expected orders in Q1"),
                                    "Status": st.column_config.TextColumn("Status", width="medium"),
                                    "Q1 Value": st.column_config.NumberColumn("Q1 Value ✏️", format="$%d", width="small", help="EDIT: Your Q1 forecast for this product line"),
                                    "SKUs": st.column_config.NumberColumn("SKUs", format="%d", width="small", help="Unique items in this product type")
                                },
                                disabled=['Customer', 'Product Type', '2025 #', '2025 $', 'Cadence', 'Q1 Est', 'Status', 'SKUs'],
                                hide_index=True,
                                use_container_width=True,
                                key=f"q1_product_editor_{tier_name}_{rep_name}",
                                height=min(500, 50 + len(display_df) * 35)
                            )
                            
                            # Store edited values and calculate totals
                            selected_total = 0
                            customer_forecasts = {}
                            
                            for _, row in edited_df.iterrows():
                                key = f"{row['Customer']}|{row['Product Type']}"
                                q1_val = int(row['Q1 Value']) if pd.notna(row['Q1 Value']) else 0
                                st.session_state[edited_key][key] = q1_val
                                
                                if row['Select']:
                                    if row['Customer'] not in customer_forecasts:
                                        customer_forecasts[row['Customer']] = 0
                                    customer_forecasts[row['Customer']] += q1_val
                                    selected_total += q1_val
                            
                            # Apply confidence percentage
                            forecast_with_conf = selected_total * conf_pct
                            
                            # Summary
                            selected_rows = len([r for _, r in edited_df.iterrows() if r['Select']])
                            st.markdown(f"""
                            <div style="background: rgba(16, 185, 129, 0.1); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 8px; padding: 12px; margin-top: 10px;">
                                <div style="font-size: 1.1rem; font-weight: 600;">
                                    {tier_name} Forecast: <span style="color: #10b981;">${selected_total:,.0f}</span> × {int(conf_pct*100)}% = 
                                    <span style="color: #10b981; font-size: 1.3rem;">${forecast_with_conf:,.0f}</span>
                                </div>
                                <div style="font-size: 0.8rem; color: #94a3b8; margin-top: 5px;">
                                    {len(customer_forecasts)} customers • {selected_rows} product lines selected
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            # === SKU DRILL-DOWN SECTION ===
                            st.markdown("---")
                            with st.expander("🔍 View SKU Details by Customer × Product", expanded=False):
                                st.caption("Select a customer and product type to see individual SKUs")
                                
                                # Get unique customer-product combinations
                                sku_options = tier_data[['Customer', 'Product_Type', 'SO_Numbers']].copy()
                                
                                # Customer selector
                                customers_in_tier = sorted(tier_data['Customer'].unique().tolist())
                                selected_customer = st.selectbox(
                                    "Customer",
                                    options=customers_in_tier,
                                    key=f"sku_cust_{tier_name}_{rep_name}"
                                )
                                
                                if selected_customer:
                                    # Product type selector for this customer
                                    cust_products = tier_data[tier_data['Customer'] == selected_customer]['Product_Type'].unique().tolist()
                                    selected_product = st.selectbox(
                                        "Product Type",
                                        options=cust_products,
                                        key=f"sku_prod_{tier_name}_{rep_name}"
                                    )
                                    
                                    if selected_product:
                                        # Get SO numbers for this customer + product
                                        match_row = tier_data[
                                            (tier_data['Customer'] == selected_customer) & 
                                            (tier_data['Product_Type'] == selected_product)
                                        ]
                                        
                                        if not match_row.empty:
                                            so_numbers = match_row.iloc[0]['SO_Numbers']
                                            
                                            if so_numbers and len(so_numbers) > 0:
                                                # Get line items for these SOs
                                                sku_items = line_items_df[line_items_df['SO_Number'].isin(so_numbers)].copy()
                                                
                                                if not sku_items.empty:
                                                    # Aggregate by Item (SKU)
                                                    sku_agg = sku_items.groupby('Item').agg({
                                                        'Quantity': 'sum',
                                                        'Item_Rate': 'mean',
                                                        'Line_Total': 'sum',
                                                        'SO_Number': 'nunique'
                                                    }).reset_index()
                                                    sku_agg.columns = ['SKU', 'Total Qty', 'Avg Rate', 'Total $', 'Orders']
                                                    sku_agg = sku_agg.sort_values('Total $', ascending=False)
                                                    
                                                    # Format for display
                                                    sku_agg['Total Qty'] = sku_agg['Total Qty'].astype(int)
                                                    sku_agg['Total $'] = sku_agg['Total $'].astype(int)
                                                    sku_agg['Avg Rate'] = sku_agg['Avg Rate'].round(2)
                                                    
                                                    # Show summary
                                                    st.markdown(f"**{selected_customer} - {selected_product}**: {len(sku_agg)} unique SKUs from {len(so_numbers)} orders")
                                                    
                                                    # Display SKU table
                                                    st.dataframe(
                                                        sku_agg,
                                                        column_config={
                                                            "SKU": st.column_config.TextColumn("SKU", width="large"),
                                                            "Total Qty": st.column_config.NumberColumn("2025 Qty", format="%d"),
                                                            "Avg Rate": st.column_config.NumberColumn("Avg Rate", format="$%.2f"),
                                                            "Total $": st.column_config.NumberColumn("2025 $", format="$%d"),
                                                            "Orders": st.column_config.NumberColumn("Orders", format="%d", help="Number of orders containing this SKU")
                                                        },
                                                        hide_index=True,
                                                        use_container_width=True,
                                                        height=min(400, 40 + len(sku_agg) * 35)
                                                    )
                                                else:
                                                    st.info("No SKU details found for this combination")
                                            else:
                                                st.info("No sales order numbers found")
                            
                            st.markdown("---")
                            
                            # Build export data
                            export_data = []
                            for cust, forecast in customer_forecasts.items():
                                export_data.append({
                                    'Customer': cust,
                                    'Confidence_Tier': tier_name,
                                    'Confidence_Pct': conf_pct,
                                    'Q1_Value': forecast,
                                    'Projected_Value': forecast * conf_pct
                                })
                            
                            if export_data:
                                reorder_buckets[f"reorder_{tier_name}"] = pd.DataFrame(export_data)
    
    # ═══════════════════════════════════════════════════════════════════════════
    # REORDER BY PRODUCT TYPE BREAKDOWN
    # ═══════════════════════════════════════════════════════════════════════════
    
    # Check if opportunities_df exists and has data
    try:
        if not opportunities_df.empty:
            st.markdown("---")
            st.markdown("#### 📊 Reorder Forecast by Product Type")
            
            # Group by Product Type
            reorder_by_type = opportunities_df.groupby('Product_Type').agg({
                'Total_Revenue': 'sum',
                'Q1_Forecast': 'sum',
                'Customer': 'nunique'
            }).reset_index()
            reorder_by_type.columns = ['Product Type', '2025 Revenue', 'Q1 Forecast', 'Customers']
            reorder_by_type = reorder_by_type.sort_values('Q1 Forecast', ascending=False)
            
            # Display as table
            st.dataframe(
                reorder_by_type,
                column_config={
                    "Product Type": st.column_config.TextColumn("Product Type", width="medium"),
                    "2025 Revenue": st.column_config.NumberColumn("2025 Revenue", format="$%d"),
                    "Q1 Forecast": st.column_config.NumberColumn("Q1 Forecast", format="$%d"),
                    "Customers": st.column_config.NumberColumn("Customers", format="%d")
                },
                hide_index=True,
                use_container_width=True
            )
            
            st.caption(f"Total Q1 Reorder Forecast: ${reorder_by_type['Q1 Forecast'].sum():,.0f} (before confidence weighting)")
    except NameError:
        pass  # opportunities_df not defined
    
    # === CALCULATE RESULTS ===
    def safe_sum(df):
        if df.empty:
            return 0
        if 'Amount_Numeric' in df.columns:
            return df['Amount_Numeric'].sum()
        elif 'Amount' in df.columns:
            return df['Amount'].sum()
        return 0
    
    def safe_sum_projected(df):
        """Sum projected values for reorder buckets"""
        if df.empty:
            return 0
        if 'Projected_Value' in df.columns:
            return df['Projected_Value'].sum()
        return 0
    
    selected_scheduled = sum(safe_sum(df) for k, df in export_buckets.items() if k in ns_categories)
    selected_pipeline = sum(safe_sum(df) for k, df in export_buckets.items() if k in hs_categories)
    
    # Calculate reorder forecast total
    selected_reorder = 0
    if reorder_buckets:
        selected_reorder = sum(safe_sum_projected(df) for df in reorder_buckets.values())
    
    total_forecast = selected_scheduled + selected_pipeline + selected_reorder
    gap_to_goal = q1_goal - total_forecast
    
    # === STICKY SUMMARY BAR (HUD STYLE) ===
    gap_class = "val-gap-behind" if gap_to_goal > 0 else "val-gap-ahead"
    gap_label = "GAP" if gap_to_goal > 0 else "AHEAD"
    gap_display = f"${abs(gap_to_goal):,.0f}"
    
    st.markdown(f"""
    <div class="sticky-forecast-bar-q1">
        <div class="sticky-item">
            <div class="sticky-label">Scheduled</div>
            <div class="sticky-val val-sched">${selected_scheduled:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item">
            <div class="sticky-label">Pipeline</div>
            <div class="sticky-val val-pipe">${selected_pipeline:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item">
            <div class="sticky-label">Reorder</div>
            <div class="sticky-val val-reorder">${selected_reorder:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item">
            <div class="sticky-label">Total Forecast</div>
            <div class="sticky-val val-total">${total_forecast:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item">
            <div class="sticky-label">{gap_label}</div>
            <div class="sticky-val {gap_class}">{gap_display}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 5: YOUR FORECAST SUMMARY
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    step5_title = "Team Forecast Summary" if is_team_view else f"{first_name}, Here's Your Q1 Forecast!"
    st.markdown(f"### 🎉 Step 5: {step5_title}")
    
    # Personalized summary message
    pct_of_goal = (total_forecast / q1_goal * 100) if q1_goal > 0 else 0
    
    subject = "The team is" if is_team_view else "You're"
    subject_has = "The team has" if is_team_view else "You've got"
    we_you = "We can" if is_team_view else "You can"
    
    if gap_to_goal <= 0:
        summary_message = f"🎉 {subject} <strong style='color: #10b981;'>${abs(gap_to_goal):,.0f} AHEAD</strong> of the ${q1_goal:,.0f} goal! Nice work!"
        summary_bg = "rgba(16, 185, 129, 0.1)"
        summary_border = "#10b981"
    elif pct_of_goal >= 75:
        summary_message = f"💪 {subject} at <strong>{pct_of_goal:.0f}%</strong> of goal — just <strong style='color: #f59e0b;'>${gap_to_goal:,.0f}</strong> to go. {we_you} close this gap!"
        summary_bg = "rgba(245, 158, 11, 0.1)"
        summary_border = "#f59e0b"
    else:
        summary_message = f"📊 {subject_has} <strong style='color: #3b82f6;'>${total_forecast:,.0f}</strong> forecasted — need <strong style='color: #ef4444;'>${gap_to_goal:,.0f}</strong> more to hit ${q1_goal:,.0f}. Let's find more opportunities!"
        summary_bg = "rgba(239, 68, 68, 0.1)"
        summary_border = "#ef4444"
    
    st.markdown(f"""
    <div style="background: {summary_bg}; border-left: 4px solid {summary_border}; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-size: 1.1rem;">{summary_message}</div>
    </div>
    """, unsafe_allow_html=True)
    
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.metric("📦 Scheduled", f"${selected_scheduled:,.0f}", help="Confirmed NetSuite orders")
    with m2:
        st.metric("🎯 Pipeline", f"${selected_pipeline:,.0f}", help="HubSpot deals to close")
    with m3:
        st.metric("🔄 Reorder", f"${selected_reorder:,.0f}", help="Historical customer opportunities")
    with m4:
        st.metric("🏁 Total Forecast", f"${total_forecast:,.0f}")
    with m5:
        if gap_to_goal > 0:
            st.metric("Gap to Goal", f"${gap_to_goal:,.0f}", delta="Behind", delta_color="inverse")
        else:
            st.metric("Ahead of Goal", f"${abs(gap_to_goal):,.0f}", delta="Ahead!", delta_color="normal")
    
    # Gauge with glass card
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        fig = create_q1_gauge(total_forecast, q1_goal, "Q1 2026 Progress to Goal")
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("#### 📊 The Breakdown")
        st.markdown(f"""
        **📦 Confirmed Orders:** ${selected_scheduled:,.0f}
        - Already in NetSuite, shipping Q1
        
        **🎯 Pipeline Deals:** ${selected_pipeline:,.0f}
        - {'Work these deals!' if not is_team_view else 'Deals to close in Q1'}
        
        **🔄 Reorder Potential:** ${selected_reorder:,.0f}
        - 🟢 Likely (75%) | 🟡 Possible (50%) | ⚪ Long Shot (25%)
        
        **🎯 Q1 Goal:** ${q1_goal:,.0f}
        """)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # === EXPORT SECTION ===
    st.markdown("---")
    st.markdown('<div class="section-header">📤 Export Q1 2026 Forecast</div>', unsafe_allow_html=True)
    
    if export_buckets or reorder_buckets:
        # Combine all selected data for export
        all_ns_data = []
        all_hs_data = []
        all_reorder_data = []
        
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
        
        # Add reorder bucket data
        if reorder_buckets:
            for key, df in reorder_buckets.items():
                if df.empty:
                    continue
                export_df = df.copy()
                export_df['Category'] = key.replace('reorder_', '').replace('_', ' ').title()
                all_reorder_data.append(export_df)
        
        col1, col2, col3 = st.columns(3)
        
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
        
        with col3:
            if all_reorder_data:
                reorder_export = pd.concat(all_reorder_data, ignore_index=True)
                csv_reorder = reorder_export.to_csv(index=False)
                st.download_button(
                    label="📥 Download Reorder Prospects (CSV)",
                    data=csv_reorder,
                    file_name=f"q1_2026_reorder_prospects_{rep_name.replace(' ', '_')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                reorder_total = reorder_export['Projected_Value'].sum() if 'Projected_Value' in reorder_export.columns else 0
                st.caption(f"{len(reorder_export)} prospects, ${reorder_total:,.0f} projected")
            else:
                st.info("No reorder prospects selected")
    else:
        st.info("Select items above to enable export")
    
    # === DEBUG INFO ===
    with st.expander("🔧 Debug: Data Summary"):
        st.write("**Data Source:** Copy of All Reps All Pipelines (Q4 2025 + Q1 2026 deals)")
        if is_team_view:
            st.write(f"**Team Reps:** {', '.join(active_team_reps)}")
        st.write(f"**Total Deals Loaded:** {len(deals_df)}")
        st.write(f"**PF Spillover:** {len(combined_pf)} orders, ${total_pf_amount:,.0f}")
        st.write(f"**PA Spillover:** {len(combined_pa)} orders, ${total_pa_amount:,.0f}")
        
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
            st.write(f"**{key}:** {len(df)} deals, ${val:,.0f}")


# Run if called directly
if __name__ == "__main__":
    main()
