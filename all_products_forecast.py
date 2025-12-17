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
    
    # Also grab the SO# (Document Number) for matching with invoices - Column B
    if len(col_names) > 1:
        so_col = col_names[1]
        if so_col in historical_df.columns:
            historical_df['SO_Number'] = historical_df[so_col].astype(str).str.strip()
    
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
    
    # Clean SO_Number for matching
    if 'SO_Number' in invoice_df.columns:
        invoice_df['SO_Number'] = invoice_df['SO_Number'].astype(str).str.strip()
        # Extract just the SO number (might be formatted like "Sales Order #12345")
        invoice_df['SO_Number'] = invoice_df['SO_Number'].str.extract(r'(\d+)', expand=False).fillna('')
    
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
    
    # Clean SO_Number
    if 'SO_Number' in line_items_df.columns:
        line_items_df['SO_Number'] = line_items_df['SO_Number'].astype(str).str.strip()
        # Extract just the number
        line_items_df['SO_Number'] = line_items_df['SO_Number'].str.extract(r'(\d+)', expand=False).fillna('')
        line_items_df = line_items_df[line_items_df['SO_Number'] != '']
    
    # Clean Item
    if 'Item' in line_items_df.columns:
        line_items_df['Item'] = line_items_df['Item'].astype(str).str.strip()
        line_items_df = line_items_df[line_items_df['Item'] != '']
        line_items_df = line_items_df[line_items_df['Item'].str.lower() != 'nan']
    
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
    
    # Clean SO numbers for matching
    orders_df['SO_Number_Clean'] = orders_df['SO_Number'].astype(str).str.extract(r'(\d+)', expand=False).fillna('')
    
    # Aggregate invoice amounts by SO#
    invoice_totals = invoices_df.groupby('SO_Number')['Invoice_Amount'].sum().reset_index()
    invoice_totals.columns = ['SO_Number_Clean', 'Invoice_Amount']
    
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
    
    # Clean SO numbers for matching
    so_numbers_clean = [str(so).strip() for so in so_numbers]
    # Extract just digits
    so_numbers_clean = [s if s.isdigit() else ''.join(filter(str.isdigit, s)) for s in so_numbers_clean]
    so_numbers_clean = [s for s in so_numbers_clean if s]  # Remove empty
    
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
    
    # Define the team reps for "All Reps" aggregate view
    TEAM_REPS = ['Alex Gonzalez', 'Jake Lynch', 'Dave Borkowski', 'Lance Mitton', 'Shopify E-commerce', 'Brad Sherman']
    
    # Add "All Reps" option at the beginning
    rep_options = ["👥 All Reps (Team View)"] + reps
    
    # Rep selector
    selected_option = st.selectbox("Select Sales Rep:", options=rep_options, key="q1_rep_selector")
    
    # Determine if we're in team view mode
    is_team_view = selected_option == "👥 All Reps (Team View)"
    
    if is_team_view:
        rep_name = "All Reps"
        # Filter to only team reps that exist in the data
        active_team_reps = [r for r in TEAM_REPS if r in reps]
        st.info(f"📊 Team View: Showing aggregate data for {len(active_team_reps)} reps: {', '.join(active_team_reps)}")
    else:
        rep_name = selected_option
        active_team_reps = [rep_name]  # Single rep
    
    # === USER-DEFINED GOAL INPUT ===
    st.markdown("### 🎯 Set Your Q1 2026 Goal")
    
    goal_key = f"q1_goal_{rep_name}"
    if goal_key not in st.session_state:
        # Default goal: higher for team view
        st.session_state[goal_key] = 5000000 if is_team_view else 1000000
    
    col1, col2 = st.columns([2, 1])
    with col1:
        q1_goal = st.number_input(
            "Enter your Q1 2026 quota/goal ($):" if not is_team_view else "Enter Team Q1 2026 quota/goal ($):",
            min_value=0,
            max_value=50000000,
            value=st.session_state[goal_key],
            step=50000,
            format="%d",
            key=f"q1_goal_input_{rep_name}"
        )
        st.session_state[goal_key] = q1_goal
    
    with col2:
        st.metric("Team Q1 Goal" if is_team_view else "Your Q1 Goal", f"${q1_goal:,.0f}")
    
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
    
    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 3: REORDER FORECAST (Historical Analysis)
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    st.markdown("### 🔄 Reorder Forecast (Historical Analysis)")
    st.caption("Customers who ordered in 2025 but have no pending orders or Q1 pipeline deals. Revenue from actual invoices.")
    
    # Initialize reorder buckets (will be populated if historical data exists)
    reorder_buckets = {}
    
    # Load historical data - aggregate if team view
    with st.spinner("Analyzing 2025 order history and invoices..."):
        if is_team_view:
            # Aggregate historical data from all team reps
            all_historical = []
            all_invoices = []
            for r in active_team_reps:
                rep_historical = load_historical_orders(main_dash, r)
                rep_invoices = load_invoices(main_dash, r)
                if not rep_historical.empty:
                    rep_historical['Rep'] = r  # Add rep column for reference
                    all_historical.append(rep_historical)
                if not rep_invoices.empty:
                    all_invoices.append(rep_invoices)
            
            historical_df = pd.concat(all_historical, ignore_index=True) if all_historical else pd.DataFrame()
            invoices_df = pd.concat(all_invoices, ignore_index=True) if all_invoices else pd.DataFrame()
        else:
            historical_df = load_historical_orders(main_dash, rep_name)
            invoices_df = load_invoices(main_dash, rep_name)
            if not historical_df.empty:
                historical_df['Rep'] = rep_name
        
        # Merge orders with invoices to get actual revenue
        if not historical_df.empty:
            historical_df = merge_orders_with_invoices(historical_df, invoices_df)
        
        # Load line items for detailed forecasting (loaded once, filtered per customer later)
        line_items_df = load_line_items(main_dash)
    
    if historical_df.empty:
        st.info("No 2025 historical orders found for this rep")
    else:
        # Calculate customer metrics (now uses Invoice_Amount for revenue)
        customer_metrics_df = calculate_customer_metrics(historical_df)
        
        # Get list of customers with pending orders (from Section 1)
        pending_customers = set()
        for key in ns_categories.keys():
            df = ns_dfs.get(key, pd.DataFrame())
            if not df.empty and 'Customer' in df.columns:
                pending_customers.update(df['Customer'].dropna().tolist())
        
        # Get list of customers in Q1 pipeline (from Section 2)
        pipeline_customers = set()
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            if not df.empty and 'Deal Name' in df.columns:
                # Extract customer name from deal name (often format: "Customer - Product")
                pipeline_customers.update(df['Deal Name'].dropna().tolist())
        
        # Identify reorder opportunities
        opportunities_df = identify_reorder_opportunities(customer_metrics_df, pending_customers, pipeline_customers)
        
        if opportunities_df.empty:
            st.success("✅ All 2025 customers already have pending orders or pipeline deals!")
        else:
            # 2025 Performance Summary - Use Invoice amounts for actual revenue
            st.markdown("#### 📊 Your 2025 Performance (Invoiced Revenue)")
            amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
            perf_col1, perf_col2, perf_col3, perf_col4 = st.columns(4)
            with perf_col1:
                st.metric("Total Revenue", f"${historical_df[amount_col].sum():,.0f}")
            with perf_col2:
                st.metric("Unique Customers", f"{historical_df['Customer'].nunique()}")
            with perf_col3:
                st.metric("Total Orders", f"{len(historical_df)}")
            with perf_col4:
                avg_order = historical_df[amount_col].mean()
                st.metric("Avg Order", f"${avg_order:,.0f}")
            
            st.markdown("---")
            
            # Customer Filter (Multi-select) - Only affects Section 3
            all_opportunity_customers = sorted(opportunities_df['Customer'].unique().tolist())
            
            filter_col1, filter_col2 = st.columns([3, 1])
            with filter_col1:
                selected_customers = st.multiselect(
                    "🔍 Filter by Customer (Section 3 only):",
                    options=all_opportunity_customers,
                    default=[],
                    placeholder="All customers (click to filter)",
                    key=f"q1_customer_filter_{rep_name}"
                )
            with filter_col2:
                if st.button("Clear Filter", key=f"q1_clear_filter_{rep_name}"):
                    st.session_state[f"q1_customer_filter_{rep_name}"] = []
                    st.rerun()
            
            # Apply customer filter
            if selected_customers:
                filtered_opportunities = opportunities_df[opportunities_df['Customer'].isin(selected_customers)]
                filtered_historical = historical_df[historical_df['Customer'].isin(selected_customers)]
                st.info(f"Showing {len(selected_customers)} selected customer(s)")
            else:
                filtered_opportunities = opportunities_df
                filtered_historical = historical_df
            
            # Get product type summary
            product_summary = get_product_type_summary(filtered_historical, filtered_opportunities)
            
            # View Toggle
            view_mode = st.radio(
                "View by:",
                options=["Confidence Tier", "Product Type"],
                horizontal=True,
                key=f"q1_view_mode_{rep_name}"
            )
            
            st.caption("💡 **Forecast** = Avg Order × Expected Q1 Orders × Confidence %. Enable **Customize** to edit values or expand **Edit Line Items** to refine by item, quantity, and rate.")
            
            st.markdown("---")
            
            # Initialize reorder buckets for tracking selections
            reorder_buckets = {}
            
            if view_mode == "Confidence Tier":
                # === VIEW BY CONFIDENCE TIER ===
                
                tiers = [
                    ('Likely', '🔴', 0.75, 'Orders 3+ times in 2025 - very likely to reorder'),
                    ('Possible', '🟡', 0.50, 'Ordered 2 times in 2025 - good chance of reorder'),
                    ('Long Shot', '⚪', 0.25, 'Ordered once in 2025 - worth reaching out')
                ]
                
                for tier_name, tier_emoji, tier_pct, tier_desc in tiers:
                    tier_customers = filtered_opportunities[filtered_opportunities['Confidence_Tier'] == tier_name]
                    
                    if tier_customers.empty:
                        continue
                    
                    tier_historical = tier_customers['Total_Revenue'].sum()
                    tier_projected = tier_customers['Projected_Value'].sum()
                    
                    checkbox_key = f"q1_reorder_{tier_name}_{rep_name}"
                    
                    is_checked = st.checkbox(
                        f"{tier_emoji} {tier_name} ({int(tier_pct*100)}%): {len(tier_customers)} customers  |  ${tier_historical:,.0f} hist  →  ${tier_projected:,.0f} projected",
                        key=checkbox_key,
                        help=tier_desc
                    )
                    
                    if is_checked:
                        with st.expander(f"🔎 View {tier_name} Customers"):
                            # Customize toggle
                            enable_edit = st.toggle("Customize", key=f"q1_reorder_tgl_{tier_name}_{rep_name}")
                            
                            if enable_edit:
                                # Session state for unselected customers
                                unselected_key = f"q1_reorder_unsel_{tier_name}_{rep_name}"
                                if unselected_key not in st.session_state:
                                    st.session_state[unselected_key] = set()
                                
                                # Session state for edited projected values
                                edited_proj_key = f"q1_reorder_edited_proj_{tier_name}_{rep_name}"
                                if edited_proj_key not in st.session_state:
                                    st.session_state[edited_proj_key] = {}
                                
                                # Row-level select/unselect buttons
                                row_col1, row_col2, row_col3 = st.columns([1, 1, 2])
                                with row_col1:
                                    if st.button("☑️ All", key=f"q1_reorder_sel_{tier_name}_{rep_name}"):
                                        st.session_state[unselected_key] = set()
                                        st.rerun()
                                with row_col2:
                                    if st.button("☐ None", key=f"q1_reorder_unsel_btn_{tier_name}_{rep_name}"):
                                        st.session_state[unselected_key] = set(tier_customers['Customer'].tolist())
                                        st.rerun()
                                
                                # Build display dataframe
                                display_data = []
                                for _, row in tier_customers.iterrows():
                                    is_selected = row['Customer'] not in st.session_state[unselected_key]
                                    # Use edited projected value if available
                                    proj_val = st.session_state[edited_proj_key].get(row['Customer'], row['Projected_Value'])
                                    row_data = {
                                        'Select': is_selected,
                                        'Customer': row['Customer'],
                                    }
                                    # Add Rep column for team view
                                    if is_team_view and 'Rep' in row and row['Rep']:
                                        row_data['Rep'] = row['Rep']
                                    row_data.update({
                                        '2025 Orders': row['Order_Count'],
                                        'Avg Order': row['Weighted_Avg_Order'],
                                        'Cadence': f"~{int(row['Cadence_Days'])}d" if pd.notna(row['Cadence_Days']) else "N/A",
                                        'Exp Q1': row['Expected_Orders_Q1'] if 'Expected_Orders_Q1' in row else 1.0,
                                        'Last Order': row['Last_Order_Date'].strftime('%Y-%m-%d') if pd.notna(row['Last_Order_Date']) else '',
                                        'Days Ago': row['Days_Since_Last'],
                                        'Forecast': proj_val
                                    })
                                    display_data.append(row_data)
                                
                                display_df = pd.DataFrame(display_data)
                                
                                # Build column config
                                col_config = {
                                    "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                    "Customer": st.column_config.TextColumn("Customer", width="medium"),
                                }
                                disabled_cols = ['Customer', '2025 Orders', 'Avg Order', 'Cadence', 'Exp Q1', 'Last Order', 'Days Ago']
                                if is_team_view and 'Rep' in display_df.columns:
                                    col_config["Rep"] = st.column_config.TextColumn("Rep", width="small")
                                    disabled_cols.append('Rep')
                                col_config.update({
                                    "2025 Orders": st.column_config.NumberColumn("Orders", width="small"),
                                    "Avg Order": st.column_config.NumberColumn("Avg Order", format="$%d"),
                                    "Cadence": st.column_config.TextColumn("Cadence", width="small"),
                                    "Exp Q1": st.column_config.NumberColumn("Exp Q1", format="%.1f", help="Expected orders in Q1 based on cadence"),
                                    "Last Order": st.column_config.TextColumn("Last Order", width="small"),
                                    "Days Ago": st.column_config.NumberColumn("Days Ago", width="small"),
                                    "Forecast": st.column_config.NumberColumn("Forecast $", format="$%d", help="Edit this value to adjust forecast")
                                })
                                
                                edited = st.data_editor(
                                    display_df,
                                    column_config=col_config,
                                    disabled=disabled_cols,
                                    hide_index=True,
                                    key=f"q1_reorder_edit_{tier_name}_{rep_name}",
                                    use_container_width=True
                                )
                                
                                # Update unselected set and edited projections
                                current_unselected = set()
                                for _, row in edited.iterrows():
                                    if not row['Select']:
                                        current_unselected.add(row['Customer'])
                                    # Store edited forecast values
                                    st.session_state[edited_proj_key][row['Customer']] = row['Forecast']
                                st.session_state[unselected_key] = current_unselected
                                
                                # Calculate selected total using EDITED values
                                selected_total = 0
                                selected_count = 0
                                for _, row in edited.iterrows():
                                    if row['Select']:
                                        selected_total += row['Forecast']
                                        selected_count += 1
                                st.caption(f"Selected: {selected_count} customers, ${selected_total:,.0f} forecast")
                                
                                # === LINE ITEM EDITING SECTION ===
                                selected_customers_list = [row['Customer'] for _, row in edited.iterrows() if row['Select']]
                                
                                if selected_customers_list and not line_items_df.empty:
                                    with st.expander("📋 Edit Line Items (Qty & Rate)"):
                                        st.caption("Adjust quantities and rates to refine your forecast. Item names are read-only.")
                                        
                                        # Session state for edited line items
                                        line_items_key = f"q1_reorder_line_items_{tier_name}_{rep_name}"
                                        if line_items_key not in st.session_state:
                                            st.session_state[line_items_key] = {}
                                        
                                        # Collect line items for all selected customers
                                        all_line_items = []
                                        for cust in selected_customers_list:
                                            cust_row = tier_customers[tier_customers['Customer'] == cust]
                                            if not cust_row.empty:
                                                so_numbers = cust_row.iloc[0].get('SO_Numbers', [])
                                                if so_numbers:
                                                    cust_items = get_customer_line_items(so_numbers, line_items_df)
                                                    if not cust_items.empty:
                                                        cust_items['Customer'] = cust
                                                        all_line_items.append(cust_items)
                                        
                                        if all_line_items:
                                            combined_items = pd.concat(all_line_items, ignore_index=True)
                                            
                                            # Build editable display with persisted edits
                                            line_display = []
                                            for _, item_row in combined_items.iterrows():
                                                item_key = f"{item_row['Customer']}|{item_row['Item']}"
                                                
                                                # Get edited values or defaults
                                                if item_key in st.session_state[line_items_key]:
                                                    qty = st.session_state[line_items_key][item_key]['Qty']
                                                    rate = st.session_state[line_items_key][item_key]['Rate']
                                                else:
                                                    # Default: proportional qty for Q1 based on expected orders
                                                    cust_row = tier_customers[tier_customers['Customer'] == item_row['Customer']]
                                                    exp_q1 = cust_row.iloc[0].get('Expected_Orders_Q1', 1.0) if not cust_row.empty else 1.0
                                                    # Scale by expected orders vs historical orders
                                                    order_count = cust_row.iloc[0].get('Order_Count', 1) if not cust_row.empty else 1
                                                    scale_factor = exp_q1 / max(order_count, 1)
                                                    qty = item_row['Total_Qty'] * scale_factor
                                                    rate = item_row['Avg_Rate']
                                                
                                                line_display.append({
                                                    'Customer': item_row['Customer'],
                                                    'Item': item_row['Item'],
                                                    'Hist Qty': item_row['Total_Qty'],
                                                    'Q1 Qty': qty,
                                                    'Rate': rate,
                                                    'Line Total': qty * rate
                                                })
                                            
                                            line_df = pd.DataFrame(line_display)
                                            
                                            edited_lines = st.data_editor(
                                                line_df,
                                                column_config={
                                                    "Customer": st.column_config.TextColumn("Customer", width="medium"),
                                                    "Item": st.column_config.TextColumn("Item", width="medium"),
                                                    "Hist Qty": st.column_config.NumberColumn("2025 Qty", format="%d", help="Historical quantity ordered in 2025"),
                                                    "Q1 Qty": st.column_config.NumberColumn("Q1 Qty", format="%.0f", help="Edit: Forecasted quantity for Q1"),
                                                    "Rate": st.column_config.NumberColumn("Rate $", format="$%.2f", help="Edit: Price per unit"),
                                                    "Line Total": st.column_config.NumberColumn("Total $", format="$%,.0f")
                                                },
                                                disabled=['Customer', 'Item', 'Hist Qty', 'Line Total'],
                                                hide_index=True,
                                                key=f"q1_line_edit_{tier_name}_{rep_name}",
                                                use_container_width=True
                                            )
                                            
                                            # Store edited values and recalculate
                                            line_item_total = 0
                                            customer_totals = {}
                                            for _, line_row in edited_lines.iterrows():
                                                item_key = f"{line_row['Customer']}|{line_row['Item']}"
                                                st.session_state[line_items_key][item_key] = {
                                                    'Qty': line_row['Q1 Qty'],
                                                    'Rate': line_row['Rate']
                                                }
                                                line_total = line_row['Q1 Qty'] * line_row['Rate']
                                                line_item_total += line_total
                                                
                                                # Accumulate by customer
                                                if line_row['Customer'] not in customer_totals:
                                                    customer_totals[line_row['Customer']] = 0
                                                customer_totals[line_row['Customer']] += line_total
                                            
                                            # Apply confidence percentage to line item total
                                            line_item_forecast = line_item_total * tier_pct
                                            st.success(f"📊 Line Item Total: ${line_item_total:,.0f} × {int(tier_pct*100)}% confidence = **${line_item_forecast:,.0f}** forecast")
                                            
                                            # Update individual customer forecasts based on line items
                                            for cust, cust_total in customer_totals.items():
                                                st.session_state[edited_proj_key][cust] = cust_total * tier_pct
                                        else:
                                            st.info("No line item detail available for selected customers")
                                
                                # Store for export with edited values
                                export_df = tier_customers[tier_customers['Customer'].isin(selected_customers_list)].copy()
                                # Update with edited forecast values
                                for idx, row in export_df.iterrows():
                                    if row['Customer'] in st.session_state[edited_proj_key]:
                                        export_df.loc[idx, 'Projected_Value'] = st.session_state[edited_proj_key][row['Customer']]
                                reorder_buckets[f"reorder_{tier_name}"] = export_df
                            else:
                                # Read-only view - shows calculated values, enable Customize to edit
                                display_data = []
                                for _, row in tier_customers.iterrows():
                                    row_data = {'Customer': row['Customer']}
                                    if is_team_view and 'Rep' in row and row['Rep']:
                                        row_data['Rep'] = row['Rep']
                                    row_data.update({
                                        '2025 Orders': row['Order_Count'],
                                        'Avg Order': row['Weighted_Avg_Order'],
                                        'Cadence': f"~{int(row['Cadence_Days'])}d" if pd.notna(row['Cadence_Days']) else "N/A",
                                        'Exp Q1': row['Expected_Orders_Q1'] if 'Expected_Orders_Q1' in row else 1.0,
                                        'Last Order': row['Last_Order_Date'].strftime('%Y-%m-%d') if pd.notna(row['Last_Order_Date']) else '',
                                        'Days Ago': row['Days_Since_Last'],
                                        'Forecast': row['Projected_Value']
                                    })
                                    display_data.append(row_data)
                                
                                display_df = pd.DataFrame(display_data)
                                col_config = {"Customer": st.column_config.TextColumn("Customer", width="medium")}
                                if is_team_view and 'Rep' in display_df.columns:
                                    col_config["Rep"] = st.column_config.TextColumn("Rep", width="small")
                                col_config.update({
                                    "2025 Orders": st.column_config.NumberColumn("Orders", width="small"),
                                    "Avg Order": st.column_config.NumberColumn("Avg Order", format="$%d"),
                                    "Cadence": st.column_config.TextColumn("Cadence", width="small"),
                                    "Exp Q1": st.column_config.NumberColumn("Exp Q1", format="%.1f"),
                                    "Last Order": st.column_config.TextColumn("Last Order", width="small"),
                                    "Days Ago": st.column_config.NumberColumn("Days Ago", width="small"),
                                    "Forecast": st.column_config.NumberColumn("Forecast $", format="$%d")
                                })
                                
                                st.dataframe(
                                    display_df,
                                    column_config=col_config,
                                    hide_index=True,
                                    use_container_width=True
                                )
                                st.caption("💡 Enable **Customize** to edit forecast values")
                                
                                # Store full tier for export
                                reorder_buckets[f"reorder_{tier_name}"] = tier_customers
            
            else:
                # === VIEW BY PRODUCT TYPE ===
                
                for product_type, data in product_summary.items():
                    checkbox_key = f"q1_reorder_prod_{product_type}_{rep_name}"
                    
                    is_checked = st.checkbox(
                        f"📦 {product_type}: {len(data['customers'])} customers  |  ${data['historical_total']:,.0f} hist  →  ${data['projected_total']:,.0f} projected",
                        key=checkbox_key
                    )
                    
                    if is_checked:
                        with st.expander(f"🔎 View {product_type} Customers"):
                            # Get customers for this product type
                            prod_customers = data['customers']
                            prod_opportunities = filtered_opportunities[filtered_opportunities['Customer'].isin(prod_customers)]
                            
                            # Customize toggle
                            enable_edit = st.toggle("Customize", key=f"q1_reorder_prod_tgl_{product_type}_{rep_name}")
                            
                            if enable_edit:
                                # Session state for unselected customers
                                unselected_key = f"q1_reorder_prod_unsel_{product_type}_{rep_name}"
                                if unselected_key not in st.session_state:
                                    st.session_state[unselected_key] = set()
                                
                                # Session state for edited projected values
                                edited_proj_key = f"q1_reorder_prod_edited_proj_{product_type}_{rep_name}"
                                if edited_proj_key not in st.session_state:
                                    st.session_state[edited_proj_key] = {}
                                
                                # Row-level buttons
                                row_col1, row_col2, row_col3 = st.columns([1, 1, 2])
                                with row_col1:
                                    if st.button("☑️ All", key=f"q1_reorder_prod_sel_{product_type}_{rep_name}"):
                                        st.session_state[unselected_key] = set()
                                        st.rerun()
                                with row_col2:
                                    if st.button("☐ None", key=f"q1_reorder_prod_unsel_btn_{product_type}_{rep_name}"):
                                        st.session_state[unselected_key] = set(prod_customers)
                                        st.rerun()
                                
                                # Build display dataframe with cadence and editable forecast
                                display_data = []
                                for _, row in prod_opportunities.iterrows():
                                    is_selected = row['Customer'] not in st.session_state[unselected_key]
                                    # Use edited projected value if available
                                    proj_val = st.session_state[edited_proj_key].get(row['Customer'], row['Projected_Value'])
                                    row_data = {
                                        'Select': is_selected,
                                        'Customer': row['Customer'],
                                    }
                                    if is_team_view and 'Rep' in row and row['Rep']:
                                        row_data['Rep'] = row['Rep']
                                    row_data.update({
                                        'Confidence': f"{row['Confidence_Tier']} ({int(row['Confidence_Pct']*100)}%)",
                                        '2025 Orders': row['Order_Count'],
                                        'Avg Order': row['Weighted_Avg_Order'],
                                        'Cadence': f"~{int(row['Cadence_Days'])}d" if pd.notna(row['Cadence_Days']) else "N/A",
                                        'Exp Q1': row['Expected_Orders_Q1'] if 'Expected_Orders_Q1' in row else 1.0,
                                        'Last Order': row['Last_Order_Date'].strftime('%Y-%m-%d') if pd.notna(row['Last_Order_Date']) else '',
                                        'Forecast': proj_val
                                    })
                                    display_data.append(row_data)
                                
                                display_df = pd.DataFrame(display_data)
                                
                                col_config = {
                                    "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                    "Customer": st.column_config.TextColumn("Customer", width="medium"),
                                }
                                disabled_cols = ['Customer', 'Confidence', '2025 Orders', 'Avg Order', 'Cadence', 'Exp Q1', 'Last Order']
                                if is_team_view and 'Rep' in display_df.columns:
                                    col_config["Rep"] = st.column_config.TextColumn("Rep", width="small")
                                    disabled_cols.append('Rep')
                                col_config.update({
                                    "Confidence": st.column_config.TextColumn("Confidence", width="small"),
                                    "2025 Orders": st.column_config.NumberColumn("Orders", width="small"),
                                    "Avg Order": st.column_config.NumberColumn("Avg Order", format="$%d"),
                                    "Cadence": st.column_config.TextColumn("Cadence", width="small"),
                                    "Exp Q1": st.column_config.NumberColumn("Exp Q1", format="%.1f"),
                                    "Last Order": st.column_config.TextColumn("Last Order", width="small"),
                                    "Forecast": st.column_config.NumberColumn("Forecast $", format="$%d", help="Edit this value to adjust forecast")
                                })
                                
                                edited = st.data_editor(
                                    display_df,
                                    column_config=col_config,
                                    disabled=disabled_cols,
                                    hide_index=True,
                                    key=f"q1_reorder_prod_edit_{product_type}_{rep_name}",
                                    use_container_width=True
                                )
                                
                                # Update unselected set and edited projections
                                current_unselected = set()
                                for _, row in edited.iterrows():
                                    if not row['Select']:
                                        current_unselected.add(row['Customer'])
                                    # Store edited forecast values
                                    st.session_state[edited_proj_key][row['Customer']] = row['Forecast']
                                st.session_state[unselected_key] = current_unselected
                                
                                # Calculate selected total using EDITED values
                                selected_total = 0
                                selected_count = 0
                                for _, row in edited.iterrows():
                                    if row['Select']:
                                        selected_total += row['Forecast']
                                        selected_count += 1
                                st.caption(f"Selected: {selected_count} customers, ${selected_total:,.0f} forecast")
                                
                                # === LINE ITEM EDITING SECTION ===
                                selected_customers_list = [row['Customer'] for _, row in edited.iterrows() if row['Select']]
                                
                                if selected_customers_list and not line_items_df.empty:
                                    with st.expander("📋 Edit Line Items (Qty & Rate)"):
                                        st.caption("Adjust quantities and rates to refine your forecast. Item names are read-only.")
                                        
                                        # Session state for edited line items
                                        line_items_key = f"q1_reorder_prod_line_items_{product_type}_{rep_name}"
                                        if line_items_key not in st.session_state:
                                            st.session_state[line_items_key] = {}
                                        
                                        # Collect line items for all selected customers
                                        all_line_items = []
                                        for cust in selected_customers_list:
                                            cust_row = prod_opportunities[prod_opportunities['Customer'] == cust]
                                            if not cust_row.empty:
                                                so_numbers = cust_row.iloc[0].get('SO_Numbers', [])
                                                if so_numbers:
                                                    cust_items = get_customer_line_items(so_numbers, line_items_df)
                                                    if not cust_items.empty:
                                                        cust_items['Customer'] = cust
                                                        cust_items['Confidence_Pct'] = cust_row.iloc[0].get('Confidence_Pct', 0.5)
                                                        all_line_items.append(cust_items)
                                        
                                        if all_line_items:
                                            combined_items = pd.concat(all_line_items, ignore_index=True)
                                            
                                            # Build editable display with persisted edits
                                            line_display = []
                                            for _, item_row in combined_items.iterrows():
                                                item_key = f"{item_row['Customer']}|{item_row['Item']}"
                                                
                                                # Get edited values or defaults
                                                if item_key in st.session_state[line_items_key]:
                                                    qty = st.session_state[line_items_key][item_key]['Qty']
                                                    rate = st.session_state[line_items_key][item_key]['Rate']
                                                else:
                                                    # Default: proportional qty for Q1 based on expected orders
                                                    cust_row = prod_opportunities[prod_opportunities['Customer'] == item_row['Customer']]
                                                    exp_q1 = cust_row.iloc[0].get('Expected_Orders_Q1', 1.0) if not cust_row.empty else 1.0
                                                    order_count = cust_row.iloc[0].get('Order_Count', 1) if not cust_row.empty else 1
                                                    scale_factor = exp_q1 / max(order_count, 1)
                                                    qty = item_row['Total_Qty'] * scale_factor
                                                    rate = item_row['Avg_Rate']
                                                
                                                line_display.append({
                                                    'Customer': item_row['Customer'],
                                                    'Item': item_row['Item'],
                                                    'Hist Qty': item_row['Total_Qty'],
                                                    'Q1 Qty': qty,
                                                    'Rate': rate,
                                                    'Line Total': qty * rate
                                                })
                                            
                                            line_df = pd.DataFrame(line_display)
                                            
                                            edited_lines = st.data_editor(
                                                line_df,
                                                column_config={
                                                    "Customer": st.column_config.TextColumn("Customer", width="medium"),
                                                    "Item": st.column_config.TextColumn("Item", width="medium"),
                                                    "Hist Qty": st.column_config.NumberColumn("2025 Qty", format="%d"),
                                                    "Q1 Qty": st.column_config.NumberColumn("Q1 Qty", format="%.0f", help="Edit: Forecasted quantity"),
                                                    "Rate": st.column_config.NumberColumn("Rate $", format="$%.2f", help="Edit: Price per unit"),
                                                    "Line Total": st.column_config.NumberColumn("Total $", format="$%,.0f")
                                                },
                                                disabled=['Customer', 'Item', 'Hist Qty', 'Line Total'],
                                                hide_index=True,
                                                key=f"q1_prod_line_edit_{product_type}_{rep_name}",
                                                use_container_width=True
                                            )
                                            
                                            # Store edited values and recalculate by customer
                                            customer_totals = {}
                                            for _, line_row in edited_lines.iterrows():
                                                item_key = f"{line_row['Customer']}|{line_row['Item']}"
                                                st.session_state[line_items_key][item_key] = {
                                                    'Qty': line_row['Q1 Qty'],
                                                    'Rate': line_row['Rate']
                                                }
                                                line_total = line_row['Q1 Qty'] * line_row['Rate']
                                                
                                                if line_row['Customer'] not in customer_totals:
                                                    customer_totals[line_row['Customer']] = 0
                                                customer_totals[line_row['Customer']] += line_total
                                            
                                            # Update forecasts with confidence applied
                                            total_forecast = 0
                                            for cust, cust_total in customer_totals.items():
                                                cust_row = prod_opportunities[prod_opportunities['Customer'] == cust]
                                                conf_pct = cust_row.iloc[0].get('Confidence_Pct', 0.5) if not cust_row.empty else 0.5
                                                cust_forecast = cust_total * conf_pct
                                                st.session_state[edited_proj_key][cust] = cust_forecast
                                                total_forecast += cust_forecast
                                            
                                            st.success(f"📊 Line Item Forecast: **${total_forecast:,.0f}** (with confidence % applied)")
                                        else:
                                            st.info("No line item detail available for selected customers")
                                
                                # Store for export with edited values
                                export_df = prod_opportunities[prod_opportunities['Customer'].isin(selected_customers_list)].copy()
                                # Update with edited forecast values
                                for idx, row in export_df.iterrows():
                                    if row['Customer'] in st.session_state[edited_proj_key]:
                                        export_df.loc[idx, 'Projected_Value'] = st.session_state[edited_proj_key][row['Customer']]
                                reorder_buckets[f"reorder_prod_{product_type}"] = export_df
                            else:
                                # Read-only view - shows calculated values, enable Customize to edit
                                display_data = []
                                for _, row in prod_opportunities.iterrows():
                                    row_data = {'Customer': row['Customer']}
                                    if is_team_view and 'Rep' in row and row['Rep']:
                                        row_data['Rep'] = row['Rep']
                                    row_data.update({
                                        'Confidence': f"{row['Confidence_Tier']} ({int(row['Confidence_Pct']*100)}%)",
                                        '2025 Orders': row['Order_Count'],
                                        'Avg Order': row['Weighted_Avg_Order'],
                                        'Cadence': f"~{int(row['Cadence_Days'])}d" if pd.notna(row['Cadence_Days']) else "N/A",
                                        'Exp Q1': row['Expected_Orders_Q1'] if 'Expected_Orders_Q1' in row else 1.0,
                                        'Last Order': row['Last_Order_Date'].strftime('%Y-%m-%d') if pd.notna(row['Last_Order_Date']) else '',
                                        'Forecast': row['Projected_Value']
                                    })
                                    display_data.append(row_data)
                                
                                display_df = pd.DataFrame(display_data)
                                col_config = {"Customer": st.column_config.TextColumn("Customer", width="medium")}
                                if is_team_view and 'Rep' in display_df.columns:
                                    col_config["Rep"] = st.column_config.TextColumn("Rep", width="small")
                                col_config.update({
                                    "Confidence": st.column_config.TextColumn("Confidence", width="small"),
                                    "2025 Orders": st.column_config.NumberColumn("Orders", width="small"),
                                    "Avg Order": st.column_config.NumberColumn("Avg Order", format="$%d"),
                                    "Cadence": st.column_config.TextColumn("Cadence", width="small"),
                                    "Exp Q1": st.column_config.NumberColumn("Exp Q1", format="%.1f"),
                                    "Last Order": st.column_config.TextColumn("Last Order", width="small"),
                                    "Forecast": st.column_config.NumberColumn("Forecast $", format="$%d")
                                })
                                
                                st.dataframe(
                                    display_df,
                                    column_config=col_config,
                                    hide_index=True,
                                    use_container_width=True
                                )
                                st.caption("💡 Enable **Customize** to edit forecast values")
                                
                                # Store for export
                                reorder_buckets[f"reorder_prod_{product_type}"] = prod_opportunities
    
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
            <div class="sticky-forecast-label">+ Reorder</div>
            <div class="sticky-forecast-value" style="color: #f59e0b; text-shadow: 0 0 20px rgba(245, 158, 11, 0.5);">${selected_reorder:,.0f}</div>
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
    
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.metric("1. Scheduled", f"${selected_scheduled:,.0f}", help="NetSuite orders with Q1 dates")
    with m2:
        st.metric("2. Pipeline", f"${selected_pipeline:,.0f}", help="HubSpot deals for Q1")
    with m3:
        st.metric("3. Reorder", f"${selected_reorder:,.0f}", help="Historical opportunity (probability-weighted)")
    with m4:
        st.metric("🏁 Total Forecast", f"${total_forecast:,.0f}", delta="Sum of 1+2+3")
    with m5:
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
        
        **Reorder Forecast:** ${selected_reorder:,.0f}
        - Historical customers with no pending orders
        - Probability-weighted by order frequency
        - 🔴 Likely (75%) | 🟡 Possible (50%) | ⚪ Long Shot (25%)
        
        **Your Q1 Goal:** ${q1_goal:,.0f}
        """)
    
    # === EXPORT SECTION ===
    st.markdown("---")
    st.markdown("### 📤 Export Q1 2026 Forecast")
    
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
