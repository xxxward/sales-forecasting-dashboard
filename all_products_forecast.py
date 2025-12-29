"""
Q1 2026 Sales Forecasting Module
High-Fidelity UI Overhaul
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# ========== DATE CONSTANTS (LOGIC PRESERVED) ==========
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


# ========== ULTRA-PREMIUM CSS ==========
def inject_custom_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap');

    /* BASE THEME */
    .stApp {
        background: #0f1116; /* Deepest charcoal */
        background-image: 
            radial-gradient(at 0% 0%, rgba(56, 189, 248, 0.08) 0px, transparent 50%),
            radial-gradient(at 100% 0%, rgba(16, 185, 129, 0.08) 0px, transparent 50%);
        font-family: 'Inter', sans-serif;
        color: #e2e8f0;
    }

    /* REMOVE PADDING & CLEANUP */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 8rem; /* Space for HUD */
        max-width: 96% !important;
    }
    header {visibility: hidden;}
    
    /* CARDS & CONTAINERS */
    .glass-panel {
        background: rgba(30, 41, 59, 0.3);
        backdrop-filter: blur(16px);
        -webkit-backdrop-filter: blur(16px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        margin-bottom: 20px;
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .glass-panel:hover {
        border-color: rgba(255, 255, 255, 0.2);
    }
    
    /* HEADINGS */
    h1, h2, h3 {
        font-family: 'Inter', sans-serif;
        letter-spacing: -0.02em;
    }
    .gradient-text {
        background: linear-gradient(135deg, #fff 0%, #cbd5e1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
    }
    .section-title {
        font-size: 1.1rem;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: #94a3b8;
        margin-top: 30px;
        margin-bottom: 15px;
        border-left: 3px solid #3b82f6;
        padding-left: 12px;
        display: flex;
        align-items: center;
    }

    /* METRICS */
    .metric-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        background: rgba(15, 23, 42, 0.6);
        border-radius: 10px;
        padding: 15px;
        border: 1px solid rgba(255,255,255,0.05);
    }
    .metric-label {
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748b;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #f1f5f9;
        font-variant-numeric: tabular-nums;
    }
    
    /* INPUTS & WIDGETS */
    div[data-testid="stSelectbox"] > div > div {
        background-color: rgba(30, 41, 59, 0.5) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        color: white !important;
    }
    div[data-testid="stNumberInput"] input {
        background-color: rgba(15, 23, 42, 0.8) !important;
        border: 1px solid rgba(59, 130, 246, 0.3) !important;
        color: white !important;
        font-family: 'Inter', monospace !important;
        font-weight: 600;
    }
    div[data-testid="stCheckbox"] label {
        color: #e2e8f0 !important;
    }
    
    /* BUTTONS */
    div.stButton > button {
        background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid rgba(255,255,255,0.1);
        color: #e2e8f0;
        font-weight: 500;
        border-radius: 6px;
        transition: all 0.2s;
    }
    div.stButton > button:hover {
        border-color: #3b82f6;
        color: #fff;
        box-shadow: 0 0 15px rgba(59, 130, 246, 0.2);
    }
    
    /* DATA TABLES */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(255,255,255,0.05);
        border-radius: 8px;
        overflow: hidden;
    }

    /* HUD FOOTER */
    .hud-footer {
        position: fixed;
        bottom: 20px;
        left: 50%;
        transform: translateX(-50%);
        width: 95%;
        max-width: 1200px;
        background: rgba(15, 23, 42, 0.9);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(59, 130, 246, 0.2);
        box-shadow: 0 0 40px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.1);
        border-radius: 16px;
        z-index: 99999;
        padding: 12px 0;
        display: flex;
        justify-content: space-evenly;
        align-items: center;
    }
    .hud-item {
        text-align: center;
        padding: 0 15px;
        position: relative;
    }
    .hud-item::after {
        content: '';
        position: absolute;
        right: 0;
        top: 20%;
        height: 60%;
        width: 1px;
        background: rgba(255,255,255,0.1);
    }
    .hud-item:last-child::after { display: none; }
    
    .hud-label {
        font-size: 0.65rem;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: #64748b;
        margin-bottom: 2px;
    }
    .hud-val {
        font-size: 1.25rem;
        font-weight: 700;
        font-family: 'Inter', sans-serif;
        font-variant-numeric: tabular-nums;
    }
    
    /* COLORS */
    .c-green { color: #34d399; text-shadow: 0 0 20px rgba(52, 211, 153, 0.2); }
    .c-blue { color: #60a5fa; text-shadow: 0 0 20px rgba(96, 165, 250, 0.2); }
    .c-amber { color: #fbbf24; text-shadow: 0 0 20px rgba(251, 191, 36, 0.2); }
    .c-red { color: #f87171; text-shadow: 0 0 20px rgba(248, 113, 113, 0.2); }
    
    /* ANIMATIONS */
    @keyframes pulse-glow {
        0% { box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.4); }
        70% { box-shadow: 0 0 0 10px rgba(59, 130, 246, 0); }
        100% { box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }
    }
    .pulse-dot {
        height: 8px;
        width: 8px;
        background-color: #3b82f6;
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
        animation: pulse-glow 2s infinite;
    }

    /* BADGES */
    .status-badge {
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: 600;
        border: 1px solid transparent;
    }
    .badge-likely { background: rgba(16, 185, 129, 0.15); color: #34d399; border-color: rgba(16, 185, 129, 0.3); }
    .badge-possible { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border-color: rgba(245, 158, 11, 0.3); }
    .badge-longshot { background: rgba(148, 163, 184, 0.15); color: #cbd5e1; border-color: rgba(148, 163, 184, 0.3); }
    </style>
    """, unsafe_allow_html=True)


# ========== GAUGE CHART (RE-STYLED) ==========
def create_q1_gauge(value, goal, title="Progress"):
    """Modern, minimalist gauge chart"""
    if goal <= 0: goal = 1
    percentage = (value / goal) * 100
    
    # High-contrast neon palette
    if percentage >= 100: bar_color = "#34d399" # Emerald
    elif percentage >= 75: bar_color = "#60a5fa" # Blue
    elif percentage >= 50: bar_color = "#fbbf24" # Amber
    else: bar_color = "#f87171" # Red
    
    max_range = max(goal * 1.1, value * 1.05)
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={
            'prefix': "$", 
            'valueformat': ",.0f",
            'font': {'size': 40, 'color': 'white', 'family': 'Inter'},
            'suffix': f" <span style='font-size:0.6em;color:#94a3b8'>/ ${(goal/1000000):.1f}M</span>"
        },
        gauge={
            'axis': {'range': [0, max_range], 'visible': False},
            'bar': {'color': bar_color, 'thickness': 0.85},
            'bgcolor': "rgba(255,255,255,0.05)",
            'borderwidth': 0,
            'steps': [{'range': [0, goal], 'color': "rgba(255,255,255,0.02)"}],
            'threshold': {
                'line': {'color': "white", 'width': 3},
                'thickness': 0.9,
                'value': goal
            }
        }
    ))
    
    fig.update_layout(
        height=220,
        margin=dict(l=20, r=20, t=30, b=10),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white', 'family': 'Inter'}
    )
    return fig

# ========== FORMAT FUNCTIONS (LOGIC PRESERVED) ==========
def get_col_by_index(df, index):
    if df is not None and len(df.columns) > index:
        return df.iloc[:, index]
    return pd.Series()

def format_ns_view(df, date_col_name):
    if df.empty: return df
    d = df.copy()
    if d.columns.duplicated().any(): d = d.loc[:, ~d.columns.duplicated()]
    if 'Internal ID' in d.columns:
        d['Link'] = d['Internal ID'].apply(lambda x: f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}" if pd.notna(x) else "")
    if 'Display_SO_Num' in d.columns: d['SO #'] = d['Display_SO_Num']
    elif 'Document Number' in d.columns: d['SO #'] = d['Document Number']
    if 'Display_Type' in d.columns: d['Type'] = d['Display_Type']
    if date_col_name == 'Promise':
        d['Ship Date'] = ''
        if 'Display_Promise_Date' in d.columns:
            pd_dates = pd.to_datetime(d['Display_Promise_Date'], errors='coerce')
            d.loc[pd_dates.notna(), 'Ship Date'] = pd_dates.dt.strftime('%Y-%m-%d')
        if 'Display_Projected_Date' in d.columns:
            proj_dates = pd.to_datetime(d['Display_Projected_Date'], errors='coerce')
            mask = (d['Ship Date'] == '') & proj_dates.notna()
            if mask.any(): d.loc[mask, 'Ship Date'] = proj_dates.loc[mask].dt.strftime('%Y-%m-%d')
    elif date_col_name == 'PA_Date':
        if 'Display_PA_Date' in d.columns:
            pa_dates = pd.to_datetime(d['Display_PA_Date'], errors='coerce')
            d['Ship Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('')
        elif 'PA_Date_Parsed' in d.columns:
            pa_dates = pd.to_datetime(d['PA_Date_Parsed'], errors='coerce')
            d['Ship Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('')
        else: d['Ship Date'] = ''
    else: d['Ship Date'] = ''
    return d.sort_values('Amount', ascending=False) if 'Amount' in d.columns else d

def format_hs_view(df):
    if df.empty: return df
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
    if 'Account Name' not in d.columns and 'Deal Name' in d.columns:
        d['Account Name'] = d['Deal Name']
    return d.sort_values('Amount_Numeric', ascending=False) if 'Amount_Numeric' in d.columns else d

# ========== HISTORICAL ANALYSIS (LOGIC PRESERVED) ==========
# (Functions: load_historical_orders, load_invoices, load_line_items, load_item_master,
#  merge_orders_with_invoices, calculate_customer_metrics, calculate_customer_product_metrics,
#  identify_reorder_opportunities, get_customer_line_items, get_product_type_summary
#  are functionally identical to the source but collapsed here for brevity in display)

def load_historical_orders(main_dash, rep_name):
    historical_df = main_dash.load_google_sheets_data("NS Sales Orders", "A:AF", version=main_dash.CACHE_VERSION)
    if historical_df.empty: return pd.DataFrame()
    col_names = historical_df.columns.tolist()
    rename_dict = {}
    if len(col_names) > 0: rename_dict[col_names[0]] = 'Internal ID'
    if len(col_names) > 1: rename_dict[col_names[1]] = 'SO_Number'
    if len(col_names) > 2: rename_dict[col_names[2]] = 'Status'
    if len(col_names) > 7: rename_dict[col_names[7]] = 'Amount'
    if len(col_names) > 8: rename_dict[col_names[8]] = 'Order Start Date'
    if len(col_names) > 17: rename_dict[col_names[17]] = 'Order Type'
    if len(col_names) > 30: rename_dict[col_names[30]] = 'Customer'
    if len(col_names) > 31: rename_dict[col_names[31]] = 'Rep Master'
    historical_df = historical_df.rename(columns=rename_dict)
    if historical_df.columns.duplicated().any(): historical_df = historical_df.loc[:, ~historical_df.columns.duplicated()]
    if 'SO_Number' in historical_df.columns: historical_df['SO_Number'] = historical_df['SO_Number'].astype(str).str.strip().str.upper()
    if 'Status' in historical_df.columns:
        historical_df['Status'] = historical_df['Status'].astype(str).str.strip()
        historical_df = historical_df[historical_df['Status'].isin(['Billed', 'Closed'])]
    else: return pd.DataFrame()
    if 'Rep Master' in historical_df.columns:
        historical_df['Rep Master'] = historical_df['Rep Master'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        historical_df = historical_df[~historical_df['Rep Master'].isin(invalid_values)]
        historical_df = historical_df[historical_df['Rep Master'] == rep_name]
    else: return pd.DataFrame()
    if 'Customer' in historical_df.columns:
        historical_df['Customer'] = historical_df['Customer'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        historical_df = historical_df[~historical_df['Customer'].isin(invalid_values)]
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '': return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try: return float(cleaned)
        except: return 0
    if 'Amount' in historical_df.columns:
        historical_df['Amount'] = historical_df['Amount'].apply(clean_numeric)
        historical_df = historical_df[historical_df['Amount'] > 0]
    if 'Order Start Date' in historical_df.columns:
        historical_df['Order Start Date'] = pd.to_datetime(historical_df['Order Start Date'], errors='coerce')
        if historical_df['Order Start Date'].notna().any():
            mask = (historical_df['Order Start Date'].dt.year < 2000) & (historical_df['Order Start Date'].notna())
            if mask.any(): historical_df.loc[mask, 'Order Start Date'] = historical_df.loc[mask, 'Order Start Date'] + pd.DateOffset(years=100)
        year_2025_start = pd.Timestamp('2025-01-01')
        year_2025_end = pd.Timestamp('2025-12-31')
        historical_df = historical_df[(historical_df['Order Start Date'] >= year_2025_start) & (historical_df['Order Start Date'] <= year_2025_end)]
    if 'Order Type' in historical_df.columns:
        historical_df['Order Type'] = historical_df['Order Type'].astype(str).str.strip()
        historical_df.loc[historical_df['Order Type'].isin(['', 'nan', 'None']), 'Order Type'] = 'Standard'
    else: historical_df['Order Type'] = 'Standard'
    return historical_df

def load_invoices(main_dash, rep_name):
    invoice_df = main_dash.load_google_sheets_data("NS Invoices", "A:U", version=main_dash.CACHE_VERSION)
    if invoice_df.empty: return pd.DataFrame()
    col_names = invoice_df.columns.tolist()
    rename_dict = {}
    if len(col_names) > 2: rename_dict[col_names[2]] = 'Invoice_Date'
    if len(col_names) > 4: rename_dict[col_names[4]] = 'SO_Number'
    if len(col_names) > 10: rename_dict[col_names[10]] = 'Invoice_Amount'
    if len(col_names) > 19: rename_dict[col_names[19]] = 'Customer'
    if len(col_names) > 20: rename_dict[col_names[20]] = 'Rep Master'
    invoice_df = invoice_df.rename(columns=rename_dict)
    if invoice_df.columns.duplicated().any(): invoice_df = invoice_df.loc[:, ~invoice_df.columns.duplicated()]
    if 'Rep Master' in invoice_df.columns:
        invoice_df['Rep Master'] = invoice_df['Rep Master'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        invoice_df = invoice_df[~invoice_df['Rep Master'].isin(invalid_values)]
        invoice_df = invoice_df[invoice_df['Rep Master'] == rep_name]
    else: return pd.DataFrame()
    if 'Customer' in invoice_df.columns:
        invoice_df['Customer'] = invoice_df['Customer'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        invoice_df = invoice_df[~invoice_df['Customer'].isin(invalid_values)]
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '': return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try: return float(cleaned)
        except: return 0
    if 'Invoice_Amount' in invoice_df.columns:
        invoice_df['Invoice_Amount'] = invoice_df['Invoice_Amount'].apply(clean_numeric)
        invoice_df = invoice_df[invoice_df['Invoice_Amount'] > 0]
    if 'Invoice_Date' in invoice_df.columns:
        invoice_df['Invoice_Date'] = pd.to_datetime(invoice_df['Invoice_Date'], errors='coerce')
        if invoice_df['Invoice_Date'].notna().any():
            mask = (invoice_df['Invoice_Date'].dt.year < 2000) & (invoice_df['Invoice_Date'].notna())
            if mask.any(): invoice_df.loc[mask, 'Invoice_Date'] = invoice_df.loc[mask, 'Invoice_Date'] + pd.DateOffset(years=100)
        year_2025_start = pd.Timestamp('2025-01-01')
        year_2025_end = pd.Timestamp('2025-12-31')
        invoice_df = invoice_df[(invoice_df['Invoice_Date'] >= year_2025_start) & (invoice_df['Invoice_Date'] <= year_2025_end)]
    if 'SO_Number' in invoice_df.columns: invoice_df['SO_Number'] = invoice_df['SO_Number'].astype(str).str.strip().str.upper()
    return invoice_df

def load_line_items(main_dash):
    line_items_df = main_dash.load_google_sheets_data("Sales Order Line Item", "A:F", version=main_dash.CACHE_VERSION)
    if line_items_df.empty: return pd.DataFrame()
    col_names = line_items_df.columns.tolist()
    rename_dict = {}
    if len(col_names) > 1: rename_dict[col_names[1]] = 'SO_Number'
    if len(col_names) > 2: rename_dict[col_names[2]] = 'Item'
    if len(col_names) > 4: rename_dict[col_names[4]] = 'Item_Rate'
    if len(col_names) > 5: rename_dict[col_names[5]] = 'Quantity'
    line_items_df = line_items_df.rename(columns=rename_dict)
    if line_items_df.columns.duplicated().any(): line_items_df = line_items_df.loc[:, ~line_items_df.columns.duplicated()]
    if 'SO_Number' in line_items_df.columns:
        line_items_df['SO_Number'] = line_items_df['SO_Number'].astype(str).str.strip().str.upper()
        line_items_df = line_items_df[line_items_df['SO_Number'] != '']
        line_items_df = line_items_df[line_items_df['SO_Number'].str.lower() != 'nan']
    if 'Item' in line_items_df.columns:
        line_items_df['Item'] = line_items_df['Item'].astype(str).str.strip()
        line_items_df = line_items_df[line_items_df['Item'] != '']
        line_items_df = line_items_df[line_items_df['Item'].str.lower() != 'nan']
        exclude_patterns = ['avatax', 'tax', 'fee', 'convenience', 'surcharge', 'handling', 'shipping', 'freight', 'fedex', 'ups ', 'usps', 'ltl', 'truckload', 'customer pickup', 'client arranged', 'generic ship', 'send to inventory', 'default shipping', 'best way', 'ground', 'next day', '2nd day', '3rd day', 'overnight', 'standard', 'saver', 'express', 'priority', 'estes', 't-force', 'ward trucking', 'old dominion', 'roadrunner', 'xpo logistics', 'abf', 'a. duie pyle', 'frontline freight', 'saia', 'dependable highway', 'cross country', 'oak harbor', 'discount', 'credit', 'adjustment', 'replacement order', 'partner discount', 'creative', 'pre-press', 'retrofit', 'press proof', 'design', 'die cut sample', 'label appl', 'application', 'changeover', 'expedite', 'rush', 'sample', 'testimonial', 'cm-for sos', 'wip', 'work in progress', 'end of group', 'other', '-not taxable-', 'fep-liner insert', 'cc payment', 'waive', 'modular plus', 'canadian business', 'canadian goods']
        exclude_exact = ['brad10', 'blake10', '420ten', 'oil10', 'welcome10', 'take10', 'jack', 'jake', 'james20off', 'lpp15', 'brad', 'davis', 'mjbiz2023', 'blackfriday10', 'danksggivingtubes', 'legends20', 'mjbizlastcall', '$100off', 'sb-45d-kit', 'sb-25d-kit', 'sb-145d-kit', 'sb-15d-kit', 'flexpack', 'bb-dml-000-00', '145d-blk-blk', 'bisonbotanics45d', 'samples2023', 'samples2023-inactive', 'jake-inactive', 'replacement order-inactive', 'every-other-label-free', 'free-application', 'single item discount', 'single line item discount', 'general discount', 'rist/howards', 'diamond creative tier', 'silver creative tier', 'platinum creative tier']
        state_pattern = re.compile(r'^[A-Z]{2}_')
        item_lower = line_items_df['Item'].str.lower()
        item_upper = line_items_df['Item'].str.upper()
        pattern_mask = item_lower.apply(lambda x: any(pattern in x for pattern in exclude_patterns))
        exact_mask = item_lower.isin([e.lower() for e in exclude_exact])
        location_mask = item_upper.apply(lambda x: bool(state_pattern.match(x)))
        exclude_mask = pattern_mask | exact_mask | location_mask
        line_items_df = line_items_df[~exclude_mask]
    def clean_numeric(value):
        if pd.isna(value) or str(value).strip() == '': return 0
        cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
        try: return float(cleaned)
        except: return 0
    if 'Item_Rate' in line_items_df.columns: line_items_df['Item_Rate'] = line_items_df['Item_Rate'].apply(clean_numeric)
    if 'Quantity' in line_items_df.columns: line_items_df['Quantity'] = line_items_df['Quantity'].apply(clean_numeric)
    line_items_df['Line_Total'] = line_items_df['Quantity'] * line_items_df['Item_Rate']
    return line_items_df

def load_item_master(main_dash):
    item_master_df = main_dash.load_google_sheets_data("Item Master", "A:C", version=main_dash.CACHE_VERSION)
    if item_master_df.empty: return {}
    col_names = item_master_df.columns.tolist()
    if len(col_names) < 3: return {}
    rename_dict = {col_names[0]: 'Item', col_names[2]: 'Description'}
    item_master_df = item_master_df.rename(columns=rename_dict)
    if 'Item' in item_master_df.columns:
        item_master_df['Item'] = item_master_df['Item'].astype(str).str.strip()
        item_master_df = item_master_df[item_master_df['Item'] != '']
        item_master_df = item_master_df[item_master_df['Item'].str.lower() != 'nan']
    if 'Description' in item_master_df.columns:
        item_master_df['Description'] = item_master_df['Description'].astype(str).str.strip().replace('nan', '')
    sku_to_desc = dict(zip(item_master_df['Item'], item_master_df['Description']))
    return sku_to_desc

def merge_orders_with_invoices(orders_df, invoices_df):
    if orders_df.empty: return orders_df
    if invoices_df.empty:
        orders_df['Invoice_Amount'] = orders_df['Amount']
        return orders_df
    orders_df['SO_Number_Clean'] = orders_df['SO_Number'].astype(str).str.strip().str.upper()
    invoices_df['SO_Number_Clean'] = invoices_df['SO_Number'].astype(str).str.strip().str.upper()
    invoice_totals = invoices_df.groupby('SO_Number_Clean')['Invoice_Amount'].sum().reset_index()
    merged = orders_df.merge(invoice_totals, on='SO_Number_Clean', how='left')
    merged['Invoice_Amount'] = merged['Invoice_Amount'].fillna(merged['Amount'])
    return merged

def calculate_customer_metrics(historical_df):
    if historical_df.empty: return pd.DataFrame()
    today = pd.Timestamp.now()
    amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
    customer_metrics = []
    for customer in historical_df['Customer'].unique():
        cust_orders = historical_df[historical_df['Customer'] == customer].copy()
        cust_orders = cust_orders.sort_values('Order Start Date')
        order_count = len(cust_orders)
        total_revenue = cust_orders[amount_col].sum()
        order_dates = cust_orders['Order Start Date'].dropna().tolist()
        weighted_sum = 0
        weight_total = 0
        for _, row in cust_orders.iterrows():
            order_date = row['Order Start Date']
            amount = row[amount_col]
            if pd.notna(order_date) and order_date.month >= 7: weight = 1.25
            else: weight = 1.0
            weighted_sum += amount * weight
            weight_total += weight
        weighted_avg = weighted_sum / weight_total if weight_total > 0 else 0
        cadence_days = None
        if len(order_dates) >= 2:
            gaps = []
            for i in range(len(order_dates) - 1):
                gap = (order_dates[i + 1] - order_dates[i]).days
                if gap > 0: gaps.append(gap)
            if gaps: cadence_days = sum(gaps) / len(gaps)
        last_order_date = cust_orders['Order Start Date'].max()
        days_since_last = (today - last_order_date).days if pd.notna(last_order_date) else 999
        product_types = cust_orders['Order Type'].value_counts().to_dict()
        product_types_str = ', '.join([f"{k} ({v})" for k, v in product_types.items()])
        if order_count >= 3:
            confidence_tier = 'Likely'
            confidence_pct = 0.75
        elif order_count >= 2:
            confidence_tier = 'Possible'
            confidence_pct = 0.50
        else:
            confidence_tier = 'Long Shot'
            confidence_pct = 0.25
        q1_days = 90
        if cadence_days and cadence_days > 0:
            expected_orders_q1 = q1_days / cadence_days
            expected_orders_q1 = min(expected_orders_q1, 6.0)
            expected_orders_q1 = max(expected_orders_q1, 1.0)
        else: expected_orders_q1 = 1.0
        projected_value = weighted_avg * expected_orders_q1 * confidence_pct
        rep_for_customer = cust_orders['Rep'].iloc[0] if 'Rep' in cust_orders.columns else ''
        so_numbers = []
        if 'SO_Number' in cust_orders.columns: so_numbers = cust_orders['SO_Number'].dropna().unique().tolist()
        customer_metrics.append({'Customer': customer, 'Rep': rep_for_customer, 'Order_Count': order_count, 'Total_Revenue': total_revenue, 'Weighted_Avg_Order': weighted_avg, 'Cadence_Days': cadence_days, 'Expected_Orders_Q1': expected_orders_q1, 'Last_Order_Date': last_order_date, 'Days_Since_Last': days_since_last, 'Product_Types': product_types_str, 'Product_Types_Dict': product_types, 'Confidence_Tier': confidence_tier, 'Confidence_Pct': confidence_pct, 'Projected_Value': projected_value, 'SO_Numbers': so_numbers})
    return pd.DataFrame(customer_metrics)

def calculate_customer_product_metrics(historical_df, line_items_df, sku_to_desc=None):
    if sku_to_desc is None: sku_to_desc = {}
    if historical_df.empty: return pd.DataFrame()
    today = pd.Timestamp.now()
    amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
    metrics = []
    for (customer, product_type), group in historical_df.groupby(['Customer', 'Order Type']):
        group = group.sort_values('Order Start Date')
        order_count = len(group)
        total_revenue = group[amount_col].sum()
        avg_order_value = total_revenue / order_count if order_count > 0 else 0
        order_dates = group['Order Start Date'].dropna().tolist()
        cadence_days = None
        if len(order_dates) >= 2:
            gaps = []
            for i in range(len(order_dates) - 1):
                gap = (order_dates[i + 1] - order_dates[i]).days
                if gap > 0: gaps.append(gap)
            if gaps: cadence_days = sum(gaps) / len(gaps)
        last_order_date = group['Order Start Date'].max()
        days_since_last = (today - last_order_date).days if pd.notna(last_order_date) else 999
        q1_days = 90
        if cadence_days and cadence_days > 0:
            expected_orders_q1 = q1_days / cadence_days
            expected_orders_q1 = min(expected_orders_q1, 6.0)
            expected_orders_q1 = max(expected_orders_q1, 1.0)
        else: expected_orders_q1 = 1.0
        if order_count >= 3:
            confidence_tier = 'Likely'
            confidence_pct = 0.75
        elif order_count >= 2:
            confidence_tier = 'Possible'
            confidence_pct = 0.50
        else:
            confidence_tier = 'Long Shot'
            confidence_pct = 0.25
        so_numbers = group['SO_Number'].dropna().unique().tolist() if 'SO_Number' in group.columns else []
        total_qty = 0
        total_line_value = 0
        avg_rate = 0
        sku_count = 0
        top_skus = ""
        if so_numbers and not line_items_df.empty:
            product_line_items = line_items_df[line_items_df['SO_Number'].isin(so_numbers)]
            if not product_line_items.empty:
                total_qty = int(product_line_items['Quantity'].sum())
                total_line_value = product_line_items['Line_Total'].sum()
                avg_rate = total_line_value / total_qty if total_qty > 0 else 0
                sku_count = product_line_items['Item'].nunique()
                sku_totals = product_line_items.groupby('Item')['Line_Total'].sum().sort_values(ascending=False)
                top_sku_list = sku_totals.head(3).index.tolist()
                top_sku_with_desc = []
                for sku in top_sku_list:
                    desc = sku_to_desc.get(sku, '')
                    if desc and desc != sku: top_sku_with_desc.append(desc)
                    else: top_sku_with_desc.append(sku)
                top_skus = ", ".join(top_sku_with_desc) if top_sku_with_desc else ""
        if total_qty > 0:
            avg_qty_per_order = total_qty / order_count
            q1_qty = int(round(avg_qty_per_order * expected_orders_q1))
            q1_value = q1_qty * avg_rate
        else:
            q1_qty = 0
            q1_value = avg_order_value * expected_orders_q1
        q1_forecast = q1_value * confidence_pct
        rep = group['Rep'].iloc[0] if 'Rep' in group.columns else ''
        metrics.append({'Customer': customer, 'Product_Type': product_type, 'Rep': rep, 'Order_Count': order_count, 'Total_Revenue': total_revenue, 'Avg_Order_Value': avg_order_value, 'Cadence_Days': cadence_days, 'Last_Order_Date': last_order_date, 'Days_Since_Last': days_since_last, 'Expected_Orders_Q1': expected_orders_q1, 'Confidence_Tier': confidence_tier, 'Confidence_Pct': confidence_pct, 'SO_Numbers': so_numbers, 'Total_Qty_2025': total_qty, 'Avg_Rate': avg_rate, 'SKU_Count': sku_count, 'Top_SKUs': top_skus, 'Q1_Qty': q1_qty, 'Q1_Value': q1_value, 'Q1_Forecast': q1_forecast})
    return pd.DataFrame(metrics)


# ========== MAIN FUNCTION ==========
def main():
    """Main function for Q1 2026 Forecasting module"""
    
    inject_custom_css()
    
    # === HERO SECTION ===
    days_until_q1 = calculate_business_days_until_q1()
    
    st.markdown(f"""
    <div style="text-align: center; padding: 40px 0;">
        <div style="display: inline-block; padding: 8px 16px; background: rgba(59, 130, 246, 0.1); border-radius: 20px; border: 1px solid rgba(59, 130, 246, 0.3); margin-bottom: 20px;">
            <span class="pulse-dot"></span> <span style="font-size: 0.8rem; font-weight: 600; color: #60a5fa; letter-spacing: 1px;">LIVE FORECASTING</span>
        </div>
        <h1 style="font-size: 4rem; font-weight: 800; background: linear-gradient(to bottom right, #ffffff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin: 0; line-height: 1;">
            Q1 2026
        </h1>
        <p style="color: #64748b; font-size: 1.2rem; font-weight: 400; margin-top: 15px; letter-spacing: 0.05em;">STRATEGIC REVENUE PROJECTION</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Hero Metrics Grid
    col_h1, col_h2, col_h3 = st.columns(3)
    with col_h1:
        st.markdown(f"""
        <div class="metric-container">
            <div class="metric-label">Fiscal Period</div>
            <div class="metric-value">Jan 1 - Mar 31</div>
        </div>
        """, unsafe_allow_html=True)
    with col_h2:
        st.markdown(f"""
        <div class="metric-container">
            <div class="metric-label">Launch Countdown</div>
            <div class="metric-value" style="color: #60a5fa;">{days_until_q1} <span style="font-size:0.5em;color:#64748b">DAYS</span></div>
        </div>
        """, unsafe_allow_html=True)
    with col_h3:
        st.markdown(f"""
        <div class="metric-container">
            <div class="metric-label">Last Sync</div>
            <div class="metric-value">{get_mst_time().strftime('%I:%M %p')}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<div style='margin-bottom: 40px'></div>", unsafe_allow_html=True)
    
    # === IMPORT FROM MAIN DASHBOARD ===
    try:
        import sales_dashboard as main_dash
        deals_df_q4, dashboard_df, invoices_df, sales_orders_df, q4_push_df = main_dash.load_all_data()
        categorize_sales_orders = main_dash.categorize_sales_orders
        deals_df = main_dash.load_google_sheets_data("Copy of All Reps All Pipelines", "A:Z", version=main_dash.CACHE_VERSION)
        
        # [Data Processing Logic preserved exactly as original...]
        if not deals_df.empty and len(deals_df.columns) >= 6:
            col_names = deals_df.columns.tolist()
            rename_dict = {}
            for col in col_names:
                if col == 'Record ID': rename_dict[col] = 'Record ID'
                elif col == 'Deal Name': rename_dict[col] = 'Deal Name'
                elif col == 'Deal Stage': rename_dict[col] = 'Deal Stage'
                elif col == 'Close Date': rename_dict[col] = 'Close Date'
                elif 'Deal Owner First Name' in col and 'Deal Owner Last Name' in col: rename_dict[col] = 'Deal Owner'
                elif col == 'Deal Owner First Name': rename_dict[col] = 'Deal Owner First Name'
                elif col == 'Deal Owner Last Name': rename_dict[col] = 'Deal Owner Last Name'
                elif col == 'Amount': rename_dict[col] = 'Amount'
                elif col == 'Close Status': rename_dict[col] = 'Status'
                elif col == 'Pipeline': rename_dict[col] = 'Pipeline'
                elif col == 'Deal Type': rename_dict[col] = 'Product Type'
                elif col == 'Pending Approval Date': rename_dict[col] = 'Pending Approval Date'
                elif col == 'Q1 2026 Spillover': rename_dict[col] = 'Q1 2026 Spillover'
                elif col == 'Account Name' or col == 'Associated Company': rename_dict[col] = 'Account Name'
                elif col == 'Company': rename_dict[col] = 'Account Name'
            deals_df = deals_df.rename(columns=rename_dict)
            if 'Deal Owner' not in deals_df.columns:
                if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                    deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                    deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
            else: deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
            def clean_numeric(value):
                if pd.isna(value) or str(value).strip() == '': return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try: return float(cleaned)
                except: return 0
            if 'Amount' in deals_df.columns: deals_df['Amount'] = deals_df['Amount'].apply(clean_numeric)
            if 'Close Date' in deals_df.columns: deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
            if 'Pending Approval Date' in deals_df.columns: deals_df['Pending Approval Date'] = pd.to_datetime(deals_df['Pending Approval Date'], errors='coerce')
            excluded_stages = ['', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 'closed lost', 'closed won', 'sales order created in NS', 'NCR', 'Shipped']
            if 'Deal Stage' in deals_df.columns:
                deals_df['Deal Stage'] = deals_df['Deal Stage'].fillna('')
                deals_df['Deal Stage'] = deals_df['Deal Stage'].astype(str).str.strip()
                deals_df = deals_df[~deals_df['Deal Stage'].str.lower().isin([s.lower() if s else '' for s in excluded_stages])]
        
    except ImportError as e:
        st.error(f"❌ Unable to import main dashboard: {e}")
        return
    except Exception as e:
        st.error(f"❌ Error loading data: {e}")
        return
    
    reps = dashboard_df['Rep Name'].tolist() if not dashboard_df.empty else []
    if not reps:
        st.warning("No reps found in Dashboard Info")
        return
    TEAM_REPS = ['Alex Gonzalez', 'Jake Lynch', 'Dave Borkowski', 'Lance Mitton', 'Shopify E-commerce', 'Brad Sherman']
    rep_options = ["👥 All Reps (Team View)"] + reps
    
    # ═══════════════════════════════════════════════════════════════════════════
    # CONTROL CENTER: IDENTITY & GOALS
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
    st.markdown('<div class="gradient-text" style="font-size: 1.5rem; margin-bottom: 20px;">👤 Identification</div>', unsafe_allow_html=True)
    
    col_sel, col_goal = st.columns([1, 1])
    
    with col_sel:
        selected_option = st.selectbox("Select Profile", options=rep_options, key="q1_rep_selector")
        
    is_team_view = selected_option == "👥 All Reps (Team View)"
    if is_team_view:
        rep_name = "All Reps"
        active_team_reps = [r for r in TEAM_REPS if r in reps]
        goal_default = 5000000
    else:
        rep_name = selected_option
        active_team_reps = [rep_name]
        goal_default = 1000000

    goal_key = f"q1_goal_{rep_name}"
    if goal_key not in st.session_state:
        st.session_state[goal_key] = goal_default
    
    with col_goal:
        q1_goal = st.number_input(
            "Q1 2026 Target ($)",
            min_value=0,
            max_value=50000000,
            value=st.session_state[goal_key],
            step=50000,
            format="%d",
            key=f"q1_goal_input_{rep_name}"
        )
        st.session_state[goal_key] = q1_goal

    st.markdown('</div>', unsafe_allow_html=True)

    # === DATA AGGREGATION (LOGIC PRESERVED) ===
    # [Logic for categorizing sales orders...]
    all_pf_spillover, all_pa_spillover, all_pf_nodate, all_pa_date, all_pa_nodate, all_pa_old = [], [], [], [], [], []
    total_pf_amount = total_pa_amount = total_pf_nodate_amount = total_pa_date_amount = total_pa_nodate_amount = total_pa_old_amount = 0
    
    for r in active_team_reps:
        so_cats = categorize_sales_orders(sales_orders_df, r)
        if not so_cats['pf_spillover'].empty:
            all_pf_spillover.append(so_cats['pf_spillover']); total_pf_amount += so_cats['pf_spillover_amount']
        if not so_cats['pa_spillover'].empty:
            all_pa_spillover.append(so_cats['pa_spillover']); total_pa_amount += so_cats['pa_spillover_amount']
        if not so_cats['pf_nodate_ext'].empty:
            all_pf_nodate.append(so_cats['pf_nodate_ext']); total_pf_nodate_amount += so_cats['pf_nodate_ext_amount']
        if not so_cats['pf_nodate_int'].empty:
            all_pf_nodate.append(so_cats['pf_nodate_int']); total_pf_nodate_amount += so_cats['pf_nodate_int_amount']
        if not so_cats['pa_date'].empty:
            all_pa_date.append(so_cats['pa_date']); total_pa_date_amount += so_cats['pa_date_amount']
        if not so_cats['pa_nodate'].empty:
            all_pa_nodate.append(so_cats['pa_nodate']); total_pa_nodate_amount += so_cats['pa_nodate_amount']
        if not so_cats['pa_old'].empty:
            all_pa_old.append(so_cats['pa_old']); total_pa_old_amount += so_cats['pa_old_amount']
    
    combined_pf = pd.concat(all_pf_spillover, ignore_index=True) if all_pf_spillover else pd.DataFrame()
    combined_pa = pd.concat(all_pa_spillover, ignore_index=True) if all_pa_spillover else pd.DataFrame()
    combined_pf_nodate = pd.concat(all_pf_nodate, ignore_index=True) if all_pf_nodate else pd.DataFrame()
    combined_pa_date = pd.concat(all_pa_date, ignore_index=True) if all_pa_date else pd.DataFrame()
    combined_pa_nodate = pd.concat(all_pa_nodate, ignore_index=True) if all_pa_nodate else pd.DataFrame()
    combined_pa_old = pd.concat(all_pa_old, ignore_index=True) if all_pa_old else pd.DataFrame()
    
    ns_categories = {
        'PF_Spillover': {'label': '📦 PF (Q1 2026 Date)', 'df': combined_pf, 'amount': total_pf_amount},
        'PA_Spillover': {'label': '⏳ PA (Q1 2026 PA Date)', 'df': combined_pa, 'amount': total_pa_amount},
        'PF_NoDate': {'label': '📦 PF (No Date)', 'df': combined_pf_nodate, 'amount': total_pf_nodate_amount},
        'PA_Date': {'label': '⏳ PA (With Date)', 'df': combined_pa_date, 'amount': total_pa_date_amount},
        'PA_NoDate': {'label': '⏳ PA (No Date)', 'df': combined_pa_nodate, 'amount': total_pa_nodate_amount},
        'PA_Old': {'label': '⚠️ PA (>2 Weeks)', 'df': combined_pa_old, 'amount': total_pa_old_amount},
    }
    
    ns_dfs = {k: format_ns_view(v['df'], 'Promise' if 'PF' in k else ('PA_Date' if k != 'PA_Old' else 'PA_Date')) for k, v in ns_categories.items()}
    
    # [Logic for HubSpot pipeline...]
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
        rep_deals = deals_df[deals_df['Deal Owner'].isin(active_team_reps)].copy()
        if 'Close Date' in rep_deals.columns:
            q1_close_mask = (rep_deals['Close Date'] >= Q1_2026_START) & (rep_deals['Close Date'] <= Q1_2026_END)
            q1_deals = rep_deals[q1_close_mask]
            q4_close_mask = (rep_deals['Close Date'] >= Q4_2025_START) & (rep_deals['Close Date'] <= Q4_2025_END)
            q4_spillover = rep_deals[q4_close_mask & (rep_deals['Q1 2026 Spillover'] == 'Q1 2026')] if 'Q1 2026 Spillover' in rep_deals.columns else pd.DataFrame()
            
            if 'Status' in q1_deals.columns:
                hs_dfs['Q1_Expect'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Expect'])
                hs_dfs['Q1_Commit'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Commit'])
                hs_dfs['Q1_BestCase'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Best Case'])
                hs_dfs['Q1_Opp'] = format_hs_view(q1_deals[q1_deals['Status'] == 'Opportunity'])
            
            if not q4_spillover.empty and 'Status' in q4_spillover.columns:
                hs_dfs['Q4_Spillover_Expect'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Expect'])
                hs_dfs['Q4_Spillover_Commit'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Commit'])
                hs_dfs['Q4_Spillover_BestCase'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Best Case'])
                hs_dfs['Q4_Spillover_Opp'] = format_hs_view(q4_spillover[q4_spillover['Status'] == 'Opportunity'])
    
    for key in hs_categories.keys():
        if key not in hs_dfs: hs_dfs[key] = pd.DataFrame()
    
    export_buckets = {}

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION: PIPELINE REVIEW
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown('<div class="section-title">01 // PIPELINE INTELLIGENCE</div>', unsafe_allow_html=True)
    
    # Global Controls
    col_ctrl_1, col_ctrl_2, col_ctrl_3 = st.columns([1, 1, 2])
    with col_ctrl_1:
        if st.button("☑️ Select All", key=f"q1_select_all_{rep_name}", use_container_width=True):
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
    with col_ctrl_2:
        if st.button("☐ Deselect All", key=f"q1_unselect_all_{rep_name}", use_container_width=True):
            for key in ns_categories.keys(): st.session_state[f"q1_chk_{key}_{rep_name}"] = False
            for key in hs_categories.keys(): st.session_state[f"q1_chk_{key}_{rep_name}"] = False
            st.rerun()
            
    # Pipeline Columns
    col_ns, col_hs = st.columns(2)
    
    # NetSuite Logic (Visuals Wrapped)
    with col_ns:
        st.markdown('<div class="glass-panel" style="border-top: 3px solid #34d399;">', unsafe_allow_html=True)
        st.markdown('<h3 style="color:#34d399; margin-bottom: 5px;">📦 NetSuite Locked</h3>', unsafe_allow_html=True)
        st.caption("Confirmed orders scheduled for Q1 delivery")
        
        for key, data in ns_categories.items():
            df = ns_dfs.get(key, pd.DataFrame())
            val = data['amount']
            checkbox_key = f"q1_chk_{key}_{rep_name}"
            
            if val > 0:
                st.markdown(f"<div style='margin-top:10px; padding: 8px; background: rgba(255,255,255,0.03); border-radius: 6px;'>", unsafe_allow_html=True)
                is_checked = st.checkbox(f"**{data['label']}** — ${val:,.0f}", key=checkbox_key)
                if is_checked:
                    with st.expander("Details"):
                        if not df.empty:
                            enable_edit = st.toggle("Custom Select", key=f"q1_tgl_{key}_{rep_name}")
                            display_cols = [c for c in ['Link', 'SO #', 'Type', 'Customer', 'Ship Date', 'Amount'] if c in df.columns]
                            
                            if enable_edit and display_cols:
                                df_edit = df.copy()
                                unselected_key = f"q1_unselected_{key}_{rep_name}"
                                if unselected_key not in st.session_state: st.session_state[unselected_key] = set()
                                id_col = 'SO #' if 'SO #' in df_edit.columns else None
                                
                                c1, c2 = st.columns(2)
                                with c1:
                                    if st.button("All", key=f"q1_row_sel_{key}_{rep_name}"):
                                        st.session_state[unselected_key] = set()
                                        st.rerun()
                                with c2:
                                    if st.button("None", key=f"q1_row_unsel_{key}_{rep_name}"):
                                        if id_col: st.session_state[unselected_key] = set(df_edit[id_col].astype(str).tolist())
                                        st.rerun()
                                
                                if id_col:
                                    df_edit.insert(0, "Select", df_edit[id_col].apply(lambda x: str(x) not in st.session_state[unselected_key]))
                                else: df_edit.insert(0, "Select", True)
                                
                                edited = st.data_editor(
                                    df_edit[['Select'] + display_cols],
                                    column_config={
                                        "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                        "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                        "Amount": st.column_config.NumberColumn("Amount", format="$%d")
                                    },
                                    disabled=display_cols, hide_index=True, key=f"q1_edit_{key}_{rep_name}"
                                )
                                if id_col:
                                    current_unselected = set()
                                    for idx, row in edited.iterrows():
                                        if not row['Select']: current_unselected.add(str(row[id_col]))
                                    st.session_state[unselected_key] = current_unselected
                                selected_rows = df.loc[edited[edited['Select']].index].copy()
                                export_buckets[key] = selected_rows
                            else:
                                st.dataframe(df[display_cols], column_config={"Link": st.column_config.LinkColumn("🔗", display_text="Open"), "Amount": st.column_config.NumberColumn("Amount", format="$%d")}, hide_index=True, use_container_width=True)
                                export_buckets[key] = df
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='opacity:0.5; margin-top:5px; font-size:0.9em;'>{data['label']}: $0</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # HubSpot Logic (Visuals Wrapped)
    with col_hs:
        st.markdown('<div class="glass-panel" style="border-top: 3px solid #60a5fa;">', unsafe_allow_html=True)
        st.markdown('<h3 style="color:#60a5fa; margin-bottom: 5px;">🎯 HubSpot Opportunities</h3>', unsafe_allow_html=True)
        st.caption("Active pipeline deals for Q1 close")
        
        for key, data in hs_categories.items():
            df = hs_dfs.get(key, pd.DataFrame())
            val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
            checkbox_key = f"q1_chk_{key}_{rep_name}"
            
            if val > 0:
                st.markdown(f"<div style='margin-top:10px; padding: 8px; background: rgba(255,255,255,0.03); border-radius: 6px;'>", unsafe_allow_html=True)
                is_checked = st.checkbox(f"**{data['label']}** — ${val:,.0f}", key=checkbox_key)
                if is_checked:
                    with st.expander("Details"):
                        if not df.empty:
                            enable_edit = st.toggle("Custom Select", key=f"q1_tgl_{key}_{rep_name}")
                            display_cols = ['Link', 'Deal ID', 'Deal Name', 'Close', 'Amount_Numeric']
                            if 'PA Date' in df.columns: display_cols.insert(4, 'PA Date')
                            
                            if enable_edit:
                                df_edit = df.copy()
                                unselected_key = f"q1_unselected_{key}_{rep_name}"
                                if unselected_key not in st.session_state: st.session_state[unselected_key] = set()
                                id_col = 'Deal ID' if 'Deal ID' in df_edit.columns else None
                                
                                c1, c2 = st.columns(2)
                                with c1:
                                    if st.button("All", key=f"q1_row_sel_{key}_{rep_name}"): st.session_state[unselected_key] = set(); st.rerun()
                                with c2:
                                    if st.button("None", key=f"q1_row_unsel_{key}_{rep_name}"):
                                        if id_col: st.session_state[unselected_key] = set(df_edit[id_col].astype(str).tolist())
                                        st.rerun()
                                
                                if id_col: df_edit.insert(0, "Select", df_edit[id_col].apply(lambda x: str(x) not in st.session_state[unselected_key]))
                                else: df_edit.insert(0, "Select", True)
                                
                                avail_cols = ['Select'] + [c for c in display_cols if c in df_edit.columns]
                                edited = st.data_editor(
                                    df_edit[avail_cols],
                                    column_config={
                                        "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                        "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                        "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")
                                    },
                                    disabled=[c for c in avail_cols if c != 'Select'], hide_index=True, key=f"q1_edit_{key}_{rep_name}"
                                )
                                if id_col:
                                    current_unselected = set()
                                    for idx, row in edited.iterrows():
                                        if not row['Select']: current_unselected.add(str(row[id_col]))
                                    st.session_state[unselected_key] = current_unselected
                                selected_rows = df.loc[edited[edited['Select']].index].copy()
                                export_buckets[key] = selected_rows
                            else:
                                avail_cols = [c for c in display_cols if c in df.columns]
                                st.dataframe(df[avail_cols], column_config={"Link": st.column_config.LinkColumn("🔗", display_text="Open"), "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")}, hide_index=True, use_container_width=True)
                                export_buckets[key] = df
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div style='opacity:0.5; margin-top:5px; font-size:0.9em;'>{data['label']}: $0</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION: REORDER AI
    # ═══════════════════════════════════════════════════════════════════════════

    st.markdown('<div class="section-title">02 // REORDER PREDICTION</div>', unsafe_allow_html=True)
    
    reorder_buckets = {}
    
    with st.spinner("Processing historical patterns..."):
        if is_team_view:
            all_historical, all_invoices = [], []
            for r in active_team_reps:
                rh = load_historical_orders(main_dash, r)
                ri = load_invoices(main_dash, r)
                if not rh.empty: rh['Rep'] = r; all_historical.append(rh)
                if not ri.empty: all_invoices.append(ri)
            historical_df = pd.concat(all_historical, ignore_index=True) if all_historical else pd.DataFrame()
            invoices_df = pd.concat(all_invoices, ignore_index=True) if all_invoices else pd.DataFrame()
        else:
            historical_df = load_historical_orders(main_dash, rep_name)
            invoices_df = load_invoices(main_dash, rep_name)
            if not historical_df.empty: historical_df['Rep'] = rep_name
            
        if not historical_df.empty: historical_df = merge_orders_with_invoices(historical_df, invoices_df)
        line_items_df = load_line_items(main_dash)
        sku_to_desc = load_item_master(main_dash)

    if historical_df.empty:
        st.info("No 2025 historical data found.")
    elif line_items_df.empty:
        st.warning("Line item data missing.")
    else:
        pending_customers, pipeline_customers = set(), set()
        for key in ns_categories.keys():
            df = ns_dfs.get(key, pd.DataFrame())
            if not df.empty and 'Customer' in df.columns: pending_customers.update(df['Customer'].dropna().tolist())
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            if not df.empty:
                for col in ['Account Name', 'Associated Company', 'Company', 'Deal Name']:
                    if col in df.columns: pipeline_customers.update(df[col].dropna().tolist()); break
        
        def normalize(name): return str(name).lower().strip() if pd.notna(name) else ''
        excluded_customers = {normalize(c) for c in pending_customers | pipeline_customers}
        excluded_customers.discard('')
        
        product_metrics_df = calculate_customer_product_metrics(historical_df, line_items_df, sku_to_desc)
        
        if not product_metrics_df.empty:
            product_metrics_df['Customer_Normalized'] = product_metrics_df['Customer'].apply(normalize)
            opportunities_df = product_metrics_df[~product_metrics_df['Customer_Normalized'].isin(excluded_customers)].copy()
            
            if opportunities_df.empty:
                st.success("All customers active in pipeline!")
            else:
                # REORDER UI
                st.markdown('<div class="glass-panel" style="border-top: 3px solid #fbbf24;">', unsafe_allow_html=True)
                
                # Metrics Row
                m1, m2, m3 = st.columns(3)
                amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
                with m1: st.metric("2025 Revenue Base", f"${historical_df[amount_col].sum():,.0f}")
                with m2: st.metric("Available Accounts", f"{opportunities_df['Customer'].nunique()}")
                with m3: st.metric("Est. Q1 Potential", f"${opportunities_df['Q1_Forecast'].sum():,.0f}")
                
                st.markdown("---")
                
                # Controls
                col_search, col_act = st.columns([2, 1])
                with col_search:
                    search_term = st.text_input("Search Accounts", placeholder="Type to filter...", key=f"reorder_search_{rep_name}")
                with col_act:
                    c1, c2 = st.columns(2)
                    with c1: 
                        if st.button("Chk All Buckets", use_container_width=True):
                            for pt in opportunities_df['Product_Type'].unique(): st.session_state[f"q1_reorder_pt_{pt}_{rep_name}"] = True
                            st.rerun()
                    with c2:
                         if st.button("Unchk All", use_container_width=True):
                            for pt in opportunities_df['Product_Type'].unique(): st.session_state[f"q1_reorder_pt_{pt}_{rep_name}"] = False
                            st.rerun()
                
                if search_term:
                    filtered_opportunities = opportunities_df[opportunities_df['Customer'].str.lower().str.contains(search_term.lower(), na=False)].copy()
                else: filtered_opportunities = opportunities_df.copy()
                
                product_type_totals = filtered_opportunities.groupby('Product_Type')['Total_Revenue'].sum().sort_values(ascending=False)
                
                # ITERATE PRODUCT TYPES
                for product_type in product_type_totals.index:
                    pt_data = filtered_opportunities[filtered_opportunities['Product_Type'] == product_type].copy()
                    if pt_data.empty: continue
                    
                    pt_projected = pt_data['Q1_Forecast'].sum()
                    pt_customers = pt_data['Customer'].nunique()
                    
                    st.markdown(f"<div style='margin-top:15px; background:rgba(0,0,0,0.2); border-radius:8px; padding:10px; border:1px solid rgba(255,255,255,0.05)'>", unsafe_allow_html=True)
                    
                    col_chk, col_info = st.columns([0.05, 0.95])
                    with col_chk:
                        is_checked = st.checkbox("", key=f"q1_reorder_pt_{product_type}_{rep_name}")
                    with col_info:
                        st.markdown(f"<span style='font-size:1.1em; font-weight:700'>{product_type}</span> <span style='color:#64748b; font-size:0.9em'>({pt_customers} accounts)</span>", unsafe_allow_html=True)
                        st.caption(f"Projected: ${pt_projected:,.0f}")
                    
                    if is_checked:
                        with st.expander("Expand Accounts", expanded=True):
                            select_key = f"q1_reorder_select_{product_type}_{rep_name}"
                            edited_key = f"q1_products_{product_type}_{rep_name}"
                            if select_key not in st.session_state: st.session_state[select_key] = set()
                            if edited_key not in st.session_state: st.session_state[edited_key] = {}
                            
                            display_data = []
                            for _, row in pt_data.iterrows():
                                key = f"{row['Customer']}|{row['Product_Type']}"
                                is_selected = key not in st.session_state[select_key]
                                q1_value = st.session_state[edited_key].get(key, int(row['Q1_Value']))
                                conf_map = {'Likely': '🟢', 'Possible': '🟡', 'Long Shot': '⚪'}
                                
                                display_data.append({
                                    'Select': is_selected,
                                    'Tier': conf_map.get(row['Confidence_Tier'], '⚪'),
                                    'Customer': row['Customer'],
                                    'Top SKUs': row.get('Top_SKUs', ''),
                                    '2025 $': int(row['Total_Revenue']),
                                    'Q1 Forecast ✏️': q1_value,
                                    '_key': key, '_conf_pct': row['Confidence_Pct']
                                })
                            
                            df_disp = pd.DataFrame(display_data)
                            edited_df = st.data_editor(
                                df_disp[['Select', 'Tier', 'Customer', 'Top SKUs', '2025 $', 'Q1 Forecast ✏️']],
                                column_config={
                                    "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                    "Tier": st.column_config.TextColumn("Prob", width="small"),
                                    "2025 $": st.column_config.NumberColumn(format="$%d"),
                                    "Q1 Forecast ✏️": st.column_config.NumberColumn(format="$%d", min_value=0)
                                },
                                hide_index=True, use_container_width=True, key=f"q1_edit_pt_{product_type}_{rep_name}"
                            )
                            
                            # Process Edits
                            new_unselected = set()
                            export_data = []
                            for idx, row in edited_df.iterrows():
                                key = display_data[idx]['_key']
                                val = row['Q1 Forecast ✏️']
                                st.session_state[edited_key][key] = val
                                if not row['Select']: new_unselected.add(key)
                                else:
                                    export_data.append({
                                        'Customer': display_data[idx]['Customer'],
                                        'Product_Type': product_type,
                                        'Top_SKUs': display_data[idx]['Top SKUs'],
                                        'Q1_Value': val,
                                        'Projected_Value': val * display_data[idx]['_conf_pct']
                                    })
                            
                            st.session_state[select_key] = new_unselected
                            if export_data: reorder_buckets[f"reorder_{product_type}"] = pd.DataFrame(export_data)
                    st.markdown("</div>", unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════════
    # CALCULATION & HUD
    # ═══════════════════════════════════════════════════════════════════════════

    def safe_sum(df): return df['Amount_Numeric'].sum() if 'Amount_Numeric' in df.columns else (df['Amount'].sum() if 'Amount' in df.columns else 0) if not df.empty else 0
    selected_scheduled = sum(safe_sum(df) for k, df in export_buckets.items() if k in ns_categories)
    selected_pipeline = sum(safe_sum(df) for k, df in export_buckets.items() if k in hs_categories)
    selected_reorder = sum(df['Projected_Value'].sum() for df in reorder_buckets.values()) if reorder_buckets else 0
    total_forecast = selected_scheduled + selected_pipeline + selected_reorder
    gap_to_goal = q1_goal - total_forecast
    
    # HUD Footer
    gap_color = "c-red" if gap_to_goal > 0 else "c-green"
    gap_label = "GAP" if gap_to_goal > 0 else "SURPLUS"
    
    st.markdown(f"""
    <div class="hud-footer">
        <div class="hud-item">
            <div class="hud-label">Scheduled</div>
            <div class="hud-val c-green">${selected_scheduled:,.0f}</div>
        </div>
        <div class="hud-item">
            <div class="hud-label">Pipeline</div>
            <div class="hud-val c-blue">${selected_pipeline:,.0f}</div>
        </div>
        <div class="hud-item">
            <div class="hud-label">Reorder</div>
            <div class="hud-val c-amber">${selected_reorder:,.0f}</div>
        </div>
        <div class="hud-item" style="border-left: 1px solid rgba(255,255,255,0.1); padding-left: 20px;">
            <div class="hud-label">Forecast</div>
            <div class="hud-val" style="color:white; font-size:1.4rem">${total_forecast:,.0f}</div>
        </div>
        <div class="hud-item">
            <div class="hud-label">{gap_label}</div>
            <div class="hud-val {gap_color}">${abs(gap_to_goal):,.0f}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION: FINAL SUMMARY
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown('<div class="section-title">03 // EXECUTIVE SUMMARY</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="glass-panel">', unsafe_allow_html=True)
    col_chart, col_txt = st.columns([1.5, 1])
    
    with col_chart:
        fig = create_q1_gauge(total_forecast, q1_goal, "Q1 Target")
        st.plotly_chart(fig, use_container_width=True)
    
    with col_txt:
        st.markdown(f"""
        <div style="padding-top:20px;">
            <div style="margin-bottom:15px; padding-bottom:15px; border-bottom:1px solid rgba(255,255,255,0.1);">
                <span class="status-badge badge-likely">CONFIRMED</span>
                <div style="display:flex; justify-content:space-between; margin-top:5px;">
                    <span>NetSuite Orders</span>
                    <span style="font-weight:700">${selected_scheduled:,.0f}</span>
                </div>
            </div>
            <div style="margin-bottom:15px; padding-bottom:15px; border-bottom:1px solid rgba(255,255,255,0.1);">
                <span class="status-badge badge-possible">ACTIVE</span>
                <div style="display:flex; justify-content:space-between; margin-top:5px;">
                    <span>Pipeline Deals</span>
                    <span style="font-weight:700">${selected_pipeline:,.0f}</span>
                </div>
            </div>
            <div style="margin-bottom:15px;">
                <span class="status-badge badge-longshot">PROJECTED</span>
                <div style="display:flex; justify-content:space-between; margin-top:5px;">
                    <span>Reorder Pattern</span>
                    <span style="font-weight:700">${selected_reorder:,.0f}</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Export Logic (Preserved)
    if total_forecast > 0:
        export_summary = []
        export_data = []
        export_summary.extend([
            {'Category': '=== Q1 2026 FORECAST ===', 'Amount': ''},
            {'Category': 'Q1 Goal', 'Amount': f"${q1_goal:,.0f}"},
            {'Category': 'Scheduled', 'Amount': f"${selected_scheduled:,.0f}"},
            {'Category': 'Pipeline', 'Amount': f"${selected_pipeline:,.0f}"},
            {'Category': 'Reorder', 'Amount': f"${selected_reorder:,.0f}"},
            {'Category': 'Total', 'Amount': f"${total_forecast:,.0f}"},
            {'Category': 'Gap', 'Amount': f"${gap_to_goal:,.0f}"},
            {'Category': '', 'Amount': ''}
        ])
        
        # [Export Loop - Logic preserved exactly...]
        for key, df in export_buckets.items():
            if df.empty: continue
            cat_val = df['Amount_Numeric'].sum() if 'Amount_Numeric' in df.columns else (df['Amount'].sum() if 'Amount' in df.columns else 0)
            if cat_val > 0:
                label = ns_categories.get(key, {}).get('label', hs_categories.get(key, {}).get('label', key))
                export_summary.append({'Category': f"{label}", 'Amount': f"${cat_val:,.0f}"})
                
                for _, row in df.iterrows():
                    if key in ns_categories:
                         export_data.append({
                            'Category': f"NS - {label}", 'ID': row.get('SO #', ''), 'Customer': row.get('Customer', ''),
                            'Amount': row.get('Amount', 0), 'Rep': row.get('Rep Master', '')
                        })
                    else:
                        export_data.append({
                            'Category': f"HS - {label}", 'ID': row.get('Deal ID', ''), 'Customer': row.get('Account Name', ''),
                            'Amount': row.get('Amount_Numeric', 0), 'Rep': row.get('Deal Owner', '')
                        })
                        
        if reorder_buckets:
            for key, df in reorder_buckets.items():
                if df.empty: continue
                export_summary.append({'Category': f"Reorder - {key}", 'Amount': f"${df['Projected_Value'].sum():,.0f}"})
                for _, row in df.iterrows():
                    export_data.append({
                        'Category': 'Reorder', 'ID': '', 'Customer': row.get('Customer', ''),
                        'Amount': row.get('Projected_Value', 0), 'Rep': rep_name
                    })

        summary_df = pd.DataFrame(export_summary)
        data_df = pd.DataFrame(export_data)
        final_csv = summary_df.to_csv(index=False) + "\n" + data_df.to_csv(index=False)
        
        st.download_button(
            label="📥 EXPORT FORECAST CSV",
            data=final_csv,
            file_name=f"q1_2026_forecast_{rep_name.replace(' ', '_')}.csv",
            mime="text/csv",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
