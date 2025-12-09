"""
Sales Forecasting Dashboard - Enhanced Version with Drill-Down Capability
Reads from Google Sheets and displays gap-to-goal analysis with interactive visualizations
Includes lead time logic for Q4/Q1 fulfillment determination and detailed order drill-downs
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
# Optional: Commission calculator module (if available)
try:
    import commission_calculator
    COMMISSION_AVAILABLE = True
except ImportError:
    COMMISSION_AVAILABLE = False

# Optional: Shipping planning module (if available)
try:
    import shipping_planning
    from importlib import reload
    reload(shipping_planning)  # Force reload to pick up any changes
    SHIPPING_PLANNING_AVAILABLE = True
except ImportError as e:
    SHIPPING_PLANNING_AVAILABLE = False
    SHIPPING_PLANNING_ERROR = str(e)
except Exception as e:
    SHIPPING_PLANNING_AVAILABLE = False
    SHIPPING_PLANNING_ERROR = f"Error loading module: {str(e)}"

# Optional: All Products Forecast module (if available)
try:
    import all_products_forecast
    from importlib import reload
    reload(all_products_forecast)  # Force reload to pick up any changes
    ALL_PRODUCTS_FORECAST_AVAILABLE = True
except ImportError as e:
    ALL_PRODUCTS_FORECAST_AVAILABLE = False
    ALL_PRODUCTS_FORECAST_ERROR = str(e)
except Exception as e:
    ALL_PRODUCTS_FORECAST_AVAILABLE = False
    ALL_PRODUCTS_FORECAST_ERROR = f"Error loading module: {str(e)}"

# Configure Plotly for dark mode compatibility
pio.templates.default = "plotly"  # Use default template that adapts to theme

# Helper function for business days calculation
def calculate_business_days_remaining():
    """
    Calculate business days from today through end of Q4 2025 (Dec 31)
    Excludes weekends and major holidays
    """
    from datetime import date, timedelta
    
    today = date.today()
    q4_end = date(2025, 12, 31)
    
    # Define holidays to exclude
    holidays = [
        date(2025, 11, 27),  # Thanksgiving
        date(2025, 11, 28),  # Day after Thanksgiving
        date(2025, 12, 25),  # Christmas
        date(2025, 12, 26),  # Day after Christmas
    ]
    
    business_days = 0
    current_date = today
    
    while current_date <= q4_end:
        # Check if it's a weekday (Monday=0, Sunday=6)
        if current_date.weekday() < 5 and current_date not in holidays:
            business_days += 1
        current_date += timedelta(days=1)
    
    return business_days

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
    /* ========================================
       GOOGLE AI STUDIO INSPIRED STYLING
       Modern glass-morphism with depth
       ======================================== */
    
    /* App-wide theming */
    .stApp {
        color-scheme: light dark;
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.05) 0%, rgba(139, 92, 246, 0.05) 100%);
    }
    
    /* Smooth transitions for everything */
    * {
        transition: all 0.3s ease !important;
    }
    
    /* ========== METRIC CARDS ========== */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.08) 0%, rgba(255, 255, 255, 0.02) 100%);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        padding: 20px;
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.15);
        overflow: visible !important;
        min-width: 140px !important;
    }
    
    [data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 40px 0 rgba(0, 0, 0, 0.25);
        border-color: rgba(59, 130, 246, 0.3);
    }
    
    [data-testid="stMetricValue"] {
        font-size: 1.8rem !important;
        font-weight: 700 !important;
        background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        overflow: visible !important;
        letter-spacing: -0.5px;
    }
    
    [data-testid="stMetricLabel"] {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.7;
        overflow: visible !important;
        white-space: normal !important;
        line-height: 1.2 !important;
    }
    
    [data-testid="stMetricDelta"] {
        font-size: 0.875rem !important;
        font-weight: 600 !important;
        white-space: normal !important;
    }
    
    /* ========== SIDEBAR STYLING ========== */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.95) 0%, rgba(30, 41, 59, 0.95) 100%);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
        color: rgba(255, 255, 255, 0.9);
    }
    
    /* ========== BUTTONS ========== */
    .stButton > button {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 12px 24px;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(59, 130, 246, 0.4);
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
    }
    
    /* ========== DATAFRAMES / TABLES ========== */
    [data-testid="stDataFrame"], .stDataFrame {
        background: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
        overflow: hidden;
    }
    
    /* Table headers */
    [data-testid="stDataFrame"] thead tr {
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.1) 0%, rgba(139, 92, 246, 0.1) 100%);
        backdrop-filter: blur(10px);
    }
    
    [data-testid="stDataFrame"] thead th {
        font-weight: 700;
        text-transform: uppercase;
        font-size: 0.75rem;
        letter-spacing: 0.5px;
        padding: 16px 12px;
        border-bottom: 2px solid rgba(59, 130, 246, 0.2);
    }
    
    /* Table rows */
    [data-testid="stDataFrame"] tbody tr:hover {
        background: rgba(59, 130, 246, 0.05);
        transform: scale(1.01);
    }
    
    [data-testid="stDataFrame"] tbody td {
        padding: 12px;
        border-bottom: 1px solid rgba(255, 255, 255, 0.05);
    }
    
    /* ========== EXPANDERS ========== */
    [data-testid="stExpander"] {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(255, 255, 255, 0.02) 100%);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
        margin: 12px 0;
    }
    
    [data-testid="stExpander"]:hover {
        border-color: rgba(59, 130, 246, 0.3);
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
    }
    
    /* ========== PROGRESS BREAKDOWN ========== */
    .progress-breakdown {
        background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
        padding: 28px;
        border-radius: 16px;
        color: white !important;
        margin: 24px 0;
        box-shadow: 0 12px 32px rgba(59, 130, 246, 0.3);
        backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .progress-breakdown:hover {
        transform: translateY(-2px);
        box-shadow: 0 16px 40px rgba(59, 130, 246, 0.4);
    }
    
    .progress-breakdown h3 {
        color: white !important;
        margin-bottom: 20px;
        font-size: 24px;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    
    .progress-item {
        display: flex;
        justify-content: space-between;
        padding: 14px 0;
        border-bottom: 1px solid rgba(255, 255, 255, 0.15);
        color: white !important;
        transition: all 0.3s ease;
    }
    
    .progress-item:hover {
        padding-left: 8px;
        border-bottom-color: rgba(255, 255, 255, 0.3);
    }
    
    .progress-item:last-child {
        border-bottom: none;
        font-weight: 700;
        font-size: 20px;
        padding-top: 18px;
        margin-top: 8px;
        border-top: 2px solid rgba(255, 255, 255, 0.3);
    }
    
    .progress-label {
        font-size: 16px;
        color: rgba(255, 255, 255, 0.95) !important;
        font-weight: 500;
    }
    
    .progress-value {
        font-size: 16px;
        font-weight: 700;
        color: white !important;
        font-family: 'SF Mono', 'Courier New', monospace;
    }
    
    /* ========== SECTION HEADERS ========== */
    .section-header {
        background: linear-gradient(135deg, rgba(59, 130, 246, 0.1) 0%, rgba(139, 92, 246, 0.1) 100%);
        backdrop-filter: blur(10px);
        padding: 16px 20px;
        border-radius: 12px;
        margin: 20px 0;
        font-weight: 700;
        border: 1px solid rgba(59, 130, 246, 0.2);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        letter-spacing: -0.3px;
    }
    
    .section-header:hover {
        border-color: rgba(59, 130, 246, 0.4);
        transform: translateX(4px);
    }
    
    /* ========== DRILL-DOWN SECTIONS ========== */
    .drill-down-section {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(255, 255, 255, 0.02) 100%);
        backdrop-filter: blur(10px);
        padding: 16px;
        border-radius: 12px;
        margin: 12px 0;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }
    
    .drill-down-section:hover {
        border-color: rgba(59, 130, 246, 0.2);
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.12);
    }
    
    /* ========== AUDIT SECTIONS ========== */
    .audit-section {
        background: linear-gradient(135deg, rgba(34, 197, 94, 0.08) 0%, rgba(59, 130, 246, 0.08) 100%);
        backdrop-filter: blur(10px);
        padding: 24px;
        border-radius: 12px;
        margin: 20px 0;
        border: 1px solid rgba(34, 197, 94, 0.2);
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }
    
    /* ========== CHANGE TRACKING ========== */
    .change-positive {
        color: #10b981;
        font-weight: 700;
        text-shadow: 0 0 8px rgba(16, 185, 129, 0.3);
    }
    
    .change-negative {
        color: #ef4444;
        font-weight: 700;
        text-shadow: 0 0 8px rgba(239, 68, 68, 0.3);
    }
    
    .change-neutral {
        color: #6b7280;
        font-weight: 600;
    }
    
    /* ========== RECONCILIATION TABLE ========== */
    .reconciliation-table {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.05) 0%, rgba(255, 255, 255, 0.02) 100%);
        backdrop-filter: blur(10px);
        padding: 20px;
        border-radius: 12px;
        margin: 16px 0;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
    }
    
    /* ========== TABS ========== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: rgba(255, 255, 255, 0.03);
        padding: 8px;
        border-radius: 12px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: transparent;
        border-radius: 8px;
        padding: 12px 24px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(59, 130, 246, 0.1);
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }
    
    /* ========== SELECT BOXES ========== */
    .stSelectbox > div > div {
        background: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .stSelectbox > div > div:hover {
        border-color: rgba(59, 130, 246, 0.3);
    }
    
    /* ========== RESPONSIVE ========== */
    @media (max-width: 768px) {
        [data-testid="stMetricValue"] {
            font-size: 1.4rem !important;
        }
        
        [data-testid="stMetric"] {
            margin-bottom: 12px !important;
            padding: 16px;
        }
        
        .progress-breakdown {
            padding: 20px;
        }
    }
    
    /* ========== ANIMATIONS ========== */
    @keyframes fadeIn {
        from {
            opacity: 0;
            transform: translateY(10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateX(-20px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    [data-testid="stMetric"],
    [data-testid="stExpander"],
    .drill-down-section {
        animation: fadeIn 0.5s ease-out;
    }
    
    /* ========== SCROLLBAR ========== */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(15, 23, 42, 0.3);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
        border-radius: 10px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
    }
    
    /* ========== TEXT SELECTION ========== */
    ::selection {
        background: rgba(59, 130, 246, 0.3);
        color: white;
    }
    
    /* ========== GLOWING CARD EFFECTS (GEMINI ENHANCEMENT) ========== */
    .metric-card-glow {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        position: relative;
        overflow: hidden;
        transition: all 0.3s ease;
    }
    
    .metric-card-glow:hover {
        transform: translateY(-4px);
    }
    
    .glow-green {
        box-shadow: 0 0 20px rgba(16, 185, 129, 0.3);
        border: 1px solid rgba(16, 185, 129, 0.5);
        animation: pulseGreen 2s infinite;
    }
    
    .glow-red {
        box-shadow: 0 0 20px rgba(239, 68, 68, 0.3);
        border: 1px solid rgba(239, 68, 68, 0.5);
        animation: pulseRed 2s infinite;
    }
    
    .glow-blue {
        box-shadow: 0 0 20px rgba(59, 130, 246, 0.3);
        border: 1px solid rgba(59, 130, 246, 0.5);
        animation: pulseBlue 2s infinite;
    }
    
    @keyframes pulseGreen {
        0%, 100% { box-shadow: 0 0 20px rgba(16, 185, 129, 0.3); }
        50% { box-shadow: 0 0 30px rgba(16, 185, 129, 0.5); }
    }
    
    @keyframes pulseRed {
        0%, 100% { box-shadow: 0 0 20px rgba(239, 68, 68, 0.3); }
        50% { box-shadow: 0 0 30px rgba(239, 68, 68, 0.5); }
    }
    
    @keyframes pulseBlue {
        0%, 100% { box-shadow: 0 0 20px rgba(59, 130, 246, 0.3); }
        50% { box-shadow: 0 0 30px rgba(59, 130, 246, 0.5); }
    }
    
    .metric-label {
        font-size: 0.8rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        opacity: 0.8;
        font-weight: 600;
    }
    
    .metric-value {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(to right, #fff, #ccc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin: 10px 0;
    }
    
    /* ========== LOCKED REVENUE BANNER ========== */
    .locked-revenue-banner {
        background: linear-gradient(90deg, #1e3a8a 0%, #172554 100%);
        padding: 15px 25px;
        border-radius: 12px;
        border-left: 6px solid #3b82f6;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    
    .locked-revenue-banner:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 12px -1px rgba(0, 0, 0, 0.2);
    }
    
    .banner-left {
        display: flex;
        align-items: center;
        gap: 12px;
    }
    
    .banner-icon {
        font-size: 24px;
    }
    
    .banner-label {
        font-size: 12px;
        text-transform: uppercase;
        opacity: 0.8;
        letter-spacing: 1px;
    }
    
    .banner-value {
        font-size: 24px;
        font-weight: 700;
        color: white;
    }
    
    .banner-right {
        text-align: right;
    }
    
    .banner-status {
        color: #4ade80;
        font-weight: 600;
    }
    </style>
    """, unsafe_allow_html=True)

# Google Sheets Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# Cache duration - 1 hour
CACHE_TTL = 3600

# Add a version number to force cache refresh when code changes
CACHE_VERSION = "v61_fix_dec31_timestamp"

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
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            st.warning(f"⚠️ No data found in {sheet_name}!{range_name}")
            return pd.DataFrame()
        
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
        error_msg = str(e)
        st.error(f"❌ Error loading data from {sheet_name}: {error_msg}")
        
        # Provide specific troubleshooting based on error type
        if "403" in error_msg or "permission" in error_msg.lower():
            st.warning("""
            **Permission Error:**
            - Make sure you've shared the Google Sheet with your service account email
            - The service account email looks like: `your-service-account@project.iam.gserviceaccount.com`
            - Share the sheet with 'Viewer' access
            """)
        elif "404" in error_msg or "not found" in error_msg.lower():
            st.warning("""
            **Sheet Not Found:**
            - Check that the spreadsheet ID is correct
            - Check that the sheet name matches exactly (case-sensitive)
            - Current spreadsheet ID: `12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk`
            """)
        elif "401" in error_msg or "authentication" in error_msg.lower():
            st.warning("""
            **Authentication Error:**
            - Your service account credentials may be invalid
            - Try regenerating the service account key in Google Cloud Console
            """)
        
        return pd.DataFrame()

def apply_q4_fulfillment_logic(deals_df):
    """
    Apply lead time logic to filter out deals that close late in Q4 
    but won't ship until Q1 based on product type
    """
    # Lead time mapping based on your image
    lead_time_map = {
        'Labeled - Labels In Stock': 10,
        'Outer Boxes': 20,
        'Non-Labeled - 1 Week Lead Time': 5,
        'Non-Labeled - 2 Week Lead Time': 10,
        'Labeled - Print & Apply': 20,
        'Non-Labeled - Custom Lead Time': 30,
        'Labeled with FEP - Print & Apply': 35,
        'Labeled - Custom Lead Time': 40,
        'Flexpack': 25,
        'Labels Only - Direct to Customer': 15,
        'Labels Only - For Inventory': 15,
        'Labeled with FEP - Labels In Stock': 25,
        'Labels Only (deprecated)': 15
    }
    
    # Calculate cutoff date for each product type
    q4_end = pd.Timestamp('2025-12-31')
    
    def get_business_days_before(end_date, business_days):
        """Calculate date that is N business days before end_date"""
        current = end_date
        days_counted = 0
        
        while days_counted < business_days:
            current -= timedelta(days=1)
            # Skip weekends (Monday=0, Sunday=6)
            if current.weekday() < 5:
                days_counted += 1
        
        return current
    
    # Add a column to track if deal counts for Q4
    deals_df['Counts_In_Q4'] = True
    deals_df['Q1_Spillover_Amount'] = 0
    
    # Check if we have a Product Type column
    if 'Product Type' in deals_df.columns:
        for product_type, lead_days in lead_time_map.items():
            cutoff_date = get_business_days_before(q4_end, lead_days)
            
            # Mark deals closing after cutoff as Q1
            mask = (
                (deals_df['Product Type'] == product_type) & 
                (deals_df['Close Date'] > cutoff_date) &
                (deals_df['Close Date'].notna())
            )
            deals_df.loc[mask, 'Counts_In_Q4'] = False
            deals_df.loc[mask, 'Q1_Spillover_Amount'] = deals_df.loc[mask, 'Amount']
            
        # Log how many deals were excluded
        excluded_count = (~deals_df['Counts_In_Q4']).sum()
        excluded_value = deals_df[~deals_df['Counts_In_Q4']]['Amount'].sum()
        
        if excluded_count > 0:
            pass  # Debug info removed
            #st.sidebar.info(f"📊 {excluded_count} deals (${excluded_value:,.0f}) deferred to Q1 2026 due to lead times")
    else:
        pass  # Debug info removed
        #st.sidebar.warning("⚠️ No 'Product Type' column found - lead time logic not applied")
    
    return deals_df

def load_all_data():
    """Load all necessary data from Google Sheets"""
    
    #st.sidebar.info("🔄 Loading data from Google Sheets...")
    
    # Load deals data - extend range to include Q1 2026 Spillover column
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
    # DEBUG: Show what we got from HubSpot
    if not deals_df.empty:
        pass  # Debug info removed
        #st.sidebar.success(f"📊 HubSpot raw data: {len(deals_df)} rows, {len(deals_df.columns)} columns")
        pass  # Debug info removed
    else:
        pass  # Debug info removed
        #st.sidebar.error("❌ No HubSpot data loaded!")
        pass
    
    # Load dashboard info (rep quotas and orders)
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoice data from NetSuite - EXTEND to include Columns T:U (Corrected Customer Name, Rep Master)
    invoices_df = load_google_sheets_data("NS Invoices", "A:U", version=CACHE_VERSION)
    
    # Load sales orders data from NetSuite - EXTEND to include Columns through AF (Calyx | External Order, Pending Approval Date, Corrected Customer Name, Rep Master)
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AF", version=CACHE_VERSION)
    
    # Clean and process deals data - FIXED VERSION to match actual sheet
    if not deals_df.empty and len(deals_df.columns) >= 6:
        # Get column names from first row
        if len(deals_df) > 0:
            # Get actual column names
            col_names = deals_df.columns.tolist()
            
            #st.sidebar.info(f"Processing {len(col_names)} HubSpot columns")
            #st.sidebar.info(f"First 10 columns: {col_names[:10]}")
            
            # Map based on ACTUAL column names from your sheet
            # Note: Column 4 appears to be "Deal Owner First Name Deal Owner Last Name" combined
            
            rename_dict = {}
            
            # Map columns by actual names (case-sensitive)
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
                    # This column has both names already combined
                    rename_dict[col] = 'Deal Owner'
                elif col == 'Deal Owner First Name':
                    rename_dict[col] = 'Deal Owner First Name'
                elif col == 'Deal Owner Last Name':
                    rename_dict[col] = 'Deal Owner Last Name'
                elif col == 'Amount':
                    rename_dict[col] = 'Amount'
                elif col == 'Close Status':
                    rename_dict[col] = 'Status'  # Map Close Status to Status
                elif col == 'Pipeline':
                    rename_dict[col] = 'Pipeline'
                elif col == 'Deal Type':
                    rename_dict[col] = 'Product Type'  # Map Deal Type to Product Type for lead time logic
                elif col == 'Average Leadtime':
                    rename_dict[col] = 'Average Leadtime'
                elif col == 'Q1 2026 Spillover':
                    rename_dict[col] = 'Q1 2026 Spillover'
            
            deals_df = deals_df.rename(columns=rename_dict)
            
            # Check if Deal Owner already exists (from combined column)
            if 'Deal Owner' not in deals_df.columns:
                # Create a combined "Deal Owner" field from First Name + Last Name if they're separate
                if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                    deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                    deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
                    #st.sidebar.success("✅ Created Deal Owner from First + Last Name")
                else:
                    pass  # Debug info removed
                    #st.sidebar.error("❌ Missing Deal Owner column!")
            else:
                pass  # Debug info removed
                #st.sidebar.success("✅ Deal Owner column already exists")
                # Clean up the Deal Owner field
                deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
            
            # Show what we have after renaming
            #st.sidebar.success(f"✅ Columns after rename: {', '.join([c for c in deals_df.columns.tolist()[:10] if c])}")
            
            # Check if we have required columns
            required_cols = ['Deal Name', 'Status', 'Close Date', 'Deal Owner', 'Amount', 'Pipeline']
            missing_cols = [col for col in required_cols if col not in deals_df.columns]
            if missing_cols:
                pass  # Debug info removed
                #st.sidebar.error(f"❌ Missing required columns: {missing_cols}")
            
            # Clean and convert amount to numeric
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
            else:
                pass  # Debug info removed
                #st.sidebar.error("❌ No Amount column found!")
            
            # Convert close date to datetime
            if 'Close Date' in deals_df.columns:
                deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
                
                # Debug: Show date range in the data
                valid_dates = deals_df['Close Date'].dropna()
                if len(valid_dates) > 0:
                    min_date = valid_dates.min()
                    max_date = valid_dates.max()
                    #st.sidebar.info(f"📅 Date range in data: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")
                    
                    # Count deals in each quarter
                    q4_2024_count = len(deals_df[(deals_df['Close Date'] >= '2024-10-01') & (deals_df['Close Date'] <= '2024-12-31')])
                    q1_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-01-01') & (deals_df['Close Date'] <= '2025-03-31')])
                    q2_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-04-01') & (deals_df['Close Date'] <= '2025-06-30')])
                    q3_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-07-01') & (deals_df['Close Date'] <= '2025-09-30')])
                    q4_2025_count = len(deals_df[(deals_df['Close Date'] >= '2025-10-01') & (deals_df['Close Date'] <= '2025-12-31')])
                    
                    #st.sidebar.info(f"Q4 2024: {q4_2024_count} | Q1 2025: {q1_2025_count} | Q2 2025: {q2_2025_count} | Q3 2025: {q3_2025_count} | Q4 2025: {q4_2025_count}")
                else:
                    pass  # Debug info removed
                    #st.sidebar.error("❌ No valid dates found in Close Date column!")
            else:
                pass  # Debug info removed
                #st.sidebar.error("❌ No Close Date column found!")
            
            # Show data before filtering
            total_deals_before = len(deals_df)
            total_amount_before = deals_df['Amount'].sum() if 'Amount' in deals_df.columns else 0
            #st.sidebar.info(f"📊 Before filtering: {total_deals_before} deals, ${total_amount_before:,.0f}")
            
            # Show unique values in Status column
            if 'Status' in deals_df.columns:
                unique_statuses = deals_df['Status'].unique()
                #st.sidebar.info(f"🏷️ Unique Status values: {', '.join([str(s) for s in unique_statuses[:10]])}")
            else:
                pass  # Debug info removed
                #st.sidebar.error("❌ No Status column found! Check 'Close Status' mapping")
            
            # FILTER: Only Q4 2025 deals (Oct 1 - Dec 31, 2025)
            # Use < Jan 1, 2026 to include all of Dec 31 regardless of timestamp
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2026-01-01')
            
            if 'Close Date' in deals_df.columns:
                before_count = len(deals_df)
                before_amount = deals_df['Amount'].sum()
                
                deals_df = deals_df[
                    (deals_df['Close Date'] >= q4_start) & 
                    (deals_df['Close Date'] < q4_end)
                ]
                after_count = len(deals_df)
                after_amount = deals_df['Amount'].sum()
                
                st.sidebar.markdown("### 📊 HubSpot Data Loaded")
                st.sidebar.caption(f"Total deals before Q4 filter: {before_count} (${before_amount:,.0f})")
                st.sidebar.caption(f"Q4 2025 deals: {after_count} (${after_amount:,.0f})")
                st.sidebar.caption(f"Filtered out: {before_count - after_count} deals")
                
                # Show breakdown by rep for Expect/Commit
                if 'Deal Owner' in deals_df.columns and 'Status' in deals_df.columns:
                    expect_commit = deals_df[deals_df['Status'].isin(['Expect', 'Commit'])]
                    st.sidebar.markdown("**Expect/Commit by Rep:**")
                    for rep in ['Brad Sherman', 'Jake Lynch', 'Dave Borkowski', 'Lance Mitton']:
                        rep_deals = expect_commit[expect_commit['Deal Owner'] == rep]
                        if not rep_deals.empty:
                            st.sidebar.caption(f"{rep}: {len(rep_deals)} deals, ${rep_deals['Amount'].sum():,.0f}")

            else:
                pass  # Debug info removed
                #st.sidebar.error("❌ Cannot apply date filter - no Close Date column")
            
            # FILTER OUT unwanted deal stages
            excluded_stages = [
                '', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 
                'closed lost', 'closed won', 'sales order created in NS', 
                'NCR', 'Shipped'
            ]
            
            # Convert Deal Stage to string and handle NaN
            if 'Deal Stage' in deals_df.columns:
                deals_df['Deal Stage'] = deals_df['Deal Stage'].fillna('')
                deals_df['Deal Stage'] = deals_df['Deal Stage'].astype(str).str.strip()
                
                # Show unique stages before filtering
                unique_stages = deals_df['Deal Stage'].unique()
                #st.sidebar.info(f"🎯 Unique Deal Stages: {', '.join([str(s) for s in unique_stages[:10]])}")
                
                # Filter out excluded stages
                deals_df = deals_df[~deals_df['Deal Stage'].str.lower().isin([s.lower() if s else '' for s in excluded_stages])]
                
                #st.sidebar.success(f"✅ After stage filter: {len(deals_df)} deals, ${deals_df['Amount'].sum():,.0f}")
            else:
                pass  # Debug info removed
                #st.sidebar.warning("⚠️ No Deal Stage column found")
            
            # Apply Q4 fulfillment logic
            deals_df = apply_q4_fulfillment_logic(deals_df)
    else:
        pass  # Debug info removed
        #st.sidebar.error(f"❌ HubSpot data has insufficient columns: {len(deals_df.columns) if not deals_df.empty else 0}")
    
    if not dashboard_df.empty:
        # Ensure we have the right column names
        if len(dashboard_df.columns) >= 3:
            dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
            
            # Remove any empty rows
            dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
            
            # Clean and convert numeric columns
            def clean_numeric(value):
                if pd.isna(value) or str(value).strip() == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            dashboard_df['Quota'] = dashboard_df['Quota'].apply(clean_numeric)
            dashboard_df['NetSuite Orders'] = dashboard_df['NetSuite Orders'].apply(clean_numeric)
    
    # Process invoice data
    if not invoices_df.empty:
        if len(invoices_df.columns) >= 15:
            # Map additional columns for Shopify identification
            rename_dict = {
                invoices_df.columns[0]: 'Invoice Number',
                invoices_df.columns[1]: 'Status',
                invoices_df.columns[2]: 'Date',
                invoices_df.columns[6]: 'Customer',
                invoices_df.columns[10]: 'Amount',
                invoices_df.columns[14]: 'Sales Rep'
            }
            
            # NEW: Map Corrected Customer Name (Column T - index 19) and Rep Master (Column U - index 20)
            if len(invoices_df.columns) > 19:
                rename_dict[invoices_df.columns[19]] = 'Corrected Customer Name'  # Column T
            if len(invoices_df.columns) > 20:
                rename_dict[invoices_df.columns[20]] = 'Rep Master'  # Column U
            
            # Try to find HubSpot Pipeline and CSM columns
            for idx, col in enumerate(invoices_df.columns):
                col_str = str(col).lower()
                if 'hubspot' in col_str and 'pipeline' in col_str:
                    rename_dict[col] = 'HubSpot_Pipeline'
                elif col_str == 'csm' or 'csm' in col_str:
                    rename_dict[col] = 'CSM'
            
            invoices_df = invoices_df.rename(columns=rename_dict)
            
            # CRITICAL: Replace Sales Rep with Rep Master and Customer with Corrected Customer Name
            # This fixes the Shopify eCommerce invoices that weren't being applied to reps correctly
            if 'Rep Master' in invoices_df.columns:
                # Rep Master is the ONLY source of truth - completely replace Sales Rep
                invoices_df['Rep Master'] = invoices_df['Rep Master'].astype(str).str.strip()
                
                # Define invalid values that should be filtered out
                invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
                
                # FILTER OUT rows where Rep Master is invalid (including #N/A)
                # These rows won't count toward any revenue
                invoices_df = invoices_df[~invoices_df['Rep Master'].isin(invalid_values)]
                
                # Now replace Sales Rep with Rep Master for all remaining rows
                invoices_df['Sales Rep'] = invoices_df['Rep Master']
                # Drop the Rep Master column since we've copied it to Sales Rep
                invoices_df = invoices_df.drop(columns=['Rep Master'])
            else:
                st.sidebar.warning("⚠️ Rep Master column not found in invoices!")
            
            if 'Corrected Customer Name' in invoices_df.columns:
                # Corrected Customer Name takes priority - replace Customer with corrected values
                invoices_df['Corrected Customer Name'] = invoices_df['Corrected Customer Name'].astype(str).str.strip()
                invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
                mask = invoices_df['Corrected Customer Name'].isin(invalid_values)
                invoices_df.loc[~mask, 'Customer'] = invoices_df.loc[~mask, 'Corrected Customer Name']
                # Drop the Corrected Customer Name column since we've copied it to Customer
                invoices_df = invoices_df.drop(columns=['Corrected Customer Name'])
            
            def clean_numeric(value):
                if pd.isna(value) or str(value).strip() == '':
                    return 0
                cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
                try:
                    return float(cleaned)
                except:
                    return 0
            
            invoices_df['Amount'] = invoices_df['Amount'].apply(clean_numeric)
            invoices_df['Date'] = pd.to_datetime(invoices_df['Date'], errors='coerce')
            
            # Filter to Q4 2025 only (10/1/2025 - 12/31/2025)
            # This should match exactly what your boss filters in the sheet
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Apply date filter
            invoices_df = invoices_df[
                (invoices_df['Date'] >= q4_start) & 
                (invoices_df['Date'] <= q4_end)
            ]
            
            # Clean up Sales Rep field
            invoices_df['Sales Rep'] = invoices_df['Sales Rep'].astype(str).str.strip()
            
            # Filter out invalid Sales Reps BEFORE groupby
            # NOTE: We DO NOT filter Amount > 0 because credit memos (negative amounts) should reduce totals
            invoices_df = invoices_df[
                (invoices_df['Sales Rep'].notna()) & 
                (invoices_df['Sales Rep'] != '') &
                (invoices_df['Sales Rep'].str.lower() != 'nan') &
                (invoices_df['Sales Rep'].str.lower() != 'house')
            ]
            
            # CRITICAL: Remove duplicate invoices if they exist (keep first occurrence)
            if 'Invoice Number' in invoices_df.columns:
                before_dedupe = len(invoices_df)
                invoices_df = invoices_df.drop_duplicates(subset=['Invoice Number'], keep='first')
                after_dedupe = len(invoices_df)
                if before_dedupe != after_dedupe:
                    st.sidebar.warning(f"⚠️ Removed {before_dedupe - after_dedupe} duplicate invoices!")
            
            # Calculate invoice totals by rep
            invoice_totals = invoices_df.groupby('Sales Rep')['Amount'].sum().reset_index()
            invoice_totals.columns = ['Rep Name', 'Invoice Total']
            
            dashboard_df['Rep Name'] = dashboard_df['Rep Name'].str.strip()
            
            dashboard_df = dashboard_df.merge(invoice_totals, on='Rep Name', how='left')
            dashboard_df['Invoice Total'] = dashboard_df['Invoice Total'].fillna(0)
            
            dashboard_df['NetSuite Orders'] = dashboard_df['Invoice Total']
            dashboard_df = dashboard_df.drop('Invoice Total', axis=1)
            
            # Add Shopify ECommerce to dashboard if it has invoices but isn't in dashboard yet
            if 'Shopify ECommerce' in invoice_totals['Rep Name'].values:
                if 'Shopify ECommerce' not in dashboard_df['Rep Name'].values:
                    shopify_total = invoice_totals[invoice_totals['Rep Name'] == 'Shopify ECommerce']['Invoice Total'].iloc[0]
                    new_shopify_row = pd.DataFrame([{
                        'Rep Name': 'Shopify ECommerce',
                        'Quota': 0,
                        'NetSuite Orders': shopify_total
                    }])
                    dashboard_df = pd.concat([dashboard_df, new_shopify_row], ignore_index=True)
    
    # Process sales orders data with NEW LOGIC
    if not sales_orders_df.empty:
        # Map column positions
        col_names = sales_orders_df.columns.tolist()
        
        rename_dict = {}
        
        # NEW: Map Internal Id column (Column A) - CRITICAL for NetSuite links
        if len(col_names) > 0:
            col_a_lower = str(col_names[0]).lower()
            if 'internal' in col_a_lower and 'id' in col_a_lower:
                rename_dict[col_names[0]] = 'Internal ID'
        
        # Find standard columns - only map FIRST occurrence
        for idx, col in enumerate(col_names):
            col_lower = str(col).lower()
            if 'status' in col_lower and 'Status' not in rename_dict.values():
                rename_dict[col] = 'Status'
            elif ('amount' in col_lower or 'total' in col_lower) and 'Amount' not in rename_dict.values():
                rename_dict[col] = 'Amount'
            elif ('sales rep' in col_lower or 'salesrep' in col_lower) and 'Sales Rep' not in rename_dict.values():
                rename_dict[col] = 'Sales Rep'
            elif 'customer' in col_lower and 'customer promise' not in col_lower and 'Customer' not in rename_dict.values():
                rename_dict[col] = 'Customer'
            elif ('doc' in col_lower or 'document' in col_lower) and 'Document Number' not in rename_dict.values():
                rename_dict[col] = 'Document Number'
        
        # Map specific columns by position (0-indexed) - be more careful
        if len(col_names) > 8 and 'Order Start Date' not in rename_dict.values():
            rename_dict[col_names[8]] = 'Order Start Date'  # Column I
        if len(col_names) > 11 and 'Customer Promise Date' not in rename_dict.values():
            rename_dict[col_names[11]] = 'Customer Promise Date'  # Column L
        if len(col_names) > 12 and 'Projected Date' not in rename_dict.values():
            rename_dict[col_names[12]] = 'Projected Date'  # Column M
        
        # NEW COLUMN POSITIONS after adding Calyx | External Order in column AC
        if len(col_names) > 28:
            rename_dict[col_names[28]] = 'Calyx External Order'  # Column AC - NEW!
        if len(col_names) > 29 and 'Pending Approval Date' not in rename_dict.values():
            rename_dict[col_names[29]] = 'Pending Approval Date'  # Column AD (was AB)
        if len(col_names) > 30:
            rename_dict[col_names[30]] = 'Corrected Customer Name'  # Column AE (was AC)
        if len(col_names) > 31:
            rename_dict[col_names[31]] = 'Rep Master'  # Column AF (was AD)
        
        # NEW: Map PI || CSM column (Column G based on screenshot)
        for idx, col in enumerate(col_names):
            col_str = str(col).lower()
            if ('pi' in col_str and 'csm' in col_str) or col_str == 'pi || csm':
                rename_dict[col] = 'PI_CSM'
                break
        
        sales_orders_df = sales_orders_df.rename(columns=rename_dict)
        
        # CRITICAL: Replace Sales Rep with Rep Master and Customer with Corrected Customer Name
        # This fixes the Shopify eCommerce orders that weren't being applied to reps correctly
        if 'Rep Master' in sales_orders_df.columns:
            # Rep Master is the ONLY source of truth - completely replace Sales Rep
            sales_orders_df['Rep Master'] = sales_orders_df['Rep Master'].astype(str).str.strip()
            
            # Define invalid values that should be filtered out
            invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
            
            # FILTER OUT rows where Rep Master is invalid (including #N/A)
            # These rows won't count toward any revenue
            sales_orders_df = sales_orders_df[~sales_orders_df['Rep Master'].isin(invalid_values)]
            
            # Now replace Sales Rep with Rep Master for all remaining rows
            sales_orders_df['Sales Rep'] = sales_orders_df['Rep Master']
            # Drop the Rep Master column since we've copied it to Sales Rep
            sales_orders_df = sales_orders_df.drop(columns=['Rep Master'])
        
        if 'Corrected Customer Name' in sales_orders_df.columns:
            # Corrected Customer Name takes priority - replace Customer with corrected values
            sales_orders_df['Customer'] = sales_orders_df['Corrected Customer Name']
            # Drop the Corrected Customer Name column since we've copied it to Customer
            sales_orders_df = sales_orders_df.drop(columns=['Corrected Customer Name'])
        
        # CRITICAL: Remove any duplicate columns that may have been created
        if sales_orders_df.columns.duplicated().any():
            pass  # Debug info removed
            #st.sidebar.warning(f"⚠️ Removed duplicate columns in Sales Orders: {sales_orders_df.columns[sales_orders_df.columns.duplicated()].tolist()}")
            sales_orders_df = sales_orders_df.loc[:, ~sales_orders_df.columns.duplicated()]
        
        # Clean numeric values
        def clean_numeric_so(value):
            value_str = str(value).strip()
            if value_str == '' or value_str == 'nan' or value_str == 'None':
                return 0
            cleaned = value_str.replace(',', '').replace('$', '').replace(' ', '')
            try:
                return float(cleaned)
            except:
                return 0
        
        if 'Amount' in sales_orders_df.columns:
            sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric_so)
        
        if 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df['Sales Rep'] = sales_orders_df['Sales Rep'].astype(str).str.strip()
        
        if 'Status' in sales_orders_df.columns:
            sales_orders_df['Status'] = sales_orders_df['Status'].astype(str).str.strip()
        
        # Convert date columns - handle 2-digit years correctly (26 = 2026, not 1926)
        date_columns = ['Order Start Date', 'Customer Promise Date', 'Projected Date', 'Pending Approval Date']
        for col in date_columns:
            if col in sales_orders_df.columns:
                # First try standard parsing
                sales_orders_df[col] = pd.to_datetime(sales_orders_df[col], errors='coerce')
                
                # Fix any dates that got parsed as 1900s (2-digit year issue)
                # If year < 2000, add 100 years (e.g., 1926 -> 2026)
                if sales_orders_df[col].notna().any():
                    mask = (sales_orders_df[col].dt.year < 2000) & (sales_orders_df[col].notna())
                    if mask.any():
                        sales_orders_df.loc[mask, col] = sales_orders_df.loc[mask, col] + pd.DateOffset(years=100)
        
        # Filter to include Pending Approval, Pending Fulfillment, AND Pending Billing/Partially Fulfilled
        if 'Status' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                sales_orders_df['Status'].isin(['Pending Approval', 'Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])
            ]
        
        # Calculate age for Old Pending Approval
        if 'Order Start Date' in sales_orders_df.columns:
            today = pd.Timestamp.now()
            
            def business_days_between(start_date, end_date):
                if pd.isna(start_date):
                    return 0
                days = pd.bdate_range(start=start_date, end=end_date).size - 1
                return max(0, days)
            
            sales_orders_df['Age_Business_Days'] = sales_orders_df['Order Start Date'].apply(
                lambda x: business_days_between(x, today)
            )
        else:
            sales_orders_df['Age_Business_Days'] = 0
        
        # Remove rows without amount or sales rep
        if 'Amount' in sales_orders_df.columns and 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[
                (sales_orders_df['Amount'] > 0) & 
                (sales_orders_df['Sales Rep'].notna()) & 
                (sales_orders_df['Sales Rep'] != '') &
                (sales_orders_df['Sales Rep'] != 'nan') &
                (~sales_orders_df['Sales Rep'].str.lower().isin(['house']))
            ]
    else:
        st.warning("Could not find required columns in NS Sales Orders")
        sales_orders_df = pd.DataFrame()
    
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
                        if 'Quota' in current_rep.columns:
                            current_val = pd.to_numeric(current_rep['Quota'].iloc[0], errors='coerce')
                            previous_val = pd.to_numeric(previous_rep['Quota'].iloc[0], errors='coerce')
                            if not pd.isna(current_val) and not pd.isna(previous_val):
                                if current_val != previous_val:
                                    rep_change['goal_change'] = current_val - previous_val
                        
                        if 'NetSuite Orders' in current_rep.columns:
                            current_val = pd.to_numeric(current_rep['NetSuite Orders'].iloc[0], errors='coerce')
                            previous_val = pd.to_numeric(previous_rep['NetSuite Orders'].iloc[0], errors='coerce')
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
        
        # Calculate all current metrics
        current_metrics = calculate_team_metrics(deals_df, dashboard_df)
        previous_metrics = calculate_team_metrics(previous['deals'], previous['dashboard'])
        
        # Helper function to calculate sales order metrics
        def calculate_so_metrics(so_df):
            metrics = {
                'pending_fulfillment': 0,
                'pending_fulfillment_no_date': 0,
                'pending_approval': 0,
                'pending_approval_no_date': 0,
                'pending_approval_old': 0
            }
            
            if so_df.empty:
                return metrics
            
            so_df = so_df.copy()
            so_df['Amount_Numeric'] = pd.to_numeric(so_df.get('Amount', 0), errors='coerce')
            
            # Q4 2025 date range for filtering
            q4_start = pd.Timestamp('2025-10-01')
            q4_end = pd.Timestamp('2025-12-31')
            
            # Parse dates
            if 'Estimated Ship Date' in so_df.columns:
                so_df['Ship_Date_Parsed'] = pd.to_datetime(so_df['Estimated Ship Date'], errors='coerce')
            else:
                so_df['Ship_Date_Parsed'] = pd.NaT
            
            # Parse Pending Approval Date for PA filtering
            if 'Pending Approval Date' in so_df.columns:
                so_df['PA_Date_Parsed'] = pd.to_datetime(so_df['Pending Approval Date'], errors='coerce')
                # Fix any dates that got parsed as 1900s (2-digit year issue: 26 -> 1926 instead of 2026)
                if so_df['PA_Date_Parsed'].notna().any():
                    mask_1900s = (so_df['PA_Date_Parsed'].dt.year < 2000) & (so_df['PA_Date_Parsed'].notna())
                    if mask_1900s.any():
                        so_df.loc[mask_1900s, 'PA_Date_Parsed'] = so_df.loc[mask_1900s, 'PA_Date_Parsed'] + pd.DateOffset(years=100)
            else:
                so_df['PA_Date_Parsed'] = pd.NaT
            
            # Pending Fulfillment
            pf_df = so_df[so_df.get('Status', '') == 'Pending Fulfillment']
            metrics['pending_fulfillment'] = pf_df[pf_df['Ship_Date_Parsed'].notna()]['Amount_Numeric'].sum()
            metrics['pending_fulfillment_no_date'] = pf_df[pf_df['Ship_Date_Parsed'].isna()]['Amount_Numeric'].sum()
            
            # Pending Approval - filtered by Pending Approval Date within Q4 2025
            pa_df = so_df[so_df.get('Status', '') == 'Pending Approval'].copy()
            
            # PA with date: must have a valid date WITHIN Q4 2025
            pa_with_q4_date = pa_df[
                (pa_df['PA_Date_Parsed'].notna()) &
                (pa_df['PA_Date_Parsed'] >= q4_start) &
                (pa_df['PA_Date_Parsed'] <= q4_end)
            ]
            metrics['pending_approval'] = pa_with_q4_date['Amount_Numeric'].sum()
            
            # PA no date: missing or invalid Pending Approval Date
            pa_no_date = pa_df[pa_df['PA_Date_Parsed'].isna()]
            metrics['pending_approval_no_date'] = pa_no_date['Amount_Numeric'].sum()
            
            # Pending Approval > 2 weeks old
            if 'Transaction Date' in so_df.columns:
                so_df['Transaction_Date_Parsed'] = pd.to_datetime(so_df['Transaction Date'], errors='coerce')
                two_weeks_ago = datetime.now() - timedelta(days=14)
                old_pa = pa_df[pa_df['Transaction_Date_Parsed'] < two_weeks_ago]
                metrics['pending_approval_old'] = old_pa['Amount_Numeric'].sum()
            
            return metrics
        
        current_so_metrics = calculate_so_metrics(sales_orders_df)
        previous_so_metrics = calculate_so_metrics(previous['sales_orders'])
        
        # Team-level changes - organized by data category
        st.markdown("#### 👥 Team Overview")
        
        # Row 1: Invoiced & Shipped
        st.markdown("**💰 Invoiced & Shipped**")
        inv_col1, inv_col2, inv_col3, inv_col4 = st.columns(4)
        
        with inv_col1:
            current_invoices = len(invoices_df) if not invoices_df.empty else 0
            previous_invoices = len(previous['invoices']) if not previous['invoices'].empty else 0
            delta_invoices = current_invoices - previous_invoices
            st.metric("Total Invoices", current_invoices, delta=delta_invoices)
        
        with inv_col2:
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
        
        with inv_col3:
            # NetSuite Orders from dashboard
            current_ns_orders = current_metrics.get('orders', 0)
            previous_ns_orders = previous_metrics.get('orders', 0)
            delta_ns = current_ns_orders - previous_ns_orders
            st.metric("NS Orders (Dashboard)", f"${current_ns_orders:,.0f}", delta=f"${delta_ns:,.0f}")
        
        with inv_col4:
            # Average invoice size
            if current_invoices > 0:
                current_avg = current_inv_total / current_invoices
            else:
                current_avg = 0
            
            if previous_invoices > 0:
                previous_avg = previous_inv_total / previous_invoices
            else:
                previous_avg = 0
            
            delta_avg = current_avg - previous_avg
            st.metric("Avg Invoice", f"${current_avg:,.0f}", delta=f"${delta_avg:,.0f}")
        
        # Row 2: Sales Orders
        st.markdown("**📦 Sales Orders**")
        so_col1, so_col2, so_col3, so_col4 = st.columns(4)
        
        with so_col1:
            current_orders = len(sales_orders_df) if not sales_orders_df.empty else 0
            previous_orders = len(previous['sales_orders']) if not previous['sales_orders'].empty else 0
            delta_orders = current_orders - previous_orders
            st.metric("Total Sales Orders", current_orders, delta=delta_orders)
        
        with so_col2:
            delta_pf = current_so_metrics['pending_fulfillment'] - previous_so_metrics['pending_fulfillment']
            st.metric("Pending Fulfillment (with date)", 
                     f"${current_so_metrics['pending_fulfillment']:,.0f}", 
                     delta=f"${delta_pf:,.0f}")
        
        with so_col3:
            delta_pa = current_so_metrics['pending_approval'] - previous_so_metrics['pending_approval']
            st.metric("Pending Approval (with date)", 
                     f"${current_so_metrics['pending_approval']:,.0f}", 
                     delta=f"${delta_pa:,.0f}")
        
        with so_col4:
            delta_pf_nd = current_so_metrics['pending_fulfillment_no_date'] - previous_so_metrics['pending_fulfillment_no_date']
            st.metric("Pending Fulfillment (no date)", 
                     f"${current_so_metrics['pending_fulfillment_no_date']:,.0f}", 
                     delta=f"${delta_pf_nd:,.0f}")
        
        # Row 3: Sales Orders Continued
        so2_col1, so2_col2, so2_col3, so2_col4 = st.columns(4)
        
        with so2_col1:
            delta_pa_nd = current_so_metrics['pending_approval_no_date'] - previous_so_metrics['pending_approval_no_date']
            st.metric("Pending Approval (no date)", 
                     f"${current_so_metrics['pending_approval_no_date']:,.0f}", 
                     delta=f"${delta_pa_nd:,.0f}")
        
        with so2_col2:
            delta_pa_old = current_so_metrics['pending_approval_old'] - previous_so_metrics['pending_approval_old']
            st.metric("Pending Approval (>2 weeks)", 
                     f"${current_so_metrics['pending_approval_old']:,.0f}", 
                     delta=f"${delta_pa_old:,.0f}")
        
        with so2_col3:
            # Total SO Amount
            if not sales_orders_df.empty and 'Amount' in sales_orders_df.columns:
                current_so_total = pd.to_numeric(sales_orders_df['Amount'], errors='coerce').sum()
            else:
                current_so_total = 0
            
            if not previous['sales_orders'].empty and 'Amount' in previous['sales_orders'].columns:
                previous_so_total = pd.to_numeric(previous['sales_orders']['Amount'], errors='coerce').sum()
            else:
                previous_so_total = 0
            
            delta_so_total = current_so_total - previous_so_total
            st.metric("Total SO Amount", f"${current_so_total:,.0f}", delta=f"${delta_so_total:,.0f}")
        
        # Row 4: HubSpot Deals
        st.markdown("**🎯 HubSpot Deals**")
        hs_col1, hs_col2, hs_col3, hs_col4 = st.columns(4)
        
        with hs_col1:
            current_deals = len(deals_df) if not deals_df.empty else 0
            previous_deals = len(previous['deals']) if not previous['deals'].empty else 0
            delta_deals = current_deals - previous_deals
            st.metric("Total Deals", current_deals, delta=delta_deals)
        
        with hs_col2:
            current_commit = current_metrics.get('expect_commit', 0)
            previous_commit = previous_metrics.get('expect_commit', 0)
            delta_commit = current_commit - previous_commit
            st.metric("HubSpot Commit", f"${current_commit:,.0f}", delta=f"${delta_commit:,.0f}")
        
        with hs_col3:
            # Calculate HubSpot Expect separately
            def get_expect_amount(df):
                if df.empty or 'Status' not in df.columns:
                    return 0
                df = df.copy()
                df['Amount_Numeric'] = pd.to_numeric(df.get('Amount', 0), errors='coerce')
                # Use Q1 2026 Spillover column (now fixed in spreadsheet)
                q4_deals = df[df.get('Q1 2026 Spillover') != 'Q1 2026']
                return q4_deals[q4_deals['Status'] == 'Expect']['Amount_Numeric'].sum()
            
            current_expect = get_expect_amount(deals_df)
            previous_expect = get_expect_amount(previous['deals'])
            delta_expect = current_expect - previous_expect
            st.metric("HubSpot Expect", f"${current_expect:,.0f}", delta=f"${delta_expect:,.0f}")
        
        with hs_col4:
            # Calculate HubSpot Best Case
            def get_best_case_amount(df):
                if df.empty or 'Status' not in df.columns:
                    return 0
                df = df.copy()
                df['Amount_Numeric'] = pd.to_numeric(df.get('Amount', 0), errors='coerce')
                # Use Q1 2026 Spillover column (now fixed in spreadsheet)
                q4_deals = df[df.get('Q1 2026 Spillover') != 'Q1 2026']
                return q4_deals[q4_deals['Status'] == 'Best Case']['Amount_Numeric'].sum()
            
            current_bc = get_best_case_amount(deals_df)
            previous_bc = get_best_case_amount(previous['deals'])
            delta_bc = current_bc - previous_bc
            st.metric("HubSpot Best Case", f"${current_bc:,.0f}", delta=f"${delta_bc:,.0f}")
        
        # Row 5: HubSpot Continued + Q1 Spillover
        hs2_col1, hs2_col2, hs2_col3, hs2_col4 = st.columns(4)
        
        with hs2_col1:
            # Calculate HubSpot Opportunity
            def get_opportunity_amount(df):
                if df.empty or 'Status' not in df.columns:
                    return 0
                df = df.copy()
                df['Amount_Numeric'] = pd.to_numeric(df.get('Amount', 0), errors='coerce')
                # Use Q1 2026 Spillover column (now fixed in spreadsheet)
                q4_deals = df[df.get('Q1 2026 Spillover') != 'Q1 2026']
                return q4_deals[q4_deals['Status'] == 'Opportunity']['Amount_Numeric'].sum()
            
            current_opp = get_opportunity_amount(deals_df)
            previous_opp = get_opportunity_amount(previous['deals'])
            delta_opp = current_opp - previous_opp
            st.metric("HubSpot Opportunity", f"${current_opp:,.0f}", delta=f"${delta_opp:,.0f}")
        
        with hs2_col2:
            current_q1 = current_metrics.get('q1_spillover_expect_commit', 0)
            previous_q1 = previous_metrics.get('q1_spillover_expect_commit', 0)
            delta_q1 = current_q1 - previous_q1
            st.metric("Q1 Spillover - Expect/Commit", f"${current_q1:,.0f}", delta=f"${delta_q1:,.0f}")
        
        # Rep-level changes
        st.markdown("#### 👤 Rep-Level Changes")
        
        if not dashboard_df.empty and not previous['dashboard'].empty:
            rep_comparison = []
            
            for rep in dashboard_df['Rep Name'].unique():
                current_rep = dashboard_df[dashboard_df['Rep Name'] == rep]
                previous_rep = previous['dashboard'][previous['dashboard']['Rep Name'] == rep]
                
                if not previous_rep.empty:
                    rep_data = {'Rep': rep}
                    
                    # NetSuite Orders change
                    if 'NetSuite Orders' in current_rep.columns:
                        current_val = pd.to_numeric(current_rep['NetSuite Orders'].iloc[0], errors='coerce')
                        previous_val = pd.to_numeric(previous_rep['NetSuite Orders'].iloc[0], errors='coerce')
                        if not pd.isna(current_val) and not pd.isna(previous_val):
                            rep_data['Current Actual'] = current_val
                            rep_data['Previous Actual'] = previous_val
                            rep_data['Δ Actual'] = current_val - previous_val
                    
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
                            'Δ Actual': '${:,.0f}'
                        }),
                        use_container_width=True
                    )
                else:
                    st.info("No significant changes in rep metrics")
            else:
                st.info("No rep-level data available for comparison")
        
    else:
        st.info("📸 No previous snapshot available. Changes will be tracked after the next refresh.")

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
    total_invoiced = 0
    if 'Amount' in filtered_invoices.columns:
        filtered_invoices['Amount_Numeric'] = pd.to_numeric(filtered_invoices['Amount'], errors='coerce')
        total_invoiced = filtered_invoices['Amount_Numeric'].sum()
    
    # Display summary metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Invoices", len(filtered_invoices))
    
    with col2:
        st.metric("Total Amount", f"${total_invoiced:,.0f}")
    
    with col3:
        if len(filtered_invoices) > 0 and total_invoiced > 0:
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
            
            # Format currency - only if we have both Amount and Amount_Numeric
            if 'Amount' in display_df.columns and 'Amount_Numeric' in filtered_invoices.columns:
                # Use the index to align properly
                display_df['Amount'] = filtered_invoices.loc[display_df.index, 'Amount_Numeric'].apply(
                    lambda x: f"${x:,.0f}" if not pd.isna(x) else ""
                )
            
            # Enhanced dataframe with column config (Gemini enhancement)
            column_config = {}
            
            if 'Amount' in display_df.columns and 'Amount_Numeric' in filtered_invoices.columns:
                max_amount = filtered_invoices['Amount_Numeric'].max()
                if max_amount > 0:
                    column_config['Amount'] = st.column_config.ProgressColumn(
                        "Invoice Value",
                        format="$%.0f",
                        min_value=0,
                        max_value=max_amount,
                        help="Size relative to largest invoice"
                    )
            
            if 'Date' in display_df.columns:
                column_config['Date'] = st.column_config.DateColumn(
                    "Invoice Date",
                    format="MMM DD, YYYY",
                    help="Date invoice was created"
                )
            
            if '🔗 NetSuite' in display_df.columns:
                column_config['🔗 NetSuite'] = st.column_config.LinkColumn(
                    "View",
                    display_text="↗️ Open",
                    help="Open in NetSuite"
                )
            
            st.dataframe(
                display_df, 
                use_container_width=True, 
                hide_index=True,
                column_config=column_config if column_config else None
            )
        else:
            st.dataframe(filtered_invoices, use_container_width=True, hide_index=True)

def build_your_own_forecast_section(metrics, quota, rep_name=None, deals_df=None, invoices_df=None, sales_orders_df=None):
    """
    Refined Interactive Forecast Builder (v6 - Robust Export Edition)
    - Captures 'Customize' selections for export
    - Includes detailed Summary + Line Item export
    - Displays SO#, Links, and Dates safely
    """
    st.markdown("### 🎯 Build Your Own Forecast")
    st.caption("Select components to include. Expand sections to see details.")
    
    # --- IN/OUT/MAYBE UPLOAD FEATURE ---
    st.markdown("---")
    with st.expander("📋 Upload Q4 Planning Status (IN/OUT/MAYBE)", expanded=False):
        st.markdown("""
        Upload a CSV or Excel file with SO#/Deal IDs and their Q4 status.
        
        **Expected format:**
        ```
        SO13501,In
        SO13502,Out
        32533096097,Maybe
        32533096098,In
        ```
        
        - Column A: ID (e.g., "SO13501" or "32533096097")
        - Column B: Status ("In", "Out", or "Maybe")
        
        **Status meanings:**
        - **In** = Automatically checked in forecast
        - **Maybe** = Automatically checked (can adjust manually)
        - **Out** = Unchecked by default
        """)
        
        uploaded_file = st.file_uploader(
            "Upload planning status file",
            type=['csv', 'xlsx', 'xls'],
            key=f"planning_upload_{rep_name}",
            help="CSV or Excel file with SO#/Deal IDs and their In/Out/Maybe status"
        )
        
        # Initialize session state for planning status
        planning_key = f'planning_status_{rep_name}'
        if planning_key not in st.session_state:
            st.session_state[planning_key] = {}
        
        # Track if we've already processed this file to prevent rerun loop
        processed_key = f'processed_file_{rep_name}'
        if processed_key not in st.session_state:
            st.session_state[processed_key] = None
        
        if uploaded_file is not None:
            # Check if this is a new file (different from last processed)
            file_id = f"{uploaded_file.name}_{uploaded_file.size}"
            
            if st.session_state[processed_key] != file_id:
                try:
                    # Read the file
                    if uploaded_file.name.endswith('.csv'):
                        # Try reading with header first
                        planning_df = pd.read_csv(uploaded_file)
                        # If no proper header, read without header
                        if planning_df.shape[1] == 2 and planning_df.columns[0] not in ['ID', 'SO', 'Deal']:
                            uploaded_file.seek(0)  # Reset file pointer
                            planning_df = pd.read_csv(uploaded_file, header=None, names=['ID', 'Status'])
                    else:
                        planning_df = pd.read_excel(uploaded_file)
                        # If no proper header, assume first two columns are ID and Status
                        if planning_df.shape[1] >= 2:
                            if planning_df.columns[0] not in ['ID', 'SO', 'Deal']:
                                planning_df.columns = ['ID', 'Status'] + list(planning_df.columns[2:])
                    
                    # Ensure we have the right columns
                    if len(planning_df.columns) >= 2:
                        # Use first two columns as ID and Status
                        if 'ID' not in planning_df.columns or 'Status' not in planning_df.columns:
                            planning_df.columns = ['ID', 'Status'] + list(planning_df.columns[2:]) if len(planning_df.columns) > 2 else ['ID', 'Status']
                        
                        # Store in session state as dict: {ID: Status}
                        # Normalize status to uppercase for consistent matching
                        st.session_state[planning_key] = dict(zip(
                            planning_df['ID'].astype(str).str.strip(),
                            planning_df['Status'].astype(str).str.strip().str.upper()
                        ))
                        
                        # Show summary
                        status_counts = planning_df['Status'].str.upper().value_counts()
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("✅ IN", status_counts.get('IN', 0))
                        with col2:
                            st.metric("⚠️ MAYBE", status_counts.get('MAYBE', 0))
                        with col3:
                            st.metric("❌ OUT", status_counts.get('OUT', 0))
                        
                        st.success(f"✅ Loaded {len(st.session_state[planning_key])} planning statuses")
                        
                        # Mark this file as processed
                        st.session_state[processed_key] = file_id
                        
                        # Note: Checkbox states will be set when rendering the checkboxes below
                        # We'll use the planning status to determine default values
                        
                        # Rerun to apply the planning status to checkboxes
                        st.rerun()
                    else:
                        st.error("❌ File must have at least 2 columns (ID and Status)")
                except Exception as e:
                    st.error(f"❌ Error reading file: {str(e)}")
            else:
                # File already processed, show summary
                if st.session_state[planning_key]:
                    from collections import Counter
                    status_counts = Counter(st.session_state[planning_key].values())
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("✅ IN", status_counts.get('IN', 0))
                    with col2:
                        st.metric("⚠️ MAYBE", status_counts.get('MAYBE', 0))
                    with col3:
                        st.metric("❌ OUT", status_counts.get('OUT', 0))
                    st.success(f"✅ Loaded {len(st.session_state[planning_key])} planning statuses")
        
        # Clear button
        if st.session_state[planning_key]:
            if st.button("🗑️ Clear Planning Status", key=f"clear_planning_{rep_name}"):
                # Clear planning status
                st.session_state[planning_key] = {}
                st.session_state[processed_key] = None
                
                # Clear all checkbox states for this rep
                keys_to_clear = [k for k in st.session_state.keys() if k.startswith(f"chk_") and k.endswith(f"_{rep_name}")]
                for key in keys_to_clear:
                    del st.session_state[key]
                
                st.rerun()
    
    st.markdown("---")
    
    # View mode toggle
    view_mode = st.radio(
        "View Mode",
        options=["📂 Category View (Separate Sections)", "📋 Consolidated View (Single Table)"],
        horizontal=True,
        key=f"view_mode_{rep_name}"
    )
    
    # Helper function to get planning status for an ID
    def get_planning_status(id_value):
        """Get planning status (IN/OUT/MAYBE) for a given SO# or Deal ID"""
        if not id_value or pd.isna(id_value):
            return None
        id_str = str(id_value).strip()
        return st.session_state[planning_key].get(id_str)
    
    # --- 1. PREPARE DATA LOCALLY ---
    
    # Helper to grab a column by Index (Safe Fallback)
    def get_col_by_index(df, index):
        if df is not None and len(df.columns) > index:
            return df.iloc[:, index]
        return pd.Series()

    # Prepare Sales Order Data
    if sales_orders_df is not None and not sales_orders_df.empty:
        # Filter for Rep
        if rep_name:
            if 'Sales Rep' in sales_orders_df.columns:
                so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
            else:
                so_data = sales_orders_df.copy() 
        else:
            so_data = sales_orders_df.copy()
            
        # --- GRAB RAW COLUMNS BY INDEX ---
        so_data['Display_SO_Num'] = get_col_by_index(so_data, 1)        # Col B: SO#
        so_data['Display_PF_Date'] = pd.to_datetime(get_col_by_index(so_data, 9), errors='coerce') # Col J: PF Date
        so_data['Display_Promise_Date'] = pd.to_datetime(get_col_by_index(so_data, 11), errors='coerce') # Col L
        so_data['Display_Projected_Date'] = pd.to_datetime(get_col_by_index(so_data, 12), errors='coerce') # Col M
        so_data['Display_Type'] = get_col_by_index(so_data, 17).fillna('Standard') # Col R: Order Type
        
        # Use renamed column if available, otherwise try by index
        if 'Pending Approval Date' in so_data.columns:
            so_data['Display_PA_Date'] = pd.to_datetime(so_data['Pending Approval Date'], errors='coerce')
        else:
            so_data['Display_PA_Date'] = pd.to_datetime(get_col_by_index(so_data, 29), errors='coerce') # Col AD: PA Date

        if 'Amount' in so_data.columns:
            so_data['Amount_Numeric'] = pd.to_numeric(so_data['Amount'], errors='coerce').fillna(0)
        else:
            so_data['Amount_Numeric'] = 0
    else:
        so_data = pd.DataFrame()

    # Prepare HubSpot Data
    if deals_df is not None and not deals_df.empty:
        if rep_name:
            hs_data = deals_df[deals_df['Deal Owner'] == rep_name].copy()
        else:
            hs_data = deals_df.copy()
            
        # Map Deal Type (Column N - Index 13)
        hs_data['Display_Type'] = get_col_by_index(hs_data, 13).fillna('Standard')
        
        # Get Pending Approval Date from Column P (index 15)
        if 'Pending Approval Date' in hs_data.columns:
            hs_data['Display_PA_Date'] = pd.to_datetime(hs_data['Pending Approval Date'], errors='coerce')
        else:
            # Fallback to column index 15 (Column P)
            hs_data['Display_PA_Date'] = pd.to_datetime(get_col_by_index(hs_data, 15), errors='coerce')

        if 'Amount' in hs_data.columns:
            hs_data['Amount_Numeric'] = pd.to_numeric(hs_data['Amount'], errors='coerce').fillna(0)
    else:
        hs_data = pd.DataFrame()

    # --- 2. CATEGORY DEFINITIONS ---
    
    invoiced_shipped = metrics.get('orders', 0)
    
    ns_categories = {
        'PF_Date_Ext':   {'label': 'Pending Fulfillment (Date) - External'},
        'PF_Date_Int':   {'label': 'Pending Fulfillment (Date) - Internal'},
        'PF_NoDate_Ext': {'label': 'PF (No Date) - External'},
        'PF_NoDate_Int': {'label': 'PF (No Date) - Internal'},
        'PA_Date':       {'label': 'Pending Approval (With Date)'},
        'PA_NoDate':     {'label': 'Pending Approval (No Date)'},
        'PA_Old':        {'label': 'Pending Approval (>2 Wks)'},
    }
    
    hs_categories = {
        'Expect':      {'label': 'HubSpot Expect'},
        'Commit':      {'label': 'HubSpot Commit'},
        'BestCase':    {'label': 'HubSpot Best Case'},
        'Opp':         {'label': 'HubSpot Opp'},
        'Q1_Expect':   {'label': 'Q1 Spillover (Expect)'},
        'Q1_Commit':   {'label': 'Q1 Spillover (Commit)'},
        'Q1_BestCase': {'label': 'Q1 Spillover (Best Case)'},
        'Q1_Opp':      {'label': 'Q1 Spillover (Opp)'},
    }

    # --- 3. CREATE DISPLAY DATAFRAMES ---
    
    # === USE CENTRALIZED CATEGORIZATION FUNCTION ===
    so_categories = categorize_sales_orders(sales_orders_df, rep_name)
    
    # Helper function to add display columns for UI
    def format_ns_view(df, date_col_name):
        if df.empty: 
            return df
        d = df.copy()
        
        # Add display columns
        if 'Internal ID' in d.columns:
            d['Link'] = d['Internal ID'].apply(lambda x: f"https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}" if pd.notna(x) else "")
        
        # Add SO# column (from Display_SO_Num)
        if 'Display_SO_Num' in d.columns:
            d['SO #'] = d['Display_SO_Num']
        
        # Add Order Type column (from Display_Type)
        if 'Display_Type' in d.columns:
            d['Type'] = d['Display_Type']
        
        # Add Classification Date based on category
        # date_col_name indicates which date field was used to classify this SO
        if date_col_name == 'Promise':
            # For PF with date: use Customer Promise Date OR Projected Date (whichever exists)
            d['Classification Date'] = '—'
            
            # Try Customer Promise Date first
            if 'Display_Promise_Date' in d.columns:
                promise_dates = pd.to_datetime(d['Display_Promise_Date'], errors='coerce')
                d.loc[promise_dates.notna(), 'Classification Date'] = promise_dates.dt.strftime('%Y-%m-%d')
            
            # Fill in with Projected Date where Promise Date is missing
            if 'Display_Projected_Date' in d.columns:
                projected_dates = pd.to_datetime(d['Display_Projected_Date'], errors='coerce')
                mask = (d['Classification Date'] == '—') & projected_dates.notna()
                if mask.any():
                    d.loc[mask, 'Classification Date'] = projected_dates.loc[mask].dt.strftime('%Y-%m-%d')
                    
        elif date_col_name == 'PA_Date':
            # For PA with date: use Pending Approval Date
            if 'Display_PA_Date' in d.columns:
                pa_dates = pd.to_datetime(d['Display_PA_Date'], errors='coerce')
                d['Classification Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('—')
            else:
                d['Classification Date'] = '—'
        else:
            # For PF/PA no date or other: show dash
            d['Classification Date'] = '—'
        
        return d.sort_values('Amount', ascending=False) if 'Amount' in d.columns else d
    
    # Map centralized categories to display dataframes
    ns_dfs = {
        'PF_Date_Ext': format_ns_view(so_categories['pf_date_ext'], 'Promise'),
        'PF_Date_Int': format_ns_view(so_categories['pf_date_int'], 'Promise'),
        'PF_NoDate_Ext': format_ns_view(so_categories['pf_nodate_ext'], 'PF_Date'),
        'PF_NoDate_Int': format_ns_view(so_categories['pf_nodate_int'], 'PF_Date'),
        'PA_Old': format_ns_view(so_categories['pa_old'], 'PA_Date'),
        'PA_Date': format_ns_view(so_categories['pa_date'], 'PA_Date'),
        'PA_NoDate': format_ns_view(so_categories['pa_nodate'], 'None')
    }

    hs_dfs = {}
    if not hs_data.empty:
        # Use the Q1 2026 Spillover column from Google Sheet (spreadsheet formula now handles PA date logic)
        # Q4 deals: NOT marked as Q1 spillover
        # Q1 deals: Explicitly marked as Q1 2026
        q4 = hs_data.get('Q1 2026 Spillover') != 'Q1 2026'
        q1 = hs_data.get('Q1 2026 Spillover') == 'Q1 2026'
        
        def format_hs_view(df):
            if df.empty: return df
            d = df.copy()
            
            # Add Deal ID column (from Record ID)
            if 'Record ID' in d.columns:
                d['Deal ID'] = d['Record ID']
            
            d['Type'] = d['Display_Type']
            d['Close'] = pd.to_datetime(d['Close Date'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('—')
            
            # Change to Pending Approval Date
            if 'Display_PA_Date' in d.columns:
                pa_dates = pd.to_datetime(d['Display_PA_Date'], errors='coerce')
                d['PA Date'] = pa_dates.dt.strftime('%Y-%m-%d').fillna('—')
            else:
                d['PA Date'] = '—'
            
            if 'Record ID' in d.columns:
                d['Link'] = d['Record ID'].apply(lambda x: f"https://app.hubspot.com/contacts/6712259/record/0-3/{x}/" if pd.notna(x) else "")
            return d.sort_values(['Type', 'Amount_Numeric'], ascending=[True, False])

        hs_dfs['Expect'] = format_hs_view(hs_data[q4 & (hs_data['Status'] == 'Expect')])
        hs_dfs['Commit'] = format_hs_view(hs_data[q4 & (hs_data['Status'] == 'Commit')])
        hs_dfs['BestCase'] = format_hs_view(hs_data[q4 & (hs_data['Status'] == 'Best Case')])
        hs_dfs['Opp'] = format_hs_view(hs_data[q4 & (hs_data['Status'] == 'Opportunity')])
        hs_dfs['Q1_Expect'] = format_hs_view(hs_data[q1 & (hs_data['Status'] == 'Expect')])
        hs_dfs['Q1_Commit'] = format_hs_view(hs_data[q1 & (hs_data['Status'] == 'Commit')])
        hs_dfs['Q1_BestCase'] = format_hs_view(hs_data[q1 & (hs_data['Status'] == 'Best Case')])
        hs_dfs['Q1_Opp'] = format_hs_view(hs_data[q1 & (hs_data['Status'] == 'Opportunity')])

    # --- 4. RENDER UI & CAPTURE SELECTIONS ---
    
    # We use this dict to store the ACTUAL dataframes to be exported
    export_buckets = {}
    
    # Check view mode
    if view_mode == "📂 Category View (Separate Sections)":
        # === ORIGINAL CATEGORY VIEW ===
        with st.container():
            col_ns, col_hs = st.columns(2)
            
            # === NETSUITE COLUMN ===
            with col_ns:
                st.markdown("#### 📦 NetSuite Orders")
                st.info(f"**Invoiced (Locked):** ${invoiced_shipped:,.0f}")
                
                for key, data in ns_categories.items():
                    # Get value for label
                    df = ns_dfs.get(key, pd.DataFrame())
                    val = df['Amount'].sum() if not df.empty and 'Amount' in df.columns else 0
                    
                    # Determine default checkbox value based on planning status
                    checkbox_key = f"chk_{key}_{rep_name}"
                    
                    # Only set default if we have planning status and this key hasn't been set yet
                    if st.session_state[planning_key] and checkbox_key not in st.session_state:
                        if not df.empty and 'SO #' in df.columns:
                            # Check planning status for items in this category
                            statuses = [get_planning_status(so_num) for so_num in df['SO #']]
                            in_count = statuses.count('IN')
                            maybe_count = statuses.count('MAYBE')
                            out_count = statuses.count('OUT')
                            
                            # Auto-check if majority are IN or MAYBE
                            if in_count + maybe_count > out_count:
                                st.session_state[checkbox_key] = True
                            else:
                                st.session_state[checkbox_key] = False
                    
                    # Always show PA_Date even if 0 to debug
                    if val > 0 or key == 'PA_Date':
                        is_checked = st.checkbox(
                            f"{data['label']}: ${val:,.0f}", 
                            key=checkbox_key
                        )
                        
                        if is_checked:
                            with st.expander(f"🔎 View Orders ({data['label']})", expanded=False):
                                if not df.empty:
                                    enable_edit = st.toggle("Customize", key=f"tgl_{key}_{rep_name}")
                                    
                                    # Display Columns
                                    display_cols = []
                                    if 'Link' in df.columns: display_cols.append('Link')
                                    if 'SO #' in df.columns: display_cols.append('SO #')
                                    if 'Type' in df.columns: display_cols.append('Type')
                                    if 'Customer' in df.columns: display_cols.append('Customer')
                                    if 'Classification Date' in df.columns: display_cols.append('Classification Date')
                                    if 'Amount' in df.columns: display_cols.append('Amount')
                                    
                                    if enable_edit and display_cols:
                                        df_edit = df.copy()
                                        
                                        # Add Status column based on planning status
                                        if 'SO #' in df_edit.columns:
                                            df_edit['Status'] = df_edit['SO #'].apply(
                                                lambda so: get_planning_status(so) if get_planning_status(so) else '—'
                                            )
                                        
                                        # Pre-fill Select column based on planning status
                                        # IMPORTANT: Only check items that have IN or MAYBE status
                                        if 'SO #' in df_edit.columns:
                                            def should_select(so_num):
                                                status = get_planning_status(so_num)
                                                # If no planning status exists, default to True (selected)
                                                if not st.session_state[planning_key] or len(st.session_state[planning_key]) == 0:
                                                    return True
                                                # If planning status exists, only select IN/MAYBE
                                                return status in ['IN', 'MAYBE']
                                            
                                            df_edit['Select'] = df_edit['SO #'].apply(should_select)
                                        else:
                                            df_edit['Select'] = True
                                        
                                        # Reorder columns: Select, Status, then display columns
                                        cols_ordered = ['Select', 'Status'] + display_cols
                                        # Only include columns that exist
                                        cols_ordered = [c for c in cols_ordered if c in df_edit.columns]
                                        
                                        edited = st.data_editor(
                                            df_edit[cols_ordered],
                                            column_config={
                                                "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                                "Status": st.column_config.SelectboxColumn(
                                                "Q4 Status", 
                                                width="small",
                                                options=['IN', 'MAYBE', 'OUT', '—'],
                                                required=False
                                            ),
                                            "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                            "SO #": st.column_config.TextColumn("SO #", width="small"),
                                            "Type": st.column_config.TextColumn("Type", width="small"),
                                            "Classification Date": st.column_config.TextColumn("Class. Date", width="small"),
                                            "Amount": st.column_config.NumberColumn("Amount", format="$%d")
                                        },
                                        disabled=['Link', 'SO #', 'Type', 'Customer', 'Classification Date', 'Amount'],
                                        hide_index=True,
                                        key=f"edit_{key}_{rep_name}",
                                        num_rows="fixed"
                                    )
                                
                                        # Update planning status from edited data
                                        if 'SO #' in edited.columns and 'Status' in edited.columns:
                                            for idx, row in edited.iterrows():
                                                so_num = str(row['SO #']).strip()
                                                status = str(row['Status']).strip().upper()
                                                if status != '—' and status in ['IN', 'MAYBE', 'OUT']:
                                                    st.session_state[planning_key][so_num] = status
                                                elif status == '—' and so_num in st.session_state[planning_key]:
                                                    # Remove from planning status if set to dash
                                                    del st.session_state[planning_key][so_num]
                                
                                        # Auto-update Select checkboxes based on Status changes
                                        if 'Status' in edited.columns and 'Select' in edited.columns:
                                            for idx in edited.index:
                                                status = str(edited.at[idx, 'Status']).strip().upper()
                                                # Auto-check if Status is IN or MAYBE
                                                if status in ['IN', 'MAYBE']:
                                                    edited.at[idx, 'Select'] = True
                                                # Auto-uncheck if Status is OUT or dash
                                                elif status in ['OUT', '—']:
                                                    edited.at[idx, 'Select'] = False
                                
                                        # Capture filtered rows for export
                                        selected_rows = edited[edited['Select']].copy()
                                        
                                        current_total = selected_rows['Amount'].sum() if 'Amount' in selected_rows.columns else 0
                                        st.caption(f"Selected: ${current_total:,.0f}")
                                        
                                        # Helpful note about auto-check
                                        if 'Status' in edited.columns:
                                            st.caption("💡 Tip: Changing Status to IN/MAYBE auto-selects the item, OUT/— auto-deselects")
                                        
                                        # ALWAYS set export_buckets in customize mode
                                        export_buckets[key] = selected_rows
                                else:
                                    # Read-only view
                                    if display_cols:
                                        df_readonly = df.copy()
                                    
                                        # Add Status column for read-only view too
                                        if 'SO #' in df_readonly.columns:
                                            df_readonly['Status'] = df_readonly['SO #'].apply(
                                                lambda so: get_planning_status(so) if get_planning_status(so) else '—'
                                            )
                                            display_readonly = ['Status'] + display_cols
                                        else:
                                            display_readonly = display_cols
                                    
                                        st.dataframe(
                                            df_readonly[display_readonly],
                                            column_config={
                                                "Status": st.column_config.TextColumn("Q4 Status", width="small"),
                                                "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                                "SO #": st.column_config.TextColumn("SO #", width="small"),
                                                "Type": st.column_config.TextColumn("Type", width="small"),
                                                "Classification Date": st.column_config.TextColumn("Class. Date", width="small"),
                                                "Amount": st.column_config.NumberColumn("Amount", format="$%d")
                                            },
                                            hide_index=True,
                                            use_container_width=True
                                        )
                                    # Capture all rows for export (with Status column)
                                    # BUT only include items with IN/MAYBE status IF planning status exists
                                    if 'SO #' in df.columns:
                                        df_export = df.copy()
                                        df_export['Status'] = df_export['SO #'].apply(
                                            lambda so: get_planning_status(so) if get_planning_status(so) else '—'
                                        )
                                        # Only filter if planning status has actual entries
                                        if st.session_state[planning_key] and len(st.session_state[planning_key]) > 0:
                                            df_export['_should_include'] = df_export['SO #'].apply(
                                                lambda so: get_planning_status(so) in ['IN', 'MAYBE']
                                            )
                                            df_export = df_export[df_export['_should_include']].drop(columns=['_should_include'])
                                        export_buckets[key] = df_export
                                    else:
                                        export_buckets[key] = df
                            
                            # Fallback: if export_buckets wasn't set by expander code, set default
                            if key not in export_buckets:
                                df_default = df.copy()
                                if 'SO #' in df_default.columns:
                                    df_default['Status'] = df_default['SO #'].apply(
                                        lambda so: get_planning_status(so) if get_planning_status(so) else '—'
                                    )
                                    # Only filter if planning status exists
                                    if st.session_state[planning_key] and len(st.session_state[planning_key]) > 0:
                                        df_default['_should_include'] = df_default['SO #'].apply(
                                            lambda so: get_planning_status(so) in ['IN', 'MAYBE']
                                        )
                                        df_default = df_default[df_default['_should_include']].drop(columns=['_should_include'])
                                export_buckets[key] = df_default

        # === HUBSPOT COLUMN ===
        with col_hs:
            st.markdown("#### 🎯 HubSpot Pipeline")
            for key, data in hs_categories.items():
                df = hs_dfs.get(key, pd.DataFrame())
                val = df['Amount_Numeric'].sum() if not df.empty else 0
                
                # Determine default checkbox value based on planning status
                checkbox_key = f"chk_{key}_{rep_name}"
                
                # Only set default if we have planning status and this key hasn't been set yet
                if st.session_state[planning_key] and checkbox_key not in st.session_state:
                    if not df.empty and 'Deal ID' in df.columns:
                        # Check planning status for items in this category
                        statuses = [get_planning_status(deal_id) for deal_id in df['Deal ID']]
                        in_count = statuses.count('IN')
                        maybe_count = statuses.count('MAYBE')
                        out_count = statuses.count('OUT')
                        
                        # Auto-check if majority are IN or MAYBE
                        if in_count + maybe_count > out_count:
                            st.session_state[checkbox_key] = True
                        else:
                            st.session_state[checkbox_key] = False
                
                if val > 0:
                    is_checked = st.checkbox(
                        f"{data['label']}: ${val:,.0f}", 
                        key=checkbox_key
                    )
                    if is_checked:
                        with st.expander(f"🔎 View Deals ({data['label']})"):
                            if not df.empty:
                                enable_edit = st.toggle("Customize", key=f"tgl_{key}_{rep_name}")
                                cols = ['Link', 'Deal ID', 'Deal Name', 'Type', 'Close', 'PA Date', 'Amount_Numeric']
                                
                                if enable_edit:
                                    df_edit = df.copy()
                                    
                                    # Add Status column based on planning status
                                    if 'Deal ID' in df_edit.columns:
                                        df_edit['Status'] = df_edit['Deal ID'].apply(
                                            lambda deal_id: get_planning_status(deal_id) if get_planning_status(deal_id) else '—'
                                        )
                                    
                                    # Pre-fill Select column based on planning status
                                    # IMPORTANT: Only check items that have IN or MAYBE status
                                    if 'Deal ID' in df_edit.columns:
                                        def should_select(deal_id):
                                            status = get_planning_status(deal_id)
                                            # If no planning status exists, default to True (selected)
                                            if not st.session_state[planning_key] or len(st.session_state[planning_key]) == 0:
                                                return True
                                            # If planning status exists, only select IN/MAYBE
                                            return status in ['IN', 'MAYBE']
                                        
                                        df_edit['Select'] = df_edit['Deal ID'].apply(should_select)
                                    else:
                                        df_edit['Select'] = True
                                    
                                    # Reorder columns: Select, Status, then display columns
                                    cols_ordered = ['Select', 'Status'] + cols
                                    # Only include columns that exist
                                    cols_ordered = [c for c in cols_ordered if c in df_edit.columns]
                                    
                                    edited = st.data_editor(
                                        df_edit[cols_ordered],
                                        column_config={
                                            "Select": st.column_config.CheckboxColumn("✓", width="small"),
                                            "Status": st.column_config.SelectboxColumn(
                                                "Q4 Status", 
                                                width="small",
                                                options=['IN', 'MAYBE', 'OUT', '—'],
                                                required=False
                                            ),
                                            "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                            "Deal ID": st.column_config.TextColumn("Deal ID", width="small"),
                                            "Type": st.column_config.TextColumn("Type", width="small"),
                                            "Close": st.column_config.TextColumn("Close Date", width="small"),
                                            "PA Date": st.column_config.TextColumn("PA Date", width="small"),
                                            "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")
                                        },
                                        disabled=['Link', 'Deal ID', 'Deal Name', 'Type', 'Close', 'PA Date', 'Amount_Numeric'],
                                        hide_index=True,
                                        key=f"edit_{key}_{rep_name}",
                                        num_rows="fixed"
                                    )
                                    
                                    # Update planning status from edited data
                                    if 'Deal ID' in edited.columns and 'Status' in edited.columns:
                                        for idx, row in edited.iterrows():
                                            deal_id = str(row['Deal ID']).strip()
                                            status = str(row['Status']).strip().upper()
                                            if status != '—' and status in ['IN', 'MAYBE', 'OUT']:
                                                st.session_state[planning_key][deal_id] = status
                                            elif status == '—' and deal_id in st.session_state[planning_key]:
                                                # Remove from planning status if set to dash
                                                del st.session_state[planning_key][deal_id]
                                    
                                    # Auto-update Select checkboxes based on Status changes
                                    if 'Status' in edited.columns and 'Select' in edited.columns:
                                        for idx in edited.index:
                                            status = str(edited.at[idx, 'Status']).strip().upper()
                                            # Auto-check if Status is IN or MAYBE
                                            if status in ['IN', 'MAYBE']:
                                                edited.at[idx, 'Select'] = True
                                            # Auto-uncheck if Status is OUT or dash
                                            elif status in ['OUT', '—']:
                                                edited.at[idx, 'Select'] = False
                                    
                                    selected_rows = edited[edited['Select']].copy()
                                    
                                    current_total = selected_rows['Amount_Numeric'].sum()
                                    st.caption(f"Selected: ${current_total:,.0f}")
                                    
                                    # Helpful note about auto-check
                                    if 'Status' in edited.columns:
                                        st.caption("💡 Tip: Changing Status to IN/MAYBE auto-selects the item, OUT/— auto-deselects")
                                    
                                    # ALWAYS set export_buckets in customize mode
                                    export_buckets[key] = selected_rows
                                else:
                                    # Read-only view
                                    df_readonly = df.copy()
                                    
                                    # Add Status column for read-only view too
                                    if 'Deal ID' in df_readonly.columns:
                                        df_readonly['Status'] = df_readonly['Deal ID'].apply(
                                            lambda deal_id: get_planning_status(deal_id) if get_planning_status(deal_id) else '—'
                                        )
                                        display_readonly = ['Status'] + cols
                                    else:
                                        display_readonly = cols
                                    
                                    st.dataframe(
                                        df_readonly[display_readonly],
                                        column_config={
                                            "Status": st.column_config.TextColumn("Q4 Status", width="small"),
                                            "Link": st.column_config.LinkColumn("🔗", display_text="Open", width="small"),
                                            "Deal ID": st.column_config.TextColumn("Deal ID", width="small"),
                                            "Type": st.column_config.TextColumn("Type", width="small"),
                                            "Close": st.column_config.TextColumn("Close Date", width="small"),
                                            "PA Date": st.column_config.TextColumn("PA Date", width="small"),
                                            "Amount_Numeric": st.column_config.NumberColumn("Amount", format="$%d")
                                        },
                                        hide_index=True,
                                        use_container_width=True
                                    )
                                    # Capture all rows for export (with Status column)
                                    # BUT only include items with IN/MAYBE status IF planning status exists
                                    if 'Deal ID' in df.columns:
                                        df_export = df.copy()
                                        df_export['Status'] = df_export['Deal ID'].apply(
                                            lambda deal_id: get_planning_status(deal_id) if get_planning_status(deal_id) else '—'
                                        )
                                        # Only filter if planning status has actual entries
                                        if st.session_state[planning_key] and len(st.session_state[planning_key]) > 0:
                                            df_export['_should_include'] = df_export['Deal ID'].apply(
                                                lambda deal_id: get_planning_status(deal_id) in ['IN', 'MAYBE']
                                            )
                                            df_export = df_export[df_export['_should_include']].drop(columns=['_should_include'])
                                        export_buckets[key] = df_export
                                    else:
                                        export_buckets[key] = df
                        
                        # Fallback: if export_buckets wasn't set by expander code, set default
                        if key not in export_buckets:
                            df_default = df.copy()
                            if 'Deal ID' in df_default.columns:
                                df_default['Status'] = df_default['Deal ID'].apply(
                                    lambda deal_id: get_planning_status(deal_id) if get_planning_status(deal_id) else '—'
                                )
                                # Only filter if planning status exists
                                if st.session_state[planning_key] and len(st.session_state[planning_key]) > 0:
                                    df_default['_should_include'] = df_default['Deal ID'].apply(
                                        lambda deal_id: get_planning_status(deal_id) in ['IN', 'MAYBE']
                                    )
                                    df_default = df_default[df_default['_should_include']].drop(columns=['_should_include'])
                            export_buckets[key] = df_default

    else:
        # === CONSOLIDATED VIEW ===
        st.markdown("### 📋 All Items (Consolidated)")
        
        # Combine all NetSuite and HubSpot dataframes with category labels
        all_items = []
        
        # Add NetSuite items
        for key, data in ns_categories.items():
            df = ns_dfs.get(key, pd.DataFrame())
            if not df.empty:
                df_with_cat = df.copy()
                df_with_cat['Category'] = data['label']
                df_with_cat['Source'] = 'NetSuite'
                # Add Status and Select columns
                if 'SO #' in df_with_cat.columns:
                    df_with_cat['Status'] = df_with_cat['SO #'].apply(
                        lambda so: get_planning_status(so) if get_planning_status(so) else '—'
                    )
                    # Default to True if no planning status, otherwise filter by IN/MAYBE
                    if not st.session_state[planning_key] or len(st.session_state[planning_key]) == 0:
                        df_with_cat['Select'] = True
                    else:
                        df_with_cat['Select'] = df_with_cat['SO #'].apply(
                            lambda so: get_planning_status(so) in ['IN', 'MAYBE']
                        )
                all_items.append(df_with_cat)
        
        # Add HubSpot items
        for key, data in hs_categories.items():
            df = hs_dfs.get(key, pd.DataFrame())
            if not df.empty:
                df_with_cat = df.copy()
                df_with_cat['Category'] = data['label']
                df_with_cat['Source'] = 'HubSpot'
                # Add Status and Select columns
                if 'Deal ID' in df_with_cat.columns:
                    df_with_cat['Status'] = df_with_cat['Deal ID'].apply(
                        lambda deal_id: get_planning_status(deal_id) if get_planning_status(deal_id) else '—'
                    )
                    # Default to True if no planning status, otherwise filter by IN/MAYBE
                    if not st.session_state[planning_key] or len(st.session_state[planning_key]) == 0:
                        df_with_cat['Select'] = True
                    else:
                        df_with_cat['Select'] = df_with_cat['Deal ID'].apply(
                            lambda deal_id: get_planning_status(deal_id) in ['IN', 'MAYBE']
                        )
                all_items.append(df_with_cat)
        
        if all_items:
            # Combine all items
            combined_df = pd.concat(all_items, ignore_index=True)
            
            # Display columns for consolidated view
            display_cols_consolidated = ['Select', 'Status', 'Category', 'Source']
            
            # Add ID column (SO# or Deal ID)
            if 'SO #' in combined_df.columns:
                combined_df['ID'] = combined_df['SO #'].fillna(combined_df.get('Deal ID', ''))
            elif 'Deal ID' in combined_df.columns:
                combined_df['ID'] = combined_df['Deal ID']
            display_cols_consolidated.append('ID')
            
            # Add common columns
            if 'Customer' in combined_df.columns:
                display_cols_consolidated.append('Customer')
            if 'Deal Name' in combined_df.columns:
                combined_df['Customer'] = combined_df['Customer'].fillna(combined_df.get('Deal Name', ''))
                if 'Customer' not in display_cols_consolidated:
                    display_cols_consolidated.append('Customer')
            
            if 'Type' in combined_df.columns:
                display_cols_consolidated.append('Type')
            if 'Classification Date' in combined_df.columns:
                combined_df['Date'] = combined_df['Classification Date'].fillna(combined_df.get('Close', ''))
            elif 'Close' in combined_df.columns:
                combined_df['Date'] = combined_df['Close']
            if 'Date' in combined_df.columns:
                display_cols_consolidated.append('Date')
            
            # Amount column
            if 'Amount' in combined_df.columns and 'Amount_Numeric' in combined_df.columns:
                combined_df['Amount_Display'] = combined_df['Amount'].fillna(combined_df['Amount_Numeric'])
            elif 'Amount' in combined_df.columns:
                combined_df['Amount_Display'] = combined_df['Amount']
            elif 'Amount_Numeric' in combined_df.columns:
                combined_df['Amount_Display'] = combined_df['Amount_Numeric']
            display_cols_consolidated.append('Amount_Display')
            
            # Filter to only columns that exist
            display_cols_consolidated = [c for c in display_cols_consolidated if c in combined_df.columns]
            
            # Editable consolidated view
            edited_consolidated = st.data_editor(
                combined_df[display_cols_consolidated],
                column_config={
                    "Select": st.column_config.CheckboxColumn("✓", width="small"),
                    "Status": st.column_config.SelectboxColumn(
                        "Q4 Status", 
                        width="small",
                        options=['IN', 'MAYBE', 'OUT', '—'],
                        required=False
                    ),
                    "Category": st.column_config.TextColumn("Category", width="medium"),
                    "Source": st.column_config.TextColumn("Source", width="small"),
                    "ID": st.column_config.TextColumn("ID", width="small"),
                    "Customer": st.column_config.TextColumn("Customer", width="medium"),
                    "Type": st.column_config.TextColumn("Type", width="small"),
                    "Date": st.column_config.TextColumn("Date", width="small"),
                    "Amount_Display": st.column_config.NumberColumn("Amount", format="$%d")
                },
                disabled=['Category', 'Source', 'ID', 'Customer', 'Type', 'Date', 'Amount_Display'],
                hide_index=True,
                key=f"consolidated_edit_{rep_name}",
                num_rows="fixed",
                height=600
            )
            
            # Update planning status from edits
            for idx, row in edited_consolidated.iterrows():
                # Determine ID and update status
                item_id = str(row.get('ID', '')).strip()
                status = str(row.get('Status', '—')).strip().upper()
                
                if item_id and status != '—' and status in ['IN', 'MAYBE', 'OUT']:
                    st.session_state[planning_key][item_id] = status
                elif item_id and status == '—' and item_id in st.session_state[planning_key]:
                    del st.session_state[planning_key][item_id]
            
            # Auto-update Select based on Status changes
            if 'Status' in edited_consolidated.columns and 'Select' in edited_consolidated.columns:
                for idx in edited_consolidated.index:
                    status = str(edited_consolidated.at[idx, 'Status']).strip().upper()
                    if status in ['IN', 'MAYBE']:
                        edited_consolidated.at[idx, 'Select'] = True
                    elif status in ['OUT', '—']:
                        edited_consolidated.at[idx, 'Select'] = False
            
            # Split back into category buckets for export
            for key in list(ns_categories.keys()) + list(hs_categories.keys()):
                label = ns_categories.get(key, hs_categories.get(key, {})).get('label', key)
                cat_items = edited_consolidated[
                    (edited_consolidated['Category'] == label) & 
                    (edited_consolidated['Select'] == True)
                ].copy()
                
                if not cat_items.empty:
                    # Map back to original dataframe structure
                    # IMPORTANT: Keep the original Amount/Amount_Numeric columns for calculations
                    if key in ns_categories:
                        # NetSuite - ensure 'Amount' column exists
                        if 'Amount' not in cat_items.columns and 'Amount_Display' in cat_items.columns:
                            cat_items['Amount'] = cat_items['Amount_Display']
                        export_buckets[key] = cat_items
                    else:
                        # HubSpot - ensure 'Amount_Numeric' column exists
                        if 'Amount_Numeric' not in cat_items.columns and 'Amount_Display' in cat_items.columns:
                            cat_items['Amount_Numeric'] = cat_items['Amount_Display']
                        export_buckets[key] = cat_items
            
            # Show selection summary
            selected_count = edited_consolidated['Select'].sum()
            selected_total = edited_consolidated[edited_consolidated['Select']]['Amount_Display'].sum()
            st.caption(f"Selected: {selected_count} items = ${selected_total:,.0f}")
            st.caption("💡 Tip: Changing Status to IN/MAYBE auto-selects the item, OUT/— auto-deselects")
        else:
            st.info("No items to display")

    # --- 5. CALCULATE RESULTS ---
    
    # Calculate totals from export buckets (which reflect custom selections)
    def safe_sum(df):
        if df.empty:
            return 0
        # Handle both Amount and Amount_Numeric columns (NS uses Amount, HS uses Amount_Numeric)
        if 'Amount_Numeric' in df.columns:
            return df['Amount_Numeric'].sum()
        elif 'Amount' in df.columns:
            return df['Amount'].sum()
        else:
            return 0
    
    selected_pending = sum(safe_sum(df) for k, df in export_buckets.items() if k in ns_categories)
    selected_pipeline = sum(safe_sum(df) for k, df in export_buckets.items() if k in hs_categories)
    
    total_forecast = invoiced_shipped + selected_pending + selected_pipeline
    gap_to_quota = quota - total_forecast
    
    st.markdown("---")
    st.markdown("### 🔮 Forecast Scenario Results")
    
    # Add locked revenue banner
    st.markdown(f"""
    <div class="locked-revenue-banner">
        <div class="banner-left">
            <span class="banner-icon">🔒</span>
            <div>
                <div class="banner-label">Locked Revenue</div>
                <div class="banner-value">${invoiced_shipped:,.0f}</div>
            </div>
        </div>
        <div class="banner-right">
            <div class="banner-label">Status</div>
            <div class="banner-status">INVOICED & SECURED</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.metric("1. Invoiced", f"${invoiced_shipped:,.0f}")
    with m2: st.metric("2. Selected Pending", f"${selected_pending:,.0f}")
    with m3: st.metric("3. Selected Pipeline", f"${selected_pipeline:,.0f}")
    with m4: st.metric("🏁 Total Forecast", f"${total_forecast:,.0f}", delta="Sum of 1+2+3")
    with m5:
        if gap_to_quota > 0:
            st.metric("Gap to Quota", f"${gap_to_quota:,.0f}", delta="Behind", delta_color="inverse")
        else:
            st.metric("Gap to Quota", f"${abs(gap_to_quota):,.0f}", delta="Ahead!", delta_color="normal")

    c1, c2 = st.columns([2, 1])
    with c1:
        # Use the enhanced sexy gauge with color zones
        fig = create_sexy_gauge(total_forecast, quota, "Progress to Quota")
        st.plotly_chart(fig, use_container_width=True)
        
    with c2:
        # Helper locally if needed
        def calculate_biz_days():
             from datetime import date, timedelta
             today = date.today()
             q4_end = date(2025, 12, 31)
             holidays = [date(2025, 11, 27), date(2025, 11, 28), date(2025, 12, 25), date(2025, 12, 26)]
             days = 0
             current = today
             while current <= q4_end:
                 if current.weekday() < 5 and current not in holidays: days += 1
                 current += timedelta(days=1)
             return days

        biz_days = calculate_biz_days()
        # Calculate based on what still needs to ship (pending orders + pipeline deals)
        items_to_ship = selected_pending + selected_pipeline
        if items_to_ship > 0 and biz_days > 0:
            required = items_to_ship / biz_days
            st.metric("Required Ship Rate", f"${required:,.0f}/day", f"{biz_days} days left")
        elif items_to_ship == 0:
            st.info("✅ No pending items to ship")
        elif gap_to_quota <= 0:
            st.success("🎉 Scenario Hits Quota!")

    # --- 6. ROBUST EXPORT FUNCTIONALITY ---
    if total_forecast > 0:
        st.markdown("---")
        
        # Initialize Lists
        export_summary = []
        export_data = []
        
        # A. Build Summary
        export_summary.append({'Category': '=== FORECAST SUMMARY ===', 'Amount': ''})
        export_summary.append({'Category': 'Quota', 'Amount': f"${quota:,.0f}"})
        export_summary.append({'Category': 'Invoiced (Always Included)', 'Amount': f"${invoiced_shipped:,.0f}"})
        export_summary.append({'Category': 'Pending Orders (Selected)', 'Amount': f"${selected_pending:,.0f}"})
        export_summary.append({'Category': 'Pipeline Deals (Selected)', 'Amount': f"${selected_pipeline:,.0f}"})
        export_summary.append({'Category': 'Total Forecast', 'Amount': f"${total_forecast:,.0f}"})
        export_summary.append({'Category': 'Gap to Goal', 'Amount': f"${gap_to_quota:,.0f}"})
        export_summary.append({'Category': '', 'Amount': ''})
        export_summary.append({'Category': '=== SELECTED COMPONENTS ===', 'Amount': ''})
        
        # Add Component Totals
        for key, df in export_buckets.items():
            # Handle both Amount and Amount_Numeric columns (NS uses Amount, HS uses Amount_Numeric)
            if 'Amount_Numeric' in df.columns:
                cat_val = df['Amount_Numeric'].sum()
            elif 'Amount' in df.columns:
                cat_val = df['Amount'].sum()
            else:
                cat_val = 0
                
            if cat_val > 0:
                label = ns_categories.get(key, hs_categories.get(key, {})).get('label', key)
                count = len(df)
                export_summary.append({'Category': f"{label} ({count} items)", 'Amount': f"${cat_val:,.0f}"})
        
        export_summary.append({'Category': '', 'Amount': ''})
        export_summary.append({'Category': '=== DETAILED LINE ITEMS ===', 'Amount': ''})
        
        # B. Build Line Items
        
        # 1. Invoices
        if invoices_df is not None and not invoices_df.empty:
            # Filter for rep if needed (using Sales Rep col)
            inv_source = invoices_df
            if rep_name and 'Sales Rep' in invoices_df.columns:
                inv_source = invoices_df[invoices_df['Sales Rep'] == rep_name]
                
            for _, row in inv_source.iterrows():
                export_data.append({
                    'Category': 'Invoice',
                    'ID': row.get('Document Number', row.get('Invoice Number', '')),
                    'Customer': row.get('Account Name', row.get('Customer', '')),
                    'Order/Deal Type': '',
                    'Date': str(row.get('Date', '')),
                    'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                    'Rep': row.get('Sales Rep', '')
                })
        
        # 2. Pending & Pipeline Items from Buckets
        for key, df in export_buckets.items():
            label = ns_categories.get(key, hs_categories.get(key, {})).get('label', key)
            
            for _, row in df.iterrows():
                # Get planning status for this item
                # First try to get it from the dataframe (if it's been edited), otherwise from session state
                if 'Status' in row and pd.notna(row['Status']) and str(row['Status']).strip() != '':
                    planning_status = str(row['Status']).strip()
                else:
                    if key in ns_categories:  # NetSuite
                        item_id_for_status = row.get('SO #', '')
                    else:  # HubSpot
                        item_id_for_status = row.get('Deal ID', row.get('Record ID', ''))
                    
                    planning_status = get_planning_status(item_id_for_status) if item_id_for_status else '—'
                    if not planning_status:
                        planning_status = '—'
                
                # Determine fields based on source type (NS vs HS)
                if key in ns_categories: # NetSuite
                    item_type = f"Sales Order - {label}"
                    item_id = row.get('SO #', row.get('Document Number', ''))
                    cust = row.get('Customer', '')
                    date_val = row.get('Classification Date', row.get('Key Date', ''))
                    deal_type = row.get('Type', row.get('Display_Type', ''))
                    # NetSuite uses 'Amount' not 'Amount_Numeric'
                    amount = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                    rep = row.get('Sales Rep', rep_name)
                else: # HubSpot
                    item_type = f"HubSpot - {label}"
                    item_id = row.get('Deal ID', row.get('Record ID', ''))
                    cust = row.get('Account Name', row.get('Deal Name', '')) # Fallback to Deal Name if Account missing
                    date_val = row.get('Close', row.get('Close Date', ''))
                    deal_type = row.get('Type', row.get('Display_Type', ''))
                    amount = pd.to_numeric(row.get('Amount_Numeric', 0), errors='coerce')
                    rep = row.get('Deal Owner', rep_name)
                
                export_data.append({
                    'Category': item_type,
                    'ID': item_id,
                    'Customer': cust,
                    'Order/Deal Type': deal_type,
                    'Date': str(date_val),
                    'Amount': amount,
                    'Q4 Status': planning_status,
                    'Rep': rep
                })

        # C. Construct CSV
        if export_data:
            summary_df = pd.DataFrame(export_summary)
            data_df = pd.DataFrame(export_data)
            
            # Format Amount in Data DF
            data_df['Amount'] = data_df['Amount'].apply(lambda x: f"${x:,.2f}")
            
            final_csv = summary_df.to_csv(index=False) + "\n" + data_df.to_csv(index=False)
            
            st.download_button(
                label="📥 Download Winning Pipeline",
                data=final_csv,
                file_name=f"winning_pipeline_{rep_name if rep_name else 'team'}_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
            st.caption(f"Export includes summary + {len(data_df)} line items.")
def display_hubspot_deals_audit(deals_df, rep_name=None):
    """
    Display audit section for HubSpot deals without amounts
    """
    st.markdown("### ⚠️ HubSpot Deals without Amounts (AUDIT!)")
    st.caption("These deals are missing amount data and need attention")
    
    if deals_df is None or deals_df.empty:
        st.info("No HubSpot deals data available")
        return
    
    # Filter by rep if specified
    if rep_name and 'Deal Owner' in deals_df.columns:
        filtered_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    else:
        filtered_deals = deals_df.copy()
    
    if filtered_deals.empty:
        st.info(f"No deals found{' for ' + rep_name if rep_name else ''}")
        return
    
    # Convert Amount to numeric and find deals without amounts
    filtered_deals['Amount_Numeric'] = pd.to_numeric(filtered_deals['Amount'], errors='coerce')
    deals_no_amount = filtered_deals[
        (filtered_deals['Amount_Numeric'].isna()) | 
        (filtered_deals['Amount_Numeric'] == 0)
    ].copy()
    
    if deals_no_amount.empty:
        st.success("✅ All deals have amounts! No issues to audit.")
        return
    
    # Show summary
    st.warning(f"⚠️ Found {len(deals_no_amount)} deals without amounts")
    
    # Break down by status
    if 'Status' in deals_no_amount.columns:
        status_categories = ['Expect', 'Commit', 'Best Case', 'Opportunity']
        
        for status in status_categories:
            status_deals = deals_no_amount[deals_no_amount['Status'] == status].copy()
            
            if not status_deals.empty:
                with st.expander(f"🔍 {status} - {len(status_deals)} deals"):
                    # Create display dataframe
                    display_data = []
                    
                    for _, row in status_deals.iterrows():
                        # Build HubSpot link if we have Record ID
                        deal_link = ""
                        record_id = row.get('Record ID', '')
                        if record_id:
                            deal_link = f"https://app.hubspot.com/contacts/6554605/deal/{record_id}"
                        
                        display_data.append({
                            'Link': deal_link,
                            'Deal Name': row.get('Deal Name', ''),
                            'Amount': '$0.00',
                            'Status': row.get('Status', ''),
                            'Pipeline': row.get('Pipeline', ''),
                            'Close Date': row.get('Close Date', ''),
                            'Product Type': row.get('Product Type', '')
                        })
                    
                    if display_data:
                        display_df = pd.DataFrame(display_data)
                        
                        # Format as clickable links
                        if 'Link' in display_df.columns:
                            display_df['Link'] = display_df['Link'].apply(
                                lambda x: f'<a href="{x}" target="_blank">View Deal</a>' if x else ''
                            )
                        
                        # Display the table with HTML links
                        st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)
                    else:
                        st.info("No deals to display")
    else:
        st.warning("Status column not found in deals data")

def calculate_team_metrics(deals_df, dashboard_df):
    """Calculate overall team metrics"""
    
    total_quota = dashboard_df['Quota'].sum()
    total_orders = dashboard_df['NetSuite Orders'].sum()
    
    # Filter for Q4 fulfillment only (spreadsheet formula now handles PA date logic)
    deals_q4 = deals_df[deals_df.get('Q1 2026 Spillover') != 'Q1 2026']
    
    # Calculate Expect/Commit forecast (Q4 only)
    expect_commit = deals_q4[deals_q4['Status'].isin(['Expect', 'Commit'])]['Amount'].sum()
    
    # Calculate Best Case/Opportunity (Q4 only)
    best_opp = deals_q4[deals_q4['Status'].isin(['Best Case', 'Opportunity'])]['Amount'].sum()
    
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

# ========== CENTRALIZED SALES ORDER CATEGORIZATION ==========
def categorize_sales_orders(sales_orders_df, rep_name=None):
    """
    SINGLE SOURCE OF TRUTH for categorizing sales orders into forecast buckets.
    
    This function ensures consistent categorization across:
    - Team Dashboard bar charts
    - Individual Rep views
    - Build Your Own Forecast section
    
    Returns a dictionary with categorized DataFrames and their amounts.
    """
    if sales_orders_df is None or sales_orders_df.empty:
        return {
            'pf_date_ext': pd.DataFrame(), 'pf_date_ext_amount': 0,
            'pf_date_int': pd.DataFrame(), 'pf_date_int_amount': 0,
            'pf_nodate_ext': pd.DataFrame(), 'pf_nodate_ext_amount': 0,
            'pf_nodate_int': pd.DataFrame(), 'pf_nodate_int_amount': 0,
            'pa_date': pd.DataFrame(), 'pa_date_amount': 0,
            'pa_nodate': pd.DataFrame(), 'pa_nodate_amount': 0,
            'pa_old': pd.DataFrame(), 'pa_old_amount': 0
        }
    
    # Filter by rep if specified
    if rep_name and 'Sales Rep' in sales_orders_df.columns:
        orders = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
    else:
        orders = sales_orders_df.copy()
    
    if orders.empty:
        return {
            'pf_date_ext': pd.DataFrame(), 'pf_date_ext_amount': 0,
            'pf_date_int': pd.DataFrame(), 'pf_date_int_amount': 0,
            'pf_nodate_ext': pd.DataFrame(), 'pf_nodate_ext_amount': 0,
            'pf_nodate_int': pd.DataFrame(), 'pf_nodate_int_amount': 0,
            'pa_date': pd.DataFrame(), 'pa_date_amount': 0,
            'pa_nodate': pd.DataFrame(), 'pa_nodate_amount': 0,
            'pa_old': pd.DataFrame(), 'pa_old_amount': 0
        }
    
    # Remove duplicate columns
    if orders.columns.duplicated().any():
        orders = orders.loc[:, ~orders.columns.duplicated()]
    
    # === ADD DISPLAY COLUMNS FOR UI ===
    # Add display columns to orders dataframe
    orders['Display_SO_Num'] = get_col_by_index(orders, 1)  # Col B: SO#
    orders['Display_Type'] = get_col_by_index(orders, 17).fillna('Standard')  # Col R: Order Type
    orders['Display_Promise_Date'] = pd.to_datetime(get_col_by_index(orders, 11), errors='coerce')  # Col L: Promise Date
    orders['Display_Projected_Date'] = pd.to_datetime(get_col_by_index(orders, 12), errors='coerce')  # Col M: Projected Date
    
    # PA Date handling
    if 'Pending Approval Date' in orders.columns:
        orders['Display_PA_Date'] = pd.to_datetime(orders['Pending Approval Date'], errors='coerce')
    else:
        orders['Display_PA_Date'] = pd.to_datetime(get_col_by_index(orders, 29), errors='coerce')  # Col AD: PA Date
    
    # Define Q4 2025 date range
    q4_start = pd.Timestamp('2025-10-01')
    q4_end = pd.Timestamp('2025-12-31')
    
    # === PENDING FULFILLMENT CATEGORIZATION ===
    pf_orders = orders[orders['Status'].isin(['Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])].copy()
    
    if not pf_orders.empty:
        # Check if dates are in Q4 range
        def has_q4_date(row):
            if pd.notna(row.get('Customer Promise Date')):
                if q4_start <= row['Customer Promise Date'] <= q4_end:
                    return True
            if pd.notna(row.get('Projected Date')):
                if q4_start <= row['Projected Date'] <= q4_end:
                    return True
            return False
        
        pf_orders['Has_Q4_Date'] = pf_orders.apply(has_q4_date, axis=1)
        
        # Check External/Internal flag
        is_ext = pd.Series(False, index=pf_orders.index)
        if 'Calyx External Order' in pf_orders.columns:
            is_ext = pf_orders['Calyx External Order'].astype(str).str.strip().str.upper() == 'YES'
        
        # Categorize PF orders
        pf_date_ext = pf_orders[(pf_orders['Has_Q4_Date'] == True) & is_ext].copy()
        pf_date_int = pf_orders[(pf_orders['Has_Q4_Date'] == True) & ~is_ext].copy()
        
        # No date means BOTH dates are missing
        no_date_mask = (
            (pf_orders['Customer Promise Date'].isna()) &
            (pf_orders['Projected Date'].isna())
        )
        pf_nodate_ext = pf_orders[no_date_mask & is_ext].copy()
        pf_nodate_int = pf_orders[no_date_mask & ~is_ext].copy()
    else:
        pf_date_ext = pf_date_int = pf_nodate_ext = pf_nodate_int = pd.DataFrame()
    
    # === PENDING APPROVAL CATEGORIZATION ===
    # Logic:
    # 1. PA with Date (within Q4): Has PA Date in Q4 2025 AND Age < 13 business days
    # 2. PA No Date: No PA Date AND Age < 13 business days  
    # 3. PA Old (>2 Weeks): Age >= 13 business days (regardless of whether they have a PA date or not)
    pa_orders = orders[orders['Status'] == 'Pending Approval'].copy()
    
    if not pa_orders.empty:
        # Check if Age_Business_Days column exists
        if 'Age_Business_Days' not in pa_orders.columns:
            pa_orders['Age_Business_Days'] = 0
        
        # Parse Pending Approval Date for all PA orders
        if 'Pending Approval Date' in pa_orders.columns:
            pa_orders['PA_Date_Parsed'] = pd.to_datetime(pa_orders['Pending Approval Date'], errors='coerce')
            
            # Fix any dates that got parsed as 1900s (2-digit year issue: 26 -> 1926 instead of 2026)
            if pa_orders['PA_Date_Parsed'].notna().any():
                mask_1900s = (pa_orders['PA_Date_Parsed'].dt.year < 2000) & (pa_orders['PA_Date_Parsed'].notna())
                if mask_1900s.any():
                    pa_orders.loc[mask_1900s, 'PA_Date_Parsed'] = pa_orders.loc[mask_1900s, 'PA_Date_Parsed'] + pd.DateOffset(years=100)
        else:
            pa_orders['PA_Date_Parsed'] = pd.NaT
        
        # Determine which orders have a valid Q4 PA Date
        has_q4_pa_date = (
            (pa_orders['PA_Date_Parsed'].notna()) &
            (pa_orders['PA_Date_Parsed'] >= q4_start) &
            (pa_orders['PA_Date_Parsed'] <= q4_end)
        )
        
        # Determine which orders have NO PA Date (or invalid/outside Q4)
        has_no_pa_date = (
            (pa_orders['PA_Date_Parsed'].isna()) |
            (pa_orders['Pending Approval Date'].astype(str).str.strip() == 'No Date') |
            (pa_orders['Pending Approval Date'].astype(str).str.strip() == '')
        )
        
        # CATEGORY 3: PA Old (>= 13 business days) - ANY PA order that is old, regardless of PA date
        # This takes priority - old orders go here first
        pa_old = pa_orders[pa_orders['Age_Business_Days'] >= 13].copy()
        
        # Only "young" orders (< 13 days) are eligible for PA with Date or PA No Date
        young_pa = pa_orders[pa_orders['Age_Business_Days'] < 13].copy()
        
        # CATEGORY 1: PA with Date - has Q4 PA date AND is NOT old (< 13 days)
        pa_date = young_pa[has_q4_pa_date.loc[young_pa.index]].copy() if not young_pa.empty else pd.DataFrame()
        
        # CATEGORY 2: PA No Date - no PA date AND is NOT old (< 13 days)
        pa_nodate = young_pa[has_no_pa_date.loc[young_pa.index]].copy() if not young_pa.empty else pd.DataFrame()
    else:
        pa_old = pa_date = pa_nodate = pd.DataFrame()
    
    # Calculate amounts
    def get_amount(df):
        return df['Amount'].sum() if not df.empty and 'Amount' in df.columns else 0
    
    return {
        'pf_date_ext': pf_date_ext,
        'pf_date_ext_amount': get_amount(pf_date_ext),
        'pf_date_int': pf_date_int,
        'pf_date_int_amount': get_amount(pf_date_int),
        'pf_nodate_ext': pf_nodate_ext,
        'pf_nodate_ext_amount': get_amount(pf_nodate_ext),
        'pf_nodate_int': pf_nodate_int,
        'pf_nodate_int_amount': get_amount(pf_nodate_int),
        'pa_date': pa_date,
        'pa_date_amount': get_amount(pa_date),
        'pa_nodate': pa_nodate,
        'pa_nodate_amount': get_amount(pa_nodate),
        'pa_old': pa_old,
        'pa_old_amount': get_amount(pa_old)
    }

def calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df=None):
    """Calculate metrics for a specific rep with detailed order lists for drill-down"""
    
    # Get rep's quota and orders
    rep_info = dashboard_df[dashboard_df['Rep Name'] == rep_name]
    
    if rep_info.empty:
        return None
    
    quota = rep_info['Quota'].iloc[0]
    orders = rep_info['NetSuite Orders'].iloc[0]
    
    # Filter deals for this rep - ALL Q4 2025 deals (regardless of spillover)
    rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    
    # Check if we have the Q1 2026 Spillover column (spreadsheet formula now handles PA date logic)
    has_spillover_column = 'Q1 2026 Spillover' in rep_deals.columns
    
    if has_spillover_column:
        # Separate deals by shipping timeline using spreadsheet formula
        rep_deals['Ships_In_Q4'] = rep_deals['Q1 2026 Spillover'] != 'Q1 2026'
        rep_deals['Ships_In_Q1'] = rep_deals['Q1 2026 Spillover'] == 'Q1 2026'
        
        # Deals that ship in Q4 2025
        rep_deals_ship_q4 = rep_deals[rep_deals['Ships_In_Q4'] == True].copy()
        
        # Deals that ship in Q1 2026 (spillover)
        rep_deals_ship_q1 = rep_deals[rep_deals['Ships_In_Q1'] == True].copy()
    else:
        # Fallback if column doesn't exist - treat all as Q4
        rep_deals_ship_q4 = rep_deals.copy()
        rep_deals_ship_q1 = pd.DataFrame()
    
    # Calculate metrics for DEALS SHIPPING IN Q4 (this counts toward quota)
    expect_commit_q4_deals = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Expect', 'Commit'])].copy()
    if expect_commit_q4_deals.columns.duplicated().any():
        expect_commit_q4_deals = expect_commit_q4_deals.loc[:, ~expect_commit_q4_deals.columns.duplicated()]
    expect_commit_q4 = expect_commit_q4_deals['Amount'].sum() if not expect_commit_q4_deals.empty else 0
    
    best_opp_q4_deals = rep_deals_ship_q4[rep_deals_ship_q4['Status'].isin(['Best Case', 'Opportunity'])].copy()
    if best_opp_q4_deals.columns.duplicated().any():
        best_opp_q4_deals = best_opp_q4_deals.loc[:, ~best_opp_q4_deals.columns.duplicated()]
    best_opp_q4 = best_opp_q4_deals['Amount'].sum() if not best_opp_q4_deals.empty else 0
    
    # Calculate metrics for Q1 SPILLOVER DEALS (closing in Q4 but shipping in Q1)
    expect_commit_q1_deals = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Expect', 'Commit'])].copy()
    if expect_commit_q1_deals.columns.duplicated().any():
        expect_commit_q1_deals = expect_commit_q1_deals.loc[:, ~expect_commit_q1_deals.columns.duplicated()]
    expect_commit_q1_spillover = expect_commit_q1_deals['Amount'].sum() if not expect_commit_q1_deals.empty else 0
    
    best_opp_q1_deals = rep_deals_ship_q1[rep_deals_ship_q1['Status'].isin(['Best Case', 'Opportunity'])].copy()
    if best_opp_q1_deals.columns.duplicated().any():
        best_opp_q1_deals = best_opp_q1_deals.loc[:, ~best_opp_q1_deals.columns.duplicated()]
    best_opp_q1_spillover = best_opp_q1_deals['Amount'].sum() if not best_opp_q1_deals.empty else 0
    
    # Total Q1 spillover
    q1_spillover_total = expect_commit_q1_spillover + best_opp_q1_spillover
    
    # === USE CENTRALIZED CATEGORIZATION FUNCTION ===
    so_categories = categorize_sales_orders(sales_orders_df, rep_name)
    
    # Extract amounts
    pending_fulfillment = so_categories['pf_date_ext_amount'] + so_categories['pf_date_int_amount']
    pending_fulfillment_no_date = so_categories['pf_nodate_ext_amount'] + so_categories['pf_nodate_int_amount']
    pending_approval = so_categories['pa_date_amount']
    pending_approval_no_date = so_categories['pa_nodate_amount']
    pending_approval_old = so_categories['pa_old_amount']
    
    # Extract detail dataframes
    pending_approval_details = so_categories['pa_date']
    pending_approval_no_date_details = so_categories['pa_nodate']
    pending_approval_old_details = so_categories['pa_old']
    pending_fulfillment_details = pd.concat([so_categories['pf_date_ext'], so_categories['pf_date_int']])
    pending_fulfillment_no_date_details = pd.concat([so_categories['pf_nodate_ext'], so_categories['pf_nodate_int']])
    
    # Total calculations - ONLY Q4 SHIPPING DEALS COUNT TOWARD QUOTA
    total_pending_fulfillment = pending_fulfillment + pending_fulfillment_no_date
    total_progress = orders + expect_commit_q4 + pending_approval + pending_fulfillment
    gap = quota - total_progress
    attainment_pct = (total_progress / quota * 100) if quota > 0 else 0
    potential_attainment = ((total_progress + best_opp_q4) / quota * 100) if quota > 0 else 0
    
    return {
        'quota': quota,
        'orders': orders,
        'expect_commit': expect_commit_q4,  # Only Q4 shipping deals
        'best_opp': best_opp_q4,  # Only Q4 shipping deals
        'gap': gap,
        'attainment_pct': attainment_pct,
        'potential_attainment': potential_attainment,
        'total_progress': total_progress,
        'pending_approval': pending_approval,
        'pending_approval_no_date': pending_approval_no_date,
        'pending_approval_old': pending_approval_old,
        'pending_fulfillment': pending_fulfillment,
        'pending_fulfillment_no_date': pending_fulfillment_no_date,
        'total_pending_fulfillment': total_pending_fulfillment,
        
        # NEW: Q1 Spillover metrics
        'q1_spillover_expect_commit': expect_commit_q1_spillover,
        'q1_spillover_best_opp': best_opp_q1_spillover,
        'q1_spillover_total': q1_spillover_total,
        
        # ALL Q4 2025 closing deals (for reference)
        'total_q4_closing_deals': len(rep_deals),
        'total_q4_closing_amount': rep_deals['Amount'].sum() if not rep_deals.empty else 0,
        
        'deals': rep_deals_ship_q4,  # Deals shipping in Q4
        
        # Add detail dataframes for drill-down
        'pending_approval_details': pending_approval_details,
        'pending_approval_no_date_details': pending_approval_no_date_details,
        'pending_approval_old_details': pending_approval_old_details,
        'pending_fulfillment_details': pending_fulfillment_details,
        'pending_fulfillment_no_date_details': pending_fulfillment_no_date_details,
        'expect_commit_deals': expect_commit_q4_deals,
        'best_opp_deals': best_opp_q4_deals,
        
        # NEW: Q1 Spillover deal details
        'expect_commit_q1_spillover_deals': expect_commit_q1_deals,
        'best_opp_q1_spillover_deals': best_opp_q1_deals,
        'all_q1_spillover_deals': rep_deals_ship_q1
    }

# ========== ENHANCED CHART FUNCTIONS (GEMINI ENHANCEMENTS) ==========

def create_sexy_gauge(current_val, target_val, title="Progress to Quota"):
    """Enhanced gauge with color zones and delta reference"""
    fig = go.Figure(go.Indicator(
        mode = "gauge+number+delta",
        value = current_val,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': title, 'font': {'size': 20, 'color': 'white'}},
        delta = {
            'reference': target_val, 
            'increasing': {'color': "#10b981"},
            'decreasing': {'color': "#ef4444"},
            'font': {'size': 16}
        },
        number = {'font': {'size': 32, 'color': 'white'}},
        gauge = {
            'axis': {
                'range': [None, target_val * 1.2], 
                'tickwidth': 1, 
                'tickcolor': "rgba(255,255,255,0.3)",
                'tickfont': {'color': 'rgba(255,255,255,0.7)', 'size': 10}
            },
            'bar': {'color': "#3b82f6", 'thickness': 0.8},
            'bgcolor': "rgba(0,0,0,0)",
            'borderwidth': 2,
            'bordercolor': "rgba(255,255,255,0.2)",
            'steps': [
                {'range': [0, target_val * 0.7], 'color': 'rgba(239, 68, 68, 0.2)'},  # Red zone
                {'range': [target_val * 0.7, target_val], 'color': 'rgba(251, 191, 36, 0.2)'}  # Yellow zone
            ],
            'threshold': {
                'line': {'color': "#10b981", 'width': 4},
                'thickness': 0.75,
                'value': target_val
            }
        }
    ))
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white'},
        height=280,
        margin=dict(l=40, r=40, t=60, b=40)
    )
    return fig

def create_pipeline_sankey(deals_df):
    """Sankey diagram showing pipeline flow from Pipeline to Status"""
    if deals_df.empty or 'Pipeline' not in deals_df.columns or 'Status' not in deals_df.columns:
        # Return empty figure if data not available
        return go.Figure()
    
    # Aggregate data: Pipeline -> Status
    df_agg = deals_df.groupby(['Pipeline', 'Status'])['Amount'].sum().reset_index()
    
    if df_agg.empty:
        return go.Figure()
    
    # Create source/target indices
    pipelines = list(df_agg['Pipeline'].unique())
    statuses = list(df_agg['Status'].unique())
    all_labels = pipelines + statuses
    
    source_indices = [pipelines.index(p) for p in df_agg['Pipeline']]
    target_indices = [len(pipelines) + statuses.index(s) for s in df_agg['Status']]
    
    # Create color map for statuses
    status_colors = {
        'Commit': 'rgba(16, 185, 129, 0.6)',
        'Expect': 'rgba(59, 130, 246, 0.6)',
        'Best Case': 'rgba(139, 92, 246, 0.6)',
        'Opp': 'rgba(251, 191, 36, 0.6)'
    }
    
    link_colors = []
    for idx in target_indices:
        status = statuses[idx - len(pipelines)]
        link_colors.append(status_colors.get(status, 'rgba(100, 116, 139, 0.4)'))
    
    fig = go.Figure(data=[go.Sankey(
        node = dict(
            pad = 15,
            thickness = 20,
            line = dict(color = "rgba(255,255,255,0.2)", width = 0.5),
            label = all_labels,
            color = "rgba(59, 130, 246, 0.8)"
        ),
        link = dict(
            source = source_indices,
            target = target_indices,
            value = df_agg['Amount'],
            color = link_colors
        )
    )])
    
    fig.update_layout(
        title_text="Pipeline Flow Analysis",
        font=dict(size=12, color='white'),
        height=500,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)'
    )
    return fig

def create_team_sunburst(dashboard_df, deals_df):
    """Enhanced sunburst chart with better colors and formatting"""
    # Prepare data structure for sunburst
    sunburst_data = []
    
    # Define distinct colors for each rep
    rep_colors = {
        'Brad Sherman': '#ef4444',      # Red
        'Jake Lynch': '#3b82f6',        # Blue
        'Dave Borkowski': '#10b981',    # Green
        'Lance Mitton': '#f59e0b',      # Amber
        'Alex Gonzalez': '#8b5cf6',     # Purple
        'Shopify ECommerce': '#ec4899'  # Pink
    }
    
    for _, rep_row in dashboard_df.iterrows():
        rep_name = rep_row['Rep Name']
        
        # Add invoiced with better label
        if 'NetSuite Orders' in rep_row and rep_row['NetSuite Orders'] > 0:
            sunburst_data.append({
                'labels': rep_name,
                'parents': '',
                'values': rep_row['NetSuite Orders'],
                'text': f"${rep_row['NetSuite Orders']:,.0f}",
                'type': 'Invoiced',
                'rep': rep_name
            })
            sunburst_data.append({
                'labels': f"{rep_name} - Invoiced",
                'parents': rep_name,
                'values': rep_row['NetSuite Orders'],
                'text': f"Invoiced: ${rep_row['NetSuite Orders']:,.0f}",
                'type': 'Invoiced',
                'rep': rep_name
            })
        
        # Add pipeline data if available
        rep_deals = deals_df[deals_df['Deal Owner'] == rep_name] if not deals_df.empty else pd.DataFrame()
        if not rep_deals.empty:
            pipeline_total = rep_deals['Amount'].sum()
            if pipeline_total > 0:
                sunburst_data.append({
                    'labels': f"{rep_name} - Pipeline",
                    'parents': rep_name,
                    'values': pipeline_total,
                    'text': f"Pipeline: ${pipeline_total:,.0f}",
                    'type': 'Pipeline',
                    'rep': rep_name
                })
    
    if not sunburst_data:
        return go.Figure()
    
    df_sunburst = pd.DataFrame(sunburst_data)
    
    # Create custom colors based on rep and type
    colors = []
    for _, row in df_sunburst.iterrows():
        rep = row['rep']
        base_color = rep_colors.get(rep, '#64748b')
        
        if row['type'] == 'Invoiced':
            # Darker shade for invoiced
            if base_color.startswith('#'):
                colors.append(base_color + 'dd')  # Add alpha
        else:
            # Lighter shade for pipeline
            if base_color.startswith('#'):
                colors.append(base_color + '88')  # More transparent
    
    fig = go.Figure(go.Sunburst(
        labels=df_sunburst['labels'],
        parents=df_sunburst['parents'],
        values=df_sunburst['values'],
        text=df_sunburst['text'],
        marker=dict(
            colors=colors,
            line=dict(color='rgba(255,255,255,0.3)', width=2)
        ),
        textfont=dict(size=14, color='white', family='Arial Black'),
        hovertemplate='<b>%{label}</b><br>%{text}<br>%{percentParent}<extra></extra>',
        branchvalues="total"
    ))
    
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white', 'size': 12},
        height=500,
        margin=dict(l=0, r=0, t=30, b=0)
    )
    return fig

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
        textposition='auto',
        textfont=dict(size=14)
    ))

    fig.add_trace(go.Bar(
        name='Expect/Commit',
        x=['Progress'],
        y=[metrics['expect_commit']],
        marker_color='#43A047',
        text=[f"${metrics['expect_commit']:,.0f}"],
        textposition='auto',
        textfont=dict(size=14)
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

# Helper function to safely grab a column by index
def get_col_by_index(df, index):
    """Safely grab a column by index with fallback"""
    if df is not None and not df.empty and len(df.columns) > index:
        return df.iloc[:, index]
    return pd.Series(dtype=object)

def create_enhanced_waterfall_chart(metrics, title, mode):
    """
    Creates a waterfall chart for forecast progress to address visibility issues with small segments.
    Each component gets its own visible bar height proportional to its value, making small segments readable.
    
    Args:
        metrics: dict with keys like 'orders', 'pending_fulfillment', etc., and 'total_quota', 'total_progress'
        title: Chart title
        mode: 'base' or 'full' to determine which components to include
    """
    # Define the steps based on mode
    if mode == "base":
        steps = [
            {'label': 'Invoiced', 'value': metrics['orders'], 'color': '#1E88E5'},
            {'label': 'Pending Fulfillment', 'value': metrics['pending_fulfillment'], 'color': '#FFC107'},
            {'label': 'Pending Approval', 'value': metrics['pending_approval'], 'color': '#FB8C00'},
            {'label': 'HubSpot Expect/Commit', 'value': metrics['expect_commit'], 'color': '#43A047'},
        ]
    elif mode == "full":
        steps = [
            {'label': 'Invoiced', 'value': metrics['orders'], 'color': '#1E88E5'},
            {'label': 'Pending Fulfillment', 'value': metrics['pending_fulfillment'], 'color': '#FFC107'},
            {'label': 'PF No Date', 'value': metrics.get('pending_fulfillment_no_date', 0), 'color': '#FFE082'},
            {'label': 'Pending Approval', 'value': metrics['pending_approval'], 'color': '#FB8C00'},
            {'label': 'PA No Date', 'value': metrics.get('pending_approval_no_date', 0), 'color': '#FFCC80'},
            {'label': 'Old PA (>2 weeks)', 'value': metrics.get('pending_approval_old', 0), 'color': '#FF9800'},
            {'label': 'HubSpot Expect/Commit', 'value': metrics['expect_commit'], 'color': '#43A047'},
        ]
    else:
        return None

    # Filter out zero-value steps to avoid clutter
    steps = [step for step in steps if step['value'] > 0]
    
    if not steps:
        return None

    # Calculate totals
    current_total = sum(step['value'] for step in steps)
    quota = metrics.get('total_quota', metrics.get('quota', 0))
    gap = quota - current_total
    
    # Create figure
    fig = go.Figure()
    
    # Add each component as a separate bar trace for full color control
    cumulative = 0
    for step in steps:
        fig.add_trace(go.Bar(
            name=step['label'],
            x=[step['label']],
            y=[step['value']],
            marker_color=step['color'],
            text=[f"${step['value']:,.0f}"],
            textposition='outside',
            textfont=dict(size=12),  # Remove fixed color to auto-adapt
            hovertemplate=f"<b>{step['label']}</b><br>${step['value']:,.0f}<br>Cumulative: ${cumulative + step['value']:,.0f}<extra></extra>",
            showlegend=True
        ))
        cumulative += step['value']
    
    # Add total bar showing cumulative sum
    fig.add_trace(go.Bar(
        name='TOTAL FORECAST',
        x=['TOTAL'],
        y=[current_total],
        marker_color='#7B1FA2',
        marker_line=dict(width=2, color='#4A148C'),
        text=[f"${current_total:,.0f}"],
        textposition='outside',
        textfont=dict(size=14, family='Arial Black'),  # Remove fixed color
        hovertemplate=f"<b>Total Forecast</b><br>${current_total:,.0f}<extra></extra>",
        showlegend=True
    ))
    
    # Add gap bar if exists
    if gap != 0:
        gap_color = '#DC3912' if gap > 0 else '#43A047'
        gap_label = 'Gap to Goal' if gap > 0 else 'Over Goal'
        fig.add_trace(go.Bar(
            name=gap_label,
            x=[gap_label],
            y=[abs(gap)],
            marker_color=gap_color,
            text=[f"${gap:,.0f}"],
            textposition='outside',
            textfont=dict(size=12),  # Remove fixed color
            hovertemplate=f"<b>{gap_label}</b><br>${gap:,.0f}<extra></extra>",
            showlegend=True
        ))
    
    # Add quota reference line
    fig.add_hline(
        y=quota,
        line_dash="dash",
        line_color="#DC3912",
        line_width=2,
        annotation_text=f"Quota Goal: ${quota:,.0f}",
        annotation_position="right"
    )
    
    # Add best case potential line if in base mode
    best_opp = metrics.get('best_opp', 0)
    if best_opp > 0:
        potential = current_total + best_opp
        fig.add_hline(
            y=potential,
            line_dash="dot",
            line_color="#FB8C00",
            line_width=2,
            annotation_text=f"Potential: ${potential:,.0f}",
            annotation_position="right"
        )
    
    # Customize layout
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=18)  # Remove fixed color to auto-adapt
        ),
        xaxis_title="Forecast Components",
        yaxis_title="Amount ($)",
        barmode='group',
        height=600,
        showlegend=True,
        legend=dict(
            orientation="v",
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.02,
            bgcolor="rgba(255,255,255,0.1)",  # Semi-transparent to work in both modes
            bordercolor="rgba(128,128,128,0.5)",
            borderwidth=1
        ),
        plot_bgcolor='rgba(0,0,0,0)',  # Transparent background
        paper_bgcolor='rgba(0,0,0,0)',  # Transparent paper
        font=dict(color=None),  # Auto font color
        yaxis=dict(
            gridcolor='rgba(128,128,128,0.2)',
            zeroline=True,
            zerolinecolor='rgba(128,128,128,0.5)',
            zerolinewidth=1
        ),
        xaxis=dict(
            tickangle=-45,
            automargin=True
        ),
        margin=dict(l=70, r=200, t=100, b=120),
        annotations=[
            dict(
                x=1.02,
                y=1.05,
                xref='paper',
                yref='paper',
                text=f"<b>Current Total:</b> ${current_total:,.0f}<br><b>Quota:</b> ${quota:,.0f}<br><b>Gap:</b> ${gap:,.0f}",
                showarrow=False,
                font=dict(size=13, color="black"),
                align="left",
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#333333",
                borderwidth=1,
                borderpad=8
            )
        ]
    )
    
    return fig

def create_status_breakdown_chart(deals_df, rep_name=None):
    """Create a pie chart showing deal distribution by status"""
    
    if rep_name:
        deals_df = deals_df[deals_df['Deal Owner'] == rep_name]
    
    # Only show Q4 deals (spreadsheet formula now handles PA date logic)
    deals_df = deals_df[deals_df.get('Q1 2026 Spillover') != 'Q1 2026']
    
    if deals_df.empty:
        return None
    
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
        title='Deal Amount by Forecast Category (Q4 Only)',
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
    
    # Only show Q4 deals (spreadsheet formula now handles PA date logic)
    deals_df = deals_df[deals_df.get('Q1 2026 Spillover') != 'Q1 2026']
    
    if deals_df.empty:
        return None
    
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
        title='Pipeline Breakdown by Forecast Category (Q4 Only)',
        color_discrete_map=color_map,
        text_auto='.2s',
        barmode='stack'
    )

    fig.update_traces(textfont_size=14, textposition='auto')

    fig.update_layout(
        height=450,
        yaxis_title="Amount ($)",
        xaxis_title="Pipeline",
        xaxis=dict(
            automargin=True,
            tickangle=-45
        ),
        yaxis=dict(automargin=True),
        margin=dict(l=50, r=50, t=80, b=100),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
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
    
    # Add Q4/Q1 indicator to color map
    timeline_df['Quarter'] = timeline_df.apply(
        lambda x: 'Q4 2025' if x.get('Q1 2026 Spillover') != 'Q1 2026' else 'Q1 2026', 
        axis=1
    )
    
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
        hover_data=['Deal Name', 'Amount', 'Pipeline', 'Quarter'],
        title='Deal Close Date Timeline',
        color_discrete_map=color_map
    )
    
    # Fixed: Use datetime object for the vertical line
    from datetime import datetime
    q4_boundary = datetime(2025, 12, 31)
    
    try:
        fig.add_vline(
            x=q4_boundary, 
            line_dash="dash", 
            line_color="red",
            annotation_text="Q4/Q1 Boundary"
        )
    except:
        pass
    
    fig.update_layout(
        height=400,
        yaxis_title="Deal Amount ($)",
        xaxis_title="Expected Close Date"
    )
    
    return fig

def create_invoice_status_chart(invoices_df, rep_name=None):
    """Create a chart showing invoice breakdown by status"""
    
    if invoices_df.empty:
        return None
    
    if rep_name:
        invoices_df = invoices_df[invoices_df['Sales Rep'] == rep_name]
    
    if invoices_df.empty:
        return None
    
    status_summary = invoices_df.groupby('Status')['Amount'].sum().reset_index()
    
    fig = px.pie(
        status_summary,
        values='Amount',
        names='Status',
        title='Invoice Amount by Status',
        hole=0.4
    )
    
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(height=400)
    
    return fig

def display_drill_down_section(title, amount, details_df, key_suffix):
    """Display a collapsible section with order details - WITH PROPER SO# AND LINKS"""
    
    item_count = len(details_df)
    with st.expander(f"{title}: ${amount:,.2f} (👀 Click to see {item_count} {'item' if item_count == 1 else 'items'})"):
        if not details_df.empty:
            # DEBUG: Check for duplicate columns
            if details_df.columns.duplicated().any():
                st.warning(f"⚠️ Duplicate columns detected: {details_df.columns[details_df.columns.duplicated()].tolist()}")
                # Remove duplicates
                details_df = details_df.loc[:, ~details_df.columns.duplicated()]
            
            try:
                # Determine data type and prepare display
                is_hubspot = 'Deal Name' in details_df.columns
                is_netsuite = 'Document Number' in details_df.columns or 'Internal ID' in details_df.columns
                
                # Create display dataframe
                display_df = pd.DataFrame()
                column_config = {}
                
                if is_hubspot and 'Record ID' in details_df.columns:
                    # HubSpot deals
                    display_df['🔗 Link'] = details_df['Record ID'].apply(
                        lambda x: f'https://app.hubspot.com/contacts/6712259/record/0-3/{x}/' if pd.notna(x) else ''
                    )
                    column_config['🔗 Link'] = st.column_config.LinkColumn(
                        "🔗 Link",
                        help="Click to view deal in HubSpot",
                        display_text="View Deal"
                    )
                    
                    # Add other HubSpot columns
                    if 'Deal Name' in details_df.columns:
                        display_df['Deal Name'] = details_df['Deal Name']
                    if 'Amount' in details_df.columns:
                        display_df['Amount'] = details_df['Amount'].apply(lambda x: f"${x:,.2f}")
                    if 'Status' in details_df.columns:
                        display_df['Status'] = details_df['Status']
                    if 'Pipeline' in details_df.columns:
                        display_df['Pipeline'] = details_df['Pipeline']
                    if 'Close Date' in details_df.columns:
                        if pd.api.types.is_datetime64_any_dtype(details_df['Close Date']):
                            display_df['Close Date'] = details_df['Close Date'].dt.strftime('%Y-%m-%d')
                        else:
                            display_df['Close Date'] = details_df['Close Date']
                    if 'Product Type' in details_df.columns:
                        display_df['Product Type'] = details_df['Product Type']
                
                elif is_netsuite:
                    # NetSuite sales orders - ALWAYS show Internal ID and create link if available
                    if 'Internal ID' in details_df.columns:
                        display_df['🔗 Link'] = details_df['Internal ID'].apply(
                            lambda x: f'https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}&whence=' if pd.notna(x) else ''
                        )
                        column_config['🔗 Link'] = st.column_config.LinkColumn(
                            "🔗 Link",
                            help="Click to view sales order in NetSuite",
                            display_text="View SO"
                        )
                        # Also show Internal ID as a regular column
                        display_df['Internal ID'] = details_df['Internal ID']
                    
                    # Add SO# (Document Number)
                    if 'Document Number' in details_df.columns:
                        display_df['SO#'] = details_df['Document Number']
                    
                    # Add other NetSuite columns
                    if 'Customer' in details_df.columns:
                        display_df['Customer'] = details_df['Customer']
                    if 'Amount' in details_df.columns:
                        display_df['Amount'] = details_df['Amount'].apply(lambda x: f"${x:,.2f}")
                    if 'Status' in details_df.columns:
                        display_df['Status'] = details_df['Status']
                    if 'Order Start Date' in details_df.columns:
                        if pd.api.types.is_datetime64_any_dtype(details_df['Order Start Date']):
                            display_df['Order Start Date'] = details_df['Order Start Date'].dt.strftime('%Y-%m-%d')
                        else:
                            display_df['Order Start Date'] = details_df['Order Start Date']
                    if 'Pending Approval Date' in details_df.columns:
                        if pd.api.types.is_datetime64_any_dtype(details_df['Pending Approval Date']):
                            display_df['Pending Approval Date'] = details_df['Pending Approval Date'].dt.strftime('%Y-%m-%d')
                        else:
                            display_df['Pending Approval Date'] = details_df['Pending Approval Date']
                    if 'Customer Promise Date' in details_df.columns:
                        if pd.api.types.is_datetime64_any_dtype(details_df['Customer Promise Date']):
                            display_df['Customer Promise Date'] = details_df['Customer Promise Date'].dt.strftime('%Y-%m-%d')
                        else:
                            display_df['Customer Promise Date'] = details_df['Customer Promise Date']
                    if 'Projected Date' in details_df.columns:
                        if pd.api.types.is_datetime64_any_dtype(details_df['Projected Date']):
                            display_df['Projected Date'] = details_df['Projected Date'].dt.strftime('%Y-%m-%d')
                        else:
                            display_df['Projected Date'] = details_df['Projected Date']
                
                # Display the dataframe
                if not display_df.empty:
                    st.dataframe(
                        display_df, 
                        use_container_width=True, 
                        hide_index=True,
                        column_config=column_config if column_config else None
                    )
                    
                    # Summary statistics
                    st.caption(f"Total: ${details_df['Amount'].sum():,.2f} | Count: {len(details_df)} items")
                else:
                    # Fallback - show available columns for debugging
                    st.warning(f"Could not format data. Available columns: {details_df.columns.tolist()}")
                    st.dataframe(details_df, use_container_width=True, hide_index=True)
                    
            except Exception as e:
                st.error(f"Error displaying data: {str(e)}")
                st.write(f"Available columns: {details_df.columns.tolist()}")
                # Show raw data as fallback
                st.dataframe(details_df.head(), use_container_width=True, hide_index=True)
        else:
            st.info("📭 Nothing to see here... yet!")

def display_progress_breakdown(metrics):
    """Display a beautiful progress breakdown card"""
    
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>💰 Section 1: The Money We Can Count On</h3>
        <div class="progress-item">
            <span class="progress-label">✅ Already Celebrating (Invoiced & Shipped)</span>
            <span class="progress-value">${metrics['orders']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">📦 In the Warehouse (Just Add Shipping Label)</span>
            <span class="progress-value">${metrics['pending_fulfillment']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">⏳ Waiting for the Magic Signature</span>
            <span class="progress-value">${metrics['pending_approval']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 Deals We're Banking On (HubSpot Expect/Commit)</span>
            <span class="progress-value">${metrics['expect_commit']:,.0f}</span>
        </div>
        <div class="progress-item">
            <span class="progress-label">🎯 THE SAFE BET TOTAL</span>
            <span class="progress-value">${metrics['total_progress']:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Add attainment info below
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Current Attainment", f"{metrics['attainment_pct']:.1f}%", 
                 delta=f"${metrics['total_progress']:,.0f} of ${metrics['quota']:,.0f}",
                 help="This is real money! 💵")
    with col2:
        st.metric("If Everything Goes Right", f"{metrics['potential_attainment']:.1f}%",
                 delta=f"+${metrics['best_opp']:,.0f} Best Case/Opp",
                 help="The optimist's view (we believe! 🌟)")

def display_reconciliation_view(deals_df, dashboard_df, sales_orders_df):
    """Show a reconciliation view to compare with boss's numbers"""
    
    st.title("🔍 Forecast Reconciliation with Boss's Numbers")
    
    # Boss's Q4 numbers from the LATEST screenshot (November 21, 2025)
    boss_rep_numbers = {
        'Jake Lynch': {
            'invoiced': 799840,
            'pending_fulfillment': 303737,
            'pending_approval': 97427,
            'hubspot': 62021,
            'total': 1263025,
            'pending_fulfillment_so_no_date': 101694,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 39174,
            'total_q4': 1263025,
            'hubspot_best_case': 413594,
            'jan_expect_commit': 0,
            'jan_best_case': 325245
        },
        'Dave Borkowski': {
            'invoiced': 395285,
            'pending_fulfillment': 200071,
            'pending_approval': 30931,
            'hubspot': 273179,
            'total': 899467,
            'pending_fulfillment_so_no_date': 68908,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 37927,
            'total_q4': 899467,
            'hubspot_best_case': 119160,
            'jan_expect_commit': 0,
            'jan_best_case': 113005
        },
        'Alex Gonzalez': {
            'invoiced': 432791,
            'pending_fulfillment': 323333,
            'pending_approval': 516,
            'hubspot': 5496,
            'total': 762136,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 25020,
            'old_pending_approval': 4900,
            'total_q4': 762136,
            'hubspot_best_case': 0,
            'jan_expect_commit': 0,
            'jan_best_case': 0
        },
        'Brad Sherman': {
            'invoiced': 194683,
            'pending_fulfillment': 39127,
            'pending_approval': 11636,
            'hubspot': 90616,
            'total': 336062,
            'pending_fulfillment_so_no_date': 4217,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 4553,
            'total_q4': 336062,
            'hubspot_best_case': 61523,
            'jan_expect_commit': 0,
            'jan_best_case': 90585
        },
        'Lance Mitton': {
            'invoiced': 28502,
            'pending_fulfillment': 1287,
            'pending_approval': 1069,
            'hubspot': 0,
            'total': 30857,
            'pending_fulfillment_so_no_date': 1914,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 0,
            'total_q4': 30857,
            'hubspot_best_case': 0,
            'jan_expect_commit': 5700,
            'jan_best_case': 8700
        },
        'House': {
            'invoiced': 0,
            'pending_fulfillment': 0,
            'pending_approval': 0,
            'hubspot': 0,
            'total': 0,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 0,
            'total_q4': 0,
            'hubspot_best_case': 0,
            'jan_expect_commit': 0,
            'jan_best_case': 0
        },
        'Shopify ECommerce': {
            'invoiced': 30621,
            'pending_fulfillment': 3374,
            'pending_approval': 171,
            'hubspot': 0,
            'total': 34166,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 2718,
            'total_q4': 34166,
            'hubspot_best_case': 0,
            'jan_expect_commit': 0,
            'jan_best_case': 1500
        }
    }
    
    # Tab selection for Rep vs Pipeline view
    tab1, tab2 = st.tabs(["By Rep", "By Pipeline"])
    
    with tab1:
        st.markdown('<div class="section-header">Section 1: Q4 Gap to Goal</div>', unsafe_allow_html=True)
        
        # Create the comparison table
        comparison_data = []
        totals = {
            'invoiced_you': 0, 'invoiced_boss': 0,
            'pf_you': 0, 'pf_boss': 0,
            'pa_you': 0, 'pa_boss': 0,
            'hs_you': 0, 'hs_boss': 0,
            'total_you': 0, 'total_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            metrics = None
            if rep_name in dashboard_df['Rep Name'].values:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics or rep_name == 'Shopify ECommerce':
                boss = boss_rep_numbers[rep_name]
                
                # Get your values
                your_invoiced = metrics['orders'] if metrics else 0
                your_pf = metrics['pending_fulfillment'] if metrics else 0
                your_pa = metrics['pending_approval'] if metrics else 0
                your_hs = metrics['expect_commit'] if metrics else 0
                your_total = your_invoiced + your_pf + your_pa + your_hs
                
                # Update totals
                totals['invoiced_you'] += your_invoiced
                totals['invoiced_boss'] += boss['invoiced']
                totals['pf_you'] += your_pf
                totals['pf_boss'] += boss['pending_fulfillment']
                totals['pa_you'] += your_pa
                totals['pa_boss'] += boss['pending_approval']
                totals['hs_you'] += your_hs
                totals['hs_boss'] += boss['hubspot']
                totals['total_you'] += your_total
                totals['total_boss'] += boss['total']
                
                comparison_data.append({
                    'Rep': rep_name,
                    'Invoiced': f"${your_invoiced:,.0f}",
                    'Invoiced (Boss)': f"${boss['invoiced']:,.0f}",
                    'Pending Fulfillment': f"${your_pf:,.0f}",
                    'PF (Boss)': f"${boss['pending_fulfillment']:,.0f}",
                    'Pending Approval': f"${your_pa:,.0f}",
                    'PA (Boss)': f"${boss['pending_approval']:,.0f}",
                    'HubSpot Expect/Commit': f"${your_hs:,.0f}",
                    'HS (Boss)': f"${boss['hubspot']:,.0f}",
                    'Total': f"${your_total:,.0f}",
                    'Total (Boss)': f"${boss['total']:,.0f}"
                })
        
        # Add totals row
        comparison_data.append({
            'Rep': 'TOTAL',
            'Invoiced': f"${totals['invoiced_you']:,.0f}",
            'Invoiced (Boss)': f"${totals['invoiced_boss']:,.0f}",
            'Pending Fulfillment': f"${totals['pf_you']:,.0f}",
            'PF (Boss)': f"${totals['pf_boss']:,.0f}",
            'Pending Approval': f"${totals['pa_you']:,.0f}",
            'PA (Boss)': f"${totals['pa_boss']:,.0f}",
            'HubSpot Expect/Commit': f"${totals['hs_you']:,.0f}",
            'HS (Boss)': f"${totals['hs_boss']:,.0f}",
            'Total': f"${totals['total_you']:,.0f}",
            'Total (Boss)': f"${totals['total_boss']:,.0f}"
        })
        
        if comparison_data:
            comparison_df = pd.DataFrame(comparison_data)
            st.dataframe(comparison_df, use_container_width=True, hide_index=True)
        
        # Section 2: Additional Orders
        st.markdown('<div class="section-header">Section 2: Additional Orders (Can be included)</div>', unsafe_allow_html=True)
        
        additional_data = []
        additional_totals = {
            'pf_no_date_you': 0, 'pf_no_date_boss': 0,
            'pa_no_date_you': 0, 'pa_no_date_boss': 0,
            'old_pa_you': 0, 'old_pa_boss': 0,
            'final_you': 0, 'final_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            metrics = None
            if rep_name in dashboard_df['Rep Name'].values:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics or rep_name == 'Shopify ECommerce':
                boss = boss_rep_numbers[rep_name]
                
                # Calculate additional metrics
                your_pf_no_date = metrics['pending_fulfillment_no_date'] if metrics else 0
                your_pa_no_date = metrics['pending_approval_no_date'] if metrics else 0
                your_old_pa = metrics['pending_approval_old'] if metrics else 0
                
                # Calculate final total
                section1_total = (metrics['orders'] + metrics['pending_fulfillment'] + 
                                 metrics['pending_approval'] + metrics['expect_commit']) if metrics else 0
                your_final = section1_total + your_pf_no_date + your_pa_no_date + your_old_pa
                
                # Update totals
                additional_totals['pf_no_date_you'] += your_pf_no_date
                additional_totals['pf_no_date_boss'] += boss['pending_fulfillment_so_no_date']
                additional_totals['pa_no_date_you'] += your_pa_no_date
                additional_totals['pa_no_date_boss'] += boss['pending_approval_so_no_date']
                additional_totals['old_pa_you'] += your_old_pa
                additional_totals['old_pa_boss'] += boss['old_pending_approval']
                additional_totals['final_you'] += your_final
                additional_totals['final_boss'] += boss['total_q4']
                
                additional_data.append({
                    'Rep': rep_name,
                    'PF SO\'s No Date': f"${your_pf_no_date:,.0f}",
                    'PF No Date (Boss)': f"${boss['pending_fulfillment_so_no_date']:,.0f}",
                    'PA SO\'s No Date': f"${your_pa_no_date:,.0f}",
                    'PA No Date (Boss)': f"${boss['pending_approval_so_no_date']:,.0f}",
                    'Old PA (>2 weeks)': f"${your_old_pa:,.0f}",
                    'Old PA (Boss)': f"${boss['old_pending_approval']:,.0f}",
                    'Total Q4': f"${your_final:,.0f}",
                    'Total Q4 (Boss)': f"${boss['total_q4']:,.0f}"
                })
        
        # Add totals row
        additional_data.append({
            'Rep': 'TOTAL',
            'PF SO\'s No Date': f"${additional_totals['pf_no_date_you']:,.0f}",
            'PF No Date (Boss)': f"${additional_totals['pf_no_date_boss']:,.0f}",
            'PA SO\'s No Date': f"${additional_totals['pa_no_date_you']:,.0f}",
            'PA No Date (Boss)': f"${additional_totals['pa_no_date_boss']:,.0f}",
            'Old PA (>2 weeks)': f"${additional_totals['old_pa_you']:,.0f}",
            'Old PA (Boss)': f"${additional_totals['old_pa_boss']:,.0f}",
            'Total Q4': f"${additional_totals['final_you']:,.0f}",
            'Total Q4 (Boss)': f"${additional_totals['final_boss']:,.0f}"
        })
        
        if additional_data:
            additional_df = pd.DataFrame(additional_data)
            st.dataframe(additional_df, use_container_width=True, hide_index=True)
        
        # Section 3: Q1 2026 Spillover
        st.markdown('<div class="section-header">Section 3: Q1 2026 Spillover (January)</div>', unsafe_allow_html=True)
        st.info("These are deals that will close in late Q4 2025 but ship in Q1 2026 due to lead times")
        
        q1_spillover_data = []
        q1_totals = {
            'expect_commit_you': 0, 'expect_commit_boss': 0,
            'best_case_you': 0, 'best_case_boss': 0,
            'total_you': 0, 'total_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            metrics = None
            if rep_name in dashboard_df['Rep Name'].values:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
            
            if metrics or rep_name == 'Shopify ECommerce':
                boss = boss_rep_numbers[rep_name]
                
                # Get Q1 spillover values - use the correct field names from calculate_rep_metrics
                your_expect_commit = metrics.get('q1_spillover_expect_commit', 0) if metrics else 0
                your_best_case = metrics.get('q1_spillover_best_opp', 0) if metrics else 0
                your_q1_total = your_expect_commit + your_best_case
                
                boss_q1_total = boss['jan_expect_commit'] + boss['jan_best_case']
                
                # Update totals
                q1_totals['expect_commit_you'] += your_expect_commit
                q1_totals['expect_commit_boss'] += boss['jan_expect_commit']
                q1_totals['best_case_you'] += your_best_case
                q1_totals['best_case_boss'] += boss['jan_best_case']
                q1_totals['total_you'] += your_q1_total
                q1_totals['total_boss'] += boss_q1_total
                
                q1_spillover_data.append({
                    'Rep': rep_name,
                    'Expect/Commit': f"${your_expect_commit:,.0f}",
                    'Expect/Commit (Boss)': f"${boss['jan_expect_commit']:,.0f}",
                    'Best Case': f"${your_best_case:,.0f}",
                    'Best Case (Boss)': f"${boss['jan_best_case']:,.0f}",
                    'Total Q1 Spillover': f"${your_q1_total:,.0f}",
                    'Total Q1 (Boss)': f"${boss_q1_total:,.0f}"
                })
        
        # Add totals row
        q1_spillover_data.append({
            'Rep': 'TOTAL',
            'Expect/Commit': f"${q1_totals['expect_commit_you']:,.0f}",
            'Expect/Commit (Boss)': f"${q1_totals['expect_commit_boss']:,.0f}",
            'Best Case': f"${q1_totals['best_case_you']:,.0f}",
            'Best Case (Boss)': f"${q1_totals['best_case_boss']:,.0f}",
            'Total Q1 Spillover': f"${q1_totals['total_you']:,.0f}",
            'Total Q1 (Boss)': f"${q1_totals['total_boss']:,.0f}"
        })
        
        if q1_spillover_data:
            q1_spillover_df = pd.DataFrame(q1_spillover_data)
            st.dataframe(q1_spillover_df, use_container_width=True, hide_index=True)
    
    with tab2:
        st.markdown("### Pipeline-Level Comparison")
        st.info("Pipeline breakdown in development - need to map invoices and sales orders to pipelines")
    
    # Summary
    st.markdown("### 📊 Key Insights")
    
    # Calculate differences first
    diff = totals['total_boss'] - totals['total_you']
    final_diff = additional_totals['final_boss'] - additional_totals['final_you']
    q1_diff = q1_totals['total_boss'] - q1_totals['total_you']
    
    # Debug: Show the actual totals being compared
    st.caption(f"Debug: Your Total Q4 = ${additional_totals['final_you']:,.0f} | Boss Total Q4 = ${additional_totals['final_boss']:,.0f} | Diff = ${abs(final_diff):,.0f}")
    st.caption(f"Debug: Your Q1 Spillover = ${q1_totals['total_you']:,.0f} | Boss Q1 Spillover = ${q1_totals['total_boss']:,.0f} | Diff = ${abs(q1_diff):,.0f}")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Section 1 Variance", f"${abs(diff):,.0f}", 
                 delta=f"{'Under' if diff > 0 else 'Over'} by ${abs(diff):,.0f}")
    
    with col2:
        st.metric("Total Q4 Variance", f"${abs(final_diff):,.0f}",
                 delta=f"{'Under' if final_diff > 0 else 'Over'} by ${abs(final_diff):,.0f}")
    
    with col3:
        st.metric("Q1 Spillover Variance", f"${abs(q1_diff):,.0f}",
                 delta=f"{'Under' if q1_diff > 0 else 'Over'} by ${abs(q1_diff):,.0f}")
    
    with col4:
        if additional_totals['final_boss'] > 0:
            # Calculate accuracy as: 100% - (percentage difference)
            accuracy = (1 - abs(final_diff) / additional_totals['final_boss']) * 100
            
            # Show color coding
            if accuracy >= 95:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Excellent match")
            elif accuracy >= 90:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Good match", delta_color="normal")
            elif accuracy >= 80:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Needs review", delta_color="off")
            else:
                st.metric("Accuracy", f"{accuracy:.1f}%", delta="Large variance", delta_color="inverse")
        else:
            st.metric("Accuracy", "N/A", delta="No boss data")

def display_team_dashboard(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display the team-level dashboard"""
   
    st.title("🎯 Team Sales Dashboard - Q4 2025")
   
    # Calculate basic metrics
    basic_metrics = calculate_team_metrics(deals_df, dashboard_df)
   
    # Aggregate full team metrics from per-rep calculations
    team_quota = basic_metrics['total_quota']
    team_best_opp = basic_metrics['best_opp']
    team_q1_spillover_expect_commit = 0  # Q1 spillover Expect/Commit only
    team_q1_spillover_best_opp = 0  # Q1 spillover Best Case/Opportunity
    
    # Filter out unwanted reps
    excluded_reps = ['House', 'house', 'HOUSE']
   
    team_invoiced = 0
    team_pf = 0
    team_pa = 0
    team_hs = 0
    team_pf_no_date = 0
    team_pa_no_date = 0
    team_old_pa = 0
    
    # FIX: Calculate team_invoiced directly from invoices_df to match Invoice Detail section
    if not invoices_df.empty and 'Amount' in invoices_df.columns:
        # Filter out House reps if needed
        if 'Sales Rep' in invoices_df.columns:
            filtered_inv = invoices_df[~invoices_df['Sales Rep'].isin(excluded_reps)].copy()
        else:
            filtered_inv = invoices_df.copy()
        
        # Calculate total invoiced amount
        filtered_inv['Amount_Numeric'] = pd.to_numeric(filtered_inv['Amount'], errors='coerce')
        team_invoiced = filtered_inv['Amount_Numeric'].sum()
    else:
        team_invoiced = 0
   
    section1_data = []
    section2_data = []
    
    for rep_name in dashboard_df['Rep Name']:
        # Skip excluded reps
        if rep_name in excluded_reps:
            continue
            
        rep_metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
        if rep_metrics:
            section1_total = (rep_metrics['orders'] + rep_metrics['pending_fulfillment'] +
                              rep_metrics['pending_approval'] + rep_metrics['expect_commit'])
            final_total = (section1_total + rep_metrics['pending_fulfillment_no_date'] +
                           rep_metrics['pending_approval_no_date'] + rep_metrics['pending_approval_old'])
           
            section1_data.append({
                'Rep': rep_name,
                'Invoiced': f"${rep_metrics['orders']:,.0f}",
                'Pending Fulfillment': f"${rep_metrics['pending_fulfillment']:,.0f}",
                'Pending Approval': f"${rep_metrics['pending_approval']:,.0f}",
                'HubSpot Expect/Commit': f"${rep_metrics['expect_commit']:,.0f}",
                'Total': f"${section1_total:,.0f}"
            })
           
            section2_data.append({
                'Rep': rep_name,
                'PF SO\'s No Date': f"${rep_metrics['pending_fulfillment_no_date']:,.0f}",
                'PA SO\'s No Date': f"${rep_metrics['pending_approval_no_date']:,.0f}",
                'Old PA (>2 weeks)': f"${rep_metrics['pending_approval_old']:,.0f}",
                'Total Q4': f"${final_total:,.0f}"
            })
           
            # Aggregate sums (exclude invoiced since we calculate it directly from invoices_df now)
            # team_invoiced += rep_metrics['orders']  # REMOVED - now calculated from raw invoices_df
            team_pf += rep_metrics['pending_fulfillment']
            team_pa += rep_metrics['pending_approval']
            team_hs += rep_metrics['expect_commit']
            team_pf_no_date += rep_metrics['pending_fulfillment_no_date']
            team_pa_no_date += rep_metrics['pending_approval_no_date']
            team_old_pa += rep_metrics['pending_approval_old']
            team_q1_spillover_expect_commit += rep_metrics.get('q1_spillover_expect_commit', 0)
            team_q1_spillover_best_opp += rep_metrics.get('q1_spillover_best_opp', 0)
   
    # Calculate team totals
    base_forecast = team_invoiced + team_pf + team_pa + team_hs
    full_forecast = base_forecast + team_pf_no_date + team_pa_no_date + team_old_pa
    base_gap = team_quota - base_forecast
    full_gap = team_quota - full_forecast
    base_attainment_pct = (base_forecast / team_quota * 100) if team_quota > 0 else 0
    full_attainment_pct = (full_forecast / team_quota * 100) if team_quota > 0 else 0
    potential_attainment = ((base_forecast + team_best_opp) / team_quota * 100) if team_quota > 0 else 0
    
    # NEW: Calculate Best Case only (not Opportunity) for optimistic gap
    deals_q4 = deals_df[deals_df.get('Q1 2026 Spillover') != 'Q1 2026'] if not deals_df.empty else pd.DataFrame()
    team_best_case = deals_q4[deals_q4['Status'] == 'Best Case']['Amount'].sum() if not deals_q4.empty and 'Status' in deals_q4.columns else 0
    
    # NEW: Optimistic Gap = Quota - (High Confidence + Best Case + PF no date + PA no date + PA >2 weeks)
    optimistic_forecast = base_forecast + team_best_case + team_pf_no_date + team_pa_no_date + team_old_pa
    optimistic_gap = team_quota - optimistic_forecast
   
    # Add total rows to data
    section1_data.append({
        'Rep': 'TOTAL',
        'Invoiced': f"${team_invoiced:,.0f}",
        'Pending Fulfillment': f"${team_pf:,.0f}",
        'Pending Approval': f"${team_pa:,.0f}",
        'HubSpot Expect/Commit': f"${team_hs:,.0f}",
        'Total': f"${base_forecast:,.0f}"
    })
   
    section2_data.append({
        'Rep': 'TOTAL',
        'PF SO\'s No Date': f"${team_pf_no_date:,.0f}",
        'PA SO\'s No Date': f"${team_pa_no_date:,.0f}",
        'Old PA (>2 weeks)': f"${team_old_pa:,.0f}",
        'Total Q4': f"${full_forecast:,.0f}"
    })
   
    # Display Q1 spillover info if applicable
    team_q1_spillover_total = team_q1_spillover_expect_commit + team_q1_spillover_best_opp
    if team_q1_spillover_total > 0:
        st.info(
            f"ℹ️ **Q1 2026 Spillover**: ${team_q1_spillover_total:,.0f} in deals closing late Q4 2025 "
            f"will ship in Q1 2026 due to product lead times. These are excluded from Q4 revenue recognition."
        )
   
    # Display key metrics with two breakdowns
    st.markdown("### 📊 Team Scorecard")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
   
    with col1:
        st.metric(
            label="🎯 Total Quota",
            value=f"${team_quota/1000:.0f}K" if team_quota < 1000000 else f"${team_quota/1000000:.1f}M",
            delta=None,
            help="Q4 2025 Sales Target"
        )
   
    with col2:
        st.metric(
            label="💪 High Confidence Forecast",
            value=f"${base_forecast/1000:.0f}K" if base_forecast < 1000000 else f"${base_forecast/1000000:.1f}M",
            delta=f"{base_attainment_pct:.1f}% of quota",
            help="Invoiced + PF (with date) + PA (with date) + HS Expect/Commit"
        )
   
    with col3:
        st.metric(
            label="📊 Full Forecast (All Sources)",
            value=f"${full_forecast/1000:.0f}K" if full_forecast < 1000000 else f"${full_forecast/1000000:.1f}M",
            delta=f"{full_attainment_pct:.1f}% of quota",
            help="Invoiced + PF (with date) + PA (with date) + HS Expect/Commit + PF (without date) + PA (without date) + PA (>2 weeks old)"
        )
    
    with col4:
        st.metric(
            label="📉 Gap to Quota",
            value=f"${base_gap/1000:.0f}K" if abs(base_gap) < 1000000 else f"${base_gap/1000000:.1f}M",
            delta=f"${-base_gap/1000:.0f}K" if base_gap < 0 else None,
            delta_color="inverse",
            help="Quota - (Invoiced + PF (with date) + PA (with date) + HS Expect/Commit)"
        )
    
    with col5:
        st.metric(
            label="📈 Optimistic Gap",
            value=f"${optimistic_gap/1000:.0f}K" if abs(optimistic_gap) < 1000000 else f"${optimistic_gap/1000000:.1f}M",
            delta=f"${-optimistic_gap/1000:.0f}K" if optimistic_gap < 0 else None,
            delta_color="inverse",
            help="Quota - (High Confidence + HS Best Case + PF (no date) + PA (no date) + PA >2 weeks)"
        )

    with col6:
        st.metric(
            label="🌟 Potential Attainment",
            value=f"{potential_attainment:.1f}%",
            delta=f"+{potential_attainment - base_attainment_pct:.1f}% upside",
            help="(High Confidence + HS Best Case/Opp) ÷ Quota"
        )
    
    # Add expandable breakdown details
    with st.expander("🔍 View Calculation Breakdowns", expanded=False):
        st.markdown("#### Detailed Component Breakdown")
        
        breakdown_col1, breakdown_col2 = st.columns(2)
        
        with breakdown_col1:
            st.markdown("##### 💪 High Confidence Forecast")
            st.markdown(f"""
            - **Invoiced:** ${team_invoiced:,.0f}
            - **Pending Fulfillment (with date):** ${team_pf:,.0f}
            - **Pending Approval (with date):** ${team_pa:,.0f}
            - **HubSpot Expect/Commit:** ${team_hs:,.0f}
            - **Total:** ${base_forecast:,.0f}
            - **% of Quota:** {base_attainment_pct:.1f}%
            """)
        
        with breakdown_col2:
            st.markdown("##### 📊 Full Forecast (All Sources)")
            st.markdown(f"""
            **High Confidence Components:**
            - Invoiced: ${team_invoiced:,.0f}
            - PF (with date): ${team_pf:,.0f}
            - PA (with date): ${team_pa:,.0f}
            - HS Expect/Commit: ${team_hs:,.0f}
            
            **Additional Sources:**
            - PF (no date): ${team_pf_no_date:,.0f}
            - PA (no date): ${team_pa_no_date:,.0f}
            - PA (>2 weeks): ${team_old_pa:,.0f}
            
            **Total:** ${full_forecast:,.0f}
            - **% of Quota:** {full_attainment_pct:.1f}%
            """)
        
        st.markdown("---")
        
        breakdown_col3, breakdown_col4 = st.columns(2)
        
        with breakdown_col3:
            st.markdown("##### 📈 Optimistic Scenario")
            st.markdown(f"""
            - **High Confidence:** ${base_forecast:,.0f}
            - **Plus: HS Best Case:** ${team_best_case:,.0f}
            - **Plus: PF (no date):** ${team_pf_no_date:,.0f}
            - **Plus: PA (no date):** ${team_pa_no_date:,.0f}
            - **Plus: PA (>2 weeks):** ${team_old_pa:,.0f}
            - **Optimistic Total:** ${optimistic_forecast:,.0f}
            - **Gap to Quota:** ${optimistic_gap:,.0f}
            """)
   
    # Invoices section and audit section
    st.markdown("---")
    
    # ========== GEMINI ENHANCEMENTS: Advanced Visualizations ==========
    st.markdown("### 📊 Advanced Team Analytics")
    
    viz_col1, viz_col2 = st.columns(2)
    
    with viz_col1:
        st.markdown("#### Pipeline Flow (Sankey)")
        try:
            if not deals_df.empty and 'Pipeline' in deals_df.columns and 'Status' in deals_df.columns:
                sankey_fig = create_pipeline_sankey(deals_df)
                if sankey_fig and hasattr(sankey_fig, 'data') and len(sankey_fig.data) > 0:
                    st.plotly_chart(sankey_fig, use_container_width=True, key="sankey_chart")
                else:
                    st.info("📭 No pipeline data available for Sankey diagram")
            else:
                st.info("📭 Missing required columns for Sankey diagram")
        except Exception as e:
            st.error(f"Error generating Sankey: {str(e)}")
            st.caption("Debug: Check that Pipeline and Status columns exist in deals data")
    
    with viz_col2:
        st.markdown("#### Team Breakdown (Sunburst)")
        try:
            if not dashboard_df.empty:
                sunburst_fig = create_team_sunburst(dashboard_df, deals_df)
                if sunburst_fig and hasattr(sunburst_fig, 'data') and len(sunburst_fig.data) > 0:
                    st.plotly_chart(sunburst_fig, use_container_width=True, key="sunburst_chart")
                else:
                    st.info("📭 No data available for Sunburst chart")
            else:
                st.info("📭 No team data available")
        except Exception as e:
            st.error(f"Error generating Sunburst: {str(e)}")
            st.caption("Debug: Check dashboard and deals data structure")
    
    st.markdown("---")
    
    # Change detection and audit section
    if st.checkbox("📊 Show Day-Over-Day Audit", value=False):
        create_dod_audit_section(deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    st.markdown("---")
    
    # Invoices section
    display_invoices_drill_down(invoices_df)
    
    st.markdown("---")
    
    # Build Your Own Forecast section
    team_metrics_for_forecast = {
        'orders': team_invoiced,
        'pending_fulfillment': team_pf,
        'pending_approval': team_pa,
        'expect_commit': team_hs,
        'pending_fulfillment_no_date': team_pf_no_date,
        'pending_approval_no_date': team_pa_no_date,
        'pending_approval_old': team_old_pa,
        'q1_spillover_expect_commit': team_q1_spillover_expect_commit,
        'q1_spillover_best_opp': team_q1_spillover_best_opp
    }
    build_your_own_forecast_section(
        team_metrics_for_forecast,
        team_quota,
        rep_name=None,
        deals_df=deals_df,
        invoices_df=invoices_df,
        sales_orders_df=sales_orders_df
    )
    
    st.markdown("---")
    
    # HubSpot Deals Audit Section
    display_hubspot_deals_audit(deals_df)
    
    st.markdown("---")
    
    # Progress bars for both breakdowns
    st.markdown("### 📈 Progress to Quota")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**💪 High Confidence Forecast Progress**")
        st.caption("Confirmed orders and forecast with dates")
        base_progress = min(base_attainment_pct / 100, 1.0)
        st.progress(base_progress)
        st.caption(f"Current: {base_attainment_pct:.1f}% | Potential: {potential_attainment:.1f}%")
   
    with col2:
        st.markdown("**📊 Full Forecast Progress**")
        st.caption("All sources including orders without dates")
        full_progress = min(full_attainment_pct / 100, 1.0)
        st.progress(full_progress)
        st.caption(f"Current: {full_attainment_pct:.1f}%")
   
    # Base Forecast Chart with Enhanced Annotations
    st.markdown("### 💪 High Confidence Forecast Breakdown")
    st.caption("Orders and deals with confirmed dates and high confidence")
    
    # Create metrics dict for base chart
    base_metrics = {
        'orders': team_invoiced,
        'pending_fulfillment': team_pf,
        'pending_approval': team_pa,
        'expect_commit': team_hs,
        'best_opp': team_best_opp,
        'total_progress': base_forecast,
        'total_quota': team_quota
    }
    
    base_chart = create_enhanced_waterfall_chart(base_metrics, "💪 High Confidence Forecast - Path to Quota", "base")
    st.plotly_chart(base_chart, use_container_width=True)

    # Full Forecast Chart with Enhanced Annotations
    st.markdown("### 📊 Full Forecast Breakdown")
    st.caption("Complete view including all orders and pending items")
    
    full_metrics = {
        'orders': team_invoiced,
        'pending_fulfillment': team_pf,
        'pending_fulfillment_no_date': team_pf_no_date,
        'pending_approval': team_pa,
        'pending_approval_no_date': team_pa_no_date,
        'pending_approval_old': team_old_pa,
        'expect_commit': team_hs,
        'best_opp': team_best_opp,
        'total_progress': base_forecast,
        'total_quota': team_quota
    }
    
    full_chart = create_enhanced_waterfall_chart(full_metrics, "📊 Full Forecast - All Sources Included", "full")
    st.plotly_chart(full_chart, use_container_width=True)

    # Other charts remain the same
    col1, col2 = st.columns(2)
   
    with col1:
        st.markdown("#### 🎯 Deal Confidence Levels")
        status_chart = create_status_breakdown_chart(deals_df)
        if status_chart:
            st.plotly_chart(status_chart, use_container_width=True)
        else:
            st.info("📭 Nothing to see here... yet!")
   
    with col2:
        st.markdown("#### 🔮 The Crystal Ball: Where Our Deals Stand")
        pipeline_chart = create_pipeline_breakdown_chart(deals_df)
        if pipeline_chart:
            st.plotly_chart(pipeline_chart, use_container_width=True)
        else:
            st.info("📭 Nothing to see here... yet!")
   
    st.markdown("### 📅 When the Magic Happens (Expected Close Dates)")
    timeline_chart = create_deals_timeline(deals_df)
    if timeline_chart:
        st.plotly_chart(timeline_chart, use_container_width=True)
    else:
        st.info("📭 Nothing to see here... yet!")
   
    if not invoices_df.empty:
        st.markdown("### 💰 Invoice Status (Show Me the Money!)")
        invoice_chart = create_invoice_status_chart(invoices_df)
        if invoice_chart:
            st.plotly_chart(invoice_chart, use_container_width=True)
   
    # Display the two sections
    st.markdown("### 👥 High Confidence Forecast by Rep")
    st.caption("Invoiced + Pending Fulfillment (with date) + Pending Approval (with date) + HubSpot Expect/Commit")
    if section1_data:
        section1_df = pd.DataFrame(section1_data)
        st.dataframe(section1_df, use_container_width=True, hide_index=True)
    else:
        st.warning("📭 No data for High Confidence Forecast")
   
    st.markdown("### 👥 Additional Forecast Items by Rep")
    st.caption("Section 1 (above) + items below = Total Q4. Items below: Pending Fulfillment (without date) + Pending Approval (without date) + Old Pending Approval (>2 weeks)")
    if section2_data:
        section2_df = pd.DataFrame(section2_data)
        st.dataframe(section2_df, use_container_width=True, hide_index=True)
    else:
        st.warning("📭 No additional forecast items")
def display_rep_dashboard(rep_name, deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Display individual rep dashboard with drill-down capability - REDESIGNED"""
    
    st.title(f"👤 {rep_name}'s Q4 2025 Forecast")
    
    # Calculate metrics with details
    metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df)
    
    if not metrics:
        st.error(f"No data found for {rep_name}")
        return
    
    # Calculate the key forecast totals
    high_confidence = metrics['total_progress']  # Invoiced + PF(date) + PA(date) + HS E/C
    
    full_forecast = (high_confidence + 
                    metrics['pending_fulfillment_no_date'] + 
                    metrics['pending_approval_no_date'] + 
                    metrics['pending_approval_old'])
    
    gap_to_quota = metrics['quota'] - high_confidence
    
    potential_attainment_value = high_confidence + metrics['best_opp']
    potential_attainment_pct = (potential_attainment_value / metrics['quota'] * 100) if metrics['quota'] > 0 else 0
    
    # NEW: Top Metrics Row (mirroring Team Scorecard)
    st.markdown("### 📊 Rep Scorecard")
    
    metric_col1, metric_col2, metric_col3, metric_col4, metric_col5 = st.columns(5)
    
    with metric_col1:
        st.metric(
            label="💰 Quota",
            value=f"${metrics['quota']/1000:.0f}K" if metrics['quota'] < 1000000 else f"${metrics['quota']/1000000:.1f}M",
            help="Your Q4 2025 sales quota"
        )
    
    with metric_col2:
        high_conf_pct = (high_confidence / metrics['quota'] * 100) if metrics['quota'] > 0 else 0
        st.metric(
            label="💪 High Confidence Forecast",
            value=f"${high_confidence/1000:.0f}K" if high_confidence < 1000000 else f"${high_confidence/1000000:.1f}M",
            delta=f"{high_conf_pct:.1f}% of quota",
            help="Invoiced & Shipped + PF (with date) + PA (with date) + HS Expect/Commit"
        )
    
    with metric_col3:
        full_forecast_pct = (full_forecast / metrics['quota'] * 100) if metrics['quota'] > 0 else 0
        st.metric(
            label="📊 Full Forecast (All Sources)",
            value=f"${full_forecast/1000:.0f}K" if full_forecast < 1000000 else f"${full_forecast/1000000:.1f}M",
            delta=f"{full_forecast_pct:.1f}% of quota",
            help="Invoiced & Shipped + PF (with date) + PA (with date) + HS Expect/Commit + PF (without date) + PA (without date) + PA (>2 weeks old)"
        )
    
    with metric_col4:
        st.metric(
            label="📉 Gap to Quota",
            value=f"${gap_to_quota/1000:.0f}K" if abs(gap_to_quota) < 1000000 else f"${gap_to_quota/1000000:.1f}M",
            delta=f"${-gap_to_quota/1000:.0f}K" if gap_to_quota < 0 else None,
            delta_color="inverse",
            help="Quota - (Invoiced & Shipped + PF (with date) + PA (with date) + HS Expect/Commit)"
        )
    
    with metric_col5:
        upside = potential_attainment_pct - high_conf_pct
        st.metric(
            label="⭐ Potential Attainment",
            value=f"{potential_attainment_pct:.1f}%",
            delta=f"+{upside:.1f}% upside",
            help="(Invoiced & Shipped + PF (with date) + PA (with date) + HS Expect/Commit + HS Best Case/Opp) ÷ Quota"
        )
    
    st.markdown("---")
    
    # Invoices section for this rep
    display_invoices_drill_down(invoices_df, rep_name)
    
    st.markdown("---")
    
    # Build Your Own Forecast section
    build_your_own_forecast_section(
        metrics,
        metrics['quota'],
        rep_name=rep_name,
        deals_df=deals_df,
        invoices_df=invoices_df,
        sales_orders_df=sales_orders_df
    )
    
    st.markdown("---")
    
    # HubSpot Deals Audit Section
    display_hubspot_deals_audit(deals_df, rep_name)
    
    st.markdown("---")
    
    # SECTION 1: What's in NetSuite with Dates and HubSpot Expect/Commit
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>💰 Section 1: What's in NetSuite with Dates and HubSpot Expect/Commit</h3>
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
            <span class="progress-label">💪 THE SAFE BET TOTAL</span>
            <span class="progress-value">${high_confidence:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Drill-down sections for Section 1
    st.markdown("#### 📊 Section 1 Details")
    
    col1, col2 = st.columns(2)
    
    with col1:
        display_drill_down_section(
            "📦 Pending Fulfillment (with date)",
            metrics['pending_fulfillment'],
            metrics.get('pending_fulfillment_details', pd.DataFrame()),
            f"{rep_name}_pf"
        )
        
        display_drill_down_section(
            "⏳ Pending Approval (with date)",
            metrics['pending_approval'],
            metrics.get('pending_approval_details', pd.DataFrame()),
            f"{rep_name}_pa"
        )
    
    with col2:
        display_drill_down_section(
            "🎯 HubSpot Expect/Commit",
            metrics['expect_commit'],
            metrics.get('expect_commit_deals', pd.DataFrame()),
            f"{rep_name}_hs"
        )
        
        display_drill_down_section(
            "🎲 Best Case/Opportunity",
            metrics['best_opp'],
            metrics.get('best_opp_deals', pd.DataFrame()),
            f"{rep_name}_bo"
        )
    
    st.markdown("---")
    
    # SECTION 2: Full Forecast
    st.markdown(f"""
    <div class="progress-breakdown">
        <h3>📊 Section 2: Full Forecast</h3>
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
            <span class="progress-label">📊 FULL FORECAST TOTAL</span>
            <span class="progress-value">${full_forecast:,.0f}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Drill-down sections for Section 2 (additional items)
    st.markdown("#### 📊 Section 2 Additional Details")
    
    warning_col1, warning_col2, warning_col3 = st.columns(3)
    
    with warning_col1:
        display_drill_down_section(
            "📦 Pending Fulfillment (without date)",
            metrics['pending_fulfillment_no_date'],
            metrics.get('pending_fulfillment_no_date_details', pd.DataFrame()),
            f"{rep_name}_pf_no_date"
        )
    
    with warning_col2:
        display_drill_down_section(
            "⏳ Pending Approval (without date)",
            metrics['pending_approval_no_date'],
            metrics.get('pending_approval_no_date_details', pd.DataFrame()),
            f"{rep_name}_pa_no_date"
        )
    
    with warning_col3:
        display_drill_down_section(
            "⏱️ Old Pending Approval (>2 weeks)",
            metrics['pending_approval_old'],
            metrics.get('pending_approval_old_details', pd.DataFrame()),
            f"{rep_name}_pa_old"
        )
    
    st.markdown("---")
    
    # Q1 2026 Spillover Details (moved from Section 3)
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
        # Sexy header with gradient
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            margin-bottom: 20px;
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
        ">
            <h1 style="
                color: white;
                font-size: 28px;
                margin: 0;
                font-weight: 800;
                text-shadow: 0 2px 10px rgba(0,0,0,0.3);
            ">📊 Calyx Command</h1>
            <p style="
                color: rgba(255,255,255,0.9);
                font-size: 14px;
                margin: 8px 0 0 0;
                font-weight: 500;
            ">Q4 2025 Sales Intelligence</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Custom navigation with icons and descriptions
        st.markdown("### 🧭 Navigation")
        
        # ERP-style navigation with CSS styling
        st.markdown("""
        <style>
        div[data-testid="stRadio"] > div {
            gap: 8px;
        }
        
        div[data-testid="stRadio"] > div > label {
            background: rgba(30, 41, 59, 0.6) !important;
            border: 1px solid rgba(71, 85, 105, 0.5) !important;
            border-left: 4px solid transparent !important;
            border-radius: 8px !important;
            padding: 12px 16px !important;
            cursor: pointer !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
            width: 100% !important;
            margin-bottom: 4px !important;
        }
        
        div[data-testid="stRadio"] > div > label:hover {
            background: rgba(51, 65, 85, 0.8) !important;
            border-color: rgba(100, 116, 139, 0.7) !important;
            transform: translateX(4px);
        }
        
        div[data-testid="stRadio"] > div > label[data-checked="true"] {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
            border: 2px solid #3b82f6 !important;
            border-left: 4px solid #60a5fa !important;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
        }
        
        div[data-testid="stRadio"] > div > label[data-checked="true"]:hover {
            transform: translateX(0);
        }
        
        div[data-testid="stRadio"] label p {
            font-size: 14px !important;
            font-weight: 600 !important;
            margin: 0 !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        # Create navigation options
        view_mode = st.radio(
            "Select View:",
            ["👥 Team Overview", "👤 Individual Rep", "🔍 Reconciliation", "🤖 AI Insights", "💰 Commission", "🧪 Concentrate Jar Forecast", "📦 All Products Forecast"],
            label_visibility="collapsed",
            key="nav_selector"
        )
        
        # Map display names back to internal names
        view_mapping = {
            "👥 Team Overview": "Team Overview",
            "👤 Individual Rep": "Individual Rep",
            "🔍 Reconciliation": "Reconciliation",
            "🤖 AI Insights": "AI Insights",
            "💰 Commission": "💰 Commission",
            "🧪 Concentrate Jar Forecast": "🧪 Concentrate Jar Forecast",
            "📦 All Products Forecast": "📦 All Products Forecast"
        }
        
        view_mode = view_mapping.get(view_mode, "Team Overview")
        
        st.markdown("---")
        
        # Sexy metrics cards for quick stats
        current_time = datetime.now()
        biz_days = calculate_business_days_remaining()
        
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, rgba(16, 185, 129, 0.2) 0%, rgba(5, 150, 105, 0.2) 100%);
            border: 1px solid rgba(16, 185, 129, 0.3);
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 15px;
        ">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                <span style="font-size: 24px;">⏱️</span>
                <div>
                    <div style="font-size: 11px; opacity: 0.7; text-transform: uppercase; letter-spacing: 1px;">Q4 Days Left</div>
                    <div style="font-size: 24px; font-weight: 700; color: #10b981;">""" + str(biz_days) + """</div>
                </div>
            </div>
            <div style="font-size: 10px; opacity: 0.6;">Business days until Dec 31, 2025</div>
        </div>
        
        <div style="
            background: linear-gradient(135deg, rgba(59, 130, 246, 0.2) 0%, rgba(37, 99, 235, 0.2) 100%);
            border: 1px solid rgba(59, 130, 246, 0.3);
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 15px;
        ">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                <span style="font-size: 24px;">🔄</span>
                <div style="flex: 1;">
                    <div style="font-size: 11px; opacity: 0.7; text-transform: uppercase; letter-spacing: 1px;">Last Sync</div>
                    <div style="font-size: 14px; font-weight: 600; color: #3b82f6;">""" + current_time.strftime('%I:%M %p') + """</div>
                </div>
            </div>
            <div style="font-size: 10px; opacity: 0.6;">Auto-refresh every hour</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Refresh button with gradient
        if st.button("🔄 Refresh Data Now", use_container_width=True):
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
    elif view_mode == "💰 Commission":
        # Commission calculator view (password protected)
        commission_calculator.display_commission_section(invoices_df, sales_orders_df)
    elif view_mode == "🧪 Concentrate Jar Forecast":
        # Concentrate Jar Forecasting view
        if SHIPPING_PLANNING_AVAILABLE:
            shipping_planning.main()
        else:
            st.error("❌ Concentrate Jar Forecasting module not found.")
            if 'SHIPPING_PLANNING_ERROR' in globals():
                st.error(f"Error details: {SHIPPING_PLANNING_ERROR}")
            st.info("Make sure shipping_planning.py is in your repository at the same level as this dashboard file.")
            st.code("Expected file location: shipping_planning.py")
            
            # Debug info
            with st.expander("🔧 Debug Information"):
                st.write("**Current working directory:**")
                import os
                st.code(os.getcwd())
                st.write("**Files in current directory:**")
                try:
                    files = os.listdir('.')
                    st.code('\n'.join([f for f in files if f.endswith('.py')]))
                except Exception as e:
                    st.error(f"Cannot list files: {e}")
    elif view_mode == "📦 All Products Forecast":
        # All Products Forecasting view
        if ALL_PRODUCTS_FORECAST_AVAILABLE:
            all_products_forecast.main()
        else:
            st.error("❌ All Products Forecast module not found.")
            if 'ALL_PRODUCTS_FORECAST_ERROR' in globals():
                st.error(f"Error details: {ALL_PRODUCTS_FORECAST_ERROR}")
            st.info("Make sure all_products_forecast.py is in your repository at the same level as this dashboard file.")
            st.code("Expected file location: all_products_forecast.py")
            
            # Debug info
            with st.expander("🔧 Debug Information"):
                st.write("**Current working directory:**")
                import os
                st.code(os.getcwd())
                st.write("**Files in current directory:**")
                try:
                    files = os.listdir('.')
                    st.code('\n'.join([f for f in files if f.endswith('.py')]))
                except Exception as e:
                    st.error(f"Cannot list files: {e}")
    else:  # Reconciliation view
        display_reconciliation_view(deals_df, dashboard_df, sales_orders_df)

if __name__ == "__main__":
    main()
