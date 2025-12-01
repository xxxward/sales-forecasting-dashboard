"""
Calyx Containers - Elite Commission Calculator
The most powerful commission tracking system ever built
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import hashlib
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==========================================
# CONFIGURATION
# ==========================================
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

ADMIN_EMAIL = "xward@calyxcontainers.com"
ADMIN_PASSWORD_HASH = hashlib.sha256("Secret2025!".encode()).hexdigest()

COMMISSION_REPS = ["Dave Borkowski", "Jake Lynch", "Brad Sherman", "Lance Mitton"]

REP_COMMISSION_RATES = {
    "Dave Borkowski": 0.05,
    "Jake Lynch": 0.07,
    "Brad Sherman": 0.07,
    "Lance Mitton": 0.07,
}

BRAD_OVERRIDE_RATE = 0.01

REP_COLORS = {
    "Dave Borkowski": "#667eea",
    "Jake Lynch": "#f093fb",
    "Brad Sherman": "#4facfe",
    "Lance Mitton": "#43e97b"
}

# ==========================================
# CUSTOM CSS
# ==========================================
def load_custom_css():
    st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: white !important;
    }
    
    .metric-card {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        margin: 10px 0;
    }
    
    .rep-card {
        background: white;
        border-radius: 20px;
        padding: 25px;
        margin: 15px 0;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.15);
        border-left: 5px solid;
        transition: transform 0.2s;
    }
    
    .rep-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 50px rgba(0, 0, 0, 0.2);
    }
    
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    section[data-testid="stSidebar"] .stMarkdown {
        color: white !important;
    }
    
    /* Make all sidebar inputs visible with solid backgrounds */
    section[data-testid="stSidebar"] [data-baseweb="select"],
    section[data-testid="stSidebar"] [data-baseweb="input"],
    section[data-testid="stSidebar"] .stMultiSelect > div > div,
    section[data-testid="stSidebar"] .stDateInput > div > div {
        background-color: white !important;
        border: 2px solid white !important;
        border-radius: 8px !important;
    }
    
    section[data-testid="stSidebar"] .stMultiSelect label,
    section[data-testid="stSidebar"] .stDateInput label {
        color: white !important;
        font-weight: 700 !important;
        font-size: 16px !important;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
    }
    
    /* Make multiselect text visible */
    section[data-testid="stSidebar"] [data-baseweb="select"] span {
        color: #333 !important;
    }
    
    /* Style sidebar buttons */
    section[data-testid="stSidebar"] .stButton > button {
        background-color: white !important;
        color: #667eea !important;
        border: 2px solid white !important;
        font-weight: 700 !important;
    }
    
    section[data-testid="stSidebar"] .stButton > button:hover {
        background-color: rgba(255, 255, 255, 0.9) !important;
        transform: scale(1.05) !important;
    }
    
    .stButton>button {
        border-radius: 10px;
        font-weight: 600;
        transition: all 0.3s;
    }
    
    .stButton>button:hover {
        transform: scale(1.05);
    }
    
    [data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: 700;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# DATA LOADING
# ==========================================

@st.cache_data(ttl=3600)
def fetch_google_sheet_data(sheet_name, range_name):
    try:
        if "gcp_service_account" not in st.secrets:
            return pd.DataFrame()

        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        if not values:
            return pd.DataFrame()
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df

    except Exception as e:
        st.error(f"Error: {str(e)}")
        return pd.DataFrame()

def process_ns_invoices(df):
    if df.empty:
        return df
    
    df.columns = df.columns.str.strip()
    
    if 'Date' in df.columns and 'Date Closed' in df.columns:
        df = df.drop(columns=['Date'])
    if 'Sales Rep' in df.columns:
        df = df.rename(columns={'Sales Rep': 'Original Sales Rep'})
    
    rename_map = {
        'Amount (Transaction Total)': 'Amount',
        'Date Closed': 'Close Date',
        'Document Number': 'Invoice',
        'Status': 'Status',
        'Amount (Transaction Tax Total)': 'Tax Amount',
        'Amount (Shipping)': 'Shipping Amount',
        'Corrected Customer Name': 'Customer',
        'Created From': 'SO Number',
        'CSM': 'CSM',
        'HubSpot Pipeline': 'Pipeline',
        'Rep Master': 'Rep Master'
    }
    
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
    
    if 'Rep Master' in df.columns:
        df['Sales Rep'] = df['Rep Master']
    
    # Clean customer names - defensive approach
    try:
        if 'Customer' in df.columns:
            df['Customer'] = df['Customer'].astype(str).str.replace('^Customer ', '', regex=True)
    except:
        pass
    
    # Clean SO Number - defensive approach
    try:
        if 'SO Number' in df.columns:
            df['SO Number'] = df['SO Number'].astype(str).str.replace('Sales Order', '').str.strip()
    except:
        pass
    
    # Parse dates - defensive approach
    try:
        if 'Close Date' in df.columns:
            df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    except:
        pass
    
    numeric_cols = ['Amount', 'Tax Amount', 'Shipping Amount']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    
    # Calculate subtotal - defensive approach
    try:
        if 'Amount' in df.columns:
            df['Subtotal'] = df['Amount']
            if 'Tax Amount' in df.columns:
                df['Subtotal'] = df['Subtotal'] - df['Tax Amount']
            if 'Shipping Amount' in df.columns:
                df['Subtotal'] = df['Subtotal'] - df['Shipping Amount']
        else:
            df['Subtotal'] = 0
    except:
        df['Subtotal'] = 0
    
    return df

def calculate_commissions(df):
    if df.empty:
        return df
    
    df['Commission Rate'] = df['Sales Rep'].map(REP_COMMISSION_RATES).fillna(0.0)
    df['Commission'] = df['Subtotal'] * df['Commission Rate']
    
    df['Brad Override'] = 0.0
    lance_mask = df['Sales Rep'] == 'Lance Mitton'
    df.loc[lance_mask, 'Brad Override'] = df.loc[lance_mask, 'Subtotal'] * BRAD_OVERRIDE_RATE
    
    df['Row_ID'] = range(len(df))
    
    return df

# ==========================================
# AUTH
# ==========================================

def verify_admin(email, password):
    if email != ADMIN_EMAIL:
        return False
    password_hash = hashlib.sha256(password.encode()).hexdigest()
    return password_hash == ADMIN_PASSWORD_HASH

def display_login():
    load_custom_css()
    
    st.markdown("""
    <div style='text-align: center; padding: 100px 20px;'>
        <h1 style='font-size: 48px; margin-bottom: 10px;'>💎 Elite Commission Tracker</h1>
        <p style='font-size: 20px; color: rgba(255,255,255,0.9); margin-bottom: 50px;'>
            The most powerful commission calculator ever built
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
        email = st.text_input("Email", placeholder="your@email.com")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        
        if st.button("🚀 Launch Dashboard", use_container_width=True):
            if verify_admin(email, password):
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ Invalid credentials")
        st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MAIN DASHBOARD
# ==========================================

def display_dashboard():
    load_custom_css()
    
    if 'selected_rows' not in st.session_state:
        st.session_state.selected_rows = set()
    
    col1, col2 = st.columns([5, 1])
    with col1:
        st.markdown("""
        <div style='padding: 20px; margin-bottom: 30px;'>
            <h1 style='font-size: 42px; margin: 0;'>💎 Elite Commission Tracker</h1>
            <p style='font-size: 18px; color: rgba(255,255,255,0.9); margin: 5px 0 0 0;'>
                Real-time commission analytics & management
            </p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()
    
    with st.spinner("🔄 Loading commission data..."):
        raw_data = fetch_google_sheet_data("NS Invoices", "A:U")
        if raw_data.empty:
            st.error("❌ Could not load data")
            return
        
        df = process_ns_invoices(raw_data)
        if df.empty:
            st.error("❌ No data after processing")
            return
    
    with st.sidebar:
        st.markdown("""
        <div style='background: rgba(255,255,255,0.2); padding: 15px; border-radius: 10px; margin-bottom: 20px; text-align: center;'>
            <h2 style='color: white; margin: 0;'>🎛️ Filters</h2>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<div style='margin: 20px 0 10px 0;'><strong style='color: white; font-size: 18px;'>📅 Date Range</strong></div>", unsafe_allow_html=True)
        if 'Close Date' in df.columns and not df['Close Date'].isna().all():
            min_date = df['Close Date'].min()
            max_date = df['Close Date'].max()
            
            date_range = st.date_input(
                "Select date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                label_visibility="collapsed"
            )
        else:
            date_range = None
        
        st.markdown("<hr style='border: 1px solid rgba(255,255,255,0.3); margin: 20px 0;'>", unsafe_allow_html=True)
        
        st.markdown("<div style='margin: 20px 0 10px 0;'><strong style='color: white; font-size: 18px;'>📊 Status</strong></div>", unsafe_allow_html=True)
        if 'Status' in df.columns:
            unique_statuses = df['Status'].unique().tolist()
            selected_statuses = st.multiselect(
                "Filter by status",
                options=unique_statuses,
                default=["Paid In Full"] if "Paid In Full" in unique_statuses else unique_statuses,
                label_visibility="collapsed"
            )
        else:
            selected_statuses = []
        
        st.markdown("<hr style='border: 1px solid rgba(255,255,255,0.3); margin: 20px 0;'>", unsafe_allow_html=True)
        
        st.markdown("<div style='margin: 20px 0 10px 0;'><strong style='color: white; font-size: 18px;'>👤 Sales Reps</strong></div>", unsafe_allow_html=True)
        selected_reps = st.multiselect(
            "Select reps to view",
            options=COMMISSION_REPS,
            default=COMMISSION_REPS,
            label_visibility="collapsed"
        )
        
        st.markdown("<hr style='border: 1px solid rgba(255,255,255,0.3); margin: 20px 0;'>", unsafe_allow_html=True)
        
        st.markdown("<div style='margin: 20px 0 10px 0;'><strong style='color: white; font-size: 18px;'>⚡ Quick Actions</strong></div>", unsafe_allow_html=True)
        if st.button("✅ Select All", use_container_width=True):
            st.session_state.selected_rows = set(df['Row_ID'].tolist()) if 'Row_ID' in df.columns else set()
            st.rerun()
        if st.button("❌ Clear All", use_container_width=True):
            st.session_state.selected_rows = set()
            st.rerun()
    
    filtered_df = df.copy()
    
    if 'Sales Rep' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Sales Rep'].isin(COMMISSION_REPS)]
        filtered_df = filtered_df[filtered_df['Sales Rep'].isin(selected_reps)]
    
    if 'Status' in filtered_df.columns and selected_statuses:
        filtered_df = filtered_df[filtered_df['Status'].isin(selected_statuses)]
    
    if 'Original Sales Rep' in filtered_df.columns:
        filtered_df = filtered_df[~filtered_df['Original Sales Rep'].astype(str).str.upper().str.contains("SHOPIFY", na=False)]
    
    if date_range and len(date_range) == 2 and 'Close Date' in filtered_df.columns:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1]) + timedelta(days=1)
        filtered_df = filtered_df[(filtered_df['Close Date'] >= start_date) & (filtered_df['Close Date'] < end_date)]
    
    filtered_df = calculate_commissions(filtered_df)
    
    if filtered_df.empty:
        st.warning("⚠️ No transactions match filters")
        return
    
    if not st.session_state.selected_rows:
        st.session_state.selected_rows = set(filtered_df['Row_ID'].tolist())
    
    included_df = filtered_df[filtered_df['Row_ID'].isin(st.session_state.selected_rows)]
    
    st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "💰 Total Sales",
            f"${included_df['Subtotal'].sum():,.0f}",
            delta=f"{len(included_df)} transactions"
        )
    
    with col2:
        st.metric(
            "💵 Total Commission",
            f"${included_df['Commission'].sum():,.2f}",
            delta=f"{included_df['Commission'].sum() / included_df['Subtotal'].sum() * 100:.1f}% avg" if included_df['Subtotal'].sum() > 0 else "0%"
        )
    
    with col3:
        st.metric(
            "🎯 Brad Override",
            f"${included_df['Brad Override'].sum():,.2f}",
            delta="1% on Lance"
        )
    
    with col4:
        st.metric(
            "📊 Selection",
            f"{len(included_df)} / {len(filtered_df)}",
            delta=f"{len(included_df)/len(filtered_df)*100:.0f}%" if len(filtered_df) > 0 else "0%"
        )
    
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    
    for rep in selected_reps:
        rep_data = filtered_df[filtered_df['Sales Rep'] == rep].copy()
        if rep_data.empty:
            continue
        
        rep_included = included_df[included_df['Sales Rep'] == rep]
        rep_color = REP_COLORS.get(rep, "#667eea")
        
        st.markdown(f"""
        <div class='rep-card' style='border-left-color: {rep_color};'>
            <h2 style='color: {rep_color}; margin: 0 0 10px 0;'>👤 {rep}</h2>
            <div style='display: flex; gap: 20px; margin-bottom: 20px;'>
                <div>
                    <span style='color: #666; font-size: 14px;'>Sales</span>
                    <div style='font-size: 24px; font-weight: 700; color: {rep_color};'>${rep_included['Subtotal'].sum():,.0f}</div>
                </div>
                <div>
                    <span style='color: #666; font-size: 14px;'>Commission</span>
                    <div style='font-size: 24px; font-weight: 700; color: {rep_color};'>${rep_included['Commission'].sum():,.2f}</div>
                </div>
                <div>
                    <span style='color: #666; font-size: 14px;'>Transactions</span>
                    <div style='font-size: 24px; font-weight: 700; color: {rep_color};'>{len(rep_included)} / {len(rep_data)}</div>
                </div>
                <div>
                    <span style='color: #666; font-size: 14px;'>Rate</span>
                    <div style='font-size: 24px; font-weight: 700; color: {rep_color};'>{REP_COMMISSION_RATES.get(rep, 0):.0%}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.expander(f"📋 View {rep}'s Transactions ({len(rep_data)})", expanded=False):
            all_selected = all(row_id in st.session_state.selected_rows for row_id in rep_data['Row_ID'].tolist())
            
            master_check = st.checkbox(
                f"Select all {rep}'s transactions",
                value=all_selected,
                key=f"master_{rep}"
            )
            
            if master_check and not all_selected:
                st.session_state.selected_rows.update(rep_data['Row_ID'].tolist())
                st.rerun()
            elif not master_check and all_selected:
                st.session_state.selected_rows -= set(rep_data['Row_ID'].tolist())
                st.rerun()
            
            st.markdown("---")
            
            cols = st.columns([0.5, 1.2, 1.2, 0.7, 2.0, 0.8, 1.0, 0.9, 1.0])
            headers = ["☑️", "Invoice", "SO #", "Status", "Customer", "Close Date", "Pipeline", "Amount", "Commission"]
            for col, header in zip(cols, headers):
                col.markdown(f"**{header}**")
            
            st.markdown("---")
            
            for _, row in rep_data.iterrows():
                row_id = row['Row_ID']
                is_selected = row_id in st.session_state.selected_rows
                
                cols = st.columns([0.5, 1.2, 1.2, 0.7, 2.0, 0.8, 1.0, 0.9, 1.0])
                
                with cols[0]:
                    if st.checkbox("", value=is_selected, key=f"cb_{row_id}", label_visibility="collapsed"):
                        if row_id not in st.session_state.selected_rows:
                            st.session_state.selected_rows.add(row_id)
                            st.rerun()
                    else:
                        if row_id in st.session_state.selected_rows:
                            st.session_state.selected_rows.remove(row_id)
                            st.rerun()
                
                with cols[1]:
                    st.text(str(row.get('Invoice', 'N/A')))
                
                with cols[2]:
                    st.text(str(row.get('SO Number', 'N/A')))
                
                with cols[3]:
                    st.text(str(row.get('Status', 'N/A'))[:12])
                
                with cols[4]:
                    customer = str(row.get('Customer', 'N/A'))
                    if customer in ['N/A', 'nan', '']:
                        customer = 'No Customer'
                    st.text(customer)
                
                with cols[5]:
                    close_date = row.get('Close Date')
                    if pd.notna(close_date):
                        st.text(close_date.strftime('%Y-%m-%d'))
                    else:
                        st.text('N/A')
                
                with cols[6]:
                    st.text(str(row.get('Pipeline', 'N/A'))[:15])
                
                with cols[7]:
                    st.text(f"${row.get('Subtotal', 0):,.0f}")
                
                with cols[8]:
                    st.text(f"${row.get('Commission', 0):,.2f}")
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='metric-card'>", unsafe_allow_html=True)
    st.markdown("### 📥 Export Data")
    
    if not included_df.empty:
        export_cols = ['Invoice', 'SO Number', 'Status', 'Customer', 'Close Date', 'CSM', 
                       'Rep Master', 'Pipeline', 'Amount', 'Subtotal', 'Commission Rate', 
                       'Commission', 'Brad Override']
        export_cols = [col for col in export_cols if col in included_df.columns]
        
        export_df = included_df[export_cols].copy()
        
        if 'Close Date' in export_df.columns:
            export_df['Close Date'] = export_df['Close Date'].dt.strftime('%Y-%m-%d')
        
        csv = export_df.to_csv(index=False)
        st.download_button(
            "📥 Download Selected Transactions (CSV)",
            data=csv,
            file_name=f"commission_report_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# MAIN APP
# ==========================================

def display_commission_section(invoices_df=None, sales_orders_df=None):
    # Ensure page config is set
    try:
        st.set_page_config(
            page_title="Elite Commission Tracker",
            page_icon="💎",
            layout="wide",
            initial_sidebar_state="expanded"
        )
    except:
        pass  # Page config already set
    
    if not st.session_state.get('authenticated', False):
        display_login()
    else:
        display_dashboard()

if __name__ == "__main__":
    st.set_page_config(
        page_title="Elite Commission Tracker",
        page_icon="💎",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    display_commission_section()
