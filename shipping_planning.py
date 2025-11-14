"""
Q4 2025 Shipping Planning Tool - UPDATED VERSION
Separates Shipped/Invoiced from Planning, shows actual document numbers
"""

import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600
CACHE_VERSION = "v1_simple"

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets - EXACT COPY from main dashboard"""
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
        
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{sheet_name}!{range_name}"
        ).execute()
        
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame()
        
        # Handle mismatched column counts
        if len(values) > 1:
            max_cols = max(len(row) for row in values)
            for row in values:
                while len(row) < max_cols:
                    row.append('')
        
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
        
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return pd.DataFrame()

def load_all_data():
    """Load data - EXACT COPY from main dashboard load_all_data function"""
    
    # Load deals
    deals_df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
    # Load dashboard info
    dashboard_df = load_google_sheets_data("Dashboard Info", "A:C", version=CACHE_VERSION)
    
    # Load invoices
    invoices_df = load_google_sheets_data("NS Invoices", "A:U", version=CACHE_VERSION)
    
    # Load sales orders
    sales_orders_df = load_google_sheets_data("NS Sales Orders", "A:AD", version=CACHE_VERSION)
    
    # Process deals - SIMPLIFIED VERSION (just the essentials)
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
            elif col == 'Amount':
                rename_dict[col] = 'Amount'
            elif col == 'Close Status':
                rename_dict[col] = 'Status'
            elif col == 'Pipeline':
                rename_dict[col] = 'Pipeline'
            elif col == 'Deal Type':
                rename_dict[col] = 'Product Type'
            elif col == 'Q1 2026 Spillover':
                rename_dict[col] = 'Q1 2026 Spillover'
        
        deals_df = deals_df.rename(columns=rename_dict)
        
        # Create Deal Owner if needed
        if 'Deal Owner' not in deals_df.columns:
            if 'Deal Owner First Name' in deals_df.columns and 'Deal Owner Last Name' in deals_df.columns:
                deals_df['Deal Owner'] = deals_df['Deal Owner First Name'].fillna('') + ' ' + deals_df['Deal Owner Last Name'].fillna('')
                deals_df['Deal Owner'] = deals_df['Deal Owner'].str.strip()
        
        # Clean Amount
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
        
        # Convert Close Date
        if 'Close Date' in deals_df.columns:
            deals_df['Close Date'] = pd.to_datetime(deals_df['Close Date'], errors='coerce')
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2026-01-01')
        
        if 'Close Date' in deals_df.columns:
            deals_df = deals_df[(deals_df['Close Date'] >= q4_start) & (deals_df['Close Date'] < q4_end)]
        
        # Add Counts_In_Q4 flag (simplified - just check if Q1 2026 Spillover column says "Q1 2026")
        if 'Q1 2026 Spillover' in deals_df.columns:
            deals_df['Counts_In_Q4'] = deals_df['Q1 2026 Spillover'] != 'Q1 2026'
        else:
            deals_df['Counts_In_Q4'] = True
    
    # Process dashboard info
    if not dashboard_df.empty and len(dashboard_df.columns) >= 3:
        dashboard_df.columns = ['Rep Name', 'Quota', 'NetSuite Orders']
        dashboard_df = dashboard_df[dashboard_df['Rep Name'].notna() & (dashboard_df['Rep Name'] != '')]
        
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
    
    # Process invoices
    if not invoices_df.empty and len(invoices_df.columns) >= 15:
        rename_dict = {
            invoices_df.columns[0]: 'Invoice Number',
            invoices_df.columns[1]: 'Status',
            invoices_df.columns[2]: 'Date',
            invoices_df.columns[6]: 'Customer',
            invoices_df.columns[10]: 'Amount',
            invoices_df.columns[14]: 'Sales Rep'
        }
        
        if len(invoices_df.columns) > 19:
            rename_dict[invoices_df.columns[19]] = 'Corrected Customer Name'
        if len(invoices_df.columns) > 20:
            rename_dict[invoices_df.columns[20]] = 'Rep Master'
        
        invoices_df = invoices_df.rename(columns=rename_dict)
        
        # Apply Rep Master override
        if 'Rep Master' in invoices_df.columns:
            mask = invoices_df['Rep Master'].notna() & (invoices_df['Rep Master'] != '')
            invoices_df.loc[mask, 'Sales Rep'] = invoices_df.loc[mask, 'Rep Master']
        
        # Clean Amount
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
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2026-01-01')
        invoices_df = invoices_df[
            (invoices_df['Date'] >= q4_start) & 
            (invoices_df['Date'] < q4_end)
        ]
    
    # Process sales orders
    if not sales_orders_df.empty and len(sales_orders_df.columns) >= 15:
        rename_dict = {
            sales_orders_df.columns[0]: 'Document Number',
            sales_orders_df.columns[1]: 'Status',
            sales_orders_df.columns[2]: 'Order Type',
            sales_orders_df.columns[3]: 'Order Start Date',
            sales_orders_df.columns[6]: 'Customer',
            sales_orders_df.columns[10]: 'Amount',
            sales_orders_df.columns[14]: 'Sales Rep'
        }
        
        if len(sales_orders_df.columns) > 24:
            rename_dict[sales_orders_df.columns[24]] = 'Corrected Customer Name'
        if len(sales_orders_df.columns) > 25:
            rename_dict[sales_orders_df.columns[25]] = 'Rep Master'
        if len(sales_orders_df.columns) > 29:
            rename_dict[sales_orders_df.columns[29]] = 'Age_Business_Days'
        
        sales_orders_df = sales_orders_df.rename(columns=rename_dict)
        
        # Apply Rep Master override
        if 'Rep Master' in sales_orders_df.columns:
            mask = sales_orders_df['Rep Master'].notna() & (sales_orders_df['Rep Master'] != '')
            sales_orders_df.loc[mask, 'Sales Rep'] = sales_orders_df.loc[mask, 'Rep Master']
        
        # Clean Amount
        def clean_numeric(value):
            if pd.isna(value) or str(value).strip() == '':
                return 0
            cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
            try:
                return float(cleaned)
            except:
                return 0
        
        sales_orders_df['Amount'] = sales_orders_df['Amount'].apply(clean_numeric)
        sales_orders_df['Order Start Date'] = pd.to_datetime(sales_orders_df['Order Start Date'], errors='coerce')
        
        # Filter to Q4 2025
        q4_start = pd.Timestamp('2025-10-01')
        q4_end = pd.Timestamp('2026-01-01')
        sales_orders_df = sales_orders_df[
            (sales_orders_df['Order Start Date'] >= q4_start) & 
            (sales_orders_df['Order Start Date'] < q4_end)
        ]
        
        # Convert Age to numeric
        if 'Age_Business_Days' in sales_orders_df.columns:
            sales_orders_df['Age_Business_Days'] = pd.to_numeric(
                sales_orders_df['Age_Business_Days'], 
                errors='coerce'
            ).fillna(0)
    
    return deals_df, dashboard_df, invoices_df, sales_orders_df

def calculate_team_metrics(deals_df, dashboard_df, invoices_df, sales_orders_df):
    """Calculate metrics for the team"""
    metrics = {
        'invoices': 0,
        'shipped_not_invoiced': 0,
        'fulfilled_not_shipped': 0,
        'pending_fulfillment': 0,
        'pending_approval': 0,
        'old_pending_approval': 0,
        'hubspot_expect': 0,
        'hubspot_commit': 0,
        'hubspot_best_case': 0,
        'hubspot_opportunity': 0,
        'q1_spillover': 0
    }
    
    # Invoices
    if invoices_df is not None and not invoices_df.empty:
        metrics['invoices'] = invoices_df['Amount'].sum()
    
    # Sales Orders
    if sales_orders_df is not None and not sales_orders_df.empty:
        if 'Status' in sales_orders_df.columns:
            metrics['shipped_not_invoiced'] = sales_orders_df[
                sales_orders_df['Status'] == 'Shipped, Not Invoiced'
            ]['Amount'].sum()
            
            metrics['fulfilled_not_shipped'] = sales_orders_df[
                sales_orders_df['Status'] == 'Fulfilled, Not Shipped'
            ]['Amount'].sum()
            
            metrics['pending_fulfillment'] = sales_orders_df[
                sales_orders_df['Status'] == 'Pending Fulfillment'
            ]['Amount'].sum()
            
            metrics['pending_approval'] = sales_orders_df[
                sales_orders_df['Status'] == 'Pending Approval'
            ]['Amount'].sum()
            
            if 'Age_Business_Days' in sales_orders_df.columns:
                metrics['old_pending_approval'] = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Approval') &
                    (sales_orders_df['Age_Business_Days'] >= 10)
                ]['Amount'].sum()
    
    # HubSpot Deals
    if deals_df is not None and not deals_df.empty and 'Status' in deals_df.columns:
        metrics['hubspot_expect'] = deals_df[deals_df['Status'] == 'Expect']['Amount'].sum()
        metrics['hubspot_commit'] = deals_df[deals_df['Status'] == 'Commit']['Amount'].sum()
        metrics['hubspot_best_case'] = deals_df[deals_df['Status'] == 'Best Case']['Amount'].sum()
        metrics['hubspot_opportunity'] = deals_df[deals_df['Status'] == 'Opportunity']['Amount'].sum()
        
        # Q1 Spillover
        metrics['q1_spillover'] = deals_df[
            (deals_df.get('Counts_In_Q4', True) == False) &
            (deals_df['Status'].isin(['Expect', 'Commit']))
        ]['Amount'].sum()
    
    return metrics

def build_shipping_plan_section(metrics, quota, rep_name, deals_df, invoices_df, sales_orders_df):
    """Build shipping plan with separated Shipped vs Planning sections"""
    
    # Initialize session state
    if 'shipped_sources' not in st.session_state:
        st.session_state.shipped_sources = {}
    if 'planning_sources' not in st.session_state:
        st.session_state.planning_sources = {}
    if 'individual_selection_mode' not in st.session_state:
        st.session_state.individual_selection_mode = {}
    if 'individual_selections' not in st.session_state:
        st.session_state.individual_selections = {}
    
    # Filter data by rep if needed
    if rep_name:
        if invoices_df is not None and not invoices_df.empty and 'Sales Rep' in invoices_df.columns:
            invoices_df = invoices_df[invoices_df['Sales Rep'] == rep_name]
        if sales_orders_df is not None and not sales_orders_df.empty and 'Sales Rep' in sales_orders_df.columns:
            sales_orders_df = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name]
    
    # ====================
    # SECTION 1: SHIPPED
    # ====================
    st.markdown("---")
    st.markdown("### 📦 Shipped (Already Counted)")
    st.caption("These items have already been shipped or invoiced and count toward your Q4 results")
    
    shipped_col1, shipped_col2 = st.columns([3, 1])
    
    with shipped_col1:
        # Invoices
        inv_key = 'Invoices'
        st.session_state.shipped_sources[inv_key] = st.checkbox(
            f"✅ Invoices: ${metrics['invoices']:,.0f}",
            value=st.session_state.shipped_sources.get(inv_key, False),
            key=f"shipped_{inv_key}_{rep_name or 'team'}"
        )
        
        # Show sample invoice numbers if selected
        if st.session_state.shipped_sources[inv_key] and invoices_df is not None and not invoices_df.empty:
            sample_invoices = invoices_df.head(3)
            st.caption(f"📄 Sample: {', '.join(sample_invoices['Invoice Number'].astype(str).tolist())}")
        
        # Shipped Not Invoiced
        sni_key = 'Shipped Not Invoiced'
        st.session_state.shipped_sources[sni_key] = st.checkbox(
            f"📦 Shipped, Not Invoiced: ${metrics['shipped_not_invoiced']:,.0f}",
            value=st.session_state.shipped_sources.get(sni_key, False),
            key=f"shipped_{sni_key}_{rep_name or 'team'}"
        )
        
        # Show sample SO numbers if selected
        if st.session_state.shipped_sources[sni_key] and sales_orders_df is not None and not sales_orders_df.empty:
            sample_sos = sales_orders_df[sales_orders_df['Status'] == 'Shipped, Not Invoiced'].head(3)
            if not sample_sos.empty:
                st.caption(f"📄 Sample: {', '.join(sample_sos['Document Number'].astype(str).tolist())}")
    
    with shipped_col2:
        shipped_total = sum([
            metrics['invoices'] if st.session_state.shipped_sources.get('Invoices', False) else 0,
            metrics['shipped_not_invoiced'] if st.session_state.shipped_sources.get('Shipped Not Invoiced', False) else 0
        ])
        
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, #2ECC71 0%, #27AE60 100%); 
                     padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 14px; margin-bottom: 5px;'>Shipped Total</div>
            <div style='color: white; font-size: 28px; font-weight: bold;'>${shipped_total:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    # ====================
    # SECTION 2: SHIPPING PLAN
    # ====================
    st.markdown("---")
    st.markdown("### 🎯 Shipping Plan (To Be Shipped)")
    st.caption("Build your plan with orders and deals that will ship this quarter")
    
    planning_col1, planning_col2 = st.columns([3, 1])
    
    with planning_col1:
        # Fulfilled Not Shipped
        fns_key = 'Fulfilled Not Shipped'
        st.session_state.planning_sources[fns_key] = st.checkbox(
            f"📋 Fulfilled, Not Shipped: ${metrics['fulfilled_not_shipped']:,.0f}",
            value=st.session_state.planning_sources.get(fns_key, False),
            key=f"planning_{fns_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[fns_key] and sales_orders_df is not None and not sales_orders_df.empty:
            sample_sos = sales_orders_df[sales_orders_df['Status'] == 'Fulfilled, Not Shipped'].head(3)
            if not sample_sos.empty:
                st.caption(f"📄 Sample: {', '.join(sample_sos['Document Number'].astype(str).tolist())}")
        
        # Pending Fulfillment
        pf_key = 'Pending Fulfillment'
        st.session_state.planning_sources[pf_key] = st.checkbox(
            f"⏳ Pending Fulfillment: ${metrics['pending_fulfillment']:,.0f}",
            value=st.session_state.planning_sources.get(pf_key, False),
            key=f"planning_{pf_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[pf_key] and sales_orders_df is not None and not sales_orders_df.empty:
            sample_sos = sales_orders_df[sales_orders_df['Status'] == 'Pending Fulfillment'].head(3)
            if not sample_sos.empty:
                st.caption(f"📄 Sample: {', '.join(sample_sos['Document Number'].astype(str).tolist())}")
        
        # Pending Approval
        pa_key = 'Pending Approval'
        st.session_state.planning_sources[pa_key] = st.checkbox(
            f"⏸️ Pending Approval: ${metrics['pending_approval']:,.0f}",
            value=st.session_state.planning_sources.get(pa_key, False),
            key=f"planning_{pa_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[pa_key] and sales_orders_df is not None and not sales_orders_df.empty:
            sample_sos = sales_orders_df[sales_orders_df['Status'] == 'Pending Approval'].head(3)
            if not sample_sos.empty:
                st.caption(f"📄 Sample: {', '.join(sample_sos['Document Number'].astype(str).tolist())}")
        
        # Old Pending Approval
        opa_key = 'Pending Approval (>2 weeks old)'
        if metrics['old_pending_approval'] > 0:
            st.session_state.planning_sources[opa_key] = st.checkbox(
                f"⚠️ Pending Approval (>2 weeks): ${metrics['old_pending_approval']:,.0f}",
                value=st.session_state.planning_sources.get(opa_key, False),
                key=f"planning_{opa_key}_{rep_name or 'team'}"
            )
            if st.session_state.planning_sources[opa_key] and sales_orders_df is not None and not sales_orders_df.empty:
                sample_sos = sales_orders_df[
                    (sales_orders_df['Status'] == 'Pending Approval') &
                    (sales_orders_df['Age_Business_Days'] >= 10)
                ].head(3)
                if not sample_sos.empty:
                    st.caption(f"📄 Sample: {', '.join(sample_sos['Document Number'].astype(str).tolist())}")
        
        st.markdown("**HubSpot Deals:**")
        
        # HubSpot Expect
        hs_exp_key = 'HubSpot Expect'
        st.session_state.planning_sources[hs_exp_key] = st.checkbox(
            f"🎯 HubSpot Expect: ${metrics['hubspot_expect']:,.0f}",
            value=st.session_state.planning_sources.get(hs_exp_key, False),
            key=f"planning_{hs_exp_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[hs_exp_key] and deals_df is not None and not deals_df.empty:
            sample_deals = deals_df[deals_df['Status'] == 'Expect'].head(3)
            if not sample_deals.empty:
                st.caption(f"📄 Sample: {', '.join(sample_deals['Record ID'].astype(str).tolist())}")
        
        # HubSpot Commit
        hs_com_key = 'HubSpot Commit'
        st.session_state.planning_sources[hs_com_key] = st.checkbox(
            f"💪 HubSpot Commit: ${metrics['hubspot_commit']:,.0f}",
            value=st.session_state.planning_sources.get(hs_com_key, False),
            key=f"planning_{hs_com_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[hs_com_key] and deals_df is not None and not deals_df.empty:
            sample_deals = deals_df[deals_df['Status'] == 'Commit'].head(3)
            if not sample_deals.empty:
                st.caption(f"📄 Sample: {', '.join(sample_deals['Record ID'].astype(str).tolist())}")
        
        # HubSpot Best Case
        hs_bc_key = 'HubSpot Best Case'
        st.session_state.planning_sources[hs_bc_key] = st.checkbox(
            f"🌟 HubSpot Best Case: ${metrics['hubspot_best_case']:,.0f}",
            value=st.session_state.planning_sources.get(hs_bc_key, False),
            key=f"planning_{hs_bc_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[hs_bc_key] and deals_df is not None and not deals_df.empty:
            sample_deals = deals_df[deals_df['Status'] == 'Best Case'].head(3)
            if not sample_deals.empty:
                st.caption(f"📄 Sample: {', '.join(sample_deals['Record ID'].astype(str).tolist())}")
        
        # HubSpot Opportunity
        hs_opp_key = 'HubSpot Opportunity'
        st.session_state.planning_sources[hs_opp_key] = st.checkbox(
            f"💎 HubSpot Opportunity: ${metrics['hubspot_opportunity']:,.0f}",
            value=st.session_state.planning_sources.get(hs_opp_key, False),
            key=f"planning_{hs_opp_key}_{rep_name or 'team'}"
        )
        if st.session_state.planning_sources[hs_opp_key] and deals_df is not None and not deals_df.empty:
            sample_deals = deals_df[deals_df['Status'] == 'Opportunity'].head(3)
            if not sample_deals.empty:
                st.caption(f"📄 Sample: {', '.join(sample_deals['Record ID'].astype(str).tolist())}")
        
        # Q1 Spillover
        q1_key = 'Q1 Spillover - Expect/Commit'
        if metrics['q1_spillover'] > 0:
            st.session_state.planning_sources[q1_key] = st.checkbox(
                f"📅 Q1 2026 Spillover (Expect/Commit): ${metrics['q1_spillover']:,.0f}",
                value=st.session_state.planning_sources.get(q1_key, False),
                key=f"planning_{q1_key}_{rep_name or 'team'}"
            )
            if st.session_state.planning_sources[q1_key] and deals_df is not None and not deals_df.empty:
                sample_deals = deals_df[
                    (deals_df.get('Counts_In_Q4', True) == False) &
                    (deals_df['Status'].isin(['Expect', 'Commit']))
                ].head(3)
                if not sample_deals.empty:
                    st.caption(f"📄 Sample: {', '.join(sample_deals['Record ID'].astype(str).tolist())}")
    
    with planning_col2:
        planning_total = sum([
            metrics['fulfilled_not_shipped'] if st.session_state.planning_sources.get('Fulfilled Not Shipped', False) else 0,
            metrics['pending_fulfillment'] if st.session_state.planning_sources.get('Pending Fulfillment', False) else 0,
            metrics['pending_approval'] if st.session_state.planning_sources.get('Pending Approval', False) else 0,
            metrics['old_pending_approval'] if st.session_state.planning_sources.get('Pending Approval (>2 weeks old)', False) else 0,
            metrics['hubspot_expect'] if st.session_state.planning_sources.get('HubSpot Expect', False) else 0,
            metrics['hubspot_commit'] if st.session_state.planning_sources.get('HubSpot Commit', False) else 0,
            metrics['hubspot_best_case'] if st.session_state.planning_sources.get('HubSpot Best Case', False) else 0,
            metrics['hubspot_opportunity'] if st.session_state.planning_sources.get('HubSpot Opportunity', False) else 0,
            metrics['q1_spillover'] if st.session_state.planning_sources.get('Q1 Spillover - Expect/Commit', False) else 0
        ])
        
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, #3498DB 0%, #2980B9 100%); 
                     padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 14px; margin-bottom: 5px;'>Plan Total</div>
            <div style='color: white; font-size: 28px; font-weight: bold;'>${planning_total:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    # ====================
    # SUMMARY
    # ====================
    st.markdown("---")
    st.markdown("### 📊 Summary")
    
    total_forecast = shipped_total + planning_total
    gap_to_quota = quota - total_forecast
    gap_percentage = (gap_to_quota / quota * 100) if quota > 0 else 0
    
    sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
    
    with sum_col1:
        st.markdown(f"""
        <div style='background: #2ECC71; padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 12px;'>Shipped</div>
            <div style='color: white; font-size: 24px; font-weight: bold;'>${shipped_total:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with sum_col2:
        st.markdown(f"""
        <div style='background: #3498DB; padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 12px;'>Plan</div>
            <div style='color: white; font-size: 24px; font-weight: bold;'>${planning_total:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with sum_col3:
        st.markdown(f"""
        <div style='background: #9B59B6; padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 12px;'>Total Forecast</div>
            <div style='color: white; font-size: 24px; font-weight: bold;'>${total_forecast:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with sum_col4:
        gap_color = '#E74C3C' if gap_to_quota > 0 else '#2ECC71'
        st.markdown(f"""
        <div style='background: {gap_color}; padding: 15px; border-radius: 10px; text-align: center;'>
            <div style='color: white; font-size: 12px;'>Gap to Quota</div>
            <div style='color: white; font-size: 24px; font-weight: bold;'>${gap_to_quota:,.0f}</div>
            <div style='color: white; font-size: 11px;'>({gap_percentage:.1f}%)</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Export functionality (simplified version without individual selection)
    st.markdown("---")
    if st.button("📥 Export Shipping Plan", type="primary"):
        export_data = []
        
        # Add shipped items
        if st.session_state.shipped_sources.get('Invoices', False) and invoices_df is not None and not invoices_df.empty:
            for _, row in invoices_df.iterrows():
                export_data.append({
                    'Category': 'Shipped',
                    'Type': 'Invoice',
                    'Document Number': row.get('Invoice Number', ''),
                    'Customer': row.get('Customer', ''),
                    'Amount': row.get('Amount', 0),
                    'Date': row.get('Date', ''),
                    'Sales Rep': row.get('Sales Rep', '')
                })
        
        if st.session_state.shipped_sources.get('Shipped Not Invoiced', False) and sales_orders_df is not None and not sales_orders_df.empty:
            for _, row in sales_orders_df[sales_orders_df['Status'] == 'Shipped, Not Invoiced'].iterrows():
                export_data.append({
                    'Category': 'Shipped',
                    'Type': 'Sales Order',
                    'Document Number': row.get('Document Number', ''),
                    'Customer': row.get('Customer', ''),
                    'Amount': row.get('Amount', 0),
                    'Date': row.get('Order Start Date', ''),
                    'Sales Rep': row.get('Sales Rep', '')
                })
        
        # Add planning items (Sales Orders)
        for status_key, status_value in [
            ('Fulfilled Not Shipped', 'Fulfilled, Not Shipped'),
            ('Pending Fulfillment', 'Pending Fulfillment'),
            ('Pending Approval', 'Pending Approval')
        ]:
            if st.session_state.planning_sources.get(status_key, False) and sales_orders_df is not None and not sales_orders_df.empty:
                for _, row in sales_orders_df[sales_orders_df['Status'] == status_value].iterrows():
                    export_data.append({
                        'Category': 'Shipping Plan',
                        'Type': 'Sales Order',
                        'Document Number': row.get('Document Number', ''),
                        'Customer': row.get('Customer', ''),
                        'Amount': row.get('Amount', 0),
                        'Date': row.get('Order Start Date', ''),
                        'Sales Rep': row.get('Sales Rep', '')
                    })
        
        # Old Pending Approval
        if st.session_state.planning_sources.get('Pending Approval (>2 weeks old)', False) and sales_orders_df is not None and not sales_orders_df.empty:
            for _, row in sales_orders_df[
                (sales_orders_df['Status'] == 'Pending Approval') &
                (sales_orders_df['Age_Business_Days'] >= 10)
            ].iterrows():
                export_data.append({
                    'Category': 'Shipping Plan',
                    'Type': 'Sales Order - Old PA',
                    'Document Number': row.get('Document Number', ''),
                    'Customer': row.get('Customer', ''),
                    'Amount': row.get('Amount', 0),
                    'Date': row.get('Order Start Date', ''),
                    'Sales Rep': row.get('Sales Rep', '')
                })
        
        # Add planning items (HubSpot)
        for status_key, status_value in [
            ('HubSpot Expect', 'Expect'),
            ('HubSpot Commit', 'Commit'),
            ('HubSpot Best Case', 'Best Case'),
            ('HubSpot Opportunity', 'Opportunity')
        ]:
            if st.session_state.planning_sources.get(status_key, False) and deals_df is not None and not deals_df.empty:
                for _, row in deals_df[deals_df['Status'] == status_value].iterrows():
                    export_data.append({
                        'Category': 'Shipping Plan',
                        'Type': f'HubSpot - {status_value}',
                        'Document Number': row.get('Record ID', ''),
                        'Customer': row.get('Deal Name', ''),
                        'Amount': row.get('Amount', 0),
                        'Date': row.get('Close Date', ''),
                        'Sales Rep': row.get('Deal Owner', '')
                    })
        
        # Q1 Spillover
        if st.session_state.planning_sources.get('Q1 Spillover - Expect/Commit', False) and deals_df is not None and not deals_df.empty:
            for _, row in deals_df[
                (deals_df.get('Counts_In_Q4', True) == False) &
                (deals_df['Status'].isin(['Expect', 'Commit']))
            ].iterrows():
                export_data.append({
                    'Category': 'Shipping Plan',
                    'Type': 'HubSpot - Q1 Spillover',
                    'Document Number': row.get('Record ID', ''),
                    'Customer': row.get('Deal Name', ''),
                    'Amount': row.get('Amount', 0),
                    'Date': row.get('Close Date', ''),
                    'Sales Rep': row.get('Deal Owner', '')
                })
        
        if export_data:
            export_df = pd.DataFrame(export_data)
            export_df['Amount'] = export_df['Amount'].apply(lambda x: f"${x:,.0f}" if pd.notna(x) else "$0")
            
            csv_data = export_df.to_csv(index=False)
            
            st.download_button(
                label="📥 Download Shipping Plan",
                data=csv_data,
                file_name=f"shipping_plan_{'team' if rep_name is None else rep_name}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
            
            st.success(f"✅ Export ready with {len(export_df)} line items")
        else:
            st.warning("No items selected for export")


def main():
    st.markdown("""
    <div style='text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                 color: white; border-radius: 10px; margin-bottom: 20px;'>
        <h3>📦 Q4 2025 Shipping Planning</h3>
        <p style='font-size: 14px; margin: 0;'>Build Your Shipping Plan - Shipped vs To Be Shipped</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Load data
    with st.spinner("🔄 Loading data..."):
        deals_df, dashboard_df, invoices_df, sales_orders_df = load_all_data()
    
    if deals_df.empty and dashboard_df.empty:
        st.error("❌ Unable to load data")
        return
    
    # Calculate metrics
    metrics = calculate_team_metrics(deals_df, dashboard_df, invoices_df, sales_orders_df)
    
    # Get total quota
    quota = dashboard_df['Quota'].sum() if not dashboard_df.empty else 5_021_440
    
    # Call the shipping plan section
    build_shipping_plan_section(
        metrics=metrics,
        quota=quota,
        rep_name=None,  # Team view
        deals_df=deals_df,
        invoices_df=invoices_df,
        sales_orders_df=sales_orders_df
    )

if __name__ == "__main__":
    main()
