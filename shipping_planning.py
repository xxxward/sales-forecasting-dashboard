"""
Q4 2025 Shipping Planning Tool
===============================
Standalone planning tool for operations team to build and track Q4 shipping plans.
This is a SANDBOX version for testing - completely separate from main sales dashboard.

Team: Xander (RevOps), Kyle (Sales VP), Cory (Supply Chain), Greg (Production), Tif (CX)
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
import gspread
from datetime import datetime, timedelta
from io import BytesIO
import json

# Page configuration
st.set_page_config(
    page_title="Q4 Shipping Planning",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .planning-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: rgba(240, 242, 246, 0.5);
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #667eea;
    }
    .conflict-warning {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
    </style>
""", unsafe_allow_html=True)

# Configuration
SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CACHE_TTL = 300  # 5 minutes
CACHE_VERSION = "v1_shipping_plan"

# Team members for assignment dropdown
TEAM_MEMBERS = ["Xander", "Kyle", "Cory", "Greg", "Tif", "Unassigned"]

# ============================================================================
# GOOGLE SHEETS CONNECTION
# ============================================================================

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name, range_name, version=CACHE_VERSION):
    """Load data from Google Sheets with caching"""
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ Missing Google Cloud credentials")
            return pd.DataFrame()
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(
            creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly']
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
        st.error(f"❌ Error loading data from {sheet_name}: {str(e)}")
        return pd.DataFrame()

def get_gspread_client():
    """Get gspread client for writing to sheets"""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
        return gc
    except Exception as e:
        st.error(f"❌ Error connecting to Google Sheets: {str(e)}")
        return None

# ============================================================================
# DATA LOADING AND PROCESSING
# ============================================================================

def clean_numeric(value):
    """Clean and convert values to numeric"""
    if pd.isna(value) or str(value).strip() == '':
        return 0
    cleaned = str(value).replace(',', '').replace('$', '').replace(' ', '').strip()
    try:
        return float(cleaned)
    except:
        return 0

def load_sales_orders():
    """Load and process Sales Orders data"""
    df = load_google_sheets_data("NS Sales Orders", "A:AD", version=CACHE_VERSION)
    
    if df.empty:
        return pd.DataFrame()
    
    # Rename columns based on your existing logic
    col_names = df.columns.tolist()
    rename_dict = {}
    
    # Map Internal ID (Column A) - note it has a SPACE not underscore
    if len(col_names) > 0:
        col_a_name = str(col_names[0])
        if 'internal' in col_a_name.lower() and 'id' in col_a_name.lower():
            rename_dict[col_names[0]] = 'Internal_ID'
    
    # Map SO Number (Column B)
    if len(col_names) > 1:
        col_b_name = str(col_names[1])
        if 'so' in col_b_name.lower() and 'number' in col_b_name.lower():
            rename_dict[col_names[1]] = 'Document_Number'
    
    # Find standard columns - look for fuzzy matches
    for idx, col in enumerate(col_names):
        col_lower = str(col).lower()
        if 'status' in col_lower and 'Status' not in rename_dict.values():
            rename_dict[col] = 'Status'
        elif 'customer' in col_lower and 'promise' not in col_lower and 'external' not in col_lower and 'Customer' not in rename_dict.values():
            rename_dict[col] = 'Customer'
        elif ('amount' in col_lower or 'total' in col_lower) and 'Amount' not in rename_dict.values():
            rename_dict[col] = 'Amount'
        elif ('sales rep' in col_lower or 'salesrep' in col_lower) and 'Sales_Rep' not in rename_dict.values():
            rename_dict[col] = 'Sales_Rep'
    
    # Map specific date columns by position
    if len(col_names) > 8:
        rename_dict[col_names[8]] = 'Order_Start_Date'
    if len(col_names) > 11:
        rename_dict[col_names[11]] = 'Customer_Promise_Date'
    if len(col_names) > 12:
        rename_dict[col_names[12]] = 'Projected_Date'
    if len(col_names) > 27:
        rename_dict[col_names[27]] = 'Pending_Approval_Date'
    
    # Map Rep Master and Corrected Customer Name
    if len(col_names) > 28:
        rename_dict[col_names[28]] = 'Corrected_Customer_Name'
    if len(col_names) > 29:
        rename_dict[col_names[29]] = 'Rep_Master'
    
    df = df.rename(columns=rename_dict)
    
    # Apply Rep Master override
    if 'Rep_Master' in df.columns:
        df['Sales_Rep'] = df['Rep_Master']
        df = df.drop(columns=['Rep_Master'])
    
    if 'Corrected_Customer_Name' in df.columns:
        df['Customer'] = df['Corrected_Customer_Name']
        df = df.drop(columns=['Corrected_Customer_Name'])
    
    # Clean and convert
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
    
    if 'Sales_Rep' in df.columns:
        df['Sales_Rep'] = df['Sales_Rep'].astype(str).str.strip()
    
    if 'Status' in df.columns:
        df['Status'] = df['Status'].astype(str).str.strip()
    
    # Convert dates
    date_columns = ['Order_Start_Date', 'Customer_Promise_Date', 'Projected_Date', 'Pending_Approval_Date']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Filter to relevant statuses
    if 'Status' in df.columns:
        df = df[df['Status'].isin(['Pending Approval', 'Pending Fulfillment', 'Pending Billing/Partially Fulfilled'])]
    
    # Remove invalid rows
    if 'Amount' in df.columns and 'Sales_Rep' in df.columns:
        df = df[
            (df['Amount'] > 0) & 
            (df['Sales_Rep'].notna()) & 
            (df['Sales_Rep'] != '') &
            (df['Sales_Rep'] != 'nan') &
            (~df['Sales_Rep'].str.lower().isin(['house']))
        ]
    
    # Add record metadata
    df['Record_Type'] = 'Sales_Order'
    
    # Try multiple possible ID columns in order of preference
    if 'Document_Number' in df.columns:
        df['Record_ID'] = df['Document_Number'].astype(str)
    elif 'Internal_ID' in df.columns:
        df['Record_ID'] = df['Internal_ID'].astype(str)
    elif len(df.columns) > 0:
        # Use the first column as fallback (usually the ID)
        df['Record_ID'] = df.iloc[:, 0].astype(str)
    else:
        df['Record_ID'] = 'SO_' + df.index.astype(str)
    
    return df

def load_invoices():
    """Load and process Invoices data"""
    df = load_google_sheets_data("NS Invoices", "A:U", version=CACHE_VERSION)
    
    if df.empty:
        return pd.DataFrame()
    
    # Rename columns
    rename_dict = {
        df.columns[0]: 'Invoice_Number',
        df.columns[1]: 'Status',
        df.columns[2]: 'Date',
        df.columns[6]: 'Customer',
        df.columns[10]: 'Amount',
        df.columns[14]: 'Sales_Rep'
    }
    
    if len(df.columns) > 19:
        rename_dict[df.columns[19]] = 'Corrected_Customer_Name'
    if len(df.columns) > 20:
        rename_dict[df.columns[20]] = 'Rep_Master'
    
    df = df.rename(columns=rename_dict)
    
    # Apply overrides
    if 'Rep_Master' in df.columns:
        df['Rep_Master'] = df['Rep_Master'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        mask = df['Rep_Master'].isin(invalid_values)
        df.loc[~mask, 'Sales_Rep'] = df.loc[~mask, 'Rep_Master']
        df = df.drop(columns=['Rep_Master'])
    
    if 'Corrected_Customer_Name' in df.columns:
        df['Corrected_Customer_Name'] = df['Corrected_Customer_Name'].astype(str).str.strip()
        invalid_values = ['', 'nan', 'None', '#N/A', '#REF!', '#VALUE!', '#ERROR!']
        mask = df['Corrected_Customer_Name'].isin(invalid_values)
        df.loc[~mask, 'Customer'] = df.loc[~mask, 'Corrected_Customer_Name']
        df = df.drop(columns=['Corrected_Customer_Name'])
    
    # Clean data
    df['Amount'] = df['Amount'].apply(clean_numeric)
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # Filter to Q4 2025
    q4_start = pd.Timestamp('2025-10-01')
    q4_end = pd.Timestamp('2025-12-31')
    df = df[(df['Date'] >= q4_start) & (df['Date'] <= q4_end)]
    
    # Clean Sales Rep
    df['Sales_Rep'] = df['Sales_Rep'].astype(str).str.strip()
    df = df[
        (df['Sales_Rep'].notna()) & 
        (df['Sales_Rep'] != '') &
        (df['Sales_Rep'].str.lower() != 'nan') &
        (df['Sales_Rep'].str.lower() != 'house')
    ]
    
    # Remove duplicates
    if 'Invoice_Number' in df.columns:
        df = df.drop_duplicates(subset=['Invoice_Number'], keep='first')
    
    # Add record metadata
    df['Record_Type'] = 'Invoice'
    
    # Use Invoice_Number if available, otherwise first column
    if 'Invoice_Number' in df.columns:
        df['Record_ID'] = df['Invoice_Number'].astype(str)
    elif len(df.columns) > 0:
        df['Record_ID'] = df.iloc[:, 0].astype(str)
    else:
        df['Record_ID'] = 'INV_' + df.index.astype(str)
    
    return df

def load_hubspot_deals():
    """Load and process HubSpot Deals data"""
    df = load_google_sheets_data("All Reps All Pipelines", "A:R", version=CACHE_VERSION)
    
    if df.empty:
        return pd.DataFrame()
    
    # Rename columns
    col_names = df.columns.tolist()
    rename_dict = {}
    
    for col in col_names:
        if col == 'Record ID':
            rename_dict[col] = 'Deal_ID'
        elif col == 'Deal Name':
            rename_dict[col] = 'Deal_Name'
        elif col == 'Deal Stage':
            rename_dict[col] = 'Deal_Stage'
        elif col == 'Close Date':
            rename_dict[col] = 'Close_Date'
        elif 'Deal Owner First Name' in col and 'Deal Owner Last Name' in col:
            rename_dict[col] = 'Deal_Owner'
        elif col == 'Deal Owner First Name':
            rename_dict[col] = 'Deal_Owner_First'
        elif col == 'Deal Owner Last Name':
            rename_dict[col] = 'Deal_Owner_Last'
        elif col == 'Amount':
            rename_dict[col] = 'Amount'
        elif col == 'Close Status':
            rename_dict[col] = 'Status'
        elif col == 'Pipeline':
            rename_dict[col] = 'Pipeline'
        elif col == 'Deal Type':
            rename_dict[col] = 'Product_Type'
    
    df = df.rename(columns=rename_dict)
    
    # Create Deal Owner if needed
    if 'Deal_Owner' not in df.columns:
        if 'Deal_Owner_First' in df.columns and 'Deal_Owner_Last' in df.columns:
            df['Deal_Owner'] = df['Deal_Owner_First'].fillna('') + ' ' + df['Deal_Owner_Last'].fillna('')
            df['Deal_Owner'] = df['Deal_Owner'].str.strip()
    else:
        df['Deal_Owner'] = df['Deal_Owner'].str.strip()
    
    # Clean Amount
    if 'Amount' in df.columns:
        df['Amount'] = df['Amount'].apply(clean_numeric)
    
    # Convert Close Date
    if 'Close_Date' in df.columns:
        df['Close_Date'] = pd.to_datetime(df['Close_Date'], errors='coerce')
    
    # Filter to Q4 2025
    q4_start = pd.Timestamp('2025-10-01')
    q4_end = pd.Timestamp('2026-01-01')
    
    if 'Close_Date' in df.columns:
        df = df[(df['Close_Date'] >= q4_start) & (df['Close_Date'] < q4_end)]
    
    # Filter out excluded stages
    excluded_stages = ['', '(Blanks)', None, 'Cancelled', 'checkout abandoned', 
                       'closed lost', 'closed won', 'sales order created in NS', 'NCR', 'Shipped']
    
    if 'Deal_Stage' in df.columns:
        df['Deal_Stage'] = df['Deal_Stage'].fillna('')
        df['Deal_Stage'] = df['Deal_Stage'].astype(str).str.strip()
        df = df[~df['Deal_Stage'].str.lower().isin([s.lower() if s else '' for s in excluded_stages])]
    
    # Add record metadata
    df['Record_Type'] = 'HubSpot_Deal'
    
    if 'Deal_ID' in df.columns:
        df['Record_ID'] = df['Deal_ID'].astype(str)
    elif len(df.columns) > 0:
        df['Record_ID'] = df.iloc[:, 0].astype(str)
    else:
        df['Record_ID'] = 'DEAL_' + df.index.astype(str)
    
    return df

def categorize_records(row):
    """Categorize each record into planning buckets"""
    if row['Record_Type'] == 'Invoice':
        return 'Invoiced_Shipped'
    
    elif row['Record_Type'] == 'Sales_Order':
        if row.get('Status') == 'Pending Fulfillment':
            # Determine expected date
            expected_date = row.get('Customer_Promise_Date') or row.get('Projected_Date')
            if pd.notna(expected_date):
                return 'PF_With_Date'
            else:
                return 'PF_No_Date'
        
        elif row.get('Status') == 'Pending Approval':
            # Check for date and age
            expected_date = row.get('Pending_Approval_Date') or row.get('Customer_Promise_Date')
            if pd.notna(expected_date):
                # Calculate age
                age_days = (pd.Timestamp.now() - expected_date).days
                if age_days > 14:
                    return 'PA_Old_Date'
                else:
                    return 'PA_With_Date'
            else:
                return 'PA_No_Date'
    
    elif row['Record_Type'] == 'HubSpot_Deal':
        status = row.get('Status', '')
        close_date = row.get('Close_Date')
        
        # Check if Q1 spillover
        is_q1 = False
        if pd.notna(close_date) and close_date >= pd.Timestamp('2025-01-01'):
            is_q1 = True
        
        if status in ['Commit', 'Expect']:
            return 'Q1_Spillover_Expect' if is_q1 else 'HubSpot_Expect'
        elif status == 'Best Case':
            return 'Q1_Spillover_Best_Case' if is_q1 else 'HubSpot_Best_Case'
        elif status == 'Opportunity':
            return 'HubSpot_Opportunity'
    
    return 'Other'

def load_all_planning_data():
    """Load all data sources and combine for planning"""
    
    with st.spinner("🔄 Loading data from NetSuite and HubSpot..."):
        # Load each source
        so_df = load_sales_orders()
        inv_df = load_invoices()
        deals_df = load_hubspot_deals()
        
        # Standardize column names for each source
        if not so_df.empty:
            # Remove duplicate columns first
            so_df = so_df.loc[:, ~so_df.columns.duplicated()]
            
            so_df = so_df.rename(columns={
                'Amount': 'Live_Amount',
                'Customer_Promise_Date': 'Live_Expected_Date',
                'Status': 'Live_Status',
                'Sales_Rep': 'Sales_Rep',
                'Customer': 'Customer'
            })
            # Use Customer Promise Date or Projected Date as expected
            if 'Live_Expected_Date' not in so_df.columns or so_df['Live_Expected_Date'].isna().all():
                if 'Projected_Date' in so_df.columns:
                    so_df['Live_Expected_Date'] = so_df['Projected_Date']
        
        if not inv_df.empty:
            # Remove duplicate columns first
            inv_df = inv_df.loc[:, ~inv_df.columns.duplicated()]
            
            inv_df = inv_df.rename(columns={
                'Amount': 'Live_Amount',
                'Date': 'Live_Expected_Date',
                'Status': 'Live_Status',
                'Sales_Rep': 'Sales_Rep',
                'Customer': 'Customer'
            })
        
        if not deals_df.empty:
            # Remove duplicate columns first
            deals_df = deals_df.loc[:, ~deals_df.columns.duplicated()]
            
            deals_df = deals_df.rename(columns={
                'Amount': 'Live_Amount',
                'Close_Date': 'Live_Expected_Date',
                'Status': 'Live_Status',
                'Deal_Owner': 'Sales_Rep',
                'Deal_Name': 'Customer'  # Use Deal Name as customer identifier
            })
        
        # Combine all sources
        all_records = []
        
        if not so_df.empty:
            all_records.append(so_df)
        if not inv_df.empty:
            all_records.append(inv_df)
        if not deals_df.empty:
            all_records.append(deals_df)
        
        if not all_records:
            return pd.DataFrame()
        
        try:
            combined_df = pd.concat(all_records, ignore_index=True)
        except Exception as e:
            st.error(f"Error during concat: {str(e)}")
            st.error("Trying to fix by selecting only needed columns...")
            
            # Select only the columns we absolutely need
            needed_cols = ['Record_Type', 'Record_ID', 'Sales_Rep', 'Customer', 
                          'Live_Amount', 'Live_Expected_Date', 'Live_Status']
            
            fixed_records = []
            for df in all_records:
                # Keep only columns that exist and are needed
                cols_to_keep = [c for c in needed_cols if c in df.columns]
                fixed_records.append(df[cols_to_keep])
            
            combined_df = pd.concat(fixed_records, ignore_index=True)
        
        # Add category
        combined_df['Category'] = combined_df.apply(categorize_records, axis=1)
        
        # Initialize planning columns
        combined_df['Override_Amount'] = None
        combined_df['Override_Ship_Date'] = None
        combined_df['Include_Flag'] = False
        combined_df['Planning_Notes'] = ''
        combined_df['Confidence_Level'] = 'Medium'
        combined_df['Assigned_To'] = 'Unassigned'
        combined_df['Live_Last_Updated'] = datetime.now()
        
        # AUTO-INCLUDE baseline items (Invoiced & Shipped)
        combined_df.loc[combined_df['Category'] == 'Invoiced_Shipped', 'Include_Flag'] = True
        
        return combined_df

# ============================================================================
# SAVING AND LOADING PLANS
# ============================================================================

def save_shipping_plan_to_sheet(version_name, planning_df):
    """Save shipping plan to Google Sheets"""
    
    try:
        gc = get_gspread_client()
        if not gc:
            st.error("Could not connect to Google Sheets")
            return False
        
        sh = gc.open_by_key(SPREADSHEET_ID)
        
        # Get or create Shipping_Plan_Active worksheet
        try:
            worksheet = sh.worksheet('Shipping_Plan_Active')
        except:
            worksheet = sh.add_worksheet('Shipping_Plan_Active', rows=1000, cols=30)
            # Add headers
            headers = [
                'Version_ID', 'Version_Name', 'Record_Type', 'Record_ID', 'Sales_Rep', 'Customer',
                'Live_Amount', 'Live_Expected_Date', 'Live_Status', 'Live_Last_Updated',
                'Override_Amount', 'Override_Ship_Date', 'Include_Flag', 'Planning_Notes',
                'Confidence_Level', 'Assigned_To', 'Category',
                'Saved_By', 'Saved_Date'
            ]
            worksheet.append_row(headers)
        
        # Generate version ID
        version_id = f"V{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        # Filter to records with includes or overrides
        records_to_save = planning_df[
            (planning_df['Include_Flag'] == True) |
            (pd.notna(planning_df['Override_Amount'])) |
            (pd.notna(planning_df['Override_Ship_Date'])) |
            (planning_df['Planning_Notes'] != '')
        ].copy()
        
        if records_to_save.empty:
            st.warning("⚠️ No records selected to save")
            return False
        
        # Prepare data
        records_to_save['Version_ID'] = version_id
        records_to_save['Version_Name'] = version_name
        records_to_save['Saved_By'] = st.session_state.get('user_name', 'Unknown')
        records_to_save['Saved_Date'] = datetime.now().isoformat()
        
        # Convert to list of lists
        save_columns = [
            'Version_ID', 'Version_Name', 'Record_Type', 'Record_ID', 'Sales_Rep', 'Customer',
            'Live_Amount', 'Live_Expected_Date', 'Live_Status', 'Live_Last_Updated',
            'Override_Amount', 'Override_Ship_Date', 'Include_Flag', 'Planning_Notes',
            'Confidence_Level', 'Assigned_To', 'Category',
            'Saved_By', 'Saved_Date'
        ]
        
        # Ensure all columns exist
        for col in save_columns:
            if col not in records_to_save.columns:
                records_to_save[col] = ''
        
        # Convert dates to strings
        for col in ['Live_Expected_Date', 'Override_Ship_Date', 'Live_Last_Updated', 'Saved_Date']:
            if col in records_to_save.columns:
                records_to_save[col] = records_to_save[col].astype(str)
        
        # Convert to list of lists
        data_to_save = records_to_save[save_columns].values.tolist()
        
        # Append to sheet
        worksheet.append_rows(data_to_save)
        
        st.session_state['last_saved_version'] = version_id
        st.session_state['last_saved_name'] = version_name
        
        return True
        
    except Exception as e:
        st.error(f"❌ Error saving to Google Sheets: {str(e)}")
        return False

def load_saved_plan_versions():
    """Load list of saved plan versions from Google Sheets"""
    try:
        gc = get_gspread_client()
        if not gc:
            return []
        
        sh = gc.open_by_key(SPREADSHEET_ID)
        
        try:
            worksheet = sh.worksheet('Shipping_Plan_Active')
            data = worksheet.get_all_records()
            
            if not data:
                return []
            
            df = pd.DataFrame(data)
            
            # Get unique versions
            if 'Version_Name' in df.columns and 'Version_ID' in df.columns:
                versions = df[['Version_ID', 'Version_Name', 'Saved_Date']].drop_duplicates()
                versions = versions.sort_values('Saved_Date', ascending=False)
                return versions['Version_Name'].tolist()
            
            return []
            
        except:
            return []
            
    except Exception as e:
        return []

# ============================================================================
# UI COMPONENTS
# ============================================================================

def display_planning_section(category_name, category_df, key_suffix):
    """Display an editable planning section for a category"""
    
    if category_df.empty:
        st.info(f"No {category_name} records")
        return category_df
    
    st.markdown(f"#### {category_name}")
    st.caption(f"{len(category_df)} records | ${category_df['Live_Amount'].sum():,.0f} total")
    
    # Prepare display dataframe
    display_df = category_df[[
        'Record_Type', 'Record_ID', 'Sales_Rep', 'Customer',
        'Live_Amount', 'Live_Expected_Date',
        'Override_Amount', 'Override_Ship_Date',
        'Include_Flag', 'Confidence_Level', 'Assigned_To', 'Planning_Notes'
    ]].copy()
    
    # Data editor
    edited_df = st.data_editor(
        display_df,
        column_config={
            "Record_Type": st.column_config.TextColumn("Type", disabled=True, width="small"),
            "Record_ID": st.column_config.TextColumn("SO/Invoice/Deal #", disabled=True, width="medium"),
            "Sales_Rep": st.column_config.TextColumn("Sales Rep", disabled=True, width="medium"),
            "Customer": st.column_config.TextColumn("Customer", disabled=True, width="medium"),
            "Live_Amount": st.column_config.NumberColumn(
                "Live Amount",
                format="$%d",
                disabled=True,
                width="small"
            ),
            "Live_Expected_Date": st.column_config.DateColumn(
                "Source Date",
                disabled=True,
                width="small"
            ),
            "Override_Amount": st.column_config.NumberColumn(
                "Override $",
                format="$%d",
                help="Override amount (leave blank to use live)",
                width="small"
            ),
            "Override_Ship_Date": st.column_config.DateColumn(
                "Override Date",
                help="Override ship date for planning",
                width="small"
            ),
            "Include_Flag": st.column_config.CheckboxColumn("Include?", width="small"),
            "Confidence_Level": st.column_config.SelectboxColumn(
                "Confidence",
                options=["High", "Medium", "Low"],
                width="small"
            ),
            "Assigned_To": st.column_config.SelectboxColumn(
                "Assigned",
                options=TEAM_MEMBERS,
                width="small"
            ),
            "Planning_Notes": st.column_config.TextColumn(
                "Notes",
                help="Planning notes and status updates",
                width="large"
            ),
        },
        hide_index=True,
        use_container_width=True,
        key=f"editor_{key_suffix}"
    )
    
    # Merge edits back to original dataframe
    category_df.update(edited_df)
    
    return category_df

def calculate_plan_summary(planning_df):
    """Calculate summary metrics for the current plan"""
    
    # Filter to included items
    included_df = planning_df[planning_df['Include_Flag'] == True].copy()
    
    if included_df.empty:
        return {
            'total_records': 0,
            'total_amount': 0,
            'by_category': {},
            'by_rep': {},
            'by_confidence': {}
        }
    
    # Calculate effective amounts
    included_df['Effective_Amount'] = included_df.apply(
        lambda row: row['Override_Amount'] if pd.notna(row['Override_Amount']) else row['Live_Amount'],
        axis=1
    )
    
    summary = {
        'total_records': len(included_df),
        'total_amount': included_df['Effective_Amount'].sum(),
        'by_category': included_df.groupby('Category')['Effective_Amount'].sum().to_dict(),
        'by_rep': included_df.groupby('Sales_Rep')['Effective_Amount'].sum().to_dict(),
        'by_confidence': included_df.groupby('Confidence_Level')['Effective_Amount'].sum().to_dict()
    }
    
    return summary

def export_plan_to_excel(version_name, planning_df, summary):
    """Export shipping plan to Excel"""
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Summary sheet
        summary_data = {
            'Metric': ['Total Records', 'Total Amount', 'Plan Version'],
            'Value': [summary['total_records'], f"${summary['total_amount']:,.0f}", version_name]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # By Category
        cat_df = pd.DataFrame(list(summary['by_category'].items()), columns=['Category', 'Amount'])
        cat_df['Amount'] = cat_df['Amount'].apply(lambda x: f"${x:,.0f}")
        cat_df.to_excel(writer, sheet_name='By Category', index=False)
        
        # By Rep
        rep_df = pd.DataFrame(list(summary['by_rep'].items()), columns=['Sales Rep', 'Amount'])
        rep_df['Amount'] = rep_df['Amount'].apply(lambda x: f"${x:,.0f}")
        rep_df.to_excel(writer, sheet_name='By Rep', index=False)
        
        # Detailed records
        included_df = planning_df[planning_df['Include_Flag'] == True].copy()
        included_df['Effective_Amount'] = included_df.apply(
            lambda row: row['Override_Amount'] if pd.notna(row['Override_Amount']) else row['Live_Amount'],
            axis=1
        )
        
        detail_columns = [
            'Record_Type', 'Record_ID', 'Sales_Rep', 'Customer',
            'Live_Amount', 'Override_Amount', 'Effective_Amount',
            'Live_Expected_Date', 'Override_Ship_Date',
            'Category', 'Confidence_Level', 'Planning_Notes', 'Assigned_To'
        ]
        
        detail_df = included_df[detail_columns].copy()
        detail_df.to_excel(writer, sheet_name='Detailed Plan', index=False)
    
    return output.getvalue()

# ============================================================================
# MAIN APP
# ============================================================================

def main():
    # Header
    st.markdown("""
    <div class="planning-header">
        <h1 style="margin:0; color: white;">📦 Q4 2025 Shipping Planning</h1>
        <p style="margin:5px 0 0 0; color: white; opacity: 0.9;">Operations Planning Tool - SANDBOX VERSION</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state
    if 'planning_data' not in st.session_state:
        st.session_state.planning_data = None
    
    if 'user_name' not in st.session_state:
        st.session_state.user_name = 'Unknown User'
    
    # Sidebar controls
    with st.sidebar:
        st.markdown("### 🎛️ Planning Controls")
        
        # User identification
        user_name = st.text_input("Your Name:", value=st.session_state.user_name)
        st.session_state.user_name = user_name
        
        st.markdown("---")
        
        # Load or create new plan
        st.markdown("#### 📂 Plan Management")
        
        # Load saved plans
        saved_versions = load_saved_plan_versions()
        
        plan_action = st.radio(
            "Select action:",
            ["Create New Plan", "Load Saved Plan"],
            key="plan_action"
        )
        
        selected_version = None
        if plan_action == "Load Saved Plan":
            if saved_versions:
                selected_version = st.selectbox("Select plan:", saved_versions)
            else:
                st.info("No saved plans found")
        
        st.markdown("---")
        
        # Refresh data
        if st.button("🔄 Refresh Live Data", use_container_width=True):
            st.cache_data.clear()
            st.session_state.planning_data = None
            st.rerun()
        
        st.markdown("---")
        
        # Save plan
        st.markdown("#### 💾 Save Current Plan")
        version_name = st.text_input(
            "Plan Name:",
            value=f"Plan_{datetime.now():%m%d_%H%M}",
            help="Give this plan a descriptive name"
        )
        
        if st.button("💾 Save Plan", type="primary", use_container_width=True):
            if st.session_state.planning_data is not None:
                with st.spinner("Saving plan..."):
                    if save_shipping_plan_to_sheet(version_name, st.session_state.planning_data):
                        st.success(f"✅ Saved: {version_name}")
                    else:
                        st.error("❌ Failed to save plan")
            else:
                st.warning("⚠️ No data to save")
    
    # Main content
    
    # Load data if not already loaded
    if st.session_state.planning_data is None:
        st.session_state.planning_data = load_all_planning_data()
    
    planning_df = st.session_state.planning_data
    
    if planning_df.empty:
        st.error("❌ No data loaded. Check your Google Sheets connection.")
        return
    
    # Display summary metrics
    st.markdown("### 📊 Current Plan Summary")
    
    # Recalculate summary from session state
    summary = calculate_plan_summary(st.session_state.planning_data)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Included Records", summary['total_records'])
    with col2:
        st.metric("Total Planned Amount", f"${summary['total_amount']:,.0f}")
    with col3:
        # Q4 Quota (hardcoded for now)
        quota = 5_021_440
        gap = quota - summary['total_amount']
        st.metric("Gap to Q4 Quota", f"${gap:,.0f}")
    with col4:
        attainment = (summary['total_amount'] / quota * 100) if quota > 0 else 0
        st.metric("Attainment", f"{attainment:.1f}%")
    
    # Tabs for different sections
    tab1, tab2, tab3, tab4 = st.tabs([
        "📦 Planning Workspace",
        "📈 Summary Views",
        "💾 Export",
        "ℹ️ Help"
    ])
    
    with tab1:
        st.markdown("### 🎯 Build Your Shipping Plan")
        st.caption("Select items to include, add dates, and assign ownership")
        
        # Calculate totals by category for the checkboxes
        category_totals = {}
        if not planning_df.empty:
            for category in planning_df['Category'].unique():
                cat_df = planning_df[planning_df['Category'] == category]
                category_totals[category] = cat_df['Live_Amount'].sum()
        
        # Create checkbox interface like "Build Your Own Forecast"
        st.markdown("#### 📦 Select Components to Include")
        col1, col2, col3 = st.columns(3)
        
        category_selections = {}
        
        with col1:
            # Always included
            st.markdown("**✅ Always Included:**")
            baseline_amount = category_totals.get('Invoiced_Shipped', 0)
            baseline_count = len(planning_df[planning_df['Category'] == 'Invoiced_Shipped'])
            baseline_included = (planning_df[(planning_df['Category'] == 'Invoiced_Shipped') & (planning_df['Include_Flag'] == True)]['Live_Amount'].sum())
            st.info(f"Invoiced & Shipped: ${baseline_amount:,.0f} ({baseline_count} items, ${baseline_included:,.0f} included)")
            
            st.markdown("**🔧 Sales Orders:**")
            category_selections['PF_With_Date'] = st.checkbox(
                f"Pending Fulfillment (with date): ${category_totals.get('PF_With_Date', 0):,.0f}",
                value=False,
                key="select_pf_date"
            )
            category_selections['PF_No_Date'] = st.checkbox(
                f"Pending Fulfillment (NO date): ${category_totals.get('PF_No_Date', 0):,.0f}",
                value=False,
                key="select_pf_nodate",
                help="CX team needs to add dates"
            )
            category_selections['PA_With_Date'] = st.checkbox(
                f"Pending Approval (with date): ${category_totals.get('PA_With_Date', 0):,.0f}",
                value=False,
                key="select_pa_date"
            )
        
        with col2:
            st.markdown("**📋 Pending Approval:**")
            category_selections['PA_No_Date'] = st.checkbox(
                f"Pending Approval (NO date): ${category_totals.get('PA_No_Date', 0):,.0f}",
                value=False,
                key="select_pa_nodate"
            )
            category_selections['PA_Old_Date'] = st.checkbox(
                f"Pending Approval (>2 weeks): ${category_totals.get('PA_Old_Date', 0):,.0f}",
                value=False,
                key="select_pa_old"
            )
            
            st.markdown("**🎯 HubSpot Deals:**")
            category_selections['HubSpot_Expect'] = st.checkbox(
                f"HubSpot Expect/Commit: ${category_totals.get('HubSpot_Expect', 0):,.0f}",
                value=False,
                key="select_hs_expect"
            )
        
        with col3:
            st.markdown("**🎯 HubSpot (cont):**")
            category_selections['HubSpot_Best_Case'] = st.checkbox(
                f"HubSpot Best Case: ${category_totals.get('HubSpot_Best_Case', 0):,.0f}",
                value=False,
                key="select_hs_bc"
            )
            category_selections['HubSpot_Opportunity'] = st.checkbox(
                f"HubSpot Opportunity: ${category_totals.get('HubSpot_Opportunity', 0):,.0f}",
                value=False,
                key="select_hs_opp"
            )
            
            st.markdown("**📅 Q1 Spillover:**")
            category_selections['Q1_Spillover_Expect'] = st.checkbox(
                f"Q1 Spillover - Expect: ${category_totals.get('Q1_Spillover_Expect', 0):,.0f}",
                value=False,
                key="select_q1_expect"
            )
            category_selections['Q1_Spillover_Best_Case'] = st.checkbox(
                f"Q1 Spillover - Best Case: ${category_totals.get('Q1_Spillover_Best_Case', 0):,.0f}",
                value=False,
                key="select_q1_bc"
            )
        
        # Apply selections button
        st.markdown("---")
        if st.button("✅ Apply Selections", type="primary", use_container_width=True):
            # Update Include_Flag based on category selections
            for category, is_selected in category_selections.items():
                if is_selected:
                    planning_df.loc[planning_df['Category'] == category, 'Include_Flag'] = True
                else:
                    planning_df.loc[planning_df['Category'] == category, 'Include_Flag'] = False
            
            # Always include invoiced/shipped
            planning_df.loc[planning_df['Category'] == 'Invoiced_Shipped', 'Include_Flag'] = True
            
            st.session_state.planning_data = planning_df
            st.success("✅ Selections applied! Scroll down to see detailed breakdown.")
            st.rerun()
        
        st.markdown("---")
        
        # Show current summary after selections
        current_included = planning_df[planning_df['Include_Flag'] == True]
        if not current_included.empty:
            st.markdown("#### 📊 Currently Included:")
            summary_col1, summary_col2, summary_col3 = st.columns(3)
            with summary_col1:
                st.metric("Items Included", len(current_included))
            with summary_col2:
                st.metric("Total Amount", f"${current_included['Live_Amount'].sum():,.0f}")
            with summary_col3:
                categories_included = current_included['Category'].nunique()
                st.metric("Categories", categories_included)
        
        st.markdown("---")
        
        # Detailed editing sections (only show if items are included)
        st.markdown("### ✏️ Edit Individual Items")
        st.caption("Fine-tune your selections, add dates, notes, and assignments")
        
        # Baseline - Always Included
        with st.expander("✅ BASELINE (Always Included)", expanded=False):
            baseline_df = planning_df[planning_df['Category'] == 'Invoiced_Shipped']
            if not baseline_df.empty:
                st.metric("Invoiced & Shipped", f"${baseline_df['Live_Amount'].sum():,.0f}")
                st.caption(f"{len(baseline_df)} invoices")
            else:
                st.info("No invoiced/shipped records")
        
        # Planning components
        st.markdown("---")
        st.markdown("#### 🔧 Planning Components")
        st.caption("Edit these sections to build your plan")
        
        # Pending Fulfillment with Date
        with st.expander("📦 Pending Fulfillment (With Date)", expanded=True):
            pf_date_df = planning_df[planning_df['Category'] == 'PF_With_Date']
            updated_pf_date = display_planning_section("PF With Date", pf_date_df, "pf_date")
            planning_df.update(updated_pf_date)
        
        # Pending Fulfillment NO Date
        with st.expander("📦 Pending Fulfillment (NO Date) - CX Input Required", expanded=True):
            pf_nodate_df = planning_df[planning_df['Category'] == 'PF_No_Date']
            updated_pf_nodate = display_planning_section("PF No Date", pf_nodate_df, "pf_nodate")
            planning_df.update(updated_pf_nodate)
        
        # Pending Approval with Date
        with st.expander("📋 Pending Approval (With Date)", expanded=False):
            pa_date_df = planning_df[planning_df['Category'] == 'PA_With_Date']
            updated_pa_date = display_planning_section("PA With Date", pa_date_df, "pa_date")
            planning_df.update(updated_pa_date)
        
        # Pending Approval NO Date
        with st.expander("📋 Pending Approval (NO Date or >2 weeks old)", expanded=True):
            pa_nodate_df = planning_df[planning_df['Category'].isin(['PA_No_Date', 'PA_Old_Date'])]
            updated_pa_nodate = display_planning_section("PA No Date/Old", pa_nodate_df, "pa_nodate")
            planning_df.update(updated_pa_nodate)
        
        # HubSpot Expect/Commit
        with st.expander("🎯 HubSpot Expect/Commit", expanded=False):
            hs_expect_df = planning_df[planning_df['Category'] == 'HubSpot_Expect']
            updated_hs_expect = display_planning_section("HubSpot Expect", hs_expect_df, "hs_expect")
            planning_df.update(updated_hs_expect)
        
        # HubSpot Best Case
        with st.expander("🎯 HubSpot Best Case", expanded=False):
            hs_bc_df = planning_df[planning_df['Category'] == 'HubSpot_Best_Case']
            updated_hs_bc = display_planning_section("HubSpot Best Case", hs_bc_df, "hs_bc")
            planning_df.update(updated_hs_bc)
        
        # Q1 Spillover
        with st.expander("📅 Q1 Spillover - Expect/Commit", expanded=False):
            q1_expect_df = planning_df[planning_df['Category'] == 'Q1_Spillover_Expect']
            updated_q1_expect = display_planning_section("Q1 Spillover Expect", q1_expect_df, "q1_expect")
            planning_df.update(updated_q1_expect)
        
        with st.expander("📅 Q1 Spillover - Best Case", expanded=False):
            q1_bc_df = planning_df[planning_df['Category'] == 'Q1_Spillover_Best_Case']
            updated_q1_bc = display_planning_section("Q1 Spillover BC", q1_bc_df, "q1_bc")
            planning_df.update(updated_q1_bc)
        
        # Update session state
        st.session_state.planning_data = planning_df
    
    with tab2:
        st.markdown("### 📊 Summary Views")
        
        # Recalculate summary with any changes
        summary = calculate_plan_summary(st.session_state.planning_data)
        
        # By Category
        st.markdown("#### By Category")
        if summary['by_category']:
            cat_data = pd.DataFrame(list(summary['by_category'].items()), columns=['Category', 'Amount'])
            cat_data = cat_data.sort_values('Amount', ascending=False)
            
            fig = go.Figure(data=[
                go.Bar(x=cat_data['Category'], y=cat_data['Amount'], marker_color='#667eea')
            ])
            fig.update_layout(
                title="Planned Amount by Category",
                xaxis_title="Category",
                yaxis_title="Amount ($)",
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)
        
        # By Rep
        st.markdown("#### By Sales Rep")
        if summary['by_rep']:
            rep_data = pd.DataFrame(list(summary['by_rep'].items()), columns=['Sales Rep', 'Amount'])
            rep_data = rep_data.sort_values('Amount', ascending=False)
            
            fig = go.Figure(data=[
                go.Bar(x=rep_data['Sales Rep'], y=rep_data['Amount'], marker_color='#764ba2')
            ])
            fig.update_layout(
                title="Planned Amount by Sales Rep",
                xaxis_title="Sales Rep",
                yaxis_title="Amount ($)",
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)
        
        # By Confidence
        st.markdown("#### By Confidence Level")
        if summary['by_confidence']:
            conf_data = pd.DataFrame(list(summary['by_confidence'].items()), columns=['Confidence', 'Amount'])
            
            fig = go.Figure(data=[
                go.Pie(labels=conf_data['Confidence'], values=conf_data['Amount'], hole=0.4)
            ])
            fig.update_layout(title="Planned Amount by Confidence Level", height=400)
            st.plotly_chart(fig, use_container_width=True)
    
    with tab3:
        st.markdown("### 💾 Export Options")
        
        summary = calculate_plan_summary(st.session_state.planning_data)
        
        st.info(f"📊 Ready to export: {summary['total_records']} records, ${summary['total_amount']:,.0f}")
        
        export_name = st.text_input("Export Name:", value=f"Q4_Shipping_Plan_{datetime.now():%Y%m%d}")
        
        if st.button("📥 Export to Excel", type="primary"):
            excel_data = export_plan_to_excel(export_name, st.session_state.planning_data, summary)
            
            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_data,
                file_name=f"{export_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    with tab4:
        st.markdown("### ℹ️ How to Use This Tool")
        
        st.markdown("""
        **Purpose:** Build and track Q4 2025 shipping plans by selecting which orders and deals to include.
        
        **Workflow:**
        1. **Load Data**: Data automatically loads from NetSuite and HubSpot
        2. **Review Baseline**: See what's already invoiced/shipped (always included)
        3. **Select Items**: 
           - Check "Include?" for items you want in the plan
           - Add override dates for items without dates
           - Add planning notes and assign to team members
        4. **Save Plan**: Give your plan a name and save it
        5. **Export**: Download Excel report to share with team
        
        **Key Features:**
        - **Live vs Override**: System shows live data but lets you override for planning
        - **Categories**: Items auto-categorize (PF, PA, HubSpot deals, etc.)
        - **Team Collaboration**: Assign items to Xander, Kyle, Cory, Greg, or Tif
        - **Confidence Levels**: Tag items as High/Medium/Low confidence
        
        **This is a SANDBOX version** - completely separate from the main sales dashboard.
        """)

if __name__ == "__main__":
    main()
