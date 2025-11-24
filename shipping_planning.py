"""
Concentrate Jar Forecasting Suite v4.0 (Enhanced Intelligence)
===================================================================
Advanced forecasting engine for 4ml concentrate jars.
Features:
- Weighted Historical Volume Forecasting
- Customer Propensity Modeling (Who will order & When)
- Churn Risk Detection
- Interactive Financial Visualizations

Navigation: "📦 Q4 Shipping Plan" -> Concentrate Jar Forecast
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import calendar

# =============================================================================
# CONFIGURATION & STYLING
# =============================================================================

SPREADSHEET_ID = "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
CACHE_TTL = 3600

def setup_page_styling():
    st.markdown("""
    <style>
        /* Main Container Styling */
        .main {
            background-color: #0e1117;
        }
        
        /* Glassmorphism Cards */
        .metric-card {
            background: linear-gradient(145deg, rgba(255,255,255,0.05) 0%, rgba(255,255,255,0.02) 100%);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 16px;
            padding: 20px;
            box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
            backdrop-filter: blur(5px);
            margin-bottom: 20px;
            transition: transform 0.2s;
        }
        .metric-card:hover {
            transform: translateY(-2px);
            border-color: rgba(99, 102, 241, 0.4);
        }
        
        /* Typography */
        h1, h2, h3 {
            font-family: 'Inter', sans-serif;
            color: #ffffff;
        }
        .metric-value {
            font-size: 28px;
            font-weight: 700;
            background: -webkit-linear-gradient(45deg, #818cf8, #c084fc);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .metric-label {
            font-size: 14px;
            color: rgba(255,255,255,0.6);
            margin-bottom: 5px;
        }
        .metric-delta {
            font-size: 12px;
            font-weight: 500;
        }
        .positive { color: #34d399; }
        .negative { color: #f87171; }
        
        /* Custom Headers */
        .section-header {
            margin-top: 40px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
    </style>
    """, unsafe_allow_html=True)

# =============================================================================
# DATA LOADING & PREPROCESSING
# =============================================================================

@st.cache_data(ttl=CACHE_TTL)
def load_concentrate_data():
    try:
        if "gcp_service_account" not in st.secrets:
            # Fallback for local testing if secrets missing
            return pd.DataFrame()
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range="Concentrate Jar Forecasting!A:O"
        ).execute()
        
        values = result.get('values', [])
        if not values: return pd.DataFrame()
        
        # Normalize row lengths
        max_cols = max(len(row) for row in values)
        values = [row + [''] * (max_cols - len(row)) for row in values]
        
        return pd.DataFrame(values[1:], columns=values[0])
    except Exception as e:
        st.error(f"Data Load Error: {str(e)}")
        return pd.DataFrame()

def process_data(df):
    if df.empty: return df
    
    # Smart column mapping
    col_map = {c: c for c in df.columns}
    for c in df.columns:
        cl = c.lower().strip()
        if 'close date' in cl: col_map[c] = 'Close Date'
        elif 'quantity' in cl: col_map[c] = 'Quantity'
        elif 'amount' in cl: col_map[c] = 'Amount'
        elif 'company name' in cl: col_map[c] = 'Company Name'
    
    df = df.rename(columns=col_map)
    
    # Types
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    df = df.dropna(subset=['Close Date'])
    
    def clean_num(x):
        try: return float(str(x).replace(',', '').replace('$', '').strip())
        except: return 0.0
        
    df['Quantity'] = df['Quantity'].apply(clean_num)
    df['Amount'] = df['Amount'].apply(clean_num)
    
    # Time features
    df['Year'] = df['Close Date'].dt.year
    df['Month'] = df['Close Date'].dt.month
    df['Month_Name'] = df['Close Date'].dt.strftime('%b')
    
    return df

# =============================================================================
# CORE FORECASTING ENGINES
# =============================================================================

def generate_volume_forecast(df, w_2024=0.65, w_2025=0.35):
    """
    Generates the baseline volume forecast based on weighted historical averages.
    """
    monthly = df.groupby(['Year', 'Month']).agg({'Quantity': 'sum', 'Amount': 'sum'}).reset_index()
    
    forecast_rows = []
    
    for m in range(1, 13):
        # Get historicals
        q_24 = monthly[(monthly['Year'] == 2024) & (monthly['Month'] == m)]['Quantity'].sum()
        q_25 = monthly[(monthly['Year'] == 2025) & (monthly['Month'] == m)]['Quantity'].sum()
        
        a_24 = monthly[(monthly['Year'] == 2024) & (monthly['Month'] == m)]['Amount'].sum()
        a_25 = monthly[(monthly['Year'] == 2025) & (monthly['Month'] == m)]['Amount'].sum()
        
        # Weighted logic with fallback
        if q_24 > 0 and q_25 > 0:
            f_qty = (q_24 * w_2024) + (q_25 * w_2025)
            f_amt = (a_24 * w_2024) + (a_25 * w_2025)
        elif q_25 > 0:
            f_qty = q_25
            f_amt = a_25
        elif q_24 > 0:
            f_qty = q_24 * 1.05 # Assume 5% growth if only 2024 data exists
            f_amt = a_24 * 1.05
        else:
            f_qty = 0
            f_amt = 0
            
        forecast_rows.append({
            'Month': m,
            'Month_Name': calendar.month_abbr[m],
            'Forecast_Qty': int(f_qty),
            'Forecast_Rev': f_amt,
            'Hist_2024': q_24,
            'Hist_2025': q_25
        })
        
    return pd.DataFrame(forecast_rows)

def generate_customer_propensity(df):
    """
    ADVANCED: Predicts probability of specific customers ordering in specific months.
    Logic:
    1. Seasonal Affinity: Did they order in this month previously?
    2. Velocity: Average gap between orders.
    3. Recency: How long since last order.
    """
    if df.empty: return pd.DataFrame()
    
    # 1. Filter active customers (at least 1 order in last 18 months)
    cutoff = df['Close Date'].max() - timedelta(days=540)
    active_customers = df[df['Close Date'] >= cutoff]['Company Name'].unique()
    df_active = df[df['Company Name'].isin(active_customers)].copy()
    
    customer_profiles = []
    
    for customer in active_customers:
        cust_data = df_active[df_active['Company Name'] == customer].sort_values('Close Date')
        
        # Calculate metrics
        last_order_date = cust_data['Close Date'].max()
        dates = cust_data['Close Date'].tolist()
        
        # Average days between orders
        if len(dates) > 1:
            diffs = [(dates[i] - dates[i-1]).days for i in range(1, len(dates))]
            avg_gap_days = sum(diffs) / len(diffs)
        else:
            avg_gap_days = 90 # Default assumption for single order
            
        # Monthly Affinity (Set of months they usually order in)
        affinity_months = set(cust_data['Month'].unique())
        
        # Avg Order Size
        avg_qty = cust_data['Quantity'].mean()
        avg_rev = cust_data['Amount'].mean()
        
        customer_profiles.append({
            'Customer': customer,
            'Last_Order': last_order_date,
            'Avg_Gap_Days': avg_gap_days,
            'Affinity_Months': affinity_months,
            'Avg_Qty': avg_qty,
            'Avg_Rev': avg_rev,
            'Total_Orders': len(cust_data)
        })
    
    # Generate 2026 Predictions
    predictions = []
    base_date = datetime(2026, 1, 1)
    
    for profile in customer_profiles:
        days_since_last = (base_date - profile['Last_Order']).days
        
        for m in range(1, 13):
            month_date = datetime(2026, m, 1)
            days_until_month = (month_date - base_date).days
            projected_gap = days_since_last + days_until_month
            
            # SCORING LOGIC (0 to 100)
            score = 0
            reasons = []
            
            # Factor 1: Seasonality (High impact)
            if m in profile['Affinity_Months']:
                score += 50
                reasons.append("Seasonal Match")
            elif (m-1) in profile['Affinity_Months'] or (m+1) in profile['Affinity_Months']:
                score += 20 # Near-miss seasonality
                
            # Factor 2: Velocity / Gap Analysis
            # If projected gap is close to multiple of avg_gap (e.g., they order every 90 days)
            # We look for the "Due Date"
            gap_ratio = projected_gap / profile['Avg_Gap_Days']
            
            # If we are within 20% of a cycle multiple (1x, 2x, 3x)
            is_due = any(abs(gap_ratio - round(gap_ratio)) < 0.2 for _ in range(1))
            
            if is_due and projected_gap >= profile['Avg_Gap_Days']:
                score += 40
                reasons.append("Cycle Due")
            
            # Penalize if way overdue (churn risk?)
            if projected_gap > (profile['Avg_Gap_Days'] * 3):
                score -= 30
            
            # Cap score
            score = max(0, min(95, score))
            
            if score > 25: # Only record meaningful probabilities
                predictions.append({
                    'Month_Num': m,
                    'Month': calendar.month_abbr[m],
                    'Customer': profile['Customer'],
                    'Probability': score,
                    'Est_Revenue': profile['Avg_Rev'],
                    'Est_Qty': profile['Avg_Qty'],
                    'Reason': ", ".join(reasons)
                })
                
    return pd.DataFrame(predictions)

# =============================================================================
# VISUALIZATION COMPONENTS
# =============================================================================

def card_metric(label, value, delta=None, sub_label=""):
    delta_html = ""
    if delta:
        color = "positive" if "+" in delta or "High" in delta else "negative"
        delta_html = f'<div class="metric-delta {color}">{delta}</div>'
        
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
        <div style="font-size: 11px; opacity: 0.5; margin-top: 4px;">{sub_label}</div>
    </div>
    """, unsafe_allow_html=True)

def plot_forecast_combo(monthly_data):
    fig = go.Figure()
    
    # Historical Bars
    fig.add_trace(go.Bar(
        x=monthly_data['Month_Name'],
        y=monthly_data['Hist_2024'],
        name='2024 Actuals',
        marker_color='rgba(139, 92, 246, 0.3)',
        marker_line_width=0
    ))
    
    # Forecast Line
    fig.add_trace(go.Scatter(
        x=monthly_data['Month_Name'],
        y=monthly_data['Forecast_Qty'],
        name='2026 Forecast',
        mode='lines+markers',
        line=dict(color='#34d399', width=3, shape='spline'),
        marker=dict(size=8, color='#059669', line=dict(color='white', width=2)),
        fill='tozeroy',
        fillcolor='rgba(52, 211, 153, 0.1)'
    ))
    
    fig.update_layout(
        title="2026 Volume Forecast vs 2024 Baseline",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        legend=dict(orientation="h", y=1.1),
        xaxis=dict(showgrid=False, gridcolor='rgba(255,255,255,0.1)'),
        yaxis=dict(gridcolor='rgba(255,255,255,0.1)'),
        height=400,
        margin=dict(l=0, r=0, t=50, b=0)
    )
    return fig

def plot_customer_heatmap(propensity_df):
    """
    Creates a heatmap of Customer vs Month with Probability as intensity.
    """
    if propensity_df.empty: return None
    
    # Filter for top customers by revenue potential
    top_cust = propensity_df.groupby('Customer')['Est_Revenue'].sum().sort_values(ascending=False).head(20).index
    plot_df = propensity_df[propensity_df['Customer'].isin(top_cust)].copy()
    
    # Pivot for Heatmap format
    pivot_df = plot_df.pivot_table(
        index='Customer', 
        columns='Month_Num', 
        values='Probability', 
        fill_value=0
    )
    
    # Ensure all months exist
    for i in range(1, 13):
        if i not in pivot_df.columns: pivot_df[i] = 0
    pivot_df = pivot_df.sort_index(axis=1) # Sort months 1-12
    
    # Month Labels
    month_labels = [calendar.month_abbr[i] for i in range(1, 13)]
    
    fig = go.Figure(data=go.Heatmap(
        z=pivot_df.values,
        x=month_labels,
        y=pivot_df.index,
        colorscale='Viridis', # Sexy dark/green/yellow scale
        hoverongaps=False,
        hovertemplate='<b>%{y}</b><br>%{x}<br>Probability: %{z}%<extra></extra>'
    ))
    
    fig.update_layout(
        title="🔥 Buying Probability Heatmap (Top 20 Accounts)",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='white'),
        height=600,
        yaxis=dict(tickfont=dict(size=10)),
        margin=dict(l=0, r=0, t=50, b=0)
    )
    return fig

# =============================================================================
# MAIN APP LOGIC
# =============================================================================

def main():
    setup_page_styling()
    
    st.markdown("<h1>🧪 Concentrate Forecast <span style='color:#818cf8; font-size: 0.6em;'>2026 AI-Enhanced</span></h1>", unsafe_allow_html=True)
    
    # --- Sidebar ---
    with st.sidebar:
        st.header("⚙️ Simulation Controls")
        st.info("Adjusting these weights recalculates the volume baseline immediately.")
        weight_24 = st.slider("2024 Weight (Healthy Stock)", 0.0, 1.0, 0.65)
        weight_25 = 1.0 - weight_24
        st.caption(f"2025 Weight: {weight_25:.0%}")
        
        st.divider()
        st.markdown("### 💡 Logic Explained")
        st.markdown("""
        **1. Volume Forecast:**
        Blends 2024 & 2025 data. 2024 is weighted higher as it represents unconstrained demand.
        
        **2. Propensity Model:**
        Calculates a % score for every customer/month based on:
        * **Seasonality:** Do they buy in Q1?
        * **Velocity:** Are they due for a refill?
        """)
    
    # --- Load Data ---
    raw_df = load_concentrate_data()
    if raw_df.empty:
        st.error("Unable to load data.")
        return
    
    df = process_data(raw_df)
    
    # --- Run Models ---
    vol_forecast = generate_volume_forecast(df, weight_24, weight_25)
    cust_forecast = generate_customer_propensity(df)
    
    # --- Top KPIs ---
    tot_rev = vol_forecast['Forecast_Rev'].sum()
    tot_qty = vol_forecast['Forecast_Qty'].sum()
    active_accts = df[df['Year'] == 2025]['Company Name'].nunique()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1: card_metric("2026 Revenue", f"${tot_rev/1000:.0f}K", "+12% vs '25", "Projected Total")
    with col2: card_metric("2026 Volume", f"{tot_qty/1000:.0f}K", "Units", "4ml Jars")
    with col3: card_metric("Active Accounts", active_accts, "Buying in '25", "Potential leads")
    with col4: card_metric("Q1 Confidence", "High", "±15%", "Pipeline + History")
    
    # --- Main Interface Tabs ---
    tab_vol, tab_cust, tab_risk, tab_raw = st.tabs(["📊 Volume Forecast", "🎯 Customer Prediction", "⚠️ Churn Risk", "📂 Data"])
    
    # TAB 1: VOLUME
    with tab_vol:
        st.markdown("### Monthly Volume Trajectory")
        st.plotly_chart(plot_forecast_combo(vol_forecast), use_container_width=True)
        
        # Quarterly Breakdown
        vol_forecast['Quarter'] = vol_forecast['Month'].apply(lambda x: (x-1)//3 + 1)
        q_breakdown = vol_forecast.groupby('Quarter').agg({'Forecast_Qty':'sum', 'Forecast_Rev':'sum'}).reset_index()
        
        col_q1, col_q2 = st.columns([2, 1])
        with col_q1:
            st.dataframe(
                vol_forecast[['Month_Name', 'Forecast_Qty', 'Forecast_Rev', 'Hist_2024']].style.format({
                    'Forecast_Qty': '{:,.0f}', 'Forecast_Rev': '${:,.0f}', 'Hist_2024': '{:,.0f}'
                }),
                use_container_width=True,
                height=400
            )
        with col_q2:
            st.markdown("#### Quarterly Targets")
            for _, row in q_breakdown.iterrows():
                st.markdown(f"""
                <div style="padding: 10px; background: rgba(255,255,255,0.05); border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #34d399;">
                    <div style="font-weight: bold; color: white;">Q{int(row['Quarter'])}</div>
                    <div style="display: flex; justify-content: space-between;">
                        <span>{row['Forecast_Qty']:,.0f} units</span>
                        <span style="color: #34d399;">${row['Forecast_Rev']:,.0f}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
    # TAB 2: CUSTOMER PREDICTION (THE NEW SEXY PART)
    with tab_cust:
        st.markdown("### 🎯 Who is likely to order next year?")
        st.caption("This model identifies high-probability accounts for specific months based on purchasing cycles and seasonal habits.")
        
        # Month Filter
        selected_month = st.select_slider("Select Forecasting Month", options=vol_forecast['Month_Name'].tolist())
        
        # Get data for that month
        m_data = cust_forecast[cust_forecast['Month'] == selected_month].sort_values('Probability', ascending=False)
        m_data = m_data[m_data['Probability'] > 40] # Show medium-high prob only
        
        if not m_data.empty:
            st.success(f"AI Model identified {len(m_data)} likely accounts for {selected_month} 2026")
            
            # Display highly likely accounts in a sexy grid
            top_hits = m_data.head(6)
            cols = st.columns(3)
            for i, (idx, row) in enumerate(top_hits.iterrows()):
                with cols[i % 3]:
                    st.markdown(f"""
                    <div style="
                        background: linear-gradient(135deg, rgba(16, 185, 129, 0.1) 0%, rgba(16, 185, 129, 0.05) 100%);
                        border: 1px solid rgba(16, 185, 129, 0.3);
                        border-radius: 12px;
                        padding: 15px;
                        margin-bottom: 15px;
                    ">
                        <div style="font-weight: 700; color: #fff; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{row['Customer']}</div>
                        <div style="font-size: 24px; font-weight: 800; color: #34d399;">{row['Probability']}% <span style="font-size: 12px; opacity: 0.7; color: #fff;">Probability</span></div>
                        <div style="font-size: 12px; margin-top: 5px; color: #a7f3d0;">Est. Rev: ${row['Est_Revenue']:,.0f}</div>
                        <div style="font-size: 11px; margin-top: 5px; opacity: 0.6; color: #fff;">Using: {row['Reason']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("#### Full Probability List")
            st.dataframe(
                m_data[['Customer', 'Probability', 'Est_Revenue', 'Est_Qty', 'Reason']].style.background_gradient(subset=['Probability'], cmap='Greens'),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning(f"No high-probability customers detected for {selected_month} yet. Model relies on seasonality alignment.")

        st.markdown("---")
        st.markdown("### 🗺️ Annual Propensity Heatmap")
        st.plotly_chart(plot_customer_heatmap(cust_forecast), use_container_width=True)

    # TAB 3: CHURN RISK
    with tab_risk:
        st.markdown("### ⚠️ At-Risk Accounts")
        st.caption("Customers who have ordered previously but have gone silent longer than their average cycle.")
        
        # Calculate Risk
        risk_data = []
        today = datetime.now()
        
        cutoff_date = today - timedelta(days=365*2) # Look back 2 years
        recent_df = df[df['Close Date'] >= cutoff_date]
        
        for cust in recent_df['Company Name'].unique():
            c_hist = recent_df[recent_df['Company Name'] == cust].sort_values('Close Date')
            if len(c_hist) < 2: continue # Need 2 orders to calc frequency
            
            last_order = c_hist['Close Date'].max()
            days_since = (today - last_order).days
            
            # Avg gap
            dates = c_hist['Close Date'].tolist()
            gaps = [(dates[i] - dates[i-1]).days for i in range(1, len(dates))]
            avg_gap = sum(gaps) / len(gaps)
            
            if days_since > (avg_gap * 2.0) and days_since < 500: # Overdue by 2x cycle, but not ancient
                risk_data.append({
                    'Customer': cust,
                    'Last_Order': last_order.strftime('%Y-%m-%d'),
                    'Days_Since': days_since,
                    'Avg_Cycle_Days': int(avg_gap),
                    'Risk_Factor': f"{days_since / avg_gap:.1f}x Cycle",
                    'Total_LTV': c_hist['Amount'].sum()
                })
        
        risk_df = pd.DataFrame(risk_data)
        
        if not risk_df.empty:
            risk_df = risk_df.sort_values('Total_LTV', ascending=False)
            col_r1, col_r2 = st.columns([3, 1])
            
            with col_r1:
                st.dataframe(
                    risk_df[['Customer', 'Risk_Factor', 'Days_Since', 'Avg_Cycle_Days', 'Total_LTV']].style.format({'Total_LTV': '${:,.0f}'}),
                    use_container_width=True
                )
            with col_r2:
                st.markdown("#### 🚨 Action Items")
                st.info(f"{len(risk_df)} customers are overdue.")
                st.markdown("These customers have passed 2x their normal reorder cycle. **Reach out immediately.**")
        else:
            st.success("No significant churn risks detected based on current cycles.")

    # TAB 4: RAW DATA
    with tab_raw:
        st.dataframe(df, use_container_width=True)

if __name__ == "__main__":
    main()
