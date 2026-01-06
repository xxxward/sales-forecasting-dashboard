"""
Q1 2026 Sales Forecasting Module
Based on Sales Dashboard architecture

KEY INSIGHT: For Q1 2026 dashboard, the primary quarter data is in the "date" buckets:
- pf_date_ext + pf_date_int = PF orders with Q1 2026 dates
- pa_date = PA orders with PA Date in Q1 2026

Spillover buckets are now for adjacent quarters:
- pf_q4_spillover/pa_q4_spillover = Q4 2025 carryover orders
- pf_q2_spillover/pa_q2_spillover = Q2 2026 forward spillover

This module imports directly from the main dashboard to reuse all data loading and categorization logic.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
# ========== STREAMLIT APP CONFIG ==========
st.set_page_config(
    page_title="Q1 2026 Forecast",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ========== DATE CONSTANTS ==========
Q1_2026_START = pd.Timestamp('2026-01-01')
Q1_2026_END = pd.Timestamp('2026-03-31')
Q4_2025_START = pd.Timestamp('2025-10-01')
Q4_2025_END = pd.Timestamp('2025-12-31')


def get_mst_time():
    """Get current time in Mountain Standard Time"""
    return datetime.now(ZoneInfo("America/Denver"))


def calculate_business_days_remaining_q1():
    """Calculate business days remaining in Q1 2026 (until Mar 31, 2026)"""
    from datetime import date
    
    today = date.today()
    q1_end = date(2026, 3, 31)
    
    # If we're past Q1, return 0
    if today > q1_end:
        return 0
    
    # Q1 2026 holidays
    holidays = [
        date(2026, 1, 1),   # New Year's Day
        date(2026, 1, 20),  # MLK Day
        date(2026, 2, 16),  # Presidents Day
    ]
    
    business_days = 0
    current_date = today
    
    while current_date <= q1_end:
        if current_date.weekday() < 5 and current_date not in holidays:
            business_days += 1
        current_date += timedelta(days=1)
    
    return business_days


# ========== CUSTOM CSS (FORECAST UI V2) ==========
def inject_custom_css():
    st.markdown(r"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=Space+Grotesk:wght@500;600;700&display=swap');

    :root{
        --bg0:#020617;
        --bg1:#0b1220;
        --bg2:#0f172a;
        --card:rgba(15,23,42,.62);
        --card2:rgba(30,41,59,.42);
        --border:rgba(148,163,184,.14);
        --border2:rgba(255,255,255,.08);
        --text:#e2e8f0;
        --muted:#94a3b8;
        --muted2:#64748b;
        --blue:#3b82f6;
        --emerald:#10b981;
        --amber:#f59e0b;
        --red:#ef4444;
        --violet:#8b5cf6;
        --pink:#ec4899;
        --radius:20px;
        --radius-sm:14px;
        --shadow: 0 18px 50px rgba(0,0,0,.55);
        --shadow-soft: 0 10px 28px rgba(0,0,0,.35);
        --glow-blue: 0 0 32px rgba(59,130,246,.18);
        --glow-emerald: 0 0 32px rgba(16,185,129,.18);
        --glow-amber: 0 0 32px rgba(245,158,11,.18);
    }

    /* ---- Streamlit chrome ---- */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* ---- App background ---- */
    .stApp{
        color: var(--text);
        font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
        background:
            radial-gradient(1200px circle at 18% 0%, rgba(59,130,246,.23), transparent 55%),
            radial-gradient(1000px circle at 82% 10%, rgba(16,185,129,.18), transparent 55%),
            radial-gradient(900px circle at 65% 95%, rgba(139,92,246,.18), transparent 55%),
            linear-gradient(180deg, #020617 0%, #020617 45%, #000 100%);
    }

    /* Subtle animated sheen */
    @keyframes sheen {
        0% { transform: translateX(-40%) translateY(-10%) rotate(8deg); opacity: .0; }
        30% { opacity: .35; }
        100% { transform: translateX(55%) translateY(12%) rotate(8deg); opacity: 0; }
    }
    .stApp::before{
        content:"";
        position: fixed;
        inset: -40%;
        pointer-events: none;
        background: linear-gradient(90deg, transparent 0%, rgba(255,255,255,.06) 45%, transparent 100%);
        filter: blur(10px);
        animation: sheen 10s ease-in-out infinite;
        z-index: 0;
    }

    /* Bring Streamlit content above pseudo elements */
    section.main > div { position: relative; z-index: 1; }

    /* ---- Layout ---- */
    .block-container{
        padding-top: 1.75rem;
        padding-bottom: 9rem;
        max-width: 1500px !important;
    }
    .main .block-container { padding-bottom: 140px !important; }

    /* Sidebar */
    section[data-testid="stSidebar"] > div{
        background:
            radial-gradient(900px circle at 30% 0%, rgba(59,130,246,.18), transparent 55%),
            linear-gradient(180deg, rgba(2,6,23,.92), rgba(15,23,42,.92));
        border-right: 1px solid var(--border);
    }
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3{
        font-family: "Space Grotesk", Inter, sans-serif;
    }

    /* Links */
    a, a:visited { color: rgba(96,165,250,.95); }
    a:hover { color: rgba(147,197,253,.95); }

    /* ---- Glass cards ---- */
    .glass-card{
        background: linear-gradient(145deg, rgba(15,23,42,.68), rgba(30,41,59,.38));
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 22px 22px;
        box-shadow: var(--shadow-soft);
        margin-bottom: 18px;
        position: relative;
        overflow: hidden;
    }
    .glass-card::after{
        content:"";
        position:absolute;
        inset:-1px;
        border-radius: var(--radius);
        padding: 1px;
        background: linear-gradient(135deg,
            rgba(59,130,246,.35),
            rgba(16,185,129,.20),
            rgba(139,92,246,.25)
        );
        -webkit-mask:
            linear-gradient(#000 0 0) content-box,
            linear-gradient(#000 0 0);
        -webkit-mask-composite: xor;
        mask-composite: exclude;
        pointer-events:none;
        opacity: .55;
    }
    .glass-card:hover{
        box-shadow: var(--shadow), var(--glow-blue);
        border-color: rgba(147,197,253,.18);
        transition: box-shadow .22s ease, border-color .22s ease;
    }

    /* Soft divider */
    .soft-divider{
        height: 1px;
        width: 100%;
        margin: 16px 0;
        background: linear-gradient(to right, transparent, rgba(148,163,184,.22), transparent);
    }

    /* ---- Titles & section headers ---- */
    .section-header{
        font-family: "Space Grotesk", Inter, sans-serif;
        font-size: 1.55rem;
        font-weight: 700;
        margin: 1.25rem 0 0.75rem 0;
        letter-spacing: .2px;
        background: linear-gradient(90deg, #fff, #cbd5e1);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        display:flex;
        align-items:center;
        gap: 10px;
    }
    .pill{
        display:inline-flex;
        align-items:center;
        gap: 8px;
        padding: 6px 12px;
        border-radius: 999px;
        border: 1px solid var(--border);
        background: rgba(2,6,23,.28);
        color: var(--muted);
        font-size: .8rem;
    }

    /* ---- Hero metrics ---- */
    .hero-metric{
        background: linear-gradient(145deg, rgba(15,23,42,.72), rgba(30,41,59,.40));
        border-radius: 16px;
        padding: 16px 18px;
        border: 1px solid rgba(148,163,184,.14);
        box-shadow: 0 10px 26px rgba(0,0,0,.30);
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    .hero-metric::before{
        content:"";
        position:absolute;
        inset:-1px;
        background: radial-gradient(500px circle at 20% 10%, rgba(59,130,246,.14), transparent 50%);
        opacity: .9;
        pointer-events:none;
    }
    .hero-label{
        font-size: .78rem;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: var(--muted);
        margin-bottom: 6px;
    }
    .hero-value{
        font-family: "Space Grotesk", Inter, sans-serif;
        font-size: 1.65rem;
        font-weight: 700;
        color: #fff;
        font-variant-numeric: tabular-nums;
    }

    /* ---- Sticky footer HUD ---- */
    .sticky-forecast-bar-q1{
        position: fixed;
        bottom: 18px;
        left: 50%;
        transform: translateX(-50%);
        width: 92%;
        max-width: 1500px;
        z-index: 99999;
        background: rgba(2,6,23,.86);
        backdrop-filter: blur(18px);
        -webkit-backdrop-filter: blur(18px);
        border: 1px solid rgba(148,163,184,.18);
        border-radius: 26px;
        padding: 12px 26px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 16px 55px rgba(0,0,0,.55), 0 0 28px rgba(59,130,246,.12);
    }
    .sticky-item{
        display:flex;
        flex-direction:column;
        align-items:center;
        flex: 1;
        min-width: 0;
    }
    .sticky-label{
        font-size: .70rem;
        color: var(--muted);
        text-transform: uppercase;
        letter-spacing: 1.2px;
        margin-bottom: 2px;
        white-space: nowrap;
    }
    .sticky-val{
        font-family: "Space Grotesk", Inter, sans-serif;
        font-size: 1.35rem;
        font-weight: 700;
        font-variant-numeric: tabular-nums;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 100%;
    }
    .val-sched{ color: #34d399; text-shadow: var(--glow-emerald); }
    .val-pipe{ color: #60a5fa; text-shadow: var(--glow-blue); }
    .val-reorder{ color: #fbbf24; text-shadow: var(--glow-amber); }
    .val-total{
        font-size: 1.55rem;
        background: linear-gradient(135deg, #fff 0%, #cbd5e1 65%, #94a3b8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .val-gap-behind{ color:#fb7185; text-shadow: 0 0 22px rgba(244,63,94,.20); }
    .val-gap-ahead{ color:#34d399; text-shadow: 0 0 22px rgba(16,185,129,.20); }

    .sticky-sep{
        width: 1px;
        height: 44px;
        background: linear-gradient(to bottom, transparent, rgba(148,163,184,.30), transparent);
        margin: 0 10px;
    }

    /* ---- Tier badges ---- */
    .tier-badge{
        display:inline-block;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: .78rem;
        font-weight: 650;
        margin-right: 8px;
        border: 1px solid rgba(148,163,184,.20);
        background: rgba(2,6,23,.25);
    }
    .tier-likely{ background: rgba(16,185,129,.16); color:#34d399; border-color: rgba(16,185,129,.30); }
    .tier-possible{ background: rgba(245,158,11,.16); color:#fbbf24; border-color: rgba(245,158,11,.30); }
    .tier-longshot{ background: rgba(148,163,184,.14); color:#cbd5e1; border-color: rgba(148,163,184,.22); }

    /* ---- Widgets ---- */
    /* Inputs */
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div{
        background: rgba(2,6,23,.28) !important;
        border: 1px solid rgba(148,163,184,.18) !important;
        border-radius: 14px !important;
        color: var(--text) !important;
        box-shadow: none !important;
    }
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stNumberInput"] input:focus{
        border-color: rgba(59,130,246,.55) !important;
        box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important;
    }

    /* Buttons */
    div.stButton > button{
        border-radius: 999px !important;
        padding: .62rem 1rem !important;
        border: 1px solid rgba(148,163,184,.18) !important;
        background: linear-gradient(135deg, rgba(59,130,246,.18), rgba(16,185,129,.12)) !important;
        color: #fff !important;
        transition: transform .14s ease, box-shadow .18s ease, border-color .18s ease;
    }
    div.stButton > button:hover{
        transform: translateY(-1px);
        border-color: rgba(59,130,246,.60) !important;
        box-shadow: 0 14px 35px rgba(0,0,0,.35), 0 0 24px rgba(59,130,246,.20);
    }

    /* Checkboxes / toggles */
    input[type="checkbox"]{ accent-color: var(--blue) !important; }
    div[data-testid="stCheckbox"] label{ color: var(--text) !important; font-weight: 520; }

    /* Expanders */
    div[data-testid="stExpander"] details{
        border: 1px solid rgba(148,163,184,.14);
        border-radius: 14px;
        background: rgba(2,6,23,.18);
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
    }
    div[data-testid="stExpander"] details:hover{
        border-color: rgba(59,130,246,.30);
    }

    /* Dataframes / Editors */
    div[data-testid="stDataFrame"], div[data-testid="stDataEditor"]{
        border: 1px solid rgba(148,163,184,.14);
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 10px 24px rgba(0,0,0,.22);
        background: rgba(2,6,23,.16);
    }

    /* Metrics */
    div[data-testid="stMetric"]{
        background: linear-gradient(145deg, rgba(15,23,42,.62), rgba(30,41,59,.34));
        border: 1px solid rgba(148,163,184,.14);
        padding: 12px 14px;
        border-radius: 16px;
        box-shadow: 0 10px 24px rgba(0,0,0,.22);
    }
    div[data-testid="stMetric"] [data-testid="stMetricLabel"]{
        color: var(--muted) !important;
    }
    div[data-testid="stMetric"] [data-testid="stMetricValue"]{
        font-family: "Space Grotesk", Inter, sans-serif !important;
        font-weight: 700 !important;
    }

    /* Responsive */
    @media (max-width: 768px){
        .sticky-forecast-bar-q1{
            width: 95%;
            padding: 10px 14px;
        }
        .sticky-val{ font-size: 1.05rem; }
        .val-total{ font-size: 1.25rem; }
        .sticky-sep{ margin: 0 6px; }
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
            'font': {'size': 42, 'color': 'white', 'family': 'Inter, sans-serif'},
            'suffix': f"<span style='font-size:18px;color:{bar_color}'> ({percentage:.0f}%)</span>"
        },
        title={
            'text': f"<span style='font-size:13px;color:#94a3b8;letter-spacing:1px'>{title.upper()}</span>",
            'font': {'size': 13}
        },
        gauge={
            'axis': {
                'range': [0, max_range], 
                'tickmode': 'array',
                'tickvals': [0, goal],
                'ticktext': ['0', 'GOAL'],
                'tickfont': {'size': 11, 'color': '#64748b'},
                'showticklabels': True
            },
            'bar': {'color': bar_color, 'thickness': 0.75},
            'bgcolor': "rgba(255,255,255,0.05)",
            'borderwidth': 0,
            'steps': [
                {'range': [0, goal], 'color': "rgba(255,255,255,0.03)"}
            ],
            'threshold': {
                'line': {'color': "#fff", 'width': 2},
                'thickness': 0.85,
                'value': goal
            }
        }
    ))
    
    fig.update_layout(
        height=280,
        margin=dict(l=25, r=25, t=40, b=15),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': 'white', 'family': 'Inter, sans-serif'}
    )
    
    return fig




# ========== EXECUTIVE VISUALS (UI ONLY) ==========
def create_forecast_composition_donut(scheduled, pipeline, reorder, title="Forecast Mix"):
    """
    Donut chart showing the mix of Scheduled / Pipeline / Reorder.
    UI-only: does not change any forecasting calculations.
    """
    scheduled = float(scheduled or 0)
    pipeline = float(pipeline or 0)
    reorder = float(reorder or 0)

    labels = ["Scheduled", "Pipeline", "Reorder"]
    values = [max(scheduled, 0), max(pipeline, 0), max(reorder, 0)]
    total = sum(values)

    # Keep chart stable even if empty
    if total <= 0:
        values = [1, 0, 0]
        total = 0

    colors = ["#34d399", "#60a5fa", "#fbbf24"]

    fig = go.Figure(
        data=[
            go.Pie(
                labels=labels,
                values=values,
                hole=0.65,
                sort=False,
                marker=dict(
                    colors=colors,
                    line=dict(color="rgba(255,255,255,0.12)", width=1),
                ),
                textinfo="percent",
                textposition="outside",
                textfont=dict(size=11, color="#cbd5e1", family="Inter, sans-serif"),
                hovertemplate="%{label}<br>$%{value:,.0f}<extra></extra>",
                pull=[0.02, 0.02, 0.02],
            )
        ]
    )

    fig.add_annotation(
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        align="center",
        text=(
            f"<span style='font-size:24px;color:white;font-family:Inter, sans-serif;'><b>${total:,.0f}</b></span>"
        ),
    )

    fig.update_layout(
        height=320,
        margin=dict(l=40, r=40, t=20, b=50),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.05,
            xanchor="center",
            x=0.5,
            font=dict(color="#cbd5e1", size=11, family="Inter, sans-serif"),
        ),
    )
    return fig


def create_forecast_waterfall(scheduled, pipeline, reorder, goal, title="Path to Goal"):
    """
    Waterfall chart that visualizes how the forecast is built vs. the goal.
    UI-only: does not change any forecasting calculations.
    """
    scheduled = float(scheduled or 0)
    pipeline = float(pipeline or 0)
    reorder = float(reorder or 0)
    goal = float(goal or 0)

    total_forecast = scheduled + pipeline + reorder
    gap_to_goal = goal - total_forecast

    x = ["Scheduled", "Pipeline", "Reorder", "Total Forecast", "Goal"]
    measure = ["relative", "relative", "relative", "total", "total"]
    y = [scheduled, pipeline, reorder, total_forecast, goal]

    # Format text labels - use K for thousands if values are large
    def format_val(v):
        if abs(v) >= 1000000:
            return f"${v/1000000:.1f}M"
        elif abs(v) >= 10000:
            return f"${v/1000:.0f}K"
        else:
            return f"${v:,.0f}"

    fig = go.Figure(
        go.Waterfall(
            name="Forecast",
            orientation="v",
            measure=measure,
            x=x,
            y=y,
            connector=dict(line=dict(color="rgba(148,163,184,0.25)", width=1)),
            text=[format_val(scheduled), format_val(pipeline), format_val(reorder), format_val(total_forecast), format_val(goal)],
            textposition="outside",
            textfont=dict(size=10, color="#cbd5e1"),
            increasing=dict(marker=dict(color="#10b981", line=dict(color="rgba(255,255,255,0.12)", width=1))),
            decreasing=dict(marker=dict(color="#ef4444", line=dict(color="rgba(255,255,255,0.12)", width=1))),
            totals=dict(marker=dict(color="rgba(100,116,139,0.5)", line=dict(color="rgba(255,255,255,0.12)", width=1))),
        )
    )

    gap_label = "Gap to goal:" if gap_to_goal > 0 else "Ahead of goal:"
    gap_color = "#fb7185" if gap_to_goal > 0 else "#34d399"

    fig.add_annotation(
        x=0.5,
        y=1.08,
        xref="paper",
        yref="paper",
        showarrow=False,
        text=f"<span style='color:#94a3b8;font-size:12px'>{gap_label}</span> "
             f"<span style='color:{gap_color};font-size:14px;font-weight:bold'>${abs(gap_to_goal):,.0f}</span>",
        font=dict(size=12, family="Inter, sans-serif"),
    )

    fig.update_layout(
        height=350,
        margin=dict(l=50, r=20, t=70, b=40),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white", family="Inter, sans-serif"),
        xaxis=dict(
            tickfont=dict(color="#cbd5e1", size=10),
            tickangle=0
        ),
        yaxis=dict(
            showgrid=True, 
            gridcolor="rgba(148,163,184,0.1)", 
            zeroline=False, 
            tickfont=dict(color="#94a3b8", size=10),
            tickformat="$,.0f"
        ),
        bargap=0.3,
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
    # Ensure Account Name is preserved if it exists
    if 'Account Name' not in d.columns and 'Deal Name' in d.columns:
        d['Account Name'] = d['Deal Name']  # Fallback to Deal Name if no Account Name
    return d.sort_values('Amount_Numeric', ascending=False) if 'Amount_Numeric' in d.columns else d


# ========== CUSTOMER NAME MATCHING FUNCTIONS ==========
import re

def normalize_customer_name(name):
    """Basic normalization: lowercase, strip whitespace"""
    if pd.isna(name) or name is None:
        return ''
    return str(name).lower().strip()

def extract_customer_keys(name):
    """
    Extract multiple matching keys from a customer name.
    Handles formats like:
    - "Acreage Holdings : Acreage Holdings: New Jersey (NJ)"
    - "Acreage Holdings: New Jersey (NJ)"
    - "Customer Name (State)"
    
    Returns a set of possible matching keys.
    """
    if pd.isna(name) or name is None:
        return set()
    
    name = str(name).strip()
    keys = set()
    
    # Add full normalized name
    normalized = name.lower().strip()
    keys.add(normalized)
    
    # Extract state code in parentheses (e.g., "(NJ)", "(MA)")
    state_match = re.search(r'\(([A-Z]{2})\)', name, re.IGNORECASE)
    if state_match:
        state_code = state_match.group(1).upper()
        keys.add(state_code)
    
    # Split by " : " (NetSuite parent : child format)
    if ' : ' in name:
        parts = name.split(' : ')
        for part in parts:
            keys.add(part.lower().strip())
            # Also add sub-parts split by ":"
            if ':' in part:
                subparts = part.split(':')
                for sp in subparts:
                    clean_sp = sp.strip().lower()
                    if len(clean_sp) > 2:  # Skip very short parts
                        keys.add(clean_sp)
    
    # Split by ":" without spaces
    if ':' in name:
        parts = name.split(':')
        for part in parts:
            clean_part = part.strip().lower()
            if len(clean_part) > 2:
                keys.add(clean_part)
    
    # Extract core name (everything before first colon or parenthesis)
    core_match = re.match(r'^([^:\(]+)', name)
    if core_match:
        core = core_match.group(1).strip().lower()
        if len(core) > 2:
            keys.add(core)
    
    # Remove empty strings
    keys.discard('')
    
    return keys

def customers_match(name1, name2):
    """
    Check if two customer names likely refer to the same customer.
    Uses fuzzy matching with multiple strategies.
    
    Returns True if they match, False otherwise.
    """
    if pd.isna(name1) or pd.isna(name2) or not name1 or not name2:
        return False
    
    n1 = str(name1).lower().strip()
    n2 = str(name2).lower().strip()
    
    # Exact match
    if n1 == n2:
        return True
    
    # One contains the other (for parent : child cases)
    if n1 in n2 or n2 in n1:
        return True
    
    # Extract and compare keys
    keys1 = extract_customer_keys(name1)
    keys2 = extract_customer_keys(name2)
    
    # Check for significant overlap in keys
    # If they share a key that's longer than just a state code, consider it a match
    common_keys = keys1 & keys2
    for key in common_keys:
        if len(key) > 3:  # More than just a state code
            return True
    
    # Check if the "location part" matches (e.g., "New Jersey (NJ)")
    # Extract everything after the last colon
    def get_location_part(name):
        if ':' in name:
            return name.split(':')[-1].strip().lower()
        return name.lower().strip()
    
    loc1 = get_location_part(str(name1))
    loc2 = get_location_part(str(name2))
    
    if loc1 == loc2 and len(loc1) > 3:
        return True
    
    # Check for matching state codes AND similar base names
    state1 = re.search(r'\(([A-Z]{2})\)', str(name1), re.IGNORECASE)
    state2 = re.search(r'\(([A-Z]{2})\)', str(name2), re.IGNORECASE)
    
    if state1 and state2 and state1.group(1).upper() == state2.group(1).upper():
        # Same state code - check if base names are similar
        base1 = re.sub(r'\([^)]+\)', '', str(name1)).replace(':', ' ').lower().split()[0] if name1 else ''
        base2 = re.sub(r'\([^)]+\)', '', str(name2)).replace(':', ' ').lower().split()[0] if name2 else ''
        if base1 and base2 and (base1 in base2 or base2 in base1):
            return True
    
    return False

def find_matching_customer(target_name, customer_list):
    """
    Find a matching customer from a list.
    Returns the matching customer name or None.
    """
    for cust in customer_list:
        if customers_match(target_name, cust):
            return cust
    return None

def build_customer_match_dict(ns_customers, hs_customers):
    """
    Build a dictionary mapping normalized names to all their variations.
    This helps with consistent matching across systems.
    
    Returns dict: {canonical_name: set of all matching names}
    """
    all_names = list(ns_customers) + list(hs_customers)
    match_groups = {}
    used = set()
    
    for name in all_names:
        if name in used:
            continue
        
        # Find all names that match this one
        group = {name}
        for other in all_names:
            if other not in used and customers_match(name, other):
                group.add(other)
        
        # Use the shortest name as canonical (often the HubSpot version)
        canonical = min(group, key=len)
        match_groups[canonical] = group
        used.update(group)
    
    return match_groups


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


def load_item_master(main_dash):
    """
    Load Item Master data for SKU descriptions
    
    Item Master tab columns:
    - Column A: Item (SKU code)
    - Column C: Description
    
    Returns a dictionary mapping SKU -> Description
    """
    
    item_master_df = main_dash.load_google_sheets_data("Item Master", "A:C", version=main_dash.CACHE_VERSION)
    
    if item_master_df.empty:
        return {}
    
    col_names = item_master_df.columns.tolist()
    
    # Column A should be Item/SKU, Column C should be Description
    if len(col_names) < 3:
        return {}
    
    rename_dict = {
        col_names[0]: 'Item',
        col_names[2]: 'Description'
    }
    
    item_master_df = item_master_df.rename(columns=rename_dict)
    
    # Clean Item column
    if 'Item' in item_master_df.columns:
        item_master_df['Item'] = item_master_df['Item'].astype(str).str.strip()
        item_master_df = item_master_df[item_master_df['Item'] != '']
        item_master_df = item_master_df[item_master_df['Item'].str.lower() != 'nan']
    
    # Clean Description column
    if 'Description' in item_master_df.columns:
        item_master_df['Description'] = item_master_df['Description'].astype(str).str.strip()
        item_master_df['Description'] = item_master_df['Description'].replace('nan', '')
    
    # Create lookup dictionary
    sku_to_desc = dict(zip(item_master_df['Item'], item_master_df['Description']))
    
    return sku_to_desc


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


def calculate_customer_product_metrics(historical_df, line_items_df, sku_to_desc=None):
    """
    Calculate metrics by Customer + Product Type combination.
    This gives accurate cadence per product line, not per customer overall.
    
    Args:
        historical_df: Historical orders dataframe
        line_items_df: Line items dataframe
        sku_to_desc: Dictionary mapping SKU codes to descriptions (from Item Master)
    
    Returns DataFrame with:
    - Customer, Product Type, Order count, Revenue, Cadence, Expected Q1 orders
    - Aggregated line item totals (qty, avg rate, total value)
    """
    
    if sku_to_desc is None:
        sku_to_desc = {}
    
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
        top_skus = ""
        
        if so_numbers and not line_items_df.empty:
            product_line_items = line_items_df[line_items_df['SO_Number'].isin(so_numbers)]
            if not product_line_items.empty:
                total_qty = int(product_line_items['Quantity'].sum())
                total_line_value = product_line_items['Line_Total'].sum()
                avg_rate = total_line_value / total_qty if total_qty > 0 else 0
                sku_count = product_line_items['Item'].nunique()
                
                # Get top 3 SKUs by total value, with descriptions from Item Master
                sku_totals = product_line_items.groupby('Item')['Line_Total'].sum().sort_values(ascending=False)
                top_sku_list = sku_totals.head(3).index.tolist()
                
                # Look up descriptions for each SKU
                top_sku_with_desc = []
                for sku in top_sku_list:
                    desc = sku_to_desc.get(sku, '')
                    if desc and desc != sku:
                        # Use description if available and different from SKU
                        top_sku_with_desc.append(desc)
                    else:
                        # Fall back to SKU code if no description
                        top_sku_with_desc.append(sku)
                
                top_skus = ", ".join(top_sku_with_desc) if top_sku_with_desc else ""
        
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
            'Top_SKUs': top_skus,
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
    days_remaining_q1 = calculate_business_days_remaining_q1()
    
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
            <div class="hero-value">{days_remaining_q1} Days</div>
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
        # Expanded range to A:Z to capture Account Name and other columns
        deals_df = main_dash.load_google_sheets_data("Copy of All Reps All Pipelines", "A:Z", version=main_dash.CACHE_VERSION)
        
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
                elif col == 'Q2 2026 Spillover':
                    rename_dict[col] = 'Q2 2026 Spillover'
                elif col == 'Q1 2026 Spillover':
                    # Handle old column name - rename to new name
                    rename_dict[col] = 'Q2 2026 Spillover'
                elif col == 'Account Name' or col == 'Associated Company':
                    rename_dict[col] = 'Account Name'
                elif col == 'Company':
                    rename_dict[col] = 'Account Name'
            
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
    # For Q1 2026 dashboard, the primary Q1 data is in the "date" buckets:
    # - pf_date_ext + pf_date_int = PF orders with Q1 2026 Promise/Projected dates
    # - pa_date = PA orders with PA Date in Q1 2026
    # Additional buckets that may convert to Q1 revenue:
    # - pf_nodate = PF orders with no date (could ship anytime)
    # - pa_nodate = PA with no date (<2 weeks old) - could convert
    # - pa_old = PA orders >2 weeks old - stale but potential
    # Spillover buckets (for reference):
    # - pf_q4_spillover/pa_q4_spillover = Q4 2025 carryover orders
    # - pf_q2_spillover/pa_q2_spillover = Q2 2026 forward spillover
    
    # Aggregate data from all active reps
    all_pf_q1 = []  # PF with Q1 dates (primary quarter)
    all_pa_q1 = []  # PA with Q1 dates (primary quarter)
    all_pf_nodate = []
    all_pa_nodate = []
    all_pa_old = []
    
    total_pf_amount = 0
    total_pa_amount = 0
    total_pf_nodate_amount = 0
    total_pa_nodate_amount = 0
    total_pa_old_amount = 0
    
    for r in active_team_reps:
        so_cats = categorize_sales_orders(sales_orders_df, r)
        
        # PF with Q1 2026 dates (External + Internal)
        if not so_cats['pf_date_ext'].empty:
            all_pf_q1.append(so_cats['pf_date_ext'])
            total_pf_amount += so_cats['pf_date_ext_amount']
        if not so_cats['pf_date_int'].empty:
            all_pf_q1.append(so_cats['pf_date_int'])
            total_pf_amount += so_cats['pf_date_int_amount']
        
        # PA with Q1 2026 PA dates
        if not so_cats['pa_date'].empty:
            all_pa_q1.append(so_cats['pa_date'])
            total_pa_amount += so_cats['pa_date_amount']
        
        # PF No Date (External + Internal combined)
        if not so_cats['pf_nodate_ext'].empty:
            all_pf_nodate.append(so_cats['pf_nodate_ext'])
            total_pf_nodate_amount += so_cats['pf_nodate_ext_amount']
        if not so_cats['pf_nodate_int'].empty:
            all_pf_nodate.append(so_cats['pf_nodate_int'])
            total_pf_nodate_amount += so_cats['pf_nodate_int_amount']
        
        # PA No Date (<2 weeks old) - moved pa_date to primary Q1 bucket above
        if not so_cats['pa_nodate'].empty:
            all_pa_nodate.append(so_cats['pa_nodate'])
            total_pa_nodate_amount += so_cats['pa_nodate_amount']
        
        # PA Old (>2 weeks)
        if not so_cats['pa_old'].empty:
            all_pa_old.append(so_cats['pa_old'])
            total_pa_old_amount += so_cats['pa_old_amount']
    
    # Combine into single dataframes
    combined_pf = pd.concat(all_pf_q1, ignore_index=True) if all_pf_q1 else pd.DataFrame()
    combined_pa = pd.concat(all_pa_q1, ignore_index=True) if all_pa_q1 else pd.DataFrame()
    combined_pf_nodate = pd.concat(all_pf_nodate, ignore_index=True) if all_pf_nodate else pd.DataFrame()
    combined_pa_nodate = pd.concat(all_pa_nodate, ignore_index=True) if all_pa_nodate else pd.DataFrame()
    combined_pa_old = pd.concat(all_pa_old, ignore_index=True) if all_pa_old else pd.DataFrame()
    
    # Map to Q1 categories - organized by certainty level
    ns_categories = {
        'PF_Q1': {'label': '📦 PF (Q1 2026 Date)', 'df': combined_pf, 'amount': total_pf_amount},
        'PA_Q1': {'label': '⏳ PA (Q1 2026 PA Date)', 'df': combined_pa, 'amount': total_pa_amount},
        'PF_NoDate': {'label': '📦 PF (No Date)', 'df': combined_pf_nodate, 'amount': total_pf_nodate_amount},
        'PA_NoDate': {'label': '⏳ PA (No Date)', 'df': combined_pa_nodate, 'amount': total_pa_nodate_amount},
        'PA_Old': {'label': '⚠️ PA (>2 Weeks)', 'df': combined_pa_old, 'amount': total_pa_old_amount},
    }
    
    # Format for display
    ns_dfs = {
        'PF_Q1': format_ns_view(combined_pf, 'Promise'),
        'PA_Q1': format_ns_view(combined_pa, 'PA_Date'),
        'PF_NoDate': format_ns_view(combined_pf_nodate, 'Promise'),
        'PA_NoDate': format_ns_view(combined_pa_nodate, 'PA_Date'),
        'PA_Old': format_ns_view(combined_pa_old, 'PA_Date'),
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
            
            # Check for spillover column (handles both old and new column names)
            spillover_col = None
            if 'Q2 2026 Spillover' in rep_deals.columns:
                spillover_col = 'Q2 2026 Spillover'
            elif 'Q1 2026 Spillover' in rep_deals.columns:
                spillover_col = 'Q1 2026 Spillover'
            
            # Q4 2025 Spillover - deals with Q4 close date that haven't been fulfilled
            # These are carryover deals from Q4 that may still convert in Q1
            q4_close_mask = (rep_deals['Close Date'] >= Q4_2025_START) & (rep_deals['Close Date'] <= Q4_2025_END)
            
            if spillover_col:
                # Q4 Spillover = Q4 close date AND marked as Q4 2025 spillover
                q4_spillover = rep_deals[q4_close_mask & (rep_deals[spillover_col] == 'Q4 2025')]
            else:
                q4_spillover = pd.DataFrame()
            
            # Debug info
            with st.expander("🔧 Debug: HubSpot Deal Counts"):
                if is_team_view:
                    st.write(f"**Team View - Reps included:** {', '.join(active_team_reps)}")
                st.write(f"**Total deals loaded:** {len(rep_deals)}")
                st.write(f"**Q1 Close Date deals:** {len(q1_deals)} (Close Date in Jan-Mar 2026)")
                st.write(f"**Q4 Spillover deals:** {len(q4_spillover)} (Q4 Close Date + Spillover flag)")
                st.write(f"**Spillover column found:** {spillover_col or 'None'}")
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
            st.caption("Orders in NetSuite - select which to include in forecast")
            
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
    # STEP 4: APPLY PIPELINE PROBABILITY WEIGHTING
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    step4_title = "Apply Pipeline Probability" if is_team_view else f"{first_name}, Apply Pipeline Probability?"
    st.markdown(f"### 🎲 Step 4: {step4_title}")
    
    st.markdown(f"""
    <div style="background: rgba(139, 92, 246, 0.1); border-left: 4px solid #8b5cf6; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-size: 1rem;">
            Want to apply <strong>probability weighting</strong> to your pipeline deals? This adjusts values based on likelihood to close.
        </div>
        <div style="color: #94a3b8; margin-top: 8px; font-size: 0.9rem;">
            <strong>Default weights:</strong> Expect = 100% | Commit = 85% | Best Case = 50% | Opportunity = 25%
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state for pipeline weighting
    pipeline_weight_key = f"pipeline_weighting_{rep_name}"
    if pipeline_weight_key not in st.session_state:
        st.session_state[pipeline_weight_key] = {
            'enabled': False,
            'expect': 100,
            'commit': 85,
            'best_case': 50,
            'opportunity': 25
        }
    
    # Toggle for enabling weighting
    pw_col1, pw_col2 = st.columns([2, 3])
    
    with pw_col1:
        apply_pipeline_weight = st.toggle(
            "Apply probability weighting to pipeline",
            value=st.session_state[pipeline_weight_key]['enabled'],
            key=f"toggle_pipeline_weight_{rep_name}",
            help="When enabled, pipeline values are multiplied by their probability percentage"
        )
        st.session_state[pipeline_weight_key]['enabled'] = apply_pipeline_weight
    
    if apply_pipeline_weight:
        with pw_col2:
            st.caption("Customize weights (or keep defaults):")
        
        # Weight customization
        wt_col1, wt_col2, wt_col3, wt_col4 = st.columns(4)
        
        with wt_col1:
            expect_wt = st.number_input(
                "Expect %", 
                min_value=0, max_value=100, 
                value=st.session_state[pipeline_weight_key]['expect'],
                key=f"wt_expect_{rep_name}"
            )
            st.session_state[pipeline_weight_key]['expect'] = expect_wt
        
        with wt_col2:
            commit_wt = st.number_input(
                "Commit %", 
                min_value=0, max_value=100, 
                value=st.session_state[pipeline_weight_key]['commit'],
                key=f"wt_commit_{rep_name}"
            )
            st.session_state[pipeline_weight_key]['commit'] = commit_wt
        
        with wt_col3:
            best_case_wt = st.number_input(
                "Best Case %", 
                min_value=0, max_value=100, 
                value=st.session_state[pipeline_weight_key]['best_case'],
                key=f"wt_bestcase_{rep_name}"
            )
            st.session_state[pipeline_weight_key]['best_case'] = best_case_wt
        
        with wt_col4:
            opp_wt = st.number_input(
                "Opportunity %", 
                min_value=0, max_value=100, 
                value=st.session_state[pipeline_weight_key]['opportunity'],
                key=f"wt_opp_{rep_name}"
            )
            st.session_state[pipeline_weight_key]['opportunity'] = opp_wt
        
        # Show impact preview
        raw_pipeline_total = sum(
            df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0 
            for k, df in export_buckets.items() if k in hs_categories
        )
        
        # Calculate weighted total
        weighted_pipeline_total = 0
        for key, df in export_buckets.items():
            if key in hs_categories and not df.empty and 'Amount_Numeric' in df.columns:
                if 'Expect' in key:
                    weighted_pipeline_total += df['Amount_Numeric'].sum() * (expect_wt / 100)
                elif 'Commit' in key:
                    weighted_pipeline_total += df['Amount_Numeric'].sum() * (commit_wt / 100)
                elif 'BestCase' in key:
                    weighted_pipeline_total += df['Amount_Numeric'].sum() * (best_case_wt / 100)
                elif 'Opp' in key:
                    weighted_pipeline_total += df['Amount_Numeric'].sum() * (opp_wt / 100)
        
        st.markdown(f"""
        <div style="background: rgba(139, 92, 246, 0.15); padding: 12px 15px; border-radius: 8px; margin-top: 10px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div>
                    <span style="color: #94a3b8;">Pipeline Raw:</span> 
                    <strong style="color: #60a5fa;">${raw_pipeline_total:,.0f}</strong>
                </div>
                <div style="font-size: 1.2rem;">→</div>
                <div>
                    <span style="color: #94a3b8;">Weighted:</span> 
                    <strong style="color: #a78bfa;">${weighted_pipeline_total:,.0f}</strong>
                </div>
                <div>
                    <span style="color: #94a3b8;">Reduction:</span> 
                    <strong style="color: #f87171;">${raw_pipeline_total - weighted_pipeline_total:,.0f}</strong>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.caption("Pipeline values will be used at face value (100%).")
    
    # ═══════════════════════════════════════════════════════════════════════════
    # SECTION 3: REORDER FORECAST (Historical Analysis)
    # ═══════════════════════════════════════════════════════════════════════════
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 5: REORDER OPPORTUNITIES
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    st.markdown(f"### 🔄 Step 5: {'Team Reorder Opportunities' if is_team_view else f'{first_name}, Find Your Reorder Opportunities'}")
    
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
    
    # Initialize data variables
    historical_df = pd.DataFrame()
    invoices_df = pd.DataFrame()
    line_items_df = pd.DataFrame()
    sku_to_desc = {}
    
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
        
        # Load line items - THIS IS THE KEY DATA
        line_items_df = load_line_items(main_dash)
        
        # Load Item Master for SKU descriptions
        sku_to_desc = load_item_master(main_dash)
    
    # Debug section - EXPANDED
    with st.expander("🔧 Debug: Data Loading Status", expanded=False):
        col1, col2, col3 = st.columns(3)
        
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
        
        with col3:
            st.write("**Item Master (SKU Descriptions):**")
            if not sku_to_desc:
                st.warning("⚠️ No Item Master loaded - check tab name 'Item Master'")
            else:
                st.success(f"✅ {len(sku_to_desc)} SKU descriptions loaded")
                # Show sample mappings
                sample_items = list(sku_to_desc.items())[:5]
                for sku, desc in sample_items:
                    st.write(f"• {sku}: {desc[:50]}..." if len(desc) > 50 else f"• {sku}: {desc}")
        
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
            if not df.empty:
                # Try multiple possible columns for customer name in HubSpot
                # Priority: Account Name > Associated Company > Company > Deal Name
                customer_col = None
                for col in ['Account Name', 'Associated Company', 'Company', 'Deal Name']:
                    if col in df.columns:
                        customer_col = col
                        break
                
                if customer_col:
                    pipeline_customers.update(df[customer_col].dropna().tolist())
        
        # Also check the raw deals_df for any customers with pipeline deals
        if not deals_df.empty:
            # Get customers from deals that are in our filtered HS categories (Q1 deals)
            for col in ['Account Name', 'Associated Company', 'Company', 'Deal Name']:
                if col in deals_df.columns:
                    # Filter to deals owned by active reps
                    if 'Deal Owner' in deals_df.columns:
                        rep_deals = deals_df[deals_df['Deal Owner'].isin(active_team_reps)]
                        pipeline_customers.update(rep_deals[col].dropna().tolist())
                    break
        
        # Debug: Show what's being excluded
        with st.expander("🔧 Debug: Customer Exclusion List", expanded=False):
            st.write(f"**Pending NS Customers ({len(pending_customers)}):**")
            st.write(sorted(list(pending_customers))[:30])
            st.write(f"**Pipeline HS Customers ({len(pipeline_customers)}):**")
            st.write(sorted(list(pipeline_customers))[:30])
            
            # Show column info
            st.write("**HubSpot DataFrame Columns:**")
            for key in list(hs_categories.keys())[:2]:
                df = hs_dfs.get(key, pd.DataFrame())
                if not df.empty:
                    st.write(f"{key}: {df.columns.tolist()}")
                    if 'Account Name' in df.columns:
                        st.write(f"Sample Account Names: {df['Account Name'].head(5).tolist()}")
                    break
            
            if not deals_df.empty:
                st.write(f"**Raw deals_df columns:** {deals_df.columns.tolist()}")
                if 'Account Name' in deals_df.columns:
                    st.write(f"Sample Account Names from deals_df: {deals_df['Account Name'].dropna().head(5).tolist()}")
        
        # Build combined set of all NS/HS customer names for fuzzy matching
        all_pipeline_customers = pending_customers | pipeline_customers
        
        # Create a function to check if a historical customer has a match in NS/HS
        def has_pipeline_match(hist_customer):
            """Check if a historical customer has a matching NS/HS entry using fuzzy matching"""
            for pipeline_cust in all_pipeline_customers:
                if customers_match(hist_customer, pipeline_cust):
                    return True
            return False
        
        # Calculate NEW product-level metrics (with SKU descriptions from Item Master)
        product_metrics_df = calculate_customer_product_metrics(historical_df, line_items_df, sku_to_desc)
        
        if product_metrics_df.empty:
            st.warning("No product metrics calculated")
        else:
            # Filter out customers with pending orders/deals using FUZZY matching
            unique_hist_customers = product_metrics_df['Customer'].unique().tolist()
            
            # Build a mapping of which historical customers have pipeline matches
            customers_with_pipeline = set()
            customer_pipeline_matches = {}  # For debugging
            
            for hist_cust in unique_hist_customers:
                for pipeline_cust in all_pipeline_customers:
                    if customers_match(hist_cust, pipeline_cust):
                        customers_with_pipeline.add(hist_cust)
                        if hist_cust not in customer_pipeline_matches:
                            customer_pipeline_matches[hist_cust] = []
                        customer_pipeline_matches[hist_cust].append(pipeline_cust)
                        break
            
            # Debug: Show fuzzy matching results
            with st.expander("🔧 Debug: Fuzzy Customer Matching", expanded=False):
                st.write(f"**Historical Customers:** {len(unique_hist_customers)}")
                st.write(f"**NS/HS Pipeline Customers:** {len(all_pipeline_customers)}")
                st.write(f"**Matched (will be excluded from reorder):** {len(customers_with_pipeline)}")
                
                if customer_pipeline_matches:
                    st.write("**Sample Matches:**")
                    for hist, matches in list(customer_pipeline_matches.items())[:10]:
                        st.write(f"• '{hist}' ↔ '{matches[0]}'")
                
                # Show unmatched customers that might need attention
                unmatched_hist = [c for c in unique_hist_customers if c not in customers_with_pipeline]
                if unmatched_hist:
                    st.write(f"**Unmatched Historical Customers ({len(unmatched_hist)}):**")
                    st.write(unmatched_hist[:15])
            
            opportunities_df = product_metrics_df[
                ~product_metrics_df['Customer'].isin(customers_with_pipeline)
            ].copy()
            
            # Debug: Show filtering stats
            total_before = product_metrics_df['Customer'].nunique()
            excluded_count = len(customers_with_pipeline)
            
            # Update debug expander with exclusion stats
            with st.expander("🔧 Debug: Customer Exclusion Details", expanded=False):
                st.write(f"**Total Historical Customers:** {total_before}")
                st.write(f"**Customers Excluded (fuzzy match to NS/HS):** {excluded_count}")
                st.write(f"**Customers Remaining (opportunities):** {opportunities_df['Customer'].nunique()}")
                
                # Show which specific customers were excluded
                st.write(f"**Excluded Customers:** {sorted(list(customers_with_pipeline))[:20]}")
            
            if opportunities_df.empty and product_metrics_df.empty:
                st.success("✅ No historical orders found for analysis.")
            else:
                # === CUSTOMER-CENTRIC REORDER SECTION (LIST VIEW) ===
                
                # Calculate 2025 metrics per customer
                amount_col = 'Invoice_Amount' if 'Invoice_Amount' in historical_df.columns else 'Amount'
                customer_summary = historical_df.groupby('Customer').agg({
                    amount_col: 'sum',
                    'SO_Number': 'count',
                    'Order Start Date': 'max'
                }).reset_index()
                customer_summary.columns = ['Customer', 'Revenue_2025', 'Order_Count', 'Last_Order']
                customer_summary['Days_Since'] = (pd.Timestamp.now() - customer_summary['Last_Order']).dt.days
                customer_summary = customer_summary.sort_values('Revenue_2025', ascending=False)
                
                # Determine status for each customer
                def get_status(row):
                    if row['Order_Count'] >= 3:
                        return ('🟢', 'Likely', 0.75)
                    elif row['Order_Count'] >= 2:
                        return ('🟡', 'Possible', 0.50)
                    return ('⚪', 'Long Shot', 0.25)
                
                customer_summary[['Status_Emoji', 'Status_Text', 'Confidence']] = customer_summary.apply(
                    lambda r: pd.Series(get_status(r)), axis=1
                )
                
                # Check NS/HS status for each customer using FUZZY MATCHING
                def has_ns_match(cust_name):
                    """Check if customer has a fuzzy match in pending NS customers"""
                    for ns_cust in pending_customers:
                        if customers_match(cust_name, ns_cust):
                            return True
                    return False
                
                def has_hs_match(cust_name):
                    """Check if customer has a fuzzy match in pipeline HS customers"""
                    for hs_cust in pipeline_customers:
                        if customers_match(cust_name, hs_cust):
                            return True
                    return False
                
                customer_summary['Has_NS'] = customer_summary['Customer'].apply(has_ns_match)
                customer_summary['Has_HS'] = customer_summary['Customer'].apply(has_hs_match)
                customer_summary['Is_Opportunity'] = ~(customer_summary['Has_NS'] | customer_summary['Has_HS'])
                
                # Initialize session state
                manual_entries_key = f"manual_entries_{rep_name}"
                if manual_entries_key not in st.session_state:
                    st.session_state[manual_entries_key] = {}
                
                reorder_selections_key = f"reorder_selections_{rep_name}"
                if reorder_selections_key not in st.session_state:
                    st.session_state[reorder_selections_key] = {}
                
                # === CONTROLS ===
                ctrl_col1, ctrl_col2, ctrl_col3 = st.columns([2, 1, 1])
                
                with ctrl_col1:
                    search_term = st.text_input(
                        "🔍 Search Customers",
                        placeholder="Type to filter customers...",
                        key=f"cust_search_{rep_name}"
                    )
                
                with ctrl_col2:
                    filter_option = st.selectbox(
                        "Filter",
                        options=["Reorder Opportunities Only", "All Customers", "Has Active Orders/Deals"],
                        key=f"cust_filter_{rep_name}"
                    )
                
                with ctrl_col3:
                    sort_option = st.selectbox(
                        "Sort by",
                        options=["2025 Revenue (High→Low)", "2025 Revenue (Low→High)", "Last Order (Recent)", "Last Order (Oldest)"],
                        key=f"cust_sort_{rep_name}"
                    )
                
                # Apply filters
                filtered_customers = customer_summary.copy()
                
                if search_term:
                    filtered_customers = filtered_customers[
                        filtered_customers['Customer'].str.lower().str.contains(search_term.lower(), na=False)
                    ]
                
                if filter_option == "Reorder Opportunities Only":
                    filtered_customers = filtered_customers[filtered_customers['Is_Opportunity']]
                elif filter_option == "Has Active Orders/Deals":
                    filtered_customers = filtered_customers[~filtered_customers['Is_Opportunity']]
                
                # Apply sort
                if sort_option == "2025 Revenue (High→Low)":
                    filtered_customers = filtered_customers.sort_values('Revenue_2025', ascending=False)
                elif sort_option == "2025 Revenue (Low→High)":
                    filtered_customers = filtered_customers.sort_values('Revenue_2025', ascending=True)
                elif sort_option == "Last Order (Recent)":
                    filtered_customers = filtered_customers.sort_values('Days_Since', ascending=True)
                else:
                    filtered_customers = filtered_customers.sort_values('Days_Since', ascending=False)
                
                # === PRE-COMPUTE ALL REORDER OPPORTUNITIES ===
                all_reorder_opps = []
                for _, cust_row in customer_summary[customer_summary['Is_Opportunity']].iterrows():
                    customer_name = cust_row['Customer']
                    confidence = cust_row['Confidence']
                    cust_products = product_metrics_df[product_metrics_df['Customer'] == customer_name].copy()
                    
                    # Get active types for this customer
                    active_types = set()
                    for key in ns_categories.keys():
                        df = ns_dfs.get(key, pd.DataFrame())
                        if not df.empty and 'Customer' in df.columns:
                            for _, row in df.iterrows():
                                if customers_match(customer_name, row.get('Customer', '')):
                                    active_types.add(str(row.get('Type', '')).lower().strip())
                    for key in hs_categories.keys():
                        df = hs_dfs.get(key, pd.DataFrame())
                        if not df.empty:
                            for _, row in df.iterrows():
                                for col in ['Account Name', 'Associated Company', 'Company']:
                                    if col in df.columns and pd.notna(row.get(col)):
                                        if customers_match(customer_name, row.get(col)):
                                            active_types.add(str(row.get('Type', '')).lower().strip())
                                            break
                    
                    # Find uncovered products
                    for _, prod_row in cust_products.iterrows():
                        prod_type = prod_row['Product_Type']
                        prod_lower = str(prod_type).lower().strip()
                        is_covered = any(prod_lower in at or at in prod_lower for at in active_types if at)
                        
                        if not is_covered:
                            all_reorder_opps.append({
                                'Customer': customer_name,
                                'Product_Type': prod_type,
                                'Q1_Value': int(prod_row['Q1_Value']),
                                'Total_Revenue': prod_row['Total_Revenue'],
                                'Expected_Orders': prod_row['Expected_Orders_Q1'],
                                'Confidence': confidence,
                                'Confidence_Tier': prod_row['Confidence_Tier'],
                                'Confidence_Pct': prod_row['Confidence_Pct'],
                                'Top_SKUs': prod_row.get('Top_SKUs', ''),
                                'Days_Since': prod_row['Days_Since_Last']
                            })
                
                # Summary stats
                total_reorder_opp_value = sum(o['Q1_Value'] for o in all_reorder_opps)
                total_reorder_weighted = sum(o['Q1_Value'] * o['Confidence'] for o in all_reorder_opps)
                
                # Count currently selected
                currently_selected = sum(1 for k, v in st.session_state.get(reorder_selections_key, {}).items() if v.get('selected', False))
                selected_value = sum(v.get('value', 0) for k, v in st.session_state.get(reorder_selections_key, {}).items() if v.get('selected', False))
                
                st.markdown(f"""
                <div style="display: flex; gap: 15px; margin: 15px 0; flex-wrap: wrap;">
                    <div style="background: rgba(16, 185, 129, 0.1); padding: 10px 15px; border-radius: 8px;">
                        <span style="color: #94a3b8;">Showing:</span> <strong style="color: white;">{len(filtered_customers)}</strong> customers
                    </div>
                    <div style="background: rgba(251, 191, 36, 0.1); padding: 10px 15px; border-radius: 8px;">
                        <span style="color: #94a3b8;">Reorder Opps:</span> <strong style="color: #fbbf24;">{len(all_reorder_opps)}</strong> products
                    </div>
                    <div style="background: rgba(251, 191, 36, 0.15); padding: 10px 15px; border-radius: 8px;">
                        <span style="color: #94a3b8;">Potential:</span> <strong style="color: #fbbf24;">${total_reorder_opp_value:,.0f}</strong>
                    </div>
                    <div style="background: rgba(16, 185, 129, 0.2); padding: 10px 15px; border-radius: 8px; border: 1px solid #10b981;">
                        <span style="color: #94a3b8;">Selected:</span> <strong style="color: #10b981;">{currently_selected}</strong> (${selected_value:,.0f})
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # === QUICK SELECT ALL REORDER OPPORTUNITIES ===
                if all_reorder_opps:
                    # Auto-expand when filter is "Reorder Opportunities Only"
                    expand_quick_actions = (filter_option == "Reorder Opportunities Only")
                    
                    with st.expander("⚡ Quick Actions: Select Reorder Opportunities", expanded=expand_quick_actions):
                        st.markdown("**Select multiple reorder opportunities at once:**")
                        
                        # Show summary table of all opportunities
                        opp_df = pd.DataFrame(all_reorder_opps)
                        opp_df_display = opp_df[['Customer', 'Product_Type', 'Q1_Value', 'Confidence_Tier', 'Days_Since']].copy()
                        opp_df_display['Q1_Value'] = opp_df_display['Q1_Value'].apply(lambda x: f"${x:,.0f}")
                        opp_df_display.columns = ['Customer', 'Product', 'Q1 Projection', 'Confidence', 'Days Since']
                        
                        st.dataframe(opp_df_display, use_container_width=True, hide_index=True, height=min(300, 35 + len(opp_df_display) * 35))
                        
                        # Count by confidence tier
                        likely_opps = [o for o in all_reorder_opps if o['Confidence_Tier'] == 'Likely']
                        possible_opps = [o for o in all_reorder_opps if o['Confidence_Tier'] == 'Possible']
                        longshot_opps = [o for o in all_reorder_opps if o['Confidence_Tier'] == 'Long Shot']
                        
                        st.caption(f"🟢 Likely: {len(likely_opps)} (${sum(o['Q1_Value'] for o in likely_opps):,.0f}) | 🟡 Possible: {len(possible_opps)} (${sum(o['Q1_Value'] for o in possible_opps):,.0f}) | 🔴 Long Shot: {len(longshot_opps)} (${sum(o['Q1_Value'] for o in longshot_opps):,.0f})")
                        
                        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)
                        
                        with btn_col1:
                            if st.button("✅ Select ALL", key=f"select_all_{rep_name}", type="primary", help="Select all reorder opportunities"):
                                for opp in all_reorder_opps:
                                    selection_key = f"{opp['Customer']}|{opp['Product_Type']}"
                                    checkbox_key = f"chk_{selection_key}_{rep_name}"
                                    # Update both the data store AND the checkbox widget key
                                    st.session_state[reorder_selections_key][selection_key] = {
                                        'selected': True,
                                        'value': opp['Q1_Value'],
                                        'confidence': opp['Confidence'],
                                        'product_type': opp['Product_Type'],
                                        'customer': opp['Customer'],
                                        'top_skus': opp['Top_SKUs'],
                                        'historical_total': opp['Total_Revenue'],
                                        'expected_orders': opp['Expected_Orders'],
                                        'confidence_tier': opp['Confidence_Tier']
                                    }
                                    st.session_state[checkbox_key] = True
                                st.success(f"Selected {len(all_reorder_opps)} opportunities!")
                                st.rerun()
                        
                        with btn_col2:
                            if st.button("🟢 Likely Only", key=f"select_likely_{rep_name}", help="Select only 'Likely' (3+ orders)"):
                                likely_count = 0
                                for opp in likely_opps:
                                    selection_key = f"{opp['Customer']}|{opp['Product_Type']}"
                                    checkbox_key = f"chk_{selection_key}_{rep_name}"
                                    st.session_state[reorder_selections_key][selection_key] = {
                                        'selected': True,
                                        'value': opp['Q1_Value'],
                                        'confidence': opp['Confidence'],
                                        'product_type': opp['Product_Type'],
                                        'customer': opp['Customer'],
                                        'top_skus': opp['Top_SKUs'],
                                        'historical_total': opp['Total_Revenue'],
                                        'expected_orders': opp['Expected_Orders'],
                                        'confidence_tier': opp['Confidence_Tier']
                                    }
                                    st.session_state[checkbox_key] = True
                                    likely_count += 1
                                st.success(f"Selected {likely_count} 'Likely' opportunities!")
                                st.rerun()
                        
                        with btn_col3:
                            if st.button("🟢🟡 Likely + Possible", key=f"select_likely_possible_{rep_name}", help="Select 'Likely' and 'Possible' (2+ orders)"):
                                count = 0
                                for opp in likely_opps + possible_opps:
                                    selection_key = f"{opp['Customer']}|{opp['Product_Type']}"
                                    checkbox_key = f"chk_{selection_key}_{rep_name}"
                                    st.session_state[reorder_selections_key][selection_key] = {
                                        'selected': True,
                                        'value': opp['Q1_Value'],
                                        'confidence': opp['Confidence'],
                                        'product_type': opp['Product_Type'],
                                        'customer': opp['Customer'],
                                        'top_skus': opp['Top_SKUs'],
                                        'historical_total': opp['Total_Revenue'],
                                        'expected_orders': opp['Expected_Orders'],
                                        'confidence_tier': opp['Confidence_Tier']
                                    }
                                    st.session_state[checkbox_key] = True
                                    count += 1
                                st.success(f"Selected {count} opportunities!")
                                st.rerun()
                        
                        with btn_col4:
                            if st.button("🗑️ Clear All", key=f"clear_all_{rep_name}", help="Clear all selections"):
                                # Clear both the data store AND the checkbox widget keys
                                for opp in all_reorder_opps:
                                    selection_key = f"{opp['Customer']}|{opp['Product_Type']}"
                                    checkbox_key = f"chk_{selection_key}_{rep_name}"
                                    if checkbox_key in st.session_state:
                                        st.session_state[checkbox_key] = False
                                st.session_state[reorder_selections_key] = {}
                                st.success("Cleared all selections!")
                                st.rerun()
                
                st.markdown("---")
                
                # === CUSTOMER LIST ===
                for _, cust_row in filtered_customers.iterrows():
                    customer_name = cust_row['Customer']
                    cust_revenue = cust_row['Revenue_2025']
                    cust_orders = cust_row['Order_Count']
                    days_since = cust_row['Days_Since']
                    status_emoji = cust_row['Status_Emoji']
                    status_text = cust_row['Status_Text']
                    confidence = cust_row['Confidence']
                    has_ns = cust_row['Has_NS']
                    has_hs = cust_row['Has_HS']
                    is_opportunity = cust_row['Is_Opportunity']
                    
                    # Build status tags
                    tags_html = ""
                    if has_ns:
                        tags_html += "<span style='background: rgba(16, 185, 129, 0.2); color: #34d399; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; margin-right: 5px;'>📦 NS</span>"
                    if has_hs:
                        tags_html += "<span style='background: rgba(96, 165, 250, 0.2); color: #60a5fa; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; margin-right: 5px;'>🎯 HS</span>"
                    if is_opportunity:
                        tags_html += "<span style='background: rgba(251, 191, 36, 0.2); color: #fbbf24; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem;'>🔄 Reorder</span>"
                    
                    # Customer expander with summary in header
                    expander_label = f"{status_emoji} **{customer_name}** — ${cust_revenue:,.0f} • {cust_orders} orders • {days_since}d ago"
                    
                    with st.expander(expander_label, expanded=False):
                        # Status tags
                        st.markdown(tags_html, unsafe_allow_html=True)
                        
                        # Get this customer's product breakdown
                        cust_products = product_metrics_df[product_metrics_df['Customer'] == customer_name].copy()
                        
                        # Find NS orders for this customer using FUZZY MATCHING
                        cust_ns_orders = []
                        for key in ns_categories.keys():
                            df = ns_dfs.get(key, pd.DataFrame())
                            if not df.empty and 'Customer' in df.columns:
                                for _, row in df.iterrows():
                                    ns_cust = row.get('Customer', '')
                                    if customers_match(customer_name, ns_cust):
                                        cust_ns_orders.append({
                                            'SO': row.get('SO #', ''),
                                            'Type': row.get('Type', key),
                                            'Amount': float(row.get('Amount', 0) or 0),
                                            'Date': str(row.get('Ship Date', ''))[:10],
                                            'NS_Customer': ns_cust  # Keep original name for debugging
                                        })
                        
                        # Find HS deals for this customer using FUZZY MATCHING
                        cust_hs_deals = []
                        for key in hs_categories.keys():
                            df = hs_dfs.get(key, pd.DataFrame())
                            if not df.empty:
                                # Try multiple customer columns
                                for _, row in df.iterrows():
                                    hs_cust = None
                                    for col in ['Account Name', 'Associated Company', 'Company', 'Deal Name']:
                                        if col in df.columns and pd.notna(row.get(col)):
                                            hs_cust = row.get(col)
                                            break
                                    
                                    if hs_cust and customers_match(customer_name, hs_cust):
                                        cust_hs_deals.append({
                                            'Deal': row.get('Deal Name', ''),
                                            'Type': row.get('Product Type', key),
                                            'Amount': float(row.get('Amount_Numeric', row.get('Amount', 0)) or 0),
                                            'Close': str(row.get('Close', row.get('Close Date', '')))[:10],
                                            'HS_Customer': hs_cust  # Keep original name for debugging
                                        })
                        
                        # Three columns layout
                        col1, col2, col3 = st.columns(3)
                        
                        # Column 1: Active NS Orders
                        with col1:
                            st.markdown("**📦 NetSuite Orders**")
                            if cust_ns_orders:
                                ns_total = sum(o['Amount'] for o in cust_ns_orders)
                                st.caption(f"Total: ${ns_total:,.0f}")
                                for order in cust_ns_orders:
                                    st.markdown(f"• {order['Type']}: ${order['Amount']:,.0f}")
                            else:
                                st.caption("None")
                        
                        # Column 2: Pipeline Deals
                        with col2:
                            st.markdown("**🎯 HubSpot Deals**")
                            if cust_hs_deals:
                                hs_total = sum(d['Amount'] for d in cust_hs_deals)
                                st.caption(f"Total: ${hs_total:,.0f}")
                                for deal in cust_hs_deals:
                                    st.markdown(f"• {deal['Type']}: ${deal['Amount']:,.0f}")
                            else:
                                st.caption("None")
                        
                        # Column 3: Reorder Opportunities
                        with col3:
                            st.markdown("**🔄 Reorder Opportunities**")
                            
                            # Add tooltip explaining the logic
                            st.caption("💡 Projected Q1 value based on order history")
                            
                            # Get product types with active NS/HS
                            active_types = set()
                            for o in cust_ns_orders:
                                active_types.add(str(o['Type']).lower().strip())
                            for d in cust_hs_deals:
                                active_types.add(str(d['Type']).lower().strip())
                            
                            reorder_opps = []
                            for _, prod_row in cust_products.iterrows():
                                prod_type = prod_row['Product_Type']
                                prod_lower = str(prod_type).lower().strip()
                                
                                # Check if covered
                                is_covered = any(prod_lower in at or at in prod_lower for at in active_types if at)
                                
                                if not is_covered:
                                    reorder_opps.append(prod_row)
                            
                            if reorder_opps:
                                for prod_row in reorder_opps:
                                    prod_type = prod_row['Product_Type']
                                    q1_value = int(prod_row['Q1_Value'])
                                    historical_total = prod_row['Total_Revenue']
                                    expected_orders = prod_row['Expected_Orders_Q1']
                                    conf_pct = prod_row['Confidence_Pct']
                                    conf_tier = prod_row['Confidence_Tier']
                                    selection_key = f"{customer_name}|{prod_type}"
                                    
                                    # Initialize if needed
                                    if selection_key not in st.session_state[reorder_selections_key]:
                                        st.session_state[reorder_selections_key][selection_key] = {
                                            'selected': False,
                                            'value': q1_value,
                                            'confidence': confidence,
                                            'product_type': prod_type,
                                            'customer': customer_name,
                                            'top_skus': prod_row.get('Top_SKUs', ''),
                                            'historical_total': historical_total,
                                            'expected_orders': expected_orders,
                                            'confidence_tier': conf_tier
                                        }
                                    
                                    # Checkbox with explanation
                                    chk_col, val_col = st.columns([2, 1])
                                    with chk_col:
                                        is_selected = st.checkbox(
                                            f"{prod_type}",
                                            value=st.session_state[reorder_selections_key][selection_key]['selected'],
                                            key=f"chk_{selection_key}_{rep_name}",
                                            help=f"2025 Total: ${historical_total:,.0f} | Expected {expected_orders:.1f} orders in Q1 | {conf_tier} ({int(conf_pct*100)}% confidence)"
                                        )
                                        st.session_state[reorder_selections_key][selection_key]['selected'] = is_selected
                                    
                                    with val_col:
                                        if is_selected:
                                            new_val = st.number_input(
                                                "$",
                                                value=st.session_state[reorder_selections_key][selection_key]['value'],
                                                min_value=0,
                                                step=500,
                                                key=f"val_{selection_key}_{rep_name}",
                                                label_visibility="collapsed",
                                                help=f"Q1 Projection (editable)"
                                            )
                                            st.session_state[reorder_selections_key][selection_key]['value'] = new_val
                                        else:
                                            st.caption(f"${q1_value:,.0f}")
                            else:
                                st.caption("All covered ✅")
                        
                        # Product breakdown table
                        st.markdown("---")
                        st.markdown("**📊 2025 Product Breakdown & Q1 Projections**")
                        st.caption("Q1 Proj = (Avg Order × Expected Q1 Orders) — Weighted = Q1 Proj × Confidence %")
                        
                        if not cust_products.empty:
                            breakdown_data = []
                            for _, prod_row in cust_products.iterrows():
                                prod_type = prod_row['Product_Type']
                                prod_lower = str(prod_type).lower().strip()
                                
                                # Check NS/HS coverage
                                in_ns = any(prod_lower in str(o['Type']).lower() or str(o['Type']).lower() in prod_lower for o in cust_ns_orders)
                                in_hs = any(prod_lower in str(d['Type']).lower() or str(d['Type']).lower() in prod_lower for d in cust_hs_deals)
                                
                                # Due status
                                cadence = prod_row['Cadence_Days']
                                days_prod = prod_row['Days_Since_Last']
                                if in_ns or in_hs:
                                    due_status = "—"
                                elif pd.notna(cadence) and cadence > 0:
                                    if days_prod > cadence * 1.5:
                                        due_status = f"🔴 {int(days_prod - cadence)}d over"
                                    elif days_prod > cadence:
                                        due_status = "🟡 Due now"
                                    else:
                                        due_status = "🟢 On track"
                                else:
                                    due_status = "⚪ TBD"
                                
                                # Confidence tier indicator
                                conf_tier = prod_row['Confidence_Tier']
                                conf_emoji = "🟢" if conf_tier == 'Likely' else ("🟡" if conf_tier == 'Possible' else "🔴")
                                
                                breakdown_data.append({
                                    'Product': prod_type,
                                    '2025 Total': f"${prod_row['Total_Revenue']:,.0f}",
                                    'Orders': int(prod_row['Order_Count']),
                                    'Exp Q1': f"{prod_row['Expected_Orders_Q1']:.1f}",
                                    'Q1 Proj': f"${prod_row['Q1_Value']:,.0f}",
                                    'Conf': f"{conf_emoji} {int(prod_row['Confidence_Pct']*100)}%",
                                    'NS': "✅" if in_ns else "—",
                                    'HS': "✅" if in_hs else "—",
                                    'Status': due_status
                                })
                            
                            st.dataframe(
                                pd.DataFrame(breakdown_data),
                                use_container_width=True,
                                hide_index=True,
                                height=min(200, 35 + len(breakdown_data) * 35)
                            )
                        
                        # Manual entry for this customer
                        st.markdown("---")
                        st.markdown("**➕ Add Manual Entry**")
                        
                        man_col1, man_col2, man_col3 = st.columns([2, 1, 1])
                        
                        product_types_list = sorted(historical_df['Order Type'].dropna().unique().tolist())
                        
                        with man_col1:
                            manual_prod = st.selectbox(
                                "Product",
                                ["Select..."] + product_types_list,
                                key=f"man_prod_{customer_name}_{rep_name}",
                                label_visibility="collapsed"
                            )
                        
                        with man_col2:
                            manual_amt = st.number_input(
                                "Amount",
                                min_value=0,
                                step=1000,
                                key=f"man_amt_{customer_name}_{rep_name}",
                                label_visibility="collapsed"
                            )
                        
                        with man_col3:
                            if st.button("➕ Add", key=f"man_add_{customer_name}_{rep_name}"):
                                if manual_prod != "Select..." and manual_amt > 0:
                                    entry_key = f"{customer_name}|{manual_prod}|manual|{len(st.session_state[manual_entries_key])}"
                                    st.session_state[manual_entries_key][entry_key] = {
                                        'customer': customer_name,
                                        'product_type': manual_prod,
                                        'amount': manual_amt,
                                        'notes': 'Manual entry',
                                        'confidence': 0.90
                                    }
                                    st.rerun()
                        
                        # Show manual entries for this customer
                        cust_manual = {k: v for k, v in st.session_state[manual_entries_key].items() 
                                      if v['customer'] == customer_name}
                        if cust_manual:
                            for entry_key, entry in cust_manual.items():
                                ent_col1, ent_col2 = st.columns([4, 1])
                                with ent_col1:
                                    st.caption(f"📝 {entry['product_type']}: ${entry['amount']:,.0f}")
                                with ent_col2:
                                    if st.button("🗑️", key=f"del_{entry_key}"):
                                        del st.session_state[manual_entries_key][entry_key]
                                        st.rerun()
                
                # === REORDER SUMMARY ===
                st.markdown("---")
                st.markdown("### 📋 Reorder Forecast Summary")
                
                # Add explanation
                st.markdown("""
                <div style="background: rgba(59, 130, 246, 0.1); padding: 12px 15px; border-radius: 8px; margin-bottom: 15px; border-left: 3px solid #3b82f6;">
                    <strong style="color: #60a5fa;">How This Works:</strong><br/>
                    <span style="color: #94a3b8; font-size: 0.9rem;">
                    • <strong>Q1 Projection</strong> = Avg order value × Expected orders in Q1 (based on cadence)<br/>
                    • <strong>Weighted Forecast</strong> = Q1 Projection × Confidence % (25-75% based on order history)<br/>
                    • Confidence: 🟢 Likely (3+ orders) = 75% | 🟡 Possible (2 orders) = 50% | 🔴 Long Shot (1 order) = 25%
                    </span>
                </div>
                """, unsafe_allow_html=True)
                
                # Calculate totals
                total_reorder_raw = 0
                total_reorder_weighted = 0
                selected_items = []
                
                for key, data in st.session_state.get(reorder_selections_key, {}).items():
                    if data.get('selected', False):
                        val = data.get('value', 0)
                        conf = data.get('confidence', 0.5)
                        total_reorder_raw += val
                        total_reorder_weighted += val * conf
                        selected_items.append({
                            'Customer': data.get('customer', ''),
                            'Product_Type': data.get('product_type', ''),
                            'Q1_Projection': val,
                            'Confidence': f"{int(conf*100)}%",
                            'Weighted_Value': val * conf,
                            'Top_SKUs': data.get('top_skus', '')
                        })
                
                # Add manual entries
                total_manual = 0
                for key, entry in st.session_state.get(manual_entries_key, {}).items():
                    amt = entry.get('amount', 0)
                    conf = entry.get('confidence', 0.90)
                    total_manual += amt * conf
                    selected_items.append({
                        'Customer': entry.get('customer', ''),
                        'Product_Type': f"{entry.get('product_type', '')} (Manual)",
                        'Q1_Projection': amt,
                        'Confidence': f"{int(conf*100)}%",
                        'Weighted_Value': amt * conf,
                        'Top_SKUs': entry.get('notes', '')
                    })
                
                total_reorder_forecast = total_reorder_weighted + total_manual
                
                # Summary metrics with better labels
                sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
                with sum_col1:
                    st.metric("Selections", len([i for i in selected_items if '(Manual)' not in i['Product_Type']]))
                with sum_col2:
                    st.metric("Manual Adds", len([i for i in selected_items if '(Manual)' in i['Product_Type']]))
                with sum_col3:
                    st.metric("Q1 Projection (Raw)", f"${total_reorder_raw + sum(e.get('amount', 0) for e in st.session_state.get(manual_entries_key, {}).values()):,.0f}",
                             help="Sum of all Q1 projections before confidence weighting")
                with sum_col4:
                    st.metric("Weighted Forecast", f"${total_reorder_forecast:,.0f}",
                             help="Q1 Projection × Confidence % — available if you enable weighting in Step 6")
                
                # Show selected items table
                if selected_items:
                    with st.expander("📝 View All Selections", expanded=True):
                        sel_df = pd.DataFrame(selected_items)
                        # Keep numeric versions for calculations
                        sel_df_display = sel_df.copy()
                        sel_df_display['Weighted_Value'] = sel_df_display['Weighted_Value'].apply(lambda x: f"${x:,.0f}")
                        sel_df_display['Q1_Projection'] = sel_df_display['Q1_Projection'].apply(lambda x: f"${x:,.0f}")
                        st.dataframe(sel_df_display[['Customer', 'Product_Type', 'Q1_Projection', 'Confidence', 'Weighted_Value']], 
                                    use_container_width=True, hide_index=True,
                                    height=min(400, 35 + len(sel_df_display) * 35))
                    
                    # Store with numeric values for calculations
                    reorder_buckets['reorder_selections'] = sel_df
    
    # ═══════════════════════════════════════════════════════════════════════════
    # STEP 6: APPLY REORDER PROBABILITY WEIGHTING
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    step6_title = "Apply Reorder Probability" if is_team_view else f"{first_name}, Apply Reorder Probability?"
    st.markdown(f"### 🎲 Step 6: {step6_title}")
    
    st.markdown(f"""
    <div style="background: rgba(251, 191, 36, 0.1); border-left: 4px solid #fbbf24; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
        <div style="font-size: 1rem;">
            Your reorder selections can be used at <strong>face value (Raw)</strong> or <strong>weighted by confidence</strong>.
        </div>
        <div style="color: #94a3b8; margin-top: 8px; font-size: 0.9rem;">
            <strong>Confidence levels:</strong> 🟢 Likely (3+ orders) = 75% | 🟡 Possible (2 orders) = 50% | 🔴 Long Shot (1 order) = 25%
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Initialize session state for reorder weighting
    reorder_weight_key = f"reorder_weighting_{rep_name}"
    if reorder_weight_key not in st.session_state:
        st.session_state[reorder_weight_key] = {'enabled': False}
    
    rw_col1, rw_col2 = st.columns([2, 3])
    
    with rw_col1:
        apply_reorder_weight = st.toggle(
            "Apply confidence weighting to reorder",
            value=st.session_state[reorder_weight_key]['enabled'],
            key=f"toggle_reorder_weight_{rep_name}",
            help="When enabled, reorder values are multiplied by their confidence percentage"
        )
        st.session_state[reorder_weight_key]['enabled'] = apply_reorder_weight
    
    # Calculate totals for display
    reorder_raw_total = 0
    reorder_weighted_total = 0
    if 'reorder_selections' in reorder_buckets and not reorder_buckets['reorder_selections'].empty:
        df = reorder_buckets['reorder_selections']
        if 'Q1_Projection' in df.columns:
            reorder_raw_total = df['Q1_Projection'].sum()
        if 'Weighted_Value' in df.columns:
            reorder_weighted_total = df['Weighted_Value'].sum()
    
    with rw_col2:
        if apply_reorder_weight:
            st.markdown(f"""
            <div style="background: rgba(251, 191, 36, 0.15); padding: 12px 15px; border-radius: 8px;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <span style="color: #94a3b8;">Raw:</span> 
                        <strong style="color: #fbbf24;">${reorder_raw_total:,.0f}</strong>
                    </div>
                    <div style="font-size: 1.2rem;">→</div>
                    <div>
                        <span style="color: #94a3b8;">Weighted:</span> 
                        <strong style="color: #a78bfa;">${reorder_weighted_total:,.0f}</strong>
                    </div>
                    <div>
                        <span style="color: #94a3b8;">Using:</span> 
                        <strong style="color: #10b981;">Weighted ✓</strong>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background: rgba(16, 185, 129, 0.15); padding: 12px 15px; border-radius: 8px;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <span style="color: #94a3b8;">Raw:</span> 
                        <strong style="color: #fbbf24;">${reorder_raw_total:,.0f}</strong>
                    </div>
                    <div style="font-size: 1.2rem;">→</div>
                    <div>
                        <span style="color: #94a3b8;">Weighted:</span> 
                        <span style="color: #64748b;">${reorder_weighted_total:,.0f}</span>
                    </div>
                    <div>
                        <span style="color: #94a3b8;">Using:</span> 
                        <strong style="color: #10b981;">Raw ✓</strong>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    
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
        if 'Weighted_Value' in df.columns:
            return df['Weighted_Value'].sum()
        if 'Projected_Value' in df.columns:
            return df['Projected_Value'].sum()
        return 0
    
    # NetSuite scheduled (always at face value)
    selected_scheduled = sum(safe_sum(df) for k, df in export_buckets.items() if k in ns_categories)
    
    # HubSpot Pipeline - apply weighting if enabled
    pipeline_weights = st.session_state.get(pipeline_weight_key, {'enabled': False})
    selected_pipeline_raw = sum(safe_sum(df) for k, df in export_buckets.items() if k in hs_categories)
    
    if pipeline_weights.get('enabled', False):
        selected_pipeline = 0
        for key, df in export_buckets.items():
            if key in hs_categories and not df.empty and 'Amount_Numeric' in df.columns:
                raw_val = df['Amount_Numeric'].sum()
                if 'Expect' in key:
                    selected_pipeline += raw_val * (pipeline_weights.get('expect', 100) / 100)
                elif 'Commit' in key:
                    selected_pipeline += raw_val * (pipeline_weights.get('commit', 85) / 100)
                elif 'BestCase' in key:
                    selected_pipeline += raw_val * (pipeline_weights.get('best_case', 50) / 100)
                elif 'Opp' in key:
                    selected_pipeline += raw_val * (pipeline_weights.get('opportunity', 25) / 100)
                else:
                    selected_pipeline += raw_val  # Default to 100% if category unknown
    else:
        selected_pipeline = selected_pipeline_raw
    
    # Calculate reorder forecast total - both RAW and WEIGHTED
    selected_reorder_weighted = 0
    selected_reorder_raw = 0
    if reorder_buckets:
        for df in reorder_buckets.values():
            if not df.empty:
                if 'Weighted_Value' in df.columns:
                    selected_reorder_weighted += df['Weighted_Value'].sum()
                elif 'Projected_Value' in df.columns:
                    selected_reorder_weighted += df['Projected_Value'].sum()
                if 'Q1_Projection' in df.columns:
                    selected_reorder_raw += df['Q1_Projection'].sum()
    
    # Apply reorder weighting based on Step 6 choice
    reorder_weights = st.session_state.get(reorder_weight_key, {'enabled': False})
    if reorder_weights.get('enabled', False):
        selected_reorder = selected_reorder_weighted
    else:
        selected_reorder = selected_reorder_raw
    
    total_forecast = selected_scheduled + selected_pipeline + selected_reorder
    gap_to_goal = q1_goal - total_forecast


    # === SIDEBAR: LIVE SCOREBOARD (UI ONLY) ===
    with st.sidebar:
        st.markdown("### 🧭 Live Scoreboard")
        st.caption("Updates instantly as you include / exclude orders, deals, and reorder rows.")

        # Progress
        progress_pct = (total_forecast / q1_goal * 100) if q1_goal > 0 else 0
        progress_val = min(max(total_forecast / q1_goal, 0), 1) if q1_goal > 0 else 0
        st.progress(progress_val)
        st.caption(f"**{progress_pct:.0f}%** of goal • Forecast: **${total_forecast:,.0f}**")

        cA, cB = st.columns(2)
        with cA:
            st.metric("📦 Scheduled", f"${selected_scheduled:,.0f}")
        with cB:
            pipeline_wt_status = "weighted" if pipeline_weights.get('enabled', False) else "raw"
            st.metric("🎯 Pipeline", f"${selected_pipeline:,.0f}",
                     help=f"Using {pipeline_wt_status} values. Raw: ${selected_pipeline_raw:,.0f}")

        cC, cD = st.columns(2)
        with cC:
            reorder_wt_status = "weighted" if reorder_weights.get('enabled', False) else "raw"
            st.metric("🔄 Reorder", f"${selected_reorder:,.0f}", 
                     help=f"Using {reorder_wt_status} values. Raw: ${selected_reorder_raw:,.0f} | Weighted: ${selected_reorder_weighted:,.0f}")
        with cD:
            if gap_to_goal > 0:
                st.metric("Gap", f"${gap_to_goal:,.0f}")
            else:
                st.metric("Ahead", f"${abs(gap_to_goal):,.0f}")

        st.markdown('<div class="soft-divider"></div>', unsafe_allow_html=True)

    # === STICKY SUMMARY BAR (HUD STYLE) ===
    gap_class = "val-gap-behind" if gap_to_goal > 0 else "val-gap-ahead"
    gap_label = "GAP" if gap_to_goal > 0 else "AHEAD"
    gap_display = f"${abs(gap_to_goal):,.0f}"
    
    pipeline_tip = f"Weighted ({pipeline_weights.get('enabled', False)}). Raw: ${selected_pipeline_raw:,.0f}"
    reorder_tip = f"Raw: ${selected_reorder_raw:,.0f} | Weighted: ${selected_reorder_weighted:,.0f}"
    
    st.markdown(f"""
    <div class="sticky-forecast-bar-q1">
        <div class="sticky-item">
            <div class="sticky-label">Scheduled</div>
            <div class="sticky-val val-sched">${selected_scheduled:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item" title="{pipeline_tip}">
            <div class="sticky-label">Pipeline{'*' if pipeline_weights.get('enabled', False) else ''}</div>
            <div class="sticky-val val-pipe">${selected_pipeline:,.0f}</div>
        </div>
        <div class="sticky-sep"></div>
        <div class="sticky-item" title="{reorder_tip}">
            <div class="sticky-label">Reorder{'*' if reorder_weights.get('enabled', False) else ''}</div>
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
    # STEP 7: YOUR FORECAST SUMMARY
    # ═══════════════════════════════════════════════════════════════════════════
    
    st.markdown("---")
    step7_title = "Team Forecast Summary" if is_team_view else f"{first_name}, Here's Your Q1 Forecast!"
    st.markdown(f"### 🎉 Step 7: {step7_title}")
    
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
        pipeline_help = f"{'Weighted' if pipeline_weights.get('enabled', False) else 'Raw'} values. Raw: ${selected_pipeline_raw:,.0f}"
        st.metric("🎯 Pipeline", f"${selected_pipeline:,.0f}", help=pipeline_help)
    with m3:
        reorder_help = f"{'Weighted' if reorder_weights.get('enabled', False) else 'Raw'} values. Raw: ${selected_reorder_raw:,.0f} | Weighted: ${selected_reorder_weighted:,.0f}"
        st.metric("🔄 Reorder", f"${selected_reorder:,.0f}", help=reorder_help)
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
        pipeline_status = "weighted" if pipeline_weights.get('enabled', False) else "raw"
        reorder_status = "weighted" if reorder_weights.get('enabled', False) else "raw"
        
        st.markdown("#### 📊 The Breakdown")
        st.markdown(f"""
        **📦 Confirmed Orders:** ${selected_scheduled:,.0f}
        - Already in NetSuite, shipping Q1
        
        **🎯 Pipeline Deals:** ${selected_pipeline:,.0f}
        - Using {pipeline_status} values{f' (raw: ${selected_pipeline_raw:,.0f})' if pipeline_weights.get('enabled', False) else ''}
        
        **🔄 Reorder Potential:** ${selected_reorder:,.0f}
        - Using {reorder_status} values{f' (raw: ${selected_reorder_raw:,.0f})' if reorder_weights.get('enabled', False) else f' (weighted: ${selected_reorder_weighted:,.0f})'}
        
        **🎯 Q1 Goal:** ${q1_goal:,.0f}
        """)
    st.markdown('</div>', unsafe_allow_html=True)


    # === EXECUTIVE VISUALS (UI ONLY) ===
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    vcol1, vcol2 = st.columns([1, 1])
    with vcol1:
        st.markdown("#### 🧩 Forecast Mix")
        st.plotly_chart(
            create_forecast_composition_donut(selected_scheduled, selected_pipeline, selected_reorder, "Forecast Mix"),
            use_container_width=True
        )
        st.caption("Mix of confirmed orders, pipeline, and reorder potential (weighted).")

    with vcol2:
        st.markdown("#### 🗺️ Path to Goal")
        st.plotly_chart(
            create_forecast_waterfall(selected_scheduled, selected_pipeline, selected_reorder, q1_goal, "Path to Goal"),
            use_container_width=True
        )
        st.caption("How the forecast builds up compared to your goal.")

    st.markdown('</div>', unsafe_allow_html=True)

    # === EXPORT SECTION (Unified - matches Sales Dashboard methodology) ===
    st.markdown("---")
    st.markdown('<div class="section-header">📤 Export Q1 2026 Forecast</div>', unsafe_allow_html=True)
    
    if total_forecast > 0:
        # Initialize Lists
        export_summary = []
        export_data = []
        
        # Get weighting status
        pipeline_weighted = pipeline_weights.get('enabled', False)
        reorder_weighted = reorder_weights.get('enabled', False)
        
        # A. Build Summary
        export_summary.append({'Category': '=== Q1 2026 FORECAST SUMMARY ===', 'Amount': ''})
        export_summary.append({'Category': 'Q1 Goal', 'Amount': f"${q1_goal:,.0f}"})
        export_summary.append({'Category': 'Scheduled Orders (NetSuite)', 'Amount': f"${selected_scheduled:,.0f}"})
        
        # Pipeline with weighting info
        if pipeline_weighted:
            export_summary.append({'Category': 'Pipeline Deals (HubSpot) - WEIGHTED', 'Amount': f"${selected_pipeline:,.0f}"})
            export_summary.append({'Category': '  → Raw Pipeline Value', 'Amount': f"${selected_pipeline_raw:,.0f}"})
            export_summary.append({'Category': f"  → Weights: Expect={pipeline_weights.get('expect',100)}% Commit={pipeline_weights.get('commit',85)}% BestCase={pipeline_weights.get('best_case',50)}% Opp={pipeline_weights.get('opportunity',25)}%", 'Amount': ''})
        else:
            export_summary.append({'Category': 'Pipeline Deals (HubSpot) - RAW', 'Amount': f"${selected_pipeline:,.0f}"})
        
        # Reorder with weighting info
        if reorder_weighted:
            export_summary.append({'Category': 'Reorder Potential - WEIGHTED', 'Amount': f"${selected_reorder:,.0f}"})
            export_summary.append({'Category': '  → Raw Reorder Value', 'Amount': f"${selected_reorder_raw:,.0f}"})
            export_summary.append({'Category': '  → Weights: Likely=75% Possible=50% LongShot=25%', 'Amount': ''})
        else:
            export_summary.append({'Category': 'Reorder Potential - RAW', 'Amount': f"${selected_reorder:,.0f}"})
            export_summary.append({'Category': '  → Weighted Value (not applied)', 'Amount': f"${selected_reorder_weighted:,.0f}"})
        
        export_summary.append({'Category': 'Total Forecast', 'Amount': f"${total_forecast:,.0f}"})
        export_summary.append({'Category': 'Gap to Goal', 'Amount': f"${gap_to_goal:,.0f}"})
        export_summary.append({'Category': '', 'Amount': ''})
        export_summary.append({'Category': '=== WEIGHTING SETTINGS ===', 'Amount': ''})
        export_summary.append({'Category': 'Pipeline Probability Applied', 'Amount': 'Yes' if pipeline_weighted else 'No'})
        export_summary.append({'Category': 'Reorder Confidence Applied', 'Amount': 'Yes' if reorder_weighted else 'No'})
        export_summary.append({'Category': '', 'Amount': ''})
        export_summary.append({'Category': '=== SELECTED COMPONENTS ===', 'Amount': ''})
        
        # Helper to strip emojis from labels for clean CSV export
        def clean_label(text):
            """Remove emojis and clean up label for CSV export"""
            import re
            # Remove emoji characters - comprehensive pattern
            emoji_pattern = re.compile("["
                u"\U0001F600-\U0001F64F"  # emoticons
                u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                u"\U0001F680-\U0001F6FF"  # transport & map symbols
                u"\U0001F1E0-\U0001F1FF"  # flags
                u"\U00002702-\U000027B0"
                u"\U000024C2-\U0001F251"
                u"\U0001f926-\U0001f937"
                u"\U00010000-\U0010ffff"
                u"\u2640-\u2642"
                u"\u2600-\u2B55"
                u"\u200d"
                u"\u23cf"
                u"\u23e9-\u23f9"  # includes ⏳ (U+23F3)
                u"\u231a-\u231b"
                u"\ufe0f"
                u"\u3030"
                "]+", flags=re.UNICODE)
            cleaned = emoji_pattern.sub('', str(text))
            return cleaned.strip()
        
        for key, df in export_buckets.items():
            if df.empty:
                continue
            # Handle both Amount and Amount_Numeric columns (NS uses Amount, HS uses Amount_Numeric)
            if 'Amount_Numeric' in df.columns:
                cat_val_raw = df['Amount_Numeric'].sum()
            elif 'Amount' in df.columns:
                cat_val_raw = df['Amount'].sum()
            else:
                cat_val_raw = 0
            
            # Apply pipeline weighting for display if enabled
            if key in hs_categories and pipeline_weighted:
                if 'Expect' in key:
                    cat_val = cat_val_raw * (pipeline_weights.get('expect', 100) / 100)
                elif 'Commit' in key:
                    cat_val = cat_val_raw * (pipeline_weights.get('commit', 85) / 100)
                elif 'BestCase' in key:
                    cat_val = cat_val_raw * (pipeline_weights.get('best_case', 50) / 100)
                elif 'Opp' in key:
                    cat_val = cat_val_raw * (pipeline_weights.get('opportunity', 25) / 100)
                else:
                    cat_val = cat_val_raw
            else:
                cat_val = cat_val_raw
                
            if cat_val_raw > 0:
                label = clean_label(ns_categories.get(key, {}).get('label', hs_categories.get(key, {}).get('label', key)))
                count = len(df)
                if key in hs_categories and pipeline_weighted and cat_val != cat_val_raw:
                    export_summary.append({'Category': f"{label} ({count} items)", 'Amount': f"${cat_val:,.0f} (raw: ${cat_val_raw:,.0f})"})
                else:
                    export_summary.append({'Category': f"{label} ({count} items)", 'Amount': f"${cat_val:,.0f}"})
        
        # Add reorder bucket totals to summary
        if reorder_buckets:
            for key, df in reorder_buckets.items():
                if df.empty:
                    continue
                # Use correct column names
                raw_val = df['Q1_Projection'].sum() if 'Q1_Projection' in df.columns else 0
                weighted_val = df['Weighted_Value'].sum() if 'Weighted_Value' in df.columns else 0
                
                if raw_val > 0 or weighted_val > 0:
                    tier_label = key.replace('reorder_', 'Reorder - ').replace('_', ' ').title()
                    count = len(df)
                    if reorder_weighted:
                        export_summary.append({'Category': f"{tier_label} ({count} items)", 'Amount': f"${weighted_val:,.0f} (raw: ${raw_val:,.0f})"})
                    else:
                        export_summary.append({'Category': f"{tier_label} ({count} items)", 'Amount': f"${raw_val:,.0f} (weighted: ${weighted_val:,.0f})"})
        
        export_summary.append({'Category': '', 'Amount': ''})
        export_summary.append({'Category': '=== DETAILED LINE ITEMS ===', 'Amount': ''})
        
        # B. Build Line Items
        
        # 1. NetSuite & HubSpot Items from export_buckets
        for key, df in export_buckets.items():
            if df.empty:
                continue
            
            label = clean_label(ns_categories.get(key, {}).get('label', hs_categories.get(key, {}).get('label', key)))
            
            # Get weight multiplier for this category
            weight_mult = 1.0
            if key in hs_categories and pipeline_weighted:
                if 'Expect' in key:
                    weight_mult = pipeline_weights.get('expect', 100) / 100
                elif 'Commit' in key:
                    weight_mult = pipeline_weights.get('commit', 85) / 100
                elif 'BestCase' in key:
                    weight_mult = pipeline_weights.get('best_case', 50) / 100
                elif 'Opp' in key:
                    weight_mult = pipeline_weights.get('opportunity', 25) / 100
            
            for _, row in df.iterrows():
                # Determine fields based on source type (NS vs HS)
                if key in ns_categories:  # NetSuite
                    item_type = f"Sales Order - {label}"
                    item_id = row.get('SO #', row.get('Document Number', ''))
                    cust = row.get('Customer', '')
                    date_val = row.get('Ship Date', row.get('Key Date', ''))
                    deal_type = row.get('Type', row.get('Display_Type', ''))
                    amount_raw = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                    amount = amount_raw  # NS always at face value
                    # Get Sales Rep
                    rep = row.get('Sales Rep', row.get('Rep Master', ''))
                else:  # HubSpot
                    item_type = f"HubSpot - {label}"
                    item_id = row.get('Deal ID', row.get('Record ID', ''))
                    cust = row.get('Account Name', row.get('Deal Name', ''))
                    date_val = row.get('Close', row.get('Close Date', ''))
                    deal_type = row.get('Type', row.get('Display_Type', ''))
                    amount_raw = pd.to_numeric(row.get('Amount_Numeric', 0), errors='coerce')
                    amount = amount_raw * weight_mult  # Apply weighting
                    # Get Deal Owner
                    rep = row.get('Deal Owner', '')
                    if pd.isna(rep) or rep is None or str(rep).strip() == '':
                        first = row.get('Deal Owner First Name', '')
                        last = row.get('Deal Owner Last Name', '')
                        if first or last:
                            rep = f"{first} {last}".strip()
                
                # Clean up date value
                if pd.isna(date_val) or date_val == '' or date_val == '—':
                    date_val = ''
                elif isinstance(date_val, pd.Timestamp):
                    date_val = date_val.strftime('%Y-%m-%d')
                elif isinstance(date_val, str):
                    if date_val and date_val != '—':
                        try:
                            parsed_date = pd.to_datetime(date_val, errors='coerce')
                            if pd.notna(parsed_date):
                                date_val = parsed_date.strftime('%Y-%m-%d')
                            else:
                                date_val = ''
                        except:
                            date_val = ''
                    else:
                        date_val = ''
                else:
                    date_val = ''
                
                # Ensure rep is a string, not NaN
                if pd.isna(rep) or rep is None:
                    rep = ''
                else:
                    rep = str(rep).strip()
                    if rep.lower() in ['nan', 'none']:
                        rep = ''
                
                export_row = {
                    'Category': item_type,
                    'ID': item_id,
                    'Customer': cust,
                    'Order/Deal Type': deal_type,
                    'Top SKUs': '',  # Not applicable for NS/HS items
                    'Date': date_val,
                    'Amount': amount,
                    'Rep': rep
                }
                
                # Add raw amount column for weighted items
                if key in hs_categories and pipeline_weighted and weight_mult != 1.0:
                    export_row['Raw Amount'] = amount_raw
                    export_row['Weight'] = f"{weight_mult:.0%}"
                
                export_data.append(export_row)
        
        # 2. Reorder Prospects from reorder_buckets
        if reorder_buckets:
            for key, df in reorder_buckets.items():
                if df.empty:
                    continue
                
                tier_label = key.replace('reorder_', 'Reorder - ').replace('_', ' ').title()
                
                for _, row in df.iterrows():
                    cust = row.get('Customer', '')
                    product_type = row.get('Product_Type', '')
                    top_skus = row.get('Top_SKUs', '')
                    conf_tier = row.get('Confidence_Tier', row.get('Confidence', ''))
                    
                    # Handle confidence - could be string "75%" or float 0.75
                    conf_val = row.get('Confidence', 0)
                    if isinstance(conf_val, str) and '%' in conf_val:
                        conf_pct = float(conf_val.replace('%', '')) / 100
                    else:
                        conf_pct = float(conf_val) if conf_val else 0
                    
                    # Use correct column names
                    q1_val = row.get('Q1_Projection', row.get('Q1_Value', 0))
                    weighted_val = row.get('Weighted_Value', row.get('Projected_Value', q1_val * conf_pct))
                    
                    # Use raw or weighted based on setting
                    if reorder_weighted:
                        amount = weighted_val
                    else:
                        amount = q1_val
                    
                    export_row = {
                        'Category': tier_label,
                        'ID': '',  # No ID for reorder prospects
                        'Customer': cust,
                        'Order/Deal Type': f"{product_type} - {conf_tier}",
                        'Top SKUs': top_skus,
                        'Date': '',  # No specific date
                        'Amount': amount,
                        'Rep': rep_name,  # Use selected rep name
                        'Confidence': f"{conf_pct:.0%}" if conf_pct else '',
                        'Raw Value': q1_val,
                        'Weighted Value': weighted_val
                    }
                    
                    export_data.append(export_row)
        
        # C. Construct CSV
        if export_data:
            summary_df = pd.DataFrame(export_summary)
            data_df = pd.DataFrame(export_data)
            
            # Format Amount in Data DF
            data_df['Amount'] = data_df['Amount'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != '' else "$0.00")
            if 'Raw Amount' in data_df.columns:
                data_df['Raw Amount'] = data_df['Raw Amount'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) else "")
            if 'Raw Value' in data_df.columns:
                data_df['Raw Value'] = data_df['Raw Value'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != 0 else "")
            if 'Weighted Value' in data_df.columns:
                data_df['Weighted Value'] = data_df['Weighted Value'].apply(lambda x: f"${x:,.2f}" if pd.notna(x) and x != 0 else "")
            
            final_csv = summary_df.to_csv(index=False) + "\n" + data_df.to_csv(index=False)
            
            st.download_button(
                label="📥 Download Q1 2026 Forecast",
                data=final_csv,
                file_name=f"q1_2026_forecast_{rep_name.replace(' ', '_')}_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
            
            # Show what's in the export
            weight_note = ""
            if pipeline_weighted or reorder_weighted:
                weight_note = " | Weighting: "
                if pipeline_weighted:
                    weight_note += "Pipeline ✓"
                if reorder_weighted:
                    weight_note += " Reorder ✓"
            
            st.caption(f"Export includes summary + {len(data_df)} line items.{weight_note}")
        else:
            st.info("No items selected for export")
    else:
        st.info("Select items above to enable export")
    
    # === DEBUG INFO ===
    with st.expander("🔧 Debug: Data Summary"):
        st.write("**Data Source:** Copy of All Reps All Pipelines (Q4 2025 + Q1 2026 deals)")
        if is_team_view:
            st.write(f"**Team Reps:** {', '.join(active_team_reps)}")
        st.write(f"**Total Deals Loaded:** {len(deals_df)}")
        
        st.write("**--- NetSuite Buckets ---**")
        st.write(f"**PF Q1 Date:** {len(combined_pf)} orders, ${total_pf_amount:,.0f}")
        st.write(f"**PA Q1 PA Date:** {len(combined_pa)} orders, ${total_pa_amount:,.0f}")
        st.write(f"**PF No Date:** {len(combined_pf_nodate)} orders, ${total_pf_nodate_amount:,.0f}")
        st.write(f"**PA No Date:** {len(combined_pa_nodate)} orders, ${total_pa_nodate_amount:,.0f}")
        st.write(f"**PA >2 Weeks:** {len(combined_pa_old)} orders, ${total_pa_old_amount:,.0f}")
        
        st.write("**--- HubSpot Buckets ---**")
        for key in hs_categories.keys():
            df = hs_dfs.get(key, pd.DataFrame())
            val = df['Amount_Numeric'].sum() if not df.empty and 'Amount_Numeric' in df.columns else 0
            st.write(f"**{key}:** {len(df)} deals, ${val:,.0f}")


# Run if called directly
if __name__ == "__main__":
    main()
