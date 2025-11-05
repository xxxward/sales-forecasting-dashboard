"""
Sales Forecasting Dashboard — Refactored
Single‑file Streamlit app for Q4 2025 sales forecasting with drill‑downs,
Q1 spillover logic, and team/rep/reconciliation views.

Notes
-----
• Requires Streamlit secrets: st.secrets["gcp_service_account"] (full JSON)
• Google Sheets: set SPREADSHEET_ID below (or via env var)
• Caching: clear via sidebar button if you change code/sheets

This refactor keeps your original business logic while consolidating
~2k lines into a maintainable structure:
  - CONFIG / STYLES / UTILITIES
  - DATA ACCESS
  - PROCESSING (deals, sales orders, invoices)
  - METRICS (team + rep)
  - VISUALS (charts & UI fragments)
  - VIEWS (Team, Rep, Reconciliation)
  - MAIN
"""

# ==================== IMPORTS ====================
import os
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ==================== CONFIG ====================
st.set_page_config(
    page_title="Sales Forecasting Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

SPREADSHEET_ID = os.getenv(
    "CALYX_SALES_SPREADSHEET_ID",
    "12s-BanWrT_N8SuB3IXFp5JF-xPYB2I-YjmYAYaWsxJk",
)
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Cache ~1 hour; bump CACHE_VERSION when logic changes to invalidate
CACHE_TTL = 3600
CACHE_VERSION = "v31"

# Quarter focus (static for 2025; adjust if needed)
Q4_START = pd.Timestamp("2025-10-01")
Q4_END = pd.Timestamp("2025-12-31")

COLORS = {
    "primary": "#1E88E5",
    "success": "#43A047",
    "warning": "#FB8C00",
    "danger": "#DC3912",
    "expect": "#1E88E5",
    "commit": "#43A047",
    "best_case": "#FB8C00",
    "opportunity": "#8E24AA",
}

LEAD_TIME_MAP = {
    "Labeled - Labels In Stock": 10,
    "Outer Boxes": 20,
    "Non-Labeled - 1 Week Lead Time": 5,
    "Non-Labeled - 2 Week Lead Time": 10,
    "Labeled - Print & Apply": 20,
    "Non-Labeled - Custom Lead Time": 30,
    "Labeled with FEP - Print & Apply": 35,
    "Labeled - Custom Lead Time": 40,
    "Flexpack": 25,
    "Labels Only - Direct to Customer": 15,
    "Labels Only - For Inventory": 15,
    "Labeled with FEP - Labels In Stock": 25,
    "Labels Only (deprecated)": 15,
}

# ==================== STYLES ====================

def load_custom_css():
    st.markdown(
        """
        <style>
        .metric-card { background:#f0f2f6; padding:20px; border-radius:10px; }
        .section-header { background:#f0f2f6; padding:10px 15px; border-radius:8px; margin:15px 0; font-weight:600; }
        .info-card { background:#e3f2fd; padding:15px; border-radius:8px; border-left:4px solid #1E88E5; margin:10px 0; }
        .progress-breakdown { background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); color:#fff; padding:24px; border-radius:14px; margin:12px 0; }
        .progress-row { display:flex; justify-content:space-between; padding:8px 0; border-bottom:1px solid rgba(255,255,255,.2); }
        .progress-row:last-child { border-bottom:none; font-weight:700; padding-top:12px; border-top:2px solid rgba(255,255,255,.35); }
        </style>
        """,
        unsafe_allow_html=True,
    )

# ==================== UTILITIES ====================

def clean_numeric(v):
    if pd.isna(v) or str(v).strip() == "":
        return 0.0
    s = str(v).replace(",", "").replace("$", "").replace(" ", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def bdays_between(start_date, end_date):
    if pd.isna(start_date):
        return 0
    return max(0, pd.bdate_range(start=start_date, end=end_date).size - 1)


def business_days_before(end_date: pd.Timestamp, n: int) -> pd.Timestamp:
    cur = end_date
    counted = 0
    while counted < n:
        cur -= timedelta(days=1)
        if cur.weekday() < 5:
            counted += 1
    return cur


def dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df

# ==================== DATA ACCESS ====================

@st.cache_data(ttl=CACHE_TTL)
def load_google_sheets_data(sheet_name: str, rng: str, version: str = CACHE_VERSION) -> pd.DataFrame:
    """Load a sheet range into a DataFrame. Pads ragged rows; safe errors."""
    try:
        if "gcp_service_account" not in st.secrets:
            st.sidebar.error("Missing Google Cloud credentials (st.secrets['gcp_service_account']).")
            return pd.DataFrame()
        creds = service_account.Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES
        )
        svc = build("sheets", "v4", credentials=creds)
        res = (
            svc.spreadsheets()
            .values()
            .get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}!{rng}")
            .execute()
        )
        vals = res.get("values", [])
        if not vals:
            return pd.DataFrame()
        max_cols = max(len(r) for r in vals) if len(vals) > 1 else len(vals[0])
        for r in vals:
            while len(r) < max_cols:
                r.append("")
        return pd.DataFrame(vals[1:], columns=vals[0])
    except Exception as e:
        st.sidebar.error(f"Error loading {sheet_name}: {e}")
        return pd.DataFrame()

# ==================== PROCESSING ====================

EXCLUDED_STAGES = {
    "", "(Blanks)", None, "Cancelled", "checkout abandoned", "closed lost",
    "closed won", "sales order created in NS", "NCR", "Shipped",
}


def apply_q1_spillover_from_lead_times(df: pd.DataFrame) -> pd.DataFrame:
    """If no explicit 'Q1 2026 Spillover' flag, infer from Product Type lead times and Close Date."""
    if "Product Type" not in df.columns or "Close Date" not in df.columns:
        df["Ships_In_Q4"] = True
        df["Ships_In_Q1"] = False
        return df

    df["Ships_In_Q4"] = True
    df["Ships_In_Q1"] = False

    for ptype, days in LEAD_TIME_MAP.items():
        cutoff = business_days_before(Q4_END, days)
        mask = (df["Product Type"] == ptype) & (df["Close Date"] > cutoff) & df["Close Date"].notna()
        df.loc[mask, "Ships_In_Q4"] = False
        df.loc[mask, "Ships_In_Q1"] = True
    return df


def process_deals(deals: pd.DataFrame) -> pd.DataFrame:
    if deals.empty:
        return deals

    # Column normalization
    rename_map = {
        "Record ID": "Record ID",
        "Deal Name": "Deal Name",
        "Deal Stage": "Deal Stage",
        "Close Date": "Close Date",
        "Close Status": "Status",
        "Amount": "Amount",
        "Pipeline": "Pipeline",
        "Deal Type": "Product Type",
        "Q1 2026 Spillover": "Q1 2026 Spillover",
    }

    # Detect combined Deal Owner col or separate first/last
    owner_col = None
    for c in deals.columns:
        cl = c.lower()
        if "deal owner first name" in cl and "deal owner last name" in cl:
            owner_col = c
            break
    if owner_col:
        rename_map[owner_col] = "Deal Owner"

    # Apply fuzzy rename
    to_rename = {}
    for c in deals.columns:
        for k, v in rename_map.items():
            if (k == c) or (k.lower() in c.lower()):
                to_rename[c] = v
                break
    deals = deals.rename(columns=to_rename)

    # Owner synthesis if needed
    if "Deal Owner" not in deals.columns:
        first = None
        last = None
        for c in deals.columns:
            cl = c.lower()
            if "deal owner first name" in cl:
                first = c
            if "deal owner last name" in cl:
                last = c
        if first and last:
            deals["Deal Owner"] = (
                deals[first].fillna("") + " " + deals[last].fillna("")
            ).str.strip()

    # Types
    if "Amount" in deals.columns:
        deals["Amount"] = deals["Amount"].apply(clean_numeric)
    if "Close Date" in deals.columns:
        deals["Close Date"] = pd.to_datetime(deals["Close Date"], errors="coerce")

    # Q4 filter
    if "Close Date" in deals.columns:
        deals = deals[(deals["Close Date"] >= Q4_START) & (deals["Close Date"] <= Q4_END)]

    # Stage filter
    if "Deal Stage" in deals.columns:
        deals["Deal Stage"] = deals["Deal Stage"].fillna("").astype(str).str.strip()
        deals = deals[~deals["Deal Stage"].str.lower().isin({(s or "").lower() for s in EXCLUDED_STAGES})]

    # Explicit spillover flag or infer
    if "Q1 2026 Spillover" in deals.columns:
        deals["Ships_In_Q4"] = deals["Q1 2026 Spillover"].astype(str).str.strip() != "Q1 2026"
        deals["Ships_In_Q1"] = ~deals["Ships_In_Q4"]
    else:
        deals = apply_q1_spillover_from_lead_times(deals)

    # Normalize Status (Expect/Commit/Best Case/Opportunity only used for HubSpot forecast rollups)
    if "Status" in deals.columns:
        deals["Status"] = deals["Status"].astype(str).str.strip()

    return dedupe_columns(deals)


def process_sales_orders(so: pd.DataFrame) -> pd.DataFrame:
    if so.empty:
        return so

    # Map known positional columns (0‑indexed)
    pos_map = {8: "Order Start Date", 11: "Customer Promise Date", 12: "Projected Date", 27: "Pending Approval Date"}
    cols = so.columns.tolist()
    for idx, name in pos_map.items():
        if len(cols) > idx:
            so.rename(columns={cols[idx]: name}, inplace=True)

    # Standard names by content
    std = {"Status": "status", "Amount": "amount", "Sales Rep": "sales rep", "Customer": "customer", "Document Number": "document"}
    for c in so.columns.tolist():
        cl = c.lower()
        if "status" in cl and "Status" not in so.columns:
            so.rename(columns={c: "Status"}, inplace=True)
        elif any(k.lower() in cl for k in ["amount", "total"]) and "Amount" not in so.columns:
            so.rename(columns={c: "Amount"}, inplace=True)
        elif ("sales rep" in cl or "salesrep" in cl) and "Sales Rep" not in so.columns:
            so.rename(columns={c: "Sales Rep"}, inplace=True)
        elif ("customer" in cl) and ("Customer" not in so.columns) and ("customer promise" not in cl):
            so.rename(columns={c: "Customer"}, inplace=True)
        elif ("doc" in cl or "document" in cl) and ("Document Number" not in so.columns):
            so.rename(columns={c: "Document Number"}, inplace=True)

    so = dedupe_columns(so)

    # Clean
    if "Amount" in so.columns:
        so["Amount"] = so["Amount"].apply(clean_numeric)
    if "Sales Rep" in so.columns:
        so["Sales Rep"] = so["Sales Rep"].astype(str).str.strip()
    if "Status" in so.columns:
        so["Status"] = so["Status"].astype(str).str.strip()
        so = so[so["Status"].isin(["Pending Approval", "Pending Fulfillment", "Pending Billing/Partially Fulfilled"])]

    # Dates
    for c in ["Order Start Date", "Customer Promise Date", "Projected Date", "Pending Approval Date"]:
        if c in so.columns:
            so[c] = pd.to_datetime(so[c], errors="coerce")

    # Age for old PAs
    if "Order Start Date" in so.columns:
        today = pd.Timestamp.now()
        so["Age_Business_Days"] = so["Order Start Date"].apply(lambda d: bdays_between(d, today))

    # Valid rows
    if {"Amount", "Sales Rep"}.issubset(so.columns):
        so = so[(so["Amount"] > 0) & so["Sales Rep"].notna() & (so["Sales Rep"].astype(str).str.strip() != "")]

    return so


def process_dashboard(dash: pd.DataFrame) -> pd.DataFrame:
    if dash.empty:
        return dash
    # Expect 3 cols: Rep, Quota, NetSuite Orders
    dash = dash.copy()
    dash.columns = dash.columns[:3]
    dash = dash.rename(columns={dash.columns[0]: "Rep Name", dash.columns[1]: "Quota", dash.columns[2]: "NetSuite Orders"})
    dash = dash[dash["Rep Name"].notna() & (dash["Rep Name"].astype(str).str.strip() != "")]
    dash["Quota"] = dash["Quota"].apply(clean_numeric)
    dash["NetSuite Orders"] = dash["NetSuite Orders"].apply(clean_numeric)
    return dash


def process_invoices(inv: pd.DataFrame) -> pd.DataFrame:
    if inv.empty:
        return inv
    inv = inv.rename(
        columns={
            inv.columns[0]: "Invoice Number",
            inv.columns[1]: "Status",
            inv.columns[2]: "Date",
            inv.columns[6]: "Customer",
            inv.columns[10]: "Amount",
            inv.columns[14]: "Sales Rep",
        }
    )
    inv["Amount"] = inv["Amount"].apply(clean_numeric)
    inv["Date"] = pd.to_datetime(inv["Date"], errors="coerce")
    inv = inv[(inv["Date"] >= Q4_START) & (inv["Date"] <= Q4_END)]
    inv["Sales Rep"] = inv["Sales Rep"].astype(str).str.strip()
    inv = inv[(inv["Amount"] > 0) & inv["Sales Rep"].notna() & (inv["Sales Rep"] != "")]
    return inv

# ==================== METRICS ====================

def calc_rep_metrics(rep: str, deals: pd.DataFrame, dash: pd.DataFrame, so: pd.DataFrame | None = None):
    info = dash[dash["Rep Name"] == rep]
    if info.empty:
        return None
    quota = float(info["Quota"].iloc[0])
    orders = float(info["NetSuite Orders"].iloc[0])

    rdeals = deals[deals["Deal Owner"] == rep].copy()
    r_q4 = rdeals[rdeals.get("Ships_In_Q4", True) == True]
    r_q1 = rdeals[rdeals.get("Ships_In_Q1", False) == True]

    expect_commit = r_q4[r_q4["Status"].isin(["Expect", "Commit"])]["Amount"].sum()
    best_opp = r_q4[r_q4["Status"].isin(["Best Case", "Opportunity"])]["Amount"].sum()

    q1_ec = r_q1[r_q1["Status"].isin(["Expect", "Commit"])]["Amount"].sum()
    q1_bo = r_q1[r_q1["Status"].isin(["Best Case", "Opportunity"])]["Amount"].sum()

    pa = pa_no_date = pa_old = pf = pf_no_date = 0.0
    pa_details = pd.DataFrame()
    pf_details = pd.DataFrame()
    pf_no_date_details = pd.DataFrame()
    pa_no_date_details = pd.DataFrame()
    pa_old_details = pd.DataFrame()

    if so is not None and not so.empty:
        ro = so[so["Sales Rep"] == rep].copy()
        if not ro.empty:
            # Pending Approval
            pa_orders = ro[ro["Status"] == "Pending Approval"].copy()
            if not pa_orders.empty:
                if "Pending Approval Date" in pa_orders.columns:
                    mask_q4 = (
                        pa_orders["Pending Approval Date"].notna()
                        & (pa_orders["Pending Approval Date"] >= Q4_START)
                        & (pa_orders["Pending Approval Date"] <= Q4_END)
                    )
                    pa_details = dedupe_columns(pa_orders[mask_q4].copy())
                    pa = float(pa_details["Amount"].sum())
                    pa_no_date_details = dedupe_columns(pa_orders[~mask_q4].copy())
                    pa_no_date = float(pa_no_date_details["Amount"].sum())
                if "Age_Business_Days" in pa_orders.columns:
                    old_mask = pa_orders["Age_Business_Days"] > 14
                    pa_old_details = dedupe_columns(pa_orders[old_mask].copy())
                    pa_old = float(pa_old_details["Amount"].sum())

            # Pending Fulfillment
            pf_orders = ro[ro["Status"].isin(["Pending Fulfillment", "Pending Billing/Partially Fulfilled"])].copy()
            if not pf_orders.empty:
                def has_q4_date(row):
                    d1 = row.get("Customer Promise Date")
                    d2 = row.get("Projected Date")
                    return (pd.notna(d1) and Q4_START <= d1 <= Q4_END) or (pd.notna(d2) and Q4_START <= d2 <= Q4_END)
                pf_orders["Has_Q4_Date"] = pf_orders.apply(has_q4_date, axis=1)
                pf_details = dedupe_columns(pf_orders[pf_orders["Has_Q4_Date"] == True].copy())
                pf = float(pf_details["Amount"].sum())
                no_date_mask = pf_orders["Customer Promise Date"].isna() & pf_orders["Projected Date"].isna()
                pf_no_date_details = dedupe_columns(pf_orders[no_date_mask].copy())
                pf_no_date = float(pf_no_date_details["Amount"].sum())

    total_progress = orders + expect_commit + pa + pf
    gap = quota - total_progress
    attn = (total_progress / quota * 100.0) if quota > 0 else 0.0
    potential = ((total_progress + best_opp) / quota * 100.0) if quota > 0 else 0.0

    return {
        "quota": quota,
        "orders": orders,
        "expect_commit": float(expect_commit),
        "best_opp": float(best_opp),
        "gap": float(gap),
        "attainment_pct": float(attn),
        "potential_attainment": float(potential),
        "total_progress": float(total_progress),
        "pending_approval": float(pa),
        "pending_approval_no_date": float(pa_no_date),
        "pending_approval_old": float(pa_old),
        "pending_fulfillment": float(pf),
        "pending_fulfillment_no_date": float(pf_no_date),
        "q1_spillover_expect_commit": float(q1_ec),
        "q1_spillover_best_opp": float(q1_bo),
        "q1_spillover_total": float(q1_ec + q1_bo),
        # Details for drill‑downs
        "pending_approval_details": pa_details,
        "pending_approval_no_date_details": pa_no_date_details,
        "pending_approval_old_details": pa_old_details,
        "pending_fulfillment_details": pf_details,
        "pending_fulfillment_no_date_details": pf_no_date_details,
        "expect_commit_deals": r_q4[r_q4["Status"].isin(["Expect", "Commit"])].copy(),
        "best_opp_deals": r_q4[r_q4["Status"].isin(["Best Case", "Opportunity"])].copy(),
        "expect_commit_q1_spillover_deals": r_q1[r_q1["Status"].isin(["Expect", "Commit"])].copy(),
        "best_opp_q1_spillover_deals": r_q1[r_q1["Status"].isin(["Best Case", "Opportunity"])].copy(),
        "all_q1_spillover_deals": r_q1.copy(),
        "total_q4_closing_deals": int(len(rdeals)),
        "total_q4_closing_amount": float(rdeals["Amount"].sum() if not rdeals.empty else 0.0),
    }


def calc_team_metrics(deals: pd.DataFrame, dash: pd.DataFrame):
    total_quota = float(dash["Quota"].sum()) if not dash.empty else 0.0
    total_orders = float(dash["NetSuite Orders"].sum()) if not dash.empty else 0.0
    deals_q4 = deals[deals.get("Ships_In_Q4", True) == True]
    expect_commit = float(deals_q4[deals_q4["Status"].isin(["Expect", "Commit"])]["Amount"].sum())
    best_opp = float(deals_q4[deals_q4["Status"].isin(["Best Case", "Opportunity"])]["Amount"].sum())
    current = expect_commit + total_orders
    gap = total_quota - current
    attn = (current / total_quota * 100) if total_quota > 0 else 0
    potential = ((current + best_opp) / total_quota * 100) if total_quota > 0 else 0

    # Spillover (if any ships in Q1)
    q1_spill = float(deals[deals.get("Ships_In_Q1", False) == True]["Amount"].sum())

    return {
        "total_quota": total_quota,
        "total_orders": total_orders,
        "expect_commit": expect_commit,
        "best_opp": best_opp,
        "gap": float(gap),
        "attainment_pct": float(attn),
        "potential_attainment": float(potential),
        "current_forecast": float(current),
        "q1_spillover": q1_spill,
    }

# ==================== VISUALS ====================

def gap_chart(metrics: dict, title: str) -> go.Figure:
    fig = go.Figure()
    orders_val = metrics.get("total_orders", metrics.get("orders", 0))
    quota_val = metrics.get("total_quota", metrics.get("quota", 0))

    fig.add_trace(
        go.Bar(
            name="NetSuite Orders",
            x=["Progress"],
            y=[orders_val],
            marker_color=COLORS["primary"],
            text=[f"${orders_val:,.0f}"],
            textposition="inside",
        )
    )
    fig.add_trace(
        go.Bar(
            name="Expect/Commit",
            x=["Progress"],
            y=[metrics.get("expect_commit", 0)],
            marker_color=COLORS["success"],
            text=[f"${metrics.get('expect_commit', 0):,.0f}"],
            textposition="inside",
        )
    )
    fig.add_trace(
        go.Scatter(
            name="Quota Goal",
            x=["Progress"],
            y=[quota_val],
            mode="markers",
            marker=dict(size=12, color=COLORS["danger"], symbol="diamond"),
        )
    )
    potential = metrics.get("expect_commit", 0) + metrics.get("best_opp", 0) + orders_val
    fig.add_trace(
        go.Scatter(
            name="Potential (if all close)",
            x=["Progress"],
            y=[potential],
            mode="markers",
            marker=dict(size=12, color=COLORS["warning"], symbol="diamond"),
        )
    )
    fig.update_layout(
        title=title,
        barmode="stack",
        height=400,
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        yaxis_title="Amount ($)",
        xaxis_title="",
    )
    return fig


def status_breakdown_chart(deals: pd.DataFrame, rep: str | None = None):
    df = deals.copy()
    if rep:
        df = df[df["Deal Owner"] == rep]
    df = df[df.get("Ships_In_Q4", True) == True]
    if df.empty:
        return None
    summary = df.groupby("Status")["Amount"].sum().reset_index()
    color_map = {"Expect": COLORS["expect"], "Commit": COLORS["commit"], "Best Case": COLORS["best_case"], "Opportunity": COLORS["opportunity"]}
    fig = px.pie(summary, values="Amount", names="Status", title="Deal Amount by Forecast Category (Q4)", color="Status", color_discrete_map=color_map, hole=0.4)
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(height=400)
    return fig


def pipeline_breakdown_chart(deals: pd.DataFrame, rep: str | None = None):
    df = deals.copy()
    if rep:
        df = df[df["Deal Owner"] == rep]
    df = df[df.get("Ships_In_Q4", True) == True]
    if df.empty:
        return None
    summary = df.groupby(["Pipeline", "Status"])["Amount"].sum().reset_index()
    color_map = {"Expect": COLORS["expect"], "Commit": COLORS["commit"], "Best Case": COLORS["best_case"], "Opportunity": COLORS["opportunity"]}
    fig = px.bar(summary, x="Pipeline", y="Amount", color="Status", title="Pipeline Breakdown (Q4)", color_discrete_map=color_map, text_auto=".2s", barmode="stack")
    fig.update_layout(height=400, yaxis_title="Amount ($)")
    return fig


def timeline_chart(deals: pd.DataFrame, rep: str | None = None):
    df = deals.copy()
    if rep:
        df = df[df["Deal Owner"] == rep]
    df = df[df["Close Date"].notna()]
    if df.empty:
        return None
    df = df.sort_values("Close Date")
    df["Quarter"] = df.apply(lambda r: "Q4 2025" if r.get("Ships_In_Q4", True) else "Q1 2026", axis=1)
    color_map = {"Expect": COLORS["expect"], "Commit": COLORS["commit"], "Best Case": COLORS["best_case"], "Opportunity": COLORS["opportunity"]}
    fig = px.scatter(df, x="Close Date", y="Amount", color="Status", size="Amount", hover_data=["Deal Name", "Amount", "Pipeline", "Quarter"], title="Deal Close Date Timeline", color_discrete_map=color_map)
    try:
        fig.add_vline(x=datetime(2025, 12, 31), line_dash="dash", line_color="red", annotation_text="Q4/Q1 Boundary")
    except Exception:
        pass
    fig.update_layout(height=400, yaxis_title="Deal Amount ($)")
    return fig


def invoice_status_chart(inv: pd.DataFrame, rep: str | None = None):
    if inv.empty:
        return None
    df = inv if rep is None else inv[inv["Sales Rep"] == rep]
    if df.empty:
        return None
    summary = df.groupby("Status")["Amount"].sum().reset_index()
    fig = px.pie(summary, values="Amount", names="Status", title="Invoice Amount by Status", hole=0.4)
    fig.update_traces(textposition="inside", textinfo="percent+label")
    fig.update_layout(height=400)
    return fig

# ---------- UI helpers ----------

def money(x: float) -> str:
    return f"${x:,.0f}"


def drilldown_table(title: str, amount: float, df: pd.DataFrame):
    with st.expander(f"{title}: {money(amount)} (Click to view {len(df)} items)"):
        if df.empty:
            st.info("No items")
            return
        df = dedupe_columns(df.copy())
        # Choose relevant columns
        if "Deal Name" in df.columns:  # HubSpot deals
            cols = [c for c in ["Deal Name", "Amount", "Status", "Pipeline", "Close Date", "Product Type"] if c in df.columns]
        else:  # Sales Orders
            cols = [c for c in ["Document Number", "Customer", "Amount", "Status", "Order Start Date", "Pending Approval Date", "Customer Promise Date", "Projected Date"] if c in df.columns]
        display = df[cols].copy()
        if "Amount" in display.columns:
            display["Amount"] = display["Amount"].apply(lambda v: f"${v:,.2f}")
        for c in display.columns:
            if "Date" in c and pd.api.types.is_datetime64_any_dtype(display[c]):
                display[c] = display[c].dt.strftime("%Y-%m-%d")
        st.dataframe(display, use_container_width=True, hide_index=True)
        st.caption(f"Total: {money(df['Amount'].sum() if 'Amount' in df.columns else 0)} | Count: {len(df)}")


def progress_breakdown(metrics: dict):
    st.markdown(
        f"""
        <div class='progress-breakdown'>
          <div class='progress-row'><span>📦 Invoiced (Orders Shipped)</span><span><b>{money(metrics['orders'])}</b></span></div>
          <div class='progress-row'><span>📤 Pending Fulfillment (with dates)</span><span><b>{money(metrics['pending_fulfillment'])}</b></span></div>
          <div class='progress-row'><span>⏳ Pending Approval (with dates)</span><span><b>{money(metrics['pending_approval'])}</b></span></div>
          <div class='progress-row'><span>✅ HubSpot Expect/Commit (Q4)</span><span><b>{money(metrics['expect_commit'])}</b></span></div>
          <div class='progress-row'><span>🎯 TOTAL PROGRESS</span><span>{money(metrics['total_progress'])}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==================== VIEWS ====================

def view_team(deals: pd.DataFrame, dash: pd.DataFrame, inv: pd.DataFrame):
    st.title("🎯 Team Sales Dashboard — Q4 2025")
    m = calc_team_metrics(deals, dash)

    if m.get("q1_spillover", 0) > 0:
        st.warning(f"📅 Q1 2026 Spillover: {money(m['q1_spillover'])} in deals expected to ship in Q1 2026.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Quota", money(m["total_quota"]))
    with c2:
        st.metric("Current Forecast", money(m["current_forecast"]), delta=f"{m['attainment_pct']:.1f}% of quota")
    with c3:
        st.metric("Gap to Goal", money(m["gap"]))
    with c4:
        st.metric("Potential Attainment", f"{m['potential_attainment']:.1f}%", delta=f"+{m['potential_attainment'] - m['attainment_pct']:.1f}% upside")

    st.markdown("### 📈 Progress to Quota")
    st.progress(min(m["attainment_pct"] / 100, 1.0))
    st.caption(f"Current: {m['attainment_pct']:.1f}% | Potential: {m['potential_attainment']:.1f}%")

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(gap_chart(m, "Team Progress to Goal"), use_container_width=True)
    with c2:
        ch = status_breakdown_chart(deals)
        st.plotly_chart(ch, use_container_width=True) if ch else st.info("No deal data for status breakdown")

    st.markdown("### 🔄 Pipeline Analysis")
    ch = pipeline_breakdown_chart(deals)
    st.plotly_chart(ch, use_container_width=True) if ch else st.info("No pipeline data")

    st.markdown("### 📅 Deal Close Timeline")
    ch = timeline_chart(deals)
    st.plotly_chart(ch, use_container_width=True) if ch else st.info("No timeline data")

    if not inv.empty:
        st.markdown("### 💰 Invoice Status Breakdown")
        ch = invoice_status_chart(inv)
        st.plotly_chart(ch, use_container_width=True) if ch else None

    st.markdown("### 👥 Rep Summary")
    rows = []
    for rep in dash["Rep Name"].tolist():
        rm = calc_rep_metrics(rep, deals, dash)
        if rm:
            rows.append({
                "Rep": rep,
                "Quota": money(rm["quota"]),
                "Orders": money(rm["orders"]),
                "Expect/Commit": money(rm["expect_commit"]),
                "Pending Approval": money(rm["pending_approval"]),
                "Pending Fulfillment": money(rm["pending_fulfillment"]),
                "Total Progress": money(rm["total_progress"]),
                "Gap": money(rm["gap"]),
                "Attainment": f"{rm['attainment_pct']:.1f}%",
                "Q1 Spillover": money(rm.get("q1_spillover_total", 0)),
            })
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def view_rep(rep: str, deals: pd.DataFrame, dash: pd.DataFrame, inv: pd.DataFrame, so: pd.DataFrame):
    st.title(f"👤 {rep} — Q4 2025 Forecast")
    m = calc_rep_metrics(rep, deals, dash, so)
    if not m:
        st.error(f"No data for {rep}")
        return

    if m.get("total_q4_closing_deals", 0) > 0:
        st.markdown(
            f"""
            <div class='info-card'>
              <b>📋 Total Q4 2025 Pipeline:</b> {m['total_q4_closing_deals']} deals worth {money(m['total_q4_closing_amount'])}<br/>
              <small>Note: {money(m.get('q1_spillover_total', 0))} expected to ship in Q1 2026 based on lead times.</small>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### 💰 Section 1: Q4 Gap to Goal Components")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: st.metric("Quota", money(m["quota"]))
    with c2: st.metric("Invoiced", money(m["orders"]))
    with c3: st.metric("Pending Fulfillment", money(m["pending_fulfillment"]))
    with c4: st.metric("Pending Approval", money(m["pending_approval"]))
    with c5: st.metric("HubSpot (Q4)", money(m["expect_commit"]))
    with c6: st.metric("Gap to Goal", money(m["gap"]))

    progress_breakdown(m)

    st.markdown("#### 📊 Section 1 Drill‑downs")
    c1, c2 = st.columns(2)
    with c1:
        drilldown_table("📤 Pending Fulfillment (with Q4 dates)", m["pending_fulfillment"], m["pending_fulfillment_details"])
        drilldown_table("⏳ Pending Approval (with dates)", m["pending_approval"], m["pending_approval_details"])
    with c2:
        drilldown_table("✅ HubSpot Expect/Commit Deals", m["expect_commit"], m["expect_commit_deals"])
        drilldown_table("🎯 Best Case/Opportunity Deals", m["best_opp"], m["best_opp_deals"])

    st.markdown("### 📊 Section 2: Additional Orders (Can be included)")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("PF SO's No Date", money(m["pending_fulfillment_no_date"]))
        drilldown_table("PF Orders Without Dates", m["pending_fulfillment_no_date"], m["pending_fulfillment_no_date_details"])
    with c2:
        st.metric("PA SO's No Date", money(m["pending_approval_no_date"]))
        drilldown_table("PA Orders Without Dates", m["pending_approval_no_date"], m["pending_approval_no_date_details"])
    with c3:
        st.metric("Old PA (>14 bdays)", money(m["pending_approval_old"]))
        drilldown_table("Old Pending Approval Orders", m["pending_approval_old"], m["pending_approval_old_details"])

    st.markdown("### 🚢 Section 3: Q1 2026 Spillover")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Expect/Commit (Q1 Ship)", money(m.get("q1_spillover_expect_commit", 0)))
        drilldown_table("Expect/Commit Deals Shipping Q1 2026", m.get("q1_spillover_expect_commit", 0), m.get("expect_commit_q1_spillover_deals", pd.DataFrame()))
    with c2:
        st.metric("Best Case/Opp (Q1 Ship)", money(m.get("q1_spillover_best_opp", 0)))
        drilldown_table("Best Case/Opp Deals Shipping Q1 2026", m.get("q1_spillover_best_opp", 0), m.get("best_opp_q1_spillover_deals", pd.DataFrame()))
    with c3:
        st.metric("Total Q1 Spillover", money(m.get("q1_spillover_total", 0)))
        drilldown_table("All Q1 2026 Spillover Deals", m.get("q1_spillover_total", 0), m.get("all_q1_spillover_deals", pd.DataFrame()))

    final_total = m["total_progress"] + m["pending_fulfillment_no_date"] + m["pending_approval_no_date"] + m["pending_approval_old"]
    st.metric("📊 FINAL TOTAL Q4", money(final_total), delta=f"Section 1: {money(m['total_progress'])} + Section 2: {money(final_total - m['total_progress'])}")

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(gap_chart(m, f"{rep}'s Progress to Goal"), use_container_width=True)
    with c2:
        ch = status_breakdown_chart(deals, rep)
        st.plotly_chart(ch, use_container_width=True) if ch else st.info("No deal data for this rep")

    st.markdown("### 🔄 Pipeline Analysis")
    ch = pipeline_breakdown_chart(deals, rep)
    st.plotly_chart(ch, use_container_width=True) if ch else st.info("No pipeline data for this rep")

    st.markdown("### 📅 Deal Close Timeline")
    ch = timeline_chart(deals, rep)
    st.plotly_chart(ch, use_container_width=True) if ch else st.info("No timeline data for this rep")


def view_reconciliation(deals: pd.DataFrame, dash: pd.DataFrame, so: pd.DataFrame):
    st.title("🔍 Forecast Reconciliation")
    st.info("Side‑by‑side comparison using Section 1 + optional Section 2 adds.")

    # Example boss snapshot — replace with your latest numbers if desired
    boss = {
        "Jake Lynch": {"invoiced": 577540, "pending_fulfillment": 263183, "pending_approval": 45220, "hubspot": 340756, "total": 1226699, "pending_fulfillment_so_no_date": 87891, "pending_approval_so_no_date": 0, "old_pending_approval": 33741, "total_q4": 1350638},
        "Dave Borkowski": {"invoiced": 237849, "pending_fulfillment": 160537, "pending_approval": 13390, "hubspot": 414768, "total": 826545, "pending_fulfillment_so_no_date": 45471, "pending_approval_so_no_date": 0, "old_pending_approval": 12244, "total_q4": 884260},
        "Alex Gonzalez": {"invoiced": 314523, "pending_fulfillment": 190865, "pending_approval": 0, "hubspot": 0, "total": 505387, "pending_fulfillment_so_no_date": 41710, "pending_approval_so_no_date": 79361, "old_pending_approval": 4900, "total_q4": 631358},
        "Brad Sherman": {"invoiced": 118211, "pending_fulfillment": 38330, "pending_approval": 28984, "hubspot": 183103, "total": 368629, "pending_fulfillment_so_no_date": 29970, "pending_approval_so_no_date": 0, "old_pending_approval": 884, "total_q4": 399482},
        "Lance Mitton": {"invoiced": 23948, "pending_fulfillment": 2027, "pending_approval": 3331, "hubspot": 11000, "total": 38356, "pending_fulfillment_so_no_date": 1613, "pending_approval_so_no_date": 0, "old_pending_approval": 60527, "total_q4": 100496},
        "House": {"invoiced": 0, "pending_fulfillment": 899, "pending_approval": 0, "hubspot": 0, "total": 0, "pending_fulfillment_so_no_date": 0, "pending_approval_so_no_date": 0, "old_pending_approval": 0, "total_q4": 0},
        "Shopify ECommerce": {"invoiced": 21348, "pending_fulfillment": 0, "pending_approval": 1174, "hubspot": 0, "total": 23421, "pending_fulfillment_so_no_date": 0, "pending_approval_so_no_date": 0, "old_pending_approval": 1544, "total_q4": 24965},
    }

    st.markdown("<div class='section-header'>Section 1: Q4 Gap to Goal</div>", unsafe_allow_html=True)

    rows = []
    totals = {k: 0.0 for k in ["invoiced_you", "invoiced_boss", "pf_you", "pf_boss", "pa_you", "pa_boss", "hs_you", "hs_boss", "total_you", "total_boss"]}

    reps = list(boss.keys())
    for rep in reps:
        rm = calc_rep_metrics(rep, deals, dash, so) if rep in dash["Rep Name"].values else None
        b = boss[rep]
        inv_you = rm["orders"] if rm else 0
        pf_you = rm["pending_fulfillment"] if rm else 0
        pa_you = rm["pending_approval"] if rm else 0
        hs_you = rm["expect_commit"] if rm else 0
        total_you = inv_you + pf_you + pa_you + hs_you

        totals["invoiced_you"] += inv_you; totals["invoiced_boss"] += b["invoiced"]
        totals["pf_you"] += pf_you; totals["pf_boss"] += b["pending_fulfillment"]
        totals["pa_you"] += pa_you; totals["pa_boss"] += b["pending_approval"]
        totals["hs_you"] += hs_you; totals["hs_boss"] += b["hubspot"]
        totals["total_you"] += total_you; totals["total_boss"] += b["total"]

        rows.append({
            "Rep": rep,
            "Invoiced": money(inv_you), "Invoiced (Boss)": money(b["invoiced"]),
            "Pending Fulfillment": money(pf_you), "PF (Boss)": money(b["pending_fulfillment"]),
            "Pending Approval": money(pa_you), "PA (Boss)": money(b["pending_approval"]),
            "HubSpot": money(hs_you), "HS (Boss)": money(b["hubspot"]),
            "Total": money(total_you), "Total (Boss)": money(b["total"]),
        })

    rows.append({
        "Rep": "TOTAL",
        "Invoiced": money(totals["invoiced_you"]), "Invoiced (Boss)": money(totals["invoiced_boss"]),
        "Pending Fulfillment": money(totals["pf_you"]), "PF (Boss)": money(totals["pf_boss"]),
        "Pending Approval": money(totals["pa_you"]), "PA (Boss)": money(totals["pa_boss"]),
        "HubSpot": money(totals["hs_you"]), "HS (Boss)": money(totals["hs_boss"]),
        "Total": money(totals["total_you"]), "Total (Boss)": money(totals["total_boss"]),
    })

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("<div class='section-header'>Section 2: Additional Orders</div>", unsafe_allow_html=True)

    rows2 = []
    totals2 = {k: 0.0 for k in ["pf_no_date_you", "pf_no_date_boss", "pa_no_date_you", "pa_no_date_boss", "old_pa_you", "old_pa_boss", "final_you", "final_boss"]}

    for rep in reps:
        rm = calc_rep_metrics(rep, deals, dash, so) if rep in dash["Rep Name"].values else None
        b = boss[rep]
        pfnd_you = rm["pending_fulfillment_no_date"] if rm else 0
        pand_you = rm["pending_approval_no_date"] if rm else 0
        oldpa_you = rm["pending_approval_old"] if rm else 0
        section1_you = (rm["orders"] + rm["pending_fulfillment"] + rm["pending_approval"] + rm["expect_commit"]) if rm else 0
        final_you = section1_you + pfnd_you + pand_you + oldpa_you

        totals2["pf_no_date_you"] += pfnd_you; totals2["pf_no_date_boss"] += b["pending_fulfillment_so_no_date"]
        totals2["pa_no_date_you"] += pand_you; totals2["pa_no_date_boss"] += b["pending_approval_so_no_date"]
        totals2["old_pa_you"] += oldpa_you; totals2["old_pa_boss"] += b["old_pending_approval"]
        totals2["final_you"] += final_you; totals2["final_boss"] += b["total_q4"]

        rows2.append({
            "Rep": rep,
            "PF SO's No Date": money(pfnd_you), "PF No Date (Boss)": money(b["pending_fulfillment_so_no_date"]),
            "PA SO's No Date": money(pand_you), "PA No Date (Boss)": money(b["pending_approval_so_no_date"]),
            "Old PA (>2 weeks)": money(oldpa_you), "Old PA (Boss)": money(b["old_pending_approval"]),
            "Total Q4": money(final_you), "Total Q4 (Boss)": money(b["total_q4"]),
        })

    rows2.append({
        "Rep": "TOTAL",
        "PF SO's No Date": money(totals2["pf_no_date_you"]), "PF No Date (Boss)": money(totals2["pf_no_date_boss"]),
        "PA SO's No Date": money(totals2["pa_no_date_you"]), "PA No Date (Boss)": money(totals2["pa_no_date_boss"]),
        "Old PA (>2 weeks)": money(totals2["old_pa_you"]), "Old PA (Boss)": money(totals2["old_pa_boss"]),
        "Total Q4": money(totals2["final_you"]), "Total Q4 (Boss)": money(totals2["final_boss"]),
    })

    st.dataframe(pd.DataFrame(rows2), use_container_width=True, hide_index=True)

    # Summary metrics
    diff1 = totals["total_boss"] - totals["total_you"]
    diff_final = totals2["final_boss"] - totals2["final_you"]
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Section 1 Variance", money(abs(diff1)), delta=("Under" if diff1 > 0 else "Over") + f" by {money(abs(diff1))}")
    with c2:
        st.metric("Total Q4 Variance", money(abs(diff_final)), delta=("Under" if diff_final > 0 else "Over") + f" by {money(abs(diff_final))}")
    with c3:
        if totals2["final_boss"] > 0:
            accuracy = (1 - abs(diff_final) / totals2["final_boss"]) * 100
            st.metric("Accuracy", f"{accuracy:.1f}%")
        else:
            st.metric("Accuracy", "N/A")

# ==================== MAIN ====================

def main():
    load_custom_css()

    # -------- Sidebar --------
    with st.sidebar:
        st.markdown("""
        <div style='text-align:center; padding:16px;'>
          <h2>CALYX</h2>
          <p style='font-size:12px; letter-spacing:3px;'>CONTAINERS</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        view = st.radio("Select View:", ["Team Overview", "Individual Rep", "Reconciliation"], index=0)
        st.markdown("---")

        st.caption(f"Spreadsheet ID: {SPREADSHEET_ID}")
        st.caption(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if st.button("🔄 Refresh Data Now"):
            st.cache_data.clear(); st.rerun()

    # -------- Data load --------
    with st.spinner("Loading Google Sheets data…"):
        deals_raw = load_google_sheets_data("All Reps All Pipelines", "A:Q")
        dash_raw = load_google_sheets_data("Dashboard Info", "A:C")
        so_raw = load_google_sheets_data("NS Sales Orders", "A:AB")
        inv_raw = load_google_sheets_data("NS Invoices", "A:Z")

    deals = process_deals(deals_raw)
    dash = process_dashboard(dash_raw)
    so = process_sales_orders(so_raw)
    inv = process_invoices(inv_raw)

    if deals.empty and dash.empty:
        st.error("❌ Unable to load required data. Check Google Sheets access and secrets.")
        return
    if deals.empty:
        st.warning("⚠️ Deals data empty — verify 'All Reps All Pipelines'.")
    if dash.empty:
        st.warning("⚠️ Dashboard info empty — verify 'Dashboard Info'.")

    # -------- Views --------
    if view == "Team Overview":
        view_team(deals, dash, inv)
    elif view == "Individual Rep":
        if dash.empty:
            st.error("No reps available.")
            return
        rep = st.selectbox("Select Rep:", options=dash["Rep Name"].tolist())
        if rep:
            view_rep(rep, deals, dash, inv, so)
    else:
        view_reconciliation(deals, dash, so)


if __name__ == "__main__":
    main()
