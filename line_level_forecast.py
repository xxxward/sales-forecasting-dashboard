"""
app_enhanced.py
---------------

Enhanced Sales Forecasting Dashboard with Line-Level Forecasting
Integrates aggregate and line-item level forecasting capabilities
"""

from __future__ import annotations

import io
import os
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

from forecasting import (
    aggregate_sales,
    compute_growth_rates,
    forecast_sales,
    prepare_dataframe,
)
from utils import clean_numeric, parse_dates
from openai_assistant import ask_ai


def load_uploaded_file(uploaded_file: st.runtime.uploaded_file_manager.UploadedFile) -> pd.DataFrame:
    """Read a file uploaded via Streamlit into a DataFrame."""
    if uploaded_file is None:
        return pd.DataFrame()
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith(".csv"):
            return pd.read_csv(uploaded_file)
        elif file_name.endswith((".xls", ".xlsx")):
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file)
        else:
            st.warning(f"Unsupported file type: {file_name}")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Failed to parse file {file_name}: {e}")
        return pd.DataFrame()


def configure_sidebar() -> Dict[str, pd.DataFrame]:
    """Render the sidebar components for data uploads and parameter selection."""
    st.sidebar.header("Data Upload & Configuration")

    data_frames: Dict[str, pd.DataFrame] = {}

    # File uploaders
    hubspot_file = st.sidebar.file_uploader(
        "Upload Hubspot data (CSV or Excel)", type=["csv", "xls", "xlsx"], key="hubspot"
    )
    netsuite_file = st.sidebar.file_uploader(
        "Upload Netsuite line-level data (CSV or Excel)", type=["csv", "xls", "xlsx"], key="netsuite"
    )
    forecast_file = st.sidebar.file_uploader(
        "Upload Boss's forecast (optional)", type=["csv", "xls", "xlsx"], key="boss_forecast"
    )

    # Load data frames
    if hubspot_file:
        data_frames["Hubspot"] = load_uploaded_file(hubspot_file)
    if netsuite_file:
        data_frames["Netsuite"] = load_uploaded_file(netsuite_file)
    if forecast_file:
        data_frames["Boss Forecast"] = load_uploaded_file(forecast_file)

    return data_frames


def select_columns(df: pd.DataFrame, context: str = "") -> Optional[Dict[str, str]]:
    """Allow the user to map columns for date, product, and value."""
    if df.empty:
        st.info(f"No data to configure{' for ' + context if context else ''}.")
        return None

    cols = df.columns.tolist()
    st.write(f"### Column Mapping{' - ' + context if context else ''}")
    date_col = st.selectbox(f"Select date column{' (' + context + ')' if context else ''}", options=cols, key=f"date_{context}")
    product_col = st.selectbox(f"Select product column{' (' + context + ')' if context else ''}", options=cols, key=f"product_{context}")
    value_col = st.selectbox(f"Select value (sales) column{' (' + context + ')' if context else ''}", options=cols, key=f"value_{context}")
    return {"date": date_col, "product": product_col, "value": value_col}


def prepare_line_level_data(
    df: pd.DataFrame,
    date_col: str,
    product_col: str,
    value_col: str,
    customer_col: Optional[str] = None,
    order_col: Optional[str] = None,
) -> pd.DataFrame:
    """Prepare line-level data with additional context columns."""
    df = df.copy()
    
    # Parse dates
    df[date_col] = parse_dates(df[date_col])
    
    # Extract temporal features
    df["Year"] = df[date_col].dt.year
    df["Quarter"] = df[date_col].dt.quarter
    df["Month"] = df[date_col].dt.month
    
    # Clean numeric values
    df[value_col] = df[value_col].apply(clean_numeric).astype(float)
    
    # Ensure product column is string type
    df[product_col] = df[product_col].astype(str)
    
    # Build output columns
    output_cols = [date_col, product_col, "Year", "Quarter", "Month", value_col]
    
    if customer_col and customer_col in df.columns:
        df[customer_col] = df[customer_col].astype(str)
        output_cols.append(customer_col)
    
    if order_col and order_col in df.columns:
        df[order_col] = df[order_col].astype(str)
        output_cols.append(order_col)
    
    return df[output_cols].copy()


def compute_line_level_forecast(
    df: pd.DataFrame,
    product_col: str,
    value_col: str,
    forecast_quarters: int = 4,
) -> pd.DataFrame:
    """Generate line-level forecasts based on historical patterns."""
    
    # Get the most recent quarter's data
    max_year = df["Year"].max()
    max_quarter = df[df["Year"] == max_year]["Quarter"].max()
    
    # Calculate average sales per product
    avg_by_product = df.groupby(product_col)[value_col].mean().reset_index()
    avg_by_product.columns = [product_col, "AvgSales"]
    
    # Calculate quarter-over-quarter growth rates
    quarterly_sales = df.groupby([product_col, "Year", "Quarter"])[value_col].sum().reset_index()
    quarterly_sales = quarterly_sales.sort_values([product_col, "Year", "Quarter"])
    
    growth_rates = {}
    for product in quarterly_sales[product_col].unique():
        product_data = quarterly_sales[quarterly_sales[product_col] == product]
        if len(product_data) >= 2:
            sales_values = product_data[value_col].values
            growth = (sales_values[-1] / sales_values[-2] - 1) if sales_values[-2] != 0 else 0
            growth_rates[product] = growth
        else:
            growth_rates[product] = 0
    
    # Generate forecasts for future quarters
    forecast_rows = []
    
    for i in range(1, forecast_quarters + 1):
        # Calculate forecast quarter
        forecast_q = max_quarter + i
        forecast_y = max_year
        
        while forecast_q > 4:
            forecast_q -= 4
            forecast_y += 1
        
        for _, row in avg_by_product.iterrows():
            product = row[product_col]
            base_sales = row["AvgSales"]
            growth = growth_rates.get(product, 0)
            
            # Apply compound growth
            forecast_sales = base_sales * ((1 + growth) ** i)
            
            forecast_rows.append({
                product_col: product,
                "Year": forecast_y,
                "Quarter": forecast_q,
                "ForecastSales": forecast_sales,
                "GrowthRate": growth,
            })
    
    forecast_df = pd.DataFrame(forecast_rows)
    return forecast_df


def render_line_level_tab(df_raw: pd.DataFrame, column_mapping: Dict[str, str]) -> None:
    """Render the line-level forecasting tab."""
    st.subheader("🔍 Line-Level Forecasting")
    
    st.write("""
    This section provides detailed line-item level forecasting, allowing you to:
    - Analyze individual transaction patterns
    - Forecast at a granular product level
    - Track customer-specific trends
    - Generate detailed projections for operational planning
    """)
    
    # Additional column selectors for line-level data
    st.write("#### Optional Additional Columns")
    cols = df_raw.columns.tolist()
    
    col1, col2 = st.columns(2)
    with col1:
        customer_col = st.selectbox(
            "Customer column (optional)",
            options=["None"] + cols,
            key="customer_col_select"
        )
        customer_col = None if customer_col == "None" else customer_col
    
    with col2:
        order_col = st.selectbox(
            "Order ID column (optional)",
            options=["None"] + cols,
            key="order_col_select"
        )
        order_col = None if order_col == "None" else order_col
    
    # Prepare line-level data
    df_line = prepare_line_level_data(
        df_raw,
        date_col=column_mapping["date"],
        product_col=column_mapping["product"],
        value_col=column_mapping["value"],
        customer_col=customer_col,
        order_col=order_col,
    )
    
    # Forecast parameters
    st.write("#### Forecast Parameters")
    col1, col2 = st.columns(2)
    
    with col1:
        forecast_quarters = st.number_input(
            "Number of quarters to forecast",
            min_value=1,
            max_value=8,
            value=4,
            step=1
        )
    
    with col2:
        confidence_level = st.slider(
            "Confidence level for predictions",
            min_value=0.8,
            max_value=0.99,
            value=0.95,
            step=0.01
        )
    
    # Generate line-level forecast
    forecast_line = compute_line_level_forecast(
        df_line,
        product_col=column_mapping["product"],
        value_col=column_mapping["value"],
        forecast_quarters=forecast_quarters,
    )
    
    # Display results in sub-tabs
    line_tabs = st.tabs([
        "Forecast Summary",
        "Detailed Forecast",
        "Historical Data",
        "Visualizations"
    ])
    
    with line_tabs[0]:
        st.write("#### Forecast Summary by Product")
        summary = forecast_line.groupby(column_mapping["product"]).agg({
            "ForecastSales": ["sum", "mean"],
            "GrowthRate": "first"
        }).reset_index()
        summary.columns = [column_mapping["product"], "Total Forecast", "Avg per Quarter", "Growth Rate"]
        summary["Total Forecast"] = summary["Total Forecast"].round(2)
        summary["Avg per Quarter"] = summary["Avg per Quarter"].round(2)
        summary["Growth Rate"] = (summary["Growth Rate"] * 100).round(2).astype(str) + "%"
        st.dataframe(summary, use_container_width=True)
    
    with line_tabs[1]:
        st.write("#### Detailed Quarterly Forecast")
        display_forecast = forecast_line.copy()
        display_forecast["ForecastSales"] = display_forecast["ForecastSales"].round(2)
        display_forecast["GrowthRate"] = (display_forecast["GrowthRate"] * 100).round(2).astype(str) + "%"
        st.dataframe(display_forecast, use_container_width=True)
        
        # Download button
        csv = display_forecast.to_csv(index=False)
        st.download_button(
            label="Download Forecast as CSV",
            data=csv,
            file_name="line_level_forecast.csv",
            mime="text/csv",
        )
    
    with line_tabs[2]:
        st.write("#### Historical Line-Level Data")
        st.write(f"Showing {len(df_line)} transactions")
        
        # Filters
        col1, col2 = st.columns(2)
        with col1:
            products = ["All"] + sorted(df_line[column_mapping["product"]].unique().tolist())
            selected_product = st.selectbox("Filter by product", products)
        
        with col2:
            years = ["All"] + sorted(df_line["Year"].unique().tolist(), reverse=True)
            selected_year = st.selectbox("Filter by year", years)
        
        # Apply filters
        filtered_df = df_line.copy()
        if selected_product != "All":
            filtered_df = filtered_df[filtered_df[column_mapping["product"]] == selected_product]
        if selected_year != "All":
            filtered_df = filtered_df[filtered_df["Year"] == selected_year]
        
        st.dataframe(filtered_df, use_container_width=True)
        
        # Summary statistics
        st.write("#### Summary Statistics")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Transactions", len(filtered_df))
        with col2:
            st.metric("Total Sales", f"${filtered_df[column_mapping['value']].sum():,.2f}")
        with col3:
            st.metric("Average Sale", f"${filtered_df[column_mapping['value']].mean():,.2f}")
        with col4:
            st.metric("Unique Products", filtered_df[column_mapping["product"]].nunique())
    
    with line_tabs[3]:
        st.write("#### Visualizations")
        
        # Historical trend
        st.write("**Historical Sales Trend**")
        monthly_sales = df_line.groupby([df_line[column_mapping["date"]].dt.to_period("M"), column_mapping["product"]])[column_mapping["value"]].sum().reset_index()
        monthly_sales[column_mapping["date"]] = monthly_sales[column_mapping["date"]].astype(str)
        pivot_monthly = monthly_sales.pivot(index=column_mapping["date"], columns=column_mapping["product"], values=column_mapping["value"]).fillna(0)
        st.line_chart(pivot_monthly)
        
        # Forecast visualization
        st.write("**Forecasted Sales by Quarter**")
        pivot_forecast = forecast_line.pivot_table(
            index=["Year", "Quarter"],
            columns=column_mapping["product"],
            values="ForecastSales",
            fill_value=0
        )
        st.bar_chart(pivot_forecast)
        
        # Growth rate comparison
        st.write("**Growth Rates by Product**")
        growth_chart_data = forecast_line.groupby(column_mapping["product"])["GrowthRate"].first().sort_values(ascending=False)
        st.bar_chart(growth_chart_data)


def main() -> None:
    """Main entry point for the Streamlit app."""
    st.set_page_config(page_title="Sales Forecasting Dashboard", layout="wide")
    st.title("📈 Sales Forecasting Dashboard")
    st.write(
        "This dashboard provides both aggregate and line-level sales forecasting capabilities. "
        "Upload your data on the left, configure the columns, and explore the results below."
    )

    # Sidebar for data upload and configuration
    data_frames = configure_sidebar()

    # Combined dataset selection
    if data_frames:
        dataset_label = st.sidebar.selectbox(
            "Select dataset for forecasting", options=list(data_frames.keys())
        )
        df_raw = data_frames[dataset_label]
    else:
        st.info("Please upload at least one dataset to begin.")
        df_raw = pd.DataFrame()

    # Column mapping for the selected dataset
    column_mapping = None
    if not df_raw.empty:
        column_mapping = select_columns(df_raw)

    # Parameter controls
    st.sidebar.subheader("Forecast Parameters")
    last_n_months = st.sidebar.number_input(
        "Number of past months to include", min_value=3, max_value=36, value=12, step=1
    )
    growth_adjustment = st.sidebar.slider(
        "Adjust forecast growth (multiplier)", min_value=0.0, max_value=2.0, value=1.0, step=0.05
    )
    custom_year = st.sidebar.number_input(
        "Forecast year (optional)", min_value=2020, max_value=2100, value=0, step=1,
        help="Leave as 0 to use the next year after the last year in your data."
    )

    # Process and display results when configuration is complete
    if df_raw.empty or column_mapping is None:
        st.stop()

    date_col = column_mapping["date"]
    product_col = column_mapping["product"]
    value_col = column_mapping["value"]

    # Prepare and aggregate data for aggregate view
    df_clean = prepare_dataframe(
        df_raw,
        date_col=date_col,
        product_col=product_col,
        value_col=value_col,
        last_n_months=last_n_months,
    )
    aggregated = aggregate_sales(
        df_clean,
        group_cols=["Product", "Year", "Quarter"],
        value_col=value_col,
    )

    # Compute growth rates
    growth_rates = compute_growth_rates(aggregated)

    # Forecast year determination
    forecast_year = None if custom_year == 0 else int(custom_year)
    forecast_df = forecast_sales(
        aggregated,
        growth_rates=growth_rates,
        forecast_year=forecast_year,
    )

    # Apply growth adjustment multiplier
    if growth_adjustment != 1.0:
        forecast_df["ForecastSales"] *= growth_adjustment

    # Layout: use tabs to separate views
    tabs = st.tabs([
        "Historical Summary",
        "Aggregate Forecast",
        "Line-Level Forecast",
        "Growth Rates",
        "Assistant"
    ])

    # Historical Summary tab
    with tabs[0]:
        st.subheader("Historical Sales Summary")
        st.write(
            "The table below shows aggregated sales by product and quarter. Use the sidebar to "
            "change how much historical data is included."
        )
        st.dataframe(aggregated)
        st.write("#### Historical Sales by Quarter")
        pivot_hist = aggregated.pivot_table(
            index=["Year", "Quarter"],
            columns="Product",
            values="Sales",
            fill_value=0,
        )
        st.bar_chart(pivot_hist)

    # Aggregate Forecast tab
    with tabs[1]:
        st.subheader("Aggregate Forecast Results")
        st.write(
            "Forecasted sales for next year by product and quarter. Adjust the growth multiplier "
            "in the sidebar to reflect optimistic or conservative scenarios."
        )
        st.dataframe(forecast_df.rename(columns={"ForecastSales": "Forecast"}))
        st.write("#### Forecast by Quarter and Product")
        pivot_forecast = forecast_df.pivot_table(
            index=["Year", "Quarter"],
            columns="Product",
            values="ForecastSales",
            fill_value=0,
        )
        st.bar_chart(pivot_forecast)

    # Line-Level Forecast tab
    with tabs[2]:
        render_line_level_tab(df_raw, column_mapping)

    # Growth Rates tab
    with tabs[3]:
        st.subheader("Growth Rates")
        st.write(
            "Average quarter-over-quarter growth rates computed from your historical data. "
            "These rates are used to generate the forecast unless you specify a custom multiplier."
        )
        st.dataframe(growth_rates.rename(columns={"AvgGrowthRate": "Average Growth Rate"}))
        st.write("#### Growth Rate by Product and Quarter")
        pivot_growth = growth_rates.pivot_table(
            index="Quarter", columns="Product", values="AvgGrowthRate", fill_value=0
        )
        st.bar_chart(pivot_growth)

    # Assistant tab
    with tabs[4]:
        st.subheader("Assistant")
        st.write(
            "Ask the assistant any questions about your data, the forecast, or the assumptions used. "
            "The assistant uses the OpenAI API and requires a valid API key set as ``OPENAI_API_KEY``."
        )
        if "assistant_history" not in st.session_state:
            st.session_state.assistant_history = []
        
        for entry in st.session_state.assistant_history:
            role = entry["role"]
            content = entry["content"]
            if role == "user":
                st.markdown(f"**You:** {content}")
            else:
                st.markdown(f"**Assistant:** {content}")
        
        question = st.text_input("Enter your question", key="assistant_input")
        if st.button("Ask", key="ask_button") and question:
            context_parts = []
            growth_summary = growth_rates.copy()
            growth_summary["AvgGrowthRate"] = (growth_summary["AvgGrowthRate"] * 100).round(2).astype(str) + "%"
            context_parts.append("Growth Rates:\n" + growth_summary.to_csv(index=False))
            
            forecast_totals = forecast_df.groupby("Product")["ForecastSales"].sum().reset_index()
            forecast_totals["ForecastSales"] = forecast_totals["ForecastSales"].round(2)
            context_parts.append("Forecast Totals:\n" + forecast_totals.to_csv(index=False))
            
            context = "\n\n".join(context_parts)
            try:
                answer = ask_ai(question, context=context)
            except Exception as e:
                answer = str(e)
            
            st.session_state.assistant_history.append({"role": "user", "content": question})
            st.session_state.assistant_history.append({"role": "assistant", "content": answer})
            st.experimental_rerun()


if __name__ == "__main__":
    main()
