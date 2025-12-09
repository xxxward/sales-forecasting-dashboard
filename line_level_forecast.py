"""
Customer Order Planning Tool
----------------------------
Simplified dashboard for sales reps to lookup customer orders and generate forecasts
"""

import pandas as pd
import streamlit as st
from datetime import datetime
import io

def load_data(uploaded_file):
    """Load the NetSuite Invoice Line Item report"""
    if uploaded_file is None:
        return pd.DataFrame()
    
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return pd.DataFrame()

def clean_numeric(value):
    """Clean currency and numeric values"""
    if pd.isna(value):
        return 0
    if isinstance(value, (int, float)):
        return float(value)
    # Remove currency symbols and commas
    cleaned = str(value).replace('$', '').replace(',', '').strip()
    try:
        return float(cleaned)
    except:
        return 0

def calculate_order_cadence(dates):
    """Calculate average days between orders"""
    if len(dates) < 2:
        return None
    
    sorted_dates = sorted(dates)
    gaps = [(sorted_dates[i+1] - sorted_dates[i]).days for i in range(len(sorted_dates)-1)]
    return sum(gaps) / len(gaps) if gaps else None

def generate_2026_forecast(customer_df, date_col, item_col, qty_col, amount_col):
    """Generate 2026 forecast based on historical patterns"""
    
    # Calculate monthly averages by product
    customer_df['Month'] = pd.to_datetime(customer_df[date_col]).dt.to_period('M')
    
    monthly_by_product = customer_df.groupby([item_col, 'Month']).agg({
        qty_col: 'sum',
        amount_col: 'sum'
    }).reset_index()
    
    # Calculate average monthly quantity and amount per product
    product_averages = monthly_by_product.groupby(item_col).agg({
        qty_col: 'mean',
        amount_col: 'mean'
    }).reset_index()
    
    product_averages.columns = ['Product', 'Avg_Monthly_Qty', 'Avg_Monthly_Amount']
    
    # Generate 2026 forecast (12 months)
    months_2026 = pd.date_range('2026-01-01', '2026-12-01', freq='MS')
    forecast_rows = []
    
    for product, avg_qty, avg_amount in product_averages.itertuples(index=False):
        for month in months_2026:
            forecast_rows.append({
                'Month': month.strftime('%B %Y'),
                'Quarter': f"Q{month.quarter}",
                'Product': product,
                'Forecasted_Qty': round(avg_qty, 0),
                'Forecasted_Amount': round(avg_amount, 2),
                'Notes': ''
            })
    
    forecast_df = pd.DataFrame(forecast_rows)
    return forecast_df

def main():
    st.set_page_config(page_title="Customer Order Planning Tool", layout="wide")
    
    st.title("📊 Customer Order Planning Tool")
    st.write("Look up customer order history and generate 2026 forecasts")
    
    # Sidebar
    st.sidebar.header("Data Upload")
    uploaded_file = st.sidebar.file_uploader(
        "Upload NetSuite Invoice Line Item Report",
        type=['csv', 'xlsx', 'xls']
    )
    
    if uploaded_file is None:
        st.info("👈 Please upload your NetSuite Invoice Line Item report to begin")
        st.stop()
    
    # Load data
    df = load_data(uploaded_file)
    
    if df.empty:
        st.error("Could not load data from file")
        st.stop()
    
    # Display available columns for user to map
    st.sidebar.subheader("Column Mapping")
    st.sidebar.write("Map your NetSuite columns:")
    
    cols = df.columns.tolist()
    
    customer_col = st.sidebar.selectbox("Customer Name Column", cols, index=0)
    date_col = st.sidebar.selectbox("Date Column", cols, index=1)
    item_col = st.sidebar.selectbox("Item/SKU Column (Column H)", cols, index=min(7, len(cols)-1))
    product_type_col = st.sidebar.selectbox("Product Type Column (Column X)", cols, index=min(23, len(cols)-1))
    qty_col = st.sidebar.selectbox("Quantity Column (Column I)", cols, index=min(8, len(cols)-1))
    amount_col = st.sidebar.selectbox("Transaction Total Column (Column N)", cols, index=min(13, len(cols)-1))
    
    # Optional order number
    order_col = st.sidebar.selectbox("Order Number Column (Optional)", ["None"] + cols)
    order_col = None if order_col == "None" else order_col
    
    # Clean data
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df[qty_col] = df[qty_col].apply(clean_numeric)
    df[amount_col] = df[amount_col].apply(clean_numeric)
    
    # Remove rows with invalid data
    df = df.dropna(subset=[customer_col, date_col])
    df = df[df[qty_col] > 0]
    
    # Customer selection
    st.sidebar.subheader("Select Customer")
    customers = sorted(df[customer_col].unique())
    selected_customer = st.sidebar.selectbox("Customer", customers)
    
    # Filter for selected customer
    customer_df = df[df[customer_col] == selected_customer].copy()
    customer_df = customer_df.sort_values(date_col, ascending=False)
    
    # Main content
    st.header(f"📈 {selected_customer}")
    
    # ===== SECTION 1: Order History =====
    st.subheader("Order History")
    
    # Prepare display columns
    display_cols = [date_col, item_col, product_type_col, qty_col, amount_col]
    if order_col:
        display_cols.insert(0, order_col)
    
    # Format the display dataframe
    history_display = customer_df[display_cols].copy()
    history_display[amount_col] = history_display[amount_col].apply(lambda x: f"${x:,.2f}")
    
    st.dataframe(history_display, use_container_width=True, height=400)
    
    # Download button for history
    csv_history = customer_df[display_cols].to_csv(index=False)
    st.download_button(
        label="📥 Download Order History",
        data=csv_history,
        file_name=f"{selected_customer}_order_history.csv",
        mime="text/csv"
    )
    
    st.divider()
    
    # ===== SECTION 2: Order Pattern Analysis =====
    st.subheader("Order Pattern Summary")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_orders = len(customer_df)
        st.metric("Total Orders", f"{total_orders:,}")
    
    with col2:
        total_units = customer_df[qty_col].sum()
        st.metric("Total Units", f"{total_units:,.0f}")
    
    with col3:
        total_revenue = customer_df[amount_col].sum()
        st.metric("Total Revenue", f"${total_revenue:,.2f}")
    
    with col4:
        avg_order_value = total_revenue / total_orders if total_orders > 0 else 0
        st.metric("Avg Order Value", f"${avg_order_value:,.2f}")
    
    col5, col6, col7 = st.columns(3)
    
    with col5:
        first_order = customer_df[date_col].min()
        st.metric("First Order", first_order.strftime('%Y-%m-%d') if pd.notna(first_order) else "N/A")
    
    with col6:
        last_order = customer_df[date_col].max()
        st.metric("Last Order", last_order.strftime('%Y-%m-%d') if pd.notna(last_order) else "N/A")
    
    with col7:
        cadence = calculate_order_cadence(customer_df[date_col].tolist())
        if cadence:
            st.metric("Order Cadence", f"Every {cadence:.0f} days")
        else:
            st.metric("Order Cadence", "N/A")
    
    # Top products
    st.write("**Top 5 Products by Quantity**")
    top_products = customer_df.groupby(item_col)[qty_col].sum().sort_values(ascending=False).head(5)
    top_products_df = pd.DataFrame({
        'Product': top_products.index,
        'Total Quantity': top_products.values
    })
    st.dataframe(top_products_df, use_container_width=True, hide_index=True)
    
    st.divider()
    
    # ===== SECTION 3: 2026 Forecast =====
    st.subheader("2026 Forecast")
    st.write("Based on historical order patterns, here's a projected forecast for 2026. You can edit the quantities before exporting.")
    
    # Generate forecast
    forecast_df = generate_2026_forecast(customer_df, date_col, item_col, qty_col, amount_col)
    
    # Create editable data editor
    st.write("**Edit forecast quantities as needed:**")
    
    edited_forecast = st.data_editor(
        forecast_df,
        use_container_width=True,
        height=400,
        column_config={
            "Forecasted_Qty": st.column_config.NumberColumn(
                "Forecasted Qty",
                help="Edit to adjust forecast",
                min_value=0,
                step=1,
                format="%d"
            ),
            "Forecasted_Amount": st.column_config.NumberColumn(
                "Forecasted Amount",
                help="Edit to adjust forecast",
                min_value=0,
                format="$%.2f"
            ),
            "Notes": st.column_config.TextColumn(
                "Notes",
                help="Add notes or adjustments"
            )
        }
    )
    
    # Summary metrics for forecast
    st.write("**2026 Forecast Summary**")
    fcol1, fcol2, fcol3 = st.columns(3)
    
    with fcol1:
        forecast_total_qty = edited_forecast['Forecasted_Qty'].sum()
        st.metric("Total Forecasted Units", f"{forecast_total_qty:,.0f}")
    
    with fcol2:
        forecast_total_revenue = edited_forecast['Forecasted_Amount'].sum()
        st.metric("Total Forecasted Revenue", f"${forecast_total_revenue:,.2f}")
    
    with fcol3:
        avg_monthly_revenue = forecast_total_revenue / 12
        st.metric("Avg Monthly Revenue", f"${avg_monthly_revenue:,.2f}")
    
    # Export options
    st.write("**Export Options**")
    
    col_export1, col_export2 = st.columns(2)
    
    with col_export1:
        # Export forecast only
        csv_forecast = edited_forecast.to_csv(index=False)
        st.download_button(
            label="📥 Download 2026 Forecast",
            data=csv_forecast,
            file_name=f"{selected_customer}_2026_forecast.csv",
            mime="text/csv"
        )
    
    with col_export2:
        # Export complete customer summary
        summary_buffer = io.StringIO()
        summary_buffer.write(f"CUSTOMER ORDER SUMMARY AND 2026 FORECAST\n")
        summary_buffer.write(f"Customer: {selected_customer}\n")
        summary_buffer.write(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
        summary_buffer.write(f"\n")
        summary_buffer.write(f"HISTORICAL SUMMARY\n")
        summary_buffer.write(f"Total Orders: {total_orders}\n")
        summary_buffer.write(f"Total Units: {total_units:,.0f}\n")
        summary_buffer.write(f"Total Revenue: ${total_revenue:,.2f}\n")
        summary_buffer.write(f"Average Order Value: ${avg_order_value:,.2f}\n")
        summary_buffer.write(f"Order Cadence: Every {cadence:.0f} days\n" if cadence else "Order Cadence: N/A\n")
        summary_buffer.write(f"\n")
        summary_buffer.write(f"2026 FORECAST\n")
        summary_buffer.write(edited_forecast.to_csv(index=False))
        
        summary_export = summary_buffer.getvalue()
        
        st.download_button(
            label="📥 Download Complete Summary",
            data=summary_export,
            file_name=f"{selected_customer}_complete_summary.csv",
            mime="text/csv"
        )

if __name__ == "__main__":
    main()
