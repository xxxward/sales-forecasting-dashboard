"""
Sales Forecasting Dashboard - Enhanced Version with Drill-Down Capability
THIS FILE CONTAINS THE UPDATED build_your_own_forecast_section FUNCTION
with enhanced drill-down visibility for Sales Orders and HubSpot Deals

Key changes:
1. Added drill-down expanders showing SO#, Deal ID, and clickable links
2. Displays detailed tables with all relevant information
3. Maintains all existing functionality while adding better visibility
"""

# This is the UPDATED version of the build_your_own_forecast_section function
# Replace the existing function (starting at line 1389) with this version

def build_your_own_forecast_section(metrics, quota, rep_name=None, deals_df=None, invoices_df=None, sales_orders_df=None):
    """
    Interactive section where users can select which data sources to include in their forecast
    NOW WITH ENHANCED DRILL-DOWN CAPABILITY showing SO#, Deal ID, and links
    """
    st.markdown("### 🎯 Build Your Own Forecast")
    st.caption("Select the components you want to include in your custom forecast calculation")
    
    # Initialize session state for individual selections if not exists
    if 'selected_individual_items' not in st.session_state:
        st.session_state.selected_individual_items = {}
    
    # Create columns for checkboxes
    col1, col2, col3 = st.columns(3)
    
    # Available data sources with their values
    sources = {
        'Invoiced & Shipped': metrics.get('orders', 0),
        'Pending Fulfillment (with date)': metrics.get('pending_fulfillment', 0),
        'Pending Approval (with date)': metrics.get('pending_approval', 0),
        'HubSpot Expect': metrics.get('expect_commit', 0) if 'expect_commit' in metrics else 0,
        'HubSpot Commit': 0,  # Will calculate separately
        'HubSpot Best Case': 0,  # Will calculate separately
        'HubSpot Opportunity': 0,  # Will calculate separately
        'Pending Fulfillment (without date)': metrics.get('pending_fulfillment_no_date', 0),
        'Pending Approval (without date)': metrics.get('pending_approval_no_date', 0),
        'Pending Approval (>2 weeks old)': metrics.get('pending_approval_old', 0),
        'Q1 Spillover - Expect/Commit': metrics.get('q1_spillover_expect_commit', 0),
        'Q1 Spillover - Best Case': metrics.get('q1_spillover_best_opp', 0)
    }
    
    # Track which categories allow individual selection
    individual_select_categories = [
        'HubSpot Expect', 'HubSpot Commit', 'HubSpot Best Case', 'HubSpot Opportunity',
        'Pending Fulfillment (without date)', 'Pending Approval (without date)', 
        'Pending Approval (>2 weeks old)', 'Q1 Spillover - Expect/Commit', 'Q1 Spillover - Best Case'
    ]
    
    # Calculate individual HubSpot categories
    if deals_df is not None and not deals_df.empty:
        if rep_name:
            rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
        else:
            rep_deals = deals_df.copy()
        
        if not rep_deals.empty and 'Status' in rep_deals.columns:
            rep_deals['Amount_Numeric'] = pd.to_numeric(rep_deals['Amount'], errors='coerce')
            
            # Filter for Q4 only
            q4_deals = rep_deals[rep_deals.get('Counts_In_Q4', True) == True]
            
            sources['HubSpot Expect'] = q4_deals[q4_deals['Status'] == 'Expect']['Amount_Numeric'].sum()
            sources['HubSpot Commit'] = q4_deals[q4_deals['Status'] == 'Commit']['Amount_Numeric'].sum()
            sources['HubSpot Best Case'] = q4_deals[q4_deals['Status'] == 'Best Case']['Amount_Numeric'].sum()
            sources['HubSpot Opportunity'] = q4_deals[q4_deals['Status'] == 'Opportunity']['Amount_Numeric'].sum()
    
    # Create checkboxes in columns with individual selection option
    selected_sources = {}
    individual_selection_mode = {}
    source_list = list(sources.keys())
    
    with col1:
        for source in source_list[0:4]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            # Add "Select Individual" option for applicable categories
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    with col2:
        for source in source_list[4:8]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    with col3:
        for source in source_list[8:]:
            selected_sources[source] = st.checkbox(
                f"{source}: ${sources[source]:,.0f}",
                value=False,
                key=f"{'team' if rep_name is None else rep_name}_{source}"
            )
            
            if source in individual_select_categories and selected_sources[source]:
                individual_selection_mode[source] = st.checkbox(
                    f"   ↳ Select individual items",
                    value=False,
                    key=f"{'team' if rep_name is None else rep_name}_{source}_individual"
                )
    
    # Show individual selection interfaces for each category
    individual_selections = {}
    
    for category, is_individual in individual_selection_mode.items():
        if is_individual:
            st.markdown(f"#### 🛒 Select Individual Items: {category}")
            
            # Get the relevant data for this category
            items_to_select = pd.DataFrame()
            
            # Sales Orders categories
            if 'Pending Fulfillment (without date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Fulfillment') &
                    (so_data['Customer Promise Date'].isna()) &
                    (so_data['Projected Date'].isna())
                ].copy()
                
            elif 'Pending Approval (without date)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                items_to_select = so_data[
                    (so_data['Status'] == 'Pending Approval') &
                    (so_data['Customer Promise Date'].isna()) &
                    (so_data['Projected Date'].isna())
                ].copy()
                
            elif 'Pending Approval (>2 weeks old)' in category and sales_orders_df is not None:
                if rep_name and 'Sales Rep' in sales_orders_df.columns:
                    so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
                else:
                    so_data = sales_orders_df.copy()
                
                if 'Age_Business_Days' in so_data.columns:
                    items_to_select = so_data[
                        (so_data['Status'] == 'Pending Approval') &
                        (so_data['Age_Business_Days'] >= 10)
                    ].copy()
                    
            # HubSpot deals categories
            elif 'HubSpot' in category and deals_df is not None:
                if rep_name:
                    hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
                else:
                    hs_deals = deals_df.copy()
                
                if not hs_deals.empty and 'Status' in hs_deals.columns:
                    hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                    q4_deals = hs_deals[hs_deals.get('Counts_In_Q4', True) == True]
                    
                    if 'Expect' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Expect'].copy()
                    elif 'Commit' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Commit'].copy()
                    elif 'Best Case' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Best Case'].copy()
                    elif 'Opportunity' in category:
                        items_to_select = q4_deals[q4_deals['Status'] == 'Opportunity'].copy()
            
            # Q1 Spillover deals
            elif 'Q1 Spillover' in category and deals_df is not None:
                if rep_name:
                    hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
                else:
                    hs_deals = deals_df.copy()
                
                if not hs_deals.empty and 'Status' in hs_deals.columns:
                    hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                    
                    # Determine status filter
                    if 'Expect/Commit' in category:
                        status_filter = ['Expect', 'Commit']
                    else:
                        status_filter = ['Best Case', 'Opportunity']
                    
                    # Use Q1 2026 Spillover column
                    if 'Q1 2026 Spillover' in hs_deals.columns:
                        items_to_select = hs_deals[
                            (hs_deals['Q1 2026 Spillover'] == 'Q1 2026') &
                            (hs_deals['Status'].isin(status_filter))
                        ].copy()
                        
                        # Debug info
                        total_spillover = hs_deals[hs_deals['Q1 2026 Spillover'] == 'Q1 2026']
                        st.caption(f"🔍 Debug: Total Q1 spillover deals = {len(total_spillover)}, {'/'.join(status_filter)} only = {len(items_to_select)}")
                        st.caption(f"Total amount in Q1 spillover {'/'.join(status_filter)} = ${items_to_select['Amount_Numeric'].sum():,.0f}")
                    else:
                        # Fallback to old logic if column doesn't exist
                        items_to_select = hs_deals[
                            (hs_deals.get('Counts_In_Q4', True) == False) &
                            (hs_deals['Status'].isin(status_filter))
                        ].copy()
                        st.caption("⚠️ Using fallback logic - Q1 2026 Spillover column not found")
            
            # ================ NEW DRILL-DOWN SECTION ================
            # Display the data with proper drill-down view BEFORE individual selection
            if not items_to_select.empty:
                st.caption(f"📊 **Viewing {len(items_to_select)} items totaling ${items_to_select['Amount' if 'Amount' in items_to_select.columns else 'Amount_Numeric'].sum():,.0f}**")
                
                # Create drill-down expander showing all details
                with st.expander(f"👀 View Detailed Breakdown (click to expand)", expanded=True):
                    # Determine if this is HubSpot or NetSuite data
                    is_hubspot = 'Deal Name' in items_to_select.columns
                    is_netsuite = 'Document Number' in items_to_select.columns or 'Internal ID' in items_to_select.columns
                    
                    # Create display dataframe with links
                    display_df = pd.DataFrame()
                    column_config = {}
                    
                    if is_hubspot and 'Record ID' in items_to_select.columns:
                        # HubSpot deals - show links and details
                        display_df['🔗 Link'] = items_to_select['Record ID'].apply(
                            lambda x: f'https://app.hubspot.com/contacts/6712259/record/0-3/{x}/' if pd.notna(x) else ''
                        )
                        column_config['🔗 Link'] = st.column_config.LinkColumn(
                            "🔗 Link",
                            help="Click to view deal in HubSpot",
                            display_text="View Deal"
                        )
                        
                        # Add Deal ID
                        display_df['Deal ID'] = items_to_select['Record ID']
                        
                        # Add other HubSpot columns
                        if 'Deal Name' in items_to_select.columns:
                            display_df['Deal Name'] = items_to_select['Deal Name']
                        if 'Account Name' in items_to_select.columns:
                            display_df['Customer'] = items_to_select['Account Name']
                        if 'Amount_Numeric' in items_to_select.columns:
                            display_df['Amount'] = items_to_select['Amount_Numeric'].apply(lambda x: f"${x:,.2f}")
                        elif 'Amount' in items_to_select.columns:
                            display_df['Amount'] = pd.to_numeric(items_to_select['Amount'], errors='coerce').apply(lambda x: f"${x:,.2f}")
                        if 'Status' in items_to_select.columns:
                            display_df['Status'] = items_to_select['Status']
                        if 'Pipeline' in items_to_select.columns:
                            display_df['Pipeline'] = items_to_select['Pipeline']
                        if 'Close Date' in items_to_select.columns:
                            if pd.api.types.is_datetime64_any_dtype(items_to_select['Close Date']):
                                display_df['Close Date'] = items_to_select['Close Date'].dt.strftime('%Y-%m-%d')
                            else:
                                display_df['Close Date'] = items_to_select['Close Date']
                        if 'Product Type' in items_to_select.columns:
                            display_df['Product Type'] = items_to_select['Product Type']
                    
                    elif is_netsuite:
                        # NetSuite sales orders - show SO# and links
                        if 'Internal ID' in items_to_select.columns:
                            display_df['🔗 Link'] = items_to_select['Internal ID'].apply(
                                lambda x: f'https://7086864.app.netsuite.com/app/accounting/transactions/salesord.nl?id={x}&whence=' if pd.notna(x) else ''
                            )
                            column_config['🔗 Link'] = st.column_config.LinkColumn(
                                "🔗 Link",
                                help="Click to view sales order in NetSuite",
                                display_text="View SO"
                            )
                            # Also show Internal ID as a regular column
                            display_df['Internal ID'] = items_to_select['Internal ID']
                        
                        # Add SO# (Document Number)
                        if 'Document Number' in items_to_select.columns:
                            display_df['SO#'] = items_to_select['Document Number']
                        
                        # Add other NetSuite columns
                        if 'Customer' in items_to_select.columns:
                            display_df['Customer'] = items_to_select['Customer']
                        if 'Amount' in items_to_select.columns:
                            display_df['Amount'] = pd.to_numeric(items_to_select['Amount'], errors='coerce').apply(lambda x: f"${x:,.2f}")
                        if 'Status' in items_to_select.columns:
                            display_df['Status'] = items_to_select['Status']
                        if 'Order Start Date' in items_to_select.columns:
                            if pd.api.types.is_datetime64_any_dtype(items_to_select['Order Start Date']):
                                display_df['Order Start Date'] = items_to_select['Order Start Date'].dt.strftime('%Y-%m-%d')
                            else:
                                display_df['Order Start Date'] = items_to_select['Order Start Date']
                        if 'Age_Business_Days' in items_to_select.columns:
                            display_df['Age (Days)'] = items_to_select['Age_Business_Days']
                    
                    # Display the formatted dataframe
                    if not display_df.empty:
                        st.dataframe(
                            display_df,
                            use_container_width=True,
                            hide_index=True,
                            column_config=column_config if column_config else None
                        )
                        
                        # Summary statistics
                        amount_col = 'Amount_Numeric' if 'Amount_Numeric' in items_to_select.columns else 'Amount'
                        total_amount = pd.to_numeric(items_to_select[amount_col], errors='coerce').sum()
                        st.caption(f"💰 **Total: ${total_amount:,.2f} | Count: {len(items_to_select)} items**")
                    else:
                        st.warning("Could not format data for display")
                        st.dataframe(items_to_select, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                st.markdown(f"##### ✅ Select items to include in your forecast:")
                st.caption(f"Check the boxes below to select specific items from the {len(items_to_select)} available")
                
            # ================ END NEW DRILL-DOWN SECTION ================
            
            # Display selection interface (existing code)
            if not items_to_select.empty:
                selected_items = []
                
                # Create a more compact selection interface
                for idx, row in items_to_select.iterrows():
                    # Determine display info based on type
                    if 'Deal Name' in row:
                        item_id = row.get('Record ID', idx)
                        item_name = row.get('Deal Name', 'Unknown')
                        item_customer = row.get('Account Name', '')
                        item_amount = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                        # Create more detailed display with Deal ID
                        display_text = f"Deal #{item_id}: {item_name} - {item_customer} - ${item_amount:,.0f}"
                    else:
                        item_id = row.get('Document Number', idx)
                        internal_id = row.get('Internal ID', '')
                        item_name = f"SO #{item_id}"
                        item_customer = row.get('Customer', '')
                        item_amount = pd.to_numeric(row.get('Amount', 0), errors='coerce')
                        # Create more detailed display with Internal ID
                        display_text = f"{item_name} (ID: {internal_id}) - {item_customer} - ${item_amount:,.0f}"
                    
                    # Checkbox for each item
                    is_selected = st.checkbox(
                        display_text,
                        value=False,
                        key=f"{'team' if rep_name is None else rep_name}_{category}_{item_id}"
                    )
                    
                    if is_selected:
                        selected_items.append({
                            'id': item_id,
                            'amount': item_amount,
                            'row': row
                        })
                
                individual_selections[category] = selected_items
                
                # Show selection summary
                if selected_items:
                    selected_total = sum(item['amount'] for item in selected_items)
                    st.success(f"✓ Selected {len(selected_items)} of {len(items_to_select)} items (${selected_total:,.2f})")
                else:
                    st.caption(f"No items selected yet - select items above to include in your forecast")
            else:
                st.info(f"No items found in this category")
    
    # Calculate custom forecast (unchanged from original)
    custom_forecast = 0
    
    for source, selected in selected_sources.items():
        if selected:
            if individual_selection_mode.get(source, False):
                # Use individual selections
                if source in individual_selections:
                    custom_forecast += sum(item['amount'] for item in individual_selections[source])
            else:
                # Use full category amount
                custom_forecast += sources[source]
    
    gap_to_quota = quota - custom_forecast
    attainment_pct = (custom_forecast / quota * 100) if quota > 0 else 0
    
    # Display results (unchanged from original)
    st.markdown("---")
    st.markdown("#### 📊 Your Custom Forecast")
    
    result_col1, result_col2, result_col3, result_col4 = st.columns(4)
    
    with result_col1:
        st.metric("Quota", f"${quota:,.0f}")
    
    with result_col2:
        st.metric("Custom Forecast", f"${custom_forecast:,.0f}")
    
    with result_col3:
        st.metric("Gap to Quota", f"${gap_to_quota:,.0f}", 
                 delta=f"${-gap_to_quota:,.0f}" if gap_to_quota < 0 else None,
                 delta_color="inverse")
    
    with result_col4:
        st.metric("Attainment", f"{attainment_pct:.1f}%")
    
    # Export functionality (keeping rest of original function - lines 1676-2068 unchanged)
    # [Rest of the export code remains exactly the same as in the original]
    
    if any(selected_sources.values()):
        st.markdown("---")
        
        # Collect data for export with summary
        export_summary = []
        export_data = []
        
        # Build summary section
        export_summary.append({
            'Category': '=== FORECAST SUMMARY ===',
            'Amount': ''
        })
        export_summary.append({
            'Category': 'Quota',
            'Amount': f"${quota:,.0f}"
        })
        export_summary.append({
            'Category': 'Custom Forecast',
            'Amount': f"${custom_forecast:,.0f}"
        })
        export_summary.append({
            'Category': 'Gap to Quota',
            'Amount': f"${gap_to_quota:,.0f}"
        })
        export_summary.append({
            'Category': 'Attainment %',
            'Amount': f"{attainment_pct:.1f}%"
        })
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        export_summary.append({
            'Category': '=== SELECTED COMPONENTS ===',
            'Amount': ''
        })
        
        # Add each selected component total
        for source, selected in selected_sources.items():
            if selected:
                if individual_selection_mode.get(source, False) and source in individual_selections:
                    # Show individual selection count
                    item_count = len(individual_selections[source])
                    item_total = sum(item['amount'] for item in individual_selections[source])
                    export_summary.append({
                        'Category': f"{source} ({item_count} items selected)",
                        'Amount': f"${item_total:,.0f}"
                    })
                else:
                    export_summary.append({
                        'Category': source,
                        'Amount': f"${sources[source]:,.0f}"
                    })
        
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        export_summary.append({
            'Category': '=== DETAILED LINE ITEMS ===',
            'Amount': ''
        })
        export_summary.append({
            'Category': '',
            'Amount': ''
        })
        
        # Get invoices data (always bulk)
        if selected_sources.get('Invoiced & Shipped', False) and invoices_df is not None:
            if rep_name and 'Sales Rep' in invoices_df.columns:
                inv_data = invoices_df[invoices_df['Sales Rep'] == rep_name].copy()
            else:
                inv_data = invoices_df.copy()
            
            if not inv_data.empty:
                for _, row in inv_data.iterrows():
                    export_data.append({
                        'Type': 'Invoice',
                        'ID': row.get('Document Number', row.get('Invoice Number', '')),
                        'Name': '',
                        'Customer': row.get('Account Name', row.get('Customer', '')),
                        'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                        'Date': row.get('Date', row.get('Transaction Date', '')),
                        'Sales Rep': row.get('Sales Rep', '')
                    })
        
        # Get sales orders data - check individual vs bulk
        if sales_orders_df is not None:
            if rep_name and 'Sales Rep' in sales_orders_df.columns:
                so_data = sales_orders_df[sales_orders_df['Sales Rep'] == rep_name].copy()
            else:
                so_data = sales_orders_df.copy()
            
            if not so_data.empty:
                # Pending Fulfillment with date (always bulk)
                if selected_sources.get('Pending Fulfillment (with date)', False):
                    pf_data = so_data[so_data['Status'] == 'Pending Fulfillment'].copy()
                    for _, row in pf_data.iterrows():
                        if pd.notna(row.get('Customer Promise Date')) or pd.notna(row.get('Projected Date')):
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Approval with date (always bulk)
                if selected_sources.get('Pending Approval (with date)', False):
                    pa_data = so_data[so_data['Status'] == 'Pending Approval'].copy()
                    for _, row in pa_data.iterrows():
                        if pd.notna(row.get('Customer Promise Date')) or pd.notna(row.get('Projected Date')):
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Fulfillment without date - check individual mode
                if selected_sources.get('Pending Fulfillment (without date)', False):
                    category = 'Pending Fulfillment (without date)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        pf_no_date = so_data[
                            (so_data['Status'] == 'Pending Fulfillment') &
                            (so_data['Customer Promise Date'].isna()) &
                            (so_data['Projected Date'].isna())
                        ].copy()
                        for _, row in pf_no_date.iterrows():
                            export_data.append({
                                'Type': 'Sales Order - Pending Fulfillment (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Pending Approval without date - check individual mode
                if selected_sources.get('Pending Approval (without date)', False):
                    category = 'Pending Approval (without date)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        pa_no_date = so_data[
                            (so_data['Status'] == 'Pending Approval') &
                            (so_data['Customer Promise Date'].isna()) &
                            (so_data['Projected Date'].isna())
                        ].copy()
                        for _, row in pa_no_date.iterrows():
                            export_data.append({
                                'Type': 'Sales Order - Pending Approval (No Date)',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                
                # Old Pending Approval - check individual mode
                if selected_sources.get('Pending Approval (>2 weeks old)', False):
                    category = 'Pending Approval (>2 weeks old)'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'Sales Order - Old Pending Approval',
                                'ID': row.get('Document Number', ''),
                                'Name': '',
                                'Customer': row.get('Customer', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Order Start Date', ''),
                                'Sales Rep': row.get('Sales Rep', '')
                            })
                    else:
                        # Bulk export
                        if 'Age_Business_Days' in so_data.columns:
                            old_pa = so_data[
                                (so_data['Status'] == 'Pending Approval') &
                                (so_data['Age_Business_Days'] >= 10)
                            ].copy()
                            for _, row in old_pa.iterrows():
                                export_data.append({
                                    'Type': 'Sales Order - Old Pending Approval',
                                    'ID': row.get('Document Number', ''),
                                    'Name': '',
                                    'Customer': row.get('Customer', ''),
                                    'Amount': pd.to_numeric(row.get('Amount', 0), errors='coerce'),
                                    'Date': row.get('Order Start Date', ''),
                                    'Sales Rep': row.get('Sales Rep', '')
                                })
        
        # Get HubSpot deals data - check individual vs bulk
        if deals_df is not None and not deals_df.empty:
            if rep_name:
                hs_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
            else:
                hs_deals = deals_df.copy()
            
            if not hs_deals.empty and 'Status' in hs_deals.columns:
                hs_deals['Amount_Numeric'] = pd.to_numeric(hs_deals['Amount'], errors='coerce')
                
                # Filter for selected categories with individual selection support
                for status_name, checkbox_name in [
                    ('Expect', 'HubSpot Expect'),
                    ('Commit', 'HubSpot Commit'),
                    ('Best Case', 'HubSpot Best Case'),
                    ('Opportunity', 'HubSpot Opportunity')
                ]:
                    if selected_sources.get(checkbox_name, False):
                        if individual_selection_mode.get(checkbox_name, False) and checkbox_name in individual_selections:
                            # Use individual selections
                            for item in individual_selections[checkbox_name]:
                                row = item['row']
                                export_data.append({
                                    'Type': f'HubSpot Deal - {status_name}',
                                    'ID': row.get('Record ID', ''),
                                    'Name': row.get('Deal Name', ''),
                                    'Customer': row.get('Account Name', ''),
                                    'Amount': item['amount'],
                                    'Date': row.get('Close Date', ''),
                                    'Sales Rep': row.get('Deal Owner', '')
                                })
                        else:
                            # Bulk export
                            status_deals = hs_deals[hs_deals['Status'] == status_name].copy()
                            for _, row in status_deals.iterrows():
                                export_data.append({
                                    'Type': f'HubSpot Deal - {status_name}',
                                    'ID': row.get('Record ID', ''),
                                    'Name': row.get('Deal Name', ''),
                                    'Customer': row.get('Account Name', ''),
                                    'Amount': row.get('Amount_Numeric', 0),
                                    'Date': row.get('Close Date', ''),
                                    'Sales Rep': row.get('Deal Owner', '')
                                })
                
                # Q1 Spillover - check individual mode
                if selected_sources.get('Q1 Spillover - Expect/Commit', False):
                    category = 'Q1 Spillover - Expect/Commit'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover (E/C)',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
                    else:
                        # Bulk export
                        q1_ec = hs_deals[
                            (hs_deals.get('Q1 2026 Spillover', '') == 'Q1 2026') &
                            (hs_deals['Status'].isin(['Expect', 'Commit']))
                        ].copy()
                        for _, row in q1_ec.iterrows():
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover (E/C)',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': row.get('Amount_Numeric', 0),
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
                
                # Q1 Spillover Best Case
                if selected_sources.get('Q1 Spillover - Best Case', False):
                    category = 'Q1 Spillover - Best Case'
                    if individual_selection_mode.get(category, False) and category in individual_selections:
                        # Use individual selections
                        for item in individual_selections[category]:
                            row = item['row']
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover (BC)',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': item['amount'],
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
                    else:
                        # Bulk export
                        q1_bc = hs_deals[
                            (hs_deals.get('Q1 2026 Spillover', '') == 'Q1 2026') &
                            (hs_deals['Status'].isin(['Best Case', 'Opportunity']))
                        ].copy()
                        for _, row in q1_bc.iterrows():
                            export_data.append({
                                'Type': 'HubSpot Deal - Q1 Spillover (BC)',
                                'ID': row.get('Record ID', ''),
                                'Name': row.get('Deal Name', ''),
                                'Customer': row.get('Account Name', ''),
                                'Amount': row.get('Amount_Numeric', 0),
                                'Date': row.get('Close Date', ''),
                                'Sales Rep': row.get('Deal Owner', '')
                            })
        
        # Create export dataframes
        if export_data:
            summary_df = pd.DataFrame(export_summary)
            detail_df = pd.DataFrame(export_data)
            
            # Sort detail by amount descending
            detail_df = detail_df.sort_values('Amount', ascending=False)
            
            # Create Excel with multiple sheets
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                detail_df.to_excel(writer, sheet_name='Details', index=False)
            
            excel_data = output.getvalue()
            
            # Download button
            st.download_button(
                label="📥 Download Custom Forecast (Excel)",
                data=excel_data,
                file_name=f"custom_forecast_{'team' if rep_name is None else rep_name}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
