def display_reconciliation_view(deals_df, dashboard_df, sales_orders_df):
    """Show a reconciliation view to compare with boss's numbers"""
    
    st.title("🔍 Forecast Reconciliation with Boss's Numbers")
    
    # Boss's Q4 numbers from the actual screenshot - CORRECTED
    boss_rep_numbers = {
        'Jake Lynch': {
            'invoiced': 518981,  # CORRECTED - was being cut off
            'pending_fulfillment': 291888, 
            'pending_approval': 42002, 
            'hubspot': 350386, 
            'total': 1203256,  # Section 1 total
            'pending_fulfillment_so_no_date': 108306,
            'pending_approval_so_no_date': 2107,
            'old_pending_approval': 33741,
            'total_q4': 1347410  # Final total
        },
        'Dave Borkowski': {
            'invoiced': 223593, 
            'pending_fulfillment': 146068, 
            'pending_approval': 15702, 
            'hubspot': 396043, 
            'total': 781406,
            'pending_fulfillment_so_no_date': 48150,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 81737,
            'total_q4': 911294
        },
        'Alex Gonzalez': {
            'invoiced': 311101, 
            'pending_fulfillment': 190589, 
            'pending_approval': 0, 
            'hubspot': 0, 
            'total': 501691,
            'pending_fulfillment_so_no_date': 3183,
            'pending_approval_so_no_date': 34846,
            'old_pending_approval': 19300,
            'total_q4': 559019
        },
        'Brad Sherman': {
            'invoiced': 107166, 
            'pending_fulfillment': 39759, 
            'pending_approval': 16878, 
            'hubspot': 211062, 
            'total': 374865,
            'pending_fulfillment_so_no_date': 35390,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1006,
            'total_q4': 411262
        },
        'Lance Mitton': {
            'invoiced': 21998, 
            'pending_fulfillment': 0, 
            'pending_approval': 2758, 
            'hubspot': 11000, 
            'total': 35756,
            'pending_fulfillment_so_no_date': 3735,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 60527,
            'total_q4': 100019
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
            'total_q4': 0
        },
        'Shopify ECommerce': {
            'invoiced': 20404, 
            'pending_fulfillment': 1406, 
            'pending_approval': 1174, 
            'hubspot': 0, 
            'total': 22984,
            'pending_fulfillment_so_no_date': 0,
            'pending_approval_so_no_date': 0,
            'old_pending_approval': 1544,
            'total_q4': 24528
        }
    }
    
    # Tab selection for Rep vs Pipeline view
    tab1, tab2 = st.tabs(["By Rep", "By Pipeline"])
    
    with tab1:
        st.markdown("### Section 1: Q4 Gap to Goal")
        
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
            if rep_name in dashboard_df['Rep Name'].values or rep_name in ['House', 'Shopify ECommerce']:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df) if rep_name in dashboard_df['Rep Name'].values else None
                
                if metrics or rep_name in ['House', 'Shopify ECommerce']:
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
        st.markdown("### Section 2: Additional Orders (Can be included)")
        
        additional_data = []
        additional_totals = {
            'pf_no_date_you': 0, 'pf_no_date_boss': 0,
            'pa_no_date_you': 0, 'pa_no_date_boss': 0,
            'old_pa_you': 0, 'old_pa_boss': 0,
            'final_you': 0, 'final_boss': 0
        }
        
        for rep_name in boss_rep_numbers.keys():
            if rep_name in dashboard_df['Rep Name'].values or rep_name in ['House', 'Shopify ECommerce']:
                metrics = calculate_rep_metrics(rep_name, deals_df, dashboard_df, sales_orders_df) if rep_name in dashboard_df['Rep Name'].values else None
                
                if metrics or rep_name in ['House', 'Shopify ECommerce']:
                    boss = boss_rep_numbers[rep_name]
                    
                    # Calculate additional metrics
                    your_pf_no_date = metrics['pending_fulfillment_no_date'] if metrics else 0
                    your_pa_no_date = metrics['pending_approval_no_date'] if metrics else 0
                    your_old_pa = metrics['pending_approval_old'] if metrics else 0
                    your_final = (metrics['total_progress'] if metrics else 0) + your_pf_no_date + your_pa_no_date + your_old_pa
                    
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
    
    with tab2:
        # Pipeline breakdown would go here with same structure
        st.markdown("### Pipeline-Level Comparison")
        st.info("Pipeline breakdown in development - need to map invoices and sales orders to pipelines")
    
    # Summary
    st.markdown("### 📊 Key Insights")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        diff = totals['total_boss'] - totals['total_you']
        st.metric("Section 1 Variance", f"${abs(diff):,.0f}", 
                 delta=f"{'Under' if diff > 0 else 'Over'} by ${abs(diff):,.0f}")
    
    with col2:
        final_diff = additional_totals['final_boss'] - additional_totals['final_you']
        st.metric("Total Q4 Variance", f"${abs(final_diff):,.0f}",
                 delta=f"{'Under' if final_diff > 0 else 'Over'} by ${abs(final_diff):,.0f}")
    
    with col3:
        accuracy = (1 - abs(final_diff) / additional_totals['final_boss']) * 100 if additional_totals['final_boss'] > 0 else 0
        st.metric("Accuracy", f"{accuracy:.1f}%")
