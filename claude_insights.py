"""
Claude AI Insights Module for Sales Dashboard
Provides interactive Q&A and automated daily change summaries
"""

import streamlit as st
import pandas as pd
from anthropic import Anthropic
from datetime import datetime, timedelta
import json

def initialize_claude():
    """Initialize Claude API client"""
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
        client = Anthropic(api_key=api_key)
        return client
    except Exception as e:
        st.error(f"Error initializing Claude: {str(e)}")
        return None

def get_pipeline_summary(deals_df, rep_name=None):
    """Generate a detailed summary of pipeline data for Claude context"""
    if rep_name:
        rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    else:
        rep_deals = deals_df.copy()
    
    if rep_deals.empty:
        return "No deals in pipeline."
    
    # Convert DataFrame to CSV string with all relevant columns
    # This gives Claude access to ALL the data in a format it can analyze
    important_columns = [
        'Deal Name', 'Amount', 'Deal Stage', 'Close Date', 'Deal Owner',
        'Account Name', 'Pipeline', 'Lead Time (Days)', 'Est. Fulfillment Date',
        'Last Modified Date', 'Create Date', 'Forecast Category'
    ]
    
    # Only include columns that exist in the DataFrame
    available_columns = [col for col in important_columns if col in rep_deals.columns]
    
    if available_columns:
        summary_df = rep_deals[available_columns].copy()
        # Convert to CSV string for Claude
        csv_data = summary_df.to_csv(index=False)
        
        # Add a header explaining the data
        data_summary = f"""
PIPELINE DATA FOR {'ALL REPS' if not rep_name else rep_name}:

Total Deals: {len(rep_deals)}
Total Pipeline Value: ${rep_deals['Amount'].sum() if 'Amount' in rep_deals.columns else 0:,.0f}

DETAILED DEAL DATA (CSV format):
{csv_data}

This CSV contains all deal information. You can analyze:
- Deal stages (Pending Fulfillment, Pending Approval, Closed Won, etc.)
- Deal amounts and timing
- Customer names
- Lead times and fulfillment dates
- Any other fields present in the data
"""
        return data_summary
    else:
        return "No deal data available."

def ask_claude(question, context_data, rep_name=None):
    """Send a question to Claude with pipeline context"""
    client = initialize_claude()
    if not client:
        return "Unable to connect to Claude AI. Please check your API key."
    
    # Build the context message
    if rep_name:
        context_intro = f"You are analyzing sales pipeline data for {rep_name}."
    else:
        context_intro = "You are analyzing team-wide sales pipeline data."
    
    system_message = f"""{context_intro}

You are a helpful sales operations analyst with access to complete pipeline data from HubSpot and NetSuite.

{context_data}

When answering questions:
1. Analyze the CSV data provided above to find specific deals
2. Be specific - cite deal names, amounts, and stages
3. For questions about "Pending Fulfillment" or "Pending Approval", look at the 'Deal Stage' column
4. Format your response clearly:
   - Use markdown headers (## and ###) to organize sections
   - Use bullet points (- or *) for lists of deals
   - Use tables when comparing multiple items
   - Bold important numbers with **$XX,XXX**
5. Always include dollar amounts when discussing deals
6. If asked about specific deal stages, filter the data by the 'Deal Stage' column and list ALL matching deals
7. Keep responses scannable - executives should be able to skim and get the key info

The CSV data contains all the information you need to answer questions accurately."""

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[
                {"role": "user", "content": question}
            ],
            system=system_message
        )
        
        return message.content[0].text
    
    except Exception as e:
        return f"Error getting response from Claude: {str(e)}"

def generate_daily_summary(deals_df, dashboard_df, team_metrics=None):
    """Generate automated daily change summary for executives"""
    client = initialize_claude()
    if not client:
        return "Unable to generate daily summary. Please check your API key."
    
    # Use pre-calculated team metrics if provided, otherwise fallback to manual calculation
    if team_metrics:
        total_goal = team_metrics['total_quota']
        total_booked = team_metrics['total_orders']
        total_pending = team_metrics['expect_commit']
        total_committed = total_booked + total_pending
        gap_to_goal = team_metrics['gap']
        percent_to_goal = team_metrics['attainment_pct']
        q1_spillover = team_metrics['q1_spillover']
        best_opp = team_metrics['best_opp']
        potential_total = total_committed + best_opp
        potential_attainment = team_metrics['potential_attainment']
    else:
        # Fallback to manual calculation if no team_metrics provided
        total_goal = 0
        total_booked = 0
        total_pending = 0
        total_committed = 0
        gap_to_goal = 0
        percent_to_goal = 0
        q1_spillover = 0
        best_opp = 0
        potential_total = 0
        potential_attainment = 0
        
        if not dashboard_df.empty:
            # Find columns dynamically
            goal_cols = [col for col in dashboard_df.columns if 'quota' in col.lower()]
            if goal_cols:
                total_goal = dashboard_df[goal_cols[0]].sum()
            
            booked_cols = [col for col in dashboard_df.columns if 'orders' in col.lower()]
            if booked_cols:
                total_booked = dashboard_df[booked_cols[0]].sum()
                
            total_committed = total_booked + total_pending
            gap_to_goal = total_goal - total_committed
            percent_to_goal = (total_committed / total_goal * 100) if total_goal > 0 else 0
    
    # Get deal stage breakdown
    stage_counts = {}
    stage_values = {}
    if 'Deal Stage' in deals_df.columns:
        for stage in deals_df['Deal Stage'].unique():
            stage_deals = deals_df[deals_df['Deal Stage'] == stage]
            stage_counts[stage] = len(stage_deals)
            stage_values[stage] = stage_deals['Amount'].sum() if 'Amount' in deals_df.columns else 0
    
    # Get deals closing this week/month
    deals_this_week = pd.DataFrame()
    deals_this_month = pd.DataFrame()
    at_risk_deals = pd.DataFrame()
    
    if 'Close Date' in deals_df.columns:
        deals_df_temp = deals_df.copy()
        deals_df_temp['Close Date'] = pd.to_datetime(deals_df_temp['Close Date'], errors='coerce')
        deals_this_week = deals_df_temp[
            (deals_df_temp['Close Date'] >= datetime.now()) & 
            (deals_df_temp['Close Date'] <= datetime.now() + timedelta(days=7))
        ]
        deals_this_month = deals_df_temp[
            (deals_df_temp['Close Date'] >= datetime.now()) & 
            (deals_df_temp['Close Date'] <= datetime.now() + timedelta(days=30))
        ]
        
        # Get at-risk deals (closing soon but still in early stages)
        if 'Deal Stage' in deals_df_temp.columns:
            at_risk_deals = deals_df_temp[
                (deals_df_temp['Deal Stage'].isin(['Qualification', 'Proposal', 'Negotiation'])) &
                (deals_df_temp['Close Date'] <= datetime.now() + timedelta(days=14))
            ]
    
    # Build rep performance section safely
    rep_performance = "No rep data available"
    if not dashboard_df.empty:
        # Find available columns for rep performance
        rep_cols = ['Rep Name']
        if goal_cols:
            rep_cols.append(goal_cols[0])
        if booked_cols:
            rep_cols.append(booked_cols[0])
        
        # Add gap column if it exists
        gap_cols = [col for col in dashboard_df.columns if 'gap' in col.lower()]
        if gap_cols:
            rep_cols.append(gap_cols[0])
        
        # Only use columns that exist
        available_rep_cols = [col for col in rep_cols if col in dashboard_df.columns]
        if available_rep_cols:
            rep_performance = dashboard_df[available_rep_cols].to_string()
    
    # Build comprehensive context
    context = f"""
Q4 2025 PERFORMANCE SNAPSHOT:
- Q4 Goal: ${total_goal:,.0f}
- Currently Booked (NetSuite Orders): ${total_booked:,.0f}  
- Expected/Committed Pipeline: ${total_pending:,.0f}
- Total Committed: ${total_committed:,.0f}
- Progress to Goal: {percent_to_goal:.1f}%
- Gap Remaining: ${gap_to_goal:,.0f}
- Best Case/Opportunity Pipeline: ${best_opp:,.0f}
- Potential Total (if all close): ${potential_total:,.0f} ({potential_attainment:.1f}% of goal)
- Q1 2026 Spillover: ${q1_spillover:,.0f}

PIPELINE BY STAGE:
{json.dumps(stage_counts, indent=2)}

VALUE BY STAGE:
{json.dumps({k: f"${v:,.0f}" for k, v in stage_values.items()}, indent=2)}

TIMING:
- Deals closing this week: {len(deals_this_week)} deals worth ${deals_this_week['Amount'].sum() if 'Amount' in deals_this_week.columns and not deals_this_week.empty else 0:,.0f}
- Deals closing this month: {len(deals_this_month)} deals worth ${deals_this_month['Amount'].sum() if 'Amount' in deals_this_month.columns and not deals_this_month.empty else 0:,.0f}
- At-risk deals (closing soon, early stage): {len(at_risk_deals)} deals worth ${at_risk_deals['Amount'].sum() if 'Amount' in at_risk_deals.columns and not at_risk_deals.empty else 0:,.0f}

TOP AT-RISK DEALS:
{at_risk_deals[['Deal Name', 'Amount', 'Deal Stage', 'Close Date', 'Deal Owner']].head(10).to_string() if not at_risk_deals.empty and all(col in at_risk_deals.columns for col in ['Deal Name', 'Amount', 'Deal Stage', 'Close Date', 'Deal Owner']) else "None"}

REP PERFORMANCE BREAKDOWN:
{rep_performance}
"""

    system_message = """You are the VP of Sales Operations writing a daily brief for the executive team and sales leadership.

Your tone: Direct, data-driven, no fluff. Like you're briefing your CEO over coffee.

Write a daily summary that answers these questions in this exact structure:

## 🎯 Bottom Line Up Front
One sentence: Are we hitting Q4 goal or not? By how much?

## 📊 Where We Stand
- Current progress vs goal (use percentages and dollars)
- What moved since yesterday/this week (be specific about wins and losses)
- Q1 spillover impact (if significant)

## 🚨 What Needs Attention TODAY
List 2-3 specific actions for sales leadership to take right now:
- Which deals need executive involvement?
- Which reps need support?
- What's about to slip through the cracks?

## 💰 The Math
Break down exactly what we need to close to hit the goal. Be specific about which stage deals need to convert.

Keep it under 400 words. Use bullet points. Be honest about the challenges. This is for decision-makers who need to act, not feel good."""

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1500,
            messages=[
                {"role": "user", "content": f"Generate today's executive sales brief based on this data:\n\n{context}"}
            ],
            system=system_message
        )
        
        return message.content[0].text
    
    except Exception as e:
        return f"Error generating summary: {str(e)}"

def display_insights_dashboard(deals_df, dashboard_df, team_metrics=None):
    """Main function to display the AI Insights dashboard"""
    
    st.markdown("## 🤖 AI-Powered Insights")
    st.markdown("---")
    
    # Tab selection
    tab1, tab2 = st.tabs(["💬 Ask Claude", "📊 Daily Summary"])
    
    with tab1:
        st.markdown("### Ask Questions About Your Pipeline")
        
        # Rep selection (optional - for individual insights)
        col1, col2 = st.columns([2, 1])
        with col1:
            view_level = st.radio(
                "View insights for:",
                ["Entire Team", "Individual Rep"],
                horizontal=True
            )
        
        selected_rep = None
        if view_level == "Individual Rep":
            with col2:
                selected_rep = st.selectbox(
                    "Select Rep:",
                    options=dashboard_df['Rep Name'].tolist() if not dashboard_df.empty else []
                )
        
        st.markdown("---")
        
        # Suggested questions
        st.markdown("**💡 Suggested Questions:**")
        suggestions = [
            "What deals are at risk of not closing this quarter?",
            "Which opportunities should I prioritize this week?",
            "What's the health of my pipeline?",
            "Are there any customers I should follow up with?"
        ]
        
        cols = st.columns(2)
        for idx, suggestion in enumerate(suggestions):
            with cols[idx % 2]:
                if st.button(suggestion, key=f"suggest_{idx}"):
                    st.session_state['suggested_question'] = suggestion
        
        st.markdown("---")
        
        # Question input
        question = st.text_area(
            "Ask your question:",
            value=st.session_state.get('suggested_question', ''),
            height=100,
            placeholder="e.g., What deals should I focus on this week?"
        )
        
        if st.button("🚀 Get Insights", type="primary"):
            if question:
                with st.spinner("Claude is analyzing your pipeline..."):
                    # Get pipeline context
                    context = get_pipeline_summary(deals_df, selected_rep)
                    
                    # Get Claude's response
                    response = ask_claude(question, context, selected_rep)
                    
                    # Display response
                    st.markdown("### 📋 Claude's Analysis:")
                    st.markdown(response)
                    
                    # Clear suggested question
                    if 'suggested_question' in st.session_state:
                        del st.session_state['suggested_question']
            else:
                st.warning("Please enter a question first.")
    
    with tab2:
        st.markdown("### Daily Pipeline Summary")
        st.caption(f"Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
        
        if st.button("🔄 Generate Fresh Summary", type="primary"):
            with st.spinner("Claude is generating your daily summary..."):
                summary = generate_daily_summary(deals_df, dashboard_df, team_metrics)
                st.markdown(summary)
                
                # Store in session state so it persists
                st.session_state['daily_summary'] = summary
                st.session_state['summary_timestamp'] = datetime.now()
        
        # Display stored summary if available
        if 'daily_summary' in st.session_state:
            st.markdown("---")
            st.markdown(st.session_state['daily_summary'])
            if 'summary_timestamp' in st.session_state:
                st.caption(f"Last generated: {st.session_state['summary_timestamp'].strftime('%I:%M %p')}")
