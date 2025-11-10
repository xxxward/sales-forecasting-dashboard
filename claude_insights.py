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
    """Generate a summary of pipeline data for Claude context"""
    if rep_name:
        rep_deals = deals_df[deals_df['Deal Owner'] == rep_name].copy()
    else:
        rep_deals = deals_df.copy()
    
    if rep_deals.empty:
        return "No deals in pipeline."
    
    summary = {
        "total_deals": len(rep_deals),
        "total_value": rep_deals['Amount'].sum() if 'Amount' in rep_deals.columns else 0,
        "deals_by_stage": rep_deals.groupby('Deal Stage')['Amount'].agg(['count', 'sum']).to_dict() if 'Deal Stage' in rep_deals.columns else {},
        "recent_activity": rep_deals.nlargest(5, 'Last Modified Date')[['Deal Name', 'Amount', 'Deal Stage', 'Close Date']].to_dict() if 'Last Modified Date' in rep_deals.columns else {}
    }
    
    return json.dumps(summary, indent=2, default=str)

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

You are a helpful sales operations analyst. Provide clear, actionable insights based on the data.

Current Pipeline Data:
{context_data}

When answering:
- Be specific and cite numbers from the data
- Focus on actionable recommendations
- Keep responses concise but thorough
- Use bullet points for clarity when appropriate
"""

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1500,
            messages=[
                {"role": "user", "content": question}
            ],
            system=system_message
        )
        
        return message.content[0].text
    
    except Exception as e:
        return f"Error getting response from Claude: {str(e)}"

def generate_daily_summary(deals_df, dashboard_df):
    """Generate automated daily change summary"""
    client = initialize_claude()
    if not client:
        return "Unable to generate daily summary. Please check your API key."
    
    # Calculate key metrics for summary
    total_pipeline = deals_df['Amount'].sum() if 'Amount' in deals_df.columns else 0
    total_deals = len(deals_df)
    
    # Get deals by stage
    stage_breakdown = deals_df.groupby('Deal Stage')['Amount'].agg(['count', 'sum']).to_dict() if 'Deal Stage' in deals_df.columns else {}
    
    # Get recent changes (last 7 days)
    if 'Last Modified Date' in deals_df.columns:
        deals_df['Last Modified Date'] = pd.to_datetime(deals_df['Last Modified Date'], errors='coerce')
        recent_changes = deals_df[deals_df['Last Modified Date'] >= datetime.now() - timedelta(days=7)]
        recent_summary = f"{len(recent_changes)} deals modified in last 7 days"
    else:
        recent_summary = "Date information not available"
    
    context = f"""
Current Pipeline Overview:
- Total Deals: {total_deals}
- Total Pipeline Value: ${total_pipeline:,.0f}
- Recent Activity: {recent_summary}

Stage Breakdown:
{json.dumps(stage_breakdown, indent=2, default=str)}
"""

    system_message = """You are a sales operations analyst providing a daily executive summary.

Create a concise daily summary that highlights:
1. Key pipeline metrics and health
2. Notable changes or trends
3. Deals requiring attention
4. Quick wins or opportunities

Keep it brief (3-4 paragraphs max) and actionable. Use a professional but friendly tone."""

    try:
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[
                {"role": "user", "content": f"Generate a daily pipeline summary based on this data:\n\n{context}"}
            ],
            system=system_message
        )
        
        return message.content[0].text
    
    except Exception as e:
        return f"Error generating summary: {str(e)}"

def display_insights_dashboard(deals_df, dashboard_df):
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
                summary = generate_daily_summary(deals_df, dashboard_df)
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
