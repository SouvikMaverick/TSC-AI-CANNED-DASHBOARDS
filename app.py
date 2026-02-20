"""
COO Dashboard - Billable HC & FTE Metrics  
A simplified Streamlit app focused on data tables without charts
"""
import streamlit as st
import pandas as pd
import json
import os
from typing import Dict, Any, List

# Page configuration
st.set_page_config(
    page_title="COO Dashboard - HC & FTE Metrics (Tables)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
    }
    h1 {
        color: #1f77b4;
        padding-bottom: 20px;
    }
    h2 {
        color: #2c3e50;
        padding-top: 20px;
    }
    h3 {
        color: #34495e;
    }
    .metric-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)


@st.cache_data(ttl=10)  # Cache for only 10 seconds to ensure fresh data
def load_data(json_path: str) -> List[Dict[str, Any]]:
    """Load and cache JSON data with short TTL for fresh updates."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_business_dataframe(data: List[Dict], metric_type: str = 'hc') -> pd.DataFrame:
    """Create a business breakdown dataframe."""
    records = []
    businesses = ['BET NA', 'HIL', 'GROWTH MARKETS', 'PLATINUM AC-CITI', 'PLATINUM AC-JPMC', 'TIME']
    
    for quarter_data in data:
        quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
        metrics = quarter_data.get('metrics', {})
        
        if not metrics:
            continue
        
        if metric_type == 'hc':
            total_key = 'total_billable_hc'
        else:
            total_key = 'total_billable_fte'
        
        for business in businesses:
            records.append({
                'Quarter': quarter,
                'Business': business,
                'Total': metrics.get(total_key, {}).get('by_business', {}).get(business, 0)
            })
    
    return pd.DataFrame(records)


def create_location_business_dataframe(data: List[Dict], metric_type: str = 'hc', location: str = 'onsite') -> pd.DataFrame:
    """Create a business breakdown dataframe for specific location (onsite/offshore)."""
    records = []
    businesses = ['BET NA', 'HIL', 'GROWTH MARKETS', 'PLATINUM AC-CITI', 'PLATINUM AC-JPMC', 'TIME']
    
    for quarter_data in data:
        quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
        metrics = quarter_data.get('metrics', {})
        
        if not metrics:
            continue
        
        if metric_type == 'hc':
            if location == 'onsite':
                location_key = 'total_onsite_hc'
            else:
                location_key = 'total_offshore_hc'
        else:
            if location == 'onsite':
                location_key = 'total_onsite_fte'
            else:
                location_key = 'total_offshore_fte'
        
        for business in businesses:
            records.append({
                'Quarter': quarter,
                'Business': business,
                'Total': metrics.get(location_key, {}).get('by_business', {}).get(business, 0)
            })
    
    return pd.DataFrame(records)


def display_metrics_cards(quarter_data: Dict, metric_type: str = 'hc'):
    """Display key metrics in card format."""
    metrics = quarter_data.get('metrics', {})
    
    if not metrics:
        st.warning("⚠️ No metrics data available for the selected quarter.")
        return
    
    if metric_type == 'hc':
        total = metrics.get('total_billable_hc', {}).get('total', 0)
        kpo = metrics.get('total_kpo_hc', {}).get('total', 0)
        non_kpo = metrics.get('total_non_kpo_hc', {}).get('total', 0)
        onsite = metrics.get('total_onsite_hc', {}).get('total', 0)
        onsite_kpo = metrics.get('onsite_kpo_hc', {}).get('total', 0)
        onsite_non_kpo = metrics.get('onsite_non_kpo_hc', {}).get('total', 0)
        offshore = metrics.get('total_offshore_hc', {}).get('total', 0)
        offshore_kpo = metrics.get('offshore_kpo_hc', {}).get('total', 0)
        offshore_non_kpo = metrics.get('offshore_non_kpo_hc', {}).get('total', 0)
        label = 'HC'
    else:
        total = metrics.get('total_billable_fte', {}).get('total', 0)
        kpo = metrics.get('total_kpo_fte', {}).get('total', 0)
        non_kpo = metrics.get('total_non_kpo_fte', {}).get('total', 0)
        onsite = metrics.get('total_onsite_fte', {}).get('total', 0)
        onsite_kpo = metrics.get('onsite_kpo_fte', {}).get('total', 0)
        onsite_non_kpo = metrics.get('onsite_non_kpo_fte', {}).get('total', 0)
        offshore = metrics.get('total_offshore_fte', {}).get('total', 0)
        offshore_kpo = metrics.get('offshore_kpo_fte', {}).get('total', 0)
        offshore_non_kpo = metrics.get('offshore_non_kpo_fte', {}).get('total', 0)
        label = 'FTE'
    
    kpo_pct = (kpo / total * 100) if total > 0 else 0
    non_kpo_pct = (non_kpo / total * 100) if total > 0 else 0
    onsite_pct = (onsite / total * 100) if total > 0 else 0
    offshore_pct = (offshore / total * 100) if total > 0 else 0
    
    # Row 1: Total, KPO, Non-KPO
    st.markdown("### 📊 Overall Metrics")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label=f"Total Billable {label}",
            value=f"{total:,.2f}",
            delta=None
        )
    
    with col2:
        st.metric(
            label=f"KPO {label}",
            value=f"{kpo:,.2f}",
            delta=f"{kpo_pct:.1f}%"
        )
    
    with col3:
        st.metric(
            label=f"Non-KPO {label}",
            value=f"{non_kpo:,.2f}",
            delta=f"{non_kpo_pct:.1f}%"
        )
    
    st.markdown("---")
    
    # Row 2: Onsite breakdown
    st.markdown("### 🏢 Onsite Metrics")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label=f"Total Onsite {label}",
            value=f"{onsite:,.2f}",
            delta=f"{onsite_pct:.1f}% of total"
        )
    
    with col2:
        onsite_kpo_pct = (onsite_kpo / onsite * 100) if onsite > 0 else 0
        st.metric(
            label=f"Onsite KPO {label}",
            value=f"{onsite_kpo:,.2f}",
            delta=f"{onsite_kpo_pct:.1f}% of onsite"
        )
    
    with col3:
        onsite_non_kpo_pct = (onsite_non_kpo / onsite * 100) if onsite > 0 else 0
        st.metric(
            label=f"Onsite Non-KPO {label}",
            value=f"{onsite_non_kpo:,.2f}",
            delta=f"{onsite_non_kpo_pct:.1f}% of onsite"
        )
    
    st.markdown("---")
    
    # Row 3: Offshore breakdown
    st.markdown("### 🌍 Offshore Metrics")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label=f"Total Offshore {label}",
            value=f"{offshore:,.2f}",
            delta=f"{offshore_pct:.1f}% of total"
        )
    
    with col2:
        offshore_kpo_pct = (offshore_kpo / offshore * 100) if offshore > 0 else 0
        st.metric(
            label=f"Offshore KPO {label}",
            value=f"{offshore_kpo:,.2f}",
            delta=f"{offshore_kpo_pct:.1f}% of offshore"
        )
    
    with col3:
        offshore_non_kpo_pct = (offshore_non_kpo / offshore * 100) if offshore > 0 else 0
        st.metric(
            label=f"Offshore Non-KPO {label}",
            value=f"{offshore_non_kpo:,.2f}",
            delta=f"{offshore_non_kpo_pct:.1f}% of offshore"
        )


def display_growth_metrics(data: List[Dict], metric_type: str = 'hc'):
    """Display growth metrics between quarters."""
    if len(data) < 2:
        return
    
    first = data[0].get('metrics', {})
    last = data[-1].get('metrics', {})
    
    # Check if metrics exist
    if not first or not last:
        st.warning("⚠️ Unable to calculate growth metrics - missing data")
        return
    
    if metric_type == 'hc':
        first_total = first.get('total_billable_hc', {}).get('total', 0)
        last_total = last.get('total_billable_hc', {}).get('total', 0)
        first_kpo = first.get('total_kpo_hc', {}).get('total', 0)
        last_kpo = last.get('total_kpo_hc', {}).get('total', 0)
        first_non_kpo = first.get('total_non_kpo_hc', {}).get('total', 0)
        last_non_kpo = last.get('total_non_kpo_hc', {}).get('total', 0)
        first_onsite = first.get('total_onsite_hc', {}).get('total', 0)
        last_onsite = last.get('total_onsite_hc', {}).get('total', 0)
        first_offshore = first.get('total_offshore_hc', {}).get('total', 0)
        last_offshore = last.get('total_offshore_hc', {}).get('total', 0)
        label = 'HC'
    else:
        first_total = first.get('total_billable_fte', {}).get('total', 0)
        last_total = last.get('total_billable_fte', {}).get('total', 0)
        first_kpo = first.get('total_kpo_fte', {}).get('total', 0)
        last_kpo = last.get('total_kpo_fte', {}).get('total', 0)
        first_non_kpo = first.get('total_non_kpo_fte', {}).get('total', 0)
        last_non_kpo = last.get('total_non_kpo_fte', {}).get('total', 0)
        first_onsite = first.get('total_onsite_fte', {}).get('total', 0)
        last_onsite = last.get('total_onsite_fte', {}).get('total', 0)
        first_offshore = first.get('total_offshore_fte', {}).get('total', 0)
        last_offshore = last.get('total_offshore_fte', {}).get('total', 0)
        label = 'FTE'
    
    total_growth = ((last_total - first_total) / first_total * 100) if first_total > 0 else 0
    kpo_growth = ((last_kpo - first_kpo) / first_kpo * 100) if first_kpo > 0 else 0
    non_kpo_growth = ((last_non_kpo - first_non_kpo) / first_non_kpo * 100) if first_non_kpo > 0 else 0
    onsite_growth = ((last_onsite - first_onsite) / first_onsite * 100) if first_onsite > 0 else 0
    offshore_growth = ((last_offshore - first_offshore) / first_offshore * 100) if first_offshore > 0 else 0
    
    st.subheader(f"📈 Growth Analysis (Q1 to Q3)")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            label=f"Total {label} Growth",
            value=f"{last_total:,.2f}",
            delta=f"{total_growth:.1f}%"
        )
    
    with col2:
        st.metric(
            label=f"KPO {label} Growth",
            value=f"{last_kpo:,.2f}",
            delta=f"{kpo_growth:.1f}%"
        )
    
    with col3:
        st.metric(
            label=f"Non-KPO {label} Growth",
            value=f"{last_non_kpo:,.2f}",
            delta=f"{non_kpo_growth:.1f}%"
        )
    
    with col4:
        st.metric(
            label=f"Onsite {label} Growth",
            value=f"{last_onsite:,.2f}",
            delta=f"{onsite_growth:.1f}%"
        )
    
    with col5:
        st.metric(
            label=f"Offshore {label} Growth",
            value=f"{last_offshore:,.2f}",
            delta=f"{offshore_growth:.1f}%"
        )


def create_fulfillment_dataframe(data: List[Dict]) -> pd.DataFrame:
    """Create a consolidated dataframe for fulfillment metrics."""
    records = []
    
    for quarter_data in data:
        quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
        metrics = quarter_data.get('metrics', {})
        
        if not metrics:
            continue
        
        records.append({
            'Quarter': quarter,
            'Total': metrics.get('total_demands', {}).get('total', 0),
            'Filled': metrics.get('filled_demands', {}).get('total', 0),
            'Open': metrics.get('open_demands', {}).get('total', 0),
            'Cancelled': metrics.get('cancelled_demands', {}).get('total', 0),
            'Expired': metrics.get('expired_demands', {}).get('total', 0),
            'Fulfillment_Rate': metrics.get('fulfillment_rate', {}).get('overall', 0)
        })
    
    return pd.DataFrame(records)


def create_fulfillment_business_dataframe(data: List[Dict]) -> pd.DataFrame:
    """Create a business breakdown dataframe for fulfillment."""
    records = []
    businesses = ['BET NA', 'HIL', 'GROWTH MARKETS', 'PLATINUM AC-CITI', 'PLATINUM AC-JPMC', 'TIME']
    
    for quarter_data in data:
        quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
        metrics = quarter_data.get('metrics', {})
        
        if not metrics:
            continue
        
        for business in businesses:
            records.append({
                'Quarter': quarter,
                'Business': business,
                'Total': metrics.get('total_demands', {}).get('by_business', {}).get(business, 0),
                'Filled': metrics.get('filled_demands', {}).get('by_business', {}).get(business, 0),
                'Open': metrics.get('open_demands', {}).get('by_business', {}).get(business, 0),
                'Fulfillment_Rate': metrics.get('fulfillment_rate', {}).get('by_business', {}).get(business, 0)
            })
    
    return pd.DataFrame(records)


def create_fulfillment_location_business_dataframe(data: List[Dict], location: str = 'onsite') -> pd.DataFrame:
    """Create a business breakdown dataframe for specific location (onsite/offshore) fulfillment."""
    records = []
    businesses = ['BET NA', 'HIL', 'GROWTH MARKETS', 'PLATINUM AC-CITI', 'PLATINUM AC-JPMC', 'TIME']
    
    for quarter_data in data:
        quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
        metrics = quarter_data.get('metrics', {})
        
        if not metrics:
            continue
        
        location_key = 'onsite_demands' if location == 'onsite' else 'offshore_demands'
        
        for business in businesses:
            # Get actual total demands by business for this location
            total_demands = metrics.get(location_key, {}).get('by_business', {}).get(business, 0)
            
            # Get actual filled and open demands by business for this location
            location_filled = metrics.get(location_key, {}).get('filled_by_business', {}).get(business, 0)
            location_open = metrics.get(location_key, {}).get('open_by_business', {}).get(business, 0)
            
            # Calculate fulfillment rate
            actionable = location_filled + location_open
            fulfillment_rate = (location_filled / actionable * 100) if actionable > 0 else 0
            
            records.append({
                'Quarter': quarter,
                'Business': business,
                'Total': total_demands,
                'Filled': location_filled,
                'Open': location_open,
                'Fulfillment_Rate': fulfillment_rate
            })
    
    return pd.DataFrame(records)


def display_fulfillment_metrics_cards(quarter_data: Dict):
    """Display fulfillment metrics in card format."""
    metrics = quarter_data.get('metrics', {})
    
    if not metrics:
        st.warning("⚠️ No fulfillment metrics data available for the selected quarter.")
        return
    
    total = metrics.get('total_demands', {}).get('total', 0)
    filled = metrics.get('filled_demands', {}).get('total', 0)
    open_count = metrics.get('open_demands', {}).get('total', 0)
    cancelled = metrics.get('cancelled_demands', {}).get('total', 0)
    expired = metrics.get('expired_demands', {}).get('total', 0)
    fulfillment_rate = metrics.get('fulfillment_rate', {}).get('overall', 0)
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.metric(
            label="Total Demands",
            value=f"{total:,}",
            delta=None
        )
    
    with col2:
        st.metric(
            label="Filled",
            value=f"{filled:,}",
            delta=f"{fulfillment_rate:.1f}%"
        )
    
    with col3:
        st.metric(
            label="Open",
            value=f"{open_count:,}",
            delta=None
        )
    
    with col4:
        st.metric(
            label="Cancelled",
            value=f"{cancelled:,}",
            delta=None
        )
    
    with col5:
        st.metric(
            label="Expired",
            value=f"{expired:,}",
            delta=None
        )
    
    with col6:
        onsite = metrics.get('onsite_demands', {}).get('total', 0)
        offshore = metrics.get('offshore_demands', {}).get('total', 0)
        offshore_pct = (offshore / total * 100) if total > 0 else 0
        st.metric(
            label="Offshore %",
            value=f"{offshore_pct:.1f}%",
            delta=f"{offshore:,} offshore"
        )


def main():
    """Main Streamlit app."""
    
    # Title and description
    st.title("📊 COO Dashboard - Workforce Metrics Analytics  ")
    st.markdown("""
    Comprehensive data tables for Billable Headcount (HC) and Full-Time Equivalent (FTE) metrics across quarters and business units.
    """)
    
    # Sidebar
    st.sidebar.title("📌 Dashboard Controls")
    st.sidebar.markdown("---")
    
    # File paths - resolve from workspace root (parent) first, then script directory
    from pathlib import Path
    script_dir = Path(__file__).resolve().parent
    workspace_root = script_dir.parent
    
    # Try workspace root first (where the complete/correct JSON is), then script directory
    hc_json_path = workspace_root / 'billable_hc_metrics.json'
    if not hc_json_path.exists():
        hc_json_path = script_dir / 'billable_hc_metrics.json'
    
    fte_json_path = workspace_root / 'billable_fte_metrics.json'
    if not fte_json_path.exists():
        fte_json_path = script_dir / 'billable_fte_metrics.json'
    
    fulfillment_json_path = workspace_root / 'fulfillment_metrics.json'
    if not fulfillment_json_path.exists():
        fulfillment_json_path = script_dir / 'fulfillment_metrics.json'
    
    hc_json_path = str(hc_json_path)
    fte_json_path = str(fte_json_path)
    fulfillment_json_path = str(fulfillment_json_path)
    
    # Load data
    try:
        hc_data = load_data(hc_json_path)
        fte_data = load_data(fte_json_path)
        
        # Load fulfillment data if available
        fulfillment_data = None
        if os.path.exists(fulfillment_json_path):
            fulfillment_data = load_data(fulfillment_json_path)
            st.sidebar.success("✅ All data loaded successfully (HC, FTE, Fulfillment)")
        else:
            st.sidebar.success("✅ HC and FTE data loaded successfully")
            st.sidebar.warning("⚠️ Fulfillment data not found")
        
        # Show which files were loaded
        st.sidebar.markdown("---")
        st.sidebar.markdown("**📁 Data Files Used:**")
        st.sidebar.caption(f"HC: {os.path.basename(hc_json_path)}")
        st.sidebar.caption(f"FTE: {os.path.basename(fte_json_path)}")
        st.sidebar.caption(f"Fulfillment: {os.path.basename(fulfillment_json_path)}")
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        return
    
    # Metric type selector
    metric_options = ['Headcount (HC)', 'Full-Time Equivalent (FTE)', 'HC vs FTE Comparison']
    
    if fulfillment_data:
        metric_options.append('Fulfillment Metrics')
    
    metric_type = st.sidebar.radio(
        "Select Metric Type",
        options=metric_options,
        index=0
    )
    
    st.sidebar.markdown("---")
    
    # Quarter selector
    if metric_type == 'Headcount (HC)':
        data = hc_data
        label = 'HC'
        metric_key = 'hc'
    elif metric_type == 'Full-Time Equivalent (FTE)':
        data = fte_data
        label = 'FTE'
        metric_key = 'fte'
    elif metric_type == 'Fulfillment Metrics':
        data = fulfillment_data if fulfillment_data else hc_data
        label = 'Fulfillment'
        metric_key = 'fulfillment'
    else:
        data = hc_data
        label = 'HC'
        metric_key = 'hc'
    
    if metric_type not in ['HC vs FTE Comparison', 'Fulfillment Metrics']:
        quarters = [f"{q['fiscal_year']} {q['quarter']}" for q in data]
        selected_quarter = st.sidebar.selectbox("Select Quarter for Detailed View", quarters, index=len(quarters)-1, key="quarter_hc_fte")
        selected_quarter_data = [q for q in data if f"{q['fiscal_year']} {q['quarter']}" == selected_quarter][0]
    elif metric_type == 'Fulfillment Metrics' and fulfillment_data:
        quarters = [f"{q['fiscal_year']} {q['quarter']}" for q in fulfillment_data]
        selected_quarter = st.sidebar.selectbox("Select Quarter for Detailed View", quarters, index=len(quarters)-1, key="quarter_fulfillment")
        selected_quarter_data = [q for q in fulfillment_data if f"{q['fiscal_year']} {q['quarter']}" == selected_quarter][0]
    
    # Business selector
    businesses = ['All', 'BET NA', 'HIL', 'GROWTH MARKETS', 'PLATINUM AC-CITI', 'PLATINUM AC-JPMC', 'TIME']
    selected_business = st.sidebar.selectbox("Filter by Business", businesses, index=0, key="business_filter")
    
    st.sidebar.markdown("---")
    st.sidebar.info(f"📅 Last Updated: {data[0]['extraction_date']}")
    
    # Main content
    if metric_type == 'HC vs FTE Comparison':
        st.header("🔄 HC vs FTE Comparison Analysis")
        
        # Comparison table
        st.subheader("📋 Detailed Comparison Table")
        
        comparison_records = []
        for i, hc_quarter in enumerate(hc_data):
            if i < len(fte_data):
                hc_metrics = hc_quarter.get('metrics', {})
                fte_metrics = fte_data[i].get('metrics', {})
                
                if not hc_metrics or not fte_metrics:
                    continue
                
                quarter = f"{hc_quarter['fiscal_year']} {hc_quarter['quarter']}"
                hc_total = hc_metrics.get('total_billable_hc', {}).get('total', 0)
                fte_total = fte_metrics.get('total_billable_fte', {}).get('total', 0)
                diff = hc_total - fte_total
                pct_diff = (diff / hc_total * 100) if hc_total > 0 else 0
                
                comparison_records.append({
                    'Quarter': quarter,
                    'Total HC': f"{hc_total:,.2f}",
                    'Total FTE': f"{fte_total:,.2f}",
                    'Difference': f"{diff:,.2f}",
                    '% Difference': f"{pct_diff:.2f}%"
                })
        
        comparison_df = pd.DataFrame(comparison_records)
        st.dataframe(comparison_df, use_container_width=True, hide_index=True)
        
        # Download button
        csv_comparison = comparison_df.to_csv(index=False)
        st.download_button(
            label="📥 Download HC vs FTE Comparison as CSV",
            data=csv_comparison,
            file_name="hc_vs_fte_comparison.csv",
            mime="text/csv",
            key="download_hc_vs_fte"
        )
        
    elif metric_type in ['Headcount (HC)', 'Full-Time Equivalent (FTE)']:
        # Detailed data tables
        st.header("📋 Data Tables by Business Unit")
        st.markdown("Comprehensive breakdown showing Overall, Onsite, and Offshore metrics across all business units")
        
        # Create dataframes for all three categories
        df_business_overall = create_business_dataframe(data, metric_key)
        df_business_onsite = create_location_business_dataframe(data, metric_key, 'onsite')
        df_business_offshore = create_location_business_dataframe(data, metric_key, 'offshore')
        
        # Helper function to create pivot table with summary rows and QTD
        def create_enhanced_pivot_table(df_business_data, data_source, metric_key, table_title):
            """Create a pivot table with VRTU, KPO, VRTU Excl KPO rows and QTD column."""
            
            # Create pivot table
            pivot_df = df_business_data.pivot(index='Business', columns='Quarter', values='Total')
            
            # Calculate QTD - difference between last two quarters per business
            qtd_values = []
            num_quarters = len(pivot_df.columns)
            
            if num_quarters >= 2:
                last_quarter = pivot_df.columns[-1]
                second_last_quarter = pivot_df.columns[-2]
                
                for business in pivot_df.index:
                    qtd_diff = pivot_df.loc[business, last_quarter] - pivot_df.loc[business, second_last_quarter]
                    qtd_values.append(qtd_diff)
            else:
                qtd_values = [0] * len(pivot_df.index)
            
            pivot_df['QTD'] = qtd_values
            
            # Add three summary rows
            # 1. VRTU - Total per quarter across all businesses
            vrtu_row = {}
            for col in pivot_df.columns:
                if col == 'QTD':
                    vrtu_row[col] = sum(qtd_values) if num_quarters >= 2 else 0
                else:
                    vrtu_row[col] = pivot_df[col].sum()
            
            # 2. KPO - KPO numbers for each quarter
            kpo_row = {}
            kpo_values_list = []
            
            if table_title == "Overall":
                for quarter_data in data_source:
                    quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
                    if metric_key == 'hc':
                        kpo_val = quarter_data['metrics']['total_kpo_hc']['total']
                    else:
                        kpo_val = quarter_data['metrics']['total_kpo_fte']['total']
                    kpo_row[quarter] = kpo_val
                    kpo_values_list.append(kpo_val)
                
                if len(kpo_values_list) >= 2:
                    kpo_row['QTD'] = kpo_values_list[-1] - kpo_values_list[-2]
                else:
                    kpo_row['QTD'] = 0
            elif table_title == "Onsite":
                # For Onsite, use onsite_kpo values
                for quarter_data in data_source:
                    quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
                    if metric_key == 'hc':
                        kpo_val = quarter_data['metrics']['onsite_kpo_hc']['total']
                    else:
                        kpo_val = quarter_data['metrics']['onsite_kpo_fte']['total']
                    kpo_row[quarter] = kpo_val
                    kpo_values_list.append(kpo_val)
                
                if len(kpo_values_list) >= 2:
                    kpo_row['QTD'] = kpo_values_list[-1] - kpo_values_list[-2]
                else:
                    kpo_row['QTD'] = 0
            elif table_title == "Offshore":
                # For Offshore, use offshore_kpo values
                for quarter_data in data_source:
                    quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
                    if metric_key == 'hc':
                        kpo_val = quarter_data['metrics']['offshore_kpo_hc']['total']
                    else:
                        kpo_val = quarter_data['metrics']['offshore_kpo_fte']['total']
                    kpo_row[quarter] = kpo_val
                    kpo_values_list.append(kpo_val)
                
                if len(kpo_values_list) >= 2:
                    kpo_row['QTD'] = kpo_values_list[-1] - kpo_values_list[-2]
                else:
                    kpo_row['QTD'] = 0
            
            # 3. VRTU Excl KPO - Total minus KPO
            vrtu_excl_kpo_row = {}
            for quarter in pivot_df.columns:
                if quarter == 'QTD':
                    vrtu_excl_kpo_row[quarter] = vrtu_row[quarter] - kpo_row[quarter]
                elif quarter in kpo_row:
                    vrtu_excl_kpo_row[quarter] = vrtu_row[quarter] - kpo_row[quarter]
            
            # Add the summary rows to the dataframe
            pivot_df.loc['VRTU'] = pd.Series(vrtu_row)
            pivot_df.loc['KPO'] = pd.Series(kpo_row)
            pivot_df.loc['VRTU Excl KPO'] = pd.Series(vrtu_excl_kpo_row)
            
            return pivot_df
        
        # Helper function for styling
        def highlight_pivot_table(df):
            """Apply highlighting to summary rows and QTD column."""
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            
            # Highlight summary rows (last 3 rows)
            summary_rows = ['VRTU', 'KPO', 'VRTU Excl KPO']
            for row in summary_rows:
                if row in df.index:
                    styles.loc[row, :] = 'background-color: #ffffcc; font-weight: bold'
            
            # Highlight QTD column for all rows
            if 'QTD' in df.columns:
                styles['QTD'] = 'background-color: #e6f3ff; font-weight: bold'
            
            # For summary rows + QTD column (intersection), use a different color
            for row in summary_rows:
                if row in df.index and 'QTD' in df.columns:
                    styles.loc[row, 'QTD'] = 'background-color: #ffcccc; font-weight: bold'
            
            return styles
        
        # Table 1: Overall Numbers
        st.markdown("---")
        st.subheader(f"📊 Overall {label} by Business Unit")
        st.caption("Total billable headcount/FTE across all locations")
        
        pivot_overall = create_enhanced_pivot_table(df_business_overall, data, metric_key, "Overall")
        styled_overall = pivot_overall.style.format("{:.2f}")
        styled_overall = styled_overall.apply(lambda x: highlight_pivot_table(pivot_overall), axis=None)
        st.dataframe(styled_overall, use_container_width=True)
        
        csv_overall = pivot_overall.to_csv()
        st.download_button(
            label=f"📥 Download Overall {label} Data as CSV",
            data=csv_overall,
            file_name=f"{metric_key}_overall_business_metrics.csv",
            mime="text/csv",
            key=f"download_overall_{metric_key}"
        )
        
        # Table 2: Onsite Numbers
        st.markdown("---")
        st.subheader(f"🏢 Onsite {label} by Business Unit")
        st.caption("Onsite workforce metrics by business unit")
        
        pivot_onsite = create_enhanced_pivot_table(df_business_onsite, data, metric_key, "Onsite")
        styled_onsite = pivot_onsite.style.format("{:.2f}")
        styled_onsite = styled_onsite.apply(lambda x: highlight_pivot_table(pivot_onsite), axis=None)
        st.dataframe(styled_onsite, use_container_width=True)
        
        csv_onsite = pivot_onsite.to_csv()
        st.download_button(
            label=f"📥 Download Onsite {label} Data as CSV",
            data=csv_onsite,
            file_name=f"{metric_key}_onsite_business_metrics.csv",
            mime="text/csv",
            key=f"download_onsite_{metric_key}"
        )
        
        # Table 3: Offshore Numbers
        st.markdown("---")
        st.subheader(f"🌍 Offshore {label} by Business Unit")
        st.caption("Offshore workforce metrics by business unit")
        
        pivot_offshore = create_enhanced_pivot_table(df_business_offshore, data, metric_key, "Offshore")
        styled_offshore = pivot_offshore.style.format("{:.2f}")
        styled_offshore = styled_offshore.apply(lambda x: highlight_pivot_table(pivot_offshore), axis=None)
        st.dataframe(styled_offshore, use_container_width=True)
        
        csv_offshore = pivot_offshore.to_csv()
        st.download_button(
            label=f"📥 Download Offshore {label} Data as CSV",
            data=csv_offshore,
            file_name=f"{metric_key}_offshore_business_metrics.csv",
            mime="text/csv",
            key=f"download_offshore_{metric_key}"
        )
    
    if metric_type == 'Fulfillment Metrics' and fulfillment_data:
        st.header("📊 Fulfillment Metrics - Demand & Resource Allocation")
        
        st.markdown("---")
        
        # Fulfillment summary table
        st.subheader("📋 Fulfillment Trends Table")
        df_fulfillment = create_fulfillment_dataframe(fulfillment_data)
        
        st.dataframe(df_fulfillment.style.format({
            'Total': '{:.0f}',
            'Filled': '{:.0f}',
            'Open': '{:.0f}',
            'Cancelled': '{:.0f}',
            'Expired': '{:.0f}',
            'Fulfillment_Rate': '{:.2f}%'
        }), use_container_width=True)
        
        csv_fulfillment = df_fulfillment.to_csv(index=False)
        st.download_button(
            label="📥 Download Fulfillment Trends as CSV",
            data=csv_fulfillment,
            file_name="fulfillment_trends.csv",
            mime="text/csv",
            key="download_fulfillment_trends"
        )
        
        st.markdown("---")
        
        # Business breakdown
        st.header("🏢 Fulfillment by Business Unit")
        df_business_fulfillment = create_fulfillment_business_dataframe(fulfillment_data)
        
        if selected_business != 'All':
            df_business_fulfillment = df_business_fulfillment[df_business_fulfillment['Business'] == selected_business]
        
        st.subheader("Fulfillment Metrics by Business")
        st.dataframe(df_business_fulfillment.style.format({
            'Total': '{:.0f}',
            'Filled': '{:.0f}',
            'Open': '{:.0f}',
            'Fulfillment_Rate': '{:.2f}%'
        }), use_container_width=True)
        
        csv_business = df_business_fulfillment.to_csv(index=False)
        st.download_button(
            label="📥 Download Business Fulfillment as CSV",
            data=csv_business,
            file_name="fulfillment_by_business.csv",
            mime="text/csv",
            key="download_business_fulfillment"
        )
        
        st.markdown("---")
        
        # Detailed tables
        st.header("📋 Detailed Fulfillment Data by Business Unit")
        
        # Create pivot tables for different metrics
        def create_fulfillment_pivot(df, metric_col, title, fulfillment_data_source, location_type='overall'):
            pivot = df.pivot(index='Business', columns='Quarter', values=metric_col)
            
            # Add QTD column
            if len(pivot.columns) >= 2:
                pivot['QTD'] = pivot.iloc[:, -1] - pivot.iloc[:, -2]
            else:
                pivot['QTD'] = 0
            
            # Add three summary rows: VRTU, KPO, VRTU Excl KPO
            # 1. VRTU - Total per quarter across all businesses
            vrtu_row = {}
            for col in pivot.columns:
                if col == 'QTD':
                    vrtu_row[col] = pivot['QTD'].sum()
                else:
                    vrtu_row[col] = pivot[col].sum()
            
            # 2. KPO - KPO numbers for each quarter (only for overall, not onsite)
            kpo_row = {}
            if location_type == 'onsite':
                # For onsite, don't show KPO data - set all to 0
                for col in pivot.columns:
                    kpo_row[col] = 0
            else:
                for quarter_data in fulfillment_data_source:
                    quarter = f"{quarter_data['fiscal_year']} {quarter_data['quarter']}"
                    if quarter in pivot.columns:
                        kpo_val = 0
                        
                        # Use overall kpo_demands for overall tables
                        if 'kpo_demands' in quarter_data['metrics']:
                            if metric_col == 'Total':
                                kpo_val = quarter_data['metrics']['kpo_demands']['total']
                            elif metric_col == 'Filled':
                                kpo_val = quarter_data['metrics']['kpo_demands']['filled']
                            elif metric_col == 'Open':
                                kpo_val = quarter_data['metrics']['kpo_demands']['open']
                            elif metric_col == 'Fulfillment_Rate':
                                # Calculate KPO fulfillment rate
                                kpo_filled = quarter_data['metrics']['kpo_demands']['filled']
                                kpo_open = quarter_data['metrics']['kpo_demands']['open']
                                kpo_actionable = kpo_filled + kpo_open
                                kpo_val = (kpo_filled / kpo_actionable * 100) if kpo_actionable > 0 else 0
                        
                        kpo_row[quarter] = kpo_val
            
            # Calculate KPO QTD
            kpo_quarters = [q for q in pivot.columns if q != 'QTD']
            if len(kpo_quarters) >= 2:
                last_q = kpo_quarters[-1]
                second_last_q = kpo_quarters[-2]
                kpo_row['QTD'] = kpo_row.get(last_q, 0) - kpo_row.get(second_last_q, 0)
            else:
                kpo_row['QTD'] = 0
            
            # 3. VRTU Excl KPO - Total minus KPO
            vrtu_excl_kpo_row = {}
            for quarter in pivot.columns:
                if quarter in kpo_row:
                    vrtu_excl_kpo_row[quarter] = vrtu_row[quarter] - kpo_row[quarter]
                else:
                    vrtu_excl_kpo_row[quarter] = vrtu_row[quarter]
            
            # Add the summary rows to the dataframe
            pivot.loc['VRTU'] = pd.Series(vrtu_row)
            pivot.loc['KPO'] = pd.Series(kpo_row)
            pivot.loc['VRTU Excl KPO'] = pd.Series(vrtu_excl_kpo_row)
            
            return pivot
        
        # Helper function for highlighting fulfillment tables
        def highlight_fulfillment_table(df):
            """Apply highlighting to summary rows (VRTU, KPO, VRTU Excl KPO) and QTD column."""
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            
            # Highlight summary rows (last 3 rows)
            summary_rows = ['VRTU', 'KPO', 'VRTU Excl KPO']
            for row in summary_rows:
                if row in df.index:
                    styles.loc[row, :] = 'background-color: #ffffcc; font-weight: bold'
            
            # Highlight QTD column for all rows
            if 'QTD' in df.columns:
                styles['QTD'] = 'background-color: #e6f3ff; font-weight: bold'
            
            # For summary rows + QTD column (intersection), use a different color
            for row in summary_rows:
                if row in df.index and 'QTD' in df.columns:
                    styles.loc[row, 'QTD'] = 'background-color: #ffcccc; font-weight: bold'
            
            return styles
        
        # Total Demands Table
        st.subheader("📊 Total Demands by Business")
        pivot_total = create_fulfillment_pivot(df_business_fulfillment, 'Total', 'Total Demands', fulfillment_data)
        styled_total = pivot_total.style.format("{:.0f}")
        styled_total = styled_total.apply(lambda x: highlight_fulfillment_table(pivot_total), axis=None)
        st.dataframe(styled_total, use_container_width=True)
        
        csv_total = pivot_total.to_csv()
        st.download_button(
            label="📥 Download Total Demands as CSV",
            data=csv_total,
            file_name="fulfillment_total_demands.csv",
            mime="text/csv",
            key="download_total_demands"
        )
        
        # Filled Demands Table
        st.markdown("---")
        st.subheader("✅ Filled Demands by Business")
        pivot_filled = create_fulfillment_pivot(df_business_fulfillment, 'Filled', 'Filled Demands', fulfillment_data)
        styled_filled = pivot_filled.style.format("{:.0f}")
        styled_filled = styled_filled.apply(lambda x: highlight_fulfillment_table(pivot_filled), axis=None)
        st.dataframe(styled_filled, use_container_width=True)
        
        csv_filled = pivot_filled.to_csv()
        st.download_button(
            label="📥 Download Filled Demands as CSV",
            data=csv_filled,
            file_name="fulfillment_filled_demands.csv",
            mime="text/csv",
            key="download_filled_demands"
        )
        
        # Open Demands Table
        st.markdown("---")
        st.subheader("⏳ Open Demands by Business")
        pivot_open = create_fulfillment_pivot(df_business_fulfillment, 'Open', 'Open Demands', fulfillment_data)
        styled_open = pivot_open.style.format("{:.0f}")
        styled_open = styled_open.apply(lambda x: highlight_fulfillment_table(pivot_open), axis=None)
        st.dataframe(styled_open, use_container_width=True)
        
        csv_open = pivot_open.to_csv()
        st.download_button(
            label="📥 Download Open Demands as CSV",
            data=csv_open,
            file_name="fulfillment_open_demands.csv",
            mime="text/csv",
            key="download_open_demands"
        )
        
        # Fulfillment Rate Table
        st.markdown("---")
        st.subheader("📊 Fulfillment Rate (%) by Business")
        pivot_rate = create_fulfillment_pivot(df_business_fulfillment, 'Fulfillment_Rate', 'Fulfillment Rate', fulfillment_data)
        styled_rate = pivot_rate.style.format("{:.2f}")
        styled_rate = styled_rate.apply(lambda x: highlight_fulfillment_table(pivot_rate), axis=None)
        st.dataframe(styled_rate, use_container_width=True)
        
        csv_rate = pivot_rate.to_csv()
        st.download_button(
            label="📥 Download Fulfillment Rate as CSV",
            data=csv_rate,
            file_name="fulfillment_rate.csv",
            mime="text/csv",
            key="download_fulfillment_rate"
        )
        
        # Onsite Demands Tables
        st.markdown("---")
        st.header("🏢 Onsite Fulfillment Metrics by Business Unit")
        
        df_business_onsite = create_fulfillment_location_business_dataframe(fulfillment_data, 'onsite')
        
        # Onsite Total Demands
        st.subheader("📊 Onsite Total Demands by Business")
        pivot_onsite_total = create_fulfillment_pivot(df_business_onsite, 'Total', 'Onsite Total Demands', fulfillment_data, 'onsite')
        styled_onsite_total = pivot_onsite_total.style.format("{:.0f}")
        styled_onsite_total = styled_onsite_total.apply(lambda x: highlight_fulfillment_table(pivot_onsite_total), axis=None)
        st.dataframe(styled_onsite_total, use_container_width=True)
        
        csv_onsite_total = pivot_onsite_total.to_csv()
        st.download_button(
            label="📥 Download Onsite Total Demands as CSV",
            data=csv_onsite_total,
            file_name="fulfillment_onsite_total_demands.csv",
            mime="text/csv",
            key="download_onsite_total"
        )
        
        # Onsite Filled Demands
        st.markdown("---")
        st.subheader("✅ Onsite Filled Demands by Business")
        pivot_onsite_filled = create_fulfillment_pivot(df_business_onsite, 'Filled', 'Onsite Filled Demands', fulfillment_data, 'onsite')
        styled_onsite_filled = pivot_onsite_filled.style.format("{:.0f}")
        styled_onsite_filled = styled_onsite_filled.apply(lambda x: highlight_fulfillment_table(pivot_onsite_filled), axis=None)
        st.dataframe(styled_onsite_filled, use_container_width=True)
        
        csv_onsite_filled = pivot_onsite_filled.to_csv()
        st.download_button(
            label="📥 Download Onsite Filled Demands as CSV",
            data=csv_onsite_filled,
            file_name="fulfillment_onsite_filled_demands.csv",
            mime="text/csv",
            key="download_onsite_filled"
        )
        
        # Onsite Open Demands
        st.markdown("---")
        st.subheader("⏳ Onsite Open Demands by Business")
        pivot_onsite_open = create_fulfillment_pivot(df_business_onsite, 'Open', 'Onsite Open Demands', fulfillment_data, 'onsite')
        styled_onsite_open = pivot_onsite_open.style.format("{:.0f}")
        styled_onsite_open = styled_onsite_open.apply(lambda x: highlight_fulfillment_table(pivot_onsite_open), axis=None)
        st.dataframe(styled_onsite_open, use_container_width=True)
        
        csv_onsite_open = pivot_onsite_open.to_csv()
        st.download_button(
            label="📥 Download Onsite Open Demands as CSV",
            data=csv_onsite_open,
            file_name="fulfillment_onsite_open_demands.csv",
            mime="text/csv",
            key="download_onsite_open"
        )
        
        # Offshore Demands Tables
        st.markdown("---")
        st.header("🌍 Offshore Fulfillment Metrics by Business Unit")
        
        df_business_offshore = create_fulfillment_location_business_dataframe(fulfillment_data, 'offshore')
        
        # Offshore Total Demands
        st.subheader("📊 Offshore Total Demands by Business")
        pivot_offshore_total = create_fulfillment_pivot(df_business_offshore, 'Total', 'Offshore Total Demands', fulfillment_data)
        styled_offshore_total = pivot_offshore_total.style.format("{:.0f}")
        styled_offshore_total = styled_offshore_total.apply(lambda x: highlight_fulfillment_table(pivot_offshore_total), axis=None)
        st.dataframe(styled_offshore_total, use_container_width=True)
        
        csv_offshore_total = pivot_offshore_total.to_csv()
        st.download_button(
            label="📥 Download Offshore Total Demands as CSV",
            data=csv_offshore_total,
            file_name="fulfillment_offshore_total_demands.csv",
            mime="text/csv",
            key="download_offshore_total"
        )
        
        # Offshore Filled Demands
        st.markdown("---")
        st.subheader("✅ Offshore Filled Demands by Business")
        pivot_offshore_filled = create_fulfillment_pivot(df_business_offshore, 'Filled', 'Offshore Filled Demands', fulfillment_data)
        styled_offshore_filled = pivot_offshore_filled.style.format("{:.0f}")
        styled_offshore_filled = styled_offshore_filled.apply(lambda x: highlight_fulfillment_table(pivot_offshore_filled), axis=None)
        st.dataframe(styled_offshore_filled, use_container_width=True)
        
        csv_offshore_filled = pivot_offshore_filled.to_csv()
        st.download_button(
            label="📥 Download Offshore Filled Demands as CSV",
            data=csv_offshore_filled,
            file_name="fulfillment_offshore_filled_demands.csv",
            mime="text/csv",
            key="download_offshore_filled"
        )
        
        # Offshore Open Demands
        st.markdown("---")
        st.subheader("⏳ Offshore Open Demands by Business")
        pivot_offshore_open = create_fulfillment_pivot(df_business_offshore, 'Open', 'Offshore Open Demands', fulfillment_data)
        styled_offshore_open = pivot_offshore_open.style.format("{:.0f}")
        styled_offshore_open = styled_offshore_open.apply(lambda x: highlight_fulfillment_table(pivot_offshore_open), axis=None)
        st.dataframe(styled_offshore_open, use_container_width=True)
        
        csv_offshore_open = pivot_offshore_open.to_csv()
        st.download_button(
            label="📥 Download Offshore Open Demands as CSV",
            data=csv_offshore_open,
            file_name="fulfillment_offshore_open_demands.csv",
            mime="text/csv",
            key="download_offshore_open"
        )
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>COO Dashboard - Workforce Metrics Analytics  </p>
        <p>Built with Streamlit • Data updated: {}</p>
    </div>
    """.format(data[0]['extraction_date']), unsafe_allow_html=True)


if __name__ == "__main__":
    main()
