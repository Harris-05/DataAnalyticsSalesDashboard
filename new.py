"""
Dynamic Business Analytics Dashboard
=====================================
A comprehensive Streamlit application for interactive sales data analysis.
Upload XLSB files, filter across multiple dimensions, and view key performance metrics instantly.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pyxlsb import open_workbook

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Business Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Add custom styling
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
    }
    .section-header {
        color: #1f77b4;
        font-size: 1.3em;
        font-weight: bold;
        margin-top: 30px;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING & CACHING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data_from_xlsb(file):
    """
    Load data from an XLSB file, specifically from the 'DATA_BK' sheet.
    
    Parameters:
    -----------
    file : UploadedFile
        The XLSB file uploaded via Streamlit
        
    Returns:
    --------
    pd.DataFrame
        Cleaned dataframe with USDAmt as numeric type, NaN rows removed
    """
    rows = []
    try:
        with open_workbook(file) as wb:
            try:
                with wb.get_sheet('DATA_BK') as sheet:
                    for row in sheet.rows():
                        rows.append([item.v for item in row])
            except Exception as e:
                st.error(f"❌ Cannot find 'DATA_BK' sheet. Available sheets: Check your file. Error: {e}")
                return pd.DataFrame()
        
        # Create DataFrame from rows
        df = pd.DataFrame(rows)
        
        # Use first row as column headers
        if len(df) > 0:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)
        
        # Convert USDAmt to numeric (primary metric)
        if "USDAmt" in df.columns:
            df["USDAmt"] = pd.to_numeric(df["USDAmt"], errors="coerce")
        else:
            st.error("❌ 'USDAmt' column not found in the data. Please verify your file.")
            return pd.DataFrame()
        
        # Drop rows with missing USDAmt values
        initial_rows = len(df)
        df = df.dropna(subset=["USDAmt"])
        removed_rows = initial_rows - len(df)
        
        if removed_rows > 0:
            st.info(f"ℹ️ Removed {removed_rows} rows with missing USDAmt values.")
        
        return df
        
    except FileNotFoundError:
        st.error("❌ File not found. Please ensure the file is properly uploaded.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        return pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR: FILE UPLOAD & INSTRUCTIONS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📂 DATA UPLOAD & INSTRUCTIONS")
    st.write("""
    ### How to Use This Dashboard:
    1. **Upload** an XLSB file from your computer
    2. **Select Filters** → Choose which columns to filter by
    3. **Filter Values** → Pick specific values for each column
    4. **View Metrics** → See KPIs and Top 10 rankings update instantly
    5. **Export Data** → Scroll down to view or download filtered data
    """)
    st.divider()
    
    # File uploader
    uploaded_file = st.file_uploader("📤 Upload XLSB File", type=['xlsb'])

# Stop execution if no file is uploaded
if not uploaded_file:
    st.info("👈 Please upload an XLSB file using the sidebar uploader to get started.")
    st.stop()

# Load the data with spinner for better UX
with st.spinner("🔄 Loading data from 'DATA_BK' sheet..."):
    df = load_data_from_xlsb(uploaded_file)

# Validate data
if df.empty:
    st.error("❌ No valid data could be loaded. Please check your file.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR: DYNAMIC FILTERING SYSTEM
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.divider()
    st.header("🎛️ DYNAMIC FILTERS")
    st.write("**Step 1: Choose dimensions to filter by**")
    
    # Suggested columns for filtering
    suggested_columns = ["Region", "BusinessSegment", "ClientSegment", "Product"]
    default_filter_columns = [col for col in suggested_columns if col in df.columns]
    
    selected_filter_columns = st.multiselect(
        "Select Filter Dimensions:",
        options=df.columns.tolist(),
        default=default_filter_columns,
        help="Choose which columns you want to filter. A dropdown will appear for each."
    )
    
    st.write("---")
    st.write("**Step 2: Select values for each dimension**")
    
    # Initialize filtered dataframe
    filtered_df = df.copy()
    
    # Generate dynamic filter dropdowns
    for col in selected_filter_columns:
        unique_vals = sorted([str(x) for x in df[col].dropna().unique()])
        
        selected_vals = st.multiselect(
            f"📌 {col}:",
            options=unique_vals,
            default=unique_vals,  # Default to all selected
            key=f"filter_{col}"
        )
        
        # Apply filter to dataframe
        if selected_vals:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

# ─────────────────────────────────────────────────────────────────────────────
# MAIN DASHBOARD HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.title("📊 BUSINESS ANALYTICS DASHBOARD")
st.write(f"*Dashboard generated for {uploaded_file.name} | Last updated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}*")

# Validate filtered data
if filtered_df.empty:
    st.warning("⚠️ No data matches your current filter selection. Try selecting more options in the sidebar.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# KEY PERFORMANCE INDICATORS (KPIs)
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("📈 KEY PERFORMANCE INDICATORS")

# Calculate KPIs
total_revenue = filtered_df["USDAmt"].sum()
transaction_count = len(filtered_df)
avg_transaction = filtered_df["USDAmt"].mean()

# Display KPIs in three columns
kpi_col1, kpi_col2, kpi_col3 = st.columns(3)

with kpi_col1:
    st.metric(
        label="💰 Total Revenue",
        value=f"${total_revenue:,.2f}",
        delta=f"{len(filtered_df)} transactions"
    )

with kpi_col2:
    st.metric(
        label="📦 Transaction Count",
        value=f"{transaction_count:,}",
        delta="filtered records"
    )

with kpi_col3:
    st.metric(
        label="📊 Average Transaction Value",
        value=f"${avg_transaction:,.2f}",
        delta="per transaction"
    )

# ─────────────────────────────────────────────────────────────────────────────
# TOP 10 VISUALIZATIONS
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
st.subheader("🏆 TOP 10 RANKINGS")
st.write("*Rankings based on sum of revenue (USDAmt) for the current filtered dataset*")

# Create three columns for the three charts
chart_col1, chart_col2, chart_col3 = st.columns(3)

# Chart 1: Top 10 Clients
with chart_col1:
    st.write("##### Top 10 Clients")
    if "CustomerName" in filtered_df.columns:
        top_clients = filtered_df.groupby("CustomerName")["USDAmt"].sum().nlargest(10).reset_index()
        top_clients = top_clients.sort_values("USDAmt", ascending=True)  # Sort for horizontal chart
        
        fig_clients = px.bar(
            top_clients,
            x="USDAmt",
            y="CustomerName",
            labels={"CustomerName": "Client", "USDAmt": "Revenue ($)"},
            text_auto=".2s"
        )
        fig_clients.update_layout(
            height=400,
            showlegend=False,
            xaxis_title="Revenue ($)",
            yaxis_title=""
        )
        st.plotly_chart(fig_clients, use_container_width=True)
    else:
        st.warning("⚠️ 'CustomerName' column not found.")

# Chart 2: Top 10 Branches
with chart_col2:
    st.write("##### Top 10 Branches")
    if "BranchName" in filtered_df.columns:
        top_branches = filtered_df.groupby("BranchName")["USDAmt"].sum().nlargest(10).reset_index()
        top_branches = top_branches.sort_values("USDAmt", ascending=True)
        
        fig_branches = px.bar(
            top_branches,
            x="USDAmt",
            y="BranchName",
            labels={"BranchName": "Branch", "USDAmt": "Revenue ($)"},
            text_auto=".2s"
        )
        fig_branches.update_layout(
            height=400,
            showlegend=False,
            xaxis_title="Revenue ($)",
            yaxis_title=""
        )
        st.plotly_chart(fig_branches, use_container_width=True)
    else:
        st.warning("⚠️ 'BranchName' column not found.")

# Chart 3: Top 10 TBMs (Territory Business Managers)
with chart_col3:
    st.write("##### Top 10 TBMs")
    if "TBM" in filtered_df.columns:
        top_tbms = filtered_df.groupby("TBM")["USDAmt"].sum().nlargest(10).reset_index()
        top_tbms = top_tbms.sort_values("USDAmt", ascending=True)
        
        fig_tbms = px.bar(
            top_tbms,
            x="USDAmt",
            y="TBM",
            labels={"TBM": "Territory Business Manager", "USDAmt": "Revenue ($)"},
            text_auto=".2s"
        )
        fig_tbms.update_layout(
            height=400,
            showlegend=False,
            xaxis_title="Revenue ($)",
            yaxis_title=""
        )
        st.plotly_chart(fig_tbms, use_container_width=True)
    else:
        st.warning("⚠️ 'TBM' column not found.")

# ─────────────────────────────────────────────────────────────────────────────
# RAW DATA VIEW
# ─────────────────────────────────────────────────────────────────────────────
st.divider()

with st.expander("📋 VIEW FILTERED DATASET", expanded=False):
    st.write(f"**Total Records: {len(filtered_df):,}**")
    st.dataframe(filtered_df, use_container_width=True, height=400)
    
    # Download button for CSV export
    csv_data = filtered_df.to_csv(index=False)
    st.download_button(
        label="📥 Download Filtered Data as CSV",
        data=csv_data,
        file_name="filtered_sales_data.csv",
        mime="text/csv"
    )

# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "💡 **Dashboard Tips:** "
    "Use sidebar filters to dynamically update all metrics and charts. "
    "Select/deselect columns to customize your analysis. "
    "Download filtered data for further analysis."
)
