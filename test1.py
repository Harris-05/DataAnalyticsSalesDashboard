import streamlit as st
import pandas as pd
from pyxlsb import open_workbook

st.set_page_config(page_title="Top 10 Sales Dashboard", layout="wide")

@st.cache_data
def load_data(file):
    rows = []
    try:
        with open_workbook(file) as wb:
            with wb.get_sheet('DATA_BK') as sheet:
                for row in sheet.rows():
                    rows.append([item.v for item in row])
        df = pd.DataFrame(rows)
        df.columns = df.iloc[0]
        df = df[1:]

        # Convert USDAmt to numeric so we can calculate the Top 10s
        if "USDAmt" in df.columns:
            df["USDAmt"] = pd.to_numeric(df["USDAmt"], errors="coerce")
        
        # Drop rows where there is no USD Amount
        df = df.dropna(subset=["USDAmt"])
        return df
    except Exception as e:
        st.error(f"Error loading file. Please ensure 'DATA_BK' sheet exists. Details: {e}")
        return pd.DataFrame()

# ---------------------------
# 1. File Upload
# ---------------------------
st.sidebar.title("📂 Data Upload")
file = st.sidebar.file_uploader("Upload your XLSB file", type=['xlsb'])

if not file:
    st.info("Please upload your data file to begin.")
    st.stop()

with st.spinner("Extracting DATA_BK..."):
    df = load_data(file)

if df.empty:
    st.error("No valid data found.")
    st.stop()

# ---------------------------
# 2. Dynamic "Slice & Dice" Filters
# ---------------------------
st.sidebar.header("🎛️ Filter Settings")
st.sidebar.write("**Step 1: Choose categories to filter by**")

# These are the core features you mentioned. We will suggest them by default if they exist in the file.
suggested_columns = ["Region", "BusinessSegment", "ClientSegment", "Product"]
default_selections = [col for col in suggested_columns if col in df.columns]

selected_filter_columns = st.sidebar.multiselect(
    "Select Columns:",
    options=df.columns.tolist(),
    default=default_selections,
    help="Pick the columns you want to filter. A new dropdown will appear for each one."
)

st.sidebar.write("---")
st.sidebar.write("**Step 2: Select specific values**")

filtered_df = df.copy()

# Generate a dropdown for every column the user picked in Step 1
for col in selected_filter_columns:
    # Get unique values, drop blanks, and sort them alphabetically
    unique_vals = sorted([str(x) for x in df[col].dropna().unique()])
    
    selected_vals = st.sidebar.multiselect(
        f"Filter by {col}:", 
        options=unique_vals, 
        default=unique_vals # Defaults to all selected so the charts don't start blank
    )
    
    # Apply the user's selection to the dataframe
    if selected_vals:
        filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

# ---------------------------
# 3. Main Dashboard (Top 10s)
# ---------------------------
st.title("🏆 Performance Dashboard")
st.write(f"**Currently analyzing {len(filtered_df):,} transactions based on your filters.**")

if filtered_df.empty:
    st.warning("No data matches your current filter combination. Try selecting more options.")
    st.stop()

st.divider()

# Set up 3 side-by-side columns for the charts
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Top 10 Clients")
    if "CustomerName" in filtered_df.columns:
        # Group by Customer, sum USD, take the top 10, and plot
        top_clients = filtered_df.groupby("CustomerName")["USDAmt"].sum().nlargest(10)
        st.bar_chart(top_clients)

with col2:
    st.subheader("Top 10 Branches")
    if "BranchName" in filtered_df.columns:
        # Group by Branch, sum USD, take the top 10, and plot
        top_branches = filtered_df.groupby("BranchName")["USDAmt"].sum().nlargest(10)
        st.bar_chart(top_branches)

with col3:
    st.subheader("Top 10 TBMs")
    if "TBM" in filtered_df.columns:
        # Group by TBM, sum USD, take the top 10, and plot
        top_tbms = filtered_df.groupby("TBM")["USDAmt"].sum().nlargest(10)
        st.bar_chart(top_tbms)

st.divider()

# ---------------------------
# 4. Filtered Raw Data View
# ---------------------------
with st.expander("📋 View Filtered Dataset"):
    st.dataframe(filtered_df, use_container_width=True)