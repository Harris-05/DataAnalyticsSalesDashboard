import streamlit as st
import pandas as pd
from pyxlsb import open_workbook

# ---------------------------
# Page Configuration
# ---------------------------
st.set_page_config(page_title="Sales Dashboard", layout="wide")

# ---------------------------
# XLSB Loader
# ---------------------------
@st.cache_data
def load_xlsb(file):
    rows = []
    try:
        with open_workbook(file) as wb:
            with wb.get_sheet('DATA_BK') as sheet:
                for row in sheet.rows():
                    rows.append([item.v for item in row])
        df = pd.DataFrame(rows)
        df.columns = df.iloc[0]  # set first row as header
        df = df[1:]  # remove header row
        return df
    except Exception as e:
        st.error(f"Error loading XLSB file: {e}")
        return pd.DataFrame()

# ---------------------------
# File Uploader
# ---------------------------
st.sidebar.header("Upload XLSB File")
xlsb_file = st.sidebar.file_uploader("Upload your XLSB file", type=["xlsb"])

if xlsb_file is None:
    st.warning("Please upload an XLSB file to continue.")
    st.stop()

# ---------------------------
# Load Data
# ---------------------------
df = load_xlsb(xlsb_file)

if df.empty:
    st.warning("No data loaded from XLSB.")
    st.stop()

# ---------------------------
# Data Cleaning & Date Parsing
# ---------------------------
df["DT"] = pd.to_numeric(df["DT"], errors="coerce")
df["USDAmt"] = pd.to_numeric(df["USDAmt"], errors="coerce")

# Convert Excel serial dates to datetime
if "DT" in df.columns:
    df["Date"] = pd.to_datetime(df["DT"], origin="1899-12-30", unit="D", errors="coerce")
else:
    st.error("DT column not found in XLSB sheet.")
    st.stop()

# Remove invalid rows
df = df.dropna(subset=["Date", "USDAmt"])

# Extract Year and Month
df["Year"] = df["Date"].dt.year
df["Month"] = df["Date"].dt.month_name()

# Month order
month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]
df["Month"] = pd.Categorical(df["Month"], categories=month_order, ordered=True)

# ---------------------------
# Sidebar Filters
# ---------------------------
st.sidebar.header("Filter Controls")

# Branch filter
branch_options = sorted(df["BranchName"].dropna().unique())
all_branches_selected = st.sidebar.checkbox("Select All Branches", value=True)
selected_branch = st.sidebar.multiselect(
    "Select Branch",
    options=branch_options,
    default=branch_options if all_branches_selected else None
)
df_branch_filtered = df[df["BranchName"].isin(selected_branch)]

# Month filter
all_months_selected = st.sidebar.checkbox("Select All Months", value=True)
selected_month = st.sidebar.multiselect(
    "Select Month",
    options=month_order,
    default=month_order if all_months_selected else None
)

# Client filter (for main dashboard)
client_options = sorted(df_branch_filtered["CustomerName"].dropna().unique())
all_clients_selected = st.sidebar.checkbox("Select All Clients (Dashboard)", value=True)
selected_client = st.sidebar.multiselect(
    "Select Client (Dashboard)",
    options=client_options,
    default=client_options if all_clients_selected else None
)

# TBM filter (for main dashboard)
if "TBM" in df.columns:
    tbm_options = sorted(df["TBM"].dropna().unique())
    all_tbms_selected = st.sidebar.checkbox("Select All TBMs (Dashboard)", value=True)
    selected_tbm = st.sidebar.multiselect(
        "Select TBM (Dashboard)",
        options=tbm_options,
        default=tbm_options if all_tbms_selected else None
    )
else:
    selected_tbm = df.get("TBM", pd.Series()).unique()  # fallback if TBM column not present

# Final filtered DataFrame for dashboard
filtered_df_dashboard = df_branch_filtered[
    df_branch_filtered["CustomerName"].isin(selected_client) &
    df_branch_filtered["Month"].isin(selected_month) &
    df_branch_filtered["TBM"].isin(selected_tbm)
]

# ---------------------------
# Quick Client View (independent)
# ---------------------------
st.sidebar.header("Quick Client View")
all_clients_full = sorted(df["CustomerName"].dropna().unique())
selected_quick_client = st.sidebar.selectbox(
    "Select Client to view details",
    options=["All Clients"] + all_clients_full
)

# ---------------------------
# Quick TBM View (independent)
# ---------------------------
if "TBM" in df.columns:
    all_tbms_full = sorted(df["TBM"].dropna().unique())
    selected_quick_tbm = st.sidebar.selectbox(
        "Select TBM to view details",
        options=["All TBMs"] + all_tbms_full
    )
else:
    selected_quick_tbm = "All TBMs"

# ---------------------------
# Apply Quick Client & Quick TBM Filters
# ---------------------------
filtered_df_quick = df.copy()

if selected_quick_client != "All Clients":
    filtered_df_quick = filtered_df_quick[filtered_df_quick["CustomerName"] == selected_quick_client]

if selected_quick_tbm != "All TBMs":
    filtered_df_quick = filtered_df_quick[filtered_df_quick["TBM"] == selected_quick_tbm]

# Also respect dashboard branch & month filters
filtered_df_quick = filtered_df_quick[
    filtered_df_quick["BranchName"].isin(selected_branch) &
    filtered_df_quick["Month"].isin(selected_month)
]

# ---------------------------
# Dashboard
# ---------------------------
st.title("📊 Sales Dashboard")
latest_year = int(df["Year"].max())
st.subheader(f"Data Overview for 2025")

if filtered_df_quick.empty:
    st.error("No data available for the selected filters or quick view.")
else:
    # KPIs
    total_sales = filtered_df_quick["USDAmt"].sum()
    total_transactions = len(filtered_df_quick)
    avg_sale = total_sales / total_transactions if total_transactions > 0 else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Sales", f"${total_sales:,.2f}")
    col2.metric("Transactions", f"{total_transactions:,}")
    col3.metric("Avg Sale Value", f"${avg_sale:,.2f}")

    st.divider()

    # Charts
    left_col, right_col = st.columns(2)

    with left_col:
        st.subheader("Monthly Sales Trend")
        monthly_sales = filtered_df_quick.groupby("Month")["USDAmt"].sum()
        st.line_chart(monthly_sales)

        st.subheader("Sales by Branch")
        branch_sales = filtered_df_quick.groupby("BranchName")["USDAmt"].sum().sort_values(ascending=True)
        st.bar_chart(branch_sales, horizontal=True)

    with right_col:
        st.subheader("Top 10 Clients")
        top_clients = filtered_df_quick.groupby("CustomerName")["USDAmt"].sum().sort_values(ascending=False).head(10)
        st.bar_chart(top_clients)

        st.subheader("🏆 Leading Branch")
        if not filtered_df_quick.empty:
            top_b = filtered_df_quick.groupby("BranchName")["USDAmt"].sum().idxmax()
            top_val = filtered_df_quick.groupby("BranchName")["USDAmt"].sum().max()
            st.success(f"**{top_b}** is the top performer with **${top_val:,.2f}** in sales.")

    # ---------------------------
    # Client Summary Table
    # ---------------------------
    st.subheader("📋 Client Summary")
    if selected_quick_client != "All Clients":
        client_summary = filtered_df_quick.groupby("Month").agg(
            Total_Sales=("USDAmt", "sum"),
            Transactions=("USDAmt", "count")
        ).reset_index()
        st.table(client_summary)

        branch_summary = filtered_df_quick.groupby("BranchName").agg(
            Total_Sales=("USDAmt", "sum"),
            Transactions=("USDAmt", "count")
        ).reset_index().sort_values("Total_Sales", ascending=False)
        st.table(branch_summary)

    # ---------------------------
    # Raw Data
    # ---------------------------
    with st.expander("View Filtered Raw Data"):
        st.dataframe(filtered_df_quick.sort_values("Date", ascending=False), use_container_width=True)
