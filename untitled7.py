import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Multi-Select for Regions
all_regions = sorted(df_original["Region"].dropna().unique())
selected_regions = st.sidebar.multiselect("Select Region(s)", options=all_regions, default=all_regions)

# Filter data by selected regions
df_filtered = df_original[df_original["Region"].isin(selected_regions)]

# Multi-Select for Categories
all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_categories = st.sidebar.multiselect("Select Category(s)", options=all_categories, default=all_categories)

# Apply category filter
df_filtered = df_filtered[df_filtered["Category"].isin(selected_categories)]

# ---- Sidebar Date Range ----
min_date = df_filtered["Order Date"].min()
max_date = df_filtered["Order Date"].max()

from_date = st.sidebar.date_input("From Date", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To Date", value=max_date, min_value=min_date, max_value=max_date)

if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

df_filtered = df_filtered[
    (df_filtered["Order Date"] >= pd.to_datetime(from_date)) & 
    (df_filtered["Order Date"] <= pd.to_datetime(to_date))
]

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# ---- KPI Calculation ----
if df_filtered.empty:
    total_sales = total_quantity = total_profit = 0
    margin_rate = 0
else:
    total_sales = df_filtered["Sales"].sum()
    total_quantity = df_filtered["Quantity"].sum()
    total_profit = df_filtered["Profit"].sum()
    margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# ---- KPI Display ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
kpi_col1.metric("Total Sales", f"${total_sales:,.2f}")
kpi_col2.metric("Total Quantity", f"{total_quantity:,.0f}")
kpi_col3.metric("Total Profit", f"${total_profit:,.2f}")
kpi_col4.metric("Margin Rate", f"{(margin_rate * 100):,.2f}%")

# ---- KPI Selection ----
st.subheader("Visualizing KPI Trends & Top Products")

if df_filtered.empty:
    st.warning("No data available for the selected filters.")
else:
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI:", options=kpi_options, horizontal=True)

    # ---- Prepare Data for Charts ----
    daily_grouped = df_filtered.groupby("Order Date").agg({
        "Sales": "sum", "Quantity": "sum", "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)
    
    # **NEW: Add 7-Day Moving Average Line**
    daily_grouped["Moving Avg"] = daily_grouped[selected_kpi].rolling(window=7).mean()

    product_grouped = df_filtered.groupby("Product Name").agg({
        "Sales": "sum", "Quantity": "sum", "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)

    top_10 = product_grouped.sort_values(by=selected_kpi, ascending=False).head(10)

    # ---- Side-by-Side Layout ----
    col_left, col_right = st.columns(2)

    with col_left:
        fig_line = px.line(
            daily_grouped, x="Order Date", y=selected_kpi, 
            title=f"{selected_kpi} Over Time (with 7-Day Avg)",
            labels={"Order Date": "Date", selected_kpi: selected_kpi},
            template="plotly_white"
        )
        fig_line.add_scatter(
            x=daily_grouped["Order Date"], y=daily_grouped["Moving Avg"],
            mode="lines", name="7-Day Moving Avg", line=dict(dash="dot", color="red")
        )
        fig_line.update_layout(height=400)
        st.plotly_chart(fig_line, use_container_width=True)

    with col_right:
        fig_bar = px.bar(
            top_10, x=selected_kpi, y="Product Name", orientation="h",
            title=f"Top 10 Products by {selected_kpi}", labels={selected_kpi: selected_kpi},
            color=selected_kpi, color_continuous_scale="Blues", template="plotly_white"
        )
        fig_bar.update_layout(height=400, yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig_bar, use_container_width=True)
