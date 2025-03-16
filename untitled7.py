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

all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

df_filtered = df_original if selected_region == "All" else df_original[df_original["Region"] == selected_region]

all_states = sorted(df_filtered["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

df_filtered = df_filtered if selected_state == "All" else df_filtered[df_filtered["State"] == selected_state]

all_categories = sorted(df_filtered["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

df_filtered = df_filtered if selected_category == "All" else df_filtered[df_filtered["Category"] == selected_category]

all_subcats = sorted(df_filtered["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

df_filtered = df_filtered if selected_subcat == "All" else df_filtered[df_filtered["Sub-Category"] == selected_subcat]

# ---- Sidebar Date Range ----
min_date, max_date = df_filtered["Order Date"].min(), df_filtered["Order Date"].max()
from_date = st.sidebar.date_input("From Date", value=min_date, min_value=min_date, max_value=max_date)
to_date = st.sidebar.date_input("To Date", value=max_date, min_value=min_date, max_value=max_date)

df_filtered = df_filtered[(df_filtered["Order Date"] >= pd.to_datetime(from_date)) & (df_filtered["Order Date"] <= pd.to_datetime(to_date))]

# ---- KPI Calculation ----
total_sales = df_filtered["Sales"].sum() if not df_filtered.empty else 0
total_quantity = df_filtered["Quantity"].sum() if not df_filtered.empty else 0
total_profit = df_filtered["Profit"].sum() if not df_filtered.empty else 0
margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# ---- KPI Display ----
st.title("SuperStore KPI Dashboard")

kpi_cols = st.columns(4)
kpi_labels = ["Sales", "Quantity Sold", "Profit", "Margin Rate"]
kpi_values = [f"${total_sales:,.2f}", f"{total_quantity:,}", f"${total_profit:,.2f}", f"{(margin_rate * 100):,.2f}%"]

for col, label, value in zip(kpi_cols, kpi_labels, kpi_values):
    col.metric(label=label, value=value)

# ---- KPI Visualization ----
if not df_filtered.empty:
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)
    
    daily_grouped = df_filtered.groupby("Order Date").agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)
    
    product_grouped = df_filtered.groupby("Product Name").agg({"Sales": "sum", "Quantity": "sum", "Profit": "sum"}).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    top_10 = product_grouped.sort_values(by=selected_kpi, ascending=False).head(10)
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        fig_line = px.line(daily_grouped, x="Order Date", y=selected_kpi, title=f"{selected_kpi} Over Time", template="plotly_white")
        st.plotly_chart(fig_line, use_container_width=True)
    
    with col_right:
        fig_bar = px.bar(top_10, x=selected_kpi, y="Product Name", orientation="h", title=f"Top 10 Products by {selected_kpi}", template="plotly_white")
        st.plotly_chart(fig_bar, use_container_width=True)
