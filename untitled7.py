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

# Apply categor
