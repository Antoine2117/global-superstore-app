import streamlit as st
import pandas as pd
import plotly.express as px

# ===== 1) Load Data =====
file_path = "Global Superstore.xls"

orders = pd.read_excel(file_path, sheet_name="Orders")
returns = pd.read_excel(file_path, sheet_name="Returns")
people = pd.read_excel(file_path, sheet_name="People")

orders = orders.merge(returns, on="Order ID", how="left")
orders = orders.merge(people, on="Region", how="left")

orders["Order Date"] = pd.to_datetime(orders["Order Date"])
orders["Year"] = orders["Order Date"].dt.year
orders["Month"] = orders["Order Date"].dt.to_period("M")

# ===== 2) Page Config =====
st.set_page_config(page_title="Global Superstore Insights", layout="wide")
st.title("📊 Global Superstore Insights Dashboard")
st.markdown("Explore **Seasonality** and **Profitability** interactively with Streamlit and Plotly.")

# ===== 3) Seasonality =====
st.header("📅 Seasonality of Sales & Profit")
years = st.multiselect("Select Year(s):", sorted(orders["Year"].unique()), default=sorted(orders["Year"].unique()))
filtered = orders[orders["Year"].isin(years)]

monthly = filtered.groupby("Month")[["Sales","Profit"]].sum().reset_index()
monthly["Month"] = monthly["Month"].astype(str)

fig1 = px.line(monthly, x="Month", y=["Sales","Profit"],
               title="Monthly Sales & Profit Trends",
               labels={"value":"USD", "variable":"Metric"},
               markers=True)
st.plotly_chart(fig1, use_container_width=True)

# ===== 4) Profitability =====
st.header("💰 Profitability Breakdown")
level = st.radio("Choose level of detail:", ["Category", "Sub-Category", "Product Name"])

profit_data = orders.groupby(level)["Profit"].sum().sort_values(ascending=False).reset_index()

fig2 = px.bar(profit_data, x="Profit", y=level,
              color="Profit",
              color_continuous_scale="Tealgrn",
              title=f"Profit by {level}",
              orientation="h")
st.plotly_chart(fig2, use_container_width=True)

# ===== 5) Insights =====
st.subheader("🔎 Key Insights")
st.markdown("""
- **Seasonality:** Sales show strong peaks in certain months.
- **Profitability:** Toggle between Category, Sub-Category, or Product.
- **Strategic Use:** Spot loss leaders vs top performers.
""")
