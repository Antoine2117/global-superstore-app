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
st.markdown("Explore **Seasonality** and **Profitability** interactively!")

# ===== 3) Seasonality =====
st.header("📅 Seasonality of Sales & Profit")

# Year range selector
min_year, max_year = int(orders["Year"].min()), int(orders["Year"].max())
year_range = st.slider("Select Year Range:", min_year, max_year, (min_year, max_year))

filtered = orders[(orders["Year"] >= year_range[0]) & (orders["Year"] <= year_range[1])]

# Group by Month + Year
seasonality = filtered.groupby(["Year", orders["Order Date"].dt.month])["Sales"].sum().reset_index()
seasonality.rename(columns={"Order Date": "Month"}, inplace=True)
seasonality["Month"] = seasonality["Month"].astype(int)

# Line chart with different colors per year
fig1 = px.line(seasonality,
               x="Month", y="Sales", color="Year",
               markers=True,
               title="Monthly Sales Trends by Year",
               labels={"Sales": "Sales (USD)", "Month": "Month"})
fig1.update_layout(xaxis=dict(tickmode="linear", tick0=1, dtick=1))

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

