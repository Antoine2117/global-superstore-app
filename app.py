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

# Year range slider
min_year, max_year = int(orders["Year"].min()), int(orders["Year"].max())
year_range = st.slider("Select Year Range:", min_year, max_year, (min_year, max_year))

# Region filter
regions = orders["Region"].unique().tolist()
selected_regions = st.multiselect("Select Region(s):", regions, default=regions)

# Filter data
filtered = orders[
    (orders["Year"] >= year_range[0]) & 
    (orders["Year"] <= year_range[1]) & 
    (orders["Region"].isin(selected_regions))
]

# Group by Year + Month
seasonality = filtered.groupby(["Year", orders["Order Date"].dt.month])["Sales"].sum().reset_index()
seasonality.rename(columns={"Order Date": "Month"}, inplace=True)
seasonality["Month"] = seasonality["Month"].astype(int)

# Plot with different colors per year
fig1 = px.line(seasonality,
               x="Month", y="Sales", color="Year",
               markers=True,
               title=f"Monthly Sales Trends ({', '.join(selected_regions)})",
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

# --- Dynamic insights for seasonality ---
if not seasonality.empty:
    month_map = {1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"May", 6:"Jun",
                 7:"Jul", 8:"Aug", 9:"Sep", 10:"Oct", 11:"Nov", 12:"Dec"}
    
    monthly_sales = seasonality.groupby("Month")["Sales"].sum()
    peak_month = month_map[int(monthly_sales.idxmax())]
    peak_value = int(monthly_sales.max())
    low_month = month_map[int(monthly_sales.idxmin())]
    low_value = int(monthly_sales.min())
    
    years_shown = f"{seasonality['Year'].min()}–{seasonality['Year'].max()}"
    trend_text = f"Between {years_shown}, peak sales were in {peak_month} (${peak_value:,}), while the lowest were in {low_month} (${low_value:,})."
else:
    trend_text = "No sales data available for the current selection."

# --- Dynamic insights for profitability ---
if not profit_data.empty:
    top_item = profit_data.iloc[0]
    bottom_item = profit_data.iloc[-1]
    top_name, top_value = str(top_item[0]), int(top_item[1])
    bottom_name, bottom_value = str(bottom_item[0]), int(bottom_item[1])
    
    profit_text = f"Top performer: {top_name} (${top_value:,}). Biggest loss maker: {bottom_name} (${bottom_value:,})."
else:
    profit_text = "No profitability data available for the current selection."

# --- Show results using individual st.write statements ---
st.write("• **Seasonality:**", trend_text)
st.write("• **Profitability:**", profit_text) 
st.write("• **Strategic Use:** Prepare for peaks and address products/regions dragging profits down.")

