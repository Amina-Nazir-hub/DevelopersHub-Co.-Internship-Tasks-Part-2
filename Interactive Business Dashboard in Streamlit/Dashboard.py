# streamlit_dashboard.py

import pandas as pd
import streamlit as st
import plotly.express as px

# -------------------------------
# Step 1: Load and clean the data
# -------------------------------

@st.cache_data
def load_data():
    df = pd.read_excel('global_superstore_2016.xlsx')

    # Convert date columns to datetime
    df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')
    df['Ship Date'] = pd.to_datetime(df['Ship Date'], errors='coerce')

    # Drop rows with missing critical values
    df = df.dropna(subset=['Sales', 'Profit', 'Customer Name', 'Region', 'Category', 'Sub-Category'])

    # Ensure Sales and Profit are numeric
    df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce')
    df['Profit'] = pd.to_numeric(df['Profit'], errors='coerce')
    df = df.dropna(subset=['Sales', 'Profit'])

    # Create additional features
    df['Year'] = df['Order Date'].dt.year
    df['Month'] = df['Order Date'].dt.month

    return df

df = load_data()

# -------------------------------
# Step 2: Sidebar filters
# -------------------------------

st.sidebar.header("Filter Options")

regions = df['Region'].unique().tolist()
categories = df['Category'].unique().tolist()
sub_categories = df['Sub-Category'].unique().tolist()

selected_region = st.sidebar.multiselect('Select Region', regions, default=regions)
selected_category = st.sidebar.multiselect('Select Category', categories, default=categories)
selected_sub = st.sidebar.multiselect('Select Sub-Category', sub_categories, default=sub_categories)

# Filter data
filtered_df = df[
    (df['Region'].isin(selected_region)) &
    (df['Category'].isin(selected_category)) &
    (df['Sub-Category'].isin(selected_sub))
]

# -------------------------------
# Step 3: Display KPIs
# -------------------------------

st.title("Global Superstore Dashboard")

total_sales = filtered_df['Sales'].sum()
total_profit = filtered_df['Profit'].sum()

col1, col2 = st.columns(2)
col1.metric("Total Sales", f"${total_sales:,.0f}")
col2.metric("Total Profit", f"${total_profit:,.0f}")

# -------------------------------
# Step 4: Top 5 Customers
# -------------------------------

st.subheader("Top 5 Customers by Sales")
top_customers = (
    filtered_df.groupby('Customer Name')['Sales']
    .sum().sort_values(ascending=False).head(5).reset_index()
)
st.bar_chart(top_customers.rename(columns={'Sales': 'Sales Amount'}), x='Customer Name', y='Sales Amount')

# -------------------------------
# Step 5: Sales and Profit Visualizations
# -------------------------------

st.subheader("Sales by Region")
fig1 = px.bar(filtered_df, x='Region', y='Sales', color='Region', title='Sales by Region')
st.plotly_chart(fig1)

st.subheader("Profit by Category")
fig2 = px.bar(filtered_df, x='Category', y='Profit', color='Category', title='Profit by Category')
st.plotly_chart(fig2)

st.subheader("Sales Trend Over Time")
sales_trend = filtered_df.groupby('Order Date')['Sales'].sum().reset_index()
fig3 = px.line(sales_trend, x='Order Date', y='Sales', title='Daily Sales Trend')
st.plotly_chart(fig3)

st.subheader("Profit Trend Over Time")
profit_trend = filtered_df.groupby('Order Date')['Profit'].sum().reset_index()
fig4 = px.line(profit_trend, x='Order Date', y='Profit', title='Daily Profit Trend')
st.plotly_chart(fig4)
