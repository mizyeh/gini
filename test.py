import streamlit as st
import pandas as pd
import plotly.express as px

# 初始化儀錶板
st.title("療程銷售儀錶板")

# 上傳 Excel 檔案
uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.rename(
        columns={"Purchase Date 年/月/日": "Purchase_Date", "療程大類": "Category", "館別": "Location", "產品名稱(群組)": "Product",
                 "購買數量": "Quantity"}, inplace=True)
    df["Purchase_Date"] = pd.to_datetime(df["Purchase_Date"], unit="D", origin="1899-12-30")
    df["Quantity"].fillna(0, inplace=True)

    # 篩選條件
    min_date, max_date = df["Purchase_Date"].min(), df["Purchase_Date"].max()
    start_date, end_date = st.date_input("選擇時間區間", [min_date, max_date], min_value=min_date, max_value=max_date)
    categories = st.multiselect("選擇療程類別", df["Category"].unique(), default=df["Category"].unique())
    locations = st.multiselect("選擇館別", df["Location"].unique(), default=df["Location"].unique())

    # 篩選數據
    filtered_df = df[(df["Purchase_Date"] >= pd.to_datetime(start_date)) &
                     (df["Purchase_Date"] <= pd.to_datetime(end_date)) &
                     (df["Category"].isin(categories)) &
                     (df["Location"].isin(locations))]

    # KPI 指標
    total_sales = filtered_df["Quantity"].sum()
    avg_sales = filtered_df["Quantity"].mean()
    st.metric("總銷售數", total_sales)
    st.metric("平均購買量", round(avg_sales, 2))

    # 視覺化圖表
    sales_trend = filtered_df.groupby("Purchase_Date")["Quantity"].sum().reset_index()
    fig_line = px.line(sales_trend, x="Purchase_Date", y="Quantity", title="銷售趨勢")
    st.plotly_chart(fig_line)

    sales_by_category = filtered_df.groupby("Category")["Quantity"].sum().reset_index()
    fig_bar = px.bar(sales_by_category, x="Category", y="Quantity", title="各療程銷售量")
    st.plotly_chart(fig_bar)

    sales_by_location = filtered_df.groupby("Location")["Quantity"].sum().reset_index()
    fig_pie = px.pie(sales_by_location, names="Location", values="Quantity", title="各館銷售佔比")
    st.plotly_chart(fig_pie)