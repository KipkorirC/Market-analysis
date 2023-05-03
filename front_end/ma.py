import streamlit as st
import pandas as pd
import plotly.express as px
import os



st.set_page_config(page_title = "sales data",
                   page_icon = ":bar_chart:",
                   layout = "wide")
pwd = os.getcwd()
df = pd.read_excel(pwd+"/clean_data/clean_first_dashboard_data.xlsx")






st.title(":bar_chart: Sales Dashboard")



total_sales = int (df["Total_purchases"].sum())
total_number_of_customers = (df["ID"].count())
average_sales_per_customer = round(total_sales/total_number_of_customers)
total_web_visits = df["NumWebVisitsMonth"].sum()
total_web_purchases = df["NumWebPurchases"].sum()
average_purchases_after_web_visits = round(total_web_purchases/total_web_visits)
average_income = round(df["Income"].mean())
left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total Sales: ")
    st.subheader(total_sales)
with middle_column:
    st.subheader("Average sales per customer: ")
    st.subheader(average_sales_per_customer)
with right_column:
    st.subheader("Average income for customers: ")
    st.subheader(f"US {average_income:,}")


st.markdown("---")
d= {"Product":["MntWines","MntFruits","MntMeatProducts","MntFishProducts", "MntSweetProducts","MntGoldProds" ], "Total" : [df["MntWines"].sum(),df["MntFruits"].sum(), df["MntMeatProducts"].sum(),df["MntFishProducts"].sum(),df["MntSweetProducts"].sum(),df["MntGoldProds"].sum()]}


product_by_totals = pd.DataFrame(data=d)
#bar charts
Education_analysis = df.groupby('Education')['ID'].count().reset_index()

## BAR CHART 1

fig_product_sales = px.bar(
    product_by_totals,
    x= "Total",
    y= "Product",
    orientation="h",
    title="<b>Product by totals</b>",
    color_discrete_sequence=["#0083B8"]*len(product_by_totals),
    template="plotly_white"
)



web_data = pd.DataFrame({"cols":["web visits","webpurchases"],"Total":[total_web_visits,total_web_purchases]})
fig_Web_visits_vs_web_purchases = px.bar(
    web_data,
    x= "cols",
    y= "Total",
    orientation="v",
    title="<b>Web stats</b>",
    color_discrete_sequence=["#0083B8"]*len(web_data),
    template="plotly_white"
)
fig = px.histogram(df, x="Age",nbins=128)
left_column ,middle_column,right_column = st.columns(3)
left_column.plotly_chart(fig_product_sales, use_container_width=True)
with middle_column :
    st.subheader("Age Distribution")
    middle_column.plotly_chart(fig,use_container_width=True)
right_column.plotly_chart(fig_Web_visits_vs_web_purchases,use_container_width=True)


Education_analysis = df.groupby('Education')['ID'].count().reset_index()
pie_chart = px.pie(Education_analysis,
                    title = "Customers Education Distribution",
                    values = 'ID',
                    names = 'Education')

total_deal_purchases = df["NumDealsPurchases"].sum()
total_catalog_purchases =df["NumCatalogPurchases"].sum()
total_store_purchases = df["NumStorePurchases"].sum()             
Purchase_type_analysis= pd.DataFrame({"cols":["Web Purchases","Deal Purchases","Catalog Purchases","Store Purchases"],"Total":[total_web_purchases,total_deal_purchases,total_catalog_purchases,total_store_purchases]})

pie_chart2 = px.pie(Purchase_type_analysis,
                    title = "Customers Purchase Type Distribution",
                    values = 'Total',
                    names = 'cols')
marital_status_data =df.groupby("Marital_Status")["ID"].count().reset_index()
fig_marital_status = px.bar(
    marital_status_data,
    x= "ID",
    y= "Marital_Status",
    orientation="h",
    title="<b>Marital Status of customers</b>",
    color_discrete_sequence=["#0083B8"]*len(marital_status_data),
    template="plotly_white"
)

left_column,middle_column,right_column = st.columns(3)

left_column.plotly_chart(pie_chart,use_container_width=True)
middle_column.plotly_chart(fig_marital_status,use_container_width=True)

right_column.plotly_chart(pie_chart2,use_container_width=True)