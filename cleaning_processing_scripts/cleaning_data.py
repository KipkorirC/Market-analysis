import pandas as pd
import os
import opendatasets as od



pwd = os.getcwd()
dataset_url = 'https://www.kaggle.com/rodsaldanha/arketing-campaign' 
od.download(dataset_url)
data_dir = './arketing-campaign'
pwd = os.getcwd()
df = pd.read_csv(pwd+"/arketing-campaign/marketing_campaign.csv", sep = ';')

# These columns are equal for all customers and will not be useful in comparison of the customers
LeastUseful_columns = ["Z_CostContact","Z_Revenue","AcceptedCmp1","AcceptedCmp2","AcceptedCmp3","AcceptedCmp4","AcceptedCmp5","Complain","Response"]

df['Income'].fillna(0,inplace = True)



df["Current_year"] = 2023
df["Age"] = df["Current_year"] - df["Year_Birth"]



df.drop(columns = ["Current_year"],inplace = True)
purchases = list(df.columns[9:15])
df['Total_purchases']= df.apply(lambda x: x[purchases[0]]+ x[purchases[1]]+x[purchases[2]]+ x[purchases[3]]+x[purchases[4]]+ x[purchases[5]], axis = 1)


Accepted = list(df.columns[20:25])
Accepted
df["Accepted_complaints"] = df.apply(lambda x: x[Accepted[0]]+x[Accepted[1]]+x[Accepted[2]]+x[Accepted[3]]+x[Accepted[4]], axis = 1)

df["Children"] = df["Kidhome"] + df["Teenhome"]


#dropping the excess columns that we will no longer need
columns_not_needed = ['Year_Birth','Dt_Customer','Recency',"AcceptedCmp1","AcceptedCmp2","AcceptedCmp3","AcceptedCmp4","AcceptedCmp5","Z_CostContact","Z_Revenue","Accepted_complaints","Complain","Response"]
data = df.copy()
cols = list(data.columns[0:9])
cols2 =  list(data.columns[15:])

cols.extend(cols2)
data.drop(columns=cols,inplace=True)
max_data = data.idxmax(axis=1)
df["favourite_product"]=max_data

df.drop(columns= columns_not_needed,inplace = True)


d= {"Product":["MntWines","MntFruits","MntMeatProducts","MntFishProducts", "MntSweetProducts","MntGoldProds" ], "Total" : [df["MntWines"].sum(),df["MntFruits"].sum(), df["MntMeatProducts"].sum(),df["MntFishProducts"].sum(),df["MntSweetProducts"].sum(),df["MntGoldProds"].sum()]}


product_by_totals = pd.DataFrame(data=d)

with pd.ExcelWriter("./clean_data/clean_first_dashboard_data.xlsx") as writer1:
    df.to_excel(writer1, sheet_name="Sheet1",index=False)
    product_by_totals.to_excel(writer1,sheet_name="Sheet2",index=False)


#df_products_by_totals = pd.DataFrame[]