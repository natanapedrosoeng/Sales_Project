import pandas as pd
import numpy as np
import hashlib
import os
import glob

# Path to read the files
sources_path = 'sources\\data\\raw'

#To list all files
files = glob.glob(os.path.join(sources_path, '*.xlsx'))

# Files
df_2013_2014 = pd.read_excel(r'sources\\data\\raw\\Base_de_dados_2013_2014_(3).xlsx')
df_2015_2016 = pd.read_excel(r'sources\\data\\raw\\Base_de_dados_2015_2016_(3).xlsx')
df_avg = pd.read_excel(r'sources\\data\\raw\\Dim_Avg_Price_2013_2016.xlsx')


# # SOURCES: Base_de_dados_2013_2014_(3) E Base_de_dados_2015_2016_(3) (ORDER DETAIL TABLES) 
# Split columns
df_2013_2014[['Country', 'Product']] = df_2013_2014['Country,Product'].str.split(',', expand=True)
            
# Order columns
df_2013_2014 = df_2013_2014[['Segment', 'Country', 'Product', 'Discount Band'
                    , 'Units Sold', 'Manufacturing Price', 'Sale Price'
                    , 'Gross Sales', 'Discounts', ' Sales', 'COGS', 'Profit'
                    , 'Date', 'Month Number', 'Month Name', 'Year'
                    , 'Product_Name']]

# Union DataFrames
df_consolidated = pd.concat([df_2013_2014, df_2015_2016])

# Renaming columns
consolidated_rename = {'Discount Band': 'DiscountBand', 'Units Sold': 'UnitsSold', 'Manufacturing Price': 'ManufacturingPrice'
                          , 'Sale Price': 'SalePrice', 'Gross Sales': 'GrossSales', ' Sales': 'Sales', 'Month Number': 'MonthNumber'
                          , 'Month Name': 'MonthName', 'Product_Name': 'ProductName'}
df_consolidated = df_consolidated.rename(columns=consolidated_rename)

# Fixing nulls in DiscountBand column
df_consolidated['DiscountBand'] = df_consolidated['DiscountBand'].fillna("Without Discount")

# Fixing the wrong words in Segment column
df_consolidated['Segment'] = df_consolidated['Segment'].replace({'Governmemt': 'Government', 'Chanel Partners': 'Channel Partners'
                                                                 , 'Enter&rise': 'Enterprise', 'Enterrise': 'Enterprise'
                                                                 , 'Smal Business': 'Small Business'})


# Creating Primary Key
#df_consolidated['PrimaryKey'] = df_consolidated.apply(lambda row: row['Segment'] + row['Country'] + row['Product'] + row['Year'], axis=1)
df_consolidated['PrimaryKey'] = df_consolidated['Segment'] + df_consolidated['Country'] + df_consolidated['Product'] + df_consolidated['Year'].astype(str)

# Creating Surrogate Key
df_consolidated['SurrogateKey'] = df_consolidated['PrimaryKey'].apply(lambda x: hashlib.md5(x.encode()).hexdigest())

#Save DataFrame df_consolidated in excel file
consolidated_name = 'sources\\data\\ready\\fact_order_details.xlsx'

# Create an object ExcelWriter to save Dataframe as a excel table
with pd.ExcelWriter(consolidated_name, engine='openpyxl', mode='w') as writer:
    df_consolidated.to_excel(writer, sheet_name='fact_order_details', index=False)

    # Cash excel table
    worksheet = writer.sheets['fact_order_details']


# SOURCE: Dim_Avg_Price_2013_2016 (AVG PRICE TABLE)
avg_rename = {'AVG Price': 'AvgPrice'}
df_avg = df_avg.rename(columns=avg_rename)

# Average Price
df_avg['AvgPriceNew'] = df_avg.groupby(['Segment', 'Country', 'Product', 'Year'])['AvgPrice'].transform('mean')
df_avg = df_avg.drop_duplicates(subset=['Segment', 'Country', 'Product', 'Year'], keep='first')
df_avg = df_avg.drop('AvgPrice', axis=1)

# Creating Primary Key in Dim_Avg
df_avg['PrimaryKey'] = df_avg['Segment'] + df_avg['Country'] + df_avg['Product'] + df_avg['Year'].astype(str)

#Creating Foreing Key
df_avg['ForeingKey'] = df_avg['PrimaryKey'].apply(lambda x: hashlib.md5(x.encode()).hexdigest())

#Save DataFrame df_avgd in excel file
avg_name = 'sources\\data\\ready\\dim_avg_price.xlsx'

# Create an object ExcelWriter to save Dataframe as a excel table
with pd.ExcelWriter(avg_name, engine='openpyxl', mode='w') as writer:
    df_avg.to_excel(writer, sheet_name='dim_avg_price', index=False)

    # Cash excel table
    worksheet = writer.sheets['dim_avg_price']
