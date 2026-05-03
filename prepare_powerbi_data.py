"""
Prepare clean, analysis-ready data for Power BI import.
Creates a structured Excel workbook with:
  - Fact table (Orders)
  - Dimension tables (Date, Geography, Product, Customer)
  - Pre-calculated summary tables for quick visuals
"""
import pandas as pd
import numpy as np
import os

BASE = os.path.dirname(__file__)
df = pd.read_excel(os.path.join(BASE, 'DMart Data Store.xlsx'), sheet_name='Order Data')

# =============================================
# 1. CLEAN & ENRICH THE FACT TABLE
# =============================================
df['Year'] = df['Order Date'].dt.year
df['Quarter'] = 'Q' + df['Order Date'].dt.quarter.astype(str)
df['Month'] = df['Order Date'].dt.month_name()
df['Month_Num'] = df['Order Date'].dt.month
df['Profit Margin %'] = round(df['Profit'] / df['Sales'] * 100, 2)
df['Discount Band'] = pd.cut(
    df['Discount'],
    bins=[-0.01, 0, 0.1, 0.2, 0.3, 1.0],
    labels=['No Discount', '1-10%', '11-20%', '21-30%', '30%+']
)
df['Revenue Band'] = pd.cut(
    df['Sales'],
    bins=[0, 50, 150, 500, 10000],
    labels=['Low (<50)', 'Medium (50-150)', 'High (150-500)', 'Premium (500+)']
)

# =============================================
# 2. CREATE DIMENSION TABLES
# =============================================

# Date Dimension
dates = pd.DataFrame({'Date': pd.date_range(df['Order Date'].min(), df['Order Date'].max())})
dates['Year'] = dates['Date'].dt.year
dates['Quarter'] = 'Q' + dates['Date'].dt.quarter.astype(str)
dates['Month'] = dates['Date'].dt.month_name()
dates['Month_Num'] = dates['Date'].dt.month
dates['Year-Quarter'] = dates['Year'].astype(str) + '-' + dates['Quarter']
dates['Year-Month'] = dates['Date'].dt.strftime('%Y-%m')

# Geography Dimension
geo = df[['Country', 'State', 'City', 'Region']].drop_duplicates().sort_values(['Region', 'Country', 'State', 'City'])
geo = geo.reset_index(drop=True)
geo.index.name = 'Geo_ID'

# Product Dimension
prod = df[['Category', 'Sub-Category', 'Product Name']].drop_duplicates().sort_values(['Category', 'Sub-Category', 'Product Name'])
prod = prod.reset_index(drop=True)
prod.index.name = 'Product_ID'

# Customer Dimension
cust = df.groupby('Customer Name').agg(
    Segment=('Segment', 'first'),
    Total_Orders=('Order ID', 'nunique'),
    Total_Sales=('Sales', 'sum'),
    Total_Profit=('Profit', 'sum'),
    First_Order=('Order Date', 'min'),
    Last_Order=('Order Date', 'max'),
    Primary_Country=('Country', lambda x: x.mode().iloc[0]),
    Primary_City=('City', lambda x: x.mode().iloc[0])
).reset_index()
cust['Avg_Order_Value'] = round(cust['Total_Sales'] / cust['Total_Orders'], 2)
cust['Profit_Margin_%'] = round(cust['Total_Profit'] / cust['Total_Sales'] * 100, 2)
cust['Customer_Tier'] = pd.cut(
    cust['Total_Sales'],
    bins=[0, 2000, 5000, 10000, 100000],
    labels=['Bronze', 'Silver', 'Gold', 'Platinum']
)

# =============================================
# 3. SUMMARY TABLES FOR POWER BI
# =============================================

# Segment Summary
seg_summary = df.groupby('Segment').agg(
    Total_Sales=('Sales', 'sum'),
    Total_Profit=('Profit', 'sum'),
    Total_Orders=('Order ID', 'nunique'),
    Total_Customers=('Customer Name', 'nunique'),
    Avg_Discount=('Discount', 'mean'),
    Total_Quantity=('Quantity', 'sum')
).reset_index()
seg_summary['Sales_Share_%'] = round(seg_summary['Total_Sales'] / seg_summary['Total_Sales'].sum() * 100, 1)
seg_summary['Profit_Margin_%'] = round(seg_summary['Total_Profit'] / seg_summary['Total_Sales'] * 100, 1)

# Country Summary
country_summary = df.groupby(['Country', 'Region']).agg(
    Total_Sales=('Sales', 'sum'),
    Total_Profit=('Profit', 'sum'),
    Total_Orders=('Order ID', 'nunique'),
    Total_Customers=('Customer Name', 'nunique')
).reset_index()
country_summary['Sales_Share_%'] = round(country_summary['Total_Sales'] / country_summary['Total_Sales'].sum() * 100, 1)
country_summary['Profit_Margin_%'] = round(country_summary['Total_Profit'] / country_summary['Total_Sales'] * 100, 1)

# SubCategory Summary
subcat_summary = df.groupby(['Category', 'Sub-Category']).agg(
    Total_Sales=('Sales', 'sum'),
    Total_Profit=('Profit', 'sum'),
    Total_Quantity=('Quantity', 'sum'),
    Avg_Discount=('Discount', 'mean')
).reset_index()
subcat_summary['Profit_Margin_%'] = round(subcat_summary['Total_Profit'] / subcat_summary['Total_Sales'] * 100, 1)
subcat_summary['Avg_Unit_Price'] = round(subcat_summary['Total_Sales'] / subcat_summary['Total_Quantity'], 0)

# Yearly Trend
yearly_trend = df.groupby('Year').agg(
    Total_Sales=('Sales', 'sum'),
    Total_Profit=('Profit', 'sum'),
    Total_Orders=('Order ID', 'nunique'),
    Total_Customers=('Customer Name', 'nunique')
).reset_index()
yearly_trend['YoY_Growth_%'] = round(yearly_trend['Total_Sales'].pct_change() * 100, 1)
yearly_trend['Profit_Margin_%'] = round(yearly_trend['Total_Profit'] / yearly_trend['Total_Sales'] * 100, 1)

# Segment x Region Matrix
seg_region = df.pivot_table(values='Sales', index='Region', columns='Segment', aggfunc='sum', fill_value=0).reset_index()

# Segment x Country Matrix
seg_country = df.pivot_table(values='Sales', index='Country', columns='Segment', aggfunc='sum', fill_value=0).reset_index()
seg_country['Total'] = seg_country[['Consumer', 'Corporate', 'Home Office']].sum(axis=1)
seg_country['Corporate_%'] = round(seg_country['Corporate'] / seg_country['Total'] * 100, 1)
seg_country = seg_country.sort_values('Total', ascending=False)

# =============================================
# 4. EXPORT TO EXCEL FOR POWER BI
# =============================================
output_path = os.path.join(BASE, 'DMart_PowerBI_Data.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Fact_Orders', index=False)
    dates.to_excel(writer, sheet_name='Dim_Date', index=False)
    geo.to_excel(writer, sheet_name='Dim_Geography')
    prod.to_excel(writer, sheet_name='Dim_Product')
    cust.to_excel(writer, sheet_name='Dim_Customer', index=False)
    seg_summary.to_excel(writer, sheet_name='Summary_Segment', index=False)
    country_summary.to_excel(writer, sheet_name='Summary_Country', index=False)
    subcat_summary.to_excel(writer, sheet_name='Summary_SubCategory', index=False)
    yearly_trend.to_excel(writer, sheet_name='Summary_YearlyTrend', index=False)
    seg_country.to_excel(writer, sheet_name='Summary_SegmentByCountry', index=False)

print(f"Power BI data file created: {output_path}")
print(f"\nSheets created:")
print(f"  - Fact_Orders         ({len(df)} rows) - Main fact table")
print(f"  - Dim_Date            ({len(dates)} rows) - Date dimension")
print(f"  - Dim_Geography       ({len(geo)} rows) - Geography dimension")
print(f"  - Dim_Product         ({len(prod)} rows) - Product dimension")
print(f"  - Dim_Customer        ({len(cust)} rows) - Customer dimension")
print(f"  - Summary_Segment     ({len(seg_summary)} rows)")
print(f"  - Summary_Country     ({len(country_summary)} rows)")
print(f"  - Summary_SubCategory ({len(subcat_summary)} rows)")
print(f"  - Summary_YearlyTrend ({len(yearly_trend)} rows)")
print(f"  - Summary_SegmentByCountry ({len(seg_country)} rows)")
print(f"\nImport 'DMart_PowerBI_Data.xlsx' into Power BI Desktop to start building your dashboard!")
