
import pandas as pd
#load Left Hand Side (LHS) Data
order_report = r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Order Report.xlsx"
pincode_zones=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Pincode Zones.xlsx"
sku_master=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - SKU Master.xlsx"

order_report_data = pd.read_excel(order_report)
pincode_zones_data= pd.read_excel(pincode_zones)
sku_master_data= pd.read_excel(sku_master)
# print(order_report_data)
# print(pincode_zones_data)
# print(sku_master_data)


# #check the shape and null values
# print(order_report_data.shape)
# print(pincode_zones_data.shape)
# print(sku_master_data.shape)
# print(order_report_data.isnull().sum())
# print(pincode_zones_data.isnull().sum())
# print(sku_master_data.isnull().sum())


# Display data types to check if they match
# print(order_report_data.dtypes)
# print(sku_master_data.dtypes)
# print(pincode_zones_data.dtypes)


# Merge the dataframes based on the common identifier (sku)
order_sku_data = pd.merge(order_report_data, sku_master_data, on='SKU', how='inner')
# print(order_sku_data)
# print(order_sku_data.shape)
# print(order_sku_data.isnull().sum())

# Add a new column 'Total Weight' by multiplying 'Weight (g)' and 'Order Qty'
order_sku_data['Total Weight(g)'] = order_sku_data['Weight (g)'] * order_sku_data['Order Qty']
# print(order_sku_data)


Invoicer=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Invoice.xlsx"
Invoicer_data=pd.read_excel(Invoicer)
# print(Invoicer_data)
# Merge the dataframes based on the common identifier (Order ID)
# print(order_sku_data.dtypes)
# print(Invoicer_data.dtypes)

order_sku_data=order_sku_data.rename(columns={'ExternOrderNo':"Order ID"})

# print("order_sku_data columns:", order_sku_data.columns)
# print("Invoicer_data columns:", Invoicer_data.columns)

data = pd.merge(order_sku_data, Invoicer_data, on="Order ID", how="inner")
# print(data)
# print(data.shape)
# print(data.dtypes)
# print(data.isnull().sum())

Rates=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Rates.xlsx"
Rates_data=pd.read_excel(Rates)
# print(Rates_data)
# print(Rates_data.dtypes)

Rates_data["Zone"]=Rates_data["Zone"].str.lower()

df=pd.merge(data,Rates_data, on="Zone", how="inner")

print(df)