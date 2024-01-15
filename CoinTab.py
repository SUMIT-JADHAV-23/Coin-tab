
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
# print(df)
# print(df.dtypes)

df["Total Weight(g)"]=df["Total Weight(g)"]/1000
df["Total Weight(g)"]=df["Total Weight(g)"].round(3)

df.rename(columns={"Total Weight(g)": "Total Weight(kg)"}, inplace=True)

# print(df['Total Weight(kg)'])

import math

def Applicable(tw, ws):
    if tw <= ws:
        aw = ws
        count = 0  # Set count to 0 when no additional slabs are needed
    else:
        slabs_add = math.ceil((tw - ws) / ws)
        count = slabs_add
        aw = ws + (slabs_add * ws)
    return aw, count

# Apply the Applicable function to create new columns
df[["Applicable Weight", "count"]] = df.apply(lambda row: pd.Series(Applicable(row['Total Weight(kg)'], row['Weight Slabs'])), axis=1)

import pandas as pd

def charges(type, fc, fac, rto, rtoa, count):
    if count == 0 and type == "Forward charges":
        charge = fc
    elif count == 0 and type == "Forward and RTO charges":
        charge = fc + rto
    elif count > 0 and type == "Forward charges":
        charge = fc + (count * fac)
    elif count > 0 and type == "Forward and RTO charges":
        charge = fc + (count * fac) + rto + (count * rtoa)
    else:
        charge = 0
    return charge

# Apply the charges function to create a new column
df["charges"] = df.apply(lambda row: charges(row["Type of Shipment"], row["Forward Fixed Charge"], row["Forward Additional Weight Slab Charge"], row["RTO Fixed Charge"], row["RTO Additional Weight Slab Charge"], row["count"]), axis=1)

# output 1
df_final = df.loc[:, ["Order ID", "AWB Code", "Total Weight(kg)", "Applicable Weight", "Charged Weight", "Weight Slabs", "Zone", "charges", "Billing Amount (Rs.)"]]

df_final = df_final.rename(columns={
    "Order ID": "Order ID",
    "AWB Code": "AWSNumber",
    "Total Weight(kg)": "Total Weight as per comp",
    "Applicable Weight": "Total Weight as per comp",
    "Charged Weight": "Total Weight as per courier company",
    "Weight Slabs": "Weight slab charged by courier",
    "Zone": "Delivery Zone",
    "charges": "Expected charges (Rs.)",
    "Billing Amount (Rs.)": "Charged billed by courier"
})
#
# result = r"E:\Study sumit\Interviwe Assignments\CoinTab\result_file.xlsx"
# df_final.to_excel(result, index=False)
# print(f"Excel file saved to {result}")

