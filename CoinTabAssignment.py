import pandas as pd
import math

order_report = r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Order Report.xlsx"
pincode_zones=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Pincode Zones.xlsx"
sku_master=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - SKU Master.xlsx"
Invoicer=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Invoice.xlsx"
Rates=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Rates.xlsx"

order_report_data = pd.read_excel(order_report)
pincode_zones_data= pd.read_excel(pincode_zones)
sku_master_data= pd.read_excel(sku_master)
Invoicer_data=pd.read_excel(Invoicer)
Rates_data=pd.read_excel(Rates)

#load And merge Left Hand Side (LHS) Data
#merge data on comman indentifier #sku
order_sku_data = pd.merge(order_report_data, sku_master_data, on='SKU', how='inner')
# print(order_sku_data)

order_sku_data['Total Weight(g)'] = order_sku_data['Weight (g)'] * order_sku_data['Order Qty']

order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"]/1000
order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"].round(3)
order_sku_data.rename(columns={"Total Weight(g)": "Total Weight(kg)"}, inplace=True)
# print(order_sku_data)

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\order_sku_data.xlsx"
# order_sku_data.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")
# print(order_sku_data.columns)

order_sku_data.drop(columns=['SKU', 'Order Qty', 'Weight (g)'],inplace=True)
# print(order_sku_data)


# Group by 'Order ID' and calculate the total weight
total_weight_per_order = order_sku_data.groupby('ExternOrderNo')['Total Weight(kg)'].sum().reset_index()
# print(total_weight_per_order)

# compare zone with courier and company
data1 = pd.DataFrame({'Customer Pincode': Invoicer_data["Customer Pincode"].unique()})
data2 = pincode_zones_data

# Merge data3 and data4 based on "Customer Pincode"
compare_zone_left = pd.merge(data1, data2, on="Customer Pincode", how='left', suffixes=('_data1', '_data2'))
# print(compare_zone_left)

# Delete duplicates
merged_data_no_duplicates = compare_zone_left.drop_duplicates(subset="Customer Pincode")
# print(merged_data_no_duplicates)


Compare_to_courier=pd.merge(Invoicer_data,merged_data_no_duplicates,on="Customer Pincode",how="inner")
# print(Compare_to_courier)


# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\Compare_to_courier.xlsx"
# Compare_to_courier.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")

Compare_to_courier.drop(columns=['Warehouse Pincode_y', 'Warehouse Pincode_x'],inplace=True)
# print(Compare_to_courier)

Compare_to_courier=Compare_to_courier.rename(columns={"Zone_y": "Delivery Zone as per X"})
# print(Compare_to_courier)

Compare_to_courier=Compare_to_courier.rename(columns={"Zone_x": "Delivery Zone charged by Courier Company"})
# print(Compare_to_courier)
#
# print(Compare_to_courier["Delivery Zone as per X"])
# print(Compare_to_courier["Delivery Zone charged by Courier Company"])



#comparare with curier
total_weight_per_order=total_weight_per_order.rename(columns={'ExternOrderNo':"Order ID"})
# print(total_weight_per_order)


#campare company/courier oder id and weight c
df1 = total_weight_per_order
df2 = Compare_to_courier
df = pd.merge(df1, df2, on="Order ID", how="inner")

# print(df.dtypes)

df=df.rename(columns={'Total Weight(kg)':'Total weight as per X (KG)'})
df=df.rename(columns={'Charged Weight':"Total weight as per Courier Company (KG)"})

# print(df.dtypes)

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\df.xlsx"
# df.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")




# calculate wight slab for company

Rates_data["Zone"]=Rates_data["Zone"].str.lower()
Rates_data1=Rates_data.rename(columns={"Zone":"Delivery Zone as per X"})
# print(Rates_data1)
df_company=pd.merge(df,Rates_data1, on="Delivery Zone as per X", how="inner")
# print(df_company.dtypes)

# calculate weight slab for courier
Rates_data2=Rates_data.rename(columns={"Zone":"Delivery Zone charged by Courier Company"})
df_courier=pd.merge(df,Rates_data2, on="Delivery Zone charged by Courier Company", how="inner")
# print(df_courier.dtypes)

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
df_company[["Weight slab as per X (KG)", "count per X"]] = df_company.apply(lambda row: pd.Series(Applicable(row['Total weight as per X (KG)'], row['Weight Slabs'])), axis=1)
df_courier[["Weight slab charged by Courier Company (KG)", "count per courier"]] = df_courier.apply(lambda row: pd.Series(Applicable(row['Total weight as per Courier Company (KG)'], row['Weight Slabs'])), axis=1)
# print(df_company)
# print(df_courier)

selected_column=["Weight slab charged by Courier Company (KG)","Order ID"]
df_courier1=df_courier[selected_column]
df_main=pd.merge(df_company,df_courier1,on="Order ID",how="inner")
# print(df_main.dtypes)



df_main = df_main[['Order ID', 'AWB Code', 'Total weight as per X (KG)','Weight slab as per X (KG)','Total weight as per Courier Company (KG)','Weight slab charged by Courier Company (KG)',"Customer Pincode",'Delivery Zone as per X','Delivery Zone charged by Courier Company','Billing Amount (Rs.)','Type of Shipment','Weight Slabs','Forward Fixed Charge','Forward Additional Weight Slab Charge','RTO Fixed Charge','RTO Additional Weight Slab Charge','count per X']]

# print(df_main.dtypes)

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\dfmain.xlsx"
# df_main.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")

#clculate charges for per order
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
df_main["Expected Charge as per X (Rs.)"] = df_main.apply(lambda row: charges(row["Type of Shipment"], row["Forward Fixed Charge"], row["Forward Additional Weight Slab Charge"], row["RTO Fixed Charge"], row["RTO Additional Weight Slab Charge"], row["count per X"]), axis=1)

# print(df_main.dtypes)
df_final=df_main
# print(df_final)

file_path= r"dffinal.xlsx"
df_final.to_excel(file_path, index=False)
print(f"Excel file saved to {file_path}")

df_final.drop(columns=["Customer Pincode",'Type of Shipment','Weight Slabs','Forward Fixed Charge','Forward Additional Weight Slab Charge','RTO Fixed Charge','RTO Additional Weight Slab Charge','count per X'],inplace=True)
print(df_final.dtypes)


df_final = df_final[['Order ID', 'AWB Code', 'Total weight as per X (KG)','Weight slab as per X (KG)','Total weight as per Courier Company (KG)','Weight slab charged by Courier Company (KG)','Delivery Zone as per X','Delivery Zone charged by Courier Company','Expected Charge as per X (Rs.)','Billing Amount (Rs.)']]
print(df_final.dtypes)

df_final=df_final.rename(columns={"Billing Amount (Rs.)":"Charges Billed by Courier Company (Rs.)"})
print(df_final.dtypes)

df_final["Difference Between Expected and Billed  (Rs.)"]=df_final["Expected Charge as per X (Rs.)"]-df_final["Charges Billed by Courier Company (Rs.)"]
print(df_final.dtypes)

# print(df_final["Difference Between Expected and Billed  (Rs.)"])


file_path= r"Output Data1.xlsx"
df_final.to_excel(file_path, index=False)
print(f"Excel file saved to {file_path}")
