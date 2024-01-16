import pandas as pd
import math

order_report = r"InputFiles/Company X - Order Report.xlsx"
pincode_zones=r"InputFiles/Company X - Pincode Zones.xlsx"
sku_master=r"InputFiles/Company X - SKU Master.xlsx"
Invoicer=r"InputFiles/Courier Company - Invoice.xlsx"
Rates=r"InputFiles/Courier Company - Rates.xlsx"

order_report_data = pd.read_excel(order_report)
pincode_zones_data= pd.read_excel(pincode_zones)
sku_master_data= pd.read_excel(sku_master)
Invoicer_data=pd.read_excel(Invoicer)
Rates_data=pd.read_excel(Rates)

#load And merge Left Hand Side (LHS) Data
#merge data on comman indentifier #sku
order_sku_data = pd.merge(order_report_data, sku_master_data, on='SKU', how='inner')
order_sku_data['Total Weight(g)'] = order_sku_data['Weight (g)'] * order_sku_data['Order Qty']
order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"]/1000
order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"].round(3)
order_sku_data.rename(columns={"Total Weight(g)": "Total Weight(kg)"}, inplace=True)

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\order_sku_data.xlsx"
# order_sku_data.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")
# print(order_sku_data.columns)

order_sku_data.drop(columns=['SKU', 'Order Qty', 'Weight (g)'],inplace=True)

# Group by 'Order ID' and calculate the total weight
total_weight_per_order = order_sku_data.groupby('ExternOrderNo')['Total Weight(kg)'].sum().reset_index()

# compare zone with courier and company
data1 = pd.DataFrame({'Customer Pincode': Invoicer_data["Customer Pincode"].unique()})
data2 = pincode_zones_data

# Merge data3 and data4 based on "Customer Pincode"
compare_zone_left = pd.merge(data1, data2, on="Customer Pincode", how='left', suffixes=('_data1', '_data2'))

# Delete duplicates
merged_data_no_duplicates = compare_zone_left.drop_duplicates(subset="Customer Pincode")
Compare_to_courier=pd.merge(Invoicer_data,merged_data_no_duplicates,on="Customer Pincode",how="inner")

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\Compare_to_courier.xlsx"
# Compare_to_courier.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")

Compare_to_courier.drop(columns=['Warehouse Pincode_y', 'Warehouse Pincode_x'],inplace=True)
Compare_to_courier=Compare_to_courier.rename(columns={"Zone_y": "Delivery Zone as per X"})
Compare_to_courier=Compare_to_courier.rename(columns={"Zone_x": "Delivery Zone charged by Courier Company"})

#comparare with curier
total_weight_per_order=total_weight_per_order.rename(columns={'ExternOrderNo':"Order ID"})

#campare company/courier oder id and weight c
df1 = total_weight_per_order
df2 = Compare_to_courier
df = pd.merge(df1, df2, on="Order ID", how="inner")

df=df.rename(columns={'Total Weight(kg)':'Total weight as per X (KG)'})
df=df.rename(columns={'Charged Weight':"Total weight as per Courier Company (KG)"})

# file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\df.xlsx"
# df.to_excel(file_path, index=False)
# print(f"Excel file saved to {file_path}")

# calculate wight slab for company
Rates_data["Zone"]=Rates_data["Zone"].str.lower()
Rates_data1=Rates_data.rename(columns={"Zone":"Delivery Zone as per X"})
df_company=pd.merge(df,Rates_data1, on="Delivery Zone as per X", how="inner")

# calculate weight slab for courier
Rates_data2=Rates_data.rename(columns={"Zone":"Delivery Zone charged by Courier Company"})
df_courier=pd.merge(df,Rates_data2, on="Delivery Zone charged by Courier Company", how="inner")

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
selected_column=["Weight slab charged by Courier Company (KG)","Order ID"]
df_courier1=df_courier[selected_column]
df_main=pd.merge(df_company,df_courier1,on="Order ID",how="inner")
df_main = df_main[['Order ID', 'AWB Code', 'Total weight as per X (KG)','Weight slab as per X (KG)','Total weight as per Courier Company (KG)','Weight slab charged by Courier Company (KG)',"Customer Pincode",'Delivery Zone as per X','Delivery Zone charged by Courier Company','Billing Amount (Rs.)','Type of Shipment','Weight Slabs','Forward Fixed Charge','Forward Additional Weight Slab Charge','RTO Fixed Charge','RTO Additional Weight Slab Charge','count per X']]

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
df_final=df_main
df_final.drop(columns=["Customer Pincode",'Type of Shipment','Weight Slabs','Forward Fixed Charge','Forward Additional Weight Slab Charge','RTO Fixed Charge','RTO Additional Weight Slab Charge','count per X'],inplace=True)
df_final = df_final[['Order ID', 'AWB Code', 'Total weight as per X (KG)','Weight slab as per X (KG)','Total weight as per Courier Company (KG)','Weight slab charged by Courier Company (KG)','Delivery Zone as per X','Delivery Zone charged by Courier Company','Expected Charge as per X (Rs.)','Billing Amount (Rs.)']]
df_final=df_final.rename(columns={"Billing Amount (Rs.)":"Charges Billed by Courier Company (Rs.)"})
df_final["Difference Between Expected and Billed  (Rs.)"]=df_final["Expected Charge as per X (Rs.)"].round(decimals=2) - df_final["Charges Billed by Courier Company (Rs.)"].round(decimals=2)

file_path= r"Output Data1.xlsx"
df_final.to_excel(file_path, index=False)
print(f"Excel file saved to {file_path}")

#output2
count_correctly_charged = 0
count_overcharged_charged = 0
count_undercharged_count = 0
total_correctly_amount = 0
total_overcharging_amount = 0
total_undercharging_amount = 0

# Assuming df_final is your DataFrame
for index, row in df_final.iterrows():
    if row["Difference Between Expected and Billed  (Rs.)"] == 0:
        count_correctly_charged += 1
        total_correctly_amount += row["Difference Between Expected and Billed  (Rs.)"]
    elif row["Difference Between Expected and Billed  (Rs.)"] < 0:
        count_overcharged_charged += 1
        total_overcharging_amount += row["Difference Between Expected and Billed  (Rs.)"]
    elif row["Difference Between Expected and Billed  (Rs.)"] > 0:
        count_undercharged_count += 1
        total_undercharging_amount += row["Difference Between Expected and Billed  (Rs.)"]
    else:
        pass

# Create a summary table
summary_table = pd.DataFrame({
    "  ":['Total Orders - Correctly Charged', 'Total Orders - Over Charged', 'Total Orders - Under Charged'],
    'Count': [count_correctly_charged, count_overcharged_charged, count_undercharged_count],
    'Amount (Rs.)': [total_correctly_amount, total_overcharging_amount, total_undercharging_amount]
}, index=['Total Orders - Correctly Charged', 'Total Orders - Over Charged', 'Total Orders - Under Charged'])

file_path= r"Output Data2.xlsx"
summary_table.to_excel(file_path, index=False)
print(f"Excel file saved to {file_path}")