import pandas as pd

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

print(df.dtypes)


