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

print(total_weight_per_order)



file_path= r"E:\Study sumit\Interviwe Assignments\CoinTab\total_weight_per_order.xlsx"
total_weight_per_order.to_excel(file_path, index=False)
print(f"Excel file saved to {file_path}")

