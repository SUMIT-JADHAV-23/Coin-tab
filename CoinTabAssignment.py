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
print(order_sku_data)

order_sku_data['Total Weight(g)'] = order_sku_data['Weight (g)'] * order_sku_data['Order Qty']

order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"]/1000
order_sku_data["Total Weight(g)"]=order_sku_data["Total Weight(g)"].round(3)
order_sku_data.rename(columns={"Total Weight(g)": "Total Weight(kg)"}, inplace=True)
print(order_sku_data)


# #RHS
#
# Courier_Invoicer=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Invoice.xlsx"
# Courier_Rates=r"E:\Study sumit\Interviwe Assignments\CoinTab\Courier Company - Rates.xlsx"
#
# Courier_Invoicer_data=pd.read_excel(Invoicer)
# Courier_Rates_data=pd.read_excel(Rates)
# # print(Courier_Invoicer_data.dtypes)
# # print(Courier_Rates_data.dtypes)
#
# Courier_Rates_data["Zone"]=Courier_Rates_data["Zone"].str.lower()
# # print(Courier_Rates_data)
# Courier_data=pd.merge ( Courier_Invoicer_data, Courier_Rates_data, on="Zone", how="inner")
# print(Courier_data)
#
#
# new1= r"E:\Study sumit\Interviwe Assignments\CoinTab\Courierdatacombine.xlsx"
# Courier_data.to_excel(new1, index=False)
# print(f"Excel file saved to {new1}")