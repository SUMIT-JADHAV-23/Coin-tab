
import pandas as pd
#load Left Hand Side (LHS) Data
order_report = r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Order Report.xlsx"
pincode_zones=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - Pincode Zones.xlsx"
sku_master=r"E:\Study sumit\Interviwe Assignments\CoinTab\Company X - SKU Master.xlsx"

ord = pd.read_excel(order_report)
pzd= pd.read_excel(pincode_zones)
smd= pd.read_excel(sku_master)
print(ord)
print(pzd)
print(smd)



