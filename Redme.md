# README
You are a data analyst tasked with verifying the charges levied by multiple courier companies used by a large e-commerce company in India (referred to as X). The charges are dependent on the weight of the product and the distance between the warehouse and the customer's delivery address. The objective is to ensure that the charges imposed by the courier partners align with the expected values calculated by X.

# Outputs of program 

The program generates two output Excel files, namely "Output Data1.xlsx" and "Output Data2.xlsx."
### In Output Data1.xlsx file
In `Output Data1.xlsx`, a detailed result is provided with columns such as below 

Output Data 1
Created a resultant in Excel file with the following columns:
- Order ID : 
- AWB Number
- Total weight as per X (KG)
- Weight slab as per X (KG)
- Total weight as per Courier Company (KG)
- Weight slab charged by Courier Company (KG)
- Delivery Zone as per X
- Delivery Zone charged by Courier Company
- Expected Charge as per X (Rs.)
- Charges Billed by Courier Company (Rs.)
- Difference Between Expected Charges and Billed Charges (Rs.)

### In Output Data2.xlsx
In `Output Data2.xlsx`, a summary table is created with information categorized into three sections:

Create a summary table

|          | Count | Amount (Rs.) |
|----------|-------|--------------|
| Total Orders - Correctly Charged  | 8     | 0            |
| Total Orders - Over Charged | 114   | -7695.7      |
| Total Orders - Under Charged    | 2     | 47.2         |

### Instruction

- In the `CoinTabAssignment.py` script, we implemented the following logical steps:
  
  - Data Reading:
    - Read the provided data from the Excel file.
  - Data Merging:
    - Merge the data sheets from `Company X - Order Report.xlsx` and `Company X - SKU Master.xlsx` based on the common 'SKU' identifier.
  - Total Weight Calculation:
    - Calculate the total weight for each unique 'SKU' in the dataset, representing the weight of individual products.
  - Total Weight per Order ID:
    - Determine the total weight per 'Order ID' by aggregating the weights of the products included in each order.
  - Zone Comparison:
    - Compare the delivery zones reported in the internal data with those mentioned in the courier company's CSV invoice.
  - Applicable Weight Calculation:
    - Determine the applicable weight for each order based on the weight slabs corresponding to the delivery zone.
  - Charge Calculation:
    - Calculate the charges for each order, considering the rate card provided by the courier company and factoring in any additional charges or fixed rates.
  - Difference Calculation:
    - Find the difference between the expected charges (as per X) and the billed charges by the courier company.
  - Summary Table Creation:
    - Generate a summary table summarizing the counts and amounts for orders that were correctly charged, overcharged, and undercharged.
    
- The script efficiently performs these operations, ensuring a thorough analysis of charging accuracy for the given e-commerce company's orders.
