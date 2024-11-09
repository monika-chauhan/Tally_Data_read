import xml.etree.ElementTree as ET

import pandas as pd 
# XML data string
tree = ET.parse('Input.xml')
root = tree.getroot()

# Function to format the date
def format_date(date_str):
    return f"{date_str[6:8]}-{date_str[4:6]}-{date_str[:4]}"

# Initialize a list to store extracted data
extracted_data = []

for voucher in root.findall(".//TALLYMESSAGE/VOUCHER"):
   
    vch_no = voucher.find("VOUCHERNUMBER").text
    vch_type = voucher.get("VCHTYPE")
    date_value = format_date(voucher.find("DATE").text) 
    if vch_type == "Receipt":
        par_entry = {
                    "Date": date_value,
                    "Transaction Type": "Parent",
                    "Vch No.": vch_no,
                    "Ref No": "NA",
                    "Ref Type": "NA",
                    "Ref Date": "NA",
                    "Debtor": "NA",
                    "Ref Amount": "NA",
                    "Amount": "NA",
                    "Particulars" :"NA",
                    "Vch Type" : vch_type,
                    "Amount Verified": "Yes"
                }
        for ledger_entry in voucher.findall(".//ALLLEDGERENTRIES.LIST"):
            total_Amount = 0
            particulars = debator = ledger_entry.find("LEDGERNAME").text if ledger_entry.find("LEDGERNAME") is not None else "NA" 
            bill_allocations = ledger_entry.findall(".//BILLALLOCATIONS.LIST")
            name = bill_allocations[0].find("NAME").text if bill_allocations[0].find("NAME") is not None else None
            if name is None:
                other_entry = {
                    "Date": date_value,
                    "Transaction Type": "Other" ,
                    "Vch No.": vch_no,
                    "Ref No": "NA",
                    "Ref Type": "NA",
                    "Ref Date": "NA",
                    "Debtor": debator,
                    "Ref Amount": "NA",
                    "Amount": ledger_entry.find("AMOUNT").text,
                    "Particulars" :particulars,
                    "Vch Type" : vch_type,
                    "Amount Verified": "NA"
                }
                extracted_data.append(other_entry) 
                            
            else:
                for bill_entry in bill_allocations:
                    child_entry = {
                    "Date": date_value,
                    "Transaction Type": "NA",
                    "Vch No.": vch_no,
                    "Ref No": "NA",
                    "Ref Type": "NA",
                    "Ref Date": "NA",
                    "Debtor": debator,
                    "Ref Amount": "NA",
                    "Amount": "NA",
                    "Particulars" :particulars,
                    "Vch Type" : vch_type,
                    "Amount Verified": "NA"
                    }
                    child_entry["Transaction Type"] = "Child" 
                    child_entry["Ref No"] = bill_entry.find("NAME").text if bill_entry.find("NAME") is not None else "NA"
                    child_entry["Ref Type"] = bill_entry.find("BILLTYPE").text if bill_entry.find("BILLTYPE") is not None else "NA"
                    child_entry["Ref Date"] = None if bill_entry.find("AMOUNT") is not None else "NA"
                    ref_amount = bill_entry.find("AMOUNT").text if bill_entry.find("AMOUNT") is not None else "NA"
                    child_entry["Ref Amount"] = ref_amount
                    total_Amount += float(ref_amount) if ref_amount else 0.0
                    extracted_data.append(child_entry) 
                par_entry["Amount"] = total_Amount
                par_entry["Debtor"]  = debator
                par_entry["Particulars"] = particulars
                extracted_data.append(par_entry)
              
df = pd.DataFrame(extracted_data)
sort_order = {"Parent": 0, "Child": 1, "Others": 2}
df['TransactionOrder'] = df['Transaction Type'].map(sort_order)
df_sorted = df.sort_values(by=['Date', 'Vch No.', 'TransactionOrder'])
# Drop the helper column after sorting
df_sorted = df_sorted.drop(columns=['TransactionOrder'])
# Display the DataFrame

print(df_sorted)
df_sorted.to_excel("output.xlsx",index=False) 
print("file Saves")        
            
            
                    