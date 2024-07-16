import pandas as pd
from lxml import etree
from datetime import datetime
import random

# Load the Excel file
df = pd.read_excel('sales_orders.xlsx')

# Replace NaN values with empty strings in the DataFrame
df = df.fillna('')

# Initialize a counter outside the loop
sono_counter = 20000

# Create the root element
root = etree.Element("NMEXML", EximID=str(random.randint(100, 999)), BranchCode="ONLINE", ACCOUNTANTCOPYID="")

# Create the Transactions element
transactions = etree.SubElement(root, "TRANSACTIONS", OnError="CONTINUE")

# Group by TRANSACTIONID to handle multiple ITEMLINE entries per SALESORDER
grouped = df.groupby('PONO')

for transaction_id, group in grouped:
    # Generate a random 6-digit number for TRANSACTIONID
    random_transaction_id = random.randint(100000, 999999)
    
    # Create a single SALESORDER element for each group
    sales_order = etree.SubElement(transactions, "SALESORDER", operation="Add", REQUESTID="1")
    etree.SubElement(sales_order, "TRANSACTIONID").text = str(random_transaction_id)
        
    # Enumerate through item rows to create ITEMLINE elements
    for idx, item_row in enumerate(group.itertuples(), start=1):
        item_line = etree.SubElement(sales_order, "ITEMLINE", operation="Add")
        etree.SubElement(item_line, "KeyID").text = str(idx - 1)  # Adjusting index to start from 0
        etree.SubElement(item_line, "ITEMNO").text = str(item_row.ITEMNO)
        etree.SubElement(item_line, "QUANTITY").text = str(item_row.QUANTITY)
        etree.SubElement(item_line, "ITEMUNIT").text = str(item_row.ITEMUNIT)
        etree.SubElement(item_line, "UNITRATIO").text = str(item_row.UNITRATIO)
        etree.SubElement(item_line, "ITEMRESERVED1").text = str(item_row.ITEMRESERVED1)
        etree.SubElement(item_line, "ITEMRESERVED2").text = str(item_row.ITEMRESERVED2)
        etree.SubElement(item_line, "ITEMRESERVED3").text = str(item_row.ITEMRESERVED3)
        etree.SubElement(item_line, "ITEMRESERVED4").text = str(item_row.ITEMRESERVED4)
        etree.SubElement(item_line, "ITEMRESERVED5").text = str(item_row.ITEMRESERVED5)
        etree.SubElement(item_line, "ITEMRESERVED6").text = str(item_row.ITEMRESERVED6)
        etree.SubElement(item_line, "ITEMRESERVED7").text = str(item_row.ITEMRESERVED7)
        etree.SubElement(item_line, "ITEMRESERVED8").text = str(item_row.ITEMRESERVED8)
        etree.SubElement(item_line, "ITEMRESERVED9").text = str(item_row.ITEMRESERVED9)
        etree.SubElement(item_line, "ITEMRESERVED10").text = str(item_row.ITEMRESERVED10)
        etree.SubElement(item_line, "ITEMOVDESC").text = str(item_row.ITEMOVDESC)
        etree.SubElement(item_line, "UNITPRICE").text = str(item_row.UNITPRICE)
        etree.SubElement(item_line, "DISCPC").text = str(item_row.DISCPC)
        etree.SubElement(item_line, "TAXCODES").text = str('T')
        etree.SubElement(item_line, "PROJECTID").text = str('TMO-1101')
        etree.SubElement(item_line, "DEPTID").text = str('ONLINE-TMS')
        etree.SubElement(item_line, "QTYSHIPPED").text = str(item_row.QTYSHIPPED)
    
    # Add the details after all ITEMLINEs for this SALESORDER
    last_item_row = group.iloc[-1]  # Get the last row in the group
    
    # Adjusting SONO format with a counter
    sono_format = f"SCO-S{datetime.now().strftime('%y%m')}-{sono_counter}"
    etree.SubElement(sales_order, "SONO").text = sono_format
    # Increment the counter
    sono_counter += 1
    # Generate today's date in YYYY-MM-DD format
    sodate_format = datetime.now().strftime('%Y-%m-%d')
    etree.SubElement(sales_order, "SODATE").text = sodate_format
    etree.SubElement(sales_order, "TAX1ID").text = str('T')  # Default value 'T'
    etree.SubElement(sales_order, "TAX1CODE").text = str('T')  # Default value 'T'
    etree.SubElement(sales_order, "TAX2CODE").text = str(last_item_row['TAX2CODE'])
    etree.SubElement(sales_order, "TAX1RATE").text = str(11)  # Default value 11
    etree.SubElement(sales_order, "TAX2RATE").text = str(last_item_row['TAX2RATE'])
    etree.SubElement(sales_order, "TAX1AMOUNT").text = str(last_item_row['TAX1AMOUNT'])
    etree.SubElement(sales_order, "TAX2AMOUNT").text = str(last_item_row['TAX2AMOUNT'])
    etree.SubElement(sales_order, "RATE").text = str(1)  # Default value 1
    etree.SubElement(sales_order, "TAXINCLUSIVE").text = str(1)  # Default value 1
    etree.SubElement(sales_order, "CUSTOMERISTAXABLE").text = str(1)  # Default value 1
    etree.SubElement(sales_order, "CASHDISCOUNT").text = str(last_item_row['CASHDISCOUNT'])
    etree.SubElement(sales_order, "CASHDISCPC").text = str(last_item_row['CASHDISCPC'])
    etree.SubElement(sales_order, "FREIGHT").text = str(last_item_row['FREIGHT'])
    etree.SubElement(sales_order, "TERMSID").text = str('C.O.D')  # Default value C.O.D
    etree.SubElement(sales_order, "SHIPVIAID").text = str(last_item_row['SHIPVIAID'])
    etree.SubElement(sales_order, "FOB").text = str(last_item_row['FOB'])
    etree.SubElement(sales_order, "ESTSHIPDATE").text = sodate_format
    etree.SubElement(sales_order, "DESCRIPTION").text = str(last_item_row['DESCRIPTION'])
    etree.SubElement(sales_order, "SHIPTO1").text = str(last_item_row['SHIPTO1'])
    etree.SubElement(sales_order, "SHIPTO2").text = str(last_item_row['SHIPTO2'])
    etree.SubElement(sales_order, "SHIPTO3").text = str(last_item_row['SHIPTO3'])
    etree.SubElement(sales_order, "SHIPTO4").text = str(last_item_row['SHIPTO4'])
    etree.SubElement(sales_order, "SHIPTO5").text = str(last_item_row['SHIPTO5'])
    etree.SubElement(sales_order, "DP").text = str(0)  # Default value 0
    etree.SubElement(sales_order, "DPACCOUNTID").text = str('TMS-210202')  # Default value TMS-210202
    etree.SubElement(sales_order, "DPUSED").text = str(last_item_row['DPUSED'])
    etree.SubElement(sales_order, "CUSTOMERID").text = str('TMO-1101')  # Default value TMO-1101
    etree.SubElement(sales_order, "PONO").text = str(last_item_row['PONO'])
    
    # Create SALESMANID element and add LASTNAME and FIRSTNAME inside it
    salesman_id = etree.SubElement(sales_order, "SALESMANID")
    etree.SubElement(salesman_id, "LASTNAME").text = str(last_item_row['LASTNAME'])
    etree.SubElement(salesman_id, "FIRSTNAME").text = str(last_item_row['FIRSTNAME'])

    etree.SubElement(sales_order, "CURRENCYNAME").text = str('IDR')  # Default value IDR

# Convert the XML tree to a string and save it to a file
xml_string = etree.tostring(root, pretty_print=True, encoding='utf-8')
with open('converted_sales_orders.xml', 'wb') as f:
    f.write(xml_string)
