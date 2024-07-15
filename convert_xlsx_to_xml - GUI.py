import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from lxml import etree
from datetime import datetime

# Function to handle file selection for Excel file
def choose_excel_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")],
        title="Select Excel File"
    )
    if excel_file_path:
        excel_file_label.config(text=f"Selected: {excel_file_path}")

# Function to handle file selection for XML export
def choose_xml_export():
    global xml_export_path
    xml_export_path = filedialog.asksaveasfilename(
        defaultextension=".xml",
        filetypes=[("XML files", "*.xml")],
        title="Save XML File As"
    )
    if xml_export_path:
        xml_export_label.config(text=f"Exporting to: {xml_export_path}")

# Function to convert selected Excel file to XML
def convert_to_xml():
    try:
        if not excel_file_path:
            messagebox.showerror("Error", "Please select an Excel file.")
            return
        
        # Load the Excel file
        df = pd.read_excel(excel_file_path)
        
        # Replace NaN values with empty strings in the DataFrame
        df = df.fillna('')
        
        # Create the root element
        root = etree.Element("NMEXML", EximID="349", BranchCode="ONLINE", ACCOUNTANTCOPYID="")
        
        # Create the Transactions element
        transactions = etree.SubElement(root, "TRANSACTIONS", OnError="CONTINUE")
        
        # Group by TRANSACTIONID to handle multiple ITEMLINE entries per SALESORDER
        grouped = df.groupby('TRANSACTIONID')
        
        for transaction_id, group in grouped:
            sales_order = etree.SubElement(transactions, "SALESORDER", operation=str(group.iloc[0]['SALESORDER operation']), REQUESTID=str(group.iloc[0]['REQUESTID']))
            etree.SubElement(sales_order, "TRANSACTIONID").text = str(group.iloc[0]['TRANSACTIONID'])
            
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
                etree.SubElement(item_line, "TAXCODES").text = str(item_row.TAXCODES)
                etree.SubElement(item_line, "PROJECTID").text = str(item_row.PROJECTID)
                etree.SubElement(item_line, "DEPTID").text = str(item_row.DEPTID)
                etree.SubElement(item_line, "QTYSHIPPED").text = str(item_row.QTYSHIPPED)
            
            # Add the details after all ITEMLINEs for this SALESORDER
            last_item_row = group.iloc[-1]  # Get the last row in the group
            
            etree.SubElement(sales_order, "SONO").text = str(last_item_row['SONO'])
            etree.SubElement(sales_order, "SODATE").text = datetime.strftime(last_item_row['SODATE'], "%Y-%m-%d")
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
            etree.SubElement(sales_order, "ESTSHIPDATE").text = datetime.strftime(last_item_row['ESTSHIPDATE'], "%Y-%m-%d")
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
        
        if not xml_export_path:
            messagebox.showerror("Error", "Please choose XML export path.")
            return
        
        # Convert the XML tree to a string and save it to the chosen file
        xml_string = etree.tostring(root, pretty_print=True, encoding='utf-8')
        with open(xml_export_path, 'wb') as f:
            f.write(xml_string)
            
        messagebox.showinfo("Conversion Complete", "Excel file converted to XML successfully!")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI Setup
root = tk.Tk()
root.title("Excel to XML Converter")

# Center the window
window_width = 500
window_height = 250
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f'{window_width}x{window_height}+{x}+{y}')

# Button to choose Excel file
excel_file_button = tk.Button(root, text="Choose Excel File", command=choose_excel_file)
excel_file_button.pack(pady=20)

# Label to display selected Excel file path
excel_file_label = tk.Label(root, text="No Excel file selected")
excel_file_label.pack()

# Button to choose XML export location
xml_export_button = tk.Button(root, text="Choose XML Export Location", command=choose_xml_export)
xml_export_button.pack(pady=20)

# Label to display selected XML export path
xml_export_label = tk.Label(root, text="No XML export location selected")
xml_export_label.pack()

# Button to convert Excel to XML
convert_button = tk.Button(root, text="Convert Excel to XML", command=convert_to_xml)
convert_button.pack(pady=20)

# Run the GUI loop
root.mainloop()