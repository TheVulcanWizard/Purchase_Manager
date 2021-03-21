

import gspread
import re
import time
from googleapiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
from tkinter import Tk
from tkinter import filedialog, messagebox
from os import path, makedirs
from shutil import move
from sys import exit
from time import sleep
from re import search
from bs4 import BeautifulSoup


def handle_error(api_request, *args, **kwargs):
    loops = 0
    max_time = 50
    while True:
        try:
            result = api_request(*args, **kwargs)
            break
        except gspread.exceptions.APIError:
            loops += 1
            timer = 1.5 ** loops
            if timer > max_time:
                timer = max_time
            print("\n\nAPIError 429\n\n" + str(APIError))
            for n in reversed(range(round(timer))):
                print(f"Trying again in: {n} seconds", end="\r")
                sleep(1)
    sleep(0.75)
    return result


start = time.time()

root = Tk().withdraw()

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

if path.exists("Credentials\\creds.json"):
    creds = ServiceAccountCredentials.from_json_keyfile_name("Credentials\\creds.json", scope)
    client = gspread.authorize(creds)
else:
    messagebox.showwarning(title="Credentials Not Found", message="Credential file cannon be located. Ensure that your \"creds.json\" file is in the \"Credentials\" directory and try again.")
    exit()

gss = client.open_by_key("xxx") #General Chemistry

pss = client.open_by_key("xxx") #Physical Chemistry

iss = client.open_by_key("xxx") #Inorganic Chemistry


filenames = filedialog.askopenfilenames(initialdir="Input Docs\\", title="Select files", filetypes=(("HTML files", "*.html"), ("All files", "*.*")))
for file_selected in filenames:
    allow_transcribe = True
    file_name = path.basename(file_selected)
    if path.exists("Transcribed Pages\\" + file_name):
        messagebox.showwarning(title="Duplicate File Detected", message="The file \"" + file_name + "\" has already been transcribed and will be ignored. To parse this file:\n\n1. Ensure all relevant PO numbers are removed from the spreadsheet\n2. Remove the duplicate file from the \"Transcribed Pages\" folder")
    else:
        with open("Input Docs\\" + file_name) as the_file:
            doc = BeautifulSoup(the_file, "html.parser")
        
        #Find which account we are charging and access the associated Google Sheet
        order_info = doc.find(id="DocGroupBox7").table
        cart_name = order_info.find_all("tr")[0].find_all("td")[1].get_text(strip=True)

        order_date_text = re.search(r'(\d+-\d+-\d+)', cart_name).group(1)
        print(order_date_text)

        account_text = order_info.find_all("tr")[6].find_all("td")[1].get_text()
        if "ILB" in account_text:
            ss = iss
        elif "PLB" in account_text:
            ss = pss
        elif "GLB" in account_text:
            ss = gss
        else:
            messagebox.showwarning(title="Course Code not Recognized", message="The course code in file \"" + file_name + "was not recognized! Ending program.")
            break
        
        existing_sheets = handle_error(ss.worksheets)

        existing_sheet_titles = [i.title for i in existing_sheets]

        requisitions = doc.find_all(class_="ForegroundContainer")
        for requisition in requisitions:
            po_container = requisition.find(class_="SupplierOnlyGroup") 
            po_number = po_container.find_all("td")[1].a.get_text(strip=True)
            print("PO Number: " + po_number + "\n-----------------------------\n")
            if po_number in existing_sheet_titles:
                allow_transcribe = False
                ignore_po = True
                messagebox.showwarning(title="Duplicate PO Number Detected", message="The PO number \"" + po_number + "\" already exists in the spreadsheet and will be ignored. This file will NOT be moved to the \"Transcribed Pages\" directory. \n\nTo parse this order, delete the PO number from the spreadsheet and run the file again. \n\nAny additional PO numbers in this file will be scanned now.")
            else:
                ignore_po = False
                rows = 1
                sheet_contents = []
                
                new_sheet = handle_error(ss.duplicate_sheet, 0, new_sheet_name=po_number)

                new_sheet_id = handle_error(ss.worksheet, po_number)._properties['sheetId']
                
                line_items_container = requisition.find_all("div", recursive=False)[1]
                line_items = line_items_container.find_all(id=re.compile("^LineItemSixPack"))
                for item in line_items:
                    rows += 1
                    item_properties = item.find_all("td")

                    #Product Name
                    product_name = item_properties[3].a.get_text(strip=True)
                    print("Product Name: " + product_name)

                    #Catalog Number
                    catalog_number = item_properties[4].get_text(strip=True)
                    print("Catalog Number: " + catalog_number)

                    #Size/Packaging
                    size_packaging = item_properties[5].div.get_text(strip=True)
                    print("Size/Packaging: " + size_packaging)

                    #Unit Price
                    unit_price = item_properties[6].span.get_text(strip=True)
                    print("Unit Price: " + unit_price) #! Need to send this as a float

                    #Quantity
                    quantity = item_properties[7].div.get_text(strip=True).split()[0]
                    print("Quantity: " + quantity) #! Need to send this as an integer

                    #Ext. Price
                    ext_price_text = item_properties[8].span.get_text(strip=True)
                    ext_price = re.sub("[^0-9\.]", "", ext_price_text)
                    print("Ext. Price: " + ext_price + "\n") #! Need to send this as a float

                    row_contents = [
                        product_name, 
                        catalog_number, 
                        size_packaging, 
                        float(unit_price), 
                        int(quantity),
                        float(ext_price),
                        0,
                        0,
                    ]
                    sheet_contents.append(row_contents)

                handle_error(new_sheet.insert_rows, sheet_contents, 2)
                handle_error(new_sheet.update, "L1", order_date_text, value_input_option='USER_ENTERED')
                print("-----------------------------\n" + str(rows-1) + " rows added to " + po_number)

                body = {
                    "requests": [
                        {#autofit columns to text
                            "autoResizeDimensions": {
                                "dimensions": {
                                    "sheetId": new_sheet_id,
                                    "dimension": "COLUMNS",
                                }
                            }
                        },
                        {#create checkboxes in "Completed"?" columns
                            'repeatCell': {
                                'cell': {
                                    'dataValidation': {
                                        'condition': {
                                            'type': 'BOOLEAN'}
                                            }
                                },
                                'range': {
                                    'sheetId': new_sheet_id, 
                                    'startRowIndex': 1, 
                                    'endRowIndex': rows,
                                    'startColumnIndex': 8,
                                    'endColumnIndex': 9,
                                    },
                                'fields': 'dataValidation'
                            }
                        },
                        {"repeatCell": {
                            "range": {
                                'sheetId': new_sheet_id, 
                                    'startRowIndex': 1, 
                                    'endRowIndex': rows,
                                    'startColumnIndex': 3,
                                    'endColumnIndex': 4,
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "numberFormat": {
                                        "type": "CURRENCY",
                                        "pattern": "\"$\"#,##0.00",
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.numberFormat"
                        }},
                        {"repeatCell": {
                            "range": {
                                'sheetId': new_sheet_id, 
                                    'startRowIndex': 1, 
                                    'endRowIndex': rows,
                                    'startColumnIndex': 5,
                                    'endColumnIndex': 6,
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "numberFormat": {
                                        "type": "CURRENCY",
                                        "pattern": "\"$\"#,##0.00",
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.numberFormat"
                        }}
                    ]
                }
                for row in range(2, rows+1):
                    body["requests"].append(
                        {# Conditional formatting for "Sent?" checkbox
                            "addConditionalFormatRule": {
                                "rule": {
                                    "ranges": [
                                        {
                                            "sheetId": new_sheet_id,
                                            'startRowIndex': row - 1, 
                                            'endRowIndex': row,
                                            'startColumnIndex': 0,
                                            'endColumnIndex': 9,
                                        }
                                    ],
                                    "booleanRule": {
                                        "condition": {
                                            "type": "CUSTOM_FORMULA",
                                            "values": [{
                                                "userEnteredValue": "=$I" + str(row) + "=TRUE"
                                            }]
                                        },
                                        "format": {
                                            "backgroundColor": {
                                                "red": 0.5,
                                                "green": 0.5,
                                                "blue": 0.5,
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    )

                handle_error(ss.batch_update, body)
                print("Formatting updated on " + po_number + "\n-----------------------------\n")
        if allow_transcribe == True:            
            source = str(path.realpath("Input Docs\\" + file_name))     
            destination = str(path.dirname(path.realpath(file_name))) + "\\Transcribed Pages"
            if not path.exists(destination):
                makedirs(destination)
            move(source, destination)

end = time.time()
print(f"Runtime = {end - start}")