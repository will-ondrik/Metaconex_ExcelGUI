from openpyxl import *
from openpyxl.styles import *
import pandas as pd
from datetime import date

def format(raw_dpp_filepath):

    # Read from the raw DPP Excel file
    df = pd.read_excel(raw_dpp_filepath, sheet_name='Sheet0')

    # Rename Column Names (Reflects Leah's Excel Sheet)
    df.rename(columns={
        "Service Address Line 1": "Service Address",
        "Service Postal Code": "Postal Code",
        "State": "Province",
        "Dispatch Status": "Status"
    }, inplace=True)

    # Delete Redundant Columns
    df.drop(columns=[
    "Service Provider", "Logistics Provider", 
    "LOB", "End Service Window", 
    "DPS Type", "Call Type", "Service Level",
    "Customer Number", "Customer Secondary Contact Name", 
    "Customer Secondary Contact Phone Number", 
    "Customer Secondary Contact Email Address", 
    "Service Address Line 2", "Service Address Line 3",
    "Service Address Line 4", "Part Number", "Part Quantity", 
    "Part Status", "Part Status Date", "Carrier Name",
    "Waybill Number", "Closure Date", "Warranty Invoice Date", 
    "Warranty Invoice Number", "Original Order BUID", 
    "Reply Code", "Service Request Type", "Product Classification", 
    "Report Description", "Service Type", "Engineer Assigned", 
    "Engineer Id", "Service Call Date",
    ], inplace=True)


   # Add New Columns for Manual Entry
    df["Zone Uplift"] = ''
    df["Overnight"] = ''
    df["Shift Uplift"] = ''
    df["PM Confirmation"] = ''
    df["Completed Date"] = ''
    df["Scheduled Date"] = ''
    df["Scheduling Notes"] = ''
    df["Missing Information"] = 'No'
    df["Admin Status"] = ''
    df["Notes"] = ''

    # Restructure Column Order
    new_column_order = [
        "Admin Status",
        "Missing Information", "Dispatch Number", 
        "Status", "Service Address", "City", 
        "Postal Code", "Province", "Country", 
        "Project Number", "Product Name",
        "Product Model", "Service Tag",
        "Service SKU", "Corrected SKU",
        "Comments to Vendor", "Zone Uplift",
        "Overnight", "Shift Uplift",
        "PM Contact", "PM Phone",
        "PM Email", "Customer Name",
        "Customer Contact Phone Number", 
        "Customer Contact Email Address",
        "Scheduled Date", "PM Confirmation",
        "Completed Date", "Notes"
    ]

    # Reorder the DataFrame columns
    df = df.reindex(columns=new_column_order) 


    # Remove Quantity from Service SKU
    text_to_remove = "Qty:1"
    df["Service SKU"] = df["Service SKU"].str.replace(text_to_remove, "")

    # Save the DataFrame to an Excel file
    current_date = date.today()
    filename = "workorders_.xlsx"
    df.to_excel(filename, sheet_name='Sheet0', index=False)

    # Load the formatted workbook
    wb = load_workbook(filename)
    ws = wb['Sheet0']

    

    # Save the formatted workbook
    wb.save(filename)

    # Return filename
    print(f'Returned filename: {filename}')
    return filename


