from openpyxl import *
from openpyxl.styles import *
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from datetime import date
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment


def merge(master_sheet_filepath, formatted_dpp_filepath):
    print(f'Master filepath: {master_sheet_filepath}')
    print(f'Formatted DPP filepath: {formatted_dpp_filepath}')

    master_wb = load_workbook(master_sheet_filepath)
    master_ws = master_wb.active
    dpp_wb = load_workbook(formatted_dpp_filepath)
    dpp_ws = dpp_wb.active

    # Get unique dispatch numbers
    m_dispatch = set(cell.value for cell in master_ws['B'] if cell.value is not None)
    d_dispatch = set(cell.value for cell in dpp_ws['B'] if cell.value is not None)
    unique_to_dpp = d_dispatch - m_dispatch

    # Ignore columns that will contain manual entries
    ignore_cols = [
        "Missing Information", "Zone Uplift", 
        "Overnight", "Shift Uplift", "PM Confirmation", 
        "Corrected SKU", "Completed Date",
        "Scheduled Date", "Admin Status"
    ]
    ignore_cols_indices = [master_ws[1].index(cell) for cell in master_ws[1] if cell.value in ignore_cols]
    
    # Find the Missing Information column index
    missing_info_col_idx = None
    for idx, cell in enumerate(master_ws[1]): 
        if cell.value == "Missing Information":
            missing_info_col_idx = idx
            break

    if missing_info_col_idx is None:
        raise ValueError("Missing Information column not found in the header row.")
    
    # Append new rows to the master sheet
    # Apply styling and formatting to new rows
    new_rows = []
    if unique_to_dpp:
        for row in dpp_ws.iter_rows(min_row=2):
            if row[1].value in unique_to_dpp:
                row_values = [cell.value for cell in row]
                master_ws.append(row_values)
                new_row = master_ws[master_ws.max_row]
                style_row(new_row, ignore_cols_indices, missing_info_col_idx)
                new_rows.append(new_row)

    # Apply formatting only to new rows
    if new_rows:
        format_cells(new_rows)

    master_wb.save(master_sheet_filepath)
    print('Merging, styling, and formatting of new rows complete.')


# Style new rows
def style_row(row, ignore_cols_indices, missing_info_col_idx):
    green_fill = PatternFill(start_color="FF90EE90", end_color="FF90EE90", fill_type="solid")
    red_fill = PatternFill(start_color="FFFFCCCB", end_color="FFFFCCCB", fill_type="solid")
    has_missing = False

    for idx, cell in enumerate(row):
        if idx in ignore_cols_indices:
            continue
        cell.fill = green_fill
        if cell.value is None:
            cell.fill = red_fill
            has_missing = True

    if has_missing:
        row[missing_info_col_idx].value = "True"
        row[missing_info_col_idx].fill = red_fill


# Format new cells
def format_cells(rows):
    # Apply text wrap and auto-fit column widths only to specified rows
    for row in rows:
        for cell in row:
            cell.alignment = Alignment(wrapText=True)
            max_length = len(str(cell.value)) if cell.value else 0
            adjusted_width = (max_length + 2) * 1.1
            cell.parent.column_dimensions[cell.column_letter].width = max(adjusted_width, cell.parent.column_dimensions[cell.column_letter].width)