from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def raw_data_to_excel(df, file_path, sheet_name):
    """
    Write a DataFrame to an Excel file in table format.
    
    Parameters:
    - df: pandas.DataFrame - The DataFrame to write to Excel.
    - file_path: str - Path to the Excel file.
    - sheet_name: str - Name of the sheet to write data to.
    """
    if os.path.exists(file_path):
        workbook = load_workbook(file_path)
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
    else:
        workbook = Workbook()
        if 'Sheet' in workbook.sheetnames:
            del workbook['Sheet']

    worksheet = workbook.create_sheet(sheet_name)

    for row in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row)

    table = Table(displayName="raw_data", ref=worksheet.dimensions)
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=True
    )
    table.tableStyleInfo = style
    worksheet.add_table(table)

#-----------Adjusting cells--------------
    for col in worksheet.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Get the column letter
        for cell in col:
            if cell.value:
                # Estimate width by multiplying character count by a width factor
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[col_letter].width = adjusted_width

    workbook.save(file_path)