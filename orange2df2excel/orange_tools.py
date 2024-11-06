from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from koboextractor import KoboExtractor
import requests
import pandas as pd
import os
from io import StringIO
import hashlib

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

def fetch_kobo_data(token, form_id, base_url="https://kf.kobotoolbox.org/api/v2"):
    """
    Fetch data from KoBoToolbox for a specified form and load it into a DataFrame using KoboExtractor.
    
    Parameters:
    - token (str): API token for KoBoToolbox.
    - form_id (str): The unique identifier of the form to fetch data from.
    - base_url (str): The base URL for the KoBoToolbox API. Default is for KoBoToolbox.

    Returns:
    - df (pandas.DataFrame): Data from KoBoToolbox in a DataFrame format.
    """
    try:
        # Initialize KoboExtractor with token and base URL
        kobo = KoboExtractor(token, base_url)
        
        # Fetch the data for the specified form
        print("Fetching data from KoBoToolbox...")
        data = kobo.get_data(form_id)
        
        # Convert the data to a DataFrame
        df = pd.json_normalize(data['results'])
        
        print("Data fetched successfully!")
        return df

    except Exception as err:
        print(f"Error fetching data: {err}")

def fetch_surveycto_data(isDataset, servername, form_or_dataset_id, username, password):
    """
    Fetch data from SurveyCTO for a specified form or dataset and load it into a DataFrame.
    
    Parameters:
    - isDataset (bool): If True, fetches data from a dataset; if False, fetches data from a form.
    - servername (str): The SurveyCTO server name (without "https://").
    - form_or_dataset_id (str): The unique ID of the form or dataset to fetch data from.
    - username (str): The SurveyCTO username for authentication.
    - password (str): The SurveyCTO password for authentication.

    Returns:
    - df (pandas.DataFrame): Data from SurveyCTO in a DataFrame format.
    """
    if isDataset:
        endpoint = f"https://{servername}.surveycto.com/api/v2/datasets/data/csv/{form_or_dataset_id}"
    else:
        endpoint = f"https://{servername}.surveycto.com/api/v1/forms/data/csv/{form_or_dataset_id}"
    
    try:
        auth = (username, password)
        
        print("Fetching data from SurveyCTO...")
        response = requests.get(endpoint, auth=auth)
        response.raise_for_status()

        df = pd.read_csv(StringIO(response.text))
        
        print("Data fetched successfully!")
        return df

    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
    except Exception as err:
        print(f"Other error occurred: {err}")

def generate_bnf_id(name, surname, dob):
    """
    Generates a unique beneficiary ID with a hash as the final component.

    Parameters:
        name (str): First name of the person.
        surname (str): Last name of the person.
        dob (str): Date of birth in 'YYYY-MM-DD' format.

    Returns:
        str: Generated unique beneficiary ID with hash included.
    """
    surname_length = len(surname)
    surname_part = surname[:3].upper().ljust(3, 'X')  # Pads with 'X' if fewer than 3 letters
    name_part = name[:3].upper().ljust(3, 'X')
    
    # Convert DOB from 'YYYY-MM-DD' to 'DDMMYY' format
    dob_parts = dob.split("-")
    dob_formatted = dob_parts[2] + dob_parts[1] + dob_parts[0][2:]  # DDMMYY format

    to_hash = f'{surname}{name}{dob_parts}'
    
    base_id = f"{surname_length}-{surname_part}-{name_part}-{dob_formatted}"
    hash_suffix = hashlib.md5(to_hash.encode()).hexdigest()
    beneficiary_id = f"{base_id}-{hash_suffix}"
    
    return beneficiary_id

