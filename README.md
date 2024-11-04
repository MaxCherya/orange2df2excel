# orange2df2xcel

A Python package with tools for working with DataFrames and saving them to Excel in a structured format.

## Installation

To install this package directly from GitHub, use:

```bash
pip install git+https://github.com/MaxCherya/orange2df2excel.git
```

## Usage

### Function: `raw_data_to_excel`

The `raw_data_to_excel` function allows you to save a pandas DataFrame to an Excel file with automatic table formatting and sheet management. If the specified Excel file exists, it will replace or add the designated sheet; if it doesn’t exist, the function creates a new file.

#### Parameters

- `df` (pandas.DataFrame): The DataFrame you want to save to Excel.
- `file_path` (str): The path to the Excel file where the DataFrame will be saved.
- `sheet_name` (str): The name of the sheet in which to write the data.

#### Example

```python
from orange2df2excel import raw_data_to_excel

# Example usage of raw_data_to_excel
raw_data_to_excel(df, "example.xlsx", "raw data")
```

### Function: `fetch_kobo_data`

The `fetch_kobo_data` function retrieves data from a specified KoBoToolbox form and loads it into a pandas DataFrame, making it easy to analyze and manipulate within Python. This function uses `KoboExtractor` to streamline the process and handle API interactions.

#### Parameters

- `token` (str): The API token for authenticating access to KoBoToolbox.
- `form_id` (str): The unique identifier of the form to retrieve data from. You can find this ID in your KoBoToolbox form settings.
- `base_url` (str, optional): The base URL for the KoBoToolbox API. It defaults to the standard KoBoToolbox URL, but you can specify a different base URL if needed.

#### Returns

- `df` (pandas.DataFrame): A DataFrame containing the fetched data, with each row representing a submission and each column a survey question or field.

#### Example

```python
from orange2df2xcel import fetch_kobo_data

# Example usage of fetch_kobo_data
api_token = "your_kobo_api_token"
form_id = "your_form_id"

# Fetch data from KoBoToolbox and store it in a DataFrame
df = fetch_kobo_data(api_token, form_id)

# Display the data
print(df.head())
```

This function provides a simple interface for retrieving KoBoToolbox data into a format suitable for data analysis, without needing to handle the API response manually.

## Requirements

- **pandas**
- **openpyxl**

## License

This project is licensed under the MIT License.