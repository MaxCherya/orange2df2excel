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

### Function: `fetch_surveycto_data`

The `fetch_surveycto_data` function retrieves data from a specified SurveyCTO form or dataset and loads it into a pandas DataFrame, allowing for easy analysis and manipulation within Python. This function dynamically adjusts the API endpoint based on whether you are fetching from a form or a dataset, making it flexible for various data retrieval tasks on SurveyCTO.

#### Parameters

- `isDataset` (bool): Determines whether to fetch data from a dataset (`True`) or a form (`False`).
- `servername` (str): The SurveyCTO server name (excluding "https://"). For example, if your server URL is `https://yourserver.surveycto.com`, use `yourserver`.
- `form_or_dataset_id` (str): The unique ID of the form or dataset to retrieve data from. This ID can be found in the SurveyCTO dashboard.
- `username` (str): The SurveyCTO username for authentication.
- `password` (str): The SurveyCTO password for authentication.

#### Returns

- `df` (pandas.DataFrame): A DataFrame containing the fetched data, where each row represents a submission, and each column corresponds to a field in the form or dataset.

#### Example

```python
from orange2df2xcel import fetch_surveycto_data

# Define SurveyCTO credentials and parameters
is_dataset = True  # Set to False if fetching from a form
servername = "your_server_name"
form_or_dataset_id = "your_form_or_dataset_id"
username = "your_username"
password = "your_password"

# Fetch data from SurveyCTO and store it in a DataFrame
df = fetch_surveycto_data(is_dataset, servername, form_or_dataset_id, username, password)

# Display the data
print(df.head())
```

This function provides a straightforward interface for retrieving data from SurveyCTO, handling authentication and endpoint selection automatically, so you don’t need to manage API interactions manually.

## Requirements

- **pandas**
- **openpyxl**
- **koboextractor**

## License

This project is licensed under the MIT License.