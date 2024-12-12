# orange2df2excel

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

### Function: `raw_data_to_excel_with_all_charts`

The `raw_data_to_excel_with_all_charts` function generates an Excel file containing raw data and a dashboard. The dashboard includes various charts (e.g., bar, line, pie, doughnut) and a section for total values. It is a flexible tool for creating visual summaries of data directly in Excel using `xlsxwriter`.

#### Parameters

- `df` (*pandas.DataFrame*): The DataFrame to be written to the Excel file.
- `file_path` (*str*): The path where the Excel file will be saved.
- `chart_config` (*dict*): A dictionary to configure the charts in the dashboard. The dictionary keys are chart types (e.g., `"bar"`, `"line"`, `"pie"`, `"doughnut"`). Each key maps to a dictionary with the following keys:
  - `'category_col'` (*str*): Column name to use as categories (e.g., x-axis or labels).
  - `'value_col'` (*str*): Column name to use as values (e.g., y-axis or data points).
- `totals` (*list, optional*): A list of column names for which totals are calculated and displayed at the top of the dashboard. If `None`, totals will not be displayed. 
  - Numeric columns: Totals are calculated as the sum.
  - Non-numeric columns: Totals represent the count of occurrences for each unique value.

#### Returns

- None. The function saves the Excel file at the specified `file_path`.

#### Behavior

1. **Raw Data Sheet**:
   - Writes the full DataFrame to the "Raw Data" sheet in the Excel file.

2. **Dashboard Sheet**:
   - **Totals Section**: Displays sums for numeric columns and counts for non-numeric columns (if specified in `totals`).
   - **Charts**: Adds charts to visualize data based on the provided `chart_config`. 
     - Bar and line charts include data labels with both category names and values (e.g., `"A: 42"`).
     - Pie and doughnut charts display percentages along with category names.

#### Example

```python
import pandas as pd

# Example data
data = {
    "Donor": ["A", "B", "A", "C", "A", "B", "C", "A", "A", "B"],
    "Sessions": [10, 5, 15, 20, 8, 12, 5, 15, 10, 8],
    "Cost": [100, 50, 75, 120, 80, 130, 60, 100, 150, 70],
}
df = pd.DataFrame(data)

# Chart configuration
chart_config = {
    "bar": {"category_col": "Donor", "value_col": "Sessions"},
    "pie": {"category_col": "Donor", "value_col": "Sessions"},
}

# Specify totals
totals = ["Sessions", "Cost", "Donor"]

# Call the function
raw_data_to_excel_with_all_charts(df, "dashboard_with_totals.xlsx", chart_config, totals)
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

### Function: `generate_bnf_id`

The `generate_bnf_id` function creates a unique beneficiary ID based on the person's name, surname, and date of birth. This ID structure includes specific information about the beneficiary, such as surname length, initials, formatted date of birth, and a unique hash component to ensure uniqueness.

#### Parameters

- `name` (str): The first name of the beneficiary.
- `surname` (str): The last name of the beneficiary.
- `dob` (str): The beneficiary's date of birth, formatted as `YYYY-MM-DD`.

#### Returns

- `beneficiary_id` (str): A unique ID for the beneficiary, structured as follows: 
  - The length of the surname.
  - The first three characters of the surname (padded with 'X' if fewer than three characters).
  - The first three characters of the name (padded with 'X' if fewer than three characters).
  - The date of birth in `DDMMYY` format.
  - A hash of the generated ID components to ensure uniqueness.

#### Example

```python
# Example usage of generate_bnf_id
name = "John"
surname = "Smith"
dob = "1990-01-01"

# Generate unique beneficiary ID
beneficiary_id = generate_bnf_id(name, surname, dob)

# Output the generated ID
print(beneficiary_id)  # Example output: "5-SMI-JOH-010190-a1b2c3d4e5f67890abcd1234567890ef"
```

This function ensures each generated ID is unique by combining structured personal data with a full hash component, allowing for consistency and minimizing the chance of duplicates even with similar input data.

### Function: `generate_session_id`

The `generate_session_id` function creates a unique session identifier based on donor name, settlement location, enumerator name, submission date, and session date. This ID structure combines key details of the session, ensuring a structured format and uniqueness for each record.

#### Parameters

- `df` (DataFrame): The input DataFrame containing the relevant session data.
- `donor_name` (str): The column name in the DataFrame that stores the donor's name.
- `location_settlement` (str): The column name in the DataFrame that stores the settlement or location name.
- `name_enumerator` (str): The column name in the DataFrame that stores the enumerator's name.
- `submission_date` (str): The column name in the DataFrame that stores the submission date, formatted as `YYYY-MM-DD`.
- `session_date` (str): The column name in the DataFrame that stores the session date, formatted as `YYYY-MM-DD`.
- `project_name` (str): A static string representing the project's name to be included in the session ID.

#### Returns

- `DataFrame`: The original DataFrame with an additional column, `session_id_sql`, containing unique session IDs for each record. The session ID is structured as follows:
  - The donor name (with spaces, colons, and commas removed, and converted to uppercase).
  - The project name.
  - The settlement location (processed similarly to the donor name).
  - The first three characters of the enumerator's name (processed similarly, padded with 'X' if fewer than three characters).
  - The submission date in `DDMMYY` format (with spaces, colons, and commas removed, and converted to uppercase).
  - The session date in `DDMMYY` format (processed similarly).

#### Example

```python
# Example DataFrame
import pandas as pd

data = {
    'donor_name': ['Donor A', 'Donor B'],
    'location_settlement': ['Settlement X', 'Settlement Y'],
    'name_enumerator': ['John Doe', 'Jane Smith'],
    'submission_date': ['2024-12-10', '2024-12-11'],
    'session_date': ['2024-12-09', '2024-12-10']
}
df = pd.DataFrame(data)

# Generate unique session IDs
project_name = 'EORE'
df = generate_session_id(df, 'donor_name', 'location_settlement', 'name_enumerator', 'submission_date', 'session_date', project_name)

# Output the DataFrame
print(df)
```

#### Example Output

| donor_name | location_settlement | name_enumerator | submission_date | session_date | session_id_sql                 |
|------------|---------------------|-----------------|-----------------|--------------|--------------------------------|
| Donor A    | Settlement X        | John Doe        | 2024-12-10      | 2024-12-09   | DONORA-EORE-SETTLEMENTX-JOH-101224-091224 |
| Donor B    | Settlement Y        | Jane Smith      | 2024-12-11      | 2024-12-10   | DONORB-EORE-SETTLEMENTY-JAN-111224-101224 |

This function ensures that each session ID is unique and follows a standardized format, making it easy to identify and track individual sessions.

### Function: `gen_encryption_key`

The `gen_encryption_key` function generates a 32-byte AES encryption key using a provided password and a randomly generated salt. This function uses PBKDF2 for key derivation to ensure a secure and unique key for each password-salt combination.

#### Parameters

- `password` (str): The password or passphrase used for key derivation.

#### Returns

- `formatted` (str): A formatted string showing both the derived key and salt values. This string includes:
  - `Key`: The 32-byte derived encryption key.
  - `Salt`: The random 32-byte salt used during key derivation, which should be securely stored to allow re-derivation of the key if needed.

#### Example

```python
# Example usage of gen_encryption_key
password = "my_secure_password"
formatted_key_salt = gen_encryption_key(password)
print(formatted_key_salt)  # Output: Key: b'...' Salt: b'...'
```

This function is useful for securely generating and displaying the key and salt, which can be stored securely for later use.

---

### Function: `encrypt_value`

The `encrypt_value` function encrypts a given value using AES encryption in CBC mode with a random initialization vector (IV). The function supports encrypting both strings and numbers, converting them to strings before encryption.

#### Parameters

- `value` (str, int, float): The plaintext value to encrypt. It can be either a string or a number.
- `key` (bytes): The 32-byte AES encryption key to use for encryption.

#### Returns

- `encrypted_value` (str): The base64-encoded encrypted value, which includes the IV and ciphertext for secure storage or transmission.

#### Example

```python
# Example usage of encrypt_value
value_to_encrypt = "sensitive_data"
key = b'some_32_byte_key_here...'  # Example key
encrypted_value = encrypt_value(value_to_encrypt, key)
print(encrypted_value)  # Output: Base64 encoded encrypted value
```

This function allows for secure encryption of sensitive information, returning a base64-encoded string for easy storage or transfer.

---

### Function: `decrypt_value`

The `decrypt_value` function decrypts a given encrypted value using AES encryption in CBC mode. It requires the base64-encoded encrypted value and the AES key that was originally used to encrypt the value.

#### Parameters

- `encrypted_value` (str): The base64-encoded encrypted value, containing both the IV and ciphertext.
- `key` (bytes): The 32-byte AES decryption key.

#### Returns

- `decrypted_value` (str): The decrypted plaintext value as a string.

#### Example

```python
# Example usage of decrypt_value
encrypted_value = "Base64_encrypted_value_here..."
key = b'some_32_byte_key_here...'  # Example key used during encryption
decrypted_value = decrypt_value(encrypted_value, key)
print(decrypted_value)  # Output: Original plaintext value
```

This function is essential for retrieving the original plaintext data by decrypting the stored encrypted value with the correct AES key.

---

### Function: `rederive_key`

The `rederive_key` function re-derives the original AES encryption key using the same password and salt that were used in the initial derivation. This is useful for accessing the encryption key without storing it directly.

#### Parameters

- `password` (str): The original password or passphrase used for key derivation.
- `salt` (bytes): The salt that was originally used during the initial key derivation.

#### Returns

- `key` (bytes): The re-derived 32-byte encryption key.

#### Example

```python
# Example usage of rederive_key
password = "my_secure_password"
salt = b'stored_salt_here...'  # Example stored salt from the initial key generation
key = rederive_key(password, salt)
print(key)  # Output: The re-derived 32-byte key
```

This function is particularly useful for securely regenerating the encryption key using the stored salt and password without needing to store the key directly.

---

### Function: `encrypt_file`

The `encrypt_file` function encrypts the contents of a specified file using AES encryption in CBC mode with PKCS7 padding. The encrypted file will include a random initialization vector (IV) at the beginning, which is necessary for decryption.

#### Parameters

- `input_file_path` (str): The path to the file that needs to be encrypted.
- `output_file_path` (str): The path where the encrypted file will be saved.
- `key` (bytes): A 32-byte AES encryption key.

#### Returns

- None

#### Example

```python
encrypt_file("plaintext.txt", "encrypted_file.en", key)
```

This function reads the input file in chunks, applies padding, and writes the encrypted data along with the IV at the beginning of the output file for secure storage or transfer.

#### Notes

- The IV is stored at the beginning of the encrypted file, which is required for decryption.
- The file is encrypted in chunks for efficient memory usage.

---

### Function: `decrypt_file`

The `decrypt_file` function decrypts a file that was encrypted using AES encryption in CBC mode and PKCS7 padding. The function reads the IV from the start of the encrypted file and removes padding after decryption to recover the original file contents.

#### Parameters

- `encrypted_file_path` (str): The path to the encrypted file that needs to be decrypted.
- `output_file_path` (str): The path where the decrypted file will be saved.
- `key` (bytes): A 32-byte AES decryption key.

#### Returns

- None

#### Example

```python
decrypt_file("encrypted_file.en", "decrypted_file.txt", key)
```

This function reads the encrypted file in chunks, decrypts it, and removes padding from the decrypted data to restore the original file contents.

#### Notes

- The function reads the IV from the start of the encrypted file; if the IV is missing or invalid, it will raise a `ValueError`.
- The decrypted output will be saved at the specified path, restoring the file to its original contents.

---

### Function: `hash_password`

The `hash_password` function securely hashes a plain-text password using bcrypt, incorporating a salt to ensure unique hash results for identical passwords. The generated hash can be stored and used to verify passwords.

#### Parameters

- `password` (str): The plain-text password to be hashed.

#### Returns

- `str`: The bcrypt hash of the password, encoded as a string.

#### Example

```python
hashed_password = hash_password("my_secure_password")
```

## Requirements

- **pandas**
- **openpyxl**
- **koboextractor**
- **cryptography**
- **pycryptodome**
- **bcrypt**

## License

This project is licensed under the MIT License.