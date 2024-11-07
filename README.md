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

---

### Function: `gen_encryption_key`

The `gen_encryption_key` function generates a 32-byte AES encryption key using a provided password and a randomly generated salt. It employs PBKDF2 for secure key derivation, ensuring a unique key for each password-salt combination.

#### Parameters

- `password` (str): The password or passphrase used for key derivation.

#### Returns

- `formatted` (str): A formatted string showing both the derived key and salt values:
  - `Key`: The 32-byte derived encryption key.
  - `Salt`: The random 32-byte salt used during key derivation, which should be stored securely to allow re-derivation of the key if needed.

#### Example

```python
# Example usage of gen_encryption_key
password = "my_secure_password"
formatted_key_salt = gen_encryption_key(password)
print(formatted_key_salt)  # Output: Key: b'...' Salt: b'...'
```

This function provides a secure way to generate and display the key and salt, which can be stored for later re-derivation.

---

### Function: `encrypt_value`

The `encrypt_value` function encrypts a given value using AES encryption in CBC mode with a randomly generated initialization vector (IV). It supports encrypting strings and numbers by converting them to strings before encryption.

#### Parameters

- `value` (str, int, float): The plaintext value to encrypt, as a string or number.
- `encoded_key` (str): The Base64-encoded 32-byte AES encryption key.

#### Returns

- `encrypted_value` (str): The base64-encoded encrypted value, which includes both the IV and ciphertext for secure storage or transmission.

#### Example

```python
# Example usage of encrypt_value
value_to_encrypt = "sensitive_data"
encoded_key = "Base64_encoded_key_here"  # Base64-encoded AES key
encrypted_value = encrypt_value(value_to_encrypt, encoded_key)
print(encrypted_value)  # Output: Base64 encoded encrypted value
```

This function allows for secure encryption of sensitive information, producing a base64-encoded output that is convenient for storage or transfer.

---

### Function: `decrypt_value`

The `decrypt_value` function decrypts an AES-encrypted value using CBC mode. It requires the base64-encoded encrypted data (which includes both IV and ciphertext) and the original AES key used for encryption.

#### Parameters

- `encrypted_value` (str): The base64-encoded encrypted value, containing both the IV and ciphertext.
- `encoded_key` (str): The Base64-encoded 32-byte AES decryption key.

#### Returns

- `decrypted_value` (str): The decrypted plaintext value as a string.

#### Example

```python
# Example usage of decrypt_value
encrypted_value = "Base64_encrypted_value_here..."
encoded_key = "Base64_encoded_key_here"  # Base64-encoded AES key used during encryption
decrypted_value = decrypt_value(encrypted_value, encoded_key)
print(decrypted_value)  # Output: Original plaintext value
```

This function retrieves the original plaintext by decrypting the stored encrypted data with the correct AES key.

---

### Function: `rederive_key`

The `rederive_key` function re-derives the original AES encryption key using the same password and salt used in the initial derivation. This is useful for accessing the encryption key without storing it directly.

#### Parameters

- `password` (str): The original password or passphrase used for key derivation.
- `salt` (bytes): The salt used during the initial key derivation.

#### Returns

- `encoded_key` (str): The Base64-encoded 32-byte re-derived encryption key.

#### Example

```python
# Example usage of rederive_key
password = "my_secure_password"
salt = b'stored_salt_here...'  # The stored salt from the initial key generation
encoded_key = rederive_key(password, salt)
print(encoded_key)  # Output: Base64-encoded re-derived 32-byte key
```

This function is helpful for regenerating the encryption key using the stored salt and password, providing the key as a Base64-encoded string for compatibility with encryption/decryption functions.

## Requirements

- **pandas**
- **openpyxl**
- **koboextractor**
- **cryptography**
- **pycryptodome**

## License

This project is licensed under the MIT License.