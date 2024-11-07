from orange2df2excel.orange_tools import gen_encryption_key, encrypt_value, decrypt_value, rederive_key
from dotenv import load_dotenv
import os
import ast

load_dotenv()

# Retrieve and evaluate `EN_KEY` and `EN_SALT` as bytes from .env
key_str = os.getenv("EN_KEY")
salt_str = os.getenv("EN_SALT")

if key_str and salt_str:
    key = ast.literal_eval(key_str)
    salt = ast.literal_eval(salt_str)
else:
    raise ValueError("EN_KEY and/or EN_SALT not found in environment variables")


email = 'test@mail.com'
password = 'H@NDIC@P'


# Encrypt and decrypt the email
encrypted_email = encrypt_value(email, key)
print("Encrypted Email:", encrypted_email)

decrypted_email = decrypt_value(encrypted_email, key)
print("Decrypted Email:", decrypted_email)

# Re-derive the key using the password and salt
derived_key = rederive_key(password=password, salt=salt)
print("Derived Key:", derived_key)