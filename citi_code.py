import json
import os
import re
import pandas as pd

# Load folder paths from config.json
with open('config.json', 'r') as config_file:
    config = json.load(config_file)
folder_path = config['CITI_FOLDER_PATH']

# List all Excel files in the folder
excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xls') or f.endswith('.xlsx')]

# Check if there are any Excel files in the folder
if not excel_files:
    raise ValueError(f"No Excel files found in folder: {folder_path}")

# Define expected columns
expected_columns = [
    'Account Number', 'Value Date', 'Customer Reference', 'Bank Reference',
    'Remittance Information', 'Transaction Amount - Debit', 'Debit / Credit'
]

# Function to match columns dynamically
def match_columns(df):
    for col in df.columns:
        for expected_col in expected_columns:
            if expected_col.lower() in str(col).lower():
                return True
    return False

# Initialize an empty DataFrame to store combined data
combined_df = pd.DataFrame()

# Process each Excel file
for file_name in excel_files:
    file_path = os.path.join(folder_path, file_name)
    
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Remove unnecessary top rows and columns
    df = df.iloc[2:, :]

    # Reset index and use the 4th row as headers
    df.columns = df.iloc[0]
    df = df.drop(2).reset_index(drop=True)

    # Clean up column names by removing NaN and trimming whitespace
    df.columns = df.columns.fillna('').str.strip()

    # Print column names for debugging
    print(f"Processing file: {file_name}")
    print("Column names after reset:")
    print(df.columns)

    # Check if the dataframe has all expected columns
    if not match_columns(df):
        print(f"Columns missing in DataFrame: {expected_columns}")
        continue  # Skip processing this file if columns are missing

    # Filter to keep only the necessary columns
    df = df[expected_columns]

    # Example: Filtering rows where Debit/Credit is "D"
    df = df[df['Debit / Credit'] == 'D']

    # Extract UTR number and vendor name from Remittance Information
    def extract_utr(text):
        if isinstance(text, str):
            match = re.search(r'UTR\s+(\w+)', text)
            return match.group(1) if match else ''
        return ''

    def extract_vendor_name(text):
        if isinstance(text, str):
            match = re.search(r'TRF\s+TO\s+(.+)', text)
            return match.group(1) if match else ''
        return ''

    # Apply extraction functions
    df['UTR'] = df['Remittance Information'].apply(extract_utr)
    df['Vendor Name'] = df['Remittance Information'].apply(extract_vendor_name)

    # Remove rows where UTR is null or empty
    df = df[df['UTR'].notna() & (df['UTR'] != '')]

    # Remove unnecessary columns
    df.drop(columns=['Remittance Information', 'Debit / Credit'], inplace=True)

    # Reorder columns
    df = df[['Account Number', 'Value Date', 'Vendor Name', 'Customer Reference', 'Bank Reference', 'UTR', 'Transaction Amount - Debit']]

    # Remove '-' from the Transaction Amount - Debit
    df['Transaction Amount - Debit'] = df['Transaction Amount - Debit'].astype(str).str.replace('-', '')

    # Append the processed DataFrame to combined_df using concat
    combined_df = pd.concat([combined_df, df], ignore_index=True)

# Save to new Excel file
output_file_path = 'processed_statement_citi.xlsx'
combined_df.to_excel(output_file_path, index=False)

print(f"All files processed and combined. The output file is saved as '{output_file_path}'.")
