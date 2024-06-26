import json
import os
import re
import pandas as pd
import pdfplumber
import fitz  
import xlsxwriter

# Load folder paths from config.json
with open('config.json', 'r') as config_file:
    config = json.load(config_file)
folder_path = config['SBI_FOLDER_PATH']

def extract_tables_from_pdfs(directory):
    tables = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    account_number = extract_account_number_fitz(pdf_path)
                    for page in pdf.pages:
                        pdf_tables = page.extract_tables()
                        for table in pdf_tables:
                            if table and len(table) > 1:
                                df = pd.DataFrame(table[1:], columns=table[0])
                                df["Account Number"] = account_number  # Add account number to each row
                                tables.append(df)
            except Exception as e:
                print(f"Error extracting tables from {filename}: {e}")
    return tables

def extract_account_number_fitz(pdf_path):
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        for line in text.split('\n'):
            if "Account Number" in line:
                account_number = line.split(':')[-1].strip()
                return account_number
    return None

def extract_account_number_from_table(df):
    for column in df.columns:
        for value in df[column]:
            if isinstance(value, str) and re.match(r'\d{15,16}', value):
                return value
    return None

def process_tables(tables):
    df_list = []
    for table in tables:
        if isinstance(table, pd.DataFrame):
            if "Account Number" not in table.columns or table["Account Number"].isnull().all():
                account_number = extract_account_number_from_table(table)
                table["Account Number"] = account_number
            df_list.append(table)

    if not df_list:
        print("No tables found in the PDFs.")
        return None

    combined_df = pd.concat(df_list, ignore_index=True)

    # Ensure the 'Value Date' column is properly formatted
    if 'Value Date' in combined_df.columns:
        combined_df['Value Date'] = combined_df['Value Date'].apply(lambda x: re.findall(r'\d{2}/\d{2}/\d{4}', x)[0] if re.findall(r'\d{2}/\d{2}/\d{4}', x) else None)

    combined_df.columns = ["Txn Date", "Value Date", "Description", "Ref No./Cheque No.", "Branch Code", "Debit", "Credit", "Balance", "Account Number"]

    combined_df = combined_df[combined_df["Description"].notnull()]
    combined_df = combined_df[combined_df["Debit"].notnull() | combined_df["Credit"].notnull()]

    combined_df = combined_df[(combined_df['Description'].notnull()) & (combined_df['Description'] != "Description")]

    columns_to_remove = ["Credit", "Balance"]
    combined_df = combined_df.drop(columns=[col for col in columns_to_remove if col in combined_df.columns])

    combined_df = combined_df[combined_df['Description'].str.contains("UTR")]

    combined_df["RTGS / NEFT"] = combined_df["Description"].str.extract(r'(NEFT|RTGS)')[0]
    combined_df["RTGS / NEFT"] = combined_df["RTGS / NEFT"].fillna("NEFT")  # Default to NEFT if blank

    combined_df["Description"] = combined_df["Description"].str.replace('\n', ' ')

    # Enhanced UTR extraction to handle both single-line and multi-line cases
    combined_df["UTR"] = combined_df["Description"].str.extract(r': (\w+)-')
    missing_utr_df = combined_df[combined_df["UTR"].isnull()]
    for index, row in missing_utr_df.iterrows():
        multiline_match = re.search(r'RTGS UTR NO[:\s]*\n*\s*(\w+)', row["Description"], re.IGNORECASE)
        if multiline_match:
            combined_df.at[index, "UTR"] = multiline_match.group(1)
        else:
            print(f"Failed to extract UTR from Description: {row['Description']}")

    # Debugging: Check rows where UTR is still missing
    missing_utr_df = combined_df[combined_df["UTR"].isnull()]
    print("Missing UTR rows:")
    print(missing_utr_df.head(10))

    combined_df = combined_df.drop(columns=["Description"])

    if 'Ref No./Cheque No.' in combined_df.columns:
        combined_df["Ref No./Cheque No."] = combined_df["Ref No./Cheque No."].str.replace('\n', ' ')
        combined_df["Ref No./Cheque No."] = combined_df["Ref No./Cheque No."].str.replace('TRANSFER TO ', '')
        combined_df[['Vendor Account Number', 'Vendor Name']] = combined_df["Ref No./Cheque No."].str.split(' / ', expand=True)
        combined_df = combined_df.drop(columns=["Ref No./Cheque No."])
    else:
        combined_df['Vendor Account Number'] = None
        combined_df['Vendor Name'] = None

    combined_df["Debit"] = combined_df["Debit"].str.replace('\n', '').str.replace('\r', '')

    combined_df = combined_df[combined_df["Debit"].str.replace(',', '').str.replace('.', '', regex=False).str.isdigit()]

    combined_df["Vendor Account Number"] = combined_df["Vendor Account Number"].astype(str)
    combined_df["Value Date"] = pd.to_datetime(combined_df["Value Date"], format='%d/%m/%Y', errors='coerce')
    combined_df["Debit"] = combined_df["Debit"].str.replace(',', '').astype(float)

    final_df = combined_df[["Account Number", "Value Date", "Vendor Account Number", "Vendor Name", "UTR", "RTGS / NEFT", "Debit"]]

    return final_df

# Directory containing PDF files

# Extract tables from PDFs
tables = extract_tables_from_pdfs(folder_path)
print(f"Extracted {len(tables)} tables from PDFs.")

# Process tables
processed_table = process_tables(tables)    

if processed_table is not None:
    # Save the processed table to an Excel file
    processed_table.to_excel("processed_statement_sbi.xlsx", index=False, engine='xlsxwriter')
    print("Data exported to processed_statement_sbi.xlsx")
else:
    print("No data to export.")
