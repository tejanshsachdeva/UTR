
## Bank Statement Processing Tool

### Overview

Processing bank statements with multiple accounts can be cumbersome and time-consuming, especially when dealing with PDF files. This tool simplifies the extraction of relevant information from PDF bank statements, such as UTR, date, debit amount, and other details, enabling easier data handling and visualization.

### How to Run the Code

#### Step 1: Download and Extract the Zip Folder

- Download the project zip folder from [this link](https://shorturl.at/2hijM).
- Extract the contents of the zip folder to your desired location.

#### Step 2: Install Python

- If you do not have Python installed on your system, download it from [this link](https://www.python.org/downloads/) and follow the installation instructions.

#### Step 3: Specify Folder Path

- Create a `config.json` file in the root directory of the project.
- Specify the folder path containing your PDF files in the `config.json` file.
  ```json
  {
      "SBI_FOLDER_PATH": "C:\\Users\\tejan\\Downloads\\UTR\\SBI\\"
  }
  ```

#### Step 4: Install Dependencies

- Open your terminal or command prompt.
- Navigate to the project directory.
- Run the following command to install the necessary dependencies:
  ```bash
  python -m pip install --upgrade pip
  pip install -r requirements.txt
  ```

#### Step 5: Execute the Script

- Run the main script to process the bank statements:
  ```bash
  python citi_code.py    #<- to process SBI Bank Statements
  python sbi_code.py     #<- to process CITI Bank Statements
  ```
- The output will be saved as `processed_statement_bank-name.xlsx` in the specified folder.

### How It Works

#### Extract Tables from PDFs

- The script iterates over each PDF in the specified directory, opens it with `pdfplumber`, and extracts tables from each page.
- The account number is added to each row of the extracted tables.

#### Extract Account Number

- The script uses PyMuPDF (`fitz`) to read the text from the PDF pages and extracts the account number.

#### Process Extracted Tables

- The extracted tables are processed to:
  - Ensure the 'Value Date' is correctly formatted.
  - Extract relevant columns and drop unnecessary ones.
  - Extract and format UTR numbers.
  - Split the reference number to obtain the vendor account number and vendor name.
  - Convert columns to appropriate data types.

### Handling Different Bank Statements

- The tool includes separate scripts for different banks:
  - **SBI Statements**: Processed using the `process_sbi_statements.py` script.
  - **Citi Bank Statements**: Since Citi bank statements are primarily in Excel format, users need to specify the path containing Citi bank Excel files. The script reads the transaction data from the "remittance information" column in the source Excel and creates an output file named `processed_statement_citi.xlsx`.

### Configuration Example

`config.json` file example:

```json
{
    "CITI_FOLDER_PATH": "C:\\Users\\tejan\\Downloads\\UTR\\CITI\\",
    "SBI_FOLDER_PATH": "C:\\Users\\tejan\\Downloads\\UTR\\SBI\\"
}
```

### Output

- The output CSV file will be named `processed_statement_bank-name.xlsx` for SBI statements and `processed_statement_citi.xlsx` for Citi bank statements.

Now you can easily work on the consolidated data for any visualizations or further analysis.
