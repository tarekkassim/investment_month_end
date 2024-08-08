import pandas as pd
import numpy as np


def allBalances(monthYear):

    # Set display options to print all columns
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    # Specify the file path
    filePath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Source Reports\210 - Settled Assets Report (History) - ' + monthYear + ' (Original).csv'
    outputPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Read the Excel file without headers
    df = pd.read_csv(filePath, header=None)

    # Extract the columns which correspond to indices
    extracted_columns = df.iloc[:, [48, 65, 76, 77, 78, 81, 84, 87]]

    # Create a new DataFrame with the extracted data
    new_df = pd.DataFrame(extracted_columns)
    new_df.columns = ['Account Number', 'Type', 'Units', 'Issuer', 'Book Value', 'Market Value', 'CUSIP', 'Accrued Interest']

    # Drop currency hedge rows
    hedgerows = ['Account Number:  000147000CAD',
                 'Account Number:  000147000EUR',
                 'Account Number:  000147000GBP',
                 'Account Number:  000147000JPY',
                 'Account Number:  000147000USD',
                 'Account Number:  06-0000/7.1',
                 'Account Number:  06-0000/7.2',
                 'Account Number:  06-0000/7.3',
                 'Account Number:  06-0000/7.4',
                 'Account Number:  06-0000/7.5'
    ]

    for x in hedgerows:
        new_df = new_df[~new_df['Account Number'].str.contains(x)]

    #Drop ACM rows
    new_df = new_df[~new_df['Issuer'].str.contains('ACM', na=False)]

    # Keep only the last 9 characters of the 'Account Number' column
    new_df['Account Number'] = new_df['Account Number'].astype(str).str.replace('Account Number:', '').str.strip()

    # Keep only the last 12 characters of the 'CUSIP' column
    new_df['CUSIP'] = new_df['CUSIP'].astype(str).str[-12:]

    # Drop rows where 'Market Value' is NaN
    new_df = new_df.dropna(subset=['Market Value'])

    #convert data types
    new_df['Account Number'] = new_df['Account Number'].astype(int)
    new_df['Units'] = new_df['Units'].astype(str).str.replace(',', '') # Remove commas from the column
    new_df['Units'] = pd.to_numeric(new_df['Units'], errors='coerce')
    new_df['Book Value'] = new_df['Book Value'].astype(str).str.replace(',', '') # Remove commas from the column
    new_df['Book Value'] = pd.to_numeric(new_df['Book Value'], errors='coerce')
    new_df['Market Value'] = new_df['Market Value'].astype(str).str.replace(',', '') # Remove commas from the column
    new_df['Market Value'] = pd.to_numeric(new_df['Market Value'], errors='coerce')
    new_df['Accrued Interest'] = new_df['Accrued Interest'].astype(str).str.replace(',', '') # Remove commas from the column
    new_df['Accrued Interest'] = pd.to_numeric(new_df['Accrued Interest'], errors='coerce')

    # print(new_df.to_string())
    # print(new_df.dtypes)

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(outputPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        new_df.to_excel(writer, sheet_name='Balances - Current Month', index=False)