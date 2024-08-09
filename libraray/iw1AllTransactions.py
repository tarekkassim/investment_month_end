import pandas as pd
import numpy as np
from datetime import datetime, timedelta


def allTransactions(monthYear):

    def get_next_month_abbr(date_str):
        # Parse the input string to a datetime object
        current_date = datetime.strptime(date_str, '%b %Y')

        # Calculate the first day of the next month
        next_month = current_date.replace(day=28) + timedelta(days=4)  # this will definitely be in the next month
        next_month = next_month.replace(day=1)  # go to the first day of the next month

        # Format the next month in the desired format and extract the first three letters of the month
        next_month_str = next_month.strftime('%b %Y')

        return next_month_str


    nextMonth = get_next_month_abbr(monthYear)

    # Set display options to print all columns
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)

    # Specify the file path
    filePath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Source Reports\315 - Cash Transaction Report - ' + monthYear + ' (Original).csv'
    outputPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Read the Excel file without headers
    df = pd.read_csv(filePath, header=None)

    # Extract the columns which correspond to indices
    extracted_columns = df.iloc[:, [31, 41, 44, 45, 46, 47, 49, 48, 50, 52, 53]]

    # Create a new DataFrame with the extracted data
    new_df = pd.DataFrame(extracted_columns)
    new_df.columns = ['Account Number', 'Date', 'Transaction Type', 'Units', 'Issuer', 'Transaction Cash Value', 'CUSIP',
                      'Details', 'Transaction Type 2', 'Details 2', 'Transaction Cash Value 2']

    # Drop currency hedge & ACM rows
    hedgerows = ['Account Number:    000147000CAD',
                 'Account Number:    000147000EUR',
                 'Account Number:    000147000GBP',
                 'Account Number:    000147000JPY',
                 'Account Number:    000147000USD',
                 'Account Number:    06-0000/7.1',
                 'Account Number:    06-0000/7.2',
                 'Account Number:    06-0000/7.3',
                 'Account Number:    06-0000/7.4',
                 'Account Number:    06-0000/7.5',
                 'Account Number:    147122010'
                 ]

    for x in hedgerows:
        new_df = new_df[~new_df['Account Number'].str.contains(x)]

    # Keep only the last 9 characters of the 'Account Number' column
    new_df['Account Number'] = new_df['Account Number'].astype(str).str.replace('Account Number:', '').str.strip()

    # Keep only the last 12 characters of the 'CUSIP' column
    new_df['CUSIP'] = new_df['CUSIP'].astype(str).str[-12:]

    # Drop rows where 'Transaction Type' is NaN
    # new_df = new_df.dropna(subset=['Transaction Type'])

    # Drop Transactions rows which are not unsettled
    new_df = new_df[
        ~((new_df['Date'].str.contains(nextMonth)) & (new_df['Details'].str.contains('PAID INTEREST|TREASURY BILLS')))]

    # convert data types
    new_df['Account Number'] = new_df['Account Number'].astype(int)
    new_df['Units'] = new_df['Units'].astype(str).str.replace(',', '')  # Remove commas from the column
    new_df['Units'] = pd.to_numeric(new_df['Units'], errors='coerce')
    new_df['Transaction Cash Value'] = new_df['Transaction Cash Value'].astype(str).str.replace(',','')  # Remove commas from the column
    new_df['Transaction Cash Value'] = pd.to_numeric(new_df['Transaction Cash Value'], errors='coerce')
    new_df['Transaction Cash Value 2'] = new_df['Transaction Cash Value 2'].astype(str).str.replace(',','')  # Remove commas from the column
    new_df['Transaction Cash Value 2'] = pd.to_numeric(new_df['Transaction Cash Value 2'], errors='coerce')
    new_df['Transaction Type'] = new_df['Transaction Type'].astype(str)
    new_df['Transaction Type 2'] = new_df['Transaction Type 2'].astype(str)
    new_df['Details'] = new_df['Details'].astype(str)
    new_df['Details 2'] = new_df['Details 2'].astype(str)

    # Combine Transaction Type columns and drop second column
    new_df['Transaction Type'] = new_df['Transaction Type'] + new_df['Transaction Type 2']
    new_df = new_df.drop(columns=['Transaction Type 2'])
    new_df['Transaction Type'] = new_df['Transaction Type'].str.replace('nan','')

    # Combine Details columns and drop second column
    new_df['Details'] = new_df['Details'] + new_df['Details 2']
    new_df = new_df.drop(columns=['Details 2'])
    new_df['Details'] = new_df['Details'].str.replace('nan','')

    # Combine Transaction Cash Value columns and drop second column
    new_df['Transaction Cash Value'] = new_df['Transaction Cash Value'].fillna(new_df['Transaction Cash Value 2'])
    new_df = new_df.drop(columns=['Transaction Cash Value 2'])

    # Replace Transaction Type for interest received
    new_df.loc[new_df['Details'].str.contains('CASH INTEREST ON DAILY BALANCE', na=False), 'Transaction Type'] = 'AIN'

    # Display the new DataFrame
    # print(new_df)

    # Print the data type of each column
    # print("\nColumn Data Types:")
    # print(new_df.dtypes)

    new_df.to_excel(outputPath, index=False, sheet_name='All Transactions')

