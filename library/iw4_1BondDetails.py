import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def bond_details(monthYear):

    def get_previous_month_abbr(date_str):
        # Parse the input string to a datetime object
        current_date = datetime.strptime(date_str, '%b %Y')

        # Calculate the first day of the previous month
        previous_month = current_date.replace(day=1) - timedelta(days=1)
        previous_month = previous_month.replace(day=1)

        # Format the previous month in the desired format and extract the first three letters of the month
        previous_month_str = previous_month.strftime('%b %Y')

        return previous_month_str

    previousMonth = get_previous_month_abbr(monthYear)

    curMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'
    preMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + previousMonth + '.xlsx'

    # Create tables for merge
    current_balance = pd.read_excel(curMthPath, sheet_name='Balances')
    previous_balance = pd.read_excel(preMthPath, sheet_name='Balances')
    current_transactions = pd.read_excel(curMthPath, sheet_name='Transactions')

    # Conditions to create new tables
    condition_current_balance = current_balance['CUSIP'].str.contains('CA', na=False)
    condition_previous_balance = previous_balance['CUSIP'].str.contains('CA', na=False)
    condition_transactions = current_transactions['CUSIP'].str.contains('CA', na=False)

    # Filter tables based on conditions
    current_balance = current_balance[condition_current_balance]
    previous_balance = previous_balance[condition_previous_balance]
    current_transactions = current_transactions[condition_transactions]

    # Filter for Account Number
    accounts = [147122001, 147122005]
    accounts_name = ['BG', 'PHN']

    for account, accounts_name in zip(accounts, accounts_name):

        # Filter tables based on account
        new_current_balance = current_balance[current_balance['Account Number'] == account].copy()
        new_previous_balance = previous_balance[previous_balance['Account Number'] == account].copy()
        new_current_transactions = current_transactions[current_transactions['Account Number'] == account].copy()

        # Change column names
        new_current_balance.rename(columns={'Market Value': 'Closing Balance'}, inplace=True)
        new_previous_balance.rename(columns={'Market Value': 'Opening Balance'}, inplace=True)

        # Pivot of transactions
        new_current_transactions = (
            new_current_transactions[new_current_transactions['Account Number'] == account].copy()
        )

        transactions = new_current_transactions.pivot_table(
            index='CUSIP',
            columns='Transaction Type',
            values='Transaction Cash Value',
            aggfunc='sum',
            fill_value=0
        )

        # Converting pivot to proper Dataframe
        transactions_accounts = transactions.reset_index()

        # multiply pivot columns by -1
        transactions_accounts.loc[:, transactions_accounts.columns != 'CUSIP'] *= -1

        # Concatenate the 'Account Number' columns from both DataFrames
        details = pd.concat([new_current_balance['CUSIP'], transactions_accounts['CUSIP']])

        # Get unique values
        details = details.drop_duplicates().reset_index(drop=True)

        # # Create a new DataFrame with the unique 'Account Number'
        details = pd.DataFrame({'CUSIP': details})

        # Lookup values from tables
        details = details.merge(new_previous_balance[['CUSIP','Opening Balance']],
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left'
                                )

        details = details.merge(transactions_accounts,
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left',
                                )

        if 'INT' in details.columns:
            details = details.drop(columns='INT')

        details = details.merge(new_current_balance[['CUSIP', 'Closing Balance']],
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left'
                                )

        details = details.fillna(0)

        # Identify columns to sum (excluding 'Account Number' and 'Closing Balance')
        columns_to_sum = [col for col in details.columns if
                          col not in ['CUSIP', 'Closing Balance']]

        # Create a new column 'Total' with the sum of the identified columns
        details['Expected Closing Balance'] = details[columns_to_sum].sum(axis=1)

        # Create a new column 'FV Change' with the difference between actual and expected closing balance
        details['FV Change'] = details['Closing Balance'] - details['Expected Closing Balance']

        # Drop Expected Closing Balance column
        details = details.drop(columns='Expected Closing Balance')

        # Export new_df to the same workbook as a new sheet
        sheet_name = f'Bond Details {accounts_name}'
        with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            details.to_excel(writer, sheet_name=sheet_name, index=False)

