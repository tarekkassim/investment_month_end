import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def interest_details(monthYear):

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
        new_current_balance.rename(columns={'Accrued Interest': 'Closing Interest Balance'}, inplace=True)
        new_previous_balance.rename(columns={'Accrued Interest': 'Opening Interest Balance'}, inplace=True)

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

        transactions_accounts = transactions_accounts[['CUSIP', 'INT']]

        # Concatenate the 'Account Number' columns from all DataFrames
        details = pd.concat([
            new_previous_balance['CUSIP'], new_current_balance['CUSIP'], transactions_accounts['CUSIP']
        ])

        # Get unique values
        details = details.drop_duplicates().reset_index(drop=True)

        # Create a new DataFrame with the unique 'Account Number'
        details = pd.DataFrame({'CUSIP': details})

        # Segregate bonds and bills
        details['Type'] = details['CUSIP'].apply(lambda x: 'Bills' if 'CA1350Z7' in x else 'Bonds')

        # Lookup values from tables

        details = details.merge(new_previous_balance[['CUSIP', 'Opening Interest Balance']],
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left'
                                )

        details = details.merge(transactions_accounts,
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left',
                                )

        details = details.merge(new_current_balance[['CUSIP', 'Closing Interest Balance']],
                                left_on='CUSIP',
                                right_on='CUSIP',
                                how='left'
                                )

        details = details.fillna(0)

        # Identify columns to sum (excluding 'Account Number' and 'Closing Balance')
        columns_to_sum = [col for col in details.columns if
                          col not in ['CUSIP', 'Type', 'Closing Interest Balance']]

        # Create a new column 'Total' with the sum of the identified columns
        details['Expected Closing Balance'] = details[columns_to_sum].sum(axis=1)

        # Create a new column 'FV Change' with the difference between actual and expected closing balance
        details['Interest Income'] = details['Closing Interest Balance'] - details['Expected Closing Balance']

        # Drop Expected Closing Balance column
        details = details.drop(columns='Expected Closing Balance')

        # Create new interest columns
        details['Interest Income - Bonds'] = details.apply(
            lambda row: row['Interest Income'] if row['Type'] == 'Bonds' else 0, axis=1)
        details['Interest Income - Bills'] = details.apply(
            lambda row: row['Interest Income'] if row['Type'] == 'Bills' else 0, axis=1)

        # Export new_df to the same workbook as a new sheet
        sheet_name = f'Int Details {accounts_name}'
        with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            details.to_excel(writer, sheet_name=sheet_name, index=False)

