import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def comparison(monthYear):

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
    previous_table = pd.read_excel(preMthPath, sheet_name='Bond Summary')

    # Transactions summary Table
    transactions_df = pd.read_excel(curMthPath, sheet_name='Transactions')

    transactions = transactions_df.pivot_table(
        index='Account Number',
        columns='Transaction Type',
        values='Transaction Cash Value',
        aggfunc='sum',
        fill_value=0
    )

    # Multiply all values in the pivot table by -1
    transactions = transactions * -1

    transactions = transactions.reset_index()

    # Balance summary table

    # Read the data from the current month Excel file
    balance_df = pd.read_excel(curMthPath, sheet_name='Balances')

    # DF for bonds
    bond_df = balance_df[(balance_df['Type'] == 'CASH EQUIVALENTS') | (balance_df['Type'] == 'FIXED INCOME')]

    # Pivot for bonds
    current_table = bond_df.pivot_table(
        index='Account Number',
        values=['Market Value', 'Accrued Interest'],
        aggfunc='sum',
        fill_value=0
    )

    # Rename the column 'Market Value' to something else, e.g., 'Total Market Value'
    current_table = current_table.rename(
        columns={'Market Value': 'Closing Balance', 'Accrued Interest': 'Closing Interest'})

    # Balance Table

    # Remove the interest column
    currentBalance = current_table.drop(columns=['Closing Interest'])
    previousBalance = previous_table.drop(columns=['Closing Interest'])

    # Merge balance with transactions
    compare_balance = currentBalance.merge(transactions,
                                           left_on='Account Number',
                                           right_on='Account Number',
                                           how='left'
                                           )

    # Rename the column 'Closing Balance' to 'Opening Balance'
    previousBalance = previousBalance.rename(columns={'Closing Balance': 'Opening Balance'})

    compare_balance = compare_balance.merge(previousBalance,
                            left_on='Account Number',
                            right_on='Account Number',
                            how='left'
                            )

    # Remove INT column
    compare_balance = compare_balance.drop(columns='INT')

    # Extract the names of all columns
    all_columns = compare_balance.columns.tolist()

    # Define the new column order
    # 'Account Number' first, 'Opening Balance' second, and 'Closing Balance' last
    # Other columns in between
    new_column_order = ['Account Number', 'Opening Balance'] + [col for col in all_columns if
                                                                col not in ['Account Number', 'Opening Balance',
                                                                            'Closing Balance']] + ['Closing Balance']

    # Reorder DataFrame columns
    compare_balance = compare_balance[new_column_order]

    # Check if 'CAS' column exists and drop it if it does
    if 'CAS' in compare_balance.columns:
        compare_balance = compare_balance.drop(columns='CAS')

    # Replace NaN values
    compare_balance = compare_balance.fillna(0)

    # Identify columns to sum (excluding 'Account Number' and 'Closing Balance')
    columns_to_sum = [col for col in compare_balance.columns if col not in ['Account Number', 'Closing Balance', 'AIN']]

    # Create a new column 'Total' with the sum of the identified columns
    compare_balance['Expected Closing Balance'] = compare_balance[columns_to_sum].sum(axis=1)

    # Create a new column 'FV Change' with the difference between actual and expected closing balance
    compare_balance['FV Change'] = compare_balance['Closing Balance'] - compare_balance['Expected Closing Balance']

    # Drop Expected Closing Balance column
    compare_balance = compare_balance.drop(columns='Expected Closing Balance')

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        compare_balance.to_excel(writer, sheet_name='Bond Comparison', index=False)

    # Interest Table

    # Remove the interest column
    current_int = current_table.drop(columns=['Closing Balance'])
    previous_int = previous_table.drop(columns=['Closing Balance'])

    # Merge balance with transactions
    compare_int = current_int.merge(transactions[['Account Number', 'INT']],
                                           left_on='Account Number',
                                           right_on='Account Number',
                                           how='left'
                                           )

    # Rename the column 'Closing Balance' to 'Opening Balance'
    previous_int = previous_int.rename(columns={'Closing Interest': 'Opening Interest'})

    compare_int = compare_int.merge(previous_int,
                                            left_on='Account Number',
                                            right_on='Account Number',
                                            how='left'
                                            )

    # Replace NaN values
    compare_int = compare_int.fillna(0)

    # Reorganize columns
    compare_int = compare_int[['Account Number', 'Opening Interest', 'INT', 'Closing Interest']]

    # Create a new column with the sum of the identified columns
    compare_int['Expected Closing Interest'] = compare_int[['Opening Interest', 'INT']].sum(axis=1)

    # Interest Income
    compare_int['Interest Income'] = compare_int['Closing Interest'] - compare_int['Expected Closing Interest']

    # Drop Expected Closing Interest column
    compare_int = compare_int.drop(columns='Expected Closing Interest')

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        compare_int.to_excel(writer, sheet_name='Interest Comparison', index=False)

