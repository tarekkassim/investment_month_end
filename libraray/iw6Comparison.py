import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def comparison(monthYear):

    curMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Create tables for merge
    current_table = pd.read_excel(curMthPath, sheet_name='Balance Summary - Cur Month')
    previous_table = pd.read_excel(curMthPath, sheet_name='Balance Summary - Prev Month')
    transactions = pd.read_excel(curMthPath, sheet_name='Transaction Summary')

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
    columns_to_sum = [col for col in compare_balance.columns if col not in ['Account Number', 'Closing Balance']]

    # Create a new column 'Total' with the sum of the identified columns
    compare_balance['Expected Closing Balance'] = compare_balance[columns_to_sum].sum(axis=1)

    # Create a new column 'FV Change' with the difference between actual and expected closing balance
    compare_balance['FV Change'] = compare_balance['Closing Balance'] - compare_balance['Expected Closing Balance']

    # Drop Expected Closing Balance column
    compare_balance = compare_balance.drop(columns='Expected Closing Balance')

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        compare_balance.to_excel(writer, sheet_name='Balance Comparison', index=False)

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

