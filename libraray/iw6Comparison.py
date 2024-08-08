import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta



def comparison(monthYear):

    #Function to create previous month
    def get_previous_month_abbr(date_str):
        # Parse the input string to a datetime object
        current_date = datetime.strptime(date_str, '%B %Y')

        # Calculate the first day of the previous month
        next_month = current_date.replace(day=1) - timedelta(days=1)  # this will definitely be in the previous month
        next_month = next_month.replace(day=1)  # go to the first day of the next month

        # Format the next month in the desired format and extract the first three letters of the month
        next_month_str = next_month.strftime('%b %Y')

        return next_month_str


    previousMonth = get_previous_month_abbr(monthYear)

    curMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Create tables for merge
    currentBalance = pd.read_excel(curMthPath, sheet_name='Balance Summary - Current Month')
    previousBalance = pd.read_excel(curMthPath, sheet_name='Balance Summary - Prev Month')
    transactions = pd.read_excel(curMthPath, sheet_name='Transaction Summary')

    compareTable = currentBalance.merge(transactions,
                            left_on='Account Number',
                            right_on='Account Number',
                            how='left'
                            )

    # Rename the column 'Closing Balance' to 'Opening Balance'
    previousBalance = previousBalance.rename(columns={'Closing Balance': 'Opening Balance'})

    compareTable = compareTable.merge(previousBalance,
                            left_on='Account Number',
                            right_on='Account Number',
                            how='left'
                            )
    #Reorder columns
    compareTable = compareTable[['Account Number', 'Opening Balance', 'RVP', 'DVP', 'INT', 'MSC', 'OIN', 'CAS','Closing Balance']]

    compareTable = compareTable.fillna(0)

    compareTable['Expected Closing Balance'] = (compareTable['Opening Balance'] +
                                                compareTable['CAS'] +
                                                compareTable['DVP'] +
                                                compareTable['MSC'] +
                                                compareTable['OIN'] +
                                                compareTable['RVP'])

    compareTable['FV Change'] = compareTable['Closing Balance'] - compareTable['Expected Closing Balance']

    compareTable = compareTable[['Account Number',
                                 'Opening Balance',
                                 'RVP', 'DVP', 'MSC', 'OIN', 'CAS',
                                 'Expected Closing Balance',
                                 'Closing Balance',
                                 'FV Change',
                                 'INT']]


    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        compareTable.to_excel(writer, sheet_name='Comparison', index=False)
