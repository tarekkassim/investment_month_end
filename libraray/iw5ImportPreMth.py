import pandas as pd
from datetime import datetime, timedelta

monthYear = 'June 2024'

def importPreMonth(monthYear):

    def get_previous_month_abbr(date_str):
        # Parse the input string to a datetime object
        current_date = datetime.strptime(date_str, '%B %Y')

        # Calculate the first day of the previous month
        previous_month = current_date.replace(day=1) - timedelta(days=1)
        previous_month = previous_month.replace(day=1)

        # Format the previous month in the desired format and extract the first three letters of the month
        previous_month_str = previous_month.strftime('%b %Y')

        return previous_month_str


    previousMonth = get_previous_month_abbr(monthYear)

    preMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + previousMonth + '.xlsx'
    curMthPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    df = pd.read_excel(preMthPath, sheet_name='Balances - Current Month')

    # Export df to the current month's workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Balances - Previous Month', index=False)

    df = pd.read_excel(preMthPath, sheet_name='Balance Summary - Current Month')

    # Export df to the current month's workbook as a new sheet
    with pd.ExcelWriter(curMthPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Balance Summary - Prev Month', index=False)