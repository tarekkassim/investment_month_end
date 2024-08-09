import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def balancePvt(monthYear):

    filePath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Read the data from the current month Excel file
    df = pd.read_excel(filePath, sheet_name='Balances - Cur Month')

    pivot_table = df.pivot_table(
        index='Account Number',
        values=['Market Value', 'Accrued Interest'],
        aggfunc='sum',
        fill_value=0
    )

    # Rename the column 'Market Value' to something else, e.g., 'Total Market Value'
    pivot_table = pivot_table.rename(columns={'Market Value': 'Closing Balance', 'Accrued Interest': 'Closing Interest'})

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pivot_table.to_excel(writer, sheet_name='Balance Summary - Cur Month', index=True)
