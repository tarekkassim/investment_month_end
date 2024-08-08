import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def transactionPvt(monthYear):

    filePath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Read the data from the current month Excel file
    df = pd.read_excel(filePath, sheet_name='All Transactions')

    pivot_table = df.pivot_table(
        index='Account Number',
        columns='Transaction Type',
        values='Transaction Cash Value',
        aggfunc='sum',
        fill_value=0
    )

    # Multiply all values in the pivot table by -1
    pivot_table = pivot_table * -1

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pivot_table.to_excel(writer, sheet_name='Transaction Summary', index=True)
