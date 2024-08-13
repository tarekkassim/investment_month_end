import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta


def balancePvt(monthYear):

    filePath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # Read the data from the current month Excel file
    df = pd.read_excel(filePath, sheet_name='Balances')

    # # DF for bonds
    # bond_df = df[(df['Type'] == 'CASH EQUIVALENTS') | (df['Type'] == 'FIXED INCOME')]
    #
    # # Pivot for bonds
    # bond_pivot_table = bond_df.pivot_table(
    #     index='Account Number',
    #     values=['Market Value', 'Accrued Interest'],
    #     aggfunc='sum',
    #     fill_value=0
    # )
    #
    # # Rename the column 'Market Value' to something else, e.g., 'Total Market Value'
    # bond_pivot_table = bond_pivot_table.rename(columns={'Market Value': 'Closing Balance', 'Accrued Interest': 'Closing Interest'})
    #
    # # Export new_df to the same workbook as a new sheet
    # with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    #     bond_pivot_table.to_excel(writer, sheet_name='Bond Summary', index=True)

    # DF for equity
    equity_df = df[(df['Type'] == 'FUNDS') & (df['Account Number'].isin([147122005, 147122009]))]

    # Pivot for equity
    equity_pivot_table = equity_df.pivot_table(
        index='Account Number',
        values=['Market Value'],
        aggfunc='sum',
        fill_value=0
    )

    # Rename the column 'Market Value' to something else, e.g., 'Total Market Value'
    equity_pivot_table = equity_pivot_table.rename(
        columns={'Market Value': 'Closing Balance'}
    )

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        equity_pivot_table.to_excel(writer, sheet_name='Equity Summary', index=True)

