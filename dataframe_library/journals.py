import pandas as pd


def journals(monthYear):
    # Import tables
    file_path = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'
    mapping_path = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Source Reports\Mapping.xlsx'
    bonds = pd.read_excel(file_path, sheet_name='Bond Comparison')
    interest = pd.read_excel(file_path, sheet_name='Interest Comparison')
    equity = pd.read_excel(file_path, sheet_name='Equity Comparison')
    cash = pd.read_excel(file_path, sheet_name='Cash Recon')
    mapping = pd.read_excel(mapping_path, sheet_name='Mapping')

    # Bond transactions
    bond_transaction_columns = [
        col for col in bonds.columns if col not in ['Account Number', 'Opening Balance', 'Closing Balance']
    ]

    bond_journals = pd.melt(
        bonds,
        id_vars=['Account Number'],
        value_vars=bond_transaction_columns,
        var_name='Type',
        value_name='Amount'
    )

    bond_journals['Source'] = 'Bonds'

    # Interest transactions

    interest_journals = pd.melt(
        interest,
        id_vars=['Account Number'],
        value_vars=['INT','Interest Income - Bonds', 'Interest Income - Bills'],
        var_name='Type',
        value_name='Amount'
    )

    interest_journals['Source'] = 'Int Payments'

    # Equity transactions
    equity_transaction_columns = [
        col for col in equity.columns if col not in ['Account Number', 'Opening Balance', 'Closing Balance']
    ]

    equity_journals = pd.melt(
        equity,
        id_vars=['Account Number'],
        value_vars=equity_transaction_columns,
        var_name='Type',
        value_name='Amount'
    )

    equity_journals['Source'] = 'Equity'

    # Cash interest transactions
    cash_interest_journals = pd.melt(
        cash,
        id_vars=['Account Number'],
        value_vars=['AIN'],
        var_name='Type',
        value_name='Amount'
    )

    cash_interest_journals['Amount'] = cash_interest_journals['Amount'] * -1

    cash_interest_journals['Source'] = 'Acc Interest'

    # Combine Journals
    all_journals = pd.concat([bond_journals, interest_journals, equity_journals, cash_interest_journals],
                             ignore_index=True)

    # Remove columns with 0 value
    all_journals = all_journals[all_journals['Amount'] != 0]
    all_journals = all_journals.dropna()

    # Merge Mapping
    all_journals = pd.merge(
        all_journals,
        mapping,
        on=['Account Number', 'Source', 'Type']
    )

    # Melt for Accounts
    all_journals = pd.melt(
        all_journals,
        id_vars=['Account Number', 'Source', 'Type', 'Amount'],
        value_vars=['Account 1', 'Account 2'],
        var_name='Account Seq',
        value_name='Account'
    )

    # Reorder Columns
    all_journals = all_journals[['Account Number', 'Source', 'Type', 'Account Seq', 'Account', 'Amount']]

    # Fix signage
    all_journals.loc[all_journals['Account Seq'] == 'Account 2', 'Amount'] *= -1

    # Sort
    all_journals = all_journals.sort_values(by=['Account Number', 'Source', 'Type', 'Account Seq'], ignore_index=True)

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        all_journals.to_excel(writer, sheet_name='Journals', index=False)

