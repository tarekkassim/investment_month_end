import pandas as pd

# Sample DataFrame
data = {
    'Account Number': [101, 102, 103],
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Opening Balance': [500, 300, 400],
    'Transaction Amount': [50, -20, 30],
    'Closing Balance': [450, 280, 430]
}
df = pd.DataFrame(data)

# Extract the names of all columns
all_columns = df.columns.tolist()

# Define the new column order
# 'Account Number' first, 'Opening Balance' second, and 'Closing Balance' last
# Other columns in between
new_column_order = ['Account Number', 'Opening Balance'] + [col for col in all_columns if col not in ['Account Number', 'Opening Balance', 'Closing Balance']] + ['Closing Balance']

# Reorder DataFrame columns
df = df[new_column_order]

print(df.to_string())