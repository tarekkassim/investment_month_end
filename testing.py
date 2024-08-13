import pandas as pd

somelist = pd.DataFrame({
    'Account Number': [1001, 1002, 1003, 1004, 1001],
    'Type': ['Debit', 'Credit', 'Debit', 'Credit', 'Debit'],
    'Amount': [150, 200, 50, 400, 150],
})

somepvt = somelist.pivot_table(
    index='Type',
    values='Amount',
    aggfunc='sum'
)

someindex = somepvt.index

someindex = pd.DataFrame(someindex)
print(someindex)