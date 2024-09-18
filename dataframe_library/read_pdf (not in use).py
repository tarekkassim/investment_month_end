import pdfplumber
import pandas as pd


def extract_cash(monthYear):
    # Specify the PDF file path
    pdf_file = (
            r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Source Reports\305 - Cash Balance Summary Report - ' + monthYear + '.pdf'
    )
    outputPath = r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + monthYear + '.xlsx'

    # List of prefixes to filter
    prefixes = ["147122001", "147122005", "147122006", "147122007", "147122009", "147122010", "147122011", "147122010",
                "147122012", "147122013", "000147000CAD"]

    # Initialize an empty list to store the lines
    filtered_lines = []

    # Open the PDF file
    with pdfplumber.open(pdf_file) as pdf:
        # Loop through the pages
        for page_number, page in enumerate(pdf.pages):
            # Extract text
            text = page.extract_text()

            # Split the text into lines
            lines = text.split("\n")

            # Filter lines that start with any of the specified prefixes
            for line in lines:
                if any(line.startswith(prefix) for prefix in prefixes):
                    filtered_lines.append(line)

    # Create a DataFrame from the filtered lines
    df = pd.DataFrame(filtered_lines, columns=["Line"])

    # Extract the account and amount from each line
    df['Account Number'] = df['Line'].apply(lambda x: x.split()[0])
    df['Closing Balance'] = df['Line'].apply(lambda x: x.split()[-1])

    # Drop line column
    df = df.drop(columns=['Line'])

    # Remove CAD from Account Number column
    df['Account Number'] = df['Account Number'].str.replace('CAD', '', regex=False)

    # Convert column type
    df['Account Number'] = df['Account Number'].astype(int)
    df['Closing Balance'] = df['Closing Balance'].astype(str).str.replace(',', '')  # Remove commas from the column
    df['Closing Balance'] = pd.to_numeric(df['Closing Balance'], errors='coerce')

    # Export new_df to the same workbook as a new sheet
    with pd.ExcelWriter(outputPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Cash Balances', index=False)
