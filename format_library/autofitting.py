from openpyxl import load_workbook


def autofit(month_year):

    # Load the workbook
    wb = load_workbook(r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + month_year + '.xlsx')

    # Function to autofit columns
    def autofit_columns(sheet):
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)  # Add a little extra space
            sheet.column_dimensions[column].width = adjusted_width

    # Iterate through all sheets and autofit columns
    for sheet in wb.worksheets:
        autofit_columns(sheet)

    # Save the updated workbook
    wb.save(r'C:\Users\ttarek\OneDrive - Tarion\Projects\Python\Investment Working - ' + month_year + '.xlsx')

