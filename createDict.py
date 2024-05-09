import openpyxl

def create_dictionary(file_path, sheet_name, column_index):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]

    # Create an empty dictionary
    dictionary = {}

    # Iterate through the rows in the specified column
    for row in sheet.iter_rows(min_row=2, min_col=column_index, values_only=True):
        value = row[0]
        row_number = row[0].row
        dictionary[value] = row_number

    return dictionary

# Example usage
file_path = '/path/to/your/excel/file.xlsx'
sheet_name = 'Sheet1'
column_index = 1

result = create_dictionary(file_path, sheet_name, column_index)