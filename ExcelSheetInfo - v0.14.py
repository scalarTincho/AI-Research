import openpyxl
import pandas as pd
from datetime import datetime

# Function to construct a dictionary mapping each cell's content in a DataFrame to its location in Excel
def construct_coord_dict_to_excel(df):
    new_dict = {}
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            new_dict[df.iat[i, j]] = to_excel_coordinates(i, j)
    return new_dict

# Function to convert 0-indexed row/col numbers to 1-indexed Excel style coordinates, e.g. "A1"
def to_excel_coordinates(row, col):
    col = col + 1  # convert 0-indexed to 1-indexed
    row = row + 1  # convert 0-indexed to 1-indexed

    col_letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_letters = chr(65 + remainder) + col_letters

    return col_letters + str(row)  # combine column letters with row number

# Function to import Excel file, process its data, and return several structured views of the data
def import_excel_file(path):
    workbook_path = path  # Path to Excel workbook to process
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)  # Load the workbook

    # Initialize dictionaries to hold various views of the data
    dfs_stripped = {}
    workbook_sheets = {}
    workbook_sheets_loc = {}
    workbook_sheets_excel = {}

    # Function to check if a string is a valid date
    def is_valid_date(date_string):
        # Date formats to check
        date_formats = ["%m/%d/%Y", "%m/%d/%y"]
        for fmt in date_formats:
            try:
                datetime.strptime(date_string, fmt)
                return True
            except ValueError:
                continue
        return False

    # Function to convert any datetimes in DataFrame to string format
    def convert_date_format(df):
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                if isinstance(df.iat[i, j], datetime):
                    df.iat[i, j] = df.iat[i, j].strftime('%m/%d/%Y')
                    return df

    # Function to map each cell's content in a DataFrame to its location
    def construct_coord_dict(df):
        new_dict = {}
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                new_dict[df.iat[i, j]] = [i, j]
        return new_dict

    # Loop over each sheet in workbook
    for sheet in workbook:
        df = pd.DataFrame(sheet.values)  # Get sheet's data as pandas DataFrame 

        workbook_sheets[sheet.title] = convert_date_format(df.copy())  # Hold a copy of original data
        workbook_sheets_loc[sheet.title] = construct_coord_dict(df.copy())  # Construct dictionary for each sheet

        for sheet_name, coord_dict in workbook_sheets_loc.items():
            # Create new dict for this sheet
            excel_dict = {}

            # Change the values in coord_dict to excel style coordinates
            for key, value in coord_dict.items():
                coord = to_excel_coordinates(*value)  # This will return the value in 'A1' style
                excel_dict[key] = f"'{sheet_name}'!{coord}"  # Now it will be in 'SheetName'!A1

            workbook_sheets_excel[sheet_name] = excel_dict  # Add the new dict to workbook_sheets_excel

        df = df.fillna("").replace(0,"")

        df = df[df.astype(bool).any(axis=1)]
        df = df.loc[:, df.astype(bool).any(axis=0)]

        dfs_stripped[sheet.title] = df

        df = df.map(lambda x: "" if isinstance(x, (int, float)) and not is_valid_date(str(x)) else x)
        dfs_stripped[sheet.title] = df

    dfs_stripped_list = {}  # will hold lists of non-empty values for each sheet

    for df_name, df in dfs_stripped.items():
        non_empty_values = df[df != ''].values.flatten()
        non_empty_values = non_empty_values[non_empty_values != '']
        non_empty_values = non_empty_values[~pd.isnull(non_empty_values)]

        # if value is datetime, convert to string format
        formatted_values = [value.strftime('%m/%d/%Y') if isinstance(value, datetime) else value for value in non_empty_values]
        dfs_stripped_list[df_name] = formatted_values 

    return workbook_sheets, workbook_sheets_loc, workbook_sheets_excel, dfs_stripped, dfs_stripped_list

# Program starts here
if __name__ == "__main__":
    # Path for test Excel file
    workbook_path = r"C:/Users/Martin56/Dropbox (Scalar Analytics)/Valuation/Sol-ti/Class B Valuation 2023.07/Company Docs/Cap Table/sol-ti-inc_2023-06-30_summary_cap_detailed_cap_ledgers.xlsx"

    workbook_sheets, workbook_sheets_loc, workbook_sheets_excel, dfs_stripped, dfs_stripped_list = import_excel_file(workbook_path)