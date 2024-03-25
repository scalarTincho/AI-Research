import openpyxl
import pandas as pd
from datetime import datetime


def construct_coord_dict_to_excel(df):
    new_dict = {}
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            new_dict[df.iat[i, j]] = to_excel_coordinates(i, j)
    return new_dict

def to_excel_coordinates(row, col):
    col = col + 1  # convert 0-indexed to 1-indexed
    row = row + 1  # convert 0-indexed to 1-indexed

    # convert column number to column letter
    col_letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_letters = chr(65 + remainder) + col_letters

    return col_letters + str(row)


def import_excel_file(path):
    
    # Replace with the path to your Excel workbook
    workbook_path = path
    
    # Load the workbook
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    
    dfs_stripped = {}
    workbook_sheets = {} 
    workbook_sheets_loc = {} 
    workbook_sheets_excel = {}
    
    
    def is_valid_date(date_string):
        date_formats = ["%m/%d/%Y", "%m/%d/%y"]  
        for fmt in date_formats:
            try:
                datetime.strptime(date_string, fmt)
                return True
            except ValueError:
                continue
        return False
    
    def convert_date_format(df):
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                if isinstance(df.iat[i, j], datetime):
                    df.iat[i, j] = df.iat[i, j].strftime('%m/%d/%Y')
                    return df

    def construct_coord_dict(df):
        new_dict = {}
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                new_dict[df.iat[i, j]] = [i, j]
        return new_dict


    for sheet in workbook:
       
        df = pd.DataFrame(sheet.values)
        

        workbook_sheets[sheet.title] = convert_date_format(df.copy())
        workbook_sheets_loc[sheet.title] = construct_coord_dict(df.copy()) 

        # Iterate over each sheet in workbook_sheets_loc
        for sheet_name, coord_dict in workbook_sheets_loc.items():
            # Create new dict for this sheet
            excel_dict = {}

            # Change the values in coord_dict to excel style coordinates
            for key, value in coord_dict.items():
                coord = to_excel_coordinates(*value)  # This will return the value in 'A1' style
                excel_dict[key] = f"'{sheet_name}'!{coord}"  # Now it will be in 'SheetName'!A1

            # Add the new dict to workbook_sheets_excel
            workbook_sheets_excel[sheet_name] = excel_dict

            
        df = df.fillna("").replace(0,"")
          
        df = df[df.astype(bool).any(axis=1)]  
        df = df.loc[:, df.astype(bool).any(axis=0)] 
        
        dfs_stripped[sheet.title] = df
        
          
        df = df.map(lambda x: "" if isinstance(x, (int, float)) and not is_valid_date(str(x)) else x)
        dfs_stripped[sheet.title] = df
    
        
    dfs_stripped_list = {}  # initialize an empty dictionary

    for df_name, df in dfs_stripped.items():
        non_empty_values = df[df != ''].values.flatten()  
        non_empty_values = non_empty_values[non_empty_values != '']  
        non_empty_values = non_empty_values[~pd.isnull(non_empty_values)]  

    
        formatted_values = [value.strftime('%m/%d/%Y') if isinstance(value, datetime) else value for value in non_empty_values]
        dfs_stripped_list[df_name] = formatted_values  

        
        return workbook_sheets, workbook_sheets_loc, workbook_sheets_excel, dfs_stripped, dfs_stripped_list
    

#Program Starts here
if __name__ == "__main__":
    
    workbook_path = r"C:\Users\Martin56\Dropbox (Scalar Analytics)\Valuation\Powerlytics, Inc(p)\IRC 409A 2023.01\Company Docs\Financials\Powerlytics - BOD Package - 1.31.2023.xlsx"
    

    workbook_sheets, workbook_sheets_loc, workbook_sheets_excel, dfs_stripped, dfs_stripped_list = import_excel_file(workbook_path)