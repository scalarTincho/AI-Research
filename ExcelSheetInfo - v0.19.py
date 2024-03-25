import openpyxl
import pandas as pd
from datetime import datetime
# # import warnings


# Function to map each cell's content in a DataFrame to its location
def construct_coord_dict(df): #Batalla Naval
    new_dict = {}
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            new_dict[df.iat[i, j]] = [i, j]
    return new_dict

# Function to convert 0-indexed row/col numbers to 1-indexed Excel style coordinates, e.g. "A1"
def to_excel_coordinates(row, col): #Batalla Naval version excel
    col = col + 1  # convert 0-indexed to 1-indexed
    row = row + 1  # convert 0-indexed to 1-indexed

    col_letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_letters = chr(65 + remainder) + col_letters

    return col_letters + str(row)  # combine column letters with row number

# Function to convert any datetimes in DataFrame to string format
def convert_date_format(df): 
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if isinstance(df.iat[i, j], datetime):
                df.iat[i, j] = df.iat[i, j].strftime('%m/%d/%Y')
    return df

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

# Function to import Excel file, process its data, and return several structured views of the data
def import_excel_file(path):
    
    workbook_path = path  # Path to Excel workbook to process
    
    #Check if file path ends with .xls:
    if path.lower().endswith('.xlsx'):
        workbook = openpyxl.load_workbook(workbook_path, data_only=True)  # Load the workbook        
        is_xlsx = True
        
    elif path.lower().endswith('.xls'):
        xls = pd.ExcelFile(path)
        workbook = xls.sheet_names
        is_xlsx = False
    

# Initialize dictionaries to hold various views of the data
    workbook_sheets = {} #The original workbook exactly like excel
    workbook_sheets_loc = {} #position of every item in [x,y] in pandas coordinates
    
    #Dictionary of sheets (as lists) ready to pass to AI (Company name removed)
    workbook_sheets_excel = {}  
    dfs_stripped = {} #Only rows and columns string data    

    # Loop over each sheet in workbook
    for sheet in workbook:
        
        #xlsx
        if is_xlsx == True:
            df = pd.DataFrame(sheet.values)  # Get sheet's data as pandas DataFrame 
            sheet_title = sheet.title
        
        #xls
        elif is_xlsx == False:
            df = pd.read_excel(xls, sheet_name=sheet)
            sheet_title = sheet
        
        '''
        The original dictionary  that holds the sheet as it is
        '''
        # Hold a copy of original data (No privacy filters)
        workbook_sheets[sheet_title] = convert_date_format(df.copy())
        
        
        # Use applymap to apply the replacement function to each element of the dataframe
        df = df.map(lambda x: x.replace(company_name, 'SUPER SECRET COMPANY NAME') if isinstance(x, str) else x)
                
       # Construct dictionary for each sheet
        workbook_sheets_loc[sheet_title] = construct_coord_dict(df)  

        for sheet_name, coord_dict in workbook_sheets_loc.items():
            # Create new dict for this sheet
            excel_dict = {}

            # Change the values in coord_dict to excel style coordinates
            for key, value in coord_dict.items():
                coord = to_excel_coordinates(*value)  # This will return the value in 'A1' style
                excel_dict[key] = f"'{sheet_name}'!{coord}"  # Now it will be in 'SheetName'!A1

            workbook_sheets_excel[sheet_name] = excel_dict  # Add the new dict to workbook_sheets_excel

        df = df.fillna("").replace(0,"")
w
        df = df[df.astype(bool).any(axis=1)]
        df = df.loc[:, df.astype(bool).any(axis=0)]
        
        
        dfs_stripped[sheet_title] = df

        df = df.map(lambda x: "" if isinstance(x, (int, float)) and not is_valid_date(str(x)) else x)
        dfs_stripped[sheet_title] = convert_date_format(df) #Convert Timestamps

    dfs_stripped_list = {}  # will hold lists of non-empty values for each sheet

    for df_name, df in dfs_stripped.items():
        non_empty_values = df[df != ''].values.flatten()
        non_empty_values = non_empty_values[non_empty_values != '']
        non_empty_values = non_empty_values[~pd.isnull(non_empty_values)]

        # if value is datetime, convert to string format
        formatted_values = [value.strftime('%m/%d/%Y') if isinstance(value, datetime) else value for value in non_empty_values]
        dfs_stripped_list[df_name] = formatted_values 

    return workbook_sheets, workbook_sheets_excel, dfs_stripped, dfs_stripped_list


# Program starts here
if __name__ == "__main__":
    # Path for test Excel file
    workbook_path = r"C:\Users\Martin56\Downloads\Test Cap. Table\Test3.xlsx"

    company_name = "Azova"    

    workbook_sheets, workbook_sheets_excel, dfs_stripped, dfs_stripped_list = import_excel_file(workbook_path)
    
    
    # # print the new workbook_sheets_excel dictionary
    print(workbook_sheets_excel.keys())
    
    store_info = ""
    for key, value in workbook_sheets_excel.items():
        store_info += f"{key} - {value}\n"
  
    
    '''
    PROMPT
So, now I need to to play the part of the best valuation analyst ever.

I'm presenting a Balance Sheet extracted from an excel file in the following format:
Value (the actual value of the cell) - Coordinates (the cell reference, that in excel would represent Row/Column)

I need you to create a condense balance sheet populating the following table just for December 2022. Present de information as a table next to each items:

Cash
Accounts Receivable
Inventory
Other Current Assets
Total Current Assets

Property plants and equipment
Goodwill & Other Indefinite-Lived Intangibles
Intangibles
Other Long Term Assets
Total Lont term Assets
Total Assets

Accounts Payable
nAccured Liabilites
Short Term Debt
Defeferred Revenue
Other Current Liabilites
Total Current Liabilities
Ling Term Debt
Other Long Term Liabilities
Total Long Term Liabilities
Total Liabilities
Equity
Total Liab&Equity


If some calculations need to be done, don't present the result bue the calculations being made using this as a guidance:
* Current assets = '2022 BS'!D9 + '2022 BS'!D10 ==>This is just an example
* Calcualtion's column value need to start with an = sign
* This need to be representes asi BS'!AB22AB24+BS'!AB22AB25+BS'!AB22AB26
* Total Liab&Equity shoulb be equal to Total Assets
* CAlculations column should match value column
    


----------------------
Dear AI. You are a valuation analyst preparing a cap. table for a company. It's name will be super secret.

The values are presented in the following structure from a list object:
Key: Value  that contains the actual value of the cell 
Value: Coordinates (the cell reference, that in excel would represent Row/Column)

Using the information below from an excel I need you to populate a table with the following structure. Please create a table.
Make this security by security and not investor by investor:

Security Name | Shares Issued and Outstanding | Issue Price/Strike Price/Price per share | Shares Issued and Outstanding Cell Reference 

Clarifications:
* Cell reference should refer to shares issued and outstanding
* In the case of options please condense them based on their issue price and oustanding number.
* "formats the Cell Reference to be copied and pasted into excel"
* "issued and outstanding shares" is not the same as "fully diluted shares"



'''