# -*- coding: utf-8 -*-
"""
Created on Tue Sep  5 12:16:32 2023

@author: Martin56
"""
import openpyxl
import pandas as pd
from datetime import datetime

def import_excel_file(path):
    
    # Replace with the path to your Excel workbook
    workbook_path = path
    
    # Load the workbook
    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
    
    #Create a variable to store the sheet names
    # sheet_names = []
    
     #Create a dictionary to store the dataframes
    # dfs = {}
    dfs_stripped = {}
    workbook_sheets = {} #store the workbook as it is
    
    # Function to check if a string represents a valid date
    def is_valid_date(date_string):
        date_formats = ["%m/%d/%Y", "%m/%d/%y"]  # Add or change formats as needed
        for fmt in date_formats:
            try:
                datetime.strptime(date_string, fmt)
                return True
            except ValueError:
                continue
        return False
    



    # #Loop through each sheet and create a dataframe with the sheet data
    for sheet in workbook:
       
        df = pd.DataFrame(sheet.values)
        
        #Create the Worksheets original book dictionary
        workbook_sheets[sheet.title] = df.copy()

        
        #Replace empty values with zero
        df = df.fillna("").replace(0,"")
          
        # Drop rows where all columns are empty
        df = df[df.astype(bool).any(axis=1)]  
        df = df.loc[:, df.astype(bool).any(axis=0)] 
        
        #create the stripped dictionary
        dfs_stripped[sheet.title] = df
        
        # Replace numeric non-date cells    
        df = df.map(lambda x: "" if isinstance(x, (int, float)) and not is_valid_date(str(x)) else x)
        dfs_stripped[sheet.title] = df
    
        
    dfs_stripped_list = {}  # initialize an empty dictionary

    # iterate over dataframes in dfs and construct non_empty_dfs
    for df_name, df in dfs_stripped.items():
        non_empty_values = df[df != ''].values.flatten()   # get non-empty values as a flattened array
        non_empty_values = non_empty_values[non_empty_values != '']  # filter out empty values
        non_empty_values = non_empty_values[~pd.isnull(non_empty_values)]  # filter out 'nan' values
    
        # convert datetime objects to string format
        formatted_values = [value.strftime('%m/%d/%Y') if isinstance(value, datetime) else value for value in non_empty_values]
        dfs_stripped_list[df_name] = formatted_values  # add to dictionary
        
    return workbook_sheets, dfs_stripped, dfs_stripped_list
    

if __name__ == "__main__":
    
    workbook_path = r"C:\Users\Martin56\Dropbox (Scalar Analytics)\Valuation\Powerlytics, Inc(p)\IRC 409A 2023.01\Company Docs\Financials\Powerlytics - BOD Package - 1.31.2023.xlsx"
    
    #Return dictionaries
    workbook_sheets, dfs_stripped, dfs_stripped_list = import_excel_file(workbook_path)
    
    
    '''
    You are a valuation analyst fetching for the correct and relevant documents for doing a business valuation as of 01/31/202.

I'm going to pass some an exel filename in <> and its sheet names in [] and separated by coma. I need you to tell me what info does this document is likely to provide provide regarding the following categories:
- Income statement as of 01/31/2021 (containing the last twelve months or twelve training months) 
- Balance sheet as of 01/31/2021
- Historical balance sheets from 20XX to 2020
- Historical Income Statement 20XX to 2020
- Projections (mention projection period contained)

Give the answer as "JSON"  and match the filename with one of the categories above in the format (Category: => Sheet Name). 
    
    '''