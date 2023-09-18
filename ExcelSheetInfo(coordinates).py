import openpyxl
import pandas as pd
from datetime import datetime

def import_excel_file(path):
    
    workbook_path = path

    workbook = openpyxl.load_workbook(workbook_path, data_only=True)
  
    dfs_stripped = {}
    workbook_sheets = {}
    
    def is_valid_date(date_string):
        date_formats = ["%m/%d/%Y", "%m/%d/%y"]
        for fmt in date_formats:
            try:
                datetime.strptime(date_string, fmt)
                return True
            except ValueError:
                continue
        return False

    for sheet in workbook:
        df = pd.DataFrame(sheet.values)
        df = df.fillna("").replace(0,"")
        df = df[df.astype(bool).any(axis=1)]
        df = df.loc[:, df.astype(bool).any(axis=0)]
 
        workbook_sheets[sheet.title] = df
        dfs_stripped[sheet.title] = df

        df = df.map(lambda x: "" if isinstance(x, (int, float)) and not is_valid_date(str(x)) else x)
        dfs_stripped[sheet.title] = df
    
    dfs_stripped_coordinates = {}

    for df_name, df in dfs_stripped.items():

        non_empty_mask = df.map(lambda x: pd.notnull(x) and x != '')
        non_empty_indices = non_empty_mask.stack()[lambda x: x].index.tolist()

        non_empty_values = df[df != ''].values.flatten()
        non_empty_values = non_empty_values[non_empty_values != '']
        non_empty_values = non_empty_values[~pd.isnull(non_empty_values)]
    
        formatted_values = [value.strftime('%m/%d/%Y') if isinstance(value, datetime) else value for value in non_empty_values]
        stripped_coordinate_list = list(zip(formatted_values, non_empty_indices))
        
        df_stripped_coordinate =  pd.DataFrame(stripped_coordinate_list, columns=['Value', 'Coordinate'])
        dfs_stripped_coordinates[df_name] = df_stripped_coordinate
        
    return workbook_sheets, dfs_stripped, dfs_stripped_coordinates
    

if __name__ == "__main__":
    
    workbook_path = r"C:\Users\Martin56\Dropbox (Scalar Analytics)\Valuation\Test Axis AI - Copy\409A 2021.03\Company Docs\Financials\Light Bio, Inc. P&L (12 months).xlsx"
    workbook_sheets, dfs_stripped, dfs_stripped_coordinates = import_excel_file(workbook_path)
    
    for key, value in dfs_stripped_coordinates.items():
        dfs_stripped_coordinates[key] = [(item["Value"].strip() if isinstance(item["Value"], str) else item["Value"], item["Coordinate"]) for index, item in value.iterrows()]
        
    NEW_DFs = {}
    
    NEW_DFs = {}
    
    for df_name, list_of_tuples in dfs_stripped_coordinates.items():
        new_df = pd.DataFrame(dtype=object)  # Initialize an empty DataFrame
        for value, (x, y) in list_of_tuples:
            new_df.loc[x, y] = str(value)  # Use loc instead of at
        NEW_DFs[df_name] = new_df
        
        
        

'''
Based on the given data structure, the calculations can be deduced into the listed structure as follows:

1. Total Revenue: This is represented as 'Total Income', (6, 0)

2. COGS: The data does not provide direct information to calculate Cost of Goods Sold.

3. Gross Profit: This is directly provided as 'Gross Profit', (7, 0)

4. OPExp (Operating Expense): This is given as 'Total Expenses', (17, 0)

5. EBITDA: There's no direct data for EBITDA calculation. However, it is typically calculated as: Gross Profit - OpEx + Depreciation + Amortization

6. Dep Exp (Depreciation Expense): It is represented as '02-505 Depreciation expense — Patents', (20, 0)

7. Amort Exp (Amortization Expense): The data does not provide direct information to calculate Amortization Expense.

8. EBIT (Earnings Before Interest and Taxes): This can be calculated as 'Net Operating Income' or Gross profit minus OpEx

9. Interest Exp (Interest Expense): This is given as '01-533 Interest Paid', (9, 0)

10. Other Exp (Other Expenses): Shown as 'Total Other Expenses',(21, 0)

11. EBT (Earnings Before Taxes): Typically calculated as EBIT - Interest Exp.

12. Taxes: There is no direct data provided for Taxes calculation.

13. Net Income: It is presented directly as 'Net Income', (23, 0) 


'''