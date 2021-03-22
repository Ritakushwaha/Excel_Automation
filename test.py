from openpyxl import Workbook, load_workbook
import pandas as pd
from pandas import ExcelWriter

_file_name = 'Test.xlsx'

wb = load_workbook(_file_name)
ws = wb.active
prod_data = {'Name' : ['Rita', 'Girwar'], 'Age' : [23, 23]}
prod_df = pd.DataFrame(prod_data)
print(prod_df)
with ExcelWriter(_file_name) as writer :
    prod_df.to_excel(writer, sheet_name='Love')
    prod_df.to_excel(writer, sheet_name='Hate')
