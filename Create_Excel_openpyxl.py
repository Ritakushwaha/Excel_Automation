# https://openpyxl.readthedocs.io/en/stable/tutorial.html

from openpyxl import Workbook, load_workbook
from datetime import datetime

month = datetime.today().month
year = datetime.today().year
dt = str(datetime.today()).split()[0]


wb = Workbook()
ws = wb.active
try:
    wb = load_workbook(f'Efforts_{month}_{year}.xlsx')
except FileNotFoundError as err:
    print(f'Efforts_{month}_{year}.xlsx' + ' not found')
    new_wb = Workbook(f'Efforts_{month}_{year}.xlsx')

names = wb.sheetnames
print(names)
if f'data_{dt}.xlsx' in names:
    print("WorkSheet already present")
else:
    wb.create_sheet(f'data_{dt}.xlsx')
wb.save(f'Efforts_{month}_{year}.xlsx')
print(names)
