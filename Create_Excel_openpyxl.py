# https://openpyxl.readthedocs.io/en/stable/tutorial.html

from openpyxl import Workbook, load_workbook
from datetime import datetime

month = datetime.today().month
year = datetime.today().year
dt = str(datetime.today()).split()[0]

wb = Workbook()
wb = load_workbook('Efforts_3_2021.xlsx')
names = wb.sheetnames
print(names)
if f'data_{dt}.xlsx' in names:
    print("WorkSheet already present")

else:
    wb.create_sheet(f'data_{dt}.xlsx')
wb.save('Efforts_3_2021.xlsx')
