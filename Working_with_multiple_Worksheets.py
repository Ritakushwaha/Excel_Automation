# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_sql_query.html#pandas.read_sql_query
# https://docs.microsoft.com/en-us/sql/machine-learning/data-exploration/python-dataframe-pandas?view=sql-server-ver15
# creator : RITA

# import all required packages
from mysql.connector import Error
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
import pandas as pd
import os
from DB_Connection import MyDB

mydb = MyDB()

_file_name = 'Multiple_worksheets.xlsx'
_sheet_name1 = 'Customer_1'
_sheet_name2 = 'Product_2'
_row = 0


def read_customer_data():
    try:
        conn = mydb.get_connection()
        sql_select_Query = "select * from Customer"
        cursor = conn.cursor()
        cursor.execute(sql_select_Query)
        records = cursor.fetchall()
        total = 0
        for row in records :
            print(row[0],row[1])
            total += 1
        print(f'Total rows in Customer Table is {total}')
    except Error as e:
        print("Error reading data from MySQL table", e)


def read_product_data():
    try:
        conn = mydb.get_connection()
        sql_select_Query = "select * from Product"
        cursor = conn.cursor()
        cursor.execute(sql_select_Query)
        records = cursor.fetchall()
        total = 0
        for row in records:
            print(row[0],row[1])
            total += 1
        print(f'Total rows in Customer Table is {total}')
    except Error as e:
        print("Error reading data from MySQL table", e)


def new_workbook(_file_name):
    print('new_workbook()')
    wb = Workbook()  # Workbook Object
    create_worksheets(wb)
    wb.save(_file_name)  # save the workbook
    wb.close()  # close the workbook


def create_worksheets(wb):
    print('create_worksheets')
    for name in wb.sheetnames:
        if name == _sheet_name1:
            print(f'{_sheet_name1} is present')
        else:
            print(f'{_sheet_name1} is not present')
            ws1 = wb.create_sheet(_sheet_name1)
            wb.save(_file_name)
        if name == _sheet_name2:
            print(f'{_sheet_name2} is present')
        else:
            print(f'{_sheet_name2} is not present')
            ws2 = wb.create_sheet(_sheet_name2)
            wb.save(_file_name)


if os.path.exists(_file_name):
    wb = load_workbook(_file_name)
    # read_customer_data()
    # read_product_data()
else:
    new_workbook(_file_name)

