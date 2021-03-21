# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_sql_query.html#pandas.read_sql_query

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
import pandas as pd
import os


# month, year, date & datetime format
month = datetime.today().month
year = datetime.today().year
_date = str(datetime.today()).split()[0]
# _minus_days = int(input("Enter days to minus from current date"))

dt = str((datetime.now()-timedelta(days = 20)).strftime("%Y-%m-%d %H:%M:%S"))


# Excel Workbook, worksheet names
_file_name = f'Efforts_{month}_{year}.xlsx'
_sheet_name = f'Data_{_date}'
_ticket_file = 'Data.xlsx'


def new_workbook(_file_name):
    print("new_workbook() : called")
    wb = Workbook()
    ws = wb.active
    ws.title = _sheet_name
    ws['A1'] = 'Ticket No'
    ws['B1'] = 'Ref No'
    ws['C1'] = 'Corp ID'
    ws['D1'] = 'Activity Type'
    ws['E1'] = 'Activity'
    ws['F1'] = 'Timecard Entry Date'
    ws['G1'] = 'Effort'
    ws['H1'] = 'Complexity'
    ws['I1'] = 'AMorAD'
    ws['J1'] = 'SOW'
    ws['K1'] = 'Project'

    col_range = ws.max_column
    for col in range(1, col_range + 1) :
        cell_header = ws.cell(1, col)
        cell_header.fill = PatternFill(start_color='e6f738', end_color='e6f738', fill_type="solid")

    wb.save(_file_name)


def read_data_excel(_ticket_file) :
    print("read_data_excel() : called")
    tickets = pd.read_excel(_ticket_file, index_col=0,usecols='A')
    return tickets


def create_list() :
    print("create_list() : called")
    # l = len(tickets)-1
    corp_id = 'rikushwa'
    project = 'TOPSI'
    complexity = 'Medium'
    activity_type = ''
    activity = ''
    reference = ''
    effort = 0
    AMorAD = ''
    SOW = ''
    my_tickets_record = []

    tickets = read_data_excel(_ticket_file)

    for ticket in tickets :
        print("ticket: ", ticket)
        if ticket == '' :
            activity_type = 'Meetings / Communication'
            activity = 'Mail Communication'
            effort = 1.5
        elif ticket == '0' :
            activity_type = 'Service-Task'
            activity = 'DSTUM'
            effort = 1.5
        elif ticket[:3] == 'CHG' :
            activity_type = 'Change Request'
            activity = 'Third party coordination'
            effort = 1
        else :
            activity_type = 'Incident'
            activity = 'Incident'
            effort = 0.75

        ticket_details = (
            [ticket, reference, corp_id, activity_type, activity, dt, effort, complexity, AMorAD, SOW, project],)
        print("ticket_details in create_list() : \n", ticket_details)
        for tickets, reference, corp_id, activity_type, activity, date, effort, complexity, AMorAD, SOW, project in ticket_details :
            my_tickets_record.append(
                [ticket, reference, corp_id, activity_type, activity, dt, effort, complexity, AMorAD, SOW, project])

    return my_tickets_record


try:
    if os.path.exists(_file_name) :
        existing_wb = load_workbook(_file_name)
        existing_ws = existing_wb.active
        print("max row ",existing_ws.max_row)
        for row in create_list() :
            existing_ws.append(row)
        existing_wb.save(_file_name)
        existing_wb.close()
    else :
        raise FileNotFoundError("File not found")
except FileNotFoundError as err:
    new_workbook(_file_name)

