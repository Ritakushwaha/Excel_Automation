# https://xlsxwriter.readthedocs.io/contents.html

import xlsxwriter
import datetime


# Create a workbook and add a worksheet.
var = str(datetime.datetime.today()).split()[0]
workbook = xlsxwriter.Workbook(f'Efforts{var}.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold' : True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format' : '$#,##0'})

# Write some data headers.
worksheet.write('A1', 'Ticket No', bold)
worksheet.write('B1', 'Ref No', bold)
worksheet.write('C1', 'Corp ID', bold)
worksheet.write('D1', 'Activity Type', bold)
worksheet.write('E1', 'Activity', bold)
worksheet.write('F1', 'Timecard Entry Date', bold)
worksheet.write('G1', 'Effort', bold)
worksheet.write('H1', 'Complexity', bold)
worksheet.write('I1', 'AMorAD', bold)
worksheet.write('J1', 'SOW', bold)
worksheet.write('K1', 'Project', bold)

tickets = ['INC123', 'INC234']
corp_id = 'rikushwa'
project = 'TOPSI'
complexity = 'Medium'
activity_type = 'Incident'
activity = 'Incident'
reference = ' '
date = str(datetime.datetime.today()).split()[0]
effort = 0.5
AMorAD = ''
SOW = ''

# Some data we want to write to the worksheet.
'''ticket_details = (
    [tickets, reference, corp_id, activity_type, activity, date, effort, complexity, AMorAD, SOW, project],
)'''

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for ticket in tickets:
    ticket_details = ([ticket, reference, corp_id, activity_type, activity, date, effort, complexity, AMorAD, SOW, project],)
    for tickets, reference, corp_id, activity_type, activity, date, effort, complexity, AMorAD, SOW, project in ticket_details :
        worksheet.write(row, col, tickets)
        worksheet.write(row, col + 1, reference)
        worksheet.write(row, col + 2, corp_id)
        worksheet.write(row, col + 3, activity_type)
        worksheet.write(row, col + 4, activity)
        worksheet.write(row, col + 5, date)
        worksheet.write(row, col + 6, effort)
        worksheet.write(row, col + 7, complexity)
        worksheet.write(row, col + 8, AMorAD)
        worksheet.write(row, col + 9, SOW)
        worksheet.write(row, col + 10, project)
        row += 1

workbook.close()
