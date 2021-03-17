# https://xlsxwriter.readthedocs.io/contents.html

import xlsxwriter
from datetime import datetime

# Create a workbook and add a worksheet.
month = datetime.today().month
year = datetime.today().year
dt = str(datetime.today()).split()[0]

workbook = xlsxwriter.Workbook(f'Efforts_{month}_{year}.xlsx')
worksheet = workbook.add_worksheet(f'data_{dt}.xlsx')

# Add a bold format to use to highlight cells.
cf = workbook.add_format({'bg_color': 'yellow'})

# Write some data headers.
worksheet.write('A1', 'Ticket No', cf)
worksheet.write('B1', 'Ref No', cf)
worksheet.write('C1', 'Corp ID', cf)
worksheet.write('D1', 'Activity Type', cf)
worksheet.write('E1', 'Activity', cf)
worksheet.write('F1', 'Timecard Entry Date', cf)
worksheet.write('G1', 'Effort', cf)
worksheet.write('H1', 'Complexity', cf)
worksheet.write('I1', 'AMorAD', cf)
worksheet.write('J1', 'SOW', cf)
worksheet.write('K1', 'Project', cf)

tickets = ['','0','CHG0056957','CHG0056935','SCTASK0081896','INC0666400','SCTASK0083834','CHG0056957','INC0662672','INC0667274','INC0667610']
l = len(tickets)-1
corp_id = 'rikushwa'
project = 'TOPSI'
complexity = 'Medium'
activity_type = ''
activity = ''
reference = ''
effort = 0
AMorAD = ''
SOW = ''
date_format = workbook.add_format({'num_format':'yyyy-mm-dd hh:mm:ss'})

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for ticket in tickets:
    if ticket == '':
        activity_type = 'Meetings / Communication'
        activity = 'Mail Communication'
        effort = 1.5
    elif ticket == '0':
        activity_type = 'Service-Task'
        activity = 'DSTUM'
        effort = 1.5
    elif ticket[:3] == 'CHG':
        activity_type = 'Change Request'
        activity = 'Third party coordination'
        effort = 1
    else:
        activity_type = 'Incident'
        activity = 'Incident'
        effort = 0.75

    ticket_details = ([ticket, reference, corp_id, activity_type, activity, dt, effort, complexity, AMorAD, SOW, project],)

    for tickets, reference, corp_id, activity_type, activity, date, effort, complexity, AMorAD, SOW, project in ticket_details :
        worksheet.write(row, col, tickets)
        worksheet.write(row, col + 1, reference)
        worksheet.write(row, col + 2, corp_id)
        worksheet.write(row, col + 3, activity_type)
        worksheet.write(row, col + 4, activity)
        worksheet.write(row, col + 5, datetime.today(),date_format)
        worksheet.write(row, col + 6, effort)
        worksheet.write(row, col + 7, complexity)
        worksheet.write(row, col + 8, AMorAD)
        worksheet.write(row, col + 9, SOW)
        worksheet.write(row, col + 10, project)
        row += 1

worksheet.write(row+1, col+6, f'=SUM(G3:G{l})')

while True:
    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        decision = input("Exception caught in workbook.close(): %s\n"
                         "Please close the file if it is open in Excel.\n"
                         "Try to write file again? Type [n]: " % e)
        if decision != 'n':
            continue
    break
