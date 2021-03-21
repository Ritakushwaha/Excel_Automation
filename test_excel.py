# looping through each ticket to generate the details
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