# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 11:24:25 2021

@author: rikushwa
"""

#!/usr/bin/python

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os
from smtplib import SMTPAuthenticationError, SMTPSenderRefused
from get_mail_data import getData


_date = str(datetime.today()).split()[0]
_file_name = ''


def send_mail_1(message):
    print('Sending Mail')
    fromaddr = getData("fromaddr")
    toaddr = getData("toaddr_balance_check")
    cc = ""
    rcpt = cc.split(",") + [toaddr]
    subject_line = f'''{getData("subject_line_balance_check")} | {_date}'''
    password = str(input("Enter Outlook Password: "))
    body = f'''Hi,\n\nHere is Check daily Balance Process Routine\n\nToday : {message}\n\nBr,\nRita Kushwaha'''


    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Cc'] = cc
    msg['Subject'] = subject_line
    msg.attach(MIMEText(body, 'plain'))


    server = smtplib.SMTP('smtp.outlook.com', 587)
    server.starttls()
    try:
        server.login(fromaddr, password)
        text = msg.as_string()
        server.sendmail(fromaddr, rcpt, text)
        print('Mail sent')
    except SMTPAuthenticationError as err:
        print("Mail not Sent","Authentication Failed: ",err)
    except SMTPSenderRefused as err1:
        print("Mail not Sent","Sender refused: ",err1)
        
    server.quit()

