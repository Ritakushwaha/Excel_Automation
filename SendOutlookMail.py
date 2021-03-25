#!/usr/bin/python

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os


_date = str(datetime.today()).split()[0]
_file_name = f'file.xlsx'


def send_mail(_file_name):
    print('Sending Mail')
    fromaddr = ""
    toaddr = ""
    cc = ""
    rcpt = cc.split(",") + [toaddr]
    subject_line = 'Capacity Report'
    password = str(input('Enter Outlook Password: '))
    body = '''Hi,\nHere is the body of mail\nBr,\nName'''
    filename = _file_name

    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Cc'] = cc
    msg['Subject'] = subject_line
    msg.attach(MIMEText(body, 'plain'))

    path = _file_name
    attachment = open(path, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.outlook.com', 587)
    server.starttls()
    server.login(fromaddr, password)
    text = msg.as_string()
    server.sendmail(fromaddr, rcpt, text)
    print('Mail sent')
    server.quit()


