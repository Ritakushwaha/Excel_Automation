# -*- coding: utf-8 -*-
"""
Created on Fri Apr 16 11:41:38 2021

@author: rikushwa
"""
import ibm_db
from DB2_connection import get_connection
from Send_Email_Check_Daily_Balance_Process import send_mail_1

def check_balance():
    print("Working instruction for following the validity of the new daily balance process")
    conn = get_connection()
    sql = f'''Query with columns in list'''

    stmt = ibm_db.exec_immediate(conn, sql)
    dictionary = ibm_db.fetch_both(stmt)
    datalist = []
    while dictionary:
        list.append([dictionary["KYTUNNISTE"].strip(), dictionary["MORMAID"].strip(), dictionary["TUOTEAVAIN"].strip(),dictionary["SALDO"].strip(),
                     dictionary["VIIMTARKSALDO"].strip(), dictionary["VIIMTARKAIKA"].strip(), dictionary["PERUSTAJA"].strip(),dictionary["PERUSTAMISAIKA"].strip(),
                     dictionary["MUUTOSAIKA"].strip(), dictionary["SNAPSHOTAIKA"].strip(), dictionary["MARTTIKYID"].strip(),dictionary["TUOTEAVAIN"].strip(),
                     dictionary["PAIVA"].strip(), dictionary["PAIVASALDOMUUTOS"].strip(), dictionary["VIIVESALDOMUUTOS"].strip(),dictionary["ALKUSALDO"].strip(),
                     dictionary["MUUTOSAIKA"].strip(), dictionary["ALKUSALDOLADATTU"].strip()])
        dictionary = ibm_db.fetch_both(stmt)
    return datalist

data = check_balance()

message = ""
if len(data)>0:
    message = f"{len(data)} Rows found\nPlease check."
    send_mail_1(message)
else:
    message = "No rows found"
    send_mail_1(message)