# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 09:07:56 2021

@author: rikushwa
"""

import ibm_db
import sys
from get_db2_data import getData


def get_connection():
    db_name = getData("db_name")
    db_host_name = getData("db_host_name")
    db_port = getData("db_port")
    db_protocol = getData("db_protocol")
    db_username = getData("db_username")
    db_password = getData("db_password")

    try:
        conn = ibm_db.connect(
            f"DATABASE = {db_name}; HOSTNAME = {db_host_name}; PORT = {db_port}; PROTOCOL = {db_protocol}; UID = {db_username}; PWD = {db_password};",
            "", "")
        return conn
    except:
        print("no connection:", ibm_db.conn_errormsg())
        sys.exit(1)

