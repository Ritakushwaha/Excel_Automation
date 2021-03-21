import mysql.connector
from mysql.connector import Error
from mysql.connector import errorcode


class MyDB:

    def get_connection(self):
        try:
            #connection to MYSQL
            conn = mysql.connector.connect(host="localhost",user="root",password="viraj200321")
            mycursor = conn.cursor()
            # creation of database if not exists
            mycursor.execute("CREATE DATABASE IF NOT EXISTS test_db")
            # connection to database
            conn = mysql.connector.connect(host="localhost", user="root", password="viraj200321",database='gym_manager')
            return conn
        except Error as err :
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR :
                print("Something is wrong with your user name or password")
            elif err.errno == errorcode.ER_BAD_DB_ERROR :
                print("Database doeasn't exists")
            else :
                print(err)
        else :
            conn.close()