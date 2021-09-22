#! python3

import pyodbc, openpyxl, re, os
from month import month
from april import april
from manager import manager
from member import member
from originator import originator
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))


namerow = conn.cursor()
namerow.execute("""SELECT [StaffName]
  FROM [dbo].[tblStaff]
  WHERE StaffEnded IS NULL""")

names = namerow.fetchall()
for lname in names: 
    filename = "C:\\Users\\jeremyshank\\Desktop\\Reports\\" + lname[0] + " Report.xlsx"
    name = re.sub("[']", "''", lname[0])
    wb = openpyxl.Workbook()
    month(conn, name, wb)
    april(conn, name, wb)
    manager(conn, name, wb)
    member(conn, name, wb)
    originator(conn, name, wb)
    wb.save(filename)
