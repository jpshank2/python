import win32com.client as wc
import pyodbc, os
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

getMissed = conn.cursor()
getMissed.execute("""SELECT DISTINCT S.StaffEMail
	  ,S.StaffName
	  ,BingoMissed
  FROM [DataWarehouse].[dbo].[Bingo] B
  INNER JOIN dbo.tblStaff S ON B.BingoCard = S.StaffBingo
  WHERE BingoMissed != 0""")

staff = getMissed.fetchall()

people = ""

if len(staff) > 0:
    for person in staff:
        people += "<li>" + person[1] + " has missed the current Bingo game " + person[2] + "times</li>"

outlook = wc.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.To = 'hgeary@bmss.com'
mail.CC = 'jeremyshank@bmss.com'
mail.Subject = 'Current Bingo Game Missed List'
mail.HTMLBody = """<p>The following people have missed the current Bingo game:</p><ul>""" + people + """</ul>"""

mail.Send()