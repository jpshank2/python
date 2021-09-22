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

if len(staff) > 0:
    for person in staff:
        if person[2] == 1:
            outlook = wc.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = person[0]
            mail.CC = 'hgeary@bmss.com'
            mail.Subject = 'Current Bingo Game Missed First Violation'
            mail.HTMLBody = """<p>""" + person[1] + """</p><p>You have missed entering your time for yesterday by our cutoff and have been kicked out of this Bingo game. Keep entering your time to avoid further penalty!</p>"""

            mail.Send()

        if person[2] == 2:
            outlook = wc.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = person[0]
            mail.CC = 'hgeary@bmss.com'
            mail.Subject = 'Current Bingo Game Missed Second Violation'
            mail.HTMLBody = """<p>""" + person[1] + """</p><p>You have missed entering your time for yesterday by our cutoff and have contributed $10 to the pot from you Christmas bonus. Keep entering your time to avoid further penalty!</p>"""

            mail.Send()

        if person[2] == 3:
            outlook = wc.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = person[0]
            mail.CC = 'hgeary@bmss.com'
            mail.Subject = 'Current Bingo Game Missed Third Violation'
            mail.HTMLBody = """<p>""" + person[1] + """</p><p>You have missed entering your time for yesterday by our cutoff and have contributed $25 to the pot from you Christmas bonus. Keep entering your time to avoid further penalty!</p>"""

            mail.Send()

        if person[2] > 3:
            outlook = wc.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = person[0]
            mail.CC = 'hgeary@bmss.com'
            mail.Subject = 'Current Bingo Game Missed Fourth or More Violation'
            mail.HTMLBody = """<p>""" + person[1] + """</p><p>You have missed entering your time for yesterday by our cutoff more than three times. Please schedule a time to meet with your Homeroom Leader to talk about your time entry habits.</p>"""

            mail.Send()