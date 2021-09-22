import pyodbc, os
import win32com.client as wc
from dotenv import load_dotenv

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

getNewStaff = conn.cursor()
getNewStaff.execute("""SELECT [StaffIndex]
      ,[StaffName]
      ,[StaffEMail]
  FROM [DataWarehouse].[dbo].[tblStaff]
  WHERE StaffEnded IS NULL AND StaffName NOT IN ('Administration', 'E-File Pool', 'Cindy Cpa', 'Robin Muncy', 'Dakota Woodson', 'Akilah Boker', 'Dianne Hart') AND StaffBingo IS NULL
  ORDER BY StaffIndex""")

newStaff = getNewStaff.fetchall()

getFreeCards = conn.cursor()
getFreeCards.execute("""SELECT DISTINCT BingoCard FROM dbo.Bingo
WHERE BingoCard NOT IN (
SELECT StaffBingo FROM dbo.tblStaff WHERE StaffBingo IS NOT NULL
)""")

freeCards = getFreeCards.fetchall()

for i in range(0, len(newStaff)):
    updateBingo = conn.cursor()
    updateBingo.execute("""UPDATE dbo.tblStaff
                            SET StaffBingo = """ + str(freeCards[i][0]) + """
                            WHERE StaffIndex = """ + str(newStaff[i][0]))
    conn.commit()

    path = "C:\\Users\\jeremyshank\\Documents\\Bingo\\Cards\\" + str(freeCards[i][0]) + ".xlsx"
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = newStaff[i][2]
    mail.Subject = 'New Bingo Card'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>' + newStaff[i][1] + ',</p><p>Your new Bingo card is attached. Rules for the Bingo Game can be found <a href="http://zeal.bmss.com/kb/time-entry-and-bingo/">on Zeal</a>.</p><p>Good luck!</p>'

    mail.Send()
