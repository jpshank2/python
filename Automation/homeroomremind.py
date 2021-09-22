import win32com.client as wc
import pyodbc, os
from datetime import datetime, timedelta

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))

homeroomLeaders = conn.cursor()
homeroomLeaders.execute("""SELECT ML.*
	  ,S.StaffName
	  ,S.StaffEMail
  FROM [DataWarehouse].[dbo].[MandMLeaders] ML
  INNER JOIN dbo.tblStaff S ON ML.StaffIndex = S.StaffIndex
  WHERE CAST(ML.Category AS int) > 1 AND CatName <> 'STEERING COMMITTEE'""")

counter = 0

leaders = homeroomLeaders.fetchall()

for leader in leaders:
    homeroomMembers = conn.cursor()
    homeroomMembers.execute("""SELECT S.StaffIndex, S.StaffName, S.StaffCode, (SELECT TOP 1 EventDate
                                FROM dbo.MandM
                                WHERE EventPerson = S.StaffName
                                ORDER BY EventDate DESC) AS [LastDate]
                            FROM dbo.tblStaff S
                            WHERE S.StaffAttribute = """ + str(leader[2]) +  
                            """AND S.StaffEnded IS NULL;""")
    members = homeroomMembers.fetchall()

    for member in members:
        if member[3] != None:
            x = member[3]
            y = datetime.now() - x
            if y.total_seconds() > 691199.0:
                outlook = wc.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)

                mail.To = 'jeremyshank@bmss.com'
                #mail.CC = 'support@bmss.com'
                mail.Subject = 'Homeroom Check In'
                mail.HTMLBody = '<p>' + leader[8] + ',</p><p>&emsp;Don\'t forget to check in with ' + member[1] + '!</p><p>Thanks!</p>'

                mail.Send()
                mail.Quit()
        else:
            outlook = wc.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = 'jeremyshank@bmss.com'
            #mail.CC = 'support@bmss.com'
            mail.Subject = 'Homeroom Check In'
            mail.HTMLBody = '<p>' + leader[8] + ',</p><p>&emsp;Don\'t forget to check in with ' + member[1] + '!</p><p>Thanks!</p>'

            mail.Send()
            mail.Quit()