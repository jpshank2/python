import win32com.client as wc
import pyodbc, os

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DEV_SERVER') + ';DATABASE=' + os.getenv('DEV_DATABASE') + ';UID=' + os.getenv('DEV_USER') + ';PWD=' + os.getenv('DEV_PASS'))
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))
outlook = wc.Dispatch('outlook.application')

homeroomLeaders = engine.cursor()
homeroomLeaders.execute("""SELECT S.StaffIndex, S.StaffName, S.StaffEMail
FROM dbo.tblStaff S
	INNER JOIN dbo.tblCategory C ON C.CatName = S.StaffName AND CatType = 'SUBDEPT'""")

leaders = homeroomLeaders.fetchall()

for leader in leaders:
    homeroomMembers = conn.cursor()
    homeroomMembers.execute("""SELECT MAX(SubDate) AS SubDate, SubSender, SubRecipient
FROM MandM.Submissions
WHERE SubType = 3 AND SubSender = """ + str(leader[0]) + """
GROUP BY SubSender, SubRecipient
HAVING CONVERT(date, MAX(SubDate)) > CONVERT(date, DATEADD(DAY, -3, GETDATE()))""")
    members = homeroomMembers.fetchall()

    checkedListForSQL = '0'

    for i in range(len(members)):
        checkedListForSQL += ', ' + str(members[i][2])

    notCheckedMembersList = engine.cursor()
    notCheckedMembersList.execute("""SELECT S.StaffIndex, S.StaffName
                                FROM dbo.tblStaff S
                                    INNER JOIN dbo.tblStaffEx SE ON SE.StaffIndex = S.StaffIndex
	                                INNER JOIN dbo.tblCategory C ON SE.StaffSubDepartment = C.Category AND C.CatType = 'SUBDEPT'
                                WHERE S.StaffIndex NOT IN (""" + checkedListForSQL + """) AND C.CatName = '""" + leader[1] + """'""")
    notCheckedMembers = notCheckedMembersList.fetchall()

    notChecked = ''

    for member in notCheckedMembers:
        notChecked += '<li>' + member[1] + '</li>'
            
    if len(notChecked) > 0:        
        mail = outlook.CreateItem(0)
        
        mail.To = leader[2]
        mail.Bcc = 'jeremyshank@bmss.com'
        mail.Subject = 'Homeroom Check In'
        mail.HTMLBody = '<p>' + leader[1] + ',</p><p>&emsp;Don\'t forget to check in with these Homeroom members and log it in the M+M Outlook app this week!</p><ul>' + notChecked + '</ul><p>Thanks!</p>'

        mail.Send()