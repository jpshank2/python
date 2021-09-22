from dotenv import load_dotenv
import win32com.client as wc
import pyodbc, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))

def email(record):
    path = "C:\\Users\\jeremyshank\\Desktop\\No Macro\\" + \
        record[0] + " Client Update Project.xlsx"
    outlook = wc.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = record[1]
    mail.Subject = 'Practice Engine Client Data Update Project'
    mail.Attachments.Add(path)
    mail.HTMLBody = '<p>' + record[0] + ',</p><p>Attached is a spreadsheet of clients you are responsible for updating.  Please complete Phase 1 below by 5pm on 11/18/20 (next Wednesday) and return to David Brown.</p><p><strong><u>Phase 1 – Review Client Level Information - to be completed by 11/18/20:</u></strong></p><ul><li>This will be done by the person currently assigned as Originator on the client.</li><li>By the end of the day 11/12/20, David Brown will email each originator a spreadsheet with only your clients.  All changes you will need to make are choices from a drop down list, which should make the process easy.</li><li>Review the list for 4 items:</li><ol><li>Entity (make sure this is correct)</li><li>Industry (make sure this is correct)</li><li>Client Partner (see attached) (change to appropriate person if necessary)</li><li>Client Manager (see attached) (change to appropriate person if necessary)</li></ol><li>Return the spreadsheet to <strong><a href="mailto:dbrown@bmss.com">David Brown</a></strong> no later than 5:00 on 11/18/20</li><li><strong>WE ONLY HAVE ONE SHOT TO CONSOLIDATE ALL INFORMATION INTO ONE SPREADSHEET AND IMPORT THE DATA INTO PE!  Therefore, it is important that you get this done on time, otherwise, each of these changes will have to be made manually, one at a time.</strong></li></ul>'

    mail.Send()
    # print(record[1][1:])


namerow = conn.cursor()

namerow.execute("""select 
	T.Originator,
	S.StaffEmail
From
(select
	(CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE (CASE WHEN O.STAFFNAME='Bill House' THEN 'Tom Albright' WHEN O.StaffName='Randy Barnes' THEN 'Greg Barnes' ELSE P.STAFFNAME END) END) as [Originator]

from tblEngagement E
	Left Join tblCategory I on E.ClientIndustry = I.Category and I.CatType = 'INDUSTRY'
	Inner Join tblClientOrigination OWN on E.Contindex=OWN.ContIndex
	Inner Join tblStaff O on OWN.StaffIndex=O.StaffIndex
	Left Join tblOwnerType Ent on E.ClientOwnership = Ent.OwnerIndex
	Inner Join tblStaff P on E.ClientPartner = P.StaffIndex
	Inner Join tblStaff M on E.ClientManager = M.StaffIndex
where E.ClientStatus NOT IN ('Lost','Internal','Suspended') and E.ClientCode<>'RLBTest' and (CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE P.STAFFNAME END) <> 'Administration'
group by (CASE WHEN O.StaffCategory='1' THEN O.StaffName ELSE (CASE WHEN O.STAFFNAME='Bill House' THEN 'Tom Albright' WHEN O.StaffName='Randy Barnes' THEN 'Greg Barnes' ELSE P.STAFFNAME END) END)
) T
Inner Join tblStaff S ON T.Originator=S.StaffName
order by T.Originator""")

names = namerow.fetchall()

for name in names:
    email(name)

# email(names[0])
