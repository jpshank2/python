import pyodbc, sys
import win32com.client as wc


try:
    conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

    emailrow = conn.cursor()
    emailrow.execute("""SELECT DISTINCT S.Email, S.Staff 
    FROM (
        SELECT E.ClientName,
            JH.Job_Name,
            CAST(SUM(TW.WIPHours) AS nvarchar) AS [Billed Hours],
            CAST(JH.Job_Budget_Hours AS nvarchar) AS [Budget Hours],
            M.StaffName AS [Staff],
            M.StaffEMail AS [Email]
        FROM tblTranWIP TW 
        INNER JOIN dbo.tblJob_Header JH ON TW.ServPeriod = JH.Job_Idx
        INNER JOIN dbo.tblEngagement E ON E.ContIndex = JH.ContIndex
        INNER JOIN dbo.tblStaff M ON M.StaffIndex = E.ClientManager
        WHERE TW.TransTypeIndex in ('1','2') 
        AND TW.WIPDate > '01/01/2018' 
        --AND DATEPART(YEAR, JH.Job_Created) = DATEPART(YEAR, GETDATE()) 
        AND JH.Job_Status NOT IN (2, 3)
        AND JH.Job_Budget_Hours <> 0
        GROUP BY JH.Job_Name, E.ClientName, JH.Job_Budget_Hours, M.StaffName, M.StaffEMail
        HAVING SUM(TW.WIPHours) > JH.Job_Budget_Hours
    ) S""")

    emails = emailrow.fetchall()

    for email in emails:
        tableRows = ""

        jobrow = conn.cursor()
        jobrow.execute("""SELECT E.ClientName,
            JH.Job_Name,
            CAST(SUM(TW.WIPHours) AS nvarchar) AS [Billed Hours],
            CAST(JH.Job_Budget_Hours AS nvarchar) AS [Budget Hours],
            M.StaffName AS [Staff],
            M.StaffEMail AS [Email]
        FROM tblTranWIP TW 
        INNER JOIN dbo.tblJob_Header JH ON TW.ServPeriod = JH.Job_Idx
        INNER JOIN dbo.tblEngagement E ON E.ContIndex = JH.ContIndex
        INNER JOIN dbo.tblStaff M ON M.StaffIndex = E.ClientManager
        WHERE TW.TransTypeIndex in ('1','2') 
        AND TW.WIPDate > '01/01/2018' 
        --AND DATEPART(YEAR, JH.Job_Created) = DATEPART(YEAR, GETDATE()) 
        AND JH.Job_Status NOT IN (2, 3)
        AND JH.Job_Budget_Hours <> 0
        AND M.StaffEMail = '""" + email[0] + """'
        GROUP BY JH.Job_Name, E.ClientName, JH.Job_Budget_Hours, M.StaffName, M.StaffEMail
        HAVING SUM(TW.WIPHours) > JH.Job_Budget_Hours""")
        
        jobs = jobrow.fetchall()

        for job in jobs:
            tableRows += """<tr><td style="margin:8px">""" + job[0] + """</td><td style="margin:8px">""" + job[1] + """</td><td style="margin:8px;text-align:center">""" + job[3] + """</td><td style="margin:8px;text-align:center">""" + job[2] + """</td></tr>"""

        # print(email)
        outlook = wc.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = 'jeremyshank@bmss.com'
        mail.Subject = 'Over Budget Jobs'
        mail.HTMLBody = """<p>""" + email[1] + """,</p><p>The following jobs are over budget, please check on them.</p><table><tr><th style="margin:8px">Client</th><th style="margin:8px">Job</th><th style="margin:8px">Budget Hours</th><th style="margin:8px">Actual Hours</th></tr>""" + tableRows + """</table><p>Thank you!</p><p>should go to """ + email[0] + """</p>"""

        mail.Send()
except Exception:
    print(sys.stderr)