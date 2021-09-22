import win32com.client as wc
import pyodbc, re, os

conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS'))


def email(record):
    noEL = conn.cursor()
    noEL.execute("""SELECT		 e.ClientCode	AS [Client ID]
			,E.ClientName	AS [Client Name]
			/*,Own.StaffName AS [Client Originator]*/
			--,SJ.JurisName AS Jurisdiction
			,Ptnr.StaffName AS [Member]
			,Mgr.StaffName AS [Client Manager]
			,Jm.StaffName AS [Job Partner]
            ,Jb.StaffName AS [Job Manager]
			,TP.WorkStatusDesc AS [Workflow Status]

			--,N.NoteText

FROM tblEngagement AS E
		
		INNER JOIN tblJob_Header  AS J	ON j.ContIndex = e.ContIndex 
		INNER JOIN tblPortfolio_Job AS TP ON J.Job_Idx = TP.Job_Idx
		LEFT JOIN tblJob_Work_Status AS ws ON ws.StatusIndex = j.Job_WorkStatus
		LEFT JOIN tblJob_Roles P ON P.Job_Idx = J.Job_Idx And P.RoleIndex = 2 
		LEFT JOIN tblJob_Roles R ON R.Job_Idx = J.Job_Idx And R.RoleIndex = 3
		LEFT JOIN tblStaff PS ON PS.StaffIndex = P.StaffIndex
		LEFT JOIN tblStaff RS ON RS.StaffIndex = R.StaffIndex
		INNER Join tblStaff AS Ptnr ON E.ClientPartner = Ptnr.StaffIndex
		INNER join tblStaff AS Mgr ON E.ClientManager = Mgr.StaffIndex
		INNER join tblStaff AS JM ON j.Job_Partner = Jm.StaffIndex
        INNER join tblStaff AS Jb ON j.Job_Manager = Jb.StaffIndex
		INNER JOIN tblStaff AS S ON j.Job_CurrentStaff = S.StaffIndex
		/*INNER join tblClientOrigination AS CO ON E.ContIndex = CO.ContIndex
		INNER join tblStaff AS OWN ON CO.StaffIndex = OWN.StaffIndex*/
		INNER JOIN tblJob_Serv JS ON JS.Job_Idx = J.Job_Idx 
		INNER JOIN tblServices SV ON SV.ServIndex = JS.ServIndex 
		LEFT JOIN tblJob_TaxReturn T ON T.Job_Idx = J.Job_Idx
		LEFT JOIN tblJob_TaxReturn_Jurisdictions zTRJ ON zTRJ.Job_Idx = J.Job_Idx 
		Inner Join tblStateJurisdiction SJ ON SJ.JurisIndex = zTRJ.JurisIndex
		LEFT JOIN (SELECT	 jh.Job_Idx 
								,MAX(jh.HistDate) AS [HistDate]
						FROM	tblJob_History jh
						GROUP BY jh.Job_Idx) jhm ON J.Job_Idx = jhm.Job_Idx 

WHERE e.ClientStatus NOT IN ('LOST', 'INTERNAL') 
AND E.ClientOffice = 'BHM' 
AND J.Job_Name = '2021 Business Personal Property Tax'
AND TP.WorkStatusDesc = 'Waiting on Signed Engagement Letter'
AND Jb.StaffName = '""" + record[0] + """'

GROUP BY e.ClientCode, e.ClientName, Ptnr.StaffName, Mgr.StaffName, Jm.StaffName, Jb.StaffName, TP.WorkStatusDesc

ORDER BY e.ClientCode""")
    noLetter = noEL.fetchall()

    noLetterList = ""
    if len(noLetter) > 0:
        for letter in noLetter:
            noLetterList += '<li>' + letter[0] + \
                '&emsp;&emsp;' + letter[1] + '</li>'
    else:
        noLetterList = "None!"

    noAD = conn.cursor()

    noAD.execute("""SELECT		 e.ClientCode	AS [Client ID]
			,E.ClientName	AS [Client Name]
			/*,Own.StaffName AS [Client Originator]*/
			--,SJ.JurisName AS Jurisdiction
			,Ptnr.StaffName AS [Member]
			,Mgr.StaffName AS [Client Manager]
			,Jm.StaffName AS [Job Partner]
            ,Jb.StaffName AS [Job Manager]
			,TP.WorkStatusDesc AS [Workflow Status]

			--,N.NoteText

FROM tblEngagement AS E
		
		INNER JOIN tblJob_Header  AS J	ON j.ContIndex = e.ContIndex 
		INNER JOIN tblPortfolio_Job AS TP ON J.Job_Idx = TP.Job_Idx
		LEFT JOIN tblJob_Work_Status AS ws ON ws.StatusIndex = j.Job_WorkStatus
		LEFT JOIN tblJob_Roles P ON P.Job_Idx = J.Job_Idx And P.RoleIndex = 2 
		LEFT JOIN tblJob_Roles R ON R.Job_Idx = J.Job_Idx And R.RoleIndex = 3
		LEFT JOIN tblStaff PS ON PS.StaffIndex = P.StaffIndex
		LEFT JOIN tblStaff RS ON RS.StaffIndex = R.StaffIndex
		INNER Join tblStaff AS Ptnr ON E.ClientPartner = Ptnr.StaffIndex
		INNER join tblStaff AS Mgr ON E.ClientManager = Mgr.StaffIndex
		INNER join tblStaff AS JM ON j.Job_Partner = Jm.StaffIndex
        INNER join tblStaff AS Jb ON j.Job_Manager = Jb.StaffIndex
		INNER JOIN tblStaff AS S ON j.Job_CurrentStaff = S.StaffIndex
		/*INNER join tblClientOrigination AS CO ON E.ContIndex = CO.ContIndex
		INNER join tblStaff AS OWN ON CO.StaffIndex = OWN.StaffIndex*/
		INNER JOIN tblJob_Serv JS ON JS.Job_Idx = J.Job_Idx 
		INNER JOIN tblServices SV ON SV.ServIndex = JS.ServIndex 
		LEFT JOIN tblJob_TaxReturn T ON T.Job_Idx = J.Job_Idx
		LEFT JOIN tblJob_TaxReturn_Jurisdictions zTRJ ON zTRJ.Job_Idx = J.Job_Idx 
		Inner Join tblStateJurisdiction SJ ON SJ.JurisIndex = zTRJ.JurisIndex
		LEFT JOIN (SELECT	 jh.Job_Idx 
								,MAX(jh.HistDate) AS [HistDate]
						FROM	tblJob_History jh
						GROUP BY jh.Job_Idx) jhm ON J.Job_Idx = jhm.Job_Idx 

WHERE e.ClientStatus NOT IN ('LOST', 'INTERNAL') 
AND E.ClientOffice = 'BHM' 
AND J.Job_Name = '2021 Business Personal Property Tax' 
--AND TP.WorkStatusDesc IN ('Waiting on Signed Engagement Letter', 'Waiting on Worksheets - AD VALOREM')
AND TP.WorkStatusDesc = 'Waiting on Worksheets - AD VALOREM'
AND Jb.StaffName = '""" + record[0] + """'

GROUP BY e.ClientCode, e.ClientName, Ptnr.StaffName, Mgr.StaffName, Jm.StaffName, Jb.StaffName, TP.WorkStatusDesc

ORDER BY e.ClientCode""")

    noADValorem = noAD.fetchall()

    noADList = ""

    if len(noADValorem) > 0:
        for AD in noADValorem:
            noADList += '<li>' + AD[0] + '&emsp;&emsp;' + AD[1] + '</li>'
    else:
        noADList = "None!"

    noNS = conn.cursor()
    noNS.execute("""SELECT		 e.ClientCode	AS [Client ID]
			,E.ClientName	AS [Client Name]
			/*,Own.StaffName AS [Client Originator]*/
			--,SJ.JurisName AS Jurisdiction
			,Ptnr.StaffName AS [Member]
			,Mgr.StaffName AS [Client Manager]
			,Jm.StaffName AS [Job Partner]
            ,Jb.StaffName AS [Job Manager]
			,TP.WorkStatusDesc AS [Workflow Status]

			--,N.NoteText

FROM tblEngagement AS E
		
		INNER JOIN tblJob_Header  AS J	ON j.ContIndex = e.ContIndex 
		INNER JOIN tblPortfolio_Job AS TP ON J.Job_Idx = TP.Job_Idx
		LEFT JOIN tblJob_Work_Status AS ws ON ws.StatusIndex = j.Job_WorkStatus
		LEFT JOIN tblJob_Roles P ON P.Job_Idx = J.Job_Idx And P.RoleIndex = 2 
		LEFT JOIN tblJob_Roles R ON R.Job_Idx = J.Job_Idx And R.RoleIndex = 3
		LEFT JOIN tblStaff PS ON PS.StaffIndex = P.StaffIndex
		LEFT JOIN tblStaff RS ON RS.StaffIndex = R.StaffIndex
		INNER Join tblStaff AS Ptnr ON E.ClientPartner = Ptnr.StaffIndex
		INNER join tblStaff AS Mgr ON E.ClientManager = Mgr.StaffIndex
		INNER join tblStaff AS JM ON j.Job_Partner = Jm.StaffIndex
        INNER join tblStaff AS Jb ON j.Job_Manager = Jb.StaffIndex
		INNER JOIN tblStaff AS S ON j.Job_CurrentStaff = S.StaffIndex
		/*INNER join tblClientOrigination AS CO ON E.ContIndex = CO.ContIndex
		INNER join tblStaff AS OWN ON CO.StaffIndex = OWN.StaffIndex*/
		INNER JOIN tblJob_Serv JS ON JS.Job_Idx = J.Job_Idx 
		INNER JOIN tblServices SV ON SV.ServIndex = JS.ServIndex 
		LEFT JOIN tblJob_TaxReturn T ON T.Job_Idx = J.Job_Idx
		LEFT JOIN tblJob_TaxReturn_Jurisdictions zTRJ ON zTRJ.Job_Idx = J.Job_Idx 
		Inner Join tblStateJurisdiction SJ ON SJ.JurisIndex = zTRJ.JurisIndex
		LEFT JOIN (SELECT	 jh.Job_Idx 
								,MAX(jh.HistDate) AS [HistDate]
						FROM	tblJob_History jh
						GROUP BY jh.Job_Idx) jhm ON J.Job_Idx = jhm.Job_Idx 

WHERE e.ClientStatus NOT IN ('LOST', 'INTERNAL') 
AND E.ClientOffice = 'BHM' 
AND J.Job_Name = '2021 Business Personal Property Tax'
AND TP.WorkStatusDesc = 'Not Started'
AND Jb.StaffName = '""" + record[0] + """'

GROUP BY e.ClientCode, e.ClientName, Ptnr.StaffName, Mgr.StaffName, Jm.StaffName, Jb.StaffName, TP.WorkStatusDesc

ORDER BY e.ClientCode""")
    notStarted = noNS.fetchall()

    notStartedList = ""
    if len(notStarted) > 0:
        for letter in notStarted:
            notStartedList += '<li>' + letter[0] + \
                '&emsp;&emsp;' + letter[1] + '</li>'
    else:
        notStartedList = "None!"

    if len(noLetter) > 0 or len(noADValorem) > 0 or len(notStarted) > 0:
        if re.search("''", record[0]) != None:
            record[0] = re.sub("''", "'", record[0])
        myFirstLast = record[0].split(' ')
        outlook = wc.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = 'esmith@bmss.com'  # record[1]
        mail.Subject = 'Please Read - Important Ad Valorem Info'
        mail.HTMLBody = '<p>Hi ' + myFirstLast[0] + ',</p><p>I just wanted to follow up again regarding clients of yours that have not provided us with a signed engagement letter, have not completed worksheets, or have not been started for the 2021 Business Personal Property Tax year:</p><u><strong>No Signed Engagement Letter:</strong></u><ul>' + \
            noLetterList + '</ul><br><u><strong>No Completed Worksheets:</strong></u><ul>' + noADList + \
            '</ul><br><u><strong>Not Started:</u></strong><ul>' + notStartedList + \
            '</ul><p>If you could please reach out to these clients to encourage them to get these items back to us, or let me know if you would like us to go ahead and prepare based on other info, that would be great!</p><p>Thanks,</p><p>Ellie Smith</p>'
        mail.Importance = 2
        mail.Send()


namerow = conn.cursor()

namerow.execute("""Select			--P.PersonTitle AS Title
				ts.StaffName AS Employee
				,ts.StaffEMail
				--,Convert(VARCHAR,TS.StaffStarted,101) as [Start Date]
				--,Convert(VARCHAR,ts.StaffEnded,101) AS [End Date]
				
	FROM       tblStaff ts 
	inner join tblGrade G ON G.GradeCode = ts.StaffCategory
	INNER JOIN tblDepartment AS D ON TS.StaffDepartment = D.DeptIdx
	inner join tblOffices AS Loc ON ts.StaFFOffice = Loc.OfficeCode
	inner join tblContactAttributes AS CA ON ts.StaffIndex = CA.ContIndex
	inner join tblContacts AS TC ON ts.StaffIndex = TC.ContIndex
	inner join tblPerson AS P ON ts.StaffIndex = P.ContIndex
	inner Join tblCategory C ON C.Category = CA.AttrValid AND C.CatType = 'HOMEROOM'


WHERE Loc.OfficeName <> ' No Selection' and TS.StaffEnded is Null AND TS.StaffManager = '-1' --AND ts.StaffName NOT IN ('Abigail Waddell', 'AJ Vanderwoude', 'Amanda Houston', 'Auston Sullivan', 'Bill Lorimer')""")

names = namerow.fetchall()
# print(names[1][0])

for name in names:
    if re.search("'", name[0]) != None:
        nick = re.sub("'", "''", name[0])
        email([nick, name[1]])
    else:
        email(name)

# for i in range(0, 15):
# 	if re.search("'", names[i][0]) != None:
# 		nick = re.sub("'", "''", name[i][0])
# 		email([nick, name[i][1]])
# 	else:
# 		email(names[i])