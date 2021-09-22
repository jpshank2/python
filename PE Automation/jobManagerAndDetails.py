from dotenv import load_dotenv
import requests, os, pyodbc, json

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

# getMASJobs = engine.cursor()

# getMASJobs.execute("""SELECT Job_Idx, JH.ContIndex, Job_PreviousJob, Job_CurrentStaff, Job_Name, Job_Code, Job_Class, Job_Status, Job_WorkStatus, Job_Office, Job_Dept, Job_MasterFile, SUM(TW.WIPOutstanding) AS [Outstanding_WIP]
# FROM dbo.tblJob_Header JH
# 	INNER JOIN tblEngagement E ON E.ContIndex = JH.ContIndex AND E.ClientStatus <> 'LOST'
# 	INNER JOIN tblTranWIP TW ON TW.ContIndex = JH.ContIndex
# WHERE Job_Template = 72 
# 	AND Job_Status NOT IN (2, 3)
# GROUP BY Job_Idx, JH.ContIndex, Job_PreviousJob, Job_CurrentStaff, Job_Name, Job_Code, Job_Class, Job_Status, Job_WorkStatus, Job_Office, Job_Dept, Job_MasterFile
# ORDER BY Job_Idx""")

# MASJobs = getMASJobs.fetchall()

# getINDClients = engine.cursor()

# getINDClients.execute("""SELECT E.ContIndex
# 	,E.ClientName
# 	,C.ContType
# 	,S.ServPartner
# 	,S.ServManager
# FROM dbo.tblEngagement E
# 	INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex
# 	INNER JOIN dbo.tblClientServices S ON S.ContIndex = E.ContIndex AND S.ServIndex = 'INDTAX'
# WHERE E.ClientStatus <> 'LOST' AND C.ContType = 1""")

# INDClients = getINDClients.fetchall()

getBUSClients = engine.cursor()

getBUSClients.execute("""SELECT E.ContIndex
	,E.ClientName
	,C.ContType
	,S.ServPartner
	,S.ServManager
FROM dbo.tblEngagement E
	INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex
	INNER JOIN dbo.tblClientServices S ON S.ContIndex = E.ContIndex AND S.ServIndex = 'BUSTAX'
WHERE E.ClientStatus <> 'LOST' AND C.ContType = 2 AND C.ContIndex NOT IN (SELECT E.ContIndex
FROM dbo.tblEngagement E
	INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex
	INNER JOIN dbo.tblClientServices S ON S.ContIndex = E.ContIndex AND S.ServIndex = 'BUSTAX'
	INNER JOIN dbo.tblJob_Header JH ON JH.ContIndex = E.ContIndex AND JH.Job_Template = 336
WHERE E.ClientStatus <> 'LOST' AND C.ContType = 2)""")

BUSClients = getBUSClients.fetchall()

# getActiveClients = engine.cursor()

# getActiveClients.execute("""SELECT E.ContIndex
# 	,E.ClientCode
# 	,C.ContType
# 	,CASE
# 		WHEN S.ServPartner = 0 THEN E.ClientPartner
# 		WHEN SP.StaffEnded IS NULL THEN S.ServPartner
# 		ELSE E.ClientPartner
# 	END AS ServPartner
# 	,CASE
# 		WHEN S.ServManager = 0 THEN E.ClientManager
# 		WHEN SM.StaffEnded IS NULL THEN S.ServManager
# 		ELSE E.ClientManager
# 	END AS ServManager
# FROM dbo.tblEngagement E
# 	INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex
# 	INNER JOIN dbo.tblClientServices S ON S.ContIndex = E.ContIndex AND S.ServIndex = 'MAS'
# 	INNER JOIN dbo.tblStaff SP ON SP.StaffIndex = S.ServPartner
# 	INNER JOIN dbo.tblStaff SM ON SM.StaffIndex = S.ServManager
# WHERE E.ClientStatus <> 'LOST' AND E.ContIndex NOT IN (SELECT E.ContIndex
# FROM dbo.tblEngagement E
# 	INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex
# 	INNER JOIN dbo.tblClientServices S ON S.ContIndex = E.ContIndex AND S.ServIndex = 'MAS'
# 	INNER JOIN dbo.tblStaff SP ON SP.StaffIndex = S.ServPartner
# 	INNER JOIN dbo.tblStaff SM ON SM.StaffIndex = S.ServManager
# 	INNER JOIN dbo.tblJob_Header JH ON JH.ContIndex = E.ContIndex AND JH.Job_Template = 262
# WHERE E.ClientStatus <> 'LOST' AND E.ContIndex < 900000)
# ORDER BY ContIndex ASC""")

# ActiveClients = getActiveClients.fetchall()

skippedClients = list()
MASskippedClients = list()

servurl = os.getenv('PE_URL')
appid = os.getenv('PE_APPID')
appkey = os.getenv('PE_APPKEY')

authurl = servurl + '/auth/connect/token'
auth = (appid, appkey)
authtype = {'grant_type': 'client_credentials', 'scope': 'pe.api'}

resptoken = requests.post(authurl, data=authtype, auth=auth)

token = resptoken.json()['access_token']

apiheader = {'Authorization': 'Bearer ' + token,
  'Content-Type': 'application/json'}

# for client in INDClients:
#     try:
#         output = requests.request('POST', servurl + '/pe/api/jobs/createtemplatejob', headers=apiheader, data=json.dumps({"JobTmp_Idx": 336,
#         "ContIndex": client[0],
#         "ServIndex": "INDTAX",
#         "Partner": client[3],
#         "Manager": client[4],
#         "Value": 1.00,
#         "PeriodStart": "2021-01-01T00:00:00.000Z",
#         "PeriodEnd": "2021-12-31T23:59:59.000Z",
#         "JobName": "2021 Tax Consulting",
#         "JobCode": "21-TxCon"}))
#         print(output.status_code)
#         if output.status_code == 200:
#           print(client[0])
#           print('\n')
#         else:
#           skippedClients.append(client[0])
#     except:
#         skippedClients.append(client[0])
#         continue

# # print('\n')

for client in BUSClients:
    try:
        output = requests.request('POST', servurl + '/pe/api/jobs/createtemplatejob', headers=apiheader, data=json.dumps({"JobTmp_Idx": 336,
        "ContIndex": client[0],
        "ServIndex": "BUSTAX",
        "Partner": client[3],
        "Manager": client[4],
        "Value": 1.00,
        "PeriodStart": "2021-01-01T00:00:00.000Z",
        "PeriodEnd": "2021-12-31T23:59:59.000Z",
        "JobName": "2021 Tax Consulting",
        "JobCode": "21-TxCon"}))
        print(output.status_code)
        if output.status_code == 200:
          print(client[0])
          print('\n')
        else:
          skippedClients.append(client[0])
    except:
        skippedClients.append(client[0])
        continue

print('\n')

# for client in ActiveClients:
#     try:
#         output = requests.request('POST', servurl + '/pe/api/jobs/createtemplatejob', headers=apiheader, data=json.dumps({"JobTmp_Idx": 262,
#         "ContIndex": client[0],
#         "ServIndex": "MAS",
#         "Partner": client[2],
#         "Manager": client[3],
#         "Value": 1.00,
#         "PeriodStart": "2021-01-01T00:00:00.000Z",
#         "PeriodEnd": "2021-12-31T23:59:59.000Z",
#         "JobName": "2021 Consulting",
#         "JobCode": "CON21"
#         }))
#         print(output.status_code)
#         if output.status_code == 200:
#           print(client[0])
#         else:
#           MASskippedClients.append(client[0])
#     except:
#         MASskippedClients.append(client[0])

# print('\n')
# print(skippedClients)
# print(MASskippedClients)
# print('\n')

# masJobs = list()

# for job in MASJobs:
#     if job[12] == 0.00:
#         output = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({"Job_Idx": job[0],
#         "ContIndex": job[1],
#         "Job_PreviousJob": job[2],
#         "Job_CurrentStaff": job[3],
#         "Job_Name": job[4],
#         "Job_Code": job[5],
#         "Job_Class": job[6],
#         "Job_Status": 3,
#         "Job_WorkStatus": job[8],
#         "Job_Office": job[9],
#         "Job_Dept": job[10],
#         "Job_MasterFile": job[11]}))

#         if output.status_code == 200:
#             print(job[0])
#             print('Original Status: ' + str(job[7]) + ' -> Final Status: 3')
#             print('\n')
#         else:
#             masJobs.append(job[0])
#     else:
#         output = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({"Job_Idx": job[0],
#         "ContIndex": job[1],
#         "Job_PreviousJob": job[2],
#         "Job_CurrentStaff": job[3],
#         "Job_Name": job[4],
#         "Job_Code": job[5],
#         "Job_Class": job[6],
#         "Job_Status": 2,
#         "Job_WorkStatus": job[8],
#         "Job_Office": job[9],
#         "Job_Dept": job[10],
#         "Job_MasterFile": job[11]}))

#         if output.status_code == 200:
#             print(job[0])
#             print('Original Status: ' + str(job[7]) + ' -> Final Status: 3')
#             print('\n')
#         else:
#             masJobs.append(job[0])

# print(masJobs)