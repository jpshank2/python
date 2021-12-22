from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

planningJobs = pd.read_sql("""SELECT JH.Job_Idx
	  ,JS.ServIndex
      ,JH.ContIndex
      ,JH.Job_Name
      ,JH.Job_Dept
      ,JH.Job_Partner
	  ,JH.Job_Manager
      ,JH.Job_InCharge
      ,JH.Job_Class
      ,JH.Job_Status
      ,JH.Job_Frequency
      ,JH.Job_Recurring
      ,JH.Job_Period_Start
      ,JH.Job_Period_End
      ,JH.Job_Finish_Deadline
      ,JH.Job_Budget_Hours
      ,JH.Job_Budget_Value
      ,JH.Job_Code
      ,JH.Job_Value
      ,JH.Job_ETCRequired
      ,JH.Job_PreviousJob
      ,JH.Job_NextJob
      ,JH.Job_LimitTimeEntry
      ,JH.Job_Complexity
      ,JH.Job_MasterFile
      ,JH.Job_Office
      ,JH.Job_WorkStatus
      ,JH.Job_CurrentStaff
      ,JH.Job_StaffingType
      ,JH.Job_Biller
      ,JH.PercentComplete
  FROM tblJob_Header JH
  inner join tblJob_Serv JS on JH.Job_Idx=JS.Job_Idx
  WHERE Job_Template=101 and Job_NextJob=0 and Job_Period_End > '12/31/2020'
  ORDER BY JH.Job_Idx
  --OFFSET 3 ROWS
  --FETCH NEXT 1 ROW ONLY
""", bmss)

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

skipped = list()

for x in range(0, planningJobs.shape[0]):
    row = planningJobs.iloc[x]
    try:
        # turn recurring off and close job
        updateOldJob = requests.request('post', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
            "Job_Idx": int(row['Job_Idx']),
            "Job_NextJob": int(row['Job_NextJob']),
            "Job_Partner": int(row['Job_Partner']),
            "Job_Manager": int(row['Job_Manager']),
            "Job_InCharge": int(row['Job_InCharge']),
            "Job_Biller": 0 if str(row['Job_Biller']) == 'nan' else int(row['Job_Biller']),
            "Job_Value": int(row['Job_Value']),
            "Job_Complexity": int(row['Job_Complexity']),
            "Job_StaffingType": int(row['Job_StaffingType']),
            "Job_LimitTimeEntry": int(row['Job_LimitTimeEntry']),
            "Job_ETCRequired": True,
            "Job_Frequency": 0,
            "Job_Recurring": False,
            "PercentComplete": int(row['PercentComplete'])
        }))

        # print(updateOldJob.json())

        if updateOldJob.status_code != 200:
            skipped.append({'stage': 'Updating old job', 'client': row['ContIndex'], 'status': updateOldJob.status_code, 'reason': updateOldJob.text})

        completeOldJob = requests.request('post', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
            "Job_Idx": int(row['Job_Idx']),
            "ContIndex": int(row['ContIndex']),
            "Job_CurrentStaff": 0,
            "Job_Name": str(row['Job_Name']),
            "Job_Code": str(row['Job_Code']),
            "Job_Class": str(row['Job_Class']),
            "Job_Status": 2,
            "Job_WorkStatus": int(row['Job_WorkStatus']),
            "Job_Dept": "UNKNOWN" if row['Job_Dept'] == '' else str(row['Job_Dept']),
            "Job_Office": "UNKNOWN" if row['Job_Office'] == '' else str(row['Job_Office']),
            "Job_MasterFile": "" if row['Job_MasterFile'] == '' or row['Job_MasterFile'] is None else row['Job_MasterFile']
        }))

        # print(completeOldJob.json())

        if completeOldJob.status_code != 200:
            skipped.append({'stage': 'Completing old job', 'client': row['ContIndex'], 'status': completeOldJob.status_code, 'reason': completeOldJob.text})
            continue

        #create new tax planning job
        # print(str(row['Job_Period_Start']))
        newJob = requests.request('post', servurl + '/pe/api/jobs/createtemplatejob', headers=apiheader, data=json.dumps({
            "JobTmp_Idx": 369,
            "ContIndex": int(row['ContIndex']),
            "ServIndex": str(row['ServIndex']),
            "Partner": int(row['Job_Partner']),
            "Manager": int(row['Job_Manager']),
            "Value": 1.00,
            "PeriodStart": str(row['Job_Period_Start']),
            "PeriodEnd": str(row['Job_Period_End']),
            "FiscalYear": row['Job_Period_Start'].year,
            "JobName": str(row['Job_Period_Start'].year) + "-Tax Planning",
            "JobCode": str(row['Job_Period_Start'].year)[-2:] + "-TXPLN"
        }))

        # print(newJob, newJob.json(), sep='\n')#.json())

        if newJob.status_code != 200:
            skipped.append({'stage': 'Creating new job', 'client': row['ContIndex'], 'status': newJob.status_code, 'reason': newJob.text})
            continue

        newJobIndex = pd.read_sql("""SELECT Job_Idx FROM dbo.tblJob_Header WHERE Job_Template = 369 AND Job_Updated_By = 'jeremyshank@bmss.com' AND ContIndex = """ + str(row['ContIndex']), bmss)

        # move job notes to new job
        getNotes = requests.request('get', servurl + '/pe/api/jobs/gethistory/' + str(row['Job_Idx']), headers=apiheader)

        allNotes = getNotes.json()

        for note in allNotes:
            if note['Type'] == 'Note':
                moveNote = requests.request('post', servurl + '/pe/api/jobs/updatenote', headers=apiheader, data=json.dumps({
                    "JobNote_Idx": int(note['Idx']),
                    "Job_Idx": int(newJobIndex.iloc[0]['Job_Idx']),
                    "NoteDate": str(note['Date']),
                    "NoteText": str(note['Text']),
                    "DeletedOn": "" if note['DeletedOn'] is None else str(note['DeletedOn']),
                    "DeletedBy": "" if note['DeletedBy'] is None else str(note['DeletedBy']),
                    "RowVer": ""
                }))

                # print(moveNote.json())

                if moveNote.status_code != 200:
                    skipped.append({'stage': 'Moving Job Note', 'client': row['ContIndex'], 'status': moveNote.status_code, 'reason': moveNote.text})
        
        print(str(row['Job_Idx']) + ' -> ' + str(newJobIndex.iloc[0]['Job_Idx']) + '\n')
    except:
        skipped.append({'client': row['ContIndex'], 'reason': 'except block'})
        continue

print(skipped)
