from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

closeFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\New and Old Job Idx from Tax Planning.xlsx'
bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

excelSheet = pd.read_excel(closeFile, 'Table2')

df = [{'job': 234011, 'point_of_error': 'except block'}, {'job': 234012, 'point_of_error': 'except block'}, {'job': 234013, 'point_of_error': 'except block'}, {'job': 234014, 'point_of_error': 'except block'}, {'job': 234015, 'point_of_error': 'except block'}, {'job': 234017, 'point_of_error': 'except block'}, {'job': 234018, 'point_of_error': 'except block'}, {'job': 234019, 'point_of_error': 'except block'}, {'job': 234021, 'point_of_error': 'except block'}, {'job': 234022, 'point_of_error': 'except block'}, {'job': 234023, 'point_of_error': 'except block'}, {'job': 234024, 'point_of_error': 'except block'}, {'job': 234025, 'point_of_error': 'except block'}, {'job': 234026, 'point_of_error': 'except block'}, {'job': 234027, 'point_of_error': 'except block'}, {'job': 234028, 'point_of_error': 'except block'}, {'job': 234029, 'point_of_error': 'except block'}, {'job': 234030, 'point_of_error': 'except block'}, {'job': 234031, 'point_of_error': 'except block'}, {'job': 234032, 'point_of_error': 'except block'}, {'job': 234033, 'point_of_error': 'except block'}, {'job': 234034, 'point_of_error': 'except block'}, {'job': 234035, 'point_of_error': 'except block'}, {'job': 234036, 'point_of_error': 'except block'}, {'job': 234037, 'point_of_error': 'except block'}, {'job': 234038, 'point_of_error': 'except block'}, {'job': 234039, 'point_of_error': 'except block'}, {'job': 234040, 'point_of_error': 'except block'}, {'job': 234041, 'point_of_error': 'except block'}, {'job': 234042, 'point_of_error': 'except block'}, {'job': 234043, 'point_of_error': 'except block'}, {'job': 234044, 'point_of_error': 'except block'}, {'job': 234045, 'point_of_error': 'except block'}, {'job': 234046, 'point_of_error': 'except block'}, {'job': 234047, 'point_of_error': 'except block'}, {'job': 234048, 'point_of_error': 'except block'}, {'job': 234049, 'point_of_error': 'except block'}, {'job': 234050, 'point_of_error': 'except block'}, {'job': 234051, 'point_of_error': 'except block'}, {'job': 234052, 'point_of_error': 'except block'}, {'job': 234053, 'point_of_error': 'except block'}, {'job': 234054, 'point_of_error': 'except block'}, {'job': 234055, 'point_of_error': 'except block'}, {'job': 234056, 'point_of_error': 'except block'}, {'job': 234057, 'point_of_error': 'except block'}, {'job': 234058, 'point_of_error': 'except block'}, {'job': 234059, 'point_of_error': 'except block'}, {'job': 234060, 'point_of_error': 'except block'}, {'job': 234061, 'point_of_error': 'except block'}, {'job': 234062, 'point_of_error': 'except block'}, {'job': 234063, 'point_of_error': 'except block'}, {'job': 234064, 'point_of_error': 'except block'}, {'job': 234065, 'point_of_error': 'except block'}, {'job': 234066, 'point_of_error': 'except block'}, {'job': 234067, 'point_of_error': 'except block'}, {'job': 234068, 'point_of_error': 'except block'}, {'job': 234069, 'point_of_error': 'except block'}, {'job': 234070, 'point_of_error': 'except block'}, {'job': 234071, 'point_of_error': 'except block'}, {'job': 234072, 'point_of_error': 'except block'}, {'job': 234073, 'point_of_error': 'except block'}, {'job': 234075, 'point_of_error': 'except block'}, {'job': 234076, 'point_of_error': 'except block'}, {'job': 234078, 'point_of_error': 'except block'}, {'job': 234079, 'point_of_error': 'except block'}, {'job': 234080, 'point_of_error': 'except block'}, {'job': 234081, 'point_of_error': 'except block'}, {'job': 234082, 'point_of_error': 'except block'}, {'job': 234083, 'point_of_error': 'except block'}, {'job': 234084, 'point_of_error': 'except block'}, {'job': 234085, 'point_of_error': 'except block'}, {'job': 234086, 'point_of_error': 'except block'}, {'job': 234087, 'point_of_error': 'except block'}, {'job': 234088, 'point_of_error': 'except block'}, {'job': 234089, 'point_of_error': 'except block'}, {'job': 234090, 'point_of_error': 'except block'}, {'job': 234091, 'point_of_error': 'except block'}, {'job': 234092, 'point_of_error': 'except block'}, {'job': 234093, 'point_of_error': 'except block'}, {'job': 234094, 'point_of_error': 'except block'}, {'job': 234095, 'point_of_error': 'except block'}, {'job': 234096, 'point_of_error': 'except block'}, {'job': 234097, 'point_of_error': 'except block'}, {'job': 233743, 'point_of_error': 'except block'}]

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

for job in df:
    oldIndex = excelSheet[excelSheet['New Job'] == job['job']].iloc[0]['Old Job']
    print(oldIndex)
    dfSQL = pd.read_sql("""SELECT JHA.Job_Status
      ,JHA.Job_WorkStatus
      ,JHA.Job_CurrentStaff
FROM tblJob_Header_Audit JHA
where JHA.Job_Updated < '12/22/2021' and JHA.Job_Idx = """ + str(oldIndex) +"""
order by JHA.Job_Updated DESC
 """, bmss)
    # print(dfSQL)
    try:
        newJobStatus = pd.read_sql("""SELECT Job_Idx, 
ContIndex,
Job_CurrentStaff,
Job_Name,
Job_Code,
Job_Class,
Job_Status,
Job_WorkStatus,
Job_Dept,
Job_Office,
Job_MasterFile
FROM dbo.tblJob_Header WHERE Job_Idx = """ + str(job['job']), bmss)

        updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
            "Job_Idx": int(job['job']),
            "ContIndex": int(newJobStatus.iloc[0]['ContIndex']),
            "Job_CurrentStaff": int(dfSQL.iloc[0]['Job_CurrentStaff']),
            "Job_Name": str(newJobStatus.iloc[0]['Job_Name']),
            "Job_Code": str(newJobStatus.iloc[0]['Job_Code']),
            "Job_Class": str(newJobStatus.iloc[0]['Job_Class']),
            "Job_Status": int(dfSQL.iloc[0]['Job_Status']),
            "Job_WorkStatus": int(dfSQL.iloc[0]['Job_WorkStatus']),
            "Job_Dept": str(newJobStatus.iloc[0]['Job_Dept']),
            "Job_Office": str(newJobStatus.iloc[0]['Job_Office']),
            "Job_Masterfile": ""
        }))

        if updateDetails.status_code != 200:
            skipped.append({'job': job['job'], 'status': updateDetails.status_code, 'reason': updateDetails.text, 'point_of_error': 'updateDetails'})
            continue
        else:
            print('Fixed Job: ' + str(job['job']))

    except:
        skipped.append({'job': job['job'], 'point_of_error': 'except block'})
        continue

print(skipped)
