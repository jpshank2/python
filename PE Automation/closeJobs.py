from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

closeFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\JobDelete.xlsx'
bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

df = pd.read_excel(closeFile, 'Sheet1')

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

for x in range(0, 1):#df.shape[0]):
    row = df.iloc[x]
    dfSQL = pd.read_sql("""SELECT Job_Partner, 
        Job_Manager, 
        ContIndex,
        Job_Name,
        Job_Code,
        Job_Class,
        Job_Office 
    FROM dbo.tblJob_Header WHERE Job_Idx = 199087""", bmss)# + str(row['Job_Idx']), bmss)
    print('yes' if dfSQL.iloc[0]['Job_Office'] == '' else 'no' )
    # try:
    #     print(row['Job_Idx'])
    #     updateManagement = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
    #         "Job_Idx": int(row['Job_Idx']),
    #         "Job_Partner": int(dfSQL.iloc[0]['Job_Partner']),
    #         "Job_Manager": int(dfSQL.iloc[0]['Job_Manager']),
    #         "Job_Recurring": False
    #     }))

    #     if updateManagement.status_code != 200:
    #         skipped.append({'job': row['Job_Idx'], 'status': updateManagement.status_code, 'reason': updateManagement.text, 'point_of_error': 'updateManagement'})

    #     # close job
    #     updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
    #         "Job_Idx": int(row['Job_Idx']),
    #         "ContIndex": int(dfSQL.iloc[0]['ContIndex']),
    #         "Job_CurrentStaff": 0,
    #         "Job_Name": str(dfSQL.iloc[0]['Job_Name']),
    #         "Job_Code": str(dfSQL.iloc[0]['Job_Code']),
    #         "Job_Class": str(dfSQL.iloc[0]['Job_Class']),
    #         "Job_Status": 2,
    #         "Job_WorkStatus": 28,
    #         "Job_Dept": "UNKNOWN",
    #         "Job_Office": "UNKNOWN" if dfSQL.iloc[0]['Job_Office'] == '' else str(dfSQL.iloc[0]['Job_Office']),
    #         "Job_Masterfile": ""
    #     }))

    #     if updateDetails.status_code != 200:
    #         skipped.append({'job': row['Job_Idx'], 'status': updateDetails.status_code, 'reason': updateDetails.text, 'point_of_error': 'updateDetails'})

    # except:
    #     skipped.append({'job': row['Job_Idx'], 'point_of_error': 'except block'})
    #     continue

print(skipped)
