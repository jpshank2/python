from dotenv import load_dotenv
import pandas as pd
import json, requests, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

closeFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\CAAS Job Delete Template.xlsx'

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

for x in range(0, df.shape[0]):
    row = df.iloc[x]
    if not pd.isnull(row.iloc[24]):
        try:
            print(row.loc['job_idx'])
            updateManagement = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
                "Job_Idx": int(row.loc['job_idx']),
                "Job_Partner": int(row.loc['PartnerIndex']),
                "Job_Manager": int(row.loc['ManagerIndex']),
                "Job_Recurring": False
            }))
            print(updateManagement.reason)
            # close job
            updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
                "Job_Idx": int(row.loc['job_idx']),
                "ContIndex": int(row.loc['ContIndex']),
                "Job_CurrentStaff": 0,
                "Job_Name": str(row.loc['Job Name']),
                "Job_Code": str(row.loc['Job_Code']),
                "Job_Class": str(row.loc['Job_Class']),
                "Job_Status": 2,
                "Job_WorkStatus": 28,
                "Job_Dept": "UNKNOWN",
                "Job_Office": "0",
                "Job_Masterfile": ""
            }))
            print(updateDetails.reason)
        except:
            skipped.append(row.loc['job_idx'])
            continue

print(skipped)
