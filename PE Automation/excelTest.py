from dotenv import load_dotenv
import pandas as pd
import json, requests, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

ogFile = r'c:\users\jeremyshank\desktop\control sheet.xlsx'
changeFile = r'c:\users\jeremyshank\desktop\variable sheet.xlsx'

df1 = pd.read_excel(ogFile, 'Sheet1')
df2 = pd.read_excel(changeFile, 'Sheet1')

# print(df1, df2, sep='\n\n')

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

for x in range(0, df1.shape[0]):
    row1 = df1.iloc[x]
    row2 = df2.iloc[x]
    if not row1.equals(row2):
        if pd.isnull(row2.iloc(4)):
            # do update scenario
            updateManagement = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
                "Job_Idx": row2[0],
                "Job_Partner": row2,
                "Job_Manager": row2,
                "Job_ETCRequired": True,
                "Job_Frequency": row2,
                "PercentComplete": row2,
                "Job_Recurring": row2}))
            updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
                "Job_Idx": row2[0],
                "ContIndex": row2[1],
                "Job_PreviousJob": row2[2],
                "Job_CurrentStaff": row2[3],
                "Job_Name": row2[4],
                "Job_Code": row2[5],
                "Job_Status": row2,
                "Job_WorkStatus": row2[8],
                "Job_Office": row2[9],
                "Job_Dept": row2[10],
                "Job_MasterFile": row2[11]}))
            print(row2)
        else:
            # set to not recurring
            updateManagement = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
                "Job_Idx": row2[0],
                "Job_Recurring": False}))
            # close job
            updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
                "Job_Idx": row2[0],
                "ContIndex": row2[1],
                "Job_PreviousJob": row2[2],
                "Job_CurrentStaff": row2[3],
                "Job_Name": row2[4],
                "Job_Code": row2[5],
                "Job_Class": row2[6],
                "Job_Status": 2,
                "Job_WorkStatus": row2[8],
                "Job_Office": row2[9],
                "Job_Dept": row2[10],
                "Job_MasterFile": row2[11]}))
