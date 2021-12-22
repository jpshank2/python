from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

df = [{
    'job': 199087,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 193681,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 203692,
    'point_of_error': 'except block'
}, {
    'job': 203691,
    'point_of_error': 'except block'
}, {
    'job': 203690,
    'point_of_error': 'except block'
}, {
    'job': 203679,
    'point_of_error': 'except block'
}, {
    'job': 203681,
    'point_of_error': 'except block'
}, {
    'job': 226657,
    'point_of_error': 'except block'
}, {
    'job': 176182,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 180535,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 180536,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 180537,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 175419,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 69378,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 206795,
    'point_of_error': 'except block'
}, {
    'job': 206796,
    'point_of_error': 'except block'
}, {
    'job': 206797,
    'point_of_error': 'except block'
}, {
    'job': 206798,
    'point_of_error': 'except block'
}, {
    'job': 206799,
    'point_of_error': 'except block'
}, {
    'job': 123212,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 143114,
    'point_of_error': 'except block'
}, {
    'job': 227589,
    'status': 400,
    'reason':
    '{"Content":"[{\\"type\\":\\"dialog\\",\\"message\\":\\"The Job could not be set to Complete. The Job has unposted WIP.\\",\\"level\\":\\"warning\\"}]","ContentType":"application/json","StatusCode":400}',
    'point_of_error': 'updateDetails'
}, {
    'job': 176392,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 176372,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 176373,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 176374,
    'status': 400,
    'reason': '["The Job_Office field is required."]',
    'point_of_error': 'updateDetails'
}, {
    'job': 200122,
    'status': 400,
    'reason':
    '{"Content":"[{\\"type\\":\\"dialog\\",\\"message\\":\\"The Job could not be set to Complete. The Job has unposted WIP.\\",\\"level\\":\\"warning\\"}]","ContentType":"application/json","StatusCode":400}',
    'point_of_error': 'updateDetails'
}]

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
    dfSQL = pd.read_sql("""SELECT Job_Partner, 
        Job_Manager, 
        ContIndex,
        Job_Name,
        Job_Code,
        Job_Class,
        Job_Office 
    FROM dbo.tblJob_Header WHERE Job_Idx = """ + str(job['job']), bmss)
    # print(dfSQL)
    try:
        print(job['job'])
        updateManagement = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
            "Job_Idx": int(job['job']),
            "Job_Partner": int(dfSQL.iloc[0]['Job_Partner']),
            "Job_Manager": int(dfSQL.iloc[0]['Job_Manager']),
            "Job_Recurring": False
        }))

        if updateManagement.status_code != 200:
            skipped.append({'job': job['job'], 'status': updateManagement.status_code, 'reason': updateManagement.text, 'point_of_error': 'updateManagement'})

        # close job
        updateDetails = requests.request('POST', servurl + '/pe/api/jobs/savedetails', headers=apiheader, data=json.dumps({
            "Job_Idx": int(job['job']),
            "ContIndex": int(dfSQL.iloc[0]['ContIndex']),
            "Job_CurrentStaff": 0,
            "Job_Name": str(dfSQL.iloc[0]['Job_Name']),
            "Job_Code": str(dfSQL.iloc[0]['Job_Code']),
            "Job_Class": str(dfSQL.iloc[0]['Job_Class']),
            "Job_Status": 2,
            "Job_WorkStatus": 28,
            "Job_Dept": "UNKNOWN",
            "Job_Office": "UNKNOWN" if dfSQL.iloc[0]['Job_Office'] == '' else str(dfSQL.iloc[0]['Job_Office']),
            "Job_Masterfile": ""
        }))

        if updateDetails.status_code != 200:
            skipped.append({'job': job['job'], 'status': updateDetails.status_code, 'reason': updateDetails.text, 'point_of_error': 'updateDetails'})

    except:
        skipped.append({'job': job['job'], 'point_of_error': 'except block'})
        continue

print(skipped)
