from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

df = pd.read_excel(r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\Caryn update.xlsx')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

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

skipped_jobs = list()

for x in range(df.shape[0]):
    row = df.iloc[x]
    dfSQL = pd.read_sql("""select
        jh.Job_Value,
        jh.Job_InCharge,
        jh.Job_Biller,
        jh.Job_Complexity,
        jh.Job_StaffingType,
        jh.Job_LimitTimeEntry,
        jh.Job_ETCRequired,
        jh.Job_Recurring,
        jh.PercentComplete,
        jh.Job_NextJob,
        jh.Job_Frequency
        from tbljob_header JH
        WHERE Job_Idx = """ + str(row['JobIdx']), bmss)
    # print(dfSQL.iloc[x])
    output = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
        "Job_Idx": int(row['JobIdx']),
        "Job_Partner": int(row['PtnrIdx']),
        "Job_Manager": int(row['MgrIdx']),
        "Job_NextJob": int(dfSQL.iloc[0]['Job_NextJob']),
        "Job_InCharge": 0 if dfSQL.iloc[0]['Job_InCharge'] is None else int(dfSQL.iloc[0]['Job_InCharge']),
        "Job_Biller": 0 if dfSQL.iloc[0]['Job_Biller'] is None else int(dfSQL.iloc[0]['Job_Biller']),
        "Job_Value": dfSQL.iloc[0]['Job_Value'],
        "Job_Complexity": int(dfSQL.iloc[0]['Job_Complexity']),
        "Job_StaffingType": int(dfSQL.iloc[0]['Job_StaffingType']),
        "Job_LimitTimeEntry": bool(dfSQL.iloc[0]['Job_LimitTimeEntry']),
        "Job_ETCRequired": True,
        "Job_Frequency": int(dfSQL.iloc[0]['Job_Frequency']),
        "Job_Recurring": bool(dfSQL.iloc[0]['Job_Recurring']),
        "PercentComplete": dfSQL.iloc[0]['PercentComplete']
    }))
    print(str(row['JobIdx']) + ' -> ' + str(output.status_code))
    if output.status_code != 200:
        skipped_jobs.append({'job': row['JobIdx'], 'status': output.status_code, 'reason': output.text})

print(skipped_jobs)
