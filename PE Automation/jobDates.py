from dotenv import load_dotenv
import requests, os, pyodbc, json

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getJobsToUpdate = engine.cursor()

getJobsToUpdate.execute("""SELECT j.job_idx
FROM tblEngagement AS E
INNER JOIN tblJob_Header  AS J    ON j.ContIndex = e.ContIndex 
WHERE e.ClientStatus NOT IN ('LOST', 'INTERNAL') AND j.Job_Status not in (2,3) and j.Job_Template = 75 and j.Job_Period_End >'12/31/2021'
ORDER BY j.Job_Idx ASC""")

jobs = getJobsToUpdate.fetchall()

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

skipped = list()

for job in jobs:
    output = requests.request('POST', servurl + '/pe/api/jobs/savedates', headers=apiheader, data=json.dumps({
        "Job_Idx": job[0],
        "Job_Period_Start": "2021-10-01",
        "Job_Period_End": "2022-09-30"
    }))

    if output.status_code != 200:
        skipped.append({'job': job[0], 'status': output.status_code, 'reason': output.text})

print(skipped)