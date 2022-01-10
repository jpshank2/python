from dotenv import load_dotenv
import requests, os, pyodbc, time, json, datetime

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')
engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

getBadJobs = engine.cursor()

getBadJobs.execute("""SELECT TOP 930 jh.Job_Idx
      ,JH.ContIndex
	  ,(case when CS.ServPartner = 0 then e.ClientPartner else cs.ServPartner end) as [Partner]
	  ,(case when cs.ServManager=0 then e.ClientManager else cs.ServManager end) as [Manager]
      ,jh.Job_InCharge
      ,jh.Job_Biller
      ,jh.Job_Frequency
      ,jh.Job_Recurring
      ,jh.Job_ETCRequired
  FROM tblJob_Header JH
  inner join tblJob_Serv JS on JH.Job_Idx=JS.Job_Idx
  inner join tblClientServices CS on JH.ContIndex=CS.ContIndex and JS.ServIndex=CS.ServIndex
  inner join tblEngagement E on JH.ContIndex=E.ContIndex
  where (job_partner=0 or job_manager=0) and Job_Status not in (-1,2,3,98) AND JH.Job_Template IN (336, 262)
  ORDER BY Job_Idx DESC""")

BadJobs = getBadJobs.fetchall()

servurl = os.getenv('PE_URL')
appid = os.getenv('PE_APPID')
appkey = os.getenv('PE_APPKEY')

authurl = servurl + '/auth/connect/token'
auth = (appid, appkey)
authtype = {'grant_type': 'client_credentials', 'scope': 'pe.api'}

resptoken = requests.post(authurl, data=authtype, auth=auth)

token = resptoken.json()['access_token']

apiheader = {'Authorization': 'Bearer ' + token,
  'Content-Type': 'application/json', 
  'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64)' 'AppleWebKit/537.36 (KHTML, like Gecko)' 'Chrome/96.0.3497.100 Safari/537.36')}

badJobs = list()

tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

for job in BadJobs:
    output = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
        "Job_Idx": job[0],
        "Job_Partner": job[2],
        "Job_Manager": job[3],
        "Job_InCharge": job[4],
        "Job_Biller": 0,
        "Job_ETCRequired": True,
        "Job_Frequency": 1,
        "Job_Recurring": 1
    }))

    if output.status_code == 200:
        print(job[0])
        print('\n')
    else:
        print(output.text)
        badJobs.append({'Job': job[0], 'Status': output.status_code, 'Reason': output.text})
    
    time.sleep(5)

    if datetime.datetime.now() > tokenExpiry:
        resptoken = requests.post(authurl, data=authtype, auth=auth)

        token = resptoken.json()['access_token']

        apiheader = {'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'}

        tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

print(badJobs)
