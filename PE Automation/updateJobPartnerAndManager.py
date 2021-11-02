from dotenv import load_dotenv
import requests, os, pyodbc, json, time

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

def update():
    engine = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

    getBadJobs = engine.cursor()

    getBadJobs.execute("""SELECT [Job_Idx]
        ,JH.ContIndex
        ,e.ClientPartner
        ,[Job_Manager]
        ,[Job_InCharge]
        ,Job_Biller
        ,[Job_Frequency]
        ,[Job_Recurring]
        ,[Job_ETCRequired]
    FROM tblJob_Header JH
    inner join tblEngagement E on JH.ContIndex=E.ContIndex
    where job_partner=1 and Job_Status not in (-1,2,3,98)
    ORDER BY Job_Idx ASC""")

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
    'Content-Type': 'application/json'}

    badJobs = list()

    for job in BadJobs:
        output = requests.request('POST', servurl + '/pe/api/jobs/savemanagement', headers=apiheader, data=json.dumps({
            "Job_Idx": job[0],
            "Job_Partner": job[2],
            "Job_Manager": job[3],
            "Job_InCharge": job[4],
            "Job_Biller": 0,
            "Job_ETCRequired": True,
            "Job_Frequency": job[6],
            "Job_Recurring": job[7]
        }))

        if output.status_code == 200:
            print(job[0])
            print('\n')
        else:
            badJobs.append({'Job': job[0], 'Status': output.status_code, 'Reason': output.text})
        
        # time.sleep(5)

    print(badJobs)

try:
    update()
except:
    update()
finally:
    update()