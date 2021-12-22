from dotenv import load_dotenv
import pandas as pd
import json, requests, os, pyodbc

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

bmss = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + os.getenv('DB_SERVER') + ';DATABASE=' + os.getenv('DB_DATABASE') + ';UID=' + os.getenv('DB_USER') + ';PWD=' + os.getenv('DB_PASS') + ';Authentication=ActiveDirectoryPassword')

df = pd.read_sql("""SELECT E.* 
            FROM dbo.tblEngagement E
                INNER JOIN dbo.tblContacts C ON C.ContIndex = E.ContIndex AND C.ContType = 2
            WHERE ClientStatus <> 'LOST' 
            AND E.ContIndex NOT IN (SELECT ContIndex FROM dbo.tblClientServices WHERE ServIndex = 'TAS')""", bmss)

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
    try:
        output = requests.request('post', servurl + '/pe/api/ClientServices/EngageService', headers=apiheader, data=json.dumps({
            "ContIndex": int(row['ContIndex']),
            "ServIndex": "TAS",
            "Partner": int(row['ClientPartner']),
            "Manager": int(row['ClientManager']),
            "InCharge": int(row['ClientInCharge'])
        }))

        if output.status_code != 200:
            skipped.append({'client': row['ContIndex'], 'status': output.status_code, 'reason': output.text})
    except:
        skipped.append({'client': row['ContIndex'], 'reason': 'except block'})
        continue

print(skipped)
