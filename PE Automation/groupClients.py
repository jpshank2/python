from dotenv import load_dotenv
import requests, os, time, json, datetime
import pandas as pd

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

groupFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\Clients to Group.xlsx'

groupSheet = pd.read_excel(groupFile, 'Sheet1')

groupSheet = groupSheet[groupSheet['New Group Client Code'].notnull()]

badGroups = list()

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

tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

for x in range(groupSheet.shape[0]):
    row = groupSheet.iloc[x]
    client = requests.request('GET', servurl + '/pe/api/clients/loadclient/' + str(row['Index']), headers=apiheader)
    client = client.json()

    client['ClientHold'] = int(row['Column1'])

    updateClientGroup = requests.request('POST', servurl + '/pe/api/clients/save', headers=apiheader, data=json.dumps(client))
    
    if updateClientGroup.status_code == 200:
        print(f"Updated client {row['Index']}")
    else:
        badGroups.append({'Client': row['Contindex'], 'Status': updateClientGroup.status_code, 'Reason': updateClientGroup.text})

    time.sleep(2)

    if datetime.datetime.now() > tokenExpiry:
        resptoken = requests.post(authurl, data=authtype, auth=auth)

        token = resptoken.json()['access_token']

        apiheader = {'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'}

        tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

print(badGroups)
