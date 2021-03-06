from dotenv import load_dotenv
import requests, os, time, json, datetime
import pandas as pd

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

clientFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\ClientUpdate.xlsx'

clientSheet = pd.read_excel(clientFile, 'ClientUpdate')

badClients = list()

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

for x in range(clientSheet.shape[0]):
    row = clientSheet.iloc[x]
    client = requests.request('GET', servurl + '/pe/api/clients/loadclient/' + str(row['Contindex']), headers=apiheader)
    client = client.json()

    del client['ClientPartner']
    del client['ClientManager']

    client['ClientPartnerIndex'] = int(row['CP'])
    client['ClientManagerIndex'] = int(row['CM'])

    updatedClient = requests.request('POST', servurl + '/pe/api/clients/save', headers=apiheader, data=json.dumps(client))
    
    if updatedClient.status_code == 200:
        print(f"Updated Client {client['ClientName']}")
    else:
        badClients.append({'Client': row['Contindex'], 'Status': updatedClient.status_code, 'Reason': updatedClient.text})

    time.sleep(5)

    if datetime.datetime.now() > tokenExpiry:
        resptoken = requests.post(authurl, data=authtype, auth=auth)

        token = resptoken.json()['access_token']

        apiheader = {'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'}

        tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

print(badClients)
