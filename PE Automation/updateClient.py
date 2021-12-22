from dotenv import load_dotenv
import json, requests, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

servurl = os.getenv('PE_URL')
appid = os.getenv('PE_APPID')
appkey = os.getenv('PE_APPKEY')

authurl = servurl + '/auth/connect/token'
auth = (appid, appkey)
authtype = {'grant_type': 'client_credentials', 'scope': 'pe.api'}

# resptoken = requests.post(authurl, data=authtype, auth=auth)

token = resptoken.json()['access_token']

apiheader = {'Authorization': 'Bearer ' + token,
  'Content-Type': 'application/json'}

skipped = list()

clientDetails = requests.request('get', servurl + '/pe/api/Clients/LoadClient/27953', headers=apiheader)

clientDetails = clientDetails.json()

clientDetails['ClientName'] = 'RLB Test Organization for Jobs, LLC'

payload = clientDetails

# print(clientDetails)

# print(skipped)

updateClient = requests.request('post', servurl + '/pe/api/Clients/Save', headers=apiheader, data=json.dumps(payload))

print(updateClient.status_code)
