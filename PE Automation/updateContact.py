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

for contact in [35551, 35553, 35555]:
    contactDetails = requests.request('get', servurl + '/pe/api/Contacts/LoadContact/' + str(contact), headers=apiheader)

    contactDetails = contactDetails.json()

    contactDetails['contactDetails']['ContName'] = 'Tim Neal'

    payload = contactDetails

    # print(skipped)

    updateContact = requests.request('post', servurl + '/pe/api/Contacts/Save', headers=apiheader, data=json.dumps(payload))

    print(updateContact.status_code)
