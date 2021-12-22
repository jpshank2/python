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

staffDetails = requests.request('get', servurl + '/pe/api/StaffAdmin/GetStaffFullDetails/219', headers=apiheader)

staffDetails = staffDetails.json()

staffDetails['staffDetails']['StaffName'] = 'David Brown'

payload = staffDetails

# print(staffDetails)

# print(skipped)

updateStaff = requests.request('post', servurl + '/pe/api/staffadmin/savestafffulldetails', headers=apiheader, data=json.dumps(payload))

print(updateStaff.status_code)
