from dotenv import load_dotenv
import requests, os, pyodbc, json

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

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

output = requests.request('GET', servurl + '/pe/api/staffmember/whois/219', headers=apiheader)

print(output.json())