from dotenv import load_dotenv
import requests, json, os

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

servurl = os.getenv('PE_URL')
authurl = servurl + '/auth/connect/token'
apiurl = servurl + '/pe/api/Clients/loadclient/2101'
appid = os.getenv('PE_APPID')
appkey = os.getenv('PE_APPKEY')

auth = (appid,appkey)
authtype = {'grant_type': 'client_credentials', 'scope': 'pe.api'}
resptoken = requests.post(authurl, data=authtype, auth=auth)
token = resptoken.json()['access_token']
apiheader = {'Authorization': 'Bearer ' + token}

meredith = requests.get(apiurl, headers=apiheader).json()

print(meredith)
