from dotenv import load_dotenv
import requests, os, time, json, datetime
import pandas as pd

load_dotenv(os.path.dirname(os.path.dirname(__file__)) + '\\.env')

serviceFile = r'C:\Users\jeremyshank\BMSS\Business Intelligence - Documents (1)\Automation Tools\ServiceUpdate.xlsx'

serviceSheet = pd.read_excel(serviceFile, 'ServiceUpdate')

badServices = list()

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

for x in range(serviceSheet.shape[0]):
    row = serviceSheet.iloc[x]
    service = requests.request('GET', f"{servurl}/pe/api/clients/getclientservicedetails/{str(row['Contindex'])}?service={row['ServIdx']}", headers=apiheader)
    service = service.json()
    service = service['ServiceDetail']

    del service['PartName']
    del service['ManName']

    service['ServPartner'] = int(row['SP'])
    service['ServManager'] = int(row['SM'])

    updateService = requests.request('POST', servurl + '/pe/api/clients/updateclientservice', headers=apiheader, data=json.dumps(service))
    
    if updateService.status_code == 200:
        print(f"Updated Service {service['ServTitle']} for client {row['Contindex']}")
    else:
        badServices.append({'Client': row['Contindex'], 'Status': updateService.status_code, 'Reason': updateService.text})

    time.sleep(5)

    if datetime.datetime.now() > tokenExpiry:
        resptoken = requests.post(authurl, data=authtype, auth=auth)

        token = resptoken.json()['access_token']

        apiheader = {'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'}

        tokenExpiry = datetime.datetime.now() + datetime.timedelta(minutes=55)

print(badServices)
