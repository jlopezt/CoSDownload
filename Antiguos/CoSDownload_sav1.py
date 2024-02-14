import requests

url = 'https://siopmgr.totemtowers.es/tsp/'
#https://siopmgr.totemtowers.es/tsp/#/identity/login
#myobj = {'somekey': 'somevalue'}

#x = requests.post(url, json = myobj)
x = requests.get(url, verify=False)

print(x.text)

