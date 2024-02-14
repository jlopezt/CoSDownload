import requests

# replacement strings
WINDOWS_LINE_ENDING = '\r\n'
UNIX_LINE_ENDING = '\n'

#/TSP/ no se envia nada
#init se envia cookie: : CSRF-TOKEN=ec4ddf10-e97e-4e79-bb28-f60daf106f88
#https://siopmgr.totemtowers.es/tsp/assets/i18n/es-ES-v2.json : CSRF-TOKEN=ec4ddf10-e97e-4e79-bb28-f60daf106f88
#https://siopmgr.totemtowers.es/tsp./assets/i18n/es-ES.json  : CSRF-TOKEN=e8724de2-0b85-4911-9262-8397b7ea31f6
#cookie:
#    nombre: CSRF-TOKEN
#    Valor:  e8724de2-0b85-4911-9262-8397b7ea31f6
#    Domain: siopmgr.totemtowers.es
#    PAth:   /
#    Expire/Max-Age: Sesio
#    Tamaño: 46
#    Prioridad: medium
#Login
#   Metodo: POST
#   URL: https://siopmgr.totemtowers.es/tsp/api/identity/login
#   carga util: {username: "jose.lopezt", password: "Jorsadi-0"}
#   respuesta:
'''   {
    "token": "8d47d567-c68f-47b5-8212-eee1b137c882",
    "userId": "caded6f6-8d1f-4bfe-b233-1bf10dcf8767",
    "loginName": "jose.lopezt",
    "userName": "24n2o9kgyr",
    "firstName": "Jose",
    "lastName": "Lopez Tola",
    "ldap": false,
    "matomoSideId": "2",
    "matomoTokenAuth": "341aefe6ab4b503f29fbf21f294d9294"
    }
'''
#Peticion de un report:
#   URL: https://siopmgr.totemtowers.es/tsp/api/entity/type/work_order/export/csv
'''
POST /tsp/api/entity/type/work_order/export/csv HTTP/1.1
Accept: application/json, text/plain, */*
Accept-Encoding: gzip, deflate, br
Accept-Language: es
Connection: keep-alive
Content-Length: 91
Content-Type: application/json;charset=UTF-8
Cookie: X-auth-token=8d47d567-c68f-47b5-8212-eee1b137c882; X-client-tenant=tsp; CSRF-TOKEN=a2adb6e3-57d4-410f-919b-625c881a89b5; _pk_id.2.e9ea=73755523c7542126.1700396420.1.1700396420.1700396420.; _pk_ses.2.e9ea=1
Host: siopmgr.totemtowers.es
Origin: https://siopmgr.totemtowers.es
Referer: https://siopmgr.totemtowers.es/tsp/
Sec-Fetch-Dest: empty
Sec-Fetch-Mode: cors
Sec-Fetch-Site: same-origin
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0
X-CSRF-TOKEN: a2adb6e3-57d4-410f-919b-625c881a89b5
X-auth-token: 8d47d567-c68f-47b5-8212-eee1b137c882
X-client-language: es-ES
X-client-login: jose.lopezt
X-client-tenant: tsp
sec-ch-ua: "Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"
sec-ch-ua-mobile: ?0
sec-ch-ua-platform: "Windows"
'''
#Respuesta
'''
{
    "documentType": "export",
    "modelId": "csv",
    "originalName": "tsp_prod_export_work_order_2023-11-19_13-34.csv",
    "name": "2673f5c6-5fef-437d-b2cd-cd0183ab9754.csv",
    "type": "text/csv",
    "lastModified": 1700397301377,
    "size": 966017,
    "customPath": null,
    "version": "1.0",
    "url": null,
    "downloadUrl": null,
    "thumbnailUrl": null,
    "id": "a59d214d-5068-46ac-b8ab-98bdf646cef2",
    "fieldName": null
}
'''
#descarga del fichero
#   URL: https://siopmgr.totemtowers.es/tsp/api/dms/download/tsp/export/csv/2673f5c6-5fef-437d-b2cd-cd0183ab9754.csv/1.0

#https://siopmgr.totemtowers.es/tsp/#/identity/login
url_hello='https://siopmgr.totemtowers.es/tsp/#/identity/login'
url_init='https://siopmgr.totemtowers.es/tsp/api/client/tsp/init'
url_login='https://siopmgr.totemtowers.es/tsp/api/identity/login'
url_csv='https://siopmgr.totemtowers.es/tsp/api/entity/type/****/export/csv'
url_download='https://siopmgr.totemtowers.es/tsp/api/dms/download/tsp/export/csv/'

url = url_hello
'''
cookie={'CSRF-TOKEN':'e8724de2-0b85-4911-9262-8397b7ea31f6'}
'''
token_init=        '9c1a58fd-cf85-4d3e-9a02-873938214802'
token_login=       '4e53bcf6-99e1-42c6-9ea8-0b20fc97d5d2'
token_auth=        '598af62b-3901-4095-bcd2-8f282d1f788c'
token_csv_request= 'fbdb403a-20a9-494b-a9a6-71abc80594c4'
token_csv_download=''

auth = {'username': 'jose.lopezt', 'password': 'Jorsadi-0'}

ficheros={
    "ficheros":[        
        {
            "nombre": "workOrdes",
            "url": "work_order",
            "nombreFich": "WorkOrdes.csv"
        },
        {
            "nombre": "sites",
            "url": "site",
            "nombreFich": "Sites.csv"
        },
        {
            "nombre": "location",
            "url": "location",
            "nombreFich": "Location.csv"
        },
        {
            "nombre": "Address",
            "url": "address",
            "nombreFich": "Address.csv"
        },
        {
            "nombre": "site_access_request",
            "url": "site_access_request",
            "nombreFich": "SiteAccessRequest"
        },
        {
            "nombre": "tenants",
            "url": "tenancy",
            "nombreFich": "Tenancies.csv"
        }
    ]
}

dirSalida="I:\\Source\\python\\CoSDownload\\CoSDownload\\Salida\\"

#Hello
headers={
'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
'Accept-Encoding':'gzip, deflate, br',
'Accept-Language':'es',
'Cache-Control':'no-cache',
'Connection':'keep-alive',
'Host':'siopmgr.totemtowers.es',
'Pragma':'no-cache',
'Sec-Fetch-Dest':'document',
'Sec-Fetch-Mode':'navigate',
'Sec-Fetch-Site':'none',
'Sec-Fetch-User':'?1',
'Upgrade-Insecure-Requests':'1',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
'sec-ch-ua':'"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
'sec-ch-ua-mobile':'?0',
'sec-ch-ua-platform':'"Windows"',   
}
#cookie={'CSRF-TOKEN':token_init}

#x = requests.get(url, headers=headers, cookies=cookie,verify=False)
x = requests.get(url, headers=headers,verify=False)

#init
url=url_init
headers={
'Accept':'application/json, text/plain, */*',
'Accept-Encoding':'gzip, deflate, br',
'Accept-Language':'es',
'Connection':'keep-alive',
'Cookie':'CSRF-TOKEN=' + token_init,
'Host':'siopmgr.totemtowers.es',
'Referer':'https://siopmgr.totemtowers.es/tsp/',
'Sec-Fetch-Dest':'empty',
'Sec-Fetch-Mode':'cors',
'Sec-Fetch-Site':'same-origin',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
'X-CSRF-TOKEN': token_init,
'X-client-language':'en-US',
'X-client-tenant':'tsp',
'sec-ch-ua':'"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
'sec-ch-ua-mobile':'?0',
'sec-ch-ua-platform':'"Windows"',
}

#x = requests.get(url, headers=headers, cookies=cookie,verify=False)
x = requests.get(url, headers=headers, verify=False)

#login
url=url_login

headers={
'Accept':'application/json, text/plain, */*',
'Accept-Encoding':'gzip, deflate, br',
'Accept-Language':'es',
#'Cache-Control':'no-cache',
'Connection':'keep-alive',
'Content-Length':'49',
'Content-Type':'application/json',
'Cookie':'CSRF-TOKEN=' + token_login,
'Host':'siopmgr.totemtowers.es',
'Origin':'https://siopmgr.totemtowers.es',
#'Pragma':'no-cache';
'Referer':'https://siopmgr.totemtowers.es/tsp/',
'Sec-Fetch-Dest':'empty',
'Sec-Fetch-Mode':'cors',
'Sec-Fetch-Site':'same-origin',
'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
'X-CSRF-TOKEN': token_login,
'X-client-language':'en-US',
'X-client-tenant':'tsp',
'sec-ch-ua':'"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
'sec-ch-ua-mobile':'?0',
'sec-ch-ua-platform':'"Windows"',
}

#cookie={'CSRF-TOKEN':token_login}
payload=auth #{"username":"jose.lopezt","password":"Jorsadi-0"}

x = requests.post(url, headers=headers, json=payload, verify=False)

if(x.status_code!=200): 
    print("login fallido")
    exit

response=x.json()
token_auth=response["token"]

############################################################
fifi=ficheros["ficheros"]
for fichero in fifi:

    print("Iniciando descarga de " + fichero["nombre"])
    #csv request
    url= url_csv.replace("****",fichero["url"])

    headers={
    'Accept':'application/json, text/plain, */*',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'es',
    'Cache-Control':'no-cache',
    'Connection':'keep-alive',
    'Content-Length':'91',
    'Content-Type':'application/json;charset=UTF-8',
    'Cookie':'X-auth-token=' + token_auth + '; X-client-tenant=tsp; CSRF-TOKEN=' + token_csv_request + '; _pk_id.2.e9ea=8bbc36756d661c7e.1700404894.1.1700404894.1700404894.; _pk_ses.2.e9ea=1',
    'Host':'siopmgr.totemtowers.es',
    'Origin':'https://siopmgr.totemtowers.es',
    'Pragma':'no-cache',
    'Referer':'https://siopmgr.totemtowers.es/tsp/',
    'Sec-Fetch-Dest':'empty',
    'Sec-Fetch-Mode':'cors',
    'Sec-Fetch-Site':'same-origin',
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
    'X-CSRF-TOKEN': token_csv_request,
    'X-auth-token': token_auth,
    'X-client-language':'es-ES',
    'X-client-login':'jose.lopezt',
    'X-client-tenant':'tsp',
    'sec-ch-ua':'"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
    'sec-ch-ua-mobile':'?0',
    'sec-ch-ua-platform':'"Windows"',
    }

    #Cookie: X-auth-token=598af62b-3901-4095-bcd2-8f282d1f788c; X-client-tenant=tsp; CSRF-TOKEN=fbdb403a-20a9-494b-a9a6-71abc80594c4; _pk_id.2.e9ea=044e0a78911ef4d7.1700407115.1.1700407115.1700407115.; _pk_ses.2.e9ea=1
    #X-CSRF-TOKEN: fbdb403a-20a9-494b-a9a6-71abc80594c4
    #X-auth-token: 598af62b-3901-4095-bcd2-8f282d1f788c

    payload={'entityType':'work_order', 'filters':[], 'facets':[], 'searchExpression':'', 'excludedIds':[]}

    x = requests.post(url, headers=headers, json=payload, verify=False)

    if(x.status_code!=200): 
        print("Solicitud de descarga fallida")
        exit

    response=x.json()
    nombreFichero=response["name"]

    #download
    url= url_download

    headers={
    'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'es',
    'Cache-Control':'no-cache',
    'Connection':'keep-alive',
    'Cookie':'X-auth-token=' + token_auth + '; X-client-tenant=tsp; CSRF-TOKEN=' + token_csv_request + '; _pk_id.2.e9ea=044e0a78911ef4d7.1700407115.1.1700407115.1700407115.; _pk_ses.2.e9ea=1',
    'Host':'siopmgr.totemtowers.es',
    'Pragma':'no-cache',
    'Referer':'https://siopmgr.totemtowers.es/tsp/',
    'Sec-Fetch-Dest':'document',
    'Sec-Fetch-Mode':'navigate',
    'Sec-Fetch-Site':'same-origin',
    'Upgrade-Insecure-Requests':'1',
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0',
    'sec-ch-ua':'"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
    'sec-ch-ua-mobile':'?0',
    'sec-ch-ua-platform':'"Windows"',
    }

    url=url_download + nombreFichero + '/1.0'
    x = requests.get(url, headers=headers, verify=False)

    nombreSalida=dirSalida + fichero["nombreFich"]
    
    contenido=x.text
    contenido=contenido.removeprefix("ï»¿")
    contenido = contenido.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

    f=open(nombreSalida,"w", encoding="utf-8-sig")
    f.write(contenido)
    f.close()
    #print(x.text)

