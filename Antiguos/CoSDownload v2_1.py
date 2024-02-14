import requests
import warnings


import os

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

warnings.filterwarnings("ignore")

# replacement strings
BASURA_INICIAL = 'ï»¿'
WINDOWS_LINE_ENDING = '\r\n'
UNIX_LINE_ENDING = '\n'

#URLs de CoS
url_COS='https://siopmgr.ppr.totemtowers.es'
"""
url_hello=url_COS + '/tsp/#/identity/login'
url_init=url_COS + '/tsp/api/client/tsp/init'
url_login=url_COS + '/tsp/api/identity/login'
url_csv=url_COS + '/tsp/api/entity/type/****/export/csv'
url_download=url_COS + '/tsp/api/dms/download/tsp/export/csv/'
"""
url_hello='https://siopmgr.ppr.totemtowers.es/tsp/#/identity/login'
url_init='https://siopmgr.totemtowers.es/tsp/api/client/tsp/init'
url_login='https://siopmgr.totemtowers.es/tsp/api/identity/login'
url_csv='https://siopmgr.totemtowers.es/tsp/api/entity/type/****/export/csv'
url_download='https://siopmgr.totemtowers.es/tsp/api/dms/download/tsp/export/csv/'

#Sharepoint URL y usuario
username = "robotFicheros@totemtowersspain.es"
password = "R0b0tF1ch3r0s"
team_site_url = "https://totemtowersspain.sharepoint.com/sites/Prueba_QLIK"

#Tokens sacados desde el navegador
token_init=        '9c1a58fd-cf85-4d3e-9a02-873938214802'
token_login=       '4e53bcf6-99e1-42c6-9ea8-0b20fc97d5d2'
token_auth=        '598af62b-3901-4095-bcd2-8f282d1f788c'
token_csv_request= 'fbdb403a-20a9-494b-a9a6-71abc80594c4'
token_csv_download=''

#Si cambia la password, hay que cambiarla aqui
auth = {'username': 'jose.lopezt', 'password': 'Jorsadi-0'}

#Lista de fcheros que va a leer
FICH={
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
            "nombreFich": "SiteAccessRequest.csv"
        },
        {
            "nombre": "tenants",
            "url": "tenancy",
            "nombreFich": "Tenancies.csv"
        },
        {
            "nombre": "Contratos tenants",
            "url": "tenant_lease",
            "nombreFich": "contratosTenants.csv"
        },
        {
            "nombre": "Acuerdos marco",
            "url": "frame_agreement",
            "nombreFich": "acuerdosMarco.csv"
        },
        {
            "nombre": "Contratos arrendamiento",
            "url": "lease",
            "nombreFich": "contratosArrendamiento.csv"
        },
        {
            "nombre": "Incidencias de contratos",
            "url": "lease_request",
            "nombreFich": "incidenciasContratos.csv"
        },        
        {
            "nombre": "Partes relacionadas",
            "url": "related_third_party",
            "nombreFich": "relatedParties.csv"
        },
        {
            "nombre": "condiciones beneficiario",
            "url": "beneficiary_condition",
            "nombreFich": "beneficiaryConditions.csv"
        },
        {
            "nombre": "Detalle bancario",
            "url": "bank_detail",
            "nombreFich": "bankDetails.csv"
        },
         {
            "nombre": "Energia",
            "url": "power",
            "nombreFich": "energia.csv"
        },
         {
            "nombre": "contratos de energia",
            "url": "power_contract",
            "nombreFich": "contratosEnergia.csv"
        },
        {
            "nombre": "Autorizaciones administrativas",
            "url": "administrative_authorization",
            "nombreFich": "autorizacionesAdministrativas.csv"
        },
        {
            "nombre": "Informe de visita",
            "url": "visit_report",
            "nombreFich": "informeVisita.csv"
        },
        {
            "nombre": "Defectos",
            "url": "snag",
            "nombreFich": "defectos.csv"
        },
        {
            "nombre": "Riesgos",
            "url": "risk_management",
            "nombreFich": "riesgos.csv"
        },
        {
            "nombre": "Tickets",
            "url": "trouble_ticket",
            "nombreFich": "tickets.csv"
        },
        {
            "nombre": "Pedidos",
            "url": "purchase_request",
            "nombreFich": "pedidos.csv"
        },
        {
            "nombre": "Hojas de acceso",
            "url": "access_sheet",
            "nombreFich": "hojasAcceso.csv"
        },
        {
            "nombre": "equipamientoTenants",
            "url": "equipment_tenant",
            "nombreFich": "equipamientoTenants.csv"
        },         
        {
            "nombre": "equipamientoTotem",
            "url": "equipment",
            "nombreFich": "equipamientoTotem.csv"
        }
    ]
}

#Directorio donde se guardaran los ficheros
dirSalida="Salida\\"

#COMENZAMOS##################################################
test_team_site_url = team_site_url
test_user_credentials = UserCredential(username,password)


def print_upload_progress(offset):
    # type: (int) -> None
    
    pass
    """
    file_size = os.path.getsize(local_path)
    print(
        "Uploaded '{0}' bytes from '{1}'...[{2}%]".format(
            offset, file_size, round(offset / file_size * 100, 2)
        )
    )
    """

ctx = ClientContext(test_team_site_url).with_credentials(test_user_credentials)

target_url = "Documentos compartidos/ExtraccionAutomaticaCoS"
target_folder = ctx.web.get_folder_by_server_relative_url(target_url)
size_chunk = 1000000
#local_path = "C:\desarrollo\python\CoSDownload\Salida\equipamientoTotem.csv"

#Hello
print("Iniciando Hello")
url = url_hello

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

if(x.status_code!=200): 
    print("Solicitud de hello fallida")
    exit

#init
print("Iniciando init")
url=url_init

headers={
'Accept':'application/json, text/plain, */*',
'Accept-Encoding':'gzip, deflate, br',
'Accept-Language':'es',
'Connection':'keep-alive',
'Cookie':'CSRF-TOKEN=' + token_init,
'Host':'siopmgr.totemtowers.es',
'Referer':url_COS + '/tsp/',
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

x = requests.get(url, headers=headers, verify=False)

if(x.status_code!=200): 
    print("Solicitud de init fallida")
    exit

#login
print("Iniciando login")
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
'Origin':url_COS + '',
#'Pragma':'no-cache';
'Referer':url_COS + '/tsp/',
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

payload=auth #{"username":"xxxxxxx","password":"xxxxx"}

x = requests.post(url, headers=headers, json=payload, verify=False)

if(x.status_code!=200): 
    print("login fallido")
    exit

response=x.json()
token_auth=response["token"]

##########################INICIO DE BUCLE PARA FICHEROS##################################
print("Iniando peticiones de decarga")
ficheros=FICH["ficheros"]
for fichero in ficheros:

    print("Iniciando solicitud de " + fichero["nombre"])
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
    'Origin':url_COS + '',
    'Pragma':'no-cache',
    'Referer':url_COS + '/tsp/',
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

    #payload={'entityType':'work_order', 'filters':[], 'facets':[], 'searchExpression':'', 'excludedIds':[]}

    payload={'entityType':fichero["url"], 'filters':[], 'facets':[], 'searchExpression':'', 'excludedIds':[]}


    x = requests.post(url, headers=headers, json=payload, verify=False)

    if(x.status_code==200): 
        response=x.json()
        nombreFichero=response["name"]

        #download
        print("Iniciando Descarga")
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
        'Referer':url_COS + '/tsp/',
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
        #contenido=contenido.removeprefix("ï»¿")
        contenido = contenido.replace(BASURA_INICIAL, '')
        contenido = contenido.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

        f=open(nombreSalida,"w", encoding="utf-8-sig")
        f.write(contenido)
        f.close()        

        with open(nombreSalida, "rb") as f:
            uploaded_file = target_folder.files.create_upload_session(
            f, size_chunk, print_upload_progress
            ).execute_query()

        print("File {0} subido correctamente".format(uploaded_file.serverRelativeUrl))

    else:
        print("Solicitud de descarga fallida")
        
