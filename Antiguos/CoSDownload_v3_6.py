import requests
import warnings

from datetime import date
from datetime import datetime

import os
import glob
import csv

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

from xlsxwriter.workbook import Workbook

warnings.filterwarnings("ignore")

# replacement strings
BASURA_INICIAL = ''#'ï»¿'
WINDOWS_LINE_ENDING = '\r\n'
UNIX_LINE_ENDING = '\n'

#URLs de CoS
url_hello='https://siopmgr.totemtowers.es/tsp/#/identity/login'
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
            "nombreFich": "WorkOrdes",
            "entityType": "work_order"
        },
        {
            "nombre": "sites",
            "url": "site",
            "nombreFich": "Sites",
            "entityType": "site"
        },
        {
            "nombre": "location",
            "url": "location",
            "nombreFich": "Location",
            "entityType": "location"
        },
        {
            "nombre": "Address",
            "url": "address",
            "nombreFich": "Address",
            "entityType": "address"
        },
        {
            "nombre": "site_access_request",
            "url": "site_access_request",
            "nombreFich": "SiteAccessRequest",
            "entityType": "site_access_request"
        },
        {
            "nombre": "tenants",
            "url": "tenancy",
            "nombreFich": "Tenancies",
            "entityType": "tenancy"
        },
        {
            "nombre": "Contratos tenants",
            "url": "tenant_lease",
            "nombreFich": "contratosTenants",
            "entityType": "tenant_lease"
        },
        {
            "nombre": "Acuerdos marco",
            "url": "frame_agreement",
            "nombreFich": "acuerdosMarco",
            "entityType": "frame_agreement"
        },
        {
            "nombre": "Contratos arrendamiento",
            "url": "lease",
            "nombreFich": "contratosArrendamiento",
            "entityType": "lease"
        },
        {
            "nombre": "Incidencias de contratos",
            "url": "lease_request",
            "nombreFich": "incidenciasContratos",
            "entityType": "lease_request"
        },        
        {
            "nombre": "Partes relacionadas",
            "url": "related_third_party",
            "nombreFich": "relatedParties",
            "entityType": "related_third_party"
        },
        {
            "nombre": "condiciones beneficiario",
            "url": "beneficiary_condition",
            "nombreFich": "beneficiaryConditions",
            "entityType": "beneficiary_condition"
        },
        {
            "nombre": "Detalle bancario",
            "url": "bank_detail",
            "nombreFich": "bankDetails",
            "entityType": "bank_detail"
        },
         {
            "nombre": "Energia",
            "url": "power",
            "nombreFich": "energia",
            "entityType": "power"
        },
         {
            "nombre": "contratos de energia",
            "url": "power_contract",
            "nombreFich": "contratosEnergia",
            "entityType": "power_contract"
        },
        {
            "nombre": "Autorizaciones administrativas",
            "url": "administrative_authorization",
            "nombreFich": "autorizacionesAdministrativas",
            "entityType": "administrative_authorization"
        },
        {
            "nombre": "Informe de visita",
            "url": "visit_report",
            "nombreFich": "informeVisita",
            "entityType": "visit_report"
        },
        {
            "nombre": "Defectos",
            "url": "snag",
            "nombreFich": "defectos",
            "entityType": "snag"
        },
        {
            "nombre": "Riesgos",
            "url": "risk_management",
            "nombreFich": "riesgos",
            "entityType": "risk_management"
        },
        {
            "nombre": "Tickets",
            "url": "trouble_ticket",
            "nombreFich": "tickets",
            "entityType": "trouble_ticket"
        },
        {
            "nombre": "Pedidos",
            "url": "purchase_request",
            "nombreFich": "pedidos",
            "entityType": "purchase_request"
        },
        {
            "nombre": "Hojas de acceso",
            "url": "access_sheet",
            "nombreFich": "hojasAcceso",
            "entityType": "access_sheet"
        }
        #,
        #{
        #    "nombre": "equipamientoTenants",
        #    "url": "equipment_tenant",
        #    "nombreFich": "equipamientoTenants",
        #    "entityType": "equipment_tenant"
        #},         
        #{
        #    "nombre": "equipamientoTotem",
        #    "url": "equipment",
        #    "nombreFich": "equipamientoTotem",
        #    "entityType": "equipment"
        #}
    ]
}

#Directorio donde se guardaran los ficheros
dirSalida="Salida/"
dirAutomatismos="Automatismos/"

target_url = "Documentos compartidos/ExtraccionAutomaticaCoS"
size_chunk = 1000000
#local_path = "C:\desarrollo\python\CoSDownload\Salida\equipamientoTotem.csv"


def csv2xlsx(path_origen,nombre_fichero,nombre_hoja=''):
    for csvfile in glob.glob(os.path.join(path_origen, nombre_fichero)):
        workbook = Workbook(csvfile[:-4] + '.xlsx')
        worksheet = workbook.add_worksheet(nombre_hoja)
        with open(csvfile, 'rt', encoding='utf8') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
        workbook.close()

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


#Día actual
today = date.today()

#Fecha actual
now = datetime.now()

hora = now.hour
minuto = now.minute
segundo = now.second
dia = today.day
mes = today.month
anno = today.year

#creo el directorio en Sharepoint
target_url_Salida = target_url + "/Reports_"  + '_' + str(anno) + '_' + str(mes) + '_' + str(dia) + '_' + str(hora) + '_' + str(minuto) + '_' + str(segundo) 
target_url_Automatismos = target_url + "/Automatismos"
folder = (
    ctx.web.default_document_library()
    .root_folder.folders.add_using_path(target_url_Salida, overwrite=True)
    .execute_query()
)
target_folder_Salida = ctx.web.get_folder_by_server_relative_url(target_url_Salida)
target_folder_Automatismos = ctx.web.get_folder_by_server_relative_url(target_url_Automatismos)

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

    payload={'entityType':fichero['entityType'], 'filters':[], 'facets':[], 'searchExpression':'', 'excludedIds':[]}

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

        #Día actual
        today = date.today()

        #Fecha actual
        now = datetime.now()

        hora = now.hour
        minuto = now.minute
        segundo = now.second
        dia = today.day
        mes = today.month
        anno = today.year

        nombreSalida=dirSalida + fichero["nombreFich"] + '_' + str(anno) + '_' + str(mes) + '_' + str(dia) + '_' + str(hora) + '_' + str(minuto) + '_' + str(segundo) + '.csv'
        nombreAutomatismos=dirAutomatismos + fichero["nombreFich"] + '.csv'
        
        contenido=x.text
        #contenido=contenido.removeprefix("ï»¿")
        contenido = contenido.replace(BASURA_INICIAL, '')
        contenido = contenido.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

        try:
            #Lo salvo y subo a la carpeta del dia/hora
            #f=open(nombreSalida,"w", encoding="utf-8-sig")
            #f=open(nombreSalida,"w", encoding="utf-8")
            f=open(nombreSalida,"w", encoding="ISO-8859-1")
            #f=open(nombreSalida,"w", encoding="windows-1252")
            f.write(contenido)

            with open(nombreSalida, "rb") as f:
                uploaded_file = target_folder_Salida.files.create_upload_session(
                f, size_chunk, print_upload_progress
                ).execute_query()

            f.close()             

            print("File {0} subido correctamente".format(uploaded_file.serverRelativeUrl))

            #Lo salvo y subo a la carpeta de automatismos
            #f=open(nombreAutomatismos,"w", encoding="utf-8-sig")
            #f=open(nombreAutomatismos,"w", encoding="utf-8")
            f=open(nombreAutomatismos,"w", encoding="ISO-8859-1")
            #f=open(nombreAutomatismos,"w", encoding="windows-1252")#cp1252
            f.write(contenido)

            with open(nombreAutomatismos, "rb") as f:
                uploaded_file = target_folder_Automatismos.files.create_upload_session(
                f, size_chunk, print_upload_progress
                ).execute_query()
            f.close() 

            #Creo y subo el xlx
            print("Convirtiendo " + fichero["nombreFich"] + ".csv a xlsx")
            csv2xlsx(dirAutomatismos,fichero["nombreFich"] + '.csv',nombre_hoja='')
            xlsxAutomatismo=nombreAutomatismos[:-4] + ".xlsx"
            
            print("Subiendo " + xlsxAutomatismo + " a SharePoint")
            with open(xlsxAutomatismo, "rb") as f:
                uploaded_file = target_folder_Automatismos.files.create_upload_session(
                f, size_chunk, print_upload_progress
                ).execute_query()

            f.close() 

            print("File {0} subido correctamente".format(uploaded_file.serverRelativeUrl))
        except Exception as e:
            print(f'caught {type(e)}: e')
            continue

    else:
        print("Solicitud de descarga fallida")
        
