#####https://platzi.com/tutoriales/1540-flask/9127-configuracion-de-powershell-para-crear-el-entorno-virtual-en-windows/

import requests
import warnings

from datetime import date
from datetime import datetime

import os
import glob
import csv
import json

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

from xlsxwriter.workbook import Workbook

warnings.filterwarnings("ignore")

# replacement strings
BASURA_INICIAL = 'ï»¿'
WINDOWS_LINE_ENDING = '\r\n'
UNIX_LINE_ENDING = '\n'

#fichero configuracion de reports a leer
reports='reports.json'

#URLs de CoS
url_hello='https://siopmgr.totemtowers.es/tsp/#/identity/login'
url_init='https://siopmgr.totemtowers.es/tsp/api/client/tsp/init'
url_login='https://siopmgr.totemtowers.es/tsp/api/identity/login'
url_reports='https://siopmgr.totemtowers.es/tsp/api/entity/type/reports/[FICHERO_URL]/read'
url_sync_Time='https://siopmgr.totemtowers.es/tsp/api/report/synchTime'
url_report ='https://siopmgr.totemtowers.es/tsp/api/report'
url_download = 'https://siopmgr.totemtowers.es/tsp/api/dms/download/tsp/export/csv/[FICHERO_DESCARGA]/1.0'

#Sharepoint URL y usuario
username = "robotFicheros@totemtowersspain.es"
password = "R0b0tF1ch3r0s"
team_site_url = "https://totemtowersspain.sharepoint.com/sites/Prueba_QLIK"


#Payloads
sync_payload = {"last_sync": 1705318200011}
report_payload = {
  "fields": None,
  "value": None,
  "count": 27932,
  "reportName": None,
  "dmsFileResponse": {
    "documentType": "export",
    "modelId": "csv",
    "originalName": "tsp_prod_report_customer_request_2024-01-15_16-57.csv",
    "name": "0dab0906-113f-4af9-9387-f4146dc2908a.csv",
    "type": "text/csv",
    "lastModified": 1705334228709,
    "size": 9862939,
    "customPath": None,
    "version": "1.0",
    "url": None,
    "downloadUrl": None,
    "thumbnailUrl": None,
    "id": "a7ce1ab6-a104-48b0-9b18-8f3379571d07",
    "fieldName": None
  },
  "facets": None
}


#Tokens sacados desde el navegador
token_init=        '9c1a58fd-cf85-4d3e-9a02-873938214802'
token_login=       '4e53bcf6-99e1-42c6-9ea8-0b20fc97d5d2'
token_auth=        '598af62b-3901-4095-bcd2-8f282d1f788c'
token_csv_request= 'fbdb403a-20a9-494b-a9a6-71abc80594c4'
token_csv_download=''

#Si cambia la password, hay que cambiarla aqui
auth = {'username': 'jose.lopezt', 'password': 'Jorsadi-0'}

#Lista de report que va a leer
'''FICH={
    "ficheros":[
        {
            "nombre": "Arbolicos",
            "url": "bd1d6f02-33b9-474b-9d5e-b6ddad0fa34f",
            "nombreFich": "CoSActividadOperativa"
        }
    ]
}
'''

try:
    f=open(reports,"r")
    txt=f.read()
    FICH=json.loads(txt)
except:
    print("No se pudo abrir el fichero" + reports)
    FICH={}
finally:
    f.close()

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
target_url_Salida = target_url + "/arboliCoS_"  + '_' + str(anno) + '_' + str(mes) + '_' + str(dia) + '_' + str(hora) + '_' + str(minuto) + '_' + str(segundo) 
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
    url= url_reports.replace("[FICHERO_URL]",fichero["url"])

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

    payload={'entityType':'reports', 'id': 'bd1d6f02-33b9-474b-9d5e-b6ddad0fa34f'}

    x = requests.post(url, headers=headers, json=payload, verify=False)

    if(x.status_code==200): 

        #sync
        print("Iniciando sync")
        url= url_sync_Time
        
        headers={
            "Accept": "application/json, text/plain, */*",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "es,es-ES;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Cookie": "X-auth-token="+token_auth+"; X-client-tenant=tsp; _pk_id.2.e9ea=e6ae610fc866f0a8.1700825629.2.1705648973.1705648973.; CSRF-TOKEN="+ token_csv_request,
            "Host": "siopmgr.totemtowers.es",
            'Origin':'https://siopmgr.totemtowers.es',
            "Pragma": "no-cache",
            "Referer": "https://siopmgr.totemtowers.es/tsp/",
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
            "X-CSRF-TOKEN": token_csv_request,
            "X-auth-token": token_auth,
            "X-client-language": "es-ES",
            "X-client-login": "jose.lopezt",
            "X-client-tenant": "tsp",
            "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Microsoft Edge";v="120"',
            "sec-ch-ua-mobile": "?0",
            "sec-ch-ua-platform": "Windows"
        }

        payload={"last_sync": 1705318200011}
        x = requests.get(url, headers=headers, json=payload, verify=False)

        if(x.status_code==200): 

            #report
            print("Iniciando report")
            url= url_report
            
            headers={
                "Accept": "application/json, text/plain, */*",
                "Accept-Encoding": "gzip, deflate, br",
                "Accept-Language": "es,es-ES;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                "Cache-Control": "no-cache",
                "Connection": "keep-alive",
                "Content-Length": "14514",
                "Cookie": "X-auth-token="+token_auth+"; X-client-tenant=tsp; _pk_id.2.e9ea=e6ae610fc866f0a8.1700825629.2.1705648973.1705648973.; CSRF-TOKEN="+token_csv_request,
                "Host": "siopmgr.totemtowers.es",
                "Pragma": "no-cache",
                "Referer": "https://siopmgr.totemtowers.es/tsp/",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
                "X-CSRF-TOKEN": token_csv_request,
                "X-auth-token": token_auth,
                "X-client-language": "es-ES",
                "X-client-login": "jose.lopezt",
                "X-client-tenant": "tsp",
                "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Microsoft Edge";v="120"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": "Windows"              
            }

            payload={
                "entityType":"customer_request","lastSeenId":None,"pageSize":250,"createNotification":False,"entityArrayNodeType":"list","advanced":True,"viewType":1,"fields":[{"title":"Código identificator proyecto","aggregate":None,"reference":"customer_request.customer_request_id"},{"title":"ID Site","aggregate":None,"reference":"customer_request.customerrequest_site.site_id"},{"title":"Código Solicitud terceros","aggregate":None,"reference":"customer_request.site_tenant_request"},{"title":"Código Site Tenant","aggregate":None,"reference":"customer_request.site_tenant_code"},{"title":"Id Emplazamiento","aggregate":None,"reference":"customer_request.customerrequest_site.location_site.location_id"},{"title":"Provincia","aggregate":None,"reference":"customer_request.customerrequest_site.location_site.location_address.state"},{"title":"Municipio","aggregate":None,"reference":"customer_request.customerrequest_site.location_site.location_address.city"},{"title":"Zona Solicitante Tenant","aggregate":None,"reference":"customer_request.tenant_requester_zone"},{"title":"Zona Totem","aggregate":None,"reference":"customer_request.customerrequest_site.zona_totem"},{"title":"totem","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.totem"},{"title":"Modalidad Compartición","aggregate":None,"reference":"customer_request.process_type"},{"title":"Submodalidad de Compartición","aggregate":None,"reference":"customer_request.sharing_subtype"},{"title":"Plan","aggregate":None,"reference":"customer_request.sharing_plan"},{"title":"Subplan","aggregate":None,"reference":"customer_request.sharing_subplan"},{"title":"Año","aggregate":None,"reference":"customer_request.year_plan"},{"title":"Cod Solicitud","aggregate":None,"reference":"customer_request.site_tenant_request"},{"title":"Estado Solicitud","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.status"},{"title":"Causa Abandono","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.process_aborted"},{"title":"Tecnologías a Instalar","aggregate":None,"reference":"customer_request.technology_install"},{"title":"Proyecto Activo","aggregate":None,"reference":"customer_request.sharing_name"},{"title":"Coste Total Adecuación","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.adequation_total_costs"},{"title":"Cost Adecuación Tenant","aggregate":None,"reference":"customer_request.customerrequest_tenancy.adequation_costs"},{"title":"Incremento Renta Tenant","aggregate":None,"reference":"customer_request.customerrequest_tenancy.oneshot_costs"},{"title":"Nombre Tenant","aggregate":None,"reference":"customer_request.customerrequest_tenancy.tenancy_relatedparty.name"},{"title":"Adecuación","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.adequation_need"},{"title":"Nuevos Costes Adecuación?","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.new_adequation_costs"},{"title":"Fecha Solicitud","aggregate":None,"reference":"customer_request.expected_date"},{"title":"Fecha Previabilidad","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_pa_sm"},{"title":"Fecha Prevista Replanteo","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.technical_visit_expected_start_date"},{"title":"Fecha Real  Replanteo","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.technical_visit_real_start_date"},{"title":"Fecha envío Planos 1","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_1"},{"title":"Inicio Estudio de Carga","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_cve_opening"},{"title":"Fin Estudio de Carga","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_cve_end"},{"title":"Inicio Pdte Solicitante 1","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_1"},{"title":"Fin Pdte Solicitante 1","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_1"},{"title":"Fecha Rechazo Planos 1","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_1"},{"title":"Fecha envío Planos 2","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_2"},{"title":"Inicio Pdte Solicitante 2","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_2"},{"title":"Fin Pdte Solicitante 2","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_2"},{"title":"Fecha Rechazo Planos 2","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_2"},{"title":"Fecha envío Planos 3","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_3"},{"title":"Inicio Pdte Solicitante 3","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_3"},{"title":"Fin Pdte Solicitante 3","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_3"},{"title":"Fecha Rechazo Planos 3","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_3"},{"title":"Fecha envío Planos 4","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_4"},{"title":"Inicio Pdte Solicitante 4","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_4"},{"title":"Fin Pdte Solicitante 4","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_4"},{"title":"Fecha Rechazo Planos 4","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_4"},{"title":"Fecha envío Planos 5","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_5"},{"title":"Inicio Pdte Solicitante 5","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_5"},{"title":"Fin Pdte Solicitante 5","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_5"},{"title":"Fecha Rechazo Planos 5","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_5"},{"title":"Fecha envío Planos 6","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_6"},{"title":"Inicio Pdte Solicitante 6","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_6"},{"title":"Fin Pdte Solicitante 6","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_6"},{"title":"Fecha Rechazo Planos 6","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_6"},{"title":"Fecha envío Planos 7","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_7"},{"title":"Inicio Pdte Solicitante 7","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_7"},{"title":"Fin Pdte Solicitante 7","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_7"},{"title":"Fecha Rechazo Planos 7","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_7"},{"title":"Fecha envío Planos 8","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_8"},{"title":"Inicio Pdte Solicitante 8","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_8"},{"title":"Fin Pdte Solicitante 8","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_8"},{"title":"Fecha Rechazo Planos 8","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_8"},{"title":"Fecha envío Planos 9","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_9"},{"title":"Inicio Pdte Solicitante 9","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_9"},{"title":"Fin Pdte Solicitante 9","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_9"},{"title":"Fecha Rechazo Planos 9","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_9"},{"title":"Fecha Rechazo Planos 10","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_reject_cap_10"},{"title":"Inicio Pdte Solicitante 10","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_wta_10"},{"title":"Fin Pdte Solicitante 10","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_wta_10"},{"title":"Fecha envío Planos 10","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_send_cap_10"},{"title":"Fecha Aprobación Planos","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_valid_cap"},{"title":"Fecha Inicio Renegociación","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_renegociation"},{"title":"Fecha Fin Renegociación","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_renegociation"},{"title":"Fecha de Recepción","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.feedback_email_reception_date_adequation"},{"title":"Fecha Inicio Legalización","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_legalization"},{"title":"Fecha Legalización","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_legalization"},{"title":"Fecha Inicio Adecuación Totem","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_rfc"},{"title":"Inicio PRL","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_prl"},{"title":"Fin PRL","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_prl"},{"title":"Inicio Adecuación Energía","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_power_management"},{"title":"Finalización Adecuación Energía","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_power_management"},{"title":"Inicio Refuerzo","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_start_reinforcement"},{"title":"Fin Refuerzo","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_reinforcement"},{"title":"Fecha Fin Adecuación Totem","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_end_adequation"},{"title":"Fecha de RFI","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.site_ready_date"},{"title":"Fecha de Necesidad","aggregate":None,"reference":"customer_request.start_date"},{"title":"Fecha de finalización esperada","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.works_expected_end_date"},{"title":"Fecha Abandono","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.date_abandon"},{"title":"Titular","aggregate":None,"reference":"customer_request.projectinstance_customerrequest.totem"},{"title":"idattivita","aggregate":None,"reference":"customer_request.miscellaneous"}],"grouping":[],"having":[],"filters":[{"show":True,"field":{"system":False,"is_open":False,"pattern":None,"can_edit":False,"field_id":3,"multiple":False,"settings":None,"can_draft":False,"essential":False,"field_key":"","field_url":None,"full_path":False,"hint_text":"","is_report":True,"max_value":None,"min_value":None,"field_data":[],"field_type":"dropdown","output_pad":1,"custom_path":"","default_now":False,"field_group":"Added for BTS","field_title":"Tipo de Solicitud","field_value":"","linkedValue":False,"placeholder":"","counter_mode":False,"field_unique":False,"regex_search":"","field_options":[{"option_id":1,"option_bold":False,"option_color":"","option_title":"Nuevo Site","option_value":1,"option_user_value":None,"option_colorize_text":False},{"option_id":5,"option_bold":False,"option_color":"","option_title":"Modificacion Site - Nueva Tenant","option_value":5,"option_user_value":None,"option_colorize_text":False},{"option_id":2,"option_bold":False,"option_color":"","option_title":"Modificacion Site - Ampliación","option_value":2,"option_user_value":None,"option_colorize_text":False},{"option_id":6,"option_bold":False,"option_color":"","option_title":"Contrato Energía","option_value":6,"option_user_value":None,"option_colorize_text":False},{"option_id":7,"option_bold":False,"option_color":"","option_title":"Cambio de Titular","option_value":7,"option_user_value":None,"option_colorize_text":False},{"option_id":8,"option_bold":False,"option_color":"","option_title":"Transferencia de Site","option_value":8,"option_user_value":None,"option_colorize_text":False},{"option_id":9,"option_bold":False,"option_color":"","option_title":"Legalización","option_value":9,"option_user_value":None,"option_colorize_text":False},{"option_id":10,"option_bold":False,"option_color":"","option_title":"Renegociación","option_value":10,"option_user_value":None,"option_colorize_text":False}],"show_in_lists":True,"field_disabled":False,"field_required":False,"override_facet":None,"pass_date_only":False,"tree_navigator":False,"field_delimiter":"","not_update_able":False,"regex_reference":"","show_in_details":True,"show_in_filters":True,"starting_number":None,"editable_in_grid":False,"field_constraint":"","future_date_only":False,"show_in_map_view":True,"field_system_name":"process_type","ignore_on_preview":False,"regex_replacement":"","show_in_card_view":True,"show_in_graph_view":False,"show_in_milestones":False,"field_related_value":"","original_field_type":None,"show_in_lists_short":True,"autoincrement_position":"","autoincrement_prefix_type":"","show_in_digital_twin_view":False},"title":"Tipo de Solicitud","values":[5,2],"fieldName":"process_type","operation":"EQUALS","reference":"customer_request.process_type","isOperationValueValid":True}],"searchExpression":"","multipleType":False,"orders":[]}
            x = requests.post(url, headers=headers, json=payload, verify=False)

            if(x.status_code==200): 
                #download 1.0
                print("Iniciando download")
                response=x.json()
                dms=response["dmsFileResponse"]
                nombreFichero=dms["originalName"]#el nombre que guardo
                nombre=dms["name"] #el nombre que pido

                url= url_download.replace("[FICHERO_DESCARGA]",nombre)

                headers={
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                "Accept-Encoding": "gzip, deflate, br",
                "Accept-Language": "es,es-ES;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                "Cache-Control": "no-cache",
                "Connection": "keep-alive",
                "Cookie": "X-auth-token="+token_auth+"; X-client-tenant=tsp; _pk_id.2.e9ea=e6ae610fc866f0a8.1700825629.2.1705648973.1705648973.; CSRF-TOKEN="+ token_csv_request,
                "Host": "siopmgr.totemtowers.es",
                "Pragma": "no-cache",
                "Referer": "https://siopmgr.totemtowers.es/tsp/",
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
                "X-CSRF-TOKEN": token_csv_request,
                "X-auth-token": token_auth,
                "X-client-language": "es-ES",
                "X-client-login": "jose.lopezt",
                "X-client-tenant": "tsp",
                "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Microsoft Edge";v="120"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": "Windows",
                "Sec-Fetch-User": "?1",
                "Upgrade-Insecure-Requests": "1"
                }

                payload={}                
                x = requests.get(url, headers=headers, json=payload, verify=False)

                if(x.status_code==200): 
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
                    contenido = contenido.replace(BASURA_INICIAL, '')
                    contenido = contenido.replace(WINDOWS_LINE_ENDING, UNIX_LINE_ENDING)

                    try:
                        #Lo salvo y subo a la carpeta del dia/hora
                        #f=open(nombreSalida,"w", encoding="utf-8-sig")
                        f=open(nombreSalida,"w", encoding="ISO-8859-1")
                        f.write(contenido)

                        with open(nombreSalida, "rb") as f:
                            uploaded_file = target_folder_Salida.files.create_upload_session(
                            f, size_chunk, print_upload_progress
                            ).execute_query()

                        f.close()             

                        print("File {0} subido correctamente".format(uploaded_file.serverRelativeUrl))

                        #Lo salvo y subo a la carpeta de automatismos
                        #f=open(nombreAutomatismos,"w", encoding="utf-8")
                        f=open(nombreAutomatismos,"w", encoding="ISO-8859-1")
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
                    
