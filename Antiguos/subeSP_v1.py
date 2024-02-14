"""
Demonstrates how to upload large file
"""

import os
import sys

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

username = "robotFicheros@totemtowersspain.es"
password = "R0b0tF1ch3r0s"
team_site_url = "https://totemtowersspain.sharepoint.com/sites/Prueba_QLIK"

#Argumentos:
#   SharePointUpload [nombre_fichero] [sharepoint_destino] [directorio_destino]
if (len (sys.argv) >= 2): 
    nombre_fichero = sys.argv[1];
    if (len (sys.argv) >= 3): 
        Sharepoint_destino = sys.argv[2];
        if (len (sys.argv) >= 4): 
            directorio_destino = sys.argv[3];
else:
    print('SharePointUpload [nombre_fichero] [sharepoint_destino] [directorio_destino]')


user_credentials = UserCredential(username,password)

def print_upload_progress(offset):
    # type: (int) -> None
    file_size = os.path.getsize(local_path)
    print(
        "Uploaded '{0}' bytes from '{1}'...[{2}%]".format(
            offset, file_size, round(offset / file_size * 100, 2)
        )
    )

ctx = ClientContext(team_site_url).with_credentials(user_credentials)

target_url = "Documentos compartidos/test"
target_folder = ctx.web.get_folder_by_server_relative_url(target_url)
size_chunk = 1000000
local_path = "C:\desarrollo\python\CoSDownload\Salida\equipamientoTotem.csv"

with open(local_path, "rb") as f:
    uploaded_file = target_folder.files.create_upload_session(
        f, size_chunk, print_upload_progress
    ).execute_query()

print("File {0} has been uploaded successfully".format(uploaded_file.serverRelativeUrl))

