import sys

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential

#Sharepoint URL y usuario
username = "robotFicheros@totemtowersspain.es"
password = "R0b0tF1ch3r0s"
Sharepoint_destino = "https://totemtowersspain.sharepoint.com/sites/Prueba_QLIK"
directorio_destino = "Documentos compartidos"
size_chunk = 1000000

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

def print_upload_progress(offset):
    print ('Progreso ' + offset)
    pass

test_user_credentials = UserCredential(username,password)

ctx = ClientContext(Sharepoint_destino).with_credentials(test_user_credentials)

#creo el directorio en Sharepoint
folder = (
    ctx.web.default_document_library()
    .root_folder.folders.add_using_path(directorio_destino, overwrite=True)
    .execute_query()
)

target_folder_Salida = ctx.web.get_folder_by_server_relative_url(directorio_destino)

with open(nombre_fichero, "rb") as f:
    uploaded_file = target_folder_Salida.files.create_upload_session(
    f, size_chunk, print_upload_progress
    ).execute_query()

f.close()