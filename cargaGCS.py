import os
import json

from google.cloud import storage

# Configuracion - Datos globales
bucket_name = "tsp-cos-inpudata"
local_directory = "Salida" #"/path/to/local/directory"
destination_directory = "D4T_Input" #"uploaded-files"

# Set the path to your service account key file
keyfile_path = 'keyfile/tgr-d4t-securestorage-dev-98a82faad55a.json' #'/path/to/keyfile.json'

# --Funciones--
def upload_to_gcs(bucket_name, local_directory, destination_directory):
    """Uploads all files in a local directory to a Google Cloud Storage bucket."""
    storage_client = storage.Client()
    bucket = storage_client.bucket(bucket_name)

    for local_file in os.listdir(local_directory):
        nombre_archivo, extension = os.path.splitext(local_file)
        #print("Nombre del fichero: " + nombre_archivo + extension)
        if(extension=='.csv'):
            local_file_path = os.path.join(local_directory, local_file)
            if os.path.isfile(local_file_path):
                dir=local_file.split('_',1)[0]
                #destination_blob_name = os.path.join(destination_directory, dir, local_file)
                destination_blob_name = destination_directory + '/' + dir + '/' + local_file
                blob = bucket.blob(destination_blob_name)
                #blob.delete()
                blob.upload_from_filename(local_file_path)
                print(f"File {local_file_path} uploaded to {destination_blob_name} in {bucket_name} bucket.")

# --MAIN--
# Set the GOOGLE_APPLICATION_CREDENTIALS environment variable
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = keyfile_path

upload_to_gcs(bucket_name, local_directory, destination_directory)
