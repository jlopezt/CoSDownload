import os
import json

from google.cloud import storage

# Configuracion - Datos globales
bucket_name = "tsp-cos-inpudata"

# Set the path to your service account key file
keyfile_path = 'keyfile/tgr-d4t-securestorage-dev-98a82faad55a.json' #'/path/to/keyfile.json'

def list_files_in_bucket(bucket_name):
    """Lists all files in a Google Cloud Storage bucket."""
    storage_client = storage.Client()
    bucket = storage_client.bucket(bucket_name)

    blobs = bucket.list_blobs()

    print(f"Files in {bucket_name} bucket:")
    for blob in blobs:
        print(blob.name)

# Set the GOOGLE_APPLICATION_CREDENTIALS environment variable
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = keyfile_path

list_files_in_bucket(bucket_name)
