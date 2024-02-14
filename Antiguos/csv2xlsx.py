#pip install XlsxWriter

import sys
import os
import glob
import csv
from xlsxwriter.workbook import Workbook

path_origen='.'
nombre_fichero='*.csv'
nombre_hoja='Hoja1'

#Argumentos:
#   csv2xlsx [path_origen] [nombre_fichero] [nombre_hoja]
if (len (sys.argv) >= 2): 
    path_origen = sys.argv[1];
    if (len (sys.argv) >= 3): 
        nombre_fichero = sys.argv[2];
        if (len (sys.argv) >= 4): nombre_hoja = sys.argv[3];
else:
    print('csv2xlsx [path_origen] [nombre_fichero] [nombre_hoja]')

#for csvfile in glob.glob(os.path.join('.', '*.csv')):
for csvfile in glob.glob(os.path.join(path_origen, nombre_fichero)):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    worksheet = workbook.add_worksheet(nombre_hoja)
    with open(csvfile, 'rt', encoding='utf8') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
    workbook.close()