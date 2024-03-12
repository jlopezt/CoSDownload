#!/bin/bash

cd /home/jlopezt/desarrollo/python/CoSDownload


echo ===========================================================================================  >> Salida/CoSDownload.log
echo >> Salida/CoSDownload.log


date >> Salida/CoSDownload.log

rm -rf Salida/*.csv

source venv/bin/activate
python3 CoSDownload_v3_7.py  >> Salida/CoSDownload.log
python3 CoSDownloadArbolito.py  >> Salida/CoSDownload.log 

#chmod 666 Salida/*.csv
#cp Salida/*.csv /home/leeFicheros/CoSDownload
deactivate

date >> Salida/CoSDownload.log

echo >> Salida/CoSDownload.log
echo ===========================================================================================  >> Salida/CoSDownload.log



