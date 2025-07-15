# ASV-exporter

Export-Skript für ASV nach WebUntis, AD, moodle und O365.

## Workflow
Das beigefügte Export-Format (```.exf```) kann direkt verwendet werden. Damit werden alle aktiven Schüler in ASV exportiert und die csv in ein 
vorgegebenes Verzeichnis gespeichert. In der aktuellen Version ```/home/svp/schuelerdaten/export.csv```. Dieses Verzeichnis kann beispielsweise
über einen Mount von einem Windows-Server auf einer Linux-Maschine im Verwaltungsnetz eingebunden sein.

Da das Skript beim Aufruf die Aktualität der Konvertierten Dateien prüft und so kaum Ressourcen zieht, kann es minütlich per Cronjob ausgeführt werden:

```
# m h  dom mon dow   command
* * * * * /home/svp/bin/asv_convert.py -q
````
