# ASV-exporter

Export-Skript für ASV nach WebUntis, AD, moodle und O365.

## Workflow
Das beigefügte Export-Format (```.exf```) kann direkt verwendet werden. Damit werden alle aktiven Schüler in ASV exportiert und die csv in ein 
vorgegebenes Verzeichnis gespeichert. In der aktuellen Version ```/home/svp/schuelerdaten/export.csv```. Dieses Verzeichnis kann beispielsweise
über einen Mount von einem Windows-Server auf einer Linux-Maschine im Verwaltungsnetz eingebunden sein.

Da das Skript beim Aufruf die Aktualität der konvertierten Dateien prüft und so kaum Ressourcen zieht (Laufzeit weniger als 0,1s wenn keine Import-Datei vorhanden, ein Import-Lauf benötigt ca. 1,3s bei 1900 Schülerdatensätzen im quiet-Modus), kann es minütlich per Cronjob ausgeführt werden:

```
# m h  dom mon dow   command
* * * * * /home/svp/bin/asv_convert.py -q
````

## Parameter
```-q``` Quiet mode, keine Ausgaben. Ohne diesen Switch zeigt das Skript den Fortschritt beim Konvertieren und die Uploads an.

```-n``` No uploads. Keine Upload-Kommandos werden ausgeführt.

```-h``` help
