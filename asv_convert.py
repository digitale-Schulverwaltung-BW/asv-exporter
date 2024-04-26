#!/usr/bin/env python3

import subprocess
import argparse
import os
import shutil
import sys
import datetime
import re

base = "/home/svp/test"
id_xref = "/home/svp/test/ID.csv"
dup_xref = "/home/svp/test/DUP.csv"
dstdir = f"{base}/ASV-Export"
webuntis = os.path.join(dstdir, "import_asv_nach_webuntis.csv")
perpustakaan = os.path.join(dstdir, "perpustakaan.csv")
netman = os.path.join(dstdir, "import_asv_nach_netman.csv")
moodle = os.path.join(dstdir, "import_asv_nach_moodle.csv")
cloudstack = os.path.join(dstdir, "import_asv_nach_cloudstack.csv")
m365dir = os.path.join(dstdir, "m365")
EFTdir = os.path.join(dstdir, "EFT")
ausweisdir = os.path.join(dstdir, "schuelerausweise")
fehlzeiten = f"{base}/ausbilder_Emails.csv" #"/var/www/fehlzeiten/uploads/ausbilder_Emails.csv"
ausbilder_stammdaten = os.path.join(dstdir, "ausbilder.csv")
post_cmd = ""  # umount /home/svp/bin/asv/mnt/"
sourcepath = f"{base}/mnt/export.csv"
destdir = "/home/svp/test/mnt"

exclude = ['1BFE0', '1BFI0', '2BFE0', '2BFE0_Absage', 'E1BT0', 'E1EG0', 'E1FI0', 'E1FI0_unklar', 
           'E1FS0', 'E1GS0', 'E1IT0', 'E1ME0', 'E1RF0', 'E2FI0', 'E2RF0', 'E2EG0', 'FTE0_DAT', 
           'FTE0_ENT', 'FTET0_ENT', 'FTET0_DAT', 'E2EG0_unklar', 'E_unklar', 'Papierkorb', 
           'FEEG_WL', 'FEET2_alt', '20_FTE0_ENT']
exclude_re = 'y_.*|.*VOR|.*VORZ|.*VZ'
python_pid=os.getpid()
infile = f"export_{python_pid}.csv"

def help_message():
    help='Convert ASV export to several import CSVs\n'
    return help

os.chdir(base)

parser = argparse.ArgumentParser(description=help_message(), formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('-q', '--quiet', action="store_true", help='quiet (non-interactive) operation', dest='q')
parser.add_argument('-n', '--noupload', action="store_true", help='no uploads to LCloud/cloudstack', dest='n')
args = parser.parse_args()

q = args.q  # quiet flag

if not os.path.isdir(f"{destdir}/alt"):
    subprocess.run(["mount", "/home/svp/bin/asv/mnt/"])

if os.path.isfile(webuntis) and os.path.getmtime(webuntis) >= os.path.getmtime(sourcepath):
    if not q: print("Up-to-date. Nothing to export.")
    exit(0)

shutil.copy(sourcepath, infile)

# Initialize dictionaries
idxlate = {}
duplate = {}

with open(id_xref, 'r') as f:
    next(f)  # Skip header
    for line in f:
        if line.strip():
            b, a = line.strip().split(';')
            if a and b:
                asvid = b
                idxlate[asvid] = a

with open(dup_xref, 'r') as f:
    for line in f:
        if line.strip():
            b, a = line.strip().split(';')
            if a and b:
                dupid = b
                duplate[dupid] = a

# Format year for O365 accounts. Exports beginning Aug 1st will be in that year, before the year minus 1
oyear = int(datetime.datetime.now().strftime("%y"))
if int(datetime.datetime.now().strftime("%m")) < 8:
    oyear -= 1
oyear = "{:02d}".format(oyear)

firstline = ""
lines = sum(1 for line in open(infile))

webfile       =open(webuntis, 'w')
moodlefile    =open(moodle, 'w')
netmanfile    =open(netman, 'w')
cloudfile     =open(cloudstack, 'w')
fehlzeitenfile=open(fehlzeiten, 'w')
####
# WebUntis
#
#             shortname: login
#     Schlüssel(extern): login für bestehende, sonst ASV-ID
webfile.write("login;shortname;idnumber;lastname;firstname;email;Klasse;birthday;Austrittsdatum;Eintrittsdatum\n")
# Moodle:
moodlefile.write("lastname;firstname;course1;username;mail\n")
# Netman:
netmanfile.write("Username;Schueler_id;Name;Vorname;Passwort;Klasse;Geburtstag\n")
# Cloudstack:
cloudfile.write("Username;Klasse;Geburtstag\n")
# FeMeSy:
fehlzeitenfile.write("ID;ausbilderMail\n")

os.system(f"rm {EFTdir}/*")
os.system(f"rm {m365dir}/*")
os.system(f"rm {ausweisdir}/*")

print("converting:")
c = 0
with open(infile, 'r') as f:
    for line in f:
        line = line.strip().split(';')
        if not firstline:
            firstline = "1"
            continue
        
        id_ = line[0].replace('"', '')

        asvid = id_
        idn = re.search('..*[-]([0-9a-f]*)$', id_).group(1)

        if id_ in duplate:
            print(f"DUP: {id_}=>{duplate[id_]}")
            idn = duplate[id_]
        idx = "asv" + idn
        svpid = idxlate.get(id_, None)
        name = line[1].replace('"', '')
        firstname = line[2].replace('"', '')
        class_ = line[6].replace('"', '')
        pclass = class_.replace('"', '')
        nclass = pclass.replace('/', '_')
        oclass = f"{oyear}_{pclass}".replace('/', '_')
        m365file      =open(f'{m365dir}/{oclass}.csv', 'a')
        EFTfile       =open(f'{EFTdir}/{oclass}.csv', 'a')
        ausweisfile   =open(f'{ausweisdir}/{nclass}.csv', 'a')

        if pclass in exclude:
            continue
        if re.match(exclude_re, pclass):
            continue
        email = line[31].replace('"', '')
        ausbilder = line[32].replace('"', '')
        if len(line)>33:
                ausbilder2 = line[33].replace('"', '').replace(' ', '')  
        else:   ausbilder2=""

        birthday = line[4].replace('"', '')
        pass_ = f"HHS-{birthday}"
        entryd = line[8].replace('"', '')
        exitd = line[9].replace('"', '')

        special_char_map = {ord('ä'):'ae', ord('ü'):'ue', ord('ö'):'oe', 
                            ord('Ä'):'Ae', ord('Ü'):'Ue', ord('Ö'):'oe',ord('ß'):'ss'}
        nname = re.sub(r'["äöüÄÖÜß]', lambda m: m.group().translate(special_char_map), name)
        nname = re.sub(r'[^\x00-\x7F]', '_', nname)
        # legacy bug. Delete the next 2 lines in 2026/2027
        if id_=="8a9041a0-89487389-0189-49e60e16-2fa1" and nname[:3]=="O_D":
            nname="o__d"
        
        username = nname[:4].lower()

        if svpid:
            username = ('{:<4}'.format(username[:4])+f"_{svpid}").replace (' ', '_')
            #username = f"{username[:4]}_{svpid}".replace(' ', '_')
            id_ = svpid
            asvid = svpid
        else:
            username = ('{:<4}'.format(username[:4])+f"-{idn}").replace (' ', '_')
            #username = f"{username[:4]}-{idn}".replace(' ', '_')

        semail = f'"{username}@stud.hhs.karlsruhe.de"' if email == '""' else email
        webfile.write    (f'"{username}";"{username}";"{id_}";"{name}";"{firstname}";"{email}";"{class_}";"{birthday}";"{exitd}";"{entryd}"\n')
        moodlefile.write (f'{name};{firstname};{class_};{username};{semail}\n')
        netmanfile.write (f'{username};{id_};{name};{firstname};{pass_};{class_};{birthday}\n')
        cloudfile.write  (f'{username};{class_};{birthday}\n')
        m365file.write   (f'{username}@stud.hhs.karlsruhe.de,{username},{oclass},{username}.{oclass},,{oclass},,,,,,,,,,\n')
        EFTfile.write    (f'"{name}","{firstname}",{username},,,\n')
        ausweisfile.write(f'"{name}";"{firstname}";"{class_}";"{birthday}";{line[23]};{line[24]};{line[25]};{line[26]};{line[27]};{line[28]};{line[29]};{line[5]};{line[22]};{line[8]};{line[9]};{username}\n')
        fehlzeitenfile.write(f'{asvid};{ausbilder}\n')
        if ausbilder2:
            fehlzeitenfile.write(f'{asvid};{ausbilder2}\n')
        
        if not q: print(f"{c}/{lines}", end='\r')
        c += 1
for file in (webfile, moodlefile, netmanfile, cloudfile, m365file, ausweisfile, fehlzeitenfile):
    file.close()

# Update file paths if necessary
for f in os.listdir(m365dir):
    file_path = os.path.join(m365dir, f)
    if os.path.isfile(file_path):
        with open(file_path, 'r+') as file:
            content = file.read()
            file.seek(0, 0)
            file.write('Benutzername,Vorname,Nachname,Anzeigename,Position,Abteilung,Telefon – Geschäftlich,Telefon (geschäftlich),Mobiltelefon,Fax,Alternative E-Mail-Adresse,Adresse,Ort,Bundesstaat,Postleitzahl,Land oder Region\n' + content)
            file.close()
        subprocess.run(['unix2dos', file_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        #if not q:
        #    print("\033[2A")

for f in os.listdir(EFTdir):
    file_path = os.path.join(EFTdir, f)
    if os.path.isfile(file_path):
        with open(file_path, 'r+') as file:            
            content = file.read()
            file.seek(0, 0)
            file.write('Nachname,Vorname,Username,Bemerkungen,Anw,Foto\n' + content)
            file.close()
        subprocess.run(['unix2dos', file_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        #if not q:
        #    print("\033[2A")

for f in os.listdir(ausweisdir):
    file_path = os.path.join(ausweisdir, f)
    if os.path.isfile(file_path):
        with open(file_path, 'r+') as file:
            content = file.read()
            file.seek(0, 0)
            file.write('Familienname;Rufname;Klasse;Geburtsdatum;Geburtsort;Schüler/in Strasse;Schüler/in Hausnummer;Schüler/in Anschrift PLZ;Schüler/in Anschrift Ort;Schüler/in Ortsteil der Anschrift;Bildungsgang Stichtag.Langform;Bildungsgang Stichtag.Kurzform;Anrede.Langform;Eintritt in diese Schule am;voraussichtlicher Austritt am;ID\n' + content)
            file.close()
        # Remove double quotes from the file if needed
        # content = content.replace('"', '')
        # with open(file_path, 'w') as file:
        #     file.write(content)

# Perform further processing and actions (not directly translatable to this context)
files = len(os.listdir(m365dir))
# duplicates = ...  # Logic to check duplicates
duplicates=subprocess.getoutput(f"/bin/cat {webuntis} |awk -F ';' '"+"{print $1}' |sort | uniq -c | grep -v ' 1 '" )

if duplicates:
    print("WARNUNG: doppelte Logins:")
    print(duplicates)
os.system(f"rm {infile}")

# Copy files to destination directories
os.system(f"cp {webuntis} {destdir}")
print("Perpustakaan export.")
os.system(f"cp {webuntis} {perpustakaan}")

# Modify and move files
os.system(f"grep -ve ';$' {fehlzeiten} > {ausbilder_stammdaten}.2")
os.rename(f"{ausbilder_stammdaten}.2", ausbilder_stammdaten)
os.system(f"cut -d';' -f1,2,4- {perpustakaan}  > {perpustakaan}.2")
with open(f"{perpustakaan}.2", 'r+') as file:
    content = file.read()
    content=content.replace("login;shortname;lastname;firstname;email;Klasse;birthday;Austrittsdatum;Eintrittsdatum",
                             "id;Ausweisnummer;Familienname;Vornamen;email;Klasse;Geburtstag;Austrittsdatum;Eintrittsdatum")
    file.seek(0,0)
    file.write(content)

os.rename(f"{perpustakaan}.2", perpustakaan)
os.system(f"mv {perpustakaan} {destdir}")
os.system(f"cp {ausbilder_stammdaten} {destdir}")
os.system(f"rm -rf {destdir}/schuelerausweise/*")
os.system(f"cp {ausweisdir}/* {destdir}/schuelerausweise/")

files += 4

# Additional logic and actions (not directly translatable to this context)
if not args.n:
    if not q: print("")
    if not q: print("")
    if not q: print("Uploading to lcloud...")
    os.chdir(dstdir)
    os.system("rm -f ../ASV-Export.zip")
    os.system("zip ../ASV-Export.zip -qr *")
    os.system(f"scp -P 52022 ../ASV-Export.zip asv-export@lcloud.hhs.karlsruhe.de:/tmp/")
    if not q: print("Upload to lcloud done.")
    if not q: print("")
    if not q: print("Upload to cloudstack...")
    os.system(f"scp ../{cloudstack} administrator@192.168.201.204:/usr/local/share/openstack-data")

if not q: print("Unmounting source.")
# Execute post command and additional actions (not directly translatable to this context)

if not q: print("Done.")
