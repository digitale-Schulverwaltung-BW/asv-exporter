#!/usr/bin/env python3
# Ubuntu 24.04: sudo apt install python3-xlsxwriter
# before: pip3 install XlsxWriter

import xlsxwriter
import subprocess
import argparse
import os
import shutil
import datetime
import re

base = "/home/svp/schuelerdaten"                # base directory
id_xref = "/home/svp/schuelerdaten/ID.csv"      # ASV-ID to SVP-ID cross reference. Set to "" if not used.
dup_xref = "/home/svp/schuelerdaten/DUP.csv"    # duplicate ASV-ID cross reference. 
dstdir = f"{base}/ASV-Export"                   # destination directory for export files

if not os.path.isdir(dstdir):
    try:
        os.makedirs(dstdir)    
    except OSError as e:
        print(f"Error creating directory {dstdir}: {e}")
        exit(1)
# Create subdirectories if they do not exist
# remove any unwanted subdirectories from the list
exports = ["WebUntis", "m365", "EFT", "schuelerausweise", "Buecherlisten", "cloudstack"]
for subdir in exports:
    subdir_path = os.path.join(dstdir, subdir)
    if not os.path.isdir(subdir_path):
        try:
            os.makedirs(subdir_path)
        except OSError as e:
            print(f"Error creating directory {subdir_path}: {e}")
            exit(1) 

## export files
## 
# Note: currently, the WebUntis export is mandatory!
webuntis = os.path.join(dstdir, "import_asv_nach_webuntis.csv")
webuntis_dst = os.path.join(dstdir+"WebUntis", "import_asv_nach_webuntis.csv")
perpustakaan = os.path.join(dstdir, "perpustakaan.csv") if "Buecherlisten" in exports else ""
netman = os.path.join(dstdir, "import_asv_nach_netman.csv") if "netman" in exports else ""
moodle = os.path.join(dstdir, "import_asv_nach_moodle.csv") if "moodle" in exports else ""
cloudstack = os.path.join(dstdir+"/cloudstack", "import_asv_nach_cloudstack.csv") if "cloudstack" in exports else ""
m365dir = os.path.join(dstdir, "m365") if "m365" in exports else ""             # one file per class
EFTdir = os.path.join(dstdir, "EFT") if "EFT" in exports else ""                # one file per class
buecherlistendir = os.path.join(dstdir, "Buecherlisten") if "Buecherlisten" in exports else ""  # one file per class
ausweisdir = os.path.join(dstdir, "schuelerausweise") if "schuelerausweise" in exports else ""  # one file per class

# student email address. Leave blank if no school-wide email addresses are used. Email addresses in the export file will take precedence.
maildomain="@stud.hhs.karlsruhe.de"

# initial password for students. This is used in the Netman export and must be changed on first login
initpass = "HHS-"  

# upload commands
upload_cmd="scp -P 52022 ../ASV-Export.zip asv-export@lcloud.hhs.karlsruhe.de:/tmp/"
upload_cloud=f"scp {cloudstack} administrator@192.168.201.204:/usr/local/share/openstack-data"

# special exports
fehlzeiten = "/var/www/fehlzeiten/uploads/ausbilder_Emails.csv"
ausbilder_stammdaten = os.path.join(dstdir, "ausbilder.csv")

# command to run after export
#
post_cmd = ""  # umount /home/svp/bin/asv/mnt/"

# The source file lives here
sourcepath = f"{base}/mnt/export.csv"

# after successful run, all exports are copied to this directory
# note: after each run, the current exports will first be moved to the subdir "alt/", removing its contents.
# That subdir 
destdir = "/home/svp/schuelerdaten/mnt"

# Exclude the following classes. First, a list of classes to exclude, then a regex for classes to exclude.
# The regex is used to exclude classes that start with 'y_' or contain 'VOR
exclude = ['1BFE0', '1BFI0', '2BFE0', '2BFE0_Absage', 'E1BT0', 'E1EG0', 'E1FI0', 'E1FI0_unklar', 
           'E1FS0', 'E1GS0', 'E1IT0', 'E1ME0', 'E1RF0', 'E2FI0', 'E2RF0', 'E2EG0', 'FTE0_DAT', 
           'FTE0_ENT', 'FTET0_ENT', 'FTET0_DAT', 'E2EG0_unklar', 'E_unklar', 'Papierkorb', 
           'FEEG_WL', 'FEET2_alt', '20_FTE0_ENT']
exclude_re = 'y_.*|.*VOR|.*VORZ|.*VZ'

## import file format specifications. The numbers correspond to the columns in the CSV file.
## matching ASV export file format included in the repository.

CSV_ID              = 0
CSV_NAME            = 1
CSV_FIRSTNAME       = 2
CSV_GENDER          = 3
CSV_CLASS           = 6
CSV_BIRTHDAY        = 4
CSV_ENTRYD          = 8
CSV_EXITD           = 9
CSV_BIRTHPLACE      = 23
CSV_STREET          = 24
CSV_NO              = 25
CSV_ZIP             = 26
CSV_TOWN            = 27
CSV_TOWN2           = 28
CSV_EDUCATION       = 29
CSV_EDUCATION_SHORT = 5
CSV_SALUTATION      = 22
CSV_EMAIL           = 31
CSV_VOLLJAEHRIG     = 32
CSV_AUSBILDER       = 33
CSV_AUSBILDER2      = 34

MIN_LINES = 10           # sanity check: quit if fewer than this lines in input file (errorneous export?)

##############################

# Thank you for reading. No user serviceable parts below. :)

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
    subprocess.run(["mount", destdir])
    if not os.path.isdir(f"{destdir}/alt"):
        try:
            os.makedirs(f"{destdir}/alt")
        except OSError as e:
            print(f"Error creating backup directory {destdir}/alt: {e}")
            exit(1)

if not os.path.isfile(sourcepath) or ( os.path.isfile(webuntis_dst) and os.path.getmtime(webuntis_dst) >= os.path.getmtime(sourcepath)):
    if not q: print("Up-to-date. Nothing to export.")
    exit(0)

num_lines = sum(1 for _ in open(sourcepath, 'rb'))
if (num_lines<MIN_LINES):
    print ("Too few lines in source file, exiting")
    exit(1)

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

webfile        = open(webuntis, 'w')   if "webuntis" in exports else None
moodlefile     = open(moodle, 'w')     if "moodle" in exports else None
netmanfile     = open(netman, 'w')     if "netman" in exports else None
cloudfile      = open(cloudstack, 'w') if "cloudstack" in exports else None
fehlzeitenfile = open(fehlzeiten, 'w') if fehlzeiten else None
####
# WebUntis
#
#             shortname: login
#     Schlüssel(extern): login für bestehende, sonst ASV-ID
if webfile!=None: webfile.write("login;shortname;idnumber;lastname;firstname;email;Klasse;birthday;Austrittsdatum;Eintrittsdatum;volljaehrig\n") 
# Moodle:
if moodlefile!=None: moodlefile.write("lastname;firstname;course1;username;mail\n")
# Netman:
if netmanfile!=None: netmanfile.write("Username;Schueler_id;Name;Vorname;Passwort;Klasse;Geburtstag\n")
# Cloudstack:
if cloudfile!=None: cloudfile.write("Username;Klasse;Geburtstag\n")
# FeMeSy:
if fehlzeitenfile!=None: fehlzeitenfile.write("ID;ausbilderMail\n")

if EFTdir:     os.system(f"rm {EFTdir}/*")
if m365dir:    os.system(f"rm {m365dir}/*")
if ausweisdir: os.system(f"rm {ausweisdir}/*")

if not q: print("converting:")
c = 0
with open(infile, 'r') as f:
    for line in f:
        line = line.strip().split(';')
        if not firstline:
            firstline = "1"
            continue
        
        id_ = line[CSV_ID].replace('"', '')

        asvid = id_
        idn = re.search('..*[-]([0-9a-f]*)$', id_).group(1)

        if id_ in duplate:
            if not q: print(f"DUP: {id_}=>{duplate[id_]}")
            idn = duplate[id_]
        idx = "asv" + idn
        svpid = idxlate.get(id_, None)
        name = line[CSV_NAME].replace('"', '')
        firstname = line[CSV_FIRSTNAME].replace('"', '')
        class_ = line[CSV_CLASS].replace('"', '')
        pclass = class_.replace('"', '')
        nclass = pclass.replace('/', '_')
        oclass = f"{oyear}_{pclass}".replace('/', '_')

        if pclass in exclude:
            continue
        if re.match(exclude_re, pclass):
            continue
        email       = line[CSV_EMAIL].replace('"', '')
        volljaehrig = line[CSV_VOLLJAEHRIG].replace('"', '')
        ausbilder   = line[CSV_AUSBILDER].replace('"', '')
        if len(line)>34:
                ausbilder2 = line[CSV_AUSBILDER2].replace('"', '').replace(' ', '')  
        else:   ausbilder2=""

        birthday = line[CSV_BIRTHDAY].replace('"', '')
        pass_ = f"{initpass}{birthday}"
        entryd = line[CSV_ENTRYD].replace('"', '')
        exitd = line[CSV_EXITD].replace('"', '')

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
        
        if maildomain:
            semail = f'"{username}{maildomain}"' if email == '""' else email
        else:
            semail = email

        if webfile!=None:    webfile.write    (f'"{username}";"{username}";"{id_}";"{name}";"{firstname}";"{email}";"{class_}";"{birthday}";"{exitd}";"{entryd}";"{volljaehrig}"\n')
        if moodlefile!=None: moodlefile.write (f'{name};{firstname};{class_};{username};{semail}\n')
        if netmanfile!=None: netmanfile.write (f'{username};{id_};{name};{firstname};{pass_};{class_};{birthday}\n')
        if cloudfile!=None:  cloudfile.write  (f'{username};{class_};{birthday}\n')
        if m365dir:
            m365file=open(f'{m365dir}/{oclass}.csv', 'a')
            m365file.write   (f'{username}{maildomain},{username},{oclass},{username}.{oclass},,{oclass},,,,,,,,,,\n')
        if EFTdir:
            EFTfile=open(f'{EFTdir}/{oclass}.csv', 'a')
            EFTfile.write (f'"{name}","{firstname}",{username},,,\n')
        if ausweisdir:
            ausweisfile   =open(f'{ausweisdir}/{nclass}.csv', 'a')
            ausweisalle   =open(f'{ausweisdir}/alle.csv', 'a')
            ausweisfile.write(f'"{name}";"{firstname}";"{class_}";"{birthday}";{line[CSV_BIRTHPLACE]};{line[CSV_STREET]};{line[CSV_NO]};{line[CSV_ZIP]};{line[CSV_TOWN]};{line[CSV_TOWN2]};{line[CSV_EDUCATION]};{line[CSV_EDUCATION_SHORT]};{line[CSV_SALUTATION]};{line[CSV_GENDER]};{line[CSV_ENTRYD]};{line[CSV_EXITD]};{username};{id_}\n')
            ausweisalle.write(f'"{name}";"{firstname}";"{class_}";"{birthday}";{line[CSV_BIRTHPLACE]};{line[CSV_STREET]};{line[CSV_NO]};{line[CSV_ZIP]};{line[CSV_TOWN]};{line[CSV_TOWN2]};{line[CSV_EDUCATION]};{line[CSV_EDUCATION_SHORT]};{line[CSV_SALUTATION]};{line[CSV_GENDER]};{line[CSV_ENTRYD]};{line[CSV_EXITD]};{username};{id_}\n')
        if fehlzeitenfile!=None: 
            fehlzeitenfile.write(f'{asvid};{ausbilder}\n')
            if ausbilder2:
                fehlzeitenfile.write(f'{asvid};{ausbilder2}\n')
        
        if not q: print(f"{c}/{lines}", end='\r')
        c += 1
for file in (webfile, moodlefile, netmanfile, cloudfile, m365file, ausweisfile, ausweisalle, fehlzeitenfile):
    if file!=None: file.close()

# Update file paths if necessary
if m365dir:
    # prepend first line in all m365 files
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

if EFTdir:
    # prepend first line in all EFT files
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

if ausweisdir:
    # prepend first line in all ausweis files
    for f in os.listdir(ausweisdir):
        file_path = os.path.join(ausweisdir, f)
        if os.path.isfile(file_path):
            with open(file_path, 'r+') as file:
                content = file.read()
                file.seek(0, 0)
                file.write('Familienname;Rufname;Klasse;Geburtsdatum;Geburtsort;Schüler/in Strasse;Schüler/in Hausnummer;Schüler/in Anschrift PLZ;Schüler/in Anschrift Ort;Schüler/in Ortsteil der Anschrift;Bildungsgang Stichtag.Langform;Bildungsgang Stichtag.Kurzform;Anrede.Langform;Geschlecht;Eintritt in diese Schule am;voraussichtlicher Austritt am;ID;UUID\n' + content)
                file.close()
            # Remove double quotes from the file if needed
            # content = content.replace('"', '')
            # with open(file_path, 'w') as file:
            #     file.write(content)

# Perform further processing and actions 
# files = len(os.listdir(m365dir))
# duplicates = ...  # Logic to check duplicates
# warning: this currently depends on the creation of the webuntis file.
if webfile!=None: 
    duplicates=subprocess.getoutput(f"/bin/cat {webuntis} |awk -F ';' '"+"{print $1}' |sort | uniq -c | grep -v ' 1 '" )

    if duplicates:
        print("WARNUNG: doppelte Logins:")
        print(duplicates)
    os.system(f"rm {infile}")

# Copy files to destination directories
if webuntis: os.system(f"cp {webuntis} {destdir}/WebUntis/")
if webuntis:
    if not q: print("Perpustakaan export.")
    if perpustakaan: os.system(f"cp {webuntis} {perpustakaan}")

# Modify and move files
if fehlzeiten: 
    os.system(f"grep -ve ';$' {fehlzeiten} > {ausbilder_stammdaten}.2")
    os.rename(f"{ausbilder_stammdaten}.2", ausbilder_stammdaten)
if perpustakaan: os.system(f"cut -d';' -f1,2,4- {perpustakaan}  > {perpustakaan}.2")
with open(f"{perpustakaan}.2", 'r+') as file:
    content = file.read()
    content=content.replace("login;shortname;lastname;firstname;email;Klasse;birthday;Austrittsdatum;Eintrittsdatum",
                             "id;Ausweisnummer;Familienname;Vornamen;email;Klasse;Geburtstag;Austrittsdatum;Eintrittsdatum")
    file.seek(0,0)
    file.write(content)


if ausweisdir:
    for f in os.listdir(ausweisdir):
        classname=(f.split('.'))[0]
        workbook = xlsxwriter.Workbook(f"{buecherlistendir}/{classname}.xlsx")
        worksheet = workbook.add_worksheet()
        worksheet.set_column(0,0,20)
        worksheet.set_column(1,1,14)
        worksheet.set_column(2,2,15)
        worksheet.set_column(3,3,10)
        worksheet.set_column(4,4,26)
        barcode = workbook.add_format()
        barcode.set_font_name('CCode39')
        bold = workbook.add_format()
        bold.set_bold()
        worksheet.set_row(0, 15, bold)
        header = ['Nachname', 'Vorname', 'Klasse', 'ID', 'Barcode']
        col=0
        for item in (header):
            worksheet.write(0, col, item)
            col += 1
        firstline = ""
        row=1
        with open(f"{ausweisdir}/{f}", 'r') as classfile:
            for line in classfile:
                if not firstline:
                    firstline = "1"
                    continue
                line = line.strip().split(';')
                name = line[0].replace('"', '')
                firstname = line[1].replace('"', '')
                id = line[16]
                nclass = line[2].replace('"', '')
                line = [name, firstname, nclass, id]
                col = 0
                for item in (line):
                    worksheet.write(row, col, item)
                    col += 1
                worksheet.write(row, col, f"=\"*\"&D{row+1}&\"*\"", barcode)
                row += 1
        workbook.close()

if perpustakaan: os.rename(f"{perpustakaan}.2", perpustakaan)
if perpustakaan: os.system(f"mv {perpustakaan} {destdir}/perpustakaan")
if ausbilder_stammdaten: os.system(f"cp {ausbilder_stammdaten} {destdir}/WebUntis")
if buecherlistendir: os.system(f"cp {buecherlistendir}/* {destdir}/Buecherlisten")
if ausweisdir:
    os.system(f"rm -rf {destdir}/schuelerausweise/*")
    os.system(f"cp {ausweisdir}/* {destdir}/schuelerausweise/")
os.system(f"mv {sourcepath} {destdir}/alt/ || mv {sourcepath} {destdir}/alt/$$.csv")

#files += 4

# Uploads if not suppressed
if not args.n:
    if not q: print("")
    if not q: print("")
    if not q: print("Uploading to lcloud...")
    os.chdir(dstdir)
    os.system("rm -f ../ASV-Export.zip")
    os.system("zip ../ASV-Export.zip -qr *")
    os.system(upload_cmd)
    if not q: print("Upload to lcloud done.")
    if not q: print("")
    if not q: print("Upload to cloudstack...")
    os.system(upload_cloud)

if not q: print("Unmounting source.")
if not q: print("Done.")
