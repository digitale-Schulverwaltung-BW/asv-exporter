#!/bin/bash

base="/home/svp/bin/asv/"
id_xref="/home/svp/bin/asv/ID.csv"
dup_xref="/home/svp/bin/asv/DUP.csv"
dstdir="ASV-Export"
webuntis="$dstdir/import_asv_nach_webuntis.csv"
perpustakaan="$dstdir/perpustakaan.csv"
#novell="$dstdir/import_asv_novell.csv"
netman="$dstdir/import_asv_nach_netman.csv"
moodle="$dstdir/import_asv_nach_moodle.csv"
m365dir="$dstdir/m365"
fehlzeiten="/var/www/fehlzeiten/uploads/ausbilder_Emails.csv"
ausbilder_stammdaten="$dstdir/ausbilder.csv"
post_cmd="" # umount /home/svp/bin/asv/mnt/"
sourcepath="/home/svp/bin/asv/mnt/export.csv"
destdir="/home/svp/bin/asv/mnt/"

exclude="1BFE0 1BFI0 2BFE0 2BFE0_Absage E1BT0 E1EG0 E1FI0 E1FI0_unklar E1FS0 E1GS0 E1IT0 E1ME0 E1RF0 E2FI0 E2RF0 E2EG0 \
FTE0_DAT FTE0_ENT FTET0_ENT FTET0_DAT E2EG0_unklar E_unklar Papierkorb FEEG_WL FEET2_alt 20_FTE0_ENT"

exclude_re='y_.*|.*VOR|.*VORZ|.*VZ'
infile="export_$$.csv"

cd $base
if [ "$1" = "-q" ]
then
	q="1"
else
	q=""
fi
[ -d /home/svp/bin/asv/mnt/alt ] || mount /home/svp/bin/asv/mnt/

if [ ! $webuntis -ot $sourcepath ]
then
	[ $q ] || echo "Up-to-date. Nothing to export."
	exit 0
fi

cp $sourcepath $infile
# WebUntis
#
#              shortname: login
#     Schlüssel(extern): login für bestehende, sonst ASV-ID
echo "login;shortname;idnumber;lastname;firstname;email;Klasse;birthday;Austrittsdatum" > $webuntis
# Moodle:
echo "lastname;firstname;course1;username;mail" > $moodle
# Novell:
# echo "username;idnumber;lastname;firstname;email;course1;city;lang" > $novell
# Netman:
echo "Username;Schueler_id;Name;Vorname;Passwort;Klasse;Geburtstag" > $netman
# FeMeSy:
echo "ID;ausbilderMail" > $fehlzeiten

rm -f $m365dir/*

declare -A idxlate
[ $q ] || echo -n "Reading old IDs..."
while IFS=';' read -ra line
do
        if [ -z $firstline ]; then firstline="1"; continue; fi
        b="${line[0]}"
        a="${line[1]}"
        if [ ! -z "$a" ]; then
                if [ ! -z "$b" ]; then
                asvid="$b"
                idxlate[$asvid]="$a"
                fi
        fi
done < $id_xref
[ $q ] || echo "done."

# quick access to the excluded list via assoc. array:
declare -A exc
for c in $exclude
do
        exc[$c]=1
done

declare -A duplate
[ $q ] || echo -n "Reading duplicate ID remappings..."
while IFS=';' read -ra line
do
        b="${line[0]}"
        a="${line[1]}"
        if [ ! -z "$a" ]; then
                if [ ! -z "$b" ]; then
                dupid="$b"
                duplate[$dupid]="$a"
                fi
        fi

done < $dup_xref

# format year for O365 accounts. Exports beginning Aug 1st will be in that year, before the year minus 1
oyear=$(date +%y)
if [ $(date +%m) -lt 8 ]; then ((oyear--)); fi
oyear=$( printf "%02d\n" $oyear)
[ $q ] || echo "Office365 year: $oyear"

firstline=""

lines=$( cat $infile | wc -l )
[ $q ] || echo "converting: "
c=0
while IFS=';' read -ra line
do
	if [ -z $firstline ]; then firstline="1"; continue; fi
	id="${line[0]}"
	id=$(echo $id|sed 's/"//g')
	idn=$(echo $id | sed 's/..*[-/]\([0-9a-f]*\)$/\1/')
	if [ ! -z "${duplate[$id]}" ]; then
		[ $q ] || echo "DUP: $id=>${duplate[$id]}"
		idn=${duplate[$id]}
	fi
	idx="asv$idn"
	svpid="${idxlate[$id]}"
	name="${line[1]}"
	firstname="${line[2]}"
	class="${line[6]}"
	pclass=$( echo $class | sed -e 's/"//g' )
	nclass=$( echo $pclass | sed -e 's-/-_-')
	oclass=$( echo "${oyear}_$pclass" | sed 's-/-_-g' )
	if [[ ${exc[$pclass]} ]]; then continue; fi
	email="${line[12]}"
	ausbilder="${line[22]}"
	birthday="${line[4]}"
	pass="HHS-$birthday"
	exitd="${line[9]}"
	nname=$(echo $name | sed -e 's/"//g' -e's/ä/ae/g' -e's/ö/oe/g' -e's/ü/ue/g' -e's/Ä/Ae/g' \
	-e's/Ö/Oe/g' -e's/Ü/Ue/g' -e's/ß/ss/g' | LANG=C sed  's/[\d128-\d255]/_/g')
	username=$(echo "$nname"|sed -e 's/\(....\).*/\1/' | tr [[:upper:]] [[:lower:]])
	if [ -n "$svpid" ]; then
		username=$(printf "%-4.4s_$svpid" "$username"| sed 's/ /_/g')
		id=$svpid
	else
		username=$(printf "%-4.4s-$idn" "$username"| sed 's/ /_/g')
		#id=$idn
	fi
	if [ "$email" = '""' ]; then semail="\"$username@stud.hhs.karlsruhe.de\""; else semail=$email; fi
	echo "\"$username\";\"$username\";\"$id\";$name;$firstname;$email;$class;$birthday;$exitd" >> $webuntis
	echo "$name;$firstname;$class;$username;$semail" >> $moodle
#	echo "$username;$id;$name;$firstname;$semail;$nclass;;" | sed -e 's/"//g' >> $novell
	echo "$username;$id;$name;$firstname;$pass;$class;$birthday" >> $netman
	echo "$username@stud.hhs.karlsruhe.de,$username,$oclass,$username.$oclass,,$oclass,,,,,,,,,," >> $m365dir/$oclass.csv
	if [ "$svpid" ]
	then
		echo "$id;$ausbilder" >> $fehlzeiten
	else
		echo "$username;$ausbilder" >> $fehlzeiten
	fi
	[ $q ] || echo -ne "$c/$lines   \r"
	((c++))
done < $infile
[ $q ] || echo ""
#iconv -f UTF-8 -t ISO-8859-1 $netman > $netman.dos
#mv $netman.dos $netman
for f in $m365dir/*
do
	sed -i '1s/^/Benutzername,Vorname,Nachname,Anzeigename,Position,Abteilung,Telefon – Geschäftlich,Telefon (geschäftlich),Mobiltelefon,Fax,Alternative E-Mail-Adresse,Adresse,Ort,Bundesstaat,Postleitzahl,Land oder Region\n/' $f
	if [ $q ]; then unix2dos $f > /dev/null 2>&1; else unix2dos $f; fi
	[ $q ] || echo -e "\033[2A"
done
files=$(ls $m365dir/*| wc -l)
duplicates=$( cat $webuntis |awk -F ';' '{print $1}' |sort | uniq -c | grep -v " 1 " )
if [ "$duplicates" ]
then
	echo "WARNUNG: doppelte Logins:"
	echo $duplicates
fi
cp $webuntis $destdir
[ $q ] || echo "Perpustakaan export."
cp $webuntis $perpustakaan
grep -v '""' $fehlzeiten > $ausbilder_stammdaten.2
grep -ve '^[0-9]' $ausbilder_stammdaten.2 > $ausbilder_stammdaten
rm $ausbilder_stammdaten.2
cut -d';' -f1,2,4- $perpustakaan  > $perpustakaan.2
sed -e 's/^login;/id;/' -e 's/shortname/Ausweisnummer/' -e 's/lastname/Familienname/' -e 's/firstname/Vornamen/' -e 's/birthday/Geburtstag/' $perpustakaan.2 > $perpustakaan
rm $perpustakaan.2
mv $perpustakaan $destdir
cp $ausbilder_stammdaten $destdir


((files=files+4))
i=1
if [ "$1" != "-n" -a "$2" != "-n" ]
then
  [ $q ] || echo ""
  [ $q ] || echo ""
  [ $q ] || echo -ne "Uploading to lcloud... \r"
  cd $dstdir
  rm -f ../ASV-Export.zip
  zip ../ASV-Export.zip -qr *
  scp -P 52022 ../ASV-Export.zip asv-export@lcloud.hhs.karlsruhe.de:/tmp/
  [ $q ] || echo "Upload to lcloud done.            "
  [ $q ] || echo ""
fi
cp -a $sourcepath $sourcepath.last

[ $q ] || echo "Unmounting source."
$post_cmd

[ $q ] || echo "Done."
rm ../$infile
