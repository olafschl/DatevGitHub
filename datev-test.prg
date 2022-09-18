* Erstellen der Buchungsdatei
*
nhandle=FCREATE("c:\aktuell\ED00001")
IF nhandle==-1	
	WAIT WINDOW "Fehler!!!"
	EXIT
ENDIF
FWRITE(nhandle,CHR(29))			&& Vorlaufbeginn Konstante = X'1D
FWRITE(nhandle,CHR(24))			&& Kennung neuer Vorlauf Konstante X'18
FWRITE(nhandle,"1")				&& Versionsnummer Konstante 1
FWRITE(nhandle,"001")			&& Datenträgernummer aus EV01
FWRITE(nhandle,"11")			&& Anwendungsnummer 11=Buchungsdatei, 13=Debitorenstammsätze
FWRITE(nhandle,"os")			&& Namenskürzel
FWRITE(nhandle,"0004503")		&& Beraternummer
FWRITE(nhandle,"00002")			&& Mandantennummer
FWRITE(nhandle,"010105")		&& Abrechnungsnummer nnnnJJ
FWRITE(nhandle,"010105")		&& Datum von
FWRITE(nhandle,"280105")		&& Datum bis
FWRITE(nhandle,"001")			&& Primanotaseite
FWRITE(nhandle,"olix")			&& Passwort
FWRITE(nhandle,SPACE(16))		&& Anwendungsinfo Konstante 16 Leerzeichen
FWRITE(nhandle,SPACE(16))		&& Input-Info Konstante 16 Leerzeichen
FWRITE(nhandle,"y")				&& Satzende
FWRITE(nhandle,CHR(181))		&& Versionskennung X'B5  fest
FWRITE(nhandle,"1")				&& Versionsnummer Konstante
FWRITE(nhandle,",")				&& Trennzeichen
FWRITE(nhandle,"4")				&& Aufgezeichnete Sachkontenlänge 4-8 zulässig
FWRITE(nhandle,",")				&& Trennzeichen
FWRITE(nhandle,"4")				&& Gespeicherte Sachkontenlänge 4-8 zulässig
FWRITE(nhandle,",")				&& Trennzeichen
FWRITE(nhandle,"SELF")			&& Produktkürzel
FWRITE(nhandle,CHR(28))			&& Feldende X'1C
FWRITE(nhandle,"y")				&& Satzende
x=Fseek(nhandle,0,2)
* Hier erfolgt die Blockgröße auf 256Byte
X=FSEEK(nhandle,0,2)
sl=250-x
datensatz=""
summe=0
FOR i=1 TO 5
	hf="+56025"  					&& +560.25 Betrag
	summe=summe+560.25
	hf=hf+"a8400"					&& Gegenkonto / Erlöskonto 8400
	hf=hf+CHR(189)+"0520032"+CHR(28)&& X'BD Belegfeld 1 Rechnungsnummer X'1C
	hF=hf+CHR(190)+"5672"+CHR(28)	&& X'BE Belegfeld 2 alternativ
	hf=hf+"d401"					&& Datum 04.01.
	hf=hf+"e1792"					&& Kontonummer Verrechnungsnummer 1792
	hf=hf+CHR(187)+"9876"+CHR(28)	&& Kostenstelle X'BB  9876=Kostenstelle chr(28)=Feldende
	hf=hf+CHR(30)+"buchtxt"+CHR(28)	&& Buchungstext max 30 Stellen chr(28)=Feldende
	hf=hf+CHR(179)+"EUR"+CHR(28)	&& Währungskennzeichen X'B3
	hf=hf+"y"						&& Satzende
	IF LEN(datensatz)+LEN(hf)>sl
		FWRITE(nhandle,datensatz)
		X=FSEEK(nhandle,0,2)
		x1=INT(x/256)
		IF x1>0
			az=256-(x-(x1*256))
		ELSE
			az=256-x
		ENDIF
		FOR i1=1 TO az
			FWRITE(nhandle,CHR(0))
		NEXT i1
		datensatz=hf
		sl=250
	ELSE
		datensatz=datensatz+hf
		hf=""
	endif
NEXT i	
IF LEN(datensatz)>0
	FWRITE(nhandle,datensatz)
endif	

hf=hf+"x"						&& Mandantenendsumme x=positiv,w=negativ
hf=hf+STRTRAN(ALLTRIM(STR(summe,12,2)),".","")	&& Summe letzten 2 STellen Kommastellen 560,25
hf=hf+"y"						&& Satzende
hf=hf+"z"						&& Mandantenende
FWRITE(nhandle,hf)
X=FSEEK(nhandle,0,2)

x1=INT(x/256)
IF x1>0
	az=256-(x-(x1*256))
ELSE
	az=256-x
ENDIF
FOR i=1 TO az
	FWRITE(nhandle,CHR(0))
NEXT i
x=FSEEK(nhandle,0,2)
bnr=INT(x/256)
FCLOSE(nhandle)
* Debitordatei erstellen
nhandle=FCREATE("c:\aktuell\ED00002")
IF nhandle==-1	
	WAIT WINDOW "Fehler!!!"
	EXIT
ENDIF
* Kopfdatei Debitor
FWRITE(nhandle,CHR(29))			&& Vorlaufbeginn Konstante = X'1D
FWRITE(nhandle,CHR(24))			&& Kennung neuer Vorlauf Konstante X'18
FWRITE(nhandle,"1")				&& Versionsnummer
FWRITE(nhandle,"001")			&& Datenträgernummer
FWRITE(nhandle,"13")			&& Anwendungsnummer OPOS-Stammdaten
FWRITE(nhandle,"os")			&& Namenskürzel
FWRITE(nhandle,"0004503")		&& Beraternummer
FWRITE(nhandle,"00002")			&& Mandantennummer
FWRITE(nhandle,"010105")		&& Abrechnungsnummer
FWRITE(nhandle,"olix")			&& Passwort
FWRITE(nhandle,SPACE(16))		&& Anwendungsinfo im Beispiel: SELF ID: 99999(=Selef ID 11579)
FWRITE(nhandle,SPACE(16))		&& Input INfo
FWRITE(nhandle,"y")				&& Satzende
FWRITE(nhandle,CHR(182))		&& Versionskennung X'B6
FWRITE(nhandle,"1")				&& Versionsnummer Konstante =1
FWRITE(nhandle,",")				&& Trennzeichen
FWRITE(nhandle,"4")				&& Länge der Sachkontennummer 4-8
FWRITE(nhandle,",")				&& Trennzeichen 
FWRITE(nhandle,"4")				&& Länge der gespeicherten Sachkontennummern 4-8
FWRITE(nhandle,",")				&& Trennzeichen
FWRITE(nhandle,"SELF")			&& Produktkürzel Konstante
FWRITE(nhandle,CHR(28))			&& Feldende X'1C
FWRITE(nhandle,"y")				&& Satzende
* Hier wird der Debitorensatz abgelegt
kontonr=10000
FWRITE(nhandle,"t")				&& Kennziffer
FWRITE(nhandle,"101")			&& Ersteingabe /Änderung
FWRITE(nhandle,CHR(30))			&& Textstart X'1E
FWRITE(nhandle,"2")				&& 1=ERsteingabe,2=Änderung immer 2senden
FWRITE(nhandle,CHR(28))			&& Textende X'1C
FWRITE(nhandle,"y")				&& Satzende
FWRITE(nhandle,"t")
FWRITE(nhandle,"102")			&& Feld Kontonummer
FWRITE(nhandle,CHR(30))
kontonr=kontonr+1
FWRITE(nhandle,ALLTRIM(STR(kontonr,9,0))) && Kontonummer
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t103")			&& Name1
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"Schlingmeier Olaf")
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t203")			&& Firma
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"Citadel Hotelsoftware GmbH")
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t106")			&& PLZ
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"48231")
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t107")			&& Ort
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"Warendorf")
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t108")			&& Strasse
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"Alter Muensterweg 29")
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"t109")			&& Anrede
FWRITE(nhandle,CHR(30))
FWRITE(nhandle,"5")				&& 1=Herrn/Frau/Frl/Firma; 2=Herrn, 3=Frau, 4=Frl; 5=Firma; 6=Eheleute; 7=Herrn und Frau
FWRITE(nhandle,CHR(28))
FWRITE(nhandle,"y")
FWRITE(nhandle,"z" ) 
X=FSEEK(nhandle,0,2)
x1=INT(x/256)
IF x1>0
	az=256-(x-(x1*256))
ELSE
	az=256-x
ENDIF

FOR i=1 TO az
	FWRITE(nhandle,CHR(0))
NEXT i
X=FSEEK(nhandle,0,2)
x1=INT(x/256)
FCLOSE(nhandle)
* EV01 DAtei erstellen
=ev01(bnr,.t.,x1)
RETURN .t.

FUNCTION ev01
PARAMETERS blocknr, Debitor,debblocknr
	nhandle=FCREATE("c:\aktuell\EV01")
	IF nhandle==-1	
		WAIT WINDOW "Fehler!!!"
		EXIT
	ENDIF
	FWRITE(nhandle,"001")  			&& Datenträger-Nummer
	FWRITE(nhandle,SPACE(3))		&& Füllzeichen
	FWRITE(nhandle,"0004503")		&& Beratenummer
	FWRITE(nhandle,"Citadel  ")		&& Beraternme
	FWRITE(nhandle," ")				&& Konstante Leereichen
	IF Debitor=.t.
		FWRITE(nhandle,"00002")			&& Anzahl Datendateien EDxxxxx
	else
		FWRITE(nhandle,"00001")			&& Anzahl Datendateien EDxxxxx
	endif
	FWRITE(nhandle,"00001")			&& Nummer der letzten Datendatei
	FWRITE(nhandle,SPACE(95))
	FWRITE(nhandle,"V")				&& Verarbeitungskennzeichen imm V
	FWRITE(nhandle,"00001")			&& Dateinummer ED00001
	FWRITE(nhandle,"11")			&& Anwendungsnummer=11 Buchungsdatei, 13=Debitorenstammatz
	FWRITE(nhandle,"os")			&& Namenskürzel
	FWRITE(nhandle,"0004503")		&& Beraternummer
	FWRITE(nhandle,"00002")			&& Mandantennummer
	FWRITE(nhandle,"010105")		&& Abrechnungsummer nnnnJJ Zyklen im Buchungsjahr 
	FWRITE(nhandle,"0000010105")	&& Datum von  mit 4 führenden Nullen auffüllen 
	FWRITE(nhandle,"280105")		&& Datum bis
	FWRITE(nhandle,"001")			&& Primanota-Seite immer 1
	FWRITE(nhandle,"olix")			&& Passwort 4-Stellen
	*FWRITE(nhandle,"00001")			&& Letzte Blocknummer 256Byte Zählbegin mit 1 (ED00001)
	FWRITE(nhandle,STRTRAN(STR(blocknr,5,0)," ","0")) && letzte Blocknummer
	FWRITE(nhandle,"001")			&& Letzte Primanota-Seitenzählung immer 001
	FWRITE(nhandle," ")				&& Korrekturkennzeichen Leerzeichen 
	FWRITE(nhandle,"1")				&& Sonderverabeitung "1" Konstante
	FWRITE(nhandle,"1,4,4,SELF    ") && Sachkontenlänge 4 oder 8 Stellen
	FWRITE(nhandle,SPACE(53))		&& Füllzeichen
	*
	FWRITE(nhandle,"V")				&& Verarbeitungskennzeichen imm V
	FWRITE(nhandle,"00002")			&& Dateinummer ED00002
	FWRITE(nhandle,"13")			&& Anwendungsnummer=11 Buchungsdatei, 13=Debitorenstammatz
	FWRITE(nhandle,"os")			&& Namenskürzel
	FWRITE(nhandle,"0004503")		&& Beraternummer
	FWRITE(nhandle,"00002")			&& Mandantennummer
	FWRITE(nhandle,"010105")		&& Abrechnungsummer nnnnJJ Zyklen im Buchungsjahr 
	FWRITE(nhandle,SPACE(10))	&& Datum von  mit 4 führenden Nullen auffüllen 
	FWRITE(nhandle,SPACE(6))		&& Datum bis
	FWRITE(nhandle,"001")			&& Primanota-Seite immer 1
	FWRITE(nhandle,"olix")			&& Passwort 4-Stellen
	*FWRITE(nhandle,"00001")			&& Letzte Blocknummer 256Byte Zählbegin mit 1 (ED00001)
	
	FWRITE(nhandle,STRTRAN(STR(debblocknr,5,0)," ","0")) && letzte Blocknummer
	FWRITE(nhandle,"001")			&& Letzte Primanota-Seitenzählung immer 001
	FWRITE(nhandle," ")				&& Korrekturkennzeichen Leerzeichen 
	FWRITE(nhandle,"1")				&& Sonderverabeitung "1" Konstante
	FWRITE(nhandle,"1,4,4,SELF    ") && Sachkontenlänge 4 oder 8 Stellen
	FWRITE(nhandle,SPACE(53))		&& Füllzeichen
	
	fCLOSE(nhandle)
	
	*
	* Vorlaufdatei ende
	*
RETURN .t.

