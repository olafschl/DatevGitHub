* Datev.prg  Testprogramm 19.09.2022 fuer branch git 
 PROCEDURE datev
	PUBLIC lcbnummer,lcbname,lcmnummer,lcpasswort,lckostenstelle,lcselfid,lcasachkl,lcgsachkl
	PUBLIC lcabgrg,lcldebitornr,lcautodebit,lceartikel,lcezahlung,lcedebitor,lcverdichtung
	PUBLIC lcvartikel,lcvzahlung,lcverz,lcazimmer,lcazahlungv,lcazahlungb,lcResZeitraum
	PUBLIC lcFirma_na, lcnetto, lceKostenstelle, lcUKassetrennen, lcKopf, lcSatz, lcSatz2, lcPrgTyp
	PUBLIC lcVKasse, lcAartikel,lcvAuslagen,lcvAusgaben,lcMwst8,lcMwSt9,lcMwSt2,lcMwst0,lcmwst3,lcASCII
	PUBLIC lcASCIIERW,lcWJ,lcaZText,lcaZTFeld,lcStKey,lcMwStGrpDeb,lcvHausbank,lcKStRoom,lcDebitorZaehler
	PUBLIC lcReport,lcDebitorImport,lcDebitorExport,lcDKopf,lcDSatz,lcASCIIANF
	LOCAL old_select,l_cError
	l_cError=ON("error")
	ON ERROR DO errhand WITH ERROR( ), MESSAGE( ), MESSAGE(1), PROGRAM( ), LINENO( )

	** ReFox 6.0 - Branding Code
	IF .F.
		_refox_ =  (9876543210)
	ENDIF

	* Modify Window Screen close
	on shutdown quit
	set date GERMAN
	set mark to "."
	set exclusive off
	set Reprocess to 5 Sconds
	set notify off
	set talk off
	set deleted on
	* set path to sys(3)+sys(2003)
	* set sysmenu off
	* _screen.WindowState=2
	
	old_select=SELECT()

	* Versions-Information der Exe-Datei auslesen und in Variable speichern
	lcdatei = ALLTRIM(substr(SYS(16,0), AT(" ", SYS(16,0),2), LEN(SYS(16,0))))
	DIMENSION versionsdaten(1)
	arrcnt = aGetFileVersion(versionsdaten, lcdatei)
	IF arrcnt # 0
		lcversion = ALLT(versionsdaten(4))
	ELSE
		lcversion = ""
	ENDIF
	ifcIT = .F.
	ifcnodemo = .F.

	* Auslesen der ini-Datei und von Balast befreien
	lcbnummer		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Beraternummer", 60)), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Beraternummer", 60)))-1)
	lcbname 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Beratername", 60)), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Beratername", 60)))-1)
	lcmnummer 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Mandantennummer", 60 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Mandantennummer", 60)))-1)
	lcPasswort 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Passwort", 4 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Passwort", 4 )))-1)
	lckostenstelle	= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Kostenstelle", 60 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Kostenstelle", 60 )))-1)
	lcselfid		= SUBSTR(ALLTRIM(ini_read("DatenTraegerKennsatz", "SELFID",1)),1, ATC(CHR(0), ALLTRIM(ini_read("DatenTraegerKennsatz", "SELFID",1)))-1)
	lcasachkl		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASachkl", 60)), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASachkl", 60)))-1)
	lcgsachkl		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "GSachkl", 60)), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "GSachkl", 60)))-1)
	lcabgrg 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "AbgRg", 1 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "AbgRg", 1)))-1)
	lcStKey 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Steuerkey", 1 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Steuerkey", 1)))-1)	
	lcldebitornr	= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "lDebitorNr", 60 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "lDebitorNr", 60 )))-1)
	lcdebitorzaehler= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Debitorzaehler", 5 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Debitorzaehler", 5 )))-1)	
	lcautodebit		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Autodebitor", 60 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Autodebitor", 60 )))-1)
	lcverdichtung	= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "Verdichtung", 60 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "Verdichtung", 60 )))-1)	
	lcPrgTyp		= SUBSTR(ALLTRIM(ini_read("DatenTraegerKennsatz", "PrgTyp",1)),1, ATC(CHR(0), ALLTRIM(ini_read("DatenTraegerKennsatz", "PrgTyp",1)))-1)
	lcASCII 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCII", 1 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCII", 1)))-1)	
	lcASCIIERW 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCIIERW", 3 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCIIERW", 3)))-1)		
	lcASCIIANF		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCIIANF", 8 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "ASCIIANF", 8)))-1)		
	lcWJ 	 		= SUBSTR(ALLTRIM(INI_Read("DatenTraegerKennsatz", "WJ", 4 )), 1, ATC(CHR(0), ALLTRIM(INI_Read("DatenTraegerKennsatz", "WJ", 4)))-1)			&& Wirtschaftsjahr

	* ASCII wenn ja wird eine zusätzliche ASCII Datei erstellt zu der Datev-Datei - für  sonstige Fibu z. B. Medussa Fibu
	* lcPrgTyp = 0 Datev 1=KHk

	lceArtikel	 	= SUBSTR(ALLTRIM(ini_read("Ersatzkonten", "Artikel",8)),1, ATC(CHR(0), ALLTRIM(ini_read("Ersatzkonten", "Artikel",8)))-1)
	lceZahlung		= SUBSTR(ALLTRIM(INI_Read("Ersatzkonten", "Zahlung", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Ersatzkonten", "Zahlung", 8)))-1)
	lceDebitor		= SUBSTR(ALLTRIM(ini_read("Ersatzkonten", "Debitor", 10)), 1, ATC(CHR(0), ALLTRIM(ini_read("Ersatzkonten", "Debitor", 10)))-1)
	lceKostenstelle	= SUBSTR(ALLTRIM(ini_read("Ersatzkonten", "Kostenstelle", 10)), 1, ATC(CHR(0), ALLTRIM(ini_read("Ersatzkonten", "Kostenstelle", 10)))-1)	

	lcverz			= SUBSTR(ALLTRIM(ini_read("PFADE","AVERZ",150)),1,ATC(CHR(0),ALLTRIM(ini_read("PFADE","AVERZ",150)))-1)
	lcdbfverz		= SUBSTR(ALLTRIM(ini_read("PFADE","DBFVERZ",150)),1,ATC(CHR(0),ALLTRIM(ini_read("PFADE","DBFVERZ",150)))-1)
	lcreport		= SUBSTR(ALLTRIM(ini_read("PFADE","Report",150)),1,ATC(CHR(0),ALLTRIM(ini_read("PFADE","Report",150)))-1)	
	lcDebitorImport	= SUBSTR(ALLTRIM(ini_read("PFADE","DebitorImport",150)),1,ATC(CHR(0),ALLTRIM(ini_read("PFADE","DebitorImport",150)))-1)	
	lcDebitorExport	= SUBSTR(ALLTRIM(ini_read("PFADE","DebitorExport",150)),1,ATC(CHR(0),ALLTRIM(ini_read("PFADE","DebitorExport",150)))-1)	
	
	lcmwst8			= SUBSTR(ALLTRIM(ini_read("MWST","8",5)),1,ATC(CHR(0),ALLTRIM(ini_read("MWST","8",5)))-1)
	lcmwst9			= SUBSTR(ALLTRIM(ini_read("MWST","0",5)),1,ATC(CHR(0),ALLTRIM(ini_read("MWST","9",5)))-1)
	lcmwst2			= SUBSTR(ALLTRIM(ini_read("MWST","2",5)),1,ATC(CHR(0),ALLTRIM(ini_read("MWST","2",5)))-1)
	lcmwst3			= SUBSTR(ALLTRIM(ini_read("MWST","3",5)),1,ATC(CHR(0),ALLTRIM(ini_read("MWST","3",5)))-1)

	lcvartikel		= SUBSTR(ALLTRIM(ini_read("Verrechnungskonto","ARTIKEL",8)),1,ATC(CHR(0),ALLTRIM(ini_read("Verrechnungskonto","ARTIKEL",8)))-1)	
	lcvzahlung 		= SUBSTR(ALLTRIM(INI_Read("Verrechnungskonto","Zahlung", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Verrechnungskonto", "Zahlung", 8)))-1)
	lcvKasse 		= SUBSTR(ALLTRIM(INI_Read("Verrechnungskonto","Kasse", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Verrechnungskonto", "Kasse", 8)))-1)
	lcvAuslagen		= SUBSTR(ALLTRIM(INI_Read("Verrechnungskonto","Auslagen", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Verrechnungskonto", "Auslagen", 8)))-1)
	lcvAusgaben		= SUBSTR(ALLTRIM(INI_Read("Verrechnungskonto","Ausgaben", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Verrechnungskonto", "Ausgaben", 8)))-1)	
	lcvHausbank		= SUBSTR(ALLTRIM(INI_Read("Verrechnungskonto","Hausbank", 8)), 1, ATC(CHR(0), ALLTRIM( INI_Read("Verrechnungskonto", "Hausbank", 8)))-1)	

	lcazimmer		= SUBSTR(ALLTRIM(ini_read("Ausnahme","Zimmer",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","Zimmer",5)))-1)
	lcazahlungv		= SUBSTR(ALLTRIM(ini_read("Ausnahme","Zahlungv",50)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","Zahlungv",50)))-1)
	lcazahlungb		= SUBSTR(ALLTRIM(ini_read("Ausnahme","Zahlungb",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","Zahlungb",5)))-1)
	
	lcResZeitraum	= SUBSTR(ALLTRIM(ini_read("Ausnahme","ResZeitraum",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","ResZeitraum",5)))-1)	
	lcFirma_na		= SUBSTR(ALLTRIM(ini_read("Ausnahme","FirmaNa",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","FirmaNa",5)))-1)				
	lcNetto			= SUBSTR(ALLTRIM(ini_read("Ausnahme","Netto",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","Netto",5)))-1)		
	lcaArtikel		= SUBSTR(ALLTRIM(ini_read("Ausnahme","Artikel",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","Artikel",5)))-1)		
	lcaZTFeld		= SUBSTR(ALLTRIM(ini_read("Ausnahme","ZTFeld",100)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","ZTFeld",100)))-1)	
	lcaZText		= SUBSTR(ALLTRIM(ini_read("Ausnahme","ZText",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","ZText",5)))-1)	
	lcMwStGrpDeb	= SUBSTR(ALLTRIM(ini_read("Ausnahme","MWSTGRPDEB",2)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","MWSTGRPDEB",2)))-1)		
	lcKStRoom		= SUBSTR(ALLTRIM(ini_read("Ausnahme","KSTRoom",1)),1,ATC(CHR(0), ALLTRIM(ini_read("Ausnahme","KSTRoom",1)))-1)	
	*WAIT WINDOW "lcaZTFeld= "+lcaZTFeld			
	&& Ztext hier wird eingeschaltet ob der Zusatztext mit ausgegeben werden soll oder nicht. (Belegfeld 2 in Datev)
	lcKopf			= SUBSTR(ALLTRIM(ini_read("Buchungssatz","Kopf",200)),1,ATC(CHR(0), ALLTRIM(ini_read("Buchungssatz","Kopf",200)))-1)
	lcSatz			= SUBSTR(ALLTRIM(ini_read("Buchungssatz","Satz",250)),1,ATC(CHR(0), ALLTRIM(ini_read("Buchungssatz","Satz",250)))-1)
	lcSatz2			= SUBSTR(ALLTRIM(ini_read("Buchungssatz","Satz2",250)),1,ATC(CHR(0), ALLTRIM(ini_read("Buchungssatz","Satz2",250)))-1)
	
	lcDKopf			= SUBSTR(ALLTRIM(ini_read("Debitorsatz","Kopf",200)),1,ATC(CHR(0), ALLTRIM(ini_read("Debitorsatz","Kopf",200)))-1)
	lcDSatz			= SUBSTR(ALLTRIM(ini_read("Debitorsatz","Satz",250)),1,ATC(CHR(0), ALLTRIM(ini_read("Debitorsatz","Satz",250)))-1)	
	
	lcMNr1			= SUBSTR(ALLTRIM(ini_read("Mandant","MNR1",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Mandant","MNR1",5)))-1)	
	lcMNr2			= SUBSTR(ALLTRIM(ini_read("Mandant","MNR2",5)),1,ATC(CHR(0), ALLTRIM(ini_read("Mandant","MNR2",5)))-1)
	* hier können noch weitere Mandanten eingebunden werden
	
	lcverz			= ALLTRIM(lcverz)
	IF RIGHT(lcverz,1)<>"\"
		lcverz=lcverz+"\"
	ENDIF
	* Öffnen Protokolldatei
	hf=lcverz+"datevp.dbf"
	IF !FILE(hf)
		WAIT WINDOW "Datei "+hf+" ist nicht vorhanden!!!"
		RETURN .t.
	ELSE
		USE &hf ORDER tag1 IN 0
	ENDIF
	
	* Öffnen der Brilliant-Parameterdatei und Auslesen der Lizenzdaten
	IF USED('param')
		SELECT param
	ELSE
		lcdbfverz=ALLTRIM(lcdbfverz)
		IF RIGHT(lcdbfverz,1)="\"
			hf= lcdbfverz+"param"
		ELSE
			hf=lcdbfverz+"\param"
		endif
		*USE (..\data\param SHARED
		USE &hf shared IN 0
	ENDIF

	* Wenn in Brilliant kein Interface, dann werden keine Daten übertragen
	ifcifc = SUBSTR(pa_lizopt, AT("I",pa_lizopt),AT(",",pa_lizopt))
	IF "T" $ ifcifc
		ifcIT = .T.
	ELSE
		ifcIT = .F.
	endif
	ifchotel = ALLTRIM(pa_hotel)
	ifccity = ALLTRIM(pa_city)
	


	DO ifclizenz
	IF 	ifcnodemo = .T. &&and ifcIT = .T.
		ifchotel = ifchotel + ", " + ifccity
	ELSE
		ifchotel = "DEMO es werden keine Daten übertragen"
	ENDIF

	DO FORM datev
	READ events
	
		
	SELECT (old_select)
	
	ON ERROR &l_cerror
	
	*
	*clear dlls
	*Release all Extended
	*Clear all
	*
	*ON SHUTDOWN
	*QUIT

RETURN

* Funktion zum Schreiben der INI - Datei
FUNCTION INI_Write
	LPARAMETERS tcwo, tcwas, tcwert
	LOCAL lcretwert, lcpfad
	DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
		STRING cSection, STRING cKey, STRING cDefault, STRING @cBuffer, ;
		INTEGER nBufferSize, STRING cINIFile
	DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
		STRING cSection, STRING cKey, STRING cValue, STRING cINIFile
	IF WritePrivStr(tcwo,tcwas, tcwert,SYS(5)+SYS(2003) + "\datev.ini") == 0
		MESSAGEBOX("Fehler beim Schreiben der INI-Datei !",16+0+0,"Fehler")
	ENDIF
RETURN(.T.)

* Funktion zum Lesen der INI - Datei
FUNCTION INI_Read
	LPARAMETERS tcwo, tcwas, tnlen
	LOCAL lcretwert, lcpfad
	lcretwert = SPACE(tnlen) + CHR(0)
	DECLARE INTEGER GetPrivateProfileString IN Win32API  AS GetPrivStr ;
		STRING cSection, STRING cKey, STRING cDefault, STRING @cBuffer, ;
		INTEGER nBufferSize, STRING cINIFile
	DECLARE INTEGER WritePrivateProfileString IN Win32API AS WritePrivStr ;
		STRING cSection, STRING cKey, STRING cValue, STRING cINIFile
	IF GetPrivStr(tcwo,tcwas, "", @lcretwert, LEN(lcretwert),SYS(5)+SYS(2003) + "\datev.ini") == 0
		MESSAGEBOX("Fehler beim Lesen der INI-Datei !",16+0+0,"Fehler "+tcwo+"  "+tcwas)
	ENDIF
RETURN(lcretwert)

FUNCTION ifclizenz
	ifcliz = SUBSTR(ALLTRIM(INI_Read("Lizenz", "Code", 50)), 1, ATC(CHR(0), ALLTRIM(INI_Read("Lizenz", "Code", 50)))-1)
	IF ifcliz = "oS-Hk$"
		newliz = 1
	ELSE
		newliz = 0
	ENDIF
	ifcliz = VAL(ifcliz)
	ncc = 0
	ctmp = alltrim(ifchotel)+"$UniDateV"+alltrim(ifccity)+"iFcuNI19"
	FOR ncount = 1 TO LEN(ctmp)
	     ncc = ncc+(ncount*ASC(SUBSTR(ctmp, ncount, 1)))*2
	ENDFOR
	ifcnodemo = (ncc==ifcliz)
	IF newliz = 1
		newliz = ncc
		writeprivstr("Lizenz", "Code", ALLTRIM(STR(newliz,12,0)), SYS(5)+SYS(2003)+"\datev.ini")
	endif
	RETURN ifcnodemo
ENDFUNC

PROCEDURE errhand
PARAMETER merror, mess, mess1, mprog, mlineno
= MESSAGEBOX('Error: '+ LTRIM(STR(merror)) + CHR(13) + ;
			"Message: " + MESS + CHR(13) + ;
			"Line of Code: " + mess1 + CHR(13) + ;
			"Line number of error: " + LTRIM(STR(mlineno)) + CHR(13) + ;
			"Program with error: " + mprog, 48, 'Programmfehler')
ENDPROC
