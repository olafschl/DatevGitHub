 * Hier erfolgt die Übergabe an KHK
 * In dieser Funktion werden alle Buchungen vom vorgegebenen 
 * Zeitraum übergeben.
 PARAMETERS Dstart,Dende
 LOCAL norder_hr,norder_hp,old_near,ok_ad,lmnr,old_near
 LOCAL caMacro, cpMacro ,ckstkonto,cpaykonto,cpaytext,debnr
 LOCAL pr_anz,pr_summe,lanrede,dsatz
 	dsatz=""
	caMacro = "Article.Ar_Lang"+g_Langnum
    cpMacro = "Paymetho.Pm_Lang"+g_Langnum
	old_near=SET("near")
	pr_anz=0
	pr_summe=0
	SET NEAR on
  	SELECT histpost
 	norder_hp=ORDER()
 	SET ORDER TO 2
	IF VAL(lcmnummer)>0
		lmnr=(lcmnummer)
	ELSE
		lmnr=1
 	ENDIF
	SELECT histpost
	SEEK dstart
	DO WHILE EOF()=.f. and histpost.hp_date>dende
 		DO case
			CASE ((histpost.hp_split=.f. and EMPTY(histpost.hp_ratecod)) or (!EMPTY(histpost.hp_ratecod) and histpost.hp_split=.t.));
			AND histpost.hp_reserid>0 AND !histpost.hp_cancel AND !EMPTY(histpost.hp_artinum)
				=SEEK(histpost.hp_artinum,"Article","Tag1")
				IF histpost.hp_reserid=366.100
					*WAIT WINDOW "hp_reserid=366.100 in postschleife"
				endif
				SELECT datev
				IF article.ar_artityp<>3 OR (histpost.hp_artinum=VAL(lcaArtikel) AND VAL(lcaArtikel)<>0)
					IF lcverdichtung="J"
						*hf=STR(VAL(lmnr),5,0)+IIF(!EMPTY(article.ar_user2),ALLTRIM(article.ar_user2),ALLTRIM(lceartikel))+STR(histpost.hp_artinum,4,0)+DTOS(histpost.hp_date)
						hf=STR(VAL(lmnr),5,0)+histpost.hp_billnum+IIF(!EMPTY(article.ar_user2),ALLTRIM(article.ar_user2),ALLTRIM(lceartikel))
						IF !SEEK(STR(VAL(lmnr),5,0)+histpost.hp_billnum+IIF(!EMPTY(article.ar_user2),ALLTRIM(article.ar_user2),ALLTRIM(lceartikel))+STR(histpost.hp_window,1),"DATEV","TAG4")
						*IF !SEEK(STR(VAL(lmnr),5,0)+IIF(!EMPTY(article.ar_user2),ALLTRIM(article.ar_user2),ALLTRIM(lceartikel))+histpost.hp_billnum ,"DATEV","TAG6")
							*WAIT WINDOW "hf= "+hf+"  nicht gefunden!!!"
							APPEND BLANK 
						ENDIF
					else
						APPEND BLANK 
					ENDIF
					IF article.ar_artityp=2
						replace datev.auslagen WITH .t.
					ELSE
						replace datev.auslagen WITH .f.
					ENDIF
					DO case
						CASE article.ar_vat = 8
							replace datev.mwst with lcmwst8
						CASE article.ar_vat = 9
							replace datev.mwst with lcmwst9
						CASE article.ar_vat = 2
							replace datev.mwst with lcmwst2
					OTHERWISE
						replace datev.mwst WITH ""
					endcase
					replace datev.artikel 	WITH histpost.hp_artinum
					replace datev.datum		WITH histres.hr_depdate  && histpost.hp_date musste gesetzt werden, damit buchungsdatum gleich rechnungsdatum ist
					replace datev.btext		WITH &caMacro
					replace datev.konto		WITH lcvartikel
					replace datev.beleg2	WITH ""
					replace datev.bkz		WITH "R"
					replace datev.rgfenster	WITH histpost.hp_window
					replace datev.mandantnr	WITH val(lmnr)				
					replace datev.rechnr	WITH histpost.hp_billnum
					replace datev.reserid	WITH histpost.hp_reserid
					replace datev.adressid	WITH histres.hr_addrid
					replace datev.abreise	WITH histres.hr_depdate
					replace datev.anreise	WITH histres.hr_arrdate
					replace datev.rname		WITH histres.hr_lname
					replace datev.rfirma	WITH histres.hr_company
					IF histpost.hp_userid="POS" OR histpost.hp_userid="POSZ2"
						replace datev.argus	WITH .t.
					endif
					*WAIT WINDOW "datev.rechnr= "+datev.rechnr
					DO case
						CASE histpost.hp_window=1
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr1)))
						CASE histpost.hp_window=2
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr2)))
							hf=SUBSTR(MLINE(histres.hr_billins,2),1,12)
							IF val(hf)>0
								replace datev.adressid	WITH VAL(hf)
							endif
						CASE histpost.hp_window=3
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr3)))
							hf=SUBSTR(MLINE(histres.hr_billins,3),1,12)
							IF val(hf)>0
								replace datev.adressid	WITH VAL(hf)
							endif
						CASE histpost.hp_window=4
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr4)))
							hf=SUBSTR(MLINE(histres.hr_billins,5),1,12)
							IF val(hf)>0
								replace datev.adressid	WITH VAL(hf)
							endif
						CASE histpost.hp_window=5
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr5)))
							hf=SUBSTR(MLINE(histres.hr_billins,6),1,12)
							IF val(hf)>0
								replace datev.adressid	WITH VAL(hf)
							endif
						CASE histpost.hp_window=6
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr6)))
							hf=SUBSTR(MLINE(histres.hr_billins,7),1,12)
							IF val(hf)>0
								replace datev.adressid	WITH VAL(hf)
							endif
					endcase	
					* gesetzt da 2006 noch nicht histpost.hp_billnum vewendet wurde
					IF !EMPTY(histpost.hp_billnum)
						replace datev.beleg1			WITH histpost.hp_billnum							
					ENDIF
					IF !empty(article.ar_user2)
						replace datev.gegenkonto 	WITH ALLTRIM(article.ar_user2)
					ELSE
						replace datev.gegenkonto	WITH ALLTRIM(lceartikel)
					ENDIF
					IF !EMPTY(article.ar_user1)
						replace datev.kostenstelle	WITH ALLTRIM(article.ar_user1)
					ELSE
						replace datev.kostenstelle	WITH ALLTRIM(lcekostenstelle)
					endif
				endif
 				IF article.ar_artityp<>3 OR (histpost.hp_artinum=VAL(lcaArtikel) AND VAL(lcaArtikel)<>0)
					replace datev.umsatz WITH datev.umsatz+histpost.hp_amount
					replace datev.mwstbetrag WITH datev.mwstbetrag+histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                     	   	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
				ENDIF
	 		CASE (!EMPTY(histpost.hp_paynum) AND histpost.hp_reserid>0 AND !histpost.hp_cancel) and histpost.hp_paynum<>param.pa_payonld;
	 			AND BETWEEN(histpost.hp_paynum,VAL(lcazahlungv),VAL(lcazahlungb))=.f.
	 			SELECT datev
	 			SET ORDER TO 2
	 			IF LOOKUP(paymetho.pm_paytyp,histpost.hp_paynum,paymetho.pm_paynum)==4
	 				SELECT address
	 				IF !EMPTY(histres.hr_compid)
	 					SEEK histres.hr_compid
	 				ELSE
	 					SEEK histres.hr_addrid
	 				ENDIF
	 				cpaykonto=ALLTRIM(STR(address.ad_compnum))
	 				IF VAL(cpaykonto)<10000  && Debitorenkonten sind immer >9999
	 					IF !EMPTY(paymetho.pm_user2)
	 						cpaykonto=ALLTRIM(paymetho.pm_user2)
	 					ELSE
	 						cpaykonto=ALLTRIM(lcedebitor)
	 						*WAIT WINDOW "lcezahlung --- 1"
	 					ENDIF
	 				ENDIF
	 				IF !EMPTY(paymetho.pm_user1)
	 					ckstkonto=ALLTRIM(paymetho.pm_user1)
	 				ELSE
	 					ckstkonto=""
	 				ENDIF
	 			ELSE
 					IF !EMPTY(paymetho.pm_user2)
 						cpaykonto=ALLTRIM(paymetho.pm_user2)
 					ELSE
 						cpaykonto=ALLTRIM(lceZahlung)
 						*WAIT WINDOW "lcezahlung  ---2"
 					ENDIF
	 				IF !EMPTY(paymetho.pm_user1)
	 					ckstkonto=ALLTRIM(paymetho.pm_user1)
	 				ELSE
	 					ckstkonto=""
	 				ENDIF
	 			ENDIF
	 			SELECT datev
	 			SET ORDER TO tag2
	 			*IF INLIST(paymetho.pm_paytyp,3,4) OR !SEEK(DTOS(histpost.hp_date)+STR(histpost.hp_paynum,2),"Datev")
	 				APPEND BLANK
	 				IF histpost.hp_reserid=14146.100
	 					*WAIT WINDOW "Reserid= 14146.100 Satz-nr. Histres= "+STR(RECNO("histres"),8,0)
	 				endif
					DO case
						CASE histpost.hp_window=1
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr1)))
						CASE histpost.hp_window=2
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr2)))
						CASE histpost.hp_window=3
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr3)))
						CASE histpost.hp_window=4
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr4)))
						CASE histpost.hp_window=5
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr5)))
						CASE histpost.hp_window=6
							replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr6)))
					endcase	
					replace datev.beleg1			WITH histpost.hp_billnum
					replace datev.rechnr			WITH histpost.hp_billnum							
					replace datev.zahlung			WITH histpost.hp_paynum
					replace datev.kostenstelle		WITH IIF(UPPER(ALLTRIM(lckostenstelle))="J",ckstkonto,"")
					replace datev.konto				WITH IIF(paymetho.pm_paytyp=4,"0",cpaykonto)
					replace datev.gegenkonto		WITH lcvzahlung
					replace datev.abreise			WITH histres.hr_depdate
					replace datev.rname		WITH histres.hr_lname
					replace datev.rfirma	WITH histres.hr_company
					replace datev.anreise	WITH histres.hr_arrdate
					replace datev.reserid			WITH histres.hr_reserid
					replace datev.adressid			with histres.hr_addrid
					IF histres.hr_reserid=0.100
						cpaytext="Pasantenbuchung"
					ELSE
						IF histpost.hp_window=2
							cpaytext=IIF(EMPTY(histres.hr_company),histres.hr_lname,ALLTRIM(histres.hr_company))
						ELSE
							cpaytext=IIF(EMPTY(histres.hr_lname),histres.hr_company,ALLTRIM(histres.hr_lname)+" "+;
								IIF(EMPTY(histres.hr_company)," ","/ "+ALLTRIM(histres.hr_company)))
						endif
					ENDIF
					replace datev.btext				WITH cpaytext
					replace datev.datum				WITH histpost.hp_date
					replace datev.bkz				WITH IIF(VAL(cpaykonto)>9999,"R","") && damit Kreditzahlung als Debitor laufen
					replace datev.umsatz			WITH datev.umsatz+(histpost.hp_amount*-1)
					replace datev.mandantnr			WITH val(lmnr)
					replace datev.rgfenster			WITH histpost.hp_window
					IF paymetho.pm_paytyp=4
						IF address.ad_compnum<10000
							IF UPPER(ALLTRIM(lcautodebit))="J"
								debnr=LOOKUP(id.id_last,"DEBITOR",id.id_code)
								IF debnr=0
									debnr=VAL(lcedebitor)
								else
									replace address.ad_compnum	WITH debnr
									replace id.id_last			WITH debnr+1
								endif
							ELSE
								*debnr=VAL(lcedebitor)
								debnr=lcedebitor
							ENDIF
						ELSE
							debnr=address.ad_compnum
						ENDIF
						*replace datev.debitornr		WITH IIF(debnr>0,ALLTRIM(STR(debnr,10,0)),ALLTRIM(lcedebitor))
						replace datev.debitornr		WITH IIF(!EMPTY(debnr),ALLTRIM(debnr),ALLTRIM(lcedebitor))
						replace datev.firma			WITH address.ad_company
						replace datev.anrede		WITH address.ad_titlcod
						replace datev.suchname		WITH address.ad_compkey
						replace datev.na_vo			WITH ALLTRIM(address.ad_lname)+", "+ALLTRIM(address.ad_fname)
						replace datev.strasse		WITH address.ad_street
						replace datev.telefon		WITH address.ad_phone
						replace datev.plz			WITH address.ad_zip
						replace datev.ort			WITH address.ad_city
						replace datev.land			WITH address.ad_country
					endif
	 			*endif
 			ENDCASE
 		ENDDO
 		SELECT histpost
 		SKIP 1 
 	ENDDO
 	
 	* Hier wird die Passantenbuchungen eingebunden wichtig für Argus Z-Abschlag
 	SELECT histpost
 	GO  top
 	SET ORDER TO tag2
 	SEEK dstart
 	DO WHILE EOF()=.f. and BETWEEN(histpost.hp_date,dstart,dende)
 		*IF histpost.hp_userid<>"POSZ2" 
 		IF histpost.hp_reserid<>0.100
 			SKIP
 			LOOP
 			* rausgenommen, da sonst die passerby buchungen aus dem hotel 
 			* nicht mit übernommen wurden
 		ENDIF
 		DO case
 			CASE histpost.hp_artinum>0 AND !histpost.hp_cancel
 				*
				=SEEK(histpost.hp_artinum,"Article","Tag1")
				SELECT datev
				IF article.ar_artityp<>3
					IF lcverdichtung="J"
						IF !SEEK(STR(VAL(lmnr),5,0)+IIF(!EMPTY(article.ar_user2),ALLTRIM(article.ar_user2),ALLTRIM(lceartikel))+DTOS(histpost.hp_date),"DATEV","TAG5")
							APPEND BLANK 
						ENDIF
					else
						APPEND BLANK 
					ENDIF
					IF article.ar_artityp=2
						replace datev.auslagen WITH .t.
					ELSE
						replace datev.auslagen WITH .f.
					ENDIF
					IF EMPTY(HISTPOST.hp_billnum)
						replace datev.beleg1 WITH ALLTRIM(STR(histpost.hp_postid,8,0))
					ELSE
						replace datev.beleg1 WITH histpost.hp_billnum
					ENDIF
					IF histpost.hp_userid="POSZ2"
						replace datev.argus		WITH .t.
					ENDIF
					DO case
						CASE article.ar_vat = 8
							replace datev.mwst with lcmwst8
						CASE article.ar_vat = 9
							replace datev.mwst with lcmwst9
						CASE article.ar_vat = 2
							replace datev.mwst with lcmwst2
					OTHERWISE
						replace datev.mwst WITH ""
					endcase
					
					replace datev.artikel 	WITH histpost.hp_artinum
					replace datev.datum		WITH histpost.hp_date
					replace datev.btext		WITH &caMacro
					replace datev.konto		WITH lcvartikel
					replace datev.beleg2	WITH ""
					replace datev.bkz		WITH "R"
					replace datev.rgfenster	WITH histpost.hp_window
					replace datev.mandantnr	WITH val(lmnr)				
					replace datev.reserid	WITH histpost.hp_reserid
					replace datev.adressid	WITH histpost.hp_addrid
					DO case
						CASE histpost.hp_window=1
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr1)))
						CASE histpost.hp_window=2
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr2)))
						CASE histpost.hp_window=3
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr3)))
						CASE histpost.hp_window=4
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr4)))
						CASE histpost.hp_window=5
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr5)))
						CASE histpost.hp_window=6
							*replace datev.beleg1	WITH ALLTRIM(STR(VAL(histres.hr_billnr6)))
					endcase								
					IF !empty(article.ar_user2)
						replace datev.gegenkonto 	WITH ALLTRIM(article.ar_user2)
						replace datev.kostenstelle	WITH ALLTRIM(article.ar_user1)
					ELSE
						replace datev.gegenkonto	WITH ALLTRIM(lceartikel)
						replace datev.kostenstelle	WITH ""
					ENDIF
				endif
 				IF article.ar_artityp<>3
					replace datev.umsatz WITH datev.umsatz+histpost.hp_amount
					replace datev.mwstbetrag WITH datev.mwstbetrag+histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                    	   	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
				ENDIF
 				
 			CASE histpost.hp_paynum>0 AND  !histpost.hp_cancel
 				*
 					*WAIT WINDOW "Passantenzahlung POSZ?? "+histpost.hp_userid
	 				SELECT datev
	 				SET ORDER TO 2
	 				IF LOOKUP(paymetho.pm_paytyp,histpost.hp_paynum,paymetho.pm_paynum)==4
	 					SELECT address
	 					IF !EMPTY(histres.hr_compid)
	 						SEEK histres.hr_compid
	 					ELSE
	 						SEEK histres.hr_addrid
	 					ENDIF
	 					cpaykonto=ALLTRIM(STR(address.ad_compnum))
	 					IF VAL(cpaykonto)<10000  && Debitorenkonten sind immer >9999
	 						IF !EMPTY(paymetho.pm_user2)
	 							cpaykonto=ALLTRIM(paymetho.pm_user2)
	 						ELSE
	 							cpaykonto=ALLTRIM(lcedebitor)
	 							*WAIT WINDOW "lcezahlung --- 1"
	 						ENDIF
	 					ENDIF
	 					IF !EMPTY(paymetho.pm_user1)
	 						ckstkonto=ALLTRIM(paymetho.pm_user1)
	 					ELSE
	 						ckstkonto=""
	 					ENDIF
	 				ELSE
 						IF !EMPTY(paymetho.pm_user2)
 							cpaykonto=ALLTRIM(paymetho.pm_user2)
 						ELSE
 							cpaykonto=ALLTRIM(lceZahlung)
 							*WAIT WINDOW "lcezahlung  ---2"
 						ENDIF
	 					IF !EMPTY(paymetho.pm_user1)
	 						ckstkonto=ALLTRIM(paymetho.pm_user1)
	 					ELSE
	 						ckstkonto=""
	 					ENDIF
	 				ENDIF
	 				SELECT datev
	 				SET ORDER TO tag2
	 				*IF INLIST(paymetho.pm_paytyp,3,4) OR !SEEK(DTOS(histpost.hp_date)+STR(histpost.hp_paynum,2),"Datev")

	 					APPEND BLANK
	 					IF EMPTY(histpost.hp_billnum)
	 						replace datev.beleg1		WITH ALLTRIM(STR(histpost.hp_postid,8,0))
	 					ELSE 
		 					replace datev.beleg1			WITH histpost.hp_billnum
		 				endif
						replace datev.zahlung			WITH histpost.hp_paynum
						replace datev.kostenstelle		WITH IIF(UPPER(ALLTRIM(lckostenstelle))="J",ckstkonto,"")
						replace datev.konto				WITH IIF(paymetho.pm_paytyp=4,"0",cpaykonto)
						replace datev.gegenkonto		WITH lcvzahlung
						replace datev.abreise			WITH histpost.hp_date
						replace datev.reserid			WITH histpost.hp_reserid
						replace datev.adressid			WITH histpost.hp_addrid
						IF histpost.hp_reserid=0.100 OR ((histpost.hp_reserid=0.200 OR histpost.hp_reserid=0.500) AND histpost.hp_userid="POSZ2")
							IF histpost.hp_userid="POSZ2"
								cpaytext="POSZ2-Abschlag"
								IF histpost.hp_reserid=0.500
									cpaytext="FW "+cpaytext
									replace datev.beleg1 WITH ALLTRIM(STR(histpost.hp_postid,8,0))
								endif
								replace datev.argus		WITH .t.
							else
								cpaytext="Pasantenbuchung"
							endif
						ELSE
							IF histpost.hp_window=2
								cpaytext=IIF(EMPTY(histres.hr_company),histres.hr_lname,ALLTRIM(histres.hr_company))
							ELSE
								cpaytext=IIF(EMPTY(histres.hr_lname),histres.hr_company,ALLTRIM(histres.hr_lname)+" "+;
									IIF(EMPTY(histres.hr_company)," ","/ "+ALLTRIM(histres.hr_company)))
							endif
						ENDIF
						replace datev.btext				WITH cpaytext
						replace datev.datum				WITH histpost.hp_date 
						replace datev.bkz				WITH IIF(VAL(cpaykonto)>9999,"R","") && damit Kreditzahlung als Debitor laufen
						replace datev.umsatz			WITH datev.umsatz+(histpost.hp_amount*-1)
						replace datev.mandantnr			WITH val(lmnr)
						replace datev.rgfenster			WITH histpost.hp_window
						* WAIT WINDOW cpaytext+"  "+STR(histpost.hp_amount,12,2)
						IF paymetho.pm_paytyp=4
							IF address.ad_compnum<10000
								IF UPPER(ALLTRIM(lcautodebit))="J"
									debnr=LOOKUP(id.id_last,"DEBITOR",id.id_code)
									IF debnr=0
										debnr=VAL(lcedebitor)
									else
										replace address.ad_compnum	WITH debnr
										replace id.id_last			WITH debnr+1
									endif
								ELSE
									debnr=VAL(lcedebitor)
								ENDIF
							ELSE
								debnr=address.ad_compnum
							ENDIF
							replace datev.debitornr		WITH ALLTRIM(STR(debnr,10,0))
							replace datev.firma			WITH address.ad_company
							replace datev.anrede		WITH address.ad_titlcod
							replace datev.suchname		WITH address.ad_compkey
							replace datev.na_vo			WITH ALLTRIM(address.ad_lname)+", "+ALLTRIM(address.ad_fname)
							replace datev.strasse		WITH address.ad_street
							replace datev.telefon		WITH address.ad_phone
							replace datev.plz			WITH address.ad_zip
							replace datev.ort			WITH address.ad_city
							replace datev.land			WITH address.ad_country
						endif
	 				*endif
 				
 		ENDCASE
 		SELECT histpost
 		SKIP
 	enddo	
 	* Hier wird die Datev.dbf aufgebaut und in die entsprechenden Datev.datei umgesetzt
 	SELECT datev
 	DELETE ALL FOR umsatz=0
 	IF lcnetto="J"
 		replace ALL umsatz WITH umsatz+mwstbetrag
 	ENDIF
 	GO top
 	*
	SET filter to
	* Hier wird die Debitorennr übertragen 
	replace ALL konto WITH debitornr FOR konto="0" AND !EMPTY(debitornr)
	* replace ALL gegenkonto WITH debitornr FOR gegenkonto="0" AND !EMPTY(debitornr)
	* Herausgenommen zu testzwecken ggf. wieder einsetzen
	
 	* Hier wird die RgNr. geprüft ob diese in Ordnung ist, da wenn eine Rechnung abgeschlossen ist und ich dann eine
 	* Artikel ins andere Fenster  verschiebe wird die rgnr mit übergeben obwohl das 2. Fenster ohne Rgnr ausgescheckt wurde.
 	DO WHILE EOF()=.f.
 		IF datev.reserid>1
 			ok_billnum=.f.
 			SELECT billnum
 			SET ORDER TO tag2
 			SEEK datev.reserid 
 			IF FOUND()
	 			DO WHILE datev.reserid = billnum.bn_reserid AND EOF()=.f.
 					IF datev.rgfenster = billnum.bn_window
 						ok_billnum=.t.
 						IF empty(datev.beleg1)
 							replace datev.beleg1	with billnum.bn_billnum
 						endif
 						EXIT
	 				ENDIF
	 				SKIP
 				ENDDO
 			endif
 			IF ok_billnum=.f.
 				replace datev.beleg1 WITH ""
 			ENDIF
 			IF EMPTY(datev.beleg1)
 				replace datev.beleg1	WITH ALLTRIM(STR(datev.reserid,12,3))
 			endif
 		ENDIF
 		SELECT datev
 		skip
 	ENDDO
 	*
	* fehlerprotokoll erstellen
	*
	* INDEX ON beleg1+STR(rgfenster,1)+STR(reserid,12,3)+STR(artikel,4,0) TO oli9
	SET ORDER TO tag3
	GO top
	DO WHILE EOF()=.f.
		IF  datev.reserid>1
			Nzahlung=0
			Nartikel=0
			nvrecno=RECNO()
			kschl=datev.beleg1+STR(datev.rgfenster,1)+STR(datev.reserid,12,3)
			DO WHILE kschl=datev.beleg1+STR(datev.rgfenster,1)+STR(datev.reserid,12,3)
				IF  datev.artikel>0
					nartikel= nartikel+datev.umsatz
				ELSE
					nzahlung=nzahlung+datev.umsatz
				ENDIF
				SKIP
			ENDDO
			IF nzahlung <> nartikel
				GO nvrecno
				DO WHILE kschl=beleg1+STR(rgfenster,1)++STR(reserid,12,3)
					replace fprotokoll WITH .t.
					replace btext WITH "Rechnung nicht Null oder Splittartikel fehlen"
					SKIP
				ENDDO
			ENDIF
	   else	
	   		sKIP
	   endif
	ENDDO
	* 
	* hier wird der Zahlungswechsel geprüft ob der in Ordnung ist
	*
	SET FILTER TO reserid=0.500
	GO top
	DO WHILE EOF()=.f.
		nzahlung=0
		nvrecno=RECNO()
		pdate=CTOD("")
		DO WHILE (pdate=datum OR EMPTY(pdate))
			pdate=datum
			nzahlung=nzahlung+umsatz
			SKIP
		ENDDO
		IF nzahlung<>0
			GO nvrecno
			DO WHILE pdate=datum
				replace fprotokoll WITH .t.
				replace btext WITH "Beim FW fehlt die Gegenbuchung"
				SKIP
			ENDDO
		ENDIF
	ENDDO
	*
	*
	* Hier wird die Buchungsdatei richtig aufgebaut und zwar so, dass 
	* kein verrechnungskonto mehr benötigt wird
	* 14101 Schlingmeier   8401 Logis 			80,00 
	* 14101 Schlingmeier   8402 Frühstück 		20,00
	*
	* Herausgenommen, da zwingend immer verrechnungskonto genommen werden muss
	* da es zwei Zahlungen für eine Rechnung geben kann
	ok_f=.f.
	IF ok_f=.t.
	*
	SET FILTER TO datev.reserid>1 OR (datev.reserid=0.100 AND datev.argus=.f.)&& Rechnung aus Brilliant
	GO top
	DO WHILE EOF()=.f.
		Nzahlung=0
		Nkonto=""
		kschl=datev.beleg1+STR(datev.rgfenster,1)+STR(datev.reserid,12,3)
		DO WHILE kschl=datev.beleg1+STR(datev.rgfenster,1)+STR(datev.reserid,12,3)
			IF datev.zahlung >0
				IF datev.konto="0" OR EMPTY(datev.konto)
					*nkonto= debitornr
					nkonto=iif(empty(datev.debitornr),lcedebitor,datev.debitornr)
				else
					nkonto = datev.konto
				endif
				replace datev.lkz	WITH .t.
			ENDIF
			IF !EMPTY(nkonto) AND datev.zahlung=0
				replace datev.konto WITH nkonto
			endif
			SKIP
		ENDDO
	ENDDO
	*
	ENDIF
	*
	SET FILTER TO 

	* hier werden nur die Umsätze aus dem Hotel nicht aus Argus übermittelt. 	
 	SELECT datev
 	SET Order to tag3
 	GO top
 	* brow
 	ntotalamount = 0
 	cbuffer=""
 	DO WHILE (!EOF("Datev"))
 		IF (datev.reserid>1 OR (datev.reserid<1 AND datev.argus=.f.)) AND datev.lkz=.f. ;
 		AND datev.fprotokoll=.f.
	 		ntotalamount = ntotalamount + datev.umsatz
 			cbuffer=cbuffer+&lcSatz+CHR(13)+CHR(10)
 		endif
 		SELECT datev
 		SKIP
 	ENDDO
 	IF LEN(cbuffer)>0
 		*
 		tDatei=ALLTRIM(lcverz)+"KHK-Hotel"+DTOS(dstart)+".txt"
 		nb=FCREATE(tdatei)
 		IF nb==-1
 			WAIT WINDOW tdatei
 			=alert(getlangtext("DATEV","TXT_CREATE_ERROR"))
 		ELSE
 			=FWRITE(nb,lckopf+CHR(13)+CHR(10))
 			=FWRITE(nb,cbuffer)
 			=FCLOSE(nb)
 		endif
 	endif
	* hier werden nur die Umsätze aus Argus übermittelt
	kasse=.f.
	* rausnehmen wenn kasse wieder aktiv ist
	
	IF thisform.arGUS.Value=1
	*
 	SELECT datev
 	SET Order to tag5
 	GO top
 	* brow
 	ntotalamount = 0
 	cbuffer=""
 	DO WHILE (!EOF("Datev"))
 		IF datev.reserid<1 AND datev.argus=.t. AND datev.lkz=.f. AND datev.fprotokoll=.f.
	 		ntotalamount = ntotalamount + datev.umsatz
 			cbuffer=cbuffer+&lcSatz+CHR(13)+CHR(10)
 		endif
 		SELECT datev
 		SKIP
 	ENDDO
 	IF LEN(cbuffer)>0
 		*
 		tDatei=ALLTRIM(lcverz)+"KHK-Kasse"+DTOS(dstart)+".txt"
 		nb=FCREATE(tdatei)
 		IF nb==-1
 			WAIT WINDOW tdatei
 			=alert(getlangtext("DATEV","TXT_CREATE_ERROR"))
 		ELSE
 			=FWRITE(nb,lckopf+CHR(13)+CHR(10))
 			=FWRITE(nb,cbuffer)
 			=FCLOSE(nb)
 		endif
 	ENDIF
 	SELECT datev
 	SET ORDER TO 3
 	* Feherprotokoll erstellen
 	* Kasse
	tdatei=alltrim(lcverz)+"KHK-K-ERR"+DTOS(dstart)
	COPY TO &tdatei FOR fprotokoll=.t. AND datev.reserid<1 AND datev.argus=.t. TYPE xls
	*REPORT FORM report\dateverrprt HEADING "Fehlerprotokoll Kassenbuchungen" FOR datev.reserid<1 AND argus=.t. AND fprotokoll=.t. PREVIEW IN SCREEN 		
	*
	endif
	* Hotel
	tdatei=ALLTRIM(lcverz)+"KHK-H-ERR"+DTOS(dstart)
	COPY TO &tdatei FOR datev.reserid>1 OR (datev.reserid=0.100 AND datev.argus=.f.) TYPE xls
	*REPORT FORM report\dateverrprt HEADING "Fehlerprotokoll Hotelbuchungen" FOR datev.reserid>1 OR (datev.reserid=0.100 AND datev.argus=.f.)and fprotokoll=.t. PREVIEW IN SCREEN 	
 	SET NEAR &old_near
 	SELECT histres
 	SET order to &norder_hr
 	SELECT histpost
 	SET order to &norder_hp
 RETURN .t.
 		