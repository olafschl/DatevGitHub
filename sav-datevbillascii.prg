* Datevbillascii

* Dies ist die neue Datevbill mit ASCII übergabe. Die alte Schnittstelle wird mit dieser abgelöst.
*
* Dateiname für Buchungsdatei muss lauten: EXTF_Buchungen.csv
*
* Dateiname für die Debitorenbeschriftung muss lauten: EXTF_Stammdaten.csv
*
* neu hinugekommen am: 16.12.2011
*

 * In dieser Funktion werden nur Rechnungen die
 * abgeschlossen sind übergeben lt. billnum.dbf
  	* Sofern nur abgeschlossene Rechnungen übergeben werden sollen, muss dann das 
 	* Buchungsdatum das Rechnungsdatum sein, da ansonsten Buchungen aus dem Vormonat 
 	* in der aktuellen übergabe übergeben wird, was bei datev ein Fehler auslöst.
 	* Beispiel: Anreise 28.12. Abreise 04.01. Buchungslauf 01.01. - 31.01 in diesem
 	* lauf werden die Buchungen vom 28.12. - 31.12. mit übernommen was korrekt ist. Das 
 	* Buchungsdatum muss aber das Rechnungsdatum sein
 	* WAIT WINDOW "lcabgrg= "+lcabgrg
 
  * In dieser Funktion werden nur Rechnungen die
 * abgeschlossen sind übergeben lt. billnum.dbf
 PARAMETERS Dstart,Dende
 LOCAL norder_hr,norder_hp,old_near,ok_ad,lmnr,old_near
 LOCAL caMacro, cpMacro ,ckstkonto,cpaykonto,cpaytext,debnr
 LOCAL pr_anz,pr_summe,lanrede,dsatz
 	* WAIT WINDOW "Datevbill"
 	dsatz=""
	caMacro = "Article.Ar_Lang"+g_Langnum
    cpMacro = "Paymetho.Pm_Lang"+g_Langnum
	old_near=SET("near")
	pr_anz=0
	pr_summe=0
	SET NEAR on
  	SELECT histpost
 	norder_hp=ORDER()
 	SET ORDER TO 1
 	SELECT histres
 	norder_hr=ORDER()
 	SET ORDER TO 2
 	GO top
 	SELECT billnum
 	*INDEX on DTOS(bn_date) TO billnum1
 	SET ORDER to tag3
 	*SEEK DTOS(dstart)
 	SEEK dstart
 	DO WHILE EOF()=.f. AND BETWEEN(billnum.bn_date,dstart,dende) 
 		IF billnum.bn_status<>"PCO"
 			SKIP
 			LOOP
 		ENDIF
 		SELECT histres
 		SEEK billnum.bn_reserid ORDER tag1
 		SELECT address
 		SEEK billnum.bn_addrid
 		IF !FOUND()
 			ok_ad=.f.
 		ELSE
 			ok_ad=.t.
 		ENDIF
 		
 		*IF VAL(lcmnummer)>0
 		*	lmnr=(lcmnummer)
 		*ELSE
	 	*	lmnr= LOOKUP(roomtype.rt_buildng,histres.hr_roomtyp,roomtype.rt_roomtyp)
	 	*ENDIF
	 	* Ausgeschaltet 24.01.2011 da Mehrmandantenfähigkeit über Feld article.ar_user3 und paymetho.pm_user3 ermittelt wird
	 		
 		SELECT histpost
 		SEEK billnum.bn_reserid
 		DO WHILE EOF()=.f. and billnum.bn_reserid=histpost.hp_reserid
 			* Als Parameter einbinden
 			IF (BETWEEN(histpost.hp_date,histres.hr_arrdate,histres.hr_depdate)=.f.;
 			AND lcResZeitraum="J") 
 				skip
 				LOOP
 			ENDIF
 			IF histpost.hp_billnum <> billnum.bn_billnum 
 				SKIP
 				LOOP
 			ENDIF
 			IF histpost.hp_billnum="2210076262"
 				WAIT WINDOW "zahlung = "+STR(histpost.hp_paynum,3,0)+"  Betrag= "+STR(histpost.hp_amount,12,2)+" addrid= "+billnum.bn_addrid
 			endif
	 		DO case
 				CASE ((histpost.hp_split=.f. and EMPTY(histpost.hp_ratecod)) or (!EMPTY(histpost.hp_ratecod) and histpost.hp_split=.t.));
 				AND histpost.hp_reserid>0 AND !histpost.hp_cancel AND !EMPTY(histpost.hp_artinum)
					=SEEK(histpost.hp_artinum,"Article","Tag1")
					lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,article.ar_user3)
					IF histpost.hp_reserid=366.100
						*WAIT WINDOW "hp_reserid=366.100 in postschleife"
					endif
					SELECT datev
					IF article.ar_artityp<>9 
						IF lcverdichtung="J"
							hfn=histpost.hp_artinum
							keyschl=histpost.hp_billnum+STR(VAL(lmnr),5,0)+STR(hfn,4,0)
							IF !SEEK(keyschl,"Datev","Tag7")
								APPEND BLANK 
							ELSE
								* wait window "!!! gefunden"
							ENDIF
						ELSE
							* keine Verdichtung
							APPEND BLANK 
						ENDIF
						&& Mehrwetsteuercode PLNUMCOD = 8 festgesetz muss geändert werden bei Steueränderung
						* IF SEEK("VATGROUP  "+STR(8,3),"Picklist","Tag3")=.t.
						IF SEEK("VATGROUP  "+STR(article.ar_vat,3),"Picklist","Tag3")=.t.
							replace datev.mwstsatz	WITH picklist.pl_numval
						ELSE
							replace datev.mwstsatz 	WITH 0
						endif
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
							CASE article.ar_vat = 3
								replace datev.mwst with lcmwst3
						OTHERWISE
							replace datev.mwst WITH ""
						endcase
						replace datev.benutzer	WITH histpost.hp_userid	
						replace datev.artikel 	WITH histpost.hp_artinum
						replace datev.datum		WITH billnum.bn_date
						replace datev.btext		WITH &caMacro
						replace datev.konto		WITH lcvartikel
						replace datev.beleg2	WITH IIF(lcaZtext="J",&lcaZTFeld,"") && Beleg2 ist ein Individuelles Feld 
						replace datev.bkz		WITH "R"
						replace datev.rgfenster	WITH billnum.bn_window
						replace datev.mandantnr	WITH val(lmnr)				
						replace datev.rechnr	WITH billnum.bn_billnum
						replace datev.reserid	WITH histpost.hp_reserid
						replace datev.adressid	WITH histres.hr_addrid
						replace datev.abreise	WITH histres.hr_depdate
						replace datev.anreise	WITH histres.hr_arrdate
						replace datev.rname		WITH histres.hr_lname
						replace datev.rfirma	WITH histres.hr_company
						replace datev.beleg1	WITH billnum.bn_billnum
						replace datev.zimmernr	WITH LOOKUP(room.rm_rmname,histres.hr_roomnum,room.rm_roomnum,"TAG1")
						IF histpost.hp_userid="POS" OR histpost.hp_userid="POSZ2"
							replace datev.argus	WITH .t.
						endif
						*WAIT WINDOW "datev.rechnr= "+datev.rechnr
						IF !empty(article.ar_user2)
							replace datev.gegenkonto 	WITH ALLTRIM(article.ar_user2)
						ELSE
							replace datev.gegenkonto	WITH ALLTRIM(lceartikel)
						ENDIF
						IF UPPER(ALLTRIM(lckostenstelle))="J"
							IF !EMPTY(article.ar_user1)
								replace datev.kostenstelle	WITH ALLTRIM(article.ar_user1)
							ELSE
								replace datev.kostenstelle	WITH ALLTRIM(lcekostenstelle)
							ENDIF
						ENDIF
					endif
	 				IF article.ar_artityp<>9
 						replace datev.umsatz WITH datev.umsatz+histpost.hp_amount
 						replace datev.mwstbetrag WITH datev.mwstbetrag+histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                        	   	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
 					ENDIF
	 			*CASE (!EMPTY(histpost.hp_paynum) AND histpost.hp_reserid>0 AND !histpost.hp_cancel) and histpost.hp_paynum<>param.pa_payonld;
	 			*AND !INLIST(histpost.hp_paynum,&lcazahlungv) AND ;
	 			*(!(histpost.hp_reserid=0.100 AND histpost.hp_window=0) AND !(histpost.hp_window=0 AND histpost.hp_reserid>1))  &&BETWEEN(histpost.hp_paynum,VAL(lcazahlungv),VAL(lcazahlungb))=.f.
	 			CASE (!EMPTY(histpost.hp_paynum) AND histpost.hp_reserid>0 AND !histpost.hp_cancel) and histpost.hp_paynum<>param.pa_payonld;
	 			AND !INLIST(histpost.hp_paynum,&lcazahlungv)
	 				
	 				SELECT datev
	 				SET ORDER TO 2
	 				IF LOOKUP(paymetho.pm_paytyp,histpost.hp_paynum,paymetho.pm_paynum)==4
	 					
	 					SELECT address
	 					SEEK billnum.bn_addrid ORDER tag1
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
	 					SELECT address
	 					SEEK billnum.bn_addrid  ORDER tag1 && gesetzt, da immer die Addresse in der datev-export datei gespeichert werden soll
	 					IF !FOUND()
	 						*WAIT WINDOW "Adresse nicht gefunden "+STR(billnum.bn_addrid,10,0)
	 					endif
	 					* eingeführt für SAP Schnittstelle 
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
	 				lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,paymetho.pm_user3)
	 				SELECT datev
	 				SET ORDER TO tag2
	 				*IF INLIST(paymetho.pm_paytyp,3,4) OR !SEEK(DTOS(histpost.hp_date)+STR(histpost.hp_paynum,2),"Datev")
	 					APPEND BLANK
	 					replace datev.benutzer			WITH histpost.hp_userid
						replace datev.rechnr			WITH billnum.bn_billnum							
						replace datev.zahlung			WITH histpost.hp_paynum
						replace datev.kostenstelle		WITH IIF(UPPER(ALLTRIM(lckostenstelle))="J",ckstkonto,"")
						replace datev.konto				WITH IIF(paymetho.pm_paytyp=4,"0",cpaykonto)
						replace datev.gegenkonto		WITH lcvzahlung
						replace datev.abreise			WITH histres.hr_depdate
						replace datev.anreise			WITH histres.hr_arrdate
						replace datev.reserid			WITH billnum.bn_reserid
						replace datev.adressid			with billnum.bn_addrid
						IF histres.hr_reserid=0.100
							cpaytext="Passantenbuchung"
						ELSE
							cpaytext= ALLTRIM(address.ad_company)+" " +ALLTRIM(address.ad_fname)+" "+ALLTRIM(address.ad_lname)
						ENDIF
						replace datev.btext				WITH cpaytext
						replace datev.datum				WITH billnum.bn_date
						replace datev.bkz				WITH "R" &&IIF(VAL(cpaykonto)>9999,"R","") && damit Kreditzahlung als Debitor laufen
						replace datev.umsatz			WITH datev.umsatz+(histpost.hp_amount*-1)
						replace datev.mandantnr			WITH val(lmnr)
						replace datev.rgfenster			WITH billnum.bn_window
						replace datev.beleg1			WITH billnum.bn_billnum
						replace datev.beleg2			WITH IIF(lcaZtext="J",&lcaZTFeld,"") && Beleg2 ist ein Individuelles Feld 
						replace datev.zimmernr			WITH LOOKUP(room.rm_rmname,histres.hr_roomnum,room.rm_roomnum,"TAG1")
						*
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
							replace datev.konto			WITH ALLTRIM(STR(debnr,10,0))  && geändert, da automatisch das Debitorenkonto gesetzt werden muss  09.12.2017
						ENDIF
						* Neu gesetzt 13.10.2017 Adressdaten werden zu allen Zahlungswegen benötigt
						replace datev.firma			WITH address.ad_company
						replace datev.firma2		WITH address.ad_departm
						replace datev.anrede		WITH address.ad_titlcod
						replace datev.suchname		WITH address.ad_compkey
						replace datev.na_vo			WITH ALLTRIM(address.ad_lname)+", "+ALLTRIM(address.ad_fname)
						replace datev.fname			WITH ALLTRIM(address.ad_fname)
						replace datev.lname			WITH ALLTRIM(address.ad_lname)
						replace datev.strasse		WITH address.ad_street
						replace datev.telefon		WITH address.ad_phone
						replace datev.telefax		WITH address.ad_fax
						replace datev.email			WITH address.ad_email
						replace datev.plz			WITH address.ad_zip
						replace datev.ort			WITH address.ad_city
						replace datev.land			WITH address.ad_country
						* endif
	 				*endif
 			ENDCASE
			SELECT histpost
 			SKIP 1
 		ENDDO
 		SELECT billnum
 		SKIP 1 
 	ENDDO
 	* Hier werden die Anzahlungen/Deposits verbucht!
 	SELECT histpost
 	GO top
 	SET ORDER TO tag2
 	SEEK dstart
 	DO WHILE EOF()=.f. and BETWEEN(histpost.hp_date,dstart,dende)
 		IF histpost.hp_window=0 AND histpost.hp_artinum=9999
 			* Hier wird der Artikel Deposit verbucht
 			=SEEK(histpost.hp_artinum,"Article","TAG1")
 			lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,article.ar_user3)
 			SELECT datev
 			APPEND BLANK
 			*
			DO case
				CASE article.ar_vat = 8
					replace datev.mwst with lcmwst8
				CASE article.ar_vat = 9
					replace datev.mwst with lcmwst9
				CASE article.ar_vat = 2
					replace datev.mwst with lcmwst2
				CASE article.ar_vat = 3
					replace datev.mwst with lcmwst3
			OTHERWISE
				replace datev.mwst WITH ""
			endcase
			* Hier muss noch auf die -reservat zugegriffen werden für Name/Firma	
			* Deposit reservierungen befinden sich in der reservat, da es sich
			* um zukünftige reservierungen handelt!
			m_lname=""
			m_company=""
			IF SEEK(histpost.hp_reserid,"Histres","TAG1")=.t.
				* Werte übergabe
				m_company=histres.hr_company
				m_lname = histres.hr_lname
			ELSE
				IF SEEK(histpost.hp_reserid,"reservat","Tag1")=.t.
					m_company = reservat.rs_company
					m_lname	  = reservat.rs_lname
				ENDIF
			ENDIF
			replace datev.benutzer	WITH histpost.hp_userid
			replace datev.beleg1	WITH STRTRAN(STR(histpost.hp_reserid,12,3),",","")
			replace datev.artikel 	WITH histpost.hp_artinum
			replace datev.datum		WITH histpost.hp_date && musste gesetzt werden, damit buchungsdatum gleich rechnungsdatum ist
			replace datev.btext		WITH "Deposit: "+ALLTRIM(m_lname)+" "+alltrim(m_company)
			replace datev.konto		WITH lcvartikel
			replace datev.beleg2	WITH IIF(lcaZtext="J",&lcaZTFeld,"") && Beleg2 ist ein Individuelles Feld 
			replace datev.bkz		WITH ""
			replace datev.rgfenster	WITH histpost.hp_window
			replace datev.mandantnr	WITH val(lmnr)				
			replace datev.rechnr	WITH STRTRAN(STR(histpost.hp_reserid,12,3),",","")
			replace datev.reserid	WITH histpost.hp_reserid
			replace datev.adressid	WITH 0
			*replace datev.abreise	WITH histres.hr_depdate
			*replace datev.anreise	WITH histres.hr_arrdate
			replace datev.rname		WITH m_lname
			replace datev.rfirma	WITH m_company
			IF !empty(article.ar_user2)
				replace datev.gegenkonto 	WITH ALLTRIM(article.ar_user2)
			ELSE
				replace datev.gegenkonto	WITH ALLTRIM(lceartikel)
			ENDIF
			IF UPPER(ALLTRIM(lckostenstelle))="J"
				IF !EMPTY(article.ar_user1)
					replace datev.kostenstelle	WITH ALLTRIM(article.ar_user1)
				ELSE
					replace datev.kostenstelle	WITH ALLTRIM(lcekostenstelle)
				ENDIF
			endif
			* IF histpost.hp_artinum<>VAL(lcaArtikel) 
			IF !INLIST(histpost.hp_artinum,&lcaArtikel)
				replace datev.umsatz WITH datev.umsatz+histpost.hp_amount
				replace datev.mwstbetrag WITH datev.mwstbetrag+histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                   	   	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
			ENDIF
 			
 			m_postid=histpost.hp_postid
 			SELECT histpost
 			SKIP
 			IF histpost.hp_postid=m_postid+1
 				* Hier wird die Zahlung Deposti verbucht
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
 				lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,paymetho.pm_user3)
 				SELECT datev
 				SET ORDER TO tag2
				APPEND BLANK
				replace datev.benutzer			WITH histpost.hp_userid
				replace datev.beleg1			WITH STRTRAN(STR(histpost.hp_reserid,12,3),",","")
				replace datev.beleg2			WITH IIF(lcaZtext="J",&lcZaTFeld,"") && Beleg2 ist ein Individuelles Feld 
				replace datev.rechnr			WITH STRTRAN(STR(histpost.hp_reserid,12,3),",","")
				replace datev.zahlung			WITH histpost.hp_paynum
				replace datev.kostenstelle		WITH IIF(UPPER(ALLTRIM(lckostenstelle))="J",ckstkonto,"")
				replace datev.konto				WITH IIF(paymetho.pm_paytyp=4,"0",cpaykonto)
				replace datev.gegenkonto		WITH lcvzahlung
				*replace datev.abreise			WITH reservat.rs_depdate
				replace datev.rname				WITH m_lname
				replace datev.rfirma			WITH m_company
				*replace datev.anreise			WITH reservat.rs_arrdate
				replace datev.reserid			WITH histpost.hp_reserid
				replace datev.adressid			with 0
				replace datev.btext				WITH "Deposit: "+ALLTRIM(m_lname)+" "+ALLTRIM(m_company)
				replace datev.datum				WITH histpost.hp_date
				replace datev.bkz				WITH "" && IIF(VAL(cpaykonto)>9999,"R","") && damit Kreditzahlung als Debitor laufen
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
							debnr=VAL(lcedebitor)
						ENDIF
					ELSE
						debnr=address.ad_compnum
					ENDIF
					replace datev.debitornr		WITH ALLTRIM(STR(debnr,10,0))
					replace datev.firma			WITH address.ad_company
					replace datev.firma2		WITH address.ad_departm
					replace datev.anrede		WITH address.ad_titlcod
					replace datev.suchname		WITH address.ad_compkey
					replace datev.na_vo			WITH ALLTRIM(address.ad_lname)+", "+ALLTRIM(address.ad_fname)
					replace datev.vorname		WITH ALLTRIM(address.ad_fname)
					replace datev.name			WITH ALLTRIM(address.ad_lname)
					replace datev.strasse		WITH address.ad_street
					replace datev.telefon		WITH address.ad_phone
					replace datev.telefax		WITH address.ad_fax
					replace datev.email			WITH address.ad_email
					replace datev.plz			WITH address.ad_zip
					replace datev.ort			WITH address.ad_city
					replace datev.land			WITH address.ad_country
				endif
 				
 			ENDIF
 		ENDIF
 		SELECT histpost
 		SKIP
 	enddo
  	* Hier werden die Auslagen 0.200 und Ausgaben 0.300 verbucht
 	SELECT histpost
 	GO top
 	SET ORDER TO tag2
 	SEEK dstart
 	*WAIT WINDOW "hier werden die Auslagen und Ausgaben verbucht!!!"
 	DO WHILE EOF()=.f. AND BETWEEN(histpost.hp_date,dstart,dende)
 		IF histpost.hp_reserid <> 0.200 and histpost.hp_reserid<>0.300 AND (histpost.hp_reserid<>0.400 or histpost.hp_cashier=99)
 		*IF histpost.hp_reserid <> 0.200 and histpost.hp_reserid<>0.300 && 05.03.2015
 			skip
 			LOOP
 		ENDIF
 		SELECT datev
 		APPEND BLANK
 		cpaykonto = LOOKUP(paymetho.pm_user2,histpost.hp_paynum,paymetho.pm_paynum)
 		ckstkonto = LOOKUP(paymetho.pm_user1,histpost.hp_paynum,paymetho.pm_paynum) 	
 		cktr	  = LOOKUP(paymetho.pm_user3,histpost.hp_paynum,paymetho.pm_paynum)  && 05.03.2015
 		
 		IF VAL(lcmnummer)>0
 			lmnr = lcmnummer
 		else
	 		lmnr = LOOKUP(paymetho.pm_user3,histpost.hp_paynum,paymetho.pm_paynum) 	
	 	ENDIF
 	
 		replace datev.benutzer			WITH histpost.hp_userid	
 		replace datev.zahlung			WITH histpost.hp_paynum
 		replace datev.abreise			WITH histpost.hp_date
 		replace datev.adressid			WITH histpost.hp_addrid
 		replace datev.kostenstelle		WITH IIF(UPPER(ALLTRIM(lckostenstelle))="J",ckstkonto,"")
 		replace datev.auslagen 			WITH IIF(histpost.hp_reserid=0.200,.t.,.f.)
 		replace datev.beleg1			WITH IIF(!EMPTY(histpost.hp_billnum),histpost.hp_billnum,ALLTRIM(STR(histpost.hp_postid,8,0)))
 		replace datev.datum				WITH histpost.hp_date
 		replace datev.btext				WITH IIF(histpost.hp_reserid=.200,"Auslage -> ","Ausgabe-> ")+ALLTRIM(histpost.hp_supplem)
		IF histpost.hp_reserid=0.200
	 		replace datev.konto				WITH IIF(histpost.hp_reserid=.200,lcvAuslagen,lcvAusgaben)
 			replace datev.gegenkonto		WITH cpaykonto
	 		replace datev.umsatz			WITH histpost.hp_amount  && *-1 da Konto Gegenkonto vertauscht wurde 			
 		Endif
 		IF histpost.hp_reserid=0.300	
	 		*replace datev.gegenkonto	WITH IIF(histpost.hp_reserid=.200,lcvAuslagen,lcvAusgaben)
 			*replace datev.	konto		WITH cpaykonto
	 		*replace datev.umsatz		WITH histpost.hp_amount *-1
	 		* Neu gesetzt 28.05.2012
	 		replace datev.gegenkonto 	WITH  cpaykonto
	 		replace datev.konto			WITH IIF(!EMPTY(histpost.hp_finacct),ALLTRIM(STR(histpost.hp_finacct)),lcvAusgaben)
	 		replace datev.umsatz		WITH histpost.hp_amount &&*-1
		endif 			
		IF histpost.hp_reserid=0.400 && Buchung von und nach Hausbank am 10.10.2014 gesetzt
			IF histpost.hp_amount >0
				replace datev.konto			WITH lcvHausbank
				replace datev.gegenkonto	WITH cpaykonto
				replace datev.umsatz		WITH histpost.hp_amount
				replace datev.btext			WITH "nach Hausbank <<- "+ALLTRIM(histpost.hp_supplem)
			ELSE
				replace datev.konto			WITH cpaykonto
				replace datev.gegenkonto	WITH lcvHausbank
				replace datev.umsatz		WITH histpost.hp_amount*-1
				replace datev.btext			WITH "von Hausbank ->> "+ALLTRIM(histpost.hp_supplem)
			ENDIF
		endif
 		replace datev.reserid			WITH histpost.hp_reserid

 		replace datev.mandantnr			WITH VAL(lmnr)
 		replace datev.rgfenster			WITH histpost.hp_window
 		SELECT histpost
 		SKIP
 	ENDDO
 	SELECT datev
 
 	* Hier wird die Passantenbuchungen eingebunden wichtig für Argus Z-Abschlag
 	SELECT histpost
 	GO  top
 	SET ORDER TO tag2
 	SEEK dstart
 	DO WHILE EOF()=.f. and BETWEEN(histpost.hp_date,dstart,dende)
 		*IF histpost.hp_userid<>"POSZ2" 
 		IF histpost.hp_reserid<>0.100 OR (histpost.hp_reserid=0.100 AND  VAL(histpost.hp_billnum)>999)
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
				IF article.ar_artityp<>9
					lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,article.ar_user3)	
					IF lcverdichtung="J"
						*
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
						CASE article.ar_vat = 3
							replace datev.mwst with lcmwst3
					OTHERWISE
						replace datev.mwst WITH ""
					endcase
					replace datev.benutzer	WITH histpost.hp_userid
					replace datev.artikel 	WITH histpost.hp_artinum
					replace datev.datum		WITH histpost.hp_date
					replace datev.btext		WITH &caMacro
					replace datev.konto		WITH lcvartikel
					replace datev.beleg2	WITH IIF(lcaZtext="J",&lcaZTFeld,"") && Beleg2 ist ein Individuelles Feld 
					replace datev.bkz		WITH ""
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
 				IF article.ar_artityp<>9
					replace datev.umsatz WITH datev.umsatz+histpost.hp_amount
					replace datev.mwstbetrag WITH datev.mwstbetrag+histpost.hp_vat0+histpost.hp_vat1+histpost.hp_vat2+histpost.hp_vat3+histpost.hp_vat4+histpost.hp_vat5+;
                    	   	histpost.hp_vat6+histpost.hp_vat7+histpost.hp_vat8+histpost.hp_vat9
				ENDIF
 				
 			CASE histpost.hp_paynum>0 AND  !histpost.hp_cancel AND !INLIST(histpost.hp_paynum,&lcazahlungv)
 				* lcaZahlungv ist eingebunden worden, damit Debitorenzahlungen nicht berücksichtigt werden.
 				* in lcazahlungv steht z.B. 13,98 d.h. Finanzweg 13=Banküberweisung und die Gegenbuchung FW=98 Payment on Ledger
 				* werden nicht berücksichtigt. Im Feld lcazahlungv werdne die Finanzwege mit Komma getrennt eingetragen die nicht
 				* übergeben werden sollen.
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
 					lmnr = IIF(VAL(lcmnummer)>0,lcmnummer,paymetho.pm_user3)	 				
	 				SELECT datev
	 				SET ORDER TO tag2
	 				*IF INLIST(paymetho.pm_paytyp,3,4) OR !SEEK(DTOS(histpost.hp_date)+STR(histpost.hp_paynum,2),"Datev")

	 					APPEND BLANK
	 					IF EMPTY(histpost.hp_billnum)
	 						IF !EMPTY(histpost.hp_ifc)
	 							replace datev.beleg1	WITH ALLTRIM(histpost.hp_ifc)
	 						else
		 						replace datev.beleg1		WITH ALLTRIM(STR(histpost.hp_postid,8,0))
		 					endif
	 					ELSE 
		 					replace datev.beleg1			WITH histpost.hp_billnum
		 				ENDIF
		 				replace datev.benutzer			WITH histpost.hp_userid
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
								cpaytext="Passantenbuchung"
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
						replace datev.bkz				WITH "" && IIF(VAL(cpaykonto)>9999,"R","") && damit Kreditzahlung als Debitor laufen
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
							replace datev.firma2		WITH address.ad_departm
							replace datev.anrede		WITH address.ad_titlcod
							replace datev.suchname		WITH address.ad_compkey
							replace datev.na_vo			WITH ALLTRIM(address.ad_lname)+", "+ALLTRIM(address.ad_fname)
							replace datev.fname			WITH address.ad_fname
							replace datev.lname			WITH address.ad_lname
							replace datev.strasse		WITH address.ad_street
							replace datev.telefon		WITH address.ad_phone
							REPLACE datev.telefax		WITH address.ad_fax
							replace datev.email			WITH address.ad_email
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
 	IF ALLTRIM(lcaArtikel)="0"
 		* WAIT WINDOW "kein Artikel als ausnahmeartikel angelegt!!! "
 	else
 		DELETE ALL FOR INLIST(datev.artikel,&lcaArtikel)  && hier werden die Ausnahmeartikel gelöscht!!!
 	endif
 	IF lcnetto="J"
 		replace ALL umsatz WITH umsatz+mwstbetrag
 	ENDIF
 	GO top
 	*
 	* Sofern nur abgeschlossene Rechnungen übergeben werden sollen, muss dann das 
 	* Buchungsdatum das Rechnungsdatum sein, da ansonsten Buchungen aus dem Vormonat 
 	* in der aktuellen übergabe übergeben wird, was bei datev ein Fehler auslöst.
 	* Beispiel: Anreise 28.12. Abreise 04.01. Buchungslauf 01.01. - 31.01 in diesem
 	* lauf werden die Buchungen vom 28.12. - 31.12. mit übernommen was korrekt ist. Das 
 	* Buchungsdatum muss aber das Rechnungsdatum sein
 	* WAIT WINDOW "lcabgrg= "+lcabgrg
 	IF ALLTRIM(UPPER(lcabgrg))="J"
 		*replace ALL datev.datum WITH datev.abreise FOR !EMPTY(datev.abreise)
 	ENDIF
 	*

	SET filter to
	
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
 				replace datev.beleg1	WITH STRTRAN(ALLTRIM(STR(datev.reserid,12,3)),",","")
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
				IF datev.reserid=168269.100
					* WAIT WINDOW "nzahlung= "+STR(nzahlung,12,2)+"  nartikel= "+STR(nartikel,12,2)
				endif
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
	* Die Funktion herausgenommen, da zwingend Verrechnungskonto angelegt werden muss
	ok=.f.
	IF ok=.t.
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
    endif
	SET FILTER TO 
	*

 	* Hier wird die Datev.dbf aufgebaut und in die entsprechenden Datev.datei umgesetzt
 	SELECT datev
 	DELETE ALL FOR umsatz=0
 	*IF lcnetto="J"
 		* replace ALL umsatz WITH umsatz+mwstbetrag  
 	*endif  // Doppelt enthalten weiter oben steht die selbe Bedingung!!!!
 	SELECT datev
 	SET Order to tag5
 	GO top

 	pr_mandant=0
 	edzaehler=0
 	msumme=0
 	sl=0
 	nh=0
 	lcBuffer=""
 	IF ALLTRIM(UPPER(lcabgrg))="J"
 		*replace ALL datev.datum WITH datev.abreise FOR !EMPTY(datev.abreise)
 	ENDIF
 	*

	SET filter to
	
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
 				replace datev.beleg1	WITH STRTRAN(ALLTRIM(STR(datev.reserid,12,3)),",","")
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
				IF datev.reserid=168269.100
					* WAIT WINDOW "nzahlung= "+STR(nzahlung,12,2)+"  nartikel= "+STR(nartikel,12,2)
				endif
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
	* Die Funktion herausgenommen, da zwingend Verrechnungskonto angelegt werden muss
	ok=.f.
	IF ok=.t.
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
    endif
	SET FILTER TO 
	*
	*BROWSE 
 	* Hier wird die Datev.dbf aufgebaut und in die entsprechenden Datev.datei umgesetzt
 	SELECT datev
 	DELETE ALL FOR umsatz=0
 	IF lcnetto="J"
 		replace ALL umsatz WITH umsatz+mwstbetrag
 	endif
 	SELECT datev
 	SET Order to tag5
 	GO top
 	*
 	*  T e s t    bitte im orgiginal ausschalten
 	*  replace ALL datev.kostenstelle WITH "280163"
 	*
 	* ende test
 	*brow
 	pr_mandant=0
 	edzaehler=0
 	msumme=0
 	sl=0
 	nh=0
 	lcBuffer=""
 	* Hier wird die Buchungsdatei erstellt
 	ASCIIDAtei = lcVerz+"EXTF_Buchung.csv"
 	nbASCII = FCREATE(ASCIIDatei)
 
 	* Hier wird die Headerzeile für die Buchungsdatei aufgebaut
 	txtHeader = '"EXTF";510;21;"Buchungsstapel";7;'+ DTOS(DATE())+STRTRAN(TIME(),":","")+"810;;"+'"SV";"";"";'+ALLTRIM(lcbnummer)+";"+ALLTRIM(lcmnummer)+";"+STR(YEAR(Dstart),4,0)+lcWJ+";" && lcWJ Wirtschaftjahr 0101 MMJJ
 	txtHeader = txtHeader + lcgSachkl+";"+DTOS(dstart)+";"+DTOS(dende)+";"+'"Rechnungen";"";'+"1;0;;"+'"EUR";;;;;;;;;'&&+CHR(13)+CHR(10)
 	txtKopf	  = "Umsatz (ohne Soll/Haben-Kz);Soll/Haben-Kennzeichen;WKZ Umsatz;Kurs;Basis-Umsatz;WKZ Basis-Umsatz;Konto;Gegenkonto (ohne BU-Schlüssel);BU-Schlüssel;Belegdatum;Belegfeld 1;Belegfeld 2;"
 	txtKopf   = txtkopf + "Skonto;Buchungstext;Postensperre;Diverse Adressnummer;Geschäftspartnerbank;Sachverhalt;Zinssperre;Beleglink;Beleginfo - Art 1;Beleginfo - Inhalt 1;Beleginfo - Art 2;Beleginfo - Inhalt 2;"
 	txtKopf   = txtkopf + "Beleginfo - Art 3;Beleginfo - Inhalt 3;Beleginfo - Art 4;Beleginfo - Inhalt 4;Beleginfo - Art 5;Beleginfo - Inhalt 5;Beleginfo - Art 6;Beleginfo - Inhalt 6;Beleginfo - Art 7;"
 	txtKopf   = txtkopf + "Beleginfo - Inhalt 7;Beleginfo - Art 8;Beleginfo - Inhalt 8;KOST1 - Kostenstelle;KOST2 - Kostenstelle;Kost-Menge;EU-Land u. UStID;EU-Steuersatz;Abw. Versteuerungsart;Sachverhalt L+L;"
 	txtKopf	  = txtkopf + "Funktionsergänzung L+L;BU 49 Hauptfunktionstyp; BU 49 Hauptfunktionsnummer;BU 49 Funktionsergänzung;Zusatzinformation - Art 1;Zusatzinformation- Inhalt 1;Zusatzinformation - Art 2;Zusatzinformation- Inhalt 2;Zusatzinformation - Art 3;"
 	txtKopf   = txtkopf +"Zusatzinformation- Inhalt 3;Zusatzinformation - Art 4;Zusatzinformation- Inhalt 4;Zusatzinformation - Art 5;Zusatzinformation- Inhalt 5;Zusatzinformation - Art 6;Zusatzinformation- Inhalt 6;"
 	txtKopf   = txtkopf +"Zusatzinformation - Art 7;Zusatzinformation- Inhalt 7;Zusatzinformation - Art 8;Zusatzinformation- Inhalt 8;Zusatzinformation - Art 9;Zusatzinformation- Inhalt 9;Zusatzinformation - Art 10;"
 	txtkopf   = txtkopf +"Zusatzinformation- Inhalt 10;Zusatzinformation - Art 11;Zusatzinformation- Inhalt 11;Zusatzinformation - Art 12;Zusatzinformation- Inhalt 12;Zusatzinformation - Art 13;Zusatzinformation- Inhalt 13;"
 	txtkopf   = txtkopf + "Zusatzinformation - Art 14;Zusatzinformation- Inhalt 14;Zusatzinformation - Art 15;Zusatzinformation- Inhalt 15;Zusatzinformation - Art 16;Zusatzinformation- Inhalt 16;Zusatzinformation - Art 17;"
 	txtkopf   = txtkopf + "Zusatzinformation- Inhalt 17;Zusatzinformation - Art 18;Zusatzinformation- Inhalt 18;Zusatzinformation - Art 19;Zusatzinformation- Inhalt 19;Zusatzinformation - Art 20;"
 	txtKopf	  = txtkopf +"Zusatzinformation- Inhalt 20;Stück;Gewicht;Zahlweise;Forderungsart;Veranlagungsjahr;Zugeordnete Fälligkeit;Skontotyp;Auftragsnummer;Buchungstyp;USt-Schlüssel (Anzahlungen);EU-Mitgliedstaat (Anzahlungen);"
 	txtKopf   = txtkopf +"Sachverhalt L+L (Anzahlungen);EU-Steuersatz;Erklöskonto;Herkunft-Kz;Leerfeld;KOST-Datum;SEPA-Mandatsreferenz;Skontosperre;Gesellschaftername;Beteiligtennummer;Identifikationsnummer;Zeichennummer;Postenspeere;Bezeichunung SoBil;KZ SolBil;Festschreibung"
	FWRITE(nbASCII,txtHeader+CHR(13)+CHR(10)+TxtKopf+CHR(13)+CHR(10))
	*
	* WAIT WINDOW "chr(13)+CHR(10) gesetzt!!!!"
	SET POINT TO ","
 	* Hier werden die Buchungsdatensätze weggeschrieben
 	SCAN
 		dsatz=""
 		dsatz=IIF(datev.umsatz<0,ALLTRIM(STR(datev.umsatz*-1,12,2)),ALLTRIM(STR(datev.umsatz,12,2)))+";"
 		dsatz=dsatz+IIF(datev.umsatz<0,'"H"','"S"')+";"+'"";;;"";'+datev.konto+";"+datev.gegenkonto+";"+'"";'+SUBSTR(DTOs(datev.datum),7,2)+SUBSTR(DTOs(Datev.datum),5,2)+";"+'"'+ALLTRIM(datev.beleg1)+'";"";;"'+ALLTRIM(datev.btext)+'";;'
		dsatz=dsatz+'""; ; ; ;"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";' +'"'+ALLTRIM(datev.kostenstelle)+'"'+';"";;"";;"";;;;;;"";"";"";'       && Beginn 16-50
		FOR zs = 51 TO 115
		 dsatz=dsatz+'"";'         &&51-  && Ausgeschaltet da nicht alle Daten zwingend sind
		NEXT zs
		dsatz = dsatz +CHR(13)+CHR(10)
		FWRITE(nbASCII,dsatz)
		Dsatz=""
		* Damit ist der Buchungsatz übernommen

 	ENDSCAN
 	IF LEN(dsatz)>0
 		WAIT WINDOW "Dsatz enthält noch daten !!!"
 	endif	
	=FCLOSE(NBASCII)

 	* Hier muss noch die Debitoren übergeben werden !!!
 	pr_mandant=0
 	nh=0
 	sl=0
 	dsatz=""
 	debitor_JN=.f. && gesetzt, da wenn kein debitor, muss die letzte
 	* Blocknummer von der Buchungsdatei genommen werden, da sonst 
 	* die letzte Blocknummer=0 ist
 	
 	* Hier wird die Headerzeile für die Debitorendatei aufgebaut
 	* SV = Stapelverarbeitung
	
	txtHeader = '"EXTF";510;16;"Debitoren/Kreditoren";4;'+ DTOS(DATE())+STRTRAN(TIME(),":","")+"810;;"+'"SV";"";"";'+ALLTRIM(lcbnummer)+";"+ALLTRIM(lcmnummer)+";"+STR(YEAR(Dstart),4,0)+lcWJ+";" && lcWJ Wirtschaftjahr 0101 MMJJ
 	txtHeader = txtHeader + lcgSachkl+";;;"+'"";"";'+";;;"+'"";;;;;;;;;'&&+CHR(13)+CHR(10)
	*
	txtKopf	= "Konto;Name (Adressattyp Unternehmen);Unternehmensgegenstand;Name (Adressattyp natürl. Person);Vorname (Adressattyp natürl. Person);Name (Adressattyp keine Angabe);Adressattyp;Kurzbezeichnung;EU-Land;EU-UStID;"
	txtKopf = txtKopf + "Anrede;Titel/Akad. Grad;Adelstitel;Namensvorsatz;Adressart;Straße;Postfach;Postleitzahl;Ort;Land;Versandzusatz;Adresszusatz;Abweichende Anrede;Abw. Zustellbezeichnung 1;Abw. Zustellbezeichnung 2;"
	txtKopf = txtKopf +"Kennz. Korrespondenzadresse;Adresse Gültig von;Adresse Gültig bis;Telefon;Bemerkung (Telefon);Telefon GL;Bemerkung (Telefon GL);E-Mail;Bemerkung (E-Mail);Internet;Bemerkung (Internet);Fax;Bemerkung (Fax);"
	txtKopf = txtKopf +"Sonstige;Bemerkung (Sonstige);Bankleitzahl 1;Bankbezeichnung 1;Bank-Kontonummer 1;Länderkennzeichen 1;IBAN-Nr. 1;IBAN1 korrekt;SWIFT-Code 1;Abw. Kontoinhaber 1;Kennz. Hauptbankverb. 1;Bankverb 1 Gültig von;"
	txtKopf = txtKopf +"Bankverb 1 Gültig bis;Bankleitzahl 2;Bankbezeichnung 2;Bank-Kontonummer 2;Länderkennzeichen 2;IBAN-Nr. 2;IBAN2 korrekt;SWIFT-Code 2;Abw. Kontoinhaber 2;Kennz. Hauptbankverb. 2;Bankverb 2 Gültig von;Bankverb 2 Gültig bis;"
	txtKopf = txtKopf +"Bankleitzahl 3;Bankbezeichnung 3;Bank-Kontonummer 3;Länderkennzeichen 3;IBAN-Nr. 3;IBAN3 korrekt;SWIFT-Code 3;Abw. Kontoinhaber 3;Kennz. Hauptbankverb. 3;Bankverb 3 Gültig von;Bankverb 3 Gültig bis;Bankleitzahl 4;"
	txtKopf = txtKopf +"Bankbezeichnung 4;Bank-Kontonummer 4;Länderkennzeichen 4;IBAN-Nr. 4;IBAN4 korrekt;SWIFT-Code 4;Abw. Kontoinhaber 4;Kennz. Hauptbankverb. 4;Bankverb 4 Gültig von;Bankverb 4 Gültig bis;Bankleitzahl 5;Bankbezeichnung 5;Bank-Kontonummer 5;"
	txtKopf = txtKopf +"Länderkennzeichen 5;IBAN-Nr. 5;IBAN5 korrekt;SWIFT-Code 5;Abw. Kontoinhaber 5;Kennz. Hauptbankverb. 5;Bankverb 5 Gültig von;Bankverb 5 Gültig bis;Leerfeld;Briefanrede;Grußformel;Kundennummer;Steuernummer;Sprache;Ansprechpartner;Vertreter;"
	txtKopf = txtKopf +"Sachbearbeiter;Diverse-Konto;Ausgabeziel;Währungssteuerung;Kreditlimit (Debitor);Zahlungsbedingung;Fälligkeit in Tagen (Debitor);Skonto in Prozent (Debitor);Kreditoren-Ziel 1 Tg.;Kreditoren-Skonto 1 %;Kreditoren-Ziel 2 Tg.;"
	txtKopf = txtKopf +"Kreditoren-Skonto 2 %;Kreditoren-Ziel 3 Brutto Tg.;Kreditoren-Ziel 4 Tg.;Kreditoren-Skonto 4 %;Kreditoren-Ziel 5 Tg.;Kreditoren-Skonto 5 %;Mahnung;Kontoauszug;Mahntext 1;Mahntext 2;Mahntext 3;Kontoauszugstext;Mahnlimit Betrag;Mahnlimit %;"
	txtKopf = txtKopf +"Zinsberechnung;Mahnzinssatz 1;Mahnzinssatz 2;Mahnzinssatz 3;Lastschrift;Verfahren;Mandantenbank;Zahlungsträger;Indiv. Feld 1;Indiv. Feld 2;Indiv. Feld 3;Indiv. Feld 4;Indiv. Feld 5;Indiv. Feld 6;Indiv. Feld 7;Indiv. Feld 8;Indiv. Feld 9;"
	txtKopf = txtKopf +"Indiv. Feld 10;Indiv. Feld 11;Indiv. Feld 12;Indiv. Feld 13;Indiv. Feld 14;Indiv. Feld 15;Abweichende Anrede (Rechnungsadresse);Adressart (Rechnungsadresse);Straße (Rechnungsadresse);Postfach (Rechnungsadresse);Postleitzahl (Rechnungsadresse);"
	txtKopf = txtKopf +"Ort (Rechnungsadresse);Land (Rechnungsadresse);Versandzusatz (Rechnungsadresse);Adresszusatz (Rechnungsadresse);Abw. Zustellbezeichnung 1 (Rechnungsadresse);Abw. Zustellbezeichnung 2 (Rechnungsadresse);Adresse Gültig von (Rechnungsadresse);"
	txtKopf = txtKopf +"Adresse Gültig bis (Rechnungsadresse);Bankleitzahl 6;Bankbezeichnung 6;Bank-Kontonummer 6;Länderkennzeichen 6;IBAN-Nr. 6;IBAN6 korrekt;SWIFT-Code 6;Abw. Kontoinhaber 6;Kennz. Hauptbankverb. 6;Bankverb 6 Gültig von;Bankverb 6 Gültig bis;"
	txtKopf = txtKopf +"Bankleitzahl 7;Bankbezeichnung 7;Bank-Kontonummer 7;Länderkennzeichen 7;IBAN-Nr. 7;IBAN7 korrekt;SWIFT-Code 7;Abw. Kontoinhaber 7;Kennz. Hauptbankverb. 7;Bankverb 7 Gültig von;Bankverb 7 Gültig bis;Bankleitzahl 8;Bankbezeichnung 8;"
	txtKopf = txtKopf +"Bank-Kontonummer 8;Länderkennzeichen 8;IBAN-Nr. 8;IBAN8 korrekt;SWIFT-Code 8;Abw. Kontoinhaber 8;Kennz. Hauptbankverb. 8;Bankverb 8 Gültig von;Bankverb 8 Gültig bis;Bankleitzahl 9;Bankbezeichnung 9;Bank-Kontonummer 9;Länderkennzeichen 9;"
	txtKopf = txtKopf +"IBAN-Nr. 9;IBAN9 korrekt;SWIFT-Code 9;Abw. Kontoinhaber 9;Kennz. Hauptbankverb. 9;Bankverb 9 Gültig von;Bankverb 9 Gültig bis;Bankleitzahl 10;Bankbezeichnung 10;Bank-Kontonummer 10;Länderkennzeichen 10;IBAN-Nr. 10;IBAN10 korrekt;SWIFT-Code 10;"
	txtKopf = txtKopf +"Abw. Kontoinhaber 10;Kennz. Hauptbankverb. 10;Bankverb 10 Gültig von;Bankverb 10 Gültig bis;Nummer Fremdsystem;Insolvent;Mandatsreferenz 1;Mandatsreferenz 2;Mandatsreferenz 3;Mandatsreferenz 4;Mandatsreferenz 5;Mandatsreferenz 6;Mandatsreferenz 7;"
	txtKopf = txtKopf +"Mandatsreferenz 8;Mandatsreferenz 9;Mandatsreferenz 10;Verknüpftes OPOS-Konto;Mahnsperre bis;Lastschriftsperre bis;Zahlungssperre bis;Gebührenberechnung;Mahngebühr 1;Mahngebühr 2;Mahngebühr 3;Pauschalenberechnung;Verzugspauschale 1;"
	txtKopf = txtKopf +"Verzugspauschale 2;Verzugspauschale 3"
	
	* Hier wird die Debitorendatei / Stammdatendatei erstellt
	ASCIIDAtei = lcVerz+"EXTF_Debitoren.csv"
	nbASCII = FCREATE(ASCIIDatei)

	FWRITE(nbASCII,txtHeader+CHR(13)+CHR(10)+TxtKopf+CHR(13)+CHR(10))

 	SELECT datev
 	SET ORDER TO tag4  && mandantnr+debitornr
 	GO top
 	SCAN
 		IF VAL(datev.debitornr)>9999
 			IF pr_mandant=0 OR pr_mandant<>datev.mandantnr
 				* Achtung hier muss der Code eingebunden werden, wenn Mehrmandantenfähig
 			ENDIF
 			dsatz=""
 
 			SET CENTURY on
 			* Hier werden die Datensätze der Debitoren gespeichert
 			dsatz = ALLTRIM(datev.debitornr)+";" 						&& Feld Kontonummer
 			dsatz = dsatz +'"'+ALLTRIM(datev.firma)+'";'				&& Firma 
 			dsatz = dsatz + '"";'										&& Unternehmensgegenstand 	
 			dsatz = dsatz + '"'+IIF(EMPTY(datev.firma),ALLTRIM(datev.lName),"")+'";'				&& Name natürliche Person
 			dsatz = dsatz + '"'+IIF(EMPTY(datev.firma),ALLTRIM(datev.fName),"")+'";'				&& Vorname natürliche Person
 			dsatz = dsatz +'"";'										&& Name (Adresstyp keine Angabe) Feld wird nicht benutzt
 			dsatz = dsatz + IIF(!EMPTY(datev.firma),'"2";','"1";')		&& Adresstyp 1 = natürliche Person, 2 = Firma
 			dsatz = dsatz + '"'	+ALLTRIM(datev.suchname)+'";'			&& Suchname
 			dsatz = dsatz + '"";'										&& EU-Kennzeichen Kennzeichen aus Länderstatistik
 			dsatz = dsatz + '"";'										&& EU-UStID
 			dsatz = dsatz + IIF(!EMPTY(datev.firma),'"Firma";','"";')	&& Anrede muss hier noch eingebunden werden in Datev-Export ist nur das Kennzeichen hinterlegt
 			dsatz = dsatz + '"";'										&& Titel Grad
 			dsatz = dsatz + '"";"";'									&& Adelstitel; Namensvorsatz
 			dsatz = dsatz + '"STR";'									&& Adressart STR= Strasse ??
 			dsatz = dsatz +'"'+ALLTRIM(datev.strasse)+'";'				&& Strasse	
 			dsatz = dsatz + '"";'										&& Postfach	
 			dsatz = dsatz + '"'+ALLTRIM(datev.plz)+'";'					&& PLZ
 			dsatz = dsatz + '"'+ALLTRIM(datev.ort)+'";'					&& Ort		
 			dsatz = dsatz + '"'+ALLTRIM(datev.land)+'";'				&& Land
 			dsatz = dsatz + '"";"";"";"";"";'							&& Versandzusatz; Addresszusatz;Abweichende Anreide;Abweichende Zustellbezeichnung 1; ... 2
 			dsatz = dsatz + "1;"										&& Kennz. Korrespondenzaddresse
 			dsatz = dsatz + ";;"										&& Adresse gültig von; ... bis
 			dsatz = dsatz + '"'+ALLTRIM(datev.telefon)+'";"";'			&& Telefon 	; Bemerkung (Telefon)
 			dsatz = dsatz +'"";"";'										&& Telefon Geschäftsleitung / Bemerkung
 			dsatz = dsatz +'"'+ALLTRIM(datev.email)+'";"";'				&& Email, Bemerkung
 			dsatz = dsatz + '"";"";'									&& Internet; Bemerkung
 			dsatz = dsatz +'"'+ALLTRIM(datev.telefax)+'";"";'			&& Fax; Bemerkung
 			dsatz = dsatz + '"";"";'									&& sonstige; Bemerkung
 			dsatz = dsatz +'"";"";"";"";"";"";"";"";"0";"";;"";"";"";"";"";'	&& Bankverbindungsdaten (41-56)
 			dsatz = dsatz +'"";"";"";"0";;;"";"";"";"";"";"";"";"";"0";;;"";"";"";'	&& Bankdaten (57-76)
 			dsatz = dsatz +'"";"";"";"";"";"0";;;"";"";"";"";"";"";"";"";"0";;;"";'	&& Bankdaten (77-96)	
 			dsatz = dsatz + '"";"";"";"";;'								&& Briefanrede (97-101)
 			dsatz = dsatz + '"'+ALLTRIM(datev.na_vo)+'";"";"";'			&& Briefanrede (102- 104)
 			dsatz = dsatz +';;"2";;'									&& OPos Allgemein (105-108 )
 			dsatz = dsatz + '"";;;'										&& Zahlungsbedingungen Debitoren  (109-111)
 			dsatz = dsatz + ";;;;;;;;;"									&& Zahlungsbedingungen Kreditoren (112-120)
 			dsatz = dsatz + ";;;;;;;;;;;;"								&& Mahnwesen Debitoren (121- 132)	
 			dsatz = dsatz + '"";"";;'									&& Lastschriften Debitoren  (133-135)
 			dsatz = dsatz + '"0";'										&& Zahlungsträger Kreditoren 136
 			dsatz = dsatz + '"";"";"";"";"";"";"";"";"";"";'			&& Individuelle Felder (137-146)
 			dsatz = dsatz + '"";"";"";"";"";'							&& Individuelle Felder (147 -151)
 			dsatz = dsatz + '"";"";"";"";"";"";"";"";"";"";"";;;'		&& Adresse - Abweichend (152-164)
 			dsatz = dsatz + '"";"";"";"";"";"";"";"";"0";;;"";"";"";"";'	&& (165-179)
 			dsatz = dsatz + '"";"";"";"";"0";;;"";"";"";"";"";"";"";"";"0";;;'&& (180-197)
 			dsatz = dsatz + '"";"";"";"";"";"";"";"";"0";;;'				&& (198 - 208)
 			dsatz = dsatz + '"";"";"";"";"";"";"";"";"0";;;'				&& (209 - 219)
 			dsatz = dsatz + '"";;"";"";"";"";"";"";"";"";"";"";'		&& (220 - 231)
 			dsatz = dsatz + '"";;;;;;;;;;;'							&& (232 - 243)
 
			FWRITE(nbASCII,dsatz+chr(13)+chr(10))  && für Testzwecke ausgeschaltet!!!!
		endif
	ENDSCAN
	=FCLOSE(NBASCII)

 	* Feherprotokoll erstellen
 	* Kasse
	tdatei=alltrim(lcverz)+"Datev-ERR"+DTOS(dstart)
	COPY TO &tdatei FOR fprotokoll=.t.  TYPE xls
	* REPORT FORM report\dateverrprt HEADING "Fehlerprotokoll " PREVIEW IN SCREEN 		
 	SET NEAR &old_near
 	SELECT histres
 	SET order to &norder_hr
 	SELECT histpost
 	SET order to &norder_hp
 RETURN .t.
