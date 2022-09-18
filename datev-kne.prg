**
** datev.fxp
**
*	Letzte Änderung 21.03.05 auf 6 Rechnungsfenster erweitert. HK
**
 PARAMETER usEmemberfield
 PRIVATE deNd
 PRIVATE dsTart
 PRIVATE ncHoice
 PRIVATE cfIle
 PRIVATE nhPorder
 PRIVATE naRea
 PRIVATE cwRitebuffer
 PRIVATE nbLockcount
 PRIVATE nlAstblockposition
 PRIVATE naBrnr
 IF  .NOT. chKlicense()
      RETURN
 ENDIF
 usEmemberfield = (PARAMETERS()<>0)
 naRea = SELECT()
 cwRitebuffer = ""
 nbLockcount = 0
 nlAstblockposition = 0
 naBrnr = 0
 nhPorder = ORDER("HistPost")
 dsTart = DATE()
 deNd = DATE()
 ncHoice = 1
 = dv_window(0)
 @ 1, 2 SAY getlangteXt("DATEV","TXT_START_DATE") SIZE 20, 22 FONT 'ARIAL',10
 @ 3, 2 SAY getlangteXt("DATEV","TXT_END_DATE") SIZE 20, 22 FONT 'ARIAL',10
 @ 5, 2 SAY getlangteXt("DATEV","TXT_ACCOUNT") SIZE 20, 22 FONT 'ARIAL',10
* = txTpanel(1,2,17,getlangteXt("DATEV","TXT_START_DATE"),15,"N")
* = txTpanel(3,2,17,getlangteXt("DATEV","TXT_END_DATE"),15)
* = txTpanel(5,2,17,"Abrechnungs-Nr.",15)
 @ 1, 19 GET dsTart SIZE 1, 12 PICTURE "@K" FONT 'ARIAL',10 VALID  .NOT. EMPTY(dsTart)
 @ 3, 19 GET deNd SIZE 1, 12 PICTURE "@K" FONT 'ARIAL',10 VALID  .NOT. EMPTY(deNd)
 @ 5, 19 GET naBrnr SIZE 1, 12 PICTURE "@K 99" FONT 'ARIAL',10 VALID  .NOT. EMPTY(naBrnr)
 @ 7, 2 GET ncHoice STYLE "N" SIZE nbUttonheight, 17 FUNCTION "*"+"H"  ;
   PICTURE "\!"+getlangteXt("COMMON","TXT_OK")+";\?"+getlangteXt("COMMON","TXT_CANCEL") FONT 'ARIAL',10 
 READ CYCLE MODAL
 = dv_window(1)
 IF (ncHoice==1)
      SELECT hiStpost
      SET ORDER TO 2
      SET NEAR ON
      IF paRam.pa_version>=6.64
           SEEK dsTart
      ELSE
           SEEK DTOS(dsTart)
      ENDIF
      SET NEAR OFF
      IF (EOF() .OR. hiStpost.hp_date>deNd)
           = alErt(getlangteXt("DATEV","TXT_NO_DATA_FOUND"))
      ELSE
           IF (yeSno(getlangteXt("DATEV","TXT_ARE_YOU_SURE")))
                lsTop = .F.
                cfIlenumber = stRzero(VAL(dv_param("DatenTraegerKennsatz", ;
                              "FileNummer","001",3)),3)
                KEYBOARD "DE"+cfIlenumber
                cfIle = GETFILE("", getlangteXt("DATEV","TXT_SELECT_DIRECTORY"),  ;
                        getlangteXt("DATEV","TXT_EXPORT"))
                IF (EMPTY(cfIle))
                     lsTop = .T.
                ELSE
                     cfIlenumber = RIGHT(cfIle, 3)
                     cfIlename = SUBSTR(cfIle, RAT("\", cfIle)+1)
                     IF (LEFT(cfIlename, 2)=="DE" .AND. VAL(cfIlenumber)> ;
                        0 .AND. VAL(cfIlenumber)<999)
                          = dv_writeparam("DatenTraegerKennsatz", ;
                            "FileNummer",cfIlenumber)
                     ELSE
                          = alErt(getlangteXt("DATEV","TXT_INVALID_FILENAME"))
                          lsTop = .T.
                     ENDIF
                ENDIF
                IF ( .NOT. lsTop)
                     IF ( .NOT. dv_dexxxfile(dsTart,deNd,cfIle))
                          = alErt(getlangteXt("DATEV","TXT_NO_DE_FILE_CREATED"))
                     ELSE
                          IF ( .NOT. dv_dv01file(dsTart,deNd,cfIle))
                               = alErt(getlangteXt("DATEV", ;
                                 "TXT_NO_DV01_FILE_CREATED"))
                          ELSE
                               = alErt(getlangteXt("DATEV","TXT_EXP_FILES_CREATED"))
                          ENDIF
                     ENDIF
                ENDIF
           ENDIF
      ENDIF
 ENDIF
 SET ORDER IN hiStpost TO nHpOrder
 SELECT (naRea)
 RETURN .T.
ENDFUNC
*
FUNCTION DV_Window
 PARAMETER naCtivate
 IF (naCtivate==0)
*      DEFINE WINDOW wdAtev AT 00, 00 SIZE 8, 40 FONT "Arial", 10 NOGROW  ;
*            NOCLOSE NOZOOM TITLE chIldtitle(teXt("DATEV", ;
*            "TXT_DATEV_EXPORT")+" V"+"2.00") SYSTEM
		DEFINE WINDOW wdatev at 0,0 size 10,30 TITLE 'DATEV Export V 2.00' ;
			CLOSE FLOAT NOGROW NOZOOM ICON file "hotel.ico"
      MOVE WINDOW wdAtev CENTER
      ACTIVATE WINDOW wdAtev
      = paNelborder()
 ELSE
      DEACTIVATE WINDOW wdAtev
      RELEASE WINDOW wdAtev
      = chIldtitle("")
 ENDIF
 RETURN .T.
ENDFUNC
*
FUNCTION DV_DV01File
 PARAMETER dsTart, deNd, cdVfile
 PRIVATE nhAndle
 PRIVATE lsUccess
 cdVfile = SUBSTR(cfIle, 1, LEN(cfIle)-6)+"DV01"
 lsUccess = .F.
 nhAndle = FCREATE(cdVfile)
 IF (nhAndle==-1)
      = alErt(getlangteXt("DATEV","TXT_CREATE_ERROR"))
 ELSE
      = FWRITE(nhAndle, dv_param("DatenTraegerKennsatz", ;
        "DatenTraegerNummer","1",6))
      = FWRITE(nhAndle, dv_param("DatenTraegerKennsatz","BeraterNummer","1",7))
      = FWRITE(nhAndle, dv_param("DatenTraegerKennsatz","BeraterName", ;
        "Berater",9))
      = FWRITE(nhAndle, " ")
      = FWRITE(nhAndle, biNair(1))
      = FWRITE(nhAndle, biNair(1))
      = FWRITE(nhAndle, SPACE(37))
      = FWRITE(nhAndle, "V")
      = FWRITE(nhAndle, biNair(VAL(RIGHT(cfIle, 3))))
      = FWRITE(nhAndle, dv_param("DatenTraegerKennsatz", ;
        "AnwendungsNummer","1",2))
      = FWRITE(nhAndle, SUBSTR(g_Userid, 1, 2))
      = FWRITE(nhAndle, PADL(ALLTRIM(dv_param("DatenTraegerKennsatz", ;
        "BeraterNummer","1",7)), 7, "0"))
      = FWRITE(nhAndle, PADL(ALLTRIM(dv_param("DatenTraegerKennsatz", ;
        "MandantenNummer","1",5)), 5, "0"))
      = FWRITE(nhAndle, PADL(LTRIM(STR(M.naBrnr)), 4, "0")+ ;
        RIGHT(STR(YEAR(dsTart), 4), 2))
      = FWRITE(nhAndle, PADL(dv_date(dsTart), 10, "0"))
      = FWRITE(nhAndle, dv_date(deNd))
      = FWRITE(nhAndle, "001")
      = FWRITE(nhAndle, dv_param("DatenTraegerKennsatz","Password","1234",4))
      = FWRITE(nhAndle, biNair(nbLockcount))
      = FWRITE(nhAndle, biNair(nlAstblockposition))
      = FWRITE(nhAndle, biNair(1))
      = FWRITE(nhAndle, SPACE(8))
      = FWRITE(nhAndle, " ")
      = FWRITE(nhAndle, "1")
      = FCLOSE(nhAndle)
      lsUccess = .T.
 ENDIF
 RETURN lsUccess
ENDFUNC
*
FUNCTION DV_DExxxFile
 PARAMETER dsTart, deNd, cfIle
 PRIVATE lsUccess
 PRIVATE nhAndle
 PRIVATE ntOtalamount
 PRIVATE cpAytext, cbIllnr
 PRIVATE csKiprooms
 = dv_exportfile()
 lsUccess = .F.
 nhAndle = FCREATE(cfIle)
 IF (nhAndle==-1)
      = alErt(getlangteXt("DATEV","TXT_CREATE_ERROR"))
 ELSE
      cbUffer = CHR(29)
      cbUffer = cbUffer+CHR(24)
      cbUffer = cbUffer+"1"
      cbUffer = cbUffer+dv_param("DatenTraegerKennsatz", ;
                "DatenTraegerNummer","1",3)
      cbUffer = cbUffer+PADL(ALLTRIM(dv_param("DatenTraegerKennsatz", ;
                "AnwendungsNummer","1",2)), 2, "0")
      cbUffer = cbUffer+SUBSTR(g_Userid, 1, 2)
      cbUffer = cbUffer+PADL(ALLTRIM(dv_param("DatenTraegerKennsatz", ;
                "BeraterNummer","1",7)), 7, "0")
      cbUffer = cbUffer+PADL(ALLTRIM(dv_param("DatenTraegerKennsatz", ;
                "MandantenNummer","1",5)), 5, "0")
      cbUffer = cbUffer+PADL(LTRIM(STR(M.naBrnr)), 4, "0")+ ;
                RIGHT(STR(YEAR(dsTart), 4), 2)
      cbUffer = cbUffer+dv_date(dsTart)
      cbUffer = cbUffer+dv_date(deNd)
      cbUffer = cbUffer+"001"
      cbUffer = cbUffer+dv_param("DatenTraegerKennsatz","Password","1234",4)
      cbUffer = cbUffer+SPACE(16)
      cbUffer = cbUffer+SPACE(16)
      cbUffer = cbUffer+"y"
      cbUffer = f_Write(nhAndle,cbUffer)
      csKiprooms = ','+ALLTRIM(dv_param("Ausnahme","Zimmer","",80))+','
      cgEgenkonto = dv_param("Zahlung","Gegenkonto","7654321",7)
      ckOnto = dv_param("Artikel","Konto","12345",5)
      caMacro = "Article.Ar_Lang"+g_Langnum
      cpMacro = "Paymetho.Pm_Lang"+g_Langnum
      SELECT hiStpost
      SET ORDER TO 2
      SET NEAR ON
      IF paRam.pa_version>=6.64
           SEEK dsTart
      ELSE
           SEEK DTOS(dsTart)
      ENDIF
      SET NEAR OFF
      DO WHILE ( .NOT. EOF("HistPost") .AND. hiStpost.hp_date<=deNd)
           ddAte = hiStpost.hp_date
           WAIT WINDOW NOWAIT "1.21"+" "+DTOC(ddAte)
           DO WHILE ( .NOT. EOF("HistPost") .AND. hiStpost.hp_date==ddAte)
                SELECT hiStres
                SET ORDER TO 1
                IF paRam.pa_version >= 6.64
                     SEEK hiStpost.hp_reserid
                ELSE
                     SEEK STR(hiStpost.hp_reserid, 12, 3)
                ENDIF
                SELECT hiStpost
                IF ','+TRIM(hiStres.hr_roomnum)+','$csKiprooms
                     SKIP 1 IN hiStpost
                     LOOP
                ENDIF
                DO CASE
                     CASE ( .NOT. EMPTY(hiStpost.hp_artinum) .AND. (EMPTY(hiStpost.hp_ratecod) .OR.  ;
                          hiStpost.hp_split) .AND. hiStpost.hp_reserid>0 .AND.  .NOT. hiStpost.hp_cancel)
                          SELECT daTev
                          SET ORDER TO 1
                          IF ( .NOT. SEEK(DTOS(ddAte)+ STR(hiStpost.hp_artinum, 4), "Datev"))
                               IF paRam.pa_version>=6.64
                                    = SEEK(hiStpost.hp_artinum, "Article")
                               ELSE
                                    = SEEK(hiStpost.hp_departm+ STR(hiStpost.hp_artinum, 4), "Article")
                               ENDIF
                               SELECT daTev
                               APPEND BLANK
                               REPLACE daTev.arTikel WITH hiStpost.hp_artinum
                               REPLACE daTev.geGenkonto WITH ALLTRIM(arTicle.ar_user2)
                               REPLACE daTev.daTum WITH ddAte
                               Replace Datev.Text With &cAMacro
                               REPLACE daTev.koNto WITH ckOnto
                          ENDIF
                          REPLACE daTev.umSatz WITH daTev.umSatz+hiStpost.hp_amount
                     CASE ( .NOT. EMPTY(hiStpost.hp_paynum) .AND. hiStpost.hp_reserid>0 .AND.  .NOT. hiStpost.hp_cancel)
                          SELECT daTev
                          SET ORDER TO 2
                          SELECT paYmetho
                          LOCATE FOR paYmetho.pm_paynum=hiStpost.hp_paynum
                          IF (usEmemberfield .AND. paYmetho.pm_paytyp==4)
                               SELECT address
                               IF ( .NOT. EMPTY(hiStres.hr_compid))
                                    LOCATE FOR ad_addrid=hiStres.hr_compid
                               ELSE
                                    LOCATE FOR ad_addrid=hiStres.hr_addrid
                               ENDIF
								IF param.pa_version > 9.03
									cbillnr = IIF(histpost.hp_window=6, histres.hr_billnr6, ;
										IIF(histpost.hp_window=5, histres.hr_billnr5, IIF(histpost.hp_window=4, ;
										histres.hr_billnr4, IIF(histpost.hr_window=3, histres.hr_billnr3, ;
										IIF(histpost.hp_window=2, histres.hr_billnr2, histres.hr_billnr1)))))
								else
                               		cbillnr = IIF(hiStpost.hp_window=3, hiStres.hr_billnr3, ;
		                               IIF(hiStpost.hp_window=2, hiStres.hr_billnr2, hiStres.hr_billnr1))
		                        endif
                               IF AT('-', cbIllnr)>0
                                    cbIllnr = SUBSTR(cbIllnr, 1, 8)
                               ENDIF
                               cbEleg1 = RIGHT(cbIllnr, 6)
                               cbEleg2 = ''
**                               cpaytext = hiStres.hr_lname
                               cpAykonto = ALLTRIM(STR(adDress.ad_member))
                               IF !EMPTY(address.ad_company)
                                    cpaytext = UPPER(address.ad_company)
                               ELSE
                                    cpaytext = UPPER(address.ad_lname)
                               ENDIF
                               IF (EMPTY(VAL(cpAykonto)))
                                    cpAykonto = ALLTRIM(paYmetho.pm_user2)
                               ENDIF
                          ELSE
                               cbEleg1 = ''
                               cbEleg2 = ''
                               cPayText = &cPMacro
                               cpAykonto = ALLTRIM(paYmetho.pm_user2)
                          ENDIF
                          IF  .NOT. SEEK(DTOS(ddAte)+ STR(hiStpost.hp_paynum, 2)+cpAykonto, "Datev") ;
                          	.OR. (usEmemberfield .AND. paYmetho.pm_paytyp==4)
                               SELECT daTev
                               APPEND BLANK
                               REPLACE daTev.zaHlung WITH hiStpost.hp_paynum
                               REPLACE daTev.geGenkonto WITH cgEgenkonto
                               REPLACE daTev.beLeg1 WITH cbEleg1, daTev.beLeg2 WITH cbEleg2
                               REPLACE daTev.daTum WITH ddAte
                               REPLACE daTev.teXt WITH cpAytext
                               REPLACE daTev.koNto WITH cpAykonto
                          ENDIF
                          REPLACE daTev.umSatz WITH daTev.umSatz+ (hiStpost.hp_amount*-1)
                ENDCASE
                SKIP 1 IN hiStpost
           ENDDO
      ENDDO
      SELECT daTev
      DELETE ALL FOR umSatz=0
      SET ORDER TO 1
      GOTO TOP
      ntOtalamount = 0
      DO WHILE ( .NOT. EOF("Datev"))
           ntOtalamount = ntOtalamount+daTev.umSatz
           cbUffer = cbUffer+IIF(daTev.umSatz>=0, "+", "-")+ stRzero(ABS(daTev.umSatz)*100,10)
           cbUffer = cbUffer+"a"+PADL(ALLTRIM(daTev.geGenkonto), 7, "0")
           cbUffer = cbUffer+"b"+PADL(ALLTRIM(daTev.beLeg1), 6, "0")
           cbUffer = cbUffer+"d"+stRzero(DAY(daTev.daTum),2)+ stRzero(MONTH(daTev.daTum),2)
           cbUffer = cbUffer+"e"+PADL(ALLTRIM(daTev.koNto), 5, "0")
           cbUffer = cbUffer+CHR(30)+PADR(ANSITOOEM(daTev.teXt), 30)+CHR(28)
           cbUffer = cbUffer+"y"
           SELECT daTev
           SKIP 1
           IF (EOF())
                IF (ntOtalamount<0)
                     cbUffer = cbUffer+"w"+stRzero(ABS(ntOtalamount)*100, 12)+"yz"
                ELSE
                     cbUffer = cbUffer+"x"+stRzero(ABS(ntOtalamount)*100, 12)+"yz"
                ENDIF
           ENDIF
           cbUffer = f_Write(nhAndle,cbUffer)
      ENDDO
      IF (LEN(cwRitebuffer)>0)
           nlAstblockposition = LEN(cwRitebuffer)-1
           cbUffer = f_Write(nhAndle,"")
      ENDIF
      = FCLOSE(nhAndle)
      lsUccess = .T.
 ENDIF
 USE IN daTev
 WAIT CLEAR
 RETURN lsUccess
ENDFUNC
*
FUNCTION Binair
 PARAMETER nnUmber
 PRIVATE nhIgh1, nhIgh2
 PRIVATE nlOw1, nlOw2
 nhIgh1 = INT(nnUmber/4096)
 nhIgh2 = INT((nnUmber-(nhIgh1*4096))/256)
 nlOw1 = INT((nnUmber-(nhIgh1*4096)-(nhIgh2*256))/16)
 nlOw2 = nnUmber-(nhIgh1*4096)-(nhIgh2*256)-(nlOw1*16)
 RETURN CHR(nlOw1*16+nlOw2)+CHR(nhIgh1*16+nhIgh2)
ENDFUNC
*
FUNCTION DV_Param
 PARAMETER csEction, cpAram, cdEfault, nlEngth
 RETURN PADR(geTparam(csEction,cpAram,"DATEV.INI",cdEfault), nlEngth)
ENDFUNC
*
FUNCTION DV_Date
 PARAMETER ddAte
 RETURN stRzero(DAY(ddAte),2)+stRzero(MONTH(ddAte),2)+ SUBSTR(STR(YEAR(ddAte), 4), 3, 2)
ENDFUNC
*
FUNCTION DV_ExportFile
 CREATE CURSOR Datev (arTikel N (4, 0), zaHlung N (2, 0), umSatz N (12,  ;
        2), geGenkonto C (7), daTum D (8), koNto C (5), teXt C (30), beLeg1 C (6), beLeg2 C (6))
 SELECT daTev
 INDEX ON DTOS(daTum)+STR(arTikel, 4) TAG taG1
 INDEX ON DTOS(daTum)+STR(zaHlung, 2)+koNto TAG taG2
 SET ORDER TO 1
 RETURN .T.
ENDFUNC
*
FUNCTION DV_WriteParam
 PARAMETER csEction, cpArameter, cvAlue
 RETURN .T.
ENDFUNC
*
FUNCTION f_Write
 PARAMETER nhAndle, cdAta
 IF ( .NOT. EMPTY(cdAta) .AND. (256-(LEN(cwRitebuffer)+LEN(cdAta))>0))
      cwRitebuffer = cwRitebuffer+cdAta
 ELSE
      = FWRITE(nhAndle, PADR(cwRitebuffer, 256, CHR(0)))
      cwRitebuffer = cdAta
      nbLockcount = nbLockcount+1
 ENDIF
 RETURN ""
ENDFUNC
*
FUNCTION ChkLicense
 PRIVATE ncC, ncOunt, ctMp, nlIc, lrEt
 ncC = 0
 ctMp = PADR(paRam.pa_hotel, 30)+"OsHk991"+PADR(paRam.pa_city, 30)+ PADR("ACCT", 8)
 FOR ncOunt = 1 TO LEN(ctMp)
      ncC = ncC+(ncOunt*ASC(SUBSTR(ctMp, ncOunt, 1)))
 ENDFOR
 nlIc = VAL(dv_param("License","Code","0",10))
 IF nlIc<>ncC
      = alErt( ;
        "Ungültiger Lizenskode für Datev Schnittstelle!;;Bitte tragen Sie den Lizenskode ein in die Datei 'DATEV.INI'." ;
        )
      lrEt = .F.
 ELSE
      lrEt = .T.
 ENDIF
 RETURN lrEt
ENDFUNC
*
