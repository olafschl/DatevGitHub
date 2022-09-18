SET ENGINEBEHAVIOR 70
SET DATE german
*SELECT histpost.hp_date, histpost.hp_artinum,sum(histpost.hp_amount), article.ar_lang3,article.ar_user2 FROM histpost;
	JOIN article ON histpost.hp_artinum=article.ar_artinum INTO CURSOR hpart GROUP BY 1,2;
	where BETWEEN(hp_date,CTOD("01.01.2019"),CTOD("31.01.2019"))
*
lcNoPayNum="99,98"
*verz = ALLTRIM(substr(SYS(16,0), AT(" ", SYS(16,0),2), LEN(SYS(16,0))))
verz = SYS(16,0)

*backupVer = SYS(16,0)+"BackupFibuExport\"
Backupver = SUBSTR(verz,1,RAT("\",Verz))+ "BackupFibu-"  &&+DTOS(Dstart)+"-" + dtos(dende)
*
WAIT WINDOW "Verz= "+verz
*
*SELECT histpost.hp_date as BelegDatum, histpost.hp_postid as BelegNR,histpost.hp_artinum,sum(histpost.hp_amount) as betrag, histpost.hp_billnum as OPNumer,;
	 article.ar_lang3 as Buchtext, Article.ar_user1 as KSt, ALLTRIM(article.ar_user2) AS Konto, article.ar_artityp,;
	 LOOKUP(picklist.pl_numval,article.ar_vat,picklist.pl_numcod) as UstSatz  FROM histpost;
	JOIN article ON histpost.hp_artinum=article.ar_artinum INTO CURSOR hbuch GROUP BY 3;
	where BETWEEN(hp_date,CTOD("01.01.2019"),CTOD("31.01.2019")) AND (hp_split = .t. OR hp_split=.f. AND ;
	EMPTY(hp_ratecod)) AND hp_cancel =.f. AND ar_artityp<>3
*
SELECT bn_date, bn_billnum, bn_reserid, bn_addrid, hp_billnum, hp_artinum FROM billnum JOIN histpost ON billnum.bn_billnum = histpost.hp_billnum INTO CURSOR bn;
	where BETWEEN(bn_date,CTOD("01.01.2019"),CTOD("31.01.2019")) AND (hp_split = .t. OR hp_split=.f. AND ;
	EMPTY(hp_ratecod)) AND hp_cancel =.f. &&AND ar_artityp<>3
BROWSE
*
SELECT bn_date, bn_billnum, bn_reserid, hp_billnum, ar_artinum, ar_artityp,ar_lang3,ALLTRIM(ar_user1) as KST, ALLTRIM(ar_user2) as Konto;
 FROM bn JOIN article ON bn.hp_artinum=article.ar_artinum INTO CURSOR hbuch;
	WHERE ar_artityp<>3
brow
*
i=2
*dsatz = "1;" + DTOC(Belegdatum) + ";" + ALLTRIM(STR(BelegNr,10,0)) + ";" + ALLTRIM(OPNumer) + ";" + ALLTRIM(BuchText) + ";" + ";"
*dsatz = dsatz + IIF(i=1,ALLTRIM(lcVArtikel),ALLTRIM(konto))+";"+IIF(i=1,ALLTRIM(konto),ALLTRIM(lcVZahlung))+";"+ALLTRIM(STR(Betrag,12,2))+";"
*dsatz = dsatz + ALLTRIM(STR(Ustsatz,2,0))+";"+CHR(13)+CHR(10)
SELECT histpost.hp_date as BelegDatum, histpost.hp_postid as BelegNr, histpost.hp_paynum as paynum, SUM(histpost.hp_amount*-1) as Betrag, histpost.hp_billnum as OPNummer,;
	paymetho.pm_lang3 as BuchText, paymetho.pm_user1 as KSt, paymetho.pm_user2 as Konto, 0 as UstSatz, ad_company, bn_billnum, bn_addrid, ad_compnum FROM histpost;
	JOIN billnum ON histpost.hp_billnum = billnum.bn_billnum join address ON address.ad_addrid = billnum.bn_addrid ;
	JOIN paymetho ON histpost.hp_paynum = paymetho.pm_paynum INTO cursor hbuch GROUP BY 2;
	where BETWEEN(hp_date,CTOD("01.01.2019"),CTOD("31.01.2019")) AND hp_reserid<>-2 AND hp_cancel=.f. AND IIF(EMPTY(lcNoPayNum)=.t.,.t.,!INLIST(histpost.hp_paynum,&lcNoPayNum))
BROWSE
*
