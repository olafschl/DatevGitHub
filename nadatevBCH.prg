*** 
*** ReFox MMII (Win) #UK970148  SCHLINGMEIER  CITADEL [VFP50]
***
*
PROCEDURE NaDatevBCH
LPARAMETERS pcarg
DO CASE
     CASE pcarg = "VERSION"
          pcarg = "Tagesabschluß Vorlaufprogramm 1.04"
     CASE pcarg = "BEFOREAUDIT"
         WAIT WINDOW "Adressdaten werden auf korrekte Debitoren geprüft!!"
         SELECT address
         SCAN
         	skey= IIF(!EMPTY(ad_company),UPPER(SUBSTR(ad_company,1,1)),UPPER(SUBSTR(ad_lname,1,1)))
         	IF skey= "S" 
         		hf= IIF( !EMPTY(ad_company), UPPER(SUBSTR(ad_company,1,3)) ,UPPER(SUBSTR(ad_lname,1,3))) 
         		IF hf="SCH"
         			skey="SCH"
         		ENDIF
         	ENDIF
         	IF skey= "S" 
         		hf=IIF(!EMPTY(ad_company),UPPER(SUBSTR(ad_company,1,2)),UPPER(SUBSTR(ad_lname,1,2)))
         		IF hf="ST"
         			skey="ST"
         		ENDIF
         	ENDIF
         	
         	DO case
         		CASE skey="A"
         			replace ad_compnum WITH 10000
         		CASE skey="B"
         			replace ad_compnum WITH 10050
         		CASE skey="C"
         			replace ad_compnum WITH 10100
         		CASE skey="D"
         			replace ad_compnum WITH 10150
         		CASE skey="E"
         			replace ad_compnum WITH 10200
         		CASE skey="F"
         			replace ad_compnum WITH 10250
         		CASE skey="G"
         			replace ad_compnum WITH 10300
         		CASE skey="H"
         			replace ad_compnum WITH 10350
         		CASE skey="I"
         			replace ad_compnum WITH 10400
         		CASE skey="J"
         			replace ad_compnum WITH 10450
         		CASE skey="K"
         			replace ad_compnum WITH 10500
         		CASE skey="L"
         			replace ad_compnum WITH 10550
         		CASE skey="M"
         			replace ad_compnum WITH 10600
         		CASE skey="N"
         			replace ad_compnum WITH 10650
         		CASE skey="O"
         			replace ad_compnum WITH 10700
         		CASE skey="P"
         			replace ad_compnum WITH 10750
         		CASE skey="Q"
         			replace ad_compnum WITH 10800
         		CASE skey="R"
         			replace ad_compnum WITH 10850
         		CASE skey=="S"
         			replace ad_compnum WITH 10900
         		CASE skey=="SCH"
         			replace ad_compnum WITH 10950
         		CASE skey="ST"
         			replace ad_compnum WITH 11000
         		CASE skey="T"
         			replace ad_compnum WITH 11050
         		CASE skey="U"
         			replace ad_compnum WITH 11100
         		CASE skey=="V"
         			replace ad_compnum WITH 11150
         		CASE skey=="W"
         			replace ad_compnum WITH 11200
         			
         		CASE skey="X"
         			replace ad_compnum WITH 11250
         		CASE skey=="Y"
         			replace ad_compnum WITH 11300
         		CASE skey=="Z"
         			replace ad_compnum WITH 11350
         	endcase
         			
         endscan
         	
         		
     CASE pcarg = "AFTERAUDIT"
     	*
ENDCASE
ENDPROC
*
*** 
*** ReFox - retrace your steps ... 
***
