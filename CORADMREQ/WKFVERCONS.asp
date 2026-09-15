<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function abrewin(xmodel,xtipo,xcktpsc) { "
  sHTML = sHTML  & " Datos= 'WKFVERCONS.ASP?MODEL=' + xmodel + '&tipo=' + xtipo + '&cktpsc=' + xcktpsc; "
 ' sHTML = sHTML  & " alert(Datos); "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=400,width=400" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 MODEL   = request.querystring("MODEL")
 tipo    = request.querystring("tipo")
 cktpsc  = request.querystring("opcion")


    sInput = "PROPIEDADES DEL PASO"
    celdas = ofv.GenCelda("","tt COLSPAN=3","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")



 

  StrSql = "SELECT * FROM WKFPASOS  WHERE WKFCOD = '" & session("WKF") & "' AND  WKFPASCOD = '" & MODEL & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS.EOF THEN


    sInput = "DESCRIPCION:"
    celdas = ofv.GenCelda("","vt","","20%","","","",sInput,"si")

    sInput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","","","2%","","","",sInput,"si")

    sInput = RS(2)
    celdas = celdas & ofv.GenCelda("","vt","","70%","","","",sInput,"si")

    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")










 CKT = request.querystring("CKT")
 WKF = request.querystring("WKF")

    CKTOPC = "NO"
 if CKT = "" then
    CKT = session("CKT")
 else
    session("CKT") = CKT
 end if 

  
 if CKT <> "" then
    CKTOPC = "SI"
  Strsql = "SELECT * FROM SEGCIRCUITO WHERE ID = " & CKT
  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rsx.eof then
     WKF         = RSX(1)
     WKFOBJTIPO  = RSX(4)
     WKFOBJCODE  = RSX(5)
     WKFCKTCODE  = RSX(6)
  END IF
 end if 

 if WKF = "" then
    WKF = session("WKF")
 else
    session("WKF") = WKF
 end if 

  WKFAUTO = "WKF0001693621131313343163"
  WKFAPRO = "WKF0001693673726313343163"
  DIM ATABLA(50,2)

   IF   CKTOPC = "SI" THEN
        IF WKFOBJTIPO = "REQ" THEN
  
                 StrSql = "Select * From REQSECTORES where codigo = '" & WKFOBJCODE  & "' AND TIPO = 'OR' "
                 set rsX0 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  
                 StrSql = "Select * From REQSECTORES where codigo = '" & WKFOBJCODE  & "' AND TIPO = 'AN' "
                 set rsX1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  
                 StrSql = "Select * From REQUERIMIENTO where codigo = '" & WKFOBJCODE  & "'  "
                 set rsX2 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)

                 CKTDESCRI = RSX2(2)
                 CKTTIPO   = "REQUERIMIENTO"
  
  		if WKF = WKFAUTO THEN
  
			WSORIG = RSX0(3)
			WUORIG = RSX2(13)
			WSDEST = RSX1(3)
			WUDEST = "NO"
  
  		END IF
  		IF WKF = WKFAPRO THEN
  
			WSORIG = RSX1(3)
			WUORIG = RSX2(14)
			WSDEST = RSX0(3)
			WUDEST = RSX2(13)
  
  		END IF
  		
  		IF WUORIG <> "NO" THEN
  		   CKTROTIPO = "USR"
  		   CKTROCODE = WUORIG
  		ELSE
  		   CKTROTIPO = "SEC"
  		   CKTROCODE = WSORIG
  		END IF
  
  		IF WUDEST <> "NO" THEN
  		   CKTRDTIPO = "USR"
  		   CKTRDCODE = WUDEST
  		ELSE
  		   CKTRDTIPO = "SEC"
  		   CKTRDCODE = WSDEST
  		END IF
  
  
  
        END IF 
  
  
   END IF


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  
  

  
  
     WKFCLASE = RS(5)


   IF   CKTOPC = "SI" THEN


    sInput = CKTTIPO
    celdas = ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    sInput = "<B>" & CKTDESCRI & " - (" & WKFOBJCODE & ") </B>"
    celdas = celdas & ofv.GenCelda("","cc","","80%","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","vt","","20%","","","",sInput,"si")

    sInput = RS(2)
    celdas = celdas & ofv.GenCelda("","vt","","80%","","","",sInput,"si")

    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")


   ELSE

    sInput = "CIRCUITO:"
    celdas = ofv.GenCelda("","vt","","10%","","","",sInput,"si")

    sInput = "<B>" & RS(1) & " - " & RS(2) & "</B>"
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

   END IF


  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF


     INIPAS = "NO"
     INISEC = "NO"
     FINPAS = "NO"
     FINSEC = "NO"
     PASOK  = "NO"
     PASRCH = "NO"

     DIM TSEC(100,2)
     ISEC  = 0
     ISEC2 = 0 
     TSC   = 0


     DIM TPASOK(200,18)
     DIM TPASORD(200,4)
     IPO  = 0
     IPR  = 0
     TPO  = 0
     TPR  = 0
     IPO2 = 0




'''''''''' CARGA PASOS '''''''''''''''''''''''' 

  StrSql = "SELECT * FROM WKFPASOS  WHERE WKFCOD = '" & session("WKF") & "' ORDER BY INI DESC, APR DESC, RCH DESC, FIN DESC "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS.EOF THEN
    DO WHILE NOT RS.EOF 


'''     REDIM PRESERVE TPASOK(IPO,8)

        WKFPASTIPO = RS(12)
 '''          WKFCLASE      <<  CLASE DEL CIRCUITO :  E - ASIGNACION EFECTIVA   M - MODELO AUTOMATIZADO
 '''          WKFPASTIPO    <<  TIPO DE PASO       :  U - NORMAL  R - RECURRENTE
 
     WXXSEN = "S"
    ''' WKFCKTCODE               ' CODIGO DEL CIRCUITO DE WORKFLOW CORRIENTE (SEGCKT)
    ''' WKF                      ' CODIGO DEL MODELO DE WORKFLOW CORRIENTE (WKFCOD)
    ''' CKTROTIPO                ' TIPO RESPONSABLE ORIGINANTE 
    ''' CKTROCODE                ' CODIGO RESPONSABLE ORIGINANTE 
    ''' CKTRDTIPO                ' TIPO RESPONSABLE DESTINATARIO
    ''' CKTRDCODE                ' CODIGO RESPONSABLE DESTINATARIO
  
 
     IF WKFPASTIPO = "R" THEN
     

           IF  CKTOPC = "SI" THEN
           
             ITAB = 0 
             itabs = 0
             StrSql = "Select * From SEGCKTRESP WHERE SEGCKTPSC = '" & RS(2) & "' AND SEGCKT = '" & WKFCKTCODE  & "' order by ORDEN "
  '''                 response.write StrSql
             set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rs3.eof then
                DO UNTIL RS3.EOF
                   ITAB = ITAB + 1
         '''          response.write itab 
                   ATABLA(ITAB,0) = RS3(4) 
                   ATABLA(ITAB,1) = RS3(5) 
                   '''if clng(sniv) = clng(rs3(5)) then
                   '''   itabs = itab
                   '''end if
                   ATABLA(ITAB,2) = RS3(6)
                rs3.movenext 
                LOOP    
             END IF            
           
           
                     IPASRC = ""
	        
	     FOR IPASR=1 TO ITAB   
	               WXXSEN = "N" 
	               IPO  = IPO + 1
	   	       TPO  = TPO + 1

		TPASOK(IPO,16) = 0	

	        PASOK = RS(10) & ""
	        PASRCH = RS(11) & ""   	  

	     	     IF NOT LEN(PASOK) > 2 THEN
	     		PASOK = "NO"
	     	     END IF
	     
	     	     IF NOT LEN(PASRCH) > 2 THEN
	     		PASRCH = "NO" 
	     	     END IF



     
		StrSql = "Select * From SEGPASO WHERE SEGCKT = '" & WKFCKTCODE & "' AND MODEL = '" & RS(2) & "' "
		StrSql = StrSql & " AND RESPONSABLE =  '" & ATABLA(IPASR,1) & "' "

		set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		IF NOT RS1.EOF THEN
		   WXXSEN = "S"
		   WXXTRESP       = RS1(7)
		   WXXCRESP       = RS1(8)
		   WXXURESP       = RS1(9)
		   TPASOK(IPO,16) = RS1(0)		
		   TPASOK(IPO,17) = RS1(5)	
		   TPASOK(IPO,18) = RS1(2)	

                '   if rs1(6) <> "" then

		      StrSql = "Select MODEL From SEGPASO WHERE SEGCKT = '" & WKFCKTCODE & "' AND SEGCKTPSC = '" & rs1(6) & "' "

		      set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		      IF NOT RS1.EOF THEN
 
	     	         IF PASOK  = RS1(0) or RS(2) = RS1(0) THEN
	     		    PASRCH = "NO" 
                         ELSEIF PASRCH = RS1(0)  THEN
	     		    PASOK  = "NO"
	     	         END IF

                      end if

                 '  end if


                ELSE
		   WXXSEN = "N"
		END IF
	   	       

           
		     TPASOK(IPO,1) = RS(2) & IPASRC         '1 cod paso

		     TPASOK(IPO,0) = RS(3) & "." & IPASR    '0 ORDEN 

		     TPASOK(IPO,2) = ATABLA(IPASR,1)         '2 RESPONSABLE  - LISTA DE RESPONSABLES AUTOMATICA
	     

	     

	     	     
	     	     TPASOK(IPO,3) = PASOK                  '3 paso siguiente ok
	     	     IPASRC = IPASR
	     	     
	     	     IF PASOK <> "NO" THEN
	     	        IF IPASR = ITAB  THEN
	                        TPASOK(IPO,3) = PASOK           
	                     ELSE
	                        TPASOK(IPO,3) = RS(2) & IPASRC                     
	                 END IF
	             END IF
	     	     
	     	     
	     	     TPASOK(IPO,4)  = PASRCH                 '4 paso siguiente rechazo 
	     	     TPASOK(IPO,5)  = RS(3) & "." & IPASR    '5 orden real
	     	     TPASOK(IPO,6)  = 0                      '6 check 
	     	     TPASOK(IPO,7)  = 0                      '7 llamados 
	     	     TPASOK(IPO,8)  = RS(4)                  '8 descripcion 
	     	     TPASOK(IPO,9)  = "N;0"                  '9 ARRAY DE LLAMADOS
	     	     TPASOK(IPO,10) = 0                     '10 INDEX PASOK
	     	     TPASOK(IPO,11) = 0                     '11 INDEX PASRCH
	     	     TPASOK(IPO,12) = 0                     '12 INDEX PASO ORD
	     	     TPASOK(IPO,13) = 0                     '13 LLAMADOS REALIZADOS    
	     	     
	     	     
	     	     TPASOK(IPO,14) = WXXSEN                '14 ACTIVADA  
	     	     TPASOK(IPO,15) = ATABLA(IPASR,0)        '15 TIPO RESPONSABLE
	     	     
		   NEXT

	     
	     
           
           ELSE
           
             IPASRC = ""
	     WXXSEN = "S"  
	     
	     FOR IPASR=1 TO 3    
	             
	               IPO  = IPO + 1
	   	       TPO  = TPO + 1

		     TPASOK(IPO,16) = RS(0)		    '16 cod id del wkfpsc 
		     TPASOK(IPO,1) = RS(2) & IPASRC         '1 cod paso

		     TPASOK(IPO,0) = RS(3) & "." & IPASR    '0 ORDEN 

		     TPASOK(IPO,2) = "L"  & IPASR           '2 RESPONSABLE  - LISTA DE RESPONSABLES AUTOMATICA
	     
	     	     PASOK = RS(10) & ""
	     	     PASRCH = RS(11) & ""
	     
	     	     IF NOT LEN(PASOK) > 2 THEN
	     		PASOK = "NO"
	     	     END IF
	     
	     	     IF NOT LEN(PASRCH) > 2 THEN
	     		PASRCH = "NO" 
	     	     END IF
	     	     
	     	     TPASOK(IPO,3) = PASOK                  '3 paso siguiente ok
	     	     IPASRC = IPASR
	     	     
	     	     IF PASOK <> "NO" THEN
	     	        IF IPASR = 3 THEN
	                        TPASOK(IPO,3) = PASOK           
	                     ELSE
	                        TPASOK(IPO,3) = RS(2) & IPASRC                     
	                     END IF
	                  END IF
	     	     
	     	     
	     	     TPASOK(IPO,4) = PASRCH                 '4 paso siguiente rechazo 
	     	     TPASOK(IPO,5) = RS(3) & "." & IPASR    '5 orden real
	     	     TPASOK(IPO,6) = 0                      '6 check 
	     	     TPASOK(IPO,7) = 0                      '7 llamados 
	     	     TPASOK(IPO,8) = RS(4)                  '8 descripcion 
	     	     TPASOK(IPO,9) = "N;0"                  '9 ARRAY DE LLAMADOS
	     	     TPASOK(IPO,10) = 0                     '10 INDEX PASOK
	     	     TPASOK(IPO,11) = 0                     '11 INDEX PASRCH
	     	     TPASOK(IPO,12) = 0                     '12 INDEX PASO ORD
	     	     TPASOK(IPO,13) = 0                     '13 LLAMADOS REALIZADOS     
	     	     TPASOK(IPO,14) = WXXSEN                '14 ACTIVADA  
	     	     TPASOK(IPO,15) = "MOD"                 '15 TIPO RESPONSABLE
	     	     
		   NEXT



           END IF 


	     

	     
	     
	     

     
     
     ELSE
     
	IF  CKTOPC = "SI" THEN

          IPO  = IPO + 1
          TPO  = TPO + 1
          WXXSEN = "N"  

	        PASOK = RS(10) & ""
	        PASRCH = RS(11) & ""   	  

	     	     IF NOT LEN(PASOK) > 2 THEN
	     		PASOK = "NO"
	     	     END IF
	     
	     	     IF NOT LEN(PASRCH) > 2 THEN
	     		PASRCH = "NO" 
	     	     END IF


		TPASOK(IPO,16) = 0		        '16 no hay correspondencia
		StrSql = "Select * From SEGPASO WHERE SEGCKT = '" & WKFCKTCODE & "' AND MODEL = '" & RS(2) & "' "
		set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		IF NOT RS1.EOF THEN
		   WXXSEN = "S"
		   WXXTRESP       = RS1(7)
		   WXXCRESP       = RS1(8)
		   WXXURESP       = RS1(9)
		   TPASOK(IPO,16) = RS1(0)	      '16 cod id del segcktpsc 
		   TPASOK(IPO,17) = RS1(5)	
		   TPASOK(IPO,18) = RS1(2)	
                 '  if rs1(6) <> "" then

		      StrSql = "Select MODEL From SEGPASO WHERE SEGCKT = '" & WKFCKTCODE & "' AND SEGCKTPSC = '" & rs1(6) & "' "
		      set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		      IF NOT RS1.EOF THEN
	     	         IF PASOK  = RS1(0) THEN
	     		    PASRCH = "NO" 
                         ELSEIF PASRCH = RS1(0)  THEN
	     		    PASOK  = "NO"
	     	         END IF

                      end if

                  ' end if
                ELSE
		   WXXSEN = "N"
		   
		END IF



	     IF WKFCLASE = "E" THEN
	     
	     	     IF RS(6) THEN
	     		TPASOK(IPO,0) = 0             '0 orden 
	     	     ELSEIF RS(9) THEN
	     		TPASOK(IPO,0) = 9999
	     	     ELSE
	     		TPASOK(IPO,0) = RS(3)
	     	     END IF
	     
	     	     TPASOK(IPO,1) = RS(2)            '1 cod paso
	     	     TPASOK(IPO,2) = "NO"             '2 responsable


		StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & RS(2) & "' "
		set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		IF NOT RS1.EOF THEN
		   TPASOK(IPO,2) = RS1(3)
		END IF
		
		TPASOK(IPO,15) = "MOD"               '15 TIPO RESPONSABLE

	     ELSE
	     
	     	     IF RS(6) THEN
	     		TPASOK(IPO,0) = 0             '0 orden 
	     	     ELSEIF RS(9) THEN
	     		TPASOK(IPO,0) = 9999
	     	     ELSE
	     		TPASOK(IPO,0) = RS(3)
	     	     END IF
	     
	        TPASOK(IPO,1) = RS(2)            '1 cod paso

    ''' CKTROTIPO                ' TIPO RESPONSABLE ORIGINANTE 
    ''' CKTROCODE                ' CODIGO RESPONSABLE ORIGINANTE 
    ''' CKTRDTIPO                ' TIPO RESPONSABLE DESTINATARIO
    ''' CKTRDCODE                ' CODIGO RESPONSABLE DESTINATARIO



		IF RS(6) THEN
		   TPASOK(IPO,2)  = CKTROCODE      'ORIGINANTE DEL CIRCUITO
		   TPASOK(IPO,15) = CKTROTIPO      '15 TIPO RESPONSABLE
		ELSEIF RS(9) THEN
		   TPASOK(IPO,2)  = CKTRDCODE      'ORIGINANTE DEL CIRCUITO
		   TPASOK(IPO,15) = CKTRDTIPO      '15 TIPO RESPONSABLE
                   IF WXXSEN = "S" AND WXXURESP <> "NO" THEN
                      TPASOK(IPO,2)  = WXXURESP    'RESPONSABLE A ASIGNAR
                      TPASOK(IPO,15) = "USR"       '15 TIPO RESPONSABLE
                   END IF
		   
		ELSE
                   TPASOK(IPO,2)  = CKTROCODE      'RESPONSABLE A ASIGNAR
                   TPASOK(IPO,15) = CKTROTIPO      '15 TIPO RESPONSABLE
                   IF WXXSEN = "S" AND WXXURESP <> "NO" THEN
                      TPASOK(IPO,2)  = WXXURESP    'RESPONSABLE A ASIGNAR
                      TPASOK(IPO,15) = "USR"       '15 TIPO RESPONSABLE
                   END IF
		END IF
		




	     END IF




	     TPASOK(IPO,3) = PASOK           '3 paso siguiente ok
	     TPASOK(IPO,4) = PASRCH          '4 paso siguiente rechazo 
	     TPASOK(IPO,5) = RS(3)           '5 orden real
	     TPASOK(IPO,6) = 0               '6 check 
	     TPASOK(IPO,7) = 0               '7 llamados 
	     TPASOK(IPO,8) = RS(4)           '8 descripcion 
	     TPASOK(IPO,9) = "N;0"           '9 ARRAY DE LLAMADOS
	     TPASOK(IPO,10) = 0              '10 INDEX PASOK
	     TPASOK(IPO,11) = 0              '11 INDEX PASRCH
	     TPASOK(IPO,12) = 0              '12 INDEX PASO ORD
	     TPASOK(IPO,13) = 0              '13 LLAMADOS REALIZADOS
	     
	     TPASOK(IPO,14) = WXXSEN           '14 ACTIVADA  

	
	
	ELSE
     
     
          IPO  = IPO + 1
          TPO  = TPO + 1
          WXXSEN = "S"  


	     IF RS(6) THEN
		TPASOK(IPO,0) = 0             '0 orden 
	     ELSEIF RS(9) THEN
		TPASOK(IPO,0) = 9999
	     ELSE
		TPASOK(IPO,0) = RS(3)
	     END IF

	     TPASOK(IPO,16) = RS(0)	      '16 cod id del wkfpsc 
	     TPASOK(IPO,1) = RS(2)            '1 cod paso
	     TPASOK(IPO,2) = "NO"             '2 responsable


	     IF WKFCLASE = "E" THEN

		StrSql = "Select * From WKFPSCRESP WHERE WKFPASCOD = '" & RS(2) & "' "
		set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		IF NOT RS1.EOF THEN
		   TPASOK(IPO,2) = RS1(3)
		END IF

	     ELSE

		IF RS(6) THEN
		   TPASOK(IPO,2) = "O"             'ORIGINANTE DEL CIRCUITO
		ELSEIF RS(9) THEN
		   TPASOK(IPO,2) = "O"             'ORIGINANTE DEL CIRCUITO
		ELSE
                   TPASOK(IPO,2) = "A"          'RESPONSABLE A ASIGNAR
		END IF


	     END IF


	     PASOK = RS(10) & ""
	     PASRCH = RS(11) & ""

	     IF NOT LEN(PASOK) > 2 THEN
		PASOK = "NO"
	     END IF

	     IF NOT LEN(PASRCH) > 2 THEN
		PASRCH = "NO" 
	     END IF

	     TPASOK(IPO,3) = PASOK           '3 paso siguiente ok
	     TPASOK(IPO,4) = PASRCH          '4 paso siguiente rechazo 
	     TPASOK(IPO,5) = RS(3)           '5 orden real
	     TPASOK(IPO,6) = 0               '6 check 
	     TPASOK(IPO,7) = 0               '7 llamados 
	     TPASOK(IPO,8) = RS(4)           '8 descripcion 
	     TPASOK(IPO,9) = "N;0"           '9 ARRAY DE LLAMADOS
	     TPASOK(IPO,10) = 0              '10 INDEX PASOK
	     TPASOK(IPO,11) = 0              '11 INDEX PASRCH
	     TPASOK(IPO,12) = 0              '12 INDEX PASO ORD
	     TPASOK(IPO,13) = 0              '13 LLAMADOS REALIZADOS
	     TPASOK(IPO,14) = WXXSEN           '14 ACTIVADA  
	     TPASOK(IPO,15) = "MOD"           '15 TIPO RESPONSABLE
	     
      END IF
     
     END IF


    RS.MOVENEXT
    LOOP
  END IF

''''''''''' FIN CARGA PASOS '''''''''''''''''''''''


''''''''''' ORDENA PASOS '''''''''''''''''''''''

IF TPO > 0 THEN 

   FOR IPO2=1 TO TPO
       IF TPASOK(IPO2,0) = 0 THEN
          IPO = IPO2
          IPO2 = 9999
          EXIT FOR
       END IF   
   NEXT 

   IF IPO2 <> 9999 THEN
      IPO = 1
   END IF





   DO 
     IF TPO = TPR THEN
        EXIT DO 

     END IF

     IPR = IPR + 1
     TPR = TPR + 1
'''     REDIM PRESERVE TPASORD(IPR,4)
     TPASORD(IPR,0) = IPO
     TPASOK(IPO,12) = IPR  

     NEXTPAS = "NO"
     IF TPASOK(IPO,3) <> "NO" THEN
        NEXTPAS = TPASOK(IPO,3)
        TPASOK(IPO,6) = 1                     '1 CHECK POR CAMINO OK
     ELSEIF TPASOK(IPO,4) <> "NO" THEN
        NEXTPAS = TPASOK(IPO,4)
        TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
     ELSE
        TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
        EXIT DO
     END IF  

     FOR IPO2=1 TO TPO
       IF TPASOK(IPO2,1) = NEXTPAS THEN
          IPO = IPO2
          IPO2 = 9999
          EXIT FOR
       END IF   

     NEXT
     
     IF IPO2 <> 9999 THEN
              IF TPASOK(IPO,6) = 1 THEN
                 TPASOK(IPO,3) = "NO"
              ELSE
                 TPASOK(IPO,4) = "NO"
              END IF 
              EXIT DO

     END IF



   LOOP 

   IF TPR > 0 THEN

      IF TPR < TPO THEN
         IPO3 = 1
         DO

           FOR IPO2=IPO3 TO TPO
               IF TPASOK(IPO2,6) = 1 THEN
                  TPASOK(IPO2,6) = 2
                  IF TPASOK(IPO2,4) <> "NO" THEN
                     IPO3 = IPO2
                     NEXTPAS = TPASOK(IPO2,4)
                     IPO = IPO2
                     IPO2 = 9999
                     EXIT FOR
                  END IF  


               END IF
           NEXT

           IF IPO2 <> 9999 THEN
              EXIT DO
           END IF

         DO 


           FOR IPO2=1 TO TPO
             IF TPASOK(IPO2,1) = NEXTPAS THEN
                IPO = IPO2
                IPO2 = 9999
                EXIT FOR
             END IF   

           NEXT

           IF IPO2 <> 9999 THEN
              IF TPASOK(IPO,6) = 1 THEN
                 TPASOK(IPO,3) = "NO"
              ELSE
                 TPASOK(IPO,4) = "NO"
              END IF 
              EXIT DO
           END IF

           IF TPASOK(IPO,6) = 0 THEN
              IPR = IPR + 1
              TPR = TPR + 1
'''              REDIM PRESERVE TPASORD(IPR,4)
              TPASORD(IPR,0) = IPO
              TPASOK(IPO,12) = IPR  
           END IF


           NEXTPAS = "NO"

           IF TPASOK(IPO,6) = 0 THEN
              
              IF TPASOK(IPO,3) <> "NO" THEN
                 NEXTPAS = TPASOK(IPO,3)
                 TPASOK(IPO,6) = 1                     '1 CHECK POR CAMINO OK
              ELSEIF TPASOK(IPO,4) <> "NO" THEN
                 NEXTPAS = TPASOK(IPO,4)
                 TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
              ELSE
                 TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
                 EXIT DO
              END IF  

           ELSEIF TPASOK(IPO,6) = 1 THEN

              IF TPASOK(IPO,4) <> "NO" THEN
                 NEXTPAS = TPASOK(IPO,4)
                 TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
              ELSE
                 TPASOK(IPO,6) = 2                     '2 CHECK POR DOS CAMINOS
                 EXIT DO
              END IF  

           ELSE

              EXIT DO
           END IF





         LOOP 


         LOOP


      END IF


      FOR IPO3=1 TO TPR

          IPO2 = TPASORD(IPO3,0) 
'''          RESPONSE.WRITE IPO3 & "-" & IPO2 & "/"
          IF TPASOK(IPO2,3) <> "NO" THEN
             FOR IPO=1 TO TPO
                 IF TPASOK(IPO,1) = TPASOK(IPO2,3) THEN
                    TPASOK(IPO2,10) = TPASOK(IPO,12)
                    IF TPASOK(IPO,7) = 0 THEN TPASOK(IPO,9) = "S"
                    TPASOK(IPO,7) = TPASOK(IPO,7) + 1
                    TPASOK(IPO,9) = TPASOK(IPO,9) & ";" & IPO3
                    EXIT FOR
                 END IF

             NEXT          


          END IF 

          IF TPASOK(IPO2,4) <> "NO" THEN
             FOR IPO=1 TO TPO
                 IF TPASOK(IPO,1) = TPASOK(IPO2,4) THEN
                    TPASOK(IPO2,11) = TPASOK(IPO,12)
                    IF TPASOK(IPO,7) = 0 THEN TPASOK(IPO,9) = "S"
                    TPASOK(IPO,7) = TPASOK(IPO,7) + 1
                    TPASOK(IPO,9) = TPASOK(IPO,9) & ";" & IPO3
                    EXIT FOR
                 END IF

             NEXT          

          END IF



      NEXT



      ISEC   = 1 
      TSC    = 1 
      IPO    = TPASORD(1,0)
      TSEC(ISEC,1) = TPASOK(IPO,2)
      TSEC(ISEC,2) = GETRESPO(TPASOK(IPO,2),TPASOK(IPO,15))
      TPASOK(IPO,2) = ISEC

      FOR IPO2=2 TO TPR
          IPO    = TPASORD(IPO2,0)
          FOR ISEC2=1 TO TSC
              IF TSEC(ISEC2,1) = TPASOK(IPO,2) THEN
                 TPASOK(IPO,2) = ISEC2
                 ISEC2 = 9999
                 EXIT FOR
              END IF 

          NEXT

          IF ISEC2 <> 9999 THEN
             ISEC   = ISEC + 1 
             TSC    = TSC + 1 
             TSEC(ISEC,1) = TPASOK(IPO,2)
             TSEC(ISEC,2) = GETRESPO(TPASOK(IPO,2),TPASOK(IPO,15))
             TPASOK(IPO,2) = ISEC
          END IF


      NEXT




      FILAS = ""

      sInput = "ORD"
      celdas = ofv.GenCelda("","tt","","","","","",sInput,"si")

      sInput = "PASO"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "CHECK"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "RESP"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "OK"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "RCH"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "LLAMADOS"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      sInput = "ARRAY"
      celdas = celdas & ofv.GenCelda("","TT","","","","","",sInput,"si")

      filas = ofv.GenRow("","","","","","","",celdas,"si")

      TOTFILAS = 0
      TOTCOLUMNAS = TSC * 8

      FOR IPO2=1 TO TPR  


          sInput = IPO2
          celdas = ofv.GenCelda("","cc","","","","","",sInput,"si")

          IPO    = TPASORD(IPO2,0)

          IF TPASOK(IPO,7) > 1 THEN 
             TFILAS = TPASOK(IPO,7) + 5
          ELSE
             TFILAS = 6
          END IF  

          TOTFILAS = TOTFILAS + TFILAS

          TPASORD(IPO2,1) = TFILAS
          TPASORD(IPO2,2) = TPASOK(IPO,9)

'''''     TPASOK(IPO,1) = RS(2)           '1 PASO
'''''     TPASOK(IPO,2) = SECT            '2 RESPONSABLE
'''''     TPASOK(IPO,3) = PASOK           '3 paso siguiente ok
'''''     TPASOK(IPO,4) = PASRCH          '4 paso siguiente rechazo 
'''''     TPASOK(IPO,5) = RS(3)           '5 orden real
'''''     TPASOK(IPO,6) = 0               '6 check 
'''''     TPASOK(IPO,7) = 0               '7 llamados 
'''''     TPASOK(IPO,8) = RS(4)           '8 descripcion 
'''''     TPASOK(IPO,9) = ""              '9 ARRAY DE LLAMADOS

          sInput = TPASOK(IPO,5) & ":" & TPASOK(IPO,1) & "-" & TPASOK(IPO,8)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,6)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,2) & ":" & TSEC(TPASOK(IPO,2),2)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,3)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,4)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,7)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          sInput = TPASOK(IPO,9)
          celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

          filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")



      NEXT


 '''  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"

 PUBLIC XMATRIX(500,500,8)

''' REDIM PRESERVE XMATRIX(TOTFILAS,TOTCOLUMNAS,4)
 IMF   = 0
 IMC   = 0
 ISF   = 0
 ISC   = 0
 BFILA = 0
 BCOL  = 0
 ICOL  = 0
 PUBLIC DF
 PUBLIC DC
 PUBLIC HF
 PUBLIC HC

 DF    = 0
 DC    = 0
 HF    = 0
 HC    = 0

 FOR IPO2=1 TO TPR

     IPO   = TPASORD(IPO2,0)



     DF    = TPASORD(IPO2,1) - 3
     HF    = TPASORD(IPO2,1) - 2
     DC    = 3
     HC    = 6

     FOR ISF=1 TO TPASORD(IPO2,1)

         FOR ISEC=1 TO TSC

           BCOL = (ISEC*8) - 8

           FOR ISC=1 TO 8
             IMF = BFILA + ISF
             IMC = BCOL + ISC
             XMATRIX(IMF,IMC,0) = "N"
             XMATRIX(IMF,IMC,1) = 0
             XMATRIX(IMF,IMC,2) = 0
             XMATRIX(IMF,IMC,3) = 0
             XMATRIX(IMF,IMC,6) = 0
             XMATRIX(IMF,IMC,7) = 0
             XMATRIX(IMF,IMC,8) = 0

             IF ISEC = TPASOK(IPO,2) THEN
                IF ISF = DF OR ISF = HF THEN
                   IF (ISC > (DC-1)) AND (ISC < (HC+1)) THEN
                      XMATRIX(IMF,IMC,0) = "R"
                      XMATRIX(IMF,IMC,4) = IPO
                   END IF
                END IF 
             END IF

             IF IMC = 1 THEN
                XMATRIX(IMF,IMC,3) = SETCONN(XMATRIX(IMF,IMC,3),1)
             END IF

             IF ISC = 8 THEN
                XMATRIX(IMF,IMC,3) = SETCONN(XMATRIX(IMF,IMC,3),100)
             END IF

             IF IMF = TOTFILAS THEN
                XMATRIX(IMF,IMC,3) = SETCONN(XMATRIX(IMF,IMC,3),1000)
             END IF


           NEXT  

         NEXT

     NEXT 


     BFILA = BFILA + TPASORD(IPO2,1)



 NEXT  

 '''''''' ARMA CONECTORES ''''''''''''

 PUBLIC DIRL  ' DIRECCCION HACIA LOS LADOS L - IZQ R - DER C - MISMA VIA
 PUBLIC DIRA  ' DIRECCCION HACIA ARRIBA-ABAJO D - ABAJO U - ARRIBA


 PUBLIC DFD
 PUBLIC DCD
 PUBLIC HFD
 PUBLIC HCD

 DFD    = 0
 DCD    = 0
 HFD    = 0
 HCD    = 0

 BFILA  = 0 
 BFILAD = 0 

 BFILADT = 0 

 FOR IPO2=1 TO TPR

     IPO     = TPASORD(IPO2,0)
     PASVA   = TPASOK(IPO,14)



     DF      = TPASORD(IPO2,1) - 3
     HF      = TPASORD(IPO2,1) - 2
     DC      = 3
     HC      = 6

     PASOK   = TPASOK(IPO,3)
     PASRCH  = TPASOK(IPO,4)
     IPASOK  = TPASOK(IPO,10)
     IPASRCH = TPASOK(IPO,11)

     IF PASOK <> "NO" THEN
        IPO3   = TPASOK(IPO,10)
        PASOK  = TPASORD(IPO3,0)
        DFD    = TPASORD(IPO3,1) - 3
        DCD    = TPASORD(IPO3,1) - 2
        HFD    = 3
        HCD    = 6

        ISEC   = TPASOK(IPO,2) 
        ISEC2  = TPASOK(PASOK,2) 

        DC     = (ISEC*8) - 8 + 5
        DCD    = (ISEC2*8) - 8 + 5
        DF     = TPASORD(IPO2,1) - 1 + BFILA

        BFILAD = 0

        FOR IPO4=1 TO (IPO3-1)
            BFILAD = BFILAD + TPASORD(IPO4,1) 
        NEXT

'''        DFD    = TPASORD(IPO3,1) - (4 + TPASOK(PASOK,13) ) + BFILAD
        DFD    = TPASORD(IPO3,1) - (TPASORD(IPO3,1) - (2 + TPASOK(PASOK,13)) ) + BFILAD
        TPASOK(PASOK,13) = TPASOK(PASOK,13) + 1  
        

        DIRL   = "N"   ' DIRECCCION HACIA LOS LADOS L - IZQ R - DER C - MISMA VIA
        DIRA   = "N"   ' DIRECCCION HACIA ARRIBA-ABAJO D - ABAJO U - ARRIBA

        IF IPO3 > IPO2 THEN
           DIRA = "D"
        ELSE
           DIRA = "U"
        END IF

        IF ISEC2 > ISEC THEN
           DIRL = "R"
        ELSEIF ISEC2 < ISEC THEN
           DIRL = "L"
        ELSE
           DIRL = "C"
        END IF
        
        IF TPASOK(PASOK,14) = "S" and TPASOK(IPO,14) = "S" THEN
           IF TPASOK(PASOK,17) = "NO"  THEN
              WXX = ARMAWAY(PASOK,IPO,6,"O")
           ELSE
              WXX = ARMAWAY(PASOK,IPO,1,"O")
           END IF 
        ELSE
           WXX = ARMAWAY(PASOK,IPO,7,"O")
        END IF

     END IF


     IF PASRCH <> "NO" THEN
        IPO3   = TPASOK(IPO,11)
        PASRCH = TPASORD(IPO3,0)
        DFD    = TPASORD(IPO3,1) - 3
        DCD    = TPASORD(IPO3,1) - 2
        HFD    = 3
        HCD    = 6

        ISEC   = TPASOK(IPO,2) 
        ISEC2  = TPASOK(PASRCH,2) 

        DC     = (ISEC*8) - 8 + 3
        DCD    = (ISEC2*8) - 8 + 3
        DF     = TPASORD(IPO2,1) - 1 + BFILA

        BFILAD = 0

        FOR IPO4=1 TO (IPO3-1)
            BFILAD = BFILAD + TPASORD(IPO4,1) 
        NEXT

'''        DFD    = TPASORD(IPO3,1) - (4 + TPASOK(PASRCH,13) ) + BFILAD

        DFD    = TPASORD(IPO3,1) - (TPASORD(IPO3,1) - (2 + TPASOK(PASRCH,13)) ) + BFILAD


        TPASOK(PASRCH,13) = TPASOK(PASRCH,13) + 1  

        DIRL   = "N"   ' DIRECCCION HACIA LOS LADOS L - IZQ R - DER C - MISMA VIA
        DIRA   = "N"   ' DIRECCCION HACIA ARRIBA-ABAJO D - ABAJO U - ARRIBA

        IF IPO3 > IPO2 THEN
           DIRA = "D"
        ELSE
           DIRA = "U"
        END IF

        IF ISEC2 > ISEC THEN
           DIRL = "R"
        ELSEIF ISEC2 < ISEC THEN
           DIRL = "L"
        ELSE
           DIRL = "C"
        END IF

      '''  WXX = ARMAWAY(PASRCH,IPO,2)
        
        IF TPASOK(PASRCH,14) = "S" AND TPASOK(IPO,14) = "S" THEN

           IF TPASOK(PASRCH,17) = "NO"  THEN
              WXX = ARMAWAY(PASRCH,IPO,6,"R")
           ELSE
              WXX = ARMAWAY(PASRCH,IPO,2,"R")
           END IF 	

	ELSE
	   WXX = ARMAWAY(PASRCH,IPO,7,"R")
	END IF


     END IF








     BFILA = BFILA + TPASORD(IPO2,1)



 NEXT  

      CELDAS = ""
      FILAS = ""
      
      colr = "#6699cc;"      

      STCB  = "COLOR:" & colr
      STB  = "BORDER-STYLE:DOTTED;BORDER-WIDTH:1;BORDER-COLOR:" & colr
      STBL = "BORDER-LEFT-STYLE:DOTTED;BORDER-LEFT-WIDTH:1;BORDER-LEFT-COLOR:" & colr
      STBR = "BORDER-RIGHT-STYLE:DOTTED;BORDER-RIGHT-WIDTH:1;BORDER-RIGHT-COLOR:" & colr
      STBT = "BORDER-TOP-STYLE:DOTTED;BORDER-TOP-WIDTH:1;BORDER-TOP-COLOR:" & colr
      STBB = "BORDER-BOTTOM-STYLE:DOTTED;BORDER-BOTTOM-WIDTH:1;BORDER-BOTTOM-COLOR:" & colr

      colr = "#F7AA22;"      

      STCU  = "COLOR:" & colr
      STRU  = "BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:" & colr
      STRLU = "BORDER-LEFT-STYLE:SOLID;BORDER-LEFT-WIDTH:2;BORDER-LEFT-COLOR:" & colr
      STRRU = "BORDER-RIGHT-STYLE:SOLID;BORDER-RIGHT-WIDTH:2;BORDER-RIGHT-COLOR:" & colr
      STRTU = "BORDER-TOP-STYLE:SOLID;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:" & colr
      STRBU = "BORDER-BOTTOM-STYLE:SOLID;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:" & colr

'      colr = "#00aa00;"      

'      STCU  = "COLOR:" & colr
'      STRU  = "BORDER-STYLE:DOTTED;BORDER-WIDTH:2;BORDER-COLOR:" & colr
'      STRLU = "BORDER-LEFT-STYLE:DOTTED;BORDER-LEFT-WIDTH:2;BORDER-LEFT-COLOR:" & colr
'      STRRU = "BORDER-RIGHT-STYLE:DOTTED;BORDER-RIGHT-WIDTH:2;BORDER-RIGHT-COLOR:" & colr
'      STRTU = "BORDER-TOP-STYLE:DOTTED;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:" & colr
'      STRBU = "BORDER-BOTTOM-STYLE:DOTTED;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:" & colr

      colr = "#bb0000;"      

      STCV  = "COLOR:" & colr
      STRV  = "BORDER-STYLE:DOTTED;BORDER-WIDTH:2;BORDER-COLOR:" & colr
      STRLV = "BORDER-LEFT-STYLE:DOTTED;BORDER-LEFT-WIDTH:2;BORDER-LEFT-COLOR:" & colr
      STRRV = "BORDER-RIGHT-STYLE:DOTTEDd;BORDER-RIGHT-WIDTH:2;BORDER-RIGHT-COLOR:" & colr
      STRTV = "BORDER-TOP-STYLE:DOTTED;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:" & colr
      STRBV = "BORDER-BOTTOM-STYLE:DOTTED;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:" & colr

      colr = "#00aa00;"      

      STCO  = "COLOR:" & colr
      STRO  = "BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:" & colr
      STRLO = "BORDER-LEFT-STYLE:SOLID;BORDER-LEFT-WIDTH:2;BORDER-LEFT-COLOR:" & colr
      STRRO = "BORDER-RIGHT-STYLE:SOLID;BORDER-RIGHT-WIDTH:2;BORDER-RIGHT-COLOR:" & colr
      STRTO = "BORDER-TOP-STYLE:SOLID;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:" & colr
      STRBO = "BORDER-BOTTOM-STYLE:SOLID;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:" & colr

      colr = "#bb0000;"      

      STCR  = "COLOR:" & colr
      STRR  = "BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:" & colr
      STRLR = "BORDER-LEFT-STYLE:SOLID;BORDER-LEFT-WIDTH:2;BORDER-LEFT-COLOR:" & colr
      STRRR = "BORDER-RIGHT-STYLE:SOLID;BORDER-RIGHT-WIDTH:2;BORDER-RIGHT-COLOR:" & colr
      STRTR = "BORDER-TOP-STYLE:SOLID;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:" & colr
      STRBR = "BORDER-BOTTOM-STYLE:SOLID;BORDER-BOTTOM-WIDTH:2;BORDER-BOTTOM-COLOR:" & colr

      colr = "#EEEEEE;"      

      STCN  = "COLOR:" & colr
      STRN  = "BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:" & colr
      STRLN = "BORDER-LEFT-STYLE:SOLID;BORDER-LEFT-WIDTH:1;BORDER-LEFT-COLOR:" & colr
      STRRN = "BORDER-RIGHT-STYLE:SOLID;BORDER-RIGHT-WIDTH:1;BORDER-RIGHT-COLOR:" & colr
      STRTN = "BORDER-TOP-STYLE:SOLID;BORDER-TOP-WIDTH:1;BORDER-TOP-COLOR:" & colr
      STRBN = "BORDER-BOTTOM-STYLE:SOLID;BORDER-BOTTOM-WIDTH:1;BORDER-BOTTOM-COLOR:" & colr

      FOR ISEC=1 TO TSC
          sInput = TSEC(ISEC,2)
          SINPUT = REPLACE(SINPUT," ","&nbsp;")
          celdas = celdas & ofv.GenCelda("","TT COLSPAN=8","CENTER","1%","","","",sInput,"si")
      NEXT

      filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

      IGNIMF = 0

      FOR IMF=1 TO TOTFILAS
          CELDAS = ""
          FOR IMC=1 TO TOTCOLUMNAS

             

             XMATCELL = XMATRIX(IMF,IMC,0)
             IF XMATCELL = "N" OR XMATCELL = "S1" OR XMATCELL = "S2"   OR XMATCELL = "C1"  OR XMATCELL = "C2"  OR XMATCELL = "C3"  OR XMATCELL = "C4" THEN

                EST = " STYLE='"
''                IF XMATRIX(IMF,IMC,6) > 0 THEN
'                   EST = EST & STCN
'                   SEN = XMATRIX(IMF,IMC,6)
'                   IF SEN > 1200 THEN
'                      SEN = SEN - 10000
'                      EST = EST & STRN
'                   END IF
'                   IF SEN > 120 THEN
'                      SEN = SEN - 1000
'                      EST = EST & STRBN
'                   END IF
'                   IF SEN > 12 THEN
'                      SEN = SEN - 100
'                      EST = EST & STRRN
'                   END IF
'                   IF SEN > 1 THEN
'                      SEN = SEN - 10
'                      EST = EST & STRTN
'                   END IF
'                   IF SEN = 1 THEN
'                      EST = EST & STRLN
'                   END IF
'
'                END IF

                IF XMATRIX(IMF,IMC,1) > 0 THEN
                   EST = EST & STCO
                   SEN = XMATRIX(IMF,IMC,1)
                   IF SEN > 1200 THEN
                      SEN = SEN - 10000
                      EST = EST & STRO
                   END IF
                   IF SEN > 120 THEN
                      SEN = SEN - 1000
                      EST = EST & STRBO
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      EST = EST & STRRO
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      EST = EST & STRTO
                   END IF
                   IF SEN = 1 THEN
                      EST = EST & STRLO
                   END IF

                END IF
                IF XMATRIX(IMF,IMC,2) > 0 THEN
                   EST = EST & STCR
                   SEN = XMATRIX(IMF,IMC,2)
                   IF SEN > 1200 THEN
                      SEN = SEN - 10000
                      EST = EST & STRR
                   END IF
                   IF SEN > 120 THEN
                      SEN = SEN - 1000
                      EST = EST & STRBR
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      EST = EST & STRRR
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      EST = EST & STRTR
                   END IF
                   IF SEN = 1 THEN
                      EST = EST & STRLR
                   END IF
                END IF 
                IF XMATRIX(IMF,IMC,3) > 0 THEN
                   EST = EST & STCB
                   SEN = XMATRIX(IMF,IMC,3)
                   IF SEN > 1200 THEN
                      SEN = SEN - 10000
                      EST = EST & STB
                   END IF
                   IF SEN > 120 THEN
                      SEN = SEN - 1000
                      EST = EST & STBB
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      EST = EST & STBR
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      EST = EST & STBT
                   END IF
                   IF SEN = 1 THEN
                      EST = EST & STBL
                   END IF
                END IF

                IF XMATRIX(IMF,IMC,6) > 0 THEN
                   EST = EST & STCU
                   SEN = XMATRIX(IMF,IMC,6)
                   IF SEN > 1200 THEN
                      SEN = SEN - 10000
                      EST = EST & STRU
                   END IF
                   IF SEN > 120 THEN
                      SEN = SEN - 1000
                      EST = EST & STRBU
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      EST = EST & STRRU
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      EST = EST & STRTU
                   END IF
                   IF SEN = 1 THEN
                      EST = EST & STRLU
                   END IF
                END IF 

                IF XMATRIX(IMF,IMC,8) > 0 THEN
                   EST = EST & STCV
                   SEN = XMATRIX(IMF,IMC,8)
                   IF SEN > 1200 THEN
                      SEN = SEN - 10000
                      EST = EST & STRV
                   END IF
                   IF SEN > 120 THEN
                      SEN = SEN - 1000
                      EST = EST & STRBV
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      EST = EST & STRRV
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      EST = EST & STRTV
                   END IF
                   IF SEN = 1 THEN
                      EST = EST & STRLV
                   END IF
                END IF 
                
                IF XMATCELL = "N" THEN 
                   EST = EST & "FONT-FAMILY:ARIAL;FONT-SIZE:1PT;' "
                   celdas = celdas & ofv.GenCelda("","AA" & EST,"","1%","","","","&nbsp;","si")
                ELSEIF XMATCELL = "S1" OR XMATCELL = "S2" THEN
                   EST = EST & "FONT-FAMILY:ARIAL;FONT-SIZE:6PT;' "
                   IPO   =  XMATRIX(IMF,IMC,4)
                   XPASO =  TPASOK(IPO,5)
                   if  XMATRIX(IMF,IMC,0) = "S1" THEN 
                       XPASO = XPASO & "<font face=webdings >4</font>"
                   else
                       XPASO = "<font face=webdings >4</font>" & XPASO
                   end if 
                   celdas = celdas & ofv.GenCelda("","AA" & EST,"CENTER","1%","","","",XPASO,"si")
                ELSE
                   EST = EST & "FONT-FAMILY:ARIAL;FONT-SIZE:6PT;' "
                   SELECT CASE XMATCELL
                          CASE "C1"
                                XPASO = "<font face=webdings COLOR=#336699 >5</font>"
                          CASE "C2"
                                XPASO = "<font face=webdings COLOR=#336699 >6</font>"
                          CASE "C3"
                                XPASO = "<font face=webdings COLOR=#336699 >4</font>"
                          CASE "C4"
                                XPASO = "<font face=webdings COLOR=#336699 >3</font>"
                   END SELECT 

                   celdas = celdas & ofv.GenCelda("","AA" & EST,"","1%","","","",XPASO,"si")
                END IF 
             ELSE
                IF (IGNIMF <> IMF) AND (IGNIMF <> (IMF - 1)) THEN
                   IPO   =  XMATRIX(IMF,IMC,4)
                   IF TPASOK(IPO,14) = "S" THEN
                      XPASO =  TPASOK(IPO,5) & ":" & TPASOK(IPO,8)
                      XPASO =  "<b>&nbsp;&nbsp;<u>" & REPLACE(XPASO," ","&nbsp;") & "</u>&nbsp;&nbsp;</b>"
                      if CKTOPC = "SI" then
                         xtipo   = "CKT"
                         xmodel  = TPASOK(IPO,16)
                         xcktpsc = TPASOK(IPO,18)
                      else
                         xtipo   = "WKF"
                         xmodel  = TPASOK(IPO,2)
                         xcktpsc = "NO"
                      end if 
                      XFUNCION = " onclick=" & chr(34) & "abrewin('" & xmodel & "','" & xtipo & "','" & xcktpsc & "');" & chr(34) & " "
                      celdas = celdas & ofv.GenCelda("","cc  COLSPAN=4 ROWSPAN=2 " & XFUNCION & " ","CENTER","1%","30","","",XPASO,"si")
                   ELSE
                      celdas = celdas & ofv.GenCelda("","AAA COLSPAN=4 ROWSPAN=2 ","CENTER","1%","","","","&nbsp;","si")
                   END IF
                   IGNIMF = IMF
                END IF 

             END IF
  
          NEXT


          filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")
      NEXT 



   sHTML = sHTML  & ofv.GenTabla("","AA STYLE='BORDER-STYLE:OUTSET;BORDER-WIDTH:1;filter:shadow(color=#dddddd,direction=125);' ","","90%","","0","","-1","-2","","",filas,"si") & "<BR>"

''''   sHTML = sHTML  & TABMAT 


   END IF



END IF

''''''''''' FIN ORDENA PASOS '''''''''''''''''''''''




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETRESPO(RESPCOD,RESPTIPO)

  sInput = "--- NO ASIGNADO ---------------------"
  
  IF RESPCOD = "NO" THEN RESPCOD = 0
  
  
  IF RESPTIPO = "MOD" THEN
  
  
  SELECT CASE RESPCOD
  
         CASE "O"
         
		 sInput = "ORIGINANTE"
         
         CASE "A"
         
		 sInput = "RESPONSABLE ASIGNADO"         
         
         CASE "L1"
         
		 sInput = "LISTA AUTOMATICA 1"                  
         CASE "L2"
         
		 sInput = "LISTA AUTOMATICA 2"                  
         CASE "L3"
         
		 sInput = "LISTA AUTOMATICA 3"                  
		 
         
         CASE ELSE
  
  
	    set cn1 = ofv.conectar(ofv.strconn5)
	    set cn2 = ofv.conectar(ofv.strconn0)
	    set cn3 = ofv.conectar(ofv.strconn2)



	  StrSql = "Select * From WKFRESP WHERE ID = 0" & RESPCOD
	  set rsT = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
	  if not rsT.eof then 

	    XRESPO = RST(3) & ""

	    if XRESPO <> "" then

	       SELECT CASE rsT(2)

		CASE "USR"
		   strsql = "Select * From USUARIOS where ID = '" & rsT(3) & "' "
		   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

		CASE "GRP"
		   strsql = "Select * From GRUPOS where ID = '" & rsT(3) & "' "
		   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

		CASE "SEC"
		   strsql = "Select * From ECO where ECO = " & rsT(3)
		   set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

	       END SELECT

	    END IF


	  end if

	  call ofv.cerrarconn(cn1)
	  call ofv.cerrarconn(cn2)
	  call ofv.cerrarconn(cn3)
  
  END SELECT
  
  ELSE

 
		 XRESPO = RESPTIPO
         
                 
  
  
	    set cn1 = ofv.conectar(ofv.strconn5)
	    set cn2 = ofv.conectar(ofv.strconn0)
	    set cn3 = ofv.conectar(ofv.strconn2)




	       SELECT CASE XRESPO

		CASE "USR"
		   strsql = "Select * From USUARIOS where ID = '" & RESPCOD & "' "
		   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

		CASE "GRP"
		   strsql = "Select * From GRUPOS where ID = '" & RESPCOD & "' "
		   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

		CASE "SEC"
		   strsql = "Select * From ECO where ECO = " & RESPCOD
		   set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
		   if not rsniv.eof then sInput = rsniv(1)
		   call ofv.cerrarconsulta(rsniv)

	       END SELECT

	

	  call ofv.cerrarconn(cn1)
	  call ofv.cerrarconn(cn2)
	  call ofv.cerrarconn(cn3)
  
  
  
  END IF

  GETRESPO = sInput

END FUNCTION

FUNCTION SETCONN(VOLD,VNEW)

S10000 = "N"
S1000  = "N"
S100   = "N"
S10    = "N"
S1     = "N"
                   SEN = VOLD
                   IF SEN > 1200 THEN
                      SEN   = SEN - 10000
                      S10000 = "S"
                   END IF
                   IF SEN > 120 THEN
                      SEN   = SEN - 1000
                      S1000 = "S"
                   END IF
                   IF SEN > 12 THEN
                      SEN = SEN - 100
                      S100 = "S"
                   END IF
                   IF SEN > 1 THEN
                      SEN = SEN - 10
                      S10 = "S"
                   END IF
                   IF SEN = 1 THEN
                      S1 = "S"
                   END IF

         SETCONN = VOLD

         IF S10000 = "S" THEN  EXIT FUNCTION

         SELECT CASE VNEW
                CASE 10000
                     IF S10000 <> "S" THEN  SETCONN = VOLD + VNEW

                CASE 1000
                     IF S1000 <> "S" THEN  SETCONN = VOLD + VNEW

                CASE 100
                     IF S100 <> "S" THEN  SETCONN = VOLD + VNEW

                CASE 10
                     IF S10 <> "S" THEN  SETCONN = VOLD + VNEW

                CASE 1  
                     IF S1 <> "S" THEN  SETCONN = VOLD + VNEW

         END SELECT


END FUNCTION

FUNCTION ARMAWAY(DESDE,HASTA,IDX,XT)

    IF DIRA = "x" THEN

    ELSE

    IF DIRL = "I" THEN

    ELSE

''''''' SALE ''''''''''

        IF DIRL = "L" THEN

           IF XT="O" THEN

              IF DIRA = "U" THEN 
        '''   RESPONSE.WRITE DIRA & " - " & XT & " - " & DFD & " - " & DCD  & " - " & DF & " - " & DC & "XX1<BR>"
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE

                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC,IDX) = SETCONN(XMATRIX(DF,DC,IDX),100)    
              END IF
           ELSE

              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC,IDX) = SETCONN(XMATRIX(DF,DC,IDX),100)    
              END IF
           END IF
        ELSEIF DIRL = "C" THEN
           IF XT="O" THEN
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1) 
              END IF
           ELSE
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1) 
              END IF
           END IF             

        ELSE
           IF XT="O" THEN
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1) 
              END IF
           ELSE
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DF,DC+1,0) = "C2"
                 XMATRIX(DF,DC+1,IDX) = SETCONN(XMATRIX(DF,DC+1,IDX),1) 
              END IF
           END IF             
        END IF 
         
''''''' VA ''''''''''
        IF DIRA = "U" THEN 
          ''' IDX=2
           IF XT="O" THEN
              FOR IM=1 TO 2 
                  XMATRIX(DF + 1,DC + IM,IDX) = SETCONN(XMATRIX(DF + 1,DC + IM,IDX),10)
              NEXT
              DC = DC + 2
           ELSE
              FOR IM=1 TO 4 
                  XMATRIX(DF + 1,DC + IM,IDX) = SETCONN(XMATRIX(DF + 1,DC + IM,IDX),10)
              NEXT
              DC = DC + 4
           END IF 


''''''' SUBE '''''''''''''''
           IMR = 1
           FOR IM=DFD TO DF
               XMATRIX(IM,DC,IDX) = SETCONN(XMATRIX(IM,DC,IDX),100)
               IF ((IM - DFD)  / (5 * IMR)) = 1 THEN
                  IMR = IMR + 1
                  IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C1"
               END IF
           NEXT
         '''  DC = DC - 2
        ELSE
      
''''''' BAJA '''''''''''''''


        IF DIRL = "L" THEN
           IF XT="O" THEN
              IMR = 1
              FOR IM=DF TO DFD-1
                  XMATRIX(IM,DC,IDX) = SETCONN(XMATRIX(IM,DC,IDX),100)
                  IF ((IM - DF)  / (5 * IMR)) = 1 THEN
                     IMR = IMR + 1
                     IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C2"
                  END IF
              NEXT
           ELSE
              IMR = 1
              FOR IM=DF TO DFD-1
                  XMATRIX(IM,DC,IDX) = SETCONN(XMATRIX(IM,DC,IDX),100)
                  IF ((IM - DF)  / (5 * IMR)) = 1 THEN
                     IMR = IMR + 1
                     IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C2"
                  END IF
              NEXT
           END IF 
        ELSEIF DIRL = "R" THEN
           IMR = 1
           FOR IM=DF+1 TO DFD-1
               XMATRIX(IM,DC+1,IDX) = SETCONN(XMATRIX(IM,DC,IDX),1)
               IF ((IM - (DF+1))  / (5 * IMR)) = 1 THEN
                  IMR = IMR + 1
                  IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C2"
               END IF
           NEXT
        ELSE
           IMR = 1
           IF XT = "O" THEN
           FOR IM=DF+1 TO DFD-1
               XMATRIX(IM,DC+1,IDX) = SETCONN(XMATRIX(IM,DC,IDX),1)
               IF ((IM - (DF+1))  / (5 * IMR)) = 1 THEN
                  IMR = IMR + 1
                  IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C2"
               END IF
           NEXT
           ELSE
          ''' FOR IM=DF TO DFD-1
           FOR IM=DF+1 TO DFD-1
               '''IDX = 3 
               XMATRIX(IM,DC+1,IDX) = SETCONN(XMATRIX(IM,DC,IDX),1)
               IF ((IM - (DF+1))  / (5 * IMR)) = 1 THEN
                  IMR = IMR + 1
                  IF IDX < 7 THEN XMATRIX(IM,DC+1,0) = "C2"
               END IF
           NEXT
           END IF
        END IF 



        END IF

''''''' VA A DESTINO ''''''''''
        IF DIRL = "R" THEN 
           IF XT="O" THEN
              IF DIRA = "U" THEN 
                 IMR = 1
                 FOR IM=DC+1 TO DCD-1
                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DC+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C3"
                     END IF
                 NEXT
              ELSE
                 IMR = 1
                 FOR IM=DC+1 TO DCD-2
                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DC+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C3"
                     END IF
                 NEXT
              END IF 
           ELSE
              IF DIRA = "U" THEN 
                 IMR = 1
                 FOR IM=DC+1 TO DCD+2
                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DC+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C3"
                     END IF
                 NEXT
              ELSE
                 IMR = 1
                 FOR IM=DC+1 TO DCD+2
                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DC+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C3"
                     END IF
                 NEXT
              END IF             
           END IF

        ELSEIF DIRL = "L" THEN 
           IF XT="O" THEN
           '   IF DIRA = "U" THEN DC = DC + 2
              IMR = 1
              FOR IM=DCD-1 TO DC
                XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                   IMR = IMR + 1
                   IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                END IF
              NEXT
            '  IF DIRA = "U" THEN DC = DC - 2
           ELSE
              '''IF DIRA = "U" THEN DC = DC + 2
              IMR = 1
              FOR IM=DCD+3 TO DC
                XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                   IMR = IMR + 1
                   IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                END IF
              NEXT
            '''  IF DIRA = "U" THEN DC = DC - 2

           END IF
        ELSE
           IF XT = "O" THEN   
             IF DIRA = "U" THEN 
            '''  DC = DC + 2
              IMR = 1
              FOR IM=DCD-1 TO DC
                  XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                  IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                     IMR = IMR + 1
                     IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                  END IF
              NEXT
            '''  DC = DC - 2
             ELSE
              DC = DC + 2
              IMR = 1
              FOR IM=DCD+1 TO DC
                  XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                  IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                     IMR = IMR + 1
                     IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                  END IF
              NEXT
              DC = DC - 2
             END IF

           ELSE
              IF DIRA = "U" THEN 
             ''' DC = DC + 2
                 IMR = 1
              '''IDX = 1 
             ''' FOR IM=DCD+1 TO DC
              '''   RESPONSE.WRITE DC & "-" & DCD & "<BR>"
                 FOR IM=DCD+3 TO DC

                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                     END IF
                 NEXT
            '''  DC = DC - 2
             ''' IDX = 2 
              ELSE
             ''' DC = DC + 2
                 IMR = 1
             ''' IDX = 1 
             ''' FOR IM=DCD+1 TO DC

                 FOR IM=DC+1 TO DCD+2

                     XMATRIX(DFD,IM,IDX) = SETCONN(XMATRIX(DFD,IM,IDX),10)
                     IF ((IM - (DCD+1))  / (5 * IMR)) = 1 THEN
                        IMR = IMR + 1
                        IF IDX < 7 THEN XMATRIX(DFD,IM,0) = "C4"
                     END IF
                 NEXT
            '''  DC = DC - 2
             ''' IDX = 2 
            END IF
           END IF 
        END IF

''''''' LLEGA ''''''''''

        IF DIRL = "L" THEN
           IF XT = "O" THEN
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DFD,DCD-1,0) = "C2"
                 XMATRIX(DFD,DCD-1,IDX) = SETCONN(XMATRIX(DFD,DCD-1,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DFD,DCD-1,0) = "C2"
                 XMATRIX(DFD,DCD-2,IDX) = SETCONN(XMATRIX(DFD,DCD-2,IDX),100)
              END IF  
           ELSE
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DFD,DCD+3,0) = "C2"
                 XMATRIX(DFD,DCD+2,IDX) = SETCONN(XMATRIX(DFD,DCD+2,IDX),100)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DFD,DCD+3,0) = "C2"
                 XMATRIX(DFD,DCD+2,IDX) = SETCONN(XMATRIX(DFD,DCD+2,IDX),100)
              END IF  
           END IF 
        ELSEIF DIRL = "C" THEN
            IF XT = "O" THEN

              IF IDX < 7 THEN XMATRIX(DFD,DCD-1,0) = "C2"
              XMATRIX(DFD,DCD-2,IDX) = SETCONN(XMATRIX(DFD,DCD-2,IDX),100)
            ELSE

              IF IDX < 7 THEN XMATRIX(DFD,DCD+3,0) = "C2"
              XMATRIX(DFD,DCD+2,IDX) = SETCONN(XMATRIX(DFD,DCD+2,IDX),100)


           END IF 
        ELSE
            IF XT = "O" THEN
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DFD,DCD,0) = "C2"
                 XMATRIX(DFD,DCD,IDX) = SETCONN(XMATRIX(DFD,DCD,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DFD,DCD-1,0) = "C2"
                 XMATRIX(DFD,DCD-1,IDX) = SETCONN(XMATRIX(DFD,DCD-1,IDX),1)
              END IF
            ELSE
              IF DIRA = "U" THEN 
                 IF IDX < 7 THEN XMATRIX(DFD,DCD+3,0) = "C2"
                 XMATRIX(DFD,DCD+3,IDX) = SETCONN(XMATRIX(DFD,DCD+3,IDX),1)
              ELSE
                 IF IDX < 7 THEN XMATRIX(DFD,DCD+3,0) = "C2"
                 XMATRIX(DFD,DCD+3,IDX) = SETCONN(XMATRIX(DFD,DCD+3,IDX),1)
              END IF
            END IF
        END IF 

     END IF
     END IF

END FUNCTION


%>