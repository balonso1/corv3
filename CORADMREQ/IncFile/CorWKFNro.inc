<%



'''''''''''''''''''''''''
FUNCTION GENSECTOROR()



  StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rsX.eof then
     XTIPOREQ = RSX(15)
     XUSRORIG = RSX(13)
     XSECTORIG = 0
     XUSRANA = RSX(14)
     XSECTANA = 0
     XASOCPRY = RSX(15)

     StrSql = "Select * From USUARIOS where ID = '" & XUSRORIG & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG = rsX2(21)
        XSECTORIGB = XSECTORIG
        XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTORIGB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "INSERT INTO PROYRESP (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,SKR,CLASE,SK1,SK2,SK3,SK4,SK5,FUNCION) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','10','" & XSECTORIG & "','" & XSECTGEREN & "','OR','','" & XUSRORIG & "','CT','U','1','0','1','1','0','3')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

        XXXX = GENSECTORACC(XSECTORIG,"SA")
        XXXX = GENSECTORACC(XSECTGEREN,"RA")
        XXXX = GENSECTORACC("999999","GL")


     END IF


     StrSql = "Select * From USUARIOS where ID = '" & XUSRANA & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTANA = rsX2(21)
        XSECTANAB = XSECTANA
        XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTANAB
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTANAB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTANAB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "INSERT INTO PROYRESP (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,SKR,CLASE,SK1,SK2,SK3,SK4,SK5,FUNCION) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','20','" & XSECTANA & "','" & XSECTGEREN & "','AN','','" & XUSRANA & "','CT','U','1','1','1','1','0','2')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

        XXXX = GENSECTORACC(XSECTANA,"SA")
        XXXX = GENSECTORACC(XSECTGEREN,"RA")
        XXXX = GENSECTORACC("999999","GL")


     END IF


     IF XASOCPRY <> "NO" THEN

        IF XASOCPRY = "REQ" THEN
           StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'AF' ORDER BY ORDEN "
           set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           DO WHILE NOT RSX1.EOF
              XXX = GENSECTORAF(RSX1(3),RSX1(7))
              RSX1.MOVENEXT
           LOOP
            
        END IF
     END IF 




  END IF



END FUNCTION

FUNCTION GENSECTORAF(XSECTOR,XUSRAFC)


  StrSql = "Select MAX(ORDEN) From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO = 'AF' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     XORDEN = 30
  ELSE

     XORDEN = RS(0)

     IF XORDEN <> "" THEN
        XORDEN = XORDEN + 10
     ELSE
        XORDEN = 30
     END IF

  END IF



     XSECTORIG = XSECTOR
     XSECTORIGB = XSECTORIG
     XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTORIGB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "INSERT INTO PROYRESP (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,SKR,CLASE,SK1,SK2,SK3,SK4,SK5,FUNCION) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XORDEN & "','" & XSECTORIG & "','" & XSECTGEREN & "','AF','',"
        STRSQL = STRSQL & "'" & XUSRAFC & "','CT','U','1','0','1','1','0','1')" 



        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

        XXXX = GENSECTORACC(XSECTORIG,"SA")
        XXXX = GENSECTORACC(XSECTGEREN,"RA")
        XXXX = GENSECTORACC("999999","GL")






END FUNCTION

FUNCTION GENSECTORACC(XSECTOR,XUSRAFC)

   XSECTORIG  = XSECTOR
   XSECTGEREN = XSECTOR

''''''''''''
  StrSql = "Select ID From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO = '" & XUSRAFC & "' AND GERENCIA = " & XSECTOR 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

''''''''''''


  StrSql = "Select MAX(ORDEN) From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO in ('AF','SA','RA','GL') "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     XORDEN = 30
  ELSE

     XORDEN = RS(0)

     IF XORDEN <> "" THEN
        XORDEN = XORDEN + 10
     ELSE
        XORDEN = 30
     END IF

  END IF







        STRSQL = "INSERT INTO PROYRESP (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,SKR,CLASE,SK1,SK2,SK3,SK4,SK5) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XORDEN & "','" & XSECTORIG & "','" & XSECTGEREN & "','" & XUSRAFC & "','',"
        STRSQL = STRSQL & "'NO','A','U','1','0','0','0','0')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql






 END IF




END FUNCTION

''''''''''''''''''''''''

function GTWKF()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "WKF"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTWKF = trim(GTsecuf)

end function

function GTWKFPSC()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "WKFPSC"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTWKFPSC = trim(GTsecuf)

end function

function GTCKT()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "CKT"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTCKT = trim(GTsecuf)

end function

function GTCKTPSC()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "CKTPSC"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTCKTPSC = trim(GTsecuf)

end function

function GTREQ()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "REQ"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTREQ = trim(GTsecuf)

end function

function GTPROY()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "PRY"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTPROY = trim(GTsecuf)

end function

function GTETP()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "ETP"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTETP = trim(GTsecuf)

end function

function GTTSK()

    sh = datepart("h",date())
    sm = datepart("n",date())
    ss = datepart("s",date())

XCodobJ = sh & sm & ss & rnd()

GThms = datediff("d",date(),cdate("12/12/2050")) & formatdatetime(now(),3)
GTses = session.sessionid
GTsecu = cstr(XCodobJ & GThms & GTses)

GTsecuf = "TSK"

for GTi=1 to len(GTsecu)
GTcs = mid(GTsecu,GTi,1)
if GTcs => "0" and GTcs <= "9" then
   GTsecuf = GTsecuf & GTcs
end if
next
 
GTTSK = trim(GTsecuf)

end function





%>