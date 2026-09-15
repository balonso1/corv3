<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

id = request.querystring("fname")
id = ofv.convertircar(id,"*"," ")
tabla = request.querystring("tabla")
volver = request.querystring("volver")
volver = ofv.convertircar(volver,"_","&")

on error resume next

set cn = ofv.conectar(ofv.strconn5)
nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

strsql = "Select * From " & tabla & " WHERE ID = " & ID
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


IF NOT RS.EOF THEN


IF TABLA = "WKFPASOS" THEN

For Each x In Request.form

    vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))


SELECT CASE X

       CASE "XXXTIPO"

             SELECT CASE vNuevoValor
                    CASE "PI"
                         STRSQL = "DELETE FROM WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','INI','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '', RCHPSC = '', APR = '0', RCH = '0', FIN = '0', INI = '1' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "PF"

                         STRSQL = "DELETE FROM WKFPSCTPC WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','FIN','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '', RCHPSC = '', APR = '0', RCH = '0', FIN = '1', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "PC"

                         STRSQL = "DELETE FROM WKFPSCTPC WHERE WKFPASCOD = '"  & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','APR','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','RCH','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '', RCHPSC = '', APR = '1', RCH = '1', FIN = '0', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


             END SELECT



       CASE "APR"


             SELECT CASE vNuevoValor
                    CASE "SI"

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','APR','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET APR = '1', FIN = '0', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "NO"

                         IF RS(8) THEN

                         STRSQL = "DELETE FROM WKFPSCTPC WHERE WKFPASCOD = '"  & session("WKFPSC") & "' AND TIPO = 'APR' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '', APR = '0', FIN = '0', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         END IF 


             END SELECT



       CASE "RCH"

             SELECT CASE vNuevoValor
                    CASE "SI"

                         STRSQL = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) "
                         STRSQL = STRSQL & "values ('" & session("WKF") & "','" & session("WKFPSC")
                         STRSQL = STRSQL & "','RCH','N','0','N')"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET RCH = '1', FIN = '0', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "NO"

                         IF RS(7) THEN

                         STRSQL = "DELETE FROM WKFPSCTPC WHERE WKFPASCOD = '"  & session("WKFPSC") & "' AND TIPO = 'RCH' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPASOS SET RCHPSC = '', RCH = '0', FIN = '0', INI = '0' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         END IF 


             END SELECT



END SELECT



NEXT




END IF 


IF TABLA = "WKFPSCTPC" THEN

For Each x In Request.form

    vNuevoValor = Request.Form(x)


SELECT CASE X

       CASE "PASOSIG"


             SELECT CASE RS(3) ' TIPO

                    CASE "INI"

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPSCTPC SET PASOSIG = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "APR"

                         STRSQL = "UPDATE WKFPASOS SET APRPSC = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPSCTPC SET PASOSIG = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                    CASE "RCH"

                         STRSQL = "UPDATE WKFPASOS SET RCHPSC = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         STRSQL = "UPDATE WKFPSCTPC SET PASOSIG = '" & vNuevoValor & "' "
                         STRSQL = STRSQL & " WHERE ID = " & ID
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


             END SELECT


END SELECT



NEXT



END IF


if cn.errors.count = 0 then
  if volver <> "" then response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if

ELSE
   response.redirect volver
END IF

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>