<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

id = request.querystring("fname")
id = ofv.convertircar(id,"*"," ")
tabla = request.querystring("tabla")
volver = request.querystring("volver")
volver = ofv.convertircar(volver,"_","&")

on error resume next

IF UCASE(LEFT(TABLA,3)) = "WKF" OR  UCASE(LEFT(TABLA,3)) = "SEG" THEN
   set cn = ofv.conectar(ofv.strconn5)
ELSE
   set cn = ofv.conectar(ofv.strconn6)
END IF

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Update " & tabla & " Set "

For Each x In Request.form
  vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
   ' response.write x & ": " & vNuevoValor & "<br>"

  if UCase(vNuevoValor) = "FALSE" then vNuevoValor = "0"
  if UCase(vNuevoValor) = "TRUE" then vNuevoValor = "1"
  nx = ucase(x)
  if (nx = "FECHAENT" or nx = "FECEMIS" or nx = "FECBAJA" or nx = "FECHATEN"  or nx = "FECVIGEN" or nx = "FECVTO" or nx = "FECVIG" or nx = "FECHAVIG" or nx = "ULTCIERRE") then



     if vNuevoValor = "" then
        vNuevoValor = clng(cdate("01-01-1900")) - nFecini
    else
        vNuevoValor = clng(cdate(vNuevoValor)) - nFecini
        


     end if
  end if

    if vNuevoValor = "NSNC" then vNuevoValor = 0

  s = "'"

  strsql = strsql & x  & " = " & s & vNuevoValor & s & ","
Next


strsql = mid(strsql,1,len(strsql)-1) & " Where " & rs(0).name & "="
strsql = strsql & s & id & s

'' response.write strsql
call ofv.crearconsultaEx(StrSql,cn,1,volver)

if cn.errors.count = 0 then
   if volver <> "" then response.redirect volver
  ' RESPONSE.WRITE VOLVER
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>