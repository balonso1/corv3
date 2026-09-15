<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

ID = request.querystring("ID")
tabla = request.querystring("tabla")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
Volver = ofv.convertircar(Volver,"*","_")

if isnull(tabla) then Tabla = "Menux"

if Ucase(Tabla) = "ATRM" then 
   set cn = ofv.conectar(ofv.strconn0)
else
   set cn = ofv.conectar(ofv.strconn1)
end if

strsql = "Select * From " & Tabla & " Where ID = " & id
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Update " & Tabla & " Set "
bEjecutar = False
For Each x In Request.form
  if isnull(Request.Form(x)) then
     vNuevoValor = "NSNC"
  else 
     vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  end if
  if Ucase(vNuevoValor) = "FALSE" then vNuevoValor = "0"
  if Ucase(vNuevoValor) = "TRUE" then vNuevoValor = "1"
  TipoValido = ofv.ValidarTipo(rs(x).name,vNuevoValor)
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
  if ValorNulo = True then exit for
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type > 199 AND rs(x).type < 300 THEN 
     s = "'"
  end if

'  response.write x & " " & s & " " & rs(x).type & " " & vNuevoValor

  xxx = rs(x).Value
  xx2 = xxx
  if isnull(xx2) then xx2 = "9999999999999999"
  
  if vNuevoValor <> "NSNC" then
    if isnull(vNuevoValor) and isnull(xxx) then
       xxx = " "
    elseif isnull(xxx) and vNuevoValor <> "" then
      bEjecutar = True
      strsql = strsql & x & " = " & s & vNuevoValor & s & " ,"
    elseif cstr(vNuevoValor) <> cstr(xx2) then
      bEjecutar = True
      strsql = strsql & x & " = " & s & vNuevoValor & s & " ,"
    end if
  else
    bEjecutar = True
    if s = "'" then
      strsql = strsql & x & " = '',"
    else
      strsql = strsql & x & " = 0,"
    end if
  end if
Next 
 
if TipoValido = False or ValorNulo = True then
  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & vNuevoValor
else
  s = ofv.QueEsx(rs(0).type)
  strsql = mid(strsql,1,len(strsql)-1) & " Where " & rs(0).name & "=" & s & id & s


  if bEjecutar then
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
  end if

  if cn.errors.count = 0 then
    response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if 
end if 

call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>