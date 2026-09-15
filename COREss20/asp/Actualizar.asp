<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

ID = request.querystring("ID")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
Tabla = request.querystring("Tabla")

if UCASE(Tabla) = "CORCRPT" THEN
   set cn = ofv.conectar(ofv.strconn1)
ELSE
   set cn = ofv.conectar(ofv.strconn0)
END IF

if Tabla = "NivelConfiden" then
   strsql = "Select * From " & Tabla & " Where NCCod = " & id
elseif Tabla = "skrdef" then
   strsql = "Select * From " & Tabla
elseif UCASE(Tabla) = "CORCRPT" then
   strsql = "Select * From " & Tabla & " Where OBJID = " & id
elseif (UCASE(Tabla) = "FERIADOS") or (UCASE(Tabla) = "RANGOS") then
   strsql = "Select * From " & Tabla & " Where ID = " & id
else
   strsql = "Select * From " & Tabla & " Where ID = '" & id & "'"
end if

set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
strsql = "Update " & Tabla & " Set "
bEjecutar = False
For Each x In Request.form
  if (not ucase(x) = "DBPWD") and  (not ucase(x) = "EMAIL") then
     vNuevoValor = UCase(CStr(ofv.convertircar(Request.Form(x),"_"," ")))
  else
     vNuevoValor = trim(Request.Form(x))
  end if
  if Ucase(vNuevoValor) = "FALSE" then vNuevoValor = "0"
  if Ucase(vNuevoValor) = "TRUE" then vNuevoValor = "1"
  TipoValido = ofv.ValidarTipo(rs(x).name,vNuevoValor)
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),vNuevoValor)
  if ValorNulo = True then exit for
  s = ofv.QueEsx(rs(x).type)
  if rs(x).type > 100 AND rs(x).type < 300 THEN 
     s = "'"
  end if
  if ucase(x) = "PASSMIN" or ucase(x) = "PASSMAX" then 
     if isnumeric(vNuevoValor) then
        if clng(vNuevoValor) > 30 then vNuevoValor = 30
        if clng(vNuevoValor) < 5 then vNuevoValor = 5
     else 
        vNuevoValor = 6
     end if
  end if


  if vNuevoValor <> "NSNC" then
     if not isnull(rs(x).Value) then
        voldvalue = cstr(rs(x).Value)
     else
        voldvalue = ""
     end if 
     if not isnull(vNuevoValor) then
        vNuevoValor = cstr(vNuevoValor)
     else
        vNuevoValor = ""
     end if 
     if ucase(x) = "DBPWD" then
'        response.write "pwd " & vNuevoValor & "' "
        if vNuevoValor <> "********" then
           if ucase(vNuevoValor) = "NULL" or vNuevoValor = "" then 
              vNuevoValor = ""
           else
              vNuevoValor = ofv.getcript("E",cstr(vNuevoValor))
           end if 
           bEjecutar = True
           strsql = strsql & x & " = " & s & vNuevoValor & s & " ,"
        end if
     else 

'response.write rs(x).name &  "-vNuevoValor " & vNuevoValor & " voldvalue " & voldvalue & " tipo " & rs(x).type & "<br>"
    if vNuevoValor <> voldvalue then
      bEjecutar = True
      strsql = strsql & x & " = " & s & vNuevoValor & s & " ,"
    elseif voldvalue = "" and vNuevoValor <> "" then
      bEjecutar = True
      strsql = strsql & x & " = " & s & vNuevoValor & s & " ,"
     end if
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
if Tabla = "skrdef" then
  strsql = mid(strsql,1,len(strsql)-1)
else
  s = ofv.QueEsx(rs(0).type)
  if rs(0).type > 199 AND rs(0).type < 300 THEN 
     s = "'"
  end if
  strsql = mid(strsql,1,len(strsql)-1) & " Where " & rs(0).name & "=" & s & id & s

end if

'    response.write bEjecutar & "strsql " & strsql

'  if id = "DEF" and trim(request.form("ID")) <> "DEF" and Tabla = "Grupos" then
'    sMensaje = "No puede modificar el ID del grupo predeterminado " & trim(request.form("ID"))
'    sVolver = "Grupos.asp"
'    sHTML = ofv.Mensajes(sMensaje,sVolver,"1")
'    response.write sHTML
'  elseif bEjecutar then

  if bEjecutar then
'    response.write "strsql " & strsql
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