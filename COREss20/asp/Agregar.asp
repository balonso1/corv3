<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

Tabla = request.querystring("Tabla")
volver = request.querystring("volver")

set cn = ofv.conectar(ofv.strconn0)

sID = trim(request.form("ID"))
set Encripter = server.createobject("CorEncriptadorEx.SHA1")
sEncript = Encripter.Encrypt(sID)

snova = 0

if Ucase(tabla) = "GRUPOS" or  Ucase(tabla) = "USUARIOS" then
   strsql = "Select * From " & Tabla & " where ID ='" & sID & "' "
   set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
   if not rs.eof then snova = 1
end if

if snova = 1 then
   response.redirect "errordup.asp?volver=" & volver & "&tabla=" & Ucase(tabla) & "&ID=" & sID
else

if Tabla = "Usuarios" then
   strsql = "Select * From skrdef"
   set rx = ofv.crearconsultaEx(StrSql,cn,1,volver)
   xlcha = clng(date()) - clng(cdate("01-01-1900"))
   xHcha1 = time()
   xHcha = ""
   for xind=1 to len(xHcha1)
       if isnumeric(mid(xHcha1,xind,1)) then
          xHcha = xHcha & mid(xHcha1,xind,1)
       end if
   next
end if



strsql = "Select * From " & Tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert Into " & Tabla & " ("
if Tabla = "Usuarios" then
   strsql = strsql & "Password,"
   strsql = strsql & "nivel,"
   strsql = strsql & "opcion,"
   strsql = strsql & "lchange,"
   strsql = strsql & "logfail,"
   strsql = strsql & "defskr,"
   strsql = strsql & "passmin,"
   strsql = strsql & "passmax,"
   strsql = strsql & "tipoesq,"
   strsql = strsql & "passhist,"
   strsql = strsql & "intfall,"
   strsql = strsql & "expira,"
   strsql = strsql & "diasexp,"
   strsql = strsql & "hchange,"
   strsql = strsql & "Rangos,"
end if

if Ucase(tabla) = "GRUPOS" then
   strsql = strsql & "sessexp,"
end if

n = 0
For Each x In Request.form
  if n = 0 then
    strsql = strsql & x
    n = 1
  else
    strsql = strsql & "," & x
  end if
Next

if Tabla = "Usuarios" then
   strsql = strsql & ",Grupo"
   strsql = strsql & ",netusuario"
end if

strsql = strsql & ") Values ("

if Tabla = "Usuarios" then
   strsql = strsql & "'" & sEncript & "',"
   strsql = strsql & "8,"
   strsql = strsql & "1,'"
   strsql = strsql & xlcha & "',"
   strsql = strsql & "0,"
   strsql = strsql & "0,"
   strsql = strsql & cint(rx("passmin")) & ","
   strsql = strsql & cint(rx("passmax")) & ","
   strsql = strsql & cint(rx("tipoesq")) & ","
   strsql = strsql & cint(rx("passhist")) & ","
   strsql = strsql & cint(rx("intfall")) & ","
   strsql = strsql & cint(rx("expira")) & ","
   strsql = strsql & cint(rx("diasexp")) & ",'"
   strsql = strsql & xhcha & "',"
   strsql = strsql & "1,"
end if

if Ucase(tabla) = "GRUPOS" then
   strsql = strsql & "'D',"
end if

n = 0
For Each x In Request.form
'  response.write "x " & x
  TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
  if TipoValido = False then exit for
  ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
  if ValorNulo = True then exit for
  s = ofv.Queesx(rs(x).type)
  if rs(x).type > 199 AND rs(x).type < 300 THEN
     s = "'"
  end if
  s = "'"
  if n = 0 then
    strsql = strsql & s & UCase(Request.Form(x)) & s
    n = 1
  else
    strsql = strsql & "," & s & UCase(Request.Form(x)) & s
  end if

  if ucase(x) = "ID" THEN NETUSUARIO = UCase(Request.Form(x))

Next

if TipoValido = False or ValorNulo = True then
  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & Request.Form(x)
else

   if Tabla = "Usuarios" then 
      strsql = strsql & ",'DEF'"
      strsql = strsql & ",'" & NETUSUARIO & "'"
   END IF
  strsql = strsql & ")"
'response.write strsql
  call ofv.crearconsultaEx(StrSql,cn,1,volver)

  if cn.errors.count = 0 then
      if Ucase(tabla) =  "USUARIOS" then session("Nusuario") = sID
      response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if
end if
end if
if Tabla = "Usuarios" then
   call ofv.cerrarconsulta(rx)
end if
call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>