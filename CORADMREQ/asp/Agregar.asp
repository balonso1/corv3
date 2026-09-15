<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

Tabla = request.querystring("Tabla")
volver = request.querystring("volver")

set cn = ofv.conectar(ofv.strconn5)


strsql = "Select * From " & Tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

strsql = "Insert Into " & Tabla & " ("
   strsql = strsql & "tipo,"
   strsql = strsql & "nivel,"
   strsql = strsql & "ESTADO,"
   strsql = strsql & "ESTANT,"
   strsql = strsql & "CALIF,"
   strsql = strsql & "CALANT,"
   strsql = strsql & "RESPONSABLE,"
   strsql = strsql & "FECHAVIG,"


n = 0
For Each x In Request.form
  if n = 0 then
    strsql = strsql & x
    n = 1
  else
    strsql = strsql & "," & x
  end if
Next


strsql = strsql & ") Values ("


   strsql = strsql & "1,"
   strsql = strsql & "8,"
   strsql = strsql & "2,"
   strsql = strsql & "1,"
   strsql = strsql & "5,"
   strsql = strsql & "5,"
   strsql = strsql & "'" & SESSION("SEGPERFIL") & "',"

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

   strsql = strsql & "'" & vNuevoValor & "',"



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
  if n = 0 then
    strsql = strsql & s & UCase(Request.Form(x)) & s
    n = 1
  else
    strsql = strsql & "," & s & UCase(Request.Form(x)) & s
  end if
Next

if TipoValido = False or ValorNulo = True then
  response.redirect "errortipo.asp?volver=" & volver & "&Columna=" & rs(x).name & "&Dato=" & Request.Form(x)
else


  strsql = strsql & ")"
'response.write strsql
  call ofv.crearconsultaEx(StrSql,cn,1,volver)

  if cn.errors.count = 0 then
      sID = trim(request.form("codigo"))
      session("NCliente") = sID
      response.redirect volver
  else
    response.write ofv.ErrorDB(cn,volver)
  end if 
end if


call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>