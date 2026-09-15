<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  if ofv.Explorador = "MSIE" then
    sHTML = ofv.MenuHeader("#cccccc","")
  else
    sHTML = ofv.MenuHeader("#ccbbaa","")
  end if

  CantRegMostrarx = 11
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  opcion = request.querystring("opcion")
  if opcion = "" then
     opcion = session("opcion")
  else
     session("opcion") = opcion
  end if
 
  DSN = request.querystring("DSN")
  if DSN = "" then
     DSN = session("DSN")
  else
     session("DSN") = DSN
  end if

  uDSN = request.querystring("uDSN")
  if uDSN = "" then
     uDSN = session("uDSN")
  else
     session("uDSN") = uDSN
  end if

  query = request.form("query")
  if query = "" then
     query = session("query")
  else
     session("query") = query
  end if

  if mid(opcion,1,1) <> "2" then
     opcion = "2" & opcion
  end if
     
  volver = "../querys.asp"
  volver2 = "querys.asp?opcion=" & opcion & "&DSN=" & DSN
 


set cn = ofv.conectar(uDSN)

'on error resume next
set rs = ofv.crearconsultaEx(query,cn,1,parametros)
if cn.errors.count > 0 then
  set errores = cn.errors.item(0)
  sDescripcion = errores.description
  sDescripcion = convertircar(sDescripcion," ","_")
  nNumero = errores.number
  volver = volver & "&Descripcion=" & sDescripcion & "&Numero=" & nNumero
  response.redirect "asp/ErrordbEx.asp?volver=" & volver
  cn.errors.clear
end if 



sHTML = sHTML & "<center><table width=70% ><tr><td CLASS=CC align=center>RESULTADOS DEL QUERY SOBRE CONECCION: " & udsn & "</td></tr>"
sHTML = sHTML & "</table>"
if Not rs.EOF then 
   rs.move(CantRegMoverx)
   sHTML = sHTML & "<table width=100% border bgcolor=#f7eedd >"
   sHTML = sHTML & "<tr>"
   for each x in rs.fields
       sHTML = sHTML & "<td class=tt width=10% >" & x.name & "</td>"
   next
   sHTML = sHTML & "</tr>"
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    posi = posi + 1
   sHTML = sHTML & "<tr>"
   for each x in rs.fields
       sHTML = sHTML & "<td class=ut  width=10% >" & x & "</td>"
   next
   sHTML = sHTML & "</tr>"
   rs.movenext
   loop
   sHTML = sHTML & "</table>"
end if

otra = "<form name=Tablas action=" & volver2
otra = otra & " method=post>"
otra = otra & "<td align=center><input class=bt type=submit value=Otra Consulta></td>"
otra = otra & "</form>"


sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",otra,"","")

'sHTML = sHTML & "<table width=70% >"
'sHTML = sHTML & "<form name=Tablas action=" & volver2
'sHTML = sHTML & " method=post><tr><td align=center><BR></td></tr>"
'sHTML = sHTML & "<tr><td align=center><input class=bt type=submit value=Otra Consulta></td></tr>"
'sHTML = sHTML & "</form>"
'sHTML = sHTML & "</table></center></body></html>"

sHTML = sHTML & "</center></body></html>"
call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

end if

  response.write sHTML



%>