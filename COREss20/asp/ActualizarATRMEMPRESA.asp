<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

nID = request.querystring("ID")
t = request.querystring("t")
Nombre = request.querystring("Nombre")
sOpcion = request.querystring("Opcion")
sGrupoUsuario = request.querystring("grupousuario")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

'response.write "gu " & sGrupoUsuario & " ID " & nid 

xdefault = request.form("default")
xtipo = "E"

dim aatrm(9)
for xind=0 to 9
    aatrm(xind) = 1
next
xind = 0
xatrm = nid
xtab = "S"

set cn = ofv.conectar(ofv.strconn0)

'strsql = "select atributo from atrm where id = " & nID 
'set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'if not rsa.eof then xatrm = rsa(0)
'call ofv.cerrarconsulta(rsa)

  stabla = "ATRMEMP"



strsql = "SELECT * From ATRMEMP Where ATRMTIPO = '" & T & "' AND ATRMUSRGRP = '" & sGrupoUsuario & "' "
strsql = strsql & " And ATRMCODEMP = " & NID
set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros) 
if rsa.eof then xtab = "N"
call ofv.cerrarconsulta(rsa)

xtipo = "A"





       aatrm(0) = request.form("CONSULTAR")
       aatrm(1) = request.form("CONSULTARD")
       aatrm(2) = request.form("ACTUALIZAR")
       aatrm(3) = request.form("ACTUALIZARD")


       if aatrm(0) = aatrm(1) and aatrm(2) = aatrm(3) then
          if xtab <> "N" then
             strsql = "delete From ATRMEMP Where ATRMTIPO = '" & T & "' AND ATRMUSRGRP = '" & sGrupoUsuario & "' "
             strsql = strsql  & " And ATRMCODEMP = " & NID
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if
       else
          if xtab = "N" then
             strsql = "insert into ATRMEMP (ATRMTIPO,ATRMUSRGRP,ATRMCODEMP,ATRMVER,"
             strsql = strsql  & "ATRMACT) values("
             strsql = strsql  & "'" & T & "','" & sGrupoUsuario & "','" & NID & "'"
             strsql = strsql  & "," & aatrm(0) & "," & aatrm(2) & ")"
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          else
             strsql = "update ATRMEMP set "
             strsql = strsql & " ATRMVER = " & aatrm(0) & ","
             strsql = strsql & " ATRMACT = " & aatrm(2)
             strsql = strsql & " Where ATRMTIPO = '" & T & "' AND ATRMUSRGRP = '" & sGrupoUsuario & "' "
             strsql = strsql & " And ATRMCODEMP = " & NID
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if  
       end if


 '  RESPONSE.WRITE STRSQL



if cn.errors.count = 0 then
  volver = volver & "&Opcion=" & sGrupoUsuario & "&t=" & t & "&Nombre=" & Nombre
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if 

call ofv.cerrarconn(cn)

%>