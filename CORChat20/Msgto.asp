
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FunDb.inc" -->
<!-- #INCLUDE FILE="IncFile/Funvarias.inc" -->
<%
response.expires=0


Xopc = request.querystring("Opc")

if Xopc <> "X" then
 
shtml = "<html><head>"
shtml = shtml & "<style>"
shtml = shtml & "p.tt {font-family:tahoma;font-size:8pt;color:#cccccc;}"
shtml = shtml & "font.tt {font-family:tahoma;font-size:8pt;color:#ff7011;}"
shtml = shtml & "</style></head>"
shtml = shtml & "<body bgcolor=#000080 text=#ffffff topmargin=0 bottommargin=0 style='border-style:outset;border-width:1;' >"
shtml = shtml & "<p  class=tt width=100% >Usuario: <font class=tt>" & Session("SegUsuario") & "</font> - Bienvenido a COR Chat 2.8  </p> "
shtml = shtml & "</body></html>"

else

Xmensaje = request.form("MsgData")
Xdestino = request.form("MsgDestino")
Xcanal = request.form("Msgcanal")
Xusuario = Session("Usuario")
Xusrdesc = Session("SegUsuario")

if Xcanal = "Publico" then
Xdestino = Xcanal
Xdestinodesc = Xcanal
Xusrdesc = Xcanal

else
set cn=ofv.conectar(ofv.strconn0)

if Xcanal = "Grupo" then
   Strsql="Select ID, descripcion From grupos where ID ='" & Xdestino & "' "
   set rs=crearconsultaEx(StrSql,cn,1,parametros)
   Xdestino = "GRP-" & Xdestino 
else
   Strsql="Select ID, descripcion From usuarios where ID ='" & Xdestino & "' "
   set rs=crearconsultaEx(StrSql,cn,1,parametros)
end if

If rs.Eof then 
  call ofv.cerrarconsulta(rs)
Else  

Xdestinodesc = rs(1)
call ofv.cerrarconsulta(rs)
end if

call ofv.cerrarconn(cn)
end if


set cn=ofv.conectar(ofv.strconn3)

YC = 1
YD = 0
Xmarca = 0

Strsql= " select * from CsMsgSearch where Msgusrid ='" & Xdestino & "' order by msgid DESC"

set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
Do until rs.Eof
   
if YC > 14 then
  Redim Preserve Xmsgid(YD)
  Xmsgid(YD) = rs(0)
  YD = YD + 1
  Xmarca = 1
end if

  rs.Movenext
  YC = YC + 1
  
Loop
call ofv.cerrarconsulta(rs)

if Xmarca = 1 then
   for YC = 0 to Ubound(Xmsgid)
   Strsql= " delete  from MsgChat where Msgid =" & Xmsgid(YC)    
'response.write strsql
   set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
   next
end if 



Strsql="insert into MsgChat (Msgusrid,Msgusrdesc,Msgtstamp,Msgdata,Msgusrorigid,Msgusroridesc) Values('" 

Strsql= Strsql & Xdestino & "','" & Xdestinodesc & "','" & now  & "','" & Xmensaje & "','" & Xusuario & "','" & Xusrdesc & "') "

' response.write strsql

set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

shtml = "<html><head>"
shtml = shtml & "<style>"
shtml = shtml & "td.tt {font-family:tahoma;font-size:8pt;color:#f78800;}"
shtml = shtml & "td.dt {font-family:tahoma;font-size:8pt;color:#eeeeee;}"
shtml = shtml & "</style></head>"

shtml = shtml & "<body bgcolor=#000080 text=#f78800 topmargin=0 bottommargin=0  onload=document.msgreset.submit(); style='border-style:outset;border-width:1;' >"



shtml = shtml & "<table><tr><td class=tt >" & now & "</td><td class=dt > Mensaje a: " & xdestinodesc & " - </td><td class=tt >" & Xmensaje & "</td></tr>"
shtml = shtml & "</table>"

 sHTML = sHTML & "<form name=msgreset action=MsgSend.asp?Xopc=" & xdestino & " target=Xcontrol  method=post>"
 shtml = shtml & "<input type=hidden size=" & wsize & " name=boton value=menu></input>" & chr(13)
 sHTML = sHTML & "</form>"
call ofv.cerrarconn(cn)

end if



response.write shtml & " </body></html>"

 
%>

