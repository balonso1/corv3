
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FunDb.inc" -->
<!-- #INCLUDE FILE="IncFile/Funvarias.inc" -->
<%
response.expires=0


wbrow = explorador()
if wbrow = "MSIE" then
   wfont = 7
   wspacing = 1
   wsize = 1
   wresto = 0
else
   wfont = 8
   wspacing = 0
   wsize = 2
   wresto = 3
end if

set cn=ofv.conectar(ofv.strconn3)



Strsql="select * from MsgChat where msgusrid = '" & Ucase(Session("usuario")) & "' "
Strsql=Strsql & " or " & "msgusrid = 'Publico' " 
Strsql=Strsql & " or " & "msgusrid = 'GRP-" & Session("Seggrp") & "' " 
Strsql=Strsql & " order by msgid DESC" 



set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)


shtml = "<html><HEAD><title>Cor Chat</title><META http-equiv='refresh' content='10' >"

shtml = shtml & "<style>"
shtml = shtml & "p.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;background-color:#114477;}"
shtml = shtml & "td.to {font-family:tahoma;font-size:" & wfont & "pt;color:#336699;background-color:#f7eedd;border-style:outset;border-width:2;}"
shtml = shtml & "td.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffffff;}"
shtml = shtml & "td.ot {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
shtml = shtml & "td.gt {font-family:tahoma;font-size:" & wfont & "pt;color:#f7eedd;}"
shtml = shtml & "td.dt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
shtml = shtml & "tr.tt {background-color:#000080;}"
shtml = shtml & "</style>"

shtml = shtml & "</head>"
shtml = shtml & "<body bgcolor=#6699cc text=#dddddd link=#ffffff alink=#ffffff vlink=#ffffff topmargin=0 bottommargin=0 >"

if rs.eof then 

shtml = shtml & "<p class=tt width=" & (100 / wsize) - wresto & "% >Usuario: " & Session("SegUsuario") & " - Ud. no ha recibido mensajes. </p> "

else

Xmarca = 0



shtml = shtml & "<table width=100% marginwidth=0 marginheight=0   cellspacing=" & wspacing & " cellpadding=2 >"
shtml = shtml & "<tr><td class=to colspan=2 >&nbsp;<b>Usuario: " & Session("Usuario") & " - " & Session("SegUsuario") & " </b></td><td class=to align=right><b>Ultimo refresco: " & now & "</b>&nbsp;</td></tr> "
do while not rs.eof

Xclase = "class=tt "

if rs(6) = "Publico" then 
Xclase = "class=ot "
elseif instr(rs(1),"GRP-") > 0 then 
Xclase = "class=gt "
end if

if Xmarca = 0 then
   Xmarca = 1
   shtml = shtml & "<tr class=tt >"
else
   shtml = shtml & "<tr>"
end if 

shtml = shtml & "<font face=tahoma size=1 color=#f788ee >"
shtml = shtml & "<td " & Xclase & " width=100 >" & rs(3) & " - </td>"

if rs(6) = "Publico" then
shtml = shtml & "<td " & Xclase & " width=150 ><b>" & Ucase(rs(6)) & "</b> - </td>"
else
shtml = shtml & "<td " & Xclase & " width=150 ><a href=MsgSend.asp?Xopc=" & rs(5) & " target=Xcontrol ><b>"& rs(6) & "</b></a>&nbsp;</td>"
end if

shtml = shtml & "<td " & Xclase & " width=450 >" & rs(4) & "</td></tr>"

rs.movenext
loop

shtml = shtml & "</table>"
shtml = shtml & "</font>"

end if

shtml = shtml & "</body></html>"


 
response.write shtml 


call ofv.cerrarconn(cn)
 
%>

