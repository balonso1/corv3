
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FunDb.inc" -->
<!-- #INCLUDE FILE="IncFile/Funvarias.inc" -->
<%
response.expires=0

if request.querystring("Xopc") <> "" then 
Xopc = request.querystring("Xopc")
else 
Xopc = "X"
end if

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

set cn=ofv.conectar(ofv.strconn0)

sHTML = "<html><head>"

shtml = shtml & "<style>"
shtml = shtml & "td.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;font-weight:bold;border-style:outset;border-width:1;}"
shtml = shtml & "td.to {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;font-weight:bold;}"
shtml = shtml & "td.ot {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
shtml = shtml & "td.dt {font-family:tahoma;font-size:" & wfont & "pt;color:#f76655;}"
shtml = shtml & "tr.tt {background-color:#000077;}"
sHTML = sHTML & "select.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "font.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;}"
sHTML = sHTML & "input.dt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
sHTML = sHTML & "input.bt{font-family:tahoma;color:#ffffff;font-size:" & wfont & "pt;background-color:#336699;}"
shtml = shtml & "</style>"

shtml = shtml & "</head>"

Strsql="Select ID, descripcion From usuarios Order By descripcion"
set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
n = 0
If rs.Eof then 
  shtml = shtml & "<body bgcolor=#6699cc topmargin=3 bottommargin=0 >"
  sHTML = sHTML & "<p align=center ><font face=tahoma size=1 color=#990000><b>NO HAY USUARIOS DISPONIBLES</b><br><br></FONT></P>"
  call ofv.cerrarconsulta(rs)
Else  

Do until rs.Eof
  Redim Preserve usuario(n)
  Redim Preserve nombre(n)
  usuario(n) = rs(0)
  nombre(n) = rs(1)
  rs.Movenext
  n = n + 1
Loop
call ofv.cerrarconsulta(rs)
end if
  shtml = shtml & "<body bgcolor=#6699cc topmargin=3 bottommargin=0 onload='document.msgsending.MsgData.focus();' >"
sHTML = sHTML & "<table id=mast width=100% border=1 marginwidth=0 marginheight=0   cellspacing=0 cellpadding=0 >"
sHTML = sHTML & "<tr><td width=18% align=center >"

sHTML = sHTML & "<table id=logo width=100% marginwidth=0 marginheight=0   cellspacing=0 cellpadding=0 background=../CorEss20/imagenes/topban.jpg style='border-style:inset;border-width:1;'>"
sHTML = sHTML & "<tr>"

sHTML = sHTML & "<td height=10 align=left width=20% style='border-style:outset;border-width:1;'>"
sHTML = sHTML & " <img src=../CorEss20/imagenes/COREssuite20.jpg ></td>"
sHTML = sHTML & "<td height=20 align=center  width=80%  class=to >"
sHTML = sHTML & "&nbsp;COR&nbsp;Chat&nbsp;2.8</td>"

sHTML = sHTML & "</tr>"
sHTML = sHTML & "</table>" & chr(13)

sHTML = sHTML & "</td><td width=82% align=center >"

sHTML = sHTML & "<table id=margen1 width=95% marginwidth=0   cellspacing=" & wspacing & " cellpadding=0 >"
sHTML = sHTML & "<form name=msgsending action=Msgto.asp?Opc=X  target=Xpanel  method=post>"
sHTML = sHTML & "<tr>"

sHTML = sHTML & "<td class=tt height=10 width=10% bgcolor=#000080 >"
sHTML = sHTML & "&nbsp;Canal</td>"

  sHTML = sHTML & "<td height=20 align=left width=20% > <font class=at> "
  sHTML = sHTML & "<select class=at  name=Msgcanal "
  sHTML = sHTML & "onchange=" & chr(34) & "if (this.selectedIndex == 0) {document.msgsending.MsgDestino.focus();} else {document.msgsending.MsgData.focus();};" & chr(34) & " >"

  if Xopc = "Publico" then 
     sHTML = sHTML & "<option value=Privado >Privado "
     sHTML = sHTML & "<option value=Publico selected>Publico "
  else
     sHTML = sHTML & "<option value=Privado selected>Privado "
     sHTML = sHTML & "<option value=Publico >Publico "
  end if

  sHTML = sHTML & "</select></font></td>"


sHTML = sHTML & "<td class=tt  height=10 width=10% bgcolor=#000080 >"
sHTML = sHTML & "&nbsp;Destino</td>"

  sHTML = sHTML & "<td height=20 align=left width=55% > <font class=at> "
  sHTML = sHTML & "<select class=at  name=MsgDestino onchange='document.msgsending.MsgData.focus();' >"

  For n = 0 To Ubound(usuario)
    If Ucase(usuario(n)) = Ucase(Xopc) Then
      sHTML = sHTML & "<option value=" & usuario(n) & " selected>" & nombre(n)
    Else
      sHTML = sHTML & "<option value=" & usuario(n) & ">" & nombre(n)
    End If
  Next
  sHTML = sHTML & "</select></font></td>"

  shtml = shtml & "<td height=20 align=left rowspan=2 width=5%  >" & chr(13)
  shtml = shtml & "<input class=bt type=submit size=10% value='Enviar' "
  shtml = shtml & "  ></td>" & chr(13)

sHTML = sHTML & "</tr>"

  sHTML =  sHTML & "<tr>" & chr(13)
sHTML = sHTML & "<td class=tt  height=10  bgcolor=#000080 >"
sHTML = sHTML & "&nbsp;Mensaje</td>"

  sHTML = sHTML & "<td height=20 align=left colspan=3>"
  sHTML = sHTML & "<input class=dt   type=text size=" & (110 / wsize) - wresto & "% name=MsgData  ></td>"




  sHTML = sHTML & "</tr></form>"

 

sHTML = sHTML & "</table>" & chr(13)

sHTML = sHTML & "</td></tr>"
sHTML = sHTML & "</table>" & chr(13)
 
response.write shtml & " </body></html>"

call ofv.cerrarconn(cn)
 
%>

