<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

if session("USRES") <> "B" then

  shtml = ofv.MenuHeader(""," bacKground='imagenes/corf5.jpg' ")
else
  shtml = ofv.MenuHeader("","")
end if



  set cn = ofv.conectar(ofv.strconn1)

  Strsql = "Select * from CorConsulta"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then
     xLicencia = rs(2)
     xEmpresa = rs(3)
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "<center><br><br><br>"

  sHTML = sHTML & "<table    width=650 marginwidth=4  cellspacing=2 cellpadding=2>"

  sHTML = sHTML & "<tr ><td align=center style='border-style:outset;border-width:1;' "
  sHTML = sHTML & " valign=middle >"
  sHTML = sHTML & "<img src=imagenes/COREssuite20.jpg border=0 "
  sHTML = sHTML & " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td   width=35%   valign=middle >"
  sHTML = sHTML & "<font face=verdana color=#000088 size=2 ><b>"
  sHTML = sHTML & "COR <font face=helvetica color=#f75500 size=4><i>e</i></font>Solution Suite 2.8</b></font>" 
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td   width=65%  valign=top >"
  sHTML = sHTML & "<font face=verdana color=#000088 size=1 ><b><br>Licencia:&nbsp;&nbsp;&nbsp;"
  sHTML = sHTML & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>" & xLicencia
  sHTML = sHTML & "</strong></b></font></td></tr>"

  sHTML = sHTML & "<tr><td align=center >&nbsp;</td><td align=center >&nbsp;</td>"
  sHTML = sHTML & "<td   width=65% ><font face=verdana color=#000088 size=1 ><b>"
  sHTML = sHTML & "Otorgada a:&nbsp; <strong>" & xEmpresa 
  sHTML = sHTML & "</strong><br></b></font></td>" 

  sHTML = sHTML & "</tr></table><br><br><br><br><br><br><br><br><br><br><br><hr size=2 width=100% color=#993333 >"





 xAviso = "Cor ESolution Suite&reg; es un producto registrado por Cor Consulting Asociados SRL."  
 xAviso = xAviso & "<br> Cor ESolution Suite&reg; y todos sus componentes son propiedad de Cor Consulting Asociados SRL."  
 xAviso = xAviso & "<br> &copy;1999 - 2005"



    sHTML = sHTML & "<table width=650 marginwidth=2 cellspacing=2 cellpadding=2>"
    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td width=35% align=left rowspan=2><br><p align=center>"
    sHTML = sHTML & "<img src=imagenes/CorLogox1.jpg border=0 width=22 height=22>"
    sHTML = sHTML & "<br>"
    sHTML = sHTML & "<font face=helvetica,tahoma color=#000099 size=2 ><b>"
    sHTML = sHTML & "Cor Consulting<br>Asociados</b></font>" 
    sHTML = sHTML & "</p></td>"
    sHTML = sHTML & "<td  width=65% align=left colspan=2><font color=#336699>"
    sHTML = sHTML & xaviso & "</font></td>"
    sHTML = sHTML & "</tr>"
    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td  align=left >"
    sHTML = sHTML & "<img src=imagenes/reparg.jpg border=0  ></td>"
    sHTML = sHTML & "<td width=65% align=left >"
    sHTML = sHTML & "<font color=#336699>&nbsp;Producto desarrollado en Argentina</font>"
    sHTML = sHTML & "</td></tr>"
    sHTML = sHTML & "</table>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>