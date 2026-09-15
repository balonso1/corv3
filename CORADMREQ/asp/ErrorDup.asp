<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()

volver = request.querystring("volver")
tabla = request.querystring("tabla")
sID = request.querystring("ID")

sHTML = "<html>" & ofv.FormStyler() & "<body bgcolor=#eeeeee><br><br><p align=center>"
sHTML = sHTML & "<font color=#000077 face=verdana size=2>Cor ESolution Suite - Error de duplicacion</p>"
sHTML = sHTML & "<br><br><br><br><p align=center><b>"
sHTML = sHTML & "El codigo </font><font color=#770000 face=verdana size=3>"
sHTML = sHTML & Ucase(sID) & " </font>"
sHTML = sHTML & "<font color=#000077 face=verdana size=2>ya existe en "
sHTML = sHTML &  tabla & "</font></b></p>"
sHTML = sHTML & "<form action=" & volver & " method=post><table width=100% marginwidth=0 "
sHTML = sHTML & "cellspacing=" & wspacing & " cellpadding=" & 1 + wspacing & "><tr>"
sHTML = sHTML & "<td width=100% align=center><input type=submit class=bt value=Volver>"
sHTML = sHTML & "</td></tr></table></form></body></html>"

response.write sHTML

%>