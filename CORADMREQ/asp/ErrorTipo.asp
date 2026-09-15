<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()

volver = request.querystring("volver")
nProced = request.querystring("Codproced")
nAccion = request.querystring("Accion")
sNombre = request.querystring("Nombre")
if nAccion <> "" then volver = volver & "&Accion=" & nAccion
if nProced <> "" then volver = volver & "&Codproced=" & nProced
if sNombre <> "" then volver = volver & "&Nombre=" & sNombre
Dato = request.querystring("Dato")
if Dato = "nsnc" then Dato = ""
Columna = request.querystring("Columna")
Columna = ofv.DescCampo(Columna)
nCodigo = request.querystring("Codobjeto")
Codigo = "&Codobjeto="
if nCodigo = "" then
  nCodigo = request.querystring("Codcomp")
  Codigo = "&Codcomp="
end if
sTipo = request.querystring("Tipobjeto")
Tipo = "&Tipobjeto="
if sTipo = "" then
  sTipo = request.querystring("Tipocomp")
  Tipo = "&Tipocomp="
end if

n = instr(volver,"?")
if n > 0 then volver = volver & Codigo & nCodigo & Tipo & sTipo

sHTML = "<html>" & ofv.FormStyler() & "<body><br><br><p align=center>"
sHTML = sHTML & "<font color=#000077 face=verdana size=2>Cor ESolution Suite - Error de datos</p>"
sHTML = sHTML & "<br><br><br><br><p align=center>"
if Dato = "" then
  sHTML = sHTML & "</font><font color=#770000 face=verdana size=3><b>" & Columna
  sHTML = sHTML & "</font></b><font color=#000077 face=verdana size=2>"
  sHTML = sHTML & " no puede ser nulo </font></p>"
else
  sHTML = sHTML & "El tipo de dato </font><b><font color=#770000 face=verdana size=3>"
  sHTML = sHTML & Dato & " no es v&aacute;lido para </font><b>"
  sHTML = sHTML & "<font color=#000077 face=verdana size=2>" & Columna & "</font></b></p>"
end if
sHTML = sHTML & "<form action=" & volver & " method=post><table width=100% marginwidth=0 "
sHTML = sHTML & "cellspacing=" & wspacing & " cellpadding=" & 1 + wspacing & "><tr>"
sHTML = sHTML & "<td width=100% align=center><input type=submit class=bt value=Volver>"
sHTML = sHTML & "</td></tr></table></form></body></html>"

response.write sHTML

%>