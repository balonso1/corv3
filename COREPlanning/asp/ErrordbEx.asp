<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

volver = request.querystring("volver")
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
if nCodigo <> "" then volver = volver & Codigo & nCodigo
if sTipo <> "" then volver = volver & Tipo & sTipo

sHTML = "<html><body><br><br><p align=center>"
sHTML = sHTML & "<font color=#000077 face=verdana size=2>Error: "
sHTML = sHTML & request.querystring("Numero") & " - "
sHTML = sHTML & ofv.convertircar(request.querystring("Descripcion"),"_"," ")
sHTML = sHTML & "</p></font><br><br><br><br>"
sHTML = sHTML & "<p align=center><a onclick=window.close() "
sHTML = sHTML & "href=" & volver & ">volver</a></p>"
sHTML = sHTML & "</body></html>"

response.write sHTML

%>