<%
  nid = request.querystring("opcion")
  t = request.querystring("t")
  o = request.querystring("o")
  volver = request.querystring("volver")

sHTML = "<html><head><title>COR ESolution Suite </title></head>"
sHTML = sHTML & "<frameset rows=20%,* FRAMESPACING=0 FRAMEBORDER=0 BORDER=0>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Cabecera SCROLLING=auto  "

sHTML = sHTML & "src=GrpUsrHead.asp?opcion=" & nid & "&t=" & t & "&volver=" & volver & " >"

sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Contenido "

if o = "A" then
   sHTML = sHTML & "src=Aplicaciones.asp?Opcion=" & nID & "&t=" & t & ">"
else
   sHTML = sHTML & "src=UsrSkrdef.asp?Opcion=" & nID & "&t=" & t & ">"
end if

sHTML = sHTML & "</frameset><noframes><body>"
sHTML = sHTML & "<p>Esta página usa marcos, pero su explorador no los admite.</p>"
sHTML = sHTML & "</body></noframes></html>"

response.write sHTML

%>