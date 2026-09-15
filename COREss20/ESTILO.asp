<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

    shtml = "<html>" & CorStyler2()
    sHTML = sHTML & "<body>" & corpaneltop()
    sHTML = sHTML & "<br><p>" & Session("CorApli") & "</p>"
    sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>