<!-- #INCLUDE FILE="IncFile/funcvarias.inc" -->

<%


set ofv = new funcvarias


Set ofv.Request = Request
Set ofv.Server = Server
Set ofv.Response = Response
Set ofv.Session = Session

'response.write " HTTP_X_FORWARDED_HOST: " & request.servervariableS("HTTP_X_FORWARDED_HOST") & "<BR>"

' response.write " remote usr: " & request.servervariableS("REMOTE_USER") & "<BR>"
'response.write " AUTH usr: " & request.servervariableS("AUTH_USER") & "<BR>"

'response.write " remote ADDR: " & request.servervariableS("REMOTE_ADDR") & "<BR>"

'response.write " ALL HTTP: " & request.servervariableS("ALL_HTTP") & "<BR>"


for each variable in request.servervariables 
    response.write "<br> " & variable & ": " 
    response.write request.servervariables(variable) 
next


%>