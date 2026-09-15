<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  if ofv.Explorador = "MSIE" then
    sHTML = ofv.MenuHeader("#cccccc","")
  else
    sHTML = ofv.MenuHeader("#ccbbaa","")
  end if

  opcion = request.querystring("opcion")

  if opcion = "" then


sHTML = sHTML & "<center><table width=70% ><tr><td CLASS=CC align=center>CONECCION</td></tr>"
sHTML = sHTML & "<form name=Tablas action=Querys.asp?opcion=D method=post>"
sHTML = sHTML & "<tr><td align=center><BR></td></tr>"
sHTML = sHTML & "<tr><td align=center><select class=at name=DSN>"
sHTML = sHTML & "<option value=strconn0>" & ofv.strconn0
sHTML = sHTML & "<option value=strconn1>" & ofv.strconn1
sHTML = sHTML & "<option value=strconn2>" & ofv.strconn2
sHTML = sHTML & "<option value=strconn3>" & ofv.strconn3
sHTML = sHTML & "<option value=strconn4>" & ofv.strconn4
sHTML = sHTML & "</select></td></tr>"
sHTML = sHTML & "<tr><td align=center><input class=bt type=submit value=Conectar></td></tr>"
sHTML = sHTML & "</form>"
sHTML = sHTML & "<form name=Tablas2 action=Querys.asp?opcion=O method=post>"
sHTML = sHTML & "<tr><td align=center><BR></td></tr>"
sHTML = sHTML & "<tr><td align=center><textarea class=at name=DSN cols=50 rows=20></textarea>"
sHTML = sHTML & "</td></tr>"
sHTML = sHTML & "<tr><td align=center><input class=bt type=submit value=Conectar></td></tr>"
sHTML = sHTML & "</form>"
sHTML = sHTML & "</table></center></body></html>"
else
   if opcion = "D" or opcion = "O" then
      xdsn = request.form("DSN")
   else
      xdsn = request.querystring("DSN")
   end if

   if opcion = "D" or opcion = "2D" then
      select case xdsn
             case "strconn0"
                  udsn = ofv.strconn0
             case "strconn1"
                  udsn = ofv.strconn1
             case "strconn2"
                  udsn = ofv.strconn2
             case "strconn3"
                  udsn = ofv.strconn3
             case "strconn4"
                  udsn = ofv.strconn4
      end select
   else
            udsn = xdsn
   end if 
sHTML = sHTML & "<center><table width=70% ><tr><td CLASS=CC align=center>QUERY SOBRE CONECCION: " & udsn & "</td></tr>"
sHTML = sHTML & "<form name=Tablas action=tablas.asp?opcion=" & opcion & "&DSN=" & xdsn & "&uDSN=" & udsn & "&cantregmover=0"
sHTML = sHTML & " method=post><tr><td align=center><BR></td></tr>"
sHTML = sHTML & "<tr><td align=center><textarea class=at name=query cols=50 rows=20>"
sHTML = sHTML & "</textarea></td></tr>"
sHTML = sHTML & "<tr><td align=center><input class=bt type=submit value=Ejecutar></td></tr>"
sHTML = sHTML & "</form>"
sHTML = sHTML & "</table></center></body></html>"

end if

end if

response.write sHTML

%>
