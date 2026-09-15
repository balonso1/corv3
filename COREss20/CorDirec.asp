<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader("Directorio de Descargas")


  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  xuser = Request.ServerVariables("LOGON_USER")
  if len(xuser) > 0 then xuser = xuser & "@"
  dir = "ftp://" & xuser & Ucase(Request.ServerVariables("SERVER_NAME")) & "/CorDownload"
  dircmp = server.htmlencode(dir) 
 

    Celda = ofv.GenCelda("","tt","","","","","","Directorio:","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "<a href=" & chr(34) & dirCMP & chr(34) & " target=_blank ><img src=iconos/ICorAbre.jpg border=0 align=center><font size=2 color=#993333 face=arial,tahoma>&nbsp;&nbsp;Descargas</font></a>"
    end if   

    Celda = ofv.GenCelda("","cc","","","","","",saccion,"si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






  sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","","","80%","","0","","5","5","#eeeeee","",Fila,"si")



  sHTML = sHTML & "</center></body></html>"
end if

Response.Write sHTML

%>