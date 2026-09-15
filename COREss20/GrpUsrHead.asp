<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

'  if ofv.explorador = "MSIE" then
'    sHTML = ofv.MenuHeader("#cccccc","")
'  else
'    sHTML = ofv.MenuHeader("#ccbbaa","")
'  end if

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  vuelta = request.querystring("volver")
  nid = request.querystring("opcion")
  t = request.querystring("t")

  set cn = ofv.conectar(ofv.strconn0)

tabla = "Usuarios"

if t = "G" then tabla = "Grupos"

strsql = "Select descripcion From " & Tabla & " where ID ='" & nid & "' "
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then

  celdas = ofv.GenCelda("","tt colspan=5","","","","","","USUARIO","si")
if t = "G" then 
  celdas = ofv.GenCelda("","tt colspan=5","","","","","","GRUPO","si")
END IF

  filas  = ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","cc","","10%","","","",nid,"si")
  celdas = celdas & ofv.GenCelda("","","","2","","","","","si")
  celdas = celdas & ofv.GenCelda("","cc","","80%","","","",rs(0),"si")
  celdas = celdas & ofv.GenCelda("","","","2","","","","","si")
  sInput = ofv.GenerarInput("","Volver","submit","","10","bt","")
  sInput = ofv.GenCelda("","","","","","","",sInput,"si")
  sAccion = ofv.convertircar(vuelta,"_","&")
  celdas = celdas & ofv.GenForm("volver","",sAccion,"_parent","post","",sInput,"si")
  filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML & "<center>"
  sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","80%","","0","","0","0","","",filas,"si")
  sHTML = sHTML & "</center></body></html>"
end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>