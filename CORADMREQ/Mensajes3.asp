<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
sVolver = request.querystring("volver")
nCampo = request.querystring("Campo")
sNombre = request.querystring("Nombre")
nAccion = request.querystring("Accion")
nProced = request.querystring("Codproced")
vuelta = request.querystring("vuelta")
nCodigo = request.querystring("Codobjeto")
sTipo = request.querystring("Tipobjeto")
Tabla = request.querystring("Tabla")
nBotones = CInt(request.querystring("Botones"))
'sID = request.querystring("ID")

sHTML = ofv.FormHeader("Atencion")

    sVolver = sVolver & "&Tabla=" & Tabla & "&Campo=" & nCampo

sInput = "<font color=#336699 face=verdana size=1><b>" & sMsg & "</b></font>"
Celda = ofv.GenCelda("","","center","","","","",sInput,"si")
Fila = ofv.GenRow("","","","","","","",Celda,"si")
Tabla = ofv.GenTabla("","","100% border","","","0","","0","0","","",Fila,"si")
Tabla = Tabla & "<hr size=1 color=#000099>"

if nBotones = 1 then
  sInput = ofv.GenerarInput("","Volver","submit","","10","bt","")
  Celda = ofv.GenCelda("","","center","50%","","","",sInput,"si")
  Formulario = ofv.GenForm("Mensajes","",sVolver & "&Eliminar=NO","","post","",Celda,"si")
else
  sInput = ofv.GenerarInput("","Eliminar","submit","","10","bt","")
  Celda = ofv.GenCelda("","","center","40%","","","",sInput,"si")
  Formulario = ofv.GenForm("Mensajes","",sVolver & "&Eliminar=SI","","post","",Celda,"si")

  sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
  Celda = ofv.GenCelda("","","center","40%","","","",sInput,"si")
  Formulario = Formulario & ofv.GenForm("Mensajesb","",sVolver & "&Eliminar=NO","","post","",Celda,"si")
end if
Fila = ofv.GenRow("","","","","","","",Formulario,"si")

Tabla = Tabla & ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

sHTML = sHTML & Tabla & "</body></html>"

response.write sHTML

%>