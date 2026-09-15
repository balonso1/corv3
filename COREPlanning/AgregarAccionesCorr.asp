<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Tabla = request.querystring("Tabla")
  vuelta = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")
  Proceso = request.querystring("Proceso")

  set cn = ofv.conectar(ofv.strconn4)

  strsql = "select * from " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR ACCION CORRECTIVA")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt MAXLENGTH=100","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  if Proceso = "si" then 
    sVolver = "../" & vuelta
  else
    sVolver = "../AccionesCorr.asp?CantRegMover=" & Regisvuelta
  end if
  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0) {alert('Los campos no deben ser nulos');"
  sFuncion = sFuncion & "return false;}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=" & sVolver
  Formulario = Formulario & ofv.GenForm("Agregar","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  if Proceso = "si" then
     sAccion = ofv.convertircar(vuelta,"_","&")
  else
     sAccion = "AccionesCorr.asp?CantRegMover=" & Regisvuelta
  end if
 
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>