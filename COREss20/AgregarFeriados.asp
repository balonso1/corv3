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

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR Feriado")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion")

  'Segmento
  Celda = ofv.gencelda("","","","","","","","","si")
  'Descripcion
  sInput = ofv.GenerarInput(rs(2).name,"","agregar","","120","dt","")
  sInput1 = ofv.GenerarInput(rs(1).name,"0","hidden","","120","dt","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput & sInput1,"si")

 ' Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if ("
  sFuncion = sFuncion & "this." & rs(2).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('Observacion no debe ser nulo');return false;} "
  sFuncion = sFuncion & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla & "&volver=../Feriados.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregar","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "feriados.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>