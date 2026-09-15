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

  set cn = ofv.conectar(ofv.strconn4)

  strsql = "select * from " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR TEMAS")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","Orden")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Codigo
  sInput = ofv.GenerarInput(rs(2).name,"0","agregar","","10","dtn MAXLENGTH=5","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0 || this."
  sFuncion = sFuncion & rs(2).name & ".value.length==0) "
  sFuncion = sFuncion & "{ alert( 'Los campos no deben ser nulos');return false;} else "
  sFuncion = sFuncion & "{if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name
  sFuncion = sFuncion & ".value < 0) {alert('Orden debe ser numerico'); this." & rs(2).name
  sFuncion = sFuncion & ".focus(); this." & rs(2).name & ".select();return false;};};" & chr(34)
  sAccion = "asp/agregar.asp?tabla=" & tabla
  sAccion = sAccion & "&volver=../Temas.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.GenForm("AgregarTema","",sAccion,"","post",sFuncion,Fila,"no")
'  fname=" & ofv.convertircar(rs(0)," ","*")
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Temas.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>