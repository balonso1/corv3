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
  ID = request.querystring("ID")
  IDx = request.querystring("IDx")
  Tema = request.querystring("Tema")
  xTema = request.querystring("xTema")
 
  set cn = ofv.conectar(ofv.strconn4)
 
  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR EVENTOS")

  'Encabezados
  Formulario = ofv.Encabezados("Evento","Secuencia","Orden")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","100","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Secuencia
  sInput = ofv.GenerarInput(rs(2).name,"","agregar","","10","dt MAXLENGTH=5","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
 
  'Codigo
  sInput = ofv.GenerarInput(rs(4).name,"","agregar","","10","dt MAXLENGTH=5","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  sInput = ofv.GenerarInput(rs(3).name,ID,"hidden","","10","","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","SI")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0 || this."
  sFuncion = sFuncion & rs(2).name & ".value.length==0 || this." & rs(4).name
  sFuncion = sFuncion & ".value.length==0){ alert( 'Los campos no deben ser nulos');"
  sFuncion = sFuncion & "return false;} else {if (isNaN(this." & rs(2).name & ".value) || this."
  sFuncion = sFuncion & rs(2).name & ".value < 0) {alert('Secuencia debe ser numerico'); this."
  sFuncion = sFuncion & rs(2).name & ".focus(); this." & rs(2).name & ".select();"
  sFuncion = sFuncion & "return false;} else {if (isNaN(this." & rs(4).name & ".value) || this."
  sFuncion = sFuncion & rs(4).name & ".value < 0) {alert('Orden debe ser numerico'); this."
  sFuncion = sFuncion & rs(4).name & ".focus(); this." & rs(4).name & ".select();"
  sFuncion = sFuncion & "return false;};};};" & chr(34)
  sAccion = "asp/agregar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  sAccion = sAccion & "&tabla=" & tabla & "&volver=../Eventos.asp?CantRegMover=" & Regisvuelta
  sAccion = sAccion & "_CantRegMostrar=10"
  Formulario = Formulario & ofv.GenForm("AgregarSub","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
 
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Eventos.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=10"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>