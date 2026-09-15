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

  sHTML = ofv.FormHeader("AGREGAR CIRCUITO")

  set cn = ofv.conectar(ofv.strconn4)

  Strsql = "Select * From " & tabla
  Clave = ofv.VectorClave(Strsql,2)
  sHTML = sHTML & ofv.ValClaveRep(Clave)

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
   
  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","Orden")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Codigo
  sInput = ofv.GenerarInput(rs(2).name,"0","agregar","","10","dtn MAXLENGTH=5" & " onblur=valida(this);","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","SI")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0 || this."
  sFuncion = sFuncion & rs(2).name & ".value.length==0) "
  sFuncion = sFuncion & "{ alert( 'Los campos no deben ser nulos');return false;} else "
  sFuncion = sFuncion & "{if (isNaN(this." & rs(2).name & ".value) || this." & rs(2).name
  sFuncion = sFuncion & ".value < 0) {alert('Orden debe ser numerico'); this." & rs(2).name
  sFuncion = sFuncion & ".focus(); this." & rs(2).name & ".select();return false;};};" & chr(34)
  sAccion = "asp/agregar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  sAccion = sAccion & "&tabla=" & tabla
  sAccion = sAccion & "&volver=../circuitos.asp?CantRegMover=" & Regisvuelta
  sAccion = sAccion & "_CantRegMostrar=10"
'  sAccion = sAccion & "&volver=../Procesosx.asp?Tipo=Circuitos_CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.GenForm("AgregarCir","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Circuitos.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=10"
'  sAccion = "Procesosx.asp?Tipo=Circuitos&CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"
end if


Response.write sHTML
'      sAccion = "AgregarSubtemas.asp?Regisvuelta=" & CantRegMoverx &  "&Opcion=si&tabla=PRODSERV&ID=" & Session("Tema") 
'      sAccion = sAccion & "&str=" & ofv.convertircar(Strsql," ","_")
'      Agregar = ofv.BotonAgregar(sAccion)

%>