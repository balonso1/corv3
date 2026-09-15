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

  sHTML = ofv.FormHeader("AGREGAR SUBTEMAS")

  set cn = ofv.conectar(ofv.strconn4)

  strsql = "Select * From " & tabla & " Where CodTema=" & ID & " Order By Indb"
  Clave = ofv.VectorClave(Strsql,3)
  sHTML = sHTML & ofv.ValClaveRep(Clave)

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  'Encabezados
  Formulario = ofv.Encabezados("Subtemas","Orden")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","110","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Codigo
  sInput = ofv.GenerarInput(rs(3).name,"0","agregar","","10","dtn MAXLENGTH=5" & " onblur=valida(this);","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  sInput = ofv.GenerarInput(rs(2).name,ID,"hidden","","10","","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0 || this."
  sFuncion = sFuncion & rs(3).name & ".value.length==0) "
  sFuncion = sFuncion & "{ alert( 'Los campos no deben ser nulos');return false;} else "
  sFuncion = sFuncion & "{if (isNaN(this." & rs(3).name & ".value) || this." & rs(3).name
  sFuncion = sFuncion & ".value < 0) {alert('Orden debe ser numerico'); this." & rs(3).name
  sFuncion = sFuncion & ".focus(); this." & rs(3).name & ".select();return false;};};" & chr(34)
  sAccion = "asp/agregar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  sAccion = sAccion & "&tabla=" & tabla & "&volver=../Subtemas.asp?CantRegMover=" & Regisvuelta
  sAccion = sAccion & "_CantRegMostrar=10"
  Formulario = Formulario & ofv.GenForm("AgregarSub","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
 
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Subtemas.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=10"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>