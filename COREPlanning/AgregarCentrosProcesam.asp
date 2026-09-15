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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn4)

  Redim Preserve Perfiles(2)
  Perfiles(1) = "N"
  Perfiles(2) = "S"
 
  strsql = "select * from " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sHTML = ofv.FormHeader("AGREGAR CENTRO DE PROCESAMIENTO")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","Existe")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","120","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Existe
  sInput = "<font class=at><select name=" & rs(2).name & "><option value=nsnc>&nbsp;"
  For n = 1 To Ubound(Perfiles)
    sInput = sInput & "<option value=" & Perfiles(n) & ">" & Perfiles(n)
  Next
  sInput = sInput & "</select>"
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
   
  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(2).name & ".options[this." & rs(2).name
  sFuncion = sFuncion & ".selectedIndex].value== 'nsnc' || this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0) { alert( 'Los campos no deben ser nulos');"
  sFuncion = sFuncion & "return false;} else {return true;};" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=../CentrosProcesam.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregar","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "CentrosProcesam.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

Response.write sHTML

%>