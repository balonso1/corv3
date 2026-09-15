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

  sHTML = ofv.FormHeader("AGREGAR CASOS DE CONTINGENCIA")

  Strsql = "Select * from Tipocont"
  Clave = ofv.VectorClave(Strsql,0)
  sHTML = sHTML & ofv.ValClaveRep(Clave)

  'Encabezados
  Formulario = ofv.Encabezados("Cod","Descripcion","Asociacion")

  Strsql = "Select * from Tipocont"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  'Codigo
  Celda = ofv.gencelda("","","","","","","","","si")
  val = " onblur=" & chr(34) & " valida(this);" & chr(34) 
  sInput = ofv.GenerarInput(rs(0).name,"","agregar","AgregarCont","5","dt MAXLENGTH=5" & val,"")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Descripcion
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","AgregarCont","70","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Asociacion
  Strsql = "Select * From Asociacion Order By descri"
  sInput = ofv.GenerarCombo("AgregarCont",rs(2).name,"",Strsql,cn,0,1,"no")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(2).name & ".options[this."
  sFuncion = sFuncion & rs(2).name & ".selectedIndex].value=='nsnc' ||  this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0) {alert('Los campos no deben ser nulos');"
  sFuncion = sFuncion & "return false;} else {if (isNaN(this." & rs(0).name & ".value) || this."
  sFuncion = sFuncion & rs(0).name & ".value < 0 ) {alert('Codigo debe ser numerico'); this."
  sFuncion = sFuncion & rs(0).name & ".focus(); this." & rs(0).name & ".select();"
  sFuncion = sFuncion & "return false;}; };" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=../Contingencias.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregar","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Contingencias.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>