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
  Subtema = request.querystring("Subtema")
  xSubtema = request.querystring("xSubtema")
  Evento = request.querystring("Evento")
  xEvento = request.querystring("xEvento")

  set cn = ofv.conectar(ofv.strconn4)
 
  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR CIRCUITO DE CONTINGENCIA")

  'Encabezados
  Formulario = ofv.Encabezados("CONTINGENCIAS")

  Celda = ofv.gencelda("","","","","","","","","si")
  Strsql = "Select Tipocont.*, Asociacion.descri From Tipocont,Asociacion "
  Strsql = Strsql & "Where Asociacion.codasoc = Tipocont.codasoc Order By Tipocont.descri"
  sID = "document.form1"
  sCombo = ofv.GenerarCombo(sID,rs(1).name,"",Strsql,cn,0,1,"no")
  Celda =  Celda & ofv.GenCelda("","","","","","","",sCombo,"si")

  sInput = ofv.GenerarInput(rs(2).name,ID,"hidden","","10","","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".options[" & rs(1).name
  sFuncion = sFuncion & ".selectedIndex].value=='nsnc') "
  sFuncion = sFuncion & "{ alert( 'Cicuito de Contingencia no puede ser nulo');this."
  sFuncion = sFuncion & rs(1).name & ".focus();return false;};" & chr(34)
  sAccion = "asp/agregar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  sAccion = sAccion & "&tabla=" & tabla & "&volver=../CircContingencia.asp?CantRegMover="
  sAccion = sAccion & Regisvuelta & "_CantRegMostrar=10"
  Formulario = Formulario & ofv.GenForm("form1","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
 
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "CircContingencia.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=10"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>