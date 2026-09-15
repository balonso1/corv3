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

  Redim Perfiles(2)
  Perfiles(1) = "N"
  Perfiles(2) = "S"
  Redim Perfiles2(3)
  Perfiles2(1) = "A"
  Perfiles2(2) = "M"
  Perfiles2(3) = "B"

  set cn = ofv.conectar(ofv.strconn4)
 
  Strsql = "Select * From " & tabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("AGREGAR PROCESOS")

  'Encabezados
  Formulario = ofv.Encabezados("Procesos","Sec.","Cri.","Pri.")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","100","dt MAXLENGTH=50","")
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

  'Secuencia
  sInput = ofv.GenerarInput(rs(2).name,"","agregar","","5","dt MAXLENGTH=5","SI")
  Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

  'Criterio
  Combo = "<select name=" & rs(3).name & "><option value=nsnc>&nbsp;"
  For n = 1 To Ubound(Perfiles)
    Combo = Combo & "<option value=" & Perfiles(n) & ">" & Perfiles(n)
  Next
  Combo = Combo & "</select></font>"
  Celda = Celda & ofv.GenCelda("","","","","","","",Combo,"si")

  'Prioridad
  Combo = "<select name=" & rs(4).name & "><option value=nsnc>&nbsp;"
  For n = 1 To Ubound(Perfiles2)
    Combo = Combo & "<option value=" & Perfiles2(n) & ">" & Perfiles2(n)
  Next
  Combo = Combo & "</select></font>"
  Celda = Celda & ofv.GenCelda("","","","","","","",Combo,"si")

  sInput = ofv.GenerarInput(rs(5).name,ID,"hidden","","10","","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","SI")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion ="onsubmit= " & chr(34) & "if (this." & rs(1).name & ".value.length==0)"
  sFuncion = sFuncion & " { alert('Proceso no debe ser nulo');this." & rs(1).name 
  sFuncion = sFuncion & ".focus();return false;}else {if (this." & rs(2).name
  sFuncion = sFuncion & ".value.length==0){ alert( 'Secuencia no debe ser nula');this."
  sFuncion = sFuncion & rs(2).name & ".value=0;return false;} else {if (isNaN(this."
  sFuncion = sFuncion & rs(2).name & ".value) || this." & rs(2).name & ".value < 0)"
  sFuncion = sFuncion & "{alert('Secuencia debe ser numerica'); this." & rs(2).name
  sFuncion = sFuncion & ".focus(); this." & rs(2).name & ".select();return false;} else "
  sFuncion = sFuncion & "{ if (this." & rs(3).name & ".options[this." & rs(3).name
  sFuncion = sFuncion & ".selectedIndex].value=='nsnc') { alert( 'Critico no debe ser nulo');"
  sFuncion = sFuncion & "this." & rs(3).name & ".focus();return false;} else { if (this."
  sFuncion = sFuncion & rs(4).name & ".options[this." & rs(4).name
  sFuncion = sFuncion & ".selectedIndex].value=='nsnc') { alert( 'Prioridad no debe ser nulo');"
  sFuncion = sFuncion & "this." & rs(4).name & ".focus();return false;};};};};};" & chr(34)
  sAccion = "asp/agregar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  sAccion = sAccion & "&tabla=" & tabla & "&volver=../Procesos.asp?CantRegMover=" & Regisvuelta
  sAccion = sAccion & "_CantRegMostrar=10"
  Formulario = Formulario & ofv.GenForm("AgregarSub","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
 
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Procesos.asp?CantRegMover=" & Regisvuelta & "&CantRegMostrar=10"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>