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

  Strsql = "Select * From Asociacion Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  n = 1
  Do until rs.Eof
    Redim Preserve Perfiles(n)
    Redim Preserve Perfilesd(n)
    Perfiles(n) = rs(0)
    Perfilesd(n) = rs(1)
    rs.Movenext
    n = n + 1
  Loop
  call ofv.cerrarconsulta(rs)

  Strsql = "Select secuen from Niveles Order By secuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    pepe = int(rs(0)) + 10
    rs.movenext
  loop
' Calculo para que siempre termine con cero
  pepe = mid(pepe,1,Len(pepe)-1) & "0"
  call ofv.cerrarconsulta(rs)

  Strsql = "Select * from Niveles Order By Codasoc,secuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sHTML = ofv.FormHeader("AGREGAR NIVELES DE ASOCIACION")

  'Encabezados
  Formulario = ofv.Encabezados("Asociacion","Sec","Descripcion")

  'Descripcion
  Celda = ofv.gencelda("","","","","","","","","si")
  sInput = "<font class=at><select name=" & rs(3).name & "><option value=nsnc>&nbsp;"
  For n = 1 To Ubound(Perfiles)
      sInput = sInput & "<option value=" & Perfiles(n) & ">" & Perfilesd(n)
  Next
  sInput = sInput & "</select>"
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Secuencia
  sInput = ofv.GenerarInput(rs(2).name,pepe,"agregar","","5","dt MAXLENGTH=5","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  'Descripcion
  sInput = ofv.GenerarInput(rs(1).name,"","agregar","","60","dt MAXLENGTH=50","")
  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

  Celda = Celda & ofv.gencelda("","","","","","","","","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(1).name
  sFuncion = sFuncion & ".value.length==0 || this." & rs(3).name & ".options[this."
  sFuncion = sFuncion & rs(3).name & ".selectedIndex].value=='nsnc' ) "
  sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}" & chr(34)
  sAccion = "asp/Agregar.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=../Niveles.asp?CantRegMover=" & Regisvuelta
  Formulario = Formulario & ofv.GenForm("Agregar","",sAccion,"","post",sFuncion,Fila,"no")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  'Aceptar y Cancelar
  sAccion = "Niveles.asp?CantRegMover=" & Regisvuelta
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>