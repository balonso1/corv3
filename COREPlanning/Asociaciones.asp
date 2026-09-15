<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader("ASOCIACIONES")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion")

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn4)

  Strsql = "Select * From Asociacion Order By Descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
    sHTML = sHTML & "<p align=center><font face=tahoma size=1 color=#990000><b>"
    sHTML = sHTML & "NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
  else
    rs.move(CantRegMoverx)
    posi=0
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

      'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Asociacion&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(1),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      'Descripcion
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"150","dt MAXLENGTH=50","")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      Celda = Celda & ofv.gencelda("","","","","","","","","si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      sAccion = ""
      if bModificar then
        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
        sAccion = sAccion & "&tabla=Asociacion&volver=" & vuelta
        sFunc = "onsubmit=" & chr(34) & "if (this." & rs(1).name & ".value.length==0 ) "
        sFunc = sFunc & "{ alert( 'Descripcion no debe ser nulo');this."
        sFunc = sFunc & rs(1).name & ".focus(); return false;} ;" & chr(34)
      end if
      Formulario = Formulario & ofv.GenForm(SForm,"",sAccion,"","post",sFunc,Fila,"si")

      posi=posi+1
      rs.Movenext
    Loop
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
  end if

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarAsociaciones.asp?Regisvuelta=" & CantRegMoverx
    sAccion = sAccion & "&Opcion=si&tabla=Asociacion"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Asociacion Order By descri "    ' orden pagina
  str2 = "Select descri From Asociacion Order By descri "  ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>