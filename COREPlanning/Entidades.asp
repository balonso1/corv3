<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader("ENTIDADES DE CONTINGENCIA")

  'Encabezados
  Formulario = ofv.Encabezados("Descripcion","Codigo")

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn4)

  Strsql = "Select * from Eco Order By Descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
    sHTML = sHTML & "<p align=center><font face=tahoma size=1 color=#990000><b>"
    sHTML = sHTML & "NO HAY ELEMENTOS ASOCIADOS</b><br><br></FONT></P>"
  else 
    rs.move(CantRegMoverx)
    posI=0
    Do While Not rs.EOF and posi < CantRegMostrarx
      sForm= "form" & posi

      'Eliminar
      if bSuprimir then
        sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
        sHref = sHref & "&Tabla=Eco&volver=" & vuelta & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs(1),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      'Descripcion
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"120","dt MAXLENGTH=50","")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      'Codigo
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt MAXLENGTH=5","SI")
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      Celda = Celda & ofv.gencelda("","","","","","","","","si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      sAccion = ""
      if bModificar then
        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
        sAccion = sAccion & "&tabla=Eco&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

      rs.Movenext
      posi=posi+1
    Loop
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
  End If

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarEntidades.asp?Regisvuelta=" & CantRegMoverx
    sAccion = sAccion & "&Opcion=si&tabla=Eco"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Eco Order By descri "    ' orden pagina
  str2 = "Select descri From Eco Order By descri "  ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>