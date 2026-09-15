
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

  sHTML = ofv.FormHeader(session("FORMN"))

  'Encabezados
  Formulario = ofv.Encabezados("Cod","descripcion")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From dominios Order By descripcion "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir AND rs(0)<>1 then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=dominios&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
     else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'codigo
    sInput = rs(0)

    Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")

    Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=dominios"
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "Agregardominios.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=dominios"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descripcion From dominios Order By descripcion"    ' orden pagina
  str2 = str1 '"Select observaciones From feriados Order By observaciones "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>